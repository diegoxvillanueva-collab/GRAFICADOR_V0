"""
S4 Graficador — Slide Builder
Construye el PPTX completo copiando slides del template y ensamblando con charts inyectados.
Opera directamente sobre el ZIP en memoria (sin win32com — corre en Linux/Railway).
"""

import zipfile
import zlib
import struct
import re
from io import BytesIO
from copy import deepcopy
from lxml import etree
from typing import List, Dict, Tuple, Optional
from itertools import groupby

from .models import Question, ParseResult
from .chart_injector import inject_chart_data


def _serialize_xml(root) -> bytes:
    """
    Serializa un elemento lxml a bytes con comillas dobles en la declaración XML.
    PowerPoint requiere: <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    lxml genera:         <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
    """
    raw = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    # Fix single quotes to double quotes in XML declaration only (first line)
    if raw.startswith(b"<?xml"):
        newline_pos = raw.find(b"?>")
        if newline_pos != -1:
            decl = raw[:newline_pos + 2]
            rest = raw[newline_pos + 2:]
            decl = decl.replace(b"'", b'"')
            raw = decl + rest
    return raw


# ──────────────────────────────────────────────────────────────────────────────
# TEMPLATE SLIDE MAPPING
# ──────────────────────────────────────────────────────────────────────────────

# Mapeo: tipo_slide -> (slide filename en template, chart filename en template)
# Basado en el análisis del Template.pptx
TEMPLATE_MAP = {
    "FREC_SIMPLE": {
        "slide_file": "ppt/slides/slide5.xml",
        "chart_files": ["ppt/charts/chart9.xml"],
        "chart_type": "clustered_col",
    },
    "FREC_MULTIPLE": {
        "slide_file": "ppt/slides/slide1.xml",
        "chart_files": ["ppt/charts/chart1.xml"],
        "chart_type": "percentStacked_col",
    },
    "FREC_MULTI_GRAFICOS": {
        "slide_file": "ppt/slides/slide3.xml",
        "chart_files": ["ppt/charts/chart3.xml", "ppt/charts/chart4.xml"],
        "chart_type": "multi_clustered_col",
    },
    "APERTURA_SIMPLE": {
        "slide_file": "ppt/slides/slide6.xml",
        "chart_files": ["ppt/charts/chart10.xml"],
        "chart_type": "percentStacked_bar",
    },
}

# Namespaces
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_PRES = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_DRAW = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
NS_SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"

CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"


# ──────────────────────────────────────────────────────────────────────────────
# SLIDE ORDERING LOGIC
# ──────────────────────────────────────────────────────────────────────────────

def _compute_slide_order(questions: List[Question]) -> List[dict]:
    """
    Calcula el orden de slides a generar basándose en capítulos, grupos y reglas de CAP_APP.

    Retorna una lista de dicts con:
        - type: "FREC" | "APERTURA"
        - chart_type: el tipo de chart a usar
        - question: Question (o None para FREC_MULTIPLE con grupo)
        - questions_group: List[Question] (para FREC_MULTIPLE)
        - titulo: título de la slide
        - capitulo: capítulo al que pertenece
    """
    slides = []

    # Agrupar por capítulo manteniendo orden de aparición
    cap_order = []
    seen_caps = set()
    for q in questions:
        if q.capitulo not in seen_caps:
            cap_order.append(q.capitulo)
            seen_caps.add(q.capitulo)

    # Recoger capítulos de apertura separados (CAP_APP)
    app_chapters = {}  # {cap_app: [questions]}
    for q in questions:
        if q.cap_app:
            app_chapters.setdefault(q.cap_app, []).append(q)

    # Generar slides por capítulo de frecuencia
    for cap in cap_order:
        cap_questions = [q for q in questions if q.capitulo == cap]

        # Agrupar por grupo_frec
        grupos = {}
        for q in cap_questions:
            grupos.setdefault(q.grupo_frec, []).append(q)

        for grupo_id in sorted(grupos.keys()):
            grupo_qs = grupos[grupo_id]
            tipo = grupo_qs[0].tipo_slide

            # --- FRECUENCIA ---
            if tipo == "FREC_SIMPLE":
                for q in grupo_qs:
                    slides.append({
                        "type": "FREC",
                        "chart_type": "FREC_SIMPLE",
                        "question": q,
                        "questions_group": None,
                        "titulo_slide": f"{q.titulo_frec} | {q.titulo_app}",
                        "titulo_chart": q.titulo_app,
                        "capitulo": cap,
                    })
                    if not q.cap_app:
                        slides.append({
                            "type": "APERTURA",
                            "chart_type": "APERTURA_SIMPLE",
                            "question": q,
                            "questions_group": None,
                            "titulo_slide": q.titulo_app,
                            "titulo_chart": q.titulo_app,
                            "capitulo": cap,
                        })

            elif tipo == "FREC_MULTIPLE":
                slides.append({
                    "type": "FREC",
                    "chart_type": "FREC_MULTIPLE",
                    "question": None,
                    "questions_group": grupo_qs,
                    "titulo_slide": f"{grupo_qs[0].titulo_frec} | {grupo_qs[0].titulo_app}",
                    "titulo_chart": grupo_qs[0].titulo_app,
                    "capitulo": cap,
                })
                for q in grupo_qs:
                    if not q.cap_app:
                        slides.append({
                            "type": "APERTURA",
                            "chart_type": "APERTURA_SIMPLE",
                            "question": q,
                            "questions_group": None,
                            "titulo_slide": q.titulo_app,
                            "titulo_chart": q.titulo_app,
                            "capitulo": cap,
                        })

            elif tipo == "FREC_MULTI_GRAFICOS":
                slides.append({
                    "type": "FREC",
                    "chart_type": "FREC_MULTI_GRAFICOS",
                    "question": None,
                    "questions_group": grupo_qs,
                    "titulo_slide": f"{grupo_qs[0].titulo_frec} | {grupo_qs[0].titulo_app}",
                    "titulo_chart": grupo_qs[0].titulo_app,
                    "capitulo": cap,
                })
                for q in grupo_qs:
                    if not q.cap_app:
                        slides.append({
                            "type": "APERTURA",
                            "chart_type": "APERTURA_SIMPLE",
                            "question": q,
                            "questions_group": None,
                            "titulo_slide": q.titulo_app,
                            "titulo_chart": q.titulo_app,
                            "capitulo": cap,
                        })

    # Capítulos de apertura separados (CAP_APP) al final
    for cap_app in sorted(app_chapters.keys()):
        for q in app_chapters[cap_app]:
            slides.append({
                "type": "APERTURA",
                "chart_type": "APERTURA_SIMPLE",
                "question": q,
                "questions_group": None,
                "titulo_slide": q.titulo_app,
                "titulo_chart": q.titulo_app,
                "capitulo": cap_app,
            })

    return slides


# ──────────────────────────────────────────────────────────────────────────────
# PPTX BUILDER
# ──────────────────────────────────────────────────────────────────────────────

class PptxBuilder:
    """
    Construye un PPTX nuevo copiando slides de un template y inyectando datos.
    Opera sobre el ZIP en memoria (BytesIO).
    """

    def __init__(self, template_bytes: bytes):
        """Carga el template PPTX en memoria."""
        self.template_zip = zipfile.ZipFile(BytesIO(template_bytes), "r")
        self.files: Dict[str, bytes] = {}  # path -> content
        self.slide_counter = 0
        self.chart_counter = 0
        self.rel_counter = 100  # Start high to avoid conflicts
        self.slide_ids = []  # (slide_path, rId) tuples for presentation.xml

        # Cargar todos los archivos del template base
        # El template puede tener archivos con CRC error (e.g., image4.jpeg)
        self._template_raw = BytesIO(template_bytes)
        for info in self.template_zip.infolist():
            try:
                self.files[info.filename] = self.template_zip.read(info.filename)
            except (zipfile.BadZipFile, KeyError, zlib.error):
                # CRC error — extraer raw bytes directamente del ZIP
                try:
                    data = self._extract_raw(info)
                    if data:
                        self.files[info.filename] = data
                except Exception:
                    pass

        # Parsear presentation.xml para obtener el último slide ID
        pres_xml = etree.fromstring(self.files["ppt/presentation.xml"])
        self._max_slide_id = 300  # Start high to avoid conflicts
        self._pres_ns = {
            "p": NS_PRES,
            "r": NS_REL,
        }

    def build(self, questions: List[Question], segment_groups=None,
              pie_pagina: str = "") -> bytes:
        """
        Construye el PPTX completo.
        """
        self._segment_groups = segment_groups
        self._pie_pagina = pie_pagina
        slide_order = _compute_slide_order(questions)

        self._remove_template_slides()

        for slide_spec in slide_order:
            self._add_slide(slide_spec)

        # Actualizar presentation.xml con las nuevas slides
        self._update_presentation_xml()

        # Actualizar [Content_Types].xml
        self._update_content_types()

        # Empaquetar como PPTX (ZIP)
        return self._package()

    def _remove_template_slides(self):
        """Remueve las slides del template del output."""
        # Encontrar todas las slides y sus charts en el template
        to_remove = []
        for path in list(self.files.keys()):
            if re.match(r"ppt/slides/slide\d+\.xml$", path):
                to_remove.append(path)
            elif re.match(r"ppt/slides/_rels/slide\d+\.xml\.rels$", path):
                to_remove.append(path)
            elif re.match(r"ppt/charts/chart\d+\.xml$", path):
                to_remove.append(path)
            elif re.match(r"ppt/charts/_rels/chart\d+\.xml\.rels$", path):
                to_remove.append(path)
            elif re.match(r"ppt/charts/style\d+\.xml$", path):
                to_remove.append(path)
            elif re.match(r"ppt/charts/colors\d+\.xml$", path):
                to_remove.append(path)

        for path in to_remove:
            del self.files[path]

    def _add_slide(self, slide_spec: dict):
        """Agrega una slide al output."""
        chart_type = slide_spec["chart_type"]
        template_info = TEMPLATE_MAP.get(chart_type)

        if template_info is None:
            return  # Tipo desconocido

        self.slide_counter += 1
        new_slide_name = f"ppt/slides/slide{self.slide_counter}.xml"
        new_slide_rels_name = f"ppt/slides/_rels/slide{self.slide_counter}.xml.rels"

        # Leer slide template del ZIP original
        src_slide = template_info["slide_file"]
        try:
            slide_xml = self.template_zip.read(src_slide)
        except (KeyError, zipfile.BadZipFile):
            return

        # Leer rels del slide template
        src_rels = src_slide.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        try:
            rels_xml = self.template_zip.read(src_rels)
        except (KeyError, zipfile.BadZipFile):
            rels_xml = self._make_empty_rels()

        # Para FREC_MULTI_GRAFICOS con N preguntas, necesitamos N charts
        if chart_type == "FREC_MULTI_GRAFICOS" and slide_spec["questions_group"]:
            n_charts = len(slide_spec["questions_group"])
            chart_paths = self._copy_multi_charts(
                template_info, slide_spec, n_charts
            )
            # Actualizar rels para apuntar a los nuevos charts
            rels_xml = self._rebuild_slide_rels_multi(rels_xml, chart_paths)
            # Actualizar slide XML para tener N graphicFrames
            slide_xml = self._rebuild_slide_multi_charts(
                slide_xml, n_charts, chart_paths, rels_xml
            )
        else:
            # Caso normal: 1 chart por slide
            src_chart = template_info["chart_files"][0]
            self.chart_counter += 1
            new_chart_name = f"ppt/charts/chart{self.chart_counter}.xml"

            # Copiar chart XML e inyectar datos
            try:
                chart_xml = self.template_zip.read(src_chart)
            except (KeyError, zipfile.BadZipFile):
                return

            chart_xml = self._inject_chart(chart_xml, slide_spec)
            self.files[new_chart_name] = chart_xml

            # Copiar chart rels
            src_chart_rels = src_chart.replace("ppt/charts/", "ppt/charts/_rels/") + ".rels"
            try:
                chart_rels = self.template_zip.read(src_chart_rels)
                new_chart_rels = new_chart_name.replace("ppt/charts/", "ppt/charts/_rels/") + ".rels"
                self.files[new_chart_rels] = chart_rels
            except (KeyError, zipfile.BadZipFile):
                pass

            # Copiar style y colors del chart si existen
            self._copy_chart_style_files(src_chart, new_chart_name)

            # Actualizar rels del slide para apuntar al nuevo chart
            rels_xml = self._update_slide_rels(rels_xml, src_chart, new_chart_name)

        # Limpiar shapes sobrantes del template (burbujas, textboxes extra, conectores)
        slide_xml = self._clean_template_shapes(slide_xml)

        if chart_type == "APERTURA_SIMPLE":
            slide_xml = self._expand_chart_position(slide_xml)

        # Agregar línea separadora entre charts en multi-gráficos
        if chart_type == "FREC_MULTI_GRAFICOS" and slide_spec.get("questions_group"):
            n = len(slide_spec["questions_group"])
            if n > 1:
                slide_xml = self._add_separator_lines(slide_xml, n)

        # Inyectar título en el slide (titulo_slide, no titulo_chart)
        titulo_slide = slide_spec.get("titulo_slide", "")
        slide_xml = self._inject_slide_title(slide_xml, titulo_slide)

        # Inyectar pie de página
        if self._pie_pagina:
            slide_xml = self._inject_footer(slide_xml, self._pie_pagina)

        # Guardar slide y rels
        self.files[new_slide_name] = slide_xml
        self.files[new_slide_rels_name] = rels_xml

        # Registrar para presentation.xml
        self._max_slide_id += 1
        self.rel_counter += 1
        self.slide_ids.append({
            "slide_path": new_slide_name,
            "slide_id": self._max_slide_id,
            "rId": f"rId{self.rel_counter}",
        })

    def _inject_chart(self, chart_xml: bytes, slide_spec: dict) -> bytes:
        """Inyecta datos en un chart XML según la especificación."""
        chart_type = slide_spec["chart_type"]
        question = slide_spec.get("question")
        questions_group = slide_spec.get("questions_group")
        titulo_chart = slide_spec.get("titulo_chart", "")

        # Determinar orientación
        orientacion = "V"
        if question:
            orientacion = getattr(question, "orientacion", "V")
        elif questions_group:
            orientacion = getattr(questions_group[0], "orientacion", "V")

        return inject_chart_data(
            chart_xml_bytes=chart_xml,
            chart_type=chart_type,
            question=question,
            questions_group=questions_group,
            title=titulo_chart,
            orientacion=orientacion,
            segment_groups=self._segment_groups,
        )

    def _copy_multi_charts(self, template_info, slide_spec, n_charts) -> List[str]:
        """Copia N charts para FREC_MULTI_GRAFICOS."""
        chart_paths = []
        src_chart = template_info["chart_files"][0]  # Usar el primer chart como base

        for i, q in enumerate(slide_spec["questions_group"]):
            self.chart_counter += 1
            new_chart_name = f"ppt/charts/chart{self.chart_counter}.xml"

            try:
                chart_xml = self.template_zip.read(src_chart)
            except (KeyError, zipfile.BadZipFile):
                continue

            # Inyectar datos de esta pregunta individual
            # Título: usar titulo_app (la pregunta) que es único por chart
            chart_xml = inject_chart_data(
                chart_xml_bytes=chart_xml,
                chart_type="FREC_MULTI_GRAFICOS",
                question=q,
                title=q.titulo_app,
            )

            self.files[new_chart_name] = chart_xml

            # Copiar rels del chart
            src_chart_rels = src_chart.replace("ppt/charts/", "ppt/charts/_rels/") + ".rels"
            try:
                chart_rels = self.template_zip.read(src_chart_rels)
                new_chart_rels = new_chart_name.replace("ppt/charts/", "ppt/charts/_rels/") + ".rels"
                self.files[new_chart_rels] = chart_rels
            except (KeyError, zipfile.BadZipFile):
                pass

            self._copy_chart_style_files(src_chart, new_chart_name)
            chart_paths.append(new_chart_name)

        return chart_paths

    def _rebuild_slide_rels_multi(self, rels_xml: bytes, chart_paths: List[str]) -> bytes:
        """Reconstruye las rels de un slide para múltiples charts."""
        root = etree.fromstring(rels_xml)

        # Remover relaciones a charts existentes
        for rel in root.findall(f"{{{NS_RELS}}}Relationship"):
            if NS_CHART_REL in rel.get("Type", ""):
                root.remove(rel)

        # Agregar relaciones a los nuevos charts
        for i, chart_path in enumerate(chart_paths):
            rel = etree.SubElement(root, f"{{{NS_RELS}}}Relationship")
            rel.set("Id", f"rIdChart{i+1}")
            rel.set("Type", NS_CHART_REL)
            rel.set("Target", f"../{chart_path.replace('ppt/', '')}")

        return _serialize_xml(root)

    def _rebuild_slide_multi_charts(self, slide_xml: bytes, n_charts: int,
                                     chart_paths: List[str], rels_xml: bytes) -> bytes:
        """Reconstruye un slide XML con N graphicFrames para multi-charts."""
        root = etree.fromstring(slide_xml)
        ns_p = {"p": NS_PRES, "a": NS_DRAW, "r": NS_REL}

        # Encontrar el spTree (shape tree)
        spTree = root.find(f".//{{{NS_PRES}}}cSld/{{{NS_PRES}}}spTree")
        if spTree is None:
            # Intentar sin namespace prefix
            spTree = root.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cSld/"
                              "{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
        if spTree is None:
            return slide_xml

        # Encontrar graphicFrames existentes
        gf_ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}graphicFrame"
        # Also check for non-namespaced
        graphic_frames = spTree.findall(f"{{{NS_PRES}}}graphicFrame")
        if not graphic_frames:
            return slide_xml

        # Usar el primer frame como template
        template_gf = graphic_frames[0]

        # Remover todos los frames existentes excepto el primero
        for gf in graphic_frames[1:]:
            spTree.remove(gf)

        # Calcular posiciones distribuidas
        # Slide width ~ 24382413 EMUs
        slide_width = 24382413
        margin = 800000  # ~0.87 inches
        gap = 200000  # gap between charts
        available_width = slide_width - 2 * margin - (n_charts - 1) * gap
        chart_width = available_width // n_charts

        for i in range(n_charts):
            if i == 0:
                gf = template_gf
            else:
                gf = deepcopy(template_gf)
                spTree.append(gf)

            # Actualizar posición y tamaño
            xfrm = gf.find(f".//{{{NS_PRES}}}xfrm")
            if xfrm is None:
                xfrm = gf.find(f".//{{{NS_DRAW}}}xfrm")
            if xfrm is not None:
                off = xfrm.find(f"{{{NS_DRAW}}}off")
                if off is not None:
                    x_pos = margin + i * (chart_width + gap)
                    off.set("x", str(x_pos))
                ext = xfrm.find(f"{{{NS_DRAW}}}ext")
                if ext is not None:
                    ext.set("cx", str(chart_width))

            # Actualizar referencia al chart
            chart_el = gf.find(f".//{{{NS_DRAW}}}graphic/{{{NS_DRAW}}}graphicData/"
                               "{http://schemas.openxmlformats.org/drawingml/2006/chart}chart")
            if chart_el is None:
                # Try with c: namespace
                chart_el = gf.find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart")
            if chart_el is not None:
                chart_el.set(f"{{{NS_REL}}}id", f"rIdChart{i+1}")

        return _serialize_xml(root)

    def _update_slide_rels(self, rels_xml: bytes, src_chart: str, new_chart: str) -> bytes:
        """Actualiza las rels del slide para apuntar al nuevo chart."""
        root = etree.fromstring(rels_xml)

        src_target = f"../{src_chart.replace('ppt/', '')}"
        new_target = f"../{new_chart.replace('ppt/', '')}"

        for rel in root.findall(f"{{{NS_RELS}}}Relationship"):
            target = rel.get("Target", "")
            if NS_CHART_REL in rel.get("Type", ""):
                # Actualizar al nuevo chart
                rel.set("Target", new_target)
                break

        return _serialize_xml(root)

    def _inject_slide_title(self, slide_xml: bytes, titulo: str) -> bytes:
        """Inyecta el título en el placeholder de texto del slide."""
        if not titulo:
            return slide_xml

        root = etree.fromstring(slide_xml)

        # Buscar shapes de texto en el slide
        for sp in root.iter(f"{{{NS_PRES}}}sp"):
            nvSpPr = sp.find(f"{{{NS_PRES}}}nvSpPr")
            if nvSpPr is not None:
                nvPr = nvSpPr.find(f"{{{NS_PRES}}}nvPr")
                if nvPr is not None:
                    ph = nvPr.find(f"{{{NS_PRES}}}ph")
                    if ph is not None:
                        ph_type = ph.get("type", "")
                        if ph_type in ("title", "ctrTitle", ""):
                            txBody = sp.find(f"{{{NS_PRES}}}txBody")
                            if txBody is None:
                                continue
                            # Limpiar TODOS los párrafos y runs, poner solo uno
                            paragraphs = txBody.findall(f"{{{NS_DRAW}}}p")
                            if paragraphs:
                                # Usar el primer párrafo, mantener su pPr
                                first_p = paragraphs[0]
                                pPr = first_p.find(f"{{{NS_DRAW}}}pPr")
                                # Encontrar un rPr template de algún run existente
                                rPr_template = None
                                for r in first_p.findall(f"{{{NS_DRAW}}}r"):
                                    rPr_template = r.find(f"{{{NS_DRAW}}}rPr")
                                    break

                                # Limpiar todos los runs y textos del primer párrafo
                                for child in list(first_p):
                                    tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                                    if tag in ("r", "br", "fld"):
                                        first_p.remove(child)

                                # Agregar un solo run con el título
                                # Formato: HelveticaNeue Std (Cuerpo) 36pt bold
                                from .models import AGORA_STYLE
                                new_r = etree.SubElement(first_p, f"{{{NS_DRAW}}}r")
                                new_rPr = etree.SubElement(new_r, f"{{{NS_DRAW}}}rPr")
                                new_rPr.set("lang", "es-AR")
                                new_rPr.set("sz", str(AGORA_STYLE["slide_title_size"]))
                                new_rPr.set("b", "1" if AGORA_STYLE["slide_title_bold"] else "0")
                                latin = etree.SubElement(new_rPr, f"{{{NS_DRAW}}}latin")
                                latin.set("typeface", AGORA_STYLE["slide_title_font"])
                                new_t = etree.SubElement(new_r, f"{{{NS_DRAW}}}t")
                                new_t.text = titulo

                                # Eliminar párrafos adicionales
                                for p in paragraphs[1:]:
                                    txBody.remove(p)

                            return _serialize_xml(root)

        return slide_xml

    def _clean_template_shapes(self, slide_xml: bytes) -> bytes:
        """
        Remueve shapes sobrantes del template: textboxes sueltos, grupos,
        conectores del template. Mantiene solo: título, footer, slideNum, graphicFrames.
        """
        root = etree.fromstring(slide_xml)
        spTree = root.find(f".//{{{NS_PRES}}}cSld/{{{NS_PRES}}}spTree")
        if spTree is None:
            return slide_xml

        to_remove = []
        for child in list(spTree):
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if tag == "sp":
                # Mantener solo placeholders conocidos (title, ftr, sldNum)
                nvSpPr = child.find(f"{{{NS_PRES}}}nvSpPr")
                if nvSpPr is not None:
                    nvPr = nvSpPr.find(f"{{{NS_PRES}}}nvPr")
                    if nvPr is not None:
                        ph = nvPr.find(f"{{{NS_PRES}}}ph")
                        if ph is not None:
                            ph_type = ph.get("type", "")
                            if ph_type in ("title", "ctrTitle", "ftr", "sldNum"):
                                continue  # mantener
                # Si no es placeholder conocido, remover (textbox suelto)
                to_remove.append(child)

            elif tag == "grpSp":
                # Remover groups (burbujas, decoraciones)
                to_remove.append(child)

            elif tag == "cxnSp":
                # Remover conectores del template (los recramos nosotros)
                to_remove.append(child)

        for el in to_remove:
            spTree.remove(el)

        return _serialize_xml(root)

    def _expand_chart_position(self, slide_xml: bytes) -> bytes:
        """
        Expande el graphicFrame del chart para usar buen ancho del slide.
        Referencia: chart en template x=1149992, slide width ~24384000 EMUs.
        """
        root = etree.fromstring(slide_xml)

        for gf in root.iter(f"{{{NS_PRES}}}graphicFrame"):
            xfrm = gf.find(f".//{{{NS_PRES}}}xfrm")
            if xfrm is None:
                xfrm = gf.find(f".//{{{NS_DRAW}}}xfrm")
            if xfrm is None:
                continue

            off = xfrm.find(f"{{{NS_DRAW}}}off")
            ext = xfrm.find(f"{{{NS_DRAW}}}ext")
            if off is not None and ext is not None:
                # Margen izq ~1cm, derecho ~1.5cm para aire
                off.set("x", "500000")       # ~0.5cm desde la izq
                off.set("y", "2100000")      # debajo del título
                ext.set("cx", "23200000")    # buen ancho con aire a la derecha
                ext.set("cy", "9800000")     # dejar espacio abajo para labels

        return _serialize_xml(root)

    def _inject_footer(self, slide_xml: bytes, footer_text: str) -> bytes:
        """Inyecta texto en el placeholder de pie de página."""
        root = etree.fromstring(slide_xml)

        for sp in root.iter(f"{{{NS_PRES}}}sp"):
            nvSpPr = sp.find(f"{{{NS_PRES}}}nvSpPr")
            if nvSpPr is None:
                continue
            nvPr = nvSpPr.find(f"{{{NS_PRES}}}nvPr")
            if nvPr is None:
                continue
            ph = nvPr.find(f"{{{NS_PRES}}}ph")
            if ph is None:
                continue
            if ph.get("type", "") == "ftr":
                txBody = sp.find(f"{{{NS_PRES}}}txBody")
                if txBody is None:
                    continue
                for p in txBody.findall(f"{{{NS_DRAW}}}p"):
                    # Limpiar runs existentes
                    for child in list(p):
                        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                        if tag_local in ("r", "br", "fld"):
                            p.remove(child)
                    # Agregar run con footer — Helvetica Neue 18pt bold
                    from .models import AGORA_STYLE
                    r = etree.SubElement(p, f"{{{NS_DRAW}}}r")
                    rPr = etree.SubElement(r, f"{{{NS_DRAW}}}rPr")
                    rPr.set("lang", "es-AR")
                    rPr.set("sz", str(AGORA_STYLE["footer_size"]))
                    rPr.set("b", "1" if AGORA_STYLE["footer_bold"] else "0")
                    latin = etree.SubElement(rPr, f"{{{NS_DRAW}}}latin")
                    latin.set("typeface", AGORA_STYLE["footer_font"])
                    t = etree.SubElement(r, f"{{{NS_DRAW}}}t")
                    t.text = footer_text
                    break
                break

        return _serialize_xml(root)

    def _add_separator_lines(self, slide_xml: bytes, n_charts: int) -> bytes:
        """
        Agrega líneas negras verticales entre charts en slides multi-gráficos.
        """
        root = etree.fromstring(slide_xml)
        spTree = root.find(f".//{{{NS_PRES}}}cSld/{{{NS_PRES}}}spTree")
        if spTree is None:
            return slide_xml

        # Buscar las posiciones reales de los graphicFrames
        gf_positions = []
        for gf in spTree.findall(f"{{{NS_PRES}}}graphicFrame"):
            xfrm = gf.find(f".//{{{NS_PRES}}}xfrm")
            if xfrm is None:
                xfrm = gf.find(f".//{{{NS_DRAW}}}xfrm")
            if xfrm is not None:
                off = xfrm.find(f"{{{NS_DRAW}}}off")
                ext = xfrm.find(f"{{{NS_DRAW}}}ext")
                if off is not None and ext is not None:
                    x = int(off.get("x", "0"))
                    cx = int(ext.get("cx", "0"))
                    y = int(off.get("y", "0"))
                    cy = int(ext.get("cy", "0"))
                    gf_positions.append((x, y, cx, cy))

        gf_positions.sort(key=lambda p: p[0])  # ordenar por x

        for i in range(1, len(gf_positions)):
            prev_x, prev_y, prev_cx, prev_cy = gf_positions[i - 1]
            curr_x, curr_y, curr_cx, curr_cy = gf_positions[i]

            # Línea en el punto medio entre el fin del chart anterior y el inicio del siguiente
            x_line = (prev_x + prev_cx + curr_x) // 2
            y_top = min(prev_y, curr_y)
            y_height = max(prev_cy, curr_cy)

            # Crear shape de línea vertical
            cxnSp = etree.SubElement(spTree, f"{{{NS_PRES}}}cxnSp")

            # nvCxnSpPr
            nvCxnSpPr = etree.SubElement(cxnSp, f"{{{NS_PRES}}}nvCxnSpPr")
            cNvPr = etree.SubElement(nvCxnSpPr, f"{{{NS_PRES}}}cNvPr")
            cNvPr.set("id", str(90000 + i))
            cNvPr.set("name", f"Separador {i}")
            etree.SubElement(nvCxnSpPr, f"{{{NS_PRES}}}cNvCxnSpPr")
            etree.SubElement(nvCxnSpPr, f"{{{NS_PRES}}}nvPr")

            # spPr
            spPr = etree.SubElement(cxnSp, f"{{{NS_PRES}}}spPr")
            xfrm = etree.SubElement(spPr, f"{{{NS_DRAW}}}xfrm")
            off = etree.SubElement(xfrm, f"{{{NS_DRAW}}}off")
            off.set("x", str(x_line))
            off.set("y", str(y_top))
            ext = etree.SubElement(xfrm, f"{{{NS_DRAW}}}ext")
            ext.set("cx", "0")
            ext.set("cy", str(y_height))

            prstGeom = etree.SubElement(spPr, f"{{{NS_DRAW}}}prstGeom")
            prstGeom.set("prst", "line")
            etree.SubElement(prstGeom, f"{{{NS_DRAW}}}avLst")

            # Línea negra 1pt
            ln = etree.SubElement(spPr, f"{{{NS_DRAW}}}ln")
            ln.set("w", "12700")  # 1pt
            fill = etree.SubElement(ln, f"{{{NS_DRAW}}}solidFill")
            srgb = etree.SubElement(fill, f"{{{NS_DRAW}}}srgbClr")
            srgb.set("val", "000000")

        return _serialize_xml(root)

    def _copy_chart_style_files(self, src_chart: str, new_chart: str):
        """Copia style y colors asociados al chart."""
        src_num = re.search(r"chart(\d+)\.xml", src_chart)
        new_num = re.search(r"chart(\d+)\.xml", new_chart)
        if not src_num or not new_num:
            return

        src_n = src_num.group(1)
        new_n = new_num.group(1)

        for prefix in ["style", "colors"]:
            src_path = f"ppt/charts/{prefix}{src_n}.xml"
            new_path = f"ppt/charts/{prefix}{new_n}.xml"
            try:
                self.files[new_path] = self.template_zip.read(src_path)
            except (KeyError, zipfile.BadZipFile):
                pass

    def _update_presentation_xml(self):
        """Actualiza presentation.xml con las nuevas slides."""
        pres_xml = etree.fromstring(self.files["ppt/presentation.xml"])

        # Encontrar sldIdLst
        sldIdLst = pres_xml.find(f"{{{NS_PRES}}}sldIdLst")
        if sldIdLst is None:
            sldIdLst = etree.SubElement(pres_xml, f"{{{NS_PRES}}}sldIdLst")

        # Limpiar slides existentes
        for sldId in sldIdLst.findall(f"{{{NS_PRES}}}sldId"):
            sldIdLst.remove(sldId)

        # Agregar nuevas slides
        for info in self.slide_ids:
            sldId = etree.SubElement(sldIdLst, f"{{{NS_PRES}}}sldId")
            sldId.set("id", str(info["slide_id"]))
            sldId.set(f"{{{NS_REL}}}id", info["rId"])

        self.files["ppt/presentation.xml"] = _serialize_xml(pres_xml)

        # Actualizar presentation.xml.rels
        pres_rels = etree.fromstring(self.files["ppt/_rels/presentation.xml.rels"])

        # Remover relaciones a slides existentes
        for rel in pres_rels.findall(f"{{{NS_RELS}}}Relationship"):
            if NS_SLIDE_REL in rel.get("Type", ""):
                pres_rels.remove(rel)

        # Agregar relaciones a nuevas slides
        for info in self.slide_ids:
            rel = etree.SubElement(pres_rels, f"{{{NS_RELS}}}Relationship")
            rel.set("Id", info["rId"])
            rel.set("Type", NS_SLIDE_REL)
            rel.set("Target", info["slide_path"].replace("ppt/", ""))

        self.files["ppt/_rels/presentation.xml.rels"] = _serialize_xml(pres_rels)

    def _update_content_types(self):
        """Actualiza [Content_Types].xml — solo registra archivos que realmente existen."""
        ct = etree.fromstring(self.files["[Content_Types].xml"])

        # Remover TODOS los overrides de slides, charts, styles y colors
        to_remove = []
        for override in ct.findall(f"{{{NS_CT}}}Override"):
            part = override.get("PartName", "")
            if re.match(r"/ppt/slides/slide\d+\.xml", part):
                to_remove.append(override)
            elif re.match(r"/ppt/charts/(chart|style|colors)\d+\.xml", part):
                to_remove.append(override)
        for el in to_remove:
            ct.remove(el)

        # Agregar overrides SOLO para archivos que existen en self.files
        for info in self.slide_ids:
            override = etree.SubElement(ct, f"{{{NS_CT}}}Override")
            override.set("PartName", f"/{info['slide_path']}")
            override.set("ContentType", CT_SLIDE)

        for path in sorted(self.files.keys()):
            if re.match(r"ppt/charts/chart\d+\.xml$", path):
                override = etree.SubElement(ct, f"{{{NS_CT}}}Override")
                override.set("PartName", f"/{path}")
                override.set("ContentType", CT_CHART)
            elif re.match(r"ppt/charts/style\d+\.xml$", path):
                override = etree.SubElement(ct, f"{{{NS_CT}}}Override")
                override.set("PartName", f"/{path}")
                override.set("ContentType", "application/vnd.ms-office.chartstyle+xml")
            elif re.match(r"ppt/charts/colors\d+\.xml$", path):
                override = etree.SubElement(ct, f"{{{NS_CT}}}Override")
                override.set("PartName", f"/{path}")
                override.set("ContentType", "application/vnd.ms-office.chartcolorstyle+xml")

        self.files["[Content_Types].xml"] = _serialize_xml(ct)

    def _extract_raw(self, info: zipfile.ZipInfo) -> Optional[bytes]:
        """Extrae un archivo del ZIP raw, bypassing CRC check."""
        import struct
        self._template_raw.seek(info.header_offset)
        header = self._template_raw.read(30)
        fname_len = struct.unpack('<H', header[26:28])[0]
        extra_len = struct.unpack('<H', header[28:30])[0]
        self._template_raw.read(fname_len + extra_len)
        raw_data = self._template_raw.read(info.compress_size)
        if info.compress_type == 8:  # deflated
            return zlib.decompress(raw_data, -15)
        return raw_data  # stored

    def _package(self) -> bytes:
        """Empaqueta todos los archivos como PPTX (ZIP)."""
        output = BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for path, content in sorted(self.files.items()):
                zf.writestr(path, content)
        return output.getvalue()

    def _make_empty_rels(self) -> bytes:
        """Crea un archivo .rels vacío."""
        root = etree.Element(f"{{{NS_RELS}}}Relationships")
        return _serialize_xml(root)


# ──────────────────────────────────────────────────────────────────────────────
# PUBLIC API
# ──────────────────────────────────────────────────────────────────────────────

def build_pptx(template_bytes: bytes, questions: List[Question],
               segment_groups=None, pie_pagina: str = "") -> bytes:
    """
    API pública: genera un PPTX con gráficos a partir del template y las preguntas.
    """
    builder = PptxBuilder(template_bytes)
    return builder.build(questions, segment_groups=segment_groups,
                         pie_pagina=pie_pagina)
