"""
S4 Graficador — Chart Injector
Cirugía XML sobre charts dentro del PPTX (ZIP).
Inyecta datos, colores y formato Ágora en los charts.
"""

from lxml import etree
from copy import deepcopy
from typing import List, Optional
from .models import Question, Answer, AGORA_STYLE, get_chart_style

# ──────────────────────────────────────────────────────────────────────────────
# NAMESPACES
# ──────────────────────────────────────────────────────────────────────────────

NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

C = NS["c"]
A = NS["a"]


def _tag(ns_prefix: str, local: str) -> str:
    return f"{{{NS[ns_prefix]}}}{local}"


def _fix_xml_decl(raw: bytes) -> bytes:
    """Arregla comillas simples a dobles en declaración XML (requisito PowerPoint)."""
    if raw.startswith(b"<?xml"):
        end = raw.find(b"?>")
        if end != -1:
            decl = raw[:end + 2].replace(b"'", b'"')
            raw = decl + raw[end + 2:]
    return raw


# ──────────────────────────────────────────────────────────────────────────────
# BUILDING BLOCKS — Formato Ágora
# ──────────────────────────────────────────────────────────────────────────────

def _build_spPr(color_hex: str) -> etree._Element:
    """<c:spPr> con fill sólido y sin borde."""
    clean = color_hex.strip("#").upper()
    spPr = etree.Element(_tag("c", "spPr"))
    fill = etree.SubElement(spPr, _tag("a", "solidFill"))
    srgb = etree.SubElement(fill, _tag("a", "srgbClr"))
    srgb.set("val", clean)
    ln = etree.SubElement(spPr, _tag("a", "ln"))
    etree.SubElement(ln, _tag("a", "noFill"))
    return spPr


def _build_dPt(idx: int, color_hex: str) -> etree._Element:
    """<c:dPt> para colorear un punto individual."""
    clean = color_hex.strip("#").upper()
    dPt = etree.Element(_tag("c", "dPt"))
    etree.SubElement(dPt, _tag("c", "idx")).set("val", str(idx))
    etree.SubElement(dPt, _tag("c", "invertIfNegative")).set("val", "0")
    etree.SubElement(dPt, _tag("c", "bubble3D")).set("val", "0")
    spPr = etree.SubElement(dPt, _tag("c", "spPr"))
    fill = etree.SubElement(spPr, _tag("a", "solidFill"))
    srgb = etree.SubElement(fill, _tag("a", "srgbClr"))
    srgb.set("val", clean)
    ln = etree.SubElement(spPr, _tag("a", "ln"))
    etree.SubElement(ln, _tag("a", "noFill"))
    return dPt


def _build_data_labels(chart_type: str) -> etree._Element:
    """<c:dLbls> con formato según tipo de chart."""
    style = get_chart_style(chart_type)
    font = style.get("data_label_font", "Roboto Condensed")
    size = style.get("data_label_size", 1400)
    bold = style.get("data_label_bold", True)
    color = style.get("data_label_color", "bg1")

    dLbls = etree.Element(_tag("c", "dLbls"))

    # Formato visual
    spPr = etree.SubElement(dLbls, _tag("c", "spPr"))
    etree.SubElement(spPr, _tag("a", "noFill"))
    ln = etree.SubElement(spPr, _tag("a", "ln"))
    etree.SubElement(ln, _tag("a", "noFill"))

    # Fuente
    txPr = etree.SubElement(dLbls, _tag("c", "txPr"))
    bodyPr = etree.SubElement(txPr, _tag("a", "bodyPr"))
    bodyPr.set("rot", "0")
    bodyPr.set("vert", "horz")
    bodyPr.set("wrap", "square")
    bodyPr.set("anchor", "ctr")
    bodyPr.set("anchorCtr", "1")
    etree.SubElement(txPr, _tag("a", "lstStyle"))
    p = etree.SubElement(txPr, _tag("a", "p"))
    pPr = etree.SubElement(p, _tag("a", "pPr"))
    defRPr = etree.SubElement(pPr, _tag("a", "defRPr"))
    defRPr.set("sz", str(size))
    defRPr.set("b", "1" if bold else "0")
    fill = etree.SubElement(defRPr, _tag("a", "solidFill"))
    etree.SubElement(fill, _tag("a", "schemeClr")).set("val", color)
    etree.SubElement(defRPr, _tag("a", "latin")).set("typeface", font)

    # Mostrar valores, no categorías
    etree.SubElement(dLbls, _tag("c", "dLblPos")).set("val", "ctr")
    etree.SubElement(dLbls, _tag("c", "showLegendKey")).set("val", "0")
    etree.SubElement(dLbls, _tag("c", "showVal")).set("val", "1")
    etree.SubElement(dLbls, _tag("c", "showCatName")).set("val", "0")
    etree.SubElement(dLbls, _tag("c", "showSerName")).set("val", "0")
    etree.SubElement(dLbls, _tag("c", "showPercent")).set("val", "0")
    etree.SubElement(dLbls, _tag("c", "showBubbleSize")).set("val", "0")

    return dLbls


def _build_str_cache(labels: List[str]) -> etree._Element:
    cache = etree.Element(_tag("c", "strCache"))
    etree.SubElement(cache, _tag("c", "ptCount")).set("val", str(len(labels)))
    for i, label in enumerate(labels):
        pt = etree.SubElement(cache, _tag("c", "pt"))
        pt.set("idx", str(i))
        v = etree.SubElement(pt, _tag("c", "v"))
        v.text = str(label)
    return cache


def _build_num_cache(values: List[float], fmt: str = "0%") -> etree._Element:
    cache = etree.Element(_tag("c", "numCache"))
    etree.SubElement(cache, _tag("c", "formatCode")).text = fmt
    etree.SubElement(cache, _tag("c", "ptCount")).set("val", str(len(values)))
    for i, val in enumerate(values):
        pt = etree.SubElement(cache, _tag("c", "pt"))
        pt.set("idx", str(i))
        v = etree.SubElement(pt, _tag("c", "v"))
        v.text = str(val) if val is not None else "0"
    return cache


def _build_cat_element(labels: List[str]) -> etree._Element:
    cat = etree.Element(_tag("c", "cat"))
    strRef = etree.SubElement(cat, _tag("c", "strRef"))
    etree.SubElement(strRef, _tag("c", "f")).text = "Sheet1!$A$1"
    strRef.append(_build_str_cache(labels))
    return cat


def _build_val_element(values: List[float], fmt: str = "0%") -> etree._Element:
    val_el = etree.Element(_tag("c", "val"))
    numRef = etree.SubElement(val_el, _tag("c", "numRef"))
    etree.SubElement(numRef, _tag("c", "f")).text = "Sheet1!$A$1"
    numRef.append(_build_num_cache(values, fmt))
    return val_el


def _build_tx_element(label: str) -> etree._Element:
    tx = etree.Element(_tag("c", "tx"))
    v = etree.SubElement(tx, _tag("c", "v"))
    v.text = label
    return tx


def _build_axis_txPr(font: str, size: int, bold: bool = True, rotation: int = 0) -> etree._Element:
    """Construye un <c:txPr> para ejes con formato Ágora."""
    txPr = etree.Element(_tag("c", "txPr"))
    bodyPr = etree.SubElement(txPr, _tag("a", "bodyPr"))
    if rotation:
        bodyPr.set("rot", str(rotation))
    bodyPr.set("vert", "horz")
    bodyPr.set("wrap", "square")
    etree.SubElement(txPr, _tag("a", "lstStyle"))
    p = etree.SubElement(txPr, _tag("a", "p"))
    pPr = etree.SubElement(p, _tag("a", "pPr"))
    defRPr = etree.SubElement(pPr, _tag("a", "defRPr"))
    defRPr.set("sz", str(size))
    defRPr.set("b", "1" if bold else "0")
    fill = etree.SubElement(defRPr, _tag("a", "solidFill"))
    etree.SubElement(fill, _tag("a", "schemeClr")).set("val", "tx1")
    etree.SubElement(defRPr, _tag("a", "latin")).set("typeface", font)
    etree.SubElement(p, _tag("a", "endParaRPr")).set("lang", "es-AR")
    return txPr


def _build_legend_txPr(font: str, size: int, bold: bool = False) -> etree._Element:
    """Construye un <c:txPr> para leyendas."""
    txPr = etree.Element(_tag("c", "txPr"))
    bodyPr = etree.SubElement(txPr, _tag("a", "bodyPr"))
    bodyPr.set("rot", "0")
    bodyPr.set("vert", "horz")
    etree.SubElement(txPr, _tag("a", "lstStyle"))
    p = etree.SubElement(txPr, _tag("a", "p"))
    pPr = etree.SubElement(p, _tag("a", "pPr"))
    defRPr = etree.SubElement(pPr, _tag("a", "defRPr"))
    defRPr.set("sz", str(size))
    defRPr.set("b", "1" if bold else "0")
    fill = etree.SubElement(defRPr, _tag("a", "solidFill"))
    etree.SubElement(fill, _tag("a", "schemeClr")).set("val", "tx1")
    etree.SubElement(defRPr, _tag("a", "latin")).set("typeface", font)
    etree.SubElement(p, _tag("a", "endParaRPr")).set("lang", "es-AR")
    return txPr


# ──────────────────────────────────────────────────────────────────────────────
# SERIES BUILDERS
# ──────────────────────────────────────────────────────────────────────────────

def _build_frec_simple_series(question: Question, chart_type: str = "FREC_SIMPLE") -> List[etree._Element]:
    """
    FREC_SIMPLE / FREC_MULTI_GRAFICOS: 1 serie, N categorías (respuestas), cada punto con su color.
    """
    ser = etree.Element(_tag("c", "ser"))
    etree.SubElement(ser, _tag("c", "idx")).set("val", "0")
    etree.SubElement(ser, _tag("c", "order")).set("val", "0")
    ser.append(_build_spPr(question.respuestas[0].color))
    etree.SubElement(ser, _tag("c", "invertIfNegative")).set("val", "0")

    for i, ans in enumerate(question.respuestas):
        ser.append(_build_dPt(i, ans.color))

    # Data labels con estilo del chart_type
    ser.append(_build_data_labels(chart_type))

    ser.append(_build_cat_element([a.label for a in question.respuestas]))
    ser.append(_build_val_element([a.total for a in question.respuestas]))

    return [ser]


def _build_frec_multiple_series(questions: List[Question]) -> List[etree._Element]:
    """
    FREC_MULTIPLE: N series (una por respuesta), M categorías (una por pregunta).
    """
    if not questions:
        return []

    ref_respuestas = questions[0].respuestas
    category_labels = []
    for q in questions:
        label = q.titulo_app if q.titulo_app else q.id_pregunta
        if len(label) > 40:
            label = label[:37] + "..."
        category_labels.append(label)

    series_list = []
    for i, ans_ref in enumerate(ref_respuestas):
        ser = etree.Element(_tag("c", "ser"))
        etree.SubElement(ser, _tag("c", "idx")).set("val", str(i))
        etree.SubElement(ser, _tag("c", "order")).set("val", str(i))
        ser.append(_build_tx_element(ans_ref.label))
        ser.append(_build_spPr(ans_ref.color))
        etree.SubElement(ser, _tag("c", "invertIfNegative")).set("val", "0")
        ser.append(_build_data_labels("FREC_MULTIPLE"))
        ser.append(_build_cat_element(category_labels))

        values = []
        for q in questions:
            if i < len(q.respuestas):
                values.append(q.respuestas[i].total)
            else:
                values.append(0.0)
        ser.append(_build_val_element(values))
        series_list.append(ser)

    return series_list


def _build_apertura_series(question: Question, segment_groups=None) -> List[etree._Element]:
    """
    APERTURA_SIMPLE: N series (una por respuesta), M categorías (segmentos).
    Barras verticales con separadores entre grupos de segmentos.
    """
    if segment_groups:
        cat_labels = []
        val_indices = []

        for gi, group in enumerate(segment_groups):
            if gi > 0:
                cat_labels.append("")
                val_indices.append(None)
            for label in group.labels:
                if label in question.segment_labels:
                    idx = question.segment_labels.index(label)
                    val_indices.append(idx)
                    cat_labels.append(label)
    else:
        cat_labels = question.segment_labels
        val_indices = list(range(len(question.segment_labels)))

    series_list = []
    for i, ans in enumerate(question.respuestas):
        ser = etree.Element(_tag("c", "ser"))
        etree.SubElement(ser, _tag("c", "idx")).set("val", str(i))
        etree.SubElement(ser, _tag("c", "order")).set("val", str(i))
        ser.append(_build_tx_element(ans.label))
        ser.append(_build_spPr(ans.color))
        etree.SubElement(ser, _tag("c", "invertIfNegative")).set("val", "0")
        ser.append(_build_data_labels("APERTURA_SIMPLE"))
        ser.append(_build_cat_element(cat_labels))

        values = []
        for vi in val_indices:
            if vi is None:
                values.append(0.0)
            else:
                values.append(ans.segment_values[vi] if vi < len(ans.segment_values) else 0.0)
        ser.append(_build_val_element(values))
        series_list.append(ser)

    return series_list


# ──────────────────────────────────────────────────────────────────────────────
# AXIS & LEGEND FORMATTERS
# ──────────────────────────────────────────────────────────────────────────────

def _format_cat_axis(plot_area, chart_type: str) -> None:
    """Formatea el eje de categorías según el tipo de chart."""
    style = get_chart_style(chart_type)
    cat_ax = plot_area.find("c:catAx", NS)
    if cat_ax is None:
        return

    font = style.get("cat_axis_font", "HelveticaNeue Std (Cuerpo)")
    size = style.get("cat_axis_size", 2200)
    bold = style.get("cat_axis_bold", False)

    # Remover txPr existente
    old_txPr = cat_ax.find("c:txPr", NS)
    if old_txPr is not None:
        cat_ax.remove(old_txPr)

    if chart_type == "APERTURA_SIMPLE":
        # Aperturas: texto vertical rotado -90°
        txPr = etree.SubElement(cat_ax, _tag("c", "txPr"))
        bodyPr = etree.SubElement(txPr, _tag("a", "bodyPr"))
        bodyPr.set("rot", "-5400000")
        bodyPr.set("vert", "horz")
        bodyPr.set("wrap", "square")
        bodyPr.set("anchor", "ctr")
        bodyPr.set("anchorCtr", "1")
        etree.SubElement(txPr, _tag("a", "lstStyle"))
        p = etree.SubElement(txPr, _tag("a", "p"))
        pPr = etree.SubElement(p, _tag("a", "pPr"))
        defRPr = etree.SubElement(pPr, _tag("a", "defRPr"))
        defRPr.set("sz", str(size))
        defRPr.set("b", "1" if bold else "0")
        fill_el = etree.SubElement(defRPr, _tag("a", "solidFill"))
        etree.SubElement(fill_el, _tag("a", "schemeClr")).set("val", "tx1")
        etree.SubElement(defRPr, _tag("a", "latin")).set("typeface", font)
        etree.SubElement(p, _tag("a", "endParaRPr")).set("lang", "es-AR")

        # tickLblPos
        tl = cat_ax.find("c:tickLblPos", NS)
        if tl is not None:
            tl.set("val", "low")
        else:
            etree.SubElement(cat_ax, _tag("c", "tickLblPos")).set("val", "low")

        # Layout de plotArea para dejar espacio inferior para labels
        layout = plot_area.find("c:layout", NS)
        if layout is None:
            layout = etree.Element(_tag("c", "layout"))
            plot_area.insert(0, layout)
        for child in list(layout):
            layout.remove(child)
        manual = etree.SubElement(layout, _tag("c", "manualLayout"))
        etree.SubElement(manual, _tag("c", "layoutTarget")).set("val", "inner")
        etree.SubElement(manual, _tag("c", "xMode")).set("val", "edge")
        etree.SubElement(manual, _tag("c", "yMode")).set("val", "edge")
        etree.SubElement(manual, _tag("c", "x")).set("val", "0.02")
        etree.SubElement(manual, _tag("c", "y")).set("val", "0.08")
        etree.SubElement(manual, _tag("c", "w")).set("val", "0.96")
        etree.SubElement(manual, _tag("c", "h")).set("val", "0.70")
    else:
        # Frecuencias: eje horizontal normal
        cat_ax.append(_build_axis_txPr(font, size, bold))


def _format_val_axis(plot_area, chart_type: str) -> None:
    """Formatea el eje de valores según el tipo de chart."""
    style = get_chart_style(chart_type)
    val_ax = plot_area.find("c:valAx", NS)
    if val_ax is None:
        return

    font = style.get("val_axis_font", "HelveticaNeue Std")
    size = style.get("val_axis_size", 2200)
    bold = style.get("val_axis_bold", True)

    # Solo formatear si el chart type define val_axis
    if "val_axis_font" not in style:
        return

    old_txPr = val_ax.find("c:txPr", NS)
    if old_txPr is not None:
        val_ax.remove(old_txPr)

    val_ax.append(_build_axis_txPr(font, size, bold))


def _format_legend(root, chart_type: str) -> None:
    """Formatea la leyenda del chart según el tipo."""
    style = get_chart_style(chart_type)
    if "legend_font" not in style:
        return

    chart = root.find("c:chart", NS)
    if chart is None:
        return

    legend = chart.find("c:legend", NS)
    if legend is None:
        return

    font = style["legend_font"]
    size = style["legend_size"]
    bold = style.get("legend_bold", False)

    # Remover txPr existente de la leyenda
    old_txPr = legend.find("c:txPr", NS)
    if old_txPr is not None:
        legend.remove(old_txPr)

    legend.append(_build_legend_txPr(font, size, bold))


# ──────────────────────────────────────────────────────────────────────────────
# CHART XML INJECTION
# ──────────────────────────────────────────────────────────────────────────────

def inject_chart_data(chart_xml_bytes: bytes, chart_type: str,
                      question: Optional[Question] = None,
                      questions_group: Optional[List[Question]] = None,
                      title: Optional[str] = None,
                      orientacion: str = "V",
                      segment_groups=None) -> bytes:
    """
    Inyecta datos en un chart XML.

    Args:
        chart_type: "FREC_SIMPLE", "FREC_MULTIPLE", "FREC_MULTI_GRAFICOS", "APERTURA_SIMPLE"
        orientacion: "V" (vertical/col) o "H" (horizontal/bar) — para frecuencias
        segment_groups: lista de SegmentGroup para separadores en aperturas
    """
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(chart_xml_bytes, parser)

    plot_area = root.find(".//c:plotArea", NS)
    if plot_area is None:
        return chart_xml_bytes

    bar_chart = plot_area.find("c:barChart", NS)
    if bar_chart is None:
        bar_chart = plot_area.find("c:bar3DChart", NS)
    if bar_chart is None:
        return chart_xml_bytes

    # Limpiar series existentes
    for ser in bar_chart.findall("c:ser", NS):
        bar_chart.remove(ser)

    # Limpiar dLbls existentes del chart (no de series)
    for dlbls in bar_chart.findall("c:dLbls", NS):
        bar_chart.remove(dlbls)

    # Construir nuevas series (con data labels tipificados)
    if chart_type == "FREC_SIMPLE" and question:
        new_series = _build_frec_simple_series(question, "FREC_SIMPLE")
    elif chart_type == "FREC_MULTIPLE" and questions_group:
        new_series = _build_frec_multiple_series(questions_group)
    elif chart_type == "FREC_MULTI_GRAFICOS" and question:
        new_series = _build_frec_simple_series(question, "FREC_MULTI_GRAFICOS")
    elif chart_type == "APERTURA_SIMPLE" and question:
        new_series = _build_apertura_series(question, segment_groups)
    else:
        return chart_xml_bytes

    # Insertar series después de varyColors
    vary_colors = bar_chart.find("c:varyColors", NS)
    if vary_colors is not None:
        insert_idx = list(bar_chart).index(vary_colors) + 1
    else:
        insert_idx = 2
    for i, ser in enumerate(new_series):
        bar_chart.insert(insert_idx + i, ser)

    # Configurar barDir y grouping
    bar_dir = bar_chart.find("c:barDir", NS)
    grouping = bar_chart.find("c:grouping", NS)

    if chart_type == "APERTURA_SIMPLE":
        if bar_dir is not None:
            bar_dir.set("val", "col")
        if grouping is not None:
            grouping.set("val", "percentStacked")
    elif chart_type in ("FREC_SIMPLE", "FREC_MULTI_GRAFICOS"):
        if bar_dir is not None:
            bar_dir.set("val", "bar" if orientacion == "H" else "col")
        if grouping is not None:
            grouping.set("val", "clustered")
    elif chart_type == "FREC_MULTIPLE":
        if bar_dir is not None:
            bar_dir.set("val", "bar" if orientacion == "H" else "col")
        if grouping is not None:
            grouping.set("val", "percentStacked")

    # Gap width
    gap = bar_chart.find("c:gapWidth", NS)
    if gap is not None:
        if chart_type == "APERTURA_SIMPLE":
            gap.set("val", str(AGORA_STYLE["apertura_gap_width"]))
        else:
            gap.set("val", str(AGORA_STYLE["frec_gap_width"]))

    # Overlap para percentStacked
    if chart_type in ("APERTURA_SIMPLE", "FREC_MULTIPLE"):
        overlap = bar_chart.find("c:overlap", NS)
        if overlap is not None:
            overlap.set("val", "100")

    # ── Formateo de ejes según tipo de chart ──
    _format_cat_axis(plot_area, chart_type)
    _format_val_axis(plot_area, chart_type)

    # ── Formateo de leyenda ──
    _format_legend(root, chart_type)

    # ── Inyectar título en el chart ──
    if title:
        _inject_chart_title(root, title, chart_type)

    # Eliminar elementos que referencian archivos del template que no copiamos
    for tag in ("c:externalData", "c:clrMapOvr", "c:printSettings"):
        el = root.find(tag, NS)
        if el is not None:
            root.remove(el)

    raw = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return _fix_xml_decl(raw)


def _inject_chart_title(root, title: str, chart_type: str = "FREC_MULTIPLE"):
    """Inyecta o actualiza el título del chart con formato según tipo."""
    style = get_chart_style(chart_type)
    font = style.get("chart_title_font", "HelveticaNeue Std")
    size = style.get("chart_title_size", 2400)
    bold = style.get("chart_title_bold", True)

    chart = root.find("c:chart", NS)
    if chart is None:
        return

    title_el = chart.find("c:title", NS)
    if title_el is not None:
        # Buscar texto en rich text y actualizar
        rich = title_el.find(".//c:rich", NS)
        if rich is not None:
            for p in rich.findall("a:p", NS):
                # Limpiar runs existentes
                for r in p.findall("a:r", NS):
                    p.remove(r)
                # Crear nuevo run con formato correcto
                new_r = etree.SubElement(p, _tag("a", "r"))
                rPr = etree.SubElement(new_r, _tag("a", "rPr"))
                rPr.set("lang", "es-AR")
                rPr.set("sz", str(size))
                rPr.set("b", "1" if bold else "0")
                fill_el = etree.SubElement(rPr, _tag("a", "solidFill"))
                etree.SubElement(fill_el, _tag("a", "schemeClr")).set("val", "tx1")
                etree.SubElement(rPr, _tag("a", "latin")).set("typeface", font)
                t = etree.SubElement(new_r, _tag("a", "t"))
                t.text = title
                return
        v = title_el.find(".//c:v", NS)
        if v is not None:
            v.text = title
    else:
        # Crear título nuevo
        title_el = etree.SubElement(chart, _tag("c", "title"))
        tx = etree.SubElement(title_el, _tag("c", "tx"))
        rich = etree.SubElement(tx, _tag("c", "rich"))

        bodyPr = etree.SubElement(rich, _tag("a", "bodyPr"))
        bodyPr.set("rot", "0")
        bodyPr.set("vert", "horz")
        bodyPr.set("wrap", "square")
        bodyPr.set("anchor", "ctr")
        bodyPr.set("anchorCtr", "1")
        etree.SubElement(rich, _tag("a", "lstStyle"))

        p = etree.SubElement(rich, _tag("a", "p"))
        pPr = etree.SubElement(p, _tag("a", "pPr"))
        pPr.set("algn", "l")
        defRPr = etree.SubElement(pPr, _tag("a", "defRPr"))
        defRPr.set("sz", str(size))
        defRPr.set("b", "1" if bold else "0")
        fill_el = etree.SubElement(defRPr, _tag("a", "solidFill"))
        etree.SubElement(fill_el, _tag("a", "schemeClr")).set("val", "tx1")
        etree.SubElement(defRPr, _tag("a", "latin")).set("typeface", font)

        r = etree.SubElement(p, _tag("a", "r"))
        rPr = etree.SubElement(r, _tag("a", "rPr"))
        rPr.set("lang", "es-AR")
        rPr.set("sz", str(size))
        rPr.set("b", "1" if bold else "0")
        fill_el2 = etree.SubElement(rPr, _tag("a", "solidFill"))
        etree.SubElement(fill_el2, _tag("a", "schemeClr")).set("val", "tx1")
        etree.SubElement(rPr, _tag("a", "latin")).set("typeface", font)
        t = etree.SubElement(r, _tag("a", "t"))
        t.text = title

        overlay = etree.SubElement(title_el, _tag("c", "overlay"))
        overlay.set("val", "0")

        chart.remove(title_el)
        chart.insert(0, title_el)

        atd = chart.find("c:autoTitleDeleted", NS)
        if atd is not None:
            atd.set("val", "0")
