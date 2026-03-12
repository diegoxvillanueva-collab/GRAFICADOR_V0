"""
S4 Graficador — Pantalla 1: Upload
Carga de Excel + config básica.  El template de Ágora está embebido en el proyecto.
"""

import streamlit as st
from pathlib import Path
from core.excel_parser import parse_excel


# Template embebido — siempre el mismo de Ágora
_TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "templates" / "Template.pptx"


def _load_template() -> bytes:
    """Carga el template de Ágora desde el proyecto."""
    return _TEMPLATE_PATH.read_bytes()


def render_step1():
    st.title("📊 S4 Graficador")
    st.caption("Ágora Consultores — Generación automática de informes")

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Archivo de datos")
        excel_file = st.file_uploader(
            "Excel de datos (.xlsx)",
            type=["xlsx"],
            help="Excel con pestañas: datos, meta_preguntas, meta_respuestas",
        )

    with col2:
        st.subheader("Configuración")
        cliente = st.text_input(
            "Cliente / Proyecto",
            value=st.session_state.get("cliente", ""),
            placeholder="Ej: Municipio de Pilar",
        )
        fecha = st.text_input(
            "Fecha del informe",
            value=st.session_state.get("fecha", ""),
            placeholder="Ej: Febrero 2026",
        )
        pie_pagina = st.text_input(
            "Pie de página",
            value=st.session_state.get("pie_pagina", ""),
            placeholder="Ej: Encuesta de opinión pública - Pilar | Febrero 2026",
        )

    st.divider()

    # Botón Leer Excel — solo necesita el Excel
    can_proceed = excel_file is not None
    if st.button("Leer Excel →", disabled=not can_proceed, type="primary", use_container_width=True):
        with st.spinner("Parseando Excel..."):
            try:
                excel_bytes = excel_file.read()
                template_bytes = _load_template()

                result = parse_excel(excel_bytes)

                # Guardar en session_state
                st.session_state.parse_result = result
                st.session_state.excel_bytes = excel_bytes
                st.session_state.template_bytes = template_bytes
                st.session_state.cliente = cliente
                st.session_state.fecha = fecha
                st.session_state.pie_pagina = pie_pagina

                # Mostrar resumen
                n_questions = len(result.questions)
                n_pending = result.color_auto_assigned
                n_warnings = len(result.warnings)

                st.success(f"Excel parseado: {n_questions} preguntas, {len(result.segment_labels)} segmentos")

                if n_pending > 0:
                    st.warning(f"{n_pending} respuestas sin color asignado. Se asignaron colores automáticos que puede editar en la siguiente pantalla.")

                if n_warnings > 0:
                    with st.expander(f"⚠️ {n_warnings} warnings"):
                        for w in result.warnings:
                            st.write(f"- {w}")

                st.session_state.step = 2
                st.rerun()

            except Exception as e:
                st.error(f"Error al parsear el Excel: {e}")
                import traceback
                st.code(traceback.format_exc())
