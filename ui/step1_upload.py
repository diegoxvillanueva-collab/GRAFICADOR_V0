"""
S4 Graficador — Pantalla 1: Upload
Carga de Excel + Template + config básica.
"""

import streamlit as st
from core.excel_parser import parse_excel


def render_step1():
    st.title("📊 S4 Graficador")
    st.caption("Ágora Consultores — Generación automática de informes")

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Archivos")
        excel_file = st.file_uploader(
            "Excel de datos (.xlsx)",
            type=["xlsx"],
            help="Excel con pestañas: datos, meta_preguntas, meta_respuestas",
        )
        template_file = st.file_uploader(
            "Template PPTX (.pptx)",
            type=["pptx"],
            help="Template de Ágora con slides de referencia",
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

    # Botón Leer Excel
    can_proceed = excel_file is not None and template_file is not None
    if st.button("Leer Excel →", disabled=not can_proceed, type="primary", use_container_width=True):
        with st.spinner("Parseando Excel..."):
            try:
                excel_bytes = excel_file.read()
                template_bytes = template_file.read()

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
