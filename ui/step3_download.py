"""
S4 Graficador — Pantalla 3: Generación y descarga
"""

import streamlit as st
import time
from core.slide_builder import build_pptx
from core.models import ParseResult


def render_step3():
    st.title("📦 Generación del PPTX")

    result: ParseResult = st.session_state.get("parse_result")
    template_bytes = st.session_state.get("template_bytes")

    if result is None or template_bytes is None:
        st.warning("No hay datos para generar. Volvé al paso 1.")
        if st.button("← Volver al inicio"):
            st.session_state.step = 1
            st.rerun()
        return

    questions = result.questions
    cliente = st.session_state.get("cliente", "informe")
    fecha = st.session_state.get("fecha", "")

    # Generar si no está generado
    if "pptx_bytes" not in st.session_state:
        progress = st.progress(0, text="Preparando...")

        try:
            progress.progress(10, text="Calculando orden de slides...")
            time.sleep(0.3)

            progress.progress(30, text="Copiando slides del template...")
            time.sleep(0.3)

            progress.progress(50, text="Inyectando gráficos...")

            pie_pagina = st.session_state.get("pie_pagina", "")
            pptx_bytes = build_pptx(template_bytes, questions,
                                     segment_groups=result.segment_groups,
                                     pie_pagina=pie_pagina)

            progress.progress(90, text="Empaquetando PPTX...")
            time.sleep(0.2)

            st.session_state.pptx_bytes = pptx_bytes
            progress.progress(100, text="¡Listo!")
            time.sleep(0.3)

        except Exception as e:
            progress.empty()
            st.error(f"Error durante la generación: {e}")
            import traceback
            st.code(traceback.format_exc())

            if st.button("← Volver a revisión"):
                st.session_state.step = 2
                st.rerun()
            return

    # Mostrar resultado
    pptx_bytes = st.session_state.pptx_bytes

    st.success(f"¡PPTX generado exitosamente! ({len(pptx_bytes) / 1024:.0f} KB)")

    # Stats
    from core.slide_builder import _compute_slide_order
    slide_order = _compute_slide_order(questions)
    n_frec = sum(1 for s in slide_order if s["type"] == "FREC")
    n_app = sum(1 for s in slide_order if s["type"] == "APERTURA")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total slides", len(slide_order))
    with col2:
        st.metric("Frecuencias", n_frec)
    with col3:
        st.metric("Aperturas", n_app)

    st.divider()

    # Nombre del archivo
    filename_parts = []
    if cliente:
        filename_parts.append(cliente.replace(" ", "_"))
    if fecha:
        filename_parts.append(fecha.replace(" ", "_"))
    filename = "_".join(filename_parts) if filename_parts else "informe"
    filename = f"{filename}.pptx"

    st.download_button(
        label=f"⬇️ Descargar {filename}",
        data=pptx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True,
    )

    st.info("📋 Próximo paso: pasar el PPTX por S3 Corrector para revisión automática.")

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Volver a revisión", use_container_width=True):
            if "pptx_bytes" in st.session_state:
                del st.session_state.pptx_bytes
            st.session_state.step = 2
            st.rerun()
    with col2:
        if st.button("🔄 Nuevo informe", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key != "authenticated":
                    del st.session_state[key]
            st.session_state.step = 1
            st.rerun()
