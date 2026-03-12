"""
S4 Graficador — Pantalla 2: Revisión
Tabla de preguntas, asignación de colores, config antes de generar.
"""

import streamlit as st
import pandas as pd
from core.models import ParseResult


def render_step2():
    st.title("📋 Revisión de preguntas")

    result: ParseResult = st.session_state.get("parse_result")
    if result is None:
        st.warning("No hay datos parseados. Volvé al paso 1.")
        if st.button("← Volver"):
            st.session_state.step = 1
            st.rerun()
        return

    questions = result.questions

    # ── Tabla resumen ──
    st.subheader("Preguntas detectadas")

    # Checkbox de inclusión
    if "include_flags" not in st.session_state:
        st.session_state.include_flags = {q.id_pregunta: True for q in questions}

    table_data = []
    for q in questions:
        table_data.append({
            "id": q.id_pregunta,
            "capítulo": q.capitulo,
            "tipo": q.tipo_slide,
            "grupo": q.grupo_frec,
            "título_frec": q.titulo_frec[:60],
            "título_app": q.titulo_app[:60],
            "respuestas": len(q.respuestas),
            "colores_pendientes": sum(1 for a in q.respuestas if a.color_pending),
        })

    df = pd.DataFrame(table_data)

    # Data editor con checkbox
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "id": st.column_config.TextColumn("ID", width="small"),
            "capítulo": st.column_config.TextColumn("Capítulo", width="medium"),
            "tipo": st.column_config.TextColumn("Tipo", width="medium"),
            "grupo": st.column_config.NumberColumn("Grupo", width="small"),
            "título_frec": st.column_config.TextColumn("Título Frec.", width="large"),
            "título_app": st.column_config.TextColumn("Título App.", width="large"),
            "respuestas": st.column_config.NumberColumn("N resp.", width="small"),
            "colores_pendientes": st.column_config.NumberColumn("🎨 Pend.", width="small"),
        },
    )

    # ── FREC_MULTI_GRAFICOS badge ──
    multi_graf = [q for q in questions if q.tipo_slide == "FREC_MULTI_GRAFICOS"]
    if multi_graf:
        st.info(f"ℹ️ {len(multi_graf)} preguntas con tipo FREC_MULTI_GRAFICOS — se generarán con N charts flexibles en una slide.")

    st.divider()

    # ── Colores pendientes ──
    pending_questions = [q for q in questions if any(a.color_pending for a in q.respuestas)]

    if pending_questions:
        st.subheader(f"🎨 Colores pendientes ({len(pending_questions)} preguntas)")
        st.caption("Estas respuestas no tenían color en meta_respuestas. Se asignaron colores automáticos que podés editar:")

        for q in pending_questions:
            pending_answers = [a for a in q.respuestas if a.color_pending]
            with st.expander(f"⚠️ {q.id_pregunta} — {q.titulo_frec[:50]} ({len(pending_answers)} pendientes)"):
                cols = st.columns(min(len(pending_answers), 4))
                for i, ans in enumerate(pending_answers):
                    col_idx = i % 4
                    with cols[col_idx]:
                        new_color = st.color_picker(
                            f"'{ans.label}'",
                            value=ans.color,
                            key=f"color_{q.id_pregunta}_{ans.label}",
                        )
                        if new_color != ans.color:
                            ans.color = new_color
    else:
        st.success("✅ Todos los colores están asignados.")

    st.divider()

    # ── Conteo de slides ──
    from core.slide_builder import _compute_slide_order
    slide_order = _compute_slide_order(questions)
    n_slides = len(slide_order)

    n_frec = sum(1 for s in slide_order if s["type"] == "FREC")
    n_app = sum(1 for s in slide_order if s["type"] == "APERTURA")

    st.metric("Slides a generar", n_slides, f"{n_frec} frecuencias + {n_app} aperturas")

    # ── Botones ──
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Volver", use_container_width=True):
            st.session_state.step = 1
            st.rerun()
    with col2:
        if st.button(
            f"Generar PPTX con {n_slides} slides →",
            type="primary",
            use_container_width=True,
        ):
            st.session_state.step = 3
            st.rerun()
