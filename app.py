"""
S4 Graficador — Entry point Streamlit
Sistema de graficación automática de informes de opinión pública.
Ágora Consultores.
"""

import streamlit as st
import os
from dotenv import load_dotenv

load_dotenv()

# ──────────────────────────────────────────────────────────────────────────────
# AUTH
# ──────────────────────────────────────────────────────────────────────────────

APP_PASSWORD = os.getenv("APP_PASSWORD", "agora2026")


def check_auth():
    """Auth simple con password."""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.title("S4 Graficador")
    st.caption("Ágora Consultores — Sistema de graficación automática")

    password = st.text_input("Contraseña", type="password")
    if st.button("Ingresar"):
        if password == APP_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta")

    return False


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="S4 Graficador — Ágora",
        page_icon="📊",
        layout="wide",
    )

    if not check_auth():
        return

    # Navigation
    if "step" not in st.session_state:
        st.session_state.step = 1

    step = st.session_state.step

    if step == 1:
        from ui.step1_upload import render_step1
        render_step1()
    elif step == 2:
        from ui.step2_review import render_step2
        render_step2()
    elif step == 3:
        from ui.step3_download import render_step3
        render_step3()


if __name__ == "__main__":
    main()
