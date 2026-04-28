# src/ui/sidebar.py

"""Sidebar con logo STARS REI, indicadores de progreso y botón de reset."""

import streamlit as st


def render_sidebar():
    """
    Renderiza el sidebar con logo STARS REI, indicadores de progreso y botón de reset.
    Lee y modifica st.session_state directamente.
    """
    st.sidebar.markdown(
        '<div class="sidebar-logo">'
        '  <div class="logo-row">'
        '    <div class="ripple-icon">'
        '      <div class="arc"></div>'
        '      <div class="arc"></div>'
        '      <div class="arc"></div>'
        '      <div class="arc"></div>'
        '    </div>'
        '    <div class="logo-text">STARS REI</div>'
        '  </div>'
        '  <div class="logo-sub">Financials Reporter</div>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.sidebar.markdown("---")

    steps = {
        1: "1. Configuración",
        2: "2. Revisión",
        3: "3. Output",
    }

    # Mostrar progreso
    for num, label in steps.items():
        if num < st.session_state.step:
            st.sidebar.success(label)
        elif num == st.session_state.step:
            st.sidebar.info(f"➡️ {label}")
        else:
            st.sidebar.text(f"   {label}")

    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Reiniciar"):
        st.session_state.step = 1
        st.session_state.results = None
        st.session_state.approved = False
        st.session_state.classification_approved = False
        st.session_state.l1_overrides = {}
        st.rerun()
