"""
RE Financials Reporter — Streamlit App
Punto de entrada principal. Solo configuración de página y routing.

Uso: streamlit run app.py
"""

import streamlit as st

from src.ui.styles import inject_styles
from src.ui.sidebar import render_sidebar
from src.ui.step1_upload import render_step1
from src.ui.step2_review import render_step2
from src.ui.step3_output import render_step3


# ── Configuración de página ──────────────────────────────────
st.set_page_config(
    page_title="STARS REI — Financials Reporter",
    page_icon="📊",
    layout="wide",
)

# ── Estilos corporativos ─────────────────────────────────────
inject_styles()

# ── Inicialización de estado ─────────────────────────────────
_SESSION_DEFAULTS = {
    "step": 1,
    "results": None,
    "approved": False,
    "classification_approved": False,
    "l1_overrides": {},
    "revalidated": False,
    "classification_finalized": False,
}
for key, default in _SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ── Sidebar ──────────────────────────────────────────────────
render_sidebar()

# ── Routing por paso ─────────────────────────────────────────
step = st.session_state.step

if step == 1:
    render_step1()
elif step == 2:
    render_step2()
elif step == 3:
    render_step3()

# ── Footer ───────────────────────────────────────────────────
st.markdown(
    '<div class="stars-footer">'
    '<strong>STARS REI</strong> — Financials Reporter<br>'
    'Real Estate Investment Management'
    '</div>',
    unsafe_allow_html=True,
)
