# src/ui/styles.py

"""CSS corporativo STARS REI."""

import streamlit as st

_STARS_CSS = """
<style>
/* === Paleta STARS REI === */
:root {
    --stars-primary: #1B3A5C;
    --stars-primary-light: #2A5580;
    --stars-accent: #D4A843;
    --stars-bg: #FFFFFF;
    --stars-bg-secondary: #F0F4F8;
    --stars-text: #1A1A2E;
    --stars-text-muted: #6B7B8D;
    --stars-border: #D1DAE3;
    --stars-success: #2E7D4F;
    --stars-warning: #D4A843;
    --stars-error: #C0392B;
}

/* Header bar */
header[data-testid="stHeader"] {
    background: linear-gradient(90deg, var(--stars-primary) 0%, var(--stars-primary-light) 100%);
}

/* Sidebar styling */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, var(--stars-primary) 0%, #142D47 100%);
}
section[data-testid="stSidebar"] * {
    color: #E8EDF2 !important;
}
section[data-testid="stSidebar"] .stMarkdown hr {
    border-color: rgba(255,255,255,0.15);
}
section[data-testid="stSidebar"] button {
    background-color: rgba(212,168,67,0.15) !important;
    border: 1px solid var(--stars-accent) !important;
    color: var(--stars-accent) !important;
}
section[data-testid="stSidebar"] button:hover {
    background-color: rgba(212,168,67,0.3) !important;
}

/* Primary buttons */
button[kind="primary"], .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--stars-primary) 0%, var(--stars-primary-light) 100%) !important;
    border: none !important;
    font-weight: 600 !important;
    letter-spacing: 0.3px;
}
button[kind="primary"]:hover {
    background: linear-gradient(135deg, var(--stars-primary-light) 0%, var(--stars-primary) 100%) !important;
    box-shadow: 0 4px 12px rgba(27,58,92,0.3) !important;
}

/* Download buttons */
.stDownloadButton > button {
    border: 1.5px solid var(--stars-primary) !important;
    color: var(--stars-primary) !important;
    font-weight: 500 !important;
    transition: all 0.2s ease;
}
.stDownloadButton > button:hover {
    background-color: var(--stars-primary) !important;
    color: white !important;
}

/* Headings */
h1 {
    color: var(--stars-primary) !important;
    font-weight: 700 !important;
    border-bottom: 3px solid var(--stars-accent);
    padding-bottom: 0.3em;
}
h2, h3 {
    color: var(--stars-primary) !important;
    font-weight: 600 !important;
}

/* Dataframes / tables */
.stDataFrame {
    border: 1px solid var(--stars-border);
    border-radius: 6px;
    overflow: hidden;
}

/* Expanders */
.streamlit-expanderHeader {
    background-color: var(--stars-bg-secondary) !important;
    border-radius: 4px;
    font-weight: 500;
    color: var(--stars-primary) !important;
}

/* Metrics */
[data-testid="stMetricValue"] {
    color: var(--stars-primary) !important;
    font-weight: 700 !important;
}

/* Alert boxes */
.stAlert [data-testid="stMarkdownContainer"] {
    font-size: 0.92rem;
}

/* File uploader */
[data-testid="stFileUploader"] {
    border: 2px dashed var(--stars-border) !important;
    border-radius: 8px;
    padding: 1rem;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--stars-primary) !important;
}

/* Selectbox / inputs */
.stSelectbox > div > div, .stTextInput > div > div > input {
    border-color: var(--stars-border) !important;
}
.stSelectbox > div > div:focus-within, .stTextInput > div > div > input:focus {
    border-color: var(--stars-primary) !important;
    box-shadow: 0 0 0 1px var(--stars-primary) !important;
}

/* Dividers */
hr {
    border-top: 1px solid var(--stars-border) !important;
}

/* Footer badge */
.stars-footer {
    text-align: center;
    padding: 2rem 0 1rem;
    color: var(--stars-text-muted);
    font-size: 0.8rem;
    border-top: 1px solid var(--stars-border);
    margin-top: 3rem;
}
.stars-footer a { color: var(--stars-primary); text-decoration: none; }

/* Logo container in sidebar */
.sidebar-logo {
    text-align: center;
    padding: 1.5rem 0 1rem;
    margin-bottom: 0.5rem;
}
.sidebar-logo .logo-row {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 12px;
}
.sidebar-logo .logo-text {
    font-size: 1.6rem;
    font-weight: 800;
    letter-spacing: 3px;
    color: #FFFFFF !important;
}
.sidebar-logo .logo-sub {
    font-size: 0.7rem;
    color: var(--stars-accent) !important;
    letter-spacing: 3px;
    text-transform: uppercase;
    margin-top: 4px;
}

/* Ripple effect — concentric arcs */
.ripple-icon {
    position: relative;
    width: 40px;
    height: 40px;
    flex-shrink: 0;
}
.ripple-icon .arc {
    position: absolute;
    border: 1.5px solid rgba(255,255,255,0.6);
    border-radius: 50%;
    border-right-color: transparent;
    border-bottom-color: transparent;
    top: 50%;
    left: 50%;
}
.ripple-icon .arc:nth-child(1) {
    width: 10px; height: 10px;
    margin-top: -5px; margin-left: -5px;
    border-color: rgba(255,255,255,0.9);
    border-right-color: transparent;
    border-bottom-color: transparent;
}
.ripple-icon .arc:nth-child(2) {
    width: 18px; height: 18px;
    margin-top: -9px; margin-left: -9px;
    border-color: rgba(255,255,255,0.7);
    border-right-color: transparent;
    border-bottom-color: transparent;
}
.ripple-icon .arc:nth-child(3) {
    width: 26px; height: 26px;
    margin-top: -13px; margin-left: -13px;
    border-color: rgba(255,255,255,0.5);
    border-right-color: transparent;
    border-bottom-color: transparent;
}
.ripple-icon .arc:nth-child(4) {
    width: 34px; height: 34px;
    margin-top: -17px; margin-left: -17px;
    border-color: rgba(255,255,255,0.3);
    border-right-color: transparent;
    border-bottom-color: transparent;
}
/* Subtle pulse animation */
@keyframes ripple-pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.6; }
}
.ripple-icon .arc:nth-child(3) { animation: ripple-pulse 3s ease-in-out infinite; }
.ripple-icon .arc:nth-child(4) { animation: ripple-pulse 3s ease-in-out 0.5s infinite; }
</style>
"""


def inject_styles():
    """Inyecta el CSS corporativo STARS REI en la app."""
    st.markdown(_STARS_CSS, unsafe_allow_html=True)
