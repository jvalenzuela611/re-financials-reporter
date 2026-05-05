# src/config.py

"""
Configuración central: constantes, catálogo de activos, opciones de clasificación,
y modelos AI disponibles.
"""

# Catálogo de activos / portafolios
ASSET_CATALOG = {
    # ── Oakland Portfolio ─────────────────────────────────────────────────────
    "Portfolio Oakland": {
        "buildings": ["Alice House JV LLC", "295 29th Street JV LLC"],
        "consolidate": True,
        "description": "Alice House + Edson House (295 29th Street) — totales consolidados",
    },
    "Alice House JV LLC": {
        "buildings": ["Alice House JV LLC"],
        "consolidate": False,
        "description": "Individual — Alice House (Oakland)",
    },
    "295 29th Street JV LLC": {
        "buildings": ["295 29th Street JV LLC"],
        "consolidate": False,
        "description": "Individual — Edson House (295 29th Street, Oakland)",
    },
    # ── Activos individuales con capital de terceros ──────────────────────────
    # Nota: las keys de este dict deben coincidir con los asset_keys de
    # ASSET_CATALOGS (src/asset_catalogs.py) para que el resolver de catálogo
    # haga match exacto y todas las cuentas vengan del catalog (sin fallback).
    "14th & Callan Street JV LLC": {
        "buildings": ["14th & Callan Street JV LLC"],
        "consolidate": False,
        "description": "Individual — 14th & Callan Street JV LLC (Sares Regis)",
    },
    "929 Mass Ave (838)": {
        "buildings": ["929 Mass Ave (838)"],
        "consolidate": False,
        "description": "Individual — 929 Mass Ave (838)",
    },
    "The Val": {
        "buildings": ["The Val"],
        "consolidate": False,
        "description": "Individual — The Val",
    },
    "J 501 Estates": {
        "buildings": ["J 501 Estates"],
        "consolidate": False,
        "description": "Individual — J 501 Estates",
    },
    "624 Yale Apartments": {
        "buildings": ["624 Yale Apartments"],
        "consolidate": False,
        "description": "Individual — 624 Yale Apartments",
    },
    "Bridge at the Blockyard": {
        "buildings": ["Bridge at the Blockyard"],
        "consolidate": False,
        "description": "Individual — Bridge at the Blockyard (CWS, 12-month budget)",
    },
    "Walnut Street Wellesley": {
        "buildings": ["Walnut Street Wellesley"],
        "consolidate": False,
        "description": "Individual — Walnut Street Wellesley / Newton Wellesley Executive Park (NWEP, JDE / Summary formats)",
    },
    # ── Ballard Portfolio (los números NUNCA se agregan — siempre individual) ─
    "KOI Apartments": {
        "buildings": ["KOI Apartments"],
        "consolidate": False,
        "description": "Individual — KOI Apartments (Ballard Portfolio)",
    },
    "The Wilcox": {
        "buildings": ["The Wilcox"],
        "consolidate": False,
        "description": "Individual — The Wilcox (Ballard Portfolio)",
    },
}

# Opciones de clasificación L1
# "Total Capex" = cuenta de capital; se trackea como L1 propia pero NO entra en el
#                 waterfall de Net Income (va debajo del NI, como NA a efectos de NI).
# "NA"          = cuenta excluida de TODOS los totales L1 (ni siquiera se tracke).
L1_OPTIONS = [
    "Income", "Operating Expenses", "Real Estate Taxes",
    "NOI", "Interest Expense", "Non-Operating Expenses", "Net Income",
    "Total Capex",
    "NA",
]

# Orden canónico de las líneas L1 al presentar el estado de resultados.
# Se usa para ordenar cuentas L3 en los Excel descargables (account_tracking, master)
# de forma que el analista las vea en el orden natural del P&L:
#   Income > OpEx > RET > NOI > Interest > Non-Op > Net Income > Capex > NA
L1_PL_ORDER = {
    "Income": 1,
    "Operating Expenses": 2,
    "Real Estate Taxes": 3,
    "NOI": 4,
    "Interest Expense": 5,
    "Non-Operating Expenses": 6,
    "Net Income": 7,
    "Total Capex": 8,
    "NA": 99,        # al fondo (y sobre cualquier L1 no catalogada)
    "Unknown": 100,
}


def l1_pl_sort_key(parent_line: str) -> int:
    """Devuelve la posición canónica P&L para un L1; L1s desconocidos van al final."""
    return L1_PL_ORDER.get(parent_line, 100)

# Directorio de masters
MASTERS_DIR = "./masters"

# Modelos AI disponibles (sincronizar con ai_analyst.py)
OPENAI_MODELS = ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "o3-mini"]
ANTHROPIC_MODELS = [
    "claude-opus-4-5",
    "claude-sonnet-4-5",
    "claude-haiku-4-5",
]
