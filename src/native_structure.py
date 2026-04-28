"""
Módulo de Estructura Nativa: detecta la jerarquía del reporte del socio
y mapea cuentas a las líneas principales del reporte financiero.

Líneas principales (según reportes reales):
  Income, Operating Expenses, Real Estate Taxes, NOI,
  Non-Operating Expenses, Interest, Net Income, Capex
"""

import re
import pandas as pd
from typing import Dict, List, Optional
from dataclasses import dataclass

try:
    from src.asset_catalogs import ASSET_CATALOGS, get_l1_for_account
except ImportError:
    # Permite importar el módulo aunque el catálogo aún no esté generado.
    ASSET_CATALOGS = {}
    def get_l1_for_account(asset_key, account_code):
        return None


# Mapeo de secciones del reporte del socio a líneas principales.
# Expandido para soportar terminología variada de distintos socios/formatos.
REPORT_LINE_RULES = [
    # (pattern en description de totales, línea del reporte, signo)
    # Nota: las reglas más específicas deben ir ANTES que las genéricas:
    # el loop hace "pattern in desc_lower" y para en el primer match.

    # === INCOME / REVENUES (específicas antes que genéricas) ===
    ("total financial income", "_sub_income", 1),  # Val: subtotal de Financial Income (NO es Income total)
    ("total rental revenue", "_sub_income", 1),    # JDE/Walnut: subtotal de Rental
    ("total revenues", "Income", 1),
    ("total revenue", "Income", 1),                # JDE/Walnut: singular "TOTAL REVENUE"
    ("total income", "Income", 1),
    # Subtotales de income — NO son el L1 total; son sub-secciones
    ("total rental income", "_sub_income", 1),
    ("total other rental income", "_sub_income", 1),
    ("total rent revenue", "_sub_income", 1),
    ("total other income", "_sub_income", 1),
    ("total concessions", "_sub_income", 1),
    ("potential rent", "_sub_income", 1),
    ("total other revenue", "_sub_income", 1),
    ("total miscellaneous", "_sub_income", 1),
    ("total laundry", "_sub_income", 1),
    ("total parking", "_sub_income", 1),
    ("total fee income", "_sub_income", 1),

    # === OPERATING EXPENSES ===
    # Val/CMC: sub-totales específicos (antes que genéricos)
    ("total noncontrollable expenses", "_sub_opex", -1),
    ("total controllable expenses", "_agg_opex", -1),  # Aggregator (rollup de Renting/Utility/Payroll/Operating) — NO es L2
    ("total renting expense", "_sub_opex", -1),
    ("total utility expense", "_sub_opex", -1),
    ("total payroll expense", "_sub_opex", -1),
    ("total operating expense", "_sub_opex", -1),  # Val: "TOTAL Operating Expense" (singular, es subgrupo no total)
    ("total operating expenses", "Operating Expenses", -1),
    ("total expenses before", "Operating Expenses", -1),  # Val: "TOTAL Expenses before RE Tax Depr Debt"
    ("total expenses", "Operating Expenses", -1),
    ("total operating costs", "Operating Expenses", -1),
    ("total payroll", "_sub_opex", -1),
    ("total salaries", "_sub_opex", -1),
    ("total personnel", "_sub_opex", -1),
    ("total staffing", "_sub_opex", -1),
    ("total utilities", "_sub_opex", -1),
    ("total turnover", "_sub_opex", -1),
    ("total make ready", "_sub_opex", -1),
    ("total unit turnover", "_sub_opex", -1),
    ("total repairs", "_sub_opex", -1),
    ("total maintenance", "_sub_opex", -1),
    ("total r&m", "_sub_opex", -1),
    ("total contract", "_sub_opex", -1),
    ("total professional", "_sub_opex", -1),
    ("total advertising", "_sub_opex", -1),
    ("total marketing", "_sub_opex", -1),
    ("total management", "_sub_opex", -1),
    ("total property g", "_sub_opex", -1),
    ("total general", "_sub_opex", -1),
    ("total administrative", "_sub_opex", -1),
    ("total admin", "_sub_opex", -1),
    ("total insurance", "_sub_opex", -1),
    ("total security", "_sub_opex", -1),
    ("total landscaping", "_sub_opex", -1),
    ("total cleaning", "_sub_opex", -1),
    ("total janitorial", "_sub_opex", -1),

    # === REAL ESTATE TAXES ===
    ("total property taxes", "Real Estate Taxes", -1),
    ("total real estate tax", "Real Estate Taxes", -1),
    ("total real estate & property tax", "Real Estate Taxes", -1),  # KOI/Wilcox YSI
    ("total real estate and property tax", "Real Estate Taxes", -1),
    ("total taxes", "Real Estate Taxes", -1),
    ("real estate & property tax", "Real Estate Taxes", -1),
    ("real estate tax", "_sub_taxes", -1),
    ("property tax", "_sub_taxes", -1),

    # === NOI ===
    ("net operating income", "NOI", 1),
    ("noi", "NOI", 1),

    # === INTEREST EXPENSE (L1 separada, NO incluida en OPEX ni en NOI) ===
    ("total interest", "Interest Expense", -1),
    ("interest expense", "Interest Expense", -1),
    ("total debt service", "Interest Expense", -1),
    ("total mortgage", "Interest Expense", -1),
    ("total interest and financing", "Interest Expense", -1),
    ("total financing expenses", "Interest Expense", -1),
    ("total financial", "Interest Expense", -1),  # Val: "TOTAL Financial" (sección de deuda/financiamiento)

    # === NON-OPERATING EXPENSES ===
    ("total company general", "Non-Operating Expenses", -1),
    ("total non-operating", "Non-Operating Expenses", -1),
    ("total non operating", "Non-Operating Expenses", -1),
    ("total other expenses", "Non-Operating Expenses", -1),
    ("total corporate", "Non-Operating Expenses", -1),

    # === CAPEX ===
    ("total capital", "Total Capex", -1),
    ("total capex", "Total Capex", -1),
    ("total capital expenditure", "Total Capex", -1),

    # === NET INCOME ===
    ("net income", "Net Income", 1),
    ("net profit", "Net Income", 1),
    ("net loss", "Net Income", 1),
]

# Yardi Variance Report (Yale) usa subtotales cuya descripción es solo el nombre del L1
# (ej: "Income", "Operating Expenses", "Non-Operating Expenses"). Para no contaminar el
# match por substring (que mataría "Miscellaneous income" = L1), estos matchean EXACTO
# y solo cuando la fila es de tipo 'total'.
EXACT_L1_TOTAL_NAMES = {
    "income": "Income",
    "operating expenses": "Operating Expenses",
    "non-operating expenses": "Non-Operating Expenses",
    "non operating expenses": "Non-Operating Expenses",
    "capital expenditures": "Total Capex",
    "investment committee (ic) capital expenditures": "Total Capex",
    "net operating income": "NOI",
    "net income": "Net Income",
    # Yardi Variance Report (Yale) — subtotales específicos
    "taxes": "Real Estate Taxes",
    "debt servicing costs": "Interest Expense",
    "ownership / partnership expenses": "Non-Operating Expenses",
    "ownership partnership expenses": "Non-Operating Expenses",
}

# Secciones de nivel 2 y su parent de nivel 1.
# Expandido para soportar terminología variada entre socios.
SECTION_TO_L1 = {
    # ── Income / Revenue ──────────────────────────────────────────
    "revenues": "Income",
    "revenue": "Income",
    "income": "Income",
    "gross revenue": "Income",
    "gross potential rent": "Income",
    "effective gross revenue": "Income",
    "net rental revenue": "Income",
    "rental income": "Income",
    "rent revenue": "Income",
    "rental revenue": "Income",
    "residential revenue": "Income",
    "residential rental revenue": "Income",
    "commercial revenue": "Income",
    "retail revenue": "Income",
    "concessions and vacancy": "Income",
    "concessions": "Income",
    "vacancy": "Income",
    "vacancy loss": "Income",
    "economic occupancy adjustments": "Income",
    "other income": "Income",
    "other revenue": "Income",
    "miscellaneous income": "Income",
    "miscellaneous revenue": "Income",
    "laundry income": "Income",
    "laundry revenue": "Income",
    "parking income": "Income",
    "parking revenue": "Income",
    "fee income": "Income",
    "fee revenue": "Income",
    "utility reimbursements": "Income",
    "utility income": "Income",
    "rubs income": "Income",
    "rubs": "Income",
    "other fees": "Income",
    "resident services": "Income",
    "bad debt": "Income",
    "bad debt recovery": "Income",
    "lease termination": "Income",

    # ── Operating Expenses ────────────────────────────────────────
    "expenses": "Operating Expenses",
    "operating expenses": "Operating Expenses",
    "operating costs": "Operating Expenses",
    "controllable expenses": "Operating Expenses",
    "noncontrollable expenses": "Operating Expenses",
    "non controllable expenses": "Operating Expenses",
    # Payroll
    "payroll and related": "Operating Expenses",
    "payroll": "Operating Expenses",
    "payroll expense": "Operating Expenses",
    "payroll expenses": "Operating Expenses",
    "salaries and wages": "Operating Expenses",
    "salaries": "Operating Expenses",
    "personnel": "Operating Expenses",
    "staffing": "Operating Expenses",
    # Leasing / Marketing
    "leasing": "Operating Expenses",
    "leasing expense": "Operating Expenses",
    "leasing expenses": "Operating Expenses",
    "marketing": "Operating Expenses",
    "marketing expense": "Operating Expenses",
    "marketing expenses": "Operating Expenses",
    "advertising": "Operating Expenses",
    "advertising and marketing": "Operating Expenses",
    # Maintenance / R&M
    "maintenance": "Operating Expenses",
    "maintenance expense": "Operating Expenses",
    "maintenance expenses": "Operating Expenses",
    "repairs and maintenance": "Operating Expenses",
    "repairs & maintenance": "Operating Expenses",
    "r&m": "Operating Expenses",
    "make ready": "Operating Expenses",
    "make ready expenses": "Operating Expenses",
    "unit turnover": "Operating Expenses",
    "unit turns": "Operating Expenses",
    "turnover": "Operating Expenses",
    "turn expense": "Operating Expenses",
    "turn expenses": "Operating Expenses",
    # Contract services
    "contract services": "Operating Expenses",
    "contracted services": "Operating Expenses",
    "contract expense": "Operating Expenses",
    "contract expenses": "Operating Expenses",
    "professional services": "Operating Expenses",
    # Utilities
    "utilities": "Operating Expenses",
    "utility expense": "Operating Expenses",
    "utility expenses": "Operating Expenses",
    # Insurance
    "insurance": "Operating Expenses",
    "insurance2": "Operating Expenses",
    "insurance expense": "Operating Expenses",
    "insurance expenses": "Operating Expenses",
    # Admin / G&A
    "administrative": "Operating Expenses",
    "administrative expense": "Operating Expenses",
    "administrative expenses": "Operating Expenses",
    "management fees": "Operating Expenses",
    "management fee": "Operating Expenses",
    "management expense": "Operating Expenses",
    "management expenses": "Operating Expenses",
    "property g & a": "Operating Expenses",
    "property g&a": "Operating Expenses",
    "general & administrative": "Operating Expenses",
    "general and administrative": "Operating Expenses",
    "g&a": "Operating Expenses",
    # Other OpEx
    "security": "Operating Expenses",
    "landscaping": "Operating Expenses",
    "grounds": "Operating Expenses",
    "cleaning": "Operating Expenses",
    "janitorial": "Operating Expenses",
    "common area": "Operating Expenses",
    "exterior maintenance": "Operating Expenses",
    "interior maintenance": "Operating Expenses",
    "fire protection": "Operating Expenses",
    "fire life safety": "Operating Expenses",
    "pool": "Operating Expenses",
    "elevator": "Operating Expenses",

    # ── Real Estate Taxes ─────────────────────────────────────────
    "property taxes": "Real Estate Taxes",
    "property tax": "Real Estate Taxes",
    "real estate taxes": "Real Estate Taxes",
    "real estate tax": "Real Estate Taxes",
    "taxes": "Real Estate Taxes",
    "tax expense": "Real Estate Taxes",

    # ── Interest Expense ──────────────────────────────────────────
    "interest": "Interest Expense",
    "interest expense": "Interest Expense",
    "interest and financing expenses": "Interest Expense",
    "interest & financing expenses": "Interest Expense",
    "financing expenses": "Interest Expense",
    "financial": "Interest Expense",
    "debt service": "Interest Expense",
    "mortgage": "Interest Expense",
    "loan expense": "Interest Expense",

    # ── Non-Operating ─────────────────────────────────────────────
    "company general & administrative": "Non-Operating Expenses",
    "company g&a": "Non-Operating Expenses",
    "non-operating expenses": "Non-Operating Expenses",
    "non operating expenses": "Non-Operating Expenses",
    "other expenses": "Non-Operating Expenses",
    "corporate expenses": "Non-Operating Expenses",
    "professional fees": "Non-Operating Expenses",
    "legal fees": "Non-Operating Expenses",
    "accounting fees": "Non-Operating Expenses",
    "unrealized gain/loss": "Non-Operating Expenses",
    "unrealized gain loss": "Non-Operating Expenses",
    "depreciation": "Non-Operating Expenses",
    "amortization": "Non-Operating Expenses",

    # ── Capital Expenditures ──────────────────────────────────────
    "capital expenditures": "Total Capex",
    "capital improvements": "Total Capex",
    "capital expense": "Total Capex",
    "capital expenses": "Total Capex",
    "capex": "Total Capex",
    "intercompany capital": "Total Capex",
    "ic capital": "Total Capex",
    "construction": "Total Capex",
    "renovations": "Total Capex",
    "rehabilitation": "Total Capex",
    "investment committee (ic) capital expenditures": "Total Capex",
    "investment committee capital expenditures": "Total Capex",

    # ── Yardi Variance Report (Yale) — secciones y sub-secciones ─
    # Income
    "rental revenue - residential": "Income",
    "rental revenue adjustments - residential": "Income",
    "rental revenue": "Income",
    # Operating Expenses — sub-secciones Yale
    "fire life and safety": "Operating Expenses",
    "unit turn expense": "Operating Expenses",
    "unit turn expenses": "Operating Expenses",
    "general repairs and maintenance": "Operating Expenses",
    "general repairs": "Operating Expenses",
    "repairs and maintenance": "Operating Expenses",
    "general and administration": "Operating Expenses",
    "insurance coverage": "Operating Expenses",
    # Non-Operating Yale
    "debt servicing costs": "Interest Expense",
    "ownership / partnership expenses": "Non-Operating Expenses",
    "ownership partnership expenses": "Non-Operating Expenses",

    # ── BCR (501 Estates) — secciones específicas ────────────────
    # Income
    "other rental income": "Income",
    # Operating Expenses — sub-secciones
    "payroll & benefits": "Operating Expenses",
    "payroll and benefits": "Operating Expenses",
    "general maintenance": "Operating Expenses",
    "make - ready/redecorating": "Operating Expenses",
    "make-ready/redecorating": "Operating Expenses",
    "make ready/redecorating": "Operating Expenses",
    "make ready redecorating": "Operating Expenses",
    "recreational amenities": "Operating Expenses",
    "advertising /marketing/promotions": "Operating Expenses",
    "advertising/marketing/promotions": "Operating Expenses",
    "advertising/marketing/promotion": "Operating Expenses",
    # Non-Operating
    "partnership/owner expenses": "Non-Operating Expenses",
    "partnership / owner expenses": "Non-Operating Expenses",
    "partnership owner expenses": "Non-Operating Expenses",
    # Interest
    "debt service": "Interest Expense",
    # Non-Op (Depreciation)
    "depreciation & amortization": "Non-Operating Expenses",
    "depreciation and amortization": "Non-Operating Expenses",
    # Capex — Construction in Progress
    "construction in progress-capital/renovation": "Total Capex",
    "construction in progress - capital/renovation": "Total Capex",
    "construction in progress-routine replacement": "Total Capex",
    "construction in progress - routine replacement": "Total Capex",
    "construction in progress": "Total Capex",
}

# ─────────────────────────────────────────────────────────────────
# FALLBACK CLASSIFICATION — Código de cuenta + Keywords descripción
# Para activos donde los headers de sección no matchean SECTION_TO_L1
# ─────────────────────────────────────────────────────────────────

# Rangos de códigos por sub-rango (Yardi estándar con overrides específicos).
# Evaluación: primer match gana. Los rangos más específicos van ANTES que los genéricos.
# Cubre:
#   - Yale/Yardi moderno (4xxx income, 5xxx opex, 5850-5899 RET, 7xxx interest, 8xxx non-op, 1xxx capex)
#   - Oakland/Alice/Edson (4xxx income, 8xxx mixto: 8010-8089 payroll, 8111-8299 opex R&M/Utilities,
#     8310 RET, 8340-8575 opex, 8610-8618 bad debt/income contra, 8620-8840 property G&A,
#     8890-8898 non-op corporate, 8901-8995 opex R&M extended, 9xxx interest)
#   - CMC (Val) usa rangos distintos (5xxx income, 6xxx opex) y tiene parser dedicado
ACCOUNT_CODE_RANGES = [
    # ── Income 3xxx (501 Estates BCR) ────────────────────────────
    (3000, 3999, "Income"),

    # ── Income (4xxx siempre income en Yardi/Oakland) ────────────
    (4000, 4999, "Income"),

    # ── Yardi moderno: 5xxx opex + 5850-5899 taxes ───────────────
    (5000, 5849, "Operating Expenses"),
    (5850, 5899, "Real Estate Taxes"),
    (5900, 5999, "Operating Expenses"),
    (6000, 6999, "Operating Expenses"),

    # ── Interest / Financing ─────────────────────────────────────
    (7000, 7999, "Interest Expense"),

    # ── Oakland 8xxx (overrides específicos antes del catch-all) ─
    (8010, 8089, "Operating Expenses"),   # Payroll
    (8111, 8199, "Operating Expenses"),   # Utilities (incl reimburses)
    (8201, 8299, "Operating Expenses"),   # Turnover, R&M, Contract Services
    (8300, 8309, "Real Estate Taxes"),    # (espacio reservado para taxes)
    (8310, 8310, "Real Estate Taxes"),    # Alice/Edson Real Estate Taxes (cuenta única)
    (8311, 8339, "Real Estate Taxes"),    # Otras taxes propiedad
    (8340, 8349, "Operating Expenses"),   # Other Taxes (Property G&A)
    (8350, 8449, "Operating Expenses"),   # Insurance + varios
    (8450, 8489, "Operating Expenses"),   # Management Fees
    (8490, 8499, "Non-Operating Expenses"),  # Legal Fees (reportado en Company G&A)
    (8500, 8599, "Operating Expenses"),   # Advertising & Marketing
    (8610, 8619, "Income"),                # Bad Debt / Write-offs (income contra en Oakland)
    (8620, 8839, "Operating Expenses"),   # Property G&A (Bank, Software, Supplies, etc.)
    (8840, 8889, "Operating Expenses"),   # Non-recurring dentro de Property G&A
    (8890, 8899, "Non-Operating Expenses"),  # Company G&A (Asset Mgmt, Audit, Franchise Tax, etc.)
    (8900, 8999, "Operating Expenses"),   # R&M extendido (8901-8995 en Oakland)

    # ── Non-Operating extendido (catch-all 9xxx menos interest) ──
    (9010, 9019, "Interest Expense"),     # Oakland: 9010 = Interest, 9010-0010 = Financing
    (9020, 9999, "Non-Operating Expenses"),

    # ── Capital Expenditures (1xxx y 2xxx Yardi) ─────────────────
    (1000, 1999, "Total Capex"),
    (2000, 2999, "Total Capex"),
]

# Keywords en descripción de cuenta (más específicos primero, primer match gana)
DESCRIPTION_KEYWORD_L1 = [
    # Real Estate Taxes — antes que OpEx genérico
    ("property tax",           "Real Estate Taxes"),
    ("real estate tax",        "Real Estate Taxes"),
    ("county tax",             "Real Estate Taxes"),
    ("city tax",               "Real Estate Taxes"),
    ("special assessment",     "Real Estate Taxes"),
    # Income / Revenue
    ("market rent",            "Income"),
    ("gross rent",             "Income"),
    ("potential rent",         "Income"),
    ("rent revenue",           "Income"),
    ("rental income",          "Income"),
    ("rental revenue",         "Income"),
    ("gain to lease",          "Income"),
    ("loss to lease",          "Income"),
    ("vacancy loss",           "Income"),
    ("vacancy",                "Income"),
    ("concession",             "Income"),
    ("other income",           "Income"),
    ("fee income",             "Income"),
    ("parking fee",            "Income"),
    ("parking income",         "Income"),
    ("storage fee",            "Income"),
    ("laundry income",         "Income"),
    ("late fee",               "Income"),
    ("nsf fee",                "Income"),
    ("application fee",        "Income"),
    ("lease termination fee",  "Income"),
    ("damage fee",             "Income"),
    ("damage fees",            "Income"),
    ("bad debt",               "Income"),
    ("rub income",             "Income"),
    ("utility reim",           "Income"),
    ("resident reimb",         "Income"),
    # ── Loan-related (orden importa: específicos antes que genéricos) ────────
    # Principal payments NO van al P&L (son movimientos de balance sheet).
    # Si aparecen mezclados en el reporte del socio, los marcamos NA.
    # Bug fix 501 Estates / Yale (feedback Florencia 23-04-2026).
    ("principal payment",      "NA"),
    ("loan principal",         "NA"),
    ("principal repayment",    "NA"),
    ("principal reduction",    "NA"),
    ("mortgage principal",     "NA"),
    # Loan cost amortization → Non-Operating (no es interest real, es amortización contable)
    ("amortization of loan",   "Non-Operating Expenses"),
    ("loan cost amortiz",      "Non-Operating Expenses"),
    ("loan fee amortiz",       "Non-Operating Expenses"),
    ("amortization loan",      "Non-Operating Expenses"),
    ("financing cost amortiz", "Non-Operating Expenses"),
    ("deferred financing",     "Non-Operating Expenses"),
    # Interest real — antes que OpEx genérico
    ("interest expense",       "Interest Expense"),
    ("mortgage interest",      "Interest Expense"),
    ("loan interest",          "Interest Expense"),
    ("debt service",           "Interest Expense"),
    # Capital — antes que OpEx genérico
    ("ic non unit",            "Total Capex"),
    ("ic construction",        "Total Capex"),
    ("ic interior",            "Total Capex"),
    ("ic exterior",            "Total Capex"),
    ("capital improvement",    "Total Capex"),
    ("capital expenditure",    "Total Capex"),
    ("construction mgmt",      "Total Capex"),
    ("construction fee",       "Total Capex"),
    ("construction management","Total Capex"),
    # Non-Operating — antes que OpEx genérico
    ("professional fee",       "Non-Operating Expenses"),
    ("o / p -",                "Non-Operating Expenses"),  # Formato "O / P - Professional Fees"
    ("o/p -",                  "Non-Operating Expenses"),
    ("legal fee",              "Non-Operating Expenses"),
    ("attorney fee",           "Non-Operating Expenses"),
    ("accounting fee",         "Non-Operating Expenses"),
    ("audit fee",              "Non-Operating Expenses"),
    ("consulting fee",         "Non-Operating Expenses"),
    ("fixed asset expense",    "Non-Operating Expenses"),
    ("depreciation",           "Non-Operating Expenses"),
    ("amortization",           "Non-Operating Expenses"),
    # Operating Expenses — Payroll
    ("payroll",                "Operating Expenses"),
    ("salary",                 "Operating Expenses"),
    ("wages",                  "Operating Expenses"),
    ("employee benefit",       "Operating Expenses"),
    ("worker comp",            "Operating Expenses"),
    ("health insurance",       "Operating Expenses"),
    ("401k",                   "Operating Expenses"),
    # Operating Expenses — Maintenance / R&M
    ("maintenance",            "Operating Expenses"),
    ("repair",                 "Operating Expenses"),
    ("hvac",                   "Operating Expenses"),
    ("plumbing",               "Operating Expenses"),
    ("electrical",             "Operating Expenses"),
    ("elevator",               "Operating Expenses"),
    ("appliance",              "Operating Expenses"),
    ("doors and windows",      "Operating Expenses"),
    ("door and window",        "Operating Expenses"),
    ("fire life",              "Operating Expenses"),
    ("fire safety",            "Operating Expenses"),
    ("pest control",           "Operating Expenses"),
    ("exterminator",           "Operating Expenses"),
    ("make ready",             "Operating Expenses"),
    ("unit turn",              "Operating Expenses"),
    ("boiler",                 "Operating Expenses"),  # 5xxx boiler = OpEx; 1xxx boiler = CapEx (code range toma precedencia)
    # Operating Expenses — Utilities
    ("electric",               "Operating Expenses"),
    ("gas utility",            "Operating Expenses"),
    ("water - ",               "Operating Expenses"),
    ("water common",           "Operating Expenses"),
    ("sewer",                  "Operating Expenses"),
    ("trash",                  "Operating Expenses"),
    ("garbage",                "Operating Expenses"),
    ("waste removal",          "Operating Expenses"),
    ("recycling",              "Operating Expenses"),
    # Operating Expenses — Marketing / Leasing
    ("marketing",              "Operating Expenses"),
    ("advertising",            "Operating Expenses"),
    ("leasing",                "Operating Expenses"),
    ("listing fee",            "Operating Expenses"),
    ("locator fee",            "Operating Expenses"),
    ("model unit",             "Operating Expenses"),
    ("signage",                "Operating Expenses"),
    # Operating Expenses — Insurance
    ("insurance",              "Operating Expenses"),
    ("liability",              "Operating Expenses"),
    # Operating Expenses — Admin / Management
    ("management fee",         "Operating Expenses"),
    ("property management",    "Operating Expenses"),
    ("administrative",         "Operating Expenses"),
    ("office supply",          "Operating Expenses"),
    ("office equipment",       "Operating Expenses"),
    ("postage",                "Operating Expenses"),
    ("telephone",              "Operating Expenses"),
    ("software",               "Operating Expenses"),
    ("bank fee",               "Operating Expenses"),
    # Operating Expenses — Contract Services
    ("landscaping",            "Operating Expenses"),
    ("landscape",              "Operating Expenses"),
    ("grounds",                "Operating Expenses"),
    ("snow removal",           "Operating Expenses"),
    ("security",               "Operating Expenses"),
    ("patrol",                 "Operating Expenses"),
    ("janitorial",             "Operating Expenses"),
    ("cleaning",               "Operating Expenses"),
    ("pool",                   "Operating Expenses"),
    ("elevator service",       "Operating Expenses"),
    ("garbage - removal",      "Operating Expenses"),
]


# Aliases: nombres alternativos detectables en los Excel -> canonical asset_key.
# Ej: el xlsb de Edson tiene "Edson House" en el header, pero la key canónica es
# "295 29th Street JV LLC". Los aliases se chequean antes del match por substring.
ASSET_ALIASES = {
    "edson house": "295 29th Street JV LLC",
    "edson": "295 29th Street JV LLC",
    "295 29th street": "295 29th Street JV LLC",
    "alice house": "Alice House JV LLC",
    "alice": "Alice House JV LLC",
    "the val": "The Val",
    "val": "The Val",
    "walnut street wellesley owner llc": "Walnut Street Wellesley",
    "walnut street wellesley owner": "Walnut Street Wellesley",
    "walnut street wellesley": "Walnut Street Wellesley",
    "walnut street": "Walnut Street Wellesley",
    "the wilcox (pn10680)": "The Wilcox",
    "the wilcox": "The Wilcox",
    "wilcox": "The Wilcox",
    "koi (pn10675)": "KOI Apartments",
    "koi apartments": "KOI Apartments",
    "koi": "KOI Apartments",
    "929 mass ave (838)": "929 Mass Ave (838)",
    "929 mass ave": "929 Mass Ave (838)",
    "929 mass": "929 Mass Ave (838)",
}


def _resolve_asset_key(ingested_files: list) -> Optional[str]:
    """
    Resuelve el asset_key (key de ASSET_CATALOGS) a partir del campo `building`
    del primer archivo ingerido.

    Orden: exact -> case-insens -> ASSET_ALIASES -> substring (cualquier dirección).
    """
    if not ingested_files or not ASSET_CATALOGS:
        return None
    raw = (getattr(ingested_files[0], 'building', '') or '').strip()
    if not raw:
        return None
    if raw in ASSET_CATALOGS:
        return raw
    raw_lower = raw.lower()
    for key in ASSET_CATALOGS:
        if key.lower() == raw_lower:
            return key
    if raw_lower in ASSET_ALIASES and ASSET_ALIASES[raw_lower] in ASSET_CATALOGS:
        return ASSET_ALIASES[raw_lower]
    for alias, target in ASSET_ALIASES.items():
        if alias in raw_lower and target in ASSET_CATALOGS:
            return target
    for key in ASSET_CATALOGS:
        kl = key.lower()
        if kl in raw_lower or raw_lower in kl:
            return key
    return None


def _classify_by_account_code(acct_code: str) -> Optional[str]:
    """
    Clasifica una cuenta L3 por rango de código numérico Yardi estándar.
    Extrae los primeros 4 dígitos (ej: '5855-0000' → 5855).
    Retorna parent_line o None si no hay match.
    """
    if not acct_code:
        return None
    m = re.match(r'^(\d{4})', str(acct_code).strip())
    if not m:
        return None
    code_num = int(m.group(1))
    for min_c, max_c, l1 in ACCOUNT_CODE_RANGES:
        if min_c <= code_num <= max_c:
            return l1
    return None


def _classify_by_description(description: str) -> Optional[str]:
    """
    Clasifica una cuenta L3 por keywords en su descripción.
    Retorna parent_line o None si no hay match.
    """
    if not description:
        return None
    desc_lower = description.lower().strip()
    for keyword, l1 in DESCRIPTION_KEYWORD_L1:
        if keyword in desc_lower:
            return l1
    return None


def _match_section_to_l1(section_desc: str, current_l1: str) -> str:
    """
    Intenta clasificar un header de sección a un L1.
    Primero busca match exacto en SECTION_TO_L1, luego substring.
    Retorna el L1 encontrado o current_l1 si no hay match.
    """
    desc = section_desc.lower().strip()
    desc_normalized = re.sub(r'\s*\d+$', '', desc).strip()

    # 1. Match exacto
    result = SECTION_TO_L1.get(desc) or SECTION_TO_L1.get(desc_normalized)
    if result:
        return result

    # 2. Substring: buscar si alguna clave del dict está contenida en el header
    #    (útil para "Leasing Expense and Advertising" → matchea "leasing")
    for key, l1 in SECTION_TO_L1.items():
        if len(key) >= 5 and key in desc:
            return l1

    return current_l1


@dataclass
class FinancialLine:
    """Una línea del reporte financiero."""
    name: str
    level: int  # 1 = principal, 2 = categoría, 3 = subcuenta
    budget_current: float
    actual_current: float
    variance_current: float
    variance_pct_current: Optional[float]
    budget_ytd: Optional[float]
    actual_ytd: Optional[float]
    variance_ytd: Optional[float]
    variance_pct_ytd: Optional[float]
    parent_line: str  # Línea L1 padre
    account_code: str = ""
    children: List = None
    partner_comment: str = ""
    section_name: str = ""
    # Posición original de la cuenta en el Excel del partner (row_num). Se usa
    # para ordenar outputs que deben respetar el orden natural del P&L del
    # partner (master, account tracking). No pisa la sort por variance que
    # hacen otros outputs (bullets, top movers).
    partner_row_order: int = 0
    # Auditoría: quién decidió el L1 final de esta cuenta. Valores:
    #   'catalog'             — asset_catalogs.py (Florencia)
    #   'section'             — heredado del section header del Excel del partner
    #   'code_range'          — fallback por rango de código Yardi
    #   'description_keyword' — fallback por keyword en descripción
    #   'unknown'             — ningún criterio matchó
    #   'user_override'       — reclassification manual del analista en Step 2
    #   'cross_side_flip'     — user_override que además cruzó Income↔Expense
    #                            (el signo se flipeó para mantener NOI consistente)
    #   'post_processing'     — reclasificación interna (ej: Financing→Non-Op por desc)
    classification_source: str = "unknown"

    def __post_init__(self):
        if self.children is None:
            self.children = []


def _extract_period_sortkey(ingested) -> int:
    """
    Extrae una clave numérica year*100+month del filename/period_label
    para ordenar archivos cronológicamente. Si no puede parsear, retorna 0.

    Soporta patrones:
      - "2025-12"  (501 Estates)
      - "12-2025"  (Val)
      - "12.25"    (Yale)
      - "Alice 12" (Oakland, año inferido del period_label o default 2025)
      - "Dec 2025" (period_label standard)
    """
    from pathlib import Path as _Path
    filename = getattr(ingested, 'filename', '') or ''
    period = getattr(ingested, 'period_label', '') or ''
    haystack = f"{filename} {period}"

    # Pattern 1: YYYY-MM o YYYY.MM (ej: "2025-12", "2025_12")
    m = re.search(r'(20\d{2})[-_.](\d{1,2})\b', haystack)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            return year * 100 + month

    # Pattern 2: MM-YYYY (ej: "12-2025")
    m = re.search(r'\b(\d{1,2})[-_](20\d{2})\b', haystack)
    if m:
        month, year = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            return year * 100 + month

    # Pattern 3: MM.YY corta (ej: "12.25" = Dec 2025)
    m = re.search(r'\b(\d{1,2})\.(2\d)\b', haystack)
    if m:
        month, short_year = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            return (2000 + short_year) * 100 + month

    # Pattern 4: MonthName YYYY o YYYY MonthName (period_label "Dec 2025")
    month_names = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
                   'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12,
                   'january':1,'february':2,'march':3,'april':4,'june':6,
                   'july':7,'august':8,'september':9,'october':10,
                   'november':11,'december':12}
    m = re.search(r'\b(' + '|'.join(month_names.keys()) + r')\w*\b.*?(20\d{2})', haystack, re.IGNORECASE)
    if m:
        month = month_names[m.group(1).lower()]
        year = int(m.group(2))
        return year * 100 + month

    # Pattern 5: solo número al final del stem (ej: "Alice 12.xlsm" → month=12)
    stem = _Path(filename).stem
    m = re.search(r'\b(\d{1,2})\s*$', stem)
    if m:
        month = int(m.group(1))
        if 1 <= month <= 12:
            year_m = re.search(r'(20\d{2})', haystack)
            year = int(year_m.group(1)) if year_m else 2025
            return year * 100 + month

    return 0


def build_financial_structure(ingested_files: list, prior_quarter_files: list = None) -> Dict:
    """
    Construye la estructura financiera completa a partir de archivos ingeridos.

    Estrategia para construir Q (trimestre) y YTD:
    - Q = Suma de los "Current" (mensuales) de TODOS los archivos del trimestre
    - YTD = Columna "YTD" del ÚLTIMO archivo (ya viene acumulado)
    - Comentarios se recolectan de TODOS los archivos (el más reciente gana)

    Los archivos se ordenan cronológicamente (por filename/period) antes de procesar,
    para que el "último" sea siempre el mes más reciente independiente del orden de upload.

    Returns:
        Dict con l1_lines, l2_sections, l3_accounts, all_comments, raw_consolidated
    """
    if not ingested_files:
        raise ValueError("No hay archivos para procesar")

    # CAMBIO: concatenar TODOS los comentarios por account_code en orden cronológico.
    # Ya no se conserva solo el más reciente; se acumulan separados por " — ".
    all_comments_raw = []
    for ingested in ingested_files:
        all_comments_raw.extend(ingested.partner_comments)

    # Agrupar por account_code, preservando orden cronológico (orden de archivos)
    from collections import OrderedDict
    comments_by_key = OrderedDict()
    for c in all_comments_raw:
        key = c.get('account_code', '') or c.get('description', '')
        if key not in comments_by_key:
            comments_by_key[key] = {
                'account_code': c.get('account_code', ''),
                'description': c.get('description', ''),
                'comment_texts': [],
                'source_files': [],
            }
        text = c.get('comment_text', '').strip()
        if text and text not in comments_by_key[key]['comment_texts']:
            comments_by_key[key]['comment_texts'].append(text)
            comments_by_key[key]['source_files'].append(c.get('source_file', ''))

    # Construir lista final con comentarios concatenados
    all_comments = []
    for key, data in comments_by_key.items():
        all_comments.append({
            'account_code': data['account_code'],
            'description': data['description'],
            'comment_text': ' — '.join(data['comment_texts']),
            'source_file': ', '.join(data['source_files']),
        })

    # Ordenar archivos cronológicamente por filename (el último = mes más reciente).
    # Crítico: el YTD se toma del ÚLTIMO archivo (Dec), así que debe ser el más reciente
    # independiente del orden de upload.
    ingested_files = sorted(ingested_files, key=_extract_period_sortkey)

    # Usar último archivo (más reciente) como BASE (preserva la estructura de filas/secciones)
    last_file = ingested_files[-1]
    df = last_file.df.copy()
    has_ytd = last_file.metadata.get('has_ytd', False)

    # --- CONSTRUIR Q: Sumar "Current" (mensual) de TODOS los archivos ---
    # Crear lookup de sumas mensuales por match_key.
    # Guardamos también un "template row" del primer archivo donde apareció la cuenta,
    # para poder rescatar cuentas que existen en M1/M2 pero desaparecen en M3 (último archivo).
    q_sums = {}  # match_key -> {budget_q, actual_q, template_row, files_seen}
    for ingested in ingested_files:
        for _, row in ingested.df.iterrows():
            key = _make_match_key(row)
            if key is None:
                continue
            if key not in q_sums:
                q_sums[key] = {
                    'budget_q': 0.0,
                    'actual_q': 0.0,
                    'template_row': row.copy(),
                    'files_seen': [],
                }
            q_sums[key]['budget_q'] += _safe_float(row.get('budget_current')) or 0.0
            q_sums[key]['actual_q'] += _safe_float(row.get('actual_current')) or 0.0
            q_sums[key]['files_seen'].append(ingested.filename)

    # Reemplazar budget_current/actual_current con la suma Q (trimestral)
    keys_in_df = set()
    for idx, row in df.iterrows():
        key = _make_match_key(row)
        if key:
            keys_in_df.add(key)
        if key and key in q_sums:
            df.at[idx, 'budget_current'] = q_sums[key]['budget_q']
            df.at[idx, 'actual_current'] = q_sums[key]['actual_q']

    # --- FIX bug Yale: rescatar cuentas que aparecen en archivos previos pero NO en el último ---
    # Si una cuenta tiene actividad en M1/M2 pero desaparece en M3, q_sums la tiene sumada
    # pero df (basado en el último archivo) no la incluye. La agregamos al final para no perderla.
    missing_keys = [k for k in q_sums.keys() if k not in keys_in_df]
    ghost_rows = []
    for key in missing_keys:
        sums = q_sums[key]
        if abs(sums['budget_q']) < 0.01 and abs(sums['actual_q']) < 0.01:
            continue
        ghost_row = sums['template_row'].copy()
        ghost_row['budget_current'] = sums['budget_q']
        ghost_row['actual_current'] = sums['actual_q']
        ghost_row['budget_ytd'] = sums['budget_q']
        ghost_row['actual_ytd'] = sums['actual_q']
        ghost_row['_ghost_account'] = True
        ghost_row['_ghost_files_seen'] = ', '.join(sums['files_seen'])
        ghost_rows.append(ghost_row)

    if ghost_rows:
        df = pd.concat([df, pd.DataFrame(ghost_rows)], ignore_index=True)

    # budget_ytd y actual_ytd ya vienen correctos del último archivo (no tocar)

    # --- PASO 1: Identificar secciones y totales ---
    current_section = "Unknown"
    current_l1 = "Unknown"

    sections = []
    for _, row in df.iterrows():
        rtype = row['row_type']

        if rtype == 'section':
            current_section = str(row['description']).strip()
            matched_l1 = _match_section_to_l1(current_section, current_l1)
            if matched_l1 != current_l1:
                current_l1 = matched_l1

        sections.append({
            'section_name': current_section,
            'report_line_l1': current_l1,
        })

    df['section_name'] = [s['section_name'] for s in sections]
    df['report_line_l1'] = [s['report_line_l1'] for s in sections]
    # Source de la clasificación inicial: viene de la herencia de sección.
    # 'unknown' si el L1 quedó en 'Unknown' (no hubo match de sección).
    df['classification_source'] = [
        'section' if l1 != 'Unknown' else 'unknown' for l1 in df['report_line_l1']
    ]

    # --- PASO 1a: Override via per-asset catalog (prioridad máxima para data rows) ---
    # El catálogo por activo (src/asset_catalogs.py) es la fuente autoritativa de
    # clasificación L1 para cuentas listadas en el budget 2026 de Florencia. Convive
    # con las reglas globales: si el (asset, código) está en el catálogo, lo usamos;
    # si no, caemos a la herencia de sección y luego al fallback por rango/keyword.
    asset_key = _resolve_asset_key(ingested_files)
    if asset_key:
        def _apply_catalog_override(row):
            if row.get('row_type') != 'data':
                return row['report_line_l1'], row['classification_source']
            code = str(row.get('account', '')).strip()
            catalog_l1 = get_l1_for_account(asset_key, code)
            if catalog_l1:
                return catalog_l1, 'catalog'
            return row['report_line_l1'], row['classification_source']
        applied = df.apply(lambda r: _apply_catalog_override(r), axis=1)
        df['report_line_l1'] = [x[0] for x in applied]
        df['classification_source'] = [x[1] for x in applied]

    # --- PASO 1b: Fallback — clasificar cuentas 'Unknown' por código o descripción ---
    # Aplica solo a filas de datos donde la sección no matcheó ningún L1.
    # Prioridad: código de cuenta (rangos Yardi) > keywords en descripción.
    def _fallback_classify(row):
        if row['report_line_l1'] != 'Unknown':
            return row['report_line_l1'], row['classification_source']
        if row['row_type'] != 'data':
            return row['report_line_l1'], row['classification_source']
        classified = _classify_by_account_code(str(row.get('account', '')))
        if classified:
            return classified, 'code_range'
        classified = _classify_by_description(str(row.get('description', '')))
        if classified:
            return classified, 'description_keyword'
        return 'Unknown', 'unknown'

    applied = df.apply(lambda r: _fallback_classify(r), axis=1)
    df['report_line_l1'] = [x[0] for x in applied]
    df['classification_source'] = [x[1] for x in applied]

    # --- PASO 2: Extraer líneas principales (L1) desde totales ---
    l1_lines = _extract_l1_lines(df)

    # --- PASO 3: Extraer categorías (L2) desde sub-totales ---
    l2_sections = _extract_l2_sections(df)

    # --- PASO 4: Extraer subcuentas (L3) con datos ---
    l3_accounts = _extract_l3_accounts(df, all_comments)

    # --- PASO 4b: Reclasificar cuentas de "Financing Expense" a Non-Operating ---
    # Dentro de la sección "Interest and Financing Expenses", cuentas como
    # "Financing Expense" (amortización de costos de financiamiento) no son
    # interés puro — el socio las reporta como Non-Operating Expenses.
    for acct in l3_accounts:
        if acct.parent_line == 'Interest Expense':
            desc_lower = acct.name.lower()
            if 'financing' in desc_lower and 'interest' not in desc_lower:
                acct.parent_line = 'Non-Operating Expenses'
                acct.classification_source = 'post_processing'

    # --- PASO 4c: Override universal — Principal Reduction / Principal Payment → NA ---
    # Bug fix 501 Estates (feedback Florencia / Q4 2025 backtest):
    # El partner reporta "Total Debt Service" que SUMA Interest + Principal Reduction.
    # La sección "Debt Service" se mapea a Interest Expense, lo que arrastra los
    # Principal Reduction (que no van al P&L — son balance-sheet movements) hacia
    # Interest, inflando el budget en ~$50k/mes. Forzamos NA por descripción.
    #
    # Esto override la herencia de sección (a diferencia de las reglas keyword en
    # KEYWORD_TO_L1 que solo aplican como fallback).
    PRINCIPAL_PATTERNS = [
        'principal reduction',
        'principal payment',
        'principal repayment',
        'loan principal',
        'mortgage principal',
        'principal due',
        'debt principal',
    ]
    for acct in l3_accounts:
        desc_lower = (acct.name or '').lower()
        # Skip if catalog autoritativo ya decidió (catálogo gana incluso sobre esto)
        if acct.classification_source == 'catalog':
            continue
        for pattern in PRINCIPAL_PATTERNS:
            if pattern in desc_lower:
                acct.parent_line = 'NA'
                acct.classification_source = 'post_processing'
                break

    # Recomputar L1 totales afectados por reclasificación desde suma de L3.
    # Además, crear L1 si no existe pero hay L3 clasificados bajo ese padre (caso Total Capex en 501 Estates).
    l1_map = {l.name: l for l in l1_lines}

    # Cuando hay catálogo resuelto, TODOS los L1 se recalculan desde L3 sum (el
    # catálogo es autoritativo). Cuando no hay catálogo, solo recomputamos los
    # L1 "below NOI" (Interest/Non-Op/Capex) que históricamente se derivaban así.
    #
    # Caso motivador: Walnut/JDE reporta "TOTAL OPERATING EXPENSES" (row 72) que
    # EXCLUYE Non-Recoverable Expenses, pero el catálogo clasifica 57299 como
    # OpEx. Sin este recálculo, L1 OpEx (503k) ≠ L3 sum (544k) y NOI calculado
    # no cuadra con el reportado.
    l1s_to_recompute = ['Interest Expense', 'Non-Operating Expenses', 'Total Capex']
    if asset_key:  # resuelto en PASO 1a
        l1s_to_recompute = ['Income', 'Operating Expenses', 'Real Estate Taxes'] + l1s_to_recompute

    for l1_name in l1s_to_recompute:
        child_l3 = [a for a in l3_accounts if a.parent_line == l1_name]
        if not child_l3:
            continue
        sum_actual_q = sum(a.actual_current for a in child_l3)
        sum_budget_q = sum(a.budget_current for a in child_l3)
        sum_actual_ytd = sum((a.actual_ytd or 0) for a in child_l3)
        sum_budget_ytd = sum((a.budget_ytd or 0) for a in child_l3)
        if l1_name in l1_map:
            l1_obj = l1_map[l1_name]
            l1_obj.actual_current = sum_actual_q if sum_actual_q else 0
            l1_obj.budget_current = sum_budget_q if sum_budget_q else 0
            l1_obj.variance_current = l1_obj.actual_current - l1_obj.budget_current
            l1_obj.variance_pct_current = _pct(l1_obj.actual_current, l1_obj.budget_current)
            l1_obj.actual_ytd = sum_actual_ytd if sum_actual_ytd else None
            l1_obj.budget_ytd = sum_budget_ytd if sum_budget_ytd else None
            l1_obj.variance_ytd = ((l1_obj.actual_ytd or 0) - (l1_obj.budget_ytd or 0)
                                   if l1_obj.actual_ytd is not None else None)
        else:
            # Crear L1 desde suma L3 (cuando el socio no tiene una fila total para este L1)
            new_l1 = FinancialLine(
                name=l1_name, level=1,
                budget_current=sum_budget_q, actual_current=sum_actual_q,
                variance_current=sum_actual_q - sum_budget_q,
                variance_pct_current=_pct(sum_actual_q, sum_budget_q),
                budget_ytd=sum_budget_ytd or None,
                actual_ytd=sum_actual_ytd or None,
                variance_ytd=(sum_actual_ytd - sum_budget_ytd) if sum_actual_ytd else None,
                variance_pct_ytd=_pct(sum_actual_ytd, sum_budget_ytd) if sum_actual_ytd else None,
                parent_line=l1_name,
            )
            l1_lines.append(new_l1)
            l1_map[l1_name] = new_l1

    # --- PASO 5: Vincular comentarios a líneas ---
    _link_comments(l3_accounts, all_comments)

    structure = {
        'l1_lines': l1_lines,
        'l2_sections': l2_sections,
        'l3_accounts': l3_accounts,
        'all_comments': all_comments,
        'raw_consolidated': df,
    }

    # --- PASO 6: NORMALIZACIÓN RET ─ RE Taxes SIEMPRE fuera de OpEx ---
    # Si el socio reportó OpEx con RET incluido (Oakland/Yale legacy), lo desagregamos
    # para consistencia en todo el output. RET queda como L1 independiente.
    _normalize_ret_split(structure)

    # --- PASO 7: Recalcular Net Income con L1 ya ajustados ---
    # Interest/Non-Op cambiaron en PASO 4b (Financing → Non-Op) y OpEx en PASO 6 (si aplicó).
    # Net Income = NOI − |Interest| − |Non-Operating|
    _recompute_net_income(structure)

    return structure


def _recompute_net_income(structure: Dict) -> None:
    """Recalcula Net Income L1 usando NOI − |Interest| − |Non-Op| con los valores finales."""
    l1_map = {l.name: l for l in structure.get('l1_lines', [])}
    noi = l1_map.get('NOI')
    ni = l1_map.get('Net Income')
    if not noi or not ni:
        return
    interest = l1_map.get('Interest Expense')
    non_op = l1_map.get('Non-Operating Expenses')
    ni.actual_current = noi.actual_current \
        - (abs(interest.actual_current) if interest else 0) \
        - (abs(non_op.actual_current) if non_op else 0)
    ni.budget_current = noi.budget_current \
        - (abs(interest.budget_current) if interest else 0) \
        - (abs(non_op.budget_current) if non_op else 0)
    ni.variance_current = ni.actual_current - ni.budget_current
    ni.variance_pct_current = _pct(ni.actual_current, ni.budget_current)
    if noi.actual_ytd is not None:
        ni.actual_ytd = noi.actual_ytd \
            - (abs(interest.actual_ytd or 0) if interest else 0) \
            - (abs(non_op.actual_ytd or 0) if non_op else 0)
    if noi.budget_ytd is not None:
        ni.budget_ytd = noi.budget_ytd \
            - (abs(interest.budget_ytd or 0) if interest else 0) \
            - (abs(non_op.budget_ytd or 0) if non_op else 0)
    if ni.actual_ytd is not None and ni.budget_ytd is not None:
        ni.variance_ytd = ni.actual_ytd - ni.budget_ytd
        ni.variance_pct_ytd = _pct(ni.actual_ytd, ni.budget_ytd)


def _normalize_ret_split(structure: Dict) -> None:
    """
    Normaliza la estructura: Real Estate Taxes SIEMPRE queda fuera del total de Operating Expenses.

    Detecta formato Oakland/legacy (OpEx partner total incluye RET) comparando:
      - diff_con_ret = |OpEx_partner − (OpEx_L3 + RET_L3)|
      - diff_sin_ret = |OpEx_partner − OpEx_L3|
    Si diff_con_ret < diff_sin_ret → OpEx incluye RET → desagregar.

    Si formato Val/CMC (RET ya separada), no toca nada.
    Muta el dict `structure` in-place.
    """
    l1_map = {l.name: l for l in structure.get('l1_lines', [])}
    opex = l1_map.get('Operating Expenses')
    taxes = l1_map.get('Real Estate Taxes')

    if not opex or not taxes or not (taxes.actual_current or 0):
        structure['ret_normalized'] = False
        structure['ret_was_inside_opex'] = False
        return

    l3_accounts = structure.get('l3_accounts', [])
    opex_l3 = sum(a.actual_current or 0 for a in l3_accounts if a.parent_line == 'Operating Expenses')
    ret_l3 = sum(a.actual_current or 0 for a in l3_accounts if a.parent_line == 'Real Estate Taxes')

    opex_abs = abs(opex.actual_current)
    diff_with_ret = abs(opex_abs - abs(opex_l3 + ret_l3))
    diff_without_ret = abs(opex_abs - abs(opex_l3))

    structure['ret_was_inside_opex'] = diff_with_ret < diff_without_ret

    if not structure['ret_was_inside_opex']:
        # Val/CMC style: ya viene normalizado
        structure['ret_normalized'] = False
        return

    # Oakland/Yale style: restar RET del OpEx partner total
    # Convención: ambos valores guardados como magnitudes con mismo signo (positivos ambos).
    opex.actual_current = opex.actual_current - taxes.actual_current
    opex.budget_current = opex.budget_current - (taxes.budget_current or 0)
    if opex.actual_ytd is not None and taxes.actual_ytd is not None:
        opex.actual_ytd = opex.actual_ytd - taxes.actual_ytd
    if opex.budget_ytd is not None and taxes.budget_ytd is not None:
        opex.budget_ytd = opex.budget_ytd - taxes.budget_ytd
    opex.variance_current = opex.actual_current - opex.budget_current
    opex.variance_pct_current = _pct(opex.actual_current, opex.budget_current)
    if opex.actual_ytd is not None and opex.budget_ytd is not None:
        opex.variance_ytd = opex.actual_ytd - opex.budget_ytd
        opex.variance_pct_ytd = _pct(opex.actual_ytd, opex.budget_ytd)
    structure['ret_normalized'] = True


def _extract_l1_lines(df: pd.DataFrame) -> List[FinancialLine]:
    """Extrae las líneas principales del reporte desde filas de totales."""
    lines = {}

    # Buscar filas tipo 'total' o líneas clave como NOI
    for _, row in df.iterrows():
        desc = str(row['description']).strip()
        desc_lower = desc.lower()
        rtype = row.get('row_type', '')

        # 1) Yardi Variance (Yale): subtotales cuya descripción es EXACTAMENTE el nombre L1
        #    (ej. "Income", "Operating Expenses"). Solo válido si row_type = 'total'.
        if rtype == 'total' and desc_lower in EXACT_L1_TOTAL_NAMES:
            line_name = EXACT_L1_TOTAL_NAMES[desc_lower]
            if line_name not in lines:
                b_c = _safe_float(row.get('budget_current'))
                a_c = _safe_float(row.get('actual_current'))
                b_y = _safe_float(row.get('budget_ytd'))
                a_y = _safe_float(row.get('actual_ytd'))
                if b_c is not None or a_c is not None:
                    lines[line_name] = FinancialLine(
                        name=line_name, level=1,
                        budget_current=b_c or 0, actual_current=a_c or 0,
                        variance_current=(a_c or 0) - (b_c or 0),
                        variance_pct_current=_pct(a_c, b_c),
                        budget_ytd=b_y, actual_ytd=a_y,
                        variance_ytd=(a_y or 0) - (b_y or 0) if a_y is not None else None,
                        variance_pct_ytd=_pct(a_y, b_y),
                        parent_line=line_name,
                    )
            continue

        # 2) Mapeo por substring (formato Alice/Edson/Val: "Total Operating Expenses", etc.)
        for pattern, line_name, sign in REPORT_LINE_RULES:
            if pattern in desc_lower and not line_name.startswith('_'):
                # Val: "TOTAL NOI Before RE Tax Depr Debt" es NOI pre-impuestos,
                # no el NOI real. El NOI verdadero viene después en "Net Operating Income".
                if line_name == 'NOI' and 'before' in desc_lower:
                    break
                if line_name not in lines:
                    b_c = _safe_float(row.get('budget_current'))
                    a_c = _safe_float(row.get('actual_current'))
                    b_y = _safe_float(row.get('budget_ytd'))
                    a_y = _safe_float(row.get('actual_ytd'))

                    if b_c is not None or a_c is not None:
                        lines[line_name] = FinancialLine(
                            name=line_name,
                            level=1,
                            budget_current=b_c or 0,
                            actual_current=a_c or 0,
                            variance_current=(a_c or 0) - (b_c or 0),
                            variance_pct_current=_pct(a_c, b_c),
                            budget_ytd=b_y,
                            actual_ytd=a_y,
                            variance_ytd=(a_y or 0) - (b_y or 0) if a_y is not None else None,
                            variance_pct_ytd=_pct(a_y, b_y),
                            parent_line=line_name,
                        )
                break

    # Calcular líneas derivadas si no existen
    if 'Income' in lines and 'Operating Expenses' in lines:
        inc = lines['Income']
        opex = lines['Operating Expenses']
        taxes = lines.get('Real Estate Taxes')

        if 'NOI' not in lines:
            noi_b = inc.budget_current - abs(opex.budget_current) - (abs(taxes.budget_current) if taxes else 0)
            noi_a = inc.actual_current - abs(opex.actual_current) - (abs(taxes.actual_current) if taxes else 0)
            lines['NOI'] = FinancialLine(
                name='NOI', level=1,
                budget_current=noi_b, actual_current=noi_a,
                variance_current=noi_a - noi_b,
                variance_pct_current=_pct(noi_a, noi_b),
                budget_ytd=None, actual_ytd=None,
                variance_ytd=None, variance_pct_ytd=None,
                parent_line='NOI',
            )

    # Calcular Net Income si no está reportado explícitamente
    if 'Net Income' not in lines and 'NOI' in lines:
        noi = lines['NOI']
        interest = lines.get('Interest Expense')
        non_op = lines.get('Non-Operating Expenses')
        ni_b = noi.budget_current - (abs(interest.budget_current) if interest else 0) - (abs(non_op.budget_current) if non_op else 0)
        ni_a = noi.actual_current - (abs(interest.actual_current) if interest else 0) - (abs(non_op.actual_current) if non_op else 0)
        ni_b_ytd = (noi.budget_ytd - (abs(interest.budget_ytd or 0) if interest else 0) - (abs(non_op.budget_ytd or 0) if non_op else 0)) if noi.budget_ytd is not None else None
        ni_a_ytd = (noi.actual_ytd - (abs(interest.actual_ytd or 0) if interest else 0) - (abs(non_op.actual_ytd or 0) if non_op else 0)) if noi.actual_ytd is not None else None
        lines['Net Income'] = FinancialLine(
            name='Net Income', level=1,
            budget_current=ni_b, actual_current=ni_a,
            variance_current=ni_a - ni_b,
            variance_pct_current=_pct(ni_a, ni_b),
            budget_ytd=ni_b_ytd, actual_ytd=ni_a_ytd,
            variance_ytd=(ni_a_ytd - ni_b_ytd) if ni_a_ytd is not None and ni_b_ytd is not None else None,
            variance_pct_ytd=_pct(ni_a_ytd, ni_b_ytd) if ni_a_ytd is not None else None,
            parent_line='Net Income',
        )

    # Orden canónico (Interest Expense después de NOI, separado de OPEX)
    order = ['Income', 'Operating Expenses', 'Real Estate Taxes', 'NOI',
             'Interest Expense', 'Non-Operating Expenses', 'Net Income']
    result = []
    for name in order:
        if name in lines:
            result.append(lines[name])
    # Agregar cualquier otra que no esté en el orden
    for name, line in lines.items():
        if name not in order:
            result.append(line)
    return result


def _extract_l2_sections(df: pd.DataFrame) -> List[FinancialLine]:
    """Extrae categorías de nivel 2 (sub-totales como Payroll, Utilities, etc.)."""
    sections = []
    for _, row in df.iterrows():
        if row['row_type'] != 'total':
            continue
        desc = str(row['description']).strip()
        desc_lower = desc.lower()

        # Solo sub-totales, no totales principales
        is_subtotal = False
        parent_l1 = None
        for pattern, line_name, sign in REPORT_LINE_RULES:
            if pattern in desc_lower:
                if line_name.startswith('_sub_'):
                    is_subtotal = True
                    parent_l1 = line_name.replace('_sub_', '').title()
                    if parent_l1 == 'Opex':
                        parent_l1 = 'Operating Expenses'
                    elif parent_l1 == 'Taxes':
                        parent_l1 = 'Real Estate Taxes'
                break

        if is_subtotal:
            b_c = _safe_float(row.get('budget_current'))
            a_c = _safe_float(row.get('actual_current'))
            b_y = _safe_float(row.get('budget_ytd'))
            a_y = _safe_float(row.get('actual_ytd'))

            sections.append(FinancialLine(
                name=desc.replace('Total ', ''),
                level=2,
                budget_current=b_c or 0,
                actual_current=a_c or 0,
                variance_current=(a_c or 0) - (b_c or 0),
                variance_pct_current=_pct(a_c, b_c),
                budget_ytd=b_y,
                actual_ytd=a_y,
                variance_ytd=(a_y or 0) - (b_y or 0) if a_y is not None else None,
                variance_pct_ytd=_pct(a_y, b_y),
                parent_line=parent_l1 or 'Unknown',
            ))

    # Ordenar por magnitud de variance
    sections.sort(key=lambda x: abs(x.variance_current), reverse=True)
    return sections


def _extract_l3_accounts(df: pd.DataFrame, comments: list) -> List[FinancialLine]:
    """Extrae subcuentas individuales con datos."""
    accounts = []
    comment_map = {c['account_code']: c['comment_text'] for c in comments if c['account_code']}

    for _, row in df.iterrows():
        if row['row_type'] != 'data':
            continue
        acct = str(row.get('account', '')).strip()
        if not acct or acct == 'nan':
            continue

        b_c = _safe_float(row.get('budget_current'))
        a_c = _safe_float(row.get('actual_current'))
        b_y = _safe_float(row.get('budget_ytd'))
        a_y = _safe_float(row.get('actual_ytd'))

        var_c = (a_c or 0) - (b_c or 0)

        # CAMBIO: ya no se eliminan cuentas con budget y actual en cero.
        # Todas las cuentas aparecen en los reportes, incluso si la variación es nula.

        # row_num del partner Excel (preserva orden natural del P&L)
        row_num = row.get('row_num')
        try:
            row_num = int(row_num) if row_num is not None else 0
        except (TypeError, ValueError):
            row_num = 0

        accounts.append(FinancialLine(
            name=str(row['description']).strip(),
            level=3,
            budget_current=b_c or 0,
            actual_current=a_c or 0,
            variance_current=var_c,
            variance_pct_current=_pct(a_c, b_c),
            budget_ytd=b_y,
            actual_ytd=a_y,
            variance_ytd=(a_y or 0) - (b_y or 0) if a_y is not None else None,
            variance_pct_ytd=_pct(a_y, b_y),
            parent_line=row.get('report_line_l1', 'Unknown'),
            account_code=acct,
            partner_comment=comment_map.get(acct, ''),
            section_name=str(row.get('section_name', '')).strip(),
            classification_source=row.get('classification_source', 'unknown'),
            partner_row_order=row_num,
        ))

    # Ordenar por magnitud de variance
    accounts.sort(key=lambda x: abs(x.variance_current), reverse=True)
    return accounts


def _link_comments(accounts: List[FinancialLine], comments: list):
    """Vincula comentarios a subcuentas por código de cuenta."""
    comment_map = {}
    for c in comments:
        key = c.get('account_code', '')
        if key:
            comment_map[key] = c['comment_text']

    for acct in accounts:
        if acct.account_code in comment_map and not acct.partner_comment:
            acct.partner_comment = comment_map[acct.account_code]


def consolidate_structures(structures: List[Dict], portfolio_name: str = "Portfolio") -> Dict:
    """
    Consolida múltiples estructuras financieras (activos) en una sola estructura
    de portafolio. Suma valores numéricos, concatena comentarios.

    Args:
        structures: Lista de outputs de build_financial_structure(), uno por activo.
        portfolio_name: Nombre del portafolio consolidado.

    Returns:
        Dict consolidado con la misma estructura (l1_lines, l2_sections, l3_accounts, etc.)
    """
    if not structures:
        raise ValueError("No hay estructuras para consolidar")
    if len(structures) == 1:
        return structures[0]

    # --- Consolidar L1 lines ---
    l1_map = {}  # name -> accumulated FinancialLine
    for struct in structures:
        for line in struct.get('l1_lines', []):
            if line.name not in l1_map:
                l1_map[line.name] = FinancialLine(
                    name=line.name, level=1,
                    budget_current=0, actual_current=0,
                    variance_current=0, variance_pct_current=None,
                    budget_ytd=0 if line.budget_ytd is not None else None,
                    actual_ytd=0 if line.actual_ytd is not None else None,
                    variance_ytd=0 if line.variance_ytd is not None else None,
                    variance_pct_ytd=None,
                    parent_line=line.name,
                )
            acc = l1_map[line.name]
            acc.budget_current += line.budget_current or 0
            acc.actual_current += line.actual_current or 0
            acc.variance_current += line.variance_current or 0
            if line.budget_ytd is not None and acc.budget_ytd is not None:
                acc.budget_ytd += line.budget_ytd or 0
            if line.actual_ytd is not None and acc.actual_ytd is not None:
                acc.actual_ytd += line.actual_ytd or 0
            if line.variance_ytd is not None and acc.variance_ytd is not None:
                acc.variance_ytd += line.variance_ytd or 0

    # Recalcular porcentajes
    for line in l1_map.values():
        line.variance_pct_current = _pct(line.actual_current, line.budget_current)
        if line.actual_ytd is not None:
            line.variance_pct_ytd = _pct(line.actual_ytd, line.budget_ytd)

    # Calcular Net Income consolidado si no está presente
    if 'Net Income' not in l1_map and 'NOI' in l1_map:
        noi = l1_map['NOI']
        interest = l1_map.get('Interest Expense')
        non_op = l1_map.get('Non-Operating Expenses')
        ni_b = noi.budget_current - (abs(interest.budget_current) if interest else 0) - (abs(non_op.budget_current) if non_op else 0)
        ni_a = noi.actual_current - (abs(interest.actual_current) if interest else 0) - (abs(non_op.actual_current) if non_op else 0)
        ni_b_ytd = (noi.budget_ytd - (abs(interest.budget_ytd or 0) if interest else 0) - (abs(non_op.budget_ytd or 0) if non_op else 0)) if noi.budget_ytd is not None else None
        ni_a_ytd = (noi.actual_ytd - (abs(interest.actual_ytd or 0) if interest else 0) - (abs(non_op.actual_ytd or 0) if non_op else 0)) if noi.actual_ytd is not None else None
        l1_map['Net Income'] = FinancialLine(
            name='Net Income', level=1,
            budget_current=ni_b, actual_current=ni_a,
            variance_current=ni_a - ni_b,
            variance_pct_current=_pct(ni_a, ni_b),
            budget_ytd=ni_b_ytd, actual_ytd=ni_a_ytd,
            variance_ytd=(ni_a_ytd - ni_b_ytd) if ni_a_ytd is not None and ni_b_ytd is not None else None,
            variance_pct_ytd=_pct(ni_a_ytd, ni_b_ytd) if ni_a_ytd is not None else None,
            parent_line='Net Income',
        )

    # Ordenar L1
    order = ['Income', 'Operating Expenses', 'Real Estate Taxes', 'NOI',
             'Interest Expense', 'Non-Operating Expenses', 'Net Income', 'Total Capex']
    l1_lines = []
    for name in order:
        if name in l1_map:
            l1_lines.append(l1_map[name])
    for name, line in l1_map.items():
        if name not in order:
            l1_lines.append(line)

    # --- Consolidar L3 accounts (prefixed by asset) ---
    all_l3 = []
    for struct in structures:
        for acct in struct.get('l3_accounts', []):
            all_l3.append(acct)
    all_l3.sort(key=lambda x: abs(x.variance_current), reverse=True)

    # --- Consolidar L2 sections ---
    l2_map = {}
    for struct in structures:
        for sec in struct.get('l2_sections', []):
            key = sec.name.lower().strip()
            if key not in l2_map:
                l2_map[key] = FinancialLine(
                    name=sec.name, level=2,
                    budget_current=0, actual_current=0,
                    variance_current=0, variance_pct_current=None,
                    budget_ytd=0 if sec.budget_ytd is not None else None,
                    actual_ytd=0 if sec.actual_ytd is not None else None,
                    variance_ytd=0 if sec.variance_ytd is not None else None,
                    variance_pct_ytd=None,
                    parent_line=sec.parent_line,
                )
            acc = l2_map[key]
            acc.budget_current += sec.budget_current or 0
            acc.actual_current += sec.actual_current or 0
            acc.variance_current += sec.variance_current or 0
            if sec.budget_ytd is not None and acc.budget_ytd is not None:
                acc.budget_ytd += sec.budget_ytd or 0
            if sec.actual_ytd is not None and acc.actual_ytd is not None:
                acc.actual_ytd += sec.actual_ytd or 0
            if sec.variance_ytd is not None and acc.variance_ytd is not None:
                acc.variance_ytd += sec.variance_ytd or 0

    for sec in l2_map.values():
        sec.variance_pct_current = _pct(sec.actual_current, sec.budget_current)
        if sec.actual_ytd is not None:
            sec.variance_pct_ytd = _pct(sec.actual_ytd, sec.budget_ytd)

    l2_sections = sorted(l2_map.values(), key=lambda x: abs(x.variance_current), reverse=True)

    # --- Consolidar comentarios ---
    all_comments = []
    for struct in structures:
        all_comments.extend(struct.get('all_comments', []))

    # --- Consolidar raw data ---
    raw_frames = [struct.get('raw_consolidated') for struct in structures
                  if struct.get('raw_consolidated') is not None]
    raw_consolidated = pd.concat(raw_frames, ignore_index=True) if raw_frames else pd.DataFrame()

    return {
        'l1_lines': l1_lines,
        'l2_sections': l2_sections,
        'l3_accounts': all_l3,
        'all_comments': all_comments,
        'raw_consolidated': raw_consolidated,
        'consolidated_from': len(structures),
    }


def _make_match_key(row) -> Optional[str]:
    """
    Crea una clave de matching para alinear filas entre archivos mensuales.
    - data rows: usa account code (ej: "4510-0000")
    - total rows: usa description normalizada (ej: "total revenues_total")
    - section/label rows: retorna None (no se suman)
    """
    row_type = str(row.get('row_type', ''))
    if row_type in ('section', 'label'):
        return None

    acct = str(row.get('account', '')).strip()
    desc = str(row.get('description', '')).strip().lower()

    if row_type == 'data' and acct and acct != 'nan' and acct != '':
        return f"acct_{acct}"
    elif row_type == 'total' and desc:
        return f"total_{desc}"
    elif desc:
        return f"desc_{desc}_{row_type}"
    return None


def _safe_float(val) -> Optional[float]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _pct(actual, budget) -> Optional[float]:
    a = _safe_float(actual)
    b = _safe_float(budget)
    if a is None or b is None or abs(b) < 0.01:
        return None
    return (a - b) / abs(b)
