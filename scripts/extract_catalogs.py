"""
One-shot: parsea los 3 Excel de budget 2026 (Alice, Edson, Val) y emite
los catálogos por activo en forma de diccionario Python listo para pegar
en src/asset_catalogs.py.

Uso:
    python scripts/extract_catalogs.py > /tmp/catalogs.py

Reglas de derivación:
- L2 = header de sección del Excel (el más reciente arriba del account).
- L1 = mapea L2 via SECTION_TO_L1 (subset). Si la L2 no está en el subset,
  se intenta substring; si no, queda None y se loggea para revisión manual.
- Overrides account-level para casos donde el Excel agrupa cuentas en una
  misma sección pero el L1 difiere (ej: Alice/Edson sección "Interest and
  Financing Expenses" tiene 9010-0000=Interest, 9010-0010=Financing→Non-Op).
"""

from __future__ import annotations
import re
import sys
from pathlib import Path

import openpyxl
from pyxlsb import open_workbook as open_xlsb

# ── Fuentes ────────────────────────────────────────────────────────────
DOWNLOADS = Path(r"C:/Users/JaimeValenzuela/Downloads")
ALICE = DOWNLOADS / "20260211 -- Alice House Cash Position - Budget v2_ FINAL stars format.xlsx"
EDSON = DOWNLOADS / "02.11.25 Edson House Budget 2026 TMG FINAL v2-tmg format.xlsb"
VAL   = DOWNLOADS / "Val Budget.xlsx"
# Walnut Street Wellesley (Lincoln Property Company, JDE format).
# Usamos el archivo de marzo como referencia del chart of accounts — los tres meses
# comparten el mismo plan de cuentas.
WALNUT = DOWNLOADS / "Budget Comparison Propiedad 03-2026.xlsx"
# The Wilcox (YSI/Yardi Budget Forecast format, tree=ysi_bf).
# Usamos el archivo de diciembre (YTD completo) para el chart of accounts.
WILCOX = DOWNLOADS / "The Wilcox Budget Comparison with Notes 12.2025.xlsx"

# ── Mapeo L2 → L1 (derivado de SECTION_TO_L1 en native_structure.py) ───
# Subset focalizado en lo que aparece en estos 3 archivos.
L2_TO_L1 = {
    # Income
    "revenues": "Income",
    "rental income": "Income",
    "concessions and vacancy": "Income",
    "other income": "Income",
    "income": "Income",
    "rental": "Income",
    "vacancy": "Income",
    "service income": "Income",
    "financial income": "Income",
    # OpEx
    "expenses": "Operating Expenses",
    "payroll and related": "Operating Expenses",
    "utilities": "Operating Expenses",
    "turnover": "Operating Expenses",
    "repairs and maintenance": "Operating Expenses",
    "contract services": "Operating Expenses",
    "advertising and marketing": "Operating Expenses",
    "management fees": "Operating Expenses",
    "property g & a": "Operating Expenses",
    "property g&a": "Operating Expenses",
    "renting": "Operating Expenses",
    "administrative": "Operating Expenses",
    "payroll expense": "Operating Expenses",
    "operating expense": "Operating Expenses",
    "utility expense": "Operating Expenses",
    "maintenance": "Operating Expenses",
    "noncontrollable": "Operating Expenses",
    "insurance": "Operating Expenses",
    # Real Estate Taxes
    "property taxes": "Real Estate Taxes",
    "taxes": "Real Estate Taxes",
    # Interest (por defecto; Financing se ajusta via override)
    "interest and financing expenses": "Interest Expense",
    "financial": "Interest Expense",
    # Non-Operating
    "company general & administrative": "Non-Operating Expenses",
    "unrealized gain/loss": "Non-Operating Expenses",
    "non operating expenses": "Non-Operating Expenses",
    "depreciation / amortization": "Non-Operating Expenses",
    # CapEx
    "capex": "Total Capex",
    "capital repairs & replacements": "Total Capex",
    "capital repairs and replacements": "Total Capex",
    "capital expenses": "Total Capex",
    # Debt Service (Edson Budget sheet)
    "debt service": "Interest Expense",
    # JDE / Lincoln Property Company (Walnut)
    "revenues:": "Income",
    "rental income": "Income",
    "other income": "Income",
    "expenses:": "Operating Expenses",
    "operating expenses": "Operating Expenses",
    "taxes & insurance": "Operating Expenses",  # split via description override (RE Tax -> RET)
    "administration": "Operating Expenses",
    "utilities": "Operating Expenses",
    "repairs & maintenance": "Operating Expenses",
    "management fee": "Operating Expenses",
    "payroll": "Operating Expenses",
    "non-recoverable expenses": "Operating Expenses",  # en JDE están arriba de NOI
    "non operating expenses": "Non-Operating Expenses",
    # YSI / Yardi Budget Forecast (Wilcox)
    "revenue": "Income",
    "cost recovery": "Income",
    "potential rent": "Income",
    "payroll": "Operating Expenses",
    "administrative salaries": "Operating Expenses",
    "leasing salaries": "Operating Expenses",
    "maintenance salaries": "Operating Expenses",
    "bonus": "Operating Expenses",
    "payroll taxes": "Operating Expenses",
    "insurance benefits": "Operating Expenses",
    "401k contribution": "Operating Expenses",
    "workers compensation": "Operating Expenses",
    "other payroll": "Operating Expenses",
    "redecorating/ make-ready": "Operating Expenses",
    "redecorating/make-ready": "Operating Expenses",
    "landscaping/ contracted srvs": "Operating Expenses",
    "landscaping/contracted srvs": "Operating Expenses",
    "leasing & marketing": "Operating Expenses",
    "general & administrative": "Operating Expenses",
    "bad debt expense": "Operating Expenses",
    "real estate & property taxes": "Real Estate Taxes",
    "interest-mrtg encumbrance": "Interest Expense",
    "professional/ partnership exp": "Non-Operating Expenses",
    "professional/partnership exp": "Non-Operating Expenses",
    "depreciation/ amortization": "Non-Operating Expenses",
    "depreciation/amortization": "Non-Operating Expenses",
    "renovation expenses": "Total Capex",
    "capitalized expenditures": "Total Capex",
}


def l2_to_l1(l2_raw: str) -> str | None:
    if not l2_raw:
        return None
    l2 = l2_raw.lower().strip()
    if l2 in L2_TO_L1:
        return L2_TO_L1[l2]
    # substring fallback
    for key, l1 in L2_TO_L1.items():
        if len(key) >= 5 and key in l2:
            return l1
    return None


# ── Overrides account-level (cuando el L1 difiere del section L1) ──────
ACCOUNT_L1_OVERRIDES = {
    # Alice/Edson: dentro de "Interest and Financing Expenses" separamos:
    # 9010-0000 = Interest Expense puro; 9010-0010 = Financing (amort costos) -> Non-Op
    ("Alice House JV LLC", "9010-0010"): "Non-Operating Expenses",
    ("295 29th Street JV LLC", "9010-0010"): "Non-Operating Expenses",
    # Walnut: sección "NON OPERATING EXPENSES" mezcla Asset Mgmt Fee (Non-Op) con
    # Mortgage Interest (Interest Expense). Split por código.
    ("Walnut Street Wellesley", "82210"): "Interest Expense",
    # 80110 Asset Mgmt Fee-External queda como Non-Operating (default del section)

    # ── THE VAL (HDFC/CMC) — fixes post-backtest Q4 2025 de Florencia ─────
    # Las 9 cuentas de "Replacements" (6543-xxx) quedaron mal bajo "Taxes"
    # porque el Excel del Val no trae un section header "Replacements" explícito:
    # los 6543 viven huérfanos entre "Taxes" (row 191) y "Subtotal Replacements"
    # (row 208). Mi extractor heredó "Taxes" → RET. En realidad son ITEMS DE
    # CAPITAL (reemplazos de bienes muebles: electrodomésticos, carpet, pisos,
    # etc.). Florencia los quiere fuera del NI.
    ("The Val", "6543-001"): "Total Capex",  # Appliances
    ("The Val", "6543-002"): "Total Capex",  # Carpet
    ("The Val", "6543-003"): "Total Capex",  # Flooring
    ("The Val", "6543-004"): "Total Capex",  # Paving
    ("The Val", "6543-005"): "Total Capex",  # Redecorating
    ("The Val", "6543-006"): "Total Capex",  # Other
    ("The Val", "6543-007"): "Total Capex",  # Misc. Apt.
    ("The Val", "6543-009"): "Total Capex",  # Reasonable Accommodations
    ("The Val", "6543-010"): "Total Capex",  # Regulatory Inspections

    # 4 cuentas Non-Op que quedaron como Interest porque el Excel del Val pone
    # el header "Non Operating Expenses" en COL C (row 233 col 3), no col B.
    # Mi extractor solo lee col B para sections, por eso las cuentas heredaron
    # la sección previa "Financial" → Interest Expense. Fix por override.
    ("The Val", "6890-000"): "Non-Operating Expenses",  # ENTITY EXPENSE (corporate)
    ("The Val", "6903-000"): "Non-Operating Expenses",  # COVID Expenses
    ("The Val", "6904-000"): "Non-Operating Expenses",  # Non Operating Other

    # 6850-000 MORTGAGE INSURANCE PREMIUM: vive bajo sección "Financial" pero
    # NO es interés. Es un costo de financiamiento no-interés → Non-Op
    # (consistente con la regla de pilot: "some financing costs are non-
    # operating, not interest").
    ("The Val", "6850-000"): "Non-Operating Expenses",
}


# Overrides por keyword en descripcion (se aplica DESPUES del section mapping).
# Util cuando una seccion agrupa cuentas con L1 distintos (Edson "Taxes & Insurance",
# "Debt Service" con principal pmts, etc.).
#
# Orden: mas especifico primero, primer match gana.
DESCRIPTION_L1_OVERRIDES = [
    # Real Estate Taxes (separar dentro de "Taxes & Insurance" de Edson)
    ("real property tax", "Real Estate Taxes"),
    ("real estate tax", "Real Estate Taxes"),
    ("cfd area tax", "Real Estate Taxes"),
    ("personal property tax", "Real Estate Taxes"),
    ("property tax consulting", "Non-Operating Expenses"),  # no es un tax, es consulting
    ("property tax", "Real Estate Taxes"),
    # Principal pmts en "Debt Service" no son P&L (son cashflow) — excluir
    ("principal-first trust", None),
    ("principal-second trust", None),
    ("principal-third trust", None),
    ("principal payment", None),
    ("2nd principal", None),
    ("aitd payable", None),
    # Unrealized gain/loss forzar Non-Op (por si cae fuera de seccion)
    ("unrealized gain", "Non-Operating Expenses"),
    ("unrealized loss", "Non-Operating Expenses"),
]


def apply_description_override(desc: str, current_l1: str | None) -> str | None:
    if not desc:
        return current_l1
    lower = desc.lower().strip()
    for kw, target in DESCRIPTION_L1_OVERRIDES:
        if kw in lower:
            return target  # may be None (signals "exclude")
    return current_l1


def _is_code(s) -> bool:
    if s is None:
        return False
    s = str(s).strip()
    return bool(re.match(r"^\d{3,4}[\.\-]", s))


def _extract_rows(rows_iter, acct_col: int, desc_col: int):
    """
    Walks rows. Tracks current section header (row with no account code but
    a description in desc_col). Emits (account_code, description, section).
    """
    current_section = ""
    for vals in rows_iter:
        # vals must be indexable; pad if short
        if vals is None:
            continue
        acct = vals[acct_col] if len(vals) > acct_col else None
        desc = vals[desc_col] if len(vals) > desc_col else None

        desc_str = str(desc).strip() if desc is not None else ""

        if _is_code(acct):
            yield str(acct).strip(), desc_str, current_section
        else:
            # Section header detection: row has a descriptive string in desc_col
            # and no account code. Skip known non-section labels.
            if desc_str and not desc_str.lower().startswith(("total ", "subtotal", "net ", "potential ")):
                # Filter out summary / non-section labels
                lower = desc_str.lower()
                noise = (
                    "allowances/vacancies", "net rental income", "physical occupancy",
                    "economic occupancy", "move-ins", "renewals", "other rental income",
                    "bad debt write off", "total apartment rent", "net operating income",
                    "total operating income", "noi before debt service", "noi after debt",
                    "total distributable", "net cash", "profit (loss)", "total cash",
                    "number of turnovers", "total unrealized",
                )
                if any(n in lower for n in noise):
                    continue
                current_section = desc_str


def parse_xlsx_rows(path: Path, sheet: str):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(list(row) if row else [])
    wb.close()
    return rows


def parse_xlsb_rows(path: Path, sheet: str):
    rows = []
    with open_xlsb(str(path)) as wb:
        with wb.get_sheet(sheet) as sh:
            for row in sh.rows():
                rows.append([c.v for c in row])
    return rows


def extract_alice():
    rows = parse_xlsx_rows(ALICE, "2026 Trended")
    # Format: col A=None, col B=code, col C=desc, col D=actual, col E=total
    # acct_col=1 (B), desc_col=2 (C)
    entries = []
    for acct, desc, section in _extract_rows(rows, acct_col=1, desc_col=2):
        entries.append((acct, desc, section))
    return entries


def extract_edson():
    rows = parse_xlsb_rows(EDSON, "Budget")
    # Format: col A=JV code, col B=Acct #, col C=desc
    # Use col A (full JV code like "4510-0000") as acct.
    # acct_col=0 (A), desc_col=2 (C)
    entries = []
    for acct, desc, section in _extract_rows(rows, acct_col=0, desc_col=2):
        entries.append((acct, desc, section))
    return entries


def extract_val():
    rows = parse_xlsx_rows(VAL, "STARS Monthly Spread")
    # Format: col A=code (like "5110-000"), col B=desc
    # acct_col=0, desc_col=1
    entries = []
    for acct, desc, section in _extract_rows(rows, acct_col=0, desc_col=1):
        entries.append((acct, desc, section))
    return entries


def extract_wilcox():
    """
    Parser YSI/Yardi Budget Forecast para Wilcox. Las secciones tienen código
    (ej '410000 REVENUE', '600002 PAYROLL'). Las cuentas L3 tienen código +
    descripción indentada + PTD/YTD values. Los totales terminan en 999/099 o
    empiezan con "TOTAL".

    Cortamos al llegar a 'ADJUSTMENTS' o '989999 NET INCOME (LOSS)' — lo que
    sigue es balance sheet, no P&L.
    """
    rows = parse_xlsx_rows(WILCOX, "Report1")
    entries = []
    current_section = ""

    def is_6digit(s):
        if s is None:
            return False
        s = str(s).strip()
        return bool(re.match(r"^\d{6}$", s))

    for vals in rows:
        if not vals:
            continue
        acct = vals[0] if len(vals) > 0 else None
        desc = vals[1] if len(vals) > 1 else None
        actual_c = vals[2] if len(vals) > 2 else None
        budget_c = vals[3] if len(vals) > 3 else None

        acct_str = str(acct).strip() if acct is not None else ""
        desc_str = str(desc).strip() if desc is not None else ""

        # Cutoff: ADJUSTMENTS / balance sheet
        if desc_str.upper() == "ADJUSTMENTS":
            break
        if acct_str == "989999" or desc_str.upper() == "NET INCOME (LOSS)":
            break

        has_vals = isinstance(actual_c, (int, float)) or isinstance(budget_c, (int, float))

        # Section header: tiene código (o no) + desc NO empty + NO valores + NO "TOTAL"
        desc_upper = desc_str.upper()
        is_total_like = desc_upper.startswith("TOTAL ")

        if desc_str and not has_vals and not is_total_like:
            # Ignorar contenedores L1 (REVENUE, OPERATING EXPENSES, NON OPERATING EXPENSES)
            # para que las sub-secciones queden como L2 "real"
            if desc_upper not in ("REVENUE", "OPERATING EXPENSES", "NON OPERATING EXPENSES"):
                current_section = desc_str
            continue

        # Data row: 6-digit code + values + no "TOTAL" en desc
        if is_6digit(acct) and has_vals and not is_total_like:
            # Limpiar indentación de la descripción (L3 tiene leading spaces)
            desc_clean = desc_str.strip()
            entries.append((acct_str, desc_clean, current_section))
            continue

        # Totals: los ignoramos para el catálogo (no son cuentas L3)

    return entries


def extract_walnut():
    """
    Parser especifico JDE para Walnut: col A=codigo 5-digitos, col C=descripcion
    (con seccion embebida via \\n). Secciones independientes en col A (sin codigo)
    o en col C (sin codigo, sin valores).
    """
    rows = parse_xlsx_rows(WALNUT, "Table 1")
    entries = []
    current_section = ""

    def is_jde_code(s):
        if s is None:
            return False
        s = str(s).strip()
        return bool(re.match(r"^\d{5}$", s))

    # JDE section noise (headers que NO son secciones reales sino labels de estructura superior)
    # NO hay noise absoluto — todas las secciones cuentan.
    noise = ("total ", "net operating income", "net income", "net income less capital")

    for vals in rows:
        if not vals:
            continue
        col_a = vals[0] if len(vals) > 0 else None
        col_c = vals[2] if len(vals) > 2 else None
        col_e = vals[4] if len(vals) > 4 else None
        col_f = vals[5] if len(vals) > 5 else None

        a_str = str(col_a).strip() if col_a is not None else ""
        c_str = str(col_c).strip() if col_c is not None else ""

        has_vals = isinstance(col_e, (int, float)) or isinstance(col_f, (int, float))

        # Caso 1: col A tiene texto no-numerico (seccion en col A) - ej "OPERATING EXPENSES", "TAXES & INSURANCE"
        if a_str and not is_jde_code(col_a) and not has_vals:
            if not any(a_str.lower().startswith(n) for n in noise):
                current_section = a_str
            continue

        # Caso 2: col C tiene seccion pura (sin codigo, sin valores) - ej "REVENUES:", "NON OPERATING EXPENSES", "CAPITAL EXPENSES"
        if not a_str and c_str and not has_vals:
            lower = c_str.lower()
            if not any(lower.startswith(n) for n in noise):
                current_section = c_str
            continue

        # Caso 3: fila de datos con codigo
        if is_jde_code(col_a):
            # La descripcion puede ser "SECCION\nDescripcion real" — actualizar seccion
            desc_lines = [ln.strip() for ln in c_str.split("\n") if ln.strip()]
            if len(desc_lines) >= 2:
                # Primera linea es seccion embebida
                embedded = desc_lines[0]
                if not any(embedded.lower().startswith(n) for n in noise):
                    current_section = embedded
                desc_final = desc_lines[1]
            else:
                desc_final = desc_lines[0] if desc_lines else ""
            entries.append((a_str, desc_final, current_section))
            continue

        # Caso 4: fila de total (col C empieza con "Total" o "Net", col A vacia o con codigo residual)
        # — la ignoramos para el catalogo (no son cuentas L3).

    return entries


# Secciones que NO pertenecen al P&L (balance sheet, cashflow, equity).
# Las cuentas debajo de estas secciones se EXCLUYEN del catálogo sin importar
# su descripción.
NON_PNL_SECTIONS = {
    "cash flow adjustments", "contributions/distributions",
    "cash flow conversion", "ending cash balance", "initial net income",
    "adjusted net income", "adjusted net income (new budget)", "cash",
    "working capital",
}


def build_catalog(asset_key: str, entries, plan: str, open_list: bool):
    accounts = {}
    unmatched_sections = set()
    for acct, desc, section in entries:
        # Hard exclude: secciones no-P&L (balance sheet, cashflow adjustments).
        if section and section.lower().strip() in NON_PNL_SECTIONS:
            continue

        # Priority: (asset, code) override > section mapping > description override
        l1 = ACCOUNT_L1_OVERRIDES.get((asset_key, acct))
        if l1 is None:
            l1 = l2_to_l1(section)
        # Apply description-based override (can force RET, or signal exclusion).
        # Excluir recoveries/reimbursements — esas son INGRESOS, no impuestos.
        desc_lower = desc.lower()
        is_recovery = any(kw in desc_lower for kw in
                          ("recov", "reimburs", "recovery", "income"))
        tax_like = ("real property tax" in desc_lower or
                    "real estate tax" in desc_lower or
                    "cfd area tax" in desc_lower or
                    "personal property tax" in desc_lower or
                    "property tax consulting" in desc_lower or
                    "property tax" in desc_lower)
        if (tax_like and not is_recovery) or \
           "principal" in desc_lower or \
           "aitd payable" in desc_lower or \
           "unrealized" in desc_lower:
            l1 = apply_description_override(desc, l1)

        if l1 is None:
            unmatched_sections.add(section)
            # Skip: caller will fall back to globals in runtime
            continue
        # First occurrence wins: preserves original P&L section when the same code
        # reappears in cashflow/adjustment sections below.
        if acct in accounts:
            continue
        accounts[acct] = {
            "desc": desc,
            "l1": l1,
            "l2": section,
        }
    return {
        "plan": plan,
        "open_list": open_list,
        "accounts": accounts,
    }, unmatched_sections


def emit_python(catalogs: dict[str, dict]) -> str:
    out = []
    out.append('"""')
    out.append("Per-asset chart of accounts. Generado desde los Excel de budget 2026 de Florencia.")
    out.append("")
    out.append("Keyed by `asset_key` (= `building` field en ingested files, = keys en ASSET_CATALOG).")
    out.append("Cada entrada mapea account_code → {desc, l1, l2}.")
    out.append("")
    out.append("Planes de cuentas:")
    out.append("- TMG (Alice, Edson): shared chart, códigos 4xxx (income) + 8xxx (opex) + 9xxx (interest).")
    out.append("- HDFC (The Val): códigos 5xxx (income) + 6xxx (opex) + algunos 1xxx/2xxx (cashflow conv).")
    out.append("")
    out.append("Si `open_list=True`, los códigos no listados caen al fallback global de native_structure.")
    out.append("Si `open_list=False`, idem — el catálogo siempre CONVIVE con las reglas globales.")
    out.append("")
    out.append("Generado automáticamente; editar a mano para overrides puntuales.")
    out.append('"""')
    out.append("")
    out.append("ASSET_CATALOGS = {")
    for asset, cat in catalogs.items():
        out.append(f'    {asset!r}: {{')
        out.append(f'        "plan": {cat["plan"]!r},')
        out.append(f'        "open_list": {cat["open_list"]!r},')
        out.append(f'        "accounts": {{')
        for code in sorted(cat["accounts"].keys()):
            info = cat["accounts"][code]
            out.append(f'            {code!r}: {{"desc": {info["desc"]!r}, "l1": {info["l1"]!r}, "l2": {info["l2"]!r}}},')
        out.append('        },')
        out.append('    },')
    out.append("}")
    out.append("")
    out.append("")
    out.append("def get_l1_for_account(asset_key: str, account_code: str) -> str | None:")
    out.append('    """L1 del (asset, code). Retorna None si no hay match."""')
    out.append("    if not asset_key or not account_code:")
    out.append("        return None")
    out.append("    cat = ASSET_CATALOGS.get(asset_key)")
    out.append("    if not cat:")
    out.append("        return None")
    out.append("    entry = cat['accounts'].get(str(account_code).strip())")
    out.append("    return entry['l1'] if entry else None")
    out.append("")
    out.append("")
    out.append("def get_l2_for_account(asset_key: str, account_code: str) -> str | None:")
    out.append('    """L2 section del (asset, code). Retorna None si no hay match."""')
    out.append("    if not asset_key or not account_code:")
    out.append("        return None")
    out.append("    cat = ASSET_CATALOGS.get(asset_key)")
    out.append("    if not cat:")
    out.append("        return None")
    out.append("    entry = cat['accounts'].get(str(account_code).strip())")
    out.append("    return entry['l2'] if entry else None")
    return "\n".join(out)


def main():
    print("[extract_catalogs] Parsing Excel files...", file=sys.stderr)
    alice_entries = extract_alice()
    edson_entries = extract_edson()
    val_entries = extract_val()
    print(f"  Alice: {len(alice_entries)} rows, Edson: {len(edson_entries)} rows, Val: {len(val_entries)} rows",
          file=sys.stderr)

    catalogs = {}
    total_unmatched = set()

    alice_cat, alice_unm = build_catalog("Alice House JV LLC", alice_entries, plan="TMG", open_list=True)
    catalogs["Alice House JV LLC"] = alice_cat
    total_unmatched |= alice_unm

    edson_cat, edson_unm = build_catalog("295 29th Street JV LLC", edson_entries, plan="TMG", open_list=True)
    catalogs["295 29th Street JV LLC"] = edson_cat
    total_unmatched |= edson_unm

    val_cat, val_unm = build_catalog("The Val", val_entries, plan="HDFC", open_list=False)
    catalogs["The Val"] = val_cat
    total_unmatched |= val_unm

    # ── Cuentas 1xxx de Val (capex / balance sheet items) ───────────────
    # Estas cuentas NO viven en el P&L del Val Budget.xlsx (mi extractor las
    # skippea via NON_PNL_SECTIONS "Cash Flow Conversion"), PERO sí aparecen
    # en los reportes MENSUALES del Val. Sin catálogo, el pipeline heredaba
    # la sección "Financial" y las clasificaba como Interest Expense.
    #
    # Feedback Florencia (backtest Q4 2025 The Val): "varias cuentas de Capex
    # las toma como Interest Expense" — 1420-000 BUILDINGS, 1464-000 FURNITURE,
    # 1315-009 CAPITAL EXP, 1470-000 APPLIANCES, 1315-000 SPECIAL ESCROW.
    #
    # Fix: inyectar entradas sintéticas en el catálogo del Val con L1=Total Capex.
    val_extra_capex = {
        '1115-000': "ESCROW",
        '1315-000': "SPECIAL ESCROW",
        '1315-009': "CAPITAL EXP OWNER IMP (1318_1319)",
        '1318-000': "OWNER IMPROVEMENTS",
        '1319-000': "CAPITAL IMPROVEMENTS",
        '1321-000': "CY-REPLACEMENT RESERVE",
        '1322-000': "REPLACEMENT RESERVE REFUND",
        '1420-000': "BUILDINGS",
        '1464-000': "FURNITURE & FIXTURES",
        '1470-000': "APPLIANCES",
    }
    n_injected = 0
    for code, desc in val_extra_capex.items():
        if code not in catalogs["The Val"]["accounts"]:
            catalogs["The Val"]["accounts"][code] = {
                "desc": desc,
                "l1": "Total Capex",
                "l2": "Cash Flow Conversion",
            }
            n_injected += 1
    print(f"  Val: inyectadas {n_injected} cuentas 1xxx sintéticas como Total Capex", file=sys.stderr)

    walnut_entries = extract_walnut()
    print(f"  Walnut: {len(walnut_entries)} rows", file=sys.stderr)
    walnut_cat, walnut_unm = build_catalog("Walnut Street Wellesley", walnut_entries, plan="JDE", open_list=True)
    catalogs["Walnut Street Wellesley"] = walnut_cat
    total_unmatched |= walnut_unm

    wilcox_entries = extract_wilcox()
    print(f"  Wilcox: {len(wilcox_entries)} rows", file=sys.stderr)
    wilcox_cat, wilcox_unm = build_catalog("The Wilcox", wilcox_entries, plan="YSI", open_list=True)
    catalogs["The Wilcox"] = wilcox_cat
    total_unmatched |= wilcox_unm

    # Stats
    for asset, cat in catalogs.items():
        print(f"  {asset}: {len(cat['accounts'])} accounts classified", file=sys.stderr)
    if total_unmatched:
        print(f"\n[warn] Sections NOT matched to any L1 (skipped):", file=sys.stderr)
        for s in sorted(total_unmatched):
            if s:
                print(f"    - {s!r}", file=sys.stderr)

    # Write to file (avoid Windows cp1252 issues with stdout)
    out_path = Path(__file__).parent.parent / "src" / "asset_catalogs.py"
    out_path.write_text(emit_python(catalogs), encoding="utf-8")
    print(f"[extract_catalogs] Wrote {out_path}", file=sys.stderr)


if __name__ == "__main__":
    main()
