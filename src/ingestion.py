"""
Módulo de Ingesta: carga archivos Excel del socio y extrae datos estructurados.
Preserva Variance Notes, distingue Current vs YTD, detecta metadata.

Soporta:
- Archivos con múltiples hojas (detecta automáticamente la hoja del income statement)
- Formatos variados de distintos socios (detección flexible de headers y columnas)
- Cualquier activo/edificio (detección genérica de nombre de propiedad)
"""

import pandas as pd
import openpyxl
import re
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field


@dataclass
class IngestedFile:
    """Resultado de ingerir un archivo Excel."""
    filename: str
    building: str
    period_label: str
    df: pd.DataFrame              # Datos con columnas estandarizadas
    partner_comments: List[Dict]  # Comentarios del socio extraídos
    section_totals: List[Dict]    # Filas de totales detectadas
    metadata: Dict[str, Any] = field(default_factory=dict)


# ─────────────────────────────────────────
# DETECCIÓN DE HOJAS (multi-sheet support)
# ─────────────────────────────────────────

# Keywords que indican un income statement / estado de resultados
_IS_KEYWORDS = [
    'budget', 'actual', 'variance', 'income', 'revenue', 'expense',
    'operating', 'noi', 'net income', 'net operating', 'total revenues',
    'total expenses', 'total operating', 'account',
    'presupuesto', 'ingresos', 'gastos', 'resultado',
]

# Keywords en nombres de hojas que claramente NO son el income statement
_SKIP_SHEET_PATTERNS = [
    r'chart',
    r'pivot',
    r'dashboard',
    r'macro',
    r'config',
    r'instructions?',
    r'cover',
    r'table\s*of\s*contents',
]

# Keywords que indican hojas de comentarios / variance notes
_COMMENTS_KEYWORDS = [
    'variance note', 'variance report', 'notes', 'comments',
    'comentarios', 'notas de varianza', 'explanations',
]


def _score_sheet_as_income_statement(ws, sheet_name: str) -> float:
    """
    Puntúa una hoja de 0-100 según su probabilidad de ser el income statement.
    Escanea las primeras filas y columnas buscando keywords financieras.
    """
    score = 0.0
    name_lower = sheet_name.lower().strip()

    # Penalizar hojas con nombres que claramente no son IS
    for pattern in _SKIP_SHEET_PATTERNS:
        if re.search(pattern, name_lower):
            score -= 20

    # Bonus por nombre descriptivo
    if any(kw in name_lower for kw in ['income', 'p&l', 'p & l', 'profit', 'loss',
                                         'estado', 'resultado', 'financial',
                                         'budget', 'actual', 'operating']):
        score += 25
    if any(kw in name_lower for kw in ['detail', 'detalle', 'summary', 'resumen']):
        score += 10
    # Bonus extra: nombres clásicos de reporte de P&L por socio
    if name_lower in ('bcr', 'budget comparison', 'budget comparison report', 'income statement'):
        score += 60
    if 'budget comparison' in name_lower:
        score += 30

    # Escanear contenido de la hoja (primeras 25 filas, 20 columnas)
    max_row = min(ws.max_row or 1, 25)
    max_col = min(ws.max_column or 1, 20)
    keywords_found = set()
    numeric_cells = 0
    account_code_cells = 0

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            val = ws.cell(row=row, column=col).value
            if val is None:
                continue

            if isinstance(val, (int, float)):
                numeric_cells += 1
                continue

            val_str = str(val).strip().lower()

            for kw in _IS_KEYWORDS:
                if kw in val_str:
                    keywords_found.add(kw)

            # Detectar account codes (patrón tipo XXXX-XXXX o numérico 4+ dígitos)
            if re.match(r'^\d{4,}[-.]?\d{0,4}$', val_str.replace(' ', '')):
                account_code_cells += 1

    score += len(keywords_found) * 5

    # Bonus especial: Budget + Actual = firma de income statement
    if 'budget' in keywords_found and 'actual' in keywords_found:
        score += 20
    if 'account' in keywords_found:
        score += 10

    if numeric_cells > 20:
        score += 15
    elif numeric_cells > 5:
        score += 8

    if account_code_cells > 3:
        score += 15

    return score


def _score_sheet_as_comments(ws, sheet_name: str) -> float:
    """
    Puntúa una hoja según su probabilidad de contener comentarios del socio
    en formato dedicado (hoja separada del IS).
    """
    score = 0.0
    name_lower = sheet_name.lower().strip()

    for kw in _COMMENTS_KEYWORDS:
        if kw in name_lower:
            score += 30

    max_row = min(ws.max_row or 1, 15)
    max_col = min(ws.max_column or 1, 10)

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            val = ws.cell(row=row, column=col).value
            if val is None:
                continue
            val_str = str(val).strip().lower()
            if any(kw in val_str for kw in ['variance note', 'explanation', 'comment',
                                              'notas', 'comentario']):
                score += 10

    return score


def detect_sheets(wb, filename: str = "") -> Dict[str, Any]:
    """
    Analiza todas las hojas del workbook y determina cuál es el income statement
    y cuál (si existe) contiene comentarios del socio en hoja separada.

    Returns:
        Dict con:
        - 'income_statement': nombre de la hoja del IS
        - 'comments_sheet': nombre de la hoja de comentarios (o None)
        - 'all_scores': scores de todas las hojas (para debugging/metadata)
    """
    is_scores = {}
    comment_scores = {}

    # Stem del filename para bonus de coincidencia (ej: "Edson 12" de "Edson 12.xlsm")
    filename_stem = Path(filename).stem.lower().strip() if filename else ""

    for name in wb.sheetnames:
        ws = wb[name]
        if (ws.max_row or 0) < 3:
            continue
        score = _score_sheet_as_income_statement(ws, name)
        # Bonus fuerte si el nombre de hoja coincide exactamente con el filename
        if filename_stem and name.lower().strip() == filename_stem:
            score += 100
        is_scores[name] = score
        comment_scores[name] = _score_sheet_as_comments(ws, name)

    if not is_scores:
        is_sheet = wb.sheetnames[0]
    else:
        is_sheet = max(is_scores, key=is_scores.get)

    comments_sheet = None
    if comment_scores:
        best_comment = max(comment_scores, key=comment_scores.get)
        if comment_scores[best_comment] > 20 and best_comment != is_sheet:
            comments_sheet = best_comment

    return {
        'income_statement': is_sheet,
        'comments_sheet': comments_sheet,
        'all_scores': {'is': is_scores, 'comments': comment_scores},
    }


# ─────────────────────────────────────────
# DETECCIÓN GENÉRICA DE BUILDING Y PERÍODO
# ─────────────────────────────────────────

_PERIOD_PATTERNS = [
    r'(?:for\s+(?:the\s+)?)?(?:month|period|quarter)\s+end(?:ing|ed)',
    r'(?:mes|periodo|trimestre)\s+(?:terminando|que\s+termina)',
    r'\b(?:january|february|march|april|may|june|july|august|september|october|november|december)\b.*\d{4}',
    r'\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b.*\d{4}',
    r'\bQ[1-4]\s*\d{4}\b',
]


def detect_building_and_period(ws, filename: str = "") -> tuple:
    """
    Extrae nombre de edificio y período de las primeras filas.
    Detección genérica: funciona con cualquier nombre de propiedad/activo.

    Estrategia:
    1. Escanea filas 1-8 buscando textos significativos en col A y col B
       (no solo col A; algunos formatos como BCR dejan col A con etiqueta
       "As of Date:" y el valor real en col B).
    2. Identifica período por keywords de fecha/período, o por valores
       datetime reales (ej: col B row 4 en BCR = datetime(2025,10,31)).
    3. Lo que no es período ni header técnico → candidato a nombre de edificio
    4. Fallback: extrae nombre del archivo; extrae período del filename si
       incluye pattern `YYYY-MM` (ej: "501 Estates ... 2025-10.xlsx").
    """
    import datetime as _dt

    building = ""
    period_label = ""
    candidates = []

    def _is_whitespace_only(v):
        """Un string de solo whitespace NO es contenido real."""
        if v is None:
            return True
        return isinstance(v, str) and not v.strip()

    # Escanear col A y col B explícitamente (antes solo fallback si col A vacía)
    for i in range(1, 9):
        for col in (1, 2):
            raw = ws.cell(row=i, column=col).value
            if _is_whitespace_only(raw):
                continue

            # datetime real (caso BCR row 4 col B = datetime(2025,10,31))
            if isinstance(raw, (_dt.datetime, _dt.date)):
                if not period_label:
                    period_label = raw.strftime("%Y-%m-%d")
                continue

            val_str = str(raw).strip()
            if len(val_str) < 3:
                continue
            val_lower = val_str.lower()

            # Detectar período
            is_period = False
            for pattern in _PERIOD_PATTERNS:
                if re.search(pattern, val_lower):
                    if not period_label:
                        period_label = val_str
                    is_period = True
                    break
            if not is_period and ('ending' in val_lower or 'period' in val_lower
                                   or 'as of date' in val_lower):
                # "As of Date:" en col A — el valor viene en col B (ya se leerá
                # en la próxima iteración del for col). NO setear period_label
                # con el texto "As of Date:", solo marcarlo como etiqueta.
                if 'as of date' in val_lower:
                    is_period = True  # evitar que quede como candidato de building
                else:
                    if not period_label:
                        period_label = val_str
                    is_period = True

            # Si no es período ni header técnico → candidato a building name
            if not is_period:
                skip_keywords = ['account', 'budget', 'actual', 'variance', 'description',
                                 'current', 'ytd', 'month', 'notes', 'total',
                                 'statement', 'profit', 'loss', 'consolidated', 'financial',
                                 'revenues', 'revenue', 'expenses', 'income', 'rental',
                                 'operating', 'interest', 'taxes', 'capex', 'payroll',
                                 'utilities', 'insurance', 'maintenance', 'management',
                                 'reporting book', 'accrual', 'location:']
                if not any(kw in val_lower for kw in skip_keywords):
                    candidates.append(val_str)

    if candidates:
        building = candidates[0]

    # Fallback filename: nombre del edificio
    if not building and filename:
        clean = Path(filename).stem
        clean = re.sub(r'\b\d{1,2}\b', '', clean)
        clean = re.sub(r'\b(Q[1-4])\b', '', clean, flags=re.IGNORECASE)
        clean = re.sub(r'\b20\d{2}\b', '', clean)
        clean = re.sub(r'[_\-]+', ' ', clean).strip()
        if clean:
            building = clean

    # Fallback filename: período (pattern YYYY-MM o similar)
    # Útil para BCR ("501 Estates ... 2025-10.xlsx") donde a veces el período
    # no aparece en el header del Excel sino solo en el nombre del archivo.
    if not period_label and filename:
        m = re.search(r'(20\d{2})[-_.](\d{1,2})\b', filename)
        if m:
            yr, mo = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12:
                period_label = f"{yr}-{mo:02d}"

    return building, period_label


# ─────────────────────────────────────────
# DETECCIÓN FLEXIBLE DE HEADERS Y COLUMNAS
# ─────────────────────────────────────────

_HEADER_PATTERNS = {
    'account': [r'account', r'acct', r'cuenta', r'c[oó]digo', r'code', r'gl\s*#', r'gl\s*account'],
    'description': [r'description', r'descripci[oó]n', r'name', r'nombre', r'detail'],
    'budget': [r'budget', r'presupuesto', r'plan', r'projected', r'forecast'],
    'actual': [r'actual', r'real', r'ejecutado'],
    'variance': [r'variance', r'varianza', r'var\.?\s', r'difference', r'diferencia'],
    'notes': [r'variance\s*notes?', r'notes?$', r'comentarios?', r'explanations?',
              r'notas?\s*(?:de\s*)?varianza'],
    'ytd_group': [r'ytd', r'year\s*to\s*date', r'acumulado', r'a[ñn]o'],
    'current_group': [r'current', r'month', r'mes\s*actual', r'periodo'],
}


def _match_header(text: str, category: str) -> bool:
    """Verifica si un texto de header coincide con algún patrón de la categoría."""
    text_lower = text.lower().strip()
    for pattern in _HEADER_PATTERNS.get(category, []):
        if re.search(pattern, text_lower):
            return True
    return False


def find_header_row(ws, max_search: int = 20) -> int:
    """
    Encuentra la fila de headers con scoring flexible.
    Soporta múltiples formatos de distintos socios.
    """
    best_row = None
    best_score = 0

    for i in range(1, max_search + 1):
        row_vals = []
        for col in range(1, 20):
            v = ws.cell(row=i, column=col).value
            if v:
                row_vals.append(str(v).strip())

        if not row_vals:
            continue

        score = 0
        has_budget = any(_match_header(v, 'budget') for v in row_vals)
        has_actual = any(_match_header(v, 'actual') for v in row_vals)
        has_account = any(_match_header(v, 'account') for v in row_vals)
        has_description = any(_match_header(v, 'description') for v in row_vals)
        has_variance = any(_match_header(v, 'variance') for v in row_vals)
        has_notes = any(_match_header(v, 'notes') for v in row_vals)

        if has_budget:
            score += 3
        if has_actual:
            score += 3
        if has_account:
            score += 2
        if has_description:
            score += 2
        if has_variance:
            score += 1
        if has_notes:
            score += 1

        # Requiere mínimo budget o actual + algo más
        if score >= 5 and (has_budget or has_actual):
            if score > best_score:
                best_score = score
                best_row = i

    if best_row:
        return best_row

    raise ValueError(
        "No se detectó fila de headers. Buscando columnas como 'Account', 'Budget', 'Actual', "
        "'Description', etc. Verifica que el archivo tenga un header de columnas reconocible."
    )


def _map_columns_flexible(ws, header_row: int) -> Dict[str, Any]:
    """
    Mapea las columnas del header de forma flexible, soportando múltiples formatos.

    Returns:
        Dict con posiciones de columnas detectadas:
        - desc_col, acct_col, budget_curr, actual_curr,
          budget_ytd, actual_ytd, notes_col, variance_col
    """
    col_list = []
    for col in range(1, 25):
        val = ws.cell(row=header_row, column=col).value
        if val:
            col_list.append((col, str(val).strip()))

    # Detectar agrupación Current vs YTD desde filas superiores al header
    ytd_start_col = None
    for search_row in range(max(1, header_row - 2), header_row):
        for col in range(1, 25):
            val = ws.cell(row=search_row, column=col).value
            if val and _match_header(str(val), 'ytd_group'):
                ytd_start_col = col
                break
        if ytd_start_col:
            break

    # Mapear cada columna a su rol
    desc_col = None
    acct_col = None
    budget_positions = []
    actual_positions = []
    variance_positions = []
    notes_col = None

    for c, name in col_list:
        if _match_header(name, 'notes') and notes_col is None:
            notes_col = c
        elif _match_header(name, 'account') and acct_col is None:
            acct_col = c
        elif _match_header(name, 'description') and desc_col is None:
            desc_col = c
        elif _match_header(name, 'budget'):
            budget_positions.append(c)
        elif _match_header(name, 'actual'):
            actual_positions.append(c)
        elif _match_header(name, 'variance'):
            variance_positions.append(c)

    # Defaults si no se encontraron
    if desc_col is None:
        desc_col = 1
    if acct_col is None:
        # Asumir columna adyacente a description, o columna 2
        acct_col = desc_col + 1 if desc_col == 1 else 2

    # Asignar Current vs YTD
    if ytd_start_col and len(budget_positions) >= 2:
        curr_budgets = [c for c in budget_positions if c < ytd_start_col]
        curr_actuals = [c for c in actual_positions if c < ytd_start_col]
        ytd_budgets = [c for c in budget_positions if c >= ytd_start_col]
        ytd_actuals = [c for c in actual_positions if c >= ytd_start_col]

        budget_curr = curr_budgets[0] if curr_budgets else budget_positions[0]
        actual_curr = curr_actuals[0] if curr_actuals else actual_positions[0]
        budget_ytd = ytd_budgets[0] if ytd_budgets else None
        actual_ytd = ytd_actuals[0] if ytd_actuals else None
    else:
        budget_curr = budget_positions[0] if budget_positions else 3
        actual_curr = actual_positions[0] if actual_positions else 4
        budget_ytd = None
        actual_ytd = None

    return {
        'desc_col': desc_col,
        'acct_col': acct_col,
        'budget_curr': budget_curr,
        'actual_curr': actual_curr,
        'budget_ytd': budget_ytd,
        'actual_ytd': actual_ytd,
        'notes_col': notes_col,
    }


# ─────────────────────────────────────────
# EXTRACCIÓN DE COMENTARIOS DE HOJA SEPARADA
# ─────────────────────────────────────────

def _extract_comments_from_sheet(ws, source_filename: str) -> List[Dict]:
    """
    Extrae comentarios del socio de una hoja dedicada de variance notes.
    Busca columnas de account code + texto de comentario.
    """
    comments = []

    # Encontrar header en las primeras filas
    header_row = None
    acct_col = None
    desc_col = None
    comment_col = None

    for i in range(1, 10):
        for col in range(1, 15):
            val = ws.cell(row=i, column=col).value
            if not val:
                continue
            val_str = str(val).strip().lower()

            if any(kw in val_str for kw in ['account', 'acct', 'código', 'code']):
                acct_col = col
                header_row = i
            elif any(kw in val_str for kw in ['description', 'descripción', 'name']):
                desc_col = col
            elif any(kw in val_str for kw in ['comment', 'note', 'explanation',
                                                'comentario', 'nota', 'variance']):
                comment_col = col

    if header_row is None or comment_col is None:
        return comments

    if acct_col is None:
        acct_col = 1
    if desc_col is None:
        desc_col = acct_col + 1

    max_row = ws.max_row or 100
    for i in range(header_row + 1, max_row + 1):
        acct_val = ws.cell(row=i, column=acct_col).value
        desc_val = ws.cell(row=i, column=desc_col).value
        comment_val = ws.cell(row=i, column=comment_col).value

        if not comment_val or not str(comment_val).strip():
            continue

        comments.append({
            'account_code': str(acct_val).strip() if acct_val else '',
            'description': str(desc_val).strip() if desc_val else '',
            'comment_text': str(comment_val).strip(),
            'source_file': source_filename,
            'row_num': i,
        })

    return comments


# ─────────────────────────────────────────
# INGESTA PRINCIPAL
# ─────────────────────────────────────────

def get_sheet_candidates(filepath: Path) -> List[Dict]:
    """
    Retorna hojas candidatas a income statement con sus scores.
    Útil para mostrar selector al usuario cuando hay múltiples hojas válidas.

    Returns:
        Lista de dicts con 'name' y 'score', ordenados por score desc.
    """
    wb = openpyxl.load_workbook(str(filepath), data_only=True)
    result = detect_sheets(wb)
    wb.close()
    scores = result['all_scores'].get('is', {})
    candidates = [
        {'name': name, 'score': score}
        for name, score in scores.items()
        if score >= 40
    ]
    candidates.sort(key=lambda x: -x['score'])
    return candidates


# ─────────────────────────────────────────
# FORMATO NWEP-SUMMARY (Resultado Walnut / EERR — Walnut Street Wellesley)
# ─────────────────────────────────────────
#
# Layout: archivo 'Resultado Walnut XX-YYYY.xlsx' o 'EERR XX-YYYY.xlsx'.
# El socio manda categorías a nivel L2 (Salaries, Utilities, Repairs &
# Maintenance, etc.) en col A y montos PTD/YTD en cols B–G. NO hay códigos
# GL — el parser sintetiza account codes 'SUM:<slug>' que se mapean al
# catálogo de Walnut Street Wellesley.

# Mapeo de descripción L2 -> (L1, L2 canónico). Las descripciones se
# normalizan a lowercase para el match. Si una descripción no está acá, la
# fila cae al fallback global (que típicamente la deja sin clasificar).
_NWEP_SUMMARY_L1_MAP = {
    # ── Income (above 'Total Income')
    'retail-base':              ('Income', 'Rental Income'),
    'commercial-base':          ('Income', 'Rental Income'),
    'percentage rents':         ('Income', 'Rental Income'),
    'lease term concessions':   ('Income', 'Concessions and Vacancy'),
    'total recovery income':    ('Income', 'Other Income'),
    'other income':             ('Income', 'Other Income'),
    'non operating income':     ('Income', 'Other Income'),
    # ── Operating Expenses (between 'Total Income' y 'Total Operating Expenses')
    'salaries':                 ('Operating Expenses', 'Payroll and Related'),
    'utilities':                ('Operating Expenses', 'Utilities'),
    'repairs & maintenance':    ('Operating Expenses', 'Repairs and Maintenance'),
    'project administration':   ('Operating Expenses', 'Property G & A'),
    'management fees':          ('Operating Expenses', 'Management Fees'),
    'insurance premiums':       ('Operating Expenses', 'Insurance'),
    'non recoverable costs':    ('Operating Expenses', 'Property G & A'),
    # ── Real Estate Taxes (L1 separada por convención)
    'real estate taxes':        ('Real Estate Taxes', 'Property Taxes'),
    # ── Below NOI — Non-Operating
    'annual audit':                     ('Non-Operating Expenses', 'Company General & Administrative'),
    'ownership costs/asset mgt fees':   ('Non-Operating Expenses', 'Company General & Administrative'),
    'ownership costs/asset mgmt':       ('Non-Operating Expenses', 'Company General & Administrative'),
    'ownership costs/asset':            ('Non-Operating Expenses', 'Company General & Administrative'),
    'organizational expenses':          ('Non-Operating Expenses', 'Company General & Administrative'),
    'organizational expense':           ('Non-Operating Expenses', 'Company General & Administrative'),
    'partnership rel-other':            ('Non-Operating Expenses', 'Company General & Administrative'),
    'tax annual rpt filing fee':        ('Non-Operating Expenses', 'Company General & Administrative'),
    'tax annual rpt filing':            ('Non-Operating Expenses', 'Company General & Administrative'),
    'project bad debt exp':             ('Non-Operating Expenses', 'Bad Debt'),
    'unrealized gain/loss':             ('Non-Operating Expenses', 'Unrealized Gain/Loss'),
    'amortize expense':                 ('Non-Operating Expenses', 'Depreciation/Amortization'),
    # ── Below NOI — Interest
    'mortgage int-perm':                ('Interest Expense', 'Interest and Financing Expenses'),
    'finan cost-legal':                 ('Interest Expense', 'Interest and Financing Expenses'),
}


def _is_nwep_summary_format(ws) -> bool:
    """
    Detecta el formato SUMMARY usado por NWEP/Walnut Street Wellesley.

    Firma:
      - Row 1, col A empieza con 'SUMMARY '
      - Row 5, col A == 'Month'
      - Row 5, col B contiene 'Actual' (típicamente 'PTD Actual')
    """
    r1 = ws.cell(row=1, column=1).value
    r5c1 = ws.cell(row=5, column=1).value
    r5c2 = ws.cell(row=5, column=2).value
    if r1 is None or r5c1 is None or r5c2 is None:
        return False
    s1 = str(r1).strip()
    s5c1 = str(r5c1).strip().lower()
    s5c2 = str(r5c2).strip().lower()
    return (
        s1.upper().startswith('SUMMARY ') and
        s5c1 == 'month' and
        'actual' in s5c2
    )


def _ingest_nwep_summary_sheet(ws, filepath: Path, sheet_info: Dict) -> IngestedFile:
    """
    Parser para formato SUMMARY de NWEP/Walnut Street Wellesley.

    Layout:
      - Row 1, col A: 'SUMMARY <nombre activo>'
      - Row 2, col A: período (datetime)
      - Row 5: headers ['Month', 'PTD Actual', 'PTD Budget', 'Variance',
                         'YTD Actual', 'YTD Budget', 'Variance']
      - Row 6+: datos. Cada fila es L2-level (no GL). Algunas filas son totales
        L1 ('Total Income', 'Total Operating Expenses', 'Net Operating Income').

    Sintetiza account codes 'SUM:<slug>' para cada descripción para que el
    catálogo pueda mapear consistentemente. Filas no listadas en el catálogo
    caen al fallback estándar.
    """
    import datetime as _dt

    # Building & período
    building = ""
    v1 = ws.cell(row=1, column=1).value
    if v1:
        building = str(v1).strip()
    period_label = ""
    v2 = ws.cell(row=2, column=1).value
    if isinstance(v2, (_dt.datetime, _dt.date)):
        period_label = v2.strftime("%Y-%m-%d")
    elif v2:
        period_label = str(v2).strip()

    # Layout fijo
    desc_col = 1
    actual_curr = 2
    budget_curr = 3
    actual_ytd = 5
    budget_ytd = 6
    header_row = 5

    # Filas que son L1 totals
    L1_TOTALS = {
        'total income': 'Income',
        'total operating expenses': 'Operating Expenses',  # primera ocurrencia (above NOI)
        'net operating income': 'NOI',
        'net income': 'Net Income',
    }

    # Filas que NO son data — son subtotales intermedios o labels que el pipeline
    # NO debe contar (la suma se reconstruye desde sus componentes).
    SKIP_LABELS = {
        'total operating expens',         # truncado: subtotal Non-Op + Interest
        'total operating expense other',  # versión no truncada del mismo subtotal
    }

    rows = []
    section_totals = []
    seen_l1 = set()

    max_row = ws.max_row or 50
    for i in range(header_row + 1, max_row + 1):
        desc_val = ws.cell(row=i, column=desc_col).value
        actual_c = ws.cell(row=i, column=actual_curr).value
        budget_c = ws.cell(row=i, column=budget_curr).value
        actual_y = ws.cell(row=i, column=actual_ytd).value
        budget_y = ws.cell(row=i, column=budget_ytd).value

        has_values = isinstance(actual_c, (int, float)) or isinstance(budget_c, (int, float))
        if desc_val is None and not has_values:
            continue

        desc_str = str(desc_val).strip() if desc_val is not None else ""
        if not desc_str:
            continue

        desc_lower = desc_str.lower()

        # Skip labels (subtotales que no son L1)
        if desc_lower in SKIP_LABELS:
            rows.append({
                'row_num': i, 'description': desc_str, 'account': "",
                'budget_current': _to_float(budget_c), 'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y), 'actual_ytd': _to_float(actual_y),
                'row_type': 'label', 'source_file': filepath.name,
            })
            continue

        is_total_label = desc_lower in L1_TOTALS

        # Segunda ocurrencia de "Total Operating Expenses" no es un L1 total real
        if is_total_label and L1_TOTALS[desc_lower] in seen_l1:
            is_total_label = False

        if is_total_label:
            l1_name = L1_TOTALS[desc_lower]
            seen_l1.add(l1_name)
            rows.append({
                'row_num': i, 'description': desc_str, 'account': "",
                'budget_current': _to_float(budget_c), 'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y), 'actual_ytd': _to_float(actual_y),
                'row_type': 'total', 'source_file': filepath.name,
            })
            section_totals.append({
                'description': desc_str,
                'budget_current': _to_float(budget_c), 'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y), 'actual_ytd': _to_float(actual_y),
                'row_num': i,
            })
        else:
            # Data row: sintetizar account code 'SUM:<slug>'
            slug = re.sub(r'[^a-z0-9]+', '-', desc_lower).strip('-')
            account = f'SUM:{slug}'
            rows.append({
                'row_num': i, 'description': desc_str, 'account': account,
                'budget_current': _to_float(budget_c), 'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y), 'actual_ytd': _to_float(actual_y),
                'row_type': 'data', 'source_file': filepath.name,
            })

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=[],
        section_totals=section_totals,
        metadata={
            'header_row': header_row,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': True,
            'sheet_used': sheet_info.get('income_statement'),
            'comments_sheet': None,
            'format': 'NWEP-Summary',
            'column_map': {
                'desc_col': desc_col, 'acct_col': None,
                'budget_curr': budget_curr, 'actual_curr': actual_curr,
                'budget_ytd': budget_ytd, 'actual_ytd': actual_ytd,
            },
        }
    )


def _is_bcr_format(ws) -> bool:
    """
    Detecta formato Budget Comparison Report (usado por 501 Estates y similares).
    Firma: row 2 contiene 'Budget Comparison Report' (con "Report" al final,
    distinto de Val/CMC que solo dice "Budget Comparison").
    """
    max_r = min(ws.max_row or 0, 6)
    for r in range(1, max_r + 1):
        for c in range(1, 3):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            vs = str(v).strip().lower()
            if 'budget comparison report' in vs:
                return True
    return False


def _ingest_bcr_sheet(ws, filepath: Path, sheet_info: Dict) -> "IngestedFile":
    """
    Parser para formato BCR (Budget Comparison Report — 501 Estates).

    Layout:
      - Row 1: nombre edificio (ej: "J 501 Estates")
      - Row 2: "Budget Comparison Report"
      - Row 3-5: metadata (Reporting Book, As of Date, Location)
      - Row 7: "Month Ending | ... | Year To Date | ..."
      - Row 8: fechas
      - Row 9 (headers): col1=(vacío), col2=Actual, col3=Budget, col4=Budget Diff,
                         col5=Budget % Var, col6=Actual(YTD), col7=Budget(YTD),
                         col8=Budget Diff(YTD), col9=Budget % Var(YTD)
      - Col 1: TODA la descripción con indentación. Para data rows formato:
               "XXXX-XXXX - Description" (ej: "3110-1110 - Market Rent")
               Para headers/totales: solo texto indentado.
    """
    building, period_label = detect_building_and_period(ws, filepath.name)

    # Encontrar header row buscando "Actual" en col 2
    header_row = None
    for r in range(5, 15):
        v = ws.cell(row=r, column=2).value
        if v and str(v).strip().lower() == 'actual':
            header_row = r
            break
    if header_row is None:
        header_row = 9

    # Layout fijo BCR
    desc_col = 1
    actual_curr = 2
    budget_curr = 3
    actual_ytd = 6
    budget_ytd = 7

    # Regex para detectar data rows: "XXXX-XXXX - Description" (formato 501 Estates)
    data_pattern = re.compile(r'^\s*(\d{3,4}[-\.]\d{3,4})\s*[-–—]\s*(.+)$')

    rows = []
    partner_comments = []
    section_totals = []

    max_row = ws.max_row or 600
    for i in range(header_row + 1, max_row + 1):
        c1 = ws.cell(row=i, column=desc_col).value
        if c1 is None:
            continue
        c1_str = str(c1)
        c1_stripped = c1_str.strip()
        if not c1_stripped:
            continue

        actual_c = ws.cell(row=i, column=actual_curr).value
        budget_c = ws.cell(row=i, column=budget_curr).value
        actual_y = ws.cell(row=i, column=actual_ytd).value
        budget_y = ws.cell(row=i, column=budget_ytd).value

        has_values = isinstance(actual_c, (int, float)) or isinstance(budget_c, (int, float))

        # Determinar tipo de fila
        data_match = data_pattern.match(c1_str)
        starts_with_total = c1_stripped.lower().startswith('total ')
        # L1 totals que no empiezan con "Total": "Net Operating Income", "Net Income",
        # "Net Income After Capital Expenditures"
        is_named_l1_total = has_values and c1_stripped.lower() in (
            'net operating income', 'net income', 'net income after capital expenditures',
        )

        if data_match:
            row_type = 'data'
            acct_str = data_match.group(1).strip()
            desc_str = data_match.group(2).strip()
        elif starts_with_total and has_values:
            row_type = 'total'
            acct_str = ""
            desc_str = c1_stripped
        elif is_named_l1_total:
            row_type = 'total'
            acct_str = ""
            desc_str = c1_stripped
        elif not has_values and not data_match:
            row_type = 'section'
            acct_str = ""
            desc_str = c1_stripped
        else:
            row_type = 'label'
            acct_str = ""
            desc_str = c1_stripped

        rows.append({
            'row_num': i,
            'description': desc_str,
            'account': acct_str,
            'budget_current': _to_float(budget_c),
            'actual_current': _to_float(actual_c),
            'budget_ytd': _to_float(budget_y),
            'actual_ytd': _to_float(actual_y),
            'row_type': row_type,
            'source_file': filepath.name,
        })

        if row_type == 'total':
            section_totals.append({
                'description': desc_str,
                'budget_current': _to_float(budget_c),
                'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y),
                'actual_ytd': _to_float(actual_y),
                'row_num': i,
            })

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'header_row': header_row,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': True,
            'sheet_used': sheet_info.get('income_statement'),
            'comments_sheet': sheet_info.get('comments_sheet'),
            'format': 'BCR',
            'column_map': {
                'desc_col': desc_col, 'acct_col': None,
                'budget_curr': budget_curr, 'actual_curr': actual_curr,
                'budget_ytd': budget_ytd, 'actual_ytd': actual_ytd,
            },
        }
    )


def _is_yardi_variance_format(ws) -> bool:
    """
    Detecta si la hoja usa el formato "Yardi Variance Report" (usado por Yale y similares).
    Firma: filas superiores contienen "Variance Report" Y headers de col A/B = "Account"/"Account Name".
    """
    max_r = min(ws.max_row or 0, 12)
    max_c = min(ws.max_column or 0, 4)
    found_variance_report = False
    found_account_headers = False
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            vs = str(v).strip().lower()
            if 'variance report' in vs:
                found_variance_report = True
            # Header exacto "account" en col 1 + "account name" en col 2
            if r <= 10 and c == 1 and vs == 'account':
                v2 = ws.cell(row=r, column=2).value
                if v2 and 'account name' in str(v2).strip().lower():
                    found_account_headers = True
    return found_variance_report and found_account_headers


def _ingest_yardi_variance_sheet(ws, filepath: Path, sheet_info: Dict) -> "IngestedFile":
    """
    Parser dedicado para el formato Yardi Variance Report (Yale).

    Layout:
      - Row 2-3: metadata (Variance Report, building name)
      - Row 4: período (ej: "Dec 2025")
      - Row 7: grupos de columnas (current period | YTD)
      - Row 8 (headers): col1=Account, col2=Account Name, col3=Actual, col4=Budget,
                         col5=$ Variance, col6=% Variance, col7=Variance Notes,
                         col8=YTD Actual, col9=YTD Budget, col10=YTD Variance, col11=YTD % Var
      - Data rows:   col1=código cuenta (ej "4010-0000"), col2=descripción, cols 3+ valores
      - Section headers: col1=nombre sección (ej "Income"), col2=vacío, sin valores
      - Subtotales:  col1=vacío, col2=nombre sección repetido, con valores
    """
    from dataclasses import dataclass
    building, period_label = detect_building_and_period(ws, filepath.name)

    # Detectar header row buscando la fila con "Account" en col 1
    header_row = None
    for r in range(1, 15):
        v = ws.cell(row=r, column=1).value
        if v and str(v).strip().lower() == 'account':
            header_row = r
            break
    if header_row is None:
        header_row = 8

    # Layout fijo por convención Yardi Variance Report
    acct_col = 1
    desc_col = 2
    actual_curr = 3
    budget_curr = 4
    notes_col = 7
    actual_ytd = 8
    budget_ytd = 9

    rows = []
    partner_comments = []
    section_totals = []

    max_row = ws.max_row or 300
    for i in range(header_row + 1, max_row + 1):
        acct_val = ws.cell(row=i, column=acct_col).value
        desc_val = ws.cell(row=i, column=desc_col).value
        actual_c = ws.cell(row=i, column=actual_curr).value
        budget_c = ws.cell(row=i, column=budget_curr).value
        actual_y = ws.cell(row=i, column=actual_ytd).value
        budget_y = ws.cell(row=i, column=budget_ytd).value
        notes_val = ws.cell(row=i, column=notes_col).value

        # Saltar filas completamente vacías
        if acct_val is None and desc_val is None and budget_c is None and actual_c is None:
            continue

        acct_str_raw = str(acct_val).strip() if acct_val else ""
        desc_str_raw = str(desc_val).strip() if desc_val else ""
        has_values = isinstance(budget_c, (int, float)) or isinstance(actual_c, (int, float))

        # Determinar si acct_str es realmente un código (formato XXXX-XXXX o similar)
        is_acct_code = bool(re.match(r'^\d{3,4}[-\s]?\d{0,4}$', acct_str_raw.replace(' ', '')))

        # Tipo de fila:
        # 1. Section header: col A tiene texto NO numérico, col B vacío, sin valores
        # 2. Subtotal: col A vacío, col B tiene texto, con valores
        # 3. Data: col A = código, col B = descripción, con valores
        if acct_str_raw and not is_acct_code and not desc_str_raw and not has_values:
            # Section header — el nombre está en col A
            row_type = 'section'
            desc_str = acct_str_raw
            acct_str = ""
        elif not acct_str_raw and desc_str_raw and has_values:
            # Subtotal / L1 total — desc en col B
            row_type = 'total'
            desc_str = desc_str_raw
            acct_str = ""
        elif is_acct_code and has_values:
            # Data row
            row_type = 'data'
            desc_str = desc_str_raw
            acct_str = acct_str_raw
        else:
            # Label / otros (ej. filas vacías parciales)
            row_type = 'label'
            desc_str = desc_str_raw or acct_str_raw
            acct_str = acct_str_raw if is_acct_code else ""

        rows.append({
            'row_num': i,
            'description': desc_str,
            'account': acct_str,
            'budget_current': _to_float(budget_c),
            'actual_current': _to_float(actual_c),
            'budget_ytd': _to_float(budget_y),
            'actual_ytd': _to_float(actual_y),
            'row_type': row_type,
            'source_file': filepath.name,
        })

        if notes_val and str(notes_val).strip():
            partner_comments.append({
                'account_code': acct_str,
                'description': desc_str,
                'comment_text': str(notes_val).strip(),
                'source_file': filepath.name,
                'row_num': i,
            })

        if row_type == 'total':
            section_totals.append({
                'description': desc_str,
                'budget_current': _to_float(budget_c),
                'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y),
                'actual_ytd': _to_float(actual_y),
                'row_num': i,
            })

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'header_row': header_row,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': True,
            'sheet_used': sheet_info.get('income_statement'),
            'comments_sheet': sheet_info.get('comments_sheet'),
            'sheets_analyzed': len(sheet_info.get('all_scores', {}).get('is', {})) or 1,
            'format': 'YardiVariance',
            'column_map': {
                'desc_col': desc_col, 'acct_col': acct_col,
                'budget_curr': budget_curr, 'actual_curr': actual_curr,
                'budget_ytd': budget_ytd, 'actual_ytd': actual_ytd,
                'notes_col': notes_col,
            },
        }
    )


def _is_cmc_format(ws) -> bool:
    """
    Detecta si la hoja usa el formato CMC Report Format (usado por The Val).
    Firma: la celda B6 (o cercana) contiene "CMC Report Format" y la fila 5
    tiene headers tipo "PTD Actual" / "PTD Budget".
    """
    max_r = min(ws.max_row or 0, 10)
    max_c = min(ws.max_column or 0, 6)
    found_cmc = False
    found_ptd = False
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            vs = str(v).strip().lower()
            if 'cmc report format' in vs:
                found_cmc = True
            if 'ptd' in vs and ('actual' in vs or 'budget' in vs):
                found_ptd = True
    return found_cmc or found_ptd


def _ingest_cmc_sheet(ws, filepath: Path, sheet_info: Dict) -> IngestedFile:
    """
    Parser específico para el formato CMC (The Val y otros activos con este layout).

    Layout CMC:
    - Row 1: nombre del edificio (ej: "The Val (832)")
    - Row 2: "Budget Comparison"
    - Row 3: "Period = Mon YYYY"
    - Row 5 (headers): col3=PTD Actual, col4=PTD Budget, col5=Variance, col6=%Var,
                       col7=YTD Actual, col8=YTD Budget, col9=Variance, col10=%Var,
                       col11=Annual, col12=Note
    - Col 1: código de cuenta (ej: "5120-000"), o vacío para section/total
    - Col 2: descripción / section header / total label (con indentación)
    """
    building, period_label = detect_building_and_period(ws, filepath.name)

    header_row = 5
    acct_col = 1
    desc_col = 2
    actual_curr = 3   # PTD Actual (mensual)
    budget_curr = 4   # PTD Budget (mensual)
    actual_ytd = 7
    budget_ytd = 8
    notes_col = 12

    rows = []
    partner_comments = []
    section_totals = []

    max_row = ws.max_row or 300
    for i in range(header_row + 1, max_row + 1):
        acct_val = ws.cell(row=i, column=acct_col).value
        desc_val = ws.cell(row=i, column=desc_col).value
        budget_c = ws.cell(row=i, column=budget_curr).value
        actual_c = ws.cell(row=i, column=actual_curr).value
        budget_y = ws.cell(row=i, column=budget_ytd).value
        actual_y = ws.cell(row=i, column=actual_ytd).value
        notes_val = ws.cell(row=i, column=notes_col).value

        if desc_val is None and acct_val is None and budget_c is None and actual_c is None:
            continue

        desc_str = str(desc_val).strip() if desc_val else ""
        acct_str = str(acct_val).strip() if acct_val else ""

        is_total = 'total' in desc_str.lower() or desc_str.strip().lower() == 'net operating income'
        is_section_header = (
            desc_str != "" and
            acct_str == "" and
            not isinstance(budget_c, (int, float)) and
            not isinstance(actual_c, (int, float)) and
            not is_total
        )
        has_data = isinstance(budget_c, (int, float)) or isinstance(actual_c, (int, float))

        row_type = "total" if is_total else ("section" if is_section_header else ("data" if has_data else "label"))

        rows.append({
            'row_num': i,
            'description': desc_str,
            'account': acct_str,
            'budget_current': _to_float(budget_c),
            'actual_current': _to_float(actual_c),
            'budget_ytd': _to_float(budget_y),
            'actual_ytd': _to_float(actual_y),
            'row_type': row_type,
            'source_file': filepath.name,
        })

        if notes_val and str(notes_val).strip():
            partner_comments.append({
                'account_code': acct_str,
                'description': desc_str,
                'comment_text': str(notes_val).strip(),
                'source_file': filepath.name,
                'row_num': i,
            })

        if is_total:
            section_totals.append({
                'description': desc_str,
                'budget_current': _to_float(budget_c),
                'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y),
                'actual_ytd': _to_float(actual_y),
                'row_num': i,
            })

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'header_row': header_row,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': True,
            'sheet_used': sheet_info['income_statement'],
            'comments_sheet': sheet_info.get('comments_sheet'),
            'sheets_analyzed': len(sheet_info.get('all_scores', {}).get('is', {})) or 1,
            'format': 'CMC',
            'column_map': {
                'desc_col': desc_col, 'acct_col': acct_col,
                'budget_curr': budget_curr, 'actual_curr': actual_curr,
                'budget_ytd': budget_ytd, 'actual_ytd': actual_ytd,
                'notes_col': notes_col,
            },
        }
    )


# ─────────────────────────────────────────
# FORMATO CWS / 12-MONTH BUDGET (Bridge at the Blockyard)
# ─────────────────────────────────────────
#
# Layout: Yardi 12-month budget con 'Tree = cws_res_is_' (Carroll/Wells/Steele).
# Firma: row 4 contiene 'Tree = cws_'. Cols A=código GL, B=descripción,
# C-N = 12 meses (Jan-Dec), O = Total anual.
#
# A diferencia de YSI/CMC (que traen un período con Actual+Budget), este formato
# es un BUDGET ANUAL — no hay actuals. La pipeline lo combina con un archivo
# separado de actuals mensuales (ver Phase 2 — pendiente).
#
# Por defecto el parser extrae el TOTAL ANUAL (col O) como budget_current. Si
# se le pasa target_period (ej. 'Q1 2026' o 'Mar 2026'), suma los meses
# correspondientes para budget_current y los meses YTD para budget_ytd.

def _is_cws_format(ws) -> bool:
    """
    Detecta formato CWS (Bridge at the Blockyard y similares).
    Firma: row 4 contiene 'Tree = cws_<algo>'.
    """
    max_r = min(ws.max_row or 0, 6)
    for r in range(1, max_r + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        s = str(v).lower()
        if 'tree' in s and '=' in s and 'cws_' in s:
            return True
    return False


def _parse_cws_target_period(target_period: Optional[str], file_year: int) -> Tuple[Optional[List[int]], Optional[List[int]]]:
    """
    Convierte target_period a (current_months, ytd_months) para el año del archivo.
    Retorna (None, None) si no se puede parsear o el año no coincide → caller
    usará el default (annual).

    Ejemplos (file_year=2026):
      'Q1 2026'      → ([1,2,3], [1,2,3])
      'Q3 2026'      → ([7,8,9], [1..9])
      'Mar 2026'     → ([3], [1,2,3])
      'October 2026' → ([10], [1..10])
      'Q1 2025'      → (None, None)  # año no coincide
    """
    if not target_period:
        return None, None

    s = str(target_period).strip().lower()

    # Extraer año del period
    year_match = re.search(r'\b(20\d{2})\b', s)
    period_year = int(year_match.group(1)) if year_match else None
    if period_year is not None and period_year != file_year:
        return None, None

    # Q1-Q4
    q_match = re.search(r'\bq([1-4])\b', s)
    if q_match:
        q = int(q_match.group(1))
        end_month = q * 3
        current = list(range(end_month - 2, end_month + 1))
        ytd = list(range(1, end_month + 1))
        return current, ytd

    # Mes único (jan-dec o january-december o "01"-"12")
    month_names = {
        'jan': 1, 'january': 1, 'feb': 2, 'february': 2, 'mar': 3, 'march': 3,
        'apr': 4, 'april': 4, 'may': 5, 'jun': 6, 'june': 6,
        'jul': 7, 'july': 7, 'aug': 8, 'august': 8, 'sep': 9, 'september': 9,
        'oct': 10, 'october': 10, 'nov': 11, 'november': 11, 'dec': 12, 'december': 12,
    }
    for name, mnum in month_names.items():
        if re.search(rf'\b{name}\b', s):
            return [mnum], list(range(1, mnum + 1))

    # Fallback YYYY-MM o MM-YYYY
    m = re.search(r'\b(20\d{2})[-_./](\d{1,2})\b', s) or re.search(r'\b(\d{1,2})[-_./](20\d{2})\b', s)
    if m:
        # Determinar cuál grupo es el mes
        g1, g2 = m.group(1), m.group(2)
        mnum = int(g2) if len(g1) == 4 else int(g1)
        if 1 <= mnum <= 12:
            return [mnum], list(range(1, mnum + 1))

    return None, None


def _ingest_cws_sheet(ws, filepath: Path, sheet_info: Dict, target_period: Optional[str] = None) -> IngestedFile:
    """
    Parser para formato CWS 12-month budget (Bridge at the Blockyard).

    Layout:
      - Row 1, col A: 'Bridge at the Blockyard (a0111048)' (building name)
      - Row 3, col A: 'Period = Jan 2026-Dec 2026'
      - Row 4, col A: 'Book = Accrual ; Tree = cws_res_is_'
      - Row 5: cols C-N = meses Jan-Dec, col O = 'Total'
      - Row 6+: data rows con código GL en col A, descripción en col B,
                12 meses en cols C-N, total anual en col O.
      - Section/total rows: col A vacía, col B con texto.

    Si target_period viene con un Q o mes (ej. 'Q1 2026', 'Mar 2026'), el
    parser suma los meses correspondientes para budget_current y budget_ytd.
    Si no, usa el TOTAL ANUAL como budget_current.

    actual_current y actual_ytd son SIEMPRE None — este archivo es budget-only.
    Los actuals deben venir de un archivo separado y la pipeline merge por
    account_code (Phase 2 — pendiente).
    """
    # Building & año
    building = ""
    v1 = ws.cell(row=1, column=1).value
    if v1:
        building = str(v1).strip()

    file_year = None
    v3 = ws.cell(row=3, column=1).value
    if v3:
        m = re.search(r'\b(20\d{2})\b', str(v3))
        if m:
            file_year = int(m.group(1))

    # Determinar columnas a sumar según target_period
    # Cols 3-14 = Jan-Dec, col 15 = Total
    current_months, ytd_months = (None, None)
    if file_year is not None:
        current_months, ytd_months = _parse_cws_target_period(target_period, file_year)

    # Si no hay target_period o no se pudo parsear → annual total (col 15)
    use_annual = current_months is None

    if use_annual:
        period_label = f"Annual {file_year}" if file_year else "Annual"
    else:
        period_label = str(target_period).strip()

    # Layout fijo
    desc_col = 2  # col B
    acct_col = 1  # col A
    month_first_col = 3  # Jan en col C
    annual_total_col = 15  # col O
    header_row = 5

    rows = []
    section_totals = []
    partner_comments = []

    max_row = ws.max_row or 250
    for i in range(header_row + 1, max_row + 1):
        c1 = ws.cell(row=i, column=acct_col).value
        c2 = ws.cell(row=i, column=desc_col).value

        # Saltar filas completamente vacías
        if c1 is None and c2 is None:
            continue

        acct_str = str(c1).strip() if c1 is not None else ""
        desc_str = str(c2).strip() if c2 is not None else ""

        is_acct_code = bool(re.match(r'^\d{4}-\d{4}$', acct_str))
        desc_upper = desc_str.upper()

        # Determinar tipo de fila
        if is_acct_code and desc_str:
            row_type = 'data'
        elif desc_str and not acct_str:
            # section o total — totales empiezan con TOTAL/NET/INCOME
            if (desc_upper.startswith('TOTAL ') or desc_upper.startswith('NET ')
                    or desc_upper.startswith('INCOME (LOSS)')):
                row_type = 'total'
            else:
                row_type = 'section'
        else:
            row_type = 'label'

        # Calcular budget_current y budget_ytd según target_period
        budget_curr_val = None
        budget_ytd_val = None
        if row_type in ('data', 'total'):
            if use_annual:
                v = ws.cell(row=i, column=annual_total_col).value
                budget_curr_val = _to_float(v)
                budget_ytd_val = budget_curr_val  # annual = ytd
            else:
                # Sumar meses específicos
                cur_sum = 0.0
                cur_any = False
                for m in current_months:
                    v = ws.cell(row=i, column=month_first_col + m - 1).value
                    fv = _to_float(v)
                    if fv is not None:
                        cur_sum += fv
                        cur_any = True
                ytd_sum = 0.0
                ytd_any = False
                for m in ytd_months:
                    v = ws.cell(row=i, column=month_first_col + m - 1).value
                    fv = _to_float(v)
                    if fv is not None:
                        ytd_sum += fv
                        ytd_any = True
                budget_curr_val = cur_sum if cur_any else None
                budget_ytd_val = ytd_sum if ytd_any else None

        rows.append({
            'row_num': i,
            'description': desc_str,
            'account': acct_str if is_acct_code else "",
            'budget_current': budget_curr_val,
            'actual_current': None,           # budget-only file
            'budget_ytd': budget_ytd_val,
            'actual_ytd': None,
            'row_type': row_type,
            'source_file': filepath.name,
        })

        if row_type == 'total':
            section_totals.append({
                'description': desc_str,
                'budget_current': budget_curr_val,
                'actual_current': None,
                'budget_ytd': budget_ytd_val,
                'actual_ytd': None,
                'row_num': i,
            })

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'header_row': header_row,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': True,
            'sheet_used': sheet_info.get('income_statement'),
            'comments_sheet': None,
            'format': 'CWS',
            'is_budget_reference': True,        # marker para Phase 2 (merge con actuals)
            'file_year': file_year,
            'extracted_period': period_label,
            'extracted_months_current': current_months,
            'extracted_months_ytd': ytd_months,
            'column_map': {
                'desc_col': desc_col, 'acct_col': acct_col,
                'month_first_col': month_first_col, 'annual_total_col': annual_total_col,
            },
        }
    )


# ─────────────────────────────────────────
# FORMATO YSI / YARDI BUDGET FORECAST (The Wilcox y similares)
# ─────────────────────────────────────────

def _is_ysi_format(ws) -> bool:
    """
    Detecta el formato YSI/Yardi Budget Forecast (usado por The Wilcox).
    Firma: celda con "Tree = ysi_bf", "Tree = ysi_cf" o "Tree = tmbr_cashflow"
    en las primeras 6 filas. Wilcox puede llegar con cualquiera de estas trees
    según el reporte que envíe el socio (cash flow vs budget forecast).

    Detección estricta: solo activa si ese marker exacto está presente. Comparte
    layout de columnas con CMC (PTD/YTD + Note en col 12) pero las secciones
    traen código de cuenta (410000 REVENUE), mientras que CMC no.
    """
    YSI_TREES = ('ysi_bf', 'ysi_cf', 'tmbr_cashflow', 'tmbr_bf', 'tmbr_cf')
    max_r = min(ws.max_row or 0, 8)
    max_c = min(ws.max_column or 0, 6)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            vs = str(v).lower()
            # "tree = X" (con o sin espacios)
            if 'tree' in vs and '=' in vs:
                for t in YSI_TREES:
                    if t in vs:
                        return True
    return False


def _detect_ysi_building(ws) -> str:
    """
    En YSI el nombre del edificio vive en A1 como 'The Wilcox (pn10680)'.
    Devuelve el nombre tal cual (con el PN entre paréntesis).
    """
    v = ws.cell(row=1, column=1).value
    if v:
        s = str(v).strip()
        if s and len(s) >= 3:
            return s
    return ""


def _detect_ysi_period(ws) -> str:
    """Período YSI vive en A3 como 'Period = Dec 2025'."""
    v = ws.cell(row=3, column=1).value
    if v:
        s = str(v).strip()
        # Extraer "Dec 2025" de "Period = Dec 2025"
        m = re.search(r'period\s*=\s*(.+)', s, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        return s
    return ""


def _ingest_ysi_sheet(ws, filepath: Path, sheet_info: Dict) -> IngestedFile:
    """
    Parser específico para formato YSI/Yardi Budget Forecast (The Wilcox).

    Layout (12 cols en ysi_bf, 11 en ysi_cf sin Notes):
    - Row 1: nombre del edificio + PN ("The Wilcox (pn10680)")
    - Row 2: "Budget Comparison"
    - Row 3: "Period = Dec 2025"
    - Row 4: "Book = Accrual,Budget ; Tree = ysi_bf"
    - Row 5 (headers): col3=PTD Actual, col4=PTD Budget, col5=Variance, col6=%Var,
                       col7=YTD Actual, col8=YTD Budget, col9=Variance, col10=%Var,
                       col11=Annual, col12=Note
    - Col 1: código de cuenta 6-dígitos (ej '410075'). Las secciones TAMBIÉN tienen
             código (ej '410000 REVENUE', '410001 RENTAL INCOME').
    - Col 2: descripción (con indentación en L3 — espacios al inicio)

    Section/total heuristic:
    - row sin valores numéricos en PTD/YTD → section header
    - row con "TOTAL" en desc o el código termina en '999' / '099' → total
    - row con valores + no es total → data (L3)

    Filtros:
    - Cortar en 'NET INCOME (LOSS)' (989999) — todo lo que sigue es balance sheet
      (ADJUSTMENTS: Assets, Liabilities & Equity), no debe entrar al P&L.
    """
    building = _detect_ysi_building(ws)
    period_label = _detect_ysi_period(ws)

    header_row = 5
    acct_col = 1
    desc_col = 2
    actual_curr = 3
    budget_curr = 4
    actual_ytd = 7
    budget_ytd = 8
    notes_col = 12  # ysi_cf no lo tiene, pero leer col 12 devuelve None

    rows = []
    partner_comments = []
    section_totals = []

    max_row = ws.max_row or 400
    cutoff_reached = False

    for i in range(header_row + 1, max_row + 1):
        if cutoff_reached:
            break

        acct_val = ws.cell(row=i, column=acct_col).value
        desc_val = ws.cell(row=i, column=desc_col).value
        budget_c = ws.cell(row=i, column=budget_curr).value
        actual_c = ws.cell(row=i, column=actual_curr).value
        budget_y = ws.cell(row=i, column=budget_ytd).value
        actual_y = ws.cell(row=i, column=actual_ytd).value
        notes_val = ws.cell(row=i, column=notes_col).value

        # Saltar filas completamente vacías
        if acct_val is None and desc_val is None and actual_c is None and budget_c is None:
            continue

        acct_str = str(acct_val).strip() if acct_val is not None else ""
        desc_str = str(desc_val).strip() if desc_val is not None else ""

        # Cortar en ADJUSTMENTS / NET INCOME (LOSS). Todo debajo es balance sheet.
        # 989999 = NET INCOME (LOSS) es la última fila P&L; procesarla y cortar.
        cutoff_after_this = (acct_str == '989999' or
                              desc_str.upper() == 'NET INCOME (LOSS)' or
                              desc_str.upper() == 'ADJUSTMENTS')

        if desc_str.upper() == 'ADJUSTMENTS' and not isinstance(actual_c, (int, float)):
            # Balance sheet label — no procesar esta fila ni las siguientes
            break

        has_values = isinstance(budget_c, (int, float)) or isinstance(actual_c, (int, float))

        # Detectar tipo de fila
        desc_upper = desc_str.upper().strip()
        is_total = (
            desc_upper.startswith('TOTAL ') or
            desc_upper in ('NET OPERATING INCOME', 'NET INCOME', 'NET INCOME (LOSS)',
                           'GROSS POTENTIAL RENT')
        )
        is_section = (
            desc_str != "" and not has_values and not is_total
        )

        if is_total and has_values:
            row_type = 'total'
        elif is_section:
            row_type = 'section'
        elif has_values:
            row_type = 'data'
        else:
            row_type = 'label'

        rows.append({
            'row_num': i,
            'description': desc_str,
            'account': acct_str if row_type == 'data' else "",
            'budget_current': _to_float(budget_c),
            'actual_current': _to_float(actual_c),
            'budget_ytd': _to_float(budget_y),
            'actual_ytd': _to_float(actual_y),
            'row_type': row_type,
            'source_file': filepath.name,
        })

        if notes_val and str(notes_val).strip():
            partner_comments.append({
                'account_code': acct_str,
                'description': desc_str,
                'comment_text': str(notes_val).strip(),
                'source_file': filepath.name,
                'row_num': i,
            })

        if row_type == 'total':
            section_totals.append({
                'description': desc_str,
                'budget_current': _to_float(budget_c),
                'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y),
                'actual_ytd': _to_float(actual_y),
                'row_num': i,
            })

        if cutoff_after_this:
            cutoff_reached = True

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'header_row': header_row,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': True,
            'sheet_used': sheet_info['income_statement'],
            'comments_sheet': sheet_info.get('comments_sheet'),
            'sheets_analyzed': len(sheet_info.get('all_scores', {}).get('is', {})) or 1,
            'format': 'YSI',
            'column_map': {
                'desc_col': desc_col, 'acct_col': acct_col,
                'budget_curr': budget_curr, 'actual_curr': actual_curr,
                'budget_ytd': budget_ytd, 'actual_ytd': actual_ytd,
                'notes_col': notes_col,
            },
        }
    )


# ─────────────────────────────────────────
# FORMATO JDE / LINCOLN PROPERTY COMPANY (Walnut Street Wellesley y similares)
# ─────────────────────────────────────────

def _is_jde_format(ws) -> bool:
    """
    Detecta el formato JDE (J.D. Edwards) usado por Lincoln Property Company.
    Firma: A1 contiene "Database: PROJ" o similar; el header de columnas está
    distribuido en múltiples filas (Actual en fila 4, período en fila 5) y
    el layout es: col A=código, col C=descripción, col E=Actual, col F=Budget.
    """
    max_r = min(ws.max_row or 0, 6)
    max_c = min(ws.max_column or 0, 8)
    markers = 0
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            vs = str(v).lower()
            if 'database: proj' in vs or 'report id' in vs:
                markers += 1
            if 'msde_data' in vs or 'format id' in vs:
                markers += 1
            if 'comparative income statement' in vs or 'monthly income statement - jde' in vs:
                markers += 2
            if 'lincoln property company' in vs:
                markers += 2
    return markers >= 2


def _detect_jde_building(ws) -> str:
    """
    En JDE el nombre del edificio vive en col D o E de las filas 1-3, como
    "Walnut Street Wellesley Owner LLC\nLincoln Property Company".
    Devuelve la primera línea que parezca un nombre de activo real.
    """
    for r in range(1, 5):
        for c in range(3, 9):
            v = ws.cell(row=r, column=c).value
            if not v:
                continue
            for line in str(v).split('\n'):
                line = line.strip()
                if len(line) < 5:
                    continue
                lower = line.lower()
                # Skip labels/metadata noise
                skip = ('database', 'report id', 'format id', 'msde', 'jde',
                        'accrual', 'comparative income statement', 'monthly income statement',
                        'lincoln property company', 'property company', 'actual', 'budget',
                        'period', 'thru', 'variance', 'through')
                if any(k in lower for k in skip):
                    continue
                # Prefer asset-looking names (LLC, LP, Owner) or plain property names
                if line.endswith(('LLC', 'LP', 'Owner', 'Inc', 'Inc.')) or line[0].isalpha():
                    return line
    return ""


def _detect_jde_period(ws) -> str:
    """Período JDE vive tipicamente en row 5 col E o F como 'Mar 2026'."""
    for r in range(3, 7):
        for c in range(4, 9):
            v = ws.cell(row=r, column=c).value
            if not v:
                continue
            s = str(v).strip()
            # Match patterns "Mar 2026", "March 2026", "01/2026", etc.
            if re.search(r'\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s*20\d{2}\b', s, re.IGNORECASE):
                return s.split('\n')[0].strip()
            if re.search(r'\b\d{1,2}[/\-]20\d{2}\b', s):
                return s.split('\n')[0].strip()
    return ""


def _ingest_jde_sheet(ws, filepath: Path, sheet_info: Dict) -> IngestedFile:
    """
    Parser específico para el formato JDE (Lincoln Property Company).

    Layout JDE:
    - Row 1-3: metadata (Database, Report ID, Format ID, nombre del edificio en col D row 2)
    - Row 4: 'Actual' (col E) y 'Current Period\\nBudget' (col F)
    - Row 5: período ('Mar 2026')
    - Row 6+: filas de datos
      - Col A: código de cuenta (5 dígitos, ej '40120') — vacía en totales/secciones
      - Col C: descripción — puede tener "SECCION\\nDescripción cuenta" en la primera
        fila de cada sección. En totales: "TOTAL <SECCION>" o "NET OPERATING INCOME".
      - Col E: Actual Current Period
      - Col F: Budget Current Period

    No trae columna YTD en este layout — el Q/YTD se construye sumando los meses
    que ingresa el usuario (la lógica ya existe en `build_financial_structure`).
    """
    building = _detect_jde_building(ws)
    period_label = _detect_jde_period(ws)

    acct_col = 1
    desc_col = 3
    actual_curr = 5
    budget_curr = 6

    rows = []
    section_totals = []
    partner_comments = []  # JDE no trae variance notes inline

    max_row = ws.max_row or 200
    for i in range(6, max_row + 1):  # Row 6 en adelante son datos
        acct_val = ws.cell(row=i, column=acct_col).value
        desc_val = ws.cell(row=i, column=desc_col).value
        actual_c = ws.cell(row=i, column=actual_curr).value
        budget_c = ws.cell(row=i, column=budget_curr).value

        # Saltar filas completamente vacías
        if acct_val is None and desc_val is None and actual_c is None and budget_c is None:
            continue

        acct_raw = str(acct_val).strip() if acct_val is not None else ""
        desc_raw = str(desc_val).strip() if desc_val is not None else ""

        # En JDE algunas secciones (REVENUES:, EXPENSES:, OPERATING EXPENSES,
        # TAXES & INSURANCE, NON OPERATING EXPENSES, CAPITAL EXPENSES) viven en col A
        # (texto, sin código numérico) o en col C (solo texto, sin código ni valores).
        acct_is_numeric = bool(re.match(r'^\d{4,6}$', acct_raw.replace('.', '').replace('-', '')))

        # Caso 1: col A texto (no código) -> section header
        if acct_raw and not acct_is_numeric and not isinstance(actual_c, (int, float)) and not isinstance(budget_c, (int, float)):
            rows.append({
                'row_num': i, 'description': acct_raw, 'account': "",
                'budget_current': None, 'actual_current': None,
                'budget_ytd': None, 'actual_ytd': None,
                'row_type': 'section', 'source_file': filepath.name,
            })
            continue

        # Descripción puede contener "SECCION\nDescripción cuenta" — emitir section virtual
        desc_lines = [ln.strip() for ln in desc_raw.split('\n') if ln.strip()]
        if len(desc_lines) >= 2 and acct_is_numeric:
            # La primera línea es un header de sección implícito; la segunda es la descripción real
            rows.append({
                'row_num': i, 'description': desc_lines[0], 'account': "",
                'budget_current': None, 'actual_current': None,
                'budget_ytd': None, 'actual_ytd': None,
                'row_type': 'section', 'source_file': filepath.name,
            })
            desc_final = desc_lines[1]
        else:
            desc_final = desc_lines[0] if desc_lines else ""

        has_values = isinstance(actual_c, (int, float)) or isinstance(budget_c, (int, float))

        # Detectar tipo de fila
        lower = desc_final.lower()
        is_total = (
            lower.startswith('total ') or
            lower in ('net operating income', 'net income', 'net income less capital')
        )
        is_section_header = (
            desc_final and not acct_is_numeric and not has_values and not is_total
        )

        if is_total and has_values:
            row_type = 'total'
            acct_out = ""
        elif acct_is_numeric and has_values:
            row_type = 'data'
            acct_out = acct_raw
        elif is_section_header:
            row_type = 'section'
            acct_out = ""
        else:
            row_type = 'label'
            acct_out = acct_raw if acct_is_numeric else ""

        rows.append({
            'row_num': i,
            'description': desc_final,
            'account': acct_out,
            'budget_current': _to_float(budget_c),
            'actual_current': _to_float(actual_c),
            'budget_ytd': None,
            'actual_ytd': None,
            'row_type': row_type,
            'source_file': filepath.name,
        })

        if row_type == 'total':
            section_totals.append({
                'description': desc_final,
                'budget_current': _to_float(budget_c),
                'actual_current': _to_float(actual_c),
                'budget_ytd': None, 'actual_ytd': None,
                'row_num': i,
            })

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'header_row': 4,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': False,
            'sheet_used': sheet_info['income_statement'],
            'comments_sheet': sheet_info.get('comments_sheet'),
            'sheets_analyzed': len(sheet_info.get('all_scores', {}).get('is', {})) or 1,
            'format': 'JDE',
            'column_map': {
                'desc_col': desc_col, 'acct_col': acct_col,
                'budget_curr': budget_curr, 'actual_curr': actual_curr,
                'budget_ytd': None, 'actual_ytd': None,
            },
        }
    )


# ─────────────────────────────────────────
# FORMATO CALLAN (14th & Callan Street JV LLC y similares)
# ─────────────────────────────────────────

def _is_callan_format(ws) -> bool:
    """
    Detecta formato Callan (14th & Callan Street JV LLC y similares).
    Señal primaria: el workbook contiene una hoja con 'income statement bud' en el nombre.
    Señal secundaria: row 6 tiene Account/Budget/Actual en cols B/C/D con super-header
    'Current' en row 5 col C.
    """
    try:
        for sname in ws.parent.sheetnames:
            sl = sname.lower()
            if 'income statement bud' in sl and 'act' in sl:
                return True
    except Exception:
        pass

    b6 = ws.cell(row=6, column=2).value
    c6 = ws.cell(row=6, column=3).value
    d6 = ws.cell(row=6, column=4).value
    c5 = ws.cell(row=5, column=3).value
    if b6 and c6 and d6 and c5:
        if (str(b6).strip().lower() == 'account' and
                str(c6).strip().lower() == 'budget' and
                str(d6).strip().lower() == 'actual' and
                str(c5).strip().lower() == 'current'):
            return True
    return False


def _ingest_callan_sheet(ws, filepath: Path, sheet_info: Dict) -> IngestedFile:
    """
    Parser para el formato Callan (Sares Regis / 14th & Callan Street JV LLC).

    Layout (sheet '03 Income statement bud v act'):
      - Row 1, col A: nombre del activo ("14th & Callan Street JV LLC")
      - Row 3, col A: período ("For the One Month Ending January 31, 2026")
      - Row 5: super-headers col C='Current', col F='YTD'
      - Row 6 (headers): col A=vacío, col B=Account, col C=Budget, col D=Actual,
                         col E=Variance, col F=Budget(YTD), col G=Actual(YTD),
                         col H=Variance(YTD), col I=Variance Notes
      - Row 7+: datos

    Tipos de fila:
      - section: col A texto, col B None (sin cuenta), sin valores numéricos
      - data:    col A descripción, col B "XXXX-XXXX", con valores numéricos
      - total:   col A "Total ..." o "Profit (loss)", col B "" (string vacío), con valores
    """
    # Encontrar la hoja correcta (budget vs actual) en el workbook
    try:
        wb = ws.parent
        for sname in wb.sheetnames:
            if 'income statement bud' in sname.lower() and 'act' in sname.lower():
                ws = wb[sname]
                sheet_info = {**sheet_info, 'income_statement': sname}
                break
    except Exception:
        pass

    # Building: row 1, col A
    building = ""
    v = ws.cell(row=1, column=1).value
    if v:
        building = str(v).strip()

    # Período: extraer de row 3 col A → "For the One Month Ending January 31, 2026"
    period_label = ""
    v3 = ws.cell(row=3, column=1).value
    if v3:
        s = str(v3).strip()
        m = re.search(r'ending\s+(.+)$', s, re.IGNORECASE)
        if m:
            period_label = m.group(1).strip()
        else:
            period_label = s

    # Fallback: TOC sheet tiene datetime en row 3 col A
    if not period_label:
        import datetime as _dt
        try:
            wb2 = ws.parent
            for toc_name in wb2.sheetnames:
                if 'table of contents' in toc_name.lower() or toc_name.startswith('00'):
                    toc_ws = wb2[toc_name]
                    toc_v = toc_ws.cell(row=3, column=1).value
                    if isinstance(toc_v, (_dt.datetime, _dt.date)):
                        period_label = toc_v.strftime("%Y-%m-%d")
                    break
        except Exception:
            pass

    # Fallback: extraer del nombre de archivo "XX-YYYY"
    if not period_label:
        mf = re.search(r'(\d{2})-(\d{4})', filepath.name)
        if mf:
            mo, yr = int(mf.group(1)), int(mf.group(2))
            if 1 <= mo <= 12:
                period_label = f"{yr}-{mo:02d}"

    # Layout de columnas fijo
    desc_col = 1
    acct_col = 2
    budget_curr = 3
    actual_curr = 4
    # col 5 = variance current (derivado, no se ingesta)
    budget_ytd = 6
    actual_ytd = 7
    # col 8 = variance ytd (derivado)
    notes_col = 9
    header_row = 6

    rows = []
    partner_comments = []
    section_totals = []

    max_row = ws.max_row or 300
    for i in range(header_row + 1, max_row + 1):
        desc_val = ws.cell(row=i, column=desc_col).value
        acct_val = ws.cell(row=i, column=acct_col).value
        budget_c = ws.cell(row=i, column=budget_curr).value
        actual_c = ws.cell(row=i, column=actual_curr).value
        budget_y = ws.cell(row=i, column=budget_ytd).value
        actual_y = ws.cell(row=i, column=actual_ytd).value
        notes_val = ws.cell(row=i, column=notes_col).value

        if desc_val is None and acct_val is None and budget_c is None and actual_c is None:
            continue

        desc_str = str(desc_val).strip() if desc_val is not None else ""
        # acct_val puede ser None (section) o "" (total) o "XXXX-XXXX" (data)
        acct_str = str(acct_val).strip() if acct_val is not None else None

        if not desc_str:
            continue

        has_values = isinstance(budget_c, (int, float)) or isinstance(actual_c, (int, float))
        is_acct_code = acct_str is not None and bool(re.match(r'^\d{3,5}-\d{3,4}$', acct_str))

        # Totals: acct es "" (string vacío) con valores, o desc empieza con "Total" / es P&L bottom
        is_total = (
            has_values and acct_str == "" and (
                desc_str.lower().startswith('total ') or
                desc_str.lower() in ('profit (loss)', 'net operating income', 'net income')
            )
        )
        # Sections: acct es None, sin valores
        is_section = acct_str is None and not has_values

        if is_total:
            row_type = 'total'
            acct_out = ""
        elif is_section:
            row_type = 'section'
            acct_out = ""
        elif is_acct_code and has_values:
            row_type = 'data'
            acct_out = acct_str
        else:
            row_type = 'label'
            acct_out = acct_str if is_acct_code else ""

        rows.append({
            'row_num': i,
            'description': desc_str,
            'account': acct_out,
            'budget_current': _to_float(budget_c),
            'actual_current': _to_float(actual_c),
            'budget_ytd': _to_float(budget_y),
            'actual_ytd': _to_float(actual_y),
            'row_type': row_type,
            'source_file': filepath.name,
        })

        notes_str = str(notes_val).strip() if notes_val is not None else ""
        if notes_str and notes_str not in ('-', 'Variance Notes'):
            partner_comments.append({
                'account_code': acct_out,
                'description': desc_str,
                'comment_text': notes_str,
                'source_file': filepath.name,
                'row_num': i,
            })

        if row_type == 'total':
            section_totals.append({
                'description': desc_str,
                'budget_current': _to_float(budget_c),
                'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y),
                'actual_ytd': _to_float(actual_y),
                'row_num': i,
            })

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'header_row': header_row,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': True,
            'sheet_used': sheet_info.get('income_statement'),
            'comments_sheet': sheet_info.get('comments_sheet'),
            'format': 'Callan',
            'column_map': {
                'desc_col': desc_col, 'acct_col': acct_col,
                'budget_curr': budget_curr, 'actual_curr': actual_curr,
                'budget_ytd': budget_ytd, 'actual_ytd': actual_ytd,
                'notes_col': notes_col,
            },
        }
    )


# ─────────────────────────────────────────
# FORMATO PDF — NWEP MONTHLY REPORT (Walnut Street Wellesley)
# ─────────────────────────────────────────
#
# El socio (Lincoln Property Company) envía mensualmente un PDF con:
#   - Page 1: SUMMARY a nivel L2 (mismo formato que el Excel 'EERR' / 'Resultado
#             Walnut') — descripciones tipo 'Salaries', 'Utilities', etc.
#   - Pages 2+: JDE detail con códigos GL (mismo formato que el Excel 'Budget
#               Comparison Propiedad') — cuentas como '40120', '51013', etc.
#
# Usamos las pages 2+ (detalle JDE) porque tienen códigos GL → 100% match con
# el catálogo de Walnut Street Wellesley (84 cuentas).


def _is_nwep_pdf_format(filepath: Path) -> bool:
    """
    Detecta PDF de monthly report de NWEP / Walnut Street Wellesley.
    Firma: page 1 contiene 'SUMMARY Newton Wellesley' Y page 2+ contiene
    'WalnutStreetWellesley' (marker JDE detail).
    """
    if filepath.suffix.lower() != '.pdf':
        return False
    try:
        import pdfplumber
        with pdfplumber.open(str(filepath)) as pdf:
            if len(pdf.pages) < 2:
                return False
            p1 = pdf.pages[0].extract_text() or ''
            p2 = pdf.pages[1].extract_text() or ''
            return ('SUMMARY Newton Wellesley' in p1 and
                    'WalnutStreetWellesley' in p2.replace(' ', ''))
    except Exception:
        return False


# Regex para data rows del JDE detail. Formato:
#   <code5> <description (sin espacios)> <actual> <budget> <variance> <%var>
#                                        <ytd_actual> <ytd_budget> <ytd_var> <%var_ytd>
# Las descripciones vienen squeezed por el text-extract del PDF (RentalInc.-Commercial),
# pero los números siempre van separados por espacios.
_PDF_NUM = r'[\(\-]?[\d,]+\.\d+\)?'
_PDF_PCT = r'-?[\d,]+\.\d+%?'
_PDF_DATA_RE = re.compile(
    rf'^(\d{{5}})\s+(.+?)\s+'
    rf'({_PDF_NUM})\s+({_PDF_NUM})\s+({_PDF_NUM})\s+({_PDF_PCT})\s+'
    rf'({_PDF_NUM})\s+({_PDF_NUM})\s+({_PDF_NUM})\s*({_PDF_PCT})?$'
)
# Regex para totales del JDE: empieza con TOTAL + 8 números (sin código)
_PDF_TOTAL_RE = re.compile(
    rf'^(TOTAL[\w&\s]+?)\s+'
    rf'({_PDF_NUM})\s+({_PDF_NUM})\s+({_PDF_NUM})\s+({_PDF_PCT})\s+'
    rf'({_PDF_NUM})\s+({_PDF_NUM})\s+({_PDF_NUM})\s*({_PDF_PCT})?$'
)


def _pdf_to_float(s: str) -> Optional[float]:
    """Convierte string del PDF (con paréntesis para negativos) a float."""
    if s is None:
        return None
    s = str(s).strip().replace(',', '').replace('$', '').replace('%', '')
    if not s:
        return None
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _ingest_nwep_pdf(filepath: Path) -> IngestedFile:
    """
    Parser para PDF monthly report de NWEP / Walnut Street Wellesley.

    Extrae el detalle JDE de pages 2+ (no el SUMMARY de page 1) porque el
    detalle tiene códigos GL que matchean 100% con el catálogo Walnut.

    Layout pages 2+:
      - Header rows con metadata (Database, ReportID, FormatID, fechas)
      - Section headers en MAYÚSCULAS (REVENUES:, RENTALINCOME, OTHERINCOME,
        OPERATINGEXPENSES, TAXES&INSURANCE, ADMINISTRATION, etc.)
      - Data rows: <code5> <desc> <actual> <budget> <var> <%var> <YTDx4>
      - Total rows: TOTAL<X> + 8 números
    """
    import pdfplumber
    import datetime as _dt

    rows = []
    section_totals = []
    partner_comments = []
    period_label = ""
    building = "Walnut Street Wellesley Owner LLC"

    with pdfplumber.open(str(filepath)) as pdf:
        # Período: page 1 line 1 (formato M/D/YYYY)
        p1 = pdf.pages[0].extract_text() or ''
        p1_lines = p1.split('\n')
        if len(p1_lines) > 1:
            m = re.search(r'(\d{1,2})/(\d{1,2})/(20\d{2})', p1_lines[1])
            if m:
                mo, day, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
                period_label = f"{yr}-{mo:02d}-{day:02d}"

        # Page 1 SUMMARY trae NOI y Net Income que el JDE detail (pages 2+) NO tiene
        # como totales — los extraemos para que figuren como L1 lines explícitos.
        # Formato page 1: '<desc> <PTD_actual> <PTD_budget> <var> <%var> <YTD_actual>...'
        SUMMARY_L1_TOTALS = ('Net Operating Income', 'Net Income')
        summary_totals_extracted = []
        for line in p1_lines:
            line_s = line.strip()
            for label in SUMMARY_L1_TOTALS:
                if line_s.startswith(label + ' '):
                    # Extraer los primeros 2 numéricos como actual_current / budget_current
                    rest = line_s[len(label):].strip()
                    nums = re.findall(rf'{_PDF_NUM}', rest)
                    if len(nums) >= 6:  # PTD actual, PTD budget, var, %var, YTD actual, YTD budget
                        summary_totals_extracted.append({
                            'description': label,
                            'actual_current': _pdf_to_float(nums[0]),
                            'budget_current': _pdf_to_float(nums[1]),
                            'actual_ytd': _pdf_to_float(nums[4]),
                            'budget_ytd': _pdf_to_float(nums[5]),
                        })
                    break

        # Procesar pages 2+ (JDE detail)
        row_num = 0
        for page in pdf.pages[1:]:
            text = page.extract_text() or ''
            for line in text.split('\n'):
                line = line.strip()
                if not line:
                    continue

                # Skip header lines (Database, PROJ, ReportID, etc.)
                if any(line.startswith(prefix) for prefix in
                       ('Database:', 'PROJ:', 'ReportID:', 'FormatID:',
                        'LincolnPropertyCompany', 'Accrual',
                        'CurrentPeriod', 'Actual', 'Thru:')):
                    continue

                row_num += 1

                # Total row
                m_tot = _PDF_TOTAL_RE.match(line)
                if m_tot:
                    desc = m_tot.group(1).strip()
                    actual_c = _pdf_to_float(m_tot.group(2))
                    budget_c = _pdf_to_float(m_tot.group(3))
                    actual_y = _pdf_to_float(m_tot.group(6))
                    budget_y = _pdf_to_float(m_tot.group(7))
                    rows.append({
                        'row_num': row_num, 'description': desc, 'account': "",
                        'budget_current': budget_c, 'actual_current': actual_c,
                        'budget_ytd': budget_y, 'actual_ytd': actual_y,
                        'row_type': 'total', 'source_file': filepath.name,
                    })
                    section_totals.append({
                        'description': desc,
                        'budget_current': budget_c, 'actual_current': actual_c,
                        'budget_ytd': budget_y, 'actual_ytd': actual_y,
                        'row_num': row_num,
                    })
                    continue

                # Data row
                m_dat = _PDF_DATA_RE.match(line)
                if m_dat:
                    code = m_dat.group(1).strip()
                    desc = m_dat.group(2).strip()
                    actual_c = _pdf_to_float(m_dat.group(3))
                    budget_c = _pdf_to_float(m_dat.group(4))
                    actual_y = _pdf_to_float(m_dat.group(7))
                    budget_y = _pdf_to_float(m_dat.group(8))
                    rows.append({
                        'row_num': row_num, 'description': desc, 'account': code,
                        'budget_current': budget_c, 'actual_current': actual_c,
                        'budget_ytd': budget_y, 'actual_ytd': actual_y,
                        'row_type': 'data', 'source_file': filepath.name,
                    })
                    continue

                # Section header: línea en mayúsculas, sin números
                if (re.match(r'^[A-Z][A-Z&\s\-/:]+$', line)
                        and 'PAGE' not in line.upper()
                        and 'PERIOD' not in line.upper()):
                    rows.append({
                        'row_num': row_num, 'description': line.rstrip(':'), 'account': "",
                        'budget_current': None, 'actual_current': None,
                        'budget_ytd': None, 'actual_ytd': None,
                        'row_type': 'section', 'source_file': filepath.name,
                    })
                    continue

                # Otherwise: label
                rows.append({
                    'row_num': row_num, 'description': line, 'account': "",
                    'budget_current': None, 'actual_current': None,
                    'budget_ytd': None, 'actual_ytd': None,
                    'row_type': 'label', 'source_file': filepath.name,
                })

        # Append NOI / Net Income totals extraídos de page 1 SUMMARY
        for t in summary_totals_extracted:
            row_num += 1
            rows.append({
                'row_num': row_num,
                'description': t['description'],
                'account': "",
                'budget_current': t['budget_current'],
                'actual_current': t['actual_current'],
                'budget_ytd': t['budget_ytd'],
                'actual_ytd': t['actual_ytd'],
                'row_type': 'total',
                'source_file': filepath.name,
            })
            section_totals.append({
                'description': t['description'],
                'budget_current': t['budget_current'],
                'actual_current': t['actual_current'],
                'budget_ytd': t['budget_ytd'],
                'actual_ytd': t['actual_ytd'],
                'row_num': row_num,
            })

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'header_row': None,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': True,
            'sheet_used': None,
            'comments_sheet': None,
            'format': 'NWEP-PDF',
            'column_map': None,
        }
    )


# ─────────────────────────────────────────
# FORMATO PDF — J 501 ESTATES CLOSE PACKAGE
# ─────────────────────────────────────────
#
# Close Package mensual de Jefferson Apartment Group (JAG) para J 501 Estates.
# El PDF contiene multiples reportes; usamos:
#   - "Budget Comparison Report": mismas filas que el Excel BCR
#     (codigos GL formato XXXX-XXXX, columnas Actual / Budget / Diff / %Var
#     para Month Ending y Year To Date).
#   - "Variance Report W/Notes": comentarios de los socios por cuenta
#     (columna "Variance Comments").


def _extract_j501_variance_comments(filepath: Path) -> List[Dict]:
    """
    Extrae comentarios del socio del 'Variance Report W/Notes' del Close
    Package de J 501 Estates.

    pdfplumber ordena el texto por coordenada Y, lo que en este PDF hace que
    el comentario aparezca en la LINEA ANTES (a veces tambien ABAJO o pegado)
    de la fila de la cuenta — por eso recolectamos linea-arriba + glued + linea-abajo.

    Devuelve dicts con las keys que espera native_structure.py:
      account_code, description, comment_text, source_file.
    """
    try:
        import pdfplumber
    except ImportError:
        return []

    NUM_TOK = re.compile(r"^-?\(?[\d,]+(?:\.\d+)?\)?$|^\(-?[\d,]+(?:\.\d+)?\)$")
    ACCT_RX = re.compile(r"^\s*(\d{4}-\d{4})\s*-?\s*(.*)$")

    def _to_money(s):
        if s is None:
            return None
        s = str(s).strip().replace(",", "").replace("$", "")
        if s in ("", "-"):
            return 0.0
        neg = s.startswith("(") and s.endswith(")")
        s = s.strip("()")
        if s.startswith("-"):
            neg = True
            s = s[1:]
        try:
            v = float(s)
        except ValueError:
            return None
        return -v if neg else v

    def _is_comment_line(s):
        s = (s or "").strip()
        if not s:
            return False
        if ACCT_RX.match(s):
            return False
        if re.match(
            r"^(Total |Net |Sub|Income|Expense|Page |Month Ending|Operating |Year To|"
            r"VARIANCE|Location|As of|Actual |Rental |Other |Payroll |General |Repairs|"
            r"Make|Recreational|Contract |Advertising|Utilities|Management|Taxes|"
            r"Insurance|Partnership|Debt |Depreciation|Construction|Non-)", s
        ):
            return False
        if "Greater than" in s or "Budget %" in s or "% Var" in s:
            return False
        if re.match(r"^[\-\d,\.\s\(\)%\$]+$", s):
            return False
        if re.search(r"MTD|YTD|\$\d|vs\.|paint|written off|@|account|budget", s, re.I):
            return True
        if re.match(r"^[A-Z][a-z]", s) and " " in s:
            return True
        return False

    comentarios: List[Dict] = []
    seen = set()

    try:
        with pdfplumber.open(str(filepath)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                if "VARIANCE REPORT" not in text.upper() and "Variance Comments" not in text:
                    continue
                lines = text.split("\n")
                consumed = set()

                for i, ln in enumerate(lines):
                    m = ACCT_RX.match(ln)
                    if not m:
                        continue
                    code = m.group(1)
                    rest = m.group(2)
                    toks = rest.split()

                    # account name = tokens before first numeric
                    name_toks = []
                    idx = 0
                    while idx < len(toks):
                        t = toks[idx]
                        if NUM_TOK.match(t):
                            break
                        if re.match(r"^-?[\d,]+\.\d+\$", t):
                            break
                        name_toks.append(t)
                        idx += 1
                    if not name_toks:
                        continue
                    cuenta = f"{code} - " + " ".join(name_toks)

                    # numeric tokens
                    num_toks = []
                    while idx < len(toks) and NUM_TOK.match(toks[idx]):
                        num_toks.append(toks[idx])
                        idx += 1
                    if len(num_toks) < 4:
                        continue

                    # comment glued on same line
                    glued_remainder = " ".join(toks[idx:]).strip()
                    glued = []
                    for t in toks:
                        if re.search(r"\d[A-Za-z\$]|[A-Za-z]\$", t):
                            mg = re.search(r"([\$A-Za-z].+)$", t)
                            if mg:
                                glued.append(mg.group(1))
                    glued_text = " ".join(glued).strip()

                    parts = []
                    used = []
                    # line(s) ABOVE
                    if i - 1 >= 0 and i - 1 not in consumed and _is_comment_line(lines[i-1]):
                        if (i - 2 >= 0 and i - 2 not in consumed
                                and _is_comment_line(lines[i-2])
                                and not ACCT_RX.match(lines[i-2].strip())):
                            parts.append(lines[i-2].strip())
                            used.append(i-2)
                        parts.append(lines[i-1].strip())
                        used.append(i-1)
                    if glued_remainder:
                        parts.append(glued_remainder)
                    elif glued_text:
                        parts.append(glued_text)

                    # below continuation: only if mid-sentence
                    if parts:
                        current = " ".join(parts).strip()
                        if not current.rstrip().endswith((".", "!", "?")):
                            for j in range(i + 1, min(i + 3, len(lines))):
                                if j in consumed:
                                    break
                                nxt = lines[j].strip()
                                if not nxt:
                                    continue
                                if ACCT_RX.match(nxt):
                                    break
                                if re.match(r"^[\-\d,\.\s\(\)%\$]+$", nxt):
                                    break
                                if re.match(r"^(Total |Net |Page |Month Ending|VARIANCE|Location|As of|Operating |Year To)", nxt):
                                    break
                                parts.append(nxt)
                                used.append(j)
                                if nxt.rstrip().endswith((".", "!", "?")):
                                    break

                    if not parts:
                        continue
                    comentario = re.sub(r"\s+", " ", " ".join(parts)).strip()
                    # cleanup: separar numero pegado a $ o letras
                    comentario = re.sub(r"(-?\d+\.?\d*)\$", r"\1 $", comentario)
                    comentario = re.sub(r"(-?\d+\.\d+)([A-Za-z])", r"\1 \2", comentario)
                    comentario = re.sub(r"^-?\d+\.\d+\s+(?=[\$A-Z])", "", comentario)
                    comentario = re.sub(r"\s+", " ", comentario).strip()
                    if len(comentario) < 8:
                        continue
                    if re.match(r"^[\-\d,\.\s\(\)%\$]+$", comentario):
                        continue

                    key = (cuenta, comentario[:60])
                    if key in seen:
                        continue
                    seen.add(key)
                    for u in used:
                        consumed.add(u)

                    comentarios.append({
                        "account_code": code,
                        "description": cuenta,
                        "comment_text": comentario,
                        "source_file": filepath.name,
                    })
    except Exception:
        return comentarios

    return comentarios


def _is_j501_pdf_format(filepath: Path) -> bool:
    """
    Detecta PDF Close Package de J 501 Estates (JAG Management).
    Firma: page 1 contiene 'J 501 Estates' Y existe una pagina con
    'Budget Comparison Report'.
    """
    if filepath.suffix.lower() != '.pdf':
        return False
    try:
        import pdfplumber
        with pdfplumber.open(str(filepath)) as pdf:
            if not pdf.pages:
                return False
            p1 = pdf.pages[0].extract_text() or ''
            if 'J 501 Estates' not in p1 and '501 Estates Apartment' not in p1:
                return False
            for page in pdf.pages[:25]:
                t = page.extract_text() or ''
                if 'Budget Comparison Report' in t:
                    return True
        return False
    except Exception:
        return False


# Linea de datos del BCR en texto PDF:
#   "3110-1110 - Market Rent  471,922.00 481,766.00 (9,844.00) (2.04) 471,922.00 481,766.00 (9,844.00) (2.04)"
_J501_DATA_RE = re.compile(
    r'^\s*(\d{4}-\d{4})\s*-\s*(.+?)\s+'
    r'(\(?-?[\d,]+\.\d{2}\)?)\s+(\(?-?[\d,]+\.\d{2}\)?)\s+'
    r'(\(?-?[\d,]+\.\d{2}\)?)\s+(\(?-?[\d,]+\.\d{2}\)?)\s+'
    r'(\(?-?[\d,]+\.\d{2}\)?)\s+(\(?-?[\d,]+\.\d{2}\)?)\s+'
    r'(\(?-?[\d,]+\.\d{2}\)?)\s*(\(?-?[\d,]+\.\d{2}\)?)?\s*$'
)
# Linea de Total con descripcion + 8 numeros
_J501_TOTAL_RE = re.compile(
    r'^\s*(Total\s+[A-Za-z][\w &\-/]*?|Net Operating Income|Net Income|'
    r'Net Income After Capital Expenditures(?: and Non-Operating Expenses)?|'
    r'Total Operating Expenses|Total Non-Operating Expenses)\s+'
    r'(\(?-?[\d,]+\.\d{2}\)?)\s+(\(?-?[\d,]+\.\d{2}\)?)\s+'
    r'(\(?-?[\d,]+\.\d{2}\)?)\s+(\(?-?[\d,]+\.\d{2}\)?)\s+'
    r'(\(?-?[\d,]+\.\d{2}\)?)\s+(\(?-?[\d,]+\.\d{2}\)?)\s+'
    r'(\(?-?[\d,]+\.\d{2}\)?)\s*(\(?-?[\d,]+\.\d{2}\)?)?\s*$'
)


def _ingest_j501_pdf(filepath: Path) -> IngestedFile:
    """
    Parser para Close Package PDF de J 501 Estates.

    Extrae el Budget Comparison Report (paginas con codigos GL XXXX-XXXX) y
    los comentarios de socios del Variance Report W/Notes.
    """
    import pdfplumber

    rows = []
    section_totals = []
    partner_comments = []
    period_label = ""
    building = "J 501 Estates"
    in_bcr = False
    section_current = ""
    row_num = 0

    # Extractor inline de comentarios del Variance Report W/Notes.
    # Mismo algoritmo que scripts-re/extraer_j501.py — inlineado para que
    # funcione en Streamlit Cloud (donde no existe el path local del skill).
    partner_comments = _extract_j501_variance_comments(filepath)
    # Tomar la as-of date del primer page header
    try:
        import pdfplumber as _pp
        with _pp.open(str(filepath)) as _pdf:
            for _p in _pdf.pages[:5]:
                _t = _p.extract_text() or ''
                _m = re.search(r'As [oO]f Date[: ]+(\d{1,2})/(\d{1,2})/(\d{2,4})', _t)
                if _m:
                    _mo, _dd, _yr = _m.group(1), _m.group(2), _m.group(3)
                    if len(_yr) == 2:
                        _yr = '20' + _yr
                    period_label = f"{_yr}-{int(_mo):02d}-{int(_dd):02d}"
                    break
    except Exception:
        pass

    with pdfplumber.open(str(filepath)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            if 'Budget Comparison Report' in text:
                in_bcr = True
                # extraer fecha si esta visible (e.g., '01/31/26')
                if not period_label:
                    md = re.search(r'(\d{2})/(\d{2})/(\d{2,4})', text)
                    if md:
                        mo, dd, yr = md.group(1), md.group(2), md.group(3)
                        if len(yr) == 2:
                            yr = '20' + yr
                        period_label = f"{yr}-{int(mo):02d}-{int(dd):02d}"
            if not in_bcr:
                continue
            # Si entramos a otra seccion del Close Package, salir del BCR
            if any(marker in text for marker in (
                'Income Statement', 'Trial Balance Report', 'Reconciliation Report',
                'Revenue v Collections', 'Deposits v Collections', 'Loan Escrow',
                'DELINQUENT AND PREPAID', 'Vendor Aging', 'RESIDENT DEPOSIT AUDIT',
                'RECONCILIATION SUMMARY', 'RENT ROLL DETAIL',
            )):
                in_bcr = False
                continue

            for raw in text.split('\n'):
                line = raw.rstrip()
                if not line.strip():
                    continue
                m = _J501_DATA_RE.match(line)
                if m:
                    row_num += 1
                    acct = m.group(1)
                    desc = m.group(2).strip()
                    actual_c = _to_float(m.group(3))
                    budget_c = _to_float(m.group(4))
                    actual_y = _to_float(m.group(7))
                    budget_y = _to_float(m.group(8))
                    rows.append({
                        'row_num': row_num,
                        'description': desc,
                        'account': acct,
                        'budget_current': budget_c,
                        'actual_current': actual_c,
                        'budget_ytd': budget_y,
                        'actual_ytd': actual_y,
                        'row_type': 'data',
                        'source_file': filepath.name,
                    })
                    continue
                mt = _J501_TOTAL_RE.match(line)
                if mt:
                    row_num += 1
                    desc = mt.group(1).strip()
                    actual_c = _to_float(mt.group(2))
                    budget_c = _to_float(mt.group(3))
                    actual_y = _to_float(mt.group(6))
                    budget_y = _to_float(mt.group(7))
                    rows.append({
                        'row_num': row_num,
                        'description': desc,
                        'account': '',
                        'budget_current': budget_c,
                        'actual_current': actual_c,
                        'budget_ytd': budget_y,
                        'actual_ytd': actual_y,
                        'row_type': 'total',
                        'source_file': filepath.name,
                    })
                    section_totals.append({
                        'description': desc,
                        'budget_current': budget_c,
                        'actual_current': actual_c,
                        'budget_ytd': budget_y,
                        'actual_ytd': actual_y,
                        'row_num': row_num,
                    })

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': True,
            'sheet_used': None,
            'comments_sheet': None,
            'format': 'J501-PDF',
            'column_map': None,
        }
    )


def ingest_single_file(filepath: Path, sheet_name: str = None, target_period: Optional[str] = None) -> IngestedFile:
    """
    Ingiere un archivo .xlsm/.xlsx/.pdf completo.
    Detecta automáticamente la hoja del income statement (multi-sheet) o
    delega a parsers específicos por formato.

    Args:
        filepath: ruta al archivo Excel o PDF
        sheet_name: si se especifica, usa esta hoja directamente (omite auto-detección).
                    Solo aplica a archivos Excel.
        target_period: período seleccionado por el usuario (solo lo usan parsers
                       multi-período como CWS).
    """
    # --- DISPATCH PDF (antes de openpyxl, que no abre PDFs) ---
    if filepath.suffix.lower() == '.pdf':
        if _is_nwep_pdf_format(filepath):
            return _ingest_nwep_pdf(filepath)
        if _is_j501_pdf_format(filepath):
            return _ingest_j501_pdf(filepath)
        raise ValueError(
            f"El archivo PDF '{filepath.name}' no coincide con ningún formato "
            f"soportado. Hoy se soportan PDFs de NWEP/Walnut Street Wellesley "
            f"(Campus at Newton Wellesley monthly reports) y de J 501 Estates "
            f"(JAG Close Package). Para otros activos, subí el archivo en formato Excel."
        )

    wb = openpyxl.load_workbook(str(filepath), data_only=True)

    # --- DETECCIÓN DE HOJAS ---
    if sheet_name:
        sheet_info = {
            'income_statement': sheet_name,
            'comments_sheet': None,
            'all_scores': {},
        }
    else:
        sheet_info = detect_sheets(wb, filename=str(filepath.name))
    ws = wb[sheet_info['income_statement']]

    # --- DISPATCH por formato ---
    # NWEP-Summary debe chequearse PRIMERO porque tiene 'PTD Actual' en headers
    # (lo que activaría el detector CMC como falso positivo). El detector NWEP
    # es estricto (requiere 'SUMMARY ' en row 1 + 'Month' en row 5) → no hay
    # falsos positivos cruzados.
    if _is_nwep_summary_format(ws):
        result = _ingest_nwep_summary_sheet(ws, filepath, sheet_info)
        wb.close()
        return result

    # CWS / 12-month budget (Bridge at the Blockyard). Detección por marker
    # 'Tree = cws_*' en row 4 — muy específica, sin falsos positivos.
    if _is_cws_format(ws):
        result = _ingest_cws_sheet(ws, filepath, sheet_info, target_period=target_period)
        wb.close()
        return result

    # Budget Comparison Report (501 Estates): checkea primero porque aunque "BCR" puede no ser
    # auto-detectada por el sheet scorer, si alguna hoja del workbook es BCR debemos usarla.
    if _is_bcr_format(ws):
        result = _ingest_bcr_sheet(ws, filepath, sheet_info)
        wb.close()
        return result
    # Callan: detección via nombre de hoja del workbook ("income statement bud...act"),
    # muy específica — no hay falsos positivos con otros formatos.
    if _is_callan_format(ws):
        result = _ingest_callan_sheet(ws, filepath, sheet_info)
        wb.close()
        return result
    # YSI debe chequearse ANTES de CMC: Wilcox tiene "PTD" en header row (lo que
    # activa _is_cmc_format por substring), pero es YSI (tree=ysi_bf) y tiene
    # layout distinto (sections con código de cuenta). La detección YSI es estricta
    # (requiere "tree = ysi_bf|ysi_cf") así que no hay falsos positivos.
    if _is_ysi_format(ws):
        result = _ingest_ysi_sheet(ws, filepath, sheet_info)
        wb.close()
        return result
    if _is_cmc_format(ws):
        result = _ingest_cmc_sheet(ws, filepath, sheet_info)
        wb.close()
        return result
    if _is_jde_format(ws):
        result = _ingest_jde_sheet(ws, filepath, sheet_info)
        wb.close()
        return result
    if _is_yardi_variance_format(ws):
        result = _ingest_yardi_variance_sheet(ws, filepath, sheet_info)
        # Extraer comentarios de hoja separada si existe
        if sheet_info.get('comments_sheet'):
            comments_ws = wb[sheet_info['comments_sheet']]
            extra = _extract_comments_from_sheet(comments_ws, filepath.name)
            result.partner_comments.extend(extra)
        wb.close()
        return result

    # Metadata
    building, period_label = detect_building_and_period(ws, filepath.name)
    header_row = find_header_row(ws)

    # --- MAPEO FLEXIBLE DE COLUMNAS ---
    col_map = _map_columns_flexible(ws, header_row)
    desc_col = col_map['desc_col']
    acct_col = col_map['acct_col']
    budget_curr = col_map['budget_curr']
    actual_curr = col_map['actual_curr']
    budget_ytd = col_map['budget_ytd']
    actual_ytd = col_map['actual_ytd']
    notes_col = col_map['notes_col']

    # Leer datos fila por fila
    rows = []
    partner_comments = []
    section_totals = []

    max_row = ws.max_row or 300

    for i in range(header_row + 1, max_row + 1):
        desc_val = ws.cell(row=i, column=desc_col).value
        acct_val = ws.cell(row=i, column=acct_col).value
        budget_c = ws.cell(row=i, column=budget_curr).value
        actual_c = ws.cell(row=i, column=actual_curr).value
        budget_y = ws.cell(row=i, column=budget_ytd).value if budget_ytd else None
        actual_y = ws.cell(row=i, column=actual_ytd).value if actual_ytd else None
        notes_val = ws.cell(row=i, column=notes_col).value if notes_col else None

        # Saltar filas completamente vacías
        if desc_val is None and acct_val is None and budget_c is None and actual_c is None:
            continue

        desc_str = str(desc_val).strip() if desc_val else ""
        acct_str = str(acct_val).strip() if acct_val else ""

        # Detectar tipo de fila
        is_total = 'total' in desc_str.lower()
        is_section_header = (
            desc_str != "" and
            acct_str == "" and
            budget_c is None and
            actual_c is None
        )
        has_data = isinstance(budget_c, (int, float)) or isinstance(actual_c, (int, float))

        row_type = "total" if is_total else ("section" if is_section_header else ("data" if has_data else "label"))

        row = {
            'row_num': i,
            'description': desc_str,
            'account': acct_str,
            'budget_current': _to_float(budget_c),
            'actual_current': _to_float(actual_c),
            'budget_ytd': _to_float(budget_y),
            'actual_ytd': _to_float(actual_y),
            'row_type': row_type,
            'source_file': filepath.name,
        }
        rows.append(row)

        # Extraer comentarios del socio (inline en la hoja del IS)
        if notes_val and str(notes_val).strip():
            partner_comments.append({
                'account_code': acct_str,
                'description': desc_str,
                'comment_text': str(notes_val).strip(),
                'source_file': filepath.name,
                'row_num': i,
            })

        # Registrar totales de sección
        if is_total:
            section_totals.append({
                'description': desc_str,
                'budget_current': _to_float(budget_c),
                'actual_current': _to_float(actual_c),
                'budget_ytd': _to_float(budget_y),
                'actual_ytd': _to_float(actual_y),
                'row_num': i,
            })

    # --- EXTRAER COMENTARIOS DE HOJA SEPARADA (si existe) ---
    if sheet_info['comments_sheet']:
        comments_ws = wb[sheet_info['comments_sheet']]
        extra_comments = _extract_comments_from_sheet(comments_ws, filepath.name)
        partner_comments.extend(extra_comments)

    wb.close()

    df = pd.DataFrame(rows)
    return IngestedFile(
        filename=filepath.name,
        building=building,
        period_label=period_label,
        df=df,
        partner_comments=partner_comments,
        section_totals=section_totals,
        metadata={
            'header_row': header_row,
            'total_rows': len(rows),
            'data_rows': len([r for r in rows if r['row_type'] == 'data']),
            'has_ytd': budget_ytd is not None,
            'sheet_used': sheet_info['income_statement'],
            'comments_sheet': sheet_info['comments_sheet'],
            'sheets_analyzed': len(sheet_info['all_scores'].get('is', {})),
            'column_map': {k: v for k, v in col_map.items() if v is not None},
        }
    )


def ingest_files(filepaths: List[Path], sheet_selections: Dict[str, List[str]] = None, target_period: Optional[str] = None) -> List[IngestedFile]:
    """
    Ingiere múltiples archivos.

    Args:
        filepaths: lista de rutas a archivos Excel
        sheet_selections: dict {filename: [sheet_name, ...]} para workbooks multi-hoja.
                          Si un archivo no está en el dict, se usa auto-detección.
        target_period: período seleccionado por el usuario (ej. 'Q1 2026', 'Mar 2026').
                       Solo lo usan parsers que soportan archivos multi-período (ej. CWS
                       12-month budget); los demás lo ignoran.
    """
    results = []
    for fp in filepaths:
        sheets = (sheet_selections or {}).get(fp.name)
        if sheets:
            for sheet in sheets:
                try:
                    result = ingest_single_file(fp, sheet_name=sheet, target_period=target_period)
                    results.append(result)
                except Exception as e:
                    raise ValueError(f"Error procesando {fp.name} [{sheet}]: {str(e)}")
        else:
            try:
                result = ingest_single_file(fp, target_period=target_period)
                results.append(result)
            except Exception as e:
                raise ValueError(f"Error procesando {fp.name}: {str(e)}")

    # Normalizar building names: si un archivo tiene nombre corto (fallback de filename)
    # y otro archivo del batch tiene un nombre completo que contiene ese nombre corto,
    # usar el nombre completo (ej: "Edson" → "295 29th Street JV LLC" si otro archivo
    # del mismo grupo tiene "295 29th Street JV LLC" y "Edson" aparece en su filename).
    # Estrategia: agrupar archivos por raíz de filename (sin número de mes).
    filename_roots = {}
    for r in results:
        root = re.sub(r'\s*\d+\s*$', '', Path(r.filename).stem).strip().lower()
        if root not in filename_roots:
            filename_roots[root] = []
        filename_roots[root].append(r)

    for root, group in filename_roots.items():
        if len(group) < 2:
            continue
        # Buscar el nombre de edificio más largo/completo del grupo
        full_name = max((r.building for r in group if r.building), key=len, default="")
        if not full_name:
            continue
        # Asignar a todos los del grupo que tengan nombre más corto o igual al fallback
        for r in group:
            if r.building != full_name and len(r.building) < len(full_name):
                r.building = full_name

    return results


def _to_float(val) -> Optional[float]:
    """Convierte valor a float, manejando formatos contables."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(',', '').replace('$', '')
    # Formato paréntesis = negativo
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    try:
        return float(s)
    except (ValueError, TypeError):
        return None
