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


def ingest_single_file(filepath: Path, sheet_name: str = None) -> IngestedFile:
    """
    Ingiere un archivo .xlsm/.xlsx completo.
    Detecta automáticamente la hoja del income statement (multi-sheet).
    Extrae: datos, comentarios del socio, totales de sección, metadata.

    Args:
        filepath: ruta al archivo Excel
        sheet_name: si se especifica, usa esta hoja directamente (omite auto-detección)
    """
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


def ingest_files(filepaths: List[Path], sheet_selections: Dict[str, List[str]] = None) -> List[IngestedFile]:
    """
    Ingiere múltiples archivos.

    Args:
        filepaths: lista de rutas a archivos Excel
        sheet_selections: dict {filename: [sheet_name, ...]} para workbooks multi-hoja.
                          Si un archivo no está en el dict, se usa auto-detección.
    """
    results = []
    for fp in filepaths:
        sheets = (sheet_selections or {}).get(fp.name)
        if sheets:
            for sheet in sheets:
                try:
                    result = ingest_single_file(fp, sheet_name=sheet)
                    results.append(result)
                except Exception as e:
                    raise ValueError(f"Error procesando {fp.name} [{sheet}]: {str(e)}")
        else:
            try:
                result = ingest_single_file(fp)
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
