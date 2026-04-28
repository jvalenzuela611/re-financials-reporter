"""
Módulo Master Excel: crea y actualiza un archivo maestro de Excel por edificio
que conserva la información financiera histórica mensual (budget tracking).

Estructura del master (por año):
- Hoja "Master": una fila por account_code con el siguiente layout de columnas:

  MONTHLY ACTUALS (T-12): M1 | M2 | M3 | M4 | M5 | M6 | M7 | M8 | M9 | M10 | M11 | M12
  Q1:  actual | budget | variance
  Q2:  actual | budget | variance
  Q3:  actual | budget | variance
  Q4:  actual | budget | variance
  YTD: actual | budget | variance

- Hoja "Partner Comments": comentarios del socio por cuenta y quarter
- Hoja "Management Commentary": input del analista (recurrente / puntual)
- Hoja "Log": registro de cada actualización

Filename: master_{building}_{year}.xlsx (un archivo por año)

Uso:
    from master_excel import update_master_excel
    master_path = update_master_excel(
        building="Alice House",
        period="Q1 2025",
        structure=structure,
        monthly_structures=[struct_m1, struct_m2, struct_m3],
        master_dir="./masters",
        user="analyst@stars.com",
    )
"""

import pandas as pd
import os
import re
from pathlib import Path
from datetime import datetime
from typing import Dict, Optional, List

from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from src.config import l1_pl_sort_key


def _sort_df_by_pl_order(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ordena un DataFrame del master siguiendo el orden NATURAL del P&L del partner:
      1) parent_line via L1_PL_ORDER (Income → OpEx → RET → NOI → Interest →
         Non-Op → Net Income → Total Capex → NA)
      2) Dentro de cada L1: section_name en orden de aparición del Excel
         (usando _partner_row_order mínimo por (L1, section_name))
      3) Dentro de cada section_name: _partner_row_order ASC (row original del Excel)
      4) Tie-break: account_code ASC

    Si _partner_row_order no está presente (masters históricos anteriores a esta
    versión), cae a (L1, account_code ASC) — compatibilidad hacia atrás.
    """
    if df is None or df.empty:
        return df
    if 'parent_line' not in df.columns:
        return df

    index_name = df.index.name
    is_indexed = index_name == 'account_code'
    tmp = df.reset_index() if is_indexed else df.copy()

    tmp['_l1_ord'] = tmp['parent_line'].fillna('').apply(l1_pl_sort_key)

    has_partner_order = '_partner_row_order' in tmp.columns

    if has_partner_order and 'section_name' in tmp.columns:
        # Orden de aparición de cada (L1, section) = min(_partner_row_order)
        tmp['_section_fillna'] = tmp['section_name'].fillna('')
        first_order = (
            tmp.groupby(['parent_line', '_section_fillna'])['_partner_row_order']
            .transform('min')
        )
        tmp['_l2_ord'] = first_order
        sort_keys = ['_l1_ord', '_l2_ord', '_partner_row_order']
        if 'account_code' in tmp.columns:
            sort_keys.append('account_code')
        tmp = tmp.sort_values(sort_keys, kind='stable')
        tmp = tmp.drop(columns=['_l1_ord', '_l2_ord', '_section_fillna',
                                '_partner_row_order'])
    else:
        # Fallback: sin partner order → L1 canónico + account_code
        sort_keys = ['_l1_ord']
        if 'account_code' in tmp.columns:
            sort_keys.append('account_code')
        tmp = tmp.sort_values(sort_keys, kind='stable')
        tmp = tmp.drop(columns=['_l1_ord'])

    if is_indexed and 'account_code' in tmp.columns:
        tmp = tmp.set_index('account_code')
    return tmp


# ─── Helpers ────────────────────────────────────────────────────

def _sanitize_building_name(building: str) -> str:
    return building.strip().replace(" ", "_").replace("(", "").replace(")", "").lower()


def _extract_year_from_period(period: str) -> int:
    match = re.search(r'(\d{4})', period)
    if not match:
        raise ValueError(f"Cannot extract year from period: {period}")
    return int(match.group(1))


def _extract_quarter_from_period(period: str) -> int:
    match = re.search(r'Q(\d)', period, re.IGNORECASE)
    if not match:
        raise ValueError(f"Cannot extract quarter from period: {period}")
    return int(match.group(1))


def _months_for_quarter(q_num: int) -> List[int]:
    start = (q_num - 1) * 3 + 1
    return [start, start + 1, start + 2]


# ─── Column definitions ─────────────────────────────────────────

_FIXED_COLS = ['description', 'parent_line', 'section_name']
_Q_LABELS = ['Q1', 'Q2', 'Q3', 'Q4']

# Colors (RRGGBB)
_COLOR_MONTHLY_HEADER = '2E86C1'   # blue header for monthly section
_COLOR_MONTHLY_CELL   = 'D6EAF8'   # light blue cell

_QUARTER_HEADER_COLORS = {         # merged Q header (one per quarter)
    'Q1': '1A5276', 'Q2': '1F618D', 'Q3': '2874A6', 'Q4': '2E86C1',
}
_QUARTER_COL_COLORS = {            # sub-column within each quarter block
    'actual':   'D6EAF8',          # light blue
    'budget':   'D5F5E3',          # light green
    'variance': 'FDEBD0',          # light orange
}
_QUARTER_COL_HEADER_COLORS = {
    'actual':   '2E86C1',
    'budget':   '27AE60',
    'variance': 'E67E22',
}

_COLOR_YTD_HEADER = '1B2631'       # dark for YTD merged header
_YTD_COL_COLORS = {
    'actual':   '85C1E9',
    'budget':   '82E0AA',
    'variance': 'F0B27A',
}
_YTD_COL_HEADER_COLORS = {
    'actual':   '1A5276',
    'budget':   '1D8348',
    'variance': 'A04000',
}

_HEADER_FONT       = Font(bold=True, size=10)
_SECTION_FONT      = Font(bold=True, size=10, color='FFFFFF')


def _canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cuando se lee el master con header=1, _apply_formatting ya habrá escrito
    etiquetas visuales cortas en fila 2 (M1, M2, actual, budget, variance…).
    Esta función las renombra de vuelta a los nombres canónicos del DataFrame
    (M1_actual, Q1_actual, Q1_budget…) para que el merge funcione correctamente.
    """
    quarters = ['Q1', 'Q2', 'Q3', 'Q4', 'YTD']
    suffixes = ['actual', 'budget', 'variance']

    rename: dict = {}
    for col in df.columns:
        col_s = str(col)
        if col_s in _FIXED_COLS:
            continue
        # Monthly actuals: M1 … M12 (integers 1-12)
        m = re.match(r'^M(\d{1,2})$', col_s)
        if m:
            rename[col_s] = f'M{m.group(1)}_actual'
            continue
        # Quarterly / YTD sub-columns: actual(.N), budget(.N), variance(.N)
        for suf_idx, suf in enumerate(suffixes):
            # pandas duplicate suffix: first = 'actual', rest = 'actual.1', 'actual.2' …
            for q_idx, q in enumerate(quarters):
                expected = suf if q_idx == 0 else f'{suf}.{q_idx}'
                if col_s == expected:
                    rename[col_s] = f'{q}_{suf}'
                    break

    return df.rename(columns=rename) if rename else df


def _build_display_columns() -> list:
    """
    Ordered list of data columns in the new layout:
    M1_actual … M12_actual | Q1_actual Q1_budget Q1_variance | … | YTD_actual YTD_budget YTD_variance
    """
    cols = [f'M{m}_actual' for m in range(1, 13)]
    for q in range(1, 5):
        cols += [f'Q{q}_actual', f'Q{q}_budget', f'Q{q}_variance']
    cols += ['YTD_actual', 'YTD_budget', 'YTD_variance']
    return cols


_DISPLAY_COLUMNS = _build_display_columns()


# ─── Computation ────────────────────────────────────────────────

def _compute_quarterly_and_ytd(df: pd.DataFrame) -> pd.DataFrame:
    """
    Computes from stored monthly actuals and quarterly budgets:
      Q_actual  = sum of monthly actuals for that quarter
      Q_variance = Q_actual - Q_budget
      YTD_actual  = sum of all monthly actuals
      YTD_budget  = sum of quarterly budgets
      YTD_variance = YTD_actual - YTD_budget
    """
    for q_num in range(1, 5):
        q_act = f'Q{q_num}_actual'
        q_bud = f'Q{q_num}_budget'
        q_var = f'Q{q_num}_variance'

        # Quarterly actual: sum monthly actuals
        m_cols = [f'M{m}_actual' for m in _months_for_quarter(q_num)]
        existing = [c for c in m_cols if c in df.columns]
        if existing:
            df[q_act] = df[existing].sum(axis=1)
            df.loc[df[existing].isna().all(axis=1), q_act] = pd.NA

        # Quarterly variance: actual - budget
        if q_act in df.columns and q_bud in df.columns:
            df[q_var] = df[q_act] - df[q_bud]
            df.loc[df[q_act].isna() & df[q_bud].isna(), q_var] = pd.NA

    # YTD actual: sum all monthly actuals
    all_m = [f'M{m}_actual' for m in range(1, 13)]
    existing_m = [c for c in all_m if c in df.columns]
    if existing_m:
        df['YTD_actual'] = df[existing_m].sum(axis=1)
        df.loc[df[existing_m].isna().all(axis=1), 'YTD_actual'] = pd.NA

    # YTD budget: sum quarterly budgets
    all_q_bud = [f'Q{q}_budget' for q in range(1, 5)]
    existing_qb = [c for c in all_q_bud if c in df.columns]
    if existing_qb:
        df['YTD_budget'] = df[existing_qb].sum(axis=1)
        df.loc[df[existing_qb].isna().all(axis=1), 'YTD_budget'] = pd.NA

    # YTD variance
    if 'YTD_actual' in df.columns and 'YTD_budget' in df.columns:
        df['YTD_variance'] = df['YTD_actual'] - df['YTD_budget']
        df.loc[df['YTD_actual'].isna() & df['YTD_budget'].isna(), 'YTD_variance'] = pd.NA

    return df


# ─── Formatting ─────────────────────────────────────────────────

def _apply_formatting(ws, n_fixed: int, n_rows: int):
    """
    Inserts two header rows above the data:
      Row 1 (section headers, merged): MONTHLY ACTUALS | Q1 | Q2 | Q3 | Q4 | YTD
      Row 2 (sub-headers):             M1…M12 | actual budget var (x4) | actual budget var
      Row 3+: data
    """
    ws.insert_rows(1)  # make room for section header row

    # col index of first data column (1-based: account_code col + fixed cols + 1)
    data_start = n_fixed + 2  # account_code counts as col 1

    # ── Section 1: Monthly Actuals ───────────────────────────────
    monthly_start = data_start
    monthly_end   = data_start + 11  # 12 months

    ws.merge_cells(start_row=1, start_column=monthly_start,
                   end_row=1, end_column=monthly_end)
    sec_cell = ws.cell(row=1, column=monthly_start)
    sec_cell.value = 'MONTHLY ACTUALS'
    sec_cell.font  = _SECTION_FONT
    sec_cell.fill  = PatternFill('solid', fgColor=_COLOR_MONTHLY_HEADER)
    sec_cell.alignment = Alignment(horizontal='center')

    for i, m in enumerate(range(1, 13)):
        col = monthly_start + i
        hdr = ws.cell(row=2, column=col)
        hdr.value = f'M{m}'
        hdr.font  = _HEADER_FONT
        hdr.fill  = PatternFill('solid', fgColor=_COLOR_MONTHLY_CELL)
        hdr.alignment = Alignment(horizontal='center')

    # ── Sections 2-5: Quarterly blocks ──────────────────────────
    q_start = monthly_end + 1
    sub_labels = ['actual', 'budget', 'variance']

    for q_idx, q_label in enumerate(_Q_LABELS):
        block_start = q_start + q_idx * 3
        block_end   = block_start + 2

        ws.merge_cells(start_row=1, start_column=block_start,
                       end_row=1, end_column=block_end)
        sec = ws.cell(row=1, column=block_start)
        sec.value = q_label
        sec.font  = _SECTION_FONT
        sec.fill  = PatternFill('solid', fgColor=_QUARTER_HEADER_COLORS[q_label])
        sec.alignment = Alignment(horizontal='center')

        for j, sub in enumerate(sub_labels):
            col = block_start + j
            hdr = ws.cell(row=2, column=col)
            hdr.value = sub
            hdr.font  = _HEADER_FONT
            hdr.fill  = PatternFill('solid', fgColor=_QUARTER_COL_HEADER_COLORS[sub])
            hdr.alignment = Alignment(horizontal='center')

    # ── Section 6: YTD ──────────────────────────────────────────
    ytd_start = q_start + 4 * 3
    ytd_end   = ytd_start + 2

    ws.merge_cells(start_row=1, start_column=ytd_start,
                   end_row=1, end_column=ytd_end)
    sec = ws.cell(row=1, column=ytd_start)
    sec.value = 'YTD'
    sec.font  = _SECTION_FONT
    sec.fill  = PatternFill('solid', fgColor=_COLOR_YTD_HEADER)
    sec.alignment = Alignment(horizontal='center')

    for j, sub in enumerate(sub_labels):
        col = ytd_start + j
        hdr = ws.cell(row=2, column=col)
        hdr.value = sub
        hdr.font  = _HEADER_FONT
        hdr.fill  = PatternFill('solid', fgColor=_YTD_COL_HEADER_COLORS[sub])
        hdr.alignment = Alignment(horizontal='center')

    # ── Fixed column headers ─────────────────────────────────────
    for col_idx in range(1, n_fixed + 2):
        ws.cell(row=2, column=col_idx).font = _HEADER_FONT

    # ── Column widths ────────────────────────────────────────────
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, min(n_rows + 3, 50)):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 8), 20)

    # ── Freeze: fixed cols + 2 header rows ──────────────────────
    ws.freeze_panes = ws.cell(row=3, column=n_fixed + 2)


# ─── Migration from old band-format masters ──────────────────────

def _try_migrate_old_master(building: str, master_dir: str, year: int) -> Optional[pd.DataFrame]:
    """Load an existing master regardless of format (old 3-band or new period-first)."""
    safe_name = _sanitize_building_name(building)

    for pattern in [f"master_{safe_name}_{year}.xlsx", f"master_{safe_name}.xlsx"]:
        old_path = Path(master_dir) / pattern
        if not old_path.exists():
            continue

        for header_row in (1, 0):
            try:
                df = pd.read_excel(
                    str(old_path), sheet_name='Master',
                    index_col='account_code',
                    dtype={'account_code': str},
                    header=header_row,
                )
                df.index = df.index.astype(str)
                df = df[[c for c in df.columns if not str(c).startswith('Unnamed')]]
                df = df[df.index.notna() & (df.index != '')]
                if any(c not in _FIXED_COLS for c in df.columns):
                    return _canonicalize_columns(df)
            except Exception:
                continue

    return None


# ─── Main function ──────────────────────────────────────────────

def update_master_excel(
    building: str,
    period: str,
    structure: Dict,
    master_dir: str = "./masters",
    user: str = "system",
    monthly_structures: Optional[List[Dict]] = None,
) -> str:
    """
    Crea o actualiza el archivo maestro de Excel para un edificio.

    Nuevo layout:
      MONTHLY ACTUALS: M1…M12 (actual only — historia anual)
      Q1 actual | Q1 budget | Q1 variance
      Q2 actual | Q2 budget | Q2 variance
      Q3 actual | Q3 budget | Q3 variance
      Q4 actual | Q4 budget | Q4 variance
      YTD actual | YTD budget | YTD variance

    Args:
        building: Nombre del edificio
        period: Período trimestral (e.g., "Q3 2025")
        structure: Output de build_financial_structure() (quarterly consolidated)
        master_dir: Directorio de masters
        user: Identificador del usuario
        monthly_structures: Lista de estructuras mensuales en orden (M1, M2, M3 del quarter).
            Si no se proveen, se usan los totales del quarterly structure como fallback.

    Returns:
        Ruta absoluta del archivo maestro.
    """
    os.makedirs(master_dir, exist_ok=True)
    safe_name = _sanitize_building_name(building)
    year      = _extract_year_from_period(period)
    q_num     = _extract_quarter_from_period(period)
    q_months  = _months_for_quarter(q_num)
    q_label   = f'Q{q_num}'

    master_path = Path(master_dir) / f"master_{safe_name}_{year}.xlsx"

    # ── Build current data ───────────────────────────────────────
    current_data: Dict[str, dict] = {}

    def _init_account(acct):
        code = acct.account_code
        if not code:
            return None
        if code not in current_data:
            current_data[code] = {
                'account_code': code,
                'description': acct.name,
                'parent_line': acct.parent_line,
                'section_name': getattr(acct, 'section_name', '') or '',
                # Guardamos partner_row_order para sortear al final con el
                # orden natural del P&L del partner
                '_partner_row_order': getattr(acct, 'partner_row_order', 0) or 0,
            }
        return code

    # Monthly actuals from monthly structures.
    # IMPORTANTE: el M column de cada estructura se determina por el MES real
    # del archivo (via period_label o filename), NO por la posición en la lista.
    # Antes se usaba `q_months[i]` donde i era el índice, lo que escribía datos
    # en el mes equivocado si la lista no venía cronológica o si el quarter se
    # detectaba mal.
    if monthly_structures:
        from src.pipeline import _parse_month_year_from_text
        for i, month_struct in enumerate(monthly_structures):
            if month_struct is None:
                continue
            # Resolver el mes real del archivo
            ms_month = None
            ms_year = None
            plabel = month_struct.get('_period_label', '') or ''
            fn = month_struct.get('_filename', '') or ''
            mo, yy = _parse_month_year_from_text(plabel)
            if not mo:
                mo, yy = _parse_month_year_from_text(fn)
            if mo:
                ms_month, ms_year = mo, yy

            # Si no pudimos detectar el mes, caer al mapeo posicional como antes
            if ms_month is None:
                ms_month = q_months[i] if i < len(q_months) else q_months[-1]

            # Validación: el mes detectado debe estar en el quarter declarado.
            # Si no, warning silencioso (la validación en UI debería haberlo
            # detectado antes, pero por si acaso preservamos los datos).
            if ms_month not in q_months:
                # El archivo no pertenece al quarter declarado → saltear
                continue

            m_col = f'M{ms_month}_actual'
            for acct in month_struct.get('l3_accounts', []):
                code = _init_account(acct)
                if code:
                    current_data[code][m_col] = acct.actual_current

    # Quarterly budget (and fallback actual) from quarterly structure
    for acct in structure.get('l3_accounts', []):
        code = _init_account(acct)
        if not code:
            continue
        current_data[code][f'{q_label}_budget'] = acct.budget_current
        if not monthly_structures:
            # Fallback: no monthly files — write quarterly actual directly
            current_data[code][f'{q_label}_actual'] = acct.actual_current

    current_df = pd.DataFrame(current_data.values())
    if not current_df.empty:
        current_df = current_df.set_index('account_code')

    # ── Load or migrate existing master ─────────────────────────
    existing_df   = pd.DataFrame()
    existing_df.index.name = 'account_code'
    preserved_sheets: Dict[str, pd.DataFrame] = {}

    if master_path.exists():
        try:
            xls = pd.ExcelFile(str(master_path))
            for sheet in xls.sheet_names:
                if sheet == 'Master':
                    continue
                if sheet == 'Log':
                    continue
                try:
                    preserved_sheets[sheet] = pd.read_excel(
                        str(master_path), sheet_name=sheet, dtype=str,
                    )
                except Exception:
                    pass

            # Read Master sheet — try header=1 first (section header in row 1)
            for hdr in (1, 0):
                try:
                    df = pd.read_excel(
                        str(master_path), sheet_name='Master',
                        index_col='account_code',
                        dtype={'account_code': str},
                        header=hdr,
                    )
                    df.index = df.index.astype(str)
                    df = df[[c for c in df.columns if not str(c).startswith('Unnamed')]]
                    df = df[df.index.notna() & (df.index != '')]
                    if any(c not in _FIXED_COLS for c in df.columns):
                        # Restore canonical column names if master was saved with visual labels
                        df = _canonicalize_columns(df)
                        existing_df = df
                        break
                except Exception:
                    continue
        except Exception:
            pass
    else:
        migrated = _try_migrate_old_master(building, master_dir, year)
        if migrated is not None:
            existing_df = migrated

    # ── Merge existing + current ─────────────────────────────────
    if existing_df.empty:
        merged = current_df.copy() if not current_df.empty else pd.DataFrame()
    else:
        # Drop columns that will be overwritten or recomputed
        cols_to_drop = []
        for m in q_months:
            cols_to_drop.append(f'M{m}_actual')
        cols_to_drop += [
            f'{q_label}_actual', f'{q_label}_budget', f'{q_label}_variance',
            'YTD_actual', 'YTD_budget', 'YTD_variance',
        ]
        existing_df = existing_df.drop(
            columns=[c for c in cols_to_drop if c in existing_df.columns],
            errors='ignore',
        )

        # _partner_row_order es una columna INTERNA de ordenamiento; va con
        # current_df (no con _FIXED_COLS que son las visibles). Lo tratamos
        # como si fuera una columna de data en el merge, y lo dropeamos antes
        # del to_excel.
        _INTERNAL_SORT_COL = '_partner_row_order'
        protected_cols = set(_FIXED_COLS) | {_INTERNAL_SORT_COL}

        hist_cols = [c for c in existing_df.columns if c not in protected_cols]
        new_cols  = [c for c in current_df.columns  if c not in protected_cols]

        merged = existing_df[hist_cols].join(
            current_df[new_cols] if not current_df.empty else pd.DataFrame(),
            how='outer',
        )

        if not current_df.empty:
            for col in _FIXED_COLS:
                if col in current_df.columns:
                    merged[col] = current_df[col].reindex(merged.index)
                    if col in existing_df.columns:
                        merged[col] = merged[col].fillna(existing_df[col])

            # Propagar la columna interna de sort si está en current_df.
            # Para rows históricas que no están en current → NaN → el sort
            # las manda al final de cada L1 (backward compat aceptable).
            if _INTERNAL_SORT_COL in current_df.columns:
                merged[_INTERNAL_SORT_COL] = current_df[_INTERNAL_SORT_COL].reindex(merged.index)

    if merged is None or (hasattr(merged, 'empty') and merged.empty):
        merged = current_df.copy()

    # ── Recompute quarterly actuals, YTD, and variances ─────────
    merged = _compute_quarterly_and_ytd(merged)

    # Sort PRIMERO (mientras _partner_row_order todavía está en el DF):
    #   1) L1 canónico P&L (Income → OpEx → RET → NOI → Interest → Non-Op →
    #      Net Income → Total Capex → NA)
    #   2) L2 (section_name) en orden de aparición del Excel del partner
    #   3) Dentro de cada L2, por partner_row_order ASC (row del Excel)
    # Así el master se lee como estado de resultados de arriba hacia abajo.
    merged = _sort_df_by_pl_order(merged)

    # ── Reorder to display layout (drop columna interna de sort) ─
    # CRÍTICO: incluir SIEMPRE las 27 columnas de display (_DISPLAY_COLUMNS),
    # aunque algunas no tengan datos (se rellenan con NaN). Si no se hace, las
    # celdas se desplazan a la izquierda y quedan bajo headers equivocados:
    # ej. correr solo Q4 → M10 actual aparece visible bajo "M1", Q4 actual bajo
    # "M4", etc. Bug detectado en backtest 501 Estates Q4 2025.
    final_fixed = [c for c in _FIXED_COLS if c in merged.columns]
    for c in _DISPLAY_COLUMNS:
        if c not in merged.columns:
            merged[c] = pd.NA
    merged = merged[final_fixed + list(_DISPLAY_COLUMNS)]

    # ── Partner Comments sheet ───────────────────────────────────
    partner_comments_df = _build_partner_comments_df(structure, q_num)
    if 'Partner Comments' in preserved_sheets and partner_comments_df is not None:
        partner_comments_df = _merge_partner_comments(
            preserved_sheets.pop('Partner Comments'), partner_comments_df, q_num,
        )
    elif 'Partner Comments' in preserved_sheets and partner_comments_df is None:
        partner_comments_df = preserved_sheets.pop('Partner Comments')

    # ── Log ──────────────────────────────────────────────────────
    log_df = pd.DataFrame(columns=['timestamp', 'user', 'period', 'accounts_updated', 'action'])
    if master_path.exists():
        try:
            log_df = pd.read_excel(str(master_path), sheet_name='Log')
        except Exception:
            pass

    new_log = pd.DataFrame([{
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'user': user,
        'period': period,
        'accounts_updated': len(current_data),
        'action': 'update' if master_path.exists() else 'create',
    }])
    log_df = pd.concat([log_df, new_log], ignore_index=True)

    # ── Write Excel ──────────────────────────────────────────────
    with pd.ExcelWriter(str(master_path), engine='openpyxl') as writer:
        merged.to_excel(writer, sheet_name='Master', index=True)

        if partner_comments_df is not None and not partner_comments_df.empty:
            partner_comments_df.to_excel(writer, sheet_name='Partner Comments', index=True)

        for sheet_name, sheet_df in preserved_sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

        log_df.to_excel(writer, sheet_name='Log', index=False)

    # ── Apply visual formatting ──────────────────────────────────
    try:
        from openpyxl import load_workbook
        wb = load_workbook(str(master_path))
        ws = wb['Master']
        _apply_formatting(ws, len(final_fixed), len(merged))
        wb.save(str(master_path))
    except Exception:
        pass

    return str(master_path.resolve())


# ─── Partner Comments ───────────────────────────────────────────

def _build_partner_comments_df(structure: Dict, q_num: int) -> Optional[pd.DataFrame]:
    rows = {}
    q_col = f'Q{q_num}_comments'
    for acct in structure.get('l3_accounts', []):
        code = acct.account_code
        if not code:
            continue
        comment = getattr(acct, 'partner_comment', '') or ''
        if not comment.strip():
            continue
        rows[code] = {
            'account_code': code,
            'description': acct.name,
            'parent_line': acct.parent_line,
            q_col: comment.strip(),
        }

    if not rows:
        return None

    df = pd.DataFrame(rows.values()).set_index('account_code')
    # Mismo orden P&L que el master principal para lectura consistente
    return _sort_df_by_pl_order(df)


def _merge_partner_comments(
    existing_df: pd.DataFrame,
    new_df: pd.DataFrame,
    q_num: int,
) -> pd.DataFrame:
    if 'account_code' in existing_df.columns:
        existing_df = existing_df.set_index('account_code')
    existing_df.index = existing_df.index.astype(str)

    q_col = f'Q{q_num}_comments'
    fixed = ['description', 'parent_line']

    existing_df = existing_df.drop(columns=[q_col], errors='ignore')

    merged = existing_df[[c for c in existing_df.columns if c not in fixed]].join(
        new_df[[c for c in new_df.columns if c not in fixed]], how='outer',
    )

    for col in fixed:
        if col in new_df.columns:
            merged[col] = new_df[col].reindex(merged.index)
            if col in existing_df.columns:
                merged[col] = merged[col].fillna(existing_df[col])

    final_fixed = [c for c in fixed if c in merged.columns]
    comment_cols = sorted(
        [c for c in merged.columns if c.endswith('_comments')],
        key=lambda x: int(re.search(r'Q(\d)', x).group(1)),
    )
    # Mismo orden P&L que el master principal
    return _sort_df_by_pl_order(merged[final_fixed + comment_cols])
