# src/pipeline.py

"""
Orquestador del pipeline de procesamiento financiero.
Recibe inputs del usuario y retorna un dict con todos los resultados.
"""

import os
import tempfile
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

from src.ingestion import ingest_files
from src.native_structure import build_financial_structure, consolidate_structures
from src.validation import validate_all
from src.output_formatter import (
    build_summary_table,
    build_bullets_by_line,
    build_alerts_section,
    build_comments_table,
    generate_full_markdown,
    generate_support_excel,
    generate_copilot_payload,
    generate_copilot_prompt,
    generate_account_tracking_excel,
)
from src.historical_analysis import build_historical_analysis
from src.master_excel import update_master_excel
from src.market_analysis import analyze_market


@dataclass
class PipelineInput:
    files_current: list          # Lista de UploadedFile de Streamlit
    files_prior: list            # Lista de UploadedFile (puede ser vacía)
    building: str
    period: str
    market_file: Optional[object] = None
    market_name: str = ""
    master_upload_path: Optional[Path] = None
    consolidate: bool = False
    sheet_selections: Optional[dict] = None


# ── Funciones privadas del módulo ──────────────────────────────


def _save_uploaded_files(uploaded_files, prefix: str) -> list:
    """Guarda archivos subidos en directorio temporal y retorna paths."""
    paths = []
    if not uploaded_files:
        return paths

    tmp_dir = tempfile.mkdtemp(prefix=f"re_financials_{prefix}_")

    for uf in uploaded_files:
        filepath = Path(tmp_dir) / uf.name
        filepath.write_bytes(uf.getbuffer())
        paths.append(filepath)

    return paths


def _parse_month_year_from_text(text: str):
    """
    Intenta extraer (month, year) de un string como:
      - "2025-10-31"       → (10, 2025)
      - "2025-10"          → (10, 2025)
      - "10/31/25"         → (10, 2025)
      - "Dec 2025"         → (12, 2025)
      - "December 31, 2025"→ (12, 2025)
      - "Q4 2025"          → (12, 2025)  (último mes del quarter)
    Retorna (month, year) o (None, None).
    """
    if not text:
        return (None, None)
    s = str(text).strip()
    lower = s.lower()

    # 1) YYYY-MM o YYYY-MM-DD (más específico — es el que BCR usa)
    m = re.search(r'\b(20\d{2})[-_./](\d{1,2})(?:[-_./]\d{1,2})?\b', s)
    if m:
        yr, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return (mo, yr)

    # 2) MM/DD/YY o MM-DD-YY (BCR row 8 col B = "10/31/25")
    m = re.search(r'\b(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})\b', s)
    if m:
        mo, _, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= mo <= 12:
            if yr < 100:
                yr += 2000
            return (mo, yr)

    # 3) "QX YYYY" directo
    m = re.search(r'\bQ([1-4])\b.*?\b(20\d{2})\b', s, re.IGNORECASE)
    if m:
        q, yr = int(m.group(1)), int(m.group(2))
        return (q * 3, yr)  # último mes del trimestre

    # 4) Month name + year
    month_map = {
        'january': 1, 'february': 2, 'march': 3, 'april': 4,
        'may': 5, 'june': 6, 'july': 7, 'august': 8,
        'september': 9, 'october': 10, 'november': 11, 'december': 12,
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4,
        'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
        'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4,
        'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8,
        'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12,
    }
    year_match = re.search(r'20\d{2}', s)
    yr = int(year_match.group()) if year_match else None
    mo = None
    # Priorizar nombres largos antes que cortos (para evitar "mar" matcheando "March")
    for name in sorted(month_map.keys(), key=len, reverse=True):
        if re.search(r'\b' + re.escape(name) + r'\b', lower):
            mo = month_map[name]
            break
    if mo and yr:
        return (mo, yr)

    return (None, None)


def _detect_period_from_files(ingested_files) -> str:
    """
    Detecta el período (ej: 'Q1 2026') a partir de los archivos ingeridos.

    Estrategia:
    1. Intenta extraer (month, year) del period_label de cada archivo. Si todos
       los archivos dan un mes válido, computa el quarter a partir del rango
       (mes_min, mes_max). Si todos los meses pertenecen al mismo quarter, usa
       ese; si abarcan múltiples, usa el del mes más reciente.
    2. Si los period_labels no alcanzan, intenta desde los nombres de archivo.
    3. Si nada funciona, retorna "" (cadena vacía) — el caller debe manejar.
       Antes había un silent fallback a datetime.now() que producía bugs.
    """
    # 1) Extraer (mes, año) de cada period_label
    months_years = []
    for f in ingested_files:
        label = getattr(f, 'period_label', '') or ''
        mo, yr = _parse_month_year_from_text(label)
        if mo and yr:
            months_years.append((mo, yr))

    # 2) Fallback: intentar desde filenames
    if not months_years:
        for f in ingested_files:
            fname = getattr(f, 'filename', '') or ''
            mo, yr = _parse_month_year_from_text(fname)
            if mo and yr:
                months_years.append((mo, yr))

    if not months_years:
        return ""  # caller debe pedir al analista que confirme manualmente

    # Usar el mes más reciente (mayor year*100+mo) para determinar el quarter
    months_years.sort(key=lambda t: t[1] * 100 + t[0])
    last_mo, last_yr = months_years[-1]
    quarter = (last_mo - 1) // 3 + 1
    return f"Q{quarter} {last_yr}"


def _parse_period_label(period: str):
    """Parse 'Q4 2025' -> (4, 2025). Retorna (None, None) si no matchea."""
    m = re.match(r'^\s*Q([1-4])\s+(20\d{2})\s*$', period or '', re.IGNORECASE)
    if not m:
        return (None, None)
    return (int(m.group(1)), int(m.group(2)))


def validate_period_coverage(period: str, ingested_files) -> dict:
    """
    Valida que el período declarado sea coherente con los archivos cargados.

    Returns:
        {
            'ok': bool,
            'period': str (normalizado),
            'quarter': int | None,
            'year': int | None,
            'expected_months': [int],
            'detected_months': [(mo, yr, filename), ...],
            'warnings': [str],
            'errors': [str],
        }
    """
    result = {
        'ok': True, 'period': period,
        'quarter': None, 'year': None,
        'expected_months': [],
        'detected_months': [],
        'warnings': [], 'errors': [],
    }
    q, yr = _parse_period_label(period)
    if not q or not yr:
        result['ok'] = False
        result['errors'].append(f"Formato de período inválido: {period!r} (esperado 'Q1 2025'..'Q4 2025').")
        return result
    result['quarter'] = q
    result['year'] = yr
    start = (q - 1) * 3 + 1
    result['expected_months'] = [start, start + 1, start + 2]

    # Detectar mes de cada archivo desde period_label o filename
    for f in ingested_files:
        src = getattr(f, 'period_label', '') or ''
        mo, yy = _parse_month_year_from_text(src)
        if not mo:
            fname = getattr(f, 'filename', '') or ''
            mo, yy = _parse_month_year_from_text(fname)
        result['detected_months'].append((mo, yy, getattr(f, 'filename', '?')))

    # Validaciones
    n_files = len(ingested_files)
    if n_files != 3:
        result['warnings'].append(
            f"Se esperan 3 archivos mensuales para un quarter; cargaste {n_files}."
        )

    # Meses detectados fuera del quarter o del año
    expected_set = set(result['expected_months'])
    for mo, yy, fn in result['detected_months']:
        if mo is None:
            result['warnings'].append(f"No pude detectar el mes de '{fn}'. Verificá el período manualmente.")
            continue
        if yy != yr:
            result['errors'].append(
                f"'{fn}' tiene año {yy}, pero el período declarado es {yr}."
            )
            result['ok'] = False
        if mo not in expected_set:
            result['errors'].append(
                f"'{fn}' es mes {mo}, fuera del quarter {q} (esperado meses {result['expected_months']})."
            )
            result['ok'] = False

    # Cobertura: ideal es cubrir los 3 meses del quarter
    detected_in_q = {mo for mo, yy, _ in result['detected_months'] if mo in expected_set and yy == yr}
    missing = [m for m in result['expected_months'] if m not in detected_in_q]
    if missing:
        result['warnings'].append(
            f"Faltan archivos para los meses {missing} del {period}."
        )
    return result


def _prev_period(period: str) -> str:
    """Deriva el período anterior: 'Q2 2025' -> 'Q1 2025', 'Q1 2025' -> 'Q4 2024'."""
    m = re.match(r'Q(\d)\s+(\d{4})', period.strip(), re.IGNORECASE)
    if m:
        q, year = int(m.group(1)), int(m.group(2))
        if q == 1:
            return f"Q4 {year - 1}"
        return f"Q{q - 1} {year}"
    return ""


def _cleanup_temp_files(paths: list):
    """Limpia archivos temporales."""
    for p in paths:
        try:
            os.unlink(p)
        except OSError:
            pass


# Lado P&L al que pertenece cada L1. Se usa para detectar reclasificaciones
# "cross-side" (Income ↔ Expense) donde la convención de signo cambia y hay
# que flipear el valor del L3 para mantener NOI consistente.
#
# Ejemplo canónico: `8162-2000 Gas Utility Reimburse Inc = -1170` vive bajo
# Operating Expenses como contra-expense. Si el analista la mueve a Income
# sin flip, el -1170 bajaría Income Y subiría OpEx (doble castigo en NOI).
# Con flip, se transforma en +1170 income y OpEx pierde el contra (neutro en NOI).
_SIDE_INCOME = "income"
_SIDE_EXPENSE = "expense"
_SIDE_FOR_L1 = {
    "Income": _SIDE_INCOME,
    "Operating Expenses": _SIDE_EXPENSE,
    "Real Estate Taxes": _SIDE_EXPENSE,
    "Interest Expense": _SIDE_EXPENSE,
    "Non-Operating Expenses": _SIDE_EXPENSE,
    "Total Capex": _SIDE_EXPENSE,
    # NA queda fuera: sus moves no flipean (la cuenta se excluye del total)
}


def _is_cross_side(old_l1: str, new_l1: str) -> bool:
    """True si la reclasificación cruza la frontera Income/Expense del P&L."""
    old_side = _SIDE_FOR_L1.get(old_l1)
    new_side = _SIDE_FOR_L1.get(new_l1)
    if old_side is None or new_side is None:
        return False
    return old_side != new_side


def _flip_l3_signs(acct) -> None:
    """Multiplica por -1 los campos numéricos de un L3 (mut in-place)."""
    for field in ("budget_current", "actual_current", "budget_ytd", "actual_ytd",
                  "variance_current", "variance_ytd"):
        val = getattr(acct, field, None)
        if val is not None:
            setattr(acct, field, -val)


def _apply_l1_overrides(structure: dict, overrides: dict):
    """
    Aplica reclasificaciones del analista a la estructura financiera.
    Modifica parent_line de cuentas L3 y recalcula L1 lines.

    Si una reclasificación cruza la frontera Income/Expense, flipea el signo
    de los valores del L3 para preservar la consistencia de NOI.
    """
    if not overrides:
        return

    # Capturar L1s afectados (antes + después) para recalcular SOLO esos.
    # L1s no tocados por overrides mantienen su total reportado por el socio.
    affected_l1s = set()
    for acct in structure.get('l3_accounts', []):
        if acct.account_code in overrides:
            old_l1 = acct.parent_line
            new_l1 = overrides[acct.account_code]
            affected_l1s.add(old_l1)
            affected_l1s.add(new_l1)
            # Flip sign BEFORE updating parent_line (lee old_l1 vs new_l1)
            crossed = _is_cross_side(old_l1, new_l1)
            if crossed:
                _flip_l3_signs(acct)
            acct.parent_line = new_l1
            # Audit trail: quién decidió esta clasificación final
            acct.classification_source = 'cross_side_flip' if crossed else 'user_override'
    affected_l1s.discard('NA')

    # Reclasificar L2 sections cuyos hijos cambiaron y recalcular L2 values
    from collections import Counter
    for section in structure.get('l2_sections', []):
        # Paso 1: identificar todos los L3 que originalmente pertenecían a esta sección (por section_name).
        section_children_raw = [
            a for a in structure.get('l3_accounts', [])
            if a.section_name and a.section_name.lower() == section.name.lower()
        ]
        if not section_children_raw:
            continue

        # Paso 2: recalcular parent_line de la sección según mayoría de sus hijos (ya con overrides aplicados).
        child_lines = [a.parent_line for a in section_children_raw if a.parent_line]
        if child_lines:
            majority = Counter(child_lines).most_common(1)[0][0]
            section.parent_line = majority

        # Paso 3 (FIX bug The Val): solo sumar al L2 los L3s cuyo parent_line coincida con el del L2.
        # Sin este filtro, si un L3 se mueve cross-side (ej. OpEx→Income con sign-flip),
        # la sección L2 original seguía sumándolo con signo invertido, descuadrando el L2.
        section_children = [
            a for a in section_children_raw if a.parent_line == section.parent_line
        ]
        if not section_children:
            section.budget_current = 0
            section.actual_current = 0
            section.variance_current = 0
            section.variance_pct_current = None
            continue

        # Recalcular valores de la sección desde sus hijos L3 (solo los que aún pertenecen)
        section.budget_current = sum(a.budget_current or 0 for a in section_children)
        section.actual_current = sum(a.actual_current or 0 for a in section_children)
        section.variance_current = section.actual_current - section.budget_current
        section.variance_pct_current = (
            section.variance_current / abs(section.budget_current)
            if section.budget_current and abs(section.budget_current) > 0.01
            else None
        )
        has_ytd = any(a.actual_ytd is not None for a in section_children)
        if has_ytd:
            section.budget_ytd = sum(a.budget_ytd or 0 for a in section_children)
            section.actual_ytd = sum(a.actual_ytd or 0 for a in section_children)
            section.variance_ytd = section.actual_ytd - section.budget_ytd
            section.variance_pct_ytd = (
                section.variance_ytd / abs(section.budget_ytd)
                if section.budget_ytd and abs(section.budget_ytd) > 0.01
                else None
            )

    # Recalcular L1 lines afectados sumando sus L3 (NA se excluye).
    l1_sums = defaultdict(lambda: {
        'budget_current': 0, 'actual_current': 0,
        'budget_ytd': 0, 'actual_ytd': 0,
    })
    # Pre-populate SOLO los L1s afectados — así líneas que queden sin hijos van a 0
    for name in affected_l1s:
        _ = l1_sums[name]

    for acct in structure.get('l3_accounts', []):
        l1 = acct.parent_line
        if l1 == 'NA':
            continue
        if l1 not in affected_l1s:
            continue
        l1_sums[l1]['budget_current'] += acct.budget_current or 0
        l1_sums[l1]['actual_current'] += acct.actual_current or 0
        if acct.budget_ytd is not None:
            l1_sums[l1]['budget_ytd'] += acct.budget_ytd or 0
        if acct.actual_ytd is not None:
            l1_sums[l1]['actual_ytd'] += acct.actual_ytd or 0

    # Actualizar SOLO L1 lines afectados por los overrides
    existing_l1_names = {l.name for l in structure.get('l1_lines', [])}
    for l1 in structure.get('l1_lines', []):
        if l1.name in l1_sums and l1.name not in ('NOI', 'Net Income'):
            s = l1_sums[l1.name]
            l1.budget_current = s['budget_current']
            l1.actual_current = s['actual_current']
            l1.variance_current = l1.actual_current - l1.budget_current
            l1.variance_pct_current = (
                l1.variance_current / abs(l1.budget_current)
                if l1.budget_current and abs(l1.budget_current) > 0.01
                else None
            )
            if s['budget_ytd'] or s['actual_ytd']:
                l1.budget_ytd = s['budget_ytd']
                l1.actual_ytd = s['actual_ytd']
                l1.variance_ytd = l1.actual_ytd - l1.budget_ytd
                l1.variance_pct_ytd = (
                    l1.variance_ytd / abs(l1.budget_ytd)
                    if l1.budget_ytd and abs(l1.budget_ytd) > 0.01
                    else None
                )

    # Crear L1 lines nuevas si los overrides movieron cuentas a un L1 que no existía
    # (caso típico: Total Capex no estaba en el reporte original pero el analista
    # reclasificó cuentas hacia ahí).
    from src.native_structure import FinancialLine
    for new_l1_name in affected_l1s:
        if new_l1_name in existing_l1_names or new_l1_name in ('NOI', 'Net Income', 'NA', 'Unknown'):
            continue
        if new_l1_name not in l1_sums:
            continue
        s = l1_sums[new_l1_name]
        new_line = FinancialLine(
            name=new_l1_name,
            level=1,
            budget_current=s['budget_current'],
            actual_current=s['actual_current'],
            variance_current=s['actual_current'] - s['budget_current'],
            variance_pct_current=(
                (s['actual_current'] - s['budget_current']) / abs(s['budget_current'])
                if s['budget_current'] and abs(s['budget_current']) > 0.01 else None
            ),
            budget_ytd=s['budget_ytd'] if s['budget_ytd'] or s['actual_ytd'] else None,
            actual_ytd=s['actual_ytd'] if s['budget_ytd'] or s['actual_ytd'] else None,
            variance_ytd=(s['actual_ytd'] - s['budget_ytd']) if s['budget_ytd'] or s['actual_ytd'] else None,
            variance_pct_ytd=None,
            parent_line=new_l1_name,
            classification_source='user_override',
        )
        structure['l1_lines'].append(new_line)

    # Recalcular NOI = Income - |OpEx| - |Real Estate Taxes| si existen los componentes
    l1_map = {l.name: l for l in structure.get('l1_lines', [])}
    if 'NOI' in l1_map and 'Income' in l1_map and 'Operating Expenses' in l1_map:
        inc = l1_map['Income']
        opex = l1_map['Operating Expenses']
        taxes = l1_map.get('Real Estate Taxes')
        noi = l1_map['NOI']
        noi.actual_current = inc.actual_current - abs(opex.actual_current) - (abs(taxes.actual_current) if taxes else 0)
        noi.budget_current = inc.budget_current - abs(opex.budget_current) - (abs(taxes.budget_current) if taxes else 0)
        noi.variance_current = noi.actual_current - noi.budget_current
        noi.variance_pct_current = (
            noi.variance_current / abs(noi.budget_current)
            if noi.budget_current and abs(noi.budget_current) > 0.01
            else None
        )
        # YTD también
        if inc.actual_ytd is not None and opex.actual_ytd is not None:
            noi.actual_ytd = inc.actual_ytd - abs(opex.actual_ytd) - (abs(taxes.actual_ytd or 0) if taxes else 0)
        if inc.budget_ytd is not None and opex.budget_ytd is not None:
            noi.budget_ytd = inc.budget_ytd - abs(opex.budget_ytd) - (abs(taxes.budget_ytd or 0) if taxes else 0)
        if noi.actual_ytd is not None and noi.budget_ytd is not None:
            noi.variance_ytd = noi.actual_ytd - noi.budget_ytd

    # Recalcular Net Income = NOI - |Interest| - |Non-Operating|
    # Capex y NA NO entran en el waterfall de NI (quedan fuera del resultado).
    if 'Net Income' in l1_map and 'NOI' in l1_map:
        noi_l = l1_map['NOI']
        interest = l1_map.get('Interest Expense')
        non_op = l1_map.get('Non-Operating Expenses')
        ni = l1_map['Net Income']
        ni.actual_current = (
            noi_l.actual_current
            - (abs(interest.actual_current) if interest else 0)
            - (abs(non_op.actual_current) if non_op else 0)
        )
        ni.budget_current = (
            noi_l.budget_current
            - (abs(interest.budget_current) if interest else 0)
            - (abs(non_op.budget_current) if non_op else 0)
        )
        ni.variance_current = ni.actual_current - ni.budget_current
        ni.variance_pct_current = (
            ni.variance_current / abs(ni.budget_current)
            if ni.budget_current and abs(ni.budget_current) > 0.01
            else None
        )
        if noi_l.actual_ytd is not None:
            ni.actual_ytd = (
                noi_l.actual_ytd
                - (abs(interest.actual_ytd or 0) if interest else 0)
                - (abs(non_op.actual_ytd or 0) if non_op else 0)
            )
        if noi_l.budget_ytd is not None:
            ni.budget_ytd = (
                noi_l.budget_ytd
                - (abs(interest.budget_ytd or 0) if interest else 0)
                - (abs(non_op.budget_ytd or 0) if non_op else 0)
            )
        if ni.actual_ytd is not None and ni.budget_ytd is not None:
            ni.variance_ytd = ni.actual_ytd - ni.budget_ytd


def revalidate_structure(results: dict) -> list:
    """
    Re-ejecuta validaciones sobre la estructura ya reclasificada (después de overrides).
    Actualiza results['validations'] in-place y retorna la nueva lista.
    """
    from src.validation import validate_all
    prior_struct = results.get('prior_structure')
    new_validations = validate_all(results['structure'], prior_struct)
    results['validations'] = new_validations
    return new_validations


def regenerate_outputs(results: dict) -> None:
    """
    Regenera TODOS los outputs derivados de la estructura (markdown, copilot_prompt,
    json_payload, excel_bytes, tracking_excel, alerts, comments_df, summary_table,
    bullets). Muta `results` in-place.

    Diseñado para invocarse DESPUÉS de que el analista aplica overrides en Step 2.
    Garantiza que todos los consumidores downstream (Step 3, AI, exports) vean la
    clasificación final — no un estado intermedio.
    """
    structure = results['structure']
    building = results.get('building', '')
    period = results.get('period', '')
    validations = results.get('validations', [])
    historical_analysis = results.get('historical_analysis', {'has_prior': False})

    # 1. Summary table + bullets (ya los recalcula step2 hoy, los incluyo por completitud)
    results['summary_table'] = build_summary_table(structure)
    results['bullets'] = build_bullets_by_line(structure)

    # 2. Alerts + comments (nota: la key usada en pipeline es 'alerts_text')
    results['alerts_text'] = build_alerts_section(validations, structure)
    results['comments_df'] = build_comments_table(structure)

    # 3. Markdown (depende de summary + bullets + alerts + comments + structure)
    results['markdown'] = generate_full_markdown(
        building=building,
        period=period,
        summary_table=results['summary_table'],
        bullets=results['bullets'],
        alerts=results['alerts_text'],
        comments_df=results['comments_df'],
        structure=structure,
    )

    # 4. Excels de soporte (ambos leen structure directamente)
    results['excel_bytes'] = generate_support_excel(
        results['summary_table'], structure, results['comments_df'], validations
    )
    results['tracking_excel'] = generate_account_tracking_excel(
        structure, historical_analysis, building, period
    )

    # 5. Copilot prompt + JSON payload (este es el CRÍTICO: el AI ve esto)
    results['json_payload'] = generate_copilot_payload(
        building, period, structure, historical_analysis
    )
    results['copilot_prompt'] = generate_copilot_prompt(
        building, period, structure, historical_analysis
    )


# ── Pipeline principal ─────────────────────────────────────────


def run_pipeline(inputs: PipelineInput) -> dict:
    """
    Ejecuta el pipeline completo y retorna un dict con todos los resultados.
    Lanza excepciones con mensajes claros si algo falla.
    """
    warnings = []

    # --- 1. GUARDAR ARCHIVOS TEMPORALES ---
    current_paths = _save_uploaded_files(inputs.files_current, "current")
    prior_paths = _save_uploaded_files(inputs.files_prior, "prior") if inputs.files_prior else []

    try:
        # --- 2. INGESTA ---
        try:
            ingested_current = ingest_files(
                current_paths,
                sheet_selections=inputs.sheet_selections,
                target_period=inputs.period,
            )
        except Exception as e:
            raise ValueError(f"Error en ingesta de archivos actuales: {e}")

        try:
            prior_period = _prev_period(inputs.period) if inputs.period else None
            ingested_prior = ingest_files(prior_paths, target_period=prior_period) if prior_paths else []
        except Exception as e:
            warnings.append(f"Ingesta trimestre anterior: {e}")
            ingested_prior = []

        # --- AUTO-DETECTAR BUILDING si el usuario no lo especificó ---
        building = inputs.building
        period = inputs.period

        if not building.strip() and ingested_current:
            building = ingested_current[0].building or "Unknown Property"

        # --- AUTO-DETECTAR PERÍODO del Excel ---
        if not period.strip() and ingested_current:
            period = _detect_period_from_files(ingested_current)

        # --- 3. ESTRUCTURA NATIVA ---
        by_building = defaultdict(list)  # populated below if consolidate mode
        try:
            if inputs.consolidate and len(ingested_current) > 0:
                for ing in ingested_current:
                    bname = ing.building or "Unknown"
                    by_building[bname].append(ing)

                if len(by_building) > 1:
                    individual_structures = {}
                    for bname, files in by_building.items():
                        individual_structures[bname] = build_financial_structure(files)
                    structure = consolidate_structures(
                        list(individual_structures.values()),
                        portfolio_name=building or "Portfolio",
                    )
                    structure['individual_structures'] = individual_structures
                else:
                    structure = build_financial_structure(ingested_current)
            else:
                structure = build_financial_structure(ingested_current)
        except Exception as e:
            raise ValueError(f"Error construyendo estructura financiera: {e}")

        try:
            prior_structure = build_financial_structure(ingested_prior) if ingested_prior else None
        except Exception as e:
            warnings.append(f"Estructura trimestre anterior: {e}")
            prior_structure = None

        # --- 4. VALIDACIÓN ---
        try:
            validations = validate_all(structure, prior_structure)
        except Exception as e:
            warnings.append(f"Validación: {e}")
            validations = []

        # --- 5. ANÁLISIS HISTÓRICO (QoQ) ---
        try:
            historical_analysis = build_historical_analysis(structure, prior_structure)
        except Exception as e:
            warnings.append(f"Análisis histórico: {e}")
            historical_analysis = {'has_prior': False}

        # --- 6. OUTPUTS ---
        try:
            summary_table = build_summary_table(structure)
            bullets = build_bullets_by_line(structure)
            alerts_text = build_alerts_section(validations, structure)
            comments_df = build_comments_table(structure)

            markdown = generate_full_markdown(
                building=building,
                period=period,
                summary_table=summary_table,
                bullets=bullets,
                alerts=alerts_text,
                comments_df=comments_df,
                structure=structure,
            )

            excel_bytes = generate_support_excel(summary_table, structure, comments_df, validations)
            tracking_excel = generate_account_tracking_excel(
                structure, historical_analysis, building, period
            )
            json_payload = generate_copilot_payload(building, period, structure, historical_analysis)
            copilot_prompt = generate_copilot_prompt(building, period, structure, historical_analysis)
        except Exception as e:
            raise ValueError(f"Error generando outputs: {e}")

        # --- 7. MASTER EXCEL (budget tracking por edificio) ---
        master_path = None
        master_bytes = None
        try:
            import shutil
            # Si el usuario subió un master existente, usarlo como base
            if inputs.master_upload_path:
                master_dir = "./masters"
                os.makedirs(master_dir, exist_ok=True)
                from src.master_excel import _sanitize_building_name, _extract_year_from_period
                safe_name = _sanitize_building_name(building)
                year = _extract_year_from_period(period)
                dest = os.path.join(master_dir, f"master_{safe_name}_{year}.xlsx")
                shutil.copy2(str(inputs.master_upload_path), dest)

            # Build per-month structures for monthly Master columns
            monthly_structures = []
            if inputs.consolidate and len(by_building) > 1:
                # Portfolio: consolidate all buildings per quarter-month position
                # by_building[building] is ordered by upload order (assumed: m1, m2, m3)
                n_months = 3
                for month_idx in range(n_months):
                    month_structs = []
                    for bname, bfiles in by_building.items():
                        if month_idx < len(bfiles):
                            try:
                                ms = build_financial_structure([bfiles[month_idx]])
                                month_structs.append(ms)
                            except Exception:
                                pass
                    if month_structs:
                        try:
                            if len(month_structs) > 1:
                                cm = consolidate_structures(
                                    month_structs, portfolio_name=building or "Portfolio"
                                )
                            else:
                                cm = month_structs[0]
                            monthly_structures.append(cm)
                        except Exception:
                            monthly_structures.append(None)
                    else:
                        monthly_structures.append(None)
            else:
                # Single building: ordenar files cronológicamente por el
                # (mes, año) que detectamos de su period_label o filename.
                # Antes se asumía que venían en orden; Streamlit puede
                # entregarlos en cualquier orden.
                from src.native_structure import _extract_period_sortkey
                ordered_files = sorted(ingested_current, key=_extract_period_sortkey)
                for ingested_file in ordered_files:
                    try:
                        month_struct = build_financial_structure([ingested_file])
                        # Propagar period_label + filename para que master_excel
                        # sepa a qué M column corresponde este mes.
                        month_struct['_period_label'] = ingested_file.period_label
                        month_struct['_filename'] = ingested_file.filename
                        monthly_structures.append(month_struct)
                    except Exception:
                        monthly_structures.append(None)

            # Si hay período anterior, escribirlo primero
            if prior_structure:
                prev_p = _prev_period(period)
                if prev_p:
                    update_master_excel(
                        building=building,
                        period=prev_p,
                        structure=prior_structure,
                        master_dir="./masters",
                        user="analyst",
                    )

            master_path = update_master_excel(
                building=building,
                period=period,
                structure=structure,
                master_dir="./masters",
                monthly_structures=monthly_structures if monthly_structures else None,
                user="analyst",
            )
            # Leer bytes del master actualizado para descarga
            if master_path and os.path.exists(master_path):
                with open(master_path, 'rb') as f:
                    master_bytes = f.read()
        except Exception as e:
            warnings.append(f"Master Excel: {e}")

        # --- 8. ANÁLISIS DE MERCADO (opcional) ---
        market = None
        if inputs.market_file:
            try:
                market_paths = _save_uploaded_files([inputs.market_file], "market")
                market = analyze_market(
                    filepath=market_paths[0],
                    report_period=period,
                    market_name=inputs.market_name or "Market",
                )
                _cleanup_temp_files(market_paths)
            except Exception as e:
                warnings.append(f"Análisis de mercado: {e}")
                market = None

        return {
            'building': building,
            'period': period,
            'ingested_files': ingested_current,
            'structure': structure,
            'prior_structure': prior_structure,
            'validations': validations,
            'historical_analysis': historical_analysis,
            'market_analysis': market,
            'summary_table': summary_table,
            'bullets': bullets,
            'alerts_text': alerts_text,
            'comments_df': comments_df,
            'markdown': markdown,
            'excel_bytes': excel_bytes,
            'tracking_excel': tracking_excel,
            'json_payload': json_payload,
            'copilot_prompt': copilot_prompt,
            'master_path': master_path,
            'master_bytes': master_bytes,
            'warnings': warnings,
        }

    finally:
        # Limpiar archivos temporales
        _cleanup_temp_files(current_paths + prior_paths)
