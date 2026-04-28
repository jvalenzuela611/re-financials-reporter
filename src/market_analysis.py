"""
Market Analysis Module — Multifamily Data Grid
Loads CoStar-style market data and produces KPIs, trends, and context
for the quarterly financial report.
"""

import pandas as pd
import re
from dataclasses import dataclass, field
from typing import Optional


# ─── Data Classes ───────────────────────────────────────────

@dataclass
class MarketSnapshot:
    """Single-quarter market snapshot."""
    period: str
    year: int
    quarter: int

    inventory_units: int = 0
    asking_rent: float = 0.0
    asking_rent_psf: float = 0.0
    asking_rent_growth: Optional[float] = None
    effective_rent: float = 0.0
    effective_rent_psf: float = 0.0
    effective_rent_growth: Optional[float] = None
    concessions_pct: float = 0.0
    vacancy_pct: float = 0.0
    vacancy_growth: Optional[float] = None
    occupancy_pct: float = 0.0
    absorption_units: int = 0
    under_construction_units: int = 0
    under_construction_pct: float = 0.0
    deliveries_units: int = 0
    deliveries_pct: float = 0.0


@dataclass
class MarketAnalysis:
    """Full market analysis result."""
    market_name: str
    current: MarketSnapshot
    prior_quarter: Optional[MarketSnapshot] = None
    prior_year: Optional[MarketSnapshot] = None
    history: list = field(default_factory=list)  # list[MarketSnapshot]
    df: Optional[pd.DataFrame] = None  # full dataframe for charting

    # Computed summaries
    rent_trend_1yr: Optional[float] = None
    vacancy_trend_1yr: Optional[float] = None
    absorption_avg_4q: Optional[float] = None
    supply_pipeline_pct: Optional[float] = None


# ─── Parsing ────────────────────────────────────────────────

_PERIOD_RE = re.compile(r'(\d{4})\s*Q(\d)', re.IGNORECASE)


def _parse_period(text: str) -> tuple:
    """Extract (year, quarter) from period string like '2025 Q4' or '2026 Q1 QTD'."""
    m = _PERIOD_RE.search(str(text).strip())
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def _safe_float(val, default=None):
    try:
        v = float(val)
        return v if pd.notna(v) else default
    except (ValueError, TypeError):
        return default


def _safe_int(val, default=0):
    try:
        v = int(float(val))
        return v if pd.notna(float(val)) else default
    except (ValueError, TypeError):
        return default


# ─── Column Mapping ─────────────────────────────────────────

_COL_MAP = {
    'period':                 [r'period', r'fecha', r'quarter'],
    'inventory_units':        [r'inventory\s*units'],
    'asking_rent':            [r'asking\s*rent\s*per\s*unit'],
    'asking_rent_psf':        [r'asking\s*rent\s*per\s*sf'],
    'asking_rent_growth':     [r'asking\s*rent\s*%?\s*growth'],
    'effective_rent':         [r'effective\s*rent\s*per\s*unit'],
    'effective_rent_psf':     [r'effective\s*rent\s*per\s*sf'],
    'effective_rent_growth':  [r'effective\s*rent\s*%?\s*growth'],
    'concessions_pct':        [r'concessions?\s*%'],
    'vacancy_pct':            [r'vacancy\s*percent'],
    'vacancy_growth':         [r'vacancy\s*%?\s*growth'],
    'occupancy_pct':          [r'occupancy\s*percent'],
    'absorption_units':       [r'absorption\s*units'],
    'under_construction':     [r'under\s*construction\s*units'],
    'under_construction_pct': [r'under\s*construction\s*percent'],
    'deliveries_units':       [r'deliveries?\s*units'],
    'deliveries_pct':         [r'deliveries?\s*percent'],
}


def _map_columns(df: pd.DataFrame) -> dict:
    """Map flexible column names to internal keys."""
    mapping = {}
    cols = list(df.columns)
    for key, patterns in _COL_MAP.items():
        for col in cols:
            col_clean = str(col).strip().lower()
            for pat in patterns:
                if re.search(pat, col_clean):
                    mapping[key] = col
                    break
            if key in mapping:
                break
    return mapping


# ─── Core Functions ─────────────────────────────────────────

def load_market_data(filepath: str) -> pd.DataFrame:
    """Load a CoStar-style multifamily data grid Excel file."""
    df = pd.read_excel(filepath)
    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _row_to_snapshot(row, col_map: dict) -> MarketSnapshot:
    period_str = str(row.get(col_map.get('period', ''), ''))
    year, quarter = _parse_period(period_str)

    return MarketSnapshot(
        period=period_str,
        year=year or 0,
        quarter=quarter or 0,
        inventory_units=_safe_int(row.get(col_map.get('inventory_units', ''), 0)),
        asking_rent=_safe_float(row.get(col_map.get('asking_rent', ''), 0), 0),
        asking_rent_psf=_safe_float(row.get(col_map.get('asking_rent_psf', ''), 0), 0),
        asking_rent_growth=_safe_float(row.get(col_map.get('asking_rent_growth', ''))),
        effective_rent=_safe_float(row.get(col_map.get('effective_rent', ''), 0), 0),
        effective_rent_psf=_safe_float(row.get(col_map.get('effective_rent_psf', ''), 0), 0),
        effective_rent_growth=_safe_float(row.get(col_map.get('effective_rent_growth', ''))),
        concessions_pct=_safe_float(row.get(col_map.get('concessions_pct', ''), 0), 0),
        vacancy_pct=_safe_float(row.get(col_map.get('vacancy_pct', ''), 0), 0),
        vacancy_growth=_safe_float(row.get(col_map.get('vacancy_growth', ''))),
        occupancy_pct=_safe_float(row.get(col_map.get('occupancy_pct', ''), 0), 0),
        absorption_units=_safe_int(row.get(col_map.get('absorption_units', ''), 0)),
        under_construction_units=_safe_int(row.get(col_map.get('under_construction', ''), 0)),
        under_construction_pct=_safe_float(row.get(col_map.get('under_construction_pct', ''), 0), 0),
        deliveries_units=_safe_int(row.get(col_map.get('deliveries_units', ''), 0)),
        deliveries_pct=_safe_float(row.get(col_map.get('deliveries_pct', ''), 0), 0),
    )


def analyze_market(
    filepath: str,
    report_period: str = "",
    market_name: str = "",
) -> MarketAnalysis:
    """
    Main entry: load market data, find the target quarter, compute trends.

    Parameters:
        filepath: Path to the Excel market data grid
        report_period: e.g. "Q1 2026" — used to align the analysis
        market_name: e.g. "Oakland" — label for display
    """
    df = load_market_data(filepath)
    col_map = _map_columns(df)

    if 'period' not in col_map:
        raise ValueError("No se encontró columna de período en el archivo de mercado.")

    # Parse all periods and sort chronologically
    periods = []
    for idx, row in df.iterrows():
        y, q = _parse_period(row[col_map['period']])
        if y and q:
            periods.append((y, q, idx))
    periods.sort(key=lambda x: (x[0], x[1]), reverse=True)

    # Build snapshots (most recent first)
    snapshots = []
    for y, q, idx in periods:
        snap = _row_to_snapshot(df.loc[idx], col_map)
        snapshots.append(snap)

    if not snapshots:
        raise ValueError("No se encontraron datos de mercado válidos.")

    # Find the target quarter from report_period
    target_y, target_q = _parse_period(report_period)

    current = None
    if target_y and target_q:
        # Find exact match or closest earlier quarter
        for s in snapshots:
            if s.year == target_y and s.quarter == target_q:
                current = s
                break
            # Also check QTD variant
            if s.year == target_y and s.quarter == target_q:
                current = s
                break

    # Fallback: use most recent
    if current is None:
        current = snapshots[0]

    # Find prior quarter and prior year same quarter
    prior_q = None
    prior_y = None
    for s in snapshots:
        if s is current:
            continue
        if prior_q is None and (s.year, s.quarter) < (current.year, current.quarter):
            prior_q = s
        if (prior_y is None
                and s.year == current.year - 1
                and s.quarter == current.quarter):
            prior_y = s
        if prior_q and prior_y:
            break

    # Compute trend summaries
    rent_trend_1yr = None
    if prior_y and prior_y.asking_rent > 0:
        rent_trend_1yr = (current.asking_rent - prior_y.asking_rent) / prior_y.asking_rent

    vacancy_trend_1yr = None
    if prior_y:
        vacancy_trend_1yr = current.vacancy_pct - prior_y.vacancy_pct

    # Average absorption last 4 quarters
    recent_4 = [s for s in snapshots if (s.year, s.quarter) <= (current.year, current.quarter)][:4]
    absorption_avg = None
    if recent_4:
        absorption_avg = sum(s.absorption_units for s in recent_4) / len(recent_4)

    # Build a clean dataframe for charting (chronological order)
    chart_data = []
    for s in reversed(snapshots):
        chart_data.append({
            'Period': s.period.replace(' QTD', ''),
            'Year': s.year,
            'Quarter': s.quarter,
            'Asking Rent': s.asking_rent,
            'Effective Rent': s.effective_rent,
            'Vacancy %': s.vacancy_pct * 100,
            'Occupancy %': s.occupancy_pct * 100,
            'Absorption': s.absorption_units,
            'Under Construction': s.under_construction_units,
            'Deliveries': s.deliveries_units,
            'Concessions %': s.concessions_pct * 100,
            'Asking Rent PSF': s.asking_rent_psf,
        })
    chart_df = pd.DataFrame(chart_data)

    return MarketAnalysis(
        market_name=market_name or "Market",
        current=current,
        prior_quarter=prior_q,
        prior_year=prior_y,
        history=snapshots,
        df=chart_df,
        rent_trend_1yr=rent_trend_1yr,
        vacancy_trend_1yr=vacancy_trend_1yr,
        absorption_avg_4q=absorption_avg,
        supply_pipeline_pct=current.under_construction_pct,
    )


# ─── Display Helpers ────────────────────────────────────────

def market_kpi_cards(analysis: MarketAnalysis) -> list:
    """Return list of dicts for KPI display: {label, value, delta, delta_color}."""
    c = analysis.current
    pq = analysis.prior_quarter

    cards = []

    # Asking Rent
    delta_rent = None
    if pq and pq.asking_rent > 0:
        delta_rent = (c.asking_rent - pq.asking_rent) / pq.asking_rent
    cards.append({
        'label': 'Asking Rent / Unit',
        'value': f'${c.asking_rent:,.0f}',
        'delta': f'{delta_rent:+.1%}' if delta_rent is not None else None,
        'delta_color': 'normal' if delta_rent and delta_rent >= 0 else 'inverse',
    })

    # Effective Rent
    delta_eff = None
    if pq and pq.effective_rent > 0:
        delta_eff = (c.effective_rent - pq.effective_rent) / pq.effective_rent
    cards.append({
        'label': 'Effective Rent / Unit',
        'value': f'${c.effective_rent:,.0f}',
        'delta': f'{delta_eff:+.1%}' if delta_eff is not None else None,
        'delta_color': 'normal' if delta_eff and delta_eff >= 0 else 'inverse',
    })

    # Vacancy
    delta_vac = None
    if pq:
        delta_vac = c.vacancy_pct - pq.vacancy_pct
    cards.append({
        'label': 'Vacancy Rate',
        'value': f'{c.vacancy_pct:.1%}',
        'delta': f'{delta_vac:+.1%}' if delta_vac is not None else None,
        'delta_color': 'inverse',  # lower vacancy = good
    })

    # Occupancy
    cards.append({
        'label': 'Occupancy Rate',
        'value': f'{c.occupancy_pct:.1%}',
        'delta': f'{(c.occupancy_pct - pq.occupancy_pct):+.1%}' if pq else None,
        'delta_color': 'normal',
    })

    # Absorption
    cards.append({
        'label': 'Absorption (units)',
        'value': f'{c.absorption_units:,}',
        'delta': f'Avg 4Q: {analysis.absorption_avg_4q:,.0f}' if analysis.absorption_avg_4q else None,
        'delta_color': 'off',
    })

    # Supply Pipeline
    cards.append({
        'label': 'Under Construction',
        'value': f'{c.under_construction_units:,} units',
        'delta': f'{c.under_construction_pct:.1%} of inventory' if c.under_construction_pct else None,
        'delta_color': 'off',
    })

    return cards


def market_summary_text(analysis: MarketAnalysis) -> str:
    """Generate a narrative paragraph summarizing market conditions."""
    c = analysis.current
    lines = []

    lines.append(
        f"**{analysis.market_name}** — {c.period}: "
        f"El mercado tiene un inventario de **{c.inventory_units:,} unidades** multifamiliares."
    )

    # Rent trend
    if analysis.rent_trend_1yr is not None:
        direction = "subieron" if analysis.rent_trend_1yr > 0 else "bajaron"
        lines.append(
            f"Las rentas asking {direction} **{abs(analysis.rent_trend_1yr):.1%}** YoY "
            f"a **${c.asking_rent:,.0f}/unidad**."
        )

    # Vacancy
    lines.append(
        f"Vacancy se ubica en **{c.vacancy_pct:.1%}** (occupancy {c.occupancy_pct:.1%})."
    )

    # Concessions
    if c.concessions_pct > 0:
        lines.append(f"Concesiones promedio: **{c.concessions_pct:.1%}**.")

    # Supply
    if c.under_construction_units > 0:
        lines.append(
            f"Pipeline de construcción: **{c.under_construction_units:,} unidades** "
            f"({c.under_construction_pct:.1%} del inventario)."
        )

    # Deliveries
    if c.deliveries_units > 0:
        lines.append(f"Entregas recientes: **{c.deliveries_units:,} unidades** este trimestre.")

    return " ".join(lines)
