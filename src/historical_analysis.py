"""
Módulo de Análisis Histórico: comparación QoQ, detección de tendencias,
aceleración de desviaciones, y tracking de cuentas nuevas/eliminadas.

Cuando se proporcionan archivos del trimestre anterior, este módulo:
1. Construye series por cuenta (Current Q vs Prior Q)
2. Calcula variación QoQ (Quarter-over-Quarter)
3. Detecta cambios de tendencia y aceleración de desviaciones
4. Genera output estructurado para Copilot y markdown

Diseñado para escalar a T12 (trailing 12 months) cuando se conecte SharePoint.
"""

import pandas as pd
from typing import Dict, List, Optional
from dataclasses import dataclass, field


@dataclass
class AccountTimeSeries:
    """Serie temporal de una cuenta individual."""
    account_code: str
    description: str
    parent_line: str  # L1 parent (Income, OpEx, etc.)
    # Current quarter
    budget_q: float
    actual_q: float
    variance_q: float
    variance_pct_q: Optional[float]
    # Prior quarter (if available)
    budget_prior_q: Optional[float] = None
    actual_prior_q: Optional[float] = None
    variance_prior_q: Optional[float] = None
    variance_pct_prior_q: Optional[float] = None
    # QoQ metrics
    qoq_actual_change: Optional[float] = None
    qoq_actual_change_pct: Optional[float] = None
    qoq_variance_change: Optional[float] = None
    # Flags
    is_new_account: bool = False
    is_removed_account: bool = False
    trend_flag: str = ""  # "", "improving", "deteriorating", "accelerating", "reversal"
    partner_comment: str = ""


@dataclass
class QoQSummary:
    """Resumen de análisis Quarter-over-Quarter por línea L1."""
    line_name: str
    actual_current_q: float
    actual_prior_q: Optional[float]
    budget_current_q: float
    budget_prior_q: Optional[float]
    variance_current_q: float
    variance_prior_q: Optional[float]
    qoq_actual_change: Optional[float] = None
    qoq_actual_change_pct: Optional[float] = None
    qoq_variance_change: Optional[float] = None
    trend_direction: str = ""  # "improving", "stable", "deteriorating"
    top_movers: List[Dict] = field(default_factory=list)


@dataclass
class AccountChangeReport:
    """Reporte estructurado de cuentas nuevas y eliminadas."""
    new_accounts: List[Dict] = field(default_factory=list)
    removed_accounts: List[Dict] = field(default_factory=list)
    total_new_impact: float = 0.0  # Impacto en $ de cuentas nuevas
    total_removed_impact: float = 0.0  # Último valor conocido de cuentas eliminadas


def build_historical_analysis(
    current_structure: Dict,
    prior_structure: Optional[Dict] = None,
) -> Dict:
    """
    Construye análisis histórico completo.

    Args:
        current_structure: output de build_financial_structure() para Q actual
        prior_structure: output de build_financial_structure() para Q anterior

    Returns:
        Dict con:
        - qoq_summary: List[QoQSummary] por línea L1
        - account_time_series: List[AccountTimeSeries] para cuentas con cambios significativos
        - account_changes: AccountChangeReport
        - trend_alerts: List[str] alertas de tendencia
    """
    result = {
        'qoq_summary': [],
        'account_time_series': [],
        'account_changes': AccountChangeReport(),
        'trend_alerts': [],
        'has_prior': prior_structure is not None,
    }

    if not prior_structure:
        # Sin histórico, solo devolver estructura básica con cuentas actuales
        result['account_time_series'] = _build_current_only_series(current_structure)
        return result

    # 1. Detectar cuentas nuevas y eliminadas
    result['account_changes'] = _detect_account_changes_structured(
        current_structure, prior_structure
    )

    # 2. Construir series QoQ por cuenta
    result['account_time_series'] = _build_qoq_series(
        current_structure, prior_structure
    )

    # 3. Construir resumen QoQ por línea L1
    result['qoq_summary'] = _build_qoq_l1_summary(
        current_structure, prior_structure, result['account_time_series']
    )

    # 4. Detectar tendencias y alertas
    result['trend_alerts'] = _detect_trends(
        result['qoq_summary'], result['account_time_series']
    )

    return result


def _build_current_only_series(structure: Dict) -> List[AccountTimeSeries]:
    """Construye series para el trimestre actual sin histórico."""
    series = []
    for acct in structure.get('l3_accounts', []):
        series.append(AccountTimeSeries(
            account_code=acct.account_code,
            description=acct.name,
            parent_line=acct.parent_line,
            budget_q=acct.budget_current,
            actual_q=acct.actual_current,
            variance_q=acct.variance_current,
            variance_pct_q=acct.variance_pct_current,
            partner_comment=acct.partner_comment,
        ))
    return series


def _detect_account_changes_structured(
    current: Dict, prior: Dict
) -> AccountChangeReport:
    """Detecta cuentas nuevas y eliminadas con impacto financiero."""
    curr_map = {a.account_code: a for a in current.get('l3_accounts', []) if a.account_code}
    prev_map = {a.account_code: a for a in prior.get('l3_accounts', []) if a.account_code}

    new_accounts = []
    total_new_impact = 0.0
    for code, acct in curr_map.items():
        if code not in prev_map:
            impact = acct.actual_current
            new_accounts.append({
                'account_code': code,
                'description': acct.name,
                'parent_line': acct.parent_line,
                'actual_current': acct.actual_current,
                'budget_current': acct.budget_current,
                'variance': acct.variance_current,
                'impact': impact,
            })
            total_new_impact += impact

    removed_accounts = []
    total_removed_impact = 0.0
    for code, acct in prev_map.items():
        if code not in curr_map:
            removed_accounts.append({
                'account_code': code,
                'description': acct.name,
                'parent_line': acct.parent_line,
                'last_actual': acct.actual_current,
                'last_budget': acct.budget_current,
            })
            total_removed_impact += acct.actual_current

    # Ordenar por impacto absoluto
    new_accounts.sort(key=lambda x: abs(x['impact']), reverse=True)
    removed_accounts.sort(key=lambda x: abs(x.get('last_actual', 0)), reverse=True)

    return AccountChangeReport(
        new_accounts=new_accounts,
        removed_accounts=removed_accounts,
        total_new_impact=total_new_impact,
        total_removed_impact=total_removed_impact,
    )


def _build_qoq_series(
    current: Dict, prior: Dict
) -> List[AccountTimeSeries]:
    """Construye series QoQ por cuenta."""
    curr_map = {a.account_code: a for a in current.get('l3_accounts', []) if a.account_code}
    prev_map = {a.account_code: a for a in prior.get('l3_accounts', []) if a.account_code}

    all_codes = set(list(curr_map.keys()) + list(prev_map.keys()))
    series = []

    for code in all_codes:
        curr = curr_map.get(code)
        prev = prev_map.get(code)

        if curr and prev:
            # Cuenta existe en ambos trimestres
            qoq_actual = curr.actual_current - prev.actual_current
            qoq_actual_pct = (
                qoq_actual / abs(prev.actual_current)
                if prev.actual_current and abs(prev.actual_current) > 0.01
                else None
            )
            qoq_var_change = curr.variance_current - prev.variance_current

            # Detectar tendencia
            trend = _classify_trend(
                curr.variance_current, prev.variance_current,
                curr.parent_line
            )

            series.append(AccountTimeSeries(
                account_code=code,
                description=curr.name,
                parent_line=curr.parent_line,
                budget_q=curr.budget_current,
                actual_q=curr.actual_current,
                variance_q=curr.variance_current,
                variance_pct_q=curr.variance_pct_current,
                budget_prior_q=prev.budget_current,
                actual_prior_q=prev.actual_current,
                variance_prior_q=prev.variance_current,
                variance_pct_prior_q=prev.variance_pct_current,
                qoq_actual_change=qoq_actual,
                qoq_actual_change_pct=qoq_actual_pct,
                qoq_variance_change=qoq_var_change,
                trend_flag=trend,
                partner_comment=curr.partner_comment,
            ))
        elif curr and not prev:
            # Cuenta nueva
            series.append(AccountTimeSeries(
                account_code=code,
                description=curr.name,
                parent_line=curr.parent_line,
                budget_q=curr.budget_current,
                actual_q=curr.actual_current,
                variance_q=curr.variance_current,
                variance_pct_q=curr.variance_pct_current,
                is_new_account=True,
                partner_comment=curr.partner_comment,
            ))
        elif prev and not curr:
            # Cuenta eliminada
            series.append(AccountTimeSeries(
                account_code=code,
                description=prev.name,
                parent_line=prev.parent_line,
                budget_q=0,
                actual_q=0,
                variance_q=0,
                variance_pct_q=None,
                budget_prior_q=prev.budget_current,
                actual_prior_q=prev.actual_current,
                variance_prior_q=prev.variance_current,
                variance_pct_prior_q=prev.variance_pct_current,
                is_removed_account=True,
            ))

    # Ordenar por magnitud de cambio QoQ en variance
    series.sort(
        key=lambda x: abs(x.qoq_variance_change or 0),
        reverse=True
    )
    return series


def _build_qoq_l1_summary(
    current: Dict,
    prior: Dict,
    account_series: List[AccountTimeSeries],
) -> List[QoQSummary]:
    """Construye resumen QoQ por línea L1."""
    curr_l1 = {l.name: l for l in current.get('l1_lines', [])}
    prev_l1 = {l.name: l for l in prior.get('l1_lines', [])}

    order = ['Income', 'Operating Expenses', 'Real Estate Taxes', 'NOI',
             'Interest Expense', 'Non-Operating Expenses', 'Net Income']

    summaries = []
    for name in order:
        c = curr_l1.get(name)
        p = prev_l1.get(name)
        if not c:
            continue

        qoq_actual = None
        qoq_actual_pct = None
        qoq_var_change = None
        trend = ""

        if p:
            qoq_actual = c.actual_current - p.actual_current
            qoq_actual_pct = (
                qoq_actual / abs(p.actual_current)
                if p.actual_current and abs(p.actual_current) > 0.01
                else None
            )
            qoq_var_change = c.variance_current - p.variance_current
            trend = _classify_l1_trend(c.variance_current, p.variance_current, name)

        # Top movers: cuentas con mayor cambio QoQ dentro de esta línea
        line_accounts = [s for s in account_series if s.parent_line == name]
        top = sorted(line_accounts, key=lambda x: abs(x.qoq_variance_change or 0), reverse=True)
        top_movers = []
        for t in top[:5]:
            if t.qoq_variance_change and abs(t.qoq_variance_change) > 100:
                top_movers.append({
                    'account_code': t.account_code,
                    'description': t.description,
                    'qoq_variance_change': t.qoq_variance_change,
                    'variance_current': t.variance_q,
                    'variance_prior': t.variance_prior_q,
                    'trend': t.trend_flag,
                })

        summaries.append(QoQSummary(
            line_name=name,
            actual_current_q=c.actual_current,
            actual_prior_q=p.actual_current if p else None,
            budget_current_q=c.budget_current,
            budget_prior_q=p.budget_current if p else None,
            variance_current_q=c.variance_current,
            variance_prior_q=p.variance_current if p else None,
            qoq_actual_change=qoq_actual,
            qoq_actual_change_pct=qoq_actual_pct,
            qoq_variance_change=qoq_var_change,
            trend_direction=trend,
            top_movers=top_movers,
        ))

    return summaries


def _detect_trends(
    qoq_summary: List[QoQSummary],
    account_series: List[AccountTimeSeries],
) -> List[str]:
    """Genera alertas de tendencia basadas en el análisis QoQ."""
    alerts = []

    # Alertas a nivel L1
    for s in qoq_summary:
        if s.trend_direction == "deteriorating":
            alerts.append(
                f"⚠️ {s.line_name}: varianza QoQ deteriorándose "
                f"(Q anterior: ${s.variance_prior_q:,.0f} → Q actual: ${s.variance_current_q:,.0f})"
            )
        if s.qoq_actual_change_pct and abs(s.qoq_actual_change_pct) > 0.20:
            direction = "aumentó" if s.qoq_actual_change_pct > 0 else "disminuyó"
            alerts.append(
                f"📊 {s.line_name}: actual {direction} "
                f"{abs(s.qoq_actual_change_pct)*100:.0f}% QoQ "
                f"(${s.qoq_actual_change:+,.0f})"
            )

    # Alertas a nivel cuenta: aceleración de desviaciones
    for acct in account_series[:20]:  # Top 20 por cambio
        if acct.trend_flag == "accelerating" and abs(acct.qoq_variance_change or 0) > 2000:
            alerts.append(
                f"🔺 {acct.description} ({acct.account_code}): "
                f"desviación acelerándose — varianza QoQ cambió ${acct.qoq_variance_change:+,.0f}"
            )
        elif acct.trend_flag == "reversal" and abs(acct.qoq_variance_change or 0) > 2000:
            alerts.append(
                f"🔄 {acct.description} ({acct.account_code}): "
                f"cambio de dirección — varianza pasó de ${acct.variance_prior_q:+,.0f} "
                f"a ${acct.variance_q:+,.0f}"
            )

    return alerts


def _classify_trend(
    variance_current: float,
    variance_prior: float,
    parent_line: str,
) -> str:
    """
    Clasifica la tendencia de una cuenta:
    - improving: varianza se acerca a 0 o se mueve favorablemente
    - deteriorating: varianza se aleja de 0 desfavorablemente
    - accelerating: deteriorating y la magnitud del cambio es grande
    - reversal: cambió de signo
    """
    if abs(variance_prior) < 0.01 and abs(variance_current) < 0.01:
        return ""

    # Cambio de signo = reversal
    if variance_prior * variance_current < 0:
        return "reversal"

    # Para income: varianza positiva = favorable
    # Para expenses: varianza negativa = favorable (gastó menos)
    is_expense = parent_line in ('Operating Expenses', 'Real Estate Taxes', 'Non-Operating Expenses', 'Interest Expense')

    if is_expense:
        # Favorable = varianza negativa (menor gasto)
        if variance_current < variance_prior:
            return "improving"
        elif abs(variance_current) > abs(variance_prior) * 1.5:
            return "accelerating"
        else:
            return "deteriorating"
    else:
        # Income/NOI: favorable = varianza positiva
        if variance_current > variance_prior:
            return "improving"
        elif abs(variance_current) > abs(variance_prior) * 1.5:
            return "accelerating"
        else:
            return "deteriorating"


def _classify_l1_trend(
    variance_current: float,
    variance_prior: float,
    line_name: str,
) -> str:
    """Clasifica tendencia a nivel L1 (simplificado)."""
    if abs(variance_prior) < 0.01 and abs(variance_current) < 0.01:
        return "stable"

    is_expense = line_name in ('Operating Expenses', 'Real Estate Taxes', 'Non-Operating Expenses')

    if is_expense:
        # Para gastos: menos gasto (varianza más negativa) = improving
        if abs(variance_current) < abs(variance_prior) * 0.9:
            return "improving"
        elif abs(variance_current) > abs(variance_prior) * 1.1:
            return "deteriorating"
        return "stable"
    else:
        # Para income/NOI: varianza actual > prior = improving
        if variance_current > variance_prior * 1.1:
            return "improving"
        elif variance_current < variance_prior * 0.9:
            return "deteriorating"
        return "stable"


def format_qoq_for_payload(analysis: Dict) -> Dict:
    """Formatea el análisis histórico para incluir en el JSON payload del Copilot."""
    if not analysis.get('has_prior'):
        return {'has_prior_quarter': False}

    payload = {
        'has_prior_quarter': True,
        'qoq_l1_summary': [],
        'account_changes': {
            'new_accounts': analysis['account_changes'].new_accounts,
            'removed_accounts': analysis['account_changes'].removed_accounts,
            'total_new_impact': analysis['account_changes'].total_new_impact,
            'total_removed_impact': analysis['account_changes'].total_removed_impact,
        },
        'trend_alerts': analysis['trend_alerts'],
        'top_qoq_movers': [],
    }

    for s in analysis.get('qoq_summary', []):
        payload['qoq_l1_summary'].append({
            'line': s.line_name,
            'actual_current_q': s.actual_current_q,
            'actual_prior_q': s.actual_prior_q,
            'variance_current_q': s.variance_current_q,
            'variance_prior_q': s.variance_prior_q,
            'qoq_actual_change': s.qoq_actual_change,
            'qoq_actual_change_pct': s.qoq_actual_change_pct,
            'trend': s.trend_direction,
            'top_movers': s.top_movers[:3],
        })

    # Top 10 cuentas con mayor cambio QoQ
    for acct in analysis.get('account_time_series', [])[:10]:
        if acct.qoq_variance_change and abs(acct.qoq_variance_change) > 100:
            payload['top_qoq_movers'].append({
                'account_code': acct.account_code,
                'description': acct.description,
                'parent_line': acct.parent_line,
                'variance_current': acct.variance_q,
                'variance_prior': acct.variance_prior_q,
                'qoq_change': acct.qoq_variance_change,
                'trend': acct.trend_flag,
                'is_new': acct.is_new_account,
                'is_removed': acct.is_removed_account,
                'partner_comment': acct.partner_comment,
            })

    return payload
