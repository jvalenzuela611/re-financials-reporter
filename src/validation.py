"""
Módulo de Validación: cuadratura de totales, detección de cuentas nuevas/eliminadas,
y otras validaciones de integridad.
"""

import pandas as pd
from typing import Dict, List, Optional
from dataclasses import dataclass


@dataclass
class ValidationResult:
    """Resultado de una validación individual."""
    check_name: str
    status: str  # "OK", "WARNING", "BLOCK"
    message: str
    details: Optional[Dict] = None


def validate_all(structure: Dict, prior_structure: Optional[Dict] = None) -> List[ValidationResult]:
    """
    Ejecuta todas las validaciones sobre la estructura financiera.

    Args:
        structure: output de build_financial_structure() para trimestre actual
        prior_structure: output de build_financial_structure() para trimestre anterior (opcional)

    Returns:
        Lista de ValidationResult
    """
    results = []

    # 1. Verificar que hay datos
    results.append(_check_data_presence(structure))

    # 2. Verificar líneas principales
    results.append(_check_l1_completeness(structure))

    # 3. Cuadratura: Total OpEx vs suma de sub-totales L2
    results.append(_check_opex_reconciliation(structure))

    # 4. Cuadratura: NOI = Income - Total OpEx - Taxes
    results.append(_check_noi_reconciliation(structure))

    # 5. Detección de cuentas nuevas/eliminadas (si hay histórico)
    if prior_structure:
        new_accts, removed_accts = _detect_account_changes(structure, prior_structure)
        results.append(_report_account_changes(new_accts, removed_accts))

    # 6. Detectar outliers (variance > 200% o budget=0 con actual > 0)
    results.append(_check_outliers(structure))

    # 7. Verificar que se extrajeron comentarios
    results.append(_check_comments(structure))

    # 8. Full P&L reconciliation: L1 totales del socio vs suma de L3 clasificados
    results.append(_check_full_pl_reconciliation(structure))

    # 9. Ghost accounts: cuentas que aparecen en M1/M2 pero no en M3 (rescatadas en consolidación)
    ghost_check = _check_ghost_accounts(structure)
    if ghost_check:
        results.append(ghost_check)

    # 10. Cuentas con descripción similar pero clasificación L1 distinta
    dup_check = _check_duplicate_descriptions(structure)
    if dup_check:
        results.append(dup_check)

    return results


def _check_ghost_accounts(structure: Dict) -> Optional[ValidationResult]:
    """
    Detecta cuentas 'fantasma': existen en al menos un archivo mensual pero
    desaparecieron del último archivo del trimestre. La consolidación las rescata
    sumando M1+M2+0, pero alertamos al analista para que valide con el socio.
    """
    raw = structure.get('raw_consolidated')
    if raw is None or '_ghost_account' not in raw.columns:
        return None

    ghost_rows = raw[raw['_ghost_account'] == True]
    if ghost_rows.empty:
        return None

    items = []
    for _, row in ghost_rows.iterrows():
        desc = str(row.get('description', '')).strip()
        acct = str(row.get('account', '')).strip()
        actual = row.get('actual_current') or 0
        budget = row.get('budget_current') or 0
        files = row.get('_ghost_files_seen', '')
        items.append({
            'account_code': acct,
            'description': desc,
            'actual': actual,
            'budget': budget,
            'files_seen': files,
        })

    n = len(items)
    return ValidationResult(
        "Cuentas fantasma (desaparecidas)",
        "WARNING",
        f"{n} cuenta(s) presente(s) en mes(es) previo(s) pero ausente(s) del último archivo. "
        f"Se sumaron de todas formas — verificar con el socio si fueron eliminadas o es error de reporte.",
        {"items": items},
    )


def _check_duplicate_descriptions(structure: Dict) -> Optional[ValidationResult]:
    """
    Detecta cuentas cuya descripción (normalizada) es similar pero que tienen
    distinta clasificación L1. Caso típico: "AMORTIZATION EXPENSE" en dos cuentas
    con código distinto, una clasificada como Interest y otra como Non-Op.

    Bug fix The Val (feedback Florencia 23-04-2026).
    """
    import re

    def _norm(s: str) -> str:
        s = (s or '').lower()
        s = re.sub(r'\([^)]*\)', '', s)
        s = re.sub(r'[^a-z0-9 ]', ' ', s)
        s = re.sub(r'\s+', ' ', s).strip()
        return s

    accts = structure.get('l3_accounts', []) or []
    by_norm = {}
    for a in accts:
        if not a.name or a.parent_line in ('NA', 'Unknown'):
            continue
        key = _norm(a.name)
        if not key or len(key) < 5:
            continue
        by_norm.setdefault(key, []).append(a)

    conflicts = []
    for key, group in by_norm.items():
        if len(group) < 2:
            continue
        l1s = {a.parent_line for a in group}
        if len(l1s) <= 1:
            continue
        conflicts.append({
            'description_norm': key,
            'accounts': [
                {
                    'code': a.account_code,
                    'description': a.name,
                    'parent_line': a.parent_line,
                    'actual': a.actual_current,
                }
                for a in group
            ],
        })

    if not conflicts:
        return None

    n = len(conflicts)
    return ValidationResult(
        "Cuentas con descripción similar — clasificación distinta",
        "WARNING",
        f"Detectado(s) {n} grupo(s) de cuenta(s) con descripción casi idéntica "
        f"pero L1 distinto. Revisar si es duplicación o reclasificación necesaria.",
        {"items": conflicts},
    )


def _check_data_presence(structure: Dict) -> ValidationResult:
    """Verifica que hay datos en la estructura."""
    l1 = structure.get('l1_lines', [])
    if not l1:
        return ValidationResult("Datos cargados", "BLOCK",
                                "No se detectaron líneas financieras principales. Verificar formato del Excel.")
    return ValidationResult("Datos cargados", "OK",
                            f"{len(l1)} líneas principales detectadas, "
                            f"{len(structure.get('l3_accounts', []))} subcuentas con datos")


def _check_l1_completeness(structure: Dict) -> ValidationResult:
    """Verifica que las líneas principales estén presentes."""
    l1_names = {l.name for l in structure.get('l1_lines', [])}
    required = {'Income', 'Operating Expenses', 'NOI'}
    missing = required - l1_names

    if missing:
        return ValidationResult("Líneas principales", "WARNING",
                                f"Faltan líneas principales: {', '.join(missing)}. "
                                "El output puede estar incompleto.")

    return ValidationResult("Líneas principales", "OK",
                            f"Líneas detectadas: {', '.join(sorted(l1_names))}")


def _detect_ret_inside_opex(structure: Dict) -> Optional[bool]:
    """
    Detecta si Real Estate Taxes está incluido dentro del total de Operating Expenses.
    Retorna True (Oakland/legacy), False (Val/CMC), o None si la detección es ambigua.

    Ambiguo = ambas fórmulas explican el total del socio con diferencia < threshold.
    """
    l1_map = {l.name: l for l in structure.get('l1_lines', [])}
    opex = l1_map.get('Operating Expenses')
    taxes = l1_map.get('Real Estate Taxes')

    if not opex or not taxes or not (taxes.actual_current or 0):
        return False

    l3_accounts = structure.get('l3_accounts', [])
    opex_l3 = sum(a.actual_current or 0 for a in l3_accounts if a.parent_line == 'Operating Expenses')
    ret_l3 = sum(a.actual_current or 0 for a in l3_accounts if a.parent_line == 'Real Estate Taxes')

    opex_abs = abs(opex.actual_current)
    diff_with_ret = abs(opex_abs - abs(opex_l3 + ret_l3))
    diff_without_ret = abs(opex_abs - abs(opex_l3))

    # Si ambas fórmulas dan resultado similar, la detección es ambigua
    ambiguity_threshold = max(1000, opex_abs * 0.005)
    if abs(diff_without_ret - diff_with_ret) < ambiguity_threshold:
        return None

    return diff_with_ret < diff_without_ret


def _check_opex_reconciliation(structure: Dict) -> ValidationResult:
    """
    Verifica OpEx (ex-taxes, normalizado) ≈ suma de L3 clasificadas bajo Operating Expenses.
    La estructura ya viene con RET separada de OpEx (_normalize_ret_split en PASO 6).
    """
    l1_map = {l.name: l for l in structure.get('l1_lines', [])}
    opex_total = l1_map.get('Operating Expenses')

    if not opex_total:
        return ValidationResult("Cuadratura OpEx", "WARNING",
                                "Total Operating Expenses no encontrado.")

    # Suma de L3 clasificadas bajo OpEx (excluye RET por clasificación)
    l3_accounts = structure.get('l3_accounts', [])
    sum_l3_opex = sum(a.actual_current or 0 for a in l3_accounts
                      if a.parent_line == 'Operating Expenses')

    diff = abs(abs(opex_total.actual_current) - abs(sum_l3_opex))
    tolerance_warn = max(abs(opex_total.actual_current) * 0.02, 500)
    tolerance_block = max(abs(opex_total.actual_current) * 0.10, 5000)

    normalized_note = " (OpEx normalizado ex-taxes)" if structure.get('ret_normalized') else ""

    if diff > tolerance_block:
        return ValidationResult("Cuadratura OpEx", "BLOCK",
                                f"Cuadratura CRÍTICA fallida{normalized_note}: "
                                f"diferencia ${diff:,.0f} excede 10%. "
                                f"OpEx ex-taxes: ${opex_total.actual_current:,.0f}, "
                                f"Suma L3 OpEx: ${sum_l3_opex:,.0f}. "
                                "Verificar clasificación de cuentas.",
                                {"difference": diff})
    if diff > tolerance_warn:
        return ValidationResult("Cuadratura OpEx", "WARNING",
                                f"Diferencia en cuadratura{normalized_note}: ${diff:,.0f}. "
                                f"OpEx ex-taxes: ${opex_total.actual_current:,.0f}, "
                                f"Suma L3 OpEx: ${sum_l3_opex:,.0f}",
                                {"difference": diff})

    return ValidationResult("Cuadratura OpEx", "OK",
                            f"Cuadratura OK ex-taxes (diferencia: ${diff:,.0f})")


def _check_noi_reconciliation(structure: Dict) -> ValidationResult:
    """
    Verifica NOI = Income - |OpEx| - |RET|.
    - Legacy (Alice/Edson): OpEx del socio ya incluye RET, se resta solo OpEx.
    - Val/CMC: RET es línea L1 separada, se resta Income - |OpEx| - |RET|.
    """
    l1_map = {l.name: l for l in structure.get('l1_lines', [])}

    income = l1_map.get('Income')
    opex = l1_map.get('Operating Expenses')
    noi = l1_map.get('NOI')
    taxes = l1_map.get('Real Estate Taxes')

    if not all([income, opex, noi]):
        return ValidationResult("Cuadratura NOI", "WARNING",
                                "Faltan líneas para validar NOI.")

    # Estructura ya normalizada: OpEx excluye RET. NOI = Income − |OpEx| − |RET|
    ret_val = abs(taxes.actual_current) if (taxes and taxes.actual_current) else 0
    expected_noi = income.actual_current - abs(opex.actual_current) - ret_val
    diff = abs(noi.actual_current - expected_noi)
    tolerance_warn = max(abs(noi.actual_current) * 0.05, 100)
    tolerance_block = max(abs(noi.actual_current) * 0.15, 5000)

    if diff > tolerance_block:
        return ValidationResult("Cuadratura NOI", "BLOCK",
                                f"Cuadratura NOI CRÍTICA: NOI calculado (${expected_noi:,.0f}) "
                                f"difiere del reportado (${noi.actual_current:,.0f}) en ${diff:,.0f}. "
                                "Verificar archivos fuente antes de continuar.",
                                {"expected": expected_noi, "reported": noi.actual_current})
    if diff > tolerance_warn:
        return ValidationResult("Cuadratura NOI", "WARNING",
                                f"NOI calculado (${expected_noi:,.0f}) difiere del reportado "
                                f"(${noi.actual_current:,.0f}) en ${diff:,.0f}",
                                {"expected": expected_noi, "reported": noi.actual_current})

    return ValidationResult("Cuadratura NOI", "OK",
                            f"NOI cuadra (diferencia: ${diff:,.0f})")


def _detect_account_changes(current: Dict, prior: Dict) -> tuple:
    """Detecta cuentas nuevas y eliminadas entre períodos."""
    curr_codes = {a.account_code for a in current.get('l3_accounts', []) if a.account_code}
    prev_codes = {a.account_code for a in prior.get('l3_accounts', []) if a.account_code}

    new_accounts = []
    for acct in current.get('l3_accounts', []):
        if acct.account_code and acct.account_code not in prev_codes:
            new_accounts.append({
                'code': acct.account_code,
                'description': acct.name,
                'actual': acct.actual_current,
                'budget': acct.budget_current,
            })

    removed_accounts = []
    for acct in prior.get('l3_accounts', []):
        if acct.account_code and acct.account_code not in curr_codes:
            removed_accounts.append({
                'code': acct.account_code,
                'description': acct.name,
            })

    return new_accounts, removed_accounts


def _report_account_changes(new_accts: list, removed_accts: list) -> ValidationResult:
    """Genera reporte de cambios en cuentas."""
    if not new_accts and not removed_accts:
        return ValidationResult("Cambios en cuentas", "OK",
                                "Sin cuentas nuevas ni eliminadas respecto al período anterior.")

    msgs = []
    if new_accts:
        msgs.append(f"{len(new_accts)} cuenta(s) nueva(s)")
    if removed_accts:
        msgs.append(f"{len(removed_accts)} cuenta(s) eliminada(s)")

    return ValidationResult("Cambios en cuentas", "WARNING",
                            f"Detectado: {', '.join(msgs)}. Revisar en detalle.",
                            {"new": new_accts, "removed": removed_accts})


def _check_outliers(structure: Dict) -> ValidationResult:
    """Detecta outliers: variance extrema o budget=0. Incluye comentarios del socio."""
    outliers = []
    comment_map = {c['account_code']: c['comment_text'] for c in structure.get('all_comments', []) if c.get('account_code')}

    for acct in structure.get('l3_accounts', []):
        reason = ""
        # Budget = 0 con actual significativo
        if abs(acct.budget_current) < 0.01 and abs(acct.actual_current) > 1000:
            reason = f"Budget $0 vs Actual ${acct.actual_current:,.0f}"

        # Variance > 300%
        elif acct.variance_pct_current and abs(acct.variance_pct_current) > 3.0:
            reason = (
                f"Variance {acct.variance_pct_current*100:,.0f}% "
                f"(${acct.variance_current:,.0f})"
            )

        if reason:
            partner_comment = acct.partner_comment or comment_map.get(acct.account_code, "")
            outliers.append({
                'account_code': acct.account_code,
                'description': acct.name,
                'parent_line': acct.parent_line,
                'budget': acct.budget_current,
                'actual': acct.actual_current,
                'variance': acct.variance_current,
                'reason': reason,
                'partner_comment': partner_comment,
            })

    if not outliers:
        return ValidationResult("Outliers", "OK", "Sin outliers extremos detectados.")

    # Build display items (backwards compatible for items key)
    display_items = []
    for o in outliers:
        line = f"{o['description']}: {o['reason']}"
        if o['partner_comment']:
            line += f" | Socio: \"{o['partner_comment']}\""
        display_items.append(line)

    return ValidationResult("Outliers", "WARNING",
                            f"{len(outliers)} outlier(s) detectado(s)",
                            {"items": display_items, "outlier_details": outliers})


def _check_full_pl_reconciliation(structure: Dict) -> ValidationResult:
    """
    Full P&L reconciliation: compare each L1 partner-reported total against
    the sum of L3 accounts classified under that L1.

    Returns details with per-L1 breakdown, total, and unclassified accounts.
    """
    l1_lines = structure.get('l1_lines', [])
    l3_accounts = structure.get('l3_accounts', [])

    if not l1_lines or not l3_accounts:
        return ValidationResult(
            "Reconciliación P&L", "WARNING",
            "Datos insuficientes para reconciliación completa.",
        )

    # Calculated lines that should not be compared directly
    calculated = {'NOI', 'Net Income'}

    # Build L3 sums by parent_line
    l3_by_parent = {}
    l3_count_by_parent = {}
    for acct in l3_accounts:
        parent = acct.parent_line or 'Unknown'
        l3_by_parent[parent] = l3_by_parent.get(parent, 0) + (acct.actual_current or 0)
        l3_count_by_parent[parent] = l3_count_by_parent.get(parent, 0) + 1

    per_l1 = []
    worst_status = "OK"
    total_partner = 0
    total_classification = 0
    total_count = 0

    # Estructura ya normalizada: OpEx L1 es ex-taxes, RET es línea independiente.
    # Cada L1 se compara con la suma de sus L3 clasificadas.
    for l1 in l1_lines:
        if l1.name in calculated:
            per_l1.append({
                'line': l1.name,
                'partner_total': l1.actual_current,
                'classification_sum': None,
                'diff': None,
                'account_count': None,
                'status': 'calc',
            })
            continue

        partner_val = l1.actual_current or 0
        class_sum = l3_by_parent.get(l1.name, 0)
        count = l3_count_by_parent.get(l1.name, 0)

        diff = partner_val - class_sum
        abs_diff = abs(diff)

        if abs_diff > 100:
            row_status = "WARNING"
        else:
            row_status = "OK"

        skip_from_total = False
        if not skip_from_total:
            if row_status == "WARNING" and worst_status == "OK":
                worst_status = "WARNING"

            total_partner += partner_val
            total_classification += class_sum
            total_count += count

        per_l1.append({
            'line': l1.name,
            'partner_total': partner_val,
            'classification_sum': class_sum,
            'diff': diff,
            'account_count': count,
            'status': row_status,
        })

    # Unclassified accounts
    unknown_sum = l3_by_parent.get('Unknown', 0) + l3_by_parent.get('unknown', 0) + l3_by_parent.get('', 0)
    unknown_count = l3_count_by_parent.get('Unknown', 0) + l3_count_by_parent.get('unknown', 0) + l3_count_by_parent.get('', 0)

    if unknown_count > 0 and worst_status == "OK":
        worst_status = "WARNING"

    details = {
        'per_l1': per_l1,
        'total_partner': total_partner,
        'total_classification': total_classification,
        'total_diff': total_partner - total_classification,
        'unclassified_count': unknown_count,
        'unclassified_sum': unknown_sum,
    }

    if worst_status == "WARNING":
        parts = []
        if any(r['status'] == 'WARNING' for r in per_l1):
            parts.append("diferencias menores en algunas líneas")
        if unknown_count:
            parts.append(f"{unknown_count} cuenta(s) sin clasificar (${unknown_sum:,.0f})")
        msg = f"Reconciliación P&L con alertas: {'; '.join(parts)}."
    else:
        msg = f"Reconciliación P&L OK. {total_count} cuentas cuadran con totales del socio."

    return ValidationResult("Reconciliación P&L", worst_status, msg, details)


def _check_comments(structure: Dict) -> ValidationResult:
    """Verifica que se extrajeron comentarios del socio."""
    comments = structure.get('all_comments', [])
    if not comments:
        return ValidationResult("Comentarios del socio", "WARNING",
                                "No se encontraron comentarios del socio (Variance Notes). "
                                "El output se generará sin ellos.")

    return ValidationResult("Comentarios del socio", "OK",
                            f"{len(comments)} comentario(s) extraído(s)")
