"""
Módulo de Output: genera la tabla resumen, bullets por línea financiera,
alertas y markdown final listo para pegar en el reporte.
"""

import pandas as pd
from typing import Dict, List, Optional
from io import BytesIO
import json

from src.config import l1_pl_sort_key


def _sort_accounts_pl_order(accounts):
    """
    Ordena una lista de L3 accounts en el orden canónico del P&L:
    (L1 order, account_code ASC). Útil para layouts legibles en Excel.
    """
    return sorted(
        accounts,
        key=lambda a: (l1_pl_sort_key(getattr(a, 'parent_line', '') or ''),
                       str(getattr(a, 'account_code', '') or '')),
    )


def _sort_accounts_partner_order(accounts):
    """
    Ordena L3 accounts siguiendo el flujo natural del P&L del partner:
      1) L1 en orden canónico (Income → OpEx → RET → NOI → Interest → Non-Op
         → Net Income → Total Capex → NA)
      2) Dentro de cada L1, orden de aparición de L2 (section_name) en el
         Excel del partner — se infiere del menor partner_row_order de cada
         L2 dentro del L1.
      3) Dentro de cada L2, por partner_row_order ASC (fila del Excel).

    Respeta la secuencia top-down que escribe el socio, pero re-agrupa las
    cuentas que el catálogo o el analista movieron a otro L1 (ej: un 6543-xxx
    del Val que estaba bajo "Taxes" y que el catálogo reclasifica como Total
    Capex — aparece al final del bloque Capex, no mezclado con los taxes).
    """
    # Para cada (L1, L2) calcular el partner_row_order mínimo — eso define
    # el orden de aparición del L2 dentro del L1.
    l2_first_row = {}  # (L1, L2) -> min row
    for a in accounts:
        key = (getattr(a, 'parent_line', '') or '',
               getattr(a, 'section_name', '') or '')
        row = getattr(a, 'partner_row_order', 0) or 0
        if key not in l2_first_row or row < l2_first_row[key]:
            l2_first_row[key] = row

    def key_fn(a):
        l1 = getattr(a, 'parent_line', '') or ''
        l2 = getattr(a, 'section_name', '') or ''
        return (
            l1_pl_sort_key(l1),                # 1) L1 canónico
            l2_first_row.get((l1, l2), 10**9), # 2) L2 orden de aparición
            getattr(a, 'partner_row_order', 0) or 0,  # 3) row dentro del L2
            str(getattr(a, 'account_code', '') or ''),  # tie-break
        )
    return sorted(accounts, key=key_fn)


# Convención de signo para el reporte: -1 = expense (mostrar negativo), 1 = revenue/income
DISPLAY_SIGN = {
    'Income': 1,
    'Operating Expenses': -1,
    'Real Estate Taxes': -1,
    'NOI': 1,
    'Interest Expense': -1,   # Línea L1 separada de OPEX y NOI
    'Interest': -1,           # Alias legacy
    'Non-Operating Expenses': -1,
    'Net Income': 1,
    'Total Capex': -1,        # Egreso de capital — se muestra negativo como OpEx/Interest
}


def _get_display_sign(line_name: str) -> int:
    """Retorna el signo de display para una línea del reporte."""
    return DISPLAY_SIGN.get(line_name, 1)


def build_summary_table(structure: Dict) -> pd.DataFrame:
    """
    Construye la tabla resumen estilo reporte del inversionista.
    Columnas: Q-Actual, Q-Budget, Variance, YTD-Actual, YTD-Budget, Variance.

    Garantías:
      - Orden canónico P&L (Income → OpEx → RET → NOI → Interest → Non-Op →
        Net Income → Total Capex), independiente del orden que trae la estructura.
      - "Total Capex" SIEMPRE aparece como última fila, aun si no hay actividad
        de capex en el período (se muestra con ceros).
    """
    # Detectar si cualquier línea tiene YTD para replicar el esquema de columnas
    lines = structure.get('l1_lines', [])
    has_ytd = any(l.actual_ytd is not None for l in lines)

    def _row_for_line(line):
        sign = _get_display_sign(line.name)
        row = {
            'Line': line.name,
            'Q Actual': line.actual_current * sign,
            'Q Budget': line.budget_current * sign,
            'Q Variance': _compute_display_variance(line.actual_current, line.budget_current, sign),
        }
        if has_ytd:
            row['YTD Actual'] = (line.actual_ytd or 0) * sign if line.actual_ytd is not None else 0
            row['YTD Budget'] = (line.budget_ytd or 0) * sign if line.budget_ytd is not None else 0
            row['YTD Variance'] = _compute_display_variance(
                line.actual_ytd or 0, line.budget_ytd or 0, sign
            ) if line.actual_ytd is not None else 0
        return row

    def _zero_row(name: str):
        sign = _get_display_sign(name)
        row = {
            'Line': name,
            'Q Actual': 0,
            'Q Budget': 0,
            'Q Variance': 0,
        }
        if has_ytd:
            row['YTD Actual'] = 0
            row['YTD Budget'] = 0
            row['YTD Variance'] = 0
        return row

    by_name = {l.name: l for l in lines}
    rows = []

    # 1) Recorrer L1s en orden P&L canónico, usando datos reales si existen
    canonical_order = [
        'Income', 'Operating Expenses', 'Real Estate Taxes', 'NOI',
        'Interest Expense', 'Non-Operating Expenses', 'Net Income',
        'Total Capex',
    ]
    for name in canonical_order:
        if name in by_name:
            rows.append(_row_for_line(by_name[name]))
        elif name == 'Total Capex':
            # Garantía: Capex SIEMPRE aparece como última fila, aun sin actividad
            rows.append(_zero_row('Total Capex'))

    # 2) Cualquier L1 fuera del orden canónico (edge case) va al final
    for line in lines:
        if line.name not in canonical_order:
            rows.append(_row_for_line(line))

    return pd.DataFrame(rows)


def build_bullets_by_line(structure: Dict, top_n: int = 3) -> str:
    """
    Genera bullets de comentario por cada línea financiera principal.
    Incluye variance, top drivers de L2/L3, y comentarios del socio.
    """
    lines = []
    l1_list = structure.get('l1_lines', [])
    l2_list = structure.get('l2_sections', [])
    l3_list = structure.get('l3_accounts', [])

    for l1 in l1_list:
        # Header de línea - usar signo display para dirección correcta
        sign = _get_display_sign(l1.name)
        display_var = _compute_display_variance(l1.actual_current, l1.budget_current, sign)
        var_str = _fmt_currency(abs(display_var))
        pct_str = _fmt_pct(l1.variance_pct_current)
        direction = "above" if display_var > 0 else "below"

        lines.append(f"**{l1.name}:** {direction} budget by {var_str} ({pct_str}).")

        # Sub-drivers L2 relevantes
        l2_relevant = [s for s in l2_list if s.parent_line == l1.name]
        l2_relevant.sort(key=lambda x: abs(x.variance_current), reverse=True)

        for sub in l2_relevant[:top_n]:
            if abs(sub.variance_current) < 100:
                continue
            sub_var = _fmt_currency(sub.variance_current)
            sub_pct = _fmt_pct(sub.variance_pct_current)
            fav = "favorable" if _is_favorable(sub.variance_current, l1.name) else "unfavorable"
            lines.append(f"  - {sub.name}: {sub_var} ({sub_pct}), {fav}.")

        # Sub-drivers L3 con mayor impacto (que no estén cubiertos por L2)
        l3_relevant = [a for a in l3_list if a.parent_line == l1.name]
        l3_relevant.sort(key=lambda x: abs(x.variance_current), reverse=True)

        shown_l3 = 0
        for acct in l3_relevant:
            if shown_l3 >= top_n:
                break
            if abs(acct.variance_current) < 500:
                continue
            acct_var = _fmt_currency(acct.variance_current)
            acct_pct = _fmt_pct(acct.variance_pct_current)
            comment = f' Partner note: "{acct.partner_comment}"' if acct.partner_comment else ""
            lines.append(
                f"  - {acct.name} ({acct.account_code}): "
                f"Actual {_fmt_currency(acct.actual_current)} vs Budget {_fmt_currency(acct.budget_current)} "
                f"(Var {acct_var}, {acct_pct}).{comment}")
            shown_l3 += 1

        lines.append("")  # Línea en blanco entre secciones

    return "\n".join(lines)


def build_alerts_section(validations: list, structure: Dict) -> str:
    """Genera sección de alertas en markdown."""
    lines = ["## Alerts"]

    # Alertas de validación
    for v in validations:
        if v.status == "WARNING":
            lines.append(f"- **WARNING:** {v.message}")
            if v.details:
                if 'new' in v.details:
                    for a in v.details['new'][:5]:
                        lines.append(f"  - NEW: {a['code']} - {a['description']} "
                                     f"(Actual: ${a.get('actual', 0):,.0f})")
                if 'removed' in v.details:
                    for a in v.details['removed'][:5]:
                        lines.append(f"  - REMOVED: {a['code']} - {a['description']}")
                if 'items' in v.details:
                    for item in v.details['items'][:5]:
                        lines.append(f"  - {item}")
        elif v.status == "BLOCK":
            lines.append(f"- **BLOCKED:** {v.message}")

    if len(lines) == 1:
        lines.append("- No critical alerts.")

    return "\n".join(lines)


def build_comments_table(structure: Dict) -> pd.DataFrame:
    """
    Construye tabla de comentarios del socio.
    CAMBIO: ahora muestra la historia completa de comentarios concatenados
    por account_code (separados por ' — ') en lugar de solo el más reciente.
    """
    comments = structure.get('all_comments', [])
    if not comments:
        return pd.DataFrame(columns=['Account', 'Description', 'Comments (all)', 'Source Files'])

    rows = []
    for c in comments:
        rows.append({
            'Account': c.get('account_code', ''),
            'Description': c.get('description', ''),
            'Comments (all)': c.get('comment_text', ''),
            'Source Files': c.get('source_file', ''),
        })
    return pd.DataFrame(rows)


def generate_full_markdown(
    building: str,
    period: str,
    summary_table: pd.DataFrame,
    bullets: str,
    alerts: str,
    comments_df: pd.DataFrame,
    structure: Dict,
) -> str:
    """Genera el documento markdown completo listo para copiar al reporte."""
    lines = []

    # Header
    lines.append(f"# {building} - Financial Summary")
    lines.append(f"**Period:** {period}")
    lines.append("")

    # Tabla resumen
    lines.append("## Portfolio Level Financials")
    lines.append("")
    lines.append(_df_to_markdown_table(summary_table))
    lines.append("")

    # Bullets
    lines.append("## Notes to Financials")
    lines.append("")
    lines.append(bullets)
    lines.append("")

    # Alertas
    lines.append(alerts)
    lines.append("")

    # Comentarios del socio
    if not comments_df.empty:
        lines.append("## Partner Comments (extracted from Variance Notes)")
        lines.append("")
        lines.append(_df_to_markdown_table(comments_df))
        lines.append("")

    return "\n".join(lines)


def generate_final_classification_excel(
    structure: Dict,
    building: str = "",
    period: str = "",
) -> bytes:
    """
    Genera un Excel con la clasificación FINAL de cuentas (L3 → L1) que se usó
    para construir el prompt del AI. Pensado como artefacto de auditoría descargable
    en Step 2, después de aprobar la clasificación.

    Contenido:
      - Hoja "Classification": una fila por cuenta L3 con:
          account_code, description, section_name (L2), parent_line (L1 final),
          classification_source (catalog | section | code_range | description_keyword |
          user_override | cross_side_flip | post_processing | unknown),
          budget_current, actual_current, variance_current,
          budget_ytd, actual_ytd, variance_ytd, partner_comment.
      - Hoja "L1 Summary": suma por L1 (útil para cuadrar contra el markdown).
      - Hoja "Source Breakdown": conteo de cuentas por fuente de clasificación
        (ayuda a ver qué % vino del catálogo vs fallback vs override manual).
    """
    buffer = BytesIO()
    l3_rows = []
    # Orden canónico P&L (Income → OpEx → RET → NOI → Interest → Non-Op → NI → Capex → NA)
    for acct in _sort_accounts_pl_order(structure.get('l3_accounts', [])):
        l3_rows.append({
            'account_code': acct.account_code,
            'description': acct.name,
            'section_name (L2)': acct.section_name,
            'parent_line (L1 FINAL)': acct.parent_line,
            'classification_source': getattr(acct, 'classification_source', 'unknown'),
            'budget_current': acct.budget_current,
            'actual_current': acct.actual_current,
            'variance_current': acct.variance_current,
            'budget_ytd': acct.budget_ytd,
            'actual_ytd': acct.actual_ytd,
            'variance_ytd': acct.variance_ytd,
            'partner_comment': acct.partner_comment,
        })
    df_l3 = pd.DataFrame(l3_rows)

    # L1 summary: suma por L1 desde L3 (debe cuadrar con lo que ve el AI)
    l1_rows = []
    if not df_l3.empty:
        grouped = df_l3.groupby('parent_line (L1 FINAL)', dropna=False).agg(
            n_accounts=('account_code', 'count'),
            sum_budget_current=('budget_current', 'sum'),
            sum_actual_current=('actual_current', 'sum'),
            sum_variance_current=('variance_current', 'sum'),
            sum_budget_ytd=('budget_ytd', 'sum'),
            sum_actual_ytd=('actual_ytd', 'sum'),
        ).reset_index()
        l1_rows = grouped.to_dict('records')
    df_l1 = pd.DataFrame(l1_rows)

    # Source breakdown: cuántas cuentas por fuente
    src_rows = []
    if not df_l3.empty:
        src = df_l3['classification_source'].value_counts().reset_index()
        src.columns = ['classification_source', 'n_accounts']
        src['pct'] = (src['n_accounts'] / src['n_accounts'].sum() * 100).round(1)
        src_rows = src.to_dict('records')
    df_src = pd.DataFrame(src_rows)

    # Meta tab
    meta = pd.DataFrame([
        {'field': 'building', 'value': building},
        {'field': 'period', 'value': period},
        {'field': 'total_l3_accounts', 'value': len(df_l3)},
        {'field': 'note', 'value': 'This is the FINAL classification used to generate the AI prompt. '
                                     'Any override by the analyst in Step 2 is reflected here.'},
    ])

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        meta.to_excel(writer, sheet_name='Meta', index=False)
        if not df_l3.empty:
            df_l3.to_excel(writer, sheet_name='Classification', index=False)
        if not df_l1.empty:
            df_l1.to_excel(writer, sheet_name='L1 Summary', index=False)
        if not df_src.empty:
            df_src.to_excel(writer, sheet_name='Source Breakdown', index=False)

    return buffer.getvalue()


def generate_support_excel(
    summary_table: pd.DataFrame,
    structure: Dict,
    comments_df: pd.DataFrame,
    validations: list,
) -> bytes:
    """Genera Excel de soporte con múltiples hojas."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Hoja 1: Resumen
        summary_table.to_excel(writer, sheet_name='Summary', index=False)

        # Hoja 2: Detalle L3 (subcuentas) — orden canónico P&L
        l3_data = []
        for acct in _sort_accounts_pl_order(structure.get('l3_accounts', [])):
            l3_data.append({
                'Account': acct.account_code,
                'Description': acct.name,
                'Report Line': acct.parent_line,
                'Budget (Current)': acct.budget_current,
                'Actual (Current)': acct.actual_current,
                'Variance': acct.variance_current,
                'Variance %': acct.variance_pct_current,
                'Budget (YTD)': acct.budget_ytd,
                'Actual (YTD)': acct.actual_ytd,
                'Partner Comment': acct.partner_comment,
            })
        if l3_data:
            pd.DataFrame(l3_data).to_excel(writer, sheet_name='Detail', index=False)

        # Hoja 3: Comentarios
        if not comments_df.empty:
            comments_df.to_excel(writer, sheet_name='Partner Comments', index=False)

        # Hoja 4: Validaciones
        val_data = [{'Check': v.check_name, 'Status': v.status, 'Message': v.message}
                    for v in validations]
        if val_data:
            pd.DataFrame(val_data).to_excel(writer, sheet_name='Validations', index=False)

        # Hoja 5: L2 Sections
        l2_data = []
        for s in structure.get('l2_sections', []):
            l2_data.append({
                'Section': s.name,
                'Parent Line': s.parent_line,
                'Budget': s.budget_current,
                'Actual': s.actual_current,
                'Variance': s.variance_current,
                'Variance %': s.variance_pct_current,
            })
        if l2_data:
            pd.DataFrame(l2_data).to_excel(writer, sheet_name='Sections L2', index=False)

    return buffer.getvalue()


def generate_account_tracking_excel(
    structure: Dict,
    historical_analysis: Optional[Dict] = None,
    building: str = "",
    period: str = "",
) -> Optional[bytes]:
    """
    Genera Excel de tracking de cuentas entre trimestres.
    Incluye:
    - Hoja 1: Resumen de cambios (cuentas nuevas y eliminadas)
    - Hoja 2: Comparación completa cuenta por cuenta (Current Q vs Prior Q)
    - Hoja 3: Series QoQ por línea L1
    - Hoja 4: Top movers (cuentas con mayor variación QoQ)
    - Hoja 5: Inventario completo de cuentas actuales (baseline para futuros trimestres)

    Retorna None si no hay análisis histórico disponible pero siempre genera
    el inventario de cuentas actuales.
    """
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book

        # --- Formatos ---
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#2F5496', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        currency_fmt = workbook.add_format({'num_format': '$#,##0', 'border': 1})
        pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1})
        text_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'})
        new_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1, 'bold': True})
        removed_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1, 'bold': True})
        title_fmt = workbook.add_format({
            'bold': True, 'font_size': 14, 'font_color': '#2F5496',
        })
        subtitle_fmt = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_color': '#404040',
        })

        has_prior = historical_analysis and historical_analysis.get('has_prior', False)

        # ═══════════════════════════════════════════
        # HOJA 1: RESUMEN DE CAMBIOS
        # ═══════════════════════════════════════════
        if has_prior:
            changes = historical_analysis.get('account_changes')
            ws1 = workbook.add_worksheet('Cambios de Cuentas')
            writer.sheets['Cambios de Cuentas'] = ws1

            row = 0
            ws1.write(row, 0, f'Tracking de Cuentas — {building} — {period}', title_fmt)
            row += 1
            ws1.write(row, 0, 'Comparación con trimestre anterior', subtitle_fmt)
            row += 2

            # Resumen
            ws1.write(row, 0, 'Cuentas nuevas:', subtitle_fmt)
            ws1.write(row, 1, len(changes.new_accounts) if changes else 0)
            ws1.write(row, 2, 'Impacto total ($):', subtitle_fmt)
            ws1.write(row, 3, changes.total_new_impact if changes else 0, currency_fmt)
            row += 1
            ws1.write(row, 0, 'Cuentas eliminadas:', subtitle_fmt)
            ws1.write(row, 1, len(changes.removed_accounts) if changes else 0)
            ws1.write(row, 2, 'Último valor ($):', subtitle_fmt)
            ws1.write(row, 3, changes.total_removed_impact if changes else 0, currency_fmt)
            row += 2

            # Tabla de cuentas nuevas
            if changes and changes.new_accounts:
                ws1.write(row, 0, 'CUENTAS NUEVAS (no existían en trimestre anterior)', subtitle_fmt)
                row += 1
                new_headers = ['Código', 'Descripción', 'Línea Reporte', 'Actual Q', 'Budget Q', 'Varianza', 'Impacto $']
                for c, h in enumerate(new_headers):
                    ws1.write(row, c, h, header_fmt)
                row += 1
                for acct in changes.new_accounts:
                    ws1.write(row, 0, acct.get('account_code', ''), new_fmt)
                    ws1.write(row, 1, acct.get('description', ''), text_fmt)
                    ws1.write(row, 2, acct.get('parent_line', ''), text_fmt)
                    ws1.write(row, 3, acct.get('actual_current', 0), currency_fmt)
                    ws1.write(row, 4, acct.get('budget_current', 0), currency_fmt)
                    ws1.write(row, 5, acct.get('variance', 0), currency_fmt)
                    ws1.write(row, 6, acct.get('impact', 0), currency_fmt)
                    row += 1
                row += 1

            # Tabla de cuentas eliminadas
            if changes and changes.removed_accounts:
                ws1.write(row, 0, 'CUENTAS ELIMINADAS (existían en trimestre anterior, ya no aparecen)', subtitle_fmt)
                row += 1
                rem_headers = ['Código', 'Descripción', 'Línea Reporte', 'Último Actual', 'Último Budget']
                for c, h in enumerate(rem_headers):
                    ws1.write(row, c, h, header_fmt)
                row += 1
                for acct in changes.removed_accounts:
                    ws1.write(row, 0, acct.get('account_code', ''), removed_fmt)
                    ws1.write(row, 1, acct.get('description', ''), text_fmt)
                    ws1.write(row, 2, acct.get('parent_line', ''), text_fmt)
                    ws1.write(row, 3, acct.get('last_actual', 0), currency_fmt)
                    ws1.write(row, 4, acct.get('last_budget', 0), currency_fmt)
                    row += 1

            # Ajustar anchos
            ws1.set_column('A:A', 14)
            ws1.set_column('B:B', 35)
            ws1.set_column('C:C', 22)
            ws1.set_column('D:G', 16)

        # ═══════════════════════════════════════════
        # HOJA 2: COMPARACIÓN COMPLETA QoQ POR CUENTA
        # ═══════════════════════════════════════════
        if has_prior:
            series = historical_analysis.get('account_time_series', [])
            # Sort en orden canónico P&L: L1 ASC, account_code ASC
            series = sorted(
                series,
                key=lambda s: (l1_pl_sort_key(getattr(s, 'parent_line', '') or ''),
                               str(getattr(s, 'account_code', '') or '')),
            )
            if series:
                comp_data = []
                for s in series:
                    status = "Nueva" if s.is_new_account else ("Eliminada" if s.is_removed_account else "Activa")
                    comp_data.append({
                        'Código': s.account_code,
                        'Descripción': s.description,
                        'Línea Reporte': s.parent_line,
                        'Status': status,
                        'Actual Q': s.actual_q,
                        'Budget Q': s.budget_q,
                        'Varianza Q': s.variance_q,
                        'Var % Q': s.variance_pct_q,
                        'Actual Prior Q': s.actual_prior_q,
                        'Budget Prior Q': s.budget_prior_q,
                        'Varianza Prior Q': s.variance_prior_q,
                        'Cambio QoQ Actual': s.qoq_actual_change,
                        'Cambio QoQ Var': s.qoq_variance_change,
                        'Tendencia': s.trend_flag.title() if s.trend_flag else '',
                        'Comentario Socio': s.partner_comment,
                    })
                comp_df = pd.DataFrame(comp_data)
                comp_df.to_excel(writer, sheet_name='Comparación QoQ', index=False)

                # Formatear la hoja
                ws2 = writer.sheets['Comparación QoQ']
                ws2.set_column('A:A', 14)
                ws2.set_column('B:B', 35)
                ws2.set_column('C:C', 22)
                ws2.set_column('D:D', 10)
                ws2.set_column('E:M', 16)
                ws2.set_column('N:N', 14)
                ws2.set_column('O:O', 40)

        # ═══════════════════════════════════════════
        # HOJA 3: RESUMEN QoQ POR LÍNEA L1
        # ═══════════════════════════════════════════
        if has_prior:
            qoq_summary = historical_analysis.get('qoq_summary', [])
            if qoq_summary:
                qoq_data = []
                for s in qoq_summary:
                    sign = _get_display_sign(s.line_name)
                    qoq_data.append({
                        'Línea': s.line_name,
                        'Actual Q (display)': s.actual_current_q * sign,
                        'Actual Prior Q (display)': (s.actual_prior_q or 0) * sign,
                        'Cambio QoQ': (s.qoq_actual_change or 0) * sign,
                        'Cambio QoQ %': s.qoq_actual_change_pct,
                        'Varianza Q': s.variance_current_q * sign,
                        'Varianza Prior Q': (s.variance_prior_q or 0) * sign,
                        'Cambio Varianza': (s.qoq_variance_change or 0) * sign,
                        'Tendencia': s.trend_direction.title() if s.trend_direction else '',
                    })
                qoq_df = pd.DataFrame(qoq_data)
                qoq_df.to_excel(writer, sheet_name='Resumen QoQ L1', index=False)

                ws3 = writer.sheets['Resumen QoQ L1']
                ws3.set_column('A:A', 22)
                ws3.set_column('B:H', 18)
                ws3.set_column('I:I', 14)

        # ═══════════════════════════════════════════
        # HOJA 4: TOP MOVERS QoQ
        # ═══════════════════════════════════════════
        if has_prior:
            series = historical_analysis.get('account_time_series', [])
            top_movers = [s for s in series
                          if s.qoq_variance_change and abs(s.qoq_variance_change) > 100
                          and not s.is_new_account and not s.is_removed_account]
            top_movers.sort(key=lambda x: abs(x.qoq_variance_change or 0), reverse=True)

            if top_movers:
                mover_data = []
                for s in top_movers[:30]:
                    mover_data.append({
                        'Código': s.account_code,
                        'Descripción': s.description,
                        'Línea': s.parent_line,
                        'Var Q Actual': s.variance_q,
                        'Var Q Anterior': s.variance_prior_q,
                        'Cambio QoQ Var': s.qoq_variance_change,
                        'Actual Q': s.actual_q,
                        'Actual Prior Q': s.actual_prior_q,
                        'Cambio QoQ Actual': s.qoq_actual_change,
                        'Tendencia': s.trend_flag.title() if s.trend_flag else '',
                        'Comentario': s.partner_comment,
                    })
                mover_df = pd.DataFrame(mover_data)
                mover_df.to_excel(writer, sheet_name='Top Movers QoQ', index=False)

                ws4 = writer.sheets['Top Movers QoQ']
                ws4.set_column('A:A', 14)
                ws4.set_column('B:B', 35)
                ws4.set_column('C:C', 22)
                ws4.set_column('D:I', 16)
                ws4.set_column('J:J', 14)
                ws4.set_column('K:K', 40)

        # ═══════════════════════════════════════════
        # HOJA 5: INVENTARIO COMPLETO (siempre se genera)
        # ═══════════════════════════════════════════
        # Orden natural del P&L del partner:
        #   1) L1 canónico (Income → OpEx → RET → NOI → Interest → Non-Op →
        #      Net Income → Capex → NA)
        #   2) Dentro de cada L1, L2 (sección) en orden de aparición en el Excel
        #   3) Dentro de cada L2, por row del Excel
        inv_data = []
        for acct in _sort_accounts_partner_order(structure.get('l3_accounts', [])):
            # Var YTD + Var % YTD (solo si hay YTD data)
            var_ytd = None
            var_pct_ytd = None
            if acct.actual_ytd is not None:
                var_ytd = (acct.actual_ytd or 0) - (acct.budget_ytd or 0)
                if acct.budget_ytd and abs(acct.budget_ytd) > 0.01:
                    var_pct_ytd = var_ytd / abs(acct.budget_ytd)
            inv_data.append({
                'Código': acct.account_code,
                'Descripción': acct.name,
                'Línea Reporte (L1)': acct.parent_line,
                'Sección (L2)': acct.section_name,
                'Budget Q': acct.budget_current,
                'Actual Q': acct.actual_current,
                'Varianza Q': acct.variance_current,
                'Var % Q': acct.variance_pct_current,
                'Budget YTD': acct.budget_ytd,
                'Actual YTD': acct.actual_ytd,
                'Varianza YTD': var_ytd,
                'Var % YTD': var_pct_ytd,
                'Comentario Socio': acct.partner_comment,
            })
        if inv_data:
            inv_df = pd.DataFrame(inv_data)
            inv_df.to_excel(writer, sheet_name='Inventario Cuentas', index=False)

            ws5 = writer.sheets['Inventario Cuentas']
            ws5.set_column('A:A', 14)   # Código
            ws5.set_column('B:B', 35)   # Descripción
            ws5.set_column('C:C', 22)   # Línea Reporte L1
            ws5.set_column('D:D', 25)   # Sección (L2)
            ws5.set_column('E:L', 16)   # Budget Q ... Var % YTD
            ws5.set_column('M:M', 40)   # Comentario Socio

        # ═══════════════════════════════════════════
        # HOJA 6: ALERTAS DE TENDENCIA
        # ═══════════════════════════════════════════
        if has_prior:
            alerts = historical_analysis.get('trend_alerts', [])
            if alerts:
                alert_data = [{'#': i+1, 'Alerta': a} for i, a in enumerate(alerts)]
                alert_df = pd.DataFrame(alert_data)
                alert_df.to_excel(writer, sheet_name='Alertas Tendencia', index=False)

                ws6 = writer.sheets['Alertas Tendencia']
                ws6.set_column('A:A', 6)
                ws6.set_column('B:B', 100)

    return buffer.getvalue()


def generate_copilot_payload(
    building: str,
    period: str,
    structure: Dict,
    historical_analysis: Optional[Dict] = None,
) -> str:
    """Genera JSON payload limpio para alimentar un copiloto de redacción."""

    def line_to_dict(line, apply_sign=False):
        sign = _get_display_sign(line.name) if apply_sign else 1
        d = {
            'name': line.name,
            'budget_current': line.budget_current * sign,
            'actual_current': line.actual_current * sign,
            'variance_current': _compute_display_variance(line.actual_current, line.budget_current, sign),
            'variance_pct_current': line.variance_pct_current,
        }
        if line.actual_ytd is not None:
            d['budget_ytd'] = line.budget_ytd * sign
            d['actual_ytd'] = line.actual_ytd * sign
            d['variance_ytd'] = _compute_display_variance(line.actual_ytd, line.budget_ytd, sign)
        if line.partner_comment:
            d['partner_comment'] = line.partner_comment
        return d

    # --- Splitear "Operating Expenses" en OpEx-only + Total Operating Expenses ---
    # En la estructura interna, L1 "Operating Expenses" = Total OpEx (incl. RET).
    # Para el prompt, necesitamos las tres líneas separadas del reporte:
    #   Operating Expenses (excl. RET), Real Estate Taxes, Total Operating Expenses.
    l1_lines = structure.get('l1_lines', [])
    l1_map = {l.name: l for l in l1_lines}
    opex_total = l1_map.get('Operating Expenses')
    ret_line = l1_map.get('Real Estate Taxes')

    l1_for_payload = []
    opex_split_ok = False
    for l in l1_lines:
        if l.name == 'Operating Expenses' and ret_line:
            # OpEx-only (excluye RET)
            opex_only_b = l.budget_current - (ret_line.budget_current or 0)
            opex_only_a = l.actual_current - (ret_line.actual_current or 0)
            opex_only_var = opex_only_a - opex_only_b

            opex_only_b_ytd = None
            opex_only_a_ytd = None
            opex_only_var_ytd = None
            if l.actual_ytd is not None and ret_line.actual_ytd is not None:
                opex_only_b_ytd = (l.budget_ytd or 0) - (ret_line.budget_ytd or 0)
                opex_only_a_ytd = (l.actual_ytd or 0) - (ret_line.actual_ytd or 0)
                opex_only_var_ytd = opex_only_a_ytd - opex_only_b_ytd

            sign = _get_display_sign('Operating Expenses')
            opex_only_dict = {
                'name': 'Operating Expenses',
                'note': 'Excludes Real Estate Taxes. See Total Operating Expenses below.',
                'budget_current': opex_only_b * sign,
                'actual_current': opex_only_a * sign,
                'variance_current': opex_only_var * sign,
                'variance_pct_current': (opex_only_var / abs(opex_only_b)) if opex_only_b and abs(opex_only_b) > 0.01 else None,
            }
            if opex_only_a_ytd is not None:
                opex_only_dict['budget_ytd'] = opex_only_b_ytd * sign
                opex_only_dict['actual_ytd'] = opex_only_a_ytd * sign
                opex_only_dict['variance_ytd'] = opex_only_var_ytd * sign

            l1_for_payload.append(opex_only_dict)

            # Validación: Total OpEx ≈ OpEx-only + RET
            check_q = abs(opex_only_var + (ret_line.variance_current or 0) - l.variance_current)
            opex_split_ok = check_q < 2  # tolerancia $1 por redondeo

        elif l.name == 'Real Estate Taxes':
            # RET normal, luego insertar Total Operating Expenses
            l1_for_payload.append(line_to_dict(l, apply_sign=True))

            # Total Operating Expenses (= OpEx-only + RET)
            if opex_total:
                total_opex_dict = line_to_dict(opex_total, apply_sign=True)
                total_opex_dict['name'] = 'Total Operating Expenses'
                total_opex_dict['note'] = 'Operating Expenses + Real Estate Taxes. Do NOT use this variance for the Operating Expenses bullet.'
                l1_for_payload.append(total_opex_dict)
        else:
            l1_for_payload.append(line_to_dict(l, apply_sign=True))

    payload = {
        'building': building,
        'period': period,
        'currency': 'USD',
        'l1_summary': l1_for_payload,
        'opex_split_validation': {
            'passed': opex_split_ok,
            'message': 'Total OpEx variance = OpEx variance + RET variance' if opex_split_ok
                       else 'WARNING: OpEx split does not reconcile within $1 tolerance',
        },
        'l2_sections': [line_to_dict(s, apply_sign=True) for s in structure.get('l2_sections', [])],
        'top_drivers': [line_to_dict(a) for a in structure.get('l3_accounts', [])[:15]],
        'partner_comments': structure.get('all_comments', []),
    }

    # Agregar datos históricos QoQ si están disponibles
    if historical_analysis:
        from src.historical_analysis import format_qoq_for_payload
        payload['historical'] = format_qoq_for_payload(historical_analysis)

    return json.dumps(payload, indent=2, ensure_ascii=False)


def generate_copilot_prompt(
    building: str,
    period: str,
    structure: Dict,
    historical_analysis: Optional[Dict] = None,
) -> str:
    """
    Genera el archivo de prompt para el agente de Copilot en SharePoint.

    Diseñado para ser subido a la carpeta del activo en SharePoint, donde el agente
    tiene acceso a todos los archivos del folder (financials .md, master .xlsx,
    reportes de trimestres anteriores) y los usa como contexto histórico.

    El prompt NO embebe los datos numéricos — le indica al agente dónde leerlos
    dentro de la carpeta. El JSON payload se mantiene inline solo como referencia
    estructurada de respaldo.
    """
    import re as _re

    # Sanitize building name para nombres de archivo (igual que master_excel.py)
    def _safe(s):
        return _re.sub(r'[^\w\s-]', '', s).strip().replace(' ', '_')

    safe_building = _safe(building)
    json_payload = generate_copilot_payload(building, period, structure, historical_analysis)

    prompt = f"""You are a real estate institutional asset management analyst at STARS Investment.

Your task: write the "Notes to Financials" section of the quarterly investor report for
{building} — {period}.

════════════════════════════════════════════════════
STEP 1 — READ YOUR SOURCE FILES (in this SharePoint folder)
════════════════════════════════════════════════════

This prompt lives inside the SharePoint folder for {building}. Before writing anything,
read and internalize the following files from this folder:

PRIMARY DATA SOURCE (current quarter):
  • financials_{safe_building}_{period.replace(' ', '_')}.md
    → This is the authoritative financial report for {period}.
      It contains the summary table, line-by-line variance analysis, alerts,
      and partner/operator comments. ALL numbers you cite must come from here.

HISTORICAL CONTEXT (prior quarters — read ALL available):
  • Any prior financials_*.md files in this folder (e.g., financials_{safe_building}_Q3_2025.md,
    financials_{safe_building}_Q2_2025.md, etc.)
    → Read these to understand: how prior variances were explained, which issues
      were flagged as one-time vs. structural, and the established narrative tone.

MULTI-QUARTER TREND DATA (MANDATORY READ):
  • master_{safe_building}_{period.replace(' ', '_')}.xlsx
    → This Excel master file is the SINGLE SOURCE OF TRUTH for multi-quarter trends.

    FILE STRUCTURE — the master has these sheets:
      Sheet "Master": One row per GL account. Columns grouped in 3 bands:
        ACTUALS:  M1|M2|M3|Q1|M4|M5|M6|Q2|M7|M8|M9|Q3|M10|M11|M12|Q4|YTD
        BUDGET:   M1|M2|M3|Q1|M4|M5|M6|Q2|M7|M8|M9|Q3|M10|M11|M12|Q4|YTD
        VARIANCE: M1|M2|M3|Q1|M4|M5|M6|Q2|M7|M8|M9|Q3|M10|M11|M12|Q4|YTD
        (M = monthly, Q = quarterly sum, YTD = year-to-date)
        Fixed columns: account_code (index), description, parent_line (L1 category)
      Sheet "Partner Comments": Operator comments by account and quarter.
      Sheet "Management Commentary": Analyst-written context (recurring and one-off items).
      Sheet "Log": Update history.

    HOW TO USE THIS FILE:
    a) Read the "Master" sheet. Group rows by parent_line to get L1-level totals.
       Use monthly columns (M1–M12) to see month-by-month evolution within each quarter.
       Use quarterly columns (Q1–Q4) for quarter-over-quarter comparison.
       Use YTD columns for cumulative year-to-date performance.
    b) Identify trends: which L1 lines have been consistently above or below budget
       across multiple quarters — cite the trend (e.g., "continuing the favorable
       trend observed since Q2").
    c) Detect reversals: if a line was favorable for 2+ quarters and is now
       unfavorable, highlight the change explicitly.
    d) Quantify cumulative impact: if a variance persists for 3+ quarters,
       note the approximate total cumulative deviation from the YTD column.
    e) Distinguish one-time vs. structural: a variance in only one quarter is
       likely one-time. If it appears in 2+ quarters, treat it as structural.
    f) Read the "Management Commentary" sheet: it contains analyst-written context
       about recurring items (e.g., tax reassessment, contract cancellations) and
       one-off events. Incorporate these explanations faithfully — they are
       AUTHORITATIVE. Do not re-explain recurring items in full; reference them
       concisely (e.g., "as noted in prior quarters").
    g) Read the "Partner Comments" sheet: operator-level notes by account and quarter.
       Use as operational color to enrich your bullets.

    If the master file is not available in this folder, fall back to the prior
    financials_*.md files for historical context.

════════════════════════════════════════════════════
STEP 2 — ESTABLISH NARRATIVE CONTINUITY
════════════════════════════════════════════════════

After reading the master file and historical financials, answer these questions
internally before writing:

1. Which variances from prior quarters were flagged as "one-time" or "timing"?
   → Check the master: did the variance resolve (disappeared in subsequent Q),
     or is it recurring (present in 2+ quarters)? Frame your language accordingly.

2. Which structural drivers were introduced in prior quarters?
   (e.g., "valet parking contract cancellation", "staffing changes", "tax reassessment")
   → Do NOT re-introduce them as new. Reference them as continuing or resolved.

3. Are there any trends visible in the master file across 2+ quarters?
   → If yes, note the trend continuation (e.g., "as has been the case since Q2...").
   → If a trend reversed this quarter, highlight it as a change.

4. What was the tone and level of detail in prior Notes to Financials?
   → Match it exactly: sentence length, degree of specificity, use of property names.

CRITICAL HIERARCHY:
  • financials_{safe_building}_{period.replace(' ', '_')}.md = AUTHORITATIVE (numbers, variances)
  • Prior quarterly files = NARRATIVE CONTEXT ONLY (never override current data)
  • If any conflict exists between current and prior files, current data prevails always.

════════════════════════════════════════════════════
STEP 3 — WRITING INSTRUCTIONS
════════════════════════════════════════════════════

FORMAT:
• One bullet per L1 line with material variance (>$1,000 Q variance).
• Start each bullet: "• **[Line Name]:** ..."
• HARD LIMIT: each bullet must be ≤200 tokens (~150-180 words). If a bullet
  exceeds this, cut the least material detail. Never exceed 200 tokens.
• Written as a continuous paragraph (no sub-bullets).
• Cover lines in this order: Income → Operating Expenses → Real Estate Taxes →
  NOI (only if something material to add) → Non-Operating Expenses → Interest →
  Capex (only if activity exists). Skip immaterial lines.

EACH BULLET STRUCTURE:
  a) State the variance: direction (favorable/unfavorable) and magnitude in USD.
  b) Quantify the 2–3 largest drivers with approximate dollar amounts.
     DO NOT just name the driver — attach a number to it.
     Example: instead of "driven by lower rents and vacancy," write
     "driven by ~$130k in effective rent shortfall, ~$50k in vacancy losses,
     and ~$40k in bad debt."
  c) Add operational color: property-specific detail (Alice House, Edson House),
     contracts, events, or operator comments from the financials file.
  d) Forward-looking framing (if applicable): will this continue or resolve?

INFORMATIONAL DENSITY (critical — this is an LP-facing document):
• Each bullet MUST contain 2–3 quantified driver amounts in addition to the
  headline variance. Pull these from the "top_drivers" and "l2_sections" in the
  JSON payload. Pick the largest by absolute variance.
• Use approximate round numbers for sub-drivers (~$40k, ~$130k, ~$8k) — investors
  care about order of magnitude, not pennies. Only the headline L1 variance should
  be precise.
• DO NOT increase bullet length to accommodate numbers. Replace vague qualitative
  phrases ("higher than expected", "lower than budgeted") with the quantified
  equivalent ("~$50k above budget", "~$25k below budget"). Same sentence count,
  more information.
• DO NOT list every sub-driver. Select only the 2–3 most material ones that
  explain ≥70% of the headline variance. Omit immaterial items entirely.

ANTI-GRANULARITY RULE:
• NEVER mention individual accounts or sub-drivers with |variance| < $5,000.
  These are noise. If no single driver exceeds $5,000, group them as
  "various smaller items" or "net of several offsetting items."
• NEVER name more than 3 drivers per bullet. If a line has 5 material drivers,
  pick the top 3 and say "among other items" for the rest.
• NEVER describe operational micro-detail that an LP investor would not care
  about (e.g., "the plumbing vendor was changed in October"). Stay at the
  category level (e.g., "R&M came in ~$12k below budget due to deferred maintenance").
• If a variance is <$5,000 at the L1 headline level, skip the bullet entirely.

STYLE RULES:
• Tone: Institutional, objective, precise. No dramatic language or exclamations.
  Write as an asset manager addressing sophisticated institutional investors.
• Numbers: Always cite the headline L1 variance precisely (e.g., "$54,356").
  For sub-drivers within bullets, use approximate round figures (e.g., "~$40k",
  "~$130k") unless the exact figure is important for context.
• Sign convention:
  - Income: positive variance = favorable, negative = unfavorable
  - Expenses (OpEx, Taxes, Interest, Non-Op): negative = favorable (less spend),
    positive = unfavorable (overspend)
  - NOI / Net Income: positive = favorable


══════════════════════════════════════════════════════════════════════
FORBIDDEN NARRATIVE PATTERNS — DO NOT WRITE
══════════════════════════════════════════════════════════════════════

The lead analyst (Florencia) has flagged these specific patterns in rejected
prior AI outputs. They MUST NOT appear in yours. This list is authoritative
and overrides any style impression you may have from the reference examples
at the end of this prompt.

A) VOCABULARY SUBSTITUTIONS
   • NEVER use "underwritten" / "underwriting" — always say "budget".
     Forbidden: "leasing below underwritten market rents"
     Correct:   "leasing below budgeted market rents"

B) CAUSAL / CLASSIFICATION INFERENCE (forbidden unless a partner comment
   uses that exact word)
   Do NOT label a variance as any of these categories on your own:
     • "structural" / "recurring" / "one-time" / "timing-related"
     • "timing/accrual issue" / "characterized as [X]"
     • "indicates [cause]"
   Rejected examples:
     - "Variance characterized as structural (not timing-related)"
     - "Indicates Q4 timing/accrual issue"
     - "Indicates consistent underperformance vs budget"
     - "Not considered a structural increase"
   Correct alternative: state magnitude + direction and stop.

C) FORWARD-LOOKING LANGUAGE (forbidden unless a partner comment explicitly
   states the forecast in those terms)
   Do NOT write predictions, expectations, or conditions about future periods.
   Rejected examples:
     - "Expected to normalize in future periods"
     - "Expected to continue unless debt terms change"
     - "Will normalize" / "Trend is expected to reverse"
     - "Capex reflects completion of planned works, not new initiatives"

D) FILLER / COMMODITY STATEMENTS (forbidden always)
   These phrases add zero information. Omit entirely.
     - "No specific operational driver identified"
     - "Reflects a mix of recurring savings and one-time repairs"
     - "Variance reflects mix of [X] + [Y]" (as a generic characterization)
     - "Overall performance was in line with expectations"

E) CROSS-LINE OFFSET NARRATIVE (forbidden unless a partner comment explicitly
   mentions a dollar-specific reclassification between two lines)
   Each L1 bullet stands alone. Do NOT write phrases like:
     - "Helped partially offset the NOI shortfall"
     - "This offset OpEx performance"
     - "Cushioned the impact on Net Income"
     - "Contributed to the bottom line"
   The ONLY allowed exception: when a partner comment says something like
   "$39,613 was reclassified from X to Y in October" — then you may mention
   that specific dollar amount in BOTH affected bullets, paraphrased in
   institutional tone.


══════════════════════════════════════════════════════════════════════
LINE ISOLATION RULE
══════════════════════════════════════════════════════════════════════

Write each bullet as if the reader will only see that bullet. No "and this
offsets X", no "brought Total OpEx close to budget", no "helped NOI hold up".
The only cross-line reference allowed is the explicit-reclassification
exception above (Forbidden Pattern E).


══════════════════════════════════════════════════════════════════════
CONCISION BY MAGNITUDE
══════════════════════════════════════════════════════════════════════

Bullet length MUST scale with variance magnitude. Padding is a violation.

  • |L1 variance| < $5,000
      → ONE sentence. Format: "[Line] was $N favorable/unfavorable to budget."
      → No drivers, no forward look, no color. Stop there.

  • $5,000 ≤ |L1 variance| < $50,000
      → 2–3 sentences. ONE named driver with approximate amount IF available
        from a partner comment or from top_drivers in the JSON. No interpretation.

  • |L1 variance| ≥ $50,000
      → Up to 5 sentences. Up to 3 named drivers with approximate amounts.
        Still no forward-looking language, no causal inference.

Concrete enforcement: if Real Estate Taxes has a $360 quarterly variance, the
correct bullet is ONE sentence, not five. Do NOT pad with speculation about
"Q4 accrual" or "timing" or "expected to normalize".


══════════════════════════════════════════════════════════════════════
SILENCE OVER FILLER
══════════════════════════════════════════════════════════════════════

If the JSON payload does not provide a specific driver (no partner comment
for this line and no top_driver above the $5k materiality threshold), your
options are:
  1. State magnitude + direction only in one sentence, and stop.
  2. If |variance| < $5,000, skip the bullet entirely.

NEVER write filler like "No specific driver identified" or "Performance was
in line with expectations". Silence is strictly better than fluff.


══════════════════════════════════════════════════════════════════════
CAPEX — MAGNITUDE ONLY
══════════════════════════════════════════════════════════════════════

For the Capex bullet (if written at all):
  • State magnitude and direction ONLY.
  • Do NOT characterize the nature of the spend ("planned works", "new
    initiatives", "routine replacements", "completion of the program", etc.)
    unless a partner comment uses those exact words.
  • If Capex activity is zero or immaterial (< $5k), skip the bullet entirely.
  • The analyst may have reclassified accounts INTO or OUT OF Capex at Step 2
    of the pipeline. The JSON payload numbers already reflect those moves —
    use them as-is. Do NOT speculate about whether a reclassification happened
    or its purpose.


══════════════════════════════════════════════════════════════════════
OPEX TOTAL-FIRST RULE (mandatory — feedback Florencia 23-Apr-2026)
══════════════════════════════════════════════════════════════════════

For the **Operating Expenses** bullet specifically (and only this one):

  1. The FIRST sentence MUST state the total Operating Expenses variance
     for the quarter (and YTD if applicable).
     Example: "Operating Expenses were $21,764 unfavorable to budget for the
     quarter, while YTD performance remains $23,099 favorable."

  2. ONLY AFTER stating the total can you list specific drivers (Payroll,
     R&M, Utilities, etc.) — and in DESCENDING order of |variance|.

  3. NEVER lead with a sub-driver (e.g., "Maintenance ran $18k over budget")
     before establishing the total. The reader must see the headline first,
     then drill down.

  4. If you list 2+ drivers, end with the closing total reference if helpful:
     "As a result, Total Operating Expenses were $X unfavorable to budget."

This rule applies to **Operating Expenses only**. For Income, Real Estate
Taxes, Non-Operating, Interest, and Capex bullets — those typically have
fewer L2 sub-lines, so the standard structure (variance + drivers) suffices.

══════════════════════════════════════════════════════════════════════
DRIVER HIERARCHY (mandatory — applies to all multi-driver bullets)
══════════════════════════════════════════════════════════════════════

When listing 2+ drivers in any bullet:
  • Order them by |variance| DESCENDING (largest dollar impact first).
  • Each driver gets a quantified amount; if you can't quantify, drop it.
  • Cap at 3 drivers per bullet. If a 4th driver is material (>$5k), use
    "and other smaller items totaling ~$Xk" as a closing aggregation.
  • The 2–3 drivers you choose MUST collectively explain ≥70% of the
    headline L1 variance.

══════════════════════════════════════════════════════════════════════
PARTNER COMMENT METRIC PRESERVATION
══════════════════════════════════════════════════════════════════════

Partner / operator variance notes often contain SPECIFIC quantitative metrics
embedded in prose. If you find numbers in partner comments, preserve them
EXACTLY — do not round, generalize, or omit:
  • Renewal rate percentages (e.g., "4.78%, 8.70%, 7.10% in Oct/Nov/Dec")
  • Vacancy / occupancy ranges (e.g., "5.0%–6.1% vs budgeted 4.8%")
  • Concession counts (e.g., "12 move-in concessions, one month free")
  • Lease counts, unit counts, days, etc.

These metrics ARE the operational color — they're more credible than
generic phrasing. Quote them verbatim or paraphrase tightly.

══════════════════════════════════════════════════════════════════════
MANDATORY: OPERATING EXPENSES vs TOTAL OPERATING EXPENSES
══════════════════════════════════════════════════════════════════════

The JSON payload contains THREE distinct expense lines. You MUST use them correctly:

  1. "Operating Expenses" — excludes Real Estate Taxes.
     → Use THIS line's variance for the "Operating Expenses" bullet.
  2. "Real Estate Taxes" — property taxes only.
     → Use THIS line's variance for the "Real Estate Taxes" bullet.
  3. "Total Operating Expenses" — equals Operating Expenses + Real Estate Taxes.
     → Do NOT write a separate bullet for this line.
     → You MAY reference it as a closing sentence in either bullet, e.g.:
       "As a result, Total Operating Expenses were $X favorable to budget."

RULES:
  • NEVER describe "Total Operating Expenses" as if it were "Operating Expenses."
  • NEVER use the Total Operating Expenses variance in the Operating Expenses bullet.
  • ALWAYS verify: Total OpEx variance ≈ OpEx variance + RET variance.
    The JSON includes a validation field ("opex_split_validation"). If it says
    the check failed, flag the discrepancy and do NOT publish the text.
  • When discussing OpEx drivers (L2 categories like Payroll, R&M, Utilities),
    they belong to the "Operating Expenses" bullet, not to Real Estate Taxes.

═══════════════════════════════════════════════════════════════════

• Cross-line reclassifications: ONLY if a partner comment in the JSON explicitly
  states a dollar-specific reclassification between two lines (see Forbidden
  Pattern E above). In that case, mention the specific dollar amount in BOTH
  affected bullets, paraphrased in institutional tone. Otherwise, LINE ISOLATION.
• Partner comments: the JSON payload contains operator variance notes. Incorporate
  them as operational color — paraphrase in institutional tone, do NOT copy
  verbatim. If a partner comment uses a specific word (e.g. "timing", "one-off",
  "reassessment"), you may reuse it — otherwise stay descriptive only.
• No account codes: never cite GL codes (e.g., "4510-0020"). Use only readable names.
• No disclaimers, footnotes, or meta-commentary.

════════════════════════════════════════════════════
STYLE / TONE REFERENCE — EXAMPLES FROM A DIFFERENT ASSET
(use ONLY for voice; ignore all numbers, properties, and periods)
════════════════════════════════════════════════════

⚠️  CRITICAL: the four paragraphs below are VERBATIM EXCERPTS from prior STARS REI
quarterly reports of a DIFFERENT ASSET (a multi-family property — used here purely
as a style reference). They are NOT about {building}. They are NOT about {period}.

You MUST ignore:
  • the property names mentioned in the examples
  • the periods cited in the examples
  • the dollar amounts and drivers in the examples

You MUST extract ONLY:
  • institutional tone
  • sentence rhythm
  • level of numerical specificity (~$Xk anchors)
  • how timing vs. structural variances are framed
  • how cross-line reclassifications are explained in BOTH affected bullets

The narrative content for **{building}** must come from the financials_*.md
files in this SharePoint folder and from the JSON payload in STEP 4 — never
from the examples below.

── [OTHER-ASSET EXAMPLE] Income ──
"Income has been generally in line with budget. The negative variance is mainly due to
higher vacancy levels and increased bad debt. Bad debt is particularly an issue at one
of the buildings, though efforts are underway to mitigate it. On a positive note, concessions
are performing better than budgeted – something we hadn't seen in a long time."

── [OTHER-ASSET EXAMPLE] Operating Expenses ──
"Operating Expenses: the positive Q3 variance reflects the same $39,613 reallocation now
recorded as negative opex. Additionally, operating expenses performed better than budget in
R&M and contract services due to the cancellation of the valet parking contract.
On the other hand, payroll has exceeded budget as a result of staffing changes implemented
during the year."

── [OTHER-ASSET EXAMPLE] Real Estate Taxes ──
"Real Estate Taxes: taxes decreased in the second half of the year following the reduction in
assessed value. This positive variance will continue increasing through year-end."

── [OTHER-ASSET EXAMPLE] Non-Operating Expenses ──
"Non-Operating Expenses: of the negative Q3 variance, $8,400 is timing related, as the
expense had been budgeted in the prior quarter. The remaining balance and what accounts for
the YTD deviation is due to higher than expected legal and accounting fees."

⚠️  REPEAT: the bullets you produce must be about **{building}** for **{period}**.
Pull the narrative content from the financials_*.md files in this SharePoint
folder and from the JSON payload in STEP 4 — never from the reference examples above.

════════════════════════════════════════════════════
STEP 4 — STRUCTURED DATA REFERENCE (JSON backup)
════════════════════════════════════════════════════

The following JSON mirrors the data in the financials .md file in structured format.
Use it as a CROSS-REFERENCE if you need precise account-level detail not captured
in the markdown narrative. The .md file remains the primary source.

{json_payload}

════════════════════════════════════════════════════
OUTPUT
════════════════════════════════════════════════════

Generate ONLY the "Notes to Financials:" section — the bullets, nothing else.
No headers, no table, no introduction, no closing remarks.
The text must be ready to paste directly into the investor report after the
Portfolio Level Financials table.

Write in English. Each bullet is a continuous paragraph.
"""

    return prompt


def build_qoq_summary_table(historical_analysis: Dict) -> Optional[pd.DataFrame]:
    """Construye tabla resumen QoQ para display en la app."""
    if not historical_analysis or not historical_analysis.get('has_prior'):
        return None

    rows = []
    for s in historical_analysis.get('qoq_summary', []):
        sign = _get_display_sign(s.line_name)
        row = {
            'Line': s.line_name,
            'Q Actual': s.actual_current_q * sign,
            'Prior Q Actual': (s.actual_prior_q or 0) * sign,
            'QoQ Change': (s.qoq_actual_change or 0) * sign,
            'QoQ %': f"{(s.qoq_actual_change_pct or 0)*100:+.1f}%",
            'Q Variance': s.variance_current_q * sign,
            'Prior Q Variance': (s.variance_prior_q or 0) * sign,
            'Trend': s.trend_direction.title() if s.trend_direction else "—",
        }
        rows.append(row)

    if not rows:
        return None
    return pd.DataFrame(rows)


def build_account_changes_table(historical_analysis: Dict) -> tuple:
    """Construye tablas de cuentas nuevas y eliminadas."""
    if not historical_analysis or not historical_analysis.get('has_prior'):
        return None, None

    changes = historical_analysis.get('account_changes')
    if not changes:
        return None, None

    new_df = None
    if changes.new_accounts:
        new_df = pd.DataFrame(changes.new_accounts)

    removed_df = None
    if changes.removed_accounts:
        removed_df = pd.DataFrame(changes.removed_accounts)

    return new_df, removed_df


# --- Helpers ---

def _fmt_currency(val: float) -> str:
    """Formatea valor como moneda."""
    if val is None:
        return "n/a"
    if val < 0:
        return f"(${abs(val):,.0f})"
    return f"${val:,.0f}"


def _fmt_pct(val: Optional[float]) -> str:
    """Formatea porcentaje."""
    if val is None:
        return "n/a"
    return f"{val * 100:+.1f}%"


def _compute_display_variance(actual: float, budget: float, sign: int) -> float:
    """
    Calcula la varianza en convención de reporte.
    Con sign=-1 (gastos): Variance = (-actual) - (-budget) = budget - actual
    Con sign=+1 (ingresos): Variance = actual - budget
    En ambos casos, varianza positiva = favorable.
    """
    return (actual * sign) - (budget * sign)


def _is_favorable(variance: float, line_name: str) -> bool:
    """
    Determina si la variance es favorable según el tipo de línea.
    Nota: esta función recibe la variance RAW (sin signo display).
    - Ingresos: + = favorable (ganaste más)
    - Gastos: - = favorable (gastaste menos)
    - NOI/Net Income: + = favorable
    """
    if line_name in ('Income',):
        return variance > 0
    if line_name in ('Operating Expenses', 'Real Estate Taxes', 'Non-Operating Expenses', 'Interest Expense'):
        return variance < 0
    return variance > 0


def _df_to_markdown_table(df: pd.DataFrame) -> str:
    """Convierte DataFrame a tabla markdown."""
    if df.empty:
        return "_No data_"

    # Formatear números
    formatted = df.copy()
    for col in formatted.columns:
        if formatted[col].dtype in ('float64', 'int64'):
            formatted[col] = formatted[col].apply(
                lambda x: f"${x:,.0f}" if pd.notna(x) else "" if abs(x) < 0.005 else f"${x:,.0f}"
                if pd.notna(x) else "")
        else:
            formatted[col] = formatted[col].fillna("")

    # Header
    headers = " | ".join(str(c) for c in formatted.columns)
    separator = " | ".join("---" for _ in formatted.columns)
    lines = [f"| {headers} |", f"| {separator} |"]

    # Rows
    for _, row in formatted.iterrows():
        vals = " | ".join(str(v) for v in row)
        lines.append(f"| {vals} |")

    return "\n".join(lines)
