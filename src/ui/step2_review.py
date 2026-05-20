# src/ui/step2_review.py

"""UI del Paso 2: revisión de validaciones, tabla resumen, clasificación de cuentas, QoQ y mercado."""

import pandas as pd
import streamlit as st

from src.config import L1_OPTIONS
from src.pipeline import _apply_l1_overrides, revalidate_structure, _is_cross_side, regenerate_outputs
from src.output_formatter import (
    build_summary_table,
    build_bullets_by_line,
    build_qoq_summary_table,
    build_account_changes_table,
    generate_final_classification_excel,
    _sort_accounts_partner_order,
)
from src.market_analysis import market_kpi_cards, market_summary_text


def render_step2():
    """
    Renderiza el Paso 2: revisión de validaciones, tabla resumen, reconciliación P&L,
    clasificación de cuentas (editor interactivo), análisis QoQ y análisis de mercado.
    """
    st.title("Revisión y Validación")
    st.markdown("Revisa los resultados antes de generar el output final. **Este paso es obligatorio.**")

    results = st.session_state.results

    if not results:
        st.warning("No hay resultados. Vuelve al Paso 1.")
        if st.button("← Volver"):
            st.session_state.step = 1
            st.rerun()
        st.stop()

    # ---- Validaciones ----
    st.subheader("Validaciones")
    validations = results['validations']
    has_block = False

    for v in validations:
        if v.status == "OK":
            st.success(f"✅ **{v.check_name}:** {v.message}")
        elif v.status == "WARNING":
            st.warning(f"⚠️ **{v.check_name}:** {v.message}")
            if v.details:
                with st.expander(f"Detalle: {v.check_name}"):
                    if 'new' in v.details and v.details['new']:
                        st.markdown("**Cuentas nuevas:**")
                        st.dataframe(pd.DataFrame(v.details['new']))
                    if 'removed' in v.details and v.details['removed']:
                        st.markdown("**Cuentas eliminadas:**")
                        st.dataframe(pd.DataFrame(v.details['removed']))
                    if 'outlier_details' in v.details:
                        outlier_df = pd.DataFrame(v.details['outlier_details'])
                        st.dataframe(outlier_df, width="stretch")
                    elif 'items' in v.details:
                        for item in v.details['items']:
                            st.text(f"  • {item}")
        elif v.status == "BLOCK":
            st.error(f"❌ **{v.check_name}:** {v.message}")
            has_block = True

    # ---- Metadata de archivos ----
    st.subheader("Archivos procesados")
    meta_rows = []
    for ing in results.get('ingested_files', []):
        meta_rows.append({
            'Archivo': ing.filename,
            'Edificio': ing.building,
            'Período': ing.period_label,
            'Hoja usada': ing.metadata.get('sheet_used', '(primera)'),
            'Hojas analizadas': ing.metadata.get('sheets_analyzed', 1),
            'Hoja comentarios': ing.metadata.get('comments_sheet') or '(inline)',
            'Filas datos': ing.metadata.get('data_rows', 0),
            'Comentarios socio': len(ing.partner_comments),
            'Tiene YTD': '✅' if ing.metadata.get('has_ytd') else '❌',
        })
    if meta_rows:
        st.dataframe(pd.DataFrame(meta_rows), width="stretch")

    # ---- Vista previa tabla resumen ----
    st.subheader("Vista previa: Tabla Resumen Financiero")
    summary = results['summary_table']
    st.dataframe(summary.style.format({
        col: "${:,.0f}" for col in summary.columns if col != 'Line'
    }), width="stretch")

    # ── ACTIVOS INDIVIDUALES (si es consolidación de portafolio) ──
    if results['structure'].get('individual_structures'):
        st.subheader("📦 Activos del Portafolio (detalle individual)")
        individual = results['structure']['individual_structures']
        tab_names = ["🏢 Consolidado"] + list(individual.keys())
        tabs = st.tabs(tab_names)

        # Tab consolidado
        with tabs[0]:
            consolidated_summary = build_summary_table(results['structure'])
            st.dataframe(consolidated_summary.style.format({
                col: "${:,.0f}" for col in consolidated_summary.columns if col != 'Line'
            }), width="stretch")

        # Tabs individuales por edificio
        for tab, (bname, struct) in zip(tabs[1:], individual.items()):
            with tab:
                ind_summary = build_summary_table(struct)
                st.dataframe(ind_summary.style.format({
                    col: "${:,.0f}" for col in ind_summary.columns if col != 'Line'
                }), width="stretch")

    # ── RECONCILIACIÓN P&L: Total Socio vs Clasificación Interna ──
    st.subheader("📊 Reconciliación P&L — Total Socio vs Suma L3")
    st.markdown(
        "Compara el total reportado por el socio (L1) con la suma de las cuentas L3 "
        "clasificadas bajo cada línea. Asegura que no se pierden cuentas en el proceso."
    )

    # Find the Full P&L reconciliation result from validations
    _pl_recon = None
    for v in validations:
        if v.check_name == "Reconciliación P&L" and v.details:
            _pl_recon = v.details
            break

    l3_accts = results['structure'].get('l3_accounts', [])

    if _pl_recon:
        recon_rows = []
        # Formulas que se muestran para las lineas calculadas (UX feedback Florencia):
        # antes mostraban "None" en Suma L3/Dif/#Ctas porque son derivadas, no
        # tienen L3s propias. Ahora explicitamos la formula.
        _CALC_FORMULAS = {
            'NOI': 'Calculado: Income − OpEx − RET',
            'Net Income': 'Calculado: NOI − Interest − Non-Op',
        }
        for row in _pl_recon['per_l1']:
            if row['status'] == 'calc':
                status_icon = 'ℹ️'
            elif row['status'] == 'OK':
                status_icon = '✅'
            elif row['status'] == 'WARNING':
                status_icon = '⚠️'
            else:
                status_icon = '❌'

            if row['status'] == 'calc':
                formula = _CALC_FORMULAS.get(row['line'], 'Línea derivada (sin L3 directas)')
                recon_rows.append({
                    'Línea L1': row['line'],
                    'Total Socio': row['partner_total'],
                    'Suma L3 (Clasificación)': formula,
                    'Diferencia': '—',
                    '# Cuentas': '—',
                    'Status': status_icon,
                })
            else:
                recon_rows.append({
                    'Línea L1': row['line'],
                    'Total Socio': row['partner_total'],
                    'Suma L3 (Clasificación)': row['classification_sum'],
                    'Diferencia': row['diff'],
                    '# Cuentas': row['account_count'],
                    'Status': status_icon,
                })

        # NOTA: Se removió la fila "TOTAL P&L" (feedback Florencia 23-04-2026):
        # mezclaba Total Capex con líneas P&L y enredaba la lectura.

        # Unclassified row
        if _pl_recon['unclassified_count'] > 0:
            recon_rows.append({
                'Línea L1': 'Sin Clasificar',
                'Total Socio': None,
                'Suma L3 (Clasificación)': _pl_recon['unclassified_sum'],
                'Diferencia': None,
                '# Cuentas': _pl_recon['unclassified_count'],
                'Status': '⚠️',
            })

        recon_df = pd.DataFrame(recon_rows)
        # Si el valor es numerico, formatear como USD. Si es string (caso de
        # las lineas calculadas que muestran la formula), dejarlo tal cual.
        def numeric_fmt(v):
            if isinstance(v, (int, float)) and v is not None:
                return f"${v:,.0f}"
            if isinstance(v, str) and v:
                return v
            return "—"
        st.dataframe(
            recon_df.style.format({
                'Total Socio': numeric_fmt,
                'Suma L3 (Clasificación)': numeric_fmt,
                'Diferencia': numeric_fmt,
            }, na_rep="—"),
            width="stretch",
        )

    # ── CUENTAS NO CLASIFICADAS (detalle) ──
    unknown_accts = [a for a in l3_accts if a.parent_line in ('Unknown', 'unknown', '')]
    if unknown_accts:
        total_unknown = sum(a.actual_current for a in unknown_accts)
        with st.expander(
            f"⚠️ {len(unknown_accts)} cuenta(s) no clasificada(s) — "
            f"Actual total: ${total_unknown:,.0f}",
            expanded=True,
        ):
            st.markdown(
                "Estas cuentas no fueron asignadas a ninguna línea L1. "
                "Identifica la sección en el Excel y agrégala a `SECTION_TO_L1` en `native_structure.py`."
            )
            unknown_df = pd.DataFrame([{
                'Código': a.account_code,
                'Descripción': a.name,
                'Budget Q': a.budget_current,
                'Actual Q': a.actual_current,
                'Varianza Q': a.variance_current,
            } for a in unknown_accts])
            st.dataframe(
                unknown_df.style.format({
                    'Budget Q': "${:,.0f}",
                    'Actual Q': "${:,.0f}",
                    'Varianza Q': "${:,.0f}",
                }),
                width="stretch",
            )

    st.markdown("---")

    # ── VALIDACIÓN DE CLASIFICACIÓN DE CUENTAS (Gate obligatorio) ──
    st.subheader("🏷️ Validación de Clasificación de Cuentas")
    st.markdown(
        "Revisa que cada cuenta esté asignada a la **línea de reporte (L1)** correcta. "
        "Si alguna clasificación es incorrecta, cámbiala en la columna **'L1 Corregido'**. "
        "Para cuentas que **no deben afectar Net Income**: usá **Total Capex** si son "
        "inversiones de capital (se trackean como L1 propia debajo del NI), o **NA** si "
        "se excluyen completamente. **Debes aprobar la clasificación para continuar.**"
    )

    l3_accounts = results['structure'].get('l3_accounts', [])

    if l3_accounts:
        # Buffer persistente de ediciones — sobrevive cambios de filtro.
        # Clave = código de cuenta; valor = L1 corregido (puede diferir o no del auto).
        if 'pending_l1_overrides' not in st.session_state:
            st.session_state.pending_l1_overrides = dict(st.session_state.get('l1_overrides') or {})

        # Construir DataFrame completo usando las ediciones pendientes como valor actual.
        # Ordenar por flujo natural del P&L del partner: L1 canónico (Income → OpEx →
        # RET → NOI → Interest → Non-Op → NI → Capex → NA) → L2 orden de aparición en
        # el Excel → partner_row_order dentro de cada L2. Esto permite al analista
        # validar las cuentas leyendo top-down, igual que el partner redactó el P&L.
        ordered_accounts = _sort_accounts_partner_order(l3_accounts)

        # Iconos por fuente de clasificación (tooltip lo aclara)
        _SOURCE_ICONS = {
            'catalog':             '📘',  # asset_catalogs.py (Florencia)
            'section':             '📄',  # heredado del section header del Excel
            'code_range':          '🔢',  # fallback por rango de código
            'description_keyword': '🔤',  # fallback por keyword
            'unknown':             '❓',  # ningún criterio matchó
            'user_override':       '✏️',  # reclassification manual
            'cross_side_flip':     '🔁',  # override que cruzó Income↔Expense
            'post_processing':     '⚙️',  # reclasificación interna
        }

        class_rows = []
        for acct in ordered_accounts:
            current = st.session_state.pending_l1_overrides.get(
                acct.account_code, acct.parent_line
            )
            auto_l1 = acct.parent_line
            modified = (current != auto_l1)
            source = getattr(acct, 'classification_source', 'unknown') or 'unknown'
            class_rows.append({
                'Origen': _SOURCE_ICONS.get(source, '·'),
                'Código': acct.account_code,
                'Descripción': acct.name,
                'L1 Asignado (auto)': auto_l1,
                'L1 Corregido': current,
                'Budget Q': acct.budget_current,
                'Actual Q': acct.actual_current,
                'Varianza': acct.variance_current,
                '_modified': modified,
                '_source': source,
            })
        class_df = pd.DataFrame(class_rows)

        # ── Resumen rápido (KPIs visuales) ────────────────────────────
        n_total = len(class_df)
        n_modified = int(class_df['_modified'].sum())
        n_unknown = int((class_df['L1 Corregido'].isin(['Unknown', 'unknown', ''])).sum())
        n_capex = int((class_df['L1 Corregido'] == 'Total Capex').sum())
        n_na = int((class_df['L1 Corregido'] == 'NA').sum())
        kpi_cols = st.columns(5)
        kpi_cols[0].metric("Total cuentas", n_total)
        kpi_cols[1].metric("✏️ Modificadas", n_modified)
        kpi_cols[2].metric("❓ Sin clasificar", n_unknown,
                           delta="revisar" if n_unknown else None,
                           delta_color="inverse" if n_unknown else "off")
        kpi_cols[3].metric("📐 Total Capex", n_capex)
        kpi_cols[4].metric("🚫 NA", n_na)

        # ── Filtros rápidos (toggles) ─────────────────────────────────
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            only_modified = st.checkbox("Solo modificadas", key="class_only_modified")
        with fc2:
            only_unknown = st.checkbox("Solo sin clasificar", key="class_only_unknown")
        with fc3:
            only_material = st.checkbox("Solo varianza > $5k", key="class_only_material")

        # ── Buscador / filtro ────────────────────────────────────────
        sc1, sc2 = st.columns([2, 1])
        with sc1:
            search = st.text_input(
                "🔍 Buscar cuenta (código o descripción)",
                key="class_search_text",
                placeholder="ej. 6209, SEWER, payroll…",
            ).strip().lower()
        with sc2:
            l1_filter = st.multiselect(
                "Filtrar por L1",
                options=L1_OPTIONS,
                default=[],
                key="class_search_l1",
            )

        display_df = class_df
        if search:
            mask = (
                class_df['Código'].astype(str).str.lower().str.contains(search, na=False)
                | class_df['Descripción'].astype(str).str.lower().str.contains(search, na=False)
            )
            display_df = class_df[mask]
        if l1_filter:
            display_df = display_df[display_df['L1 Corregido'].isin(l1_filter)]
        if only_modified:
            display_df = display_df[display_df['_modified']]
        if only_unknown:
            display_df = display_df[display_df['L1 Corregido'].isin(['Unknown', 'unknown', ''])]
        if only_material:
            display_df = display_df[display_df['Varianza'].abs() >= 5000]
        # Ocultar columnas internas antes de mostrar
        display_df = display_df.drop(columns=['_modified', '_source'], errors='ignore')
        display_df = display_df.reset_index(drop=True)

        active_filters = []
        if search: active_filters.append(f"texto='{search}'")
        if l1_filter: active_filters.append(f"L1={l1_filter}")
        if only_modified: active_filters.append("solo modificadas")
        if only_unknown: active_filters.append("solo sin clasificar")
        if only_material: active_filters.append("|var|≥$5k")
        if active_filters:
            st.caption(
                f"Mostrando **{len(display_df)}** de **{len(class_df)}** "
                f"cuentas — filtros: {', '.join(active_filters)}"
            )

        if len(display_df) == 0:
            st.info("Sin resultados para el filtro actual.")
            edited_df = display_df
        else:
            # Key del editor depende del filtro para evitar conflictos de estado interno
            # cuando el set de filas visibles cambia entre reruns.
            editor_key = (
                f"classification_editor::{search}::{','.join(sorted(l1_filter))}"
                f"::{int(only_modified)}{int(only_unknown)}{int(only_material)}"
            )
            edited_df = st.data_editor(
                display_df,
                column_config={
                    'Origen': st.column_config.TextColumn(
                        'Origen', disabled=True, width='small',
                        help="📘 catálogo · 📄 sección · 🔢 código · 🔤 keyword · ❓ desconocido · ✏️ override · 🔁 cross-side flip · ⚙️ post-processing"
                    ),
                    'Código': st.column_config.TextColumn('Código', disabled=True, width='small'),
                    'Descripción': st.column_config.TextColumn('Descripción', disabled=True, width='large'),
                    'L1 Asignado (auto)': st.column_config.TextColumn('L1 Auto', disabled=True, width='medium'),
                    'L1 Corregido': st.column_config.SelectboxColumn(
                        'L1 Corregido',
                        options=L1_OPTIONS,
                        required=True,
                        width='medium',
                        help="Selecciona la línea L1 correcta para esta cuenta",
                    ),
                    'Budget Q': st.column_config.NumberColumn('Budget Q', format="$%d", disabled=True),
                    'Actual Q': st.column_config.NumberColumn('Actual Q', format="$%d", disabled=True),
                    'Varianza': st.column_config.NumberColumn('Varianza', format="$%d", disabled=True),
                },
                width="stretch",
                hide_index=True,
                num_rows="fixed",
                key=editor_key,
            )

            # Persistir ediciones visibles en el buffer — edits sobre filas filtradas
            # se conservan cuando el usuario cambia el filtro.
            for _, row in edited_df.iterrows():
                code = row['Código']
                new_l1 = row['L1 Corregido']
                auto_l1 = row['L1 Asignado (auto)']
                if new_l1 != auto_l1:
                    st.session_state.pending_l1_overrides[code] = new_l1
                else:
                    # Si volvió al auto, eliminar del buffer
                    st.session_state.pending_l1_overrides.pop(code, None)

        # Contar cambios sobre el buffer completo (no solo lo visible)
        auto_by_code = {r['Código']: r['L1 Asignado (auto)'] for r in class_rows}
        desc_by_code = {r['Código']: r['Descripción'] for r in class_rows}
        actual_by_code = {r['Código']: r['Actual Q'] for r in class_rows}
        changes_made = []
        cross_side_count = 0
        for code, new_l1 in st.session_state.pending_l1_overrides.items():
            auto = auto_by_code.get(code)
            if auto and new_l1 != auto:
                crosses = _is_cross_side(auto, new_l1)
                if crosses:
                    cross_side_count += 1
                changes_made.append({
                    'code': code,
                    'desc': desc_by_code.get(code, ''),
                    'from': auto,
                    'to': new_l1,
                    'crosses_side': crosses,
                    'actual_q': actual_by_code.get(code, 0) or 0,
                })

        if changes_made:
            st.info(f"📝 {len(changes_made)} reclasificación(es) pendiente(s):")
            for ch in changes_made:
                base = (
                    f"- **{ch['code']}** ({ch['desc']}): "
                    f"~~{ch['from']}~~ → **{ch['to']}**"
                )
                if ch['crosses_side']:
                    flipped = -(ch['actual_q'] or 0)
                    base += (
                        f"  💱 _cambio de signo_: "
                        f"Actual Q `${ch['actual_q']:,.0f}` → `${flipped:,.0f}` "
                        f"(Income ↔ Expense)"
                    )
                st.markdown(base)
            if cross_side_count:
                st.caption(
                    f"💱 {cross_side_count} reclasificación(es) cruzan Income↔Expense: "
                    "el signo de la cuenta se invierte automáticamente al aprobar "
                    "para mantener la consistencia de NOI."
                )

        # Botones de clasificación
        cl_col1, cl_col2 = st.columns(2)
        with cl_col1:
            if st.button(
                "✅ Aprobar clasificación" + (f" ({len(changes_made)} cambios)" if changes_made else ""),
                type="primary",
                width="stretch",
                key="approve_classification",
            ):
                # Guardar overrides desde el buffer completo (no solo lo visible)
                overrides = {
                    code: new_l1
                    for code, new_l1 in st.session_state.pending_l1_overrides.items()
                    if auto_by_code.get(code) and new_l1 != auto_by_code[code]
                }
                st.session_state.l1_overrides = overrides
                st.session_state.classification_approved = True

                # Si hay cambios, re-aplicarlos a la estructura
                if overrides:
                    _apply_l1_overrides(results['structure'], overrides)

                # Revalidar cuadratura con la estructura reclasificada.
                # ORDEN CRÍTICO: primero revalidate (actualiza results['validations']),
                # luego regenerate_outputs (que usa validations para armar alerts + markdown).
                revalidate_structure(results)

                # Regenerar TODOS los outputs desde la estructura post-override:
                # markdown, copilot_prompt, json_payload, excel_bytes, tracking_excel,
                # alerts_text, comments_df, summary_table, bullets. Garantiza que
                # Step 3 y el AI usen la clasificación FINAL, no un estado intermedio.
                regenerate_outputs(results)

                st.session_state.results = results
                st.session_state.revalidated = True
                # Re-abrir finalización: el usuario debe confirmar de nuevo
                st.session_state.classification_finalized = False

                st.rerun()

        with cl_col2:
            if st.session_state.classification_approved:
                st.success("✅ Clasificación aprobada por el analista")
            else:
                st.warning("⏳ Pendiente de aprobación")

        # ── Descarga de clasificación final (solo tras aprobación) ──
        # Artefacto de auditoría: muestra L3 → L1 FINAL usado para generar el prompt
        # del AI, incluyendo la fuente de cada decisión (catalog / section / fallback /
        # user_override / cross_side_flip). Para revisar antes de ir a Step 3.
        if st.session_state.classification_approved:
            st.markdown("#### 📥 Clasificación final (auditoría)")
            st.caption(
                "Excel con la clasificación L3 → L1 **final** que el AI usará para generar "
                "los notes. Incluye la fuente de cada decisión (catálogo, override del "
                "analista, etc.) para revisar antes de proceder al Step 3."
            )
            try:
                building_name = results.get('building', 'asset')
                period_name = results.get('period', 'period')
                classification_bytes = generate_final_classification_excel(
                    results['structure'], building=building_name, period=period_name
                )
                safe = "".join(c if c.isalnum() else "_" for c in f"{building_name}_{period_name}").strip("_")
                st.download_button(
                    label="📥 Descargar clasificación final (.xlsx)",
                    data=classification_bytes,
                    file_name=f"classification_final_{safe}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    key="download_final_classification",
                )
            except Exception as e:
                st.error(f"No se pudo generar el Excel de clasificación final: {e}")

        # ── Revalidación post-aprobación ──
        if st.session_state.classification_approved and st.session_state.get('revalidated'):
            na_accts = [a for a in l3_accounts if a.parent_line == 'NA']
            st.markdown("#### 🔁 Revalidación post-clasificación")
            if na_accts:
                na_sum = sum(a.actual_current for a in na_accts)
                st.caption(
                    f"ℹ️ {len(na_accts)} cuenta(s) marcada(s) como **NA** "
                    f"(${na_sum:,.0f} actual) — excluidas del cálculo de Net Income."
                )

            # Mostrar estado actualizado de las validaciones críticas
            new_vals = results.get('validations', [])
            critical_checks = {'Cuadratura OpEx', 'Cuadratura NOI', 'Reconciliación P&L', 'Líneas principales'}
            for v in new_vals:
                if v.check_name not in critical_checks:
                    continue
                if v.status == 'OK':
                    st.success(f"✅ **{v.check_name} (revalidado):** {v.message}")
                elif v.status == 'WARNING':
                    st.warning(f"⚠️ **{v.check_name} (revalidado):** {v.message}")
                elif v.status == 'BLOCK':
                    st.error(f"❌ **{v.check_name} (revalidado):** {v.message}")

            # Tabla de reconciliación P&L actualizada
            new_recon = None
            for v in new_vals:
                if v.check_name == "Reconciliación P&L" and v.details:
                    new_recon = v.details
                    break

            if new_recon:
                with st.expander("📊 Reconciliación P&L actualizada (post-reclasificación)", expanded=True):
                    recon_rows2 = []
                    for row in new_recon['per_l1']:
                        status_icon = {'calc': 'ℹ️', 'OK': '✅', 'WARNING': '⚠️'}.get(row['status'], '❌')
                        recon_rows2.append({
                            'Línea L1': row['line'],
                            'Total Socio': row['partner_total'],
                            'Suma L3 (post)': row['classification_sum'],
                            'Diferencia': row['diff'],
                            '# Cuentas': row['account_count'],
                            'Status': status_icon,
                        })
                    # Sin fila TOTAL P&L (ver nota arriba).
                    recon_df2 = pd.DataFrame(recon_rows2)
                    numeric_fmt = lambda v: f"${v:,.0f}" if isinstance(v, (int, float)) and v is not None else "—"
                    st.dataframe(
                        recon_df2.style.format({
                            'Total Socio': numeric_fmt,
                            'Suma L3 (post)': numeric_fmt,
                            'Diferencia': numeric_fmt,
                        }, na_rep="—"),
                        width="stretch",
                    )

            # ── Segundo paso: re-ajuste iterativo antes de finalizar ──
            st.markdown("#### 🔧 Ajuste adicional (opcional)")
            if not st.session_state.get('classification_finalized'):
                st.info(
                    "Si algo quedó pendiente, edita la tabla de clasificación arriba ☝️ "
                    "y vuelve a presionar **Aprobar clasificación** para revalidar. "
                    "Cuando estés conforme, presiona **Finalizar clasificación** para "
                    "bloquearla y ver el desglose de Operating Expenses."
                )
                if st.button(
                    "✔️ Finalizar clasificación",
                    type="primary",
                    width="stretch",
                    key="finalize_classification",
                ):
                    st.session_state.classification_finalized = True
                    st.rerun()
            else:
                st.success("🔒 Clasificación finalizada. Desglose de Operating Expenses disponible abajo.")
                if st.button(
                    "🔓 Reabrir clasificación",
                    width="stretch",
                    key="reopen_classification",
                ):
                    st.session_state.classification_finalized = False
                    st.rerun()

            # ── Desglose de Operating Expenses (post-finalización) ──
            if st.session_state.get('classification_finalized'):
                st.markdown("#### 📋 Desglose de Operating Expenses")
                opex_accts = [a for a in l3_accounts if a.parent_line == 'Operating Expenses']

                if not opex_accts:
                    st.warning("No hay cuentas clasificadas como **Operating Expenses**.")
                else:
                    opex_rows = []
                    for a in opex_accts:
                        opex_rows.append({
                            'Código': a.account_code,
                            'Descripción': a.name,
                            'Sección nativa': getattr(a, 'section_name', '') or '—',
                            'Budget Q': a.budget_current,
                            'Actual Q': a.actual_current,
                            'Varianza Q': a.variance_current,
                        })
                    # Totales
                    total_budget = sum(a.budget_current or 0 for a in opex_accts)
                    total_actual = sum(a.actual_current or 0 for a in opex_accts)
                    total_var = sum(a.variance_current or 0 for a in opex_accts)
                    opex_rows.append({
                        'Código': '',
                        'Descripción': f'TOTAL ({len(opex_accts)} cuentas)',
                        'Sección nativa': '',
                        'Budget Q': total_budget,
                        'Actual Q': total_actual,
                        'Varianza Q': total_var,
                    })
                    opex_df = pd.DataFrame(opex_rows)
                    st.dataframe(
                        opex_df.style.format({
                            'Budget Q': "${:,.0f}",
                            'Actual Q': "${:,.0f}",
                            'Varianza Q': "${:,.0f}",
                        }),
                        width="stretch",
                        hide_index=True,
                    )

                    # Comparación vs total socio
                    opex_line = next(
                        (l for l in results['structure'].get('l1_lines', [])
                         if l.name == 'Operating Expenses'),
                        None,
                    )
                    if opex_line is not None:
                        partner_total = opex_line.actual_current
                        diff = total_actual - partner_total
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Total Socio (L1)", f"${partner_total:,.0f}")
                        c2.metric("Suma L3 (desglose)", f"${total_actual:,.0f}")
                        c3.metric("Diferencia", f"${diff:,.0f}",
                                  delta_color="inverse" if abs(diff) > 100 else "normal")

    st.markdown("---")

    # ---- Análisis QoQ (si hay datos del trimestre anterior) ----
    hist = results.get('historical_analysis')
    if hist and hist.get('has_prior'):
        st.subheader("Análisis Quarter-over-Quarter (QoQ)")

        qoq_table = build_qoq_summary_table(hist)
        if qoq_table is not None:
            st.dataframe(qoq_table.style.format({
                col: "${:,.0f}" for col in qoq_table.columns
                if col not in ('Line', 'QoQ %', 'Trend')
            }), width="stretch")

        # Alertas de tendencia
        trend_alerts = hist.get('trend_alerts', [])
        if trend_alerts:
            with st.expander(f"📈 Alertas de tendencia ({len(trend_alerts)})", expanded=True):
                for alert in trend_alerts:
                    st.markdown(f"- {alert}")

        # Cuentas nuevas/eliminadas
        new_df, removed_df = build_account_changes_table(hist)
        if new_df is not None or removed_df is not None:
            with st.expander("🔄 Cambios en cuentas (nuevas/eliminadas)", expanded=False):
                if new_df is not None and not new_df.empty:
                    st.markdown(f"**Cuentas nuevas ({len(new_df)}):**")
                    st.dataframe(new_df, width="stretch")
                if removed_df is not None and not removed_df.empty:
                    st.markdown(f"**Cuentas eliminadas ({len(removed_df)}):**")
                    st.dataframe(removed_df, width="stretch")

    # ---- ANÁLISIS DE MERCADO ----
    market = results.get('market_analysis')
    if market:
        st.markdown("---")
        st.subheader(f"📈 Análisis de Mercado — {market.market_name}")
        st.markdown(market_summary_text(market))

        # KPI Cards
        kpis = market_kpi_cards(market)
        kpi_cols = st.columns(len(kpis))
        for col, kpi in zip(kpi_cols, kpis):
            with col:
                st.metric(
                    label=kpi['label'],
                    value=kpi['value'],
                    delta=kpi['delta'],
                    delta_color=kpi['delta_color'],
                )

        # Charts (últimos 5 años = 20 trimestres)
        chart_df = market.df
        if chart_df is not None and not chart_df.empty:
            chart_recent = chart_df.tail(20).copy()
            chart_recent = chart_recent.set_index('Period')

            chart_tab1, chart_tab2, chart_tab3, chart_tab4 = st.tabs([
                "🏠 Rentas", "📊 Vacancy & Occupancy",
                "🏗️ Supply Pipeline", "📋 Data Table",
            ])

            with chart_tab1:
                st.line_chart(
                    chart_recent[['Asking Rent', 'Effective Rent']],
                    use_container_width=True,
                )
                if 'Concessions %' in chart_recent.columns:
                    st.bar_chart(
                        chart_recent[['Concessions %']],
                        use_container_width=True,
                    )

            with chart_tab2:
                st.line_chart(
                    chart_recent[['Vacancy %', 'Occupancy %']],
                    use_container_width=True,
                )

            with chart_tab3:
                supply_cols = ['Under Construction', 'Deliveries']
                available_cols = [c for c in supply_cols if c in chart_recent.columns]
                if available_cols:
                    st.bar_chart(
                        chart_recent[available_cols],
                        use_container_width=True,
                    )
                if 'Absorption' in chart_recent.columns:
                    st.line_chart(
                        chart_recent[['Absorption']],
                        use_container_width=True,
                    )

            with chart_tab4:
                st.dataframe(
                    chart_recent.reset_index().style.format({
                        'Asking Rent': '${:,.0f}',
                        'Effective Rent': '${:,.0f}',
                        'Vacancy %': '{:.1f}%',
                        'Occupancy %': '{:.1f}%',
                        'Concessions %': '{:.2f}%',
                    }),
                    width="stretch",
                )

    # ---- Comentarios del socio ----
    comments_df = results['comments_df']
    if not comments_df.empty:
        with st.expander(f"💬 Comentarios del socio ({len(comments_df)} notas)", expanded=False):
            st.dataframe(comments_df, width="stretch")

    # ---- Botones ----
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Volver a Configuración"):
            st.session_state.step = 1
            st.rerun()
    with col2:
        if has_block:
            st.error("⛔ Hay validaciones bloqueantes. Corrige los archivos y vuelve a cargar.")
        elif not st.session_state.classification_approved:
            st.warning("⏳ Debes aprobar la clasificación de cuentas antes de continuar (sección arriba ☝️).")
        elif not st.session_state.get('classification_finalized'):
            st.warning("⏳ Debes **finalizar** la clasificación antes de continuar (sección arriba ☝️).")
        else:
            if st.button("✅ Aprobar y generar output", type="primary", width="stretch"):
                st.session_state.approved = True
                st.session_state.step = 3
                st.rerun()
