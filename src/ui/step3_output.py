# src/ui/step3_output.py

"""UI del Paso 3: Management Commentary, análisis AI, botones de descarga."""

import pandas as pd
import streamlit as st
from datetime import datetime

from src.ai_analyst import generate_ai_analysis, build_ai_prompt, score_ai_output, OPENAI_MODELS, ANTHROPIC_MODELS
from src.management_commentary import (
    load_commentary,
    add_entry as add_commentary_entry,
    delete_entry as delete_commentary_entry,
    get_commentary_for_prompt,
    CommentaryEntry,
)
from src.master_excel import _extract_year_from_period as _year_from_period
from src.market_analysis import market_kpi_cards, market_summary_text


def render_step3():
    """
    Renderiza el Paso 3: Management Commentary, análisis AI, botones de descarga.
    """
    st.title("Output Final — Financials del Reporte")

    results = st.session_state.results
    if not results or not st.session_state.approved:
        st.warning("Debes aprobar en el Paso 2 antes de generar el output.")
        if st.button("← Volver"):
            st.session_state.step = 2
            st.rerun()
        st.stop()

    # ---- Tabla Resumen ----
    st.subheader("Tabla Resumen Financiero")
    summary = results['summary_table']
    st.dataframe(summary.style.format({
        col: "${:,.0f}" for col in summary.columns if col != 'Line'
    }), width="stretch")

    # ---- Bullets por línea (rule-based) ----
    with st.expander("📝 Notes to Financials (rule-based)", expanded=False):
        st.markdown(results['bullets'])

    # ── MANAGEMENT COMMENTARY — Memory Layer ──
    st.markdown("---")
    st.subheader("Management Commentary")
    st.markdown(
        "Agrega contexto cualitativo para enriquecer las Notes to Financials. "
        "Estos comentarios se guardan en el Master Excel y se reutilizan como "
        "contexto en futuros trimestres."
    )

    _mc_building = results['building']
    _mc_period = results['period']
    try:
        _mc_year = _year_from_period(_mc_period)
    except Exception:
        _mc_year = 2025

    # Load existing commentary
    if 'mc_entries' not in st.session_state:
        st.session_state.mc_entries = load_commentary(_mc_building, _mc_year)

    mc_entries = st.session_state.mc_entries

    # ── ORDEN LÓGICO (feedback Florencia 23-04-2026) ─────────────────────────
    # Reordena entries para que el analista los vea en flujo natural:
    #   1) Recurrentes primero (contexto persistente), luego Puntuales (eventos)
    #   2) Dentro de cada grupo: orden P&L canónico (Income → OpEx → RET → NOI →
    #      Interest → Non-Op → Net Income → Total Capex)
    #   3) Dentro de cada L1: período más reciente primero
    from src.config import l1_pl_sort_key
    def _mc_sort_key(e):
        type_order = 0 if e.comment_type == 'recurring' else 1
        return (type_order, l1_pl_sort_key(e.line_name), e.period or '')
    mc_entries_sorted = sorted(mc_entries, key=_mc_sort_key)

    # Display existing entries con orden lógico + agrupado visual
    if mc_entries_sorted:
        # Mantener el mapeo de # visible → idx real para el botón de eliminar
        st.session_state['_mc_display_to_real_idx'] = {
            i + 1: mc_entries.index(e) for i, e in enumerate(mc_entries_sorted)
        }

        # Agrupar por tipo para mostrar con headers
        recurring = [e for e in mc_entries_sorted if e.comment_type == 'recurring']
        one_off = [e for e in mc_entries_sorted if e.comment_type != 'recurring']

        def _render_group(entries, header, badge_color):
            if not entries:
                return
            st.markdown(
                f"<div style='margin-top:8px;margin-bottom:4px;'>"
                f"<span style='background:{badge_color};color:white;padding:3px 10px;"
                f"border-radius:12px;font-size:12px;font-weight:600;'>{header}</span>"
                f" <span style='color:#666;font-size:12px;'>({len(entries)})</span>"
                f"</div>", unsafe_allow_html=True,
            )
            rows = []
            for e in entries:
                visible_idx = next(
                    i + 1 for i, x in enumerate(mc_entries_sorted) if x is e
                )
                rows.append({
                    '#': visible_idx,
                    'Línea L1': e.line_name,
                    'Cuenta': e.account_code or '—',
                    'Período': e.period,
                    'Comentario': e.text,
                    'Autor': e.author,
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        _render_group(recurring, "🔁 RECURRENTES", "#0d9488")
        _render_group(one_off, "⚡ PUNTUALES", "#ea580c")

        # Delete entry — usa el # visible y resuelve al índice real
        del_col1, del_col2 = st.columns([1, 3])
        with del_col1:
            del_idx_visible = st.number_input(
                "# a eliminar", min_value=1, max_value=len(mc_entries_sorted),
                value=1, key="mc_del_idx",
            )
        with del_col2:
            if st.button("Eliminar entrada", key="mc_delete"):
                real_idx = st.session_state['_mc_display_to_real_idx'].get(
                    del_idx_visible, del_idx_visible - 1
                )
                st.session_state.mc_entries = delete_commentary_entry(
                    _mc_building, _mc_year, real_idx,
                )
                st.rerun()
    else:
        st.info("No hay comentarios guardados para este activo/año.")

    # Add new entry
    with st.expander("Agregar comentario", expanded=not mc_entries):
        L1_CHOICES = [
            "Income", "Operating Expenses", "Real Estate Taxes",
            "NOI", "Interest Expense", "Non-Operating Expenses", "Net Income",
            "Total Capex",
        ]
        mc_col1, mc_col2 = st.columns(2)
        with mc_col1:
            mc_line = st.selectbox("Línea L1", L1_CHOICES, key="mc_line")
            mc_type = st.radio(
                "Tipo", ["Recurrente", "Evento puntual"],
                horizontal=True, key="mc_type",
            )
        with mc_col2:
            mc_account = st.text_input(
                "Cuenta GL (opcional, para L3)",
                placeholder="e.g., 5010-0001",
                key="mc_account",
            )
            mc_author = st.text_input("Autor", value="Analyst", key="mc_author")

        mc_text = st.text_area(
            "Comentario",
            placeholder="Ej: El contrato de valet parking fue cancelado en Alice House. "
                        "El ahorro se refleja en Contract Services.",
            height=80,
            key="mc_text",
        )

        if st.button("Agregar", key="mc_add", type="primary"):
            if mc_text.strip():
                new_entry = CommentaryEntry(
                    account_code=mc_account.strip(),
                    line_name=mc_line,
                    level="L3" if mc_account.strip() else "L1",
                    period=_mc_period,
                    comment_type="recurring" if mc_type == "Recurrente" else "one_off",
                    text=mc_text.strip(),
                    author=mc_author.strip() or "Analyst",
                    timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                )
                st.session_state.mc_entries = add_commentary_entry(
                    _mc_building, _mc_year, new_entry,
                )
                st.rerun()
            else:
                st.warning("Escribe un comentario antes de agregar.")

    # ---- Alertas ----
    if results.get('alerts_text'):
        st.subheader("Alertas")
        st.markdown(results['alerts_text'])

    # ---- Comentarios del socio ----
    comments_df = results['comments_df']
    if not comments_df.empty:
        st.subheader(f"Comentarios del Socio ({len(comments_df)} notas)")
        st.dataframe(comments_df, width="stretch")

    # ---- Market Context (resumen compacto en output) ----
    market = results.get('market_analysis')
    if market:
        st.subheader(f"📈 Market Context — {market.market_name}")
        st.markdown(market_summary_text(market))
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

    # ── ANALYST BRIEFING — Contexto libre antes de generar prompt ──
    st.markdown("---")
    st.subheader("Analyst Briefing")
    st.markdown(
        "Escribe contexto general sobre el activo para este trimestre. "
        "Este texto se incorpora directamente al prompt del AI y al archivo "
        "descargable del Copilot Prompt. Usa este espacio para explicar eventos "
        "relevantes, cambios operativos, o cualquier detalle que el modelo necesite "
        "para generar notas precisas."
    )

    if 'analyst_briefing' not in st.session_state:
        st.session_state.analyst_briefing = ""

    analyst_briefing = st.text_area(
        "Briefing del analista",
        value=st.session_state.analyst_briefing,
        placeholder=(
            "Ej: En Q3 2025, Alice House tuvo cambio de property manager. "
            "Se cancelaron contratos de valet parking y laundry. "
            "Ocupación bajó a 92% por renovaciones en 4 unidades. "
            "Legal fees incluyen costos de litigio con contratista anterior."
        ),
        height=150,
        key="analyst_briefing_input",
    )
    st.session_state.analyst_briefing = analyst_briefing

    # ---- Descargas ----
    st.markdown("---")
    st.subheader("Descargas")

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📥 Descargar Markdown",
            data=results['markdown'],
            file_name=f"financials_{results['building']}_{results['period']}.md",
            mime="text/markdown",
            width="stretch",
        )

    with col2:
        st.download_button(
            "📥 Descargar Excel de Soporte",
            data=results['excel_bytes'],
            file_name=f"financials_support_{results['building']}_{results['period']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    col3, col4 = st.columns(2)

    with col3:
        st.download_button(
            "📥 Descargar JSON (Copilot Payload)",
            data=results['json_payload'],
            file_name=f"copilot_payload_{results['building']}_{results['period']}.json",
            mime="application/json",
            width="stretch",
        )

    with col4:
        # Build enriched prompt for download (includes briefing + commentary + market)
        _mc_text_dl = get_commentary_for_prompt(_mc_building, _mc_year)
        _prompt_dl = build_ai_prompt(
            copilot_prompt=results['copilot_prompt'],
            market_analysis=results.get('market_analysis'),
            custom_instructions=st.session_state.get('analyst_briefing', ''),
            management_commentary=_mc_text_dl,
        )
        st.download_button(
            "🤖 Descargar Prompt para Copilot",
            data=_prompt_dl,
            file_name=f"copilot_prompt_{results['building']}_{results['period']}.txt",
            mime="text/plain",
            width="stretch",
            help="Archivo de texto con instrucciones + datos + briefing del analista para pegar en ChatGPT/Claude",
        )

    # Fila adicional: Tracking de cuentas + Master Excel
    col5, col6 = st.columns(2)
    with col5:
        if results.get('tracking_excel'):
            st.download_button(
                "📊 Descargar Tracking de Cuentas",
                data=results['tracking_excel'],
                file_name=f"account_tracking_{results['building']}_{results['period']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                help="Excel con: cambios de cuentas (nuevas/eliminadas), comparación QoQ, top movers, inventario completo de cuentas",
            )

    with col6:
        if results.get('master_bytes'):
            st.download_button(
                "📥 Descargar Master Excel (actualizado)",
                data=results['master_bytes'],
                file_name=f"master_{results['building']}_{results['period']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                help="Master Excel con columnas mensuales + trimestrales + YTD. Reutilízalo en la siguiente iteración subiéndolo en Step 1.",
            )

    # ---- Vista del Markdown completo ----
    with st.expander("📄 Ver Markdown completo (para copiar)", expanded=False):
        st.code(results['markdown'], language="markdown")

    # ---- Build enriched prompt (with briefing + commentary + market) ----
    mc_text = get_commentary_for_prompt(_mc_building, _mc_year)
    enriched_prompt = build_ai_prompt(
        copilot_prompt=results['copilot_prompt'],
        market_analysis=results.get('market_analysis'),
        custom_instructions=st.session_state.get('analyst_briefing', ''),
        management_commentary=mc_text,
    )

    # ---- Vista del Prompt para Copilot ----
    with st.expander("🤖 Ver Prompt para Copilot (para copiar)", expanded=False):
        st.markdown(
            "Copia este texto completo y pégalo en ChatGPT, Claude u otro LLM para generar "
            "los *Notes to Financials* con el estilo institucional del reporte real. "
            "Incluye el Analyst Briefing y Management Commentary si los escribiste."
        )
        st.code(enriched_prompt, language="text")

    # ---- Quality Scorer: pegar output del LLM y evaluarlo ----
    st.markdown("---")
    st.subheader("🎯 Quality Score del output del LLM")
    st.markdown(
        "Pegá acá el texto que te devolvió ChatGPT / Claude / Gemini / Copilot. "
        "El sistema valida automáticamente: idioma, estructura de bullets, "
        "lenguaje prohibido (forward-looking, cross-line, 'underwritten'), "
        "regla OpEx total-first, cobertura de líneas L1 materiales."
    )
    pasted_output = st.text_area(
        "Pegar Notes to Financials generadas por el LLM:",
        value=st.session_state.get('pasted_ai_output', ''),
        height=240,
        key='pasted_ai_output_input',
        placeholder='Notes to Financials:\n\n• **Income:** ...',
    )
    score_col1, score_col2 = st.columns([1, 4])
    with score_col1:
        if st.button("Evaluar", type="primary", width="stretch"):
            st.session_state['pasted_ai_output'] = pasted_output
            st.session_state['ai_score_result'] = score_ai_output(
                pasted_output, results.get('structure')
            )

    score_result = st.session_state.get('ai_score_result')
    if score_result:
        score = score_result['score']
        grade = score_result['grade']
        # Color por grade
        if grade == 'A':
            color = '#16a34a'
        elif grade == 'B':
            color = '#65a30d'
        elif grade == 'C':
            color = '#ca8a04'
        elif grade == 'D':
            color = '#ea580c'
        else:
            color = '#dc2626'

        st.markdown(
            f"<div style='display:flex;align-items:center;gap:24px;padding:16px;"
            f"border-radius:8px;background:{color}15;border-left:4px solid {color};'>"
            f"<div style='font-size:48px;font-weight:700;color:{color};'>{grade}</div>"
            f"<div><div style='font-size:24px;font-weight:600;'>{score} / 100</div>"
            f"<div style='color:#666;font-size:13px;'>"
            f"{sum(1 for c in score_result['checks'] if c['passed'])}/"
            f"{len(score_result['checks'])} checks pasados, "
            f"{len(score_result['penalties'])} penalización(es)"
            f"</div></div></div>",
            unsafe_allow_html=True,
        )

        with st.expander("Ver detalle de checks", expanded=False):
            for c in score_result['checks']:
                icon = "✅" if c['passed'] else "❌"
                st.markdown(f"{icon} **{c['name']}** — {c['detail']}")

        if score_result['penalties']:
            with st.expander(
                f"Ver {len(score_result['penalties'])} penalización(es) detectada(s)",
                expanded=False
            ):
                pen_df = pd.DataFrame(score_result['penalties'])
                st.dataframe(pen_df, width="stretch")
