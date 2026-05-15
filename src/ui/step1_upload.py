# src/ui/step1_upload.py

"""UI del Paso 1: selección de activo, upload de archivos, validación de período, botón Procesar."""

from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

from src.config import ASSET_CATALOG
from src.pipeline import (
    PipelineInput, run_pipeline, _save_uploaded_files,
    _parse_month_year_from_text, validate_period_coverage,
)


def _preview_file_period(file) -> tuple:
    """
    Lee rápido las primeras filas de un archivo para detectar (mes, año).
    Usa period_label del header si existe, sino cae al filename.
    No ejecuta la ingesta completa — solo abre el workbook y escanea rows 1-10.

    Returns: (month, year, source_str) donde source_str indica de dónde salió.
    """
    fname = getattr(file, 'name', '?')
    # 1) Intentar desde filename (patterns YYYY-MM)
    mo, yr = _parse_month_year_from_text(fname)
    if mo and yr:
        return (mo, yr, f"filename: {fname}")

    # 2) Escanear primeras filas del Excel
    try:
        # Streamlit files: leer bytes y abrir con openpyxl
        import io as _io
        file.seek(0)
        data = file.read()
        file.seek(0)
        wb = openpyxl.load_workbook(_io.BytesIO(data), data_only=True, read_only=True)
        # Saltar hojas tipo 'Comments' / 'Notes' / etc. al buscar fecha — el caso
        # 929 Mass tiene una hoja "Comments" con fechas en filas iniciales que
        # antes confundia la deteccion (la fuente quedaba mal aunque el ingestor
        # despues elegia bien la hoja del IS).
        _SKIP_PREVIEW = ('comment', 'comentario', 'note', 'notas', 'variance note',
                         'explanation', 'cover', 'instruction', 'instructions')
        ordered = sorted(
            wb.sheetnames[:8],
            key=lambda n: any(k in n.lower() for k in _SKIP_PREVIEW),
        )
        for sheet_name in ordered:
            if any(k in sheet_name.lower() for k in _SKIP_PREVIEW):
                continue
            ws = wb[sheet_name]
            for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if i > 10:
                    break
                for cell in (row[:4] if row else ()):
                    if cell is None:
                        continue
                    text = str(cell)
                    mo, yr = _parse_month_year_from_text(text)
                    if mo and yr:
                        wb.close()
                        return (mo, yr, f"sheet '{sheet_name}' row {i}: {text[:40]}")
        wb.close()
    except Exception:
        pass

    return (None, None, "no detectado")


def _suggest_period_from_files(files) -> tuple:
    """
    Dada una lista de archivos, sugiere (quarter, year) basado en los meses
    detectados. Usa el mes más reciente para definir el quarter.
    """
    detections = []
    for f in files:
        mo, yr, _src = _preview_file_period(f)
        if mo and yr:
            detections.append((mo, yr))
    if not detections:
        return (None, None)
    detections.sort(key=lambda t: t[1] * 100 + t[0])
    last_mo, last_yr = detections[-1]
    return ((last_mo - 1) // 3 + 1, last_yr)


class _MiniIng:
    """Mini wrapper que expone period_label + filename para validate_period_coverage."""
    def __init__(self, filename, period_label):
        self.filename = filename
        self.period_label = period_label


def render_step1():
    """
    Paso 1: selección de activo, upload de archivos, VALIDACIÓN DE PERÍODO
    (gate obligatorio antes de Procesar), botón Procesar.
    """
    st.title("Generador de Financials — Reporte Trimestral")
    st.markdown("Carga los archivos del trimestre para generar la sección financiera del reporte.")

    # ── Catálogo de activos / portafolios ──
    asset_options = list(ASSET_CATALOG.keys())
    selected_asset = st.selectbox(
        "Selecciona el activo o portafolio",
        options=asset_options,
        index=0,
        help="Elige un portafolio (consolida múltiples edificios) o un activo individual.",
        key="asset_selector",
    )

    asset_config = ASSET_CATALOG[selected_asset]
    building = selected_asset
    consolidate_portfolio = asset_config["consolidate"]

    if consolidate_portfolio:
        st.info(
            f"**Modo Portafolio:** {asset_config['description']}. "
            f"Carga los archivos de: **{', '.join(asset_config['buildings'])}**. "
            "El programa los agrupará por edificio y consolidará los totales."
        )
    else:
        st.caption(f"📍 {asset_config['description']}")

    st.markdown("### Archivos del trimestre actual")
    files_current = st.file_uploader(
        "Cargar archivos del trimestre (Excel o PDF)",
        type=["xlsm", "xlsx", "pdf"],
        accept_multiple_files=True,
        help="Carga los 3 archivos mensuales del trimestre (ej: Oct, Nov, Dec). "
             "PDF soportado para Walnut Street Wellesley (Campus at Newton Wellesley monthly report).",
        key="files_current",
    )

    with st.expander("📁 Trimestre anterior (opcional - para comparación QoQ)", expanded=False):
        files_prior = st.file_uploader(
            "Cargar archivos del trimestre anterior (Excel o PDF)",
            type=["xlsm", "xlsx", "pdf"],
            accept_multiple_files=True,
            help="Para generar comparación trimestre vs trimestre anterior",
            key="files_prior",
        )

    with st.expander("📈 Datos de mercado (opcional - análisis de mercado)", expanded=False):
        st.markdown(
            "Carga un archivo tipo **CoStar Multifamily Data Grid** (.xlsx) "
            "con datos trimestrales del mercado donde está ubicado el activo."
        )
        market_col1, market_col2 = st.columns([2, 1])
        with market_col1:
            market_file = st.file_uploader(
                "Archivo de datos de mercado",
                type=["xlsx", "xlsm"],
                accept_multiple_files=False,
                key="market_file",
            )
        with market_col2:
            market_name = st.text_input(
                "Nombre del mercado",
                value="",
                placeholder="Ej: Oakland, San Francisco, Austin...",
                key="market_name",
                help="Nombre del mercado/submarket para etiquetar gráficas y KPIs",
            )

    with st.expander("📂 Master Excel (opcional - budget tracking histórico)", expanded=False):
        st.markdown(
            "Carga el **Master Excel** existente para que el programa agregue "
            "el nuevo trimestre a las columnas históricas. Si no cargas uno, "
            "se creará un master nuevo."
        )
        master_upload = st.file_uploader(
            "Archivo Master Excel existente",
            type=["xlsx"],
            accept_multiple_files=False,
            key="master_upload",
            help="El Master Excel de períodos anteriores. Se le agregarán las columnas del nuevo trimestre.",
        )

    st.markdown("---")

    # ═══════════════════════════════════════════════════════════════
    # GATE DE VALIDACIÓN DE PERÍODO (obligatorio)
    # ═══════════════════════════════════════════════════════════════
    st.markdown("### 📅 Validación del período")

    period_confirmed = False
    period_str = ""
    period_ok_to_proceed = False

    if not files_current:
        st.info("Cargá al menos un archivo arriba para validar el período.")
    else:
        # Preview: detectar mes/año de cada archivo
        preview_rows = []
        detected_for_suggestion = []
        for f in files_current:
            mo, yr, src = _preview_file_period(f)
            preview_rows.append({
                'Archivo': f.name,
                'Mes detectado': mo if mo else '—',
                'Año detectado': yr if yr else '—',
                'Fuente': src,
            })
            if mo and yr:
                detected_for_suggestion.append((mo, yr))

        preview_df = pd.DataFrame(preview_rows)
        st.caption("Detección automática del período cubierto por cada archivo:")
        st.dataframe(preview_df, width="stretch", hide_index=True)

        # Sugerencia de quarter + year
        sug_q, sug_y = _suggest_period_from_files(files_current)

        pcol1, pcol2 = st.columns([1, 1])
        with pcol1:
            qopts = ['Q1', 'Q2', 'Q3', 'Q4']
            default_q_idx = (sug_q - 1) if sug_q else 0
            chosen_q = st.selectbox(
                "Trimestre",
                options=qopts,
                index=default_q_idx,
                key="period_quarter",
                help="Selecciona el trimestre que cubren los archivos cargados.",
            )
        with pcol2:
            default_y = sug_y if sug_y else 2025
            chosen_y = st.number_input(
                "Año",
                min_value=2015, max_value=2099,
                value=int(default_y),
                step=1, key="period_year",
                help="Año del trimestre reportado.",
            )

        period_str = f"{chosen_q} {int(chosen_y)}"

        if sug_q and sug_y:
            sug_str = f"Q{sug_q} {sug_y}"
            if period_str == sug_str:
                st.caption(f"✅ Coincide con la detección automática: **{sug_str}**")
            else:
                st.caption(f"ℹ️ Detección automática sugería: **{sug_str}** (vos elegiste {period_str})")

        # Validar coherencia con los archivos cargados
        mini_ings = []
        for f, row in zip(files_current, preview_rows):
            mo, yr = row['Mes detectado'], row['Año detectado']
            plabel = f"{yr}-{mo:02d}" if isinstance(mo, int) and isinstance(yr, int) else ""
            mini_ings.append(_MiniIng(filename=f.name, period_label=plabel))

        validation = validate_period_coverage(period_str, mini_ings)

        for err in validation['errors']:
            st.error(f"❌ {err}")
        for warn in validation['warnings']:
            st.warning(f"⚠️ {warn}")
        if validation['ok'] and not validation['warnings']:
            st.success(
                f"✅ Los {len(files_current)} archivos cubren correctamente **{period_str}** "
                f"(meses {validation['expected_months']})."
            )

        # Checkbox de confirmación obligatorio
        period_confirmed = st.checkbox(
            f"✅ Confirmo que el período correcto es **{period_str}** y los archivos corresponden a esos meses.",
            key="period_confirmed",
            value=False,
        )
        period_ok_to_proceed = period_confirmed and validation['ok']

    st.markdown("---")

    # ── Botón Procesar (gated por validación de período) ──
    if st.button(
        "▶️ Procesar",
        type="primary",
        width="stretch",
        disabled=not period_ok_to_proceed,
        help=("Primero carga los archivos y confirmá el período arriba."
              if not period_ok_to_proceed else None),
    ):
        if not files_current:
            st.error("⚠️ Carga al menos un archivo Excel del trimestre actual.")
        elif not period_ok_to_proceed:
            st.error("⚠️ Confirmá el período arriba antes de procesar.")
        else:
            with st.spinner("Procesando archivos..."):
                try:
                    master_path_tmp = None
                    if master_upload:
                        tmp_paths = _save_uploaded_files([master_upload], "master_base")
                        master_path_tmp = tmp_paths[0] if tmp_paths else None

                    inputs = PipelineInput(
                        files_current=files_current,
                        files_prior=files_prior,
                        building=building,
                        period=period_str,  # período CONFIRMADO por el analista
                        market_file=market_file,
                        market_name=market_name,
                        master_upload_path=master_path_tmp,
                        consolidate=consolidate_portfolio,
                    )
                    results = run_pipeline(inputs)

                    for w in results.get('warnings', []):
                        st.warning(f"⚠️ {w}")

                    st.session_state.results = results
                    st.session_state.step = 2
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error procesando archivos: {str(e)}")
                    st.exception(e)
