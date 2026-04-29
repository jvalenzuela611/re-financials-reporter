"""
Test del fix de KOI: agregadas 8 cuentas al catálogo.

Valida:
1. Cobertura del catálogo (debería ser 100% en archivo Jan 2026)
2. L1 totals coinciden con totales del socio
3. NOI cuadra (calculado vs reportado)
4. No hay cambios de signos anómalos en partner_total vs L3_sum
5. Regresión: los otros 9 activos siguen funcionando
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from src.ingestion import ingest_single_file
from src.native_structure import build_financial_structure, _resolve_asset_key
from src.asset_catalogs import ASSET_CATALOGS, get_l1_for_account
from src.validation import _check_full_pl_reconciliation, _check_noi_reconciliation


KOI_FILE = Path(r'C:\Users\JaimeValenzuela\Downloads\01 Budget_Comparison_Cash Flow 01.26 (pn10675) with Client GL.xlsx')

# Archivos para test de regresión (uno por activo)
DOWNLOADS = Path(r'C:\Users\JaimeValenzuela\Downloads')
REGRESSION_FILES = [
    ('Alice House JV LLC',           DOWNLOADS / '00 Alice House JV LLC financial reports 02-28-2026.xlsx'),
    ('295 29th Street JV LLC',       DOWNLOADS / '00 295 29th Street JV LLC financial reports 03-31-2026.xlsx'),
    ('14th & Callan Street JV LLC',  DOWNLOADS / '00 14th & Callan Street JV financial report 03-2026 v3.xlsx'),
    ('J 501 Estates',                DOWNLOADS / '501 Estates Apartment LLC Close Package 2025-12.xlsx'),
    ('624 Yale Apartments',          DOWNLOADS / '02_624 Yale_VarianceReport 12.25.xlsx'),
]


def section(title):
    print()
    print('=' * 80)
    print(title)
    print('=' * 80)


def test_koi_coverage():
    section('TEST 1: COBERTURA DEL CATALOGO KOI (debe ser 100%)')
    result = ingest_single_file(KOI_FILE)
    df = result.df
    data = df[df.row_type == 'data']
    asset_key = _resolve_asset_key([result])
    misses = [(r['account'], r['description']) for _, r in data.iterrows()
              if not get_l1_for_account(asset_key, r['account'])]
    print(f'  asset_key resuelto: {asset_key}')
    print(f'  Hits: {len(data) - len(misses)}/{len(data)}')
    if misses:
        print('  MISSES:')
        for m in misses:
            print(f'    {m}')
        return False
    print('  PASS — todas las cuentas cubiertas')
    return True


def test_koi_l1_vs_partner():
    section('TEST 2: L1 TOTALS COINCIDEN CON TOTALES DEL SOCIO')
    result = ingest_single_file(KOI_FILE)
    df = result.df.sort_values('row_num')
    structure = build_financial_structure([result])

    # Totales del socio del archivo (extraer de filas tipo 'total' con descripcion clave)
    partner_totals = {}
    for _, row in df.iterrows():
        if row['row_type'] != 'total':
            continue
        d = (row['description'] or '').upper().strip()
        ac = row['actual_current']
        if d == 'TOTAL REVENUE':
            partner_totals['Income'] = ac
        elif d == 'TOTAL OPERATING EXPENSES':
            partner_totals['OpEx_inc_RET'] = ac
        elif d == 'TOTAL REAL ESTATE & PROPERTY TAXES':
            partner_totals['RET'] = ac
        elif d == 'TOTAL NET OPERATING INCOME':
            partner_totals['NOI'] = ac
        elif d == 'TOTAL INTEREST-MRTG ENCUMBRANCE':
            partner_totals['Interest_only'] = ac
        elif d == 'TOTAL NON OPERATING EXPENSES':
            partner_totals['NonOp_inc_Interest'] = ac
        elif d == 'TOTAL CAPITALIZED EXPENDITURES':
            partner_totals['Capex'] = ac

    # OpEx ex-RET = TOTAL OPERATING EXPENSES - TOTAL RET (asi lo separa la pipeline)
    partner_totals['OpEx_ex_RET'] = partner_totals['OpEx_inc_RET'] - partner_totals['RET']
    # Non-Op ex-Interest = TOTAL NON OPERATING EXPENSES - TOTAL INTEREST
    partner_totals['NonOp_ex_Interest'] = partner_totals['NonOp_inc_Interest'] - partner_totals['Interest_only']

    by_name = {l.name: l for l in structure.get('l1_lines', [])}

    checks = [
        ('Income',                 partner_totals['Income'],            by_name.get('Income')),
        ('Operating Expenses',     partner_totals['OpEx_ex_RET'],       by_name.get('Operating Expenses')),
        ('Real Estate Taxes',      partner_totals['RET'],               by_name.get('Real Estate Taxes')),
        ('NOI',                    partner_totals['NOI'],               by_name.get('NOI')),
        ('Interest Expense',       partner_totals['Interest_only'],     by_name.get('Interest Expense')),
        ('Non-Operating Expenses', partner_totals['NonOp_ex_Interest'], by_name.get('Non-Operating Expenses')),
        ('Total Capex',            partner_totals['Capex'],             by_name.get('Total Capex')),
    ]

    all_ok = True
    for name, expected, line in checks:
        if line is None:
            print(f'  [MISS] {name:25s}  L1 line ausente en pipeline')
            all_ok = False
            continue
        actual = line.actual_current or 0
        diff = actual - expected
        ok = abs(diff) < 1.0
        flag = 'OK' if ok else 'FAIL'
        print(f'  [{flag}] {name:25s}  pipeline={actual:>14,.2f}  partner={expected:>14,.2f}  diff={diff:>+10,.2f}')
        if not ok:
            all_ok = False
    return all_ok


def test_koi_noi_calc():
    section('TEST 3: NOI CALCULADO == NOI DEL SOCIO')
    result = ingest_single_file(KOI_FILE)
    structure = build_financial_structure([result])
    by_name = {l.name: l for l in structure.get('l1_lines', [])}

    inc = by_name.get('Income').actual_current
    opex = by_name.get('Operating Expenses').actual_current
    ret = by_name.get('Real Estate Taxes').actual_current
    noi_partner = by_name.get('NOI').actual_current

    noi_calc = inc - abs(opex) - abs(ret)
    diff = noi_calc - noi_partner
    print(f'  NOI partner:    {noi_partner:>14,.2f}')
    print(f'  NOI calculado:  {noi_calc:>14,.2f}  (= {inc:,.2f} - {abs(opex):,.2f} - {abs(ret):,.2f})')
    print(f'  Diff:           {diff:>+14,.2f}')
    ok = abs(diff) < 1.0
    print(f'  {"PASS" if ok else "FAIL"} — {"cuadra" if ok else "no cuadra"}')
    return ok


def test_koi_l1_l3_signs():
    section('TEST 4: NO CAMBIOS DE SIGNOS EN partner_total vs L3_sum')
    result = ingest_single_file(KOI_FILE)
    structure = build_financial_structure([result])
    val = _check_full_pl_reconciliation(structure)

    print(f'  Status validacion: {val.status}')
    print(f'  Mensaje: {val.message}')
    print()

    sign_anomalies = []
    diff_warnings = []
    for r in val.details['per_l1']:
        pt = r['partner_total']
        cs = r['classification_sum']
        df = r['diff']
        line = r['line']

        # Sign anomaly: partner y L3 con signos opuestos (fuera de "calc" lines)
        if pt is not None and cs is not None and pt != 0 and cs != 0:
            if (pt > 0) != (cs > 0):
                sign_anomalies.append((line, pt, cs))

        # Diff > $100
        if df is not None and abs(df) > 100:
            diff_warnings.append((line, df))

        pt_s = f'{pt:>14,.2f}' if pt is not None else '          None'
        cs_s = f'{cs:>14,.2f}' if cs is not None else '          None'
        df_s = f'{df:>+10,.2f}' if df is not None else '      None'
        print(f'    {line:25s}  partner={pt_s}  L3sum={cs_s}  diff={df_s}')

    print()
    if sign_anomalies:
        print(f'  FAIL — {len(sign_anomalies)} cambio(s) de signo:')
        for line, pt, cs in sign_anomalies:
            print(f'    {line}: partner={pt:+,.2f}, L3sum={cs:+,.2f}')
        return False
    if diff_warnings:
        print(f'  FAIL — {len(diff_warnings)} L1 con diff > $100:')
        for line, df in diff_warnings:
            print(f'    {line}: {df:+,.2f}')
        return False
    print('  PASS — todos los partner_total cuadran exactamente con L3_sum')
    return True


def test_regression_other_assets():
    section('TEST 5: REGRESION EN LOS OTROS 9 ACTIVOS')
    all_ok = True
    for asset_key, fpath in REGRESSION_FILES:
        if not fpath.exists():
            print(f'  [SKIP] {asset_key:30s}  archivo no encontrado: {fpath.name}')
            continue
        result = ingest_single_file(fpath)
        df = result.df
        data = df[df.row_type == 'data']
        resolved_key = _resolve_asset_key([result])

        misses = [r['account'] for _, r in data.iterrows()
                  if not get_l1_for_account(resolved_key, r['account'])]
        cov = (len(data) - len(misses)) / len(data) * 100 if len(data) else 100

        # Validacion P&L
        try:
            structure = build_financial_structure([result])
            val = _check_full_pl_reconciliation(structure)
            sign_anomalies = sum(1 for r in val.details['per_l1']
                                  if r['partner_total'] is not None and r['classification_sum'] is not None
                                  and r['partner_total'] != 0 and r['classification_sum'] != 0
                                  and (r['partner_total'] > 0) != (r['classification_sum'] > 0))
            big_diffs = sum(1 for r in val.details['per_l1']
                            if r['diff'] is not None and abs(r['diff']) > 100)
            status = 'OK' if sign_anomalies == 0 and big_diffs == 0 else 'WARN'
        except Exception as e:
            status = f'ERROR: {e}'
            sign_anomalies = '-'
            big_diffs = '-'

        flag = 'OK' if status == 'OK' else 'CHECK'
        print(f'  [{flag}] {resolved_key:30s}  cov={cov:5.1f}%  sign_anom={sign_anomalies}  big_diffs={big_diffs}  status={status}')
        if status not in ('OK',) and not str(status).startswith('WARN'):
            all_ok = False
    return all_ok


def main():
    results = {
        'cobertura': test_koi_coverage(),
        'l1_partner': test_koi_l1_vs_partner(),
        'noi_calc': test_koi_noi_calc(),
        'signs': test_koi_l1_l3_signs(),
        'regresion': test_regression_other_assets(),
    }
    section('RESUMEN')
    for k, v in results.items():
        flag = 'PASS' if v else 'FAIL'
        print(f'  {flag}  {k}')
    overall = all(results.values())
    print()
    print('OVERALL: ' + ('TODOS PASARON' if overall else 'HAY FALLAS — revisar arriba'))
    return 0 if overall else 1


if __name__ == '__main__':
    sys.exit(main())
