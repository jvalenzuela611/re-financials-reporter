"""
Generador de catálogo para Blockyard (Bridge at the Blockyard).

Lee el archivo Blockyard 12-month directamente (sin pasar por la ingesta,
que aún no tiene parser CWS) y emite el snippet Python para insertar en
src/asset_catalogs.py.

Layout del archivo:
- Row 1 col A: 'Bridge at the Blockyard (a0111048)'
- Row 4 col A: 'Book = Accrual ; Tree = cws_res_is_'
- Row 5 cols C-N: meses Jan-Dec, col O = 'Total'
- Row 6+: data rows con código GL en col A y descripción en col B
- Section/total rows: col A vacía, col B con texto
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import openpyxl  # noqa
from scripts.gen_new_catalogs import SECTION_MAP_BLOCKYARD, emit_python


BLOCKYARD_FILE = Path(r'C:\Users\JaimeValenzuela\Downloads\Blockyard Revised Budget 2026.xlsm')


def build_blockyard_catalog():
    wb = openpyxl.load_workbook(str(BLOCKYARD_FILE), data_only=True)
    ws = wb['Report1']

    accounts = {}
    current_l1, current_l2 = None, None
    unmapped_sections = set()
    data_rows = 0

    for r in range(6, ws.max_row + 1):
        c1 = ws.cell(row=r, column=1).value
        c2 = ws.cell(row=r, column=2).value

        # Section/total: col A vacía, col B con texto
        if not c1 and c2:
            label = str(c2).strip().upper()
            # Ignorar totales (no se mapean — solo cambian el contexto)
            if label.startswith('TOTAL ') or label.startswith('NET ') or label.startswith('INCOME (LOSS)'):
                continue
            mapping = SECTION_MAP_BLOCKYARD.get(label)
            if mapping:
                current_l1, current_l2 = mapping
            else:
                unmapped_sections.add(label)
            continue

        # Data row: col A = código, col B = descripción
        if c1 and c2 and re.match(r'^\d{4}-\d{4}$', str(c1).strip()):
            code = str(c1).strip()
            desc = str(c2).strip()
            data_rows += 1
            if code in accounts:
                continue
            if current_l1 is None:
                continue
            accounts[code] = {
                'desc': desc,
                'l1': current_l1,
                'l2': current_l2,
            }

    wb.close()
    return {
        'asset_key': 'Bridge at the Blockyard',
        'plan': 'CWS',
        'accounts': accounts,
        'unmapped_sections': unmapped_sections,
        'data_rows': data_rows,
    }


def main():
    cat = build_blockyard_catalog()
    print(f'asset_key:        {cat["asset_key"]}')
    print(f'plan:             {cat["plan"]}')
    print(f'data rows:        {cat["data_rows"]}')
    print(f'cuentas mapeadas: {len(cat["accounts"])}')
    if cat['unmapped_sections']:
        print(f'WARN secciones sin mapping: {cat["unmapped_sections"]}')
    print()
    print('=' * 60)
    print('# CODIGO PYTHON PARA ASSET_CATALOGS')
    print('=' * 60)
    print()
    print(emit_python(cat['asset_key'], cat['plan'], cat['accounts']))


if __name__ == '__main__':
    main()
