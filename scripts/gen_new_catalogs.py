"""
Generador de catálogos de cuentas para activos nuevos.

Imprime el código Python de los entries de ASSET_CATALOGS para que se inserten
en src/asset_catalogs.py. Usa el contexto de sección de las hojas ingeridas
para asignar L1/L2 a cada cuenta.

Activos cubiertos:
- 14th & Callan Street JV LLC (formato Callan / Sares Regis)
- J 501 Estates (formato BCR)
- 624 Yale Apartments (formato YardiVariance)
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from src.ingestion import ingest_single_file


# Mapping de nombre de sección → (L1, L2). Conserva las convenciones de naming
# usadas por el catálogo Alice (TMG) para que rollups y formato sean consistentes.
SECTION_MAP_CALLAN = {
    'Rental Income': ('Income', 'Rental Income'),
    'Concessions and Vacancy': ('Income', 'Concessions and Vacancy'),
    'Retail Income': ('Income', 'Other Income'),
    'Other Income': ('Income', 'Other Income'),
    'Payroll and Related': ('Operating Expenses', 'Payroll and Related'),
    'Utilities': ('Operating Expenses', 'Utilities'),
    'Turnover': ('Operating Expenses', 'Turnover'),
    'Repairs and Maintenance': ('Operating Expenses', 'Repairs and Maintenance'),
    'Contract Services': ('Operating Expenses', 'Contract Services'),
    'Insurance': ('Operating Expenses', 'Insurance'),
    'Advertising and Marketing': ('Operating Expenses', 'Advertising and Marketing'),
    'Management Fees': ('Operating Expenses', 'Management Fees'),
    'Property G & A': ('Operating Expenses', 'Property G & A'),
    'Property Taxes': ('Real Estate Taxes', 'Property Taxes'),
    'Company General & Administrative': ('Non-Operating Expenses', 'Company General & Administrative'),
    'Unrealized Gain/Loss': ('Non-Operating Expenses', 'Unrealized Gain/Loss'),
    'Interest and Financing Expenses': ('Interest Expense', 'Interest and Financing Expenses'),
}

SECTION_MAP_501 = {
    'Rental Income': ('Income', 'Rental Income'),
    'Other Rental Income': ('Income', 'Other Income'),
    'Other Income': ('Income', 'Other Income'),
    'Payroll & Benefits': ('Operating Expenses', 'Payroll and Related'),
    'General Maintenance': ('Operating Expenses', 'Repairs and Maintenance'),
    'Repairs & Maintenance': ('Operating Expenses', 'Repairs and Maintenance'),
    'Make - Ready/Redecorating': ('Operating Expenses', 'Turnover'),
    'Recreational Amenities': ('Operating Expenses', 'Repairs and Maintenance'),
    'Contract Services': ('Operating Expenses', 'Contract Services'),
    'Advertising /Marketing/Promotions': ('Operating Expenses', 'Advertising and Marketing'),
    'General & Administrative': ('Operating Expenses', 'Property G & A'),
    'Utilities': ('Operating Expenses', 'Utilities'),
    'Management Fees': ('Operating Expenses', 'Management Fees'),
    'Insurance': ('Operating Expenses', 'Insurance'),
    'Taxes': ('Real Estate Taxes', 'Property Taxes'),
    'Partnership/Owner Expenses': ('Non-Operating Expenses', 'Company General & Administrative'),
    'Debt Service': ('Interest Expense', 'Interest and Financing Expenses'),
    'Depreciation & Amortization': ('Non-Operating Expenses', 'Depreciation/Amortization'),
    'Construction in Progress-Capital/Renovation': ('Total Capex', 'CapEx'),
    'Construction in Progress-Routine Replacement': ('Total Capex', 'CapEx'),
}

SECTION_MAP_BLOCKYARD = {
    # Income
    'RENTAL INCOME':                          ('Income', 'Rental Income'),
    'OTHER INCOME':                           ('Income', 'Other Income'),
    # Operating Expenses
    'PAYROLL & BENEFITS':                     ('Operating Expenses', 'Payroll and Related'),
    'MARKETING & ADVERTISING':                ('Operating Expenses', 'Advertising and Marketing'),
    'TURNOVER COSTS':                         ('Operating Expenses', 'Turnover'),
    'REPAIRS & MAINTENANCE':                  ('Operating Expenses', 'Repairs and Maintenance'),
    'PROFESSIONAL/CONTRACT SERVICES':         ('Operating Expenses', 'Contract Services'),
    'GENERAL & ADMINISTRATIVE EXPENSES':      ('Operating Expenses', 'Property G & A'),
    'UTILITIES':                              ('Operating Expenses', 'Utilities'),
    'INSURANCE':                              ('Operating Expenses', 'Insurance'),
    'MANAGEMENT FEES':                        ('Operating Expenses', 'Management Fees'),
    # RET (separado L1)
    'TAXES':                                  ('Real Estate Taxes', 'Property Taxes'),
    # Below NOI
    'DEBT SERVICE':                           ('Interest Expense', 'Interest and Financing Expenses'),
    'OTHER NON-OPERATING COSTS':              ('Non-Operating Expenses', 'Company General & Administrative'),
    'DEPRECIATION & AMORTIZATION':            ('Non-Operating Expenses', 'Depreciation/Amortization'),
    # Capex
    'INTERIOR IMPROVEMENTS':                  ('Total Capex', 'CapEx'),
    'EXTERIOR IMPROVEMENTS':                  ('Total Capex', 'CapEx'),
    'INTERIOR IMPROVEMENTS-UPFRONT':          ('Total Capex', 'CapEx'),
    'FURNITURE/FIXTURES/EQUIPMENT-UPFRONT':   ('Total Capex', 'CapEx'),
    'OTHER INTERIOR IMRPOVEMENTS-UPFRONT':    ('Total Capex', 'CapEx'),  # typo en el archivo
    'EXTERIOR IMPROVEMENTS-UPFRONT':          ('Total Capex', 'CapEx'),
    'BUILDINGS & IMPROVEMENTS-UPFRONT':       ('Total Capex', 'CapEx'),
    'OTHER EXTERIOR IMPROVEMENTS-UPFRONT':    ('Total Capex', 'CapEx'),
    'CAPITAL EXPENDITURES':                   ('Total Capex', 'CapEx'),
}


SECTION_MAP_YALE = {
    'Rental Revenue - Residential': ('Income', 'Rental Income'),
    'Rental Revenue Adjustments - Residential': ('Income', 'Concessions and Vacancy'),
    'Other Revenue': ('Income', 'Other Income'),
    'Payroll': ('Operating Expenses', 'Payroll and Related'),
    'Marketing': ('Operating Expenses', 'Advertising and Marketing'),
    'Repairs and Maintenance': ('Operating Expenses', 'Repairs and Maintenance'),
    'General Repairs and Maintenance': ('Operating Expenses', 'Repairs and Maintenance'),
    'Landscaping': ('Operating Expenses', 'Repairs and Maintenance'),
    'Fire Life and Safety': ('Operating Expenses', 'Repairs and Maintenance'),
    'Unit Turn Expense': ('Operating Expenses', 'Turnover'),
    'Utilities': ('Operating Expenses', 'Utilities'),
    'Management Fees': ('Operating Expenses', 'Management Fees'),
    'General and Administration': ('Operating Expenses', 'Property G & A'),
    'Insurance Coverage': ('Operating Expenses', 'Insurance'),
    'Taxes': ('Real Estate Taxes', 'Property Taxes'),
    'Debt Servicing Costs': ('Interest Expense', 'Interest and Financing Expenses'),
    'Ownership / Partnership Expenses': ('Non-Operating Expenses', 'Company General & Administrative'),
    'Capital Expenditures': ('Total Capex', 'CapEx'),
    'Investment Committee (IC) Capital Expenditures': ('Total Capex', 'CapEx'),
}


def build_catalog(filepaths, section_map: dict, plan: str) -> dict:
    """
    Ingiere uno o varios archivos del mismo activo, recorre filas en orden, y
    genera el dict de cuentas asignando L1/L2 según la sección activa.

    Si una cuenta aparece en varios archivos, se conserva la PRIMERA ocurrencia
    encontrada — los archivos posteriores solo agregan cuentas nuevas. Esto evita
    pisar un L1/L2 ya validado con uno potencialmente menos preciso.

    Args:
        filepaths: Path único o lista de Paths del mismo activo.
        section_map: dict de mapeo sección → (L1, L2).
        plan: nombre del plan de cuentas.
    """
    if isinstance(filepaths, (str, Path)):
        filepaths = [Path(filepaths)]
    else:
        filepaths = [Path(f) for f in filepaths]

    accounts = {}
    unmapped_sections = set()
    asset_key = None
    total_data_rows = 0

    for fp in filepaths:
        result = ingest_single_file(fp)
        if asset_key is None:
            asset_key = result.building
        df = result.df.sort_values('row_num').reset_index(drop=True)
        total_data_rows += int((df['row_type'] == 'data').sum())

        current_l1, current_l2 = None, None
        for _, row in df.iterrows():
            rt = row['row_type']
            desc = (row['description'] or '').strip()
            if rt == 'section':
                mapping = section_map.get(desc)
                if mapping:
                    current_l1, current_l2 = mapping
                else:
                    unmapped_sections.add(desc)
            elif rt == 'data':
                acct = (row['account'] or '').strip()
                if not acct or acct in accounts:
                    continue
                if current_l1 is None:
                    continue
                accounts[acct] = {
                    'desc': desc,
                    'l1': current_l1,
                    'l2': current_l2,
                }

    return {
        'asset_key': asset_key,
        'plan': plan,
        'accounts': accounts,
        'unmapped_sections': unmapped_sections,
        'data_rows_in_file': total_data_rows,
        'files_processed': len(filepaths),
    }


def emit_python(asset_key: str, plan: str, accounts: dict) -> str:
    """Genera el snippet Python de un entry de ASSET_CATALOGS."""
    lines = [f"    '{asset_key}': {{"]
    lines.append(f"        \"plan\": '{plan}',")
    lines.append(f"        \"open_list\": True,")
    lines.append(f"        \"accounts\": {{")
    for code, info in accounts.items():
        desc = info['desc'].replace("'", "\\'")
        l1 = info['l1']
        l2 = info['l2']
        lines.append(
            f"            '{code}': {{\"desc\": '{desc}', "
            f"\"l1\": '{l1}', \"l2\": '{l2}'}},"
        )
    lines.append(f"        }},")
    lines.append(f"    }},")
    return "\n".join(lines)


def main():
    DOWNLOADS = Path(r'C:\Users\JaimeValenzuela\Downloads')
    targets = [
        {
            'tag': 'Callan',
            'files': [
                DOWNLOADS / '00 14th & Callan Street JV financial report 01-2026.xlsx',
                DOWNLOADS / '00 14th & Callan Street JV financial report 02-2026.xlsx',
                DOWNLOADS / '00 14th & Callan Street JV financial report 03-2026 v3.xlsx',
            ],
            'map': SECTION_MAP_CALLAN,
            'plan': 'SaresRegis',
        },
        {
            'tag': '501 Estates',
            'files': [
                DOWNLOADS / '501 Estates Apartment LLC Close Package 2025-10.xlsx',
                DOWNLOADS / '501 Estates Apartment LLC Close Package 2025-11.xlsx',
                DOWNLOADS / '501 Estates Apartment LLC Close Package 2025-12.xlsx',
            ],
            'map': SECTION_MAP_501,
            'plan': 'BCR',
        },
        {
            'tag': '624 Yale',
            'files': [
                DOWNLOADS / '02_624 Yale_VarianceReport 10.25.xlsx',
                DOWNLOADS / '02_624 Yale_VarianceReport 11.25.xlsx',
                DOWNLOADS / '02_624 Yale_VarianceReport 12.25.xlsx',
            ],
            'map': SECTION_MAP_YALE,
            'plan': 'YardiVariance',
        },
    ]

    cats = []
    for t in targets:
        print(f"\n{'='*60}\n{t['tag']}\n{'='*60}")
        cat = build_catalog(t['files'], t['map'], t['plan'])
        cats.append((t, cat))
        print(f"  asset_key: {cat['asset_key']}")
        print(f"  files merged: {cat['files_processed']}")
        print(f"  data rows total: {cat['data_rows_in_file']}")
        print(f"  accounts mapeadas (unicas): {len(cat['accounts'])}")
        if cat['unmapped_sections']:
            print(f"  WARN secciones sin mapping: {cat['unmapped_sections']}")

    print(f"\n\n{'#'*60}\n# CODIGO PYTHON PARA ASSET_CATALOGS\n{'#'*60}\n")
    for t, cat in cats:
        print(emit_python(cat['asset_key'], cat['plan'], cat['accounts']))
        print()


if __name__ == '__main__':
    main()
