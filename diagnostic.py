# diagnostic.py

"""Script de diagnóstico para validar ingesta de archivos Excel."""

import sys
import os
from pathlib import Path

# Agregar raíz del proyecto al path de forma portátil
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.ingestion import ingest_files

# Leer paths desde variable de entorno o usar defaults relativos
FILES_ENV = os.environ.get("DIAGNOSTIC_FILES", "")
if FILES_ENV:
    files = [Path(p.strip()) for p in FILES_ENV.split(",") if p.strip()]
else:
    # Default: buscar xlsm en ./test_data/ si existe
    test_dir = PROJECT_ROOT / "test_data"
    files = list(test_dir.glob("*.xlsm")) if test_dir.exists() else []

if not files:
    print("No se encontraron archivos. Usar: DIAGNOSTIC_FILES='path1,path2' python diagnostic.py")
    sys.exit(0)

for fpath in files:
    print(f"\n{'='*60}")
    print(f"FILE: {fpath.name}")
    results = ingest_files([Path(fpath)])
    if not results:
        print("  ERROR: No results returned")
        continue
    r = results[0]
    print(f"  Building: {r.building}")
    print(f"  Period:   {r.period_label}")
    print(f"  has_ytd:  {r.metadata.get('has_ytd')}")
    print(f"  Rows in df: {len(r.df)}")

    totals = r.df[r.df['row_type'] == 'total']
    print(f"  Total rows ({len(totals)}):")
    for _, row in totals.iterrows():
        desc = str(row['description'])[:50]
        print(f"    [{row.get('report_line_l1','?')}] {desc}: "
              f"budget={row.get('budget_current','')}, actual={row.get('actual_current','')}")

    sections = r.df[r.df['row_type'] == 'section']
    print(f"  Sections ({len(sections)}):")
    for _, row in sections.iterrows():
        print(f"    {row['description']}")
