"""
Inserta (o reemplaza) los catálogos de los 3 nuevos activos en src/asset_catalogs.py
mergeando todos los meses disponibles para máxima cobertura de cuentas.

Si los asset_keys ya existen, los reemplaza in-place.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from scripts.gen_new_catalogs import (
    build_catalog,
    emit_python,
    SECTION_MAP_CALLAN,
    SECTION_MAP_501,
    SECTION_MAP_YALE,
)

DOWNLOADS = Path(r'C:\Users\JaimeValenzuela\Downloads')

TARGETS = [
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


def remove_existing_block(src: str, asset_key: str) -> str:
    """
    Elimina el bloque existente de un asset_key. El bloque empieza en la línea
    `    'asset_key': {` y termina en la línea `    },` que cierra el dict del activo.
    """
    pattern = re.compile(
        r"    '" + re.escape(asset_key) + r"': \{.*?\n    \},\n\n",
        re.DOTALL,
    )
    return pattern.sub("", src, count=1)


def main():
    target_file = ROOT / 'src' / 'asset_catalogs.py'
    src = target_file.read_text(encoding='utf-8')
    original_size = len(src)

    cats = []
    for t in TARGETS:
        cat = build_catalog(t['files'], t['map'], t['plan'])
        cats.append((cat, t['plan']))
        # Si ya existe, eliminar el bloque viejo
        if f"'{cat['asset_key']}': {{" in src:
            src = remove_existing_block(src, cat['asset_key'])
            print(f"  -> reemplazando '{cat['asset_key']}'")

    # Construir bloque a insertar
    blocks = [emit_python(cat['asset_key'], plan, cat['accounts']) for cat, plan in cats]
    new_blocks = "\n".join(blocks).rstrip()

    marker = "    },\n\n}\n\n\ndef get_l1_for_account"
    if marker not in src:
        print("ERROR: marker de cierre no encontrado")
        return 2

    replacement = "    },\n\n" + new_blocks + "\n\n}\n\n\ndef get_l1_for_account"
    new_src = src.replace(marker, replacement)

    target_file.write_text(new_src, encoding='utf-8')
    print(f"OK: tamaño final {len(new_src)} chars (delta vs original: {len(new_src)-original_size:+d})")
    print(f"Activos:")
    for cat, plan in cats:
        print(f"  - {cat['asset_key']} ({plan}, {len(cat['accounts'])} cuentas, {cat['files_processed']} archivos merged)")
    return 0


if __name__ == '__main__':
    sys.exit(main())
