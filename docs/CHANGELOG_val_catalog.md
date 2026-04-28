# Changelog — The Val catalog

Rastreo de cambios al catálogo de cuentas de **The Val** (plan HDFC/CMC).

---

## 2026-04-22 (tarde) — Fix #2: cuentas 1xxx Capex mal clasificadas como Interest

### Contexto

Florencia detectó en Step 2 (panel de clasificación) que varias cuentas `1xxx`
(capital / balance sheet) del Val aparecían con `L1 Auto = Interest Expense`:

> *"en The Val varias cuentas de Capex las toma como Interest Expense"*

Cuentas concretas del screenshot:
- `1420-000 BUILDINGS`   (actual $-42,807, budget $0)
- `1464-000 FURNITURE & FIXTURES`   (actual $-29,400)
- `1315-009 CAPITAL EXP OWNER IMP (1318_1319)`   (actual $26,779)
- `1470-000 APPLIANCES`   (actual $-15,838)
- `1315-000 SPECIAL ESCROW`
- (solo `6820-000 INTEREST` con $732,549 es interés real)

### Diagnóstico

Las cuentas `1xxx` viven bajo **"Cash Flow Conversion"** en el Val Budget.xlsx.
El extractor las skippea via `NON_PNL_SECTIONS` ("cash flow conversion"), así
que NO quedan en el catálogo del Val.

Pero en los **reportes mensuales** del Val (formato CMC), esas mismas cuentas
sí se ingieren — y como no están en el catálogo, el pipeline cae al fallback
de herencia de sección. La sección anterior fue "Financial" (donde viven los
6820-xxx de interés real) → las `1xxx` heredaron Interest Expense.

### Fix aplicado

Inyecté **10 entradas sintéticas** en el catálogo del Val con L1=Total Capex.
Se agregan en [scripts/extract_catalogs.py](../scripts/extract_catalogs.py)
en el dict `val_extra_capex`, que se mergea al catálogo después del
`build_catalog` normal. L2 = `"Cash Flow Conversion"` (sección real del partner).

| Account code | Descripción | L1 antes (wrong) | L1 después (correct) |
|---|---|---|---|
| `1115-000` | ESCROW | Interest Expense (inherit) | **Total Capex** |
| `1315-000` | SPECIAL ESCROW | Interest Expense (inherit) | **Total Capex** |
| `1315-009` | CAPITAL EXP OWNER IMP (1318_1319) | Interest Expense (inherit) | **Total Capex** |
| `1318-000` | OWNER IMPROVEMENTS | Interest Expense (inherit) | **Total Capex** |
| `1319-000` | CAPITAL IMPROVEMENTS | Interest Expense (inherit) | **Total Capex** |
| `1321-000` | CY-REPLACEMENT RESERVE | Interest Expense (inherit) | **Total Capex** |
| `1322-000` | REPLACEMENT RESERVE REFUND | Interest Expense (inherit) | **Total Capex** |
| `1420-000` | BUILDINGS | Interest Expense (inherit) | **Total Capex** |
| `1464-000` | FURNITURE & FIXTURES | Interest Expense (inherit) | **Total Capex** |
| `1470-000` | APPLIANCES | Interest Expense (inherit) | **Total Capex** |

### Impacto en distribuciones L1 del Val

| L1 | Fix #1 (después del 1er pase) | Fix #2 (post-1xxx) | Δ |
|---|---:|---:|---:|
| Income | 37 | 37 | 0 |
| Operating Expenses | 106 | 106 | 0 |
| Real Estate Taxes | 4 | 4 | 0 |
| Interest Expense | 12 | 12 | 0 |
| Non-Operating Expenses | 6 | 6 | 0 |
| Total Capex | 9 | **19** | **+10** |
| **Total** | **174** | **184** | **+10** |

Las 12 Interest accounts son ahora solo los **intereses reales** del Val
(6820-000..6820-400 + 6830-000 Swap). Confirmado contra el screenshot de
Florencia: `6820-000 INTEREST $732,549` es la única línea de interés real,
las demás eran ruido de herencia.

### Regresión verificada
Alice (249) / Edson (485) / Walnut (59) / Wilcox (150) — distribuciones L1
idénticas. Total en el sistema: 1,127 cuentas (antes 1,117 → +10 del Val).

### Nota sobre cuentas 2xxx

Las cuentas `2322-001 PRINCIPAL PAYMENTS` y `2322-002 2ND PRINCIPAL PAYMENTS`
también viven en "Cash Flow Conversion" y podrían aparecer en reportes
mensuales. Son **repagos de deuda** (liability, no expense) — si aparecen
en un backtest futuro con clasificación errónea, se agregan con L1=NA
(excluido del NI) o L1=Total Capex según preferencia de Florencia.

---

## 2026-04-22 (mañana) — Fix #1 post-backtest Q4 2025 (feedback Florencia)

### Contexto

Florencia revisó el output del AI sobre Q4 2025 de The Val y detectó que
**13 cuentas estaban mal clasificadas en el catálogo**, distorsionando los
totales L1 (Real Estate Taxes, Interest Expense) que alimentan al AI. Las
dos menciones concretas en su feedback:

- *"~$47k in emergency maintenance (replacements/redecorating) — Correcto
  como comentario, pero no debería estar en OpEx (es NA/CapEx)"*
- *"Real Estate Taxes were $57,963 unfavorable to budget — Incorrecto;
  está incluyendo replacements"*
- *"Interest Expense was $61,814 favorable in Q4 — Incorrecto; incluye
  cuentas mal clasificadas como interest"*

### Diagnóstico

Dos bugs estructurales de extracción, no errores de catálogo intencionales:

**Bug 1 — Sección "Replacements" no explícita en el Excel del Val.**
El Excel budget del Val pone las 9 cuentas `6543-001..6543-010` (Replacements)
entre row 191 ("Taxes") y row 208 ("Subtotal Replacements"), sin un section
header "Replacements" propio. El extractor heredó "Taxes" como sección para
esas cuentas → las clasificó como Real Estate Taxes.

**Bug 2 — Header "Non Operating Expenses" en col C, no col B.**
El Excel del Val pone el section header "Non Operating Expenses" en la celda
**row 233 col C** (no col B). El extractor para formato HDFC usa `desc_col=1`
(col B) → no detectó el header → 4 cuentas bajo esa sección heredaron la
previa ("Financial") y quedaron como Interest Expense.

### Fix aplicado

Account-level overrides en `ACCOUNT_L1_OVERRIDES` dentro de
[scripts/extract_catalogs.py](../scripts/extract_catalogs.py) — 13 entradas
explícitas, auditables por account code.

| Account code | Descripción | L1 antes (wrong) | L1 después (correct) | Causa raíz |
|---|---|---|---|---|
| `6543-001` | Appliances | Real Estate Taxes | **Total Capex** | Bug 1 |
| `6543-002` | Carpet | Real Estate Taxes | **Total Capex** | Bug 1 |
| `6543-003` | Flooring | Real Estate Taxes | **Total Capex** | Bug 1 |
| `6543-004` | Paving | Real Estate Taxes | **Total Capex** | Bug 1 |
| `6543-005` | Redecorating | Real Estate Taxes | **Total Capex** | Bug 1 |
| `6543-006` | Other | Real Estate Taxes | **Total Capex** | Bug 1 |
| `6543-007` | Misc. Apt. | Real Estate Taxes | **Total Capex** | Bug 1 |
| `6543-009` | Reasonable Accommodations | Real Estate Taxes | **Total Capex** | Bug 1 |
| `6543-010` | Regulatory Inspections | Real Estate Taxes | **Total Capex** | Bug 1 |
| `6850-000` | MORTGAGE INSURANCE PREMIUM | Interest Expense | **Non-Operating Expenses** | Bug 2 + regla pilot (financing ≠ interest) |
| `6890-000` | ENTITY EXPENSE | Interest Expense | **Non-Operating Expenses** | Bug 2 |
| `6903-000` | COVID Expenses | Interest Expense | **Non-Operating Expenses** | Bug 2 |
| `6904-000` | Non Operating Other | Interest Expense | **Non-Operating Expenses** | Bug 2 |

### Impacto en distribuciones L1 de The Val

| L1 | Antes | Después | Δ |
|---|---:|---:|---:|
| Income | 37 | 37 | 0 |
| Operating Expenses | 106 | 106 | 0 |
| Real Estate Taxes | 13 | **4** | **-9** |
| Interest Expense | 16 | **12** | **-4** |
| Non-Operating Expenses | 2 | **6** | **+4** |
| Total Capex | 0 | **9** | **+9** |
| **Total** | **174** | **174** | 0 |

El total de cuentas no cambia — solo la asignación de L1 para las 13 afectadas.

### Efecto esperado sobre Q4 2025 del Val al re-correr el pipeline

- **NOI** sube (los 9 replacements salen de OpEx/RET arriba del NOI → ahora están debajo en Capex).
- **Net Income** cambia porque Mortgage Insurance, Entity Expense, COVID, Non-Op Other bajan Interest Expense y suben Non-Op Expense. El NI neto no cambia (ambos suman al mismo hijo antes de NI), pero la narrativa del AI ahora separa correctamente interest "real" de non-op.
- **Real Estate Taxes bullet** debería quedar chico (~solo el 6710-000) y sin inflar por replacements.
- **Interest Expense bullet** debería mostrar solo los 12 accounts de interés real (6820-xxx + 6830-000 Swap).

### Regresión verificada
Catálogos de los otros 4 activos (Alice House JV LLC, 295 29th Street JV LLC,
Walnut Street Wellesley, The Wilcox) mantienen sus distribuciones L1 idénticas.
Total de cuentas en el sistema: 1,117 (sin cambio).

### Cómo regenerar

```bash
python scripts/extract_catalogs.py
```

Los overrides viven en el script, así que el `src/asset_catalogs.py` generado
reflejará siempre los fixes mientras el script no cambie.

### Decisiones intencionales (confirmadas por Florencia, NO son bugs)

- **`6709-000 GROUND LEASE` (Val) → Real Estate Taxes** — el partner lo
  reporta bajo la sección "Taxes" y el plan de cuentas normalizado 2026
  lo mantiene ahí. Es una decisión de Florencia, no un error de herencia.
- **`8750-0000 Ground Lease` (Edson/295 29th Street JV LLC) → Interest
  Expense** — el partner lo reporta bajo "Debt Service". Idem: parte del
  catálogo normalizado 2026.

Cualquier cambio futuro a estas clasificaciones debe venir de Florencia.

### Fixes paralelos al AI prompt (no catálogo)

Los fixes a la **narrativa del AI** (no clasificar timing/structural sin
source, no mezclar líneas, no forward-looking, no filler, etc.) se hicieron
aparte en [src/output_formatter.py](../src/output_formatter.py) dentro de
`generate_copilot_prompt` — ver sección "FORBIDDEN NARRATIVE PATTERNS".
Esos fixes aplican a TODOS los activos, no solo Val.
