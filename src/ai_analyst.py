"""
AI Analyst Module — Connects to OpenAI or Anthropic APIs
to generate "Notes to Financials" and market analysis commentary
directly within the Streamlit app.
"""

import json
from typing import Optional, Dict

_SYSTEM_PROMPT = """\
You are a financial reporting synthesis engine for third-party capital in real estate asset management.

Your task is to generate quarterly "Notes to Financials" for building administration reports.

Do NOT:
- Infer causes
- Add business knowledge
- Explain trends unless explicitly explained in the source
- Classify items as structural, temporary, one-time, or recurring unless explicitly stated
- Add forward-looking language unless explicitly stated

If the source does not say it, OMIT it.

FORWARD-LOOKING STATEMENTS
Include forward-looking language ONLY if the source explicitly states continuation, normalization, or one-time nature. Never assume normalization, improvement, persistence, or reversion.

TIMING vs STRUCTURAL
Do NOT classify any variance as "timing" or "structural" unless the source text explicitly uses that language. No implied classification based on experience or patterns.

INCOME STATEMENT LOGIC
Specific case (Oakland / Alice-Edson only; NOT global):
- Gross Potential Rent + Concessions: analyze together as Net Effective Rent
- Vacancy + Loss/Gain to Lease: analyze together as Retention Rate
This logic is asset-specific and must NOT be applied to other assets unless explicitly instructed.
All other income and expense lines: analyze individually. Do NOT combine. Do NOT reinterpret.

TERMINOLOGY (STRICT)
If the source says "budget", use "budget". Do NOT replace "budget" with "underwriting". Never assume underwriting context unless explicitly stated.

NON-OPERATING / NON-OPEX ITEMS
Financing costs, legal expenses, taxes:
- Do NOT comment on recurrence, persistence, or dependency
- Do NOT add explanations (e.g., legal timelines)
- Only restate what is explicitly written

NUMBERS VS NARRATIVE
Do NOT convert numerical movements into qualitative statements unless explicitly stated. Example: a declining vacancy number does NOT imply "occupancy improved" unless written.

SCOPE CONTROL
Do NOT generalize asset-specific comments to portfolios. Do NOT extrapolate one building's explanation to others.

FORMAT RULES — STRICT BULLET FORMAT (NON-NEGOTIABLE)
- Output header must always be exactly: "Notes to Financials:"
- Output MUST be a bullet list, one bullet per L1 line commented.
- Each bullet MUST start with the L1 line name in bold followed by a colon (e.g., "**Income:**", "**Operating Expenses:**", "**Real Estate Taxes:**", "**Non-Operating Expenses:**").
- Each bullet MUST be 2–4 sentences MAXIMUM. Never more.
- Do NOT write long paragraphs. Do NOT write prose essays. Bullets are short and narrative.
- No special symbols (e.g., ~, tildes, bullets inside bullets).
- No sub-bullets. No nested lists.
- Text must be clean and ready for literal copy-paste into investor reports.
- Output language: ENGLISH only. Never Spanish, never mixed.

OPEX TOTAL-FIRST RULE (mandatory — feedback Florencia 23-Apr-2026)
For the Operating Expenses bullet ONLY:
- The FIRST sentence MUST state the total Operating Expenses variance for the quarter (and YTD if available). Example: "Operating Expenses were $21,764 unfavorable to budget for the quarter."
- ONLY AFTER stating the total can you list drivers (Payroll, R&M, Utilities).
- Drivers MUST be listed in DESCENDING order of |variance| (largest first).
- NEVER lead with a sub-driver before stating the total.
- This rule applies to Operating Expenses only — Income, Taxes, Non-Op, Interest, Capex have fewer sub-lines and don't need this enforcement.

DRIVER HIERARCHY (mandatory)
- When listing 2+ drivers in ANY bullet, order them by |variance| DESCENDING.
- Cap at 2 drivers per bullet (this overrides general "top-3" rules elsewhere).
- The 1–2 drivers chosen MUST collectively explain ≥70% of the L1 variance.

PARTNER COMMENT METRIC PRESERVATION
If partner / operator notes contain specific numbers (renewal rates 4.78%/8.70%/7.10%, vacancy 5.0%–6.1%, "12 move-in concessions", lease counts, days), preserve them VERBATIM or paraphrase tightly. Do NOT round, generalize, or omit these — they are credible operational color.

NUMBER GRANULARITY (STRICT)
- Round all dollar amounts to the nearest $5,000 or $10,000 (e.g., "$130,000", "$50,000", "$40,000"). Never write "$54,356" or "$219,766".
- Never use the tilde "~" before a number. Write the number clean.
- Never list more than TWO numerical drivers per bullet. If a line has 3+ sub-variances, mention only the top 1–2 by magnitude.
- Do NOT break down a single L1 variance into 3+ sub-components with individual dollar amounts (e.g., NEVER write "~$35k in R&M + ~$25k in payroll + ~$14k in utilities + ~$26k in G&A savings"). Pick the dominant driver and mention it once, rounded.
- Only mention variances that are MATERIAL — roughly >$10,000 or >5% of the L1 line total. Immaterial movements must be omitted.

NARRATIVE STYLE (MANAGEMENT COMMENTARY TONE)
- The output must read like management commentary to investors, NOT like an audit memo or variance reconciliation.
- Lead with the business "why" (when present in Management Commentary), not with a list of GL account movements.
- If the Management Commentary explains a cause (e.g., "appealing process was successful", "staffing model changed"), that cause IS the bullet. Supporting numbers come after, rounded, and only if material.
- If the Management Commentary does NOT explain a line, write a minimal factual bullet (line name + direction + rounded magnitude) and STOP. Do not invent color.

BLACKLISTED PHRASES (NEVER USE)
The following phrases are forbidden unless copied verbatim from the Management Commentary:
- "structural in nature"
- "expected to continue"
- "may moderate"
- "tied to turnover levels"
- "pressured results"
- "momentum"
- "dependent on sustained leasing"
- "not yet fully materialized"
- "reflecting seasonality"
- "gradual improvement"
- "partially offset by favorable performance"
- Any phrase classifying a variance as "structural", "temporary", "one-time", or "recurring" on its own authority.
If none of these are in the source, they cannot appear in the output.

RECURRING EVENTS ACROSS QUARTERS
If the Management Commentary marks an item as "recurring" (e.g., tax reassessment, insurance renewal,
long-term vacancy in specific units), reference it concisely without repeating the full explanation
from previous quarters. Example: "Real estate taxes remain elevated following the 2025 reassessment
(as noted in prior quarters)." Do NOT re-explain the cause in detail if it was already covered.
If an event is marked "one_off", treat it as isolated to the current period — do not carry it forward.

ROLE & CONTEXT
The reports are sent to external investors (third-party capital). The narrative standard follows the internal criteria of the reporting lead. Your output must match historical investor reports in tone, structure, and discipline.

FINAL SELF-CHECK (MANDATORY)
Before finalizing output, validate:
- Every statement is explicitly supported by source text
- No inferred causality exists
- No forward-looking language without explicit support
- No business knowledge has been injected
- Terminology matches the source exactly

CORE PRINCIPLE (NON-NEGOTIABLE)
You are NOT an analyst providing opinions. You are a faithful synthesis system. Faithfulness to source overrides completeness or elegance.

────────────────────────────────────────────────────
GOLD-STANDARD OUTPUT EXAMPLE (match this format, tone, and brevity)
────────────────────────────────────────────────────

Notes to Financials:

- **Income:** Q4 was the quarter more affected by lower income when comparing to the budget. This is driven especially by lower effective rent than expected, which accounts for around $130,000 of the deviation. $50,000 are represented by a higher vacancy and an additional $40,000 come from bad debt. The quarterly deviation represents most of the deviation for the year as we had assumed we would be able to reduce concessions throughout the year, with the lowest amount forecasted in Q4, which was not the case.

- **Operating Expenses:** The negative variance for Q4 was mostly driven by higher-than-expected payroll and R&M expenses. We made changes to staffing where we added a dedicated manager to each property, which was previously shared between the two properties. On R&M, we had to make additional repairs during Q4, to a window and the trash chute. Additionally, more appliances than anticipated in Q4's budget were purchased. On a YTD basis, the deviation was compensated by lower than expected utilities and contract services expenses.

- **Real Estate Taxes:** In July we were notified by the County Assessor that the appealing process was successful and that we would see a decrease in the assessed values for the properties. This resulted in a reduction of expenses for the second half of the year.

- **Non-Operating Expenses:** The deviation is represented by Alice House's loan amortization.

────────────────────────────────────────────────────

Notice how the gold-standard example:
- Uses bold L1 labels followed by a colon.
- Keeps each bullet to 2–4 sentences.
- Rounds numbers to the nearest $5k–$10k.
- Leads with the business "why" (appealing process, staffing changes, specific repairs) when the commentary provides it.
- Keeps bullets short when the commentary is thin (Non-Operating Expenses: one sentence).
- Uses natural investor-facing language. No "structural", no "tied to turnover", no "may moderate".
- Never breaks down a single L1 line into 3+ sub-dollar-amounts.

Your output MUST match this example in format, tone, brevity, and number granularity.\
"""


def _build_market_context(market_analysis) -> str:
    """Build market context string from MarketAnalysis object."""
    if not market_analysis:
        return ""

    c = market_analysis.current
    parts = [
        "\n────────────────────────────────────────────────────",
        "MARKET CONTEXT — MULTIFAMILY MARKET DATA",
        "────────────────────────────────────────────────────\n",
        f"Market: {market_analysis.market_name}",
        f"Period: {c.period}",
        f"Inventory: {c.inventory_units:,} units",
        f"Asking Rent: ${c.asking_rent:,.0f}/unit (${c.asking_rent_psf:.2f}/SF)",
        f"Effective Rent: ${c.effective_rent:,.0f}/unit (${c.effective_rent_psf:.2f}/SF)",
    ]

    if market_analysis.rent_trend_1yr is not None:
        parts.append(f"Rent Growth YoY: {market_analysis.rent_trend_1yr:+.1%}")
    if c.concessions_pct:
        parts.append(f"Concessions: {c.concessions_pct:.1%}")

    parts.append(f"Vacancy: {c.vacancy_pct:.1%}")
    parts.append(f"Occupancy: {c.occupancy_pct:.1%}")

    if market_analysis.vacancy_trend_1yr is not None:
        parts.append(f"Vacancy Change YoY: {market_analysis.vacancy_trend_1yr:+.1%}")

    parts.append(f"Absorption: {c.absorption_units:,} units")
    if market_analysis.absorption_avg_4q:
        parts.append(f"Avg Absorption (4Q): {market_analysis.absorption_avg_4q:,.0f} units")

    parts.append(f"Under Construction: {c.under_construction_units:,} units ({c.under_construction_pct:.1%} of inventory)")
    parts.append(f"Deliveries: {c.deliveries_units:,} units")

    parts.append(
        "\nThis market data is provided for reference only. "
        "Do NOT use it to infer, contextualize, or explain property-level variances. "
        "Only reference market conditions if the source financial report or management comments explicitly do so."
    )

    return "\n".join(parts)


def build_ai_prompt(
    copilot_prompt: str,
    market_analysis=None,
    custom_instructions: str = "",
    management_commentary: str = "",
) -> str:
    """
    Build the full prompt for the AI, combining:
    1. The existing copilot prompt (context + style + data + OUTPUT instruction)
    2. Management commentary (analyst memory layer)
    3. Market context (if available)
    4. Custom user instructions from the analyst briefing

    ORDEN CRÍTICO: las 3 piezas del analista (2, 3, 4) deben aparecer ANTES de
    la sección final "OUTPUT" del copilot_prompt — si van después, el LLM lee
    la instrucción OUTPUT primero y puede empezar a redactar sin haber leído
    el contexto del analista.

    Estructura del prompt final:
      [copilot_prompt: STEP 1..4] (data + estilo + JSON)
      [MANAGEMENT COMMENTARY]   ← autoritativo del analista
      [MARKET CONTEXT]
      [ADDITIONAL INSTRUCTIONS] ← briefing del analista
      [copilot_prompt: OUTPUT]  ← instrucción final de qué generar
    """
    mc_block = ""
    if management_commentary and management_commentary.strip():
        mc_block = (
            "\n════════════════════════════════════════════════════\n"
            "MANAGEMENT COMMENTARY — ANALYST INPUT\n"
            "════════════════════════════════════════════════════\n"
            "The following management commentary was provided by the asset management team.\n"
            "Use this context to enrich your Notes to Financials. These are AUTHORITATIVE\n"
            "explanations — incorporate them faithfully. Distinguish between recurring items\n"
            "and one-off events as labeled.\n\n"
            f"{management_commentary.strip()}"
        )

    market_ctx = _build_market_context(market_analysis) or ""

    briefing_block = ""
    if custom_instructions and custom_instructions.strip():
        briefing_block = (
            "\n════════════════════════════════════════════════════\n"
            "ADDITIONAL INSTRUCTIONS FROM ANALYST\n"
            "════════════════════════════════════════════════════\n"
            "The following is operational context for the current period written by the\n"
            "analyst in charge of this asset. It is AUTHORITATIVE — incorporate it faithfully\n"
            "into the relevant L1 bullets. If it mentions specific drivers, amounts, dates, or\n"
            "causes, use them verbatim (paraphrased to institutional tone).\n\n"
            f"{custom_instructions.strip()}"
        )

    extra_blocks = "\n\n".join(b for b in (mc_block, market_ctx, briefing_block) if b.strip())

    # Si no hay contexto adicional, devolver el prompt base tal cual
    if not extra_blocks:
        return copilot_prompt

    # Insertar extras ANTES de la sección OUTPUT del copilot_prompt.
    # El marker debe coincidir con el que escribe generate_copilot_prompt.
    output_marker = (
        "════════════════════════════════════════════════════\n"
        "OUTPUT\n"
        "════════════════════════════════════════════════════"
    )
    if output_marker in copilot_prompt:
        pre, output_section = copilot_prompt.split(output_marker, 1)
        return f"{pre.rstrip()}\n\n{extra_blocks}\n\n{output_marker}{output_section}"

    # Fallback (si la estructura del copilot_prompt cambia): append al final
    # pero con un warning visible para el LLM.
    return (
        f"{copilot_prompt}\n\n{extra_blocks}\n\n"
        "(NOTE: analyst context above must be incorporated into the bullets above.)"
    )


def call_openai(
    prompt: str,
    api_key: str,
    model: str = "gpt-4o",
    temperature: float = 0.2,
    max_tokens: int = 4000,
) -> str:
    """Call OpenAI API and return the response text."""
    from openai import OpenAI

    client = OpenAI(api_key=api_key)

    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "system",
                "content": _SYSTEM_PROMPT,
            },
            {"role": "user", "content": prompt},
        ],
        temperature=temperature,
        max_tokens=max_tokens,
    )

    return response.choices[0].message.content


def call_anthropic(
    prompt: str,
    api_key: str,
    model: str = "claude-sonnet-4-5",
    temperature: float = 0.2,
    max_tokens: int = 4000,
) -> str:
    """Call Anthropic Claude API and return the response text."""
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)

    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        temperature=temperature,
        system=_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
    )

    return response.content[0].text


def generate_ai_analysis(
    copilot_prompt: str,
    api_key: str,
    provider: str = "openai",
    model: str = "",
    market_analysis=None,
    custom_instructions: str = "",
    management_commentary: str = "",
    temperature: float = 0.2,
    max_tokens: int = 4000,
) -> str:
    """
    Main entry point: build prompt, call AI API, return generated text.

    Parameters:
        copilot_prompt: The pre-built copilot prompt with financial data
        api_key: API key for the selected provider
        provider: "openai" or "anthropic"
        model: Model name override (empty = default)
        market_analysis: Optional MarketAnalysis object
        custom_instructions: Additional analyst instructions
        management_commentary: Formatted management commentary text
        temperature: Generation temperature
        max_tokens: Max tokens for response
    """
    full_prompt = build_ai_prompt(
        copilot_prompt, market_analysis, custom_instructions, management_commentary,
    )

    if provider == "openai":
        default_model = "gpt-4o"
        return call_openai(
            prompt=full_prompt,
            api_key=api_key,
            model=model or default_model,
            temperature=temperature,
            max_tokens=max_tokens,
        )
    elif provider == "anthropic":
        default_model = "claude-sonnet-4-20250514"
        return call_anthropic(
            prompt=full_prompt,
            api_key=api_key,
            model=model or default_model,
            temperature=temperature,
            max_tokens=max_tokens,
        )
    else:
        raise ValueError(f"Provider '{provider}' not supported. Use 'openai' or 'anthropic'.")


# ─── Quality Scorer ─────────────────────────────────────────

# Forbidden phrases (case-insensitive). Detected → -1 punto cada una.
_FORBIDDEN_PHRASES = [
    "underwritten", "underwriting",
    "structural in nature", "structural increase",
    "expected to continue", "expected to normalize",
    "will normalize", "may moderate", "trend is expected",
    "tied to turnover", "pressured results", "momentum",
    "dependent on sustained", "not yet fully materialized",
    "reflecting seasonality", "gradual improvement",
    "partially offset by favorable",
    "helped offset", "helped partially offset",
    "this offset", "cushioned the impact",
    "contributed to the bottom line",
    "no specific operational driver", "no specific driver identified",
    "performance was in line with expectations",
    "characterized as structural", "indicates timing",
    "indicates consistent", "timing/accrual issue",
    "timing-related", "timing related",
    "completion of planned works", "completion of the program",
]

# L1 lines esperados en el output (al menos uno)
_EXPECTED_L1_LABELS = [
    "Income", "Operating Expenses", "Real Estate Taxes",
    "NOI", "Net Income",
    "Interest Expense", "Non-Operating Expenses",
    "Total Capex", "Capex",
]


def score_ai_output(output_text: str, structure: Optional[Dict] = None) -> Dict:
    """
    Evalúa la calidad del output del LLM contra las reglas del prompt.

    Returns:
        {
          'score': int (0-100),
          'grade': 'A'/'B'/'C'/'D'/'F',
          'checks': [{'name': str, 'passed': bool, 'detail': str}],
          'penalties': [{'rule': str, 'evidence': str, 'severity': int}],
        }
    """
    import re

    if not output_text or not output_text.strip():
        return {
            'score': 0, 'grade': 'F',
            'checks': [{'name': 'Output present', 'passed': False, 'detail': 'Empty output'}],
            'penalties': [],
        }

    text = output_text.strip()
    text_lower = text.lower()
    checks = []
    penalties = []

    # 1) Idioma: debe ser inglés (heurística: español tiene "ñ", "más", "según", etc.)
    spanish_markers = re.findall(r'\b(según|más|también|ñ|presentó|impulsada|debido a|debido al)\b', text_lower)
    if spanish_markers:
        checks.append({'name': 'Language = English', 'passed': False,
                       'detail': f"Spanish markers found: {set(spanish_markers)}"})
        penalties.append({'rule': 'Language', 'evidence': str(set(spanish_markers)), 'severity': 15})
    else:
        checks.append({'name': 'Language = English', 'passed': True, 'detail': '—'})

    # 2) Estructura de bullets: al menos un bullet con "**LineName:**"
    bullet_pattern = re.compile(r'\*\*([A-Z][A-Za-z &/-]+):\*\*')
    bullets = bullet_pattern.findall(text)
    if not bullets:
        checks.append({'name': 'Bullet structure', 'passed': False,
                       'detail': 'No "**LineName:**" bullets detected'})
        penalties.append({'rule': 'Structure', 'evidence': 'no bullets', 'severity': 20})
    else:
        checks.append({'name': 'Bullet structure', 'passed': True,
                       'detail': f"{len(bullets)} bullet(s): {bullets}"})

    # 3) Header "Notes to Financials:"
    has_header = bool(re.search(r'notes to financials\s*:', text_lower))
    checks.append({'name': 'Header present', 'passed': has_header,
                   'detail': '—' if has_header else 'Missing "Notes to Financials:" header'})
    if not has_header:
        penalties.append({'rule': 'Header', 'evidence': 'missing', 'severity': 5})

    # 4) Forbidden phrases
    found_forbidden = []
    for phrase in _FORBIDDEN_PHRASES:
        if phrase.lower() in text_lower:
            found_forbidden.append(phrase)
    if found_forbidden:
        checks.append({'name': 'No forbidden phrases', 'passed': False,
                       'detail': f"Found: {found_forbidden}"})
        for p in found_forbidden:
            penalties.append({'rule': 'Forbidden phrase', 'evidence': p, 'severity': 5})
    else:
        checks.append({'name': 'No forbidden phrases', 'passed': True, 'detail': '—'})

    # 5) OpEx total-first rule
    opex_match = re.search(r'\*\*Operating Expenses:\*\*\s*([^•\n]{0,500})', text)
    if opex_match:
        opex_first_section = opex_match.group(1)
        # Buscar dollar amount en las primeras 2 oraciones (≈ primeros 200 chars)
        first_sentences = opex_first_section[:250]
        has_dollar_first = bool(re.search(r'\$\d{1,3}(?:,\d{3})*(?:\.\d+)?', first_sentences))
        # Detectar si lo primero es un sub-driver (ej. "Maintenance ran...", "Payroll exceeded...")
        starts_with_subdriver = bool(re.match(
            r'\s*(maintenance|payroll|r&m|utilities|repairs|professional|legal|insurance|contract)',
            first_sentences.lower()
        ))
        if has_dollar_first and not starts_with_subdriver:
            checks.append({'name': 'OpEx total-first', 'passed': True, 'detail': '—'})
        else:
            checks.append({'name': 'OpEx total-first', 'passed': False,
                           'detail': 'OpEx bullet does not lead with the total variance'})
            penalties.append({'rule': 'OpEx total-first', 'evidence': first_sentences[:80], 'severity': 8})
    else:
        # No hay bullet de OpEx — neutro
        checks.append({'name': 'OpEx total-first', 'passed': True,
                       'detail': '(No Operating Expenses bullet)'})

    # 6) Forward-looking detection (lenguaje futuro)
    forward_patterns = [
        r'\bwill\s+(continue|normalize|reverse|improve|moderate|persist)\b',
        r'\bexpected\s+to\b',
        r'\banticipated\s+to\b',
        r'\bshould\s+(continue|reverse|normalize)\b',
        r'\bin\s+future\s+periods\b',
    ]
    forward_hits = []
    for pat in forward_patterns:
        m = re.search(pat, text_lower)
        if m:
            forward_hits.append(m.group(0))
    if forward_hits:
        checks.append({'name': 'No forward-looking language', 'passed': False,
                       'detail': f"Forward language detected: {forward_hits}"})
        for f in forward_hits:
            penalties.append({'rule': 'Forward-looking', 'evidence': f, 'severity': 5})
    else:
        checks.append({'name': 'No forward-looking language', 'passed': True, 'detail': '—'})

    # 7) Cross-line offset language
    cross_line = []
    for pat in ('helped offset', 'partially offset', 'offset opex', 'offset noi',
                'cushioned the impact', 'contributed to the bottom line'):
        if pat in text_lower:
            cross_line.append(pat)
    if cross_line:
        checks.append({'name': 'No cross-line narrative', 'passed': False,
                       'detail': f"Found: {cross_line}"})
        for p in cross_line:
            penalties.append({'rule': 'Cross-line', 'evidence': p, 'severity': 4})
    else:
        checks.append({'name': 'No cross-line narrative', 'passed': True, 'detail': '—'})

    # 8) Cobertura de L1: si tenemos structure, verificar que todos los L1 con
    #    |variance| > $5k tengan bullet
    if structure:
        l1_lines = structure.get('l1_lines') or []
        missing = []
        for l1 in l1_lines:
            if l1.name in ('NOI', 'Net Income'):
                continue
            var = abs(l1.variance_current or 0)
            if var < 5000:
                continue
            if l1.name not in bullets and l1.name.replace('Total ', '') not in bullets:
                missing.append(f"{l1.name} (${var:,.0f})")
        if missing:
            checks.append({'name': 'L1 coverage', 'passed': False,
                           'detail': f"Missing material L1s: {missing}"})
            for m in missing:
                penalties.append({'rule': 'L1 coverage', 'evidence': m, 'severity': 6})
        else:
            checks.append({'name': 'L1 coverage', 'passed': True, 'detail': '—'})

    # ── Cálculo de score ──
    score = 100
    for p in penalties:
        score -= p['severity']
    score = max(0, min(100, score))

    if score >= 90:
        grade = 'A'
    elif score >= 80:
        grade = 'B'
    elif score >= 70:
        grade = 'C'
    elif score >= 60:
        grade = 'D'
    else:
        grade = 'F'

    return {
        'score': score,
        'grade': grade,
        'checks': checks,
        'penalties': penalties,
    }


# ─── Available Models ───────────────────────────────────────

OPENAI_MODELS = [
    "gpt-4o",
    "gpt-4o-mini",
    "gpt-4-turbo",
    "o3-mini",
]

ANTHROPIC_MODELS = [
    "claude-opus-4-5",
    "claude-sonnet-4-5",
    "claude-haiku-4-5",
]
