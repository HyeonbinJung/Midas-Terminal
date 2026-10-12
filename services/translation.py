from __future__ import annotations

import asyncio
import inspect
import json
import re
from dataclasses import replace
from typing import Any
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from googletrans import Translator

from services.research_engine import (
    Catalyst,
    CompanyBackground,
    HighlightArgument,
    RiskFactor,
)


# Smaller chunks are much more reliable for long SEC filing paragraphs than the
# 3k+ character chunks used in v5. googletrans remains the primary translator.
MAX_TRANSLATION_CHARS = 1400
FINE_TRANSLATION_CHARS = 520
TRANSLATION_RETRIES = 3
GOOGLE_FALLBACK_TIMEOUT = 15


async def _maybe_await(value: Any) -> Any:
    if inspect.isawaitable(value):
        return await value
    return value


def _split_long_text(text: str, limit: int = MAX_TRANSLATION_CHARS) -> list[str]:
    """Split filing prose on natural boundaries while preserving every character."""
    text = (text or "").strip()
    if not text:
        return []
    if len(text) <= limit:
        return [text]

    chunks: list[str] = []
    remaining = text
    while len(remaining) > limit:
        candidates = [
            remaining.rfind(". ", 0, limit),
            remaining.rfind("; ", 0, limit),
            remaining.rfind(": ", 0, limit),
            remaining.rfind(", ", 0, limit),
            remaining.rfind(" ", 0, limit),
        ]
        cut = max(candidates)
        if cut < max(80, limit // 3):
            cut = limit
        else:
            cut += 1
        chunks.append(remaining[:cut].strip())
        remaining = remaining[cut:].strip()
    if remaining:
        chunks.append(remaining)
    return chunks


def _english_heavy(text: str) -> bool:
    if not text:
        return False
    letters = len(re.findall(r"[A-Za-z]", text))
    korean = len(re.findall(r"[가-힣]", text))
    return letters >= 18 and letters > korean * 1.25


def _looks_translated(source: str, translated: str) -> bool:
    source = (source or "").strip()
    translated = (translated or "").strip()
    if not translated:
        return False
    if translated != source:
        # A useful Korean result normally contains Hangul. Numeric/date-only strings
        # are allowed to remain identical.
        if re.search(r"[가-힣]", translated):
            return True
        if not _english_heavy(source):
            return True
    return not _english_heavy(source)


async def _googletrans_once(active: Any, text: str) -> str:
    try:
        item = await _maybe_await(active.translate(text, src="en", dest="ko"))
        return str(getattr(item, "text", item) or "").strip()
    except Exception:
        return ""


def _google_public_fallback_sync(text: str) -> str:
    """Best-effort Google Translate fallback using the public translate endpoint.

    This is only used when googletrans returns the English source unchanged. It is
    deliberately non-fatal: an outage must never prevent the DOCX from exporting.
    """
    try:
        params = urlencode({
            "client": "gtx",
            "sl": "en",
            "tl": "ko",
            "dt": "t",
            "q": text,
        })
        req = Request(
            "https://translate.googleapis.com/translate_a/single?" + params,
            headers={"User-Agent": "Mozilla/5.0 MidasCapitalResearch/5.1"},
        )
        with urlopen(req, timeout=GOOGLE_FALLBACK_TIMEOUT) as response:
            payload = json.loads(response.read().decode("utf-8"))
        pieces = payload[0] if isinstance(payload, list) and payload else []
        result = "".join(str(piece[0]) for piece in pieces if isinstance(piece, list) and piece and piece[0])
        return result.strip()
    except Exception:
        return ""


async def _translate_piece(active: Any, text: str) -> str:
    """Translate one chunk with progressively smaller retries.

    v5 treated one unchanged field as a hard error, causing the entire export to
    fail. v5.1 instead retries at sentence-sized granularity, then a secondary
    Google endpoint, and finally returns the source as a last-resort *without*
    aborting the report.
    """
    source = (text or "").strip()
    if not source:
        return source
    if not _english_heavy(source):
        # Dates, tickers, numeric ranges, etc. do not need translation.
        return source

    for attempt in range(TRANSLATION_RETRIES + 1):
        translated = await _googletrans_once(active, source)
        if _looks_translated(source, translated):
            return translated
        if attempt < TRANSLATION_RETRIES:
            await asyncio.sleep(0.25 * (attempt + 1))

    # Long paragraphs sometimes fail as one request but translate correctly when
    # split into much smaller sentence fragments.
    if len(source) > FINE_TRANSLATION_CHARS:
        fine_parts = _split_long_text(source, FINE_TRANSLATION_CHARS)
        translated_parts: list[str] = []
        any_changed = False
        for part in fine_parts:
            value = await _googletrans_once(active, part)
            if _looks_translated(part, value):
                translated_parts.append(value)
                any_changed = True
            else:
                fallback = await asyncio.to_thread(_google_public_fallback_sync, part)
                if _looks_translated(part, fallback):
                    translated_parts.append(fallback)
                    any_changed = True
                else:
                    translated_parts.append(part)
        if any_changed:
            return " ".join(x for x in translated_parts if x).strip()

    fallback = await asyncio.to_thread(_google_public_fallback_sync, source)
    if _looks_translated(source, fallback):
        return fallback

    # Never raise here. A temporary Google translation outage should not turn a
    # valid English research report into an HTTP 400/504. The next export gets a
    # fresh chance because the translation cache key is versioned in main.py.
    print("[TRANSLATION WARN] upstream returned source text unchanged; exporting best-effort field")
    return source


async def _translate_values_async(values: list[str]) -> list[str]:
    translator = Translator()

    async def run(active: Any) -> list[str]:
        # Keep concurrency low; Google Translate is more stable with a small number
        # of simultaneous long-text requests.
        semaphore = asyncio.Semaphore(2)

        async def translate_value(value: str) -> str:
            if not value or not value.strip():
                return value
            async with semaphore:
                parts = _split_long_text(value)
                translated_parts = [await _translate_piece(active, part) for part in parts]
                return " ".join(part for part in translated_parts if part).strip() or value

        return list(await asyncio.gather(*(translate_value(value) for value in values)))

    if hasattr(translator, "__aenter__") and hasattr(translator, "__aexit__"):
        async with translator as active:
            return await run(active)

    try:
        return await run(translator)
    finally:
        close = getattr(translator, "close", None)
        if callable(close):
            await _maybe_await(close())


def _translate_strings(values: list[str]) -> list[str]:
    """Translate every non-empty research string while preserving order."""
    return asyncio.run(_translate_values_async(values))


def translate_research_components_to_korean(
    background: CompanyBackground,
    catalysts: list[Catalyst],
    risks: list[RiskFactor],
    highlights: list[HighlightArgument],
) -> tuple[CompanyBackground, list[Catalyst], list[RiskFactor], list[HighlightArgument]]:
    """Translate all automatically populated narrative research fields to Korean.

    Numbers, tickers, valuation ranges and URLs are preserved. Fixed template labels
    are localized separately by research_engine after the document has been filled.
    """
    background_values = [
        background.founding_listing,
        background.segments_revenue_mix,
        background.revenue_geography,
        background.key_customers_channels,
        background.competitors_market_share,
        background.cost_structure_supply_chain,
        background.ownership_major_shareholders,
        background.management_incentives,
    ]

    catalyst_texts: list[str] = []
    for item in catalysts:
        catalyst_texts.extend([item.catalyst, item.expected_date, item.verify])

    risk_texts: list[str] = []
    for item in risks:
        risk_texts.extend([item.risk, item.mitigation])

    highlight_texts: list[str] = []
    for item in highlights:
        highlight_texts.extend([item.argument, item.evidence_source])

    all_values = background_values + catalyst_texts + risk_texts + highlight_texts
    translated = _translate_strings(all_values)
    cursor = 0

    bg_vals = translated[cursor:cursor + 8]
    cursor += 8
    translated_background = CompanyBackground(
        founding_listing=bg_vals[0],
        segments_revenue_mix=bg_vals[1],
        revenue_geography=bg_vals[2],
        key_customers_channels=bg_vals[3],
        competitors_market_share=bg_vals[4],
        cost_structure_supply_chain=bg_vals[5],
        ownership_major_shareholders=bg_vals[6],
        management_incentives=bg_vals[7],
    )

    translated_catalysts: list[Catalyst] = []
    for item in catalysts:
        title, expected_date, verify = translated[cursor:cursor + 3]
        cursor += 3
        translated_catalysts.append(
            replace(item, catalyst=title, expected_date=expected_date, verify=verify)
        )

    level_map = {"High": "높음", "Med": "중간", "Low": "낮음"}
    translated_risks: list[RiskFactor] = []
    for item in risks:
        risk_name, mitigation = translated[cursor:cursor + 2]
        cursor += 2
        translated_risks.append(
            replace(
                item,
                risk=risk_name,
                likelihood=level_map.get(item.likelihood, item.likelihood),
                severity=level_map.get(item.severity, item.severity),
                mitigation=mitigation,
            )
        )

    translated_highlights: list[HighlightArgument] = []
    for item in highlights:
        argument, evidence = translated[cursor:cursor + 2]
        cursor += 2
        translated_highlights.append(
            replace(item, argument=argument, evidence_source=evidence)
        )

    return translated_background, translated_catalysts, translated_risks, translated_highlights
