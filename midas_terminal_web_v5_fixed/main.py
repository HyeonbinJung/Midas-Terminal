from __future__ import annotations

import asyncio
import math
import os
import threading
import time
from collections import deque
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Awaitable, Callable, Generic, TypeVar
from uuid import uuid4

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# Import services only after .env has been loaded because the research engine reads
# SEC / peer settings at import time.
from services.news_kr import collect_kr_news  # noqa: E402
from services.news_us import collect_us_news  # noqa: E402
from services.research_engine import (  # noqa: E402
    Catalyst,
    CompanyBackground,
    HighlightArgument,
    MarketData,
    RiskFactor,
    SecResearch,
    build_catalysts,
    build_company_background,
    build_highlights,
    build_risk_factors,
    build_key_financials_payload,
    build_filled_template_from_components,
    fetch_market_data,
    fetch_sec_research,
    select_peer_companies,
)
from services.translation import translate_research_components_to_korean  # noqa: E402

STATIC_DIR = BASE_DIR / "static"
GENERATED_DIR = BASE_DIR / "generated"
GENERATED_DIR.mkdir(exist_ok=True)

MARKET_CACHE_SECONDS = max(15, int(os.getenv("MARKET_CACHE_SECONDS", "60")))
RESEARCH_CACHE_SECONDS = max(60, int(os.getenv("RESEARCH_CACHE_SECONDS", "900")))
PEER_CACHE_SECONDS = max(300, int(os.getenv("PEER_CACHE_SECONDS", "3600")))
SEC_CACHE_SECONDS = max(900, int(os.getenv("SEC_CACHE_SECONDS", "86400")))
NEWS_CACHE_SECONDS = max(30, int(os.getenv("NEWS_CACHE_SECONDS", "300")))
MARKET_TIMEOUT_SECONDS = max(10, int(os.getenv("MARKET_TIMEOUT_SECONDS", "50")))
RESEARCH_COMPONENT_TIMEOUT_SECONDS = max(15, int(os.getenv("RESEARCH_COMPONENT_TIMEOUT_SECONDS", "75")))
REPORT_CONCURRENCY = max(1, min(int(os.getenv("REPORT_CONCURRENCY", "2")), 8))
YAHOO_CONCURRENCY = max(1, min(int(os.getenv("YAHOO_CONCURRENCY", "6")), 16))
GENERATED_STALE_SECONDS = max(300, int(os.getenv("GENERATED_STALE_SECONDS", "3600")))
TRANSLATION_TIMEOUT_SECONDS = max(60, int(os.getenv("TRANSLATION_TIMEOUT_SECONDS", "300")))

T = TypeVar("T")


class AsyncTTLCache(Generic[T]):
    """Small in-process TTL cache with per-key single-flight locking.

    If 20 users ask for AAPL at the same time, one upstream request is executed and
    the other 19 wait for and reuse that result.
    """

    def __init__(self, max_items: int = 256) -> None:
        self.max_items = max_items
        self._values: dict[str, tuple[float, T]] = {}
        self._locks: dict[str, asyncio.Lock] = {}

    def _prune(self) -> None:
        now = time.monotonic()
        expired = [key for key, (expires, _) in self._values.items() if expires <= now]
        for key in expired:
            self._values.pop(key, None)
        if len(self._values) > self.max_items:
            # Dict preserves insertion order; evict oldest inserted entries first.
            for key in list(self._values)[: len(self._values) - self.max_items]:
                self._values.pop(key, None)
        if len(self._locks) > self.max_items * 4:
            removable = [
                key for key, lock in self._locks.items()
                if key not in self._values and not lock.locked()
            ]
            for key in removable[: self.max_items * 2]:
                self._locks.pop(key, None)

    async def get_or_create(
        self,
        key: str,
        ttl_seconds: int,
        factory: Callable[[], Awaitable[T]],
    ) -> tuple[T, bool]:
        key = key.upper().strip()
        now = time.monotonic()
        cached = self._values.get(key)
        if cached and cached[0] > now:
            return cached[1], True

        lock = self._locks.setdefault(key, asyncio.Lock())
        async with lock:
            now = time.monotonic()
            cached = self._values.get(key)
            if cached and cached[0] > now:
                return cached[1], True
            value = await factory()
            self._values[key] = (time.monotonic() + ttl_seconds, value)
            self._prune()
            return value, False


class SlidingWindowLimiter:
    """In-process per-IP rate limiter suitable for the current single-PC/ngrok setup."""

    def __init__(self) -> None:
        self._events: dict[tuple[str, str], deque[float]] = {}
        self._lock = threading.Lock()

    def check(self, ip: str, bucket: str, limit: int, window_seconds: int) -> int | None:
        now = time.monotonic()
        cutoff = now - window_seconds
        key = (ip, bucket)
        with self._lock:
            events = self._events.setdefault(key, deque())
            while events and events[0] <= cutoff:
                events.popleft()
            if len(events) >= limit:
                retry = max(1, int(window_seconds - (now - events[0])) + 1)
                return retry
            events.append(now)
            # Opportunistic cleanup so an ngrok URL cannot grow this dictionary forever.
            if len(self._events) > 4096:
                stale_keys = [k for k, q in self._events.items() if not q or q[-1] <= cutoff]
                for stale in stale_keys[:1024]:
                    self._events.pop(stale, None)
        return None


@dataclass(slots=True)
class ResearchBundle:
    data: MarketData
    sec: SecResearch
    background: CompanyBackground
    peers: list[MarketData]
    catalysts: list[Catalyst]
    risks: list[RiskFactor]
    highlights: list[HighlightArgument]
    warnings: list[str]


@dataclass(slots=True)
class TranslatedResearch:
    background: CompanyBackground
    catalysts: list[Catalyst]
    risks: list[RiskFactor]
    highlights: list[HighlightArgument]


market_cache: AsyncTTLCache[MarketData] = AsyncTTLCache(max_items=256)
research_cache: AsyncTTLCache[ResearchBundle] = AsyncTTLCache(max_items=128)
sec_cache: AsyncTTLCache[SecResearch] = AsyncTTLCache(max_items=256)
peer_cache: AsyncTTLCache[list[MarketData]] = AsyncTTLCache(max_items=256)
news_cache: AsyncTTLCache[dict[str, Any]] = AsyncTTLCache(max_items=8)
translation_cache: AsyncTTLCache[TranslatedResearch] = AsyncTTLCache(max_items=128)
rate_limiter = SlidingWindowLimiter()
report_semaphore = asyncio.Semaphore(REPORT_CONCURRENCY)
yahoo_semaphore = asyncio.Semaphore(YAHOO_CONCURRENCY)
report_ticker_locks: dict[str, asyncio.Lock] = {}

app = FastAPI(title="Midas Capital Research Terminal", version="5.1.0")


class NoCacheStaticFiles(StaticFiles):
    """StaticFiles variant for local/ngrok development.

    Users often replace the project folder while keeping the same 127.0.0.1/ngrok
    origin open. Browser revalidation can otherwise leave an older app.js/index
    pair in memory, which made newly-added F2 Key Financials look as if they did
    not exist. Force static assets to be revalidated on every request.
    """

    async def get_response(self, path: str, scope: dict[str, Any]):
        response = await super().get_response(path, scope)
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response


app.mount("/static", NoCacheStaticFiles(directory=STATIC_DIR), name="static")


def _client_ip(request: Request) -> str:
    # ngrok forwards the original client IP. Fall back to the direct socket address.
    forwarded_values = [x.strip() for x in request.headers.get("x-forwarded-for", "").split(",") if x.strip()]
    if forwarded_values:
        # Prefer the last proxy-appended address so a client-supplied first value
        # cannot trivially bypass the per-IP limiter through ngrok.
        return forwarded_values[-1]
    return request.client.host if request.client else "unknown"


def _enforce_rate_limit(
    request: Request,
    bucket: str,
    *,
    limit: int,
    window_seconds: int,
) -> None:
    retry = rate_limiter.check(_client_ip(request), bucket, limit, window_seconds)
    if retry is not None:
        raise HTTPException(
            status_code=429,
            detail=f"Too many {bucket} requests. Try again in {retry}s.",
            headers={"Retry-After": str(retry)},
        )


def _finite(value: Any) -> float | None:
    try:
        number = float(value)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError):
        return None


def _market_payload(data: MarketData) -> dict[str, Any]:
    history = data.price_history
    change_pct = None
    previous_close = None
    if history is not None and not history.empty and "Close" in history.columns:
        closes = history["Close"].dropna()
        if len(closes) >= 2:
            previous_close = _finite(closes.iloc[-2])
            current = data.current_price or _finite(closes.iloc[-1])
            if previous_close and current is not None:
                change_pct = (current / previous_close - 1.0) * 100

    targets = data.analyst_targets or {}
    lowered = {str(k).lower(): v for k, v in targets.items()}

    def target(*keys: str) -> float | None:
        for key in keys:
            val = _finite(lowered.get(key.lower()))
            if val is not None:
                return val
        return None

    return {
        "ticker": data.ticker,
        "company_name": data.company_name,
        "exchange": data.exchange,
        "sector": data.sector,
        "industry": data.industry,
        "currency": data.currency,
        "current_price": data.current_price,
        "previous_close": previous_close,
        "change_pct": change_pct,
        "high_52w": data.high_52w,
        "low_52w": data.low_52w,
        "market_cap": data.market_cap,
        "shares_outstanding": data.shares_outstanding,
        "float_shares": data.float_shares,
        "beta": data.beta,
        "dividend_yield_pct": data.dividend_yield_pct,
        "avg_daily_value": data.avg_daily_value,
        "target_mean": target("mean", "average", "targetMeanPrice"),
        "target_median": target("median", "targetMedianPrice"),
        "target_high": target("high", "targetHighPrice"),
        "target_low": target("low", "targetLowPrice"),
    }


def _peer_payload(peer: MarketData) -> dict[str, Any]:
    info = peer.info or {}
    return {
        "ticker": peer.ticker,
        "company_name": peer.company_name,
        "market_cap": peer.market_cap,
        "sector": peer.sector,
        "industry": peer.industry,
        "current_price": peer.current_price,
        "revenue_growth": _finite(info.get("revenueGrowth")),
        "ebitda_margin": _finite(info.get("ebitdaMargins")),
        "beta": peer.beta,
    }


def _serialize_news(category_results: Any) -> dict[str, Any]:
    categories = []
    total = 0
    for category, items in category_results.items():
        serial = []
        for item in items:
            serial.append(
                {
                    "title": item.title,
                    "url": item.url,
                    "source": item.source,
                    "published_at": item.published_at.isoformat(),
                    "keyword": item.keyword,
                }
            )
        total += len(serial)
        categories.append({"name": category, "items": serial})
    return {"total": total, "categories": categories}


async def _get_market_data(ticker: str) -> MarketData:
    symbol = ticker.upper().strip()

    async def factory() -> MarketData:
        async with yahoo_semaphore:
            return await asyncio.wait_for(
                asyncio.to_thread(fetch_market_data, symbol),
                timeout=MARKET_TIMEOUT_SECONDS,
            )

    value, _ = await market_cache.get_or_create(symbol, MARKET_CACHE_SECONDS, factory)
    return value


async def _get_sec_data(ticker: str) -> SecResearch:
    symbol = ticker.upper().strip()

    async def factory() -> SecResearch:
        return await asyncio.wait_for(
            asyncio.to_thread(fetch_sec_research, symbol),
            timeout=RESEARCH_COMPONENT_TIMEOUT_SECONDS,
        )

    value, _ = await sec_cache.get_or_create(symbol, SEC_CACHE_SECONDS, factory)
    return value


async def _get_peer_data(data: MarketData) -> list[MarketData]:
    symbol = data.ticker.upper().strip()

    async def factory() -> list[MarketData]:
        async with yahoo_semaphore:
            return await asyncio.wait_for(
                asyncio.to_thread(select_peer_companies, data, 3),
                timeout=RESEARCH_COMPONENT_TIMEOUT_SECONDS,
            )

    value, _ = await peer_cache.get_or_create(symbol, PEER_CACHE_SECONDS, factory)
    return value


async def _get_research_bundle(ticker: str) -> ResearchBundle:
    symbol = ticker.upper().strip()

    async def factory() -> ResearchBundle:
        data = await _get_market_data(symbol)
        warnings: list[str] = []

        sec_result, peer_result = await asyncio.gather(
            _get_sec_data(data.ticker),
            _get_peer_data(data),
            return_exceptions=True,
        )

        if isinstance(sec_result, Exception):
            sec = SecResearch()
            warnings.append(f"SEC enrichment unavailable: {type(sec_result).__name__}")
        else:
            sec = sec_result
            if not sec.annual_url:
                warnings.append("No SEC annual filing was resolved for this ticker.")

        if isinstance(peer_result, Exception):
            peers: list[MarketData] = []
            warnings.append(f"Peer selection unavailable: {type(peer_result).__name__}")
        else:
            peers = peer_result
            if len(peers) < 3:
                warnings.append(f"Only {len(peers)} eligible peer(s) were resolved.")

        background, catalysts, risks, highlights = await asyncio.gather(
            asyncio.to_thread(build_company_background, data, sec),
            asyncio.to_thread(build_catalysts, data, sec),
            asyncio.to_thread(build_risk_factors, data, sec),
            asyncio.to_thread(build_highlights, data, peers, sec),
        )
        if not catalysts:
            warnings.append("No catalyst candidates could be generated from the available public data.")
        if not risks:
            warnings.append("No risk-factor rows could be generated from the available public data.")
        if not highlights:
            warnings.append("No fact-based highlight rows could be generated from the available public data.")
        return ResearchBundle(
            data=data, sec=sec, background=background, peers=peers, catalysts=catalysts,
            risks=risks, highlights=highlights, warnings=warnings,
        )

    value, _ = await research_cache.get_or_create(symbol, RESEARCH_CACHE_SECONDS, factory)
    return value


async def _get_korean_research(bundle: ResearchBundle) -> TranslatedResearch:
    symbol = bundle.data.ticker.upper().strip() + ":KO:V5_1"

    async def factory() -> TranslatedResearch:
        values = await asyncio.wait_for(
            asyncio.to_thread(
                translate_research_components_to_korean,
                bundle.background, bundle.catalysts, bundle.risks, bundle.highlights,
            ),
            timeout=TRANSLATION_TIMEOUT_SECONDS,
        )
        background, catalysts, risks, highlights = values
        return TranslatedResearch(background, catalysts, risks, highlights)

    value, _ = await translation_cache.get_or_create(symbol, RESEARCH_CACHE_SECONDS, factory)
    return value


def _safe_delete(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except Exception as exc:
        print(f"[CLEANUP] could not delete {path.name}: {type(exc).__name__}: {exc}")


def _cleanup_stale_generated() -> None:
    cutoff = time.time() - GENERATED_STALE_SECONDS
    for path in GENERATED_DIR.glob("*.docx"):
        try:
            if path.stat().st_mtime < cutoff:
                path.unlink(missing_ok=True)
        except OSError:
            continue


@app.on_event("startup")
async def startup_cleanup() -> None:
    await asyncio.to_thread(_cleanup_stale_generated)


@app.middleware("http")
async def security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    if request.url.path.startswith("/api/") or request.url.path == "/":
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


@app.get("/")
async def index() -> FileResponse:
    return FileResponse(
        STATIC_DIR / "index.html",
        headers={"Cache-Control": "no-store, no-cache, must-revalidate, max-age=0"},
    )


@app.get("/favicon.ico", include_in_schema=False)
async def favicon() -> FileResponse:
    return FileResponse(STATIC_DIR / "midas-logo.png", media_type="image/png")


@app.get("/api/health")
async def health() -> dict[str, Any]:
    return {
        "status": "ok",
        "service": "Midas Capital Research Terminal",
        "version": "5.1.0",
        "report_concurrency": REPORT_CONCURRENCY,
        "yahoo_concurrency": YAHOO_CONCURRENCY,
    }


@app.get("/api/market/{ticker}")
async def market(ticker: str, request: Request) -> dict[str, Any]:
    _enforce_rate_limit(request, "market", limit=60, window_seconds=60)
    try:
        data = await _get_market_data(ticker)
        return _market_payload(data)
    except asyncio.TimeoutError as exc:
        raise HTTPException(status_code=504, detail="Market-data request timed out.") from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/research/{ticker}")
async def research(ticker: str, request: Request) -> dict[str, Any]:
    _enforce_rate_limit(request, "research", limit=10, window_seconds=60)
    try:
        bundle = await _get_research_bundle(ticker)
        sec = bundle.sec
        background = bundle.background
        status = "partial" if bundle.warnings else "complete"
        return {
            "status": status,
            "warnings": bundle.warnings,
            "market": _market_payload(bundle.data),
            "key_financials": build_key_financials_payload(bundle.data),
            "background": {
                "founding_listing": background.founding_listing,
                "segments_revenue_mix": background.segments_revenue_mix,
                "revenue_geography": background.revenue_geography,
                "key_customers_channels": background.key_customers_channels,
                "competitors_market_share": background.competitors_market_share,
                "cost_structure_supply_chain": background.cost_structure_supply_chain,
                "ownership_major_shareholders": background.ownership_major_shareholders,
                "management_incentives": background.management_incentives,
            },
            "peers": [_peer_payload(p) for p in bundle.peers],
            "highlights": [
                {"argument": h.argument, "evidence_source": h.evidence_source}
                for h in bundle.highlights
            ],
            "catalysts": [
                {
                    "catalyst": c.catalyst,
                    "expected_date": c.expected_date,
                    "probability": c.probability,
                    "price_impact": c.price_impact,
                    "verify": c.verify,
                }
                for c in bundle.catalysts
            ],
            "risks": [
                {
                    "risk": r.risk,
                    "likelihood": r.likelihood,
                    "severity": r.severity,
                    "tp_impact": r.tp_impact,
                    "mitigation": r.mitigation,
                }
                for r in bundle.risks
            ],
            "sources": {
                "annual_form": sec.annual_form,
                "annual_filing_date": sec.annual_filing_date,
                "annual_url": sec.annual_url,
                "proxy_form": sec.proxy_form,
                "proxy_filing_date": sec.proxy_filing_date,
                "proxy_url": sec.proxy_url,
                "event_form": sec.event_form,
                "event_filing_date": sec.event_filing_date,
                "event_url": sec.event_url,
            },
        }
    except asyncio.TimeoutError as exc:
        raise HTTPException(status_code=504, detail="Research request timed out upstream.") from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/news/{region}")
async def news(region: str, request: Request) -> dict[str, Any]:
    _enforce_rate_limit(request, "news", limit=20, window_seconds=60)
    region = region.upper().strip()
    if region not in {"KR", "US"}:
        raise HTTPException(status_code=400, detail="region must be KR or US")

    async def factory() -> dict[str, Any]:
        raw = await (collect_kr_news() if region == "KR" else collect_us_news())
        return _serialize_news(raw)

    try:
        value, _ = await news_cache.get_or_create(region, NEWS_CACHE_SECONDS, factory)
        return value
    except Exception as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc


@app.post("/api/report/{ticker}")
async def report(ticker: str, request: Request, language: str = "en") -> FileResponse:
    # Export is intentionally the strictest endpoint: it creates a local file and may
    # trigger research if the cache is cold.
    _enforce_rate_limit(request, "report", limit=5, window_seconds=600)
    language = language.lower().strip()
    if language not in {"en", "ko"}:
        raise HTTPException(status_code=400, detail="language must be en or ko")
    safe_symbol = "".join(ch for ch in ticker.upper() if ch.isalnum() or ch in ".-^")[:24] or "REPORT"
    filename = (
        f"Midas_{safe_symbol.replace('^', '')}_{'KO' if language == 'ko' else 'EN'}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')[:-3]}_{uuid4().hex[:10]}.docx"
    )
    path = GENERATED_DIR / filename
    ticker_lock = report_ticker_locks.setdefault(safe_symbol, asyncio.Lock())

    try:
        bundle = await _get_research_bundle(ticker)
        background = bundle.background
        catalysts = bundle.catalysts
        risks = bundle.risks
        highlights = bundle.highlights
        if language == "ko":
            translated = await _get_korean_research(bundle)
            background = translated.background
            catalysts = translated.catalysts
            risks = translated.risks
            highlights = translated.highlights

        async with report_semaphore:
            # Serializing identical ticker exports avoids needless simultaneous writes
            # to the same template structures while UUID filenames still guarantee no
            # collision between users.
            async with ticker_lock:
                await asyncio.to_thread(
                    build_filled_template_from_components,
                    path,
                    bundle.data,
                    background,
                    bundle.peers,
                    bundle.sec,
                    catalysts,
                    risks,
                    highlights,
                    language,
                )
    except asyncio.TimeoutError as exc:
        _safe_delete(path)
        raise HTTPException(status_code=504, detail="Report research timed out upstream.") from exc
    except Exception as exc:
        _safe_delete(path)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return FileResponse(
        path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        background=BackgroundTask(_safe_delete, path),
        headers={
            "X-Midas-Research-Status": "partial" if bundle.warnings else "complete",
            "X-Midas-Report-Language": language,
        },
    )


if __name__ == "__main__":
    import uvicorn

    # No --reload here by default: reload mode is convenient for development but less
    # predictable under multiple users/ngrok. Use `python -m uvicorn main:app --reload`
    # explicitly when editing code.
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=False)
