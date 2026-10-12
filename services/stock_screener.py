from __future__ import annotations

import math
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any

import numpy as np
import yfinance as yf

# Curated liquid US universe. Sector hints let us prune requests before touching Yahoo.
# The actual classification metrics are fetched from yfinance at search time.
UNIVERSE: dict[str, list[str]] = {
    "Technology": ["AAPL","MSFT","NVDA","AVGO","ORCL","CRM","ADBE","AMD","INTC","QCOM","TXN","AMAT","MU","NOW","PLTR","PANW","CRWD","SNPS","CDNS","INTU"],
    "Communication Services": ["GOOGL","META","NFLX","DIS","CMCSA","TMUS","VZ","T","CHTR","EA","TTWO","RBLX","SPOT","PINS"],
    "Consumer Cyclical": ["AMZN","TSLA","HD","LOW","MCD","NKE","SBUX","BKNG","MAR","GM","F","ABNB","ORLY","AZO","ROST","TJX"],
    "Consumer Defensive": ["WMT","COST","PG","KO","PEP","PM","MO","CL","MDLZ","KHC","GIS","KMB","KR","TGT"],
    "Financial Services": ["JPM","BAC","WFC","C","GS","MS","BLK","SCHW","AXP","V","MA","PYPL","COF","SYF","CME","ICE","SPGI","MCO"],
    "Healthcare": ["LLY","UNH","JNJ","ABBV","MRK","PFE","TMO","ABT","DHR","AMGN","GILD","ISRG","VRTX","REGN","CVS","CI","HCA","BSX"],
    "Industrials": ["GE","CAT","DE","HON","RTX","LMT","BA","UPS","FDX","UNP","CSX","ETN","PH","MMM","EMR","WM","URI","GWW"],
    "Energy": ["XOM","CVX","COP","EOG","SLB","MPC","PSX","OXY","VLO","KMI","WMB","HAL","DVN","FANG"],
    "Utilities": ["NEE","SO","DUK","CEG","AEP","SRE","D","EXC","XEL","PCG"],
    "Real Estate": ["PLD","AMT","EQIX","WELL","SPG","O","PSA","DLR","CCI","VICI"],
    "Basic Materials": ["LIN","SHW","FCX","NEM","APD","ECL","NUE","DOW","DD","VMC","MLM"],
}

SECTOR_SCORE_MAP = {
    "Utilities": 15,
    "Consumer Defensive": 22,
    "Real Estate": 30,
    "Healthcare": 36,
    "Financial Services": 48,
    "Industrials": 60,
    "Basic Materials": 70,
    "Energy": 76,
    "Consumer Cyclical": 82,
    "Communication Services": 86,
    "Technology": 92,
}

CYCLICALITY_SCORE_MAP = {
    "Defensive": 20,
    "Secular Growth": 58,
    "Cyclical": 86,
    "Unknown": 50,
}


def _finite(v: Any) -> float | None:
    try:
        x = float(v)
        return x if math.isfinite(x) else None
    except (TypeError, ValueError):
        return None


def _ratio(v: Any) -> float | None:
    return _finite(v)


def _clip01(x: float | None, fallback: float = 0.5) -> float:
    if x is None or not math.isfinite(x):
        return fallback
    return max(0.0, min(1.0, x))


def _scale(value: float | None, low: float, high: float, *, invert: bool = False, fallback: float = 0.5) -> int:
    if value is None or not math.isfinite(value) or high == low:
        return int(round(fallback * 100))
    x = (value - low) / (high - low)
    if invert:
        x = 1.0 - x
    return int(round(_clip01(x, fallback) * 100))


def _class_size(mc: float | None) -> str:
    if mc is None:
        return "Unknown"
    if mc >= 200e9:
        return "Mega"
    if mc >= 10e9:
        return "Large"
    if mc >= 2e9:
        return "Mid"
    if mc >= 300e6:
        return "Small"
    return "Micro"


def _class_vol(vol: float | None) -> str:
    if vol is None:
        return "Unknown"
    if vol < 0.25:
        return "Low"
    if vol < 0.45:
        return "Medium"
    return "High"


def _class_beta(beta: float | None) -> str:
    if beta is None:
        return "Unknown"
    if beta < 0.8:
        return "Low"
    if beta <= 1.2:
        return "Medium"
    return "High"


def _class_style(pe: float | None, rev_growth: float | None, fcf_yield: float | None) -> str:
    if pe is not None and pe > 0 and pe <= 12:
        return "Deep Value"
    if rev_growth is not None and rev_growth >= 0.18:
        return "Growth"
    if pe is not None and pe > 0 and pe <= 18 and (fcf_yield is None or fcf_yield >= 0.03):
        return "Value"
    if (rev_growth is not None and rev_growth >= 0.08) and (pe is None or pe <= 30):
        return "GARP"
    return "GARP"


def _class_growth(rev_growth: float | None) -> str:
    if rev_growth is None:
        return "Unknown"
    if rev_growth >= 0.20:
        return "High"
    if rev_growth >= 0.05:
        return "Medium"
    return "Low"


def _class_profitability(roe: float | None, net_margin: float | None) -> str:
    if net_margin is not None and net_margin < 0:
        return "Loss-making"
    if (roe is not None and roe >= 0.20) or (net_margin is not None and net_margin >= 0.15):
        return "High Quality"
    if roe is None and net_margin is None:
        return "Unknown"
    return "Average"


def _class_leverage(net_debt_ebitda: float | None, net_debt: float | None) -> str:
    if net_debt is not None and net_debt <= 0:
        return "Net Cash"
    if net_debt_ebitda is None:
        return "Unknown"
    if net_debt_ebitda <= 2.5:
        return "Moderate"
    return "High"


def _class_cashflow(fcf: float | None, fcf_margin: float | None, capex_ratio: float | None) -> str:
    if fcf is not None and fcf < 0:
        return "Cash Burn"
    if capex_ratio is not None and capex_ratio >= 0.12:
        return "CapEx Heavy"
    if fcf is not None and fcf > 0 and fcf_margin is not None and fcf_margin >= 0.10:
        return "FCF Compounder"
    if fcf is None:
        return "Unknown"
    return "Positive FCF"


def _class_dividend(div_yield: float | None) -> str:
    if div_yield is None or div_yield < 0.005:
        return "Non-dividend"
    if div_yield >= 0.04:
        return "High Yield"
    return "Dividend"


def _class_momentum(ret_6m: float | None) -> str:
    if ret_6m is None:
        return "Unknown"
    if ret_6m >= 0.10:
        return "Positive"
    if ret_6m <= -0.10:
        return "Negative"
    return "Neutral"


def _class_liquidity(adtv: float | None) -> str:
    if adtv is None:
        return "Unknown"
    if adtv >= 500e6:
        return "High"
    if adtv >= 50e6:
        return "Medium"
    return "Low"


def _class_cyclicality(sector: str) -> str:
    if sector in {"Consumer Defensive", "Utilities", "Healthcare"}:
        return "Defensive"
    if sector in {"Technology", "Communication Services"}:
        return "Secular Growth"
    return "Cyclical"


def _history_stats(ticker: str) -> tuple[float | None, float | None, float | None, float | None]:
    try:
        h = yf.Ticker(ticker).history(period="1y", interval="1d", auto_adjust=True, actions=False)
        if h is None or h.empty or "Close" not in h.columns:
            return None, None, None, None
        close = h["Close"].dropna()
        rets = close.pct_change().dropna()
        vol = float(rets.std() * np.sqrt(252)) if len(rets) >= 30 else None
        ret_6m = float(close.iloc[-1] / close.iloc[max(0, len(close) - 126)] - 1) if len(close) >= 30 else None
        ret_12m = float(close.iloc[-1] / close.iloc[0] - 1) if len(close) >= 30 else None
        adtv = None
        if "Volume" in h.columns:
            dollar = (h["Close"] * h["Volume"]).dropna().tail(63)
            if len(dollar):
                adtv = float(dollar.mean())
        return vol, ret_6m, ret_12m, adtv
    except Exception:
        return None, None, None, None


def _price_history_3y(ticker: str) -> list[dict[str, Any]]:
    """Compact weekly 3-year adjusted close series for the five returned matches."""
    try:
        h = yf.Ticker(ticker).history(period="3y", interval="1wk", auto_adjust=True, actions=False)
        if h is None or h.empty or "Close" not in h.columns:
            return []
        close = h["Close"].dropna()
        out: list[dict[str, Any]] = []
        for idx, value in close.items():
            try:
                out.append({"date": idx.strftime("%Y-%m-%d"), "close": round(float(value), 4)})
            except Exception:
                continue
        return out
    except Exception:
        return []


def _scores(
    market_cap: float | None,
    sector: str,
    vol: float | None,
    beta: float | None,
    pe: float | None,
    rev_growth: float | None,
    roe: float | None,
    net_margin: float | None,
    net_debt: float | None,
    nde: float | None,
    fcf_yield: float | None,
    fcf_margin: float | None,
    capex_ratio: float | None,
    div_yield: float | None,
    ret_6m: float | None,
    adtv: float | None,
    cyclicality: str,
) -> dict[str, int]:
    size_score = _scale(math.log10(max(market_cap or 3e8, 3e8)), math.log10(3e8), math.log10(2e12))
    sector_score = int(SECTOR_SCORE_MAP.get(sector, 50))
    volatility_score = _scale(vol, 0.15, 0.75)  # high vol => high score
    beta_score = _scale(beta, 0.5, 2.0)  # high beta => high score

    growth_component = _clip01(None if rev_growth is None else (rev_growth + 0.05) / 0.40)
    pe_component = _clip01(None if pe is None or pe <= 0 else (min(pe, 55.0) - 8.0) / (55.0 - 8.0), 0.5)
    fcf_value_component = _clip01(None if fcf_yield is None else fcf_yield / 0.08, 0.5)
    style_score = int(round((0.5 * growth_component + 0.35 * pe_component + 0.15 * (1.0 - fcf_value_component)) * 100))

    growth_score = _scale(rev_growth, -0.05, 0.35)
    roe_score = _scale(roe, -0.05, 0.35)
    margin_score = _scale(net_margin, -0.05, 0.30)
    profitability_score = int(round((roe_score + margin_score) / 2))

    if net_debt is not None and net_debt <= 0:
        leverage_score = 8
    else:
        leverage_score = _scale(nde, 0.0, 5.0)  # more leverage => higher score

    fcf_margin_score = _scale(fcf_margin, -0.05, 0.20)
    fcf_yield_score = _scale(fcf_yield, -0.02, 0.08)
    capex_burden = _scale(capex_ratio, 0.00, 0.20)  # high capex burden => high raw score
    cash_flow_score = int(round((0.45 * fcf_margin_score) + (0.35 * fcf_yield_score) + (0.20 * (100 - capex_burden))))

    dividend_score = _scale(div_yield, 0.00, 0.06)
    momentum_score = _scale(ret_6m, -0.30, 0.50)
    liquidity_score = _scale(math.log10(max(adtv or 1e6, 1e6)), 6.0, 10.0)
    cyclicality_score = int(CYCLICALITY_SCORE_MAP.get(cyclicality, 50))
    return {
        "market_cap": size_score,
        "sector": sector_score,
        "volatility": volatility_score,
        "beta": beta_score,
        "style": style_score,
        "growth": growth_score,
        "profitability": profitability_score,
        "leverage": leverage_score,
        "cash_flow": cash_flow_score,
        "dividend": dividend_score,
        "momentum": momentum_score,
        "liquidity": liquidity_score,
        "cyclicality": cyclicality_score,
    }


def _snapshot(ticker: str, sector_hint: str) -> dict[str, Any] | None:
    try:
        t = yf.Ticker(ticker)
        try:
            info = t.info or {}
        except Exception:
            info = {}
        vol, ret_6m, ret_12m, adtv = _history_stats(ticker)
        market_cap = _finite(info.get("marketCap"))
        if market_cap is None:
            try:
                market_cap = _finite(t.fast_info.get("market_cap"))
            except Exception:
                pass
        if market_cap is None:
            return None

        sector = str(info.get("sector") or sector_hint)
        industry = str(info.get("industry") or "")
        beta = _finite(info.get("beta"))
        pe = _finite(info.get("forwardPE")) or _finite(info.get("trailingPE"))
        rev_growth = _ratio(info.get("revenueGrowth"))
        roe = _ratio(info.get("returnOnEquity"))
        net_margin = _ratio(info.get("profitMargins"))
        revenue = _finite(info.get("totalRevenue"))
        fcf = _finite(info.get("freeCashflow"))
        fcf_margin = (fcf / revenue) if fcf is not None and revenue not in (None, 0) else None
        fcf_yield = (fcf / market_cap) if fcf is not None and market_cap else None
        total_debt = _finite(info.get("totalDebt"))
        cash = _finite(info.get("totalCash"))
        ebitda = _finite(info.get("ebitda"))
        net_debt = (total_debt - cash) if total_debt is not None and cash is not None else None
        nde = (net_debt / ebitda) if net_debt is not None and ebitda not in (None, 0) and ebitda > 0 else None
        capex = _finite(info.get("capitalExpenditures"))
        capex_ratio = abs(capex) / revenue if capex is not None and revenue not in (None, 0) else None
        div_yield = _ratio(info.get("dividendYield"))
        if div_yield is not None and div_yield > 0.25:
            div_yield /= 100.0
        cyclicality = _class_cyclicality(sector)
        return {
            "ticker": ticker,
            "name": str(info.get("shortName") or info.get("longName") or ticker),
            "sector": sector,
            "industry": industry,
            "market_cap": market_cap,
            "beta": beta,
            "annualized_volatility": vol,
            "forward_pe": pe,
            "revenue_growth": rev_growth,
            "roe": roe,
            "net_margin": net_margin,
            "fcf_yield": fcf_yield,
            "fcf": fcf,
            "fcf_margin": fcf_margin,
            "net_debt": net_debt,
            "net_debt_ebitda": nde,
            "capex_ratio": capex_ratio,
            "dividend_yield": div_yield,
            "ret_6m": ret_6m,
            "ret_12m": ret_12m,
            "avg_daily_value": adtv,
            "labels": {
                "market_cap": _class_size(market_cap),
                "volatility": _class_vol(vol),
                "beta": _class_beta(beta),
                "style": _class_style(pe, rev_growth, fcf_yield),
                "growth": _class_growth(rev_growth),
                "profitability": _class_profitability(roe, net_margin),
                "leverage": _class_leverage(nde, net_debt),
                "cash_flow": _class_cashflow(fcf, fcf_margin, capex_ratio),
                "dividend": _class_dividend(div_yield),
                "momentum": _class_momentum(ret_6m),
                "liquidity": _class_liquidity(adtv),
                "cyclicality": cyclicality,
            },
            "scores": _scores(
                market_cap, sector, vol, beta, pe, rev_growth, roe, net_margin,
                net_debt, nde, fcf_yield, fcf_margin, capex_ratio, div_yield,
                ret_6m, adtv, cyclicality,
            ),
        }
    except Exception:
        return None


def screen_stocks(filters: dict[str, str], limit: int = 5, max_candidates: int = 55) -> dict[str, Any]:
    selected_sector = filters.get("sector", "Any")
    pairs: list[tuple[str, str]] = []
    for sector, tickers in UNIVERSE.items():
        if selected_sector != "Any" and sector != selected_sector:
            continue
        pairs.extend((ticker, sector) for ticker in tickers)

    if len(pairs) > max_candidates:
        step = len(pairs) / max_candidates
        pairs = [pairs[int(i * step)] for i in range(max_candidates)]

    rows: list[dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=10) as pool:
        futures = {pool.submit(_snapshot, t, s): (t, s) for t, s in pairs}
        for fut in as_completed(futures):
            row = fut.result()
            if row:
                rows.append(row)

    def matches(row: dict[str, Any]) -> bool:
        labels = row.get("labels", {})
        if selected_sector != "Any" and row.get("sector") != selected_sector:
            return False
        for key, wanted in filters.items():
            if key == "sector" or wanted in (None, "", "Any"):
                continue
            if labels.get(key) != wanted:
                return False
        return True

    exact = [r for r in rows if matches(r)]
    exact.sort(key=lambda r: ((r.get("market_cap") or 0), (r.get("avg_daily_value") or 0)), reverse=True)
    selected = exact[: max(1, min(limit, 10))]

    if selected:
        with ThreadPoolExecutor(max_workers=min(5, len(selected))) as pool:
            hist_futures = {pool.submit(_price_history_3y, r["ticker"]): r for r in selected}
            for fut in as_completed(hist_futures):
                hist_futures[fut]["price_history_3y"] = fut.result()

    return {
        "count_scanned": len(rows),
        "count_matched": len(exact),
        "results": selected,
        "filters": filters,
        "note": (
            "Metrics and classifications are calculated from current yfinance / Yahoo Finance data. "
            "The 3D score graph uses deterministic 0-100 normalization for all 13 screening dimensions; "
            "Sector and Cyclicality scores are visualization mappings, while other scores come from live raw metrics."
        ),
    }
