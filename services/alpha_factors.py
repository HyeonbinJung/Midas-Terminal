from __future__ import annotations

import math
import re
from typing import Any

import numpy as np
import pandas as pd
import yfinance as yf


def _f(key: str, label: str, unit: str, description: str) -> dict[str, str]:
    return {"key": key, "label": label, "unit": unit, "description": description}


FACTOR_CATALOG: dict[str, list[dict[str, str]]] = {
    "Momentum": [
        _f("momentum_1m", "1M Return", "percent", "Quarter-end price return over the prior ~1 month."),
        _f("momentum_3m", "3M Return", "percent", "Quarter-end price return over the prior ~3 months."),
        _f("momentum_6m", "6M Return", "percent", "Quarter-end price return over the prior ~6 months."),
        _f("momentum_12m", "12M Return", "percent", "Quarter-end price return over the prior ~12 months."),
        _f("momentum_12_1m", "12-1M Momentum", "percent", "Return from ~12 months ago to ~1 month ago, excluding the latest month."),
        _f("rsi_14d", "RSI (14D)", "score", "14-trading-day Relative Strength Index sampled at each quarter end."),
        _f("macd_pct", "MACD", "percent", "12D EMA minus 26D EMA, expressed as a percentage of price at each quarter end."),
    ],
    "Value / Multiple": [
        _f("pe_ttm", "P/E (TTM)", "multiple", "Quarter-end price divided by trailing-four-quarter diluted EPS."),
        _f("forward_pe", "Forward P/E", "multiple", "Current Yahoo forward P/E; latest-quarter snapshot because historical consensus is not exposed."),
        _f("ev_ebitda_ttm", "EV / EBITDA (TTM)", "multiple", "Estimated quarter-end enterprise value divided by trailing-four-quarter EBITDA."),
        _f("ev_revenue_ttm", "EV / Revenue (TTM)", "multiple", "Estimated quarter-end enterprise value divided by trailing-four-quarter revenue."),
        _f("price_book", "P / Book", "multiple", "Estimated quarter-end market cap divided by stockholders' equity."),
        _f("price_fcf_ttm", "P / FCF (TTM)", "multiple", "Estimated quarter-end market cap divided by trailing-four-quarter free cash flow."),
        _f("fcf_yield_ttm", "FCF Yield (TTM)", "percent", "Trailing-four-quarter free cash flow divided by estimated quarter-end market cap."),
    ],
    "Growth": [
        _f("revenue_growth_yoy", "Revenue Growth YoY", "percent", "Quarterly revenue growth versus the same fiscal quarter one year earlier."),
        _f("ebitda_growth_yoy", "EBITDA Growth YoY", "percent", "Quarterly EBITDA growth versus the same fiscal quarter one year earlier."),
        _f("eps_growth_yoy", "EPS Growth YoY", "percent", "Diluted EPS growth versus the same fiscal quarter one year earlier."),
        _f("fcf_growth_yoy", "FCF Growth YoY", "percent", "Quarterly free-cash-flow growth versus the same fiscal quarter one year earlier."),
        _f("revenue_cagr_3y", "3Y Revenue CAGR", "percent", "Trailing-four-quarter revenue CAGR versus three years earlier when enough quarterly history exists."),
    ],
    "Profitability / Quality": [
        _f("gross_margin_ttm", "Gross Margin", "percent", "Trailing-four-quarter gross profit divided by revenue."),
        _f("ebitda_margin_ttm", "EBITDA Margin", "percent", "Trailing-four-quarter EBITDA divided by revenue."),
        _f("operating_margin_ttm", "Operating Margin", "percent", "Trailing-four-quarter operating income divided by revenue."),
        _f("net_margin_ttm", "Net Margin", "percent", "Trailing-four-quarter net income divided by revenue."),
        _f("roe_ttm", "ROE", "percent", "Trailing-four-quarter net income divided by average/current stockholders' equity."),
        _f("roa_ttm", "ROA", "percent", "Trailing-four-quarter net income divided by average/current total assets."),
        _f("roic_ttm", "ROIC", "percent", "Approximate NOPAT divided by average/current invested capital (equity + debt - cash)."),
    ],
    "Cash Flow Quality": [
        _f("cfo_net_income", "CFO / Net Income", "multiple", "Trailing-four-quarter operating cash flow divided by trailing net income."),
        _f("fcf_margin_ttm", "FCF Margin", "percent", "Trailing-four-quarter free cash flow divided by revenue."),
        _f("fcf_conversion", "FCF Conversion", "percent", "Trailing-four-quarter FCF divided by trailing-four-quarter EBITDA."),
        _f("accrual_ratio", "Accrual Ratio", "percent", "Trailing net income minus CFO, divided by average/current total assets."),
    ],
    "Leverage / Balance Sheet": [
        _f("debt_equity", "Debt / Equity", "multiple", "Quarter-end total debt divided by stockholders' equity."),
        _f("net_debt_ebitda", "Net Debt / EBITDA", "multiple", "Quarter-end net debt divided by trailing-four-quarter EBITDA."),
        _f("interest_coverage", "Interest Coverage", "multiple", "Trailing EBIT divided by absolute trailing interest expense."),
        _f("current_ratio", "Current Ratio", "multiple", "Quarter-end current assets divided by current liabilities."),
        _f("cash_assets", "Cash / Assets", "percent", "Quarter-end cash and short-term investments divided by total assets."),
    ],
    "Efficiency": [
        _f("asset_turnover", "Asset Turnover", "multiple", "Trailing revenue divided by average/current total assets."),
        _f("inventory_turnover", "Inventory Turnover", "multiple", "Trailing cost of revenue divided by average/current inventory."),
        _f("receivable_turnover", "Receivable Turnover", "multiple", "Trailing revenue divided by average/current accounts receivable."),
        _f("cash_conversion_cycle", "Cash Conversion Cycle", "days", "Approximate DIO + DSO - DPO using trailing revenue/COGS and average/current working-capital balances."),
    ],
    "Volatility / Risk": [
        _f("beta_1y", "Beta", "multiple", "Rolling daily-return beta versus SPY using up to ~252 trading days."),
        _f("volatility_30d", "30D Volatility", "percent", "Annualized standard deviation of the prior ~30 trading-day returns."),
        _f("volatility_90d", "90D Volatility", "percent", "Annualized standard deviation of the prior ~90 trading-day returns."),
        _f("max_drawdown_1y", "Max Drawdown", "percent", "Worst peak-to-trough drawdown over the prior ~252 trading days."),
        _f("downside_volatility_90d", "Downside Volatility", "percent", "Annualized standard deviation of negative daily returns over the prior ~90 trading days."),
    ],
    "Size / Liquidity": [
        _f("market_cap", "Market Cap", "currency", "Estimated market capitalization at each quarter end."),
        _f("adv_63d", "ADV", "currency", "Average daily dollar trading value over the prior ~63 trading days."),
        _f("volume_growth", "Volume Growth", "percent", "Recent ~21D average share volume versus the preceding ~63D average."),
        _f("bid_ask_spread", "Bid-Ask Spread", "percent", "Current bid-ask spread as a percentage of midpoint; latest-quarter snapshot only."),
    ],
    "Estimate Revision": [
        _f("eps_revision_1m", "EPS Revision 1M", "percent", "Current-quarter consensus EPS estimate change versus ~30 days ago when Yahoo exposes EPS trend data."),
        _f("eps_revision_3m", "EPS Revision 3M", "percent", "Current-quarter consensus EPS estimate change versus ~90 days ago when Yahoo exposes EPS trend data."),
        _f("revenue_revision", "Revenue Revision", "percent", "Revenue-estimate revision when Yahoo exposes a historical comparison; otherwise unavailable."),
        _f("target_price_revision", "Target Price Revision", "percent", "Target-price revision when a historical target comparison is exposed; otherwise unavailable."),
    ],
    "Earnings Surprise": [
        _f("eps_surprise", "EPS Surprise", "percent", "Reported EPS versus consensus EPS estimate on historical earnings dates."),
        _f("revenue_surprise", "Revenue Surprise", "percent", "Reported revenue versus consensus revenue estimate when exposed by Yahoo earnings history."),
        _f("post_earnings_drift", "Post-Earnings Drift", "percent", "Approximate 20-trading-day price return following each reported earnings date."),
    ],
    "Sentiment": [
        _f("analyst_rating_change", "Analyst Rating Change", "score", "Quarterly net upgrade/downgrade signal from Yahoo analyst actions."),
        _f("short_interest", "Short Interest", "percent", "Current short percent of float; latest-quarter snapshot only."),
        _f("news_sentiment", "News Sentiment", "score", "Simple lexical score across currently exposed Yahoo news headlines; latest-quarter snapshot only."),
        _f("insider_buying", "Insider Buying", "percent", "Net disclosed insider purchase value as a percentage of estimated market cap when transaction data is exposed."),
    ],
    "Shareholder Return": [
        _f("buyback_yield", "Buyback Yield", "percent", "Trailing-four-quarter common-stock repurchases divided by quarter-end market cap."),
        _f("dividend_yield", "Dividend Yield", "percent", "Trailing-four-quarter cash dividends divided by quarter-end market cap."),
        _f("net_payout_yield", "Net Payout Yield", "percent", "Trailing dividends + buybacks - share issuance, divided by quarter-end market cap."),
    ],
    "Capital Allocation": [
        _f("capex_revenue", "CapEx / Revenue", "percent", "Absolute trailing-four-quarter capital expenditure divided by trailing revenue."),
        _f("rd_revenue", "R&D / Revenue", "percent", "Trailing-four-quarter R&D expense divided by trailing revenue."),
        _f("ma_spending", "M&A Spending", "percent", "Absolute trailing-four-quarter acquisition/business-purchase cash outflow divided by trailing revenue."),
        _f("sbc_revenue", "SBC / Revenue", "percent", "Trailing-four-quarter stock-based compensation divided by trailing revenue."),
    ],
}

CATALOG_BY_KEY = {item["key"]: item for group in FACTOR_CATALOG.values() for item in group}


def _finite(v: Any) -> float | None:
    try:
        x = float(v)
        return x if math.isfinite(x) else None
    except (TypeError, ValueError):
        return None


def _norm_label(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def _line(df: pd.DataFrame | None, aliases: list[str]) -> pd.Series:
    if df is None or df.empty:
        return pd.Series(dtype="float64")
    normalized = {_norm_label(idx): idx for idx in df.index}
    chosen = None
    for alias in aliases:
        key = _norm_label(alias)
        if key in normalized:
            chosen = normalized[key]
            break
    if chosen is None:
        for alias in aliases:
            needle = _norm_label(alias)
            for key, idx in normalized.items():
                if needle and (needle in key or key in needle):
                    chosen = idx
                    break
            if chosen is not None:
                break
    if chosen is None:
        return pd.Series(dtype="float64")
    out = pd.to_numeric(df.loc[chosen], errors="coerce")
    out.index = pd.to_datetime(out.index).tz_localize(None)
    return out.sort_index()


def _period_series(series: pd.Series, periods: pd.PeriodIndex) -> pd.Series:
    if series.empty:
        return pd.Series(index=periods, dtype="float64")
    tmp = series.copy()
    tmp.index = pd.PeriodIndex(pd.to_datetime(tmp.index), freq="Q")
    tmp = tmp.groupby(level=0).last()
    return tmp.reindex(periods).astype("float64")


def _safe_div(a: pd.Series, b: pd.Series) -> pd.Series:
    out = a.astype(float) / b.astype(float).replace(0, np.nan)
    return out.replace([np.inf, -np.inf], np.nan)


def _yoy(series: pd.Series) -> pd.Series:
    prior = series.shift(4)
    return (series - prior) / prior.abs().replace(0, np.nan)


def _ttm(series: pd.Series) -> pd.Series:
    return series.rolling(4, min_periods=4).sum()


def _avg_balance(series: pd.Series) -> pd.Series:
    return ((series + series.shift(4)) / 2.0).combine_first(series)


def _latest_only(periods: pd.PeriodIndex, value: float | None) -> pd.Series:
    out = pd.Series(index=periods, dtype="float64")
    if len(periods) and value is not None:
        out.loc[periods[-1]] = value
    return out


def _price_at_or_before(close: pd.Series, when: pd.Timestamp, trading_days_back: int = 0) -> float | None:
    hist = close.loc[:when]
    if hist.empty:
        return None
    idx = len(hist) - 1 - trading_days_back
    if idx < 0:
        return None
    return _finite(hist.iloc[idx])


def _rsi_at(close: pd.Series, when: pd.Timestamp, lookback: int = 14) -> float | None:
    s = close.loc[:when].tail(lookback + 35)
    if len(s) < lookback + 1:
        return None
    delta = s.diff()
    gain = delta.clip(lower=0).ewm(alpha=1 / lookback, adjust=False).mean()
    loss = (-delta.clip(upper=0)).ewm(alpha=1 / lookback, adjust=False).mean()
    if loss.empty or gain.empty:
        return None
    l = _finite(loss.iloc[-1])
    g = _finite(gain.iloc[-1])
    if l in (None, 0):
        return 100.0 if g and g > 0 else 50.0
    rs = g / l
    return 100.0 - (100.0 / (1.0 + rs))


def _macd_pct_at(close: pd.Series, when: pd.Timestamp) -> float | None:
    s = close.loc[:when].tail(90)
    if len(s) < 26:
        return None
    ema12 = s.ewm(span=12, adjust=False).mean().iloc[-1]
    ema26 = s.ewm(span=26, adjust=False).mean().iloc[-1]
    price = s.iloc[-1]
    if price == 0:
        return None
    return _finite((ema12 - ema26) / price)


def _price_factor_series(close: pd.Series, volume: pd.Series, spy_close: pd.Series, quarter_dates: list[pd.Timestamp]) -> dict[str, pd.Series]:
    idx = pd.PeriodIndex([d.to_period("Q") for d in quarter_dates], freq="Q")
    keys = [
        "momentum_1m", "momentum_3m", "momentum_6m", "momentum_12m", "momentum_12_1m",
        "rsi_14d", "macd_pct", "beta_1y", "volatility_30d", "volatility_90d",
        "max_drawdown_1y", "downside_volatility_90d", "adv_63d", "volume_growth",
    ]
    result = {k: pd.Series(index=idx, dtype="float64") for k in keys}
    stock_ret = close.pct_change().dropna()
    spy_ret = spy_close.pct_change().dropna() if not spy_close.empty else pd.Series(dtype="float64")
    aligned = pd.concat([stock_ret.rename("stock"), spy_ret.rename("spy")], axis=1).dropna()
    dollar_volume = (close * volume.reindex(close.index)).dropna() if not volume.empty else pd.Series(dtype="float64")

    for period, d in zip(idx, quarter_dates):
        current = _price_at_or_before(close, d)
        for key, days in (("momentum_1m", 21), ("momentum_3m", 63), ("momentum_6m", 126), ("momentum_12m", 252)):
            past = _price_at_or_before(close, d, days)
            if current is not None and past not in (None, 0):
                result[key].loc[period] = current / past - 1.0
        p12 = _price_at_or_before(close, d, 252)
        p1 = _price_at_or_before(close, d, 21)
        if p12 not in (None, 0) and p1 is not None:
            result["momentum_12_1m"].loc[period] = p1 / p12 - 1.0
        rsi = _rsi_at(close, d)
        if rsi is not None:
            result["rsi_14d"].loc[period] = rsi
        macd = _macd_pct_at(close, d)
        if macd is not None:
            result["macd_pct"].loc[period] = macd

        for key, days in (("volatility_30d", 30), ("volatility_90d", 90)):
            window = stock_ret.loc[:d].tail(days)
            if len(window) >= max(15, days // 2):
                result[key].loc[period] = float(window.std() * np.sqrt(252))
        down = stock_ret.loc[:d].tail(90)
        down = down[down < 0]
        if len(down) >= 10:
            result["downside_volatility_90d"].loc[period] = float(down.std() * np.sqrt(252))
        prices = close.loc[:d].tail(253)
        if len(prices) >= 40:
            dd = prices / prices.cummax() - 1.0
            result["max_drawdown_1y"].loc[period] = float(dd.min())
        beta_window = aligned.loc[:d].tail(252)
        if len(beta_window) >= 40:
            var = float(beta_window["spy"].var())
            if var > 0:
                result["beta_1y"].loc[period] = float(beta_window["stock"].cov(beta_window["spy"]) / var)
        dv = dollar_volume.loc[:d].tail(63)
        if len(dv) >= 10:
            result["adv_63d"].loc[period] = float(dv.mean())
        if not volume.empty:
            vhist = volume.loc[:d].dropna()
            recent = vhist.tail(21)
            prior = vhist.iloc[max(0, len(vhist) - 84):max(0, len(vhist) - 21)]
            if len(recent) >= 10 and len(prior) >= 20 and prior.mean() != 0:
                result["volume_growth"].loc[period] = float(recent.mean() / prior.mean() - 1.0)
    return result


def _statement_frames(t: yf.Ticker) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    def get(kind: str) -> pd.DataFrame:
        try:
            if kind == "income":
                return t.get_income_stmt(freq="quarterly")
            if kind == "balance":
                return t.get_balance_sheet(freq="quarterly")
            return t.get_cash_flow(freq="quarterly")
        except Exception:
            try:
                attr = {"income": "quarterly_income_stmt", "balance": "quarterly_balance_sheet", "cash": "quarterly_cashflow"}[kind]
                value = getattr(t, attr)
                return value if isinstance(value, pd.DataFrame) else pd.DataFrame()
            except Exception:
                return pd.DataFrame()
    return get("income"), get("balance"), get("cash")




def _annual_revenue_cagr_3y(t: yf.Ticker) -> float | None:
    try:
        df = t.get_income_stmt(freq="yearly")
    except Exception:
        try:
            df = t.financials
        except Exception:
            return None
    rev = _line(df, ["Total Revenue", "Operating Revenue"])
    clean = rev.dropna()
    if len(clean) < 4:
        return None
    clean = clean.sort_index()
    start = _finite(clean.iloc[-4]); end = _finite(clean.iloc[-1])
    if start is None or end is None or start <= 0 or end <= 0:
        return None
    return _finite((end / start) ** (1.0 / 3.0) - 1.0)

def _serialize_series(series: pd.Series, periods: pd.PeriodIndex) -> list[float | None]:
    vals: list[float | None] = []
    for p in periods:
        v = _finite(series.get(p))
        vals.append(round(v, 8) if v is not None else None)
    return vals


def _diagnostic(values: pd.Series, forward_ret: pd.Series) -> dict[str, Any]:
    clean = values.dropna()
    latest = _finite(clean.iloc[-1]) if not clean.empty else None
    mean = _finite(clean.mean()) if not clean.empty else None
    std = _finite(clean.std(ddof=0)) if len(clean) >= 2 else None
    z = ((latest - mean) / std) if latest is not None and mean is not None and std not in (None, 0) else None
    joined = pd.concat([values.rename("factor"), forward_ret.rename("fwd")], axis=1).dropna()
    corr = None
    if len(joined) >= 4 and joined["factor"].nunique() > 1 and joined["fwd"].nunique() > 1:
        corr = _finite(joined["factor"].rank().corr(joined["fwd"].rank()))
    return {"latest": latest, "mean": mean, "zscore": _finite(z), "forward_3m_spearman": corr, "observations": int(len(joined)), "coverage": int(clean.size)}


def _safe_dataframe_call(t: yf.Ticker, method: str, *args: Any, **kwargs: Any) -> pd.DataFrame:
    try:
        fn = getattr(t, method)
        value = fn(*args, **kwargs)
        return value if isinstance(value, pd.DataFrame) else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def _first_numeric(row: pd.Series, names: list[str]) -> float | None:
    mapping = {_norm_label(c): c for c in row.index}
    for name in names:
        key = _norm_label(name)
        if key in mapping:
            v = _finite(row[mapping[key]])
            if v is not None:
                return v
    return None


def _eps_revisions(t: yf.Ticker) -> tuple[float | None, float | None]:
    df = _safe_dataframe_call(t, "get_eps_trend")
    if df.empty:
        try:
            value = getattr(t, "eps_trend")
            if isinstance(value, pd.DataFrame):
                df = value
        except Exception:
            pass
    if df.empty:
        return None, None
    row = None
    for idx in df.index:
        if str(idx).lower() in {"0q", "currentqtr", "currentquarter"}:
            row = df.loc[idx]
            break
    if row is None:
        row = df.iloc[0]
    current = _first_numeric(row, ["current"])
    d30 = _first_numeric(row, ["30daysAgo", "30 days ago", "30daysago"])
    d90 = _first_numeric(row, ["90daysAgo", "90 days ago", "90daysago"])
    rev1 = (current / d30 - 1.0) if current is not None and d30 not in (None, 0) else None
    rev3 = (current / d90 - 1.0) if current is not None and d90 not in (None, 0) else None
    return _finite(rev1), _finite(rev3)


def _earnings_event_series(t: yf.Ticker, close: pd.Series, periods: pd.PeriodIndex) -> tuple[pd.Series, pd.Series, pd.Series]:
    eps_out = pd.Series(index=periods, dtype="float64")
    rev_out = pd.Series(index=periods, dtype="float64")
    drift_out = pd.Series(index=periods, dtype="float64")
    df = _safe_dataframe_call(t, "get_earnings_dates", limit=24)
    if df.empty:
        return eps_out, rev_out, drift_out
    tmp = df.copy()
    try:
        tmp.index = pd.to_datetime(tmp.index).tz_localize(None)
    except Exception:
        return eps_out, rev_out, drift_out
    for d, row in tmp.sort_index().iterrows():
        p = pd.Timestamp(d).to_period("Q")
        if p not in periods:
            continue
        estimate = _first_numeric(row, ["EPS Estimate", "EPSEstimate", "Estimate"])
        actual = _first_numeric(row, ["Reported EPS", "ReportedEPS", "EPS Actual"])
        surprise = _first_numeric(row, ["Surprise(%)", "Surprise %", "SurprisePercent"])
        if surprise is not None:
            # Yahoo may expose 5.2 for 5.2% or 0.052 depending on version.
            if abs(surprise) > 2:
                surprise /= 100.0
            eps_out.loc[p] = surprise
        elif estimate not in (None, 0) and actual is not None:
            eps_out.loc[p] = (actual - estimate) / abs(estimate)
        rev_est = _first_numeric(row, ["Revenue Estimate", "RevenueEstimate"])
        rev_actual = _first_numeric(row, ["Reported Revenue", "ReportedRevenue", "Revenue Actual"])
        if rev_est not in (None, 0) and rev_actual is not None:
            rev_out.loc[p] = (rev_actual - rev_est) / abs(rev_est)
        after = close.loc[close.index >= pd.Timestamp(d)]
        if len(after) >= 21:
            start = _finite(after.iloc[0])
            end = _finite(after.iloc[20])
            if start not in (None, 0) and end is not None:
                drift_out.loc[p] = end / start - 1.0
    return eps_out, rev_out, drift_out


def _analyst_change_series(t: yf.Ticker, periods: pd.PeriodIndex) -> pd.Series:
    out = pd.Series(index=periods, dtype="float64")
    df = _safe_dataframe_call(t, "get_upgrades_downgrades")
    if df.empty:
        try:
            value = getattr(t, "upgrades_downgrades")
            if isinstance(value, pd.DataFrame):
                df = value
        except Exception:
            pass
    if df.empty:
        return out
    tmp = df.copy()
    try:
        tmp.index = pd.to_datetime(tmp.index).tz_localize(None)
    except Exception:
        date_col = next((c for c in tmp.columns if "date" in str(c).lower()), None)
        if date_col is None:
            return out
        parsed = pd.to_datetime(tmp[date_col], errors="coerce")
        try:
            parsed = parsed.dt.tz_localize(None)
        except Exception:
            pass
        tmp.index = parsed
    action_col = next((c for c in tmp.columns if _norm_label(c) in {"action", "gradeaction"}), None)
    if action_col is None:
        return out
    scores: dict[pd.Period, list[float]] = {}
    for d, row in tmp.iterrows():
        if pd.isna(d):
            continue
        p = pd.Timestamp(d).to_period("Q")
        if p not in periods:
            continue
        text = str(row[action_col]).lower()
        score = 1.0 if "up" in text else -1.0 if "down" in text else 0.0
        scores.setdefault(p, []).append(score)
    for p, vals in scores.items():
        out.loc[p] = float(sum(vals))
    return out


def _news_sentiment(t: yf.Ticker) -> float | None:
    try:
        news = t.news or []
    except Exception:
        return None
    positive = {"beat", "beats", "surge", "record", "upgrade", "upgraded", "growth", "profit", "profits", "strong", "bullish", "wins", "win", "raises", "raised", "outperform"}
    negative = {"miss", "misses", "cut", "cuts", "downgrade", "downgraded", "loss", "losses", "weak", "lawsuit", "probe", "fall", "falls", "drops", "drop", "bearish", "warning"}
    scores: list[float] = []
    for item in news[:50]:
        if not isinstance(item, dict):
            continue
        content = item.get("content") if isinstance(item.get("content"), dict) else item
        title = str(content.get("title") or item.get("title") or "")
        summary = str(content.get("summary") or item.get("summary") or "")
        words = re.findall(r"[a-z]+", (title + " " + summary).lower())
        if not words:
            continue
        p = sum(1 for w in words if w in positive)
        n = sum(1 for w in words if w in negative)
        if p or n:
            scores.append((p - n) / max(1, p + n))
    return float(np.mean(scores)) if scores else None


def _insider_buying_series(t: yf.Ticker, periods: pd.PeriodIndex, market_cap: pd.Series, close: pd.Series) -> pd.Series:
    out = pd.Series(index=periods, dtype="float64")
    df = _safe_dataframe_call(t, "get_insider_transactions")
    if df.empty:
        try:
            value = getattr(t, "insider_transactions")
            if isinstance(value, pd.DataFrame):
                df = value
        except Exception:
            pass
    if df.empty:
        return out
    date_col = next((c for c in df.columns if "date" in str(c).lower() or "start" in str(c).lower()), None)
    text_col = next((c for c in df.columns if "transaction" in str(c).lower() or "text" in str(c).lower()), None)
    value_col = next((c for c in df.columns if _norm_label(c) in {"value", "transactionvalue"}), None)
    shares_col = next((c for c in df.columns if "share" in str(c).lower()), None)
    if date_col is None or text_col is None:
        return out
    accum: dict[pd.Period, float] = {}
    for _, row in df.iterrows():
        d = pd.to_datetime(row[date_col], errors="coerce")
        if pd.isna(d):
            continue
        p = pd.Timestamp(d).to_period("Q")
        if p not in periods:
            continue
        text = str(row[text_col]).lower()
        sign = 1.0 if any(k in text for k in ["purchase", "buy", "acquisition"]) else -1.0 if any(k in text for k in ["sale", "sell", "disposition"]) else 0.0
        if sign == 0:
            continue
        value = _finite(row[value_col]) if value_col else None
        if value is None and shares_col:
            shares = abs(_finite(row[shares_col]) or 0.0)
            px = _price_at_or_before(close, pd.Timestamp(d))
            value = shares * px if shares and px is not None else None
        if value is not None:
            accum[p] = accum.get(p, 0.0) + sign * abs(value)
    for p, value in accum.items():
        mc = _finite(market_cap.get(p))
        if mc not in (None, 0):
            out.loc[p] = value / mc
    return out


def build_alpha_factor_payload(ticker: str) -> dict[str, Any]:
    symbol = ticker.upper().strip()
    if not symbol:
        raise ValueError("Ticker is required.")

    t = yf.Ticker(symbol)
    try:
        info = t.info or {}
    except Exception:
        info = {}

    hist = t.history(period="3y", interval="1d", auto_adjust=True, actions=False)
    if hist is None or hist.empty or "Close" not in hist.columns:
        raise ValueError(f"No 3-year price history returned for {symbol}.")
    close = hist["Close"].dropna().copy()
    close.index = pd.to_datetime(close.index).tz_localize(None)
    if "Volume" in hist.columns:
        volume = hist["Volume"].dropna().astype(float).copy()
        volume.index = pd.to_datetime(volume.index).tz_localize(None)
        volume = volume.reindex(close.index)
    else:
        volume = pd.Series(dtype="float64")

    qlast = close.groupby(close.index.to_period("Q")).tail(1).tail(12)
    quarter_dates = [pd.Timestamp(x) for x in qlast.index]
    periods = pd.PeriodIndex([d.to_period("Q") for d in quarter_dates], freq="Q")
    quarter_price = pd.Series([float(v) for v in qlast.values], index=periods, dtype="float64")

    try:
        spy = yf.Ticker("SPY").history(period="3y", interval="1d", auto_adjust=True, actions=False)
        spy_close = spy["Close"].dropna().copy() if spy is not None and not spy.empty and "Close" in spy.columns else pd.Series(dtype="float64")
        if not spy_close.empty:
            spy_close.index = pd.to_datetime(spy_close.index).tz_localize(None)
    except Exception:
        spy_close = pd.Series(dtype="float64")

    income, balance, cashflow = _statement_frames(t)

    revenue = _period_series(_line(income, ["Total Revenue", "Operating Revenue"]), periods)
    cost_revenue = _period_series(_line(income, ["Cost Of Revenue", "Cost of Revenue", "Cost Of Goods Sold"]), periods)
    gross_profit = _period_series(_line(income, ["Gross Profit"]), periods)
    operating_income = _period_series(_line(income, ["Operating Income"]), periods)
    ebit = _period_series(_line(income, ["EBIT"]), periods).combine_first(operating_income)
    net_income = _period_series(_line(income, ["Net Income", "Net Income Common Stockholders"]), periods)
    pretax_income = _period_series(_line(income, ["Pretax Income", "Income Before Tax"]), periods)
    tax_provision = _period_series(_line(income, ["Tax Provision", "Income Tax Expense"]), periods)
    interest_expense = _period_series(_line(income, ["Interest Expense", "Interest Expense Non Operating"]), periods)
    ebitda = _period_series(_line(income, ["EBITDA", "Normalized EBITDA"]), periods)
    eps = _period_series(_line(income, ["Diluted EPS", "Basic EPS"]), periods)
    diluted_shares = _period_series(_line(income, ["Diluted Average Shares", "Basic Average Shares"]), periods)
    rd = _period_series(_line(income, ["Research And Development", "Research Development"]), periods)

    equity = _period_series(_line(balance, ["Stockholders Equity", "Common Stock Equity", "Total Equity Gross Minority Interest"]), periods)
    total_assets = _period_series(_line(balance, ["Total Assets"]), periods)
    inventory = _period_series(_line(balance, ["Inventory"]), periods)
    receivables = _period_series(_line(balance, ["Accounts Receivable", "Receivables", "Net Receivables"]), periods)
    payables = _period_series(_line(balance, ["Accounts Payable", "Payables And Accrued Expenses", "Payables"]), periods)
    debt = _period_series(_line(balance, ["Total Debt"]), periods)
    if debt.isna().all():
        current_debt = _period_series(_line(balance, ["Current Debt", "Current Debt And Capital Lease Obligation"]), periods)
        long_debt = _period_series(_line(balance, ["Long Term Debt", "Long Term Debt And Capital Lease Obligation"]), periods)
        debt = current_debt.fillna(0) + long_debt.fillna(0)
        debt[(current_debt.isna()) & (long_debt.isna())] = np.nan
    cash = _period_series(_line(balance, ["Cash Cash Equivalents And Short Term Investments", "Cash And Cash Equivalents", "Cash Financial"]), periods)
    current_assets = _period_series(_line(balance, ["Current Assets", "Total Current Assets"]), periods)
    current_liabilities = _period_series(_line(balance, ["Current Liabilities", "Total Current Liabilities"]), periods)
    bs_shares = _period_series(_line(balance, ["Ordinary Shares Number", "Share Issued"]), periods)

    cfo = _period_series(_line(cashflow, ["Operating Cash Flow", "Total Cash From Operating Activities"]), periods)
    capex = _period_series(_line(cashflow, ["Capital Expenditure", "Capital Expenditures"]), periods)
    free_cf = _period_series(_line(cashflow, ["Free Cash Flow"]), periods)
    if free_cf.isna().all() and (not cfo.isna().all() or not capex.isna().all()):
        free_cf = cfo + capex.where(capex <= 0, -capex)
    repurchase = _period_series(_line(cashflow, ["Repurchase Of Capital Stock", "Repurchase Of Stock", "Common Stock Purchase"]), periods)
    dividends = _period_series(_line(cashflow, ["Cash Dividends Paid", "Common Stock Dividend Paid"]), periods)
    issuance = _period_series(_line(cashflow, ["Issuance Of Capital Stock", "Common Stock Issuance"]), periods)
    acquisitions = _period_series(_line(cashflow, ["Net Business Purchases", "Acquisitions Net", "Purchase Of Business"]), periods)
    sbc = _period_series(_line(cashflow, ["Stock Based Compensation", "Stock Based Compensation Expense"]), periods)

    revenue_ttm = _ttm(revenue)
    cost_revenue_ttm = _ttm(cost_revenue)
    gross_profit_ttm = _ttm(gross_profit)
    operating_income_ttm = _ttm(operating_income)
    ebit_ttm = _ttm(ebit)
    net_income_ttm = _ttm(net_income)
    pretax_ttm = _ttm(pretax_income)
    tax_ttm = _ttm(tax_provision)
    interest_ttm = _ttm(interest_expense).abs()
    ebitda_ttm = _ttm(ebitda)
    eps_ttm = _ttm(eps)
    fcf_ttm = _ttm(free_cf)
    cfo_ttm = _ttm(cfo)
    capex_ttm = _ttm(capex).abs()
    rd_ttm = _ttm(rd).abs()
    acquisitions_ttm = _ttm(acquisitions).abs()
    sbc_ttm = _ttm(sbc).abs()

    shares = bs_shares.combine_first(diluted_shares)
    current_shares = _finite(info.get("sharesOutstanding"))
    if current_shares is not None:
        shares = shares.fillna(current_shares)
    market_cap = quarter_price * shares
    current_mc = _finite(info.get("marketCap"))
    current_price = _finite(info.get("currentPrice")) or _finite(close.iloc[-1])
    if market_cap.isna().all() and current_mc and current_price:
        market_cap = quarter_price * (current_mc / current_price)
    elif current_mc is not None and len(periods):
        market_cap.loc[periods[-1]] = current_mc

    net_debt = debt - cash
    enterprise_value = market_cap + net_debt

    avg_assets = _avg_balance(total_assets)
    avg_equity = _avg_balance(equity)
    avg_inventory = _avg_balance(inventory)
    avg_receivables = _avg_balance(receivables)
    avg_payables = _avg_balance(payables)
    invested_capital = equity + debt - cash
    avg_invested = _avg_balance(invested_capital)
    tax_rate = _safe_div(tax_ttm, pretax_ttm).clip(lower=0, upper=0.5)
    nopat = ebit_ttm * (1.0 - tax_rate.fillna(0.21))

    factors: dict[str, pd.Series] = _price_factor_series(close, volume, spy_close, quarter_dates)
    factors.update({
        "market_cap": market_cap,
        "pe_ttm": _safe_div(quarter_price, eps_ttm),
        "ev_ebitda_ttm": _safe_div(enterprise_value, ebitda_ttm),
        "ev_revenue_ttm": _safe_div(enterprise_value, revenue_ttm),
        "price_book": _safe_div(market_cap, equity),
        "price_fcf_ttm": _safe_div(market_cap, fcf_ttm),
        "fcf_yield_ttm": _safe_div(fcf_ttm, market_cap),
        "revenue_growth_yoy": _yoy(revenue),
        "ebitda_growth_yoy": _yoy(ebitda),
        "eps_growth_yoy": _yoy(eps),
        "fcf_growth_yoy": _yoy(free_cf),
        "revenue_cagr_3y": (revenue_ttm / revenue_ttm.shift(12)).where((revenue_ttm > 0) & (revenue_ttm.shift(12) > 0)).pow(1 / 3) - 1.0,
        "gross_margin_ttm": _safe_div(gross_profit_ttm, revenue_ttm),
        "ebitda_margin_ttm": _safe_div(ebitda_ttm, revenue_ttm),
        "operating_margin_ttm": _safe_div(operating_income_ttm, revenue_ttm),
        "net_margin_ttm": _safe_div(net_income_ttm, revenue_ttm),
        "roe_ttm": _safe_div(net_income_ttm, avg_equity),
        "roa_ttm": _safe_div(net_income_ttm, avg_assets),
        "roic_ttm": _safe_div(nopat, avg_invested),
        "cfo_net_income": _safe_div(cfo_ttm, net_income_ttm),
        "fcf_margin_ttm": _safe_div(fcf_ttm, revenue_ttm),
        "fcf_conversion": _safe_div(fcf_ttm, ebitda_ttm),
        "accrual_ratio": _safe_div(net_income_ttm - cfo_ttm, avg_assets),
        "debt_equity": _safe_div(debt, equity),
        "net_debt_ebitda": _safe_div(net_debt, ebitda_ttm),
        "interest_coverage": _safe_div(ebit_ttm, interest_ttm),
        "current_ratio": _safe_div(current_assets, current_liabilities),
        "cash_assets": _safe_div(cash, total_assets),
        "asset_turnover": _safe_div(revenue_ttm, avg_assets),
        "inventory_turnover": _safe_div(cost_revenue_ttm, avg_inventory),
        "receivable_turnover": _safe_div(revenue_ttm, avg_receivables),
        "cash_conversion_cycle": (_safe_div(avg_inventory, cost_revenue_ttm) + _safe_div(avg_receivables, revenue_ttm) - _safe_div(avg_payables, cost_revenue_ttm)) * 365.0,
        "capex_revenue": _safe_div(capex_ttm, revenue_ttm),
        "rd_revenue": _safe_div(rd_ttm, revenue_ttm),
        "ma_spending": _safe_div(acquisitions_ttm, revenue_ttm),
        "sbc_revenue": _safe_div(sbc_ttm, revenue_ttm),
    })

    # Current-only metrics Yahoo does not expose historically.
    annual_rev_cagr = _annual_revenue_cagr_3y(t)
    if len(periods) and annual_rev_cagr is not None and pd.isna(factors["revenue_cagr_3y"].iloc[-1]):
        factors["revenue_cagr_3y"].loc[periods[-1]] = annual_rev_cagr

    factors["forward_pe"] = _latest_only(periods, _finite(info.get("forwardPE")))
    bid = _finite(info.get("bid")); ask = _finite(info.get("ask"))
    mid = (bid + ask) / 2.0 if bid is not None and ask is not None and (bid + ask) != 0 else None
    factors["bid_ask_spread"] = _latest_only(periods, ((ask - bid) / mid) if mid and bid is not None and ask is not None else None)
    short_pct = _finite(info.get("shortPercentOfFloat"))
    if short_pct is not None and short_pct > 1.0:
        short_pct /= 100.0
    factors["short_interest"] = _latest_only(periods, short_pct)
    factors["news_sentiment"] = _latest_only(periods, _news_sentiment(t))

    eps1, eps3 = _eps_revisions(t)
    factors["eps_revision_1m"] = _latest_only(periods, eps1)
    factors["eps_revision_3m"] = _latest_only(periods, eps3)
    factors["revenue_revision"] = _latest_only(periods, None)
    factors["target_price_revision"] = _latest_only(periods, None)

    eps_surprise, revenue_surprise, drift = _earnings_event_series(t, close, periods)
    factors["eps_surprise"] = eps_surprise
    factors["revenue_surprise"] = revenue_surprise
    factors["post_earnings_drift"] = drift
    factors["analyst_rating_change"] = _analyst_change_series(t, periods)
    factors["insider_buying"] = _insider_buying_series(t, periods, market_cap, close)

    # Shareholder-return factors: normalize sign conventions so repurchases/dividends are positive payouts.
    buyback_q = (-repurchase).where(repurchase < 0, repurchase.abs())
    dividend_q = (-dividends).where(dividends < 0, dividends.abs())
    issuance_q = issuance.where(issuance > 0, issuance.abs())
    buyback_ttm = _ttm(buyback_q)
    dividend_ttm = _ttm(dividend_q)
    issuance_ttm = _ttm(issuance_q)
    factors["buyback_yield"] = _safe_div(buyback_ttm, market_cap)
    factors["dividend_yield"] = _safe_div(dividend_ttm, market_cap)
    factors["net_payout_yield"] = _safe_div(dividend_ttm + buyback_ttm - issuance_ttm, market_cap)

    # Ensure every catalog key exists even when Yahoo cannot expose the underlying history.
    for key in CATALOG_BY_KEY:
        factors.setdefault(key, pd.Series(index=periods, dtype="float64"))

    forward_3m = quarter_price.shift(-1) / quarter_price - 1.0
    diagnostics = {key: _diagnostic(series, forward_3m) for key, series in factors.items()}

    # Preserve the native daily 3-year adjusted close path for the Alpha overlay.
    # Factors are quarterly, but price should remain continuous rather than being
    # collapsed to quarter-end or weekly observations.
    price_history = [{"date": d.strftime("%Y-%m-%d"), "close": round(float(v), 4)} for d, v in close.items()]
    available_fundamental_quarters = int(max(
        revenue.notna().sum(), equity.notna().sum(), free_cf.notna().sum(), eps.notna().sum(), ebitda.notna().sum(), total_assets.notna().sum()
    ))
    available_factor_count = sum(1 for key in CATALOG_BY_KEY if factors[key].notna().any())

    return {
        "ticker": symbol,
        "company_name": str(info.get("longName") or info.get("shortName") or symbol),
        "sector": str(info.get("sector") or ""),
        "industry": str(info.get("industry") or ""),
        "quarters": [str(p) for p in periods],
        "quarter_dates": [d.strftime("%Y-%m-%d") for d in quarter_dates],
        "catalog": FACTOR_CATALOG,
        "series": {key: _serialize_series(factors[key], periods) for key in CATALOG_BY_KEY},
        "diagnostics": diagnostics,
        "price_history_3y": price_history,
        "availability": {
            "price_quarters": len(periods),
            "fundamental_quarters": available_fundamental_quarters,
            "available_factors": available_factor_count,
            "catalog_factors": len(CATALOG_BY_KEY),
        },
        "note": (
            "Price-derived factors use the 3-year market history. Statement-derived factors use the quarterly depth Yahoo Finance exposes. "
            "Forward P/E, bid-ask spread, short interest, news sentiment and EPS revision are current/latest-quarter snapshots. "
            "Revenue revision and target-price revision remain blank when Yahoo does not expose a historical comparison; blanks are intentional rather than fabricated data."
        ),
    }
