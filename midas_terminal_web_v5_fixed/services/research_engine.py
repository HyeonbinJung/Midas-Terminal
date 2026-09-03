from __future__ import annotations

import asyncio
import math
import os
import re
import statistics
import tempfile
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
import requests
import yfinance as yf
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph


BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")
DEFAULT_TEMPLATE = BASE_DIR / "templates" / "Midas Template.docx"
TEMPLATE_ENV = os.getenv("MIDAS_TEMPLATE_PATH", "").strip()

SEC_USER_AGENT = os.getenv(
    "SEC_USER_AGENT",
    "MidasCapitalResearch/1.0 research@example.com",
).strip()
SEC_TIMEOUT_SECONDS = max(5, min(int(os.getenv("SEC_TIMEOUT_SECONDS", "20")), 60))
PEER_CANDIDATE_LIMIT = max(6, min(int(os.getenv("PEER_CANDIDATE_LIMIT", "12")), 30))
SEC_MIN_INTERVAL_SECONDS = max(0.11, min(float(os.getenv("SEC_MIN_INTERVAL_SECONDS", "0.13")), 1.0))
BACKGROUND_CELL_CHARS = max(350, min(int(os.getenv("BACKGROUND_CELL_CHARS", "900")), 1800))
REVENUE_GROWTH_HIGHLIGHT_PCT = max(5.0, min(float(os.getenv("REVENUE_GROWTH_HIGHLIGHT_PCT", "20")), 100.0))
CATALYST_MAX_ROWS = max(1, min(int(os.getenv("CATALYST_MAX_ROWS", "5")), 5))
RISK_MAX_ROWS = max(1, min(int(os.getenv("RISK_MAX_ROWS", "6")), 6))
HIGHLIGHT_MAX_ROWS = max(1, min(int(os.getenv("HIGHLIGHT_MAX_ROWS", "5")), 6))

SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik:010d}.json"
SEC_ARCHIVES_BASE = "https://www.sec.gov/Archives/edgar/data"

# Objective/public-source fields are auto-filled from Yahoo Finance and SEC EDGAR.
# Recommendation / Thesis and all DCF sections remain analyst-owned. Catalysts are generated
# automatically from scheduled events, consensus fundamentals, and SEC-announced milestones.


@dataclass(slots=True)
class MarketData:
    ticker: str
    company_name: str = ""
    exchange: str = ""
    sector: str = ""
    industry: str = ""
    currency: str = ""
    current_price: float | None = None
    high_52w: float | None = None
    low_52w: float | None = None
    avg_daily_value: float | None = None
    market_cap: float | None = None
    shares_outstanding: float | None = None
    float_shares: float | None = None
    beta: float | None = None
    dividend_yield_pct: float | None = None
    annual_income: pd.DataFrame | None = None
    annual_balance: pd.DataFrame | None = None
    annual_cashflow: pd.DataFrame | None = None
    quarterly_income: pd.DataFrame | None = None
    quarterly_balance: pd.DataFrame | None = None
    price_history: pd.DataFrame | None = None
    revenue_estimate: pd.DataFrame | None = None
    earnings_estimate: pd.DataFrame | None = None
    analyst_targets: dict[str, Any] | None = None
    recommendations_summary: pd.DataFrame | None = None
    info: dict[str, Any] | None = None
    major_holders: pd.DataFrame | None = None
    institutional_holders: pd.DataFrame | None = None
    mutualfund_holders: pd.DataFrame | None = None
    insider_roster_holders: pd.DataFrame | None = None
    upgrades_downgrades: pd.DataFrame | None = None
    calendar: Any = None
    selected_peers: list[str] | None = None


@dataclass(slots=True)
class SecResearch:
    cik: int | None = None
    annual_form: str = ""
    annual_filing_date: str = ""
    annual_url: str = ""
    annual_text: str = ""
    proxy_form: str = ""
    proxy_filing_date: str = ""
    proxy_url: str = ""
    proxy_text: str = ""
    event_form: str = ""
    event_filing_date: str = ""
    event_url: str = ""
    event_text: str = ""


@dataclass(slots=True)
class Catalyst:
    catalyst: str = ""
    expected_date: str = ""
    probability: str = ""
    price_impact: str = ""
    verify: str = ""


@dataclass(slots=True)
class RiskFactor:
    risk: str = ""
    likelihood: str = ""
    severity: str = ""
    tp_impact: str = ""
    mitigation: str = ""


@dataclass(slots=True)
class HighlightArgument:
    argument: str = ""
    evidence_source: str = ""


@dataclass(slots=True)
class CompanyBackground:
    founding_listing: str = ""
    segments_revenue_mix: str = ""
    revenue_geography: str = ""
    key_customers_channels: str = ""
    competitors_market_share: str = ""
    cost_structure_supply_chain: str = ""
    ownership_major_shareholders: str = ""
    management_incentives: str = ""


INCOME_ALIASES: dict[str, tuple[str, ...]] = {
    "Revenue": ("Total Revenue", "Operating Revenue", "Revenue"),
    "EBITDA": ("EBITDA", "Normalized EBITDA"),
    "Operating income": ("Operating Income", "Total Operating Income As Reported"),
    "Net income": ("Net Income", "Net Income Common Stockholders", "Net Income Including Noncontrolling Interests"),
    "EPS": ("Diluted EPS", "Basic EPS"),
    "Pretax income": ("Pretax Income", "Income Before Tax"),
    "Tax provision": ("Tax Provision", "Income Tax Expense"),
    "Diluted shares": ("Diluted Average Shares", "Basic Average Shares"),
}

BALANCE_ALIASES: dict[str, tuple[str, ...]] = {
    "Cash and equivalents": (
        "Cash Cash Equivalents And Short Term Investments",
        "Cash And Cash Equivalents",
        "Cash Financial",
        "Cash",
    ),
    "Accounts receivable": ("Accounts Receivable", "Receivables"),
    "Inventory": ("Inventory",),
    "Total current assets": ("Current Assets", "Total Current Assets"),
    "Property, plant and equipment": ("Net PPE", "Property Plant Equipment", "Gross PPE"),
    "Goodwill and intangibles": (
        "Goodwill And Other Intangible Assets",
        "Goodwill",
        "Other Intangible Assets",
    ),
    "Total assets": ("Total Assets",),
    "Accounts payable": ("Accounts Payable", "Payables"),
    "Total current liabilities": ("Current Liabilities", "Total Current Liabilities"),
    "Total debt (incl. leases)": ("Total Debt",),
    "Lease liabilities": (
        "Capital Lease Obligations",
        "Long Term Capital Lease Obligation",
        "Current Capital Lease Obligation",
    ),
    "Total liabilities": ("Total Liabilities Net Minority Interest", "Total Liabilities"),
    "Total equity": (
        "Stockholders Equity",
        "Total Equity Gross Minority Interest",
        "Common Stock Equity",
    ),
    "Minority interest": ("Minority Interest", "Minority Interests"),
    "Preferred equity": ("Preferred Stock Equity", "Preferred Stock"),
    "Ordinary shares": ("Ordinary Shares Number", "Share Issued"),
}

CASHFLOW_ALIASES: dict[str, tuple[str, ...]] = {
    "Free cash flow": ("Free Cash Flow",),
}


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    return result if math.isfinite(result) else None


def _safe_div(a: float | None, b: float | None) -> float | None:
    if a is None or b is None or abs(b) < 1e-12:
        return None
    return a / b


def _first_number(*values: Any) -> float | None:
    for value in values:
        number = _to_float(value)
        if number is not None:
            return number
    return None


def _safe_attr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        value = getattr(obj, name)
        return value if value is not None else default
    except Exception:
        return default


def _safe_mapping_get(mapping: Any, key: str, default: Any = None) -> Any:
    try:
        if mapping is None:
            return default
        if hasattr(mapping, "get"):
            value = mapping.get(key, default)
        else:
            value = mapping[key]
        return value if value is not None else default
    except Exception:
        return default


def _safe_df(value: Any) -> pd.DataFrame:
    return value.copy() if isinstance(value, pd.DataFrame) else pd.DataFrame()


def _get_statement(ticker: yf.Ticker, *names: str) -> pd.DataFrame:
    for name in names:
        try:
            value = getattr(ticker, name)
            if isinstance(value, pd.DataFrame) and not value.empty:
                return value.copy()
        except Exception:
            continue
    return pd.DataFrame()


def _get_optional_df(ticker: yf.Ticker, *method_names: str) -> pd.DataFrame:
    for name in method_names:
        try:
            method = getattr(ticker, name)
            value = method() if callable(method) else method
            if isinstance(value, pd.DataFrame) and not value.empty:
                return value.copy()
        except Exception:
            continue
    return pd.DataFrame()


def _get_calendar(ticker: yf.Ticker) -> Any:
    """Best-effort yfinance calendar fetch across version differences."""
    for name in ("get_calendar", "calendar"):
        try:
            obj = getattr(ticker, name)
            value = obj() if callable(obj) else obj
            if value is not None:
                if isinstance(value, pd.DataFrame):
                    return value.copy()
                if isinstance(value, (dict, pd.Series)):
                    return value.copy() if hasattr(value, "copy") else value
        except Exception:
            continue
    return {}


def _normalize_symbol(raw: str) -> str:
    symbol = raw.strip().upper()
    if not re.fullmatch(r"[A-Z0-9.\-^=]{1,24}", symbol):
        raise ValueError("Ticker 형식이 올바르지 않습니다. 예: TSLA, AAPL, BRK-B, 005930.KS")
    return symbol


def _resolve_template_path() -> Path:
    candidates: list[Path] = []
    if TEMPLATE_ENV:
        candidates.append(Path(TEMPLATE_ENV).expanduser())
    candidates.extend(
        [
            DEFAULT_TEMPLATE,
            BASE_DIR / "Midas Template.docx",
        ]
    )
    for path in candidates:
        if path.exists() and path.is_file():
            return path
    raise FileNotFoundError(
        "Midas Template.docx를 찾지 못했습니다. "
        "templates/Midas Template.docx에 두거나 .env의 MIDAS_TEMPLATE_PATH를 지정하세요."
    )


def _get_info(ticker: yf.Ticker) -> dict[str, Any]:
    try:
        info = ticker.get_info()
        return info if isinstance(info, dict) else {}
    except Exception:
        try:
            info = ticker.info
            return info if isinstance(info, dict) else {}
        except Exception:
            return {}


def _history(ticker: yf.Ticker, period: str) -> pd.DataFrame:
    try:
        df = ticker.history(period=period, auto_adjust=False, actions=False)
        if isinstance(df, pd.DataFrame):
            return df.copy()
    except Exception:
        pass
    return pd.DataFrame()


def _analyst_targets(ticker: yf.Ticker) -> dict[str, Any]:
    for getter in ("get_analyst_price_targets", "analyst_price_targets"):
        try:
            obj = getattr(ticker, getter)
            value = obj() if callable(obj) else obj
            if isinstance(value, dict):
                return value
            if isinstance(value, pd.Series):
                return value.to_dict()
        except Exception:
            continue
    return {}


def fetch_market_data(raw_ticker: str) -> MarketData:
    symbol = _normalize_symbol(raw_ticker)
    ticker = yf.Ticker(symbol)

    info = _get_info(ticker)
    fast = _safe_attr(ticker, "fast_info", {})
    hist_1y = _history(ticker, "1y")
    hist_3y = _history(ticker, "3y")

    current_from_hist = None
    if not hist_1y.empty and "Close" in hist_1y.columns:
        current_from_hist = _to_float(hist_1y["Close"].dropna().iloc[-1]) if not hist_1y["Close"].dropna().empty else None

    current_price = _first_number(
        info.get("currentPrice"),
        info.get("regularMarketPrice"),
        _safe_mapping_get(fast, "last_price"),
        current_from_hist,
    )

    market_cap = _first_number(
        info.get("marketCap"),
        _safe_mapping_get(fast, "market_cap"),
    )

    shares = _first_number(
        info.get("sharesOutstanding"),
        _safe_mapping_get(fast, "shares"),
    )

    if market_cap is None and current_price is not None and shares is not None:
        market_cap = current_price * shares

    high_52w = _first_number(info.get("fiftyTwoWeekHigh"), _safe_mapping_get(fast, "year_high"))
    low_52w = _first_number(info.get("fiftyTwoWeekLow"), _safe_mapping_get(fast, "year_low"))

    if not hist_1y.empty and "High" in hist_1y.columns and high_52w is None:
        high_52w = _to_float(hist_1y["High"].max())
    if not hist_1y.empty and "Low" in hist_1y.columns and low_52w is None:
        low_52w = _to_float(hist_1y["Low"].min())

    avg_daily_value = None
    if not hist_1y.empty and {"Close", "Volume"}.issubset(hist_1y.columns):
        recent = hist_1y[["Close", "Volume"]].dropna().tail(30)
        if not recent.empty:
            avg_daily_value = _to_float((recent["Close"] * recent["Volume"]).mean())

    dividend_yield = _to_float(info.get("dividendYield"))
    if dividend_yield is not None:
        dividend_yield = dividend_yield * 100 if abs(dividend_yield) <= 1 else dividend_yield

    annual_income = _get_statement(ticker, "income_stmt", "financials")
    annual_balance = _get_statement(ticker, "balance_sheet")
    annual_cashflow = _get_statement(ticker, "cashflow", "cash_flow")
    quarterly_income = _get_statement(ticker, "quarterly_income_stmt", "quarterly_financials")
    quarterly_balance = _get_statement(ticker, "quarterly_balance_sheet")

    revenue_estimate = _get_optional_df(ticker, "get_revenue_estimate", "revenue_estimate")
    earnings_estimate = _get_optional_df(ticker, "get_earnings_estimate", "earnings_estimate")
    recommendations_summary = _get_optional_df(ticker, "get_recommendations_summary", "recommendations_summary")
    major_holders = _get_optional_df(ticker, "get_major_holders", "major_holders")
    institutional_holders = _get_optional_df(ticker, "get_institutional_holders", "institutional_holders")
    mutualfund_holders = _get_optional_df(ticker, "get_mutualfund_holders", "mutualfund_holders")
    insider_roster_holders = _get_optional_df(ticker, "get_insider_roster_holders", "insider_roster_holders")
    upgrades_downgrades = _get_optional_df(ticker, "get_upgrades_downgrades", "upgrades_downgrades")
    calendar = _get_calendar(ticker)

    company_name = str(info.get("longName") or info.get("shortName") or symbol)
    exchange = str(info.get("fullExchangeName") or info.get("exchange") or _safe_mapping_get(fast, "exchange", ""))
    sector = str(info.get("sector") or "")
    industry = str(info.get("industry") or "")
    currency = str(info.get("financialCurrency") or info.get("currency") or _safe_mapping_get(fast, "currency", ""))

    has_any_data = any(
        [
            current_price is not None,
            market_cap is not None,
            not annual_income.empty,
            not annual_balance.empty,
        ]
    )
    if not has_any_data:
        raise ValueError(f"{symbol} 데이터를 Yahoo Finance에서 찾지 못했습니다.")

    return MarketData(
        ticker=symbol,
        company_name=company_name,
        exchange=exchange,
        sector=sector,
        industry=industry,
        currency=currency,
        current_price=current_price,
        high_52w=high_52w,
        low_52w=low_52w,
        avg_daily_value=avg_daily_value,
        market_cap=market_cap,
        shares_outstanding=shares,
        float_shares=_to_float(info.get("floatShares")),
        beta=_to_float(info.get("beta")),
        dividend_yield_pct=dividend_yield,
        annual_income=annual_income,
        annual_balance=annual_balance,
        annual_cashflow=annual_cashflow,
        quarterly_income=quarterly_income,
        quarterly_balance=quarterly_balance,
        price_history=hist_3y if not hist_3y.empty else hist_1y,
        revenue_estimate=revenue_estimate,
        earnings_estimate=earnings_estimate,
        analyst_targets=_analyst_targets(ticker),
        recommendations_summary=recommendations_summary,
        info=info,
        major_holders=major_holders,
        institutional_holders=institutional_holders,
        mutualfund_holders=mutualfund_holders,
        insider_roster_holders=insider_roster_holders,
        upgrades_downgrades=upgrades_downgrades,
        calendar=calendar,
    )


def fetch_peer_data(raw_ticker: str) -> MarketData:
    """Fetch only the fields needed to score and populate peer-comparison tables.

    Unlike fetch_market_data(), this intentionally skips quarterly statements, analyst
    estimates, holders, recommendations, and other expensive endpoints.
    """
    symbol = _normalize_symbol(raw_ticker)
    ticker = yf.Ticker(symbol)
    info = _get_info(ticker)
    fast = _safe_attr(ticker, "fast_info", {})

    annual_income = _get_statement(ticker, "income_stmt", "financials")
    annual_balance = _get_statement(ticker, "balance_sheet")
    history = _history(ticker, "5y")

    current_from_hist = None
    if not history.empty and "Close" in history.columns:
        closes = history["Close"].dropna()
        if not closes.empty:
            current_from_hist = _to_float(closes.iloc[-1])

    current_price = _first_number(
        info.get("currentPrice"),
        info.get("regularMarketPrice"),
        _safe_mapping_get(fast, "last_price"),
        current_from_hist,
    )
    shares = _first_number(info.get("sharesOutstanding"), _safe_mapping_get(fast, "shares"))
    market_cap = _first_number(info.get("marketCap"), _safe_mapping_get(fast, "market_cap"))
    if market_cap is None and current_price is not None and shares is not None:
        market_cap = current_price * shares

    has_any_data = any((market_cap is not None, not annual_income.empty, not annual_balance.empty))
    if not has_any_data:
        raise ValueError(f"{symbol} peer fundamentals unavailable")

    return MarketData(
        ticker=symbol,
        company_name=str(info.get("longName") or info.get("shortName") or symbol),
        exchange=str(info.get("fullExchangeName") or info.get("exchange") or _safe_mapping_get(fast, "exchange", "")),
        sector=str(info.get("sector") or ""),
        industry=str(info.get("industry") or ""),
        currency=str(info.get("financialCurrency") or info.get("currency") or _safe_mapping_get(fast, "currency", "")),
        current_price=current_price,
        market_cap=market_cap,
        shares_outstanding=shares,
        beta=_to_float(info.get("beta")),
        annual_income=annual_income,
        annual_balance=annual_balance,
        price_history=history,
        info=info,
    )


def _format_price(value: float | None) -> str:
    if value is None:
        return ""
    decimals = 4 if abs(value) < 1 else 2
    return f"{value:,.{decimals}f}"


def _format_compact_money(value: float | None, currency: str = "") -> str:
    if value is None:
        return ""
    sign = "-" if value < 0 else ""
    n = abs(value)
    if n >= 1e12:
        text = f"{n / 1e12:.2f}tn"
    elif n >= 1e9:
        text = f"{n / 1e9:.2f}bn"
    elif n >= 1e6:
        text = f"{n / 1e6:.1f}m"
    else:
        text = f"{n:,.0f}"
    return f"{sign}{text}{(' ' + currency) if currency else ''}"


def _format_millions(value: float | None) -> str:
    if value is None:
        return ""
    scaled = value / 1_000_000
    if abs(scaled) >= 100:
        return f"{scaled:,.0f}"
    if abs(scaled) >= 10:
        return f"{scaled:,.1f}"
    return f"{scaled:,.2f}"


def _format_shares_m(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value / 1_000_000:,.1f}m"


def _format_pct_ratio(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value * 100:.1f}%"


def _format_pct_value(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.2f}%"


def _format_ratio(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.2f}x"


def _format_number(value: float | None, decimals: int = 2) -> str:
    if value is None:
        return ""
    return f"{value:,.{decimals}f}"


def _all_tables(parent: DocumentObject | _Cell) -> Iterable[Table]:
    for table in parent.tables:
        yield table
        for row in table.rows:
            for cell in row.cells:
                yield from _all_tables(cell)


def _set_paragraph_text(paragraph: Paragraph, text: str) -> None:
    if not text:
        return
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def _set_cell_text(cell: _Cell, text: str) -> None:
    if text == "":
        return
    paragraph = cell.paragraphs[0]
    _set_paragraph_text(paragraph, text)
    for extra in cell.paragraphs[1:]:
        for run in extra.runs:
            run.text = ""


def _find_table_by_first_cell(doc: DocumentObject, first_cell: str) -> Table | None:
    needle = first_cell.strip().lower()
    for table in _all_tables(doc):
        if not table.rows or not table.rows[0].cells:
            continue
        text = table.rows[0].cells[0].text.strip().lower()
        if text == needle:
            return table
    return None


def _find_table_containing_row(doc: DocumentObject, label: str) -> Table | None:
    needle = label.strip().lower()
    for table in _all_tables(doc):
        for row in table.rows:
            if row.cells and row.cells[0].text.strip().lower() == needle:
                return table
    return None


def _row_map(table: Table) -> dict[str, int]:
    result: dict[str, int] = {}
    for i, row in enumerate(table.rows):
        if row.cells:
            result[row.cells[0].text.strip().lower()] = i
    return result


def _column_for_year(df: pd.DataFrame, year: int) -> Any | None:
    if df is None or df.empty:
        return None
    matches: list[Any] = []
    for col in df.columns:
        try:
            ts = pd.Timestamp(col)
            if ts.year == year:
                matches.append(col)
        except Exception:
            continue
    if not matches:
        return None
    return max(matches, key=lambda c: pd.Timestamp(c))


def _line_value(df: pd.DataFrame | None, aliases: Iterable[str], column: Any) -> float | None:
    if df is None or df.empty or column is None or column not in df.columns:
        return None

    normalized_index = {str(idx).strip().lower(): idx for idx in df.index}
    for alias in aliases:
        actual = normalized_index.get(alias.strip().lower())
        if actual is not None:
            try:
                return _to_float(df.at[actual, column])
            except Exception:
                try:
                    return _to_float(df.loc[actual, column])
                except Exception:
                    pass
    return None


def _annual_value(df: pd.DataFrame | None, aliases: Iterable[str], year: int) -> float | None:
    if df is None or df.empty:
        return None
    return _line_value(df, aliases, _column_for_year(df, year))


def _combined_annual_value(df: pd.DataFrame | None, alias_groups: Iterable[Iterable[str]], year: int) -> float | None:
    values: list[float] = []
    for aliases in alias_groups:
        value = _annual_value(df, aliases, year)
        if value is not None:
            values.append(value)
    return sum(values) if values else None


def _price_at_or_before(history: pd.DataFrame | None, date_value: Any) -> float | None:
    if history is None or history.empty or "Close" not in history.columns:
        return None
    try:
        target = pd.Timestamp(date_value)
        idx = pd.to_datetime(history.index)
        if getattr(idx, "tz", None) is not None:
            idx = idx.tz_localize(None)
        target = target.tz_localize(None) if target.tzinfo is not None else target
        closes = pd.Series(history["Close"].values, index=idx).dropna()
        closes = closes[closes.index <= target]
        return _to_float(closes.iloc[-1]) if not closes.empty else None
    except Exception:
        return None


def _quarter_columns(data: MarketData) -> list[Any]:
    cols: set[Any] = set()
    for df in (data.quarterly_income, data.quarterly_balance):
        if df is None or df.empty:
            continue
        for col in df.columns:
            try:
                cols.add(pd.Timestamp(col))
            except Exception:
                continue
    return sorted(cols)[-6:]


def _quarter_label(col: Any) -> str:
    ts = pd.Timestamp(col)
    quarter = (ts.month - 1) // 3 + 1
    return f"{quarter}Q{ts.year}"


def _quarter_value(df: pd.DataFrame | None, aliases: Iterable[str], date_value: Any) -> float | None:
    if df is None or df.empty:
        return None
    target = pd.Timestamp(date_value)
    for col in df.columns:
        try:
            if pd.Timestamp(col) == target:
                return _line_value(df, aliases, col)
        except Exception:
            continue
    return None


def _ltm_sum(df: pd.DataFrame | None, aliases: Iterable[str], ending: Any) -> float | None:
    if df is None or df.empty:
        return None
    end_ts = pd.Timestamp(ending)
    cols: list[Any] = []
    for col in df.columns:
        try:
            ts = pd.Timestamp(col)
            if ts <= end_ts:
                cols.append(col)
        except Exception:
            continue
    cols.sort(key=lambda c: pd.Timestamp(c), reverse=True)
    values: list[float] = []
    for col in cols:
        value = _line_value(df, aliases, col)
        if value is not None:
            values.append(value)
        if len(values) == 4:
            break
    return sum(values) if len(values) == 4 else None


def _effective_tax_rate(data: MarketData, year: int) -> float | None:
    pretax = _annual_value(data.annual_income, INCOME_ALIASES["Pretax income"], year)
    tax = _annual_value(data.annual_income, INCOME_ALIASES["Tax provision"], year)
    rate = _safe_div(tax, pretax)
    if rate is None or rate < 0 or rate > 0.60:
        return None
    return rate


def _roic(data: MarketData, year: int) -> float | None:
    op_income = _annual_value(data.annual_income, INCOME_ALIASES["Operating income"], year)
    debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], year)
    equity = _annual_value(data.annual_balance, BALANCE_ALIASES["Total equity"], year)
    cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], year)
    if op_income is None or debt is None or equity is None:
        return None
    tax_rate = _effective_tax_rate(data, year)
    nopat = op_income * (1 - tax_rate) if tax_rate is not None else op_income
    invested = debt + equity - (cash or 0.0)
    return _safe_div(nopat, invested)


def _roe(data: MarketData, year: int) -> float | None:
    ni = _annual_value(data.annual_income, INCOME_ALIASES["Net income"], year)
    equity = _annual_value(data.annual_balance, BALANCE_ALIASES["Total equity"], year)
    prev_equity = _annual_value(data.annual_balance, BALANCE_ALIASES["Total equity"], year - 1)
    if ni is None or equity is None:
        return None
    avg_equity = (equity + prev_equity) / 2 if prev_equity is not None else equity
    return _safe_div(ni, avg_equity)


def _estimate_avg(df: pd.DataFrame | None, period: str) -> float | None:
    if df is None or df.empty:
        return None
    candidates = [period, period.lower(), period.upper()]
    for idx in candidates:
        if idx in df.index:
            for col in ("avg", "Avg", "average", "Average"):
                if col in df.columns:
                    return _to_float(df.at[idx, col])
    # Some versions expose Period as a normal column.
    for period_col in ("period", "Period"):
        if period_col in df.columns:
            rows = df[df[period_col].astype(str).str.lower() == period.lower()]
            if not rows.empty:
                for col in ("avg", "Avg", "average", "Average"):
                    if col in rows.columns:
                        return _to_float(rows.iloc[0][col])
    return None


_SEC_TICKER_CACHE: dict[str, int] | None = None
_SEC_RATE_LOCK = threading.Lock()
_SEC_LAST_REQUEST_AT = 0.0


def _sec_wait_for_slot() -> None:
    """Serialize SEC requests across all worker threads to respect fair-access limits."""
    global _SEC_LAST_REQUEST_AT
    with _SEC_RATE_LOCK:
        now = time.monotonic()
        wait_for = SEC_MIN_INTERVAL_SECONDS - (now - _SEC_LAST_REQUEST_AT)
        if wait_for > 0:
            time.sleep(wait_for)
        _SEC_LAST_REQUEST_AT = time.monotonic()


def _sec_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": SEC_USER_AGENT,
            "Accept-Encoding": "gzip, deflate",
            "Accept": "application/json,text/html,application/xhtml+xml,*/*",
        }
    )
    return session


def _sec_get_json(session: requests.Session, url: str) -> dict[str, Any]:
    _sec_wait_for_slot()
    response = session.get(url, timeout=SEC_TIMEOUT_SECONDS)
    response.raise_for_status()
    value = response.json()
    return value if isinstance(value, dict) else {}


def _sec_get_text(session: requests.Session, url: str) -> str:
    _sec_wait_for_slot()
    response = session.get(url, timeout=SEC_TIMEOUT_SECONDS)
    response.raise_for_status()
    return response.text


def _sec_ticker_map(session: requests.Session) -> dict[str, int]:
    global _SEC_TICKER_CACHE
    if _SEC_TICKER_CACHE is not None:
        return _SEC_TICKER_CACHE

    mapping: dict[str, int] = {}
    try:
        payload = _sec_get_json(session, SEC_TICKERS_URL)
        for item in payload.values():
            if not isinstance(item, dict):
                continue
            ticker = str(item.get("ticker") or "").upper().strip()
            cik = item.get("cik_str")
            if ticker and cik is not None:
                try:
                    mapping[ticker] = int(cik)
                except (TypeError, ValueError):
                    continue
    except Exception as exc:
        print(f"[SEC] ticker map unavailable: {type(exc).__name__}: {exc}")

    _SEC_TICKER_CACHE = mapping
    return mapping


def _filing_url(cik: int, accession: str, primary_document: str) -> str:
    accession_clean = accession.replace("-", "")
    return f"{SEC_ARCHIVES_BASE}/{cik}/{accession_clean}/{primary_document}"


def _latest_filing_from_submissions(
    submissions: dict[str, Any],
    accepted_forms: tuple[str, ...],
) -> tuple[str, str, str, str] | None:
    recent = _safe_mapping_get(_safe_mapping_get(submissions, "filings", {}), "recent", {})
    if not isinstance(recent, dict):
        return None

    forms = recent.get("form") or []
    accession_numbers = recent.get("accessionNumber") or []
    primary_documents = recent.get("primaryDocument") or []
    filing_dates = recent.get("filingDate") or []

    size = min(len(forms), len(accession_numbers), len(primary_documents), len(filing_dates))
    for idx in range(size):
        form = str(forms[idx] or "").upper().strip()
        if form not in accepted_forms:
            continue
        accession = str(accession_numbers[idx] or "").strip()
        primary = str(primary_documents[idx] or "").strip()
        filing_date = str(filing_dates[idx] or "").strip()
        if accession and primary:
            return form, filing_date, accession, primary
    return None


def _filing_html_to_text(html_text: str) -> str:
    if not html_text:
        return ""
    soup = BeautifulSoup(html_text, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()
    text = soup.get_text(" ", strip=True)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def fetch_sec_research(raw_ticker: str) -> SecResearch:
    symbol = _normalize_symbol(raw_ticker).split(".", 1)[0]
    result = SecResearch()
    session = _sec_session()

    try:
        cik = _sec_ticker_map(session).get(symbol)
        if cik is None:
            return result
        result.cik = cik

        submissions = _sec_get_json(session, SEC_SUBMISSIONS_URL.format(cik=cik))
        annual = _latest_filing_from_submissions(submissions, ("10-K", "10-K/A", "20-F", "20-F/A", "40-F", "40-F/A"))
        proxy = _latest_filing_from_submissions(submissions, ("DEF 14A", "DEF 14A/A"))
        event = _latest_filing_from_submissions(submissions, ("8-K", "8-K/A", "6-K", "6-K/A"))

        if annual:
            form, filing_date, accession, primary = annual
            result.annual_form = form
            result.annual_filing_date = filing_date
            result.annual_url = _filing_url(cik, accession, primary)
            try:
                result.annual_text = _filing_html_to_text(_sec_get_text(session, result.annual_url))
            except Exception as exc:
                print(f"[SEC] annual filing text unavailable: {type(exc).__name__}: {exc}")

        if proxy:
            form, filing_date, accession, primary = proxy
            result.proxy_form = form
            result.proxy_filing_date = filing_date
            result.proxy_url = _filing_url(cik, accession, primary)
            try:
                result.proxy_text = _filing_html_to_text(_sec_get_text(session, result.proxy_url))
            except Exception as exc:
                print(f"[SEC] proxy text unavailable: {type(exc).__name__}: {exc}")

        if event:
            form, filing_date, accession, primary = event
            result.event_form = form
            result.event_filing_date = filing_date
            result.event_url = _filing_url(cik, accession, primary)
            try:
                result.event_text = _filing_html_to_text(_sec_get_text(session, result.event_url))
            except Exception as exc:
                print(f"[SEC] event filing text unavailable: {type(exc).__name__}: {exc}")

    except Exception as exc:
        print(f"[SEC] research unavailable for {symbol}: {type(exc).__name__}: {exc}")

    return result


def _sentences(text: str) -> list[str]:
    if not text:
        return []
    normalized = re.sub(r"\s+", " ", text).strip()
    parts = re.split(r"(?<=[.!?])\s+(?=[A-Z0-9(])", normalized)
    return [p.strip() for p in parts if 45 <= len(p.strip()) <= 900]


def _keyword_extract(
    text: str,
    keywords: tuple[str, ...],
    *,
    max_sentences: int = 4,
    max_chars: int | None = None,
) -> str:
    if not text:
        return ""
    max_chars = max_chars or BACKGROUND_CELL_CHARS
    keys = tuple(k.lower() for k in keywords)
    scored: list[tuple[int, int, str]] = []
    seen: set[str] = set()

    for idx, sentence in enumerate(_sentences(text)):
        lower = sentence.lower()
        score = sum(1 for keyword in keys if keyword in lower)
        if score <= 0:
            continue
        normalized = re.sub(r"[^a-z0-9]+", "", lower)[:220]
        if normalized in seen:
            continue
        seen.add(normalized)
        scored.append((score, -idx, sentence))

    scored.sort(reverse=True)
    chosen = [item[2] for item in scored[: max_sentences * 3]]
    # Restore filing order among the highest-scoring candidates.
    chosen.sort(key=lambda sentence: text.find(sentence))
    chosen = chosen[:max_sentences]

    output = " ".join(chosen).strip()
    if len(output) > max_chars:
        output = output[: max_chars - 1].rstrip() + "…"
    return output


def _business_section(text: str) -> str:
    if not text:
        return ""
    patterns = (
        (r"ITEM\s+1[.:\s-]+BUSINESS", r"ITEM\s+1A[.:\s-]+RISK\s+FACTORS"),
        (r"ITEM\s+4[.:\s-]+INFORMATION\s+ON\s+THE\s+COMPANY", r"ITEM\s+5[.:\s-]+OPERATING"),
    )
    candidates: list[str] = []
    for start_pattern, end_pattern in patterns:
        starts = list(re.finditer(start_pattern, text, flags=re.IGNORECASE))
        for start in starts:
            tail = text[start.end() :]
            end = re.search(end_pattern, tail, flags=re.IGNORECASE)
            section = tail[: end.start()] if end else tail[:160_000]
            section = section.strip()
            # TOC matches are usually tiny; the real Item section is much longer.
            if len(section) >= 1_200:
                candidates.append(section[:180_000])
    if candidates:
        return max(candidates, key=len)
    return text[:180_000]


def _risk_section(text: str) -> str:
    """Best-effort extraction of Item 1A / Risk Factors from annual filings."""
    if not text:
        return ""
    patterns = (
        (r"ITEM\s+1A[.:\s-]+RISK\s+FACTORS", r"ITEM\s+1B[.:\s-]+"),
        (r"ITEM\s+1A[.:\s-]+RISK\s+FACTORS", r"ITEM\s+2[.:\s-]+PROPERTIES"),
        (r"RISK\s+FACTORS", r"ITEM\s+4[.:\s-]+INFORMATION\s+ON\s+THE\s+COMPANY"),
    )
    candidates: list[str] = []
    for start_pattern, end_pattern in patterns:
        starts = list(re.finditer(start_pattern, text, flags=re.IGNORECASE))
        for start in starts:
            tail = text[start.end():]
            end = re.search(end_pattern, tail, flags=re.IGNORECASE)
            section = (tail[:end.start()] if end else tail[:220_000]).strip()
            if len(section) >= 1_500:
                candidates.append(section[:220_000])
    if candidates:
        return max(candidates, key=len)
    # Some foreign/private issuers use non-standard headings. Keep a bounded fallback.
    hit = re.search(r"risk\s+factors", text, flags=re.IGNORECASE)
    if hit:
        return text[hit.end():hit.end()+180_000]
    return ""


def _proxy_compensation_section(text: str) -> str:
    if not text:
        return ""
    matches = list(
        re.finditer(
            r"executive\s+compensation|compensation\s+discussion|incentive\s+compensation",
            text,
            flags=re.IGNORECASE,
        )
    )
    if matches:
        # Prefer a later occurrence to avoid a table-of-contents hit.
        match = matches[-1] if len(matches) > 1 else matches[0]
        return text[match.start() : match.start() + 80_000]
    return text[:120_000]


def _format_listing_history(data: MarketData, business_text: str) -> str:
    parts: list[str] = []
    founding = _keyword_extract(
        business_text,
        ("founded", "incorporated", "organized", "formed", "established"),
        max_sentences=2,
        max_chars=520,
    )
    if founding:
        parts.append(founding)

    info = data.info or {}
    first_trade = info.get("firstTradeDateEpochUtc") or info.get("firstTradeDate")
    first_trade_text = ""
    try:
        if first_trade is not None:
            ts = float(first_trade)
            first_trade_text = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")
    except (TypeError, ValueError, OSError):
        pass

    listing = f"Listed as {data.ticker}"
    if data.exchange:
        listing += f" on {data.exchange}"
    if first_trade_text:
        listing += f"; Yahoo Finance first-trade date: {first_trade_text}"
    listing += "."
    parts.append(listing)

    return " ".join(parts)[:BACKGROUND_CELL_CHARS]


def _format_holder_table(df: pd.DataFrame | None, limit: int = 5) -> list[str]:
    if df is None or df.empty:
        return []
    frame = df.copy().reset_index(drop=False)
    lower_cols = {str(col).lower(): col for col in frame.columns}
    holder_col = next((col for key, col in lower_cols.items() if "holder" in key or key in ("name", "organization")), None)
    pct_col = next((col for key, col in lower_cols.items() if "pct" in key or "percent" in key), None)
    shares_col = next((col for key, col in lower_cols.items() if "share" in key and "pct" not in key), None)
    value_col = next((col for key, col in lower_cols.items() if key == "value" or "market value" in key), None)

    if holder_col is None:
        return []

    lines: list[str] = []
    for _, row in frame.head(limit).iterrows():
        holder = str(row.get(holder_col, "") or "").strip()
        if not holder or holder.lower() == "nan":
            continue
        details: list[str] = []
        if pct_col is not None:
            pct = _to_float(row.get(pct_col))
            if pct is not None:
                if abs(pct) <= 1:
                    pct *= 100
                details.append(f"{pct:.2f}%")
        if shares_col is not None:
            shares = _to_float(row.get(shares_col))
            if shares is not None:
                details.append(f"{shares / 1_000_000:,.1f}m shares")
        if value_col is not None and not details:
            value = _to_float(row.get(value_col))
            if value is not None:
                details.append(_format_compact_money(value))
        lines.append(holder + (f" ({', '.join(details)})" if details else ""))
    return lines


def _format_major_holders(data: MarketData) -> str:
    parts: list[str] = []
    institutional = _format_holder_table(data.institutional_holders, limit=5)
    if institutional:
        parts.append("Top reported institutional holders: " + "; ".join(institutional) + ".")

    major = data.major_holders
    if major is not None and not major.empty:
        flat: list[str] = []
        for _, row in major.reset_index(drop=False).head(6).iterrows():
            vals = [str(v).strip() for v in row.tolist() if str(v).strip() not in ("", "nan", "None")]
            if vals:
                flat.append(" ".join(vals))
        if flat:
            parts.append("Ownership summary: " + "; ".join(flat) + ".")

    return " ".join(parts)[:BACKGROUND_CELL_CHARS]


def _format_management(data: MarketData, proxy_text: str) -> str:
    info = data.info or {}
    officers = info.get("companyOfficers") or []
    lines: list[str] = []
    if isinstance(officers, list):
        for officer in officers[:6]:
            if not isinstance(officer, dict):
                continue
            name = str(officer.get("name") or "").strip()
            title = str(officer.get("title") or "").strip()
            if name and title:
                lines.append(f"{name} — {title}")
    officer_text = "Management: " + "; ".join(lines) + "." if lines else ""

    incentives = _keyword_extract(
        _proxy_compensation_section(proxy_text),
        (
            "annual incentive",
            "long-term incentive",
            "equity award",
            "performance share",
            "stock award",
            "bonus",
            "executive compensation",
            "compensation program",
        ),
        max_sentences=3,
        max_chars=620,
    )
    result = " ".join(x for x in (officer_text, incentives) if x)
    return result[:BACKGROUND_CELL_CHARS]


def build_company_background(data: MarketData, sec: SecResearch) -> CompanyBackground:
    info = data.info or {}
    business = _business_section(sec.annual_text)
    fallback_summary = str(info.get("longBusinessSummary") or "").strip()
    source_text = business or fallback_summary

    segments = _keyword_extract(
        source_text,
        ("reportable segment", "operating segment", "segments", "segment revenue", "segment net sales", "revenue mix"),
        max_sentences=4,
    )
    if not segments and fallback_summary:
        segments = fallback_summary[:BACKGROUND_CELL_CHARS]

    geography = _keyword_extract(
        source_text,
        ("geographic", "international", "foreign", "outside the united states", "americas", "europe", "china", "asia", "country"),
        max_sentences=4,
    )
    customers = _keyword_extract(
        source_text,
        ("customer", "customers", "distributor", "reseller", "channel", "direct sales", "retail", "wholesale"),
        max_sentences=4,
    )
    competitors = _keyword_extract(
        source_text,
        ("competition", "competitive", "competitor", "compete", "market share"),
        max_sentences=4,
    )
    if competitors and "market share" not in competitors.lower():
        competitors += " Public filings do not necessarily disclose third-party market-share estimates."
        competitors = competitors[:BACKGROUND_CELL_CHARS]

    costs = _keyword_extract(
        source_text,
        ("supply chain", "supplier", "suppliers", "manufacturing", "contract manufacturer", "raw material", "cost of sales", "cost of revenue", "procurement"),
        max_sentences=4,
    )

    ownership = _format_major_holders(data)
    management = _format_management(data, sec.proxy_text)

    return CompanyBackground(
        founding_listing=_format_listing_history(data, source_text),
        segments_revenue_mix=segments or "Not reliably disclosed by the available free structured APIs; review the latest annual filing.",
        revenue_geography=geography or "Not reliably disclosed by the available free structured APIs; review the latest annual filing.",
        key_customers_channels=customers or "No material customer/channel disclosure was reliably extracted from the latest annual filing.",
        competitors_market_share=competitors or "No reliable competitor/market-share disclosure was extracted from the latest annual filing.",
        cost_structure_supply_chain=costs or "No reliable supply-chain/cost-structure disclosure was extracted from the latest annual filing.",
        ownership_major_shareholders=ownership or "Ownership data was not available from Yahoo Finance at build time.",
        management_incentives=management or "Management/incentive data was not available from Yahoo Finance or the latest proxy at build time.",
    )



def _calendar_values(calendar: Any, *keys: str) -> list[Any]:
    if calendar is None:
        return []
    wanted = {key.lower() for key in keys}
    values: list[Any] = []
    try:
        if isinstance(calendar, dict):
            for key, value in calendar.items():
                if str(key).strip().lower() in wanted:
                    if isinstance(value, (list, tuple, set, pd.Index)):
                        values.extend(list(value))
                    else:
                        values.append(value)
        elif isinstance(calendar, pd.Series):
            for key, value in calendar.items():
                if str(key).strip().lower() in wanted:
                    values.append(value)
        elif isinstance(calendar, pd.DataFrame):
            for idx in calendar.index:
                if str(idx).strip().lower() in wanted:
                    values.extend(calendar.loc[idx].tolist())
            for col in calendar.columns:
                if str(col).strip().lower() in wanted:
                    values.extend(calendar[col].tolist())
    except Exception:
        return values
    return [value for value in values if value is not None]


def _future_calendar_date(calendar: Any, *keys: str) -> datetime | None:
    now = datetime.now(timezone.utc)
    candidates: list[datetime] = []
    for raw in _calendar_values(calendar, *keys):
        if isinstance(raw, (list, tuple, set, pd.Index)):
            raws = list(raw)
        else:
            raws = [raw]
        for value in raws:
            try:
                ts = pd.Timestamp(value)
                if pd.isna(ts):
                    continue
                if ts.tzinfo is None:
                    ts = ts.tz_localize("UTC")
                else:
                    ts = ts.tz_convert("UTC")
                dt = ts.to_pydatetime()
                if dt >= now:
                    candidates.append(dt)
            except Exception:
                continue
    return min(candidates) if candidates else None


def _event_move_range(data: MarketData) -> tuple[float, float]:
    history = data.price_history
    if history is None or history.empty or "Close" not in history.columns:
        beta = abs(data.beta or 1.0)
        return (max(2.0, min(5.0, 2.0 * beta)), max(4.0, min(10.0, 5.0 * beta)))
    try:
        returns = history["Close"].dropna().pct_change().dropna().abs() * 100.0
        if returns.empty:
            raise ValueError
        q75 = float(returns.quantile(0.75))
        q95 = float(returns.quantile(0.95))
        low = max(2.0, min(12.0, q75 * 1.8))
        high = max(low + 1.0, min(25.0, q95 * 1.5))
        return round(low, 1), round(high, 1)
    except Exception:
        return (3.0, 7.0)


def _impact_text(low: float, high: float, direction: int = 0) -> str:
    low = abs(low)
    high = max(abs(high), low)
    prefix = "+" if direction > 0 else "-" if direction < 0 else "±"
    return f"{prefix}{low:.1f}–{high:.1f}%"


def _extract_expected_window(sentence: str) -> str:
    text = re.sub(r"\s+", " ", sentence or "").strip()
    current_year = datetime.now().year
    patterns = (
        r"\b(?:Q[1-4]\s*20\d{2}|[1-4]Q\s*20\d{2})\b",
        r"\b(?:first|second)\s+half\s+(?:of\s+)?20\d{2}\b",
        r"\b(?:FY|fiscal\s+year)\s*20\d{2}\b",
        r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+20\d{2}\b",
        r"\b20\d{2}\b",
    )
    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            candidate = match.group(0)
            year_match = re.search(r"20\d{2}", candidate)
            if year_match and int(year_match.group(0)) >= current_year:
                return candidate
    return "TBD / per filing"


def _sec_milestone_catalyst(sec: SecResearch, data: MarketData) -> Catalyst | None:
    source_text = sec.event_text or sec.annual_text
    if not source_text:
        return None
    event_sentences = _sentences(source_text[:180_000])
    keyword_groups: list[tuple[str, tuple[str, ...]]] = [
        ("Strategic transaction / closing milestone", ("acquisition", "merger", "business combination", "transaction", "closing")),
        ("Regulatory approval / decision", ("regulatory approval", "fda", "approval", "regulator")),
        ("Product launch / commercialization milestone", ("launch", "commercialization", "commercialize", "new product")),
        ("Capacity expansion / facility milestone", ("capacity", "facility", "construction", "plant", "manufacturing expansion")),
        ("Geographic / channel expansion milestone", ("expand", "expansion", "new market", "international", "distribution agreement")),
    ]
    forward_terms = ("expected to", "expects to", "plan to", "plans to", "intend to", "intends to", "anticipate", "will ")
    best: tuple[int, str, str] | None = None
    for idx, sentence in enumerate(event_sentences):
        lower = sentence.lower()
        forward = any(term in lower for term in forward_terms)
        for title, keywords in keyword_groups:
            hits = sum(1 for keyword in keywords if keyword in lower)
            if not hits:
                continue
            future_window = _extract_expected_window(sentence)
            score = hits * 3 + (3 if forward else 0) + (3 if future_window != "TBD / per filing" else 0) - min(idx // 60, 3)
            if best is None or score > best[0]:
                best = (score, title, sentence)
    if best is None or best[0] < 4:
        return None
    _, title, sentence = best
    expected = _extract_expected_window(sentence)
    lower_sentence = sentence.lower()
    probability = "75%" if "expected to" in lower_sentence or "expects to" in lower_sentence or "will " in lower_sentence else "65%"
    low, high = _event_move_range(data)
    source_label = sec.event_form or sec.annual_form or "SEC filing"
    source_date = sec.event_filing_date or sec.annual_filing_date
    verify = (
        f"Verify against {source_label} {source_date} and the next company update; "
        f"confirm the announced milestone, timing and any disclosed financial/KPI impact."
    )
    return Catalyst(title, expected, probability, _impact_text(low, high, 0), verify[:360])


def _target_price_value(data: MarketData, *keys: str) -> float | None:
    targets = data.analyst_targets or {}
    lowered = {str(k).lower(): v for k, v in targets.items()}
    for key in keys:
        value = _to_float(lowered.get(key.lower()))
        if value is not None:
            return value
    return None


def build_catalysts(data: MarketData, sec: SecResearch) -> list[Catalyst]:
    """Generate factual, rule-based catalyst rows without an LLM/API key.

    Probability and price-impact fields are model estimates, not facts. Verification text
    explicitly tells the analyst which reported data or primary filing confirms the event.
    """
    catalysts: list[Catalyst] = []
    low_move, high_move = _event_move_range(data)

    # 1) Company-specific announced milestone from the most recent 8-K/6-K or annual filing.
    milestone = _sec_milestone_catalyst(sec, data)
    if milestone:
        catalysts.append(milestone)

    # 2) Scheduled earnings is the highest-confidence recurring catalyst.
    earnings_date = _future_calendar_date(data.calendar, "Earnings Date", "EarningsDate")
    earnings_when = earnings_date.strftime("%Y-%m-%d") if earnings_date else "Next earnings (date TBC)"
    catalysts.append(
        Catalyst(
            catalyst="Next earnings release / guidance update",
            expected_date=earnings_when,
            probability="100%" if earnings_date else "95%",
            price_impact=_impact_text(low_move, high_move, 0),
            verify="Compare reported revenue, EPS, margins and forward guidance with current consensus; verify in the earnings release and 10-Q/6-K.",
        )
    )

    actual_year = _latest_actual_year(data) or (datetime.now().year - 1)
    revenue_now = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], actual_year)
    revenue_prev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], actual_year - 1)
    revenue_est = _estimate_avg(data.revenue_estimate, "0y") or _estimate_avg(data.revenue_estimate, "+1y")
    prior_growth = _safe_div(revenue_now - revenue_prev, revenue_prev) if revenue_now is not None and revenue_prev is not None else None
    forward_growth = _safe_div(revenue_est - revenue_now, revenue_now) if revenue_est is not None and revenue_now is not None else None
    if forward_growth is not None:
        delta = forward_growth - prior_growth if prior_growth is not None else 0.0
        if delta >= 0.03:
            name = "Revenue growth acceleration vs prior year"
            direction = 1
        elif delta <= -0.03:
            name = "Revenue growth deceleration vs prior year"
            direction = -1
        else:
            name = "Revenue growth delivery vs consensus"
            direction = 0
        verify = f"Verify FY{actual_year + 1} revenue growth against consensus {forward_growth * 100:.1f}%"
        if prior_growth is not None:
            verify += f" versus FY{actual_year} actual growth {prior_growth * 100:.1f}%."
        else:
            verify += "."
        catalysts.append(
            Catalyst(name, f"FY{actual_year + 1} / next results", "65%", _impact_text(low_move, high_move, direction), verify)
        )

    eps_now = _annual_value(data.annual_income, INCOME_ALIASES["EPS"], actual_year)
    eps_est = _estimate_avg(data.earnings_estimate, "0y") or _estimate_avg(data.earnings_estimate, "+1y")
    if eps_est is not None and eps_now is not None:
        if eps_now <= 0 < eps_est:
            name, direction = "EPS / profitability inflection to positive", 1
            detail = f"Consensus EPS moves from {eps_now:.2f} to {eps_est:.2f}."
        else:
            eps_growth = _safe_div(eps_est - eps_now, abs(eps_now)) if abs(eps_now) > 1e-9 else None
            if eps_growth is not None and eps_growth >= 0.20:
                name, direction = "EPS growth / operating leverage delivery", 1
            elif eps_growth is not None and eps_growth <= -0.20:
                name, direction = "EPS pressure / estimate downside risk", -1
            else:
                name, direction = "EPS delivery vs consensus", 0
            detail = f"Verify consensus EPS {eps_est:.2f} versus FY{actual_year} actual EPS {eps_now:.2f}."
        catalysts.append(
            Catalyst(name, f"FY{actual_year + 1} / next results", "65%", _impact_text(low_move, high_move, direction), detail)
        )

    # 5) Street target dispersion/revisions can create a measurable re-rating catalyst.
    mean_target = _target_price_value(data, "mean", "average", "targetMeanPrice")
    if mean_target is not None and data.current_price not in (None, 0):
        gap = (mean_target / data.current_price - 1.0) * 100.0
        if abs(gap) >= 5.0:
            direction = 1 if gap > 0 else -1
            impact_low = max(2.0, min(abs(gap) * 0.35, 12.0))
            impact_high = max(impact_low + 1.0, min(abs(gap) * 0.70, 25.0))
            catalysts.append(
                Catalyst(
                    catalyst="Street estimate / target-price re-rating",
                    expected_date="Next 6–12 months",
                    probability="55%",
                    price_impact=_impact_text(impact_low, impact_high, direction),
                    verify=f"Track analyst upgrades/downgrades and target revisions; current mean target {mean_target:.2f} implies {gap:+.1f}% versus spot.",
                )
            )

    # Fallback KPI catalyst so most liquid equities populate a full five-row table.
    if len(catalysts) < CATALYST_MAX_ROWS:
        ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], actual_year)
        revenue = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], actual_year)
        fcf = _annual_value(data.annual_cashflow, CASHFLOW_ALIASES["Free cash flow"], actual_year)
        margin = _safe_div(ebitda, revenue)
        verify_parts = []
        if margin is not None:
            verify_parts.append(f"FY{actual_year} EBITDA margin {margin * 100:.1f}%")
        if fcf is not None:
            verify_parts.append(f"FY{actual_year} FCF {_format_millions(fcf)}m")
        verify = " and ".join(verify_parts) if verify_parts else "reported EBITDA margin and free cash flow"
        catalysts.append(
            Catalyst(
                catalyst="Margin / cash-flow execution",
                expected_date="Next 2 earnings",
                probability="60%",
                price_impact=_impact_text(low_move, high_move, 0),
                verify=f"Verify whether {verify} improves or deteriorates in subsequent results and guidance.",
            )
        )

    # Preserve priority while removing duplicates and respecting the template's five rows.
    deduped: list[Catalyst] = []
    seen: set[str] = set()
    for item in catalysts:
        key = re.sub(r"[^a-z0-9]+", "", item.catalyst.lower())
        if not item.catalyst or key in seen:
            continue
        seen.add(key)
        deduped.append(item)
        if len(deduped) >= CATALYST_MAX_ROWS:
            break
    return deduped


def _risk_impact_text(data: MarketData, severity: str) -> str:
    low, high = _event_move_range(data)
    factor = {"High": 1.55, "Med": 1.0, "Low": 0.65}.get(severity, 1.0)
    impact_low = max(2.0, min(20.0, low * factor))
    impact_high = max(impact_low + 1.0, min(30.0, high * factor))
    return f"-{impact_low:.1f}–{impact_high:.1f}%"


def _level_from_score(score: int) -> str:
    if score >= 4:
        return "High"
    if score >= 2:
        return "Med"
    return "Low"


def build_risk_factors(data: MarketData, sec: SecResearch) -> list[RiskFactor]:
    """Populate six rule-based risks from financial trends and SEC risk disclosures."""
    risks: list[RiskFactor] = []
    actual_year = _latest_actual_year(data) or (datetime.now().year - 1)
    revenue = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], actual_year)
    prev_revenue = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], actual_year - 1)
    prev2_revenue = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], actual_year - 2)
    growth = _safe_div(revenue - prev_revenue, prev_revenue) if revenue is not None and prev_revenue is not None else None
    prev_growth = _safe_div(prev_revenue - prev2_revenue, prev2_revenue) if prev_revenue is not None and prev2_revenue is not None else None

    growth_score = 0
    if growth is not None and growth < 0:
        growth_score += 4
    elif growth is not None and growth < 0.05:
        growth_score += 2
    if growth is not None and prev_growth is not None and growth < prev_growth - 0.08:
        growth_score += 2
    growth_likelihood = _level_from_score(growth_score)
    growth_severity = "High" if growth is not None and growth < 0 else ("Med" if growth_score >= 2 else "Low")
    growth_detail = f"FY{actual_year} revenue growth {growth * 100:.1f}%" if growth is not None else "quarterly/annual revenue growth"
    risks.append(RiskFactor(
        "Revenue growth / demand slowdown",
        growth_likelihood,
        growth_severity,
        _risk_impact_text(data, growth_severity),
        f"Monitor {growth_detail} versus consensus and company guidance; escalate if two consecutive reporting periods miss the expected growth path. Source: filings / Yahoo Finance.",
    ))

    ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], actual_year)
    prev_ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], actual_year - 1)
    margin = _safe_div(ebitda, revenue)
    prev_margin = _safe_div(prev_ebitda, prev_revenue)
    margin_delta = margin - prev_margin if margin is not None and prev_margin is not None else None
    margin_score = 0
    if margin is not None and margin < 0:
        margin_score += 4
    if margin_delta is not None and margin_delta <= -0.05:
        margin_score += 4
    elif margin_delta is not None and margin_delta <= -0.02:
        margin_score += 2
    margin_likelihood = _level_from_score(margin_score)
    margin_severity = "High" if margin_score >= 4 else ("Med" if margin_score >= 2 else "Low")
    margin_detail = f"FY{actual_year} EBITDA margin {margin * 100:.1f}%" if margin is not None else "EBITDA/gross margin"
    risks.append(RiskFactor(
        "Margin / profitability compression",
        margin_likelihood,
        margin_severity,
        _risk_impact_text(data, margin_severity),
        f"Track {margin_detail}, gross margin, pricing and operating-expense growth; compare each result with prior guidance and peer margins. Source: filings / Yahoo Finance.",
    ))

    debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], actual_year)
    cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], actual_year)
    current_assets = _annual_value(data.annual_balance, BALANCE_ALIASES["Total current assets"], actual_year)
    current_liabilities = _annual_value(data.annual_balance, BALANCE_ALIASES["Total current liabilities"], actual_year)
    net_debt = debt - cash if debt is not None and cash is not None else None
    leverage = _safe_div(net_debt, ebitda) if net_debt is not None and ebitda not in (None, 0) else None
    current_ratio = _safe_div(current_assets, current_liabilities)
    balance_score = 0
    if leverage is not None and leverage >= 4:
        balance_score += 5
    elif leverage is not None and leverage >= 2:
        balance_score += 3
    if current_ratio is not None and current_ratio < 1:
        balance_score += 2
    balance_likelihood = _level_from_score(balance_score)
    balance_severity = "High" if balance_score >= 4 else ("Med" if balance_score >= 2 else "Low")
    metrics = []
    if leverage is not None:
        metrics.append(f"net debt/EBITDA {leverage:.2f}x")
    if current_ratio is not None:
        metrics.append(f"current ratio {current_ratio:.2f}x")
    metrics_text = ", ".join(metrics) or "net debt, liquidity and maturity schedule"
    risks.append(RiskFactor(
        "Balance-sheet / refinancing risk",
        balance_likelihood,
        balance_severity,
        _risk_impact_text(data, balance_severity),
        f"Monitor {metrics_text}, interest expense, covenant headroom and debt maturities; verify material financing changes in 8-K/6-K filings.",
    ))

    risk_text = _risk_section(sec.annual_text) or _business_section(sec.annual_text)
    filing_specs = [
        (
            "Customer / channel concentration",
            ("customer concentration", "significant customer", "major customer", "distributor", "channel partner", "customer accounted", "customers accounted"),
            "Track top-customer/channel exposure, renewals and disclosed concentration percentages; verify changes in 10-Q/10-K customer disclosures.",
            "High",
        ),
        (
            "Supply-chain / input-cost disruption",
            ("supply chain", "single source", "sole source", "supplier", "shortage", "raw material", "manufacturing disruption"),
            "Monitor lead times, inventory, gross margin, supplier concentration and management commentary on shortages/cost inflation.",
            "Med",
        ),
        (
            "Regulatory / legal / compliance risk",
            ("regulatory", "regulation", "litigation", "lawsuit", "government investigation", "antitrust", "compliance", "privacy", "export control"),
            "Monitor regulator/court announcements and new 8-K/10-Q disclosures; track any quantified fine, remedy, restriction or revenue exposure.",
            "High",
        ),
        (
            "Competitive / technology / pricing pressure",
            ("competition", "competitive", "pricing pressure", "price competition", "technological change", "market share", "compete"),
            "Monitor peer revenue growth, pricing, product launches, market-share commentary and the company’s gross/EBITDA margin trend.",
            "High",
        ),
    ]
    for title, keywords, monitor, default_severity in filing_specs:
        excerpt = _keyword_extract(risk_text, keywords, max_sentences=2, max_chars=420)
        if not excerpt and title.startswith("Customer"):
            # Customer concentration is not universal; skip when no evidence exists.
            continue
        likelihood = "Med" if excerpt else "Low"
        severity = default_severity if excerpt else "Med"
        source_note = " SEC annual filing flags: " + excerpt if excerpt else " No material specific disclosure was extracted; retain as a structural monitoring item."
        risks.append(RiskFactor(title, likelihood, severity, _risk_impact_text(data, severity), (monitor + source_note)[:700]))

    # Preserve priority and fill to the six-row template with broad structural risks if needed.
    fallback_specs = [
        ("Competitive / technology / pricing pressure", "Med", "High", "Monitor peer growth, pricing, product cadence and market-share commentary."),
        ("Regulatory / legal / compliance risk", "Low", "High", "Monitor regulator, court and company filing updates for quantified restrictions or liabilities."),
        ("Supply-chain / operating execution risk", "Low", "Med", "Monitor inventory, supplier commentary, delivery timing and gross-margin variance."),
    ]
    seen: set[str] = set()
    deduped: list[RiskFactor] = []
    for item in risks:
        key = re.sub(r"[^a-z0-9]+", "", item.risk.lower())
        if key in seen:
            continue
        seen.add(key)
        deduped.append(item)
        if len(deduped) >= RISK_MAX_ROWS:
            return deduped
    for title, likelihood, severity, monitor in fallback_specs:
        key = re.sub(r"[^a-z0-9]+", "", title.lower())
        if key in seen:
            continue
        seen.add(key)
        deduped.append(RiskFactor(title, likelihood, severity, _risk_impact_text(data, severity), monitor))
        if len(deduped) >= RISK_MAX_ROWS:
            break
    return deduped


def build_highlights(data: MarketData, peers: list[MarketData], sec: SecResearch) -> list[HighlightArgument]:
    """Build fact-based key arguments; this does not generate a Buy/Sell thesis."""
    items: list[HighlightArgument] = []
    year = _latest_actual_year(data) or (datetime.now().year - 1)
    revenue = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], year)
    prev_revenue = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], year - 1)
    growth = _safe_div(revenue - prev_revenue, prev_revenue) if revenue is not None and prev_revenue is not None else None
    if growth is not None:
        adjective = "strong" if growth >= 0.20 else "positive" if growth >= 0 else "negative"
        items.append(HighlightArgument(
            f"FY{year} revenue growth was {adjective} at {growth * 100:.1f}% YoY.",
            f"FY{year} revenue {_format_millions(revenue)}m vs FY{year-1} {_format_millions(prev_revenue)}m · Yahoo Finance / annual filing.",
        ))

    ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], year)
    margin = _safe_div(ebitda, revenue)
    prev_ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], year - 1)
    prev_margin = _safe_div(prev_ebitda, prev_revenue)
    if margin is not None:
        change = (margin - prev_margin) * 100 if prev_margin is not None else None
        direction = f", {change:+.1f}ppt YoY" if change is not None else ""
        items.append(HighlightArgument(
            f"FY{year} EBITDA margin was {margin * 100:.1f}%{direction}.",
            f"EBITDA {_format_millions(ebitda)}m / revenue {_format_millions(revenue)}m · Yahoo Finance / annual filing.",
        ))

    fcf = _annual_value(data.annual_cashflow, CASHFLOW_ALIASES["Free cash flow"], year)
    if fcf is not None:
        tone = "positive" if fcf >= 0 else "negative"
        items.append(HighlightArgument(
            f"FY{year} free cash flow was {tone} at {_format_millions(fcf)}m {data.currency or ''}.".strip(),
            f"Reported free cash flow · Yahoo Finance / annual cash-flow statement.",
        ))

    debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], year)
    cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], year)
    if debt is not None and cash is not None:
        net_debt = debt - cash
        state = "net debt" if net_debt >= 0 else "net cash"
        items.append(HighlightArgument(
            f"Balance sheet carried {state} of {_format_millions(abs(net_debt))}m {data.currency or ''} at FY{year}.".strip(),
            f"Debt {_format_millions(debt)}m and cash {_format_millions(cash)}m · Yahoo Finance / annual balance sheet.",
        ))

    # Compare the latest available EV/EBITDA with the selected peer median when data exists.
    target_metric = _year_peer_metrics(data, year).get("ev / ebitda (x)")
    peer_values = [
        _year_peer_metrics(peer, year).get("ev / ebitda (x)")
        for peer in peers[:3]
    ]
    peer_values = [v for v in peer_values if v is not None and math.isfinite(v) and v > 0]
    if target_metric is not None and target_metric > 0 and peer_values:
        median = statistics.median(peer_values)
        gap = (target_metric / median - 1.0) * 100 if median else None
        descriptor = "premium" if gap is not None and gap >= 0 else "discount"
        items.append(HighlightArgument(
            f"FY{year} EV/EBITDA was {target_metric:.1f}x versus selected-peer median {median:.1f}x ({abs(gap or 0):.1f}% {descriptor}).",
            "Year-end price + reported financials for target and automatically selected peers · Yahoo Finance.",
        ))

    mean_target = _target_price_value(data, "mean", "average", "targetMeanPrice")
    if mean_target is not None and data.current_price not in (None, 0):
        gap = (mean_target / data.current_price - 1.0) * 100
        items.append(HighlightArgument(
            f"Street mean target implies {gap:+.1f}% versus the current share price.",
            f"Mean target {mean_target:.2f} vs spot {data.current_price:.2f} · Yahoo Finance analyst-target aggregation.",
        ))

    if sec.annual_url and len(items) < HIGHLIGHT_MAX_ROWS:
        items.append(HighlightArgument(
            f"Latest primary annual filing available: {sec.annual_form} filed {sec.annual_filing_date}.",
            f"SEC EDGAR · {sec.annual_url}",
        ))

    # De-duplicate and keep the first five fact patterns by priority.
    deduped: list[HighlightArgument] = []
    seen: set[str] = set()
    for item in items:
        key = re.sub(r"[^a-z0-9]+", "", item.argument.lower())[:180]
        if not item.argument or key in seen:
            continue
        seen.add(key)
        deduped.append(item)
        if len(deduped) >= HIGHLIGHT_MAX_ROWS:
            break
    return deduped


def _region_for_yahoo(data: MarketData) -> str:
    country = str((data.info or {}).get("country") or "").strip().lower()
    mapping = {
        "united states": "US",
        "canada": "CA",
        "united kingdom": "GB",
        "south korea": "KR",
        "korea": "KR",
        "japan": "JP",
        "germany": "DE",
        "france": "FR",
        "australia": "AU",
        "hong kong": "HK",
        "singapore": "SG",
        "india": "IN",
        "china": "CN",
        "taiwan": "TW",
    }
    return mapping.get(country, "US")


def _extract_symbols_from_frame(frame: pd.DataFrame | None) -> list[str]:
    if frame is None or frame.empty:
        return []
    symbols: list[str] = []
    columns = {str(c).lower(): c for c in frame.columns}
    symbol_col = next((c for key, c in columns.items() if key in ("symbol", "ticker") or "symbol" in key), None)
    if symbol_col is not None:
        symbols.extend(str(x).upper().strip() for x in frame[symbol_col].tolist())
    for idx in frame.index.tolist():
        candidate = str(idx).upper().strip()
        if re.fullmatch(r"[A-Z0-9.\-]{1,20}", candidate):
            symbols.append(candidate)
    return [x for x in symbols if x and x != "NAN"]


def _industry_candidate_symbols(data: MarketData) -> list[str]:
    """Return a compact list of plausible peers without fetching full fundamentals.

    yfinance's Industry/Sector constructors changed across versions. Calling them with
    only the key works on the 0.2.x line used by this project and remains compatible
    with newer releases; region filtering is handled separately below.
    """
    info = data.info or {}
    industry_key = str(info.get("industryKey") or "").strip()
    region = _region_for_yahoo(data)
    candidates: list[str] = []

    if industry_key and hasattr(yf, "Industry"):
        try:
            industry = yf.Industry(industry_key)
            for attr in ("top_companies", "top_growth_companies", "top_performing_companies"):
                try:
                    frame = getattr(industry, attr)
                    if callable(frame):
                        frame = frame()
                    candidates.extend(_extract_symbols_from_frame(frame if isinstance(frame, pd.DataFrame) else None))
                except Exception:
                    continue
        except Exception as exc:
            print(f"[PEER] Industry API unavailable: {type(exc).__name__}: {exc}")

    # Fallback/expansion: Yahoo screener by exact industry display name.
    industry_name = str(data.industry or "").strip()
    if industry_name and hasattr(yf, "EquityQuery") and hasattr(yf, "screen"):
        try:
            query_parts: list[Any] = [yf.EquityQuery("eq", ["industry", industry_name])]
            if region == "US":
                query_parts.append(yf.EquityQuery("eq", ["region", "us"]))
            query = query_parts[0] if len(query_parts) == 1 else yf.EquityQuery("and", query_parts)
            response = yf.screen(query, size=min(60, PEER_CANDIDATE_LIMIT * 3))
            if isinstance(response, dict):
                for quote in response.get("quotes", []) or []:
                    if not isinstance(quote, dict):
                        continue
                    symbol = str(quote.get("symbol") or quote.get("ticker") or "").upper().strip()
                    if symbol:
                        candidates.append(symbol)
        except Exception as exc:
            print(f"[PEER] Screener unavailable: {type(exc).__name__}: {exc}")

    # Sector fallback only when industry endpoints return too few names.
    if len(set(candidates)) < 6:
        sector_key = str(info.get("sectorKey") or "").strip()
        if sector_key and hasattr(yf, "Sector"):
            try:
                sector = yf.Sector(sector_key)
                frame = sector.top_companies
                if callable(frame):
                    frame = frame()
                candidates.extend(_extract_symbols_from_frame(frame if isinstance(frame, pd.DataFrame) else None))
            except Exception as exc:
                print(f"[PEER] Sector API unavailable: {type(exc).__name__}: {exc}")

    deduped: list[str] = []
    seen: set[str] = set()
    for symbol in candidates:
        symbol = symbol.upper().strip()
        if symbol == data.ticker.upper() or symbol in seen:
            continue
        if not re.fullmatch(r"[A-Z0-9.\-]{1,20}", symbol):
            continue
        # For a US-listed target, reject obvious foreign-suffix symbols before any
        # expensive fundamentals request. Hyphenated US classes (BRK-B) remain valid.
        if region == "US" and "." in symbol:
            continue
        seen.add(symbol)
        deduped.append(symbol)
    return deduped[:PEER_CANDIDATE_LIMIT]


def _peer_is_eligible(target: MarketData, candidate: MarketData) -> bool:
    target_region = _region_for_yahoo(target)
    info = candidate.info or {}
    exchange = (candidate.exchange or str(info.get("exchange") or "")).upper()
    quote_type = str(info.get("quoteType") or "EQUITY").upper()

    if quote_type not in ("EQUITY", ""):  # reject funds, indexes, futures, etc.
        return False
    if not candidate.market_cap or candidate.market_cap <= 0:
        return False
    if target.sector and candidate.sector and target.sector.lower() != candidate.sector.lower():
        return False

    if target_region == "US":
        if "." in candidate.ticker:
            return False
        bad_exchange_tokens = ("OTC", "PINK", "GREY", "PNK", "YHD")
        if any(token in exchange for token in bad_exchange_tokens):
            return False
        # US-listed ADRs can still be valid peers (e.g. a foreign issuer trading on
        # Nasdaq/NYSE), so filter by listing venue rather than issuer domicile.

    return True

def _latest_actual_year(data: MarketData) -> int | None:
    for year in range(datetime.now().year, 2018, -1):
        if _column_for_year(data.annual_income, year) is not None:
            return year
    return None


def _peer_score(target: MarketData, candidate: MarketData) -> float:
    score = 0.0
    penalties = 0.0
    target_year = _latest_actual_year(target) or 2025
    candidate_year = _latest_actual_year(candidate) or target_year
    year = min(target_year, candidate_year)

    target_revenue = _annual_value(target.annual_income, INCOME_ALIASES["Revenue"], year)
    cand_revenue = _annual_value(candidate.annual_income, INCOME_ALIASES["Revenue"], year)
    target_ebitda = _annual_value(target.annual_income, INCOME_ALIASES["EBITDA"], year)
    cand_ebitda = _annual_value(candidate.annual_income, INCOME_ALIASES["EBITDA"], year)
    target_prev = _annual_value(target.annual_income, INCOME_ALIASES["Revenue"], year - 1)
    cand_prev = _annual_value(candidate.annual_income, INCOME_ALIASES["Revenue"], year - 1)

    def log_distance(a: float | None, b: float | None, weight: float) -> None:
        nonlocal score, penalties
        if a is not None and b is not None and a > 0 and b > 0:
            score += weight * abs(math.log(a / b))
        else:
            penalties += weight * 1.5

    log_distance(target.market_cap, candidate.market_cap, 0.38)
    log_distance(target_revenue, cand_revenue, 0.24)

    tg = _safe_div(target_revenue - target_prev, target_prev) if target_revenue is not None and target_prev is not None else None
    cg = _safe_div(cand_revenue - cand_prev, cand_prev) if cand_revenue is not None and cand_prev is not None else None
    if tg is not None and cg is not None:
        score += 0.18 * min(abs(tg - cg) * 4.0, 2.0)
    else:
        penalties += 0.18

    tm = _safe_div(target_ebitda, target_revenue)
    cm = _safe_div(cand_ebitda, cand_revenue)
    if tm is not None and cm is not None:
        score += 0.15 * min(abs(tm - cm) * 3.0, 2.0)
    else:
        penalties += 0.15

    if target.beta is not None and candidate.beta is not None:
        score += 0.05 * min(abs(target.beta - candidate.beta), 2.0)

    if target.industry and candidate.industry and target.industry.lower() != candidate.industry.lower():
        score += 0.8
    elif target.sector and candidate.sector and target.sector.lower() != candidate.sector.lower():
        score += 1.4

    return score + penalties


def select_peer_companies(data: MarketData, count: int = 3) -> list[MarketData]:
    symbols = _industry_candidate_symbols(data)
    if not symbols:
        return []

    candidates: list[tuple[float, MarketData]] = []
    fetch_cap = min(len(symbols), max(count * 4, 10))
    for symbol in symbols[:fetch_cap]:
        try:
            candidate = fetch_peer_data(symbol)
            if not candidate.company_name or candidate.ticker == data.ticker:
                continue
            if not _peer_is_eligible(data, candidate):
                continue
            candidates.append((_peer_score(data, candidate), candidate))
        except Exception as exc:
            print(f"[PEER] {symbol} skipped: {type(exc).__name__}: {exc}")

    candidates.sort(key=lambda item: item[0])
    selected: list[MarketData] = []
    seen_names: set[str] = set()
    for _, candidate in candidates:
        company_key = re.sub(r"\W+", "", candidate.company_name.lower())
        if company_key and company_key in seen_names:
            continue
        seen_names.add(company_key)
        selected.append(candidate)
        if len(selected) >= count:
            break
    return selected


def _year_peer_metrics(data: MarketData, year: int) -> dict[str, float | None]:
    revenue = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], year)
    ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], year)
    op_income = _annual_value(data.annual_income, INCOME_ALIASES["Operating income"], year)
    net_income = _annual_value(data.annual_income, INCOME_ALIASES["Net income"], year)
    shares = _annual_value(data.annual_income, INCOME_ALIASES["Diluted shares"], year)
    if shares is None:
        shares = _annual_value(data.annual_balance, BALANCE_ALIASES["Ordinary shares"], year)
    date_col = _column_for_year(data.annual_balance, year)
    if date_col is None:
        date_col = _column_for_year(data.annual_income, year)
    price = _price_at_or_before(data.price_history, date_col) if date_col is not None else None
    market_cap = price * shares if price is not None and shares is not None else None
    debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], year)
    cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], year)
    minority = _annual_value(data.annual_balance, BALANCE_ALIASES["Minority interest"], year)
    preferred = _annual_value(data.annual_balance, BALANCE_ALIASES["Preferred equity"], year)
    net_debt = debt - cash if debt is not None and cash is not None else None
    enterprise_value = market_cap + (net_debt or 0.0) + (minority or 0.0) + (preferred or 0.0) if market_cap is not None else None
    return {
        "market cap": market_cap,
        "enterprise value": enterprise_value,
        "revenue": revenue,
        "ebitda": ebitda,
        "operating income": op_income,
        "net income": net_income,
        "ev / sales (x)": _safe_div(enterprise_value, revenue),
        "ev / ebitda (x)": _safe_div(enterprise_value, ebitda),
        "p / e (x)": _safe_div(market_cap, net_income),
    }


def _format_peer_metric(label: str, value: float | None) -> str:
    if value is None:
        return ""
    if label.endswith("(x)"):
        return _format_ratio(value)
    return _format_millions(value)


def _fill_header(doc: DocumentObject, data: MarketData) -> None:
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text.startswith("Company Name:"):
            _set_paragraph_text(paragraph, f"Company Name: {data.company_name} (Ticker: {data.ticker})")
        elif text.startswith("Ticker ") and "Exchange" in text and "GICS sector" in text:
            sector_industry = " / ".join(x for x in (data.sector, data.industry) if x) or "______"
            _set_paragraph_text(
                paragraph,
                f"Ticker {data.ticker} · Exchange {data.exchange or '______'} · "
                f"GICS sector / industry {sector_industry} · Reporting currency {data.currency or '______'}",
            )

    # Top recommendation box: only objective fields are populated.
    for table in doc.tables:
        if len(table.rows) == 3 and len(table.columns) == 4 and table.cell(0, 0).text.strip() == "Recommendation":
            if data.current_price is not None:
                _set_cell_text(table.cell(1, 1), _format_price(data.current_price))
            _set_cell_text(table.cell(2, 1), datetime.now().strftime("%m / %d / %Y"))
            break


def _fill_market_profile(doc: DocumentObject, data: MarketData) -> None:
    table = _find_table_by_first_cell(doc, "ITEM")
    # There are several ITEM tables; identify the 2-column one containing Closing price.
    for candidate in _all_tables(doc):
        rows = _row_map(candidate)
        if "closing price" in rows and len(candidate.columns) == 2:
            table = candidate
            break
    if table is None:
        return

    rows = _row_map(table)
    values = {
        "closing price": _format_price(data.current_price),
        "currency": data.currency,
        "52-week high · low": (
            f"{_format_price(data.high_52w)} · {_format_price(data.low_52w)}"
            if data.high_52w is not None and data.low_52w is not None
            else ""
        ),
        "average daily value traded": _format_compact_money(data.avg_daily_value, data.currency),
        "market cap": _format_compact_money(data.market_cap, data.currency),
        "shares outstanding": _format_shares_m(data.shares_outstanding),
        "free float": _format_shares_m(data.float_shares),
        "beta": _format_number(data.beta, 2),
        "dividend yield": _format_pct_value(data.dividend_yield_pct),
    }
    for label, value in values.items():
        if value and label in rows:
            _set_cell_text(table.cell(rows[label], 1), value)



def _fill_highlights(doc: DocumentObject, highlights: list[HighlightArgument]) -> None:
    if not highlights:
        return
    for table in _all_tables(doc):
        if len(table.columns) != 3 or len(table.rows) < 2:
            continue
        header = [cell.text.strip().upper() for cell in table.rows[0].cells]
        if header[:3] != ["#", "ARGUMENT", "EVIDENCE · SOURCE"]:
            continue
        max_rows = min(HIGHLIGHT_MAX_ROWS, len(table.rows) - 1, len(highlights))
        for idx in range(max_rows):
            item = highlights[idx]
            _set_cell_text(table.cell(idx + 1, 0), str(idx + 1))
            _set_cell_text(table.cell(idx + 1, 1), item.argument)
            _set_cell_text(table.cell(idx + 1, 2), item.evidence_source)
        break


def _fill_risk_factors(doc: DocumentObject, risks: list[RiskFactor]) -> None:
    if not risks:
        return
    for table in _all_tables(doc):
        if len(table.columns) != 5 or len(table.rows) < 2:
            continue
        header = [cell.text.strip().upper() for cell in table.rows[0].cells]
        if header[:5] != ["RISK", "LIKELIHOOD", "SEVERITY", "TP IMPACT", "MITIGATION · MONITORED INDICATOR"]:
            continue
        for idx, item in enumerate(risks[: min(RISK_MAX_ROWS, len(table.rows) - 1)], start=1):
            values = (item.risk, item.likelihood, item.severity, item.tp_impact, item.mitigation)
            for col, value in enumerate(values):
                if value:
                    _set_cell_text(table.cell(idx, col), str(value))
        break


def _fill_company_background(doc: DocumentObject, background: CompanyBackground) -> None:
    labels = {
        "founding · listing history": background.founding_listing,
        "segments and revenue mix": background.segments_revenue_mix,
        "revenue by geography": background.revenue_geography,
        "key customers · channels": background.key_customers_channels,
        "competitors and market share": background.competitors_market_share,
        "cost structure · supply chain": background.cost_structure_supply_chain,
        "ownership and major shareholders": background.ownership_major_shareholders,
        "management and incentives": background.management_incentives,
    }
    for table in _all_tables(doc):
        rows = _row_map(table)
        if "founding · listing history" not in rows or len(table.columns) != 2:
            continue
        for label, value in labels.items():
            if label in rows and value:
                _set_cell_text(table.cell(rows[label], 1), value)
        break


def _fill_catalysts(doc: DocumentObject, catalysts: list[Catalyst]) -> None:
    if not catalysts:
        return
    table = None
    for candidate in _all_tables(doc):
        if len(candidate.columns) == 5 and len(candidate.rows) >= 2:
            header = [cell.text.strip().upper() for cell in candidate.rows[0].cells]
            if header[:5] == ["CATALYST", "EXPECTED DATE", "PROB.", "PRICE IMPACT", "HOW WE VERIFY IT"]:
                table = candidate
                break
    if table is None:
        return

    for out_row, item in enumerate(catalysts[: min(CATALYST_MAX_ROWS, len(table.rows) - 1)], start=1):
        values = (item.catalyst, item.expected_date, item.probability, item.price_impact, item.verify)
        for col, value in enumerate(values):
            if value:
                _set_cell_text(table.cell(out_row, col), str(value))


def _numeric_text_value(text: str) -> float | None:
    clean = (text or "").strip().replace(",", "")
    # Only style cells that begin with a numeric sign/value. This avoids colouring
    # prose such as "From -2% to +3%" as an entire paragraph.
    match = re.match(r"^([+\-−]?\d+(?:\.\d+)?)", clean)
    if not match:
        return None
    try:
        return float(match.group(1).replace("−", "-"))
    except ValueError:
        return None


def _style_numeric_cells(doc: DocumentObject) -> None:
    negative_red = RGBColor(0xD8, 0x3B, 0x3B)

    for table in _all_tables(doc):
        rows = _row_map(table)
        growth_row = rows.get("revenue growth (%)")
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                value = _numeric_text_value(cell.text)
                if value is None:
                    continue
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if value < 0:
                            run.font.color.rgb = negative_red
                        if growth_row is not None and row_idx == growth_row and col_idx > 0 and value >= REVENUE_GROWTH_HIGHLIGHT_PCT:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW



def build_key_financials_payload(data: MarketData) -> dict[str, Any]:
    """Return the same key financials used by the DOCX as a web-friendly payload.

    Monetary values are displayed in millions to match the template. Raw numeric
    values are included for conditional styling in the browser.
    """
    income_years = [2023, 2024, 2025, 2026, 2027]
    balance_years = [2023, 2024, 2025, 2026, 2027]

    actual_revenue: dict[int, float | None] = {
        year: _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], year)
        for year in (2022, 2023, 2024, 2025)
    }

    income_rows: dict[str, list[dict[str, Any]]] = {
        label: [] for label in (
            "Revenue", "Revenue growth (%)", "EBITDA", "EBITDA margin (%)",
            "Operating income", "Net income", "EPS", "Free cash flow",
            "ROIC (%)", "Net debt / EBITDA (x)",
        )
    }

    def add_income(label: str, raw: float | None, display: str, kind: str) -> None:
        income_rows[label].append({"raw": raw, "display": display or "", "kind": kind})

    for year in (2023, 2024, 2025):
        revenue = actual_revenue.get(year)
        prev_revenue = actual_revenue.get(year - 1)
        growth = _safe_div(revenue - prev_revenue, prev_revenue) if revenue is not None and prev_revenue is not None else None
        ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], year)
        op_income = _annual_value(data.annual_income, INCOME_ALIASES["Operating income"], year)
        net_income = _annual_value(data.annual_income, INCOME_ALIASES["Net income"], year)
        eps = _annual_value(data.annual_income, INCOME_ALIASES["EPS"], year)
        fcf = _annual_value(data.annual_cashflow, CASHFLOW_ALIASES["Free cash flow"], year)
        cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], year)
        debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], year)
        net_debt = debt - cash if debt is not None and cash is not None else None
        roic = _roic(data, year)
        leverage = _safe_div(net_debt, ebitda)

        add_income("Revenue", None if revenue is None else revenue / 1_000_000, _format_millions(revenue), "money")
        add_income("Revenue growth (%)", None if growth is None else growth * 100, _format_pct_ratio(growth), "growth_pct")
        add_income("EBITDA", None if ebitda is None else ebitda / 1_000_000, _format_millions(ebitda), "money")
        margin = _safe_div(ebitda, revenue)
        add_income("EBITDA margin (%)", None if margin is None else margin * 100, _format_pct_ratio(margin), "pct")
        add_income("Operating income", None if op_income is None else op_income / 1_000_000, _format_millions(op_income), "money")
        add_income("Net income", None if net_income is None else net_income / 1_000_000, _format_millions(net_income), "money")
        add_income("EPS", eps, _format_number(eps, 2), "number")
        add_income("Free cash flow", None if fcf is None else fcf / 1_000_000, _format_millions(fcf), "money")
        add_income("ROIC (%)", None if roic is None else roic * 100, _format_pct_ratio(roic), "pct")
        add_income("Net debt / EBITDA (x)", leverage, _format_ratio(leverage), "ratio")

    previous_revenue = actual_revenue.get(2025)
    for year, period in ((2026, "0y"), (2027, "+1y")):
        revenue = _estimate_avg(data.revenue_estimate, period)
        eps = _estimate_avg(data.earnings_estimate, period)
        growth = _safe_div(revenue - previous_revenue, previous_revenue) if revenue is not None and previous_revenue is not None else None
        add_income("Revenue", None if revenue is None else revenue / 1_000_000, _format_millions(revenue), "money")
        add_income("Revenue growth (%)", None if growth is None else growth * 100, _format_pct_ratio(growth), "growth_pct")
        for label, kind in (
            ("EBITDA", "money"), ("EBITDA margin (%)", "pct"),
            ("Operating income", "money"), ("Net income", "money"),
        ):
            add_income(label, None, "", kind)
        add_income("EPS", eps, _format_number(eps, 2), "number")
        for label, kind in (("Free cash flow", "money"), ("ROIC (%)", "pct"), ("Net debt / EBITDA (x)", "ratio")):
            add_income(label, None, "", kind)
        if revenue is not None:
            previous_revenue = revenue

    balance_labels = (
        "Cash and equivalents", "Accounts receivable", "Inventory", "Total current assets",
        "Property, plant and equipment", "Goodwill and intangibles", "Total assets",
        "Accounts payable", "Total current liabilities", "Total debt (incl. leases)",
        "Total liabilities", "Total equity", "Net debt", "Net debt / EBITDA (x)",
        "Current ratio (x)", "ROE (%)", "ROIC (%)",
    )
    balance_rows: dict[str, list[dict[str, Any]]] = {label: [] for label in balance_labels}

    def add_balance(label: str, raw: float | None, display: str, kind: str) -> None:
        balance_rows[label].append({"raw": raw, "display": display or "", "kind": kind})

    for year in (2023, 2024, 2025):
        cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], year)
        debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], year)
        ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], year)
        current_assets = _annual_value(data.annual_balance, BALANCE_ALIASES["Total current assets"], year)
        current_liabilities = _annual_value(data.annual_balance, BALANCE_ALIASES["Total current liabilities"], year)
        net_debt = debt - cash if debt is not None and cash is not None else None
        goodwill = _annual_value(data.annual_balance, ("Goodwill And Other Intangible Assets",), year)
        if goodwill is None:
            goodwill = _combined_annual_value(data.annual_balance, (("Goodwill",), ("Other Intangible Assets",)), year)

        monetary = {
            "Cash and equivalents": cash,
            "Accounts receivable": _annual_value(data.annual_balance, BALANCE_ALIASES["Accounts receivable"], year),
            "Inventory": _annual_value(data.annual_balance, BALANCE_ALIASES["Inventory"], year),
            "Total current assets": current_assets,
            "Property, plant and equipment": _annual_value(data.annual_balance, BALANCE_ALIASES["Property, plant and equipment"], year),
            "Goodwill and intangibles": goodwill,
            "Total assets": _annual_value(data.annual_balance, BALANCE_ALIASES["Total assets"], year),
            "Accounts payable": _annual_value(data.annual_balance, BALANCE_ALIASES["Accounts payable"], year),
            "Total current liabilities": current_liabilities,
            "Total debt (incl. leases)": debt,
            "Total liabilities": _annual_value(data.annual_balance, BALANCE_ALIASES["Total liabilities"], year),
            "Total equity": _annual_value(data.annual_balance, BALANCE_ALIASES["Total equity"], year),
            "Net debt": net_debt,
        }
        for label, value in monetary.items():
            add_balance(label, None if value is None else value / 1_000_000, _format_millions(value), "money")
        leverage = _safe_div(net_debt, ebitda)
        current_ratio = _safe_div(current_assets, current_liabilities)
        roe = _roe(data, year)
        roic = _roic(data, year)
        add_balance("Net debt / EBITDA (x)", leverage, _format_ratio(leverage), "ratio")
        add_balance("Current ratio (x)", current_ratio, _format_ratio(current_ratio), "ratio")
        add_balance("ROE (%)", None if roe is None else roe * 100, _format_pct_ratio(roe), "pct")
        add_balance("ROIC (%)", None if roic is None else roic * 100, _format_pct_ratio(roic), "pct")

    # Forecast balance-sheet values are intentionally left blank because Yahoo's
    # public consensus endpoints do not reliably provide a full forecast balance sheet.
    for _year in (2026, 2027):
        for label in balance_labels:
            kind = "ratio" if "(x)" in label else "pct" if "(%)" in label else "money"
            add_balance(label, None, "", kind)

    return {
        "currency": data.currency,
        "units": "millions except per-share figures and ratios",
        "income": {
            "years": [f"FY{year}{'A' if year <= 2025 else 'E'}" for year in income_years],
            "rows": [{"label": label, "values": values} for label, values in income_rows.items()],
        },
        "balance": {
            "years": [f"FY{year}{'A' if year <= 2025 else 'E'}" for year in balance_years],
            "rows": [{"label": label, "values": values} for label, values in balance_rows.items()],
        },
        "revenue_growth_highlight_pct": REVENUE_GROWTH_HIGHLIGHT_PCT,
    }

def _fill_income_statement(doc: DocumentObject, data: MarketData) -> None:
    table = _find_table_containing_row(doc, "Revenue growth (%)")
    if table is None:
        return
    rows = _row_map(table)
    years = {2023: 1, 2024: 2, 2025: 3, 2026: 4, 2027: 5}

    actual_revenue: dict[int, float | None] = {}
    for year in (2022, 2023, 2024, 2025):
        actual_revenue[year] = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], year)

    for year in (2023, 2024, 2025):
        col = years[year]
        revenue = actual_revenue[year]
        ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], year)
        op_income = _annual_value(data.annual_income, INCOME_ALIASES["Operating income"], year)
        net_income = _annual_value(data.annual_income, INCOME_ALIASES["Net income"], year)
        eps = _annual_value(data.annual_income, INCOME_ALIASES["EPS"], year)
        fcf = _annual_value(data.annual_cashflow, CASHFLOW_ALIASES["Free cash flow"], year)
        cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], year)
        debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], year)
        net_debt = debt - cash if debt is not None and cash is not None else None

        values = {
            "revenue": _format_millions(revenue),
            "revenue growth (%)": _format_pct_ratio(_safe_div(revenue - actual_revenue.get(year - 1), actual_revenue.get(year - 1)) if revenue is not None and actual_revenue.get(year - 1) is not None else None),
            "ebitda": _format_millions(ebitda),
            "ebitda margin (%)": _format_pct_ratio(_safe_div(ebitda, revenue)),
            "operating income": _format_millions(op_income),
            "net income": _format_millions(net_income),
            "eps": _format_number(eps, 2),
            "free cash flow": _format_millions(fcf),
            "roic (%)": _format_pct_ratio(_roic(data, year)),
            "net debt / ebitda (x)": _format_ratio(_safe_div(net_debt, ebitda)),
        }
        for label, value in values.items():
            if value and label in rows:
                _set_cell_text(table.cell(rows[label], col), value)

    # Best-effort analyst estimates: yfinance usually exposes 0y / +1y revenue and EPS.
    estimate_map = [(2026, "0y", 4), (2027, "+1y", 5)]
    previous_revenue = actual_revenue.get(2025)
    for year, period, col in estimate_map:
        revenue = _estimate_avg(data.revenue_estimate, period)
        eps = _estimate_avg(data.earnings_estimate, period)
        if revenue is not None and "revenue" in rows:
            _set_cell_text(table.cell(rows["revenue"], col), _format_millions(revenue))
        if revenue is not None and previous_revenue is not None and "revenue growth (%)" in rows:
            _set_cell_text(table.cell(rows["revenue growth (%)"], col), _format_pct_ratio(_safe_div(revenue - previous_revenue, previous_revenue)))
        if eps is not None and "eps" in rows:
            _set_cell_text(table.cell(rows["eps"], col), _format_number(eps, 2))
        if revenue is not None:
            previous_revenue = revenue


def _fill_balance_sheet(doc: DocumentObject, data: MarketData) -> None:
    table = _find_table_containing_row(doc, "Cash and equivalents")
    if table is None or len(table.columns) != 6:
        # Find the annual 6-column balance table rather than the quarterly 7-column one.
        for candidate in _all_tables(doc):
            rows = _row_map(candidate)
            if "cash and equivalents" in rows and "current ratio (x)" in rows and len(candidate.columns) == 6:
                table = candidate
                break
    if table is None:
        return

    rows = _row_map(table)
    years = {2023: 1, 2024: 2, 2025: 3}

    for year, col in years.items():
        cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], year)
        debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], year)
        ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], year)
        current_assets = _annual_value(data.annual_balance, BALANCE_ALIASES["Total current assets"], year)
        current_liabilities = _annual_value(data.annual_balance, BALANCE_ALIASES["Total current liabilities"], year)
        net_debt = debt - cash if debt is not None and cash is not None else None

        goodwill = _annual_value(data.annual_balance, ("Goodwill And Other Intangible Assets",), year)
        if goodwill is None:
            goodwill = _combined_annual_value(
                data.annual_balance,
                (("Goodwill",), ("Other Intangible Assets",)),
                year,
            )

        values = {
            "cash and equivalents": _format_millions(cash),
            "accounts receivable": _format_millions(_annual_value(data.annual_balance, BALANCE_ALIASES["Accounts receivable"], year)),
            "inventory": _format_millions(_annual_value(data.annual_balance, BALANCE_ALIASES["Inventory"], year)),
            "total current assets": _format_millions(current_assets),
            "property, plant and equipment": _format_millions(_annual_value(data.annual_balance, BALANCE_ALIASES["Property, plant and equipment"], year)),
            "goodwill and intangibles": _format_millions(goodwill),
            "total assets": _format_millions(_annual_value(data.annual_balance, BALANCE_ALIASES["Total assets"], year)),
            "accounts payable": _format_millions(_annual_value(data.annual_balance, BALANCE_ALIASES["Accounts payable"], year)),
            "total current liabilities": _format_millions(current_liabilities),
            "total debt (incl. leases)": _format_millions(debt),
            "total liabilities": _format_millions(_annual_value(data.annual_balance, BALANCE_ALIASES["Total liabilities"], year)),
            "total equity": _format_millions(_annual_value(data.annual_balance, BALANCE_ALIASES["Total equity"], year)),
            "net debt": _format_millions(net_debt),
            "net debt / ebitda (x)": _format_ratio(_safe_div(net_debt, ebitda)),
            "current ratio (x)": _format_ratio(_safe_div(current_assets, current_liabilities)),
            "roe (%)": _format_pct_ratio(_roe(data, year)),
            "roic (%)": _format_pct_ratio(_roic(data, year)),
        }
        for label, value in values.items():
            if value and label in rows:
                _set_cell_text(table.cell(rows[label], col), value)


def _fill_quarterly_valuation(doc: DocumentObject, data: MarketData) -> None:
    table = None
    for candidate in _all_tables(doc):
        rows = _row_map(candidate)
        if "share price" in rows and "enterprise value" in rows and "ltm ebitda" in rows and len(candidate.columns) == 7:
            table = candidate
            break
    if table is None:
        return

    rows = _row_map(table)
    quarters = _quarter_columns(data)
    if not quarters:
        return

    # Keep only the latest six, shown chronologically.
    quarters = quarters[-6:]
    for out_col, quarter in enumerate(quarters, start=1):
        if out_col >= len(table.columns):
            break
        _set_cell_text(table.cell(0, out_col), _quarter_label(quarter))

        price = _price_at_or_before(data.price_history, quarter)
        shares = _quarter_value(data.quarterly_income, INCOME_ALIASES["Diluted shares"], quarter)
        if shares is None:
            shares = _quarter_value(data.quarterly_balance, BALANCE_ALIASES["Ordinary shares"], quarter)
        debt = _quarter_value(data.quarterly_balance, BALANCE_ALIASES["Total debt (incl. leases)"], quarter)
        lease = _quarter_value(data.quarterly_balance, BALANCE_ALIASES["Lease liabilities"], quarter)
        cash = _quarter_value(data.quarterly_balance, BALANCE_ALIASES["Cash and equivalents"], quarter)
        minority = _quarter_value(data.quarterly_balance, BALANCE_ALIASES["Minority interest"], quarter)
        preferred = _quarter_value(data.quarterly_balance, BALANCE_ALIASES["Preferred equity"], quarter)
        market_cap = price * shares if price is not None and shares is not None else None
        net_debt = debt - cash if debt is not None and cash is not None else None
        enterprise_value = market_cap + (net_debt or 0.0) + (minority or 0.0) + (preferred or 0.0) if market_cap is not None else None
        ltm_ebitda = _ltm_sum(data.quarterly_income, INCOME_ALIASES["EBITDA"], quarter)
        ltm_net_income = _ltm_sum(data.quarterly_income, INCOME_ALIASES["Net income"], quarter)
        ev_ebitda = _safe_div(enterprise_value, ltm_ebitda)
        pe = _safe_div(market_cap, ltm_net_income)

        values = {
            "share price": _format_price(price),
            "diluted shares (m)": _format_millions(shares),
            "market cap": _format_millions(market_cap),
            "total debt": _format_millions(debt),
            "lease liabilities": _format_millions(lease),
            "cash and equivalents": _format_millions(cash),
            "net debt": _format_millions(net_debt),
            "minorities · preferred": _format_millions((minority or 0.0) + (preferred or 0.0)) if minority is not None or preferred is not None else "",
            "enterprise value": _format_millions(enterprise_value),
            "ltm ebitda": _format_millions(ltm_ebitda),
            "ev / ebitda (x)": _format_ratio(ev_ebitda),
            "p / e (x)": _format_ratio(pe),
        }
        for label, value in values.items():
            if value and label in rows:
                _set_cell_text(table.cell(rows[label], out_col), value)


def _fill_target_company_peer_columns(
    doc: DocumentObject,
    data: MarketData,
    peers: list[MarketData],
) -> None:
    table = None
    for candidate in _all_tables(doc):
        if len(candidate.columns) == 13 and len(candidate.rows) >= 11:
            rows = _row_map(candidate)
            if "market cap" in rows and "ev / sales (x)" in rows:
                table = candidate
                break
    if table is None:
        return

    companies = [data] + peers[:3]
    rows = _row_map(table)
    years = (2023, 2024, 2025)

    for company_idx, company in enumerate(companies):
        start_col = 1 + company_idx * 3
        for col in range(start_col, min(start_col + 3, len(table.columns))):
            _set_cell_text(table.cell(0, col), company.ticker)

        for year_offset, year in enumerate(years):
            col = start_col + year_offset
            if col >= len(table.columns):
                break
            metrics = _year_peer_metrics(company, year)
            for label, value in metrics.items():
                if label in rows and value is not None:
                    _set_cell_text(table.cell(rows[label], col), _format_peer_metric(label, value))


def _fill_peer_median_average(doc: DocumentObject, peers: list[MarketData]) -> None:
    if not peers:
        return
    table = None
    for candidate in _all_tables(doc):
        if len(candidate.columns) == 7 and len(candidate.rows) >= 10:
            header = " ".join(cell.text.lower() for cell in candidate.rows[0].cells)
            rows = _row_map(candidate)
            if "median" in header and "average" in header and "ev / sales (x)" in rows:
                table = candidate
                break
    if table is None:
        return

    rows = _row_map(table)
    for year_offset, year in enumerate((2023, 2024, 2025)):
        metric_sets = [_year_peer_metrics(peer, year) for peer in peers]
        for label in rows:
            values = [m.get(label) for m in metric_sets]
            clean = [float(v) for v in values if v is not None and math.isfinite(float(v))]
            if not clean:
                continue
            median_value = statistics.median(clean)
            average_value = statistics.fmean(clean)
            median_col = 1 + year_offset
            average_col = 4 + year_offset
            _set_cell_text(table.cell(rows[label], median_col), _format_peer_metric(label, median_value))
            _set_cell_text(table.cell(rows[label], average_col), _format_peer_metric(label, average_value))


def _recommendation_counts(summary: pd.DataFrame | None) -> tuple[int, int, int, int] | None:
    if summary is None or summary.empty:
        return None
    row = summary.iloc[0]
    lower = {str(k).lower(): _to_float(v) or 0.0 for k, v in row.items()}
    strong_buy = int(lower.get("strongbuy", 0))
    buy = int(lower.get("buy", 0))
    hold = int(lower.get("hold", 0))
    sell = int(lower.get("sell", 0))
    strong_sell = int(lower.get("strongsell", 0))
    buys = strong_buy + buy
    sells = sell + strong_sell
    total = buys + hold + sells
    return total, buys, hold, sells


def _target_value(targets: dict[str, Any] | None, *keys: str) -> float | None:
    if not targets:
        return None
    lowered = {str(k).lower(): v for k, v in targets.items()}
    for key in keys:
        value = _to_float(lowered.get(key.lower()))
        if value is not None:
            return value
    return None



def _fill_analyst_actions(doc: DocumentObject, data: MarketData) -> None:
    table = None
    for candidate in _all_tables(doc):
        if len(candidate.columns) == 6 and len(candidate.rows) >= 9 and candidate.cell(0, 0).text.strip() == "INSTITUTION":
            table = candidate
            break
    if table is None:
        return

    df = data.upgrades_downgrades
    if df is None or df.empty:
        return

    frame = df.copy().reset_index()
    lower = {str(col).lower(): col for col in frame.columns}
    firm_col = next((col for key, col in lower.items() if key in ("firm", "institution", "brokerage") or "firm" in key), None)
    to_col = next((col for key, col in lower.items() if "tograde" in key or key == "rating"), None)
    from_col = next((col for key, col in lower.items() if "fromgrade" in key), None)
    action_col = next((col for key, col in lower.items() if key == "action" or "action" in key), None)
    date_col = next((col for key, col in lower.items() if "date" in key or key in ("index", "gradedate")), None)

    if firm_col is None:
        return

    for out_row, (_, row) in enumerate(frame.head(6).iterrows(), start=1):
        firm = str(row.get(firm_col, "") or "").strip()
        if not firm or firm.lower() == "nan":
            continue
        _set_cell_text(table.cell(out_row, 0), firm)

        if date_col is not None:
            raw_date = row.get(date_col)
            try:
                date_value = pd.to_datetime(raw_date)
                _set_cell_text(table.cell(out_row, 1), date_value.strftime("%Y-%m-%d"))
            except Exception:
                pass

        rating = str(row.get(to_col, "") or "").strip() if to_col is not None else ""
        if rating and rating.lower() != "nan":
            _set_cell_text(table.cell(out_row, 2), rating)

        action = str(row.get(action_col, "") or "").strip() if action_col is not None else ""
        previous = str(row.get(from_col, "") or "").strip() if from_col is not None else ""
        if action and action.lower() != "nan":
            _set_cell_text(table.cell(out_row, 4), action)
        note_parts = []
        if previous and previous.lower() != "nan":
            note_parts.append(f"From {previous}")
        if action and action.lower() != "nan":
            note_parts.append(action)
        if note_parts:
            _set_cell_text(table.cell(out_row, 5), " · ".join(note_parts))


def _fill_analyst_summary(doc: DocumentObject, data: MarketData) -> None:
    table = None
    for candidate in _all_tables(doc):
        if len(candidate.columns) == 6 and len(candidate.rows) >= 9 and candidate.cell(0, 0).text.strip() == "INSTITUTION":
            table = candidate
            break
    if table is None:
        return

    targets = data.analyst_targets or {}
    mean = _target_value(targets, "mean", "average", "targetMeanPrice")
    median = _target_value(targets, "median", "targetMedianPrice")
    high = _target_value(targets, "high", "targetHighPrice")
    low = _target_value(targets, "low", "targetLowPrice")
    counts = _recommendation_counts(data.recommendations_summary)

    if counts:
        total, buys, hold, sells = counts
        _set_cell_text(table.cell(7, 1), f"n = {total}")
        _set_cell_text(table.cell(7, 2), f"Buy {buys} / Hold {hold} / Sell {sells}")
    if mean is not None:
        _set_cell_text(table.cell(7, 3), _format_price(mean))
        if data.current_price is not None:
            _set_cell_text(table.cell(7, 4), _format_pct_ratio(_safe_div(mean - data.current_price, data.current_price)))
    if median is not None:
        _set_cell_text(table.cell(8, 3), _format_price(median))
    if high is not None or low is not None:
        high_text = _format_price(high) if high is not None else "—"
        low_text = _format_price(low) if low is not None else "—"
        _set_cell_text(table.cell(8, 2), f"High {high_text} / low {low_text}")


def _fill_source(doc: DocumentObject, sec: SecResearch) -> None:
    for table in _all_tables(doc):
        rows = _row_map(table)
        if "annual report · filings" in rows and "other" in rows and len(table.columns) == 2:
            if sec.annual_url:
                filing_label = f"{sec.annual_form} · {sec.annual_filing_date} · {sec.annual_url}"
                _set_cell_text(table.cell(rows["annual report · filings"], 1), filing_label[:1500])

            other_sources: list[str] = []
            if sec.event_url:
                other_sources.append(f"SEC {sec.event_form} · {sec.event_filing_date} · {sec.event_url}")
            if sec.proxy_url:
                other_sources.append(f"SEC {sec.proxy_form} · {sec.proxy_filing_date} · {sec.proxy_url}")
            other_sources.append(f"Yahoo Finance via yfinance · {datetime.now().strftime('%Y-%m-%d')}")
            _set_cell_text(table.cell(rows["other"], 1), "\n".join(other_sources)[:1500])
            break



def _localize_document_to_korean(doc: DocumentObject) -> None:
    """Localize fixed template labels after all English-keyed fill routines run.

    Automatically generated prose is translated with googletrans before this stage;
    this function handles deterministic template labels/headings without sending
    boilerplate text to an external translator.
    """
    exact_map = {
        "ARGUMENT": "핵심 논거",
        "ITEM": "항목",
        "VALUE": "값",
        "EVIDENCE · SOURCE": "근거 · 출처",
        "CATALYST": "촉매",
        "EXPECTED DATE": "예상 시점",
        "PROB.": "확률",
        "PRICE IMPACT": "주가 영향",
        "HOW WE VERIFY IT": "확인 방법",
        "RISK": "리스크",
        "LIKELIHOOD": "발생 가능성",
        "SEVERITY": "심각도",
        "TP IMPACT": "목표주가 영향",
        "MITIGATION · MONITORED INDICATOR": "대응 · 모니터링 지표",
        "Founding · listing history": "설립 · 상장 이력",
        "Segments and revenue mix": "사업부문 · 매출 구성",
        "Revenue by geography": "지역별 매출",
        "Key customers · channels": "주요 고객 · 판매 채널",
        "Competitors and market share": "경쟁사 · 시장점유율",
        "Cost structure · supply chain": "비용 구조 · 공급망",
        "Ownership and major shareholders": "소유구조 · 주요 주주",
        "Management and incentives": "경영진 · 보상체계",
        "Revenue": "매출액",
        "Revenue growth (%)": "매출 성장률 (%)",
        "EBITDA margin (%)": "EBITDA 마진 (%)",
        "Operating income": "영업이익",
        "Net income": "순이익",
        "Free cash flow": "잉여현금흐름",
        "Net debt / EBITDA (x)": "순부채 / EBITDA (x)",
        "Cash and equivalents": "현금 및 현금성자산",
        "Accounts receivable": "매출채권",
        "Inventory": "재고자산",
        "Total current assets": "유동자산 합계",
        "Property, plant and equipment": "유형자산",
        "Goodwill and intangibles": "영업권 및 무형자산",
        "Total assets": "자산총계",
        "Accounts payable": "매입채무",
        "Total current liabilities": "유동부채 합계",
        "Total debt (incl. leases)": "총부채성 차입금 (리스 포함)",
        "Total liabilities": "부채총계",
        "Total equity": "자본총계",
        "Net debt": "순부채",
        "Current ratio (x)": "유동비율 (x)",
        "Closing price": "종가",
        "Currency": "통화",
        "Average daily value traded": "평균 일일 거래대금",
        "Market cap": "시가총액",
        "Shares outstanding": "발행주식수",
        "Free float": "유통주식수",
        "Dividend yield": "배당수익률",
        "Annual report · filings": "연차보고서 · 공시",
        "Earnings call transcript": "실적발표 컨퍼런스콜",
        "Industry report": "산업 보고서",
        "Government · trade data": "정부 · 무역 데이터",
        "Channel checks · interviews": "채널 체크 · 인터뷰",
        "Other": "기타",
        "High / med / low": "높음 / 중간 / 낮음",
        "Buy / Hold / Sell": "매수 / 중립 / 매도",
        "INSTITUTION": "기관",
        "DATE": "일자",
        "RATING": "의견",
        "TARGET PRICE": "목표주가",
        "CHANGE": "변경",
        "NOTE": "비고",
        "Street average": "시장 평균",
        "Median · range": "중앙값 · 범위",
    }
    phrase_map = {
        "03 — COMPANY BACKGROUND": "03 — 기업 개요",
        "04 — CATALYSTS": "04 — 촉매",
        "06 — ANALYST RECOMMENDATIONS": "06 — 애널리스트 의견",
        "07 — RISK FACTORS": "07 — 리스크 요인",
        "08 — MONITORING AND SOURCES": "08 — 모니터링 및 출처",
        "Highlights — key arguments": "핵심 포인트 — 주요 논거",
        "2.1 Key financials — income statement": "2.1 주요 재무지표 — 손익계산서",
        "2.2 Key financials — balance sheet": "2.2 주요 재무지표 — 재무상태표",
        "Latest published note per institution. Fill the average and median rows at the bottom.": "기관별 최신 공개 의견입니다. 하단에는 평균과 중앙값을 표시합니다.",
        "What would make us wrong:": "우리의 판단을 틀리게 만드는 조건:",
        "Business model, revenue mix, customers, competitive structure, recent history. Explain how the company makes money before anything else.": "사업모델, 매출 구성, 고객, 경쟁구도 및 최근 이력을 정리합니다. 무엇보다 회사가 어떻게 수익을 창출하는지 먼저 설명합니다.",
    }

    def localize_paragraph(paragraph: Paragraph) -> None:
        original = paragraph.text.strip()
        if not original:
            return
        if original in exact_map:
            _set_paragraph_text(paragraph, exact_map[original])
            return
        if original.startswith("Company Name:"):
            _set_paragraph_text(paragraph, original.replace("Company Name:", "회사명:", 1).replace("Ticker:", "티커:"))
            return
        if original.startswith("Ticker ") and "Exchange" in original:
            replaced = original.replace("Ticker ", "티커 ", 1).replace("Exchange ", "거래소 ").replace("GICS sector / industry", "GICS 섹터 / 산업").replace("Reporting currency", "보고 통화")
            _set_paragraph_text(paragraph, replaced)
            return
        for english, korean in phrase_map.items():
            if english in original:
                _set_paragraph_text(paragraph, original.replace(english, korean))
                return

    for paragraph in doc.paragraphs:
        localize_paragraph(paragraph)
    for table in _all_tables(doc):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    localize_paragraph(paragraph)

def fill_document(
    template_path: Path,
    output_path: Path,
    data: MarketData,
    *,
    background: CompanyBackground,
    peers: list[MarketData],
    sec: SecResearch,
    catalysts: list[Catalyst] | None = None,
    risks: list[RiskFactor] | None = None,
    highlights: list[HighlightArgument] | None = None,
    language: str = "en",
) -> None:
    doc = Document(str(template_path))
    _fill_header(doc, data)
    _fill_market_profile(doc, data)
    _fill_highlights(doc, highlights or [])
    _fill_company_background(doc, background)
    _fill_catalysts(doc, catalysts or [])
    _fill_risk_factors(doc, risks or [])
    _fill_income_statement(doc, data)
    _fill_balance_sheet(doc, data)
    _fill_quarterly_valuation(doc, data)
    # DCF sections are intentionally left untouched.
    _fill_target_company_peer_columns(doc, data, peers)
    _fill_peer_median_average(doc, peers)
    _fill_analyst_actions(doc, data)
    _fill_analyst_summary(doc, data)
    _fill_source(doc, sec)
    _style_numeric_cells(doc)
    if language.lower().strip() == "ko":
        _localize_document_to_korean(doc)
    doc.save(str(output_path))


def build_filled_template_from_components(
    output_path: Path,
    data: MarketData,
    background: CompanyBackground,
    peers: list[MarketData],
    sec: SecResearch,
    catalysts: list[Catalyst] | None = None,
    risks: list[RiskFactor] | None = None,
    highlights: list[HighlightArgument] | None = None,
    language: str = "en",
) -> MarketData:
    """Render a DOCX from already-fetched research data.

    The web server uses this path so Research -> Export does not hit Yahoo/SEC again.
    """
    template_path = _resolve_template_path()
    data.selected_peers = [peer.ticker for peer in peers[:3]]
    fill_document(
        template_path,
        output_path,
        data,
        background=background,
        peers=peers[:3],
        sec=sec,
        catalysts=catalysts or build_catalysts(data, sec),
        risks=risks or build_risk_factors(data, sec),
        highlights=highlights or build_highlights(data, peers[:3], sec),
        language=language,
    )
    return data


def build_filled_template(raw_ticker: str, output_path: Path) -> MarketData:
    """Backward-compatible one-shot builder used by the Discord version."""
    data = fetch_market_data(raw_ticker)

    # SEC and Yahoo are independent: either one may fail without blocking the whole build.
    sec = fetch_sec_research(data.ticker)
    background = build_company_background(data, sec)

    try:
        peers = select_peer_companies(data, count=3)
    except Exception as exc:
        print(f"[PEER] automatic peer selection failed: {type(exc).__name__}: {exc}")
        peers = []

    catalysts = build_catalysts(data, sec)
    risks = build_risk_factors(data, sec)
    highlights = build_highlights(data, peers, sec)
    return build_filled_template_from_components(output_path, data, background, peers, sec, catalysts, risks, highlights)
