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
from docx.shared import Pt, RGBColor
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
    quarterly_cashflow: pd.DataFrame | None = None
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


@dataclass(slots=True)
class FinancialSignal:
    priority: int
    tag: str
    en: str
    ko: str
    severity: str = "NEUTRAL"
    category: str = ""
    mandatory: bool = False


FINANCIAL_COMMENT_MAX = max(4, min(int(os.getenv("FINANCIAL_COMMENT_MAX", "10")), 14))


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
    quarterly_cashflow = _get_statement(ticker, "quarterly_cashflow", "quarterly_cash_flow")

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
        quarterly_cashflow=quarterly_cashflow,
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



def _statement_period_columns(df: pd.DataFrame | None, *, quarterly: bool) -> list[Any]:
    """Return all usable statement columns in chronological order.

    Yahoo normally exposes 4-5 annual periods and 5-8 quarters depending on the
    security.  We intentionally keep every dated column Yahoo returns instead of
    truncating the detailed Research view.
    """
    if df is None or df.empty:
        return []
    dated: list[tuple[pd.Timestamp, Any]] = []
    for col in df.columns:
        try:
            ts = pd.Timestamp(col)
        except Exception:
            continue
        if pd.isna(ts):
            continue
        dated.append((ts, col))
    dated.sort(key=lambda item: item[0])
    # Deduplicate equivalent timestamps while preserving the original dataframe key.
    out: list[Any] = []
    seen: set[pd.Timestamp] = set()
    for ts, col in dated:
        key = ts.normalize()
        if key in seen:
            continue
        seen.add(key)
        out.append(col)
    return out


def _statement_period_label(col: Any, *, quarterly: bool) -> str:
    ts = pd.Timestamp(col)
    if quarterly:
        q = (ts.month - 1) // 3 + 1
        return f"{q}Q{ts.year}A"
    return f"FY{ts.year}A"


def _statement_kind(label: str) -> str:
    key = label.lower()
    if "eps" in key or "earnings per share" in key:
        return "number"
    if "shares" in key or "share issued" in key or "stock issued" in key:
        return "shares"
    return "money"


def _statement_display(value: float | None, kind: str) -> tuple[float | None, str]:
    if value is None:
        return None, ""
    if kind == "number":
        return value, _format_number(value, 2)
    if kind == "shares":
        scaled = value / 1_000_000
        return scaled, _format_number(scaled, 2)
    scaled = value / 1_000_000
    return scaled, _format_millions(value)


def _serialize_statement(df: pd.DataFrame | None, *, quarterly: bool) -> dict[str, Any]:
    cols = _statement_period_columns(df, quarterly=quarterly)
    if df is None or df.empty or not cols:
        return {"periods": [], "rows": [], "row_count": 0}
    rows: list[dict[str, Any]] = []
    for idx in df.index:
        label = str(idx).strip()
        if not label:
            continue
        kind = _statement_kind(label)
        values: list[dict[str, Any]] = []
        has_value = False
        for col in cols:
            try:
                value = _to_float(df.at[idx, col])
            except Exception:
                try:
                    value = _to_float(df.loc[idx, col])
                except Exception:
                    value = None
            raw, display = _statement_display(value, kind)
            if raw is not None:
                has_value = True
            values.append({"raw": raw, "display": display, "kind": kind})
        if has_value:
            rows.append({"label": label, "values": values})
    return {
        "periods": [_statement_period_label(col, quarterly=quarterly) for col in cols],
        "rows": rows,
        "row_count": len(rows),
    }


def _chart_series(
    df: pd.DataFrame | None,
    metrics: list[tuple[str, tuple[str, ...]]],
    *,
    quarterly: bool,
) -> dict[str, Any]:
    cols = _statement_period_columns(df, quarterly=quarterly)
    periods = [_statement_period_label(col, quarterly=quarterly) for col in cols]
    series: list[dict[str, Any]] = []
    for label, aliases in metrics:
        values: list[float | None] = []
        found = False
        for col in cols:
            value = _line_value(df, aliases, col)
            scaled = value / 1_000_000 if value is not None else None
            values.append(scaled)
            found = found or scaled is not None
        if found:
            series.append({"label": label, "values": values})
    return {"periods": periods, "series": series}


def _income_overview_chart(data: MarketData, *, quarterly: bool) -> dict[str, Any]:
    df = data.quarterly_income if quarterly else data.annual_income
    metrics = [
        ("Revenue", ("Total Revenue", "Operating Revenue", "Revenue")),
        ("Gross Profit", ("Gross Profit",)),
        ("EBITDA", ("EBITDA", "Normalized EBITDA")),
        ("Operating Income", ("Operating Income", "Total Operating Income As Reported")),
        ("Pretax Income", ("Pretax Income", "Income Before Tax")),
        ("Net Income", ("Net Income", "Net Income Common Stockholders", "Net Income Including Noncontrolling Interests")),
    ]
    return _chart_series(df, metrics, quarterly=quarterly)


def _balance_overview_chart(data: MarketData, *, quarterly: bool) -> dict[str, Any]:
    df = data.quarterly_balance if quarterly else data.annual_balance
    metrics = [
        ("Cash & ST Investments", BALANCE_ALIASES["Cash and equivalents"]),
        ("Current Assets", BALANCE_ALIASES["Total current assets"]),
        ("Total Assets", BALANCE_ALIASES["Total assets"]),
        ("Total Debt", BALANCE_ALIASES["Total debt (incl. leases)"]),
        ("Current Liabilities", BALANCE_ALIASES["Total current liabilities"]),
        ("Total Liabilities", BALANCE_ALIASES["Total liabilities"]),
        ("Total Equity", BALANCE_ALIASES["Total equity"]),
    ]
    return _chart_series(df, metrics, quarterly=quarterly)


def _cashflow_overview_chart(data: MarketData, *, quarterly: bool) -> dict[str, Any]:
    df = data.quarterly_cashflow if quarterly else data.annual_cashflow
    metrics = [
        ("Operating Cash Flow", ("Operating Cash Flow", "Total Cash From Operating Activities")),
        ("Capital Expenditure", ("Capital Expenditure", "Capital Expenditures")),
        ("Free Cash Flow", ("Free Cash Flow",)),
        ("Stock Based Compensation", ("Stock Based Compensation",)),
        ("Dividends Paid", ("Cash Dividends Paid", "Common Stock Dividend Paid")),
        ("Share Repurchases", ("Repurchase Of Capital Stock", "Repurchase Of Stock")),
    ]
    return _chart_series(df, metrics, quarterly=quarterly)


def _expense_breakdown(data: MarketData, *, quarterly: bool) -> dict[str, Any]:
    """Expose the operating-cost lines Yahoo actually provides.

    This is intentionally evidence-preserving: if Yahoo does not break SG&A into
    sales/marketing and G&A for a company, those child lines remain absent rather
    than being estimated.
    """
    df = data.quarterly_income if quarterly else data.annual_income
    cols = _statement_period_columns(df, quarterly=quarterly)
    periods = [_statement_period_label(col, quarterly=quarterly) for col in cols]
    candidates: list[tuple[str, tuple[str, ...]]] = [
        ("Cost of Revenue", ("Cost Of Revenue", "Cost Of Goods Sold")),
        ("Total Operating Expense", ("Operating Expense", "Total Operating Expenses")),
        ("SG&A", ("Selling General And Administration", "Selling General Administrative")),
        ("Selling & Marketing", ("Selling And Marketing Expense", "Selling And Marketing")),
        ("General & Administrative", ("General And Administrative Expense", "General And Administrative")),
        ("R&D", ("Research And Development", "Research Development")),
        ("D&A in Income Statement", ("Depreciation And Amortization In Income Statement",)),
        ("Other Operating Expense", ("Other Operating Expenses", "Other Operating Expense")),
        ("Restructuring / M&A", ("Restructuring And Merger Acquisition", "Restructuring")),
    ]
    series: list[dict[str, Any]] = []
    revenue_aliases = INCOME_ALIASES["Revenue"]
    for label, aliases in candidates:
        vals: list[float | None] = []
        pcts: list[float | None] = []
        found = False
        for col in cols:
            value = _line_value(df, aliases, col)
            revenue = _line_value(df, revenue_aliases, col)
            scaled = abs(value) / 1_000_000 if value is not None else None
            pct_rev = abs(value) / abs(revenue) * 100 if value is not None and revenue not in (None, 0) else None
            vals.append(scaled)
            pcts.append(pct_rev)
            found = found or scaled is not None
        if found:
            series.append({"label": label, "values": vals, "pct_revenue": pcts})

    latest: list[dict[str, Any]] = []
    if cols:
        for item in series:
            value = item["values"][-1] if item["values"] else None
            pct_rev = item["pct_revenue"][-1] if item["pct_revenue"] else None
            if value is not None:
                latest.append({"label": item["label"], "value": value, "pct_revenue": pct_rev})
        latest.sort(key=lambda item: abs(item.get("value") or 0), reverse=True)
    return {"periods": periods, "series": series, "latest": latest}


def _peer_metric_value(data: MarketData, key: str) -> float | None:
    info = data.info or {}
    direct = _to_float(info.get(key))
    if direct is not None:
        return direct
    return None


def _latest_statement_value(df: pd.DataFrame | None, aliases: Iterable[str]) -> float | None:
    cols = _statement_period_columns(df, quarterly=False)
    if not cols:
        return None
    return _line_value(df, aliases, cols[-1])


def _peer_revenue_series(data: MarketData, *, quarterly: bool, limit: int) -> list[dict[str, Any]]:
    df = data.quarterly_income if quarterly else data.annual_income
    cols = _statement_period_columns(df, quarterly=quarterly)
    if limit > 0:
        cols = cols[-limit:]
    out: list[dict[str, Any]] = []
    for col in cols:
        value = _line_value(df, INCOME_ALIASES["Revenue"], col)
        out.append({
            "period": _statement_period_label(col, quarterly=quarterly),
            "value": None if value is None else value / 1_000_000,
        })
    return out


def _normalized_price_series(data: MarketData) -> list[dict[str, Any]]:
    history = data.price_history
    if history is None or history.empty or "Close" not in history.columns:
        return []
    try:
        closes = history[["Close"]].copy()
        closes.index = pd.to_datetime(closes.index)
        if getattr(closes.index, "tz", None) is not None:
            closes.index = closes.index.tz_localize(None)
        closes = closes.dropna()
        if closes.empty:
            return []
        # Weekly Friday closes keep a 3-year, 8-security comparison compact and readable.
        weekly = closes["Close"].resample("W-FRI").last().dropna()
        if weekly.empty:
            weekly = closes["Close"].dropna()
        base = _to_float(weekly.iloc[0])
        if base is None or abs(base) < 1e-12:
            return []
        return [
            {"date": idx.strftime("%Y-%m-%d"), "value": float(value) / base * 100.0}
            for idx, value in weekly.items()
            if _to_float(value) is not None
        ]
    except Exception:
        return []


def build_peer_analysis_payload(companies: list[MarketData]) -> dict[str, Any]:
    """Build the F6 peer-comparison workstation payload.

    The first company is the subject security and up to seven following companies are
    analyst-selected peers. Values are best-effort Yahoo/yfinance observations; no
    missing multiple or margin is estimated from unsupported data.
    """
    rows: list[dict[str, Any]] = []
    for data in companies[:8]:
        info = data.info or {}
        market_cap = data.market_cap
        total_debt = _first_number(info.get("totalDebt"), _latest_statement_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"]))
        cash = _first_number(info.get("totalCash"), _latest_statement_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"]))
        enterprise_value = _first_number(info.get("enterpriseValue"))
        if enterprise_value is None and market_cap is not None:
            enterprise_value = market_cap + (total_debt or 0.0) - (cash or 0.0)
        ebitda = _first_number(info.get("ebitda"), _latest_statement_value(data.annual_income, INCOME_ALIASES["EBITDA"]))
        free_cash_flow = _first_number(info.get("freeCashflow"), _latest_statement_value(data.annual_cashflow, CASHFLOW_ALIASES["Free cash flow"]))
        net_debt = None
        if total_debt is not None or cash is not None:
            net_debt = (total_debt or 0.0) - (cash or 0.0)
        net_debt_ebitda = _safe_div(net_debt, ebitda)
        fcf_yield = _safe_div(free_cash_flow, market_cap)
        if fcf_yield is not None:
            fcf_yield *= 100.0

        rows.append({
            "ticker": data.ticker,
            "company_name": data.company_name,
            "sector": data.sector,
            "industry": data.industry,
            "currency": data.currency,
            "current_price": data.current_price,
            "market_cap": market_cap,
            "enterprise_value": enterprise_value,
            "net_debt": net_debt,
            "trailing_pe": _to_float(info.get("trailingPE")),
            "forward_pe": _to_float(info.get("forwardPE")),
            "ev_revenue": _to_float(info.get("enterpriseToRevenue")),
            "ev_ebitda": _to_float(info.get("enterpriseToEbitda")),
            "price_sales": _to_float(info.get("priceToSalesTrailing12Months")),
            "price_book": _to_float(info.get("priceToBook")),
            "revenue_growth": _to_float(info.get("revenueGrowth")),
            "earnings_growth": _to_float(info.get("earningsGrowth")),
            "gross_margin": _to_float(info.get("grossMargins")),
            "ebitda_margin": _to_float(info.get("ebitdaMargins")),
            "operating_margin": _to_float(info.get("operatingMargins")),
            "net_margin": _to_float(info.get("profitMargins")),
            "roe": _to_float(info.get("returnOnEquity")),
            "roa": _to_float(info.get("returnOnAssets")),
            "net_debt_ebitda": net_debt_ebitda,
            "fcf_yield": fcf_yield,
            "beta": data.beta,
            "annual_revenue": _peer_revenue_series(data, quarterly=False, limit=3),
            "quarterly_revenue": _peer_revenue_series(data, quarterly=True, limit=8),
            "normalized_price": _normalized_price_series(data),
        })

    return {
        "subject": rows[0]["ticker"] if rows else "",
        "company_count": len(rows),
        "peer_count": max(0, len(rows) - 1),
        "companies": rows,
        "notes": [
            "Revenue is shown in each company's reporting currency, USD-equivalent conversion is not applied.",
            "Price performance is weekly close normalized to 100 at each security's first available observation in the 3-year window.",
            "Yahoo Finance field coverage differs by issuer; unavailable metrics remain blank rather than being estimated.",
        ],
    }

def build_key_financials_payload(data: MarketData) -> dict[str, Any]:
    """Detailed annual + quarterly financial statement payload for the web UI.

    The Research view now exposes every financial statement line Yahoo/yfinance
    returns, plus focused annual and quarterly charts.  Monetary statement values
    and share counts are normalized to millions; EPS remains per-share.
    """
    annual_income_detail = _serialize_statement(data.annual_income, quarterly=False)
    quarterly_income_detail = _serialize_statement(data.quarterly_income, quarterly=True)
    annual_balance_detail = _serialize_statement(data.annual_balance, quarterly=False)
    quarterly_balance_detail = _serialize_statement(data.quarterly_balance, quarterly=True)
    annual_cashflow_detail = _serialize_statement(data.annual_cashflow, quarterly=False)
    quarterly_cashflow_detail = _serialize_statement(data.quarterly_cashflow, quarterly=True)

    consensus = []
    for year, period in ((2026, "0y"), (2027, "+1y")):
        rev = _estimate_avg(data.revenue_estimate, period)
        eps = _estimate_avg(data.earnings_estimate, period)
        consensus.append({
            "period": f"FY{year}E",
            "revenue_raw": None if rev is None else rev / 1_000_000,
            "revenue": _format_millions(rev),
            "eps_raw": eps,
            "eps": _format_number(eps, 2),
        })

    return {
        "currency": data.currency,
        "units": "MM except EPS/per-share values; share-count rows are MM shares",
        "consensus": consensus,
        "detailed": {
            "income": {
                "annual": annual_income_detail,
                "quarterly": quarterly_income_detail,
                "annual_chart": _income_overview_chart(data, quarterly=False),
                "quarterly_chart": _income_overview_chart(data, quarterly=True),
                "expense_annual": _expense_breakdown(data, quarterly=False),
                "expense_quarterly": _expense_breakdown(data, quarterly=True),
            },
            "balance": {
                "annual": annual_balance_detail,
                "quarterly": quarterly_balance_detail,
                "annual_chart": _balance_overview_chart(data, quarterly=False),
                "quarterly_chart": _balance_overview_chart(data, quarterly=True),
            },
            "cashflow": {
                "annual": annual_cashflow_detail,
                "quarterly": quarterly_cashflow_detail,
                "annual_chart": _cashflow_overview_chart(data, quarterly=False),
                "quarterly_chart": _cashflow_overview_chart(data, quarterly=True),
            },
        },
        "counts": {
            "income_annual": annual_income_detail["row_count"],
            "income_quarterly": quarterly_income_detail["row_count"],
            "balance_annual": annual_balance_detail["row_count"],
            "balance_quarterly": quarterly_balance_detail["row_count"],
            "cashflow_annual": annual_cashflow_detail["row_count"],
            "cashflow_quarterly": quarterly_cashflow_detail["row_count"],
        },
        "revenue_growth_highlight_pct": REVENUE_GROWTH_HIGHLIGHT_PCT,
    }


def _growth_rate(current: float | None, previous: float | None) -> float | None:
    if current is None or previous is None or abs(previous) < 1e-12:
        return None
    return current / previous - 1.0


def _infer_financial_severity(tag: str) -> str:
    t = tag.strip().upper()
    critical = (
        "CRITICAL BALANCE SHEET STRESS", "PERSISTENT CASH BURN", "STRUCTURAL CONTRACTION",
        "PROFITABILITY COLLAPSE", "VERY HIGH LEVERAGE WITH EARNINGS DECLINE",
    )
    if any(token in t for token in critical):
        return "CRITICAL"
    severe = (
        "SIGNIFICANT CONTRACTION", "SEVERE ", "OPERATING LOSS", "NET LOSS",
        "BALANCE SHEET DETERIORATION", "VERY HIGH LEVERAGE", "VERY HIGH LIABILITY",
        "WORKING-CAPITAL STRESS", "EXPENSIVE WITH DETERIORATING FUNDAMENTALS",
        "MULTI-YEAR MARGIN EROSION",
    )
    if any(token in t for token in severe):
        return "SEVERE"
    warning = (
        "WARNING", "RISK", "DETERIORATION", "DECELERATION", "NEGATIVE OPERATING LEVERAGE",
        "BUILD-UP", "COMPRESSION", "DILUTION", "LOW CASH COVERAGE", "TIGHT LIQUIDITY",
        "HIGH LEVERAGE", "HIGH LIABILITY", "LOW ROIC", "PRESSURE", "DIVERGENCE",
        "IMPAIRMENT", "PREMIUM AT RISK", "VALUATION TRAP", "LOW CAPITAL RETURNS",
        "VERY HIGH EV / EBITDA", "VERY HIGH P / E",
    )
    if any(token in t for token in warning):
        return "WARNING"
    weak = (
        "LOW GROWTH", "WEAK ", "MILD CONTRACTION", "LOW FCF", "LOW LIQUIDITY",
        "ELEVATED LEVERAGE", "ELEVATED INTANGIBLES",
    )
    if any(token in t for token in weak):
        return "WEAK"
    positive = (
        "HIGH-QUALITY", "HYPER GROWTH", "HIGH GROWTH", "STRONG GROWTH",
        "GROWTH ACCELERATION", "POSITIVE OPERATING LEVERAGE", "TURNAROUND", "ACCRETION",
        "EXCEPTIONAL", "STRONG FCF", "HIGH EARNINGS QUALITY", "STRONG ROIC",
        "BALANCE SHEET STRENGTHENING", "STRONG LIQUIDITY", "HEALTHY LIQUIDITY",
        "FULL DEBT CASH COVERAGE", "EFFICIENCY", "PRODUCTIVITY", "LEVERAGE IMPROVEMENT",
        "DELEVERAGING", "LOW LEVERAGE", "NET CASH", "AGGRESSIVE BUYBACK",
        "GROWING INTO VALUATION", "FUNDAMENTAL / PRICE DIVERGENCE",
        "RELATIVE VALUATION DISCOUNT", "GROWTH-ADJUSTED DISCOUNT", "SUSTAINED HIGH GROWTH",
    )
    if any(token in t for token in positive):
        return "POSITIVE"
    return "NEUTRAL"


def _signal(
    priority: int,
    tag: str,
    en: str,
    ko: str,
    *,
    severity: str | None = None,
    category: str = "",
    mandatory: bool = False,
) -> FinancialSignal:
    return FinancialSignal(
        priority=priority, tag=tag, en=en, ko=ko,
        severity=(severity or _infer_financial_severity(tag)).upper(),
        category=category, mandatory=mandatory,
    )


def _dedupe_signals(signals: list[FinancialSignal], limit: int | None = None) -> list[FinancialSignal]:
    ordered = sorted(enumerate(signals), key=lambda item: (item[1].priority, item[0]))
    result: list[FinancialSignal] = []
    seen: set[str] = set()
    for _, item in ordered:
        key = item.tag.strip().upper()
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(item)
        if limit is not None and len(result) >= limit:
            break
    return result


def _select_financial_signals(
    signals: list[FinancialSignal],
    section: str,
    limit: int = FINANCIAL_COMMENT_MAX,
) -> list[FinancialSignal]:
    """Keep the most important signals while guaranteeing core diagnostics per table."""
    unique = _dedupe_signals(signals)
    by_tag = {item.tag.upper(): item for item in unique}

    core_groups: dict[str, list[set[str]]] = {
        "income": [
            {"HYPER GROWTH", "HIGH GROWTH", "STRONG GROWTH", "MODERATE GROWTH", "LOW GROWTH", "MILD CONTRACTION", "SIGNIFICANT CONTRACTION", "STRUCTURAL CONTRACTION"},
            {"GROWTH QUALITY WARNING", "HIGH-QUALITY GROWTH", "PROFITABILITY COLLAPSE", "MULTI-YEAR MARGIN EROSION", "EBITDA MARGIN", "OPERATING LOSS", "OPERATING TURNAROUND", "OPERATING INCOME PRESSURE"},
            {"PERSISTENT CASH BURN", "CASH BURN", "EXCEPTIONAL CASH GENERATION", "HEALTHY CASH GENERATION", "MODERATE CASH GENERATION", "WEAK CASH GENERATION", "CASH CONVERSION WARNING", "HIGH EARNINGS QUALITY"},
        ],
        "balance": [
            {"LIQUIDITY RISK", "TIGHT LIQUIDITY", "ADEQUATE LIQUIDITY", "HEALTHY LIQUIDITY", "EXCESS LIQUIDITY", "FULL DEBT CASH COVERAGE", "STRONG LIQUIDITY", "LOW CASH COVERAGE", "LIQUIDITY DETERIORATION"},
            {"CRITICAL BALANCE SHEET STRESS", "VERY HIGH LEVERAGE WITH EARNINGS DECLINE", "VERY HIGH LEVERAGE", "HIGH LEVERAGE", "ELEVATED LEVERAGE", "LOW LEVERAGE", "NET CASH", "LEVERAGE DETERIORATION", "DELEVERAGING"},
            {"LEVERAGE-DRIVEN ROE", "HIGH-QUALITY CAPITAL RETURNS", "ROE PROFILE"},
            {"LEVERAGE-DRIVEN ROE", "HIGH-QUALITY CAPITAL RETURNS", "ROIC PROFILE"},
        ],
        "valuation": [
            {"EV / EBITDA N/M", "VERY LOW EV / EBITDA", "LOW EV / EBITDA", "HIGH EV / EBITDA", "VERY HIGH EV / EBITDA", "EV / EBITDA PROFILE", "SECTOR-SPECIFIC VALUATION"},
            {"P / E N/M", "VERY LOW P / E", "LOW P / E", "HIGH P / E", "VERY HIGH P / E", "P / E PROFILE"},
            {"FUNDAMENTAL / PRICE DIVERGENCE", "MULTIPLE COMPRESSION", "MULTIPLE EXPANSION", "EV / EBITDA COMPRESSION", "EV / EBITDA EXPANSION", "GROWING INTO VALUATION", "ENTERPRISE MULTIPLE COMPRESSION", "ENTERPRISE MULTIPLE EXPANSION", "VALUATION TRAP RISK", "PREMIUM AT RISK"},
        ],
    }

    selected: list[FinancialSignal] = []
    selected_tags: set[str] = set()

    def add(item: FinancialSignal | None) -> None:
        if item is None or item.tag.upper() in selected_tags or len(selected) >= limit:
            return
        selected.append(item)
        selected_tags.add(item.tag.upper())

    # Explicit mandatory signals first.
    for item in unique:
        if item.mandatory:
            add(item)

    # Guarantee one diagnostic from each core group, even if it has a low raw priority.
    for group in core_groups.get(section, []):
        candidates = [by_tag[tag] for tag in group if tag in by_tag]
        if candidates:
            add(min(candidates, key=lambda x: x.priority))

    # Then fill the remaining slots by raw risk/importance priority.
    for item in unique:
        add(item)
        if len(selected) >= limit:
            break

    return sorted(selected, key=lambda x: x.priority)


def _latest_template_actual_year(data: MarketData) -> int | None:
    # The stock-pitch template currently exposes FY2023A-FY2025A actual columns.
    years = [year for year in (2023, 2024, 2025) if _column_for_year(data.annual_income, year) is not None]
    return max(years) if years else None


def _is_financial_company(data: MarketData) -> bool:
    sector = (data.sector or "").strip().lower()
    industry = (data.industry or "").strip().lower()
    financial_tokens = ("financial", "bank", "insurance", "credit service", "capital market", "asset management")
    return "financial" in sector or any(token in industry for token in financial_tokens)


def _income_financial_signals(data: MarketData) -> list[FinancialSignal]:
    year = _latest_template_actual_year(data)
    if year is None:
        return []
    prev = year - 1
    prev2 = year - 2
    prev3 = year - 3
    rev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], year)
    prev_rev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], prev)
    prev2_rev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], prev2)
    prev3_rev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], prev3)
    rev_g = _growth_rate(rev, prev_rev)
    prev_rev_g = _growth_rate(prev_rev, prev2_rev)
    prev2_rev_g = _growth_rate(prev2_rev, prev3_rev)

    ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], year)
    prev_ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], prev)
    prev2_ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], prev2)
    ebitda_g = _growth_rate(ebitda, prev_ebitda)
    margin = _safe_div(ebitda, rev)
    prev_margin = _safe_div(prev_ebitda, prev_rev)
    prev2_margin = _safe_div(prev2_ebitda, prev2_rev)

    op_income = _annual_value(data.annual_income, INCOME_ALIASES["Operating income"], year)
    prev_op = _annual_value(data.annual_income, INCOME_ALIASES["Operating income"], prev)
    op_g = _growth_rate(op_income, prev_op) if prev_op is not None and prev_op > 0 else None
    net_income = _annual_value(data.annual_income, INCOME_ALIASES["Net income"], year)
    prev_ni = _annual_value(data.annual_income, INCOME_ALIASES["Net income"], prev)
    ni_g = _growth_rate(net_income, prev_ni) if prev_ni is not None and prev_ni > 0 else None
    eps = _annual_value(data.annual_income, INCOME_ALIASES["EPS"], year)
    prev_eps = _annual_value(data.annual_income, INCOME_ALIASES["EPS"], prev)
    eps_g = _growth_rate(eps, prev_eps) if prev_eps is not None and prev_eps > 0 else None
    shares = _annual_value(data.annual_income, INCOME_ALIASES["Diluted shares"], year)
    prev_shares = _annual_value(data.annual_income, INCOME_ALIASES["Diluted shares"], prev)
    shares_g = _growth_rate(shares, prev_shares)

    fcf = _annual_value(data.annual_cashflow, CASHFLOW_ALIASES["Free cash flow"], year)
    prev_fcf = _annual_value(data.annual_cashflow, CASHFLOW_ALIASES["Free cash flow"], prev)
    prev2_fcf = _annual_value(data.annual_cashflow, CASHFLOW_ALIASES["Free cash flow"], prev2)
    fcf_g = _growth_rate(fcf, prev_fcf) if prev_fcf is not None and prev_fcf > 0 else None
    fcf_margin = _safe_div(fcf, rev)
    conversion = _safe_div(fcf, net_income) if net_income is not None and net_income > 0 else None
    roic = _roic(data, year)
    cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], year)
    debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], year)
    net_debt = debt - cash if debt is not None and cash is not None else None
    nd_ebitda = _safe_div(net_debt, ebitda) if ebitda is not None and ebitda > 0 else None

    next_rev_est = _estimate_avg(data.revenue_estimate, "0y")
    next2_rev_est = _estimate_avg(data.revenue_estimate, "+1y")
    next_eps_est = _estimate_avg(data.earnings_estimate, "0y")
    forward_rev_g = _growth_rate(next_rev_est, rev)
    forward2_rev_g = _growth_rate(next2_rev_est, next_rev_est)
    forward_eps_g = _growth_rate(next_eps_est, eps) if eps is not None and eps > 0 else None

    signals: list[FinancialSignal] = []

    # Forward consensus diagnostics for the estimate columns already written into the template.
    if forward_rev_g is not None:
        if forward_rev_g < 0:
            signals.append(_signal(4, "FORWARD REVENUE CONTRACTION",
                f"Consensus revenue for FY{year+1} implies a {abs(forward_rev_g)*100:.1f}% decline from FY{year}, pointing to expected top-line contraction ahead.",
                f"컨센서스 기준 FY{year+1} 매출이 FY{year} 대비 {abs(forward_rev_g)*100:.1f}% 감소할 것으로 예상되어 향후 외형 위축 가능성이 반영되고 있다.", severity="WARNING"))
        elif rev_g is not None and (forward_rev_g - rev_g) <= -0.10:
            signals.append(_signal(7, "FORWARD GROWTH NORMALIZATION",
                f"Consensus FY{year+1} revenue growth of {forward_rev_g*100:.1f}% is {abs(forward_rev_g-rev_g)*100:.1f} percentage points below FY{year}, indicating a meaningful expected normalization in growth.",
                f"컨센서스 FY{year+1} 매출 성장률이 {forward_rev_g*100:.1f}%로 FY{year} 대비 {abs(forward_rev_g-rev_g)*100:.1f}%p 낮아져 향후 성장률 정상화가 예상된다.", severity="WEAK"))
        elif rev_g is not None and (forward_rev_g - rev_g) >= 0.05:
            signals.append(_signal(12, "FORWARD GROWTH ACCELERATION",
                f"Consensus FY{year+1} revenue growth of {forward_rev_g*100:.1f}% is above FY{year} growth, indicating expected re-acceleration in the top line.",
                f"컨센서스 FY{year+1} 매출 성장률이 {forward_rev_g*100:.1f}%로 FY{year}보다 높아 향후 외형 성장 재가속이 예상된다.", severity="POSITIVE"))

    if forward_rev_g is not None and forward2_rev_g is not None and forward2_rev_g < forward_rev_g - 0.08:
        signals.append(_signal(9, "TWO-YEAR FORWARD DECELERATION",
            f"Consensus revenue growth slows from {forward_rev_g*100:.1f}% in FY{year+1} to {forward2_rev_g*100:.1f}% in FY{year+2}, suggesting that growth expectations continue to normalize beyond the next fiscal year.",
            f"컨센서스 매출 성장률이 FY{year+1} {forward_rev_g*100:.1f}%에서 FY{year+2} {forward2_rev_g*100:.1f}%로 둔화될 것으로 예상되어 성장률 정상화가 다음 회계연도 이후에도 이어질 가능성이 있다.", severity="WEAK"))

    if forward_eps_g is not None and forward_rev_g is not None:
        if forward_eps_g >= forward_rev_g + 0.10:
            signals.append(_signal(13, "FORWARD EARNINGS LEVERAGE",
                f"Consensus EPS growth of {forward_eps_g*100:.1f}% materially exceeds expected revenue growth of {forward_rev_g*100:.1f}%, implying anticipated margin expansion and/or share-count accretion.",
                f"컨센서스 EPS 성장률 {forward_eps_g*100:.1f}%가 예상 매출 성장률 {forward_rev_g*100:.1f}%를 크게 상회해 향후 margin 개선 또는 주식수 감소 효과가 기대되고 있다.", severity="POSITIVE"))
        elif forward_eps_g <= forward_rev_g - 0.10:
            signals.append(_signal(8, "FORWARD EARNINGS LAG",
                f"Consensus EPS growth of {forward_eps_g*100:.1f}% trails expected revenue growth by {(forward_rev_g-forward_eps_g)*100:.1f} percentage points, suggesting that top-line growth may not fully translate into per-share earnings.",
                f"컨센서스 EPS 성장률이 예상 매출 성장률을 {(forward_rev_g-forward_eps_g)*100:.1f}%p 하회해 향후 외형 성장이 주당이익 증가로 충분히 연결되지 않을 가능성이 있다.", severity="WARNING"))

    # Multi-year and cross-statement diagnostics. These are intentionally high-priority
    # because they identify persistent deterioration that one-year point estimates can miss.
    if rev_g is not None and prev_rev_g is not None and prev2_rev_g is not None:
        growth_series = (prev2_rev_g, prev_rev_g, rev_g)
        if all(g >= 0.20 for g in growth_series):
            signals.append(_signal(6, "SUSTAINED HIGH GROWTH",
                f"Revenue growth has remained above 20% for three consecutive annual periods through FY{year}, indicating persistent rather than one-off high growth.",
                f"FY{year}까지 3개 연속 연간 구간에서 매출 성장률이 20%를 상회해 일회성이 아닌 구조적인 고성장세가 이어지고 있다.",
                severity="POSITIVE"))
        if all(g < 0 for g in growth_series):
            signals.append(_signal(1, "STRUCTURAL CONTRACTION",
                f"Revenue has contracted for three consecutive annual periods through FY{year}, suggesting that the top-line weakness is persistent rather than temporary.",
                f"FY{year}까지 3개 연속 연간 구간에서 매출이 감소해 외형 위축이 일시적이라기보다 구조적으로 이어지고 있을 가능성이 높다.",
                severity="CRITICAL"))
        elif prev2_rev_g > prev_rev_g > rev_g > 0 and (prev2_rev_g - rev_g) >= 0.10:
            signals.append(_signal(4, "MULTI-YEAR GROWTH DECELERATION",
                f"Revenue growth has slowed in each of the last three annual periods, falling {((prev2_rev_g-rev_g)*100):.1f} percentage points from the earlier rate to FY{year}.",
                f"최근 3개 연간 구간에서 매출 성장률이 연속 둔화되며 초기 수준 대비 {((prev2_rev_g-rev_g)*100):.1f}%p 낮아져 성장 모멘텀 약화가 구조적으로 이어지고 있다.",
                severity="WARNING"))

    if margin is not None and prev_margin is not None:
        one_year_margin_delta = margin - prev_margin
        if one_year_margin_delta <= -0.05:
            signals.append(_signal(1, "PROFITABILITY COLLAPSE",
                f"EBITDA margin compressed {abs(one_year_margin_delta)*10000:.0f} bps in FY{year}, representing a severe one-year deterioration in operating profitability.",
                f"FY{year} EBITDA margin이 전년 대비 {abs(one_year_margin_delta)*10000:.0f}bp 축소되어 영업 수익성이 단기간에 크게 악화됐다.",
                severity="CRITICAL"))
    if margin is not None and prev_margin is not None and prev2_margin is not None:
        if prev2_margin > prev_margin > margin and (prev2_margin - margin) >= 0.03:
            signals.append(_signal(3, "MULTI-YEAR MARGIN EROSION",
                f"EBITDA margin has declined for three consecutive fiscal-year observations and is {((prev2_margin-margin)*10000):.0f} bps below FY{prev2}, indicating persistent profitability erosion.",
                f"EBITDA margin이 3개 연도 관측치에서 연속 하락해 FY{prev2} 대비 {((prev2_margin-margin)*10000):.0f}bp 낮아졌으며 수익성 악화가 지속되고 있다.",
                severity="SEVERE"))

    if rev_g is not None and ebitda_g is not None and rev_g < -0.10 and ebitda_g < -0.10:
        signals.append(_signal(2, "FUNDAMENTAL DETERIORATION",
            f"Revenue and EBITDA both declined by double digits in FY{year}, indicating broad deterioration across both business scale and operating earnings.",
            f"FY{year} 매출과 EBITDA가 모두 두 자릿수 감소해 외형과 영업이익 양쪽에서 전반적인 펀더멘털 악화가 나타나고 있다.",
            severity="SEVERE"))

    if fcf is not None and prev_fcf is not None and fcf < 0 and prev_fcf < 0:
        signals.append(_signal(1, "PERSISTENT CASH BURN",
            f"Free cash flow has remained negative for at least two consecutive fiscal years through FY{year}, indicating persistent post-investment cash consumption.",
            f"FY{year}까지 최소 2개 회계연도 연속 free cash flow가 마이너스를 기록해 투자지출 이후 현금 소진이 지속되고 있다.",
            severity="CRITICAL"))

    if rev_g is not None and rev_g >= 0.10 and fcf is not None and fcf < 0:
        signals.append(_signal(4, "GROWTH / CASH FLOW DIVERGENCE",
            f"Revenue grew {rev_g*100:.1f}% in FY{year}, but free cash flow remained negative, indicating that reported growth is not yet translating into internally generated cash.",
            f"FY{year} 매출은 {rev_g*100:.1f}% 성장했지만 free cash flow가 마이너스를 기록해 외형 성장이 아직 내부 현금창출로 연결되지 못하고 있다.",
            severity="WARNING"))

    if net_income is not None and net_income > 0 and fcf is not None and fcf < 0:
        signals.append(_signal(3, "EARNINGS / CASH FLOW DIVERGENCE",
            f"The company reported positive net income in FY{year} while free cash flow was negative, creating a material divergence between accounting earnings and cash generation.",
            f"FY{year} 순이익은 흑자지만 free cash flow는 마이너스로 회계상 이익과 실제 현금창출 간 괴리가 크게 나타나고 있다.",
            severity="SEVERE"))

    if rev_g is not None and rev_g >= 0.20 and roic is not None and roic < 0.08:
        signals.append(_signal(5, "GROWTH WITH LOW CAPITAL RETURNS",
            f"Revenue is growing above 20%, but ROIC is only {roic*100:.1f}%, indicating that rapid expansion is being achieved with relatively weak returns on invested capital.",
            f"매출은 20% 이상 빠르게 성장하고 있지만 ROIC가 {roic*100:.1f}%에 그쳐 높은 성장 속도 대비 투자자본 수익성이 낮다.",
            severity="WARNING"))

    # Composite growth-quality branches receive the highest priority.
    if rev_g is not None and margin is not None and prev_margin is not None:
        margin_delta = margin - prev_margin
        if rev_g >= 0.20 and margin_delta >= 0.01:
            signals.append(_signal(5, "HIGH-QUALITY GROWTH",
                f"Revenue grew {rev_g*100:.1f}% in FY{year} while EBITDA margin expanded {margin_delta*10000:.0f} bps, indicating that rapid top-line growth is translating into stronger operating profitability.",
                f"FY{year} 매출이 {rev_g*100:.1f}% 성장하는 동시에 EBITDA margin이 {margin_delta*10000:.0f}bp 확대되어, 빠른 외형 성장이 수익성 개선으로 이어지고 있다."))
        elif rev_g >= 0.20 and margin_delta <= -0.03:
            signals.append(_signal(2, "GROWTH QUALITY WARNING",
                f"Revenue grew {rev_g*100:.1f}% in FY{year}, but EBITDA margin compressed {abs(margin_delta)*10000:.0f} bps, suggesting that the current growth rate is being achieved with materially weaker profitability.",
                f"FY{year} 매출은 {rev_g*100:.1f}% 성장했지만 EBITDA margin이 {abs(margin_delta)*10000:.0f}bp 축소되어, 높은 외형 성장에도 성장의 수익성이 크게 약화되고 있다."))

    if rev_g is not None:
        if rev_g >= 0.30:
            signals.append(_signal(20, "HYPER GROWTH",
                f"Revenue increased {rev_g*100:.1f}% in FY{year}, placing the company in a very high-growth profile with rapid top-line expansion.",
                f"FY{year} 매출 성장률이 {rev_g*100:.1f}%로 30%를 상회하며 매우 빠른 외형 성장을 보이고 있다."))
        elif rev_g >= 0.20:
            signals.append(_signal(20, "HIGH GROWTH",
                f"Revenue increased {rev_g*100:.1f}% in FY{year}, reflecting a strong high-growth top-line profile.",
                f"FY{year} 매출이 {rev_g*100:.1f}% 증가하며 높은 수준의 외형 성장세를 유지하고 있다."))
        elif rev_g >= 0.10:
            signals.append(_signal(20, "STRONG GROWTH",
                f"Revenue increased {rev_g*100:.1f}% in FY{year}, representing healthy double-digit top-line growth.",
                f"FY{year} 매출이 {rev_g*100:.1f}% 증가하며 견조한 두 자릿수 성장을 기록하고 있다."))
        elif rev_g >= 0.05:
            signals.append(_signal(20, "MODERATE GROWTH",
                f"Revenue increased {rev_g*100:.1f}% in FY{year}, indicating moderate but positive top-line expansion.",
                f"FY{year} 매출이 {rev_g*100:.1f}% 증가하며 완만한 외형 성장을 보이고 있다."))
        elif rev_g >= 0:
            signals.append(_signal(20, "LOW GROWTH",
                f"Revenue increased only {rev_g*100:.1f}% in FY{year}, suggesting a relatively mature or slow-growth top-line profile.",
                f"FY{year} 매출 성장률이 {rev_g*100:.1f}%에 그쳐 외형 성장 속도는 낮은 수준이다."))
        elif rev_g > -0.10:
            signals.append(_signal(10, "MILD CONTRACTION",
                f"Revenue declined {abs(rev_g)*100:.1f}% in FY{year}, indicating a moderate contraction in the top line.",
                f"FY{year} 매출이 {abs(rev_g)*100:.1f}% 감소하며 단기적인 외형 위축이 나타나고 있다."))
        else:
            signals.append(_signal(5, "SIGNIFICANT CONTRACTION",
                f"Revenue declined {abs(rev_g)*100:.1f}% in FY{year}, signaling a material contraction in business activity.",
                f"FY{year} 매출이 {abs(rev_g)*100:.1f}% 감소하며 뚜렷한 사업 위축이 나타나고 있다."))

    if rev_g is not None and prev_rev_g is not None:
        delta = rev_g - prev_rev_g
        if delta >= 0.05:
            signals.append(_signal(15, "GROWTH ACCELERATION",
                f"Revenue growth accelerated by {delta*100:.1f} percentage points versus FY{prev}, strengthening the company's top-line momentum.",
                f"매출 성장률이 FY{prev} 대비 {delta*100:.1f}%p 확대되며 성장 모멘텀이 강해지고 있다."))
        elif delta <= -0.05 and rev_g > 0:
            signals.append(_signal(8, "GROWTH DECELERATION",
                f"Revenue remains positive, but growth slowed by {abs(delta)*100:.1f} percentage points versus FY{prev}, indicating a meaningful loss of top-line momentum.",
                f"매출은 성장하고 있으나 성장률이 FY{prev} 대비 {abs(delta)*100:.1f}%p 둔화되며 성장 모멘텀이 약화되고 있다."))

    if ebitda_g is not None and rev_g is not None:
        spread = ebitda_g - rev_g
        if spread >= 0.10:
            signals.append(_signal(7, "POSITIVE OPERATING LEVERAGE",
                f"EBITDA grew {ebitda_g*100:.1f}%, outpacing revenue growth by {spread*100:.1f} percentage points and demonstrating positive operating leverage.",
                f"EBITDA가 {ebitda_g*100:.1f}% 증가해 매출 성장률을 {spread*100:.1f}%p 상회하며 뚜렷한 operating leverage가 나타나고 있다."))
        elif spread <= -0.10:
            signals.append(_signal(6, "NEGATIVE OPERATING LEVERAGE",
                f"EBITDA growth lagged revenue by {abs(spread)*100:.1f} percentage points, indicating that cost growth is absorbing a meaningful portion of the top-line expansion.",
                f"EBITDA 성장률이 매출 성장률을 {abs(spread)*100:.1f}%p 하회해 비용 증가가 외형 성장의 상당 부분을 상쇄하고 있다."))

    if margin is not None:
        if margin >= 0.30:
            text_en, text_ko = "EBITDA margin exceeds 30%, indicating exceptional operating profitability.", "EBITDA margin이 30%를 상회하며 매우 높은 영업 수익성을 보이고 있다."
        elif margin >= 0.20:
            text_en, text_ko = "EBITDA margin is above 20%, indicating strong operating profitability.", "EBITDA margin이 20%를 상회하며 강한 영업 수익성을 유지하고 있다."
        elif margin >= 0.10:
            text_en, text_ko = "EBITDA margin remains in a healthy double-digit range.", "EBITDA margin이 두 자릿수의 양호한 수준을 유지하고 있다."
        elif margin >= 0.05:
            text_en, text_ko = "EBITDA margin is relatively low, leaving meaningful room for cost and operating-efficiency improvement.", "EBITDA margin이 비교적 낮아 비용 구조 및 운영 효율성 개선 여지가 존재한다."
        elif margin >= 0:
            text_en, text_ko = "EBITDA margin is near breakeven, indicating limited operating-profitability headroom.", "EBITDA margin이 손익분기점에 가까워 영업 수익성 여력이 제한적이다."
        else:
            text_en, text_ko = "EBITDA is negative, indicating that the current operating model is not yet generating positive EBITDA.", "EBITDA가 마이너스로 현재 영업 구조에서 양의 EBITDA를 창출하지 못하고 있다."
        if margin >= 0.20:
            margin_severity = "POSITIVE"
        elif margin >= 0.10:
            margin_severity = "NEUTRAL"
        elif margin >= 0.05:
            margin_severity = "WEAK"
        elif margin >= 0:
            margin_severity = "WARNING"
        else:
            margin_severity = "SEVERE"
        signals.append(_signal(22 if margin >= 0 else 3, "EBITDA MARGIN", text_en, text_ko, severity=margin_severity))

    if op_income is not None and prev_op is not None:
        if prev_op <= 0 < op_income:
            signals.append(_signal(4, "OPERATING TURNAROUND",
                f"Operating income turned positive in FY{year}, marking a meaningful transition from operating losses to profitability.",
                f"FY{year} 영업이익이 흑자로 전환되며 본업 기준의 의미 있는 턴어라운드가 나타났다."))
        elif op_income < 0:
            signals.append(_signal(4, "OPERATING LOSS",
                f"The company remains operating-loss making in FY{year}, so the core business has not yet reached operating profitability.",
                f"FY{year} 영업손실을 기록하고 있어 본업의 수익성 확보에는 아직 도달하지 못했다."))
        elif op_g is not None and op_g < 0:
            signals.append(_signal(12, "OPERATING INCOME PRESSURE",
                f"Operating income declined {abs(op_g)*100:.1f}% despite remaining profitable, indicating weaker operating earnings versus FY{prev}.",
                f"영업흑자는 유지했지만 영업이익이 FY{prev} 대비 {abs(op_g)*100:.1f}% 감소해 본업의 이익 창출력이 약화됐다."))

    if net_income is not None and prev_ni is not None:
        if prev_ni <= 0 < net_income:
            signals.append(_signal(4, "NET INCOME TURNAROUND",
                f"Net income turned positive in FY{year}, indicating a bottom-line return to profitability.",
                f"FY{year} 순이익이 흑자로 전환되며 bottom-line 기준의 턴어라운드가 나타났다."))
        elif net_income < 0:
            signals.append(_signal(4, "NET LOSS",
                f"The company reported a net loss in FY{year}, indicating that operating and non-operating results have not yet translated into positive bottom-line earnings.",
                f"FY{year} 순손실을 기록해 최종적인 이익 창출에는 아직 도달하지 못했다."))
        elif ni_g is not None and op_g is not None:
            divergence = ni_g - op_g
            if divergence <= -0.15:
                signals.append(_signal(11, "BOTTOM-LINE DIVERGENCE",
                    f"Net income growth trailed operating-income growth by {abs(divergence)*100:.1f} percentage points, warranting a review of interest, tax and non-operating items.",
                    f"순이익 성장률이 영업이익 성장률을 {abs(divergence)*100:.1f}%p 하회해 이자, 세금 및 비영업 항목의 영향을 점검할 필요가 있다."))

    if eps_g is not None and ni_g is not None:
        diff = eps_g - ni_g
        if diff >= 0.05:
            signals.append(_signal(13, "EPS ACCRETION",
                f"EPS growth exceeded net-income growth by {diff*100:.1f} percentage points, consistent with share-count reduction or buyback accretion.",
                f"EPS 성장률이 순이익 성장률을 {diff*100:.1f}%p 상회해 주식수 감소 또는 자사주 매입에 따른 주당가치 증대 효과가 나타나고 있다."))
        elif diff <= -0.05 or (shares_g is not None and shares_g >= 0.05):
            signals.append(_signal(9, "DILUTION DRAG",
                f"EPS is growing more slowly than net income, indicating that share-count expansion is diluting the conversion of earnings growth into per-share value.",
                f"EPS 성장률이 순이익 성장률을 하회해 주식수 증가가 이익 성장을 주당가치 증가로 전환하는 데 부담으로 작용하고 있다."))

    if fcf is not None:
        if fcf < 0:
            signals.append(_signal(3, "CASH BURN",
                f"Free cash flow was negative in FY{year}, indicating that the company is currently consuming cash after capital expenditures.",
                f"FY{year} free cash flow가 마이너스를 기록해 투자지출 이후 현금을 소진하고 있다.", severity="SEVERE"))
        elif fcf_margin is not None and fcf_margin >= 0.20:
            signals.append(_signal(14, "EXCEPTIONAL CASH GENERATION",
                f"Free-cash-flow margin was {fcf_margin*100:.1f}%, demonstrating exceptionally strong cash generation relative to revenue.",
                f"FCF margin이 {fcf_margin*100:.1f}%로 매출 대비 매우 우수한 현금창출력을 보이고 있다.", severity="POSITIVE"))
        elif fcf_margin is not None and fcf_margin >= 0.10:
            signals.append(_signal(18, "HEALTHY CASH GENERATION",
                f"Free-cash-flow margin was {fcf_margin*100:.1f}%, indicating healthy conversion of revenue into discretionary cash flow.",
                f"FCF margin이 {fcf_margin*100:.1f}%로 매출을 재량 현금흐름으로 전환하는 능력이 양호하다.", severity="POSITIVE"))
        elif fcf_margin is not None and fcf_margin >= 0.05:
            signals.append(_signal(20, "MODERATE CASH GENERATION",
                f"Free-cash-flow margin was {fcf_margin*100:.1f}%, indicating positive but moderate cash generation relative to revenue.",
                f"FCF margin이 {fcf_margin*100:.1f}%로 양의 현금흐름을 창출하고 있으나 매출 대비 현금창출력은 중간 수준이다.", severity="NEUTRAL"))
        elif fcf_margin is not None:
            signals.append(_signal(12, "WEAK CASH GENERATION",
                f"Free-cash-flow margin was only {fcf_margin*100:.1f}%, leaving a limited cash buffer relative to revenue.",
                f"FCF margin이 {fcf_margin*100:.1f}%에 그쳐 매출 대비 현금창출 여력이 제한적이다.", severity="WEAK"))

        if fcf_g is not None and fcf_g >= 0.15:
            signals.append(_signal(16, "STRONG FCF GROWTH",
                f"Free cash flow increased {fcf_g*100:.1f}% in FY{year}, strengthening internally generated funding capacity.",
                f"FY{year} free cash flow가 {fcf_g*100:.1f}% 증가하며 내부 현금창출 여력이 강화되고 있다."))

    if conversion is not None:
        if conversion >= 1.20:
            signals.append(_signal(12, "HIGH EARNINGS QUALITY",
                f"Free cash flow was {conversion*100:.0f}% of net income, with cash generation exceeding reported earnings and supporting high earnings quality.",
                f"Free cash flow가 순이익의 {conversion*100:.0f}% 수준으로 회계상 이익을 상회하는 현금을 창출해 earnings quality가 높다."))
        elif conversion < 0.50:
            signals.append(_signal(7, "CASH CONVERSION WARNING",
                f"Free cash flow was only {conversion*100:.0f}% of net income, indicating weak conversion of accounting earnings into cash.",
                f"Free cash flow가 순이익의 {conversion*100:.0f}%에 그쳐 회계상 이익의 현금 전환력이 낮다."))

    if roic is not None:
        if roic >= 0.25:
            signals.append(_signal(18, "EXCEPTIONAL ROIC", f"ROIC of {roic*100:.1f}% indicates exceptional efficiency in converting invested capital into operating returns.", f"ROIC가 {roic*100:.1f}%로 투자자본 대비 매우 높은 수익을 창출하고 있다."))
        elif roic >= 0.15:
            signals.append(_signal(20, "STRONG ROIC", f"ROIC of {roic*100:.1f}% reflects strong capital efficiency.", f"ROIC가 {roic*100:.1f}%로 우수한 자본 효율성을 보이고 있다."))
        elif roic < 0.08:
            signals.append(_signal(10, "LOW ROIC", f"ROIC of {roic*100:.1f}% is relatively low, suggesting limited returns on the capital committed to the business.", f"ROIC가 {roic*100:.1f}%로 낮아 투입 자본 대비 수익성이 제한적이다."))

    if nd_ebitda is not None:
        if nd_ebitda < 0:
            signals.append(_signal(24, "NET CASH", "The company holds more cash than debt, providing substantial balance-sheet flexibility.", "현금이 부채를 상회하는 순현금 상태로 높은 재무적 유연성을 보유하고 있다."))
        elif nd_ebitda > 4:
            signals.append(_signal(3, "HIGH LEVERAGE", f"Net debt stands at {nd_ebitda:.2f}x EBITDA, indicating a high leverage burden relative to operating earnings.", f"Net Debt/EBITDA가 {nd_ebitda:.2f}x로 영업이익 대비 높은 레버리지 부담을 보이고 있다."))

    return _select_financial_signals(signals, "income")


def _balance_financial_signals(data: MarketData) -> list[FinancialSignal]:
    year = _latest_template_actual_year(data)
    if year is None:
        return []
    prev = year - 1
    prev2 = year - 2
    rev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], year)
    prev_rev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], prev)
    prev2_rev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], prev2)
    rev_g = _growth_rate(rev, prev_rev)
    prev_rev_g = _growth_rate(prev_rev, prev2_rev)
    ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], year)
    prev_ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], prev)
    ebitda_g = _growth_rate(ebitda, prev_ebitda)

    def bal(label: str, y: int = year) -> float | None:
        return _annual_value(data.annual_balance, BALANCE_ALIASES[label], y)

    cash, prev_cash, prev2_cash = bal("Cash and equivalents"), bal("Cash and equivalents", prev), bal("Cash and equivalents", prev2)
    debt, prev_debt, prev2_debt = bal("Total debt (incl. leases)"), bal("Total debt (incl. leases)", prev), bal("Total debt (incl. leases)", prev2)
    ar, prev_ar, prev2_ar = bal("Accounts receivable"), bal("Accounts receivable", prev), bal("Accounts receivable", prev2)
    inv, prev_inv, prev2_inv = bal("Inventory"), bal("Inventory", prev), bal("Inventory", prev2)
    ca, cl = bal("Total current assets"), bal("Total current liabilities")
    prev_ca, prev_cl = bal("Total current assets", prev), bal("Total current liabilities", prev)
    ppe, prev_ppe = bal("Property, plant and equipment"), bal("Property, plant and equipment", prev)
    assets, prev_assets = bal("Total assets"), bal("Total assets", prev)
    ap, prev_ap = bal("Accounts payable"), bal("Accounts payable", prev)
    liabilities = bal("Total liabilities")
    equity, prev_equity = bal("Total equity"), bal("Total equity", prev)
    goodwill = _annual_value(data.annual_balance, ("Goodwill And Other Intangible Assets",), year)
    if goodwill is None:
        goodwill = _combined_annual_value(data.annual_balance, (("Goodwill",), ("Other Intangible Assets",)), year)
    net_debt = debt - cash if debt is not None and cash is not None else None
    prev_net_debt = prev_debt - prev_cash if prev_debt is not None and prev_cash is not None else None
    nd_ebitda = _safe_div(net_debt, ebitda) if ebitda is not None and ebitda > 0 else None
    prev_nd_ebitda = _safe_div(prev_net_debt, prev_ebitda) if prev_ebitda is not None and prev_ebitda > 0 else None
    current_ratio = _safe_div(ca, cl)
    prev_current_ratio = _safe_div(prev_ca, prev_cl)
    roe = _roe(data, year)
    roic = _roic(data, year)

    cash_g = _growth_rate(cash, prev_cash)
    prev_cash_g = _growth_rate(prev_cash, prev2_cash)
    debt_g = _growth_rate(debt, prev_debt)
    prev_debt_g = _growth_rate(prev_debt, prev2_debt)
    ar_g = _growth_rate(ar, prev_ar)
    prev_ar_g = _growth_rate(prev_ar, prev2_ar)
    inv_g = _growth_rate(inv, prev_inv)
    prev_inv_g = _growth_rate(prev_inv, prev2_inv)
    ppe_g = _growth_rate(ppe, prev_ppe)
    asset_g = _growth_rate(assets, prev_assets)
    ap_g = _growth_rate(ap, prev_ap)
    equity_g = _growth_rate(equity, prev_equity) if prev_equity is not None and prev_equity > 0 else None

    signals: list[FinancialSignal] = []
    is_financial = _is_financial_company(data)
    if is_financial:
        signals.append(_signal(1, "SECTOR-SPECIFIC INTERPRETATION",
            "Standard industrial-company liquidity and leverage thresholds are suppressed because financial companies require sector-specific balance-sheet metrics such as regulatory capital, funding mix and asset quality.",
            "금융주는 일반 기업용 유동성 및 레버리지 기준 대신 규제자본, 조달구조, 자산건전성 등 섹터 특화 지표가 중요하므로 표준 산업기업 기준의 일부 자동 판정을 적용하지 않았다."))

    # Cross-metric stress tests.
    if not is_financial and cash_g is not None and debt_g is not None and current_ratio is not None:
        if cash_g <= -0.20 and debt_g >= 0.20 and current_ratio < 1.20:
            signals.append(_signal(1, "CRITICAL BALANCE SHEET STRESS",
                f"Cash fell {abs(cash_g)*100:.1f}% while debt rose {debt_g*100:.1f}% and the current ratio is only {current_ratio:.2f}x, indicating simultaneous pressure on liquidity and leverage.",
                f"현금이 {abs(cash_g)*100:.1f}% 감소하는 동안 부채는 {debt_g*100:.1f}% 증가했고 Current ratio가 {current_ratio:.2f}x에 그쳐 유동성과 레버리지 양쪽에서 동시에 강한 압박이 나타나고 있다.",
                severity="CRITICAL"))

    if not is_financial and nd_ebitda is not None and nd_ebitda >= 5 and ebitda_g is not None and ebitda_g < 0:
        signals.append(_signal(1, "VERY HIGH LEVERAGE WITH EARNINGS DECLINE",
            f"Net debt / EBITDA is {nd_ebitda:.2f}x while EBITDA declined {abs(ebitda_g)*100:.1f}%, creating a particularly weak combination of elevated debt and shrinking earnings capacity.",
            f"Net Debt/EBITDA가 {nd_ebitda:.2f}x인 가운데 EBITDA도 {abs(ebitda_g)*100:.1f}% 감소해 높은 부채 부담과 이익창출력 약화가 동시에 나타나고 있다.",
            severity="CRITICAL"))

    if ar_g is not None and inv_g is not None and rev_g is not None and cash_g is not None:
        if (ar_g - rev_g) >= 0.15 and (inv_g - rev_g) >= 0.15 and cash_g < 0:
            signals.append(_signal(2, "WORKING-CAPITAL STRESS",
                "Accounts receivable and inventory are both growing materially faster than revenue while cash is declining, indicating a broad working-capital drain rather than a single-line anomaly.",
                "매출채권과 재고가 모두 매출보다 크게 빠르게 증가하는 동시에 현금이 감소해 단일 항목 문제가 아니라 전반적인 운전자본 부담이 확대되고 있다.",
                severity="SEVERE"))

    if current_ratio is not None and prev_current_ratio is not None and (current_ratio - prev_current_ratio) <= -0.30:
        signals.append(_signal(5, "LIQUIDITY DETERIORATION",
            f"The current ratio fell from {prev_current_ratio:.2f}x to {current_ratio:.2f}x, indicating a meaningful year-over-year weakening in short-term liquidity.",
            f"Current ratio가 전년 {prev_current_ratio:.2f}x에서 {current_ratio:.2f}x로 하락해 단기 유동성이 의미 있게 약화됐다.",
            severity="WARNING"))

    if cash_g is not None and prev_cash_g is not None and cash_g < 0 and prev_cash_g < 0 and cash is not None and prev2_cash not in (None, 0):
        total_cash_drop = cash / prev2_cash - 1.0
        if total_cash_drop <= -0.25:
            signals.append(_signal(4, "MULTI-YEAR CASH DEPLETION",
                f"Cash has declined for two consecutive years and is {abs(total_cash_drop)*100:.1f}% below FY{prev2}, indicating persistent erosion of the liquidity buffer.",
                f"현금이 2년 연속 감소해 FY{prev2} 대비 {abs(total_cash_drop)*100:.1f}% 낮아졌으며 유동성 완충력이 지속적으로 약화되고 있다.",
                severity="WARNING"))

    if debt_g is not None and prev_debt_g is not None and debt_g > 0.10 and prev_debt_g > 0.10:
        signals.append(_signal(5, "MULTI-YEAR DEBT BUILD-UP",
            f"Total debt has increased by more than 10% in each of the last two annual periods, indicating a sustained increase in balance-sheet leverage.",
            f"총부채가 최근 2개 연간 구간에서 각각 10% 이상 증가해 balance sheet 레버리지가 지속적으로 확대되고 있다.",
            severity="WARNING"))

    if ar_g is not None and prev_ar_g is not None and rev_g is not None and prev_rev_g is not None:
        if (ar_g - rev_g) >= 0.10 and (prev_ar_g - prev_rev_g) >= 0.10:
            signals.append(_signal(4, "PERSISTENT RECEIVABLE BUILD-UP",
                "Accounts receivable has outgrown revenue by at least 10 percentage points in two consecutive annual periods, increasing the likelihood of a structural collection-quality issue.",
                "매출채권 증가율이 2개 연속 연간 구간에서 매출 성장률을 10%p 이상 상회해 구조적인 채권 회수 효율 저하 가능성이 높아지고 있다.",
                severity="WARNING"))

    if inv_g is not None and prev_inv_g is not None and rev_g is not None and prev_rev_g is not None:
        if (inv_g - rev_g) >= 0.10 and (prev_inv_g - prev_rev_g) >= 0.10:
            signals.append(_signal(4, "PERSISTENT INVENTORY BUILD-UP",
                "Inventory has outgrown revenue by at least 10 percentage points in two consecutive annual periods, raising the risk of persistent excess inventory or slower sell-through.",
                "재고 증가율이 2개 연속 연간 구간에서 매출 성장률을 10%p 이상 상회해 과잉재고 또는 판매회전 둔화가 구조적으로 이어질 가능성이 있다.",
                severity="WARNING"))

    if not is_financial and cash_g is not None and debt_g is not None:
        if cash_g > 0.05 and debt_g < -0.05:
            signals.append(_signal(4, "BALANCE SHEET STRENGTHENING",
                "Cash increased while total debt declined, indicating a clear improvement in financial flexibility and balance-sheet strength.",
                "현금은 증가하고 총부채는 감소해 재무적 유연성과 balance sheet 건전성이 동시에 개선되고 있다."))
        elif cash_g < -0.05 and debt_g > 0.05:
            signals.append(_signal(2, "BALANCE SHEET DETERIORATION",
                "Cash declined while total debt increased, indicating a deterioration in balance-sheet flexibility.",
                "현금 감소와 총부채 증가가 동시에 나타나 balance sheet의 재무적 유연성이 약화되고 있다."))

    cash_debt = _safe_div(cash, debt) if debt is not None and debt > 0 else None
    if not is_financial and cash_debt is not None:
        if cash_debt >= 1:
            signals.append(_signal(18, "FULL DEBT CASH COVERAGE", f"Cash equals {cash_debt*100:.0f}% of total debt, providing enough liquidity to cover the entire debt balance.", f"현금이 총부채의 {cash_debt*100:.0f}% 수준으로 전체 부채를 상쇄할 수 있는 매우 강한 유동성을 보유하고 있다."))
        elif cash_debt >= 0.70:
            signals.append(_signal(20, "STRONG LIQUIDITY", f"Cash covers {cash_debt*100:.0f}% of total debt, indicating a strong liquidity buffer.", f"현금이 총부채의 {cash_debt*100:.0f}%를 커버해 높은 유동성 완충력을 보유하고 있다."))
        elif cash_debt < 0.20:
            signals.append(_signal(10, "LOW CASH COVERAGE", f"Cash covers only {cash_debt*100:.0f}% of total debt, leaving the company relatively dependent on ongoing cash generation and external funding access.", f"현금이 총부채의 {cash_debt*100:.0f}%에 불과해 지속적인 현금창출과 외부 자금조달 접근성에 대한 의존도가 상대적으로 높다."))

    if ar_g is not None and rev_g is not None:
        spread = ar_g - rev_g
        if spread >= 0.15:
            signals.append(_signal(5, "RECEIVABLE BUILD-UP", f"Accounts receivable grew {spread*100:.1f} percentage points faster than revenue, raising a collection-quality and working-capital warning.", f"매출채권 증가율이 매출 성장률을 {spread*100:.1f}%p 상회해 채권 회수 속도와 매출의 질을 점검할 필요가 있다."))
        elif spread <= -0.05:
            signals.append(_signal(22, "RECEIVABLE EFFICIENCY", f"Accounts receivable grew {abs(spread)*100:.1f} percentage points slower than revenue, indicating improved receivables efficiency.", f"매출채권 증가율이 매출 성장률을 {abs(spread)*100:.1f}%p 하회해 매출채권 운용 효율성이 개선되고 있다."))

    if inv_g is not None and rev_g is not None:
        spread = inv_g - rev_g
        if spread >= 0.30:
            signals.append(_signal(3, "SEVERE INVENTORY BUILD-UP", f"Inventory growth exceeded revenue growth by {spread*100:.1f} percentage points, signaling a material inventory build-up and potential demand mismatch.", f"재고 증가율이 매출 성장률을 {spread*100:.1f}%p 상회해 과잉재고 또는 수요 둔화 가능성이 크게 높아지고 있다."))
        elif spread >= 0.15:
            signals.append(_signal(6, "INVENTORY BUILD-UP", f"Inventory grew {spread*100:.1f} percentage points faster than revenue, warranting monitoring for demand slowdown or overproduction.", f"재고 증가율이 매출 성장률을 {spread*100:.1f}%p 상회해 수요 둔화 또는 재고 누적 여부를 모니터링할 필요가 있다."))
        elif spread <= -0.05:
            signals.append(_signal(23, "INVENTORY EFFICIENCY", f"Inventory grew more slowly than revenue by {abs(spread)*100:.1f} percentage points, indicating improving inventory efficiency.", f"재고 증가율이 매출 성장률을 {abs(spread)*100:.1f}%p 하회해 재고 운용 효율성이 개선되고 있다."))

    if not is_financial and current_ratio is not None:
        if current_ratio >= 3:
            signals.append(_signal(24, "EXCESS LIQUIDITY", f"The current ratio is {current_ratio:.2f}x, indicating very strong short-term liquidity, though excess working capital may reduce capital efficiency.", f"Current ratio가 {current_ratio:.2f}x로 단기 유동성은 매우 우수하나 과도한 유동자산이 자본 효율성을 낮출 가능성도 있다."))
        elif current_ratio >= 1.5:
            signals.append(_signal(21, "HEALTHY LIQUIDITY", f"The current ratio of {current_ratio:.2f}x indicates a healthy short-term liquidity position.", f"Current ratio가 {current_ratio:.2f}x로 단기 지급능력이 건전한 수준이다."))
        elif current_ratio >= 1.2:
            signals.append(_signal(18, "ADEQUATE LIQUIDITY", f"The current ratio of {current_ratio:.2f}x is adequate but provides a more limited liquidity cushion.", f"Current ratio가 {current_ratio:.2f}x로 수용 가능한 수준이나 유동성 완충력은 크지 않다."))
        elif current_ratio >= 1:
            signals.append(_signal(9, "TIGHT LIQUIDITY", f"The current ratio of {current_ratio:.2f}x indicates a relatively tight short-term liquidity position.", f"Current ratio가 {current_ratio:.2f}x로 단기 지급 여력이 다소 타이트하다."))
        else:
            signals.append(_signal(2, "LIQUIDITY RISK", f"The current ratio is {current_ratio:.2f}x, with current liabilities exceeding current assets and creating a short-term liquidity warning.", f"Current ratio가 {current_ratio:.2f}x로 유동부채가 유동자산을 상회해 단기 유동성 위험을 모니터링할 필요가 있다."))

    if ppe_g is not None and rev_g is not None:
        spread = ppe_g - rev_g
        if spread >= 0.10:
            signals.append(_signal(14, "CAPACITY INVESTMENT", f"PP&E grew {spread*100:.1f} percentage points faster than revenue, suggesting significant capacity investment whose future revenue productivity should be monitored.", f"PP&E 증가율이 매출 성장률을 {spread*100:.1f}%p 상회해 선행 설비투자가 향후 매출 성장으로 연결되는지 확인할 필요가 있다."))
        elif spread <= -0.10:
            signals.append(_signal(20, "ASSET PRODUCTIVITY", f"Revenue growth outpaced PP&E growth by {abs(spread)*100:.1f} percentage points, indicating improving productivity of the existing asset base.", f"매출 성장률이 PP&E 증가율을 {abs(spread)*100:.1f}%p 상회해 기존 설비자산의 생산성이 개선되고 있다."))

    goodwill_assets = _safe_div(goodwill, assets)
    goodwill_equity = _safe_div(goodwill, equity) if equity is not None and equity > 0 else None
    if goodwill_assets is not None:
        if goodwill_assets > 0.60:
            signals.append(_signal(2, "VERY HIGH ACQUISITION EXPOSURE", f"Goodwill and intangibles represent {goodwill_assets*100:.1f}% of total assets, making reported book value highly dependent on acquired intangible value and future impairment assumptions.", f"Goodwill 및 무형자산이 총자산의 {goodwill_assets*100:.1f}%를 차지해 장부가치가 인수자산 가치와 향후 impairment 가정에 크게 의존하고 있다.", severity="SEVERE"))
        elif goodwill_assets > 0.40:
            signals.append(_signal(5, "HIGH ACQUISITION EXPOSURE", f"Goodwill and intangibles represent {goodwill_assets*100:.1f}% of total assets, making acquisition performance and impairment risk material to book value.", f"Goodwill 및 무형자산이 총자산의 {goodwill_assets*100:.1f}%를 차지해 과거 M&A 성과와 impairment risk가 장부가치에 중요한 변수로 작용한다."))
        elif goodwill_assets >= 0.25:
            signals.append(_signal(12, "ELEVATED INTANGIBLES", f"Goodwill and intangibles account for {goodwill_assets*100:.1f}% of assets, creating elevated sensitivity to future impairment charges.", f"Goodwill 및 무형자산이 총자산의 {goodwill_assets*100:.1f}%를 차지해 향후 impairment에 대한 민감도가 비교적 높다."))
    if goodwill_equity is not None and goodwill_equity > 1:
        signals.append(_signal(4, "IMPAIRMENT SENSITIVITY", f"Goodwill and intangibles exceed book equity at {goodwill_equity*100:.0f}% of equity, so a large impairment could materially affect reported net worth.", f"Goodwill 및 무형자산이 자기자본의 {goodwill_equity*100:.0f}%로 장부상 자기자본을 상회해 대규모 impairment 발생 시 equity에 미치는 영향이 클 수 있다."))

    if asset_g is not None and rev_g is not None:
        spread = asset_g - rev_g
        if spread >= 0.10:
            signals.append(_signal(15, "ASSET BUILD AHEAD OF GROWTH", f"Total assets grew {spread*100:.1f} percentage points faster than revenue, making monetization of the expanded asset base an important execution point.", f"총자산 증가율이 매출 성장률을 {spread*100:.1f}%p 상회해 확대된 자산이 향후 수익으로 전환되는지 확인할 필요가 있다."))
        elif spread <= -0.10:
            signals.append(_signal(22, "IMPROVING ASSET PRODUCTIVITY", f"Revenue grew {abs(spread)*100:.1f} percentage points faster than total assets, indicating improving asset productivity.", f"매출 성장률이 총자산 증가율을 {abs(spread)*100:.1f}%p 상회해 자산 활용 효율성이 개선되고 있다."))

    if ap_g is not None and rev_g is not None and ap_g - rev_g >= 0.15:
        if cash_g is not None and cash_g < 0:
            signals.append(_signal(6, "PAYABLES / LIQUIDITY WARNING", f"Accounts payable is growing materially faster than revenue while cash is declining, warranting a review of supplier-payment timing and liquidity pressure.", f"매입채무가 매출보다 빠르게 증가하는 동시에 현금이 감소해 공급업체 지급 지연 또는 유동성 압박 여부를 점검할 필요가 있다."))
        else:
            signals.append(_signal(17, "WORKING-CAPITAL FUNDING", f"Accounts payable is growing faster than revenue, which may reflect more efficient use of supplier financing but should be monitored for payment-term changes.", f"매입채무가 매출보다 빠르게 증가해 공급업체 결제조건을 활용한 운전자본 효율화 가능성이 있으나 지급조건 변화는 모니터링할 필요가 있다."))

    if not is_financial and debt_g is not None and ebitda_g is not None:
        spread = debt_g - ebitda_g
        if spread >= 0.10:
            signals.append(_signal(6, "LEVERAGE BUILD-UP", f"Debt growth exceeded EBITDA growth by {spread*100:.1f} percentage points, indicating increasing leverage relative to earnings capacity.", f"부채 증가율이 EBITDA 성장률을 {spread*100:.1f}%p 상회해 이익창출력 대비 재무 레버리지가 확대되고 있다."))
        elif spread <= -0.10:
            signals.append(_signal(22, "LEVERAGE IMPROVEMENT", f"EBITDA growth outpaced debt growth by {abs(spread)*100:.1f} percentage points, reducing leverage relative to operating earnings.", f"EBITDA 성장률이 부채 증가율을 {abs(spread)*100:.1f}%p 상회해 영업이익 대비 레버리지 부담이 감소하고 있다."))

    liabilities_assets = _safe_div(liabilities, assets)
    if not is_financial and liabilities_assets is not None:
        if liabilities_assets > 0.90:
            signals.append(_signal(3, "VERY HIGH LIABILITY LOAD", f"Liabilities represent {liabilities_assets*100:.1f}% of total assets, indicating a highly leveraged balance-sheet structure.", f"총부채가 총자산의 {liabilities_assets*100:.1f}%를 차지해 매우 높은 balance-sheet leverage를 보이고 있다."))
        elif liabilities_assets >= 0.75:
            signals.append(_signal(8, "HIGH LIABILITY LOAD", f"Liabilities represent {liabilities_assets*100:.1f}% of total assets, indicating elevated balance-sheet leverage.", f"총부채가 총자산의 {liabilities_assets*100:.1f}%를 차지해 부채 의존도가 높은 수준이다."))
        elif liabilities_assets < 0.40:
            signals.append(_signal(25, "CONSERVATIVE CAPITAL STRUCTURE", f"Liabilities represent only {liabilities_assets*100:.1f}% of total assets, indicating a conservative capital structure.", f"총부채가 총자산의 {liabilities_assets*100:.1f}%에 그쳐 보수적인 자본구조를 유지하고 있다."))

    if equity is not None:
        if equity < 0:
            signals.append(_signal(4, "NEGATIVE EQUITY", "Book equity is negative, so accumulated losses, repurchases and other capital-account movements should be reviewed before interpreting leverage or ROE mechanically.", "장부상 자기자본이 마이너스로 누적손실, 자사주 매입 등 자본계정 변동의 원인을 확인한 뒤 레버리지와 ROE를 해석할 필요가 있다."))
        elif equity_g is not None and equity_g >= 0.10:
            signals.append(_signal(23, "EQUITY BASE EXPANSION", f"Book equity increased {equity_g*100:.1f}% in FY{year}, strengthening the company's accounting capital base.", f"FY{year} 자기자본이 {equity_g*100:.1f}% 증가하며 장부상 자본 기반이 강화되고 있다."))
        elif equity_g is not None and equity_g <= -0.10:
            signals.append(_signal(8, "EQUITY BASE EROSION", f"Book equity declined {abs(equity_g)*100:.1f}% in FY{year}, reducing the accounting capital cushion.", f"FY{year} 자기자본이 {abs(equity_g)*100:.1f}% 감소해 장부상 자본 완충력이 약화됐다.", severity="WARNING"))

    if equity_g is not None and debt_g is not None and equity_g < -0.05 and debt_g > 0.05:
        signals.append(_signal(4, "CAPITAL STRUCTURE DETERIORATION",
            "Book equity is shrinking while debt is increasing, indicating a less favorable shift in the capital structure.",
            "자기자본이 감소하는 동시에 부채가 증가해 자본구조가 불리한 방향으로 이동하고 있다.",
            severity="WARNING"))

    if not is_financial and nd_ebitda is not None:
        if nd_ebitda < 0:
            signals.append(_signal(18, "NET CASH", "The company is in a net-cash position, providing substantial financial flexibility.", "현금이 부채를 상회하는 순현금 상태로 높은 재무적 유연성을 보유하고 있다."))
        elif nd_ebitda <= 1.5:
            signals.append(_signal(21, "LOW LEVERAGE", f"Net debt is only {nd_ebitda:.2f}x EBITDA, indicating a healthy leverage profile.", f"Net Debt/EBITDA가 {nd_ebitda:.2f}x로 낮아 건전한 레버리지 구조를 유지하고 있다."))
        elif nd_ebitda >= 5:
            signals.append(_signal(2, "VERY HIGH LEVERAGE", f"Net debt is {nd_ebitda:.2f}x EBITDA, indicating a very high debt burden relative to earnings capacity.", f"Net Debt/EBITDA가 {nd_ebitda:.2f}x로 이익창출력 대비 매우 높은 부채 부담을 보이고 있다."))
        elif nd_ebitda >= 4:
            signals.append(_signal(4, "HIGH LEVERAGE", f"Net debt is {nd_ebitda:.2f}x EBITDA, representing a high leverage burden.", f"Net Debt/EBITDA가 {nd_ebitda:.2f}x로 높은 재무 레버리지를 보이고 있다."))
        elif nd_ebitda >= 3:
            signals.append(_signal(9, "ELEVATED LEVERAGE", f"Net debt is {nd_ebitda:.2f}x EBITDA, an elevated level that warrants monitoring.", f"Net Debt/EBITDA가 {nd_ebitda:.2f}x로 다소 높은 수준이어서 레버리지 추이를 모니터링할 필요가 있다."))
        if prev_nd_ebitda is not None:
            delta = nd_ebitda - prev_nd_ebitda
            if delta <= -0.50:
                signals.append(_signal(10, "DELEVERAGING", f"Net debt / EBITDA improved by {abs(delta):.2f}x year over year, indicating meaningful deleveraging.", f"Net Debt/EBITDA가 전년 대비 {abs(delta):.2f}x 하락하며 의미 있는 deleveraging이 진행되고 있다."))
            elif delta >= 0.50:
                signals.append(_signal(6, "LEVERAGE DETERIORATION", f"Net debt / EBITDA increased by {delta:.2f}x year over year, indicating a material build-up in leverage.", f"Net Debt/EBITDA가 전년 대비 {delta:.2f}x 상승하며 재무 레버리지가 빠르게 확대되고 있다."))

    if roe is not None and equity is not None and equity > 0:
        if roe >= 0.25:
            roe_en, roe_ko, roe_sev = f"ROE is {roe*100:.1f}%, indicating exceptionally high returns on book equity.", f"ROE가 {roe*100:.1f}%로 장부상 자기자본 대비 매우 높은 수익성을 보이고 있다.", "POSITIVE"
        elif roe >= 0.15:
            roe_en, roe_ko, roe_sev = f"ROE is {roe*100:.1f}%, indicating strong shareholder capital efficiency.", f"ROE가 {roe*100:.1f}%로 주주자본 활용 효율성이 우수하다.", "POSITIVE"
        elif roe >= 0.10:
            roe_en, roe_ko, roe_sev = f"ROE is {roe*100:.1f}%, representing a healthy but not exceptional return on equity.", f"ROE가 {roe*100:.1f}%로 양호하지만 매우 높은 수준은 아닌 자기자본 수익성을 보이고 있다.", "NEUTRAL"
        elif roe >= 0.05:
            roe_en, roe_ko, roe_sev = f"ROE is {roe*100:.1f}%, indicating relatively weak returns on shareholder equity.", f"ROE가 {roe*100:.1f}%로 주주자본 대비 수익성이 비교적 낮다.", "WEAK"
        else:
            roe_en, roe_ko, roe_sev = f"ROE is only {roe*100:.1f}%, indicating poor returns on shareholder equity.", f"ROE가 {roe*100:.1f}%에 그쳐 주주자본 대비 수익성이 낮다.", "WARNING"
        signals.append(_signal(19 if roe_sev in {"POSITIVE", "NEUTRAL"} else 10, "ROE PROFILE", roe_en, roe_ko, severity=roe_sev))

    if roic is not None:
        if roic >= 0.20:
            roic_en, roic_ko, roic_sev = f"ROIC is {roic*100:.1f}%, indicating excellent returns on invested capital.", f"ROIC가 {roic*100:.1f}%로 투자자본 대비 매우 우수한 수익성을 보이고 있다.", "POSITIVE"
        elif roic >= 0.15:
            roic_en, roic_ko, roic_sev = f"ROIC is {roic*100:.1f}%, indicating strong capital efficiency.", f"ROIC가 {roic*100:.1f}%로 우수한 자본 효율성을 보이고 있다.", "POSITIVE"
        elif roic >= 0.10:
            roic_en, roic_ko, roic_sev = f"ROIC is {roic*100:.1f}%, indicating reasonable invested-capital efficiency.", f"ROIC가 {roic*100:.1f}%로 투자자본 효율성은 대체로 양호한 수준이다.", "NEUTRAL"
        elif roic >= 0.08:
            roic_en, roic_ko, roic_sev = f"ROIC is {roic*100:.1f}%, leaving limited excess return on invested capital before considering the cost of capital.", f"ROIC가 {roic*100:.1f}%로 자본비용을 고려하기 전 기준에서도 초과수익 여력이 크지 않다.", "WEAK"
        else:
            roic_en, roic_ko, roic_sev = f"ROIC is only {roic*100:.1f}%, indicating weak returns on the capital committed to the business.", f"ROIC가 {roic*100:.1f}%에 그쳐 사업에 투입된 자본 대비 수익성이 낮다.", "WARNING"
        signals.append(_signal(18 if roic_sev in {"POSITIVE", "NEUTRAL"} else 9, "ROIC PROFILE", roic_en, roic_ko, severity=roic_sev))

    if not is_financial and roe is not None and roic is not None and nd_ebitda is not None and roe >= 0.25 and roic < 0.10 and nd_ebitda > 3:
        signals.append(_signal(3, "LEVERAGE-DRIVEN ROE", f"ROE is {roe*100:.1f}%, but ROIC is only {roic*100:.1f}% with net debt / EBITDA above 3x, suggesting that high shareholder returns may be amplified by financial leverage.", f"ROE는 {roe*100:.1f}%로 높지만 ROIC가 {roic*100:.1f}%에 그치고 Net Debt/EBITDA가 3x를 상회해 높은 ROE가 재무 레버리지에 의해 확대되고 있을 가능성이 있다."))
    elif not is_financial and roe is not None and roic is not None and roe >= 0.20 and roic >= 0.15:
        signals.append(_signal(16, "HIGH-QUALITY CAPITAL RETURNS", f"ROE of {roe*100:.1f}% and ROIC of {roic*100:.1f}% jointly indicate strong returns on both shareholder and invested capital.", f"ROE {roe*100:.1f}%와 ROIC {roic*100:.1f}%를 동시에 기록해 주주자본과 투자자본 모두에서 높은 수익성을 보이고 있다."))

    return _select_financial_signals(signals, "balance")


def _all_quarter_dates(data: MarketData) -> list[pd.Timestamp]:
    cols: set[pd.Timestamp] = set()
    for df in (data.quarterly_income, data.quarterly_balance):
        if df is None or df.empty:
            continue
        for col in df.columns:
            try:
                cols.add(pd.Timestamp(col))
            except Exception:
                continue
    return sorted(cols)


def _prior_year_quarter(dates: list[pd.Timestamp], latest: pd.Timestamp) -> pd.Timestamp | None:
    latest_q = (latest.month - 1) // 3 + 1
    matches = [d for d in dates if d.year == latest.year - 1 and (d.month - 1) // 3 + 1 == latest_q]
    if matches:
        return max(matches)
    older = [d for d in dates if d < latest]
    return older[-4] if len(older) >= 4 else (older[0] if older else None)


def _quarter_valuation_snapshot(data: MarketData, quarter: pd.Timestamp) -> dict[str, float | None]:
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
    ltm_eps = _safe_div(ltm_net_income, shares) if ltm_net_income is not None and shares is not None and shares > 0 else None
    ev_ebitda = _safe_div(enterprise_value, ltm_ebitda) if ltm_ebitda is not None and ltm_ebitda > 0 else None
    pe = _safe_div(market_cap, ltm_net_income) if ltm_net_income is not None and ltm_net_income > 0 else None
    return {
        "price": price, "shares": shares, "debt": debt, "lease": lease, "cash": cash,
        "minority_preferred": (minority or 0.0) + (preferred or 0.0) if minority is not None or preferred is not None else None,
        "market_cap": market_cap, "net_debt": net_debt, "enterprise_value": enterprise_value,
        "ltm_ebitda": ltm_ebitda, "ltm_net_income": ltm_net_income, "ltm_eps": ltm_eps,
        "ev_ebitda": ev_ebitda, "pe": pe,
    }


def _valuation_financial_signals(data: MarketData, peers: list[MarketData] | None = None) -> list[FinancialSignal]:
    dates = _all_quarter_dates(data)
    if not dates:
        return []
    latest = dates[-1]
    previous = _prior_year_quarter(dates, latest)
    cur = _quarter_valuation_snapshot(data, latest)
    old = _quarter_valuation_snapshot(data, previous) if previous is not None else {}
    label = _quarter_label(latest)
    signals: list[FinancialSignal] = []
    is_financial = _is_financial_company(data)
    if is_financial:
        signals.append(_signal(1, "SECTOR-SPECIFIC VALUATION",
            "EV / EBITDA is not used as a primary automatic valuation signal for financial companies; P / E and sector-specific book-value or capital metrics are more informative.",
            "금융주는 EV/EBITDA보다 P/E와 P/B, 규제자본 등 섹터 특화 지표가 더 유의미하므로 EV/EBITDA 기반 자동 valuation 판정을 적용하지 않았다."))

    price_g = _growth_rate(cur.get("price"), old.get("price")) if old else None
    eps_g = _growth_rate(cur.get("ltm_eps"), old.get("ltm_eps")) if old and (old.get("ltm_eps") or 0) > 0 else None
    shares_g = _growth_rate(cur.get("shares"), old.get("shares")) if old else None
    mcap_g = _growth_rate(cur.get("market_cap"), old.get("market_cap")) if old else None
    ev_g = _growth_rate(cur.get("enterprise_value"), old.get("enterprise_value")) if old else None
    ebitda_g = _growth_rate(cur.get("ltm_ebitda"), old.get("ltm_ebitda")) if old and (old.get("ltm_ebitda") or 0) > 0 else None

    actual_year = _latest_template_actual_year(data)
    annual_rev_g = None
    annual_nd_ebitda = None
    if actual_year is not None:
        annual_rev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], actual_year)
        prior_annual_rev = _annual_value(data.annual_income, INCOME_ALIASES["Revenue"], actual_year - 1)
        annual_rev_g = _growth_rate(annual_rev, prior_annual_rev)
        annual_ebitda = _annual_value(data.annual_income, INCOME_ALIASES["EBITDA"], actual_year)
        annual_debt = _annual_value(data.annual_balance, BALANCE_ALIASES["Total debt (incl. leases)"], actual_year)
        annual_cash = _annual_value(data.annual_balance, BALANCE_ALIASES["Cash and equivalents"], actual_year)
        if annual_debt is not None and annual_cash is not None and annual_ebitda is not None and annual_ebitda > 0:
            annual_nd_ebitda = (annual_debt - annual_cash) / annual_ebitda

    # Current peer valuation/growth medians, using each peer's latest available quarter.
    peer_ev_multiples: list[float] = []
    peer_pe_multiples: list[float] = []
    peer_ebitda_growth: list[float] = []
    peer_eps_growth: list[float] = []
    for peer in peers or []:
        peer_dates = _all_quarter_dates(peer)
        if not peer_dates:
            continue
        peer_latest = peer_dates[-1]
        peer_prior = _prior_year_quarter(peer_dates, peer_latest)
        peer_cur = _quarter_valuation_snapshot(peer, peer_latest)
        peer_old = _quarter_valuation_snapshot(peer, peer_prior) if peer_prior is not None else {}
        v = peer_cur.get("ev_ebitda")
        if v is not None and math.isfinite(float(v)) and v > 0:
            peer_ev_multiples.append(float(v))
        v = peer_cur.get("pe")
        if v is not None and math.isfinite(float(v)) and v > 0:
            peer_pe_multiples.append(float(v))
        if peer_old and (peer_old.get("ltm_ebitda") or 0) > 0:
            g = _growth_rate(peer_cur.get("ltm_ebitda"), peer_old.get("ltm_ebitda"))
            if g is not None and math.isfinite(float(g)):
                peer_ebitda_growth.append(float(g))
        if peer_old and (peer_old.get("ltm_eps") or 0) > 0:
            g = _growth_rate(peer_cur.get("ltm_eps"), peer_old.get("ltm_eps"))
            if g is not None and math.isfinite(float(g)):
                peer_eps_growth.append(float(g))

    peer_ev_median = statistics.median(peer_ev_multiples) if peer_ev_multiples else None
    peer_pe_median = statistics.median(peer_pe_multiples) if peer_pe_multiples else None
    peer_ebitda_g_median = statistics.median(peer_ebitda_growth) if peer_ebitda_growth else None
    peer_eps_g_median = statistics.median(peer_eps_growth) if peer_eps_growth else None

    if price_g is not None and eps_g is not None:
        spread = price_g - eps_g
        if price_g < 0 and eps_g > 0:
            signals.append(_signal(3, "FUNDAMENTAL / PRICE DIVERGENCE", f"The share price declined {abs(price_g)*100:.1f}% while LTM EPS increased {eps_g*100:.1f}%, creating a notable divergence between improving earnings and market valuation.", f"주가는 {abs(price_g)*100:.1f}% 하락한 반면 LTM EPS는 {eps_g*100:.1f}% 증가해 펀더멘털 개선과 시장 valuation 사이의 괴리가 확대되고 있다."))
        elif spread >= 0.20:
            signals.append(_signal(7, "MULTIPLE EXPANSION", f"Share-price appreciation outpaced LTM EPS growth by {spread*100:.1f} percentage points, indicating that multiple expansion has been a major driver of returns.", f"주가 상승률이 LTM EPS 성장률을 {spread*100:.1f}%p 상회해 실적 성장보다 multiple expansion이 주가 상승을 크게 견인했다."))
        elif spread <= -0.20:
            signals.append(_signal(5, "MULTIPLE COMPRESSION", f"LTM EPS growth outpaced the share price by {abs(spread)*100:.1f} percentage points, indicating meaningful equity-multiple compression.", f"LTM EPS 성장률이 주가 상승률을 {abs(spread)*100:.1f}%p 상회해 equity multiple compression이 나타나고 있다."))

    if shares_g is not None:
        if shares_g <= -0.05:
            signals.append(_signal(15, "AGGRESSIVE BUYBACK", f"Diluted shares declined {abs(shares_g)*100:.1f}% versus the comparison period, providing meaningful buyback-driven support to per-share value.", f"희석주식수가 비교 기간 대비 {abs(shares_g)*100:.1f}% 감소해 적극적인 자사주 매입이 주당가치를 지원하고 있다."))
        elif shares_g >= 0.10:
            signals.append(_signal(4, "SEVERE DILUTION", f"Diluted shares increased {shares_g*100:.1f}%, creating substantial shareholder dilution.", f"희석주식수가 {shares_g*100:.1f}% 증가해 상당한 shareholder dilution이 발생하고 있다."))
        elif shares_g >= 0.05:
            signals.append(_signal(7, "HIGH DILUTION", f"Diluted shares increased {shares_g*100:.1f}%, creating a meaningful dilution drag on per-share value.", f"희석주식수가 {shares_g*100:.1f}% 증가해 주당가치에 의미 있는 희석 부담이 발생하고 있다."))
        elif shares_g >= 0.02:
            signals.append(_signal(16, "MODERATE DILUTION", f"Diluted shares increased {shares_g*100:.1f}%, indicating moderate dilution.", f"희석주식수가 {shares_g*100:.1f}% 증가해 완만한 희석이 발생하고 있다."))

    if mcap_g is not None and eps_g is not None:
        spread = mcap_g - eps_g
        if spread >= 0.20:
            signals.append(_signal(11, "EQUITY MULTIPLE EXPANSION", f"Market-cap growth exceeded LTM EPS growth by {spread*100:.1f} percentage points, indicating a richer equity valuation multiple.", f"시가총액 증가율이 LTM EPS 성장률을 {spread*100:.1f}%p 상회해 equity valuation multiple이 확대되고 있다."))
        elif spread <= -0.20:
            signals.append(_signal(10, "EQUITY VALUATION COMPRESSION", f"LTM EPS growth exceeded market-cap growth by {abs(spread)*100:.1f} percentage points, indicating equity valuation compression.", f"LTM EPS 성장률이 시가총액 증가율을 {abs(spread)*100:.1f}%p 상회해 equity valuation compression이 나타나고 있다."))

    if not is_financial and ev_g is not None and ebitda_g is not None:
        spread = ev_g - ebitda_g
        if spread >= 0.20:
            signals.append(_signal(8, "ENTERPRISE MULTIPLE EXPANSION", f"Enterprise value grew {spread*100:.1f} percentage points faster than LTM EBITDA, indicating strong expansion in the enterprise valuation multiple.", f"Enterprise value 증가율이 LTM EBITDA 성장률을 {spread*100:.1f}%p 상회해 enterprise valuation multiple이 크게 확대되고 있다."))
        elif spread <= -0.20:
            signals.append(_signal(6, "ENTERPRISE MULTIPLE COMPRESSION", f"LTM EBITDA growth outpaced enterprise-value growth by {abs(spread)*100:.1f} percentage points, indicating material EV/EBITDA compression.", f"LTM EBITDA 성장률이 enterprise value 증가율을 {abs(spread)*100:.1f}%p 상회해 EV/EBITDA multiple이 의미 있게 압축되고 있다."))

    ltm_ebitda = cur.get("ltm_ebitda")
    if not is_financial and ltm_ebitda is not None:
        if ltm_ebitda <= 0:
            signals.append(_signal(2, "EV / EBITDA N/M", "LTM EBITDA is non-positive, so EV / EBITDA is not a meaningful valuation measure for the latest period.", "LTM EBITDA가 0 이하이므로 최신 기간의 EV/EBITDA는 유의미한 valuation 지표로 사용하기 어렵다."))
        elif ebitda_g is not None:
            if ebitda_g >= 0.25:
                signals.append(_signal(17, "STRONG EBITDA MOMENTUM", f"LTM EBITDA increased {ebitda_g*100:.1f}%, indicating strong recent earnings momentum.", f"LTM EBITDA가 {ebitda_g*100:.1f}% 증가해 최근 이익 모멘텀이 강한 수준이다."))
            elif ebitda_g < 0:
                signals.append(_signal(6, "EBITDA DETERIORATION", f"LTM EBITDA declined {abs(ebitda_g)*100:.1f}%, indicating weakening recent operating fundamentals.", f"LTM EBITDA가 {abs(ebitda_g)*100:.1f}% 감소해 최근 영업 펀더멘털이 약화되고 있다."))

    ev_ebitda = cur.get("ev_ebitda")
    if not is_financial and ev_ebitda is not None:
        if ev_ebitda < 5:
            signals.append(_signal(18, "VERY LOW EV / EBITDA", f"The company trades at {ev_ebitda:.2f}x LTM EV / EBITDA, a very low absolute multiple that requires distinguishing genuine undervaluation from structural fundamental risk.", f"회사는 LTM EV/EBITDA {ev_ebitda:.2f}x에 거래되고 있어 절대적으로 매우 낮은 multiple이지만 구조적 저평가와 펀더멘털 리스크를 구분할 필요가 있다."))
        elif ev_ebitda < 8:
            signals.append(_signal(21, "LOW EV / EBITDA", f"The company trades at a relatively low {ev_ebitda:.2f}x LTM EV / EBITDA multiple.", f"회사는 LTM EV/EBITDA {ev_ebitda:.2f}x로 비교적 낮은 valuation multiple에 거래되고 있다."))
        elif ev_ebitda > 20:
            signals.append(_signal(8, "VERY HIGH EV / EBITDA", f"The company trades at {ev_ebitda:.2f}x LTM EV / EBITDA, embedding substantial expectations for future growth and margin expansion.", f"회사는 LTM EV/EBITDA {ev_ebitda:.2f}x로 높은 성장과 수익성 개선 기대가 상당 부분 반영된 valuation에 거래되고 있다."))
        elif ev_ebitda > 16:
            signals.append(_signal(13, "HIGH EV / EBITDA", f"The {ev_ebitda:.2f}x LTM EV / EBITDA multiple represents a relatively high absolute valuation.", f"LTM EV/EBITDA가 {ev_ebitda:.2f}x로 절대적인 valuation 수준이 높은 편이다."))
        else:
            signals.append(_signal(24, "EV / EBITDA PROFILE", f"The company trades at {ev_ebitda:.2f}x LTM EV / EBITDA, an intermediate absolute valuation that should be interpreted primarily against peers and growth quality.", f"회사는 LTM EV/EBITDA {ev_ebitda:.2f}x에 거래되고 있어 절대 수준은 중간 범위이며 peer 및 성장의 질과 함께 해석하는 것이 적절하다.", severity="NEUTRAL"))

    pe = cur.get("pe")
    if cur.get("ltm_net_income") is not None and (cur.get("ltm_net_income") or 0) <= 0:
        signals.append(_signal(2, "P / E N/M", "LTM net income is non-positive, so P / E is not a meaningful valuation measure for the latest period.", "LTM 순이익이 0 이하이므로 최신 기간의 P/E는 유의미한 valuation 지표로 사용하기 어렵다."))
    elif pe is not None:
        if pe < 8:
            signals.append(_signal(17, "VERY LOW P / E", f"The company trades at {pe:.2f}x LTM P / E, a very low earnings multiple that may reflect either undervaluation or elevated structural risk.", f"회사는 LTM P/E {pe:.2f}x로 매우 낮은 earnings multiple에 거래되고 있어 저평가와 구조적 리스크를 함께 점검할 필요가 있다."))
        elif pe < 12:
            signals.append(_signal(20, "LOW P / E", f"The company trades at a relatively low {pe:.2f}x LTM P / E multiple.", f"회사는 LTM P/E {pe:.2f}x로 비교적 낮은 earnings multiple에 거래되고 있다."))
        elif pe > 35:
            signals.append(_signal(7, "VERY HIGH P / E", f"The company trades at {pe:.2f}x LTM P / E, implying very high market expectations for future earnings growth.", f"회사는 LTM P/E {pe:.2f}x로 향후 이익 성장에 대한 시장 기대가 매우 높게 반영되어 있다."))
        elif pe > 25:
            signals.append(_signal(13, "HIGH P / E", f"The {pe:.2f}x LTM P / E multiple embeds a meaningful growth premium.", f"LTM P/E가 {pe:.2f}x로 성장 기대에 대한 상당한 premium이 반영되어 있다."))
        else:
            signals.append(_signal(24, "P / E PROFILE", f"The company trades at {pe:.2f}x LTM P / E, an intermediate earnings multiple that is best interpreted relative to peers and expected earnings growth.", f"회사는 LTM P/E {pe:.2f}x에 거래되고 있어 earnings multiple은 중간 범위이며 peer 및 향후 이익 성장률과 함께 해석하는 것이 적절하다.", severity="NEUTRAL"))

    # Detect low-multiple value traps and expensive companies with weakening fundamentals.
    if not is_financial and ((pe is not None and pe < 12) or (ev_ebitda is not None and ev_ebitda < 8)):
        weak_fundamental = (annual_rev_g is not None and annual_rev_g < 0) or (ebitda_g is not None and ebitda_g < 0)
        leveraged = annual_nd_ebitda is not None and annual_nd_ebitda >= 3
        if weak_fundamental and leveraged:
            signals.append(_signal(2, "VALUATION TRAP RISK",
                "The stock screens at a low headline multiple, but declining fundamentals and elevated leverage raise the risk that the discount reflects structural weakness rather than mispricing.",
                "표면적인 valuation multiple은 낮지만 펀더멘털 둔화와 높은 레버리지가 동시에 나타나 현재 discount가 단순 저평가가 아니라 구조적 약점을 반영할 가능성이 있다.",
                severity="SEVERE"))

    if not is_financial and ((pe is not None and pe > 25) or (ev_ebitda is not None and ev_ebitda > 16)) and ebitda_g is not None and ebitda_g < 0:
        signals.append(_signal(2, "EXPENSIVE WITH DETERIORATING FUNDAMENTALS",
            "The company retains a premium valuation while LTM EBITDA is declining, creating an unfavorable combination of high expectations and weakening operating performance.",
            "높은 valuation premium이 유지되는 가운데 LTM EBITDA가 감소해 높은 시장 기대와 영업 펀더멘털 약화가 동시에 나타나고 있다.",
            severity="SEVERE"))

    cash = cur.get("cash")
    market_cap = cur.get("market_cap")
    cash_mcap = _safe_div(cash, market_cap)
    if cash_mcap is not None:
        if cash_mcap >= 0.50:
            signals.append(_signal(14, "VERY CASH RICH", f"Cash equals {cash_mcap*100:.1f}% of market capitalization, making the operating-business value materially lower than the headline equity value.", f"현금이 시가총액의 {cash_mcap*100:.1f}%에 달해 operating business에 부여된 순수 가치가 표면적인 equity value보다 크게 낮다."))
        elif cash_mcap >= 0.30:
            signals.append(_signal(19, "CASH-RICH VALUATION", f"Cash represents {cash_mcap*100:.1f}% of market capitalization, providing a substantial valuation and liquidity cushion.", f"현금이 시가총액의 {cash_mcap*100:.1f}%를 차지해 valuation 및 유동성 측면에서 상당한 완충력을 제공한다."))

    mp_ev = _safe_div(cur.get("minority_preferred"), cur.get("enterprise_value"))
    if mp_ev is not None and mp_ev > 0.10:
        signals.append(_signal(12, "MATERIAL NON-COMMON CLAIMS", f"Minority interests and preferred claims represent {mp_ev*100:.1f}% of enterprise value, making non-common capital claims material to the EV bridge.", f"비지배지분 및 우선주가 enterprise value의 {mp_ev*100:.1f}%를 차지해 EV bridge에서 보통주 외 자본청구권의 영향이 크다."))

    # Relative valuation against automatically selected peers.
    if not is_financial and ev_ebitda is not None and peer_ev_median is not None and peer_ev_median > 0:
        rel = ev_ebitda / peer_ev_median - 1.0
        if rel <= -0.20:
            signals.append(_signal(9, "RELATIVE VALUATION DISCOUNT",
                f"LTM EV / EBITDA is {abs(rel)*100:.1f}% below the peer median ({ev_ebitda:.2f}x vs. {peer_ev_median:.2f}x).",
                f"LTM EV/EBITDA가 {ev_ebitda:.2f}x로 peer median {peer_ev_median:.2f}x 대비 {abs(rel)*100:.1f}% 할인되어 거래되고 있다.", severity="POSITIVE"))
            if ebitda_g is not None and peer_ebitda_g_median is not None and ebitda_g >= peer_ebitda_g_median + 0.05:
                signals.append(_signal(4, "GROWTH-ADJUSTED DISCOUNT",
                    f"The company trades at a {abs(rel)*100:.1f}% EV / EBITDA discount to peers while LTM EBITDA growth exceeds the peer median by {(ebitda_g-peer_ebitda_g_median)*100:.1f} percentage points, strengthening the relative-value case.",
                    f"Peer 대비 EV/EBITDA는 {abs(rel)*100:.1f}% 할인되어 있는 반면 LTM EBITDA 성장률은 peer median을 {(ebitda_g-peer_ebitda_g_median)*100:.1f}%p 상회해 성장 조정 기준의 상대적 valuation 매력이 높다.", severity="POSITIVE"))
        elif rel >= 0.20:
            signals.append(_signal(10, "RELATIVE VALUATION PREMIUM",
                f"LTM EV / EBITDA is {rel*100:.1f}% above the peer median ({ev_ebitda:.2f}x vs. {peer_ev_median:.2f}x), requiring superior growth or profitability to justify the premium.",
                f"LTM EV/EBITDA가 {ev_ebitda:.2f}x로 peer median {peer_ev_median:.2f}x 대비 {rel*100:.1f}% premium에 거래되고 있어 높은 성장 또는 수익성으로 premium을 정당화할 필요가 있다.", severity="WARNING"))
            if ebitda_g is not None and peer_ebitda_g_median is not None and ebitda_g <= peer_ebitda_g_median - 0.05:
                signals.append(_signal(3, "PREMIUM AT RISK",
                    f"The company trades at a {rel*100:.1f}% EV / EBITDA premium to peers despite LTM EBITDA growth trailing the peer median by {(peer_ebitda_g_median-ebitda_g)*100:.1f} percentage points, increasing multiple-compression risk.",
                    f"Peer 대비 EV/EBITDA는 {rel*100:.1f}% premium인 반면 LTM EBITDA 성장률은 peer median을 {(peer_ebitda_g_median-ebitda_g)*100:.1f}%p 하회해 multiple compression 위험이 높다.", severity="WARNING"))

    if pe is not None and peer_pe_median is not None and peer_pe_median > 0:
        rel = pe / peer_pe_median - 1.0
        if rel <= -0.20 and eps_g is not None and peer_eps_g_median is not None and eps_g >= peer_eps_g_median + 0.05:
            signals.append(_signal(5, "EARNINGS GROWTH-ADJUSTED DISCOUNT",
                f"P / E is {abs(rel)*100:.1f}% below the peer median while LTM EPS growth exceeds peers by {(eps_g-peer_eps_g_median)*100:.1f} percentage points, indicating an attractive growth-adjusted earnings multiple.",
                f"P/E가 peer median 대비 {abs(rel)*100:.1f}% 할인된 가운데 LTM EPS 성장률은 peer를 {(eps_g-peer_eps_g_median)*100:.1f}%p 상회해 이익 성장 조정 기준의 valuation 매력이 존재한다.", severity="POSITIVE"))
        elif rel >= 0.20 and eps_g is not None and peer_eps_g_median is not None and eps_g <= peer_eps_g_median - 0.05:
            signals.append(_signal(4, "P / E PREMIUM AT RISK",
                f"P / E is {rel*100:.1f}% above the peer median while LTM EPS growth trails peers by {(peer_eps_g_median-eps_g)*100:.1f} percentage points, weakening support for the earnings premium.",
                f"P/E가 peer median 대비 {rel*100:.1f}% premium인 반면 LTM EPS 성장률은 peer를 {(peer_eps_g_median-eps_g)*100:.1f}%p 하회해 earnings premium의 정당성이 약화되고 있다.", severity="WARNING"))

    net_debt = cur.get("net_debt")
    if net_debt is not None and net_debt < 0:
        signals.append(_signal(18, "NET CASH", "The latest valuation bridge reflects a net-cash position, reducing enterprise value relative to market capitalization.", "최신 valuation bridge에서 순현금 상태를 보이며 market capitalization 대비 enterprise value를 낮추는 요인으로 작용한다."))

    if not is_financial and old and cur.get("ev_ebitda") is not None and old.get("ev_ebitda") is not None:
        mult_delta = cur["ev_ebitda"] - old["ev_ebitda"]
        if mult_delta <= -2:
            signals.append(_signal(10, "EV / EBITDA COMPRESSION", f"EV / EBITDA fell from {old['ev_ebitda']:.2f}x to {cur['ev_ebitda']:.2f}x, showing a material compression in the enterprise valuation multiple.", f"EV/EBITDA가 {old['ev_ebitda']:.2f}x에서 {cur['ev_ebitda']:.2f}x로 하락해 enterprise valuation multiple이 의미 있게 압축됐다."))
        elif mult_delta >= 2:
            signals.append(_signal(11, "EV / EBITDA EXPANSION", f"EV / EBITDA rose from {old['ev_ebitda']:.2f}x to {cur['ev_ebitda']:.2f}x, showing a material expansion in the enterprise valuation multiple.", f"EV/EBITDA가 {old['ev_ebitda']:.2f}x에서 {cur['ev_ebitda']:.2f}x로 상승해 enterprise valuation multiple이 의미 있게 확대됐다."))

    if old and cur.get("pe") is not None and old.get("pe") is not None and eps_g is not None:
        if eps_g > 0 and cur["pe"] < old["pe"]:
            signals.append(_signal(9, "GROWING INTO VALUATION", f"LTM EPS increased while P / E declined from {old['pe']:.2f}x to {cur['pe']:.2f}x, indicating that earnings growth is absorbing valuation pressure.", f"LTM EPS가 증가하는 동시에 P/E가 {old['pe']:.2f}x에서 {cur['pe']:.2f}x로 하락해 이익 성장을 통해 기존 valuation 부담을 소화하고 있다."))

    return _select_financial_signals(signals, "valuation")


def _insert_paragraph_after(anchor: Any, parent: Any) -> Paragraph:
    from docx.oxml import OxmlElement
    element = OxmlElement("w:p")
    anchor.addnext(element)
    return Paragraph(element, parent)


def _financial_section_diagnosis(signals: list[FinancialSignal], language: str) -> str:
    counts = {key: 0 for key in ("CRITICAL", "SEVERE", "WARNING", "WEAK", "POSITIVE", "NEUTRAL")}
    for item in signals:
        counts[item.severity if item.severity in counts else "NEUTRAL"] += 1
    negatives = counts["CRITICAL"] + counts["SEVERE"] + counts["WARNING"] + counts["WEAK"]
    positives = counts["POSITIVE"]
    if counts["CRITICAL"]:
        skew = "CRITICAL"
    elif counts["SEVERE"]:
        skew = "NEGATIVE"
    elif counts["WARNING"] >= 2 and positives == 0:
        skew = "CAUTIOUS"
    elif positives >= 2 and negatives == 0:
        skew = "POSITIVE"
    else:
        skew = "MIXED"
    if language.lower().strip() == "ko":
        return (f"종합 진단 — {skew} | Critical {counts['CRITICAL']} · Severe {counts['SEVERE']} · "
                f"Warning {counts['WARNING']} · Weak {counts['WEAK']} · Positive {counts['POSITIVE']}")
    return (f"Diagnosis — {skew} | Critical {counts['CRITICAL']} · Severe {counts['SEVERE']} · "
            f"Warning {counts['WARNING']} · Weak {counts['WEAK']} · Positive {counts['POSITIVE']}")


def _severity_label(severity: str, language: str) -> str:
    severity = severity.upper()
    if language.lower().strip() != "ko":
        return severity
    return {
        "CRITICAL": "매우 심각", "SEVERE": "심각", "WARNING": "경고",
        "WEAK": "약함", "POSITIVE": "긍정", "NEUTRAL": "중립",
    }.get(severity, severity)


def _insert_financial_readout(table: Table, signals: list[FinancialSignal], *, language: str) -> None:
    if not signals:
        return
    language = language.lower().strip()
    parent = table._parent
    anchor = table._tbl
    title = _insert_paragraph_after(anchor, parent)
    anchor = title._p
    title.paragraph_format.space_before = Pt(4)
    title.paragraph_format.space_after = Pt(2)
    title.paragraph_format.keep_with_next = True
    run = title.add_run("AUTOMATED FINANCIAL READOUT" if language != "ko" else "자동 재무 분석")
    run.bold = True
    run.font.size = Pt(8.5)

    diagnosis = _insert_paragraph_after(anchor, parent)
    anchor = diagnosis._p
    diagnosis.paragraph_format.left_indent = Pt(8)
    diagnosis.paragraph_format.space_before = Pt(0)
    diagnosis.paragraph_format.space_after = Pt(2)
    diagnosis.paragraph_format.keep_with_next = True
    d_run = diagnosis.add_run(_financial_section_diagnosis(signals, language))
    d_run.bold = True
    d_run.italic = True
    d_run.font.size = Pt(7.5)

    for item in signals:
        p = _insert_paragraph_after(anchor, parent)
        anchor = p._p
        p.paragraph_format.left_indent = Pt(8)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.keep_together = True
        p.paragraph_format.line_spacing = 1.0
        tag = p.add_run(f"[{_severity_label(item.severity, language)} | {item.tag}] ")
        tag.bold = True
        tag.font.size = Pt(7.5)
        body = p.add_run(item.ko if language == "ko" else item.en)
        body.font.size = Pt(8.0)


def _append_automated_financial_commentary(doc: DocumentObject, data: MarketData, peers: list[MarketData] | None = None, *, language: str) -> None:
    income_table = _find_table_containing_row(doc, "Revenue growth (%)")
    if income_table is not None:
        _insert_financial_readout(income_table, _income_financial_signals(data), language=language)

    balance_table = None
    for candidate in _all_tables(doc):
        rows = _row_map(candidate)
        if "cash and equivalents" in rows and "current ratio (x)" in rows and len(candidate.columns) == 6:
            balance_table = candidate
            break
    if balance_table is not None:
        _insert_financial_readout(balance_table, _balance_financial_signals(data), language=language)

    valuation_table = None
    for candidate in _all_tables(doc):
        rows = _row_map(candidate)
        if "share price" in rows and "enterprise value" in rows and "ltm ebitda" in rows and len(candidate.columns) == 7:
            valuation_table = candidate
            break
    if valuation_table is not None:
        _insert_financial_readout(valuation_table, _valuation_financial_signals(data, peers or []), language=language)


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
        ev_ebitda = _safe_div(enterprise_value, ltm_ebitda) if ltm_ebitda is not None and ltm_ebitda > 0 else None
        pe = _safe_div(market_cap, ltm_net_income) if ltm_net_income is not None and ltm_net_income > 0 else None

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
            "ev / ebitda (x)": "N/M" if ltm_ebitda is not None and ltm_ebitda <= 0 else _format_ratio(ev_ebitda),
            "p / e (x)": "N/M" if ltm_net_income is not None and ltm_net_income <= 0 else _format_ratio(pe),
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
    _append_automated_financial_commentary(doc, data, peers, language=language)
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
