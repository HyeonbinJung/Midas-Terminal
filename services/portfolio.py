from __future__ import annotations

import hmac
import math
import os
import re
import sqlite3
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = Path(os.getenv("PORTFOLIO_DB_PATH", str(DATA_DIR / "midas_portfolio.sqlite3"))).expanduser()
TRADE_PASSWORD = os.getenv("PORTFOLIO_TRADE_PASSWORD", "3824")
INITIAL_CASH = float(os.getenv("PORTFOLIO_INITIAL_CASH", "1000000"))
MAX_PARTICIPATION_RATE = max(0.005, min(float(os.getenv("PORTFOLIO_MAX_PARTICIPATION", "0.05")), 0.25))
NY_TZ = ZoneInfo("America/New_York")
_TICKER_RE = re.compile(r"^[A-Z0-9.\-]{1,15}$")
_DB_LOCK = threading.RLock()
_QUOTE_LOCK = threading.RLock()
_QUOTE_CACHE: dict[str, tuple[float, "Quote"]] = {}

try:
    import exchange_calendars as xcals
except Exception:  # pragma: no cover
    xcals = None


@dataclass(slots=True)
class Quote:
    ticker: str
    price: float
    bid: float
    ask: float
    bid_size: float | None = None
    ask_size: float | None = None
    volume: float | None = None
    avg_volume: float | None = None
    minute_volume: float | None = None
    day_open: float | None = None
    day_high: float | None = None
    day_low: float | None = None
    previous_close: float | None = None
    currency: str = "USD"
    exchange: str = ""
    company_name: str = ""
    quoted_at: str = ""
    source: str = "Yahoo Finance / yfinance"
    spread_source: str = "YAHOO_BID_ASK"

    @property
    def mid(self) -> float:
        return (self.bid + self.ask) / 2 if self.bid > 0 and self.ask > 0 else self.price

    @property
    def spread(self) -> float:
        return max(0.0, self.ask - self.bid)

    @property
    def spread_bps(self) -> float:
        return (self.spread / self.mid * 10000.0) if self.mid > 0 else 0.0

    @property
    def day_change_pct(self) -> float | None:
        if self.previous_close and self.previous_close > 0:
            return (self.price / self.previous_close - 1.0) * 100.0
        return None


def _now_utc() -> datetime:
    return datetime.now(timezone.utc)


def _iso(dt: datetime | None = None) -> str:
    return (dt or _now_utc()).astimezone(timezone.utc).isoformat(timespec="seconds")


def _connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=10, isolation_level=None)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=5000")
    return conn


def _ensure_column(conn: sqlite3.Connection, table: str, name: str, ddl: str) -> None:
    cols = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    if name not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")


def init_portfolio_db() -> None:
    with _DB_LOCK, _connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS portfolio_account (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                initial_cash REAL NOT NULL,
                cash REAL NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS portfolio_positions (
                ticker TEXT PRIMARY KEY,
                quantity REAL NOT NULL CHECK (quantity >= 0),
                avg_cost REAL NOT NULL CHECK (avg_cost >= 0),
                realized_pnl REAL NOT NULL DEFAULT 0,
                last_price REAL,
                last_price_at TEXT,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS portfolio_orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                side TEXT NOT NULL CHECK (side IN ('BUY','SELL')),
                ticker TEXT NOT NULL,
                order_type TEXT NOT NULL CHECK (order_type IN ('MARKET','LIMIT')),
                quantity REAL NOT NULL CHECK (quantity > 0),
                limit_price REAL,
                filled_quantity REAL NOT NULL DEFAULT 0,
                avg_fill_price REAL,
                status TEXT NOT NULL,
                tif TEXT NOT NULL DEFAULT 'DAY',
                note TEXT NOT NULL DEFAULT '',
                cancel_reason TEXT NOT NULL DEFAULT '',
                last_fill_bucket TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_portfolio_orders_status ON portfolio_orders(status, ticker);
            CREATE INDEX IF NOT EXISTS idx_portfolio_orders_time ON portfolio_orders(created_at DESC);

            CREATE TABLE IF NOT EXISTS portfolio_transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                executed_at TEXT NOT NULL,
                side TEXT NOT NULL CHECK (side IN ('BUY', 'SELL')),
                ticker TEXT NOT NULL,
                quantity REAL NOT NULL CHECK (quantity > 0),
                price REAL NOT NULL CHECK (price > 0),
                gross_amount REAL NOT NULL CHECK (gross_amount >= 0),
                realized_pnl REAL NOT NULL DEFAULT 0,
                cash_after REAL NOT NULL,
                note TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_portfolio_transactions_time ON portfolio_transactions(executed_at DESC);

            CREATE TABLE IF NOT EXISTS portfolio_snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                captured_at TEXT NOT NULL,
                cash REAL NOT NULL,
                positions_value REAL NOT NULL,
                total_equity REAL NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_portfolio_snapshots_time ON portfolio_snapshots(captured_at ASC);
            """
        )
        _ensure_column(conn, "portfolio_orders", "last_fill_bucket", "TEXT")

        # Migrate v5.10 ledgers without losing the existing shared portfolio.
        for name, ddl in (
            ("order_id", "INTEGER"),
            ("order_type", "TEXT NOT NULL DEFAULT 'MARKET'"),
            ("limit_price", "REAL"),
            ("reference_price", "REAL"),
            ("bid", "REAL"),
            ("ask", "REAL"),
            ("slippage_bps", "REAL NOT NULL DEFAULT 0"),
            ("participation_pct", "REAL"),
            ("liquidity_source", "TEXT NOT NULL DEFAULT ''"),
        ):
            _ensure_column(conn, "portfolio_transactions", name, ddl)

        now = _iso()
        conn.execute(
            "INSERT OR IGNORE INTO portfolio_account(id, initial_cash, cash, created_at, updated_at) VALUES (1, ?, ?, ?, ?)",
            (INITIAL_CASH, INITIAL_CASH, now, now),
        )
        if not conn.execute("SELECT COUNT(*) AS n FROM portfolio_snapshots").fetchone()["n"]:
            conn.execute(
                "INSERT INTO portfolio_snapshots(captured_at, cash, positions_value, total_equity) VALUES (?, ?, 0, ?)",
                (now, INITIAL_CASH, INITIAL_CASH),
            )


def verify_trade_password(password: str) -> bool:
    return hmac.compare_digest(str(password or ""), str(TRADE_PASSWORD))


def normalize_ticker(ticker: str) -> str:
    symbol = str(ticker or "").strip().upper()
    if not _TICKER_RE.match(symbol):
        raise ValueError("Enter a valid U.S. ticker symbol.")
    return symbol


def market_status(now: datetime | None = None) -> dict[str, Any]:
    now_utc = (now or _now_utc()).astimezone(timezone.utc)
    now_et = now_utc.astimezone(NY_TZ)
    current = pd.Timestamp(now_utc).floor("min")
    if xcals is not None:
        try:
            cal = xcals.get_calendar("XNYS")
            is_open = bool(cal.is_open_on_minute(current))
            start = (pd.Timestamp(now_utc.date()) - pd.Timedelta(days=7)).strftime("%Y-%m-%d")
            end = (pd.Timestamp(now_utc.date()) + pd.Timedelta(days=10)).strftime("%Y-%m-%d")
            schedule = cal.schedule.loc[start:end]
            session_open = session_close = None
            for _, row in schedule.iterrows():
                op, cl = pd.Timestamp(row["open"]), pd.Timestamp(row["close"])
                if op <= current <= cl:
                    session_open, session_close = op, cl
                    break
                if current < op and session_open is None:
                    session_open, session_close = op, cl
            return {
                "is_open": is_open,
                "market": "U.S. REGULAR SESSION",
                "timezone": "America/New_York",
                "now_et": now_et.isoformat(timespec="seconds"),
                "session_open_et": session_open.tz_convert("America/New_York").isoformat() if session_open is not None else None,
                "session_close_et": session_close.tz_convert("America/New_York").isoformat() if session_close is not None else None,
                "calendar": "XNYS",
                "note": "Orders/fills are processed only during the NYSE regular session, including holidays and early closes.",
            }
        except Exception:
            pass
    minutes = now_et.hour * 60 + now_et.minute
    is_open = now_et.weekday() < 5 and (9 * 60 + 30) <= minutes < 16 * 60
    return {
        "is_open": is_open,
        "market": "U.S. REGULAR SESSION",
        "timezone": "America/New_York",
        "now_et": now_et.isoformat(timespec="seconds"),
        "session_open_et": None,
        "session_close_et": None,
        "calendar": "WEEKDAY_FALLBACK",
        "note": "Fallback regular-session check. Install exchange-calendars for holiday/early-close accuracy.",
    }


def _finite_positive(value: Any) -> float | None:
    try:
        v = float(value)
    except (TypeError, ValueError):
        return None
    return v if math.isfinite(v) and v > 0 else None


def _safe_get(obj: Any, key: str) -> Any:
    try:
        if hasattr(obj, "get"):
            return obj.get(key)
    except Exception:
        pass
    try:
        return getattr(obj, key)
    except Exception:
        return None


def _fetch_yfinance_quote_uncached(ticker: str) -> Quote:
    """Fetch a best-effort Yahoo quote.

    Yahoo/yfinance is not an exchange-certified real-time or Level-II feed. When
    Yahoo does not expose usable bid/ask, an indicative spread is synthesized
    from price and day range and clearly marked as such in the payload.
    """
    symbol = normalize_ticker(ticker)
    import yfinance as yf

    obj = yf.Ticker(symbol)
    fast: Any = {}
    info: dict[str, Any] = {}
    quoted_at = _iso()
    try:
        fast = obj.fast_info
    except Exception:
        fast = {}

    def pick(*keys: str) -> float | None:
        for key in keys:
            v = _finite_positive(_safe_get(fast, key))
            if v is not None:
                return v
        for key in keys:
            v = _finite_positive(info.get(key))
            if v is not None:
                return v
        return None

    # info is needed for bid/ask/volume on most yfinance versions.
    try:
        info = obj.info or {}
    except Exception:
        info = {}

    price = pick("last_price", "lastPrice", "regular_market_price", "regularMarketPrice", "currentPrice")
    previous_close = pick("previous_close", "previousClose", "regularMarketPreviousClose")
    day_open = pick("open", "regularMarketOpen")
    day_high = pick("day_high", "dayHigh", "regularMarketDayHigh")
    day_low = pick("day_low", "dayLow", "regularMarketDayLow")
    volume = pick("last_volume", "volume", "regularMarketVolume")
    avg_volume = pick("three_month_average_volume", "ten_day_average_volume", "averageVolume", "averageVolume10days")
    bid = pick("bid")
    ask = pick("ask")
    bid_size = pick("bidSize")
    ask_size = pick("askSize")
    currency = str(_safe_get(fast, "currency") or info.get("currency") or "USD")
    exchange = str(_safe_get(fast, "exchange") or info.get("exchange") or info.get("fullExchangeName") or "")
    company_name = str(info.get("shortName") or info.get("longName") or "")
    minute_volume = None

    try:
        hist = obj.history(period="1d", interval="1m", prepost=False, auto_adjust=False)
        if hist is not None and not hist.empty:
            closes = hist.get("Close")
            if price is None and closes is not None:
                valid = closes.dropna()
                if not valid.empty:
                    price = _finite_positive(valid.iloc[-1])
                    try:
                        quoted_at = pd.Timestamp(valid.index[-1]).to_pydatetime().astimezone(timezone.utc).isoformat(timespec="seconds")
                    except Exception:
                        pass
            vols = hist.get("Volume")
            if vols is not None:
                valid_v = vols.dropna()
                if not valid_v.empty:
                    minute_volume = _finite_positive(valid_v.iloc[-1]) or 0.0
    except Exception:
        pass

    if price is None:
        raise RuntimeError(f"Yahoo Finance did not return a tradable price for {symbol}.")
    if currency and currency.upper() != "USD":
        raise RuntimeError(f"{symbol} is quoted in {currency}, not USD. The shared portfolio accepts U.S.-dollar securities only.")

    spread_source = "YAHOO_BID_ASK"
    if not bid or not ask or ask < bid:
        # Indicative fallback: 2-40 bps half-spread depending on observed day range.
        range_pct = ((day_high - day_low) / price) if day_high and day_low and price > 0 else 0.01
        half_spread_bps = max(1.0, min(20.0, 1.5 + range_pct * 350.0))
        half = price * half_spread_bps / 10000.0
        bid, ask = max(0.000001, price - half), price + half
        spread_source = "INDICATIVE_MODEL"

    return Quote(
        ticker=symbol, price=float(price), bid=float(bid), ask=float(ask),
        bid_size=bid_size, ask_size=ask_size, volume=volume, avg_volume=avg_volume,
        minute_volume=minute_volume, day_open=day_open, day_high=day_high, day_low=day_low,
        previous_close=previous_close, currency=currency or "USD", exchange=exchange,
        company_name=company_name, quoted_at=quoted_at, spread_source=spread_source,
    )


def fetch_yfinance_quote(ticker: str, *, force: bool = False) -> Quote:
    symbol = normalize_ticker(ticker)
    now = time.monotonic()
    if not force:
        with _QUOTE_LOCK:
            cached = _QUOTE_CACHE.get(symbol)
            if cached and now - cached[0] <= 3.0:
                return cached[1]
    quote = _fetch_yfinance_quote_uncached(symbol)
    with _QUOTE_LOCK:
        _QUOTE_CACHE[symbol] = (time.monotonic(), quote)
        if len(_QUOTE_CACHE) > 64:
            oldest = sorted(_QUOTE_CACHE.items(), key=lambda kv: kv[1][0])[:16]
            for key, _ in oldest:
                _QUOTE_CACHE.pop(key, None)
    return quote


def quote_payload(q: Quote) -> dict[str, Any]:
    return {
        "ticker": q.ticker, "price": q.price, "bid": q.bid, "ask": q.ask,
        "mid": q.mid, "spread": q.spread, "spread_bps": q.spread_bps,
        "spread_source": q.spread_source, "bid_size": q.bid_size, "ask_size": q.ask_size,
        "volume": q.volume, "avg_volume": q.avg_volume, "minute_volume": q.minute_volume,
        "day_open": q.day_open, "day_high": q.day_high, "day_low": q.day_low,
        "previous_close": q.previous_close, "day_change_pct": q.day_change_pct,
        "currency": q.currency, "exchange": q.exchange, "company_name": q.company_name,
        "quoted_at": q.quoted_at, "source": q.source,
    }


def fetch_chart_data(ticker: str, interval: str = "1m") -> dict[str, Any]:
    symbol = normalize_ticker(ticker)
    interval = str(interval or "1m").lower()
    mapping = {
        "1m": ("1d", "1m", "1-MINUTE"),
        "5m": ("5d", "5m", "5-MINUTE"),
        "15m": ("1mo", "15m", "15-MINUTE"),
        "1d": ("1y", "1d", "DAILY"),
        "1wk": ("5y", "1wk", "WEEKLY"),
        "1mo": ("10y", "1mo", "MONTHLY"),
    }
    if interval not in mapping:
        raise ValueError("Interval must be one of 1m, 5m, 15m, 1d, 1wk, 1mo.")
    period, yf_interval, label = mapping[interval]
    import yfinance as yf

    hist = yf.Ticker(symbol).history(period=period, interval=yf_interval, prepost=False, auto_adjust=False)
    if hist is None or hist.empty:
        raise RuntimeError(f"No chart data returned for {symbol} ({interval}).")
    hist = hist.dropna(subset=["Close"]).copy()
    if len(hist) > 1200:
        hist = hist.iloc[-1200:]
    typical = (hist["High"].astype(float) + hist["Low"].astype(float) + hist["Close"].astype(float)) / 3.0
    volume_series = hist.get("Volume", pd.Series(index=hist.index, dtype=float)).fillna(0).astype(float)
    cumulative_volume = volume_series.cumsum().replace(0, pd.NA)
    vwap = (typical * volume_series).cumsum() / cumulative_volume
    sma20 = hist["Close"].astype(float).rolling(20, min_periods=1).mean()
    ema20 = hist["Close"].astype(float).ewm(span=20, adjust=False).mean()

    rows: list[dict[str, Any]] = []
    for i, (idx, row) in enumerate(hist.iterrows()):
        ts = pd.Timestamp(idx)
        try:
            ts = ts.tz_convert("UTC") if ts.tzinfo else ts.tz_localize("UTC")
        except Exception:
            pass
        rows.append({
            "time": ts.isoformat(),
            "open": float(row["Open"]), "high": float(row["High"]), "low": float(row["Low"]), "close": float(row["Close"]),
            "volume": float(volume_series.iloc[i]),
            "vwap": None if pd.isna(vwap.iloc[i]) else float(vwap.iloc[i]),
            "sma20": float(sma20.iloc[i]), "ema20": float(ema20.iloc[i]),
        })
    return {
        "ticker": symbol, "interval": interval, "interval_label": label, "period": period,
        "rows": rows, "updated_at": _iso(),
        "source": "Yahoo Finance / yfinance; intraday availability and latency are subject to Yahoo limits.",
    }


def _account_row(conn: sqlite3.Connection) -> sqlite3.Row:
    row = conn.execute("SELECT * FROM portfolio_account WHERE id=1").fetchone()
    if row is None:
        raise RuntimeError("Portfolio account is not initialized.")
    return row


def _position_rows(conn: sqlite3.Connection) -> list[sqlite3.Row]:
    return conn.execute("SELECT * FROM portfolio_positions WHERE quantity > 0 ORDER BY ticker").fetchall()


def _reserved_cash(conn: sqlite3.Connection, exclude_order_id: int | None = None) -> float:
    sql = """
        SELECT COALESCE(SUM((quantity-filled_quantity)*limit_price),0) AS r
        FROM portfolio_orders
        WHERE side='BUY' AND order_type='LIMIT' AND status IN ('OPEN','PARTIAL')
    """
    params: tuple[Any, ...] = ()
    if exclude_order_id is not None:
        sql += " AND id<>?"
        params = (exclude_order_id,)
    return float(conn.execute(sql, params).fetchone()["r"] or 0.0)


def _reserved_sell_qty(conn: sqlite3.Connection, ticker: str, exclude_order_id: int | None = None) -> float:
    sql = """
        SELECT COALESCE(SUM(quantity-filled_quantity),0) AS r
        FROM portfolio_orders
        WHERE side='SELL' AND ticker=? AND order_type='LIMIT' AND status IN ('OPEN','PARTIAL')
    """
    params: list[Any] = [ticker]
    if exclude_order_id is not None:
        sql += " AND id<>?"
        params.append(exclude_order_id)
    return float(conn.execute(sql, tuple(params)).fetchone()["r"] or 0.0)


def _record_snapshot(conn: sqlite3.Connection, *, force: bool = False) -> None:
    account = _account_row(conn)
    rows = _position_rows(conn)
    positions_value = sum(float(r["quantity"]) * float(r["last_price"] or r["avg_cost"]) for r in rows)
    total_equity = float(account["cash"]) + positions_value
    if not force:
        latest = conn.execute("SELECT captured_at FROM portfolio_snapshots ORDER BY id DESC LIMIT 1").fetchone()
        if latest:
            try:
                age = (_now_utc() - datetime.fromisoformat(latest["captured_at"]).astimezone(timezone.utc)).total_seconds()
                if age < 60:
                    return
            except Exception:
                pass
    conn.execute(
        "INSERT INTO portfolio_snapshots(captured_at, cash, positions_value, total_equity) VALUES (?, ?, ?, ?)",
        (_iso(), float(account["cash"]), positions_value, total_equity),
    )


def _liquidity_capacity(q: Quote) -> tuple[float, float, str]:
    """Return (max shares per fill cycle, reference minute volume, source)."""
    if q.minute_volume is not None and q.minute_volume > 0:
        ref = float(q.minute_volume)
        source = "LAST_1M_VOLUME"
    elif q.avg_volume and q.avg_volume > 0:
        ref = float(q.avg_volume) / 390.0
        source = "ADV_390"
    elif q.volume and q.volume > 0:
        ref = float(q.volume) / 195.0  # conservative half-session fallback
        source = "DAY_VOLUME_ESTIMATE"
    else:
        ref = 1000.0
        source = "FALLBACK_1000_SHARES"
    capacity = max(1.0, ref * MAX_PARTICIPATION_RATE)
    return capacity, ref, source


def _execution_estimate(side: str, requested_qty: float, q: Quote, limit_price: float | None = None) -> dict[str, Any]:
    capacity, ref_minute_volume, liquidity_source = _liquidity_capacity(q)
    fill_qty = min(float(requested_qty), capacity)
    participation = min(1.0, fill_qty / max(ref_minute_volume, 1.0))
    # Impact beyond bid/ask. The square-root model rises with participation and
    # with day range, but is capped to avoid absurd simulated fills.
    day_range_bps = ((q.day_high - q.day_low) / q.price * 10000.0) if q.day_high and q.day_low and q.price > 0 else 100.0
    impact_bps = min(250.0, 1.0 + 28.0 * math.sqrt(max(participation, 0.0)) + 0.015 * min(day_range_bps, 1000.0) * math.sqrt(max(participation, 0.0)))
    base = q.ask if side == "BUY" else q.bid
    execution = base * (1.0 + impact_bps / 10000.0) if side == "BUY" else base * (1.0 - impact_bps / 10000.0)
    marketable = True
    if limit_price is not None:
        if side == "BUY":
            marketable = q.ask <= limit_price + 1e-10
            execution = min(execution, limit_price)
        else:
            marketable = q.bid >= limit_price - 1e-10
            execution = max(execution, limit_price)
    total_slippage_bps = (abs(execution - q.mid) / q.mid * 10000.0) if q.mid > 0 else impact_bps
    return {
        "marketable": marketable,
        "fill_quantity": fill_qty if marketable else 0.0,
        "execution_price": round(float(execution), 6),
        "reference_price": q.mid,
        "slippage_bps": total_slippage_bps,
        "participation_pct": participation * 100.0,
        "liquidity_source": liquidity_source,
        "capacity": capacity,
        "reference_minute_volume": ref_minute_volume,
    }


def _apply_fill(conn: sqlite3.Connection, order: sqlite3.Row, q: Quote, qty: float, price: float, estimate: dict[str, Any]) -> dict[str, Any]:
    symbol = order["ticker"]
    side = order["side"]
    qty = float(qty)
    gross = price * qty
    now = _iso()
    account = _account_row(conn)
    cash = float(account["cash"])
    pos = conn.execute("SELECT * FROM portfolio_positions WHERE ticker=?", (symbol,)).fetchone()
    old_qty = float(pos["quantity"]) if pos else 0.0
    old_avg = float(pos["avg_cost"]) if pos else 0.0
    old_realized = float(pos["realized_pnl"]) if pos else 0.0
    realized = 0.0

    if side == "BUY":
        qty = min(qty, max(0.0, cash / price))
        if qty <= 1e-10:
            raise ValueError("Insufficient cash for any additional fill.")
        gross = price * qty
        new_qty = old_qty + qty
        new_avg = ((old_qty * old_avg) + gross) / new_qty
        new_cash = cash - gross
        new_realized = old_realized
    else:
        qty = min(qty, old_qty)
        if qty <= 1e-10:
            raise ValueError("No available shares for this sell fill.")
        gross = price * qty
        realized = (price - old_avg) * qty
        new_qty = max(0.0, old_qty - qty)
        new_avg = old_avg if new_qty > 1e-10 else 0.0
        new_cash = cash + gross
        new_realized = old_realized + realized

    conn.execute("UPDATE portfolio_account SET cash=?, updated_at=? WHERE id=1", (new_cash, now))
    if new_qty <= 1e-10:
        conn.execute("DELETE FROM portfolio_positions WHERE ticker=?", (symbol,))
    else:
        conn.execute(
            """
            INSERT INTO portfolio_positions(ticker, quantity, avg_cost, realized_pnl, last_price, last_price_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(ticker) DO UPDATE SET quantity=excluded.quantity, avg_cost=excluded.avg_cost,
                realized_pnl=excluded.realized_pnl, last_price=excluded.last_price,
                last_price_at=excluded.last_price_at, updated_at=excluded.updated_at
            """,
            (symbol, new_qty, new_avg, new_realized, q.price, q.quoted_at, now),
        )

    conn.execute(
        """
        INSERT INTO portfolio_transactions(
            executed_at, side, ticker, quantity, price, gross_amount, realized_pnl, cash_after, note,
            order_id, order_type, limit_price, reference_price, bid, ask, slippage_bps,
            participation_pct, liquidity_source
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (now, side, symbol, qty, price, gross, realized, new_cash, order["note"], order["id"],
         order["order_type"], order["limit_price"], estimate["reference_price"], q.bid, q.ask,
         estimate["slippage_bps"], estimate["participation_pct"], estimate["liquidity_source"]),
    )

    prior_filled = float(order["filled_quantity"] or 0.0)
    prior_avg = float(order["avg_fill_price"] or 0.0)
    total_filled = prior_filled + qty
    avg_fill = ((prior_filled * prior_avg) + qty * price) / total_filled
    remaining = max(0.0, float(order["quantity"]) - total_filled)
    status = "FILLED" if remaining <= 1e-8 else "PARTIAL"
    fill_bucket = _now_utc().strftime("%Y-%m-%dT%H:%MZ")
    conn.execute(
        "UPDATE portfolio_orders SET filled_quantity=?, avg_fill_price=?, status=?, updated_at=?, last_fill_bucket=? WHERE id=?",
        (total_filled, avg_fill, status, now, fill_bucket, order["id"]),
    )
    _record_snapshot(conn, force=True)
    return {
        "order_id": order["id"], "side": side, "ticker": symbol, "quantity": qty,
        "price": price, "gross_amount": gross, "realized_pnl": realized, "cash_after": new_cash,
        "executed_at": now, "slippage_bps": estimate["slippage_bps"],
        "participation_pct": estimate["participation_pct"], "liquidity_source": estimate["liquidity_source"],
        "bid": q.bid, "ask": q.ask, "reference_price": estimate["reference_price"],
        "remaining_quantity": remaining, "status": status,
    }


def _expire_day_orders(conn: sqlite3.Connection) -> None:
    now_utc = _now_utc()
    now_et = now_utc.astimezone(NY_TZ)
    today_et = now_et.date().isoformat()
    session_close_utc: datetime | None = None
    if xcals is not None:
        try:
            cal = xcals.get_calendar("XNYS")
            schedule = cal.schedule.loc[today_et:today_et]
            if not schedule.empty:
                close_ts = pd.Timestamp(schedule.iloc[0]["close"])
                session_close_utc = close_ts.to_pydatetime().astimezone(timezone.utc)
        except Exception:
            session_close_utc = None
    rows = conn.execute("SELECT id, created_at FROM portfolio_orders WHERE status IN ('OPEN','PARTIAL') AND tif='DAY'").fetchall()
    for row in rows:
        try:
            created_date = datetime.fromisoformat(row["created_at"]).astimezone(NY_TZ).date().isoformat()
            expired = created_date < today_et
            if created_date == today_et:
                if session_close_utc is not None:
                    expired = now_utc >= session_close_utc
                else:
                    expired = now_et.hour >= 16
            if expired:
                conn.execute("UPDATE portfolio_orders SET status='EXPIRED', cancel_reason='DAY order expired at session close', updated_at=? WHERE id=?", (_iso(), row["id"]))
        except Exception:
            continue


def process_open_orders() -> list[str]:
    init_portfolio_db()
    warnings: list[str] = []
    if not market_status().get("is_open"):
        with _DB_LOCK, _connect() as conn:
            _expire_day_orders(conn)
        return warnings

    with _DB_LOCK, _connect() as conn:
        _expire_day_orders(conn)
        rows = conn.execute("SELECT * FROM portfolio_orders WHERE status IN ('OPEN','PARTIAL') ORDER BY id").fetchall()
    current_bucket = _now_utc().strftime("%Y-%m-%dT%H:%MZ")
    for row in rows:
        try:
            if row["last_fill_bucket"] == current_bucket:
                continue
            q = fetch_yfinance_quote(row["ticker"])
            remaining = float(row["quantity"]) - float(row["filled_quantity"] or 0.0)
            if remaining <= 1e-8:
                continue
            est = _execution_estimate(row["side"], remaining, q, float(row["limit_price"]) if row["limit_price"] is not None else None)
            if not est["marketable"] or est["fill_quantity"] <= 0:
                continue
            with _DB_LOCK, _connect() as conn:
                conn.execute("BEGIN IMMEDIATE")
                try:
                    fresh = conn.execute("SELECT * FROM portfolio_orders WHERE id=?", (row["id"],)).fetchone()
                    if fresh is None or fresh["status"] not in ("OPEN", "PARTIAL"):
                        conn.execute("ROLLBACK")
                        continue
                    if fresh["last_fill_bucket"] == current_bucket:
                        conn.execute("ROLLBACK")
                        continue
                    remaining = float(fresh["quantity"]) - float(fresh["filled_quantity"] or 0.0)
                    qty = min(remaining, float(est["fill_quantity"]))
                    if fresh["side"] == "SELL":
                        pos = conn.execute("SELECT quantity FROM portfolio_positions WHERE ticker=?", (fresh["ticker"],)).fetchone()
                        held = float(pos["quantity"]) if pos else 0.0
                        available = max(0.0, held - _reserved_sell_qty(conn, fresh["ticker"], exclude_order_id=fresh["id"]))
                        qty = min(qty, available)
                    else:
                        cash = float(_account_row(conn)["cash"])
                        available_cash = max(0.0, cash - _reserved_cash(conn, exclude_order_id=fresh["id"]))
                        qty = min(qty, available_cash / max(float(est["execution_price"]), 1e-12))
                    if qty > 1e-10:
                        _apply_fill(conn, fresh, q, qty, float(est["execution_price"]), est)
                    conn.execute("COMMIT")
                except Exception:
                    conn.execute("ROLLBACK")
                    raise
        except Exception as exc:
            warnings.append(f"ORDER #{row['id']} {row['ticker']}: {exc}")
    return warnings


def submit_order(*, side: str, ticker: str, quantity: float, password: str, order_type: str = "MARKET", limit_price: float | None = None, tif: str = "DAY", note: str = "") -> dict[str, Any]:
    init_portfolio_db()
    if not verify_trade_password(password):
        raise PermissionError("Invalid portfolio trading password.")
    if not market_status().get("is_open"):
        raise RuntimeError("U.S. regular market is closed. New orders can be placed only during the regular session.")

    side = str(side or "").strip().upper()
    order_type = str(order_type or "MARKET").strip().upper()
    tif = str(tif or "DAY").strip().upper()
    if side not in {"BUY", "SELL"}:
        raise ValueError("Side must be BUY or SELL.")
    if order_type not in {"MARKET", "LIMIT"}:
        raise ValueError("Order type must be MARKET or LIMIT.")
    if tif not in {"DAY", "GTC"}:
        raise ValueError("TIF must be DAY or GTC.")
    stored_tif = "IOC" if order_type == "MARKET" else tif
    symbol = normalize_ticker(ticker)
    qty = float(quantity)
    if not math.isfinite(qty) or qty <= 0 or qty > 1_000_000_000:
        raise ValueError("Quantity must be greater than 0.")
    lp = None
    if order_type == "LIMIT":
        lp = _finite_positive(limit_price)
        if lp is None:
            raise ValueError("A positive limit price is required for LIMIT orders.")

    q = fetch_yfinance_quote(symbol, force=True)
    now = _iso()
    with _DB_LOCK, _connect() as conn:
        conn.execute("BEGIN IMMEDIATE")
        try:
            account = _account_row(conn)
            cash = float(account["cash"])
            if side == "BUY":
                reserved = _reserved_cash(conn)
                required = qty * (lp if lp is not None else q.ask * 1.03)
                if required > cash - reserved + 1e-8:
                    raise ValueError(f"Insufficient buying power. Est. required ${required:,.2f}; available ${(cash-reserved):,.2f} after open orders.")
            else:
                pos = conn.execute("SELECT quantity FROM portfolio_positions WHERE ticker=?", (symbol,)).fetchone()
                held = float(pos["quantity"]) if pos else 0.0
                available = max(0.0, held - _reserved_sell_qty(conn, symbol))
                if qty > available + 1e-8:
                    raise ValueError(f"Insufficient available shares. Order {qty:g}; available {available:g} after open sell orders.")
            cur = conn.execute(
                """
                INSERT INTO portfolio_orders(created_at, updated_at, side, ticker, order_type, quantity, limit_price, filled_quantity, avg_fill_price, status, tif, note)
                VALUES (?, ?, ?, ?, ?, ?, ?, 0, NULL, 'OPEN', ?, ?)
                """,
                (now, now, side, symbol, order_type, qty, lp, stored_tif, str(note or "")[:240]),
            )
            order_id = int(cur.lastrowid)
            order = conn.execute("SELECT * FROM portfolio_orders WHERE id=?", (order_id,)).fetchone()
            estimate = _execution_estimate(side, qty, q, lp)
            fill: dict[str, Any] | None = None
            if estimate["marketable"] and estimate["fill_quantity"] > 0:
                fill = _apply_fill(conn, order, q, float(estimate["fill_quantity"]), float(estimate["execution_price"]), estimate)
                order = conn.execute("SELECT * FROM portfolio_orders WHERE id=?", (order_id,)).fetchone()
            # Market orders are IOC in this simulator: an unfilled remainder is cancelled.
            if order_type == "MARKET" and order["status"] != "FILLED":
                reason = "Market IOC remainder cancelled after simulated liquidity cap"
                conn.execute("UPDATE portfolio_orders SET status='PARTIAL_CANCELLED', cancel_reason=?, updated_at=? WHERE id=?", (reason, _iso(), order_id))
            conn.execute("COMMIT")
        except Exception:
            conn.execute("ROLLBACK")
            raise

    with _DB_LOCK, _connect() as conn:
        order_out = dict(conn.execute("SELECT * FROM portfolio_orders WHERE id=?", (order_id,)).fetchone())
    return {
        "order": order_out, "fill": fill, "quote": quote_payload(q),
        "portfolio": get_portfolio_payload(refresh_quotes=False, process_orders=False),
    }


def cancel_order(*, order_id: int, password: str) -> dict[str, Any]:
    init_portfolio_db()
    if not verify_trade_password(password):
        raise PermissionError("Invalid portfolio trading password.")
    with _DB_LOCK, _connect() as conn:
        conn.execute("BEGIN IMMEDIATE")
        try:
            row = conn.execute("SELECT * FROM portfolio_orders WHERE id=?", (int(order_id),)).fetchone()
            if row is None:
                raise ValueError("Order not found.")
            if row["status"] not in ("OPEN", "PARTIAL"):
                raise ValueError(f"Order #{order_id} is already {row['status']}.")
            conn.execute("UPDATE portfolio_orders SET status='CANCELLED', cancel_reason='Cancelled by PM', updated_at=? WHERE id=?", (_iso(), int(order_id)))
            conn.execute("COMMIT")
        except Exception:
            conn.execute("ROLLBACK")
            raise
    return {"order_id": int(order_id), "status": "CANCELLED", "portfolio": get_portfolio_payload(refresh_quotes=False, process_orders=False)}


def refresh_position_quotes() -> list[str]:
    init_portfolio_db()
    warnings: list[str] = []
    with _DB_LOCK, _connect() as conn:
        symbols = [r["ticker"] for r in _position_rows(conn)]
    for symbol in symbols:
        try:
            q = fetch_yfinance_quote(symbol)
            with _DB_LOCK, _connect() as conn:
                conn.execute("UPDATE portfolio_positions SET last_price=?, last_price_at=?, updated_at=? WHERE ticker=?", (q.price, q.quoted_at, _iso(), symbol))
        except Exception as exc:
            warnings.append(f"{symbol}: {exc}")
    if market_status().get("is_open"):
        with _DB_LOCK, _connect() as conn:
            _record_snapshot(conn, force=False)
    return warnings


def get_portfolio_payload(*, refresh_quotes: bool = True, transaction_limit: int = 150, process_orders: bool = True) -> dict[str, Any]:
    init_portfolio_db()
    warnings: list[str] = []
    if process_orders:
        warnings.extend(process_open_orders())
    if refresh_quotes:
        warnings.extend(refresh_position_quotes())
    with _DB_LOCK, _connect() as conn:
        _expire_day_orders(conn)
        account = _account_row(conn)
        positions = _position_rows(conn)
        tx_rows = conn.execute(
            """
            SELECT id, executed_at, side, ticker, quantity, price, gross_amount, realized_pnl, cash_after, note,
                   order_id, order_type, limit_price, reference_price, bid, ask, slippage_bps, participation_pct, liquidity_source
            FROM portfolio_transactions ORDER BY id DESC LIMIT ?
            """, (max(1, min(int(transaction_limit), 500)),)
        ).fetchall()
        order_rows = conn.execute("SELECT * FROM portfolio_orders ORDER BY id DESC LIMIT 200").fetchall()
        history = conn.execute("SELECT captured_at, cash, positions_value, total_equity FROM portfolio_snapshots ORDER BY id DESC LIMIT 750").fetchall()
        total_realized = float(conn.execute("SELECT COALESCE(SUM(realized_pnl),0) AS total FROM portfolio_transactions").fetchone()["total"] or 0.0)
        reserved_cash = _reserved_cash(conn)

    position_payload: list[dict[str, Any]] = []
    positions_value = unrealized_pnl = 0.0
    largest_weight = 0.0
    for row in positions:
        qty, avg = float(row["quantity"]), float(row["avg_cost"])
        last = float(row["last_price"] or avg)
        value, cost = qty * last, qty * avg
        unrl = value - cost
        positions_value += value
        unrealized_pnl += unrl
        position_payload.append({
            "ticker": row["ticker"], "quantity": qty, "avg_cost": avg, "last_price": last,
            "last_price_at": row["last_price_at"], "market_value": value, "cost_basis": cost,
            "unrealized_pnl": unrl, "return_pct": (unrl / cost * 100.0) if cost > 0 else 0.0,
            "realized_pnl": float(row["realized_pnl"] or 0),
        })

    cash = float(account["cash"])
    total_equity = cash + positions_value
    initial_cash = float(account["initial_cash"])
    total_pnl = total_equity - initial_cash
    for item in position_payload:
        item["weight_pct"] = item["market_value"] / total_equity * 100.0 if total_equity > 0 else 0.0
        item["available_to_sell"] = item["quantity"]
        largest_weight = max(largest_weight, item["weight_pct"])
    # compute reserved sell quantities without leaking connections
    with _DB_LOCK, _connect() as conn:
        for item in position_payload:
            item["available_to_sell"] = max(0.0, item["quantity"] - _reserved_sell_qty(conn, item["ticker"]))

    open_orders = [dict(r) for r in order_rows if r["status"] in ("OPEN", "PARTIAL")]
    return {
        "account": {
            "scope": "MIDAS CAPITAL SHARED PORTFOLIO", "initial_cash": initial_cash, "cash": cash,
            "reserved_cash": reserved_cash, "buying_power": max(0.0, cash - reserved_cash),
            "positions_value": positions_value, "total_equity": total_equity, "total_pnl": total_pnl,
            "total_return_pct": (total_pnl / initial_cash * 100.0) if initial_cash > 0 else 0.0,
            "unrealized_pnl": unrealized_pnl, "realized_pnl": total_realized,
            "cash_pct": (cash / total_equity * 100.0) if total_equity > 0 else 0.0,
            "gross_exposure_pct": (positions_value / total_equity * 100.0) if total_equity > 0 else 0.0,
            "largest_position_pct": largest_weight, "open_order_count": len(open_orders),
            "created_at": account["created_at"], "updated_at": account["updated_at"],
        },
        "positions": position_payload,
        "orders": [dict(r) for r in order_rows],
        "open_orders": open_orders,
        "transactions": [dict(r) for r in tx_rows],
        "history": [dict(r) for r in reversed(history)],
        "market_status": market_status(), "warnings": warnings,
        "execution_model": {
            "max_participation_pct": MAX_PARTICIPATION_RATE * 100.0,
            "market_order_tif": "IOC",
            "slippage": "Bid/ask + square-root impact using recent 1-minute volume or ADV fallback",
            "partial_fill": "Per refresh cycle, simulated fill is capped at max participation of recent minute liquidity",
            "disclaimer": "Simulation only. Yahoo/yfinance is not Level II or exchange-certified real-time data.",
        },
        "quote_source": "Yahoo Finance / yfinance. Quotes can be delayed or throttled and are not exchange-certified real-time data.",
    }


# Backward-compatible v5.10 entry point: now routes through the execution engine.
def execute_trade(*, side: str, ticker: str, quantity: float, password: str, note: str = "") -> dict[str, Any]:
    result = submit_order(side=side, ticker=ticker, quantity=quantity, password=password, order_type="MARKET", note=note)
    fill = result.get("fill") or {}
    return {"trade": fill, "portfolio": result["portfolio"], "order": result["order"], "quote": result["quote"]}
