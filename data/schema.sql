-- MIDAS Capital shared virtual trading schema (v5.11)
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
    note TEXT NOT NULL DEFAULT '',
    order_id INTEGER,
    order_type TEXT NOT NULL DEFAULT 'MARKET',
    limit_price REAL,
    reference_price REAL,
    bid REAL,
    ask REAL,
    slippage_bps REAL NOT NULL DEFAULT 0,
    participation_pct REAL,
    liquidity_source TEXT NOT NULL DEFAULT ''
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
