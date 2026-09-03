# Midas Capital Research Terminal — Web v5

Midas-branded FastAPI research terminal for market data, SEC filing enrichment, peer comparison, news clipping and automated Word research-template population.

## v5 highlights
- F2 Research now loads **Key Financials** when `RUN` is pressed:
  - Income Statement FY2023A-FY2027E
  - Balance Sheet FY2023A-FY2027E
  - Negative values shown in red
  - Revenue Growth >= configured threshold highlighted
  - Mobile-safe internal horizontal scrolling
- Automatic Company Background, Highlights, Catalysts and 6 Risk Factor rows.
- Automatic Peer 1 / Peer 2 / Peer 3 selection and peer financial comparison.
- English or Korean DOCX export.
- Korean export uses `googletrans.Translator` for **all automatically populated narrative research text**, including long SEC filing excerpts and full risk monitoring paragraphs; fixed template labels in the populated sections are also localized.
- DCF, Recommendation and analyst-authored Thesis remain analyst-owned.

## Reliability / multi-user protections
- UUID DOCX filenames
- Browser export-click lock
- POST report endpoint
- Bounded report generation (`REPORT_CONCURRENCY`, default 2)
- Bounded Yahoo concurrency (`YAHOO_CONCURRENCY`, default 6)
- Same-ticker export serialization
- Market / Research / Peer / SEC / News TTL caches
- Per-key single-flight cache for simultaneous identical requests
- Global SEC request pacing
- Per-IP rate limiting, including ngrok forwarded addresses
- Generated DOCX cleanup after delivery and stale-file cleanup at startup
- Research `PARTIAL` warnings when upstream enrichment is incomplete

## Install / run on Windows
```powershell
cd "path\to\midas_terminal_web_v5"
python -m pip install -r requirements.txt
copy .env.example .env
python main.py
```

Open:
```text
http://127.0.0.1:8000
```

Development auto-reload:
```powershell
python -m uvicorn main:app --host 127.0.0.1 --port 8000 --reload
```

External ngrok access, from a second PowerShell window:
```powershell
ngrok http 8000
```

## SEC User-Agent
Edit `.env` and replace the example email:
```env
SEC_USER_AGENT=MidasCapitalResearch/5.0 your_real_email@example.com
```
This is an SEC request-identification header, not an API key.

## Korean translation
`requirements.txt` includes:
```text
googletrans==4.0.2
```
The report selector on F2 supports `ENGLISH` and `한국어`. Korean export sends automatically populated English research prose through googletrans, then creates the DOCX. Numbers, tickers, formulas and valuation ranges are retained rather than translated.

Long filing excerpts may take longer than English export. Default translation timeout is 180 seconds and can be changed in `.env`:
```env
TRANSLATION_TIMEOUT_SECONDS=300
```

## Main API routes
- `GET /api/market/AAPL`
- `GET /api/research/AAPL` — includes `key_financials`
- `GET /api/news/KR`
- `GET /api/news/US`
- `POST /api/report/AAPL?language=en`
- `POST /api/report/AAPL?language=ko`
- `GET /api/health`

## Default rate limits
- Market: 60 / minute / IP
- Research: 10 / minute / IP
- News: 20 / minute / IP
- Export: 5 / 10 minutes / IP

This remains a single-machine architecture when run on a PC + ngrok. For a much larger user base, move caches/rate limits to Redis and DOCX generation to worker processes.


## v5.1 hotfix
If F2 Key Financials previously did not appear, v5.1 disables browser/static caching and version-busts the frontend assets. Restart the server and reload the page normally; a hard refresh should no longer be required.

Korean export no longer aborts because one Google Translate request returns unchanged English. The service retries smaller chunks and uses a best-effort Google fallback before allowing export to continue.
