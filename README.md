<div align="center">

# 🏛️ MIDAS Terminal

**Automated Equity Research & Financial Intelligence Platform**

<img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" height="26">
<img src="https://img.shields.io/badge/FastAPI-009688?style=for-the-badge&logo=fastapi&logoColor=white" height="26">
<img src="https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black" height="26">
<img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" height="26">
<img src="https://img.shields.io/badge/Plotly-3F4F75?style=for-the-badge&logo=plotly&logoColor=white" height="26">
<img src="https://img.shields.io/badge/SQLite-003B57?style=for-the-badge&logo=sqlite&logoColor=white" height="26">
<img src="https://img.shields.io/badge/SEC%20EDGAR-1F2937?style=for-the-badge" height="26">

[Overview](#-overview) · [Features](#-key-features) · [Tech Stack](#-tech-stack) · [Pipeline](#-research-pipeline) · [Getting Started](#-getting-started) · [Structure](#%EF%B8%8F-project-structure) · [Limitations](#-limitations)

</div>

---

## 📌 Overview

MIDAS Terminal is a full-stack financial research platform designed to streamline fundamental equity analysis by integrating market data, SEC filings, financial statements, comparable-company analysis, and automated research report generation.

Built with **Python, FastAPI, and JavaScript**, the system automates repetitive research workflows while preserving analyst control over investment theses, valuation assumptions, and final recommendations.

## ✨ Key Features

| Feature | Description |
|---|---|
| 📊 **Financial Market Terminal** | Ticker-based stock analysis · market capitalization, beta, price changes and valuation-related metrics · analyst price targets and market statistics |
| 🔎 **Automated Fundamental Research** | SEC EDGAR filing integration · automated company background research · income statement and balance sheet analysis · comparable-company selection · evidence-based investment highlights |
| 🧠 **Research Intelligence** | Rule-based catalyst identification · financial and business risk-factor screening · source-linked research evidence · U.S. and Korean financial news monitoring |
| 📝 **Automated Research Reports** | Structured equity research reports in Word format · English and Korean report generation · automated financial statement and research-section population · analyst-controlled DCF, investment thesis and recommendation sections |

## 🧰 Tech Stack

| Component | Technologies |
|---|---|
| Backend | Python, FastAPI, Starlette, Pydantic, Uvicorn |
| Frontend | JavaScript, HTML5, CSS3 |
| Charts & Visualization | Plotly.js (bundled locally) |
| Financial Data | Yahoo Finance, yfinance |
| Regulatory Data | SEC EDGAR |
| Data Processing | Pandas, NumPy |
| Persistence | SQLite (WAL mode) |
| Market Calendar | exchange-calendars (NYSE sessions, holidays, early closes) |
| Web Extraction | BeautifulSoup4, Requests, aiohttp |
| Document Generation | python-docx |
| Translation | googletrans |
| Configuration | python-dotenv |
| Performance | asyncio, TTL Caching, Semaphores, ThreadPoolExecutor |
| Deployment | Render, ngrok |

## 🔄 Research Pipeline

```text
Ticker → Market Data → SEC Filings → Financial Analysis → Peer Comparison → Research Signals → DOCX Export
```

## ⚙️ Engineering Highlights

MIDAS implements asynchronous data retrieval, single-flight TTL caching, concurrency controls, API rate limiting, and automated report-file cleanup.

These mechanisms reduce redundant upstream requests and support more reliable report generation in a single-instance deployment.

## 🚀 Getting Started

```powershell
python -m pip install -r requirements.txt
copy .env.example .env      # set SEC_USER_AGENT to your contact email
python main.py
```

Open **http://127.0.0.1:8000**. For development auto-reload:

```powershell
python -m uvicorn main:app --host 127.0.0.1 --port 8000 --reload
```

External access from a second terminal: `ngrok http 8000`. `render.yaml` is included for Render deployment.

## 🗂️ Project Structure

```text
main.py                     FastAPI app, API routes, caching, rate limiting, DOCX export
services/
  research_engine.py        SEC EDGAR enrichment, financials, peers, catalysts, risks
  stock_screener.py         Stock screener
  alpha_factors.py          Quarterly alpha-factor calculations
  portfolio.py              SQLite-backed virtual portfolio
  news_us.py / news_kr.py   U.S. and Korean news collection
  translation.py            Korean report translation (googletrans)
static/                     Terminal UI (index.html, app.js, styles.css, plotly.min.js)
templates/                  Word research-report template
data/schema.sql             Portfolio database schema
CHANGES_v*.md               Release notes per version
```

## 🎯 Project Objective

To make institutional-style equity research workflows more accessible, structured, and efficient through financial data engineering and research automation.

MIDAS is designed to assist fundamental analysts, not replace human investment judgment.

## ⚠️ Limitations

Research outputs depend on the availability and accuracy of public financial data. Automated catalysts and risk indicators are heuristic research aids rather than validated forecasts. The v5.1 implementation does not autonomously generate DCF valuations or investment recommendations.
