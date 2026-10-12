from __future__ import annotations

import asyncio
import html
import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

import aiohttp


US_NEWS_PER_KEYWORD = max(1, min(int(os.getenv("US_NEWS_PER_KEYWORD", "5")), 5))
US_NEWS_LOOKBACK_HOURS = max(1, min(int(os.getenv("US_NEWS_LOOKBACK_HOURS", "24")), 168))
US_GOOGLE_RESULTS_TO_SCAN = max(10, min(int(os.getenv("US_GOOGLE_RESULTS_TO_SCAN", "40")), 100))

ET_TZ = ZoneInfo("America/New_York")
GOOGLE_NEWS_RSS = "https://news.google.com/rss/search"

MAX_TOTAL_EMBED_CHARS = 5800
MAX_FIELD_VALUE_CHARS = 1000


US_CATEGORIES: dict[str, list[tuple[str, str]]] = {
    "💼 Buyside": [
        ("Private Equity", '"private equity" deal OR acquisition OR buyout'),
        ("PE Buyout", '"buyout" "private equity"'),
        ("VC", '"venture capital" funding OR investment'),
    ],
    "💰 Financing": [
        ("Equity Offering", '"public offering" OR "stock offering" company'),
        ("Financing", '"financing" company raises capital'),
        ("Convertible", '"convertible notes" OR "convertible debt" company'),
    ],
    "🤝 M&A": [
        ("M&A", 'merger OR acquisition "M&A"'),
        ("Acquisition", 'company acquisition deal'),
        ("Sale", 'company sale strategic review buyer'),
        ("Investment", 'strategic investment company funding round'),
    ],
    "📈 Equity Market": [
        ("IPO", 'IPO OR "initial public offering"'),
        ("Stocks", 'stocks shares surge plunge market movers'),
    ],
    "🏗️ Restructuring": [
        ("Restructuring", 'company restructuring layoffs bankruptcy turnaround'),
    ],
}


@dataclass(slots=True)
class NewsItem:
    title: str
    url: str
    source: str
    published_at: datetime
    keyword: str


class GoogleNewsError(RuntimeError):
    pass


def clean_text(value: str) -> str:
    return html.unescape(re.sub(r"<[^>]+>", "", value or "")).strip()


def safe_source(url: str) -> str:
    try:
        host = (urlparse(url).hostname or "").lower()
        if host.startswith("www."):
            host = host[4:]
        return host or "Google News"
    except Exception:
        return "Google News"


def parse_pub_date(value: str) -> datetime:
    dt = parsedate_to_datetime(value)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(ET_TZ)


def google_time_filter(hours: int) -> str:
    return "when:1d" if hours <= 24 else "when:7d"


def shorten(text: str, limit: int = 120) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[: limit - 1] + "…"


def normalize_title(title: str) -> str:
    title = re.sub(r"\s+", " ", title).strip().lower()
    return re.sub(r"[^\w]", "", title)


async def search_google_news(
    session: aiohttp.ClientSession,
    keyword: str,
    query: str,
    *,
    limit: int,
    lookback_hours: int,
) -> list[NewsItem]:
    params = {
        "q": f"{query} {google_time_filter(lookback_hours)}",
        "hl": "en-US",
        "gl": "US",
        "ceid": "US:en",
    }
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/151 Safari/537.36"
        )
    }

    try:
        async with session.get(
            GOOGLE_NEWS_RSS,
            params=params,
            headers=headers,
            timeout=aiohttp.ClientTimeout(total=15),
        ) as response:
            if response.status != 200:
                body = (await response.text())[:300]
                raise GoogleNewsError(f"Google News HTTP {response.status}: {body}")
            xml_text = await response.text()
    except asyncio.TimeoutError as exc:
        raise GoogleNewsError("Google News request timed out.") from exc
    except aiohttp.ClientError as exc:
        raise GoogleNewsError(f"Google News connection error: {exc}") from exc

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise GoogleNewsError("Could not parse Google News RSS response.") from exc

    cutoff = datetime.now(ET_TZ) - timedelta(hours=lookback_hours)
    results: list[NewsItem] = []
    seen_urls: set[str] = set()
    seen_titles: set[str] = set()

    for raw in root.findall("./channel/item")[:US_GOOGLE_RESULTS_TO_SCAN]:
        title = clean_text(raw.findtext("title") or "")
        url = (raw.findtext("link") or "").strip()
        pub_date = (raw.findtext("pubDate") or "").strip()
        source_node = raw.find("source")
        source = clean_text(
            source_node.text
            if source_node is not None and source_node.text
            else ""
        )

        if not title or not url or not pub_date:
            continue

        try:
            published_at = parse_pub_date(pub_date)
        except Exception:
            continue

        if published_at < cutoff:
            continue

        if source and title.endswith(f" - {source}"):
            title = title[: -(len(source) + 3)].strip()

        title_key = normalize_title(title)
        if url in seen_urls or title_key in seen_titles:
            continue

        seen_urls.add(url)
        seen_titles.add(title_key)
        results.append(
            NewsItem(
                title=title,
                url=url,
                source=source or safe_source(url),
                published_at=published_at,
                keyword=keyword,
            )
        )

        if len(results) >= limit:
            break

    return results


async def collect_us_news() -> dict[str, list[NewsItem]]:
    category_results: dict[str, list[NewsItem]] = {
        category: [] for category in US_CATEGORIES
    }

    connector = aiohttp.TCPConnector(limit=10)
    async with aiohttp.ClientSession(connector=connector) as session:
        semaphore = asyncio.Semaphore(6)

        async def search_one(
            category: str,
            keyword: str,
            query: str,
        ) -> tuple[str, list[NewsItem]]:
            async with semaphore:
                try:
                    items = await search_google_news(
                        session,
                        keyword,
                        query,
                        limit=US_NEWS_PER_KEYWORD,
                        lookback_hours=US_NEWS_LOOKBACK_HOURS,
                    )
                    return category, items
                except Exception as exc:
                    print(f"[US ERROR] {category} / {keyword}: {exc}")
                    return category, []

        tasks = [
            search_one(category, keyword, query)
            for category, searches in US_CATEGORIES.items()
            for keyword, query in searches
        ]
        collected = await asyncio.gather(*tasks)

    for category, items in collected:
        category_results[category].extend(items)

    for category, items in category_results.items():
        items.sort(key=lambda item: item.published_at, reverse=True)
        unique_items: list[NewsItem] = []
        seen_urls: set[str] = set()
        seen_titles: set[str] = set()

        for item in items:
            title_key = normalize_title(item.title)
            if item.url in seen_urls or title_key in seen_titles:
                continue
            seen_urls.add(item.url)
            seen_titles.add(title_key)
            unique_items.append(item)

        category_results[category] = unique_items

    return category_results


def build_article_entry(item: NewsItem) -> str:
    return (
        f"• **{item.keyword}** · "
        f"[{shorten(item.title, 95)}]({item.url}) "
        f"— {shorten(item.source, 24)}"
    )


def build_category_text(items: list[NewsItem], char_limit: int) -> str:
    if not items:
        return "No recent results"

    lines: list[str] = []
    for item in items:
        entry = build_article_entry(item)
        candidate = "\n".join(lines + [entry])
        if len(candidate) > char_limit:
            break
        lines.append(entry)

    return "\n".join(lines) if lines else "No displayable results"
