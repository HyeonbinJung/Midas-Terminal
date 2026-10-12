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


NEWS_PER_KEYWORD = max(1, min(int(os.getenv("NEWS_PER_KEYWORD", "5")), 5))
NEWS_LOOKBACK_HOURS = max(1, min(int(os.getenv("NEWS_LOOKBACK_HOURS", "24")), 168))
GOOGLE_RESULTS_TO_SCAN = max(10, min(int(os.getenv("GOOGLE_RESULTS_TO_SCAN", "40")), 100))

KST = ZoneInfo("Asia/Seoul")
GOOGLE_NEWS_RSS = "https://news.google.com/rss/search"

# Discord 한 메시지의 모든 Embed 텍스트 합계 제한은 6000자.
# 안전 마진을 두고 5800자 이하로 맞춤.
MAX_TOTAL_EMBED_CHARS = 5800
MAX_FIELD_VALUE_CHARS = 1000


KR_CATEGORIES: dict[str, list[tuple[str, str]]] = {
    "💼 Buyside": [
        ("사모펀드", "사모펀드"),
        ("PE", "PE 사모펀드"),
        ("VC", "VC 벤처캐피탈 투자"),
    ],
    "💰 자금조달": [
        ("자금조달", "기업 자금조달"),
        ("유상증자", "유상증자"),
        ("CB발행", "CB 발행 전환사채"),
    ],
    "🤝 M&A": [
        ("M&A", "M&A 인수합병"),
        ("투자유치", "투자유치"),
        ("인수", "기업 인수"),
        ("매각", "기업 매각"),
    ],
    "📈 주식 시황": [
        ("IPO", "IPO 기업공개"),
        ("특징주", "특징주"),
    ],
    "🏗️ 구조조정": [
        ("구조조정", "기업 구조조정"),
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
    return dt.astimezone(KST)


def google_time_filter(hours: int) -> str:
    return "when:1d" if hours <= 24 else "when:7d"


def shorten(text: str, limit: int = 120) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[: limit - 1] + "…"


def normalize_title(title: str) -> str:
    title = re.sub(r"\s+", " ", title).strip().lower()
    return re.sub(r"[^\w가-힣]", "", title)


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
        "hl": "ko",
        "gl": "KR",
        "ceid": "KR:ko",
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
        raise GoogleNewsError("Google News 요청 시간이 초과되었습니다.") from exc
    except aiohttp.ClientError as exc:
        raise GoogleNewsError(f"Google News 연결 오류: {exc}") from exc

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise GoogleNewsError("Google News RSS 응답을 해석하지 못했습니다.") from exc

    cutoff = datetime.now(KST) - timedelta(hours=lookback_hours)
    results: list[NewsItem] = []
    seen_urls: set[str] = set()
    seen_titles: set[str] = set()

    for raw in root.findall("./channel/item")[:GOOGLE_RESULTS_TO_SCAN]:
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


async def collect_kr_news() -> dict[str, list[NewsItem]]:
    category_results: dict[str, list[NewsItem]] = {
        category: [] for category in KR_CATEGORIES
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
                        limit=NEWS_PER_KEYWORD,
                        lookback_hours=NEWS_LOOKBACK_HOURS,
                    )
                    return category, items
                except Exception as exc:
                    print(f"[KR ERROR] {category} / {keyword}: {exc}")
                    return category, []

        tasks = [
            search_one(category, keyword, query)
            for category, searches in KR_CATEGORIES.items()
            for keyword, query in searches
        ]
        collected = await asyncio.gather(*tasks)

    for category, items in collected:
        category_results[category].extend(items)

    # 카테고리 내부에서 중복 제거 후 최신순 정렬
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
    # 한 메시지에 최대한 많이 넣기 위해 한 줄 포맷 사용
    return (
        f"• **{item.keyword}** · "
        f"[{shorten(item.title, 95)}]({item.url}) "
        f"— {shorten(item.source, 24)}"
    )


def build_category_text(items: list[NewsItem], char_limit: int) -> str:
    if not items:
        return "최근 검색 결과 없음"

    lines: list[str] = []
    for item in items:
        entry = build_article_entry(item)
        candidate = "\n".join(lines + [entry])
        if len(candidate) > char_limit:
            break
        lines.append(entry)

    return "\n".join(lines) if lines else "표시 가능한 검색 결과 없음"
