"""Strict, immutable discovery of official selection-arrangement notices."""

from __future__ import annotations

import hashlib
import re
from dataclasses import asdict, dataclass
from difflib import unified_diff
from html.parser import HTMLParser
from typing import ClassVar
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

from .notice import parse_notice

DEFAULT_NOTICE_INDEX_URL = "https://jwc.hitwh.edu.cn/ks/list.htm"
ARRANGEMENT_MARKERS = ("选课时间安排", "各类课程选课时间", "课程选课安排")
ARRANGEMENT_TITLE_RE = re.compile(
    r"^(?:关于)?20\d{2}年(?:春季|夏季|秋季|冬季)学期.*(?:各类课程选课时间安排|各类课程选课安排).*通知$"
)


@dataclass(frozen=True)
class OfficialNoticeLink:
    title: str
    url: str


def _approved_host(url: str, official_hosts: tuple[str, ...]) -> bool:
    host = (urlparse(url).hostname or "").lower()
    return host in {item.lower() for item in official_hosts}


def _download_html(
    url: str, *, official_hosts: tuple[str, ...], timeout_seconds: int = 10
) -> str:
    if not _approved_host(url, official_hosts):
        raise ValueError("notice source is not an approved official host")
    request = Request(url, headers={"User-Agent": "academic-course-selection/0.1"})
    with urlopen(request, timeout=timeout_seconds) as response:
        final_url = response.geturl() if hasattr(response, "geturl") else url
        if not _approved_host(final_url, official_hosts):
            raise ValueError(
                "notice download redirected outside the approved official host"
            )
        payload = response.read()
        charset = response.headers.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace")


class _NoticeListParser(HTMLParser):
    def __init__(self, base_url: str) -> None:
        super().__init__()
        self.base_url = base_url
        self.links: list[OfficialNoticeLink] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        attributes = dict(attrs)
        title = (attributes.get("title") or "").strip()
        href = (attributes.get("href") or "").strip()
        if href and ARRANGEMENT_TITLE_RE.fullmatch(title):
            self.links.append(
                OfficialNoticeLink(title=title, url=urljoin(self.base_url, href))
            )


class _ArticleTextParser(HTMLParser):
    BLOCK_TAGS: ClassVar[frozenset[str]] = frozenset(
        {"br", "p", "div", "tr", "td", "th", "li", "h1", "h2", "h3"}
    )
    VOID_TAGS: ClassVar[frozenset[str]] = frozenset(
        {
            "area",
            "base",
            "br",
            "col",
            "embed",
            "hr",
            "img",
            "input",
            "link",
            "meta",
            "param",
            "source",
            "track",
            "wbr",
        }
    )

    def __init__(self) -> None:
        super().__init__()
        self.depth = 0
        self.parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        classes = (dict(attrs).get("class") or "").split()
        if self.depth:
            if tag.lower() in self.BLOCK_TAGS:
                self.parts.append("\n")
            if tag.lower() not in self.VOID_TAGS:
                self.depth += 1
        elif "wp_articlecontent" in classes:
            self.depth = 1

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if self.depth and tag.lower() in self.BLOCK_TAGS:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if not self.depth:
            return
        if tag.lower() in self.BLOCK_TAGS:
            self.parts.append("\n")
        self.depth -= 1

    def handle_data(self, data: str) -> None:
        if self.depth:
            self.parts.append(data)

    def text(self) -> str:
        lines = [re.sub(r"\s+", "", line) for line in "".join(self.parts).splitlines()]
        return "\n".join(line for line in lines if line)


def parse_official_notice_links(
    index_html: str, *, index_url: str
) -> tuple[OfficialNoticeLink, ...]:
    parser = _NoticeListParser(index_url)
    parser.feed(index_html)
    unique: dict[str, OfficialNoticeLink] = {}
    for link in parser.links:
        unique.setdefault(link.url, link)
    return tuple(unique.values())


def parse_official_notice_article(article_html: str, *, title: str) -> str:
    parser = _ArticleTextParser()
    parser.feed(article_html)
    body = parser.text()
    if not body:
        raise ValueError("official notice article has no readable content")
    return f"{title}\n{body}"


def _notice_index_pages(index_url: str) -> tuple[str, str]:
    parsed = urlparse(index_url)
    second_path = re.sub(r"/list(?:1)?\.htm$", "/list2.htm", parsed.path)
    if second_path == parsed.path:
        return index_url, index_url
    return index_url, parsed._replace(path=second_path).geturl()


def discover_official_notice_candidates(
    index_url: str = DEFAULT_NOTICE_INDEX_URL,
    *,
    official_hosts: tuple[str, ...] = ("jwc.hitwh.edu.cn",),
    timeout_seconds: int = 10,
) -> list[dict]:
    """Find the newest matching arrangement notice on the first two list pages."""
    for page_url in _notice_index_pages(index_url):
        index_html = _download_html(
            page_url, official_hosts=official_hosts, timeout_seconds=timeout_seconds
        )
        links = parse_official_notice_links(index_html, index_url=page_url)
        if not links:
            continue
        link = links[0]
        article_html = _download_html(
            link.url, official_hosts=official_hosts, timeout_seconds=timeout_seconds
        )
        text = parse_official_notice_article(article_html, title=link.title)
        return [candidate_from_text(link.url, text, official_hosts=official_hosts)]
    return []


def candidate_from_text(
    source_url: str, text: str, *, official_hosts: tuple[str, ...]
) -> dict:
    host = (urlparse(source_url).hostname or "").lower()
    if host not in {item.lower() for item in official_hosts}:
        raise ValueError("notice source is not an approved official host")
    notice = parse_notice(text, source_url=source_url, source_kind="official")
    if not any(marker in notice.title for marker in ARRANGEMENT_MARKERS):
        raise ValueError("article is not a course-selection time arrangement")
    payload = asdict(notice)
    payload["content_hash"] = hashlib.sha256(text.encode("utf-8")).hexdigest()
    payload["version_id"] = hashlib.sha256(f"{source_url}\n{text}".encode()).hexdigest()
    payload["status"] = "candidate"
    payload["query_eligible"] = not notice.missing_fields and bool(notice.windows)
    return payload


def notice_diff(previous: dict, candidate: dict) -> str:
    return "\n".join(
        unified_diff(
            previous.get("source_text", "").splitlines(),
            candidate.get("source_text", "").splitlines(),
            fromfile="confirmed",
            tofile="candidate",
            lineterm="",
        )
    )
