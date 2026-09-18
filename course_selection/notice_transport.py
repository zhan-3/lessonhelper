"""Network and browser transports for reading selection notices.

Every call in this module performs I/O.  Parsing and domain rules stay in
:mod:`course_selection.notice` and :mod:`course_selection.notice_discovery`, so
the application core can depend on those modules without gaining the ability to
open a socket or drive a browser.

The application core never imports this module.  It receives the two callables
it needs as injected dependencies, which keeps ``workbench_service`` testable
with a plain function instead of a patched transport.
"""

from __future__ import annotations

import re
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from .notice_discovery import (
    DEFAULT_NOTICE_INDEX_URL,
    approved_host,
    candidate_from_text,
    parse_official_notice_article,
    parse_official_notice_links,
)

_USER_AGENT = "academic-course-selection/0.1"


class _TextExtractor(HTMLParser):
    """Collect the visible text of an arbitrary HTML page."""

    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        value = data.strip()
        if value:
            self.parts.append(value)

    def text(self) -> str:
        return "\n".join(self.parts)


def _require_http_url(source_url: str) -> None:
    parsed = urlparse(source_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("通知链接必须是 HTTP 或 HTTPS 地址")


def fetch_notice_text(source_url: str, *, timeout_seconds: int = 10) -> str:
    """Read a public notice page over plain HTTP and return its visible text."""
    _require_http_url(source_url)
    request = Request(source_url, headers={"User-Agent": _USER_AGENT})
    with urlopen(request, timeout=timeout_seconds) as response:
        payload = response.read()
        charset = response.headers.get_content_charset() or "utf-8"
    parser = _TextExtractor()
    parser.feed(payload.decode(charset, errors="replace"))
    return parser.text()


def fetch_notice_text_in_browser(
    source_url: str,
    *,
    profile_root: Path = Path(".private/course-progress"),
    browser: str = "chromium",
    login_timeout_seconds: int = 600,
) -> str:
    """Read a notice in the visible, persistent academic browser session."""
    from playwright.sync_api import sync_playwright

    from course_progress.session import AcademicBrowserSession

    _require_http_url(source_url)
    if login_timeout_seconds <= 0:
        raise ValueError("认证等待时间必须大于 0")

    with sync_playwright() as playwright, AcademicBrowserSession(
        playwright,
        browser_name=browser,
        profile_root=profile_root,
        persistent=False,
    ) as session:
        page = session.open_authenticated(
            source_url, timeout_seconds=login_timeout_seconds
        )
        page.wait_for_timeout(500)
        text = page.locator("body").inner_text(timeout=10_000).strip()
        if not text:
            raise ValueError("通知页面没有读取到正文")
        return text


def _download_html(
    url: str, *, official_hosts: tuple[str, ...], timeout_seconds: int = 10
) -> str:
    if not approved_host(url, official_hosts):
        raise ValueError("notice source is not an approved official host")
    request = Request(url, headers={"User-Agent": _USER_AGENT})
    with urlopen(request, timeout=timeout_seconds) as response:
        final_url = response.geturl() if hasattr(response, "geturl") else url
        if not approved_host(final_url, official_hosts):
            raise ValueError(
                "notice download redirected outside the approved official host"
            )
        payload = response.read()
        charset = response.headers.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace")


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
