import argparse
import csv
import re
import sys
import time
import warnings
from email.utils import parseaddr
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import requests
from bs4 import BeautifulSoup

try:
    from ddgs import DDGS  # preferred package name
except Exception:
    try:
        from duckduckgo_search import DDGS  # backward compatibility
    except Exception:
        DDGS = None


DDG_HTML_URL = "https://html.duckduckgo.com/html/"
JINA_PROXY_PREFIX = "https://r.jina.ai/http://"
PROJECT_ROOT = Path(__file__).resolve().parents[1]
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)


def browser_headers() -> dict[str, str]:
    return {
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate",
        "Connection": "keep-alive",
    }


def parse_ddg_href(href: str) -> str | None:
    if not href:
        return None
    if href.startswith("//"):
        href = "https:" + href
    if href.startswith("http://") or href.startswith("https://"):
        parsed = urlparse(href)
        if "duckduckgo.com" in parsed.netloc and parsed.path.startswith("/l/"):
            qs = parse_qs(parsed.query)
            target = qs.get("uddg", [None])[0]
            if target:
                return unquote(target)
        return href
    return None


def unique_keep_order(urls: list[str], max_results: int) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for u in urls:
        if not u or not u.startswith("http"):
            continue
        low = u.lower()
        if "html.duckduckgo.com/html/?" in low:
            continue
        if "external-content.duckduckgo.com/" in low:
            continue
        if low.endswith(".ico"):
            continue
        if u in seen:
            continue
        seen.add(u)
        out.append(u)
        if len(out) >= max_results:
            break
    return out


def ddgs_search(query: str, max_results: int) -> list[str]:
    if DDGS is None:
        raise RuntimeError("duckduckgo-search is not installed")

    urls: list[str] = []
    with DDGS() as ddgs:
        for item in ddgs.text(query, region="wt-wt", safesearch="off", max_results=max_results):
            url = item.get("href") or item.get("url") or ""
            if url:
                urls.append(url)
    return unique_keep_order(urls, max_results)


def duckduckgo_html_search(session: requests.Session, query: str, max_results: int) -> list[str]:
    headers = browser_headers()
    headers["Referer"] = "https://html.duckduckgo.com/"

    resp = session.post(
        DDG_HTML_URL,
        data={"q": query, "b": ""},
        headers=headers,
        timeout=25,
    )
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")
    urls: list[str] = []

    for tag in soup.select("a.result__a[href], a[data-testid='result-title-a'][href], a.result-link[href]"):
        url = parse_ddg_href(tag.get("href", ""))
        if url:
            urls.append(url)

    return unique_keep_order(urls, max_results)


def jina_ddg_search(session: requests.Session, query: str, max_results: int) -> list[str]:
    """
    Remote fetch fallback: gets DDG HTML through r.jina.ai.
    Useful when local network/IP is blocked by search engines.
    """
    ddg_url = f"html.duckduckgo.com/html/?q={requests.utils.quote(query)}"
    proxy_url = JINA_PROXY_PREFIX + ddg_url
    resp = session.get(proxy_url, timeout=30)
    resp.raise_for_status()
    text = resp.text

    urls: list[str] = []
    # Capture markdown links from r.jina.ai response.
    for m in re.finditer(r"\]\((https?://[^\s)]+)\)", text):
        raw = m.group(1)
        url = parse_ddg_href(raw)
        if url:
            urls.append(url)

    return unique_keep_order(urls, max_results)


def is_profile_like_linkedin(url: str) -> bool:
    if "linkedin.com" not in url:
        return False
    blocked = ("/company/", "/school/", "/jobs/", "/feed/", "/search/")
    return "/in/" in url and not any(p in url for p in blocked)


def build_query(university: str, professor: str, *extra: str) -> str:
    parts: list[str] = []
    for p in (university, professor, *extra):
        t = p.strip() if p else ""
        if t:
            parts.append(t)
    return " ".join(parts)


def save_results_csv(
    output_path: Path,
    university: str,
    city: str,
    professor: str,
    keyword: str,
    search_query: str,
    engine: str,
    urls: list[str],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "university",
                "city",
                "professor",
                "keyword",
                "search_query",
                "engine",
                "rank",
                "url",
                "is_linkedin_profile",
            ],
        )
        writer.writeheader()
        for i, url in enumerate(urls, start=1):
            writer.writerow(
                {
                    "university": university,
                    "city": city,
                    "professor": professor,
                    "keyword": keyword,
                    "search_query": search_query,
                    "engine": engine,
                    "rank": i,
                    "url": url,
                    "is_linkedin_profile": is_profile_like_linkedin(url),
                }
            )


EMAIL_RE = re.compile(
    r"(?<![a-zA-Z0-9._%+-])([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})(?![a-zA-Z0-9._%+-])"
)
# International + common local formats (best-effort, not exhaustive).
PHONE_RE = re.compile(
    r"(?:\+?\d{1,4}[\s.-]?)?(?:\(?\d{2,4}\)?[\s.-]?)?\d{2,4}[\s.-]?\d{2,4}[\s.-]?\d{2,6}(?:[\s.-]?\d{2,4})?"
)


def _normalize_email(raw: str) -> str | None:
    raw = raw.strip().strip("<>\"'").lower()
    if raw.startswith("mailto:"):
        raw = raw[7:]
    _, addr = parseaddr(raw)
    addr = (addr or raw).strip().lower()
    if not addr or "@" not in addr:
        return None
    if addr.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".ico")):
        return None
    if any(x in addr for x in ("example.com", "test@", "yourname@", "email@", "name@")):
        return None
    return addr


def extract_emails(text: str, limit: int = 20) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for m in EMAIL_RE.finditer(text):
        norm = _normalize_email(m.group(1))
        if not norm or norm in seen:
            continue
        seen.add(norm)
        out.append(norm)
        if len(out) >= limit:
            break
    return out


def extract_phones(text: str, limit: int = 15) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for m in PHONE_RE.finditer(text):
        s = re.sub(r"\s+", " ", m.group(0)).strip()
        digits = re.sub(r"\D", "", s)
        if len(digits) < 9 or len(digits) > 16:
            continue
        if s in seen:
            continue
        seen.add(s)
        out.append(s)
        if len(out) >= limit:
            break
    return out


def jina_wrap_page_url(url: str) -> str:
    u = url.strip()
    if u.startswith("https://"):
        return "https://r.jina.ai/" + u
    if u.startswith("http://"):
        return "https://r.jina.ai/http://" + u[len("http://") :]
    return "https://r.jina.ai/http://" + u


def fetch_page_plaintext(
    session: requests.Session,
    url: str,
    *,
    via_jina: bool,
    max_bytes: int = 600_000,
    timeout: int = 25,
) -> tuple[str, str, int, str]:
    """
    Returns (title, plaintext, http_status, note).
    """
    fetch_url = jina_wrap_page_url(url) if via_jina else url
    try:
        resp = session.get(fetch_url, timeout=timeout, stream=True)
        status = resp.status_code
        chunks: list[bytes] = []
        total = 0
        for chunk in resp.iter_content(chunk_size=64_000):
            if not chunk:
                continue
            chunks.append(chunk)
            total += len(chunk)
            if total >= max_bytes:
                break
        raw = b"".join(chunks)
        text = raw.decode(resp.encoding or "utf-8", errors="replace")
    except requests.RequestException as exc:
        return "", "", 0, str(exc)

    if via_jina:
        title = ""
        m = re.search(r"^Title:\s*(.+)$", text, re.MULTILINE)
        if m:
            title = m.group(1).strip()
        m2 = re.search(r"(?:Markdown Content:)([\s\S]*)", text)
        body = m2.group(1) if m2 else text
        return title, body, status, ""
    soup = BeautifulSoup(text, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    title_tag = soup.find("title")
    title = title_tag.get_text(strip=True) if title_tag else ""
    body = soup.get_text("\n", strip=True)
    return title, body, status, ""


def enrichment_url_order(urls: list[str], university: str) -> list[str]:
    """Prefer likely university / academic pages before social noise."""

    def score(item: tuple[int, str]) -> tuple[int, int]:
        idx, u = item
        host = urlparse(u).netloc.lower()
        uni = re.sub(r"[^a-z0-9]+", " ", university.lower()).split()
        uni_tokens = [t for t in uni if len(t) >= 4][:6]
        s = 50
        if any(t in host for t in uni_tokens):
            s -= 20
        if any(h in host for h in (".edu", ".ac.", "edu.pk", "univ", "university")):
            s -= 15
        if "linkedin.com" in host:
            s += 5
        if "facebook.com" in host or "instagram.com" in host or "twitter.com" in host or "x.com" in host:
            s += 8
        return (s, idx)

    return [u for _, u in sorted(enumerate(urls), key=score)]


def default_enriched_path(output_path: Path) -> Path:
    return output_path.with_name(output_path.stem + "_enriched" + output_path.suffix)


def save_enriched_csv(
    output_path: Path,
    university: str,
    city: str,
    professor: str,
    keyword: str,
    search_query: str,
    engine: str,
    rows: list[dict[str, str]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "university",
        "city",
        "professor",
        "keyword",
        "search_query",
        "engine",
        "rank",
        "url",
        "is_linkedin_profile",
        "source",
        "page_title",
        "http_status",
        "fetch_note",
        "emails",
        "phones",
    ]
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


LEADER_ROLE_KEYWORDS = [
    "vice chancellor",
    "chancellor",
    "rector",
    "dean",
    "principal",
    "provost",
    "president",
    "director",
    "registrar",
    "head of department",
    "hod",
    "chairman",
    "chairperson",
]
LEADER_ROLE_RE = re.compile(
    r"\b(" + "|".join(re.escape(x) for x in sorted(LEADER_ROLE_KEYWORDS, key=len, reverse=True)) + r")\b",
    re.IGNORECASE,
)
NAME_RE = re.compile(r"\b[A-Z][A-Za-z'`.-]+(?:\s+[A-Z][A-Za-z'`.-]+){1,4}\b")
ROLE_ALIASES: dict[str, str] = {
    "vc": "vice chancellor",
    "vice-chancellor": "vice chancellor",
    "head": "head of department",
}
NAME_PREFIX_RE = re.compile(
    r"^(?:(?:Dr|Prof|Professor|Mr|Ms|Mrs|Miss|Cdre|Rear|Admiral|Vice|Brig|Lt|Gen)\.?\s+)+",
    re.IGNORECASE,
)
BAD_NAME_TOKENS = {
    "university",
    "faculty",
    "department",
    "office",
    "contact",
    "school",
    "admissions",
    "campus",
    "leadership",
    "marketing",
    "management",
    "information",
    "technology",
    "computer",
    "science",
    "health",
    "sciences",
    "academic",
    "director",
    "rector",
    "dean",
    "religious",
    "affairs",
    "appointments",
    "deanery",
    "registrar",
    "chancellor",
    "principal",
    "controller",
    "software",
    "development",
    "college",
    "law",
    "faculty",
    "king",
    "santo",
    "vice",
    "pro",
    "advisory",
    "utilization",
    "regarding",
    "announcement",
    "notice",
    "update",
    "policy",
    "application",
    "admission",
    "medal",
    "holders",
    "holder",
    "detail",
    "list",
    "lists",
    "awardees",
}
BAD_NAME_PHRASES = (
    "the rector",
    "the registrar",
    "the chancellor",
    "the vice-chancellor",
    "the vice chancellor",
    "the pro-vice-chancellor",
    "the pro vice chancellor",
    "new deanery appointments",
    "religious affairs",
    "assistant principal",
    "software development",
    "canon law",
    "king's college",
    "kings college",
    "medal holders",
    "more detail",
    "rector's list",
    "dean's list",
    "scholarship awardees",
)


def _normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def _is_name_like(candidate: str, university: str) -> bool:
    c = _normalize_space(candidate)
    c = NAME_PREFIX_RE.sub("", c).strip()
    if len(c) < 5:
        return False
    if len(c.split()) < 2:
        return False
    low = c.lower()
    if low.startswith("the "):
        return False
    if any(p in low for p in BAD_NAME_PHRASES):
        return False
    words = re.sub(r"[^a-z0-9]+", " ", low).split()
    if all(w in BAD_NAME_TOKENS for w in words):
        return False
    if any(w in BAD_NAME_TOKENS for w in words):
        return False
    uni_tokens = [t for t in re.sub(r"[^a-z0-9]+", " ", university.lower()).split() if len(t) >= 4]
    if any(t in low for t in uni_tokens):
        return False
    return True


def normalize_role_label(role: str) -> str:
    r = _normalize_space(role).lower()
    return ROLE_ALIASES.get(r, r)


def role_matches(requested: str, discovered: str) -> bool:
    req = normalize_role_label(requested)
    got = normalize_role_label(discovered)
    if req == got:
        return True
    # allow specific broader role queries such as "dean" matching "dean of ..."
    if got.startswith(req + " "):
        return True
    return False


def clean_name_candidate(candidate: str) -> str:
    c = _normalize_space(candidate)
    c = NAME_PREFIX_RE.sub("", c).strip()
    # trim trailing short honorific remnants
    c = re.sub(r"\b(?:SI|HI|HI\(M\)|M)\b\.?$", "", c, flags=re.IGNORECASE).strip()
    return c


def get_university_preset(university: str, city: str) -> dict[str, object]:
    """
    Optional curated hints for universities with known site structure.
    Returns:
      {
        "domain": str,
        "seed_urls": list[str],
      }
    """
    u = _normalize_space(university).lower()
    c = _normalize_space(city).lower()

    if "fast" in u or "nuces" in u:
        base = "https://www.nu.edu.pk"
        campus_map = {
            "karachi": "Karachi",
            "lahore": "Lahore",
            "islamabad": "Islamabad",
            "peshawar": "Peshawar",
            "chiniot faisalabad": "Chiniot-Faisalabad",
            "faisalabad": "Chiniot-Faisalabad",
            "chiniot": "Chiniot-Faisalabad",
            "multan": "Multan",
        }
        campus = None
        for k, v in campus_map.items():
            if k in c:
                campus = v
                break

        seed_urls = [base]
        if campus == "Multan":
            seed_urls.extend(
                [
                    f"{base}/MultanCampus/Index",
                    f"{base}/MultanCampus/Faculty",
                    f"{base}/MultanCampus/ContactUs",
                ]
            )
        elif campus:
            seed_urls.extend(
                [
                    f"{base}/Campus/{campus}",
                    f"{base}/Campus/{campus}/Faculty",
                    f"{base}/Campus/{campus}/ContactUs",
                    f"{base}/Campus/{campus}/RectorLists",
                    f"{base}/Campus/{campus}/DeanLists",
                ]
            )

        return {"domain": "nu.edu.pk", "seed_urls": seed_urls}

    return {"domain": "", "seed_urls": []}


def extract_leadership_mentions(text: str, university: str, max_mentions: int = 120) -> list[tuple[str, str, str]]:
    """
    Extract (role, name, snippet) from page plaintext.
    This is heuristic and best-effort.
    """
    out: list[tuple[str, str, str]] = []
    lines = text.splitlines()
    for i, raw in enumerate(lines):
        line = _normalize_space(raw)
        if not line or len(line) < 10 or len(line) > 220:
            continue
        role_m = LEADER_ROLE_RE.search(line)
        if not role_m:
            continue
        role = role_m.group(1).lower()
        # Name can be on same line or neighboring lines (common in table/card layouts).
        start = max(0, i - 2)
        end = min(len(lines), i + 3)
        context_lines = [_normalize_space(x) for x in lines[start:end] if _normalize_space(x)]
        context_text = " | ".join(context_lines)
        low_ctx = context_text.lower()
        if any(
            x in low_ctx
            for x in (
                "rector's list",
                "dean's list",
                "medal holders",
                "scholarship awardees",
                "more detail",
            )
        ):
            continue
        names = NAME_RE.findall(context_text)
        picked = ""
        for n in names:
            cleaned = clean_name_candidate(n)
            if _is_name_like(cleaned, university):
                picked = cleaned
                break
        if not picked:
            continue
        out.append((role, picked, context_text[:220]))
        if len(out) >= max_mentions:
            break
    return out


def save_leaders_csv(output_path: Path, university: str, city: str, leaders: list[dict[str, str]]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "university",
                "city",
                "role",
                "name",
                "mentions",
                "current_score",
                "source_urls",
                "snippets",
            ],
        )
        writer.writeheader()
        for row in leaders:
            writer.writerow(
                {
                    "university": university,
                    "city": city,
                    "role": row["role"],
                    "name": row["name"],
                    "mentions": row["mentions"],
                    "current_score": row.get("current_score", "0"),
                    "source_urls": row["source_urls"],
                    "snippets": row["snippets"],
                }
            )


def discover_university_leaders(
    session: requests.Session,
    university: str,
    engine: str,
    *,
    city: str = "",
    official_domain: str = "",
    uni_website: str = "",
    max_search_results: int = 15,
    max_pages: int = 8,
    via_jina: bool = False,
    fetch_delay: float = 0.8,
    strict_current: bool = False,
) -> tuple[list[dict[str, str]], str]:
    """
    Discover likely leadership names for a university from search results + page scraping.
    Returns (leaders, engine_used_for_search).
    """
    city_clean = _normalize_space(city)
    preset = get_university_preset(university, city_clean)
    preset_domain = str(preset.get("domain", "") or "").strip().lower()
    preset_seed_urls = [str(x) for x in list(preset.get("seed_urls", []))]

    discover_query = f"{university} {city_clean} rector dean vice chancellor leadership administration".strip()
    if preset_domain:
        discover_query = f"{discover_query} site:{preset_domain}".strip()
    urls, engine_used = search_urls(session, discover_query, max_search_results, engine)

    def host(u: str) -> str:
        return urlparse(u).netloc.lower().replace("www.", "")

    uni_tokens = [t for t in re.sub(r"[^a-z0-9]+", " ", university.lower()).split() if len(t) >= 4]

    def domain_score(netloc: str) -> int:
        s = 0
        if any(t in netloc for t in uni_tokens):
            s += 4
        if any(t in netloc for t in ("edu", "ac.", "university", "nu.edu.pk", "fast")):
            s += 3
        if "linkedin." in netloc or "facebook." in netloc or "twitter." in netloc or "wikipedia." in netloc:
            s -= 4
        return s

    official_domain = official_domain.strip().lower().replace("www.", "")
    if not official_domain and preset_domain:
        official_domain = preset_domain
    if urls:
        best_url = max(urls, key=lambda u: domain_score(host(u)))
        if (not official_domain) and domain_score(host(best_url)) >= 3:
            official_domain = host(best_url)

    university_l = university.lower()
    domain_hints = {
        "fast": "nu.edu.pk",
        "bahria": "bahria.edu.pk",
        "comsats": "comsats.edu.pk",
        "nust": "nust.edu.pk",
    }
    if not official_domain:
        for key, dom in domain_hints.items():
            if key in university_l:
                official_domain = dom
                break

    if official_domain:
        site_query = f"site:{official_domain} rector dean vice chancellor registrar office directory"
        scoped_urls, scoped_engine = search_urls(session, site_query, max_search_results, engine)
        if scoped_urls:
            urls = scoped_urls
            engine_used = scoped_engine

    if official_domain:
        urls = [u for u in urls if host(u) == official_domain or host(u).endswith("." + official_domain)] or urls
    ordered = enrichment_url_order(urls, university)

    targets: list[str] = []
    seen: set[str] = set()

    def norm(u: str) -> str:
        p = urlparse(u)
        return f"{p.scheme}://{p.netloc}{p.path}".lower().rstrip("/")

    if uni_website.strip():
        u = uni_website.strip()
        seen.add(norm(u))
        targets.append(u)

    for u in preset_seed_urls:
        n = norm(u)
        if n in seen:
            continue
        seen.add(n)
        targets.append(u)

    for u in ordered:
        if len(targets) >= max_pages + (1 if uni_website.strip() else 0):
            break
        n = norm(u)
        if n in seen:
            continue
        seen.add(n)
        targets.append(u)

    historical_terms = (
        "former ",
        "past ",
        "history",
        "historical",
        "announced",
        "appointment",
        "appointments",
        "news",
        "event",
        "activity",
        "sympathizes",
        "launches",
        "conference",
        "seminar",
        "workshop",
    )
    current_terms = (
        "office directory",
        "contact",
        "rector@",
        "registrar@",
        "dean@",
        "pro-rector",
        "vice chancellor",
        "registrar",
        "dean",
        "rector",
        "director",
    )

    def current_score(snippet: str, source_url: str) -> int:
        s = 0
        low = snippet.lower()
        host_path = f"{urlparse(source_url).netloc}{urlparse(source_url).path}".lower()

        if any(t in low for t in historical_terms):
            s -= 4
        if "retired" in low:
            s -= 1

        if any(t in low for t in current_terms):
            s += 3
        if "@" in snippet and (".edu" in snippet.lower() or ".ac." in snippet.lower()):
            s += 3
        if "office-directory" in host_path or "office directory" in low:
            s += 4
        if "about" in host_path or "administration" in host_path or "leadership" in host_path:
            s += 1
        if any(tok in host_path for tok in uni_tokens):
            s += 2
        if official_domain and (host_path.startswith(official_domain) or f".{official_domain}" in host_path):
            s += 3
        return s

    grouped: dict[tuple[str, str], dict[str, object]] = {}
    for u in targets:
        time.sleep(max(fetch_delay, 0))
        title, body, status, note = fetch_page_plaintext(session, u, via_jina=via_jina)
        if status != 200 or not body:
            continue
        mentions = extract_leadership_mentions(body, university)
        for role, name, snippet in mentions:
            key = (role, name)
            if key not in grouped:
                grouped[key] = {
                    "role": role,
                    "name": name,
                    "mentions": 0,
                    "current_score": 0,
                    "sources": set(),
                    "snippets": [],
                }
            row = grouped[key]
            row["mentions"] = int(row["mentions"]) + 1
            row["current_score"] = int(row["current_score"]) + current_score(snippet, u)
            cast_sources = row["sources"]
            if isinstance(cast_sources, set):
                cast_sources.add(u)
            cast_snippets = row["snippets"]
            if isinstance(cast_snippets, list) and len(cast_snippets) < 3:
                cast_snippets.append(snippet)

    leaders: list[dict[str, str]] = []
    for _, row in grouped.items():
        sources = sorted(row["sources"]) if isinstance(row["sources"], set) else []
        snippets = row["snippets"] if isinstance(row["snippets"], list) else []
        leaders.append(
            {
                "role": str(row["role"]).title(),
                "name": str(row["name"]),
                "mentions": str(row["mentions"]),
                "current_score": str(row.get("current_score", 0)),
                "source_urls": " | ".join(sources),
                "snippets": " || ".join(snippets),
            }
        )

    if city_clean:
        ct = city_clean.lower()
        city_filtered: list[dict[str, str]] = []
        for r in leaders:
            hay = " ".join([r.get("role", ""), r.get("name", ""), r.get("source_urls", ""), r.get("snippets", "")]).lower()
            if ct in hay:
                city_filtered.append(r)
        if city_filtered:
            leaders = city_filtered

    # Keep likely current names first; drop clearly historical/noisy entries.
    min_score = 5 if strict_current else 2
    leaders = [r for r in leaders if int(r.get("current_score", "0")) >= min_score]
    if strict_current:
        strict_terms = ("office directory", "contact", "administration", "registrar@", "rector@", "dean@")
        strict_filtered: list[dict[str, str]] = []
        for r in leaders:
            hay = " ".join([r.get("source_urls", ""), r.get("snippets", "")]).lower()
            if any(t in hay for t in strict_terms):
                strict_filtered.append(r)
        if strict_filtered:
            leaders = strict_filtered
    leaders.sort(key=lambda r: (-int(r["current_score"]), -int(r["mentions"]), r["role"], r["name"]))
    return leaders, engine_used


def search_urls(session: requests.Session, query: str, max_results: int, engine: str) -> tuple[list[str], str]:
    max_results = max(1, max_results)

    if engine == "ddgs":
        return ddgs_search(query, max_results), "ddgs"

    if engine == "duckduckgo_html":
        return duckduckgo_html_search(session, query, max_results), "duckduckgo_html"

    if engine == "jina_ddg":
        return jina_ddg_search(session, query, max_results), "jina_ddg"

    # auto
    for name in ("ddgs", "duckduckgo_html", "jina_ddg"):
        try:
            if name == "ddgs":
                urls = ddgs_search(query, max_results)
            elif name == "jina_ddg":
                urls = jina_ddg_search(session, query, max_results)
            else:
                urls = duckduckgo_html_search(session, query, max_results)
        except Exception:
            continue
        if urls:
            return urls, name
    return [], ""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Free LinkedIn URL finder from search results (DDGS + DDG HTML fallback)."
    )
    parser.add_argument("-q", "--query", help="Full search query (overrides name fields)")
    parser.add_argument("-u", "--university", help="University name")
    parser.add_argument("--city", help="Campus city filter, e.g. Karachi, Islamabad, Lahore")
    parser.add_argument(
        "--official-domain",
        help="Restrict discovery to official university domain, e.g. nu.edu.pk, bahria.edu.pk",
    )
    parser.add_argument("-p", "--professor", help="Professor/person name")
    parser.add_argument("--keyword", default="linkedin", help='Single extra term (default: "linkedin")')
    parser.add_argument("-k", "--keywords", nargs="+", metavar="TERM", help="Multiple extra terms")
    parser.add_argument("--max-results", type=int, default=10, help="Maximum URLs (default: 10)")
    parser.add_argument("--delay-seconds", type=float, default=1.0, help="Delay before request (default: 1.0)")
    parser.add_argument("--output", default="output/google_linkedin_results.csv", help="CSV output path")
    parser.add_argument(
        "--engine",
        choices=("auto", "ddgs", "duckduckgo_html", "jina_ddg"),
        default="auto",
        help="auto = DDGS -> DDG HTML -> Jina DDG mirror (default: auto)",
    )
    parser.add_argument(
        "--enrich",
        action="store_true",
        help="Fetch result pages and extract emails/phones; keeps --output unchanged and also writes an enriched CSV",
    )
    parser.add_argument(
        "--uni-website",
        dest="uni_website",
        help="Optional university/staff page URL to fetch first (in addition to search URLs)",
    )
    parser.add_argument(
        "--enrich-max-pages",
        type=int,
        default=8,
        help="Max search-result URLs to fetch for contact scraping (default: 8). --uni-website is extra.",
    )
    parser.add_argument(
        "--enrich-fetch-delay",
        type=float,
        default=1.0,
        help="Pause between page fetches in seconds (default: 1.0)",
    )
    parser.add_argument(
        "--enrich-via-jina",
        action="store_true",
        help="Fetch each page through r.jina.ai (slower; helps when sites block your IP)",
    )
    parser.add_argument(
        "--enriched-output",
        help="Path for enriched CSV (default: <output_stem>_enriched.csv next to --output)",
    )
    parser.add_argument(
        "--discover-leaders",
        action="store_true",
        help="Discover leadership names (rector/dean/etc.) from university pages and save a leaders CSV",
    )
    parser.add_argument(
        "--role",
        help="Leadership role to auto-select discovered name, e.g. rector, dean, vice chancellor",
    )
    parser.add_argument(
        "--leaders-output",
        default="output/university_leaders.csv",
        help="Output CSV for discovered leadership names (default: output/university_leaders.csv)",
    )
    parser.add_argument(
        "--discover-max-pages",
        type=int,
        default=8,
        help="Max pages to scrape for leader discovery (default: 8)",
    )
    parser.add_argument(
        "--strict-current",
        action="store_true",
        help="Stricter filter for current office holders (drops weak/noisy matches)",
    )
    return parser.parse_args()


def prompt_if_missing(value: str | None, label: str) -> str:
    if value and value.strip():
        return value.strip()
    return input(f"Enter {label}: ").strip()


def prompt_optional(label: str) -> str:
    if not sys.stdin.isatty():
        return ""
    return input(f"Enter {label} (press Enter to skip): ").strip()


def resolve_output_path(path_str: str) -> Path:
    p = Path(path_str)
    if p.is_absolute():
        return p
    return PROJECT_ROOT / p


def select_leader_interactively(leaders: list[dict[str, str]], requested_role: str = "") -> tuple[str, str]:
    """
    Returns (selected_name, selected_role).
    - If requested_role is provided, pick first matching leader automatically.
    - Otherwise prompt user to choose by number / role / full name.
    """
    if not leaders:
        return "", ""

    if requested_role:
        matches = [r for r in leaders if role_matches(requested_role, r["role"])]
        if matches:
            return matches[0]["name"], matches[0]["role"]
        return "", ""

    if not sys.stdin.isatty():
        # Non-interactive context: best-ranked leader.
        return leaders[0]["name"], leaders[0]["role"]

    print("\nChoose one leader to search URLs for:")
    for idx, row in enumerate(leaders[:20], start=1):
        print(f"{idx:02d}. {row['role']} -> {row['name']} (mentions: {row['mentions']})")
    print("Type number (e.g. 1), role (e.g. rector), or exact name.")
    choice = input("Your choice: ").strip()

    if not choice:
        return leaders[0]["name"], leaders[0]["role"]

    if choice.isdigit():
        i = int(choice)
        if 1 <= i <= min(len(leaders), 20):
            row = leaders[i - 1]
            return row["name"], row["role"]

    role_matches_found = [r for r in leaders if role_matches(choice, r["role"])]
    if role_matches_found:
        row = role_matches_found[0]
        return row["name"], row["role"]

    # Fallback: treat as typed name
    return choice, "manual"


def main() -> int:
    args = parse_args()
    # Keep terminal output clean for users.
    warnings.filterwarnings(
        "ignore",
        category=RuntimeWarning,
        message=r"This package \(`duckduckgo_search`\) has been renamed to `ddgs`!.*",
    )

    if DDGS is None and args.engine in ("auto", "ddgs"):
        print("Missing dependency: duckduckgo-search")
        print("Install it with: pip install duckduckgo-search")
        print("You can still run with: --engine duckduckgo_html")
        if args.engine == "ddgs":
            return 1

    session = requests.Session()
    session.headers.update(browser_headers())

    role_requested = (args.role or "").strip().lower()
    city_requested = _normalize_space(args.city or "")
    official_domain_requested = (args.official_domain or "").strip().lower().replace("www.", "")
    university_for_discovery = (args.university or "").strip()
    discovered_leaders: list[dict[str, str]] = []
    discovered_engine = ""

    auto_discover_needed = (not args.query) and (not (args.professor or "").strip())
    should_discover = args.discover_leaders or role_requested or auto_discover_needed

    if should_discover:
        if not university_for_discovery:
            university_for_discovery = prompt_if_missing(args.university, "university name")
        if not city_requested:
            city_requested = _normalize_space(prompt_optional("city name"))
        print(f"Discovering leaders for: {university_for_discovery}")
        try:
            discovered_leaders, discovered_engine = discover_university_leaders(
                session,
                university_for_discovery,
                args.engine,
                city=city_requested,
                official_domain=official_domain_requested,
                uni_website=(args.uni_website or ""),
                max_search_results=max(12, args.max_results),
                max_pages=max(1, args.discover_max_pages),
                via_jina=args.enrich_via_jina,
                fetch_delay=max(args.enrich_fetch_delay, 0),
                strict_current=args.strict_current,
            )
        except requests.RequestException as exc:
            print(f"Leader discovery network error: {exc}")
            return 1

        leaders_path = resolve_output_path(args.leaders_output)
        save_leaders_csv(leaders_path, university_for_discovery, city_requested, discovered_leaders)
        print(f"Leaders CSV saved to: {leaders_path.resolve()}")
        if discovered_leaders:
            print("\nTop discovered leaders:")
            for idx, row in enumerate(discovered_leaders[:12], start=1):
                print(
                    f"{idx:02d}. {row['role']} -> {row['name']} "
                    f"(current_score: {row.get('current_score', '0')}, mentions: {row['mentions']})"
                )
        else:
            print("No leaders discovered (try --uni-website with direct staff page).")
            if sys.stdin.isatty():
                if not official_domain_requested:
                    maybe_domain = prompt_optional("official university domain (e.g. nu.edu.pk)")
                    if maybe_domain:
                        official_domain_requested = maybe_domain.lower().replace("www.", "").strip()
                        try:
                            discovered_leaders, discovered_engine = discover_university_leaders(
                                session,
                                university_for_discovery,
                                args.engine,
                                city=city_requested,
                                official_domain=official_domain_requested,
                                uni_website=(args.uni_website or ""),
                                max_search_results=max(12, args.max_results),
                                max_pages=max(1, args.discover_max_pages),
                                via_jina=args.enrich_via_jina,
                                fetch_delay=max(args.enrich_fetch_delay, 0),
                                strict_current=args.strict_current,
                            )
                        except requests.RequestException as exc:
                            print(f"Leader discovery retry (domain) network error: {exc}")
                            return 1
                        save_leaders_csv(leaders_path, university_for_discovery, city_requested, discovered_leaders)
                        print(f"Leaders CSV updated: {leaders_path.resolve()}")
                        if discovered_leaders:
                            print("\nTop discovered leaders (domain retry):")
                            for idx, row in enumerate(discovered_leaders[:12], start=1):
                                print(
                                    f"{idx:02d}. {row['role']} -> {row['name']} "
                                    f"(current_score: {row.get('current_score', '0')}, mentions: {row['mentions']})"
                                )
                retry_site = ""
                if not discovered_leaders:
                    retry_site = prompt_optional("official leadership/office URL")
                if retry_site:
                    try:
                        discovered_leaders, discovered_engine = discover_university_leaders(
                            session,
                            university_for_discovery,
                            args.engine,
                            city=city_requested,
                            official_domain=official_domain_requested,
                            uni_website=retry_site,
                            max_search_results=max(12, args.max_results),
                            max_pages=max(1, args.discover_max_pages),
                            via_jina=args.enrich_via_jina,
                            fetch_delay=max(args.enrich_fetch_delay, 0),
                            strict_current=args.strict_current,
                        )
                    except requests.RequestException as exc:
                        print(f"Leader discovery retry network error: {exc}")
                        return 1
                    save_leaders_csv(leaders_path, university_for_discovery, city_requested, discovered_leaders)
                    print(f"Leaders CSV updated: {leaders_path.resolve()}")
                    if discovered_leaders:
                        print("\nTop discovered leaders (retry):")
                        for idx, row in enumerate(discovered_leaders[:12], start=1):
                            print(
                                f"{idx:02d}. {row['role']} -> {row['name']} "
                                f"(current_score: {row.get('current_score', '0')}, mentions: {row['mentions']})"
                            )
            if auto_discover_needed and not (args.professor or "").strip():
                if discovered_leaders:
                    pass
                else:
                    print("Stopping here because no leader names were found automatically.")
                    print("Tip: add --city and/or --uni-website to improve discovery.")
                    return 2

    if args.query and args.query.strip():
        search_query = args.query.strip()
        keyword = ""
        university = (args.university or "N/A").strip()
        professor = (args.professor or "N/A").strip()
    else:
        university = university_for_discovery or prompt_if_missing(args.university, "university name")
        professor = (args.professor or "").strip()
        if not professor and discovered_leaders:
            selected_name, selected_role = select_leader_interactively(discovered_leaders, role_requested)
            if selected_name:
                professor = selected_name
                if role_requested:
                    print(f"Using discovered name for role '{role_requested}': {professor}")
                else:
                    print(f"Selected leader: {professor} ({selected_role})")
            elif role_requested:
                available = ", ".join(sorted({r["role"] for r in discovered_leaders}))
                print(f"No discovered match for role '{role_requested}'. Available roles: {available}")
                return 2
        elif not professor and should_discover and not discovered_leaders:
            print("Could not find leadership names automatically.")
            print("Try one of these:")
            print("- provide city with --city")
            print("- provide official page with --uni-website (office directory / administration page)")
            print("- increase --discover-max-pages")
            return 2

        if not professor:
            professor = prompt_if_missing(args.professor, "professor name")
        if args.keywords:
            keyword = " ".join([t.strip() for t in args.keywords if t and t.strip()])
        else:
            keyword = (args.keyword or "linkedin").strip() or "linkedin"
        search_query = build_query(university, city_requested, professor, keyword)

    print(f"Searching query: {search_query}")
    print(f"Engine mode: {args.engine}")
    time.sleep(max(args.delay_seconds, 0))

    try:
        urls, engine_used = search_urls(session, search_query, args.max_results, args.engine)
    except requests.RequestException as exc:
        print(f"Network error: {exc}")
        return 1
    except RuntimeError as exc:
        print(str(exc))
        return 1

    if not urls:
        print("No URLs found. Try:")
        print("- add stronger keyword: -k \"site:linkedin.com/in\" linkedin")
        print("- use full query: -q \"Asif Khaliq Bahria University site:linkedin.com/in\"")
        print("- run later or from another network")
        return 2

    linkedin_profiles = [u for u in urls if is_profile_like_linkedin(u)]

    print(f"\nResults from: {engine_used}")
    print(f"Top {len(urls)} URLs:\n")
    for i, url in enumerate(urls, start=1):
        print(f"{i:02d}. {url}")

    print("\nLinkedIn profile candidates:\n")
    if linkedin_profiles:
        for i, url in enumerate(linkedin_profiles, start=1):
            print(f"{i:02d}. {url}")
    else:
        print("No direct linkedin.com/in/... link in this result set.")

    out_path = resolve_output_path(args.output)
    save_results_csv(
        output_path=out_path,
        university=university,
        city=city_requested,
        professor=professor,
        keyword=keyword,
        search_query=search_query,
        engine=engine_used,
        urls=urls,
    )
    print(f"\nSaved results to: {out_path.resolve()}")

    if not args.enrich and sys.stdin.isatty():
        more = input("Do you want more info (emails/phones) for this selected person? [y/N]: ").strip().lower()
        if more in ("y", "yes"):
            args.enrich = True

    if args.enrich:
        rank_by_url = {u: str(i) for i, u in enumerate(urls, start=1)}
        targets: list[tuple[str, str, str]] = []
        seen_u: set[str] = set()

        def norm(u: str) -> str:
            p = urlparse(u)
            return f"{p.scheme}://{p.netloc}{p.path}".lower().rstrip("/")

        uw = (args.uni_website or "").strip()
        if uw:
            targets.append((uw, "user_uni_site", "0"))
            seen_u.add(norm(uw))

        fetched_search = 0
        for u in enrichment_url_order(urls, university):
            if fetched_search >= max(1, args.enrich_max_pages):
                break
            nu = norm(u)
            if nu in seen_u:
                continue
            seen_u.add(nu)
            targets.append((u, "search_result", rank_by_url[u]))
            fetched_search += 1

        enriched_rows: list[dict[str, str]] = []
        for url, source, rank in targets:
            time.sleep(max(args.enrich_fetch_delay, 0))
            title, body, status, note = fetch_page_plaintext(
                session, url, via_jina=args.enrich_via_jina
            )
            emails = extract_emails(body)
            phones = extract_phones(body)
            if "linkedin.com" in urlparse(url).netloc.lower() and status == 200 and len(body) < 800:
                note = (note + "; " if note else "") + "linkedin_page_often_login_wall_for_scrapers"

            enriched_rows.append(
                {
                    "university": university,
                    "city": city_requested,
                    "professor": professor,
                    "keyword": keyword,
                    "search_query": search_query,
                    "engine": engine_used,
                    "rank": rank,
                    "url": url,
                    "is_linkedin_profile": str(is_profile_like_linkedin(url)),
                    "source": source,
                    "page_title": title.replace("\n", " ").strip()[:500],
                    "http_status": str(status),
                    "fetch_note": note[:500],
                    "emails": "; ".join(emails),
                    "phones": "; ".join(phones),
                }
            )

        enriched_path = resolve_output_path(args.enriched_output) if args.enriched_output else default_enriched_path(out_path)
        save_enriched_csv(
            output_path=enriched_path,
            university=university,
            city=city_requested,
            professor=professor,
            keyword=keyword,
            search_query=search_query,
            engine=engine_used,
            rows=enriched_rows,
        )
        all_emails = sorted({e for r in enriched_rows for e in r["emails"].split("; ") if e})
        all_phones = sorted({p for r in enriched_rows for p in r["phones"].split("; ") if p})
        print(f"\nEnriched CSV saved to: {enriched_path.resolve()}")
        if all_emails:
            print(f"Unique emails found ({len(all_emails)}): {', '.join(all_emails[:10])}{' ...' if len(all_emails) > 10 else ''}")
        else:
            print("No emails extracted (pages may block bots or use images for contact info).")
        if all_phones:
            print(f"Unique phones found ({len(all_phones)}): {', '.join(all_phones[:10])}{' ...' if len(all_phones) > 10 else ''}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
