#!/usr/bin/env python3
"""Fetch latest PARP7 / TIPARP / RBN-2397 items from PubMed, ClinicalTrials.gov and
industry news (Google News RSS, which surfaces BioSpace and other outlets), update
data/updates.json, and regenerate the live block inside news.html.

Runs daily via .github/workflows/parp7-daily-update.yml.

No third-party deps — uses the stdlib so it works on any GitHub Actions runner.
"""
from __future__ import annotations

import json
import re
import sys
import html
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from urllib import request, parse
from urllib.error import URLError, HTTPError

ROOT = Path(__file__).resolve().parent.parent
DATA_FILE = ROOT / "data" / "updates.json"
HTML_FILE = ROOT / "news.html"

PUBMED_ESEARCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
PUBMED_ESUMMARY = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
CT_GOV_API = "https://clinicaltrials.gov/api/v2/studies"
GOOGLE_NEWS_RSS = "https://news.google.com/rss/search"

# Terms used across the different sources
PUBMED_TERM = "PARP7[Title/Abstract] OR TIPARP[Title/Abstract]"
CT_TERM = "PARP7 OR RBN-2397 OR TIPARP"
NEWS_QUERY = 'PARP7 OR RBN-2397 OR "TIPARP inhibitor"'

MARK_START = "<!-- LIVE-NEWS:START -->"
MARK_END = "<!-- LIVE-NEWS:END -->"

USER_AGENT = "PARP7-InfoSite-DailyBot/1.0 (+https://github.com/suguizhuan/sugz)"


def http_get_json(url: str, params: dict) -> dict:
    query = parse.urlencode(params)
    req = request.Request(f"{url}?{query}", headers={"User-Agent": USER_AGENT})
    with request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


def http_get_text(url: str, params: dict) -> str:
    query = parse.urlencode(params)
    req = request.Request(f"{url}?{query}", headers={"User-Agent": USER_AGENT})
    with request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8", errors="replace")


def fetch_pubmed(days: int = 3, max_items: int = 15) -> list[dict]:
    """PubMed articles mentioning PARP7 / TIPARP in title/abstract, recent `days`."""
    try:
        search = http_get_json(
            PUBMED_ESEARCH,
            {
                "db": "pubmed",
                "term": PUBMED_TERM,
                "reldate": str(days),
                "datetype": "pdat",
                "sort": "date",
                "retmax": str(max_items),
                "retmode": "json",
            },
        )
    except (URLError, HTTPError, TimeoutError, ValueError) as e:
        print(f"[pubmed] esearch failed: {e}", file=sys.stderr)
        return []

    ids = search.get("esearchresult", {}).get("idlist", [])
    if not ids:
        return []

    try:
        summary = http_get_json(
            PUBMED_ESUMMARY,
            {"db": "pubmed", "id": ",".join(ids), "retmode": "json"},
        )
    except (URLError, HTTPError, TimeoutError, ValueError) as e:
        print(f"[pubmed] esummary failed: {e}", file=sys.stderr)
        return []

    result = summary.get("result", {})
    items: list[dict] = []
    for pmid in ids:
        rec = result.get(pmid, {})
        if not rec:
            continue
        journal = rec.get("fulljournalname") or rec.get("source") or ""
        pubdate = rec.get("pubdate") or rec.get("epubdate") or ""
        items.append(
            {
                "id": f"pubmed-{pmid}",
                "date": pubdate,
                "category": "Paper",
                "title": rec.get("title", "").strip(),
                "summary": journal,
                "tags": ["pubmed"],
                "url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
                "pmid": pmid,
            }
        )
    return items


def fetch_clinicaltrials(max_items: int = 15) -> list[dict]:
    """Recently updated ClinicalTrials.gov studies mentioning PARP7 / RBN-2397."""
    try:
        params = {
            "query.term": CT_TERM,
            "pageSize": str(max_items),
            "sort": "LastUpdatePostDate:desc",
            "format": "json",
        }
        payload = http_get_json(CT_GOV_API, params)
    except (URLError, HTTPError, TimeoutError, ValueError) as e:
        print(f"[ct.gov] fetch failed: {e}", file=sys.stderr)
        return []

    items: list[dict] = []
    for study in payload.get("studies", []):
        proto = study.get("protocolSection", {})
        ident = proto.get("identificationModule", {})
        nct = ident.get("nctId")
        title = ident.get("briefTitle", "")
        status = proto.get("statusModule", {}).get("overallStatus", "")
        phase = ",".join(proto.get("designModule", {}).get("phases", [])) or "N/A"
        updated = (
            proto.get("statusModule", {})
            .get("lastUpdatePostDateStruct", {})
            .get("date", "")
        )
        if not nct:
            continue
        items.append(
            {
                "id": f"ct-{nct}",
                "date": updated,
                "category": "Trial",
                "title": title,
                "summary": f"Status: {status} · Phase: {phase}",
                "tags": ["clinicaltrials.gov", status.lower().replace(" ", "-")],
                "url": f"https://clinicaltrials.gov/study/{nct}",
                "nct": nct,
            }
        )
    return items


def fetch_news(max_items: int = 15) -> list[dict]:
    """Industry / press news via Google News RSS (covers BioSpace, Endpoints, etc.)."""
    try:
        xml_text = http_get_text(
            GOOGLE_NEWS_RSS,
            {"q": NEWS_QUERY, "hl": "en-US", "gl": "US", "ceid": "US:en"},
        )
    except (URLError, HTTPError, TimeoutError) as e:
        print(f"[news] fetch failed: {e}", file=sys.stderr)
        return []

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as e:
        print(f"[news] parse failed: {e}", file=sys.stderr)
        return []

    items: list[dict] = []
    for item in root.iterfind(".//item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        pub = (item.findtext("pubDate") or "").strip()
        source_el = item.find("source")
        source = (source_el.text or "").strip() if source_el is not None else ""
        if not title or not link:
            continue
        # Stable id from the link/title
        slug = re.sub(r"[^a-z0-9]+", "-", (source + "-" + title).lower())[:80].strip("-")
        items.append(
            {
                "id": f"news-{slug}",
                "date": pub,
                "category": "News",
                "title": title,
                "summary": source or "News",
                "tags": ["news", source.lower().replace(" ", "-")] if source else ["news"],
                "url": link,
                "source": source,
            }
        )
        if len(items) >= max_items:
            break
    return items


def merge_new(existing: list[dict], incoming: list[dict], max_keep: int = 50) -> list[dict]:
    """Prepend new items (dedup by id) and keep the most recent `max_keep`."""
    seen = {it["id"] for it in existing}
    new_items = [it for it in incoming if it["id"] not in seen]
    merged = new_items + existing
    return merged[:max_keep]


def update_data(featured_only: bool = False) -> dict:
    if DATA_FILE.exists():
        data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    else:
        data = {"featured": [], "news": [], "pubmed": [], "trials": []}

    if not featured_only:
        data["news"] = merge_new(data.get("news", []), fetch_news())
        data["pubmed"] = merge_new(data.get("pubmed", []), fetch_pubmed())
        data["trials"] = merge_new(data.get("trials", []), fetch_clinicaltrials())

    data["generated_at"] = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    DATA_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return data


def esc(s: str) -> str:
    return html.escape(s or "", quote=True)


def render_html_block(data: dict) -> str:
    now = data.get("generated_at", datetime.now(timezone.utc).isoformat())
    parts: list[str] = []
    parts.append(f'<p class="updated-at">最后更新：<time datetime="{esc(now)}">{esc(now)}</time></p>')

    # Featured (curated)
    featured = data.get("featured", [])
    if featured:
        parts.append('<h3 class="live-h3">重点关注</h3>')
        parts.append('<div class="card-grid">')
        for it in featured:
            tags = " ".join(
                f'<span class="badge indigo">{esc(t)}</span>' for t in it.get("tags", [])[:4]
            )
            parts.append(
                f'<div class="card">'
                f'<span class="badge">{esc(it.get("category",""))}</span> '
                f'<span style="color:var(--text-muted);font-size:12px;margin-left:6px;">{esc(it.get("date",""))}</span>'
                f'<h4 style="margin:8px 0 6px;font-size:16px;">{esc(it.get("title",""))}</h4>'
                f'<p style="font-size:14px;color:var(--text-muted);margin:0 0 10px;">{esc(it.get("summary",""))}</p>'
                f'<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px;">{tags}</div>'
                f'<a class="more" href="{esc(it.get("url","#"))}" target="_blank" rel="noopener">来源 →</a>'
                f'</div>'
            )
        parts.append('</div>')

    # News (Google News / BioSpace crawler)
    news = data.get("news", [])
    parts.append('<h3 class="live-h3">近期新闻（BioSpace / 爬虫）</h3>')
    if news:
        parts.append('<ul class="news-list">')
        for it in news[:15]:
            src = it.get("source") or it.get("summary") or "News"
            parts.append(
                f'<li class="news-item">'
                f'<a class="news-title" href="{esc(it.get("url","#"))}" target="_blank" rel="noopener">{esc(it.get("title",""))}</a>'
                f'<span class="news-meta"><span class="news-src">{esc(src)}</span> · {esc(it.get("date",""))}</span>'
                f'</li>'
            )
        parts.append('</ul>')
    else:
        parts.append('<p style="color:var(--text-muted);font-size:14px;">今日暂无新的新闻条目（脚本将在下次运行时继续尝试）。</p>')

    # Trials
    trials = data.get("trials", [])
    parts.append('<h3 class="live-h3">近期 ClinicalTrials.gov 更新</h3>')
    if trials:
        parts.append(
            '<div class="table-wrap"><table class="data"><thead><tr>'
            '<th>NCT</th><th>标题</th><th>状态/期</th><th>更新</th></tr></thead><tbody>'
        )
        for it in trials[:12]:
            parts.append(
                f'<tr><td><a href="{esc(it.get("url","#"))}" target="_blank" rel="noopener">{esc(it.get("nct",""))}</a></td>'
                f'<td>{esc(it.get("title",""))}</td>'
                f'<td>{esc(it.get("summary",""))}</td>'
                f'<td>{esc(it.get("date",""))}</td></tr>'
            )
        parts.append('</tbody></table></div>')
    else:
        parts.append('<p style="color:var(--text-muted);font-size:14px;">暂无匹配的临床试验更新。</p>')

    # PubMed papers
    pubmed = data.get("pubmed", [])
    parts.append('<h3 class="live-h3">近期 PubMed 文献</h3>')
    if pubmed:
        parts.append('<ul class="pubmed-list">')
        for it in pubmed[:15]:
            parts.append(
                f'<li><a href="{esc(it.get("url","#"))}" target="_blank" rel="noopener">{esc(it.get("title",""))}</a>'
                f' <span class="pm-meta">— {esc(it.get("summary",""))} · {esc(it.get("date",""))}</span></li>'
            )
        parts.append('</ul>')
    else:
        parts.append('<p style="color:var(--text-muted);font-size:14px;">今日暂无新的 PubMed 条目（脚本将在下次运行时继续尝试）。</p>')

    return "\n".join(parts)


def inject_html(block: str) -> None:
    if not HTML_FILE.exists():
        print(f"[html] news.html not found at {HTML_FILE}", file=sys.stderr)
        return
    text = HTML_FILE.read_text(encoding="utf-8")
    if MARK_START not in text or MARK_END not in text:
        print("[html] markers not found; leave file untouched", file=sys.stderr)
        return
    pattern = re.compile(
        re.escape(MARK_START) + r".*?" + re.escape(MARK_END),
        re.DOTALL,
    )
    replacement = f"{MARK_START}\n{block}\n{MARK_END}"
    new_text = pattern.sub(replacement, text)
    if new_text != text:
        HTML_FILE.write_text(new_text, encoding="utf-8")
        print(f"[html] news.html updated ({len(block)} chars in block)")


def main(argv: list[str]) -> int:
    featured_only = "--featured-only" in argv
    data = update_data(featured_only=featured_only)
    block = render_html_block(data)
    inject_html(block)
    print(
        f"OK: featured={len(data.get('featured',[]))} "
        f"news={len(data.get('news',[]))} "
        f"pubmed={len(data.get('pubmed',[]))} trials={len(data.get('trials',[]))}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
