#!/usr/bin/env python3
"""Fetch latest TLR7/8 antagonist items from PubMed and ClinicalTrials.gov,
update data/updates.json, and regenerate the "最新进展" block inside research.html.

Runs daily via .github/workflows/tlr78-daily-update.yml.

No third-party deps — uses the stdlib so it works on any GitHub Actions runner.
"""
from __future__ import annotations

import json
import re
import sys
import html
from datetime import datetime, timezone
from pathlib import Path
from urllib import request, parse
from urllib.error import URLError, HTTPError

ROOT = Path(__file__).resolve().parent.parent
DATA_FILE = ROOT / "data" / "updates.json"
HTML_FILE = ROOT / "research.html"

PUBMED_ESEARCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
PUBMED_ESUMMARY = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
CT_GOV_API = "https://clinicaltrials.gov/api/v2/studies"

MARK_START = "<!-- LATEST-UPDATES:START -->"
MARK_END = "<!-- LATEST-UPDATES:END -->"

# Focused query: TLR7 or TLR8 together with antagonism / inhibition / autoimmune context.
PUBMED_TERM = (
    "((TLR7[Title/Abstract] OR TLR8[Title/Abstract]) "
    "AND (antagonist[Title/Abstract] OR inhibitor[Title/Abstract] OR "
    "inhibition[Title/Abstract] OR lupus[Title/Abstract]))"
)
CT_TERM = "TLR7 antagonist OR TLR8 antagonist OR enpatoran OR afimetoran OR MHV370 OR E-6742"

USER_AGENT = "TLR78-InfoSite-DailyBot/1.0 (+https://github.com/suguizhuan/sugz)"


def http_get_json(url: str, params: dict) -> dict:
    query = parse.urlencode(params)
    req = request.Request(f"{url}?{query}", headers={"User-Agent": USER_AGENT})
    with request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


def fetch_pubmed(days: int = 3, max_items: int = 15) -> list[dict]:
    """Return PubMed articles about TLR7/8 antagonism, updated in the last `days`."""
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
    except (URLError, HTTPError, TimeoutError) as e:
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
    except (URLError, HTTPError, TimeoutError) as e:
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


def fetch_clinicaltrials(max_items: int = 10) -> list[dict]:
    """Return recently updated ClinicalTrials.gov studies on TLR7/8 antagonists."""
    try:
        params = {
            "query.term": CT_TERM,
            "pageSize": str(max_items),
            "sort": "LastUpdatePostDate:desc",
            "format": "json",
        }
        req = request.Request(
            f"{CT_GOV_API}?{parse.urlencode(params)}",
            headers={"User-Agent": USER_AGENT},
        )
        with request.urlopen(req, timeout=30) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except (URLError, HTTPError, TimeoutError) as e:
        print(f"[ct.gov] fetch failed: {e}", file=sys.stderr)
        return []

    items: list[dict] = []
    for study in payload.get("studies", []):
        proto = study.get("protocolSection", {})
        nct = proto.get("identificationModule", {}).get("nctId")
        title = proto.get("identificationModule", {}).get("briefTitle", "")
        status = proto.get("statusModule", {}).get("overallStatus", "")
        phase = ",".join(proto.get("designModule", {}).get("phases", [])) or "N/A"
        updated = proto.get("statusModule", {}).get("lastUpdatePostDateStruct", {}).get("date", "")
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


def merge_new(existing: list[dict], incoming: list[dict], max_keep: int = 40) -> list[dict]:
    """Prepend new items (dedup by id) and keep the most recent `max_keep`."""
    seen = {it["id"] for it in existing}
    new_items = [it for it in incoming if it["id"] not in seen]
    merged = new_items + existing
    return merged[:max_keep]


def update_data(featured_only: bool = False) -> dict:
    if DATA_FILE.exists():
        data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    else:
        data = {"featured": [], "pubmed": [], "trials": []}

    if not featured_only:
        pubmed = fetch_pubmed()
        trials = fetch_clinicaltrials()
        data["pubmed"] = merge_new(data.get("pubmed", []), pubmed)
        data["trials"] = merge_new(data.get("trials", []), trials)

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
            tags = " ".join(f'<span class="badge indigo">{esc(t)}</span>' for t in it.get("tags", [])[:4])
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

    # Trials
    trials = data.get("trials", [])
    if trials:
        parts.append('<h3 class="live-h3">近期 ClinicalTrials.gov 更新</h3>')
        parts.append('<div class="table-wrap"><table class="data"><thead><tr>'
                     '<th>NCT</th><th>标题</th><th>状态/期</th><th>更新</th></tr></thead><tbody>')
        for it in trials[:10]:
            parts.append(
                f'<tr><td><a href="{esc(it.get("url","#"))}" target="_blank" rel="noopener">{esc(it.get("nct",""))}</a></td>'
                f'<td>{esc(it.get("title",""))}</td>'
                f'<td>{esc(it.get("summary",""))}</td>'
                f'<td>{esc(it.get("date",""))}</td></tr>'
            )
        parts.append('</tbody></table></div>')

    # PubMed papers
    pubmed = data.get("pubmed", [])
    if pubmed:
        parts.append('<h3 class="live-h3">近期 PubMed 文献</h3>')
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
        print(f"[html] research.html not found at {HTML_FILE}", file=sys.stderr)
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
        print(f"[html] research.html updated ({len(block)} chars in block)")


def main(argv: list[str]) -> int:
    featured_only = "--featured-only" in argv
    data = update_data(featured_only=featured_only)
    block = render_html_block(data)
    inject_html(block)
    print(
        f"OK: featured={len(data.get('featured',[]))} "
        f"pubmed={len(data.get('pubmed',[]))} trials={len(data.get('trials',[]))}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
