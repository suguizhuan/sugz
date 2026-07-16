#!/usr/bin/env python3
"""每日自动抓取康百达在研项目各靶点的增量动态，写入 data/monitor.json 的 `auto` 区。

- 临床（clinical）来源：ClinicalTrials.gov API v2（按 LastUpdatePostDate 倒序）
- 文章（article）来源：PubMed E-utilities（按发布日期倒序 + reldate 时间窗）

设计要点：
- 仅使用标准库，任何 GitHub Actions runner 均可运行。
- 幂等：按 id 去重合并，不会重复写入；每个靶点保留最近 N 条。
- 自动条目写入 data["auto"][<target_id>]，与人工策展的 `updates` 分开，
  前端渲染时合并展示并标注“自动抓取”。
- 单个来源失败不影响其它来源与靶点。
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib import request, parse
from urllib.error import URLError, HTTPError

ROOT = Path(__file__).resolve().parent.parent
DATA_FILE = ROOT / "data" / "monitor.json"

PUBMED_ESEARCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
PUBMED_ESUMMARY = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
CT_GOV_API = "https://clinicaltrials.gov/api/v2/studies"

USER_AGENT = "KBD-TargetMonitor-DailyBot/1.0 (+https://github.com/suguizhuan/sugz)"

MAX_KEEP_PER_TARGET = 30      # 每个靶点每来源保留的最近条数上限
PUBMED_RELDATE_DAYS = 3       # PubMed 时间窗（天）
PUBMED_MAX = 8
CT_MAX = 8

# 每个靶点的检索关键词（ClinicalTrials.gov 的 query.term 与 PubMed 的 term）
SEARCH = {
    "KBD1921": {"ct": "PARP7", "pubmed": "PARP7"},
    "KBD4409": {"ct": "TLR7 OR TLR8 OR enpatoran OR afimetoran",
                 "pubmed": "(TLR7/8 inhibitor) OR enpatoran OR afimetoran"},
    "KBD2430": {"ct": "PARP1 selective inhibitor", "pubmed": "PARP1 selective inhibitor brain"},
    "KBD2188": {"ct": "GPC3 radioligand OR GPC3 radioconjugate OR GPC3",
                 "pubmed": "GPC3 (radioligand OR radioconjugate OR radiopharmaceutical)"},
    "KBD1183": {"ct": "apelin receptor agonist OR APJ agonist",
                 "pubmed": "apelin receptor agonist"},
    "KBD1180": {"ct": "barzolvolimab OR (KIT inhibitor urticaria)",
                 "pubmed": "barzolvolimab OR (c-Kit urticaria)"},
    "KBD1298": {"ct": "DNA-PK inhibitor", "pubmed": "DNA-PK inhibitor"},
}


def _get_json(url: str, params: dict) -> dict | None:
    try:
        req = request.Request(
            f"{url}?{parse.urlencode(params)}", headers={"User-Agent": USER_AGENT}
        )
        with request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except (URLError, HTTPError, TimeoutError, ValueError) as e:
        print(f"[warn] request failed {url}: {e}", file=sys.stderr)
        return None


def fetch_trials(target_id: str, term: str) -> list[dict]:
    payload = _get_json(
        CT_GOV_API,
        {
            "query.term": term,
            "pageSize": str(CT_MAX),
            "sort": "LastUpdatePostDate:desc",
            "format": "json",
        },
    )
    if not payload:
        return []
    items: list[dict] = []
    for study in payload.get("studies", []):
        proto = study.get("protocolSection", {})
        ident = proto.get("identificationModule", {})
        status = proto.get("statusModule", {})
        nct = ident.get("nctId")
        if not nct:
            continue
        phases = proto.get("designModule", {}).get("phases", []) or []
        overall = status.get("overallStatus", "")
        updated = status.get("lastUpdatePostDateStruct", {}).get("date", "")
        title = ident.get("briefTitle", "") or ident.get("officialTitle", "")
        parts = []
        if overall:
            parts.append(f"状态：{overall}")
        if phases:
            parts.append("期别：" + "/".join(phases))
        if updated:
            parts.append(f"更新：{updated}")
        items.append(
            {
                "id": f"ct-{nct}-{target_id}",
                "date": (updated or "")[:10],
                "module": "clinical",
                "title": f"[{nct}] {title}",
                "body": [" · ".join(parts)] if parts else [],
                "sources": [
                    {"label": f"ClinicalTrials.gov（{nct}）",
                     "url": f"https://clinicaltrials.gov/study/{nct}"}
                ],
            }
        )
    return items


def fetch_pubmed(target_id: str, term: str) -> list[dict]:
    search = _get_json(
        PUBMED_ESEARCH,
        {
            "db": "pubmed",
            "term": term,
            "reldate": str(PUBMED_RELDATE_DAYS),
            "datetype": "pdat",
            "sort": "date",
            "retmax": str(PUBMED_MAX),
            "retmode": "json",
        },
    )
    if not search:
        return []
    ids = search.get("esearchresult", {}).get("idlist", [])
    if not ids:
        return []
    summary = _get_json(
        PUBMED_ESUMMARY, {"db": "pubmed", "id": ",".join(ids), "retmode": "json"}
    )
    if not summary:
        return []
    result = summary.get("result", {})
    items: list[dict] = []
    for pmid in ids:
        rec = result.get(pmid, {})
        if not rec:
            continue
        journal = rec.get("fulljournalname") or rec.get("source") or ""
        pubdate = rec.get("pubdate") or rec.get("epubdate") or ""
        # 归一化日期为 YYYY-MM-DD（PubMed 常见 "2026 Jul 8" 或 "2026 Jul"）
        iso = _pubdate_to_iso(pubdate)
        items.append(
            {
                "id": f"pubmed-{pmid}-{target_id}",
                "date": iso,
                "module": "article",
                "title": rec.get("title", "").strip(),
                "body": [journal] if journal else [],
                "sources": [
                    {"label": f"PubMed（PMID {pmid}）",
                     "url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"}
                ],
            }
        )
    return items


_MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}


def _pubdate_to_iso(s: str) -> str:
    parts = (s or "").split()
    if not parts:
        return ""
    year = parts[0]
    if not (year.isdigit() and len(year) == 4):
        return s
    month = _MONTHS.get(parts[1], 1) if len(parts) > 1 else 1
    day = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else 1
    return f"{year}-{month:02d}-{day:02d}"


def merge(existing: list[dict], incoming: list[dict]) -> list[dict]:
    by_id = {it["id"]: it for it in existing}
    for it in incoming:
        by_id[it["id"]] = it  # 覆盖更新（日期/状态可能变化）
    merged = list(by_id.values())
    merged.sort(key=lambda x: str(x.get("date", "")), reverse=True)
    return merged[:MAX_KEEP_PER_TARGET]


def main() -> int:
    if not DATA_FILE.exists():
        print(f"[error] {DATA_FILE} not found", file=sys.stderr)
        return 1
    data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    auto = data.get("auto") or {}

    target_ids = {t["id"] for t in data.get("targets", [])}
    total_new = 0
    for tid, kw in SEARCH.items():
        if tid not in target_ids:
            continue
        incoming = fetch_trials(tid, kw["ct"]) + fetch_pubmed(tid, kw["pubmed"])
        before = {it["id"] for it in auto.get(tid, [])}
        merged = merge(auto.get(tid, []), incoming)
        total_new += len([it for it in merged if it["id"] not in before])
        if merged:
            auto[tid] = merged

    data["auto"] = auto
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    data.setdefault("meta", {})["auto_generated_at"] = now

    DATA_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"OK: auto targets={len(auto)} new_items={total_new} at {now}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
