from __future__ import annotations

from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

NS = {
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _safe_text(v, limit=500):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()[:limit]


def _load_shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.findall(".//x:t", NS)) for si in root.findall("x:si", NS)]


def _get_sheets(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall("rel:Relationship", NS)}
    out = {}
    for sh in wb.findall("x:sheets/x:sheet", NS):
        rid = sh.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        out[sh.attrib.get("name")] = ("xl/" + relmap.get(rid, "").lstrip("/")).replace("xl//", "xl/")
    return out


def _colnum(col):
    n = 0
    for c in col:
        n = n * 26 + ord(c) - 64
    return n


def _ref_rc(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref or "")
    return (int(m.group(2)), _colnum(m.group(1))) if m else (0, 0)


def _cell_value(c, shared):
    t = c.attrib.get("t")
    v = c.find("x:v", NS)
    if t == "inlineStr":
        return "".join(x.text or "" for x in c.findall(".//x:t", NS))
    if v is None:
        return ""
    raw = v.text or ""
    if t == "s":
        return shared[int(raw)] if raw.isdigit() and int(raw) < len(shared) else raw
    return raw


def _is_review_type(value: str) -> bool:
    low = _safe_text(value, 200).lower().replace("ё", "е")
    return "отзыв" in low and "коммент" not in low


def _is_comment_type(value: str) -> bool:
    low = _safe_text(value, 200).lower().replace("ё", "е")
    return any(term in low for term in ["коммент", "обсужд", "встраив", "serm", "топ", "выдач"])


def _read_sheet_matrix(z, path, shared, maxcols=60):
    root = ET.fromstring(z.read(path))
    rows = {}
    for row in root.findall("x:sheetData/x:row", NS):
        rnum = int(row.attrib.get("r", "0"))
        vals = [""] * maxcols
        for c in row.findall("x:c", NS):
            _, col = _ref_rc(c.attrib.get("r", ""))
            if 1 <= col <= maxcols:
                vals[col - 1] = _cell_value(c, shared)
        rows[rnum] = vals
    return rows


def extract_plan_fact_from_orm_excel(path: Path) -> dict:
    """Best-effort extraction from an ORM placement/report Excel.

    Universal intent: detect a sheet with placement rows and derive plan/fact.
    Current v0.2 supports the tested layout but keeps warnings for PM.
    """
    with zipfile.ZipFile(path, "r") as z:
        shared = _load_shared_strings(z)
        sheets = _get_sheets(z)
        pub_name = next((s for s in sheets if "публика" in s.lower()), None)
        if not pub_name:
            return {"source_file": path.name, "status": "missing_required", "error": "No publication sheet detected"}

        rows = _read_sheet_matrix(z, sheets[pub_name], shared, maxcols=30)
        plan_row = rows.get(7, [])
        plan_text = plan_row[2] if len(plan_row) > 2 else ""

        plan_reviews = 0
        plan_comments = 0
        m = re.search(r"Отзывы\s*[-–]\s*(\d+)", str(plan_text), flags=re.I)
        if m:
            plan_reviews = int(m.group(1))
        m = re.search(r"комментарии\s*[-–]\s*(\d+)", str(plan_text), flags=re.I)
        if m:
            plan_comments = int(m.group(1))

        def _num(value):
            try:
                return int(float(value)) if str(value).strip() else 0
            except Exception:
                return 0

        plan_by_post_type = {
            "ОС": _num(plan_row[7] if len(plan_row) > 7 else 0),
            "СС": _num(plan_row[8] if len(plan_row) > 8 else 0),
            "ЦС": _num(plan_row[9] if len(plan_row) > 9 else 0),
            "ПС": _num(plan_row[10] if len(plan_row) > 10 else 0),
            "Всего": _num(plan_row[11] if len(plan_row) > 11 else 0),
        }

        format_counts = Counter()
        post_type_counts = Counter()
        views_total = 0
        link_count = 0

        for rnum, vals in rows.items():
            if rnum < 19 or not any(vals):
                continue
            placement_type = vals[0] if len(vals) > 0 else ""
            link = vals[3] if len(vals) > 3 else ""
            post_type = vals[9] if len(vals) > 9 else ""
            if placement_type or link:
                format_counts[_safe_text(placement_type, 80)] += 1
                if post_type:
                    post_type_counts[_safe_text(post_type, 20)] += 1
                if link:
                    link_count += 1
                try:
                    views_total += float(vals[7]) if vals[7] != "" else 0
                except Exception:
                    pass

        fact_reviews = sum(count for typ, count in format_counts.items() if _is_review_type(typ))
        fact_comments = sum(count for typ, count in format_counts.items() if _is_comment_type(typ))
        fact_total = fact_reviews + fact_comments

        return {
            "source_file": path.name,
            "sheet": pub_name,
            "status": "ready_with_caveat",
            "brand": rows.get(5, ["", "", ""])[2],
            "period_excel": rows.get(6, ["", "", ""])[2],
            "plan_text": plan_text,
            "plan_reviews": plan_reviews,
            "plan_comments": plan_comments,
            "plan_total_from_text": plan_reviews + plan_comments,
            "plan_by_post_type": plan_by_post_type,
            "fact_reviews": fact_reviews,
            "fact_comments": fact_comments,
            "fact_total": fact_total,
            "campaign_publications_count": fact_total,
            "campaign_publications_by_type": dict(format_counts),
            "campaign_fact_method": "publication_sheet_rows",
            "campaign_fact_source_sheet": pub_name,
            "campaign_vs_organic_status": "ready_with_caveat" if fact_total else "missing_required",
            "format_counts": dict(format_counts),
            "post_type_counts": dict(post_type_counts),
            "views_total": int(views_total),
            "link_count": link_count,
            "qa_flags": [
                "Confirm which plan total is client-facing.",
                "Do not explain variance without PM confirmation.",
                "Links and views require QA before final client use.",
            ],
        }
