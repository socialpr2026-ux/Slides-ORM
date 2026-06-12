from __future__ import annotations

from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

MONTHS = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]

MONTH_BY_NUMBER = {f"{idx:02d}": month for idx, month in enumerate(MONTHS, start=1)}

STANDARD_ROW_ORDER = [
    "Отзывы (отзовики, аптеки) на топ-площадках без покупки",
    "Отзывы с покупкой",
    "Комментарии на площадках в топ-20",
    "Комментарии в открытых обсуждениях (симптоматика, диагнозы, препараты)",
    "Комментарии с нативным фото продукта",
    "Мониторинг и реагирование, PPT-отчет",
]

SERVICE_ROW_NAME = "Мониторинг и реагирование, PPT-отчет"

STOP_ROW_PATTERNS = [
    r"^sov", r"^sov%", r"прирост\s+sov", r"стоимость", r"бюджет",
    r"кол-во просмотров", r"количество просмотров", r"источник данных",
    r"упоминания? в месяц"
]

TOTAL_ROW_HINTS = [
    "итого всех единиц", "итого единиц", "итого размещений", "итого"
]


def _safe_text(v, limit=1000):
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
        n = n * 26 + ord(c.upper()) - 64
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


def _read_sheet_matrix(z, path, shared, maxcols=80):
    root = ET.fromstring(z.read(path))
    rows = {}
    dim = root.find("x:dimension", NS)
    for row in root.findall("x:sheetData/x:row", NS):
        rnum = int(row.attrib.get("r", "0"))
        vals = [""] * maxcols
        for c in row.findall("x:c", NS):
            _, col = _ref_rc(c.attrib.get("r", ""))
            if 1 <= col <= maxcols:
                vals[col - 1] = _cell_value(c, shared)
        rows[rnum] = vals
    return rows, dim.attrib.get("ref", "") if dim is not None else ""


def _is_number(v):
    try:
        float(v)
        return str(v).strip() != ""
    except Exception:
        return False


def _to_num(v):
    try:
        f = float(v)
        if f.is_integer():
            return int(f)
        return round(f, 2)
    except Exception:
        return ""


def _sum_values(values):
    total = 0
    has_any = False
    for value in values:
        if _is_number(value):
            total += float(value)
            has_any = True
    if not has_any:
        return ""
    return int(total) if total.is_integer() else round(total, 2)


def _month_base(label):
    text = _safe_text(label)
    for m in MONTHS:
        if m.lower() in text.lower():
            return m
    return ""


def _period_month(period_hint: str) -> str:
    text = _safe_text(period_hint).lower()
    date_match = re.search(r"\d{1,2}[.](\d{1,2})[.]\d{2,4}", text)
    if date_match:
        return MONTH_BY_NUMBER.get(date_match.group(1).zfill(2), "")
    for month in MONTHS:
        if month.lower() in text:
            return month
        stem = month.lower()[:-1]
        if stem and stem in text:
            return month
    return ""


def _previous_month(month: str) -> str:
    if month not in MONTHS:
        return ""
    return MONTHS[(MONTHS.index(month) - 1) % len(MONTHS)]


def _is_fact_col(label):
    return "факт" in _safe_text(label).lower()


def _is_plan_col(label):
    return "план" in _safe_text(label).lower()


def _is_total_col(label):
    label = _safe_text(label).lower()
    return "всего" in label or "total" in label


def _find_media_plan_sheet(sheets, rows_by_sheet):
    for name in sheets:
        low = name.lower()
        if "медиаплан" in low or low.startswith("mp") or "media" in low:
            return name
    for name, rows in rows_by_sheet.items():
        text = " ".join(_safe_text(c, 100) for vals in rows.values() for c in vals[:5] if c)
        if "медиаплан" in text.lower() or "наименование работ" in text.lower():
            return name
    return next(iter(sheets))


def _find_header_row(rows):
    best = None
    for rnum, vals in rows.items():
        first = _safe_text(vals[0]).lower()
        row_text = " ".join(_safe_text(v).lower() for v in vals[:24])
        if "наименование работ" in row_text:
            return rnum
        if "репутационная поддержка" in first:
            return rnum
        month_count = sum(1 for v in vals[:30] if _month_base(v))
        if month_count >= 3:
            best = rnum if best is None else min(best, rnum)
    if best is not None:
        return best
    raise ValueError("Media plan header row not found")


def _detect_layout(header):
    month_cols = []
    total_cols = []
    for idx, value in enumerate(header):
        label = _safe_text(value)
        if _is_total_col(label):
            total_cols.append({"index": idx, "label": label})
            continue
        month = _month_base(label)
        if month:
            month_cols.append({
                "index": idx,
                "label": label,
                "month": month,
                "is_plan": _is_plan_col(label),
                "is_fact": _is_fact_col(label),
            })
    has_plan_fact_pairs = any(c["is_plan"] or c["is_fact"] for c in month_cols)
    return {
        "has_plan_fact_pairs": has_plan_fact_pairs,
        "month_cols": month_cols,
        "total_cols": total_cols,
    }


def _select_last_two_month_cols(layout, rows):
    month_cols = layout["month_cols"]
    if layout["has_plan_fact_pairs"]:
        fact_cols = [c for c in month_cols if c["is_fact"]]
        selected = fact_cols[-2:] if len(fact_cols) >= 2 else fact_cols[-2:]
    else:
        non_empty = []
        for c in month_cols:
            has_val = False
            for vals in rows.values():
                if c["index"] < len(vals) and _is_number(vals[c["index"]]):
                    has_val = True
                    break
            if has_val:
                non_empty.append(c)
        selected = non_empty[-2:] if len(non_empty) >= 2 else non_empty[-2:]
    return selected


def _empty_month_col(month: str) -> dict:
    return {
        "index": None,
        "label": month,
        "month": month,
        "is_plan": False,
        "is_fact": False,
        "missing_in_source": True,
    }


def _select_period_month_cols(layout, rows, period_hint: str = ""):
    current_month = _period_month(period_hint)
    if not current_month:
        return _select_last_two_month_cols(layout, rows)
    target_months = [_previous_month(current_month), current_month]
    selected = []
    for month in target_months:
        candidates = [c for c in layout["month_cols"] if c["month"] == month]
        if layout["has_plan_fact_pairs"]:
            fact_candidates = [c for c in candidates if c["is_fact"]]
            candidates = fact_candidates or candidates
        if candidates:
            selected.append(candidates[-1])
        else:
            selected.append(_empty_month_col(month))
    return selected


def _should_stop_activity_row(name):
    low = name.lower()
    return any(re.search(pattern, low) for pattern in STOP_ROW_PATTERNS)


def _is_total_row(name):
    low = name.lower()
    return any(hint in low for hint in TOTAL_ROW_HINTS)


def standardize_format_name(raw_name: str) -> str:
    low = _safe_text(raw_name, 1000).lower()
    if "мониторинг" in low or "ppt" in low or "презентац" in low or ("отчет" in low and "sov" not in low):
        return SERVICE_ROW_NAME
    # IMPORTANT: classify comment rows before review rows, because many comment rows include the word "отзывам".
    if "нативным фото" in low or "фото продукта" in low or "с фото" in low:
        return "Комментарии с нативным фото продукта"
    if "топ-20" in low or "топ 20" in low or "serm" in low or "выдач" in low or "вопросы и ответы" in low:
        return "Комментарии на площадках в топ-20"
    if "встраивания" in low or "обсуждения" in low or "симптом" in low or "сегмент" in low:
        return "Комментарии в открытых обсуждениях (симптоматика, диагнозы, препараты)"
    if "отзыв" in low and ("без условия покупки" in low or "без покупки" in low):
        return "Отзывы (отзовики, аптеки) на топ-площадках без покупки"
    if "отзыв" in low and ("покуп" in low or "условием покупки" in low or "аптека.ру" in low or "apteka.ru" in low):
        return "Отзывы с покупкой"
    if "отзыв" in low:
        return "Отзывы (отзовики, аптеки) на топ-площадках без покупки"
    return _safe_text(raw_name, 220)


def _extract_activity_rows(rows, header_row, selected_cols):
    out = []
    for rnum in sorted(rows):
        if rnum <= header_row:
            continue
        vals = rows[rnum]
        raw_name = _safe_text(vals[0], 600)
        if not raw_name:
            continue
        if _should_stop_activity_row(raw_name):
            break
        if _is_total_row(raw_name):
            # total rows do not go to slide 2 body. Slide uses standardized rows + service row.
            continue
        has_numeric = any(c.get("index") is not None and c["index"] < len(vals) and _is_number(vals[c["index"]]) for c in selected_cols)
        if not has_numeric:
            continue
        values = [
            _to_num(vals[c["index"]] if c.get("index") is not None and c["index"] < len(vals) else "")
            if c.get("index") is not None else 0
            for c in selected_cols
        ]
        out.append({
            "source_row": rnum,
            "source_name": raw_name,
            "name": standardize_format_name(raw_name),
            "values": values,
            "total": _sum_values(values),
            "is_service_row": standardize_format_name(raw_name) == SERVICE_ROW_NAME,
        })
    return out


def _dedupe_and_order(rows):
    by_name = {}
    source_map = {}
    for row in rows:
        name = row["name"]
        if name not in by_name:
            by_name[name] = dict(row)
            source_map[name] = [row["source_name"]]
        else:
            # If duplicate standardized rows appear, sum values by month.
            current = by_name[name]
            values = []
            max_len = max(len(current.get("values", [])), len(row.get("values", [])))
            for idx in range(max_len):
                a = current.get("values", [""] * max_len)[idx] if idx < len(current.get("values", [])) else ""
                b = row.get("values", [""] * max_len)[idx] if idx < len(row.get("values", [])) else ""
                values.append(_sum_values([a, b]))
            current["values"] = values
            current["total"] = _sum_values(values)
            source_map[name].append(row["source_name"])
    ordered = []
    for name in STANDARD_ROW_ORDER:
        if name in by_name:
            item = by_name[name]
            item["source_names"] = source_map[name]
            ordered.append(item)
    # add any unknown rows before service row, if present
    unknowns = [row for name, row in by_name.items() if name not in STANDARD_ROW_ORDER]
    if unknowns:
        service = [r for r in ordered if r["name"] == SERVICE_ROW_NAME]
        ordered = [r for r in ordered if r["name"] != SERVICE_ROW_NAME] + unknowns + service
    return ordered


def _ensure_service_row(rows, selected_cols, default_values=None):
    if default_values is None:
        default_values = [1, 1]
    rows = [row for row in rows if row["name"] != SERVICE_ROW_NAME]
    values = []
    for idx, col in enumerate(selected_cols):
        if col.get("index") is None:
            values.append(0)
        else:
            values.append(default_values[idx] if idx < len(default_values) else 1)
    service_row = {
        "source_row": None,
        "source_name": "service row default",
        "source_names": ["service row default"],
        "name": SERVICE_ROW_NAME,
        "values": values,
        "total": _sum_values(values),
        "is_service_row": True,
        "added_by_script": True,
    }
    return rows + [service_row]


def _is_missing_value(value) -> bool:
    text = _safe_text(value).lower()
    return text in {"", "нет данных", "н/д", "na", "n/a", "-"}


def _selected_month_index(selected_cols: list[dict], month_label: str = "") -> int | None:
    if month_label:
        target = _month_base(month_label) or _safe_text(month_label)
        for idx, col in enumerate(selected_cols):
            if col.get("month") == target or _safe_text(col.get("label")) == target:
                return idx
    missing = [idx for idx, col in enumerate(selected_cols) if col.get("missing_in_source")]
    if missing:
        return missing[-1]
    return len(selected_cols) - 1 if selected_cols else None


def apply_fact_overrides_to_media_plan_table(media_plan_table: dict, fact_data: dict, *, month_label: str = "") -> dict:
    """Override the selected fact month in slide 2 from a project month fact table.

    Project ORM rows are the authoritative monthly fact. Media-plan values can
    contain planning mistakes, so the selected month is replaced for matching
    standardized publication rows. Service rows are not part of publication fact.
    """
    selected_cols = list(media_plan_table.get("selected_columns") or [])
    fact_rows = dict((fact_data or {}).get("fact_rows_by_name") or {})
    target_idx = _selected_month_index(selected_cols, month_label or (fact_data or {}).get("sheet", ""))
    override_result = {
        "status": "skipped",
        "reason": "",
        "target_month": "",
        "source_file": (fact_data or {}).get("source_file", ""),
        "source_sheet": (fact_data or {}).get("sheet", ""),
        "total_fact": int((fact_data or {}).get("total_fact") or (fact_data or {}).get("campaign_publications_count") or 0),
        "method": (fact_data or {}).get("campaign_fact_method") or (fact_data or {}).get("method", ""),
        "applied_count": 0,
        "overridden_rows": [],
        "warnings": list((fact_data or {}).get("warnings") or []),
    }
    if target_idx is None:
        override_result["reason"] = "No selected month column to fill."
        media_plan_table["fact_overrides"] = override_result
        return media_plan_table
    if not fact_rows:
        override_result["reason"] = "No fact rows found in configured fact source."
        media_plan_table["fact_overrides"] = override_result
        return media_plan_table

    target_month = selected_cols[target_idx].get("month") or selected_cols[target_idx].get("label") or ""
    override_result["target_month"] = target_month
    applied_names = set()
    for row in media_plan_table.get("full_rows") or []:
        name = row.get("name")
        if name not in fact_rows:
            continue
        values = list(row.get("values") or [])
        while len(values) <= target_idx:
            values.append("")
        original = values[target_idx]
        values[target_idx] = fact_rows[name]
        row["values"] = values
        row["total"] = _sum_values(values)
        row["fact_override"] = {
            "month": target_month,
            "original_value": original,
            "source_file": override_result["source_file"],
            "source_sheet": override_result["source_sheet"],
            "method": override_result["method"],
        }
        applied_names.add(name)
        override_result["overridden_rows"].append({"name": name, "original_value": original, "fact_value": fact_rows[name]})

    existing_full_names = {row.get("name") for row in media_plan_table.get("full_rows") or []}
    missing_rows = []
    for name in STANDARD_ROW_ORDER:
        if name == SERVICE_ROW_NAME or name in existing_full_names:
            continue
        value = int(fact_rows.get(name) or 0)
        if value <= 0:
            continue
        values = [0 if c.get("missing_in_source") else "" for c in selected_cols]
        while len(values) <= target_idx:
            values.append("")
        values[target_idx] = value
        missing_rows.append({
            "source_row": None,
            "source_name": "project ORM fact row",
            "source_names": ["project ORM fact row"],
            "name": name,
            "values": values,
            "total": _sum_values(values),
            "is_service_row": False,
            "added_by_script": True,
            "fact_override": {
                "month": target_month,
                "original_value": "",
                "source_file": override_result["source_file"],
                "source_sheet": override_result["source_sheet"],
                "method": override_result["method"],
            },
        })
        applied_names.add(name)
        override_result["overridden_rows"].append({"name": name, "original_value": "", "fact_value": value})

    if missing_rows:
        full_rows = [row for row in media_plan_table.get("full_rows") or [] if row.get("name") != SERVICE_ROW_NAME]
        service_rows = [row for row in media_plan_table.get("full_rows") or [] if row.get("name") == SERVICE_ROW_NAME]
        media_plan_table["full_rows"] = _dedupe_and_order(full_rows + missing_rows + service_rows)
        max_slide_rows = max(len((media_plan_table.get("slide_table") or {}).get("rows") or []), len(STANDARD_ROW_ORDER))
        slide_rows = media_plan_table["full_rows"][:max_slide_rows]
        service = [row for row in media_plan_table["full_rows"] if row.get("name") == SERVICE_ROW_NAME]
        if service and all(row.get("name") != SERVICE_ROW_NAME for row in slide_rows):
            slide_rows = slide_rows[: max_slide_rows - 1] + service[:1]
        media_plan_table.setdefault("slide_table", {})["rows"] = slide_rows
        media_plan_table["slide_rows_count"] = len(slide_rows)
        media_plan_table["full_rows_count"] = len(media_plan_table["full_rows"])
        media_plan_table["overflow_rows"] = media_plan_table["full_rows"][max_slide_rows:]
        media_plan_table["overflow_rows_count"] = len(media_plan_table["overflow_rows"])

    for row in media_plan_table.get("full_rows") or []:
        name = row.get("name")
        if name == SERVICE_ROW_NAME or name not in STANDARD_ROW_ORDER or name in fact_rows:
            continue
        values = list(row.get("values") or [])
        while len(values) <= target_idx:
            values.append("")
        original = values[target_idx]
        values[target_idx] = 0
        row["values"] = values
        row["total"] = _sum_values(values)
        row["fact_override"] = {
            "month": target_month,
            "original_value": original,
            "source_file": override_result["source_file"],
            "source_sheet": override_result["source_sheet"],
            "method": override_result["method"],
        }
        override_result["overridden_rows"].append({"name": name, "original_value": original, "fact_value": 0})

    for row in media_plan_table.get("slide_table", {}).get("rows") or []:
        name = row.get("name")
        if name == SERVICE_ROW_NAME or name not in STANDARD_ROW_ORDER:
            continue
        values = list(row.get("values") or [])
        while len(values) <= target_idx:
            values.append("")
        original = values[target_idx]
        values[target_idx] = fact_rows.get(name, 0)
        row["values"] = values
        row["total"] = _sum_values(values)
        row["fact_override"] = {
            "month": target_month,
            "original_value": original,
            "source_file": override_result["source_file"],
            "source_sheet": override_result["source_sheet"],
            "method": override_result["method"],
        }
        applied_names.add(name)

    override_result["applied_count"] = len(applied_names)
    override_result["status"] = "applied" if applied_names else "not_needed"
    if not applied_names:
        override_result["reason"] = "No matching standard publication rows were present."
    media_plan_table["fact_overrides"] = override_result
    media_plan_table["qa_notes"].append(
        "Configured project ORM month fact overrides the selected media-plan fact month for matching publication rows."
    )
    return media_plan_table


def extract_media_plan_table(path: Path, max_rows_for_slide=6, period_hint: str = "") -> dict:
    with zipfile.ZipFile(path, "r") as z:
        shared = _load_shared_strings(z)
        sheets = _get_sheets(z)
        rows_by_sheet = {}
        ranges = {}
        for name, sheet_path in sheets.items():
            rows, dim = _read_sheet_matrix(z, sheet_path, shared)
            rows_by_sheet[name] = rows
            ranges[name] = dim

        sheet_name = _find_media_plan_sheet(sheets, rows_by_sheet)
        rows = rows_by_sheet[sheet_name]
        header_row = _find_header_row(rows)
        header = rows[header_row]
        layout = _detect_layout(header)
        selected_cols = _select_period_month_cols(layout, rows, period_hint)
        activity_rows = _extract_activity_rows(rows, header_row, selected_cols)
        activity_rows = _dedupe_and_order(activity_rows)
        activity_rows = _ensure_service_row(activity_rows, selected_cols)

        # no empty rows: if there are too few rows, service row is already present; if still fewer than capacity,
        # repeat is forbidden, so slide body can have fewer rows only if template capacity is reduced in future.
        # Current template has exactly 6 standard body rows.
        selected_header = [c["month"] for c in selected_cols]
        total_label = "Всего"

        slide_rows = activity_rows[:max_rows_for_slide]
        overflow_rows = activity_rows[max_rows_for_slide:]

        return {
            "source_file": path.name,
            "sheet": sheet_name,
            "sheet_range": ranges.get(sheet_name, ""),
            "header_row": header_row,
            "layout": layout,
            "selected_columns": selected_cols,
            "selected_header": selected_header,
            "total_mode": "sum_selected_months",
            "total_label": total_label,
            "full_rows_count": len(activity_rows),
            "slide_rows_count": len(slide_rows),
            "overflow_rows_count": len(overflow_rows),
            "slide_table": {
                "headers": ["Публикации в соцмедиа"] + selected_header + [total_label],
                "rows": slide_rows,
            },
            "full_rows": activity_rows,
            "overflow_rows": overflow_rows,
            "status": "ready" if selected_cols and activity_rows else "missing_required",
            "qa_notes": [
                "Last three column headers must be centered.",
                "Total column label is always 'Всего'.",
                "Total is computed as the sum of the two selected month columns.",
                "Rows are standardized to presentation-level labels and source names are preserved.",
                "No empty rows are allowed.",
                "Service row 'Мониторинг и реагирование, PPT-отчет' is always last.",
            ],
        }
