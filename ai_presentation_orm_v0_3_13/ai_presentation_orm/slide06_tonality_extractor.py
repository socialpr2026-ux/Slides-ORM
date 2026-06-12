from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path
import re

from .slide03_data_extractor import _brand_name_matches, _canonical_brand_name, _clean_int, _is_thematic_metric_label, _norm_brand_name, _period_label_from_content, _period_to_caption
from .slide05_sov_extractor import _source_label
from .xlsx_safe_reader import open_workbook_safe


TONE_KEYS = ("positive", "neutral", "negative")
TONE_LABELS = {
    "positive": "Позитив",
    "neutral": "Нейтрально",
    "negative": "Негатив",
}


def _lower(value) -> str:
    return str(value or "").strip().lower()


def _norm_tone(value) -> str:
    text = _lower(value)
    if "позит" in text or "полож" in text:
        return "positive"
    if "негат" in text or "отриц" in text:
        return "negative"
    if "нейтр" in text:
        return "neutral"
    return ""


def _header_index(headers, *names: str):
    lowered = [_lower(h) for h in headers]
    for name in names:
        needle = name.lower()
        for idx, header in enumerate(lowered):
            if header == needle:
                return idx
    for name in names:
        needle = name.lower()
        for idx, header in enumerate(lowered):
            if needle in header:
                return idx
    return None


def _find_header(ws, required: list[str], max_rows: int = 80):
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows), values_only=True), start=1):
        headers = list(row)
        lowered = [_lower(h) for h in headers]
        if all(any(req.lower() in h for h in lowered) for req in required):
            return idx, headers
    raise ValueError(f"Header with required columns not found: {required}")


def _find_messages_header(ws, max_rows: int = 80):
    return _find_header(ws, ["Тональность"], max_rows=max_rows)


def _object_values(value) -> list[str]:
    text = str(value or "").strip()
    if not text:
        return []
    return [part.strip() for part in re.split(r"[,;|]+", text) if part.strip()]


def _value_mentions_brand(value, project_brand: str) -> bool:
    if _brand_name_matches(str(value or ""), project_brand):
        return True
    return any(_brand_name_matches(part, project_brand) for part in _object_values(value))


def _row_mentions_project(values: list, project_brand: str, candidate_indices: list[int | None]) -> bool:
    if not _lower(project_brand):
        return False
    for idx in candidate_indices:
        if idx is None or idx >= len(values):
            continue
        if _value_mentions_brand(values[idx], project_brand):
            return True
    return False


def _date_key(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    text = str(value or "").replace("\xa0", " ").strip()
    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%d.%m.%Y")
        except Exception:
            pass
    return text


def _date_sort_key(value: str):
    try:
        return datetime.strptime(value, "%d.%m.%Y")
    except Exception:
        return datetime.max


def _period_dates(period_raw: str) -> list[str]:
    text = str(period_raw or "")
    import re

    match = re.search(
        r"(\d{1,2})[.](\d{1,2})[.](\d{2,4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?\s*[–-]\s*"
        r"(\d{1,2})[.](\d{1,2})[.](\d{2,4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?",
        text,
    )
    if not match:
        return []
    d1, m1, y1, d2, m2, y2 = [int(part) for part in match.groups()]
    y1 = 2000 + y1 if y1 < 100 else y1
    y2 = 2000 + y2 if y2 < 100 else y2
    start = datetime(y1, m1, d1)
    end = datetime(y2, m2, d2)
    if end < start:
        return []
    out = []
    current = start
    while current <= end:
        out.append(current.strftime("%d.%m.%Y"))
        current += timedelta(days=1)
    return out


def _safe_pct(part: int, total: int) -> float:
    return part / total if total else 0.0


def _filter_competitive_brand_rows(rows: list[dict], project_brand: str, competitor_brands: list[str] | None = None) -> list[dict]:
    competitors = [str(item).strip() for item in (competitor_brands or []) if str(item or "").strip()]
    canonical_rows = []
    for row in rows:
        item = dict(row)
        item["brand"] = _canonical_brand_name(item.get("brand", ""), project_brand, competitors)
        canonical_rows.append(item)
    if not competitors:
        return _aggregate_tonality_rows(canonical_rows)
    allowed = {_norm_brand_name(project_brand), *{_norm_brand_name(item) for item in competitors}}
    allowed.discard("")
    filtered = [row for row in canonical_rows if _norm_brand_name(row.get("brand", "")) in allowed]
    return _aggregate_tonality_rows(filtered)


def _aggregate_tonality_rows(rows: list[dict]) -> list[dict]:
    grouped: dict[str, dict] = {}
    for row in rows:
        brand = str(row.get("brand") or "").strip()
        if not brand:
            continue
        target = grouped.setdefault(brand, {"brand": brand, "total": 0, "positive": 0, "neutral": 0, "negative": 0})
        for key in ["total", "positive", "neutral", "negative"]:
            target[key] += int(row.get(key, 0) or 0)
    out = []
    for item in grouped.values():
        total = int(item.get("total") or 0) or int(item.get("positive", 0) or 0) + int(item.get("neutral", 0) or 0) + int(item.get("negative", 0) or 0)
        item["total"] = total
        item["positive_share"] = _safe_pct(int(item.get("positive", 0) or 0), total)
        item["neutral_share"] = _safe_pct(int(item.get("neutral", 0) or 0), total)
        item["negative_share"] = _safe_pct(int(item.get("negative", 0) or 0), total)
        out.append(item)
    return out


def _extract_brand_tonality_from_messages(wb, project_brand: str, competitor_brands: list[str] | None = None) -> tuple[list[dict], dict]:
    if "Сообщения" not in wb.sheetnames:
        return [], {}
    ws = wb["Сообщения"]
    header_row, headers = _find_messages_header(ws, max_rows=80)
    tone_idx = _header_index(headers, "Тональность")
    object_indices = [
        idx for idx in (
            _header_index(headers, "Объекты", "Объект"),
            _header_index(headers, "Теги", "Тег"),
        )
        if idx is not None
    ]
    if tone_idx is None or not object_indices:
        return [], {}

    counts = defaultdict(lambda: Counter({key: 0 for key in TONE_KEYS}))
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        values = list(row)
        tone = _norm_tone(values[tone_idx] if tone_idx < len(values) else "")
        if not tone:
            continue
        objects = []
        for idx in object_indices:
            if idx < len(values):
                objects.extend(_object_values(values[idx]))
        for name in objects:
            counts[name][tone] += 1

    rows = []
    for name, counter in counts.items():
        total = int(sum(counter.values()))
        if not name or total <= 0:
            continue
        rows.append({
            "brand": name,
            "total": total,
            "positive": int(counter["positive"]),
            "neutral": int(counter["neutral"]),
            "negative": int(counter["negative"]),
            "positive_share": _safe_pct(counter["positive"], total),
            "neutral_share": _safe_pct(counter["neutral"], total),
            "negative_share": _safe_pct(counter["negative"], total),
        })
    rows = sorted(_filter_competitive_brand_rows(rows, project_brand, competitor_brands), key=lambda item: item["total"], reverse=True)
    project_row = next((item for item in rows if _brand_name_matches(item.get("brand", ""), project_brand)), {})
    return rows, project_row


def _extract_brand_tonality(wb, project_brand: str, competitor_brands: list[str] | None = None) -> tuple[list[dict], dict]:
    message_rows, message_project_row = _extract_brand_tonality_from_messages(wb, project_brand, competitor_brands)
    if message_project_row:
        return message_rows, message_project_row
    if "Теги" not in wb.sheetnames:
        return message_rows, message_project_row
    ws = wb["Теги"]
    header_row, headers = _find_header(ws, ["Тег", "Сообщения", "Позитив", "Негатив"])
    tag_idx = _header_index(headers, "Тег", "Теги")
    total_idx = _header_index(headers, "Сообщения")
    pos_idx = _header_index(headers, "Позитив", "Позитивные")
    neu_idx = _header_index(headers, "Нейтрально", "Нейтральные", "Нейтрал")
    neg_idx = _header_index(headers, "Негатив", "Негативные")
    if None in (tag_idx, total_idx, pos_idx, neu_idx, neg_idx):
        raise ValueError("Brand tonality columns are incomplete on sheet 'Теги'")

    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        values = list(row)
        name = str(values[tag_idx] if tag_idx < len(values) and values[tag_idx] is not None else "").strip()
        if not name:
            continue
        low = name.lower()
        if low == "дата":
            break
        if low == "всего" or _is_thematic_metric_label(name):
            continue
        total = _clean_int(values[total_idx] if total_idx < len(values) else 0)
        positive = _clean_int(values[pos_idx] if pos_idx < len(values) else 0)
        neutral = _clean_int(values[neu_idx] if neu_idx < len(values) else 0)
        negative = _clean_int(values[neg_idx] if neg_idx < len(values) else 0)
        if total <= 0 and positive + neutral + negative <= 0:
            continue
        total = total or positive + neutral + negative
        rows.append({
            "brand": name,
            "total": total,
            "positive": positive,
            "neutral": neutral,
            "negative": negative,
            "positive_share": _safe_pct(positive, total),
            "neutral_share": _safe_pct(neutral, total),
            "negative_share": _safe_pct(negative, total),
        })
    rows = sorted(_filter_competitive_brand_rows(rows, project_brand, competitor_brands), key=lambda item: item["total"], reverse=True)
    project_row = next((item for item in rows if _brand_name_matches(item.get("brand", ""), project_brand)), {})
    if not project_row:
        if message_project_row:
            return message_rows, message_project_row
    return rows, project_row


def _extract_daily_project_tonality(wb, project_brand: str, period_raw: str) -> tuple[list[dict], dict]:
    if "Сообщения" not in wb.sheetnames:
        return _extract_daily_fallback(wb, period_raw)
    ws = wb["Сообщения"]
    header_row, headers = _find_messages_header(ws, max_rows=80)
    date_idx = _header_index(headers, "Дата", "Время публикации", "Время")
    tone_idx = _header_index(headers, "Тональность")
    brand_idx = _header_index(headers, project_brand)
    object_idx = _header_index(headers, "Объекты", "Объект")
    tag_idx = _header_index(headers, "Теги", "Тег")
    filter_indices = [idx for idx in (object_idx, tag_idx) if idx is not None]
    if brand_idx is None and not filter_indices:
        return _extract_daily_fallback(wb, period_raw)

    by_day = defaultdict(lambda: Counter({key: 0 for key in TONE_KEYS}))
    top_sources = Counter()
    top_sites = Counter()
    top_message_types = Counter()
    samples = {key: [] for key in TONE_KEYS}
    source_idx = _header_index(headers, "Тип источника", "Тип площадки")
    site_idx = _header_index(headers, "Источник", "Площадка", "Где пишет")
    message_type_idx = _header_index(headers, "Тип сообщения", "Тип")
    text_idx = _header_index(headers, "Текст", "Текст сообщения")
    link_idx = _header_index(headers, "Ссылка на сообщение", "URL сообщения", "Url сообщения", "Ссылка")

    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        values = list(row)
        has_brand_column_hit = brand_idx is not None and brand_idx < len(values) and str(values[brand_idx] or "").strip()
        has_object_hit = _row_mentions_project(values, project_brand, filter_indices)
        if not (has_brand_column_hit or has_object_hit):
            continue
        tone = _norm_tone(values[tone_idx] if tone_idx is not None and tone_idx < len(values) else "")
        if not tone:
            continue
        date = _date_key(values[date_idx] if date_idx is not None and date_idx < len(values) else "")
        if not date:
            continue
        by_day[date][tone] += 1
        if source_idx is not None and source_idx < len(values) and values[source_idx]:
            top_sources[str(values[source_idx]).strip()] += 1
        if site_idx is not None and site_idx < len(values) and values[site_idx]:
            top_sites[str(values[site_idx]).strip()] += 1
        if message_type_idx is not None and message_type_idx < len(values) and values[message_type_idx]:
            top_message_types[str(values[message_type_idx]).strip()] += 1
        if text_idx is not None and text_idx < len(values) and values[text_idx] and len(samples[tone]) < 5:
            samples[tone].append({
                "text": str(values[text_idx]).strip()[:220],
                "link": str(values[link_idx]).strip() if link_idx is not None and link_idx < len(values) and values[link_idx] else "",
                "site": str(values[site_idx]).strip() if site_idx is not None and site_idx < len(values) and values[site_idx] else "",
                "source_type": str(values[source_idx]).strip() if source_idx is not None and source_idx < len(values) and values[source_idx] else "",
            })

    dates = _period_dates(period_raw) or sorted(by_day.keys(), key=_date_sort_key)
    daily_rows = []
    for date in dates:
        counter = by_day[date]
        daily_rows.append({
            "date": date,
            "positive": int(counter["positive"]),
            "neutral": int(counter["neutral"]),
            "negative": int(counter["negative"]),
            "total": int(sum(counter.values())),
        })
    return daily_rows, {
        "method": "messages_sheet_project_brand_filter",
        "top_sources": dict(top_sources.most_common(5)),
        "top_sites": dict(top_sites.most_common(8)),
        "top_message_types": dict(top_message_types.most_common(5)),
        "samples": samples,
    }


def _extract_daily_fallback(wb, period_raw: str) -> tuple[list[dict], dict]:
    if "Тональность" not in wb.sheetnames:
        return [], {"method": "missing_daily_tonality"}
    ws = wb["Тональность"]
    header_row, headers = _find_header(ws, ["Дата", "Позитив", "Негатив"], max_rows=80)
    date_idx = _header_index(headers, "Дата")
    pos_idx = _header_index(headers, "Позитив", "Позитивные")
    neu_idx = _header_index(headers, "Нейтрально", "Нейтральные", "Нейтрал")
    neg_idx = _header_index(headers, "Негатив", "Негативные")
    raw = {}
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        values = list(row)
        date = _date_key(values[date_idx] if date_idx is not None and date_idx < len(values) else "")
        if not date or date.lower() == "всего":
            continue
        raw[date] = {
            "date": date,
            "positive": _clean_int(values[pos_idx] if pos_idx is not None and pos_idx < len(values) else 0),
            "neutral": _clean_int(values[neu_idx] if neu_idx is not None and neu_idx < len(values) else 0),
            "negative": _clean_int(values[neg_idx] if neg_idx is not None and neg_idx < len(values) else 0),
        }
    dates = _period_dates(period_raw) or sorted(raw.keys(), key=_date_sort_key)
    daily_rows = []
    for date in dates:
        item = raw.get(date, {"date": date, "positive": 0, "neutral": 0, "negative": 0})
        item["total"] = item["positive"] + item["neutral"] + item["negative"]
        daily_rows.append(item)
    return daily_rows, {"method": "fallback_tonality_sheet_total_query"}


def _period_from_wb(wb) -> str:
    for sheet_name in ["Сводные данные", "Содержание", "Тональность"]:
        if sheet_name in wb.sheetnames:
            period = _period_label_from_content(wb[sheet_name])
            if period:
                return period
    return ""


def _peak_day(daily_rows: list[dict]) -> dict:
    return max(daily_rows or [], key=lambda item: item.get("total", 0), default={})


def _build_signals(project_row: dict, daily_rows: list[dict], method_meta: dict) -> dict:
    total = int(project_row.get("total", 0) or 0)
    positive = int(project_row.get("positive", 0) or 0)
    neutral = int(project_row.get("neutral", 0) or 0)
    negative = int(project_row.get("negative", 0) or 0)
    return {
        "dominant_tone": max(TONE_KEYS, key=lambda key: project_row.get(key, 0) or 0) if total else "",
        "positive_share": _safe_pct(positive, total),
        "neutral_share": _safe_pct(neutral, total),
        "negative_share": _safe_pct(negative, total),
        "peak_day": _peak_day(daily_rows),
        "top_sources": method_meta.get("top_sources", {}),
        "top_sites": method_meta.get("top_sites", {}),
        "top_message_types": method_meta.get("top_message_types", {}),
        "samples": method_meta.get("samples", {}),
    }


def _project_sheet_with_month(wb, month_sheet: str):
    if month_sheet and month_sheet in wb.sheetnames:
        return wb[month_sheet], month_sheet
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 35), values_only=True):
            values = [str(value or "").strip().lower() for value in row]
            if any("тип размещения" in value for value in values) and any("ссылка" in value or "площадка" in value for value in values):
                return ws, name
    return None, ""


def _compact_section_label(value) -> str:
    text = str(value or "").strip().lower().replace("ё", "е")
    return re.sub(r"[^а-яa-z0-9]+", "", text)


def _is_project_section_header(values: list) -> bool:
    first = _compact_section_label(values[0] if values else "")
    if not first:
        return False
    section_labels = {
        "отзывы",
        "реагирования",
        "топ",
        "негатив",
        "мониторинг",
        "резерв",
        "фото",
        "встраивания",
        "комментарии",
        "посев",
    }
    if first in section_labels:
        return True
    raw_first = str(values[0] or "").strip().lower().replace("ё", "е") if values else ""
    if "тип размещения" in raw_first or "период" == raw_first or "ключевые сообщения" in raw_first:
        return True
    non_empty = [str(value or "").strip() for value in values if str(value or "").strip()]
    has_url = any("http" in item.lower() for item in non_empty)
    is_compact_section = first in section_labels or any(first.startswith(label) for label in section_labels if len(label) > 3)
    return bool(non_empty) and len(non_empty) <= 3 and not has_url and is_compact_section


def _is_project_publication_row(values: list) -> bool:
    if not values or _is_project_section_header(values):
        return False
    first = str(values[0] or "").strip().lower()
    if not first or "тип размещения" in first:
        return False
    joined = " ".join(str(value or "") for value in values).strip()
    if not joined:
        return False
    has_url = "http" in joined.lower()
    has_status = any(str(value or "").strip().lower().replace("ё", "е") in {"размещено", "опубликовано"} for value in values)
    return has_url or has_status or len(joined) > 80


def _extract_project_negative_count(project_orm_path: Path | None, month_sheet: str = "") -> dict:
    if not project_orm_path or not project_orm_path.exists():
        return {"status": "missing_project_orm"}
    try:
        wb = open_workbook_safe(project_orm_path)
        ws, sheet_name = _project_sheet_with_month(wb, month_sheet)
        if ws is None:
            return {"status": "missing_month_sheet"}
        negative_count = 0
        section_found = False
        inside_negative = False
        for row in ws.iter_rows(values_only=True):
            values = list(row)
            first = _compact_section_label(values[0] if values else "")
            if "негатив" in first and _is_project_section_header(values):
                section_found = True
                inside_negative = True
                continue
            if inside_negative and _is_project_section_header(values):
                inside_negative = False
            if inside_negative and _is_project_publication_row(values):
                negative_count += 1
        return {
            "status": "ready" if section_found else "section_not_found",
            "negative_count": negative_count,
            "source_sheet": sheet_name,
        }
    except Exception as exc:
        return {"status": "error", "error": str(exc)}


def _apply_project_negative_override(
    brand_rows: list[dict],
    project_row: dict,
    project_brand: str,
    override_meta: dict,
) -> tuple[list[dict], dict]:
    status = override_meta.get("status")
    if status != "ready":
        return brand_rows, project_row
    new_negative = max(int(override_meta.get("negative_count") or 0), 0)
    if not project_row:
        return brand_rows, project_row
    old_negative = int(project_row.get("negative") or 0)
    if new_negative == old_negative:
        return brand_rows, project_row

    def adjusted(row: dict) -> dict:
        total = int(row.get("total") or 0)
        positive = int(row.get("positive") or 0)
        neutral = int(row.get("neutral") or 0)
        negative = min(new_negative, max(total - positive, 0)) if total else new_negative
        neutral = max(total - positive - negative, 0) if total else max(neutral + old_negative - negative, 0)
        out = dict(row)
        out.update({
            "positive": positive,
            "neutral": neutral,
            "negative": negative,
            "positive_share": _safe_pct(positive, total),
            "neutral_share": _safe_pct(neutral, total),
            "negative_share": _safe_pct(negative, total),
            "negative_source": "project_orm_negative_section",
        })
        return out

    project_norm = _norm_brand_name(project_brand)
    new_project_row = adjusted(project_row)
    new_brand_rows = []
    for row in brand_rows:
        if _norm_brand_name(row.get("brand", "")) == project_norm:
            new_brand_rows.append(dict(new_project_row, brand=row.get("brand") or project_brand))
        else:
            new_brand_rows.append(row)
    return new_brand_rows, new_project_row


def extract_slide06_tonality(
    *,
    analytics_path: Path,
    project_brand: str,
    source_system: str = "",
    competitor_brands: list[str] | None = None,
    project_orm_path: Path | None = None,
    month_sheet: str = "",
) -> dict:
    wb = open_workbook_safe(analytics_path)
    period_raw = _period_from_wb(wb)
    month_prep, month_caption, month_title = _period_to_caption(period_raw)
    brand_rows, project_row = _extract_brand_tonality(wb, project_brand, competitor_brands)
    daily_rows, method_meta = _extract_daily_project_tonality(wb, project_brand, period_raw)
    orm_negative_meta = _extract_project_negative_count(project_orm_path, month_sheet)
    brand_rows, project_row = _apply_project_negative_override(brand_rows, project_row, project_brand, orm_negative_meta)
    if orm_negative_meta.get("status") == "ready":
        method_meta["project_orm_negative_override"] = orm_negative_meta
    daily_total = sum(int(row.get("total", 0) or 0) for row in daily_rows)
    project_total = int(project_row.get("total", 0) or 0)
    totals_match = daily_total == project_total
    ready = bool(project_row and daily_rows)

    return {
        "source_file": analytics_path.name,
        "source_system": source_system or "Brand Analytics",
        "source_label": _source_label(source_system or "Brand Analytics"),
        "project_brand": project_brand,
        "period_raw": period_raw,
        "month_prepositional": month_prep,
        "month_caption": month_caption,
        "month_title": month_title,
        "brand_rows": brand_rows,
        "project_row": project_row,
        "daily_rows": daily_rows,
        "daily_method": method_meta,
        "project_orm_negative_override": orm_negative_meta,
        "totals_check": {
            "project_total_from_tags": project_total,
            "project_total_from_messages_daily": daily_total,
            "matches": totals_match,
            "method": method_meta.get("method", ""),
        },
        "signals": _build_signals(project_row, daily_rows, method_meta),
        "caveats": [] if totals_match else ["Дневная динамика по листу сообщений не сходится с тональностью по тегам."],
        "methodology_status": "ready" if ready and totals_match else "blocked",
    }
