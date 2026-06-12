
from __future__ import annotations

import re
from pathlib import Path

from .xlsx_safe_reader import open_workbook_safe


def _clean_int(value):
    if value in (None, ""):
        return 0
    if isinstance(value, str):
        value = value.replace("\n", "").replace(" ", "").replace("\xa0", "").replace("—", "0").replace("-", "0")
    try:
        return int(float(value))
    except Exception:
        return 0


def _clean_float(value):
    if value in (None, ""):
        return 0.0
    if isinstance(value, str):
        value = value.replace("%", "").replace(",", ".").replace(" ", "").replace("\xa0", "")
    try:
        f = float(value)
        return f / 100 if f > 1 else f
    except Exception:
        return 0.0


def _period_label_from_content(ws):
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 12), values_only=True):
        row_vals = list(row)
        if len(row_vals) >= 3 and str(row_vals[1] or "").strip().lower() == "период:":
            return str(row_vals[2] or "")
        for value in row_vals:
            text = str(value or "").strip()
            if text.lower().startswith("период:"):
                return text.split(":", 1)[1].strip()
    return ""


def _period_to_caption(period: str) -> tuple[str, str, str]:
    low = (period or "").lower()
    date_match = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{2,4}).*?(\d{1,2})\.\2\.\3", low)
    if date_match:
        start_day, month_num, year_raw, end_day = date_match.groups()
        month_names = {
            "01": ("январе", "января", "ЯНВАРЕ"),
            "02": ("феврале", "февраля", "ФЕВРАЛЕ"),
            "03": ("марте", "марта", "МАРТЕ"),
            "04": ("апреле", "апреля", "АПРЕЛЕ"),
            "05": ("мае", "мая", "МАЕ"),
            "06": ("июне", "июня", "ИЮНЕ"),
            "07": ("июле", "июля", "ИЮЛЕ"),
            "08": ("августе", "августа", "АВГУСТЕ"),
            "09": ("сентябре", "сентября", "СЕНТЯБРЕ"),
            "10": ("октябре", "октября", "ОКТЯБРЕ"),
            "11": ("ноябре", "ноября", "НОЯБРЕ"),
            "12": ("декабре", "декабря", "ДЕКАБРЕ"),
        }
        month_key = month_num.zfill(2)
        if month_key in month_names:
            prep, genitive, title = month_names[month_key]
            year = year_raw if len(year_raw) == 4 else "20" + year_raw
            return prep, f"{int(start_day)}–{int(end_day)} {genitive} {year}", title
    word_range = re.search(
        r"(\d{1,2})\s*[–—-]\s*(\d{1,2})\s+"
        r"(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)"
        r"\s+(\d{4})",
        low,
    )
    if word_range:
        start_day, end_day, month_word, year = word_range.groups()
        month_names_by_word = {
            "января": ("январе", "ЯНВАРЕ"),
            "февраля": ("феврале", "ФЕВРАЛЕ"),
            "марта": ("марте", "МАРТЕ"),
            "апреля": ("апреле", "АПРЕЛЕ"),
            "мая": ("мае", "МАЕ"),
            "июня": ("июне", "ИЮНЕ"),
            "июля": ("июле", "ИЮЛЕ"),
            "августа": ("августе", "АВГУСТЕ"),
            "сентября": ("сентябре", "СЕНТЯБРЕ"),
            "октября": ("октябре", "ОКТЯБРЕ"),
            "ноября": ("ноябре", "НОЯБРЕ"),
            "декабря": ("декабре", "ДЕКАБРЕ"),
        }
        prep, title = month_names_by_word[month_word]
        return prep, f"{int(start_day)}–{int(end_day)} {month_word} {year}", title
    mapping = [
        ("январ", "январе", "январь"), ("феврал", "феврале", "февраль"),
        ("март", "марте", "март"), ("апрел", "апреле", "апрель"),
        ("май", "мае", "май"), ("июн", "июне", "июнь"),
        ("июл", "июле", "июль"), ("август", "августе", "август"),
        ("сентябр", "сентябре", "сентябрь"), ("октябр", "октябре", "октябрь"),
        ("ноябр", "ноябре", "ноябрь"), ("декабр", "декабре", "декабрь"),
    ]
    year = "2026" if "2026" in low else "2025" if "2025" in low else ""
    for key, prep, nomin in mapping:
        if key in low:
            return prep, f"{nomin} {year}".strip(), nomin.upper()
    return "периоде", period or "период", "ПЕРИОДЕ"


def _find_metrics_sheet(wb):
    candidates = ["Показатели брендов", "Показатели", "Теги", "Продукты"]
    for name in candidates:
        if name in wb.sheetnames:
            try:
                _find_header(wb[name])
                return wb[name], name
            except ValueError:
                continue
    # fallback: sheet with object/messages columns
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 15), values_only=True):
            vals = [str(v or "").strip().lower() for v in row]
            if ("объект" in vals or "тег" in vals or "" in vals) and "сообщения" in vals:
                return ws, name
    raise ValueError("Metrics sheet with brand indicators not found")


def _find_header(ws):
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 80), values_only=True), start=1):
        vals = list(row)
        lowered = [str(v or "").strip().lower() for v in vals]
        if ("объект" in lowered or "" in lowered) and "сообщения" in lowered and "аудитория" in lowered:
            return idx, vals
        if ("продукт" in lowered or "тег" in lowered) and "сообщения" in lowered:
            return idx, vals
    raise ValueError("Metrics header not found")


def _get_index(headers, *names, default=None):
    lowered = {str(h or "").strip().lower(): i for i, h in enumerate(headers)}
    for name in names:
        if name.lower() in lowered:
            return lowered[name.lower()]
    return default


def _is_thematic_metric_label(value: str) -> bool:
    low = str(value or "").strip().lower().replace("ё", "е")
    return (
        not low
        or "тематика" in low
        or "тематичес" in low
        or "обсуждение" in low
        or "детских ноотроп" in low
        or "категори" in low
        or "без брендов" in low
        or "поисков" in low
        or "запрос" in low
        or "не считаем" in low
    )


def _brand_from_product_name(value: str, project_brand: str = "", brand_names: list[str] | None = None) -> str:
    text = str(value or "").strip()
    low = text.lower().replace("ё", "е")
    candidates = [project_brand] + list(brand_names or [])
    for brand in candidates:
        needle = str(brand or "").strip().lower().replace("ё", "е")
        if needle and needle in low:
            return brand
    return text[:60]


def _aggregate_brand_rows(rows: list[dict]) -> list[dict]:
    aggregated: dict[str, dict] = {}
    for row in rows:
        brand = row.get("brand") or ""
        if not brand:
            continue
        target = aggregated.setdefault(brand, {"brand": brand, "mentions": 0, "sm_index": 0, "audience": 0, "views": 0, "engagement": 0, "comments": 0, "likes": 0, "reposts": 0, "sov": 0})
        for key in ["mentions", "sm_index", "audience", "views", "engagement", "comments", "likes", "reposts"]:
            target[key] += int(row.get(key, 0) or 0)
    return list(aggregated.values())


def _norm_brand_name(value: str) -> str:
    text = str(value or "").strip().lower().replace("ё", "е")
    text = re.sub(r"[®™]", "", text)
    return re.sub(r"\s+", " ", text)


def _brand_name_matches(value: str, target: str) -> bool:
    left = _norm_brand_name(value)
    right = _norm_brand_name(target)
    if not left or not right:
        return False
    if left == right:
        return True
    return min(len(left), len(right)) >= 4 and (left in right or right in left)


def _canonical_brand_name(value: str, project_brand: str = "", competitor_brands: list[str] | None = None) -> str:
    if _brand_name_matches(value, project_brand):
        return project_brand
    for brand in competitor_brands or []:
        if _brand_name_matches(value, brand):
            return str(brand or "").strip()
    return str(value or "").strip()


def _filter_competitive_rows(rows: list[dict], project_brand: str = "", competitor_brands: list[str] | None = None) -> list[dict]:
    competitors = [str(item).strip() for item in (competitor_brands or []) if str(item or "").strip()]
    allowed_names = [str(project_brand or "").strip(), *competitors]
    allowed_names = [item for item in allowed_names if item]
    canonical_rows = []
    for row in rows:
        item = dict(row)
        item["brand"] = _canonical_brand_name(item.get("brand", ""), project_brand, competitors)
        canonical_rows.append(item)
    if not competitors:
        return _aggregate_brand_rows(canonical_rows)
    filtered = [
        row
        for row in canonical_rows
        if any(_brand_name_matches(row.get("brand", ""), allowed) for allowed in allowed_names)
    ]
    return _aggregate_brand_rows(filtered)


def extract_brand_mentions_sov(
    analytics_path: Path,
    project_brand: str = "",
    top_n: int = 6,
    competitor_brands: list[str] | None = None,
) -> dict:
    wb = open_workbook_safe(analytics_path)
    content_ws = wb["Содержание"] if "Содержание" in wb.sheetnames else wb[wb.sheetnames[0]]
    period_raw = _period_label_from_content(content_ws)
    month_prep, month_caption, month_title = _period_to_caption(period_raw)

    ws, sheet_name = _find_metrics_sheet(wb)
    header_row, headers = _find_header(ws)
    object_idx = _get_index(headers, "Объект", "Тег", "Продукт", default=1)
    mentions_idx = _get_index(headers, "Сообщения")
    audience_idx = _get_index(headers, "Аудитория")
    sm_idx = _get_index(headers, "СМ Индекс", default=None)
    views_idx = _get_index(headers, "Просмотры", default=None)
    engagement_idx = _get_index(headers, "Вовлечённость", default=None)
    comments_idx = _get_index(headers, "Комментарии", default=None)
    likes_idx = _get_index(headers, "Лайки", default=None)
    reposts_idx = _get_index(headers, "Репосты", default=None)
    sov_idx = _get_index(headers, "Share of voice", "SOV", "Доля голоса", default=None)

    if mentions_idx is None:
        raise ValueError("Required metrics column 'Сообщения' not found")

    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        values = list(row)
        brand = values[object_idx] if object_idx is not None and len(values) > object_idx else None
        if not brand:
            continue
        brand = str(brand).strip()
        if str(sheet_name).lower() == "продукты":
            brand = _brand_from_product_name(brand, project_brand)
        if brand.lower() in {"дата", "итого"}:
            break
        if brand.lower() == "всего":
            continue
        if _is_thematic_metric_label(brand):
            continue
        mentions = _clean_int(values[mentions_idx] if len(values) > mentions_idx else 0)
        if mentions == 0:
            continue
        row_item = {
            "brand": brand,
            "mentions": mentions,
            "sm_index": _clean_int(values[sm_idx] if sm_idx is not None and len(values) > sm_idx else 0),
            "audience": _clean_int(values[audience_idx] if audience_idx is not None and len(values) > audience_idx else 0),
            "views": _clean_int(values[views_idx] if views_idx is not None and len(values) > views_idx else 0),
            "engagement": _clean_int(values[engagement_idx] if engagement_idx is not None and len(values) > engagement_idx else 0),
            "comments": _clean_int(values[comments_idx] if comments_idx is not None and len(values) > comments_idx else 0),
            "likes": _clean_int(values[likes_idx] if likes_idx is not None and len(values) > likes_idx else 0),
            "reposts": _clean_int(values[reposts_idx] if reposts_idx is not None and len(values) > reposts_idx else 0),
            "sov": _clean_float(values[sov_idx] if sov_idx is not None and len(values) > sov_idx else 0),
        }
        rows.append(row_item)

    if str(sheet_name).lower() == "продукты":
        rows = _aggregate_brand_rows(rows)

    rows = _filter_competitive_rows(rows, project_brand, competitor_brands)
    total_mentions = sum(r["mentions"] for r in rows)
    # Comparative Medialogia reports may have empty Share of voice column: compute from messages.
    for row in rows:
        if not row.get("sov") and total_mentions:
            row["sov"] = row["mentions"] / total_mentions

    rows_by_mentions = sorted(rows, key=lambda x: x["mentions"], reverse=True)
    rows_by_audience = sorted(rows, key=lambda x: x["audience"], reverse=True)
    top_mentions = rows_by_mentions[:top_n]
    top_audience = rows_by_audience[:top_n]

    project = next((r for r in rows if project_brand and r["brand"].lower() == project_brand.lower()), None)
    rank_mentions = next((i + 1 for i, r in enumerate(rows_by_mentions) if project and r["brand"] == project["brand"]), None)
    rank_audience = next((i + 1 for i, r in enumerate(rows_by_audience) if project and r["brand"] == project["brand"]), None)
    leader_mentions = rows_by_mentions[0] if rows_by_mentions else {}
    leader_audience = rows_by_audience[0] if rows_by_audience else {}
    gap_to_mentions_leader = round(leader_mentions.get("mentions", 0) / project["mentions"], 1) if project and project.get("mentions") else None

    return {
        "source_file": analytics_path.name,
        "source_sheet": sheet_name,
        "project_brand": project_brand,
        "period_raw": period_raw,
        "month_prepositional": month_prep,
        "month_caption": month_caption,
        "month_title": month_title,
        "methodology_status": "ready_with_caveat",
        "sov_methodology_note": "SOV рассчитан по сообщениям сравнительного отчета, если Share of voice в выгрузке не заполнен.",
        "total_mentions_competitive_set": total_mentions,
        "all_brand_rows": rows,
        "top_brands": top_mentions,
        "top_audience_brands": top_audience,
        "project_brand_row": project,
        "project_brand_rank": rank_mentions,
        "project_audience_rank": rank_audience,
        "mentions_leader": leader_mentions,
        "audience_leader": leader_audience,
        "gap_to_mentions_leader": gap_to_mentions_leader,
        "chart_mentions": {
            "categories": [r["brand"] for r in top_mentions],
            "mentions": [r["mentions"] for r in top_mentions],
        },
        "chart_audience": {
            "categories": [r["brand"] for r in top_audience],
            "audience": [r["audience"] for r in top_audience],
        },
    }


# ---------- v0.3.13: analytical context for reasoning ----------

def _header_match(value: str, required: str) -> bool:
    v = str(value or "").strip().lower()
    r = required.lower()
    # Exact match is safest. Allow structured header variants like "Теги / Показатель".
    return v == r or v.startswith(r + " /")


def _find_header_row_by_columns(ws, required_names, max_rows=80):
    required = [r.lower() for r in required_names]
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows), values_only=True), start=1):
        vals = [str(v or "").strip().lower() for v in row]
        if all(any(_header_match(v, req) for v in vals) for req in required):
            return idx, list(row)
    return None, []


def _idx(headers, *names, default=None):
    low = {str(h or "").strip().lower(): i for i, h in enumerate(headers)}
    for name in names:
        name = name.lower()
        for h, i in low.items():
            if h == name or name in h:
                return i
    return default


def _rows_from_table(ws, header_row, headers, max_take=10, numeric_col="Сообщения"):
    if not header_row:
        return []
    name_i = 1 if len(headers) > 1 else 0
    # Better semantic indices
    for cand in ["Площадка", "Источник", "Тематика", "Слова", "Слово", "Теги", "Тег", "Группа", "Тип"]:
        i = _idx(headers, cand, default=None)
        if i is not None:
            name_i = i
            break
    msg_i = _idx(headers, numeric_col, "Сообщения", default=None)
    aud_i = _idx(headers, "Аудитория", default=None)
    type_i = _idx(headers, "Тип", default=None)
    out = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        vals = list(row)
        name = vals[name_i] if name_i < len(vals) else None
        if not name or str(name).strip().lower() == "итого":
            continue
        name_text = str(name).strip()
        name_low = name_text.lower()
        if name_low == "дата":
            break
        if name_low == "всего" or "не считаем" in name_low or name_low.startswith("тематические запросы"):
            continue
        if re.fullmatch(r"\d+([.,]\d+)?", name_low) or re.fullmatch(r"\d{1,2}[.]\d{1,2}[.]\d{4}", name_low):
            continue
        messages = _clean_int(vals[msg_i] if msg_i is not None and msg_i < len(vals) else 0)
        audience = _clean_int(vals[aud_i] if aud_i is not None and aud_i < len(vals) else 0)
        typ = str(vals[type_i]).strip() if type_i is not None and type_i < len(vals) and vals[type_i] else ""
        if messages == 0 and audience == 0:
            continue
        out.append({"name": name_text, "type": typ, "messages": messages, "audience": audience})
    out = sorted(out, key=lambda x: (x["messages"], x["audience"]), reverse=True)
    return out[:max_take]


def _meaningful_words(rows):
    stop = {"отзыв", "оценка", "день", "раз", "ответ", "посол", "лета", "цена", "руб"}
    return [r for r in rows if r["name"].lower() not in stop][:8]


def extract_analytical_context(analytical_path: Path) -> dict:
    """Read a Medialogia analytical report and extract causal/context signals for slide 03.

    This function does not replace comparative metrics. It provides the why-layer for the final insight:
    themes, platforms, platform types, words, tags and source concentration.
    """
    wb = open_workbook_safe(analytical_path)
    content_ws = wb["Содержание"] if "Содержание" in wb.sheetnames else wb[wb.sheetnames[0]]
    period_raw = _period_label_from_content(content_ws)

    result = {
        "source_file": analytical_path.name,
        "source_system": "Медиалогия",
        "period_raw": period_raw,
        "themes": [],
        "platforms": [],
        "platform_types": [],
        "words": [],
        "tags": [],
        "reason_signals": [],
    }

    if "Тематики" in wb.sheetnames:
        ws = wb["Тематики"]
        hr, headers = _find_header_row_by_columns(ws, ["Тематика", "Сообщения"])
        result["themes"] = _rows_from_table(ws, hr, headers, max_take=8)

    platform_sheet = "Площадки" if "Площадки" in wb.sheetnames else "Источники" if "Источники" in wb.sheetnames else ""
    if platform_sheet:
        ws = wb[platform_sheet]
        required = ["Площадка", "Сообщения"] if platform_sheet == "Площадки" else ["Источник", "Сообщения"]
        hr, headers = _find_header_row_by_columns(ws, required)
        result["platforms"] = _rows_from_table(ws, hr, headers, max_take=8)

    if "Типы площадок" in wb.sheetnames:
        ws = wb["Типы площадок"]
        # This sheet has values as columns in one total row.
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
            vals = list(row)
            names = [str(v or "").strip() for v in vals]
            if "Соцсети" in names and "Всего" in names:
                header = names
                # next row is usually totals
                next_row = next(ws.iter_rows(min_row=row[0].row + 1, max_row=row[0].row + 1, values_only=True), None) if False else None
                break
        # fallback simple parse by visible rows
        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), start=1):
            vals = [str(v or "").strip() for v in row]
            if "Соцсети" in vals and "Всего" in vals:
                headers = vals
                total_vals = list(next(ws.iter_rows(min_row=r_idx + 1, max_row=r_idx + 1, values_only=True)))
                items = []
                for i, h in enumerate(headers):
                    if h and h != "Всего":
                        items.append({"name": h, "messages": _clean_int(total_vals[i] if i < len(total_vals) else 0), "audience": 0})
                result["platform_types"] = sorted(items, key=lambda x: x["messages"], reverse=True)[:8]
                break

    words_sheet = "Слова" if "Слова" in wb.sheetnames else "Популярные слова" if "Популярные слова" in wb.sheetnames else ""
    if words_sheet:
        ws = wb[words_sheet]
        required = ["Слова", "Сообщения"] if words_sheet == "Слова" else ["Слово", "Сообщения"]
        hr, headers = _find_header_row_by_columns(ws, required)
        result["words"] = _meaningful_words(_rows_from_table(ws, hr, headers, max_take=20))

    if "Теги" in wb.sheetnames:
        ws = wb["Теги"]
        hr, headers = _find_header_row_by_columns(ws, ["Теги", "Сообщения"])
        if not hr:
            hr, headers = _find_header_row_by_columns(ws, ["Тег", "Сообщения"])
        result["tags"] = _rows_from_table(ws, hr, headers, max_take=8)

    # Build concise reason signals from extracted data.
    top_platforms = [p for p in result["platforms"] if p.get("messages", 0) > 0][:3]
    top_types = [p for p in result["platform_types"] if p.get("messages", 0) > 0][:3]
    relevant_themes = [t for t in result["themes"] if t["name"].lower() not in {"прочее", "юмор", "знакомство и общение"}][:3]
    words = [w["name"] for w in result["words"][:5]]
    tags = [t for t in result["tags"] if t.get("messages", 0) > 0][:3]

    if top_platforms:
        result["reason_signals"].append({
            "type": "platforms",
            "text": "ключевые площадки: " + ", ".join([p["name"] for p in top_platforms])
        })
    if top_types:
        result["reason_signals"].append({
            "type": "platform_types",
            "text": "основные типы площадок: " + ", ".join([p["name"] for p in top_types])
        })
    if relevant_themes:
        result["reason_signals"].append({
            "type": "themes",
            "text": "заметные темы: " + ", ".join([t["name"] for t in relevant_themes])
        })
    if words:
        result["reason_signals"].append({
            "type": "words",
            "text": "частые контексты: " + ", ".join(words)
        })
    if tags:
        result["reason_signals"].append({
            "type": "tags",
            "text": "теги / сегменты: " + ", ".join([t["name"] for t in tags])
        })

    return result
