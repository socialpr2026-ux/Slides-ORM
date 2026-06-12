from __future__ import annotations

from collections import Counter
from pathlib import Path

from .slide03_data_extractor import _clean_int, _is_thematic_metric_label, _period_label_from_content, _period_to_caption
from .slide05_sov_extractor import _source_label
from .xlsx_safe_reader import open_workbook_safe


SOURCE_TYPE_ORDER = ["Соцсети", "Мессенджеры", "Отзывы", "Видео", "Другое", "Блоги", "Форумы"]


def _lower(value) -> str:
    return str(value or "").strip().lower()


def _header_index(headers, *names: str):
    lowered = [_lower(item) for item in headers]
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


def _row_mentions_project(values: list, project_brand: str, candidate_indices: list[int | None]) -> bool:
    needle = _lower(project_brand)
    if not needle:
        return False
    for idx in candidate_indices:
        if idx is None or idx >= len(values):
            continue
        if needle in _lower(values[idx]):
            return True
    return False


def _column_mentions_project(ws, header_row: int, column_idx: int, project_brand: str, max_scan_rows: int = 500) -> bool:
    needle = _lower(project_brand)
    if not needle:
        return False
    max_row = min(ws.max_row, header_row + max_scan_rows)
    for row in ws.iter_rows(min_row=header_row + 1, max_row=max_row, values_only=True):
        values = list(row)
        if column_idx < len(values) and needle in _lower(values[column_idx]):
            return True
    return False


def _project_object_filter_indices(ws, header_row: int, headers: list, project_brand: str) -> list[int]:
    if not project_brand:
        return []
    exact_names = {"объект", "объекты", "объект ba", "объекты ba", "бренд", "бренды", "тег", "теги"}
    excluded_fragments = ("роль", "role", "тип", "type", "тональность", "sentiment", "категор", "category")
    candidates: list[int] = []
    for idx, header in enumerate(headers):
        label = _lower(header).replace("ё", "е")
        if not label:
            continue
        is_exact = label in exact_names
        is_safe_object = ("объект" in label or "тег" in label) and not any(fragment in label for fragment in excluded_fragments)
        if is_exact or is_safe_object:
            candidates.append(idx)
    validated = [
        idx for idx in dict.fromkeys(candidates)
        if _column_mentions_project(ws, header_row, idx, project_brand)
    ]
    return validated


def _period_from_wb(wb) -> str:
    for sheet_name in ["Сводные данные", "Содержание", "Источники"]:
        if sheet_name in wb.sheetnames:
            period = _period_label_from_content(wb[sheet_name])
            if period:
                return period
    return ""


def _fmt_int(value: int | float) -> str:
    return f"{int(round(float(value))):,}".replace(",", " ")


def _fmt_pct(value: float, digits: int = 1) -> str:
    return f"{float(value) * 100:.{digits}f}%".replace(".", ",")


def _messages_word(value: int | float) -> str:
    number = abs(int(round(float(value))))
    if 11 <= number % 100 <= 14:
        return "сообщений"
    last = number % 10
    if last == 1:
        return "сообщение"
    if 2 <= last <= 4:
        return "сообщения"
    return "сообщений"


def _join_ru(items: list[str]) -> str:
    clean = [str(item).strip() for item in items if str(item or "").strip()]
    if not clean:
        return ""
    if len(clean) == 1:
        return clean[0]
    return ", ".join(clean[:-1]) + " и " + clean[-1]


def _map_source_type(value: str) -> str:
    text = _lower(value)
    if "соцсет" in text or "микроблог" in text:
        return "Соцсети"
    if "мессенджер" in text:
        return "Мессенджеры"
    if "отзыв" in text:
        return "Отзывы"
    if "блог" in text:
        return "Блоги"
    if "форум" in text:
        return "Форумы"
    if "видео" in text:
        return "Видео"
    if "сми" in text or "медиа" in text:
        return "Другое"
    return "Другое"


def _source_display(value: str) -> str:
    text = str(value or "").strip()
    low = text.lower()
    normalized_low = low.replace("ё", "е")
    exact_mapping = {
        "озон": "Ozon",
        "sprosivracha.com - форум": "Sprosivracha.com",
        "мамлайф": "Mom.Life",
        "бебиблог (babyblog.ru)": "Babyblog.ru",
        "женский форум беби.ру": "Baby.ru",
        "likee - сообщество для коротких видео": "Likee",
        "женский журнал woman.ru": "Woman.ru",
        "medihost.ru - форум": "Medihost.ru",
        "большой вопрос - все вопросы и ответы!": "Bolshoyvopros.ru",
    }
    if normalized_low in exact_mapping:
        return exact_mapping[normalized_low]
    if "да-да новости" in normalized_low or "dadanews" in normalized_low:
        return "dadanews.ru"
    if "большой вопрос" in normalized_low or "bolshoyvopros" in normalized_low:
        return "Bolshoyvopros.ru"
    if "мамлайф" in normalized_low or "mom.life" in normalized_low:
        return "Mom.Life"
    if "бебиблог" in normalized_low or "babyblog" in normalized_low:
        return "Babyblog.ru"
    if "беби.ру" in normalized_low or "baby.ru" in normalized_low:
        return "Baby.ru"
    if "sprosivracha" in normalized_low:
        return "Sprosivracha.com"
    if "woman.ru" in normalized_low:
        return "Woman.ru"
    if "medihost" in normalized_low:
        return "Medihost.ru"
    if "likee" in normalized_low:
        return "Likee"
    if "wildberries" in normalized_low or "вайлдбер" in normalized_low:
        return "Wildberries"
    if normalized_low in {"ozon", "ozon.ru"}:
        return "Ozon"
    mapping = {
        "telegram.org": "Telegram",
        "telegram.me": "Telegram",
        "t.me": "Telegram",
        "vk.com": "ВКонтакте",
        "threads.com": "Threads",
        "instagram.com": "Instagram",
        "youtube.com": "YouTube",
        "dzen.ru": "Дзен",
        "ok.ru": "Одноклассники",
        "ozon.ru": "Ozon",
        "prodoctorov.ru": "ПроДокторов",
        "wildberries.ru": "Wildberries",
        "facebook.com": "Facebook",
        "mom.life": "Mom.Life",
        "apteka.ru": "Apteka.ru",
        "tiktok.com": "TikTok",
        "market.yandex.ru": "Яндекс.Маркет",
    }
    for needle, label in mapping.items():
        if needle in low:
            return label
    if low.startswith("http"):
        import re

        match = re.search(r"https?://(?:www\.)?([^/]+)", low)
        if match:
            return match.group(1)
    if "." in text:
        return text.split("/")[0]
    return text or "Не указано"


def _merge_source_rows_by_label(rows: list[dict]) -> list[dict]:
    merged: dict[str, dict] = {}
    for row in rows:
        label = _source_display(row.get("label") or row.get("source") or "")
        item = merged.setdefault(label, {"source": row.get("source", label), "label": label, "messages": 0, "positive": 0, "neutral": 0, "negative": 0})
        item["messages"] += int(row.get("messages", 0) or 0)
        item["positive"] += int(row.get("positive", 0) or 0)
        item["neutral"] += int(row.get("neutral", 0) or 0)
        item["negative"] += int(row.get("negative", 0) or 0)
    return sorted(merged.values(), key=lambda item: item["messages"], reverse=True)


def _platform_from_url_or_source(url: str, fallback: str = "") -> str:
    return _source_display(url or fallback)


def _short_text(value: str, limit: int = 38) -> str:
    text = " ".join(str(value or "").split())
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def _truthy_tag(value) -> bool:
    if value in (None, "", 0, "0"):
        return False
    if isinstance(value, str):
        text = value.strip().lower()
        return text not in {"", "0", "нет", "false"}
    return True


def _message_tag_columns(headers: list) -> tuple[list[int], list[int]]:
    processed_idx = _header_index(headers, "Обработано")
    start_idx = processed_idx + 1 if processed_idx is not None else 40
    brand_cols = []
    thematic_cols = []
    for idx, header in enumerate(headers):
        if idx < start_idx:
            continue
        label = str(header or "").strip()
        if not label:
            continue
        if _is_thematic_metric_label(label):
            thematic_cols.append(idx)
        else:
            brand_cols.append(idx)
    return brand_cols, thematic_cols


def _row_has_any(values: list, indexes: list[int]) -> bool:
    return any(idx < len(values) and _truthy_tag(values[idx]) for idx in indexes)


def _extract_messages_rollup(wb, project_brand: str = "") -> dict:
    if "Сообщения" not in wb.sheetnames:
        return {"total": 0, "type_counter": Counter(), "source_counter": Counter(), "review_sources": Counter(), "community_rows": [], "filter": {"scope": "all_messages", "status": "missing_messages_sheet"}}
    ws = wb["Сообщения"]
    header_row, headers = _find_messages_header(ws, max_rows=80)
    source_idx = _header_index(headers, "Источник", "Площадка", "Где пишет")
    type_idx = _header_index(headers, "Тип источника", "Тип площадки", "Тип")
    community_idx = _header_index(headers, "Место публикации", "Где пишет", "Площадка")
    community_url_idx = _header_index(headers, "Url места публикации", "URL места публикации", "Ссылка на блог", "Ссылка на сообщение")
    audience_idx = _header_index(headers, "Аудитория", "Аудитория блога", "Аудитория автора", "Просмотры")
    object_filter_indices = _project_object_filter_indices(ws, header_row, headers, project_brand)
    filter_by_project_object = bool(project_brand and object_filter_indices)
    brand_tag_cols, thematic_tag_cols = ([], []) if filter_by_project_object else _message_tag_columns(headers)
    filter_by_brand_tags = bool(brand_tag_cols)
    total = 0
    skipped_thematic_only = 0
    type_counter: Counter[str] = Counter()
    source_counter: Counter[str] = Counter()
    review_sources: Counter[str] = Counter()
    community_counter: dict[tuple[str, str], dict] = {}
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        values = list(row)
        if not any(values):
            continue
        if filter_by_project_object:
            if not _row_mentions_project(values, project_brand, object_filter_indices):
                continue
        elif filter_by_brand_tags:
            has_brand_tag = _row_has_any(values, brand_tag_cols)
            if not has_brand_tag:
                if _row_has_any(values, thematic_tag_cols):
                    skipped_thematic_only += 1
                continue
        total += 1
        source = str((values[source_idx] if source_idx is not None and source_idx < len(values) else "") or "").strip()
        raw_type = str((values[type_idx] if type_idx is not None and type_idx < len(values) else "") or source or "").strip()
        mapped_type = _map_source_type(raw_type)
        type_counter[mapped_type] += 1
        if source:
            source_counter[source] += 1
            if mapped_type == "Отзывы":
                review_sources[source] += 1
        if mapped_type in {"Соцсети", "Мессенджеры"}:
            community = str((values[community_idx] if community_idx is not None and community_idx < len(values) else "") or "").strip()
            url = str((values[community_url_idx] if community_url_idx is not None and community_url_idx < len(values) else "") or "").strip()
            if community and url:
                key = (community, url)
                item = community_counter.setdefault(key, {
                    "community": _short_text(community, 42),
                    "community_full": community,
                    "url": url,
                    "platform": _platform_from_url_or_source(url, source),
                    "messages": 0,
                    "audience": 0,
                })
                item["messages"] += 1
                item["audience"] = max(
                    int(item.get("audience", 0) or 0),
                    _clean_int(values[audience_idx] if audience_idx is not None and audience_idx < len(values) else 0),
                )
    community_rows = sorted(community_counter.values(), key=lambda item: item["messages"], reverse=True)[:10]
    return {
        "total": total,
        "type_counter": type_counter,
        "source_counter": source_counter,
        "review_sources": review_sources,
        "community_rows": community_rows,
        "filter": {
            "scope": "project_object_messages" if filter_by_project_object else ("brand_tagged_messages_excluding_thematic" if filter_by_brand_tags else "all_messages"),
            "brand_tag_columns": [str(headers[idx] or "").strip() for idx in brand_tag_cols],
            "thematic_tag_columns": [str(headers[idx] or "").strip() for idx in thematic_tag_cols],
            "project_object_columns": [str(headers[idx] or "").strip() for idx in object_filter_indices] if filter_by_project_object else [],
            "skipped_thematic_only_messages": skipped_thematic_only,
        },
    }


def _extract_sources_sheet(wb, total_messages: int, fallback_counter: Counter[str], *, prefer_sheet: bool = True) -> list[dict]:
    rows = []
    if prefer_sheet and "Источники" in wb.sheetnames:
        ws = wb["Источники"]
        header_row, headers = _find_header(ws, ["Источник", "Сообщения"], max_rows=30)
        source_idx = _header_index(headers, "Источник")
        messages_idx = _header_index(headers, "Сообщения")
        pos_idx = _header_index(headers, "Позитив")
        neu_idx = _header_index(headers, "Нейтрально")
        neg_idx = _header_index(headers, "Негатив")
        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
            values = list(row)
            first_non_empty = next((str(value).strip() for value in values if str(value or "").strip()), "")
            if first_non_empty.lower() == "дата":
                break
            source = str(values[source_idx] if source_idx is not None and source_idx < len(values) else "").strip()
            if not source:
                continue
            if source.lower() == "дата":
                break
            if source.isdigit():
                continue
            messages = _clean_int(values[messages_idx] if messages_idx is not None and messages_idx < len(values) else 0)
            if messages <= 0:
                continue
            rows.append({
                "source": source,
                "label": _source_display(source),
                "messages": messages,
                "positive": _clean_int(values[pos_idx] if pos_idx is not None and pos_idx < len(values) else 0),
                "neutral": _clean_int(values[neu_idx] if neu_idx is not None and neu_idx < len(values) else 0),
                "negative": _clean_int(values[neg_idx] if neg_idx is not None and neg_idx < len(values) else 0),
            })
    if not rows and fallback_counter:
        rows = [
            {"source": source, "label": _source_display(source), "messages": count}
            for source, count in fallback_counter.most_common(12)
        ]
    rows = _merge_source_rows_by_label(rows)
    rows = sorted(rows, key=lambda item: item["messages"], reverse=True)
    selected = rows[:10]
    selected_total = sum(item["messages"] for item in selected)
    other = max(int(total_messages or 0) - selected_total, 0)
    if other:
        selected.append({"source": "other", "label": "Другие", "messages": other})
    total = sum(item["messages"] for item in selected) or total_messages
    for item in selected:
        item["share"] = item["messages"] / total if total else 0
    return selected


def _extract_communities(wb, max_rows: int = 10) -> list[dict]:
    if "Сообщества" not in wb.sheetnames:
        return []
    ws = wb["Сообщества"]
    header_row, headers = _find_header(ws, ["Сообщество", "Сообщения", "Аудитория"], max_rows=30)
    community_idx = _header_index(headers, "Сообщество")
    url_idx = _header_index(headers, "Url сообщества", "URL сообщества")
    messages_idx = _header_index(headers, "Сообщения")
    audience_idx = _header_index(headers, "Аудитория")
    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        values = list(row)
        community = str(values[community_idx] if community_idx is not None and community_idx < len(values) else "").strip()
        if not community:
            continue
        messages = _clean_int(values[messages_idx] if messages_idx is not None and messages_idx < len(values) else 0)
        audience = _clean_int(values[audience_idx] if audience_idx is not None and audience_idx < len(values) else 0)
        url = str(values[url_idx] if url_idx is not None and url_idx < len(values) else "").strip()
        if messages <= 0 and audience <= 0:
            continue
        rows.append({
            "community": _short_text(community, 42),
            "community_full": community,
            "url": url,
            "platform": _platform_from_url_or_source(url),
            "messages": messages,
            "audience": audience,
        })
        if len(rows) >= max_rows:
            break
    return rows


def _type_rows(type_counter: Counter[str]) -> list[dict]:
    rows = []
    total = sum(type_counter.values())
    for name in SOURCE_TYPE_ORDER:
        messages = int(type_counter.get(name, 0))
        if messages <= 0:
            continue
        rows.append({"type": name, "messages": messages, "share": messages / total if total else 0})
    for name, messages in type_counter.most_common():
        if name not in SOURCE_TYPE_ORDER and messages > 0:
            rows.append({"type": name, "messages": int(messages), "share": messages / total if total else 0})
    return sorted(rows, key=lambda item: item["messages"], reverse=True)


def _build_insight(data: dict) -> str:
    total = int(data.get("total_messages", 0) or 0)
    types_by_name = {row["type"]: row for row in data.get("platform_type_rows", [])}
    social = int(types_by_name.get("Соцсети", {}).get("messages", 0) or 0)
    messengers = int(types_by_name.get("Мессенджеры", {}).get("messages", 0) or 0)
    reviews = int(types_by_name.get("Отзывы", {}).get("messages", 0) or 0)
    communication_core = social + messengers
    core_share = communication_core / total if total else 0
    review_share = reviews / total if total else 0
    top_sources = [row for row in data.get("source_rows", []) if row.get("label") != "Другие"][:2]
    top_source_count = sum(int(row.get("messages", 0) or 0) for row in top_sources)
    top_source_share = top_source_count / total if total else 0
    top_source_parts = []
    for row in top_sources:
        label = str(row.get("label", "")).strip()
        messages = int(row.get("messages", 0) or 0)
        share = messages / total if total else 0
        if label:
            top_source_parts.append(f"{label} ({_fmt_int(messages)} {_messages_word(messages)}, {_fmt_pct(share)})")
    top_source_text = " и ".join(top_source_parts) if top_source_parts else "ключевые источники"
    top_source_labels = {str(row.get("label", "")).strip() for row in top_sources}
    review_source_rows = [
        row
        for row in data.get("source_rows", [])
        if str(row.get("label", "")).strip()
        and row.get("label") != "Другие"
        and str(row.get("label", "")).strip() not in top_source_labels
    ][:2]
    review_source_labels = [str(row.get("label", "")).strip() for row in review_source_rows]
    review_source_count = sum(int(row.get("messages", 0) or 0) for row in review_source_rows)
    review_source_sentence = ""
    if review_source_labels and review_source_count:
        review_source_sentence = (
            f" Внутри отзывного контура заметны {_join_ru(review_source_labels)}: "
            f"{_fmt_int(review_source_count)} {_messages_word(review_source_count)} "
            f"({_fmt_pct(review_source_count / total if total else 0)})."
        )

    return (
        f"Чаще всего бренды упоминают в соцсетях и мессенджерах: эти каналы дают "
        f"{_fmt_int(communication_core)} {_messages_word(communication_core)} из {_fmt_int(total)} "
        f"({_fmt_pct(core_share)}). Видимость формируют {top_source_text}. Отзывы занимают "
        f"{_fmt_int(reviews)} {_messages_word(reviews)} ({_fmt_pct(review_share)}). "
        f"Соцсети создают знание и интерес, а отзывы работают ближе к покупке."
        f"{review_source_sentence}"
    )


def extract_slide07_sources(
    *,
    analytics_path: Path,
    project_brand: str,
    source_system: str = "",
) -> dict:
    wb = open_workbook_safe(analytics_path)
    period_raw = _period_from_wb(wb)
    month_prep, month_caption, month_title = _period_to_caption(period_raw)
    rollup = _extract_messages_rollup(wb, project_brand)
    total_messages = int(rollup["total"] or 0)
    platform_type_rows = _type_rows(rollup["type_counter"])
    filter_scope = (rollup.get("filter") or {}).get("scope", "all_messages")
    filtered_from_messages = filter_scope in {"brand_tagged_messages_excluding_thematic", "project_object_messages"}
    source_rows = _extract_sources_sheet(wb, total_messages, rollup["source_counter"], prefer_sheet=not filtered_from_messages)
    community_rows = rollup.get("community_rows") or ([] if filtered_from_messages else _extract_communities(wb))
    data = {
        "source_file": analytics_path.name,
        "source_system": source_system or "Brand Analytics",
        "source_label": _source_label(source_system or "Brand Analytics"),
        "project_brand": project_brand,
        "period_raw": period_raw,
        "month_prepositional": month_prep,
        "month_caption": month_caption,
        "month_title": month_title,
        "total_messages": total_messages,
        "platform_type_rows": platform_type_rows,
        "source_rows": source_rows,
        "community_rows": community_rows,
        "review_sources": rollup["review_sources"].most_common(6),
        "message_filter": rollup.get("filter") or {},
        "methodology": {
            "platform_types": "Brand-tagged messages excluding thematic-only rows grouped by Тип источника" if filtered_from_messages else "Сообщения grouped by Тип источника with reference categories",
            "top_sources": "Brand-tagged messages excluding thematic-only rows grouped by Источник" if filtered_from_messages else "Источники top by Сообщения with other bucket from total messages",
            "communities": "Brand-tagged social and messenger rows grouped by Место публикации" if filtered_from_messages else "Сообщества top by Сообщения",
        },
    }
    data["insight_text"] = _build_insight(data)
    data["methodology_status"] = "ready" if total_messages and platform_type_rows and source_rows else "blocked"
    return data
