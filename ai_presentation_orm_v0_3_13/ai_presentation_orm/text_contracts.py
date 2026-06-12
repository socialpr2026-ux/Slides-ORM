from __future__ import annotations

import re
from typing import Any


FORBIDDEN_MAIN_TEXT_PHRASES = (
    "Brand Analytics фиксирует",
    "категорийная тематика",
    "площадочная структура требует разделять роли каналов",
    "Вывод построен по",
    "по листу рейтингов",
    ".xlsx",
    ".pptx",
)


def as_int(value: Any, default: int = 0) -> int:
    if value is None or value == "":
        return default
    try:
        return int(round(float(value)))
    except Exception:
        return default


def as_float(value: Any, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    try:
        return float(value)
    except Exception:
        return default


def fmt_int(value: Any) -> str:
    return f"{as_int(value):,}".replace(",", " ")


def fmt_pct(value: Any, digits: int = 1) -> str:
    val = as_float(value)
    if abs(val) <= 1:
        val *= 100
    if digits == 0:
        return f"{val:.0f}%".replace(".", ",")
    return f"{val:.{digits}f}%".replace(".", ",")


def fmt_pp(value: Any, digits: int = 1, signed: bool = True) -> str:
    val = as_float(value)
    sign = "+" if signed and val > 0 else ""
    if digits == 0:
        body = f"{val:.0f}".replace(".", ",")
    else:
        body = f"{val:.{digits}f}".replace(".", ",")
    return f"{sign}{body} п.п."


def plural_ru(n: Any, one: str, few: str, many: str) -> str:
    n = abs(as_int(n))
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


def join_ru(items: list[str]) -> str:
    clean = [str(item).strip() for item in items if str(item or "").strip()]
    if not clean:
        return ""
    if len(clean) == 1:
        return clean[0]
    return ", ".join(clean[:-1]) + " и " + clean[-1]


def period_in(data: dict, default: str = "отчетном периоде") -> str:
    value = str(data.get("month_prepositional") or default).strip()
    if value.lower().startswith("в "):
        return value
    return f"в {value}"


def period_start(data: dict, default: str = "В отчетном периоде") -> str:
    text = period_in(data, default.lower())
    return text[:1].upper() + text[1:]


def rank_text(rank: Any) -> str:
    rank_i = as_int(rank)
    return f"{rank_i}-е место" if rank_i else "заметную позицию"


def _brand(data: dict, fallback: str = "бренд") -> str:
    return str(data.get("project_brand") or fallback).strip() or fallback


def _project_row(data: dict) -> dict:
    return data.get("project_brand_row") or data.get("project_row") or {}


def _top_competitor(rows: list[dict], brand: str) -> dict:
    brand_low = brand.lower()
    for row in rows:
        if str(row.get("brand", "")).strip().lower() != brand_low:
            return row
    return {}


def _message_word(n: Any) -> str:
    return plural_ru(n, "сообщение", "сообщения", "сообщений")


def _mention_word(n: Any) -> str:
    return plural_ru(n, "упоминание", "упоминания", "упоминаний")


def _publication_word(n: Any) -> str:
    return plural_ru(n, "публикация", "публикации", "публикаций")


def campaign_required_warning(split: dict) -> str:
    if split.get("status") == "verified":
        return ""
    return "Факт кампании требует проверки в проектной ORM-таблице перед финальной версией."


def slide03_category_insight(data: dict) -> str:
    brand = _brand(data)

    # Support canonical model data merge
    cm = data.get("canonical_model") if isinstance(data, dict) else None
    if cm is not None:
        tb = cm.get("target_brand", {})
        sov_cm = cm.get("competitive_sov_scope", {})
        sov_target = sov_cm.get("target", {})
        orm_camp = cm.get("orm_campaign", {})
        total = as_int(sov_cm.get("total_with_campaign"))
        mentions = as_int(tb.get("with_campaign"))
        sov = sov_target.get("sov_with_campaign")
        rank = data.get("project_brand_rank")
        competitors = sov_cm.get("competitors", [])
        leader = competitors[0] if competitors else {}
        campaign = as_int(orm_camp.get("total"))
        organic = as_int(tb.get("organic"))
    else:
        row = _project_row(data)
        if not row:
            return "Целевой бренд не сопоставлен с таблицей показателей. Нужно проверить название бренда и конкурентный набор."
        total = as_int(data.get("total_mentions_competitive_set"))
        mentions = as_int(row.get("mentions"))
        sov = row.get("sov")
        rank = data.get("project_brand_rank")
        leader = data.get("mentions_leader") or (data.get("top_brands") or [{}])[0]
        split = data.get("campaign_split") or {}
        campaign = as_int(split.get("campaign_publications_count") or split.get("campaign_mentions"))
        organic = as_int(split.get("organic_mentions"))

    first = (
        f"{period_start(data)} бренд с учетом кампании занял {rank_text(rank)}: "
        f"{fmt_int(mentions)} {_message_word(mentions)} и SOV {fmt_pct(sov)}."
    )
    if total:
        first = (
            f"{period_start(data)} в конкурентном инфополе набора выделено {fmt_int(total)} "
            f"{_mention_word(total)}. {brand} с учетом кампании занял {rank_text(rank)}: "
            f"{fmt_int(mentions)} {_message_word(mentions)} и SOV {fmt_pct(sov)}."
        )

    leader_text = ""
    if leader and str(leader.get("brand", "")).strip() and str(leader.get("brand")).strip().lower() != brand.lower():
        gap_raw = as_float(leader.get("sov", 0)) - as_float(sov)
        gap_text = f" Разница с лидером — {fmt_pp(gap_raw * 100 if abs(gap_raw) <= 1 else gap_raw, signed=False)}" if gap_raw > 0 else ""
        leader_text = (
            f" Лидер — {leader.get('brand')} с {fmt_int(leader.get('mentions', 0))} "
            f"{_message_word(leader.get('mentions', 0))} и SOV {fmt_pct(leader.get('sov', 0))}.{gap_text}"
        )

    if campaign:
        campaign_text = (
            f" В кампании учтено {fmt_int(campaign)} {_publication_word(campaign)}. "
            f"Органическая часть брендового инфополя — {fmt_int(organic)} {_mention_word(organic)}."
        )
    else:
        campaign_text = " Кампания является частью отчета, но подтвержденный факт публикаций не найден в текущих вводных."

    position_hint = (
        " Это сильная стартовая база для закрепления доли в обсуждениях."
        if as_int(rank) and as_int(rank) <= 2
        else " Задача кампании — усиливать присутствие в контекстах, где бренд уже сравнивают с конкурентами."
    )
    split = data.get("campaign_split") or {}
    warning = campaign_required_warning(split)
    return " ".join(part for part in [first, leader_text, campaign_text, position_hint, warning] if part).strip()


def slide04_problem_insights(data: dict) -> list[str]:
    brand = _brand(data)
    total = as_int(data.get("total_topic_messages"))
    contexts = [
        str(item.get("rail_phrase") or item.get("name") or "").strip()
        for item in (data.get("chart_terms") or [])[:3]
    ]
    if not contexts:
        contexts = [str(item).strip() for item in (data.get("right_contexts") or [])[:3]]
    context_text = join_ru(contexts) or "ключевые пользовательские ситуации"
    signals = data.get("insight_signals") or {}
    brand_ctx = signals.get("brand_mentions") or {}
    brand_mentions = as_int(brand_ctx.get("brand_mentions"))
    coverage_text = ""
    if brand_mentions and total:
        coverage_text = (
            f" {brand} присутствует в {fmt_int(brand_mentions)} из них "
            f"({fmt_pct(brand_mentions / total)}), поэтому есть пространство для более заметной связи бренда с проблемой и выбором решения."
        )

    fact = (
        f"{period_start(data)} получено {fmt_int(total)} {_message_word(total)} по ключевым контекстам. "
        f"Основные поводы обсуждений — {context_text}.{coverage_text}"
    )
    action = (
        "Рекомендация — выводить материалы кампании в уже существующие обсуждения: "
        "показывать связь продукта с симптомом, моментом выбора, курсом применения и обращением за советом."
    )
    return [fact, action]


def slide05_sov_insight(data: dict) -> str:
    brand = _brand(data)

    # Support canonical model data merge
    cm = data.get("canonical_model") if isinstance(data, dict) else None
    if cm is not None:
        tb = cm.get("target_brand", {})
        sov_cm = cm.get("competitive_sov_scope", {})
        sov_target = sov_cm.get("target", {})
        orm_camp = cm.get("orm_campaign", {})
        rank = data.get("project_rank")
        sov = sov_target.get("sov_with_campaign")
        rows = sov_cm.get("competitors", [])
        competitor = rows[0] if rows else {}
        campaign = as_int(orm_camp.get("total"))
        organic_sov = sov_target.get("sov_without_campaign")
        organic_mentions = as_int(tb.get("organic"))
        lift = sov_target.get("sov_lift_pp")
    else:
        rows = data.get("all_brand_rows") or data.get("pie_rows") or []
        rank = data.get("project_rank")
        sov = data.get("project_sov")
        competitor = _top_competitor(rows, brand)
        split = data.get("campaign_split") or {}
        campaign = as_int(split.get("campaign_publications_count") or split.get("campaign_materials"))
        organic_sov = split.get("organic_sov")
        organic_mentions = as_int(split.get("organic_mentions"))
        lift = split.get("sov_lift_pp")

    if competitor and str(competitor.get("brand", "")).strip():
        if as_int(rank) == 1:
            position = (
                f"Доля голоса бренда {brand} в группе прямых конкурентов составила {fmt_pct(sov)}. "
                f"Ближайший конкурент — {competitor.get('brand')} с {fmt_pct(competitor.get('sov', 0))}."
            )
        else:
            position = (
                f"Доля голоса бренда {brand} в группе прямых конкурентов составила {fmt_pct(sov)}. "
                f"Лидер — {competitor.get('brand')} с {fmt_pct(competitor.get('sov', 0))}."
            )
    else:
        position = f"Доля голоса бренда {brand} в группе прямых конкурентов составила {fmt_pct(sov)}."

    if campaign and organic_sov is not None and lift is not None:
        organic_detail = ""
        if organic_mentions:
            organic_detail = f", {fmt_int(organic_mentions)} {_mention_word(organic_mentions)}"
        return (
            f"{position} При этом без учета репутационной кампании показатель был бы ниже "
            f"({fmt_pct(organic_sov)}{organic_detail}). "
            f"Разница между сценариями — {fmt_pp(lift, signed=False)} "
            f"В расчете учтено {fmt_int(campaign)} {_publication_word(campaign)}."
        )
    return (
        f"{position} Кампания является обязательной частью отчета, но подтвержденный факт публикаций не найден. "
        "Сценарное сравнение нужно пересчитать после проверки проектной ORM-таблицы."
    )


def slide06_tonality_insight(data: dict) -> list[str]:
    brand = _brand(data)

    # Support canonical model data merge
    cm = data.get("canonical_model") if isinstance(data, dict) else None
    if cm is not None:
        tonality_cm = cm.get("target_tonality", {})
        total = as_int(tonality_cm.get("positive", 0) + tonality_cm.get("neutral", 0) + tonality_cm.get("negative", 0))
        pos = as_int(tonality_cm.get("positive"))
        neu = as_int(tonality_cm.get("neutral"))
        neg = as_int(tonality_cm.get("negative"))
    else:
        row = data.get("project_row") or {}
        total = as_int(row.get("total"))
        pos = as_int(row.get("positive"))
        neu = as_int(row.get("neutral"))
        neg = as_int(row.get("negative"))

    pos_pct = pos / total if total else 0
    neu_pct = neu / total if total else 0
    neg_pct = neg / total if total else 0
    positive_line = (
        f"Доля позитивных упоминаний {period_in(data)} у бренда {brand} — {fmt_pct(pos_pct, 0)} "
        f"({fmt_int(pos)} {_message_word(pos)}). "
        "Позитивные сообщения отражают опыт применения, удобство продукта и общий положительный эффект."
    )
    if neg:
        negative_line = (
            f"Негативные упоминания бренда {brand} — {fmt_pct(neg_pct, 0)} ({fmt_int(neg)} {_message_word(neg)}). "
            "Эти сообщения требуют отдельного просмотра: они показывают зоны контроля в карточках, ответах и пользовательском опыте."
        )
    else:
        negative_line = f"Негативные упоминания бренда {brand} не зафиксированы. Важно сохранять мониторинг вопросов и нейтральных сомнений."
    neutral_line = (
        f"Нейтральные упоминания составили {fmt_pct(neu_pct, 0)} ({fmt_int(neu)} {_message_word(neu)}): "
        "это вопросы, уточнения и обсуждения выбора без выраженной оценки."
    )
    return [positive_line, negative_line, neutral_line]


def slide07_sources_insight(data: dict) -> str:
    # Support canonical model data merge
    cm = data.get("canonical_model") if isinstance(data, dict) else None
    if cm is not None:
        platform = cm.get("platform_scope", {})
        data["platform_type_rows"] = platform.get("platform_type_rows", data.get("platform_type_rows", []))
        data["source_rows"] = platform.get("source_rows", data.get("source_rows", []))
        data["total_messages"] = platform.get("total", data.get("total_messages", 0))

    total = as_int(data.get("total_messages"))
    rows = {row.get("type"): row for row in data.get("platform_type_rows") or []}
    social = as_int(rows.get("Соцсети", {}).get("messages"))
    messengers = as_int(rows.get("Мессенджеры", {}).get("messages"))
    reviews = as_int(rows.get("Отзывы", {}).get("messages"))
    core = social + messengers
    top_sources = [row for row in (data.get("source_rows") or []) if row.get("label") != "Другие"][:2]
    top_source_text = join_ru([
        f"{row.get('label')} ({fmt_int(row.get('messages'))} {_message_word(row.get('messages'))})"
        for row in top_sources
        if row.get("label")
    ]) or "ключевые источники"
    review_share = reviews / total if total else 0
    review_sources = data.get("review_sources") or []
    review_source_text = join_ru([
        f"{source} ({fmt_int(count)})"
        for source, count in review_sources[:3]
        if source
    ])
    review_detail = f" — прежде всего это отзывы на {review_source_text}" if review_source_text else ""
    return (
        f"{period_start(data)} более половины всех обсуждений ({fmt_pct(core / total if total else 0)}) составили обсуждения в соцсетях "
        f"и мессенджерах. По количеству упоминаний лидируют {top_source_text}. "
        "Бренды обсуждают в сообществах и чатах: "
        "пользователи задают вопросы о применении, сравнивают препараты и делятся опытом.\n"
        f"Отзывы составили {fmt_pct(review_share)} инфополя{review_detail}: суммарно {fmt_int(reviews)} "
        f"{_message_word(reviews)}."
    )


def slide08_seeding_summary(data: dict) -> str:
    # Support canonical model data merge
    cm = data.get("canonical_model") if isinstance(data, dict) else None
    if cm is not None:
        orm_camp = cm.get("orm_campaign", {})
        data["total_materials"] = orm_camp.get("total", data.get("total_materials", 0))
        data["views_metrics"] = {
            "received_views_total": orm_camp.get("views_total", 0),
            "has_views": orm_camp.get("views_total", 0) > 0,
            "metric_label": "Просмотры",
        }

    total = as_int(data.get("total_materials"))
    reviews = as_int(data.get("review_materials"))
    comments = as_int(data.get("comment_materials"))
    metrics = data.get("views_metrics") or {}
    views = as_int(metrics.get("selected_views_total") or metrics.get("received_views_total"))
    platform_rows = data.get("platform_rows") or []
    top_platforms = join_ru([str(row.get("platform") or row.get("label") or "").strip() for row in platform_rows[:3]])
    placement = (
        f" Основные точки размещения: {top_platforms}."
        if top_platforms
        else ""
    )
    views_text = (
        f" Материалы получили {fmt_int(views)} просмотров."
        if views
        else " Просмотры не выводятся как факт: в текущих вводных поле просмотров не заполнено."
    )
    return (
        f"{period_start(data)} в проектной таблице учтено {fmt_int(total)} {_publication_word(total)} кампании: "
        f"{fmt_int(comments)} {_message_word(comments)} в обсуждениях и {fmt_int(reviews)} "
        f"{plural_ru(reviews, 'отзыв', 'отзыва', 'отзывов')}.{placement}{views_text}"
    )


def slide11_ratings_insight(data: dict) -> str:
    brand = _brand(data)
    start = data.get("start_summary") or {}
    current = data.get("current_summary") or {}
    caveat = str(data.get("ratings_period_caveat") or "").strip()
    caveat_prefix = f"{caveat} " if caveat else ""
    if current.get("data_status") == "missing":
        return (
            f"{caveat_prefix}На старте по бренду {brand} зафиксировано {fmt_int(start.get('reviews_total'))} "
            "отзывов. В текущем периоде рейтинговые значения не заполнены, поэтому динамику карточек нужно подтвердить отдельно."
        )
    start_pos = as_int(start.get("positive_cards"))
    cur_pos = as_int(current.get("positive_cards"))
    start_neg = as_int(start.get("negative_cards"))
    cur_neg = as_int(current.get("negative_cards"))
    start_no = as_int(start.get("no_reviews"))
    cur_no = as_int(current.get("no_reviews"))
    start_reviews = as_int(start.get("reviews_total"))
    cur_reviews = as_int(current.get("reviews_total"))
    reviews_delta = cur_reviews - start_reviews
    placed = as_int((data.get("agency_review_placements") or {}).get("placed_reviews"))
    neg_text = (
        f"Негативных карточек стало {fmt_int(cur_neg)} против {fmt_int(start_neg)} на старте."
        if cur_neg
        else f"Негативных карточек в текущем периоде не зафиксировано. На старте было {fmt_int(start_neg)}."
    )
    return (
        f"{caveat_prefix}В текущем периоде у бренда {brand} {fmt_int(cur_pos)} позитивных карточек против {fmt_int(start_pos)} на старте. "
        f"{neg_text} Карточек без отзывов — {fmt_int(cur_no)} против {fmt_int(start_no)}. "
        f"Количество отзывов изменилось с {fmt_int(start_reviews)} до {fmt_int(cur_reviews)} "
        f"({reviews_delta:+,}).".replace(",", " ")
        + (f" В кампании учтено {fmt_int(placed)} свежих отзывов." if placed else "")
    )


def build_final_conclusions(project: dict, slide_data_by_number: dict[int, dict]) -> dict:
    slide04 = slide_data_by_number.get(4, {})
    slide05 = slide_data_by_number.get(5, {})
    slide06 = slide_data_by_number.get(6, {})
    slide07 = slide_data_by_number.get(7, {})
    slide08 = slide_data_by_number.get(8, {})
    slide11 = slide_data_by_number.get(11, {})

    # Merge canonical model data into each slide data dict
    cm = slide_data_by_number.get("canonical_model")
    if cm is not None:
        for d in [slide05, slide06, slide07, slide08]:
            d["canonical_model"] = cm

    return {
        "slide03_05": slide05_sov_insight(slide05),
        "slide04": " ".join(slide04_problem_insights(slide04)),
        "slide07": slide07_sources_insight(slide07),
        "slide08": slide08_seeding_summary(slide08),
        "slide11": slide11_ratings_insight(slide11),
        "tone": slide06_tonality_insight(slide06),
    }


def plain_segments(text: str, *, lead_bold: bool = True) -> list[dict[str, Any]]:
    clean = normalize_text(text)
    if not lead_bold:
        return [{"text": clean, "bold": False}]
    match = re.match(r"([^:]{3,70}:)(.*)", clean)
    if not match:
        return [{"text": clean, "bold": False}]
    return [
        {"text": match.group(1) + " ", "bold": True},
        {"text": match.group(2).strip(), "bold": False},
    ]


def normalize_text(text: str) -> str:
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    return text


def qa_text_contract(text: str, *, slide_number: int) -> dict[str, Any]:
    raw = str(text or "")
    low = raw.lower()
    blockers = []
    forbidden = [phrase for phrase in FORBIDDEN_MAIN_TEXT_PHRASES if phrase.lower() in low]
    if forbidden:
        blockers.append(
            f"Slide {slide_number:02d} main text contains forbidden source/method phrase: "
            + ", ".join(forbidden)
        )
    return {
        "blockers": blockers,
        "warnings": [],
        "checks": {
            "no_source_or_methodology_in_main_text": not forbidden,
        },
    }
