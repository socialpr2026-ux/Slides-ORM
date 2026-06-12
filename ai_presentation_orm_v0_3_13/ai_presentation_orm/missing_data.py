MISSING_DATA_QUESTIONS = {
    "project": ("проектные данные", "Нужны бренд, клиент, период и категория.", "бриф / PM-комментарий / отчетный Excel"),
    "plan_fact": ("план-факт", "Нужны план, факт, выполнение и причины расхождений.", "медиаплан / рабочий ORM-отчет / Placement"),
    "brand_mentions": ("упоминания брендов", "Нужны бренды и количество сообщений.", "Insight / Медиалогия / Brand Analytics"),
    "sov": ("SOV", "Нужен SOV, тип расчета, период и методология.", "Insight / аналитический отчет"),
    "themes": ("тематики", "Нужны темы, объемы и источник данных.", "Insight / BA / Медиалогия"),
    "sentiment": ("тональность", "Нужны позитив, нейтрал, негатив и источник разметки.", "Insight / аналитический отчет"),
    "sources": ("источники", "Нужны площадки, сообщения, аудитория / просмотры.", "Insight / BA / Медиалогия"),
    "seeding": ("посев / размещения", "Нужны публикации, площадки, просмотры, ссылки.", "рабочий ORM-отчет / Placement"),
    "examples": ("примеры", "Нужны реальные тексты, ссылки или скриншоты.", "рабочая таблица размещений"),
    "ratings": ("рейтинги / карточки", "Нужны старт и финал рейтингов / отзывов.", "Review / карточки / скриншоты"),
    "conclusions": ("выводы", "Нужны факт, интерпретация, действие, оговорка.", "Insight layer после загрузки данных"),
}


def build_missing_data_report(slide_readiness: dict) -> dict:
    items = []
    for slide in slide_readiness.get("slides", []):
        for entity in slide.get("missing_entities", []):
            block, need, source = MISSING_DATA_QUESTIONS.get(
                entity,
                (entity, f"Нужны данные для {entity}.", "PM / источник данных")
            )
            items.append({
                "slide_number": slide.get("slide_number"),
                "role": slide.get("role"),
                "block": block,
                "missing": need,
                "why_needed": f"Без этого слайд {slide.get('role')} нельзя собрать без риска.",
                "where_to_get": source,
                "impact_if_missing": "Слайд будет partial / missing_required или потребует удаления / оговорки.",
            })
    return {"items": items}
