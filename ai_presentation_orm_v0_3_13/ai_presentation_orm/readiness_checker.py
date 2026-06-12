from __future__ import annotations

from collections import Counter

from .canonical_model import SLIDE_REQUIRED_ENTITIES


def _available_entities(data_inventory: dict) -> set[str]:
    available = set()
    for file_entry in data_inventory.get("files", []):
        for entity in file_entry.get("detected_entities", []):
            available.add(entity)
        for sheet in file_entry.get("sheets", []):
            for entity in sheet.get("detected_entities", []):
                available.add(entity)
    return available


def build_slide_readiness(template_map: dict, data_inventory: dict, file_roles: dict) -> dict:
    available = _available_entities(data_inventory)
    slides = []

    for slide in template_map.get("slides", []):
        role = slide.get("role_guess", "unknown")
        required = SLIDE_REQUIRED_ENTITIES.get(role, [])
        missing = [entity for entity in required if entity not in available]

        if not required:
            status = "partial"
        elif not missing:
            status = "ready_with_caveat"
        elif len(missing) < len(required):
            status = "partial"
        else:
            status = "missing_required"

        slides.append({
            "slide_number": slide.get("slide_number"),
            "role": role,
            "required_entities": required,
            "available_entities": sorted(list(available.intersection(required))),
            "missing_entities": missing,
            "status": status,
            "note": _status_note(role, status, missing),
        })

    return {
        "available_entities": sorted(list(available)),
        "status_counts": dict(Counter(slide["status"] for slide in slides)),
        "slides": slides,
    }


def _status_note(role: str, status: str, missing: list[str]) -> str:
    if status == "ready_with_caveat":
        return "Данные найдены, перед генерацией нужен QA и методологические оговорки."
    if status == "partial":
        return "Данных хватает частично, потребуется ручной mapping или уточнение."
    if status == "missing_required":
        return "Не хватает обязательных данных: " + ", ".join(missing)
    return "Нужна дополнительная проверка."
