from __future__ import annotations

import re
from typing import Any, Iterable


FONT_FAMILY = "Arial"
TITLE_FONT_SIZE = 28.0
BODY_FONT_SIZE = 11.0
BODY_LINE_SPACING = 1.0
CHART_FONT_SIZE = 10.0
TABLE_FONT_SIZE = 9.0
SOURCE_FONT_SIZE = 9.0

TEXT_SPACE_AFTER_PT = 0.0
TABLE_CELL_VERTICAL_ANCHOR = "middle"

SLIDE_CONTENT_LEFT_EMU = 424661

OTHER_BUCKET_LABELS = {
    "другие",
    "другое",
    "остальные",
    "остальное",
    "прочие",
    "прочее",
    "other",
    "others",
}


def normalize_label(value: Any) -> str:
    text = str(value or "").strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", text)


def is_other_bucket(value: Any) -> bool:
    text = normalize_label(value)
    if text in OTHER_BUCKET_LABELS:
        return True
    return text.startswith("друг") or text.startswith("осталь") or text.startswith("проч")


def format_int_spaces(value: int | float | None) -> str:
    if value is None:
        return "0"
    return f"{int(round(float(value))):,}".replace(",", " ")


def sort_other_last(
    rows: Iterable[dict[str, Any]] | None,
    *,
    label_key: str,
    value_key: str | None = None,
    descending: bool = True,
) -> list[dict[str, Any]]:
    def numeric(row: dict[str, Any]) -> float:
        if not value_key:
            return 0.0
        try:
            return float(row.get(value_key, 0) or 0)
        except Exception:
            return 0.0

    def key(row: dict[str, Any]) -> tuple[int, float, str]:
        value = numeric(row)
        score = -value if descending else value
        return (1 if is_other_bucket(row.get(label_key)) else 0, score, normalize_label(row.get(label_key)))

    return sorted(list(rows or []), key=key)


STYLE_FILLER_PHRASES = [
    "в рамках",
    "осуществляется",
    "имеет место",
    "на текущий момент",
    "данный",
    "является важным",
    "брендовое окружение концентрируется",
    "сценарий применения",
]

CAMPAIGN_CAUSAL_PHRASES = [
    "благодаря кампании",
    "в результате кампании",
    "кампания повлияла",
    "кампания влияет",
    "кампания привела",
    "кампания обеспечила",
    "эффект кампании",
    "прирост от кампании",
    "кампания добавила",
]


def _sentence_words(sentence: str) -> int:
    return len(re.findall(r"[A-Za-zА-Яа-яЁё0-9%+,.]+", sentence))


def qa_slide_copy_principles(
    text: str,
    *,
    slide_number: int,
    allow_campaign_causality: bool = False,
    max_sentence_words: int = 34,
) -> dict[str, Any]:
    """Apply practical slide-copy checks inspired by "Пиши, сокращай".

    The function blocks only high-risk issues, such as unsupported campaign
    causality. Style imperfections are warnings because generated slide copy can
    still be valid when it contains necessary technical terms.
    """
    raw = str(text or "")
    low = normalize_label(raw)
    blockers: list[str] = []
    warnings: list[str] = []

    filler_left = [phrase for phrase in STYLE_FILLER_PHRASES if phrase in low]
    if filler_left:
        warnings.append(
            f"Slide {slide_number:02d} copy has bureaucratic or heavy wording: "
            + ", ".join(filler_left[:6])
        )

    campaign_causal_left = [phrase for phrase in CAMPAIGN_CAUSAL_PHRASES if phrase in low]
    if campaign_causal_left and not allow_campaign_causality:
        blockers.append(
            f"Slide {slide_number:02d} must not present campaign contribution as proven causality: "
            + ", ".join(campaign_causal_left[:6])
        )

    sentences = [item.strip() for item in re.split(r"(?<=[.!?])\s+|\n+", raw) if item.strip()]
    long_sentences = [sentence for sentence in sentences if _sentence_words(sentence) > max_sentence_words]
    if long_sentences:
        warnings.append(
            f"Slide {slide_number:02d} has long sentences; simplify before client delivery."
        )

    leader_hits = len(re.findall(r"\bлидер\w*|\bлидир\w*", low, flags=re.IGNORECASE))
    if leader_hits > 2:
        warnings.append(f"Slide {slide_number:02d} repeats leader wording too often.")

    repeated_but = len(re.findall(r"\bно\b", low))
    if repeated_but:
        warnings.append(f"Slide {slide_number:02d} uses 'но'; check that it does not weaken a positive conclusion.")

    return {
        "blockers": blockers,
        "warnings": warnings,
        "checks": {
            "no_bureaucracy_or_cliche": not filler_left,
            "no_unsupported_campaign_causality": not campaign_causal_left or allow_campaign_causality,
            "short_direct_sentences": not long_sentences,
            "no_duplicated_leader_wording": leader_hits <= 2,
            "no_unneeded_but": repeated_but == 0,
        },
    }


def merge_copy_qa(target: dict[str, Any], copy_qa: dict[str, Any], *, checks_key: str = "copy_principles") -> None:
    target.setdefault("blockers", []).extend(copy_qa.get("blockers") or [])
    target.setdefault("warnings", []).extend(copy_qa.get("warnings") or [])
    target.setdefault("checks", {})[checks_key] = copy_qa.get("checks") or {}
