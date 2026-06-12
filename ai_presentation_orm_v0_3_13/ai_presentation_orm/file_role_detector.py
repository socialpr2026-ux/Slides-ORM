from __future__ import annotations

from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

from .utils import safe_text, path_matches_any

NS = {
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


ROLE_KEYWORDS = {
    "media_plan": [
        "медиаплан", "наименование работ", "репутационная поддержка бренда",
        "всего флайт", "месяц (план)", "месяц (факт)", "единиц размещения"
    ],
    "analytical_report": [
        "sov", "тональность", "упоминания", "инфополе", "источники", "аудитория", "сюжеты",
        "медиалогия", "brand analytics", "аналитический отчет", "аналитический отчёт"
    ],
    "placement_table": [
        "размещено", "ссылка", "текст сообщения", "просмотров получено", "вовлечение",
        "комментарии", "отзывы", "площадка"
    ],
    "report_excel": [
        "текущий отчет", "план", "факт", "публикации", "отчет", "отчёт", "итого", "просмотры"
    ],
    "ratings_table": [
        "рейтинг", "карточ", "sku", "оценка", "звезды", "отзывы с покупкой"
    ],
    "brief_or_notes": [
        "бриф", "claims", "ограничения", "ключевые сообщения", "задача", "комментарий pm"
    ],
    "raw_export": [
        "сообщения", "авторы", "сообщества", "география", "теги", "популярные слова",
        "персоны", "юрлица", "продукты"
    ],
}


def collect_input_files(config: dict) -> list[Path]:
    """Collect input files.

    Prefer explicit inputs.files for reproducible project runs.
    Fall back to glob patterns only when explicit files are not provided.
    """
    inputs_cfg = config.get("inputs", {})
    explicit_files = inputs_cfg.get("files") or []
    exclude_patterns = inputs_cfg.get("exclude_patterns", [])

    if explicit_files:
        files = [Path(p) for p in explicit_files]
    else:
        input_dir = Path(config["paths"]["input_dir"])
        include_patterns = inputs_cfg.get("include_patterns", ["*.xlsx", "*.csv", "*.txt", "*.pptx"])
        files = []
        for pattern in include_patterns:
            files.extend(input_dir.glob(pattern))

    unique = []
    seen = set()
    for path in files:
        if not path.is_file():
            continue
        if path_matches_any(path, exclude_patterns):
            continue
        resolved = path.resolve()
        if resolved not in seen:
            unique.append(path)
            seen.add(resolved)
    return sorted(unique, key=lambda p: p.name.lower())


def _load_shared_strings(zip_file: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall("x:si", NS):
        parts = [t.text or "" for t in si.findall(".//x:t", NS)]
        out.append("".join(parts))
    return out


def _sample_xlsx_text(path: Path, max_chars: int = 6000) -> str:
    texts = []
    try:
        with zipfile.ZipFile(path, "r") as z:
            shared = _load_shared_strings(z)
            texts.extend(shared[:400])
            # Also include workbook sheet names
            if "xl/workbook.xml" in z.namelist():
                root = ET.fromstring(z.read("xl/workbook.xml"))
                for sh in root.findall("x:sheets/x:sheet", NS):
                    texts.append(sh.attrib.get("name", ""))
    except Exception:
        return ""
    return safe_text(" ".join(texts), max_chars).lower()


def _sample_text_file(path: Path, max_chars: int = 6000) -> str:
    try:
        return safe_text(path.read_text(encoding="utf-8", errors="ignore"), max_chars).lower()
    except Exception:
        return ""


def _score_roles(text: str) -> dict[str, int]:
    scores = {role: 0 for role in ROLE_KEYWORDS}
    low = text.lower()
    for role, keywords in ROLE_KEYWORDS.items():
        for keyword in keywords:
            if keyword.lower() in low:
                scores[role] += 1
    return scores


def detect_file_roles(input_files: list[Path], template_path: Path) -> dict:
    result = {
        "template_pptx": str(template_path),
        "files": []
    }
    for path in input_files:
        suffix = path.suffix.lower()
        role = "unknown"
        text_sample = ""
        scores = {}
        if path.resolve() == template_path.resolve():
            role = "template_pptx"
        elif suffix in [".xlsx", ".xlsm"]:
            text_sample = _sample_xlsx_text(path)
            scores = _score_roles(text_sample + " " + path.name.lower())
            role = max(scores, key=scores.get) if scores and max(scores.values()) > 0 else "unknown_excel"
        elif suffix == ".csv":
            text_sample = _sample_text_file(path)
            scores = _score_roles(text_sample + " " + path.name.lower())
            role = max(scores, key=scores.get) if scores and max(scores.values()) > 0 else "unknown_csv"
        elif suffix == ".txt":
            text_sample = _sample_text_file(path)
            scores = _score_roles(text_sample + " " + path.name.lower())
            role = max(scores, key=scores.get) if scores and max(scores.values()) > 0 else "brief_or_notes"
        elif suffix == ".pptx":
            role = "reference_or_old_presentation"

        result["files"].append({
            "file": str(path),
            "name": path.name,
            "suffix": suffix,
            "role_guess": role,
            "role_scores": scores,
            "sample_preview": safe_text(text_sample, 400),
        })
    return result
