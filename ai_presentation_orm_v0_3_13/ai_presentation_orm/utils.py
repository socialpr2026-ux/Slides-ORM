import re
from pathlib import Path


def safe_text(value, limit: int = 500) -> str:
    if value is None:
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text[:limit]


def emu_to_cm(value) -> float:
    return round(value / 360000, 2)


def path_matches_any(path: Path, patterns: list[str]) -> bool:
    return any(path.match(pattern) for pattern in patterns or [])
