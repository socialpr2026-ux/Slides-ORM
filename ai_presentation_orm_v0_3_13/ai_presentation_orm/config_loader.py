from pathlib import Path

try:
    import yaml
except ModuleNotFoundError:  # pragma: no cover - runtime fallback for bundled packages
    yaml = None


def _scalar(value: str):
    text = value.strip()
    if not text:
        return ""
    if (text.startswith('"') and text.endswith('"')) or (text.startswith("'") and text.endswith("'")):
        return text[1:-1]
    low = text.lower()
    if low == "true":
        return True
    if low == "false":
        return False
    if low in {"null", "none"}:
        return None
    try:
        return int(text)
    except ValueError:
        return text


def _line_indent(line: str) -> int:
    return len(line) - len(line.lstrip(" "))


def _next_content(lines: list[str], start: int):
    for idx in range(start, len(lines)):
        stripped = lines[idx].strip()
        if stripped and not stripped.startswith("#"):
            return idx, stripped
    return None, ""


def _parse_list(lines: list[str], index: int, indent: int):
    out = []
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            index += 1
            continue
        current_indent = _line_indent(line)
        if current_indent < indent:
            break
        if current_indent != indent or not stripped.startswith("- "):
            break
        out.append(_scalar(stripped[2:].strip()))
        index += 1
    return out, index


def _parse_dict(lines: list[str], index: int = 0, indent: int = 0):
    out = {}
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            index += 1
            continue
        current_indent = _line_indent(line)
        if current_indent < indent:
            break
        if current_indent > indent:
            index += 1
            continue
        if ":" not in stripped:
            index += 1
            continue
        key, value = stripped.split(":", 1)
        key = key.strip()
        value = value.strip()
        if value:
            out[key] = _scalar(value)
            index += 1
            continue
        next_idx, next_stripped = _next_content(lines, index + 1)
        if next_idx is None:
            out[key] = {}
            index += 1
            continue
        next_indent = _line_indent(lines[next_idx])
        if next_indent <= current_indent:
            out[key] = {}
            index += 1
        elif next_stripped.startswith("- "):
            out[key], index = _parse_list(lines, next_idx, next_indent)
        else:
            out[key], index = _parse_dict(lines, next_idx, next_indent)
    return out, index


def _fallback_safe_load(text: str) -> dict:
    parsed, _ = _parse_dict(text.splitlines())
    return parsed


def load_config(path: str) -> dict:
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"Config not found: {config_path}")
    with config_path.open("r", encoding="utf-8") as f:
        raw_text = f.read()
        config = yaml.safe_load(raw_text) if yaml else _fallback_safe_load(raw_text)
        config = config or {}

    required = ["project", "paths", "processing", "rules"]
    for key in required:
        if key not in config:
            raise ValueError(f"Missing config section: {key}")

    if "output_dir" not in config["paths"]:
        raise ValueError("paths.output_dir is required")
    if "template_pptx" not in config["paths"]:
        raise ValueError("paths.template_pptx is required")
    if "input_dir" not in config["paths"]:
        raise ValueError("paths.input_dir is required")

    return config
