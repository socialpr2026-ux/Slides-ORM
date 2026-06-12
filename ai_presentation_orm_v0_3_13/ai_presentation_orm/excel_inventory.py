from __future__ import annotations

from pathlib import Path
import csv
import re
import zipfile
import xml.etree.ElementTree as ET

from .utils import safe_text

NS = {
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


ENTITY_KEYWORDS = {
    "media_plan": ["медиаплан", "наименование работ", "репутационная поддержка бренда", "всего флайт", "единиц размещения"],
    "project": ["проект", "бренд", "период", "клиент", "кампания"],
    "plan_fact": ["план", "факт", "выполнение", "расхождение", "публикации"],
    "brand_mentions": ["упоминания", "сообщения", "бренды", "number of mentions"],
    "sov": ["sov", "доля голоса", "share of voice", "sov%"],
    "themes": ["темы", "тематика", "инфополе", "контексты", "диагнозы", "проблема"],
    "sentiment": ["тональность", "позитив", "негатив", "нейтрал"],
    "sources": ["источники", "площадки", "тип источника", "сообщества", "аудитория"],
    "seeding": ["посев", "размещено", "ссылка", "просмотры", "вовлечение"],
    "examples": ["примеры", "текст сообщения", "отзывы", "комментарии"],
    "ratings": ["рейтинг", "карточ", "оценка", "sku", "отзывы с покупкой"],
    "methodology": ["источник данных", "методология", "исключен спам", "период"],
}


def _col_to_num(col: str) -> int:
    n = 0
    for c in col:
        n = n * 26 + ord(c.upper()) - 64
    return n


def _cell_ref_to_rc(ref: str) -> tuple[int, int]:
    match = re.match(r"([A-Z]+)(\d+)", ref or "")
    return (int(match.group(2)), _col_to_num(match.group(1))) if match else (0, 0)


def _load_shared_strings(z: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall("x:si", NS):
        out.append("".join([t.text or "" for t in si.findall(".//x:t", NS)]))
    return out


def _get_workbook_sheets(z: zipfile.ZipFile) -> list[dict]:
    wb_root = ET.fromstring(z.read("xl/workbook.xml"))
    rel_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rels = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rel_root.findall("rel:Relationship", NS)}
    sheets = []
    for sh in wb_root.findall("x:sheets/x:sheet", NS):
        name = sh.attrib.get("name", "")
        rid = sh.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rels.get(rid, "")
        path = "xl/" + target.lstrip("/")
        path = path.replace("xl//", "xl/")
        sheets.append({"name": name, "path": path})
    return sheets


def _read_cell_value(c, shared_strings: list[str]) -> str:
    t = c.attrib.get("t")
    v = c.find("x:v", NS)
    if t == "inlineStr":
        return "".join([x.text or "" for x in c.findall(".//x:t", NS)])
    if v is None:
        return ""
    raw = v.text or ""
    if t == "s":
        try:
            return shared_strings[int(raw)]
        except Exception:
            return raw
    if t == "b":
        return "TRUE" if raw == "1" else "FALSE"
    return raw


def _sample_sheet(z: zipfile.ZipFile, sheet_path: str, shared_strings: list[str], max_rows: int, max_cols: int) -> dict:
    if sheet_path not in z.namelist():
        return {"dimension": "", "sample_rows": [], "detected_entities": [], "error": f"{sheet_path} not found"}
    root = ET.fromstring(z.read(sheet_path))
    dim_el = root.find("x:dimension", NS)
    dimension = dim_el.attrib.get("ref", "") if dim_el is not None else ""

    sample_rows = []
    text_blob = []
    for row in root.findall("x:sheetData/x:row", NS):
        rnum = int(row.attrib.get("r", "0"))
        if rnum > max_rows:
            break
        vals = [""] * max_cols
        for c in row.findall("x:c", NS):
            _, col = _cell_ref_to_rc(c.attrib.get("r", ""))
            if 1 <= col <= max_cols:
                val = safe_text(_read_cell_value(c, shared_strings), 160)
                vals[col - 1] = val
                if val:
                    text_blob.append(val)
        if any(vals):
            sample_rows.append(vals)

    detected = _detect_entities(" ".join(text_blob + [dimension]))
    return {"dimension": dimension, "sample_rows": sample_rows, "detected_entities": detected}


def _detect_entities(text: str) -> list[str]:
    low = text.lower()
    found = []
    for entity, keywords in ENTITY_KEYWORDS.items():
        if any(keyword.lower() in low for keyword in keywords):
            found.append(entity)
    return found


def _inventory_xlsx(path: Path, config: dict) -> dict:
    max_rows = int(config["processing"].get("max_sample_rows_per_sheet", 15))
    max_cols = int(config["processing"].get("max_sample_cols_per_sheet", 18))
    entry = {"file": str(path), "name": path.name, "type": "xlsx", "sheets": [], "detected_entities": []}
    try:
        with zipfile.ZipFile(path, "r") as z:
            shared = _load_shared_strings(z)
            for sheet in _get_workbook_sheets(z):
                sample = _sample_sheet(z, sheet["path"], shared, max_rows, max_cols)
                sheet_entry = {
                    "name": sheet["name"],
                    "path": sheet["path"],
                    "range": sample.get("dimension", ""),
                    "sample_rows": sample.get("sample_rows", []),
                    "detected_entities": sample.get("detected_entities", []),
                }
                entry["sheets"].append(sheet_entry)
        all_entities = []
        for sheet in entry["sheets"]:
            all_entities.extend(sheet["detected_entities"])
        entry["detected_entities"] = sorted(set(all_entities))
    except Exception as exc:
        entry["error"] = str(exc)
    return entry


def _inventory_csv(path: Path, config: dict) -> dict:
    max_rows = int(config["processing"].get("max_sample_rows_per_sheet", 15))
    entry = {"file": str(path), "name": path.name, "type": "csv", "sample_rows": [], "detected_entities": []}
    try:
        with path.open("r", encoding="utf-8", errors="ignore", newline="") as f:
            reader = csv.reader(f)
            for i, row in enumerate(reader):
                if i >= max_rows:
                    break
                entry["sample_rows"].append([safe_text(c, 160) for c in row[:18]])
        text = " ".join([" ".join(row) for row in entry["sample_rows"]])
        entry["detected_entities"] = _detect_entities(text)
    except Exception as exc:
        entry["error"] = str(exc)
    return entry


def _inventory_txt(path: Path) -> dict:
    entry = {"file": str(path), "name": path.name, "type": "txt", "sample_preview": "", "detected_entities": []}
    try:
        text = path.read_text(encoding="utf-8", errors="ignore")
        entry["sample_preview"] = safe_text(text, 1500)
        entry["detected_entities"] = _detect_entities(text)
    except Exception as exc:
        entry["error"] = str(exc)
    return entry


def build_data_inventory(input_files: list[Path], config: dict) -> dict:
    inventory = {"files": []}
    for path in input_files:
        suffix = path.suffix.lower()
        if suffix in [".xlsx", ".xlsm"]:
            inventory["files"].append(_inventory_xlsx(path, config))
        elif suffix == ".csv":
            inventory["files"].append(_inventory_csv(path, config))
        elif suffix == ".txt":
            inventory["files"].append(_inventory_txt(path))
        elif suffix == ".pptx":
            inventory["files"].append({"file": str(path), "name": path.name, "type": "pptx", "detected_entities": []})
        else:
            inventory["files"].append({"file": str(path), "name": path.name, "type": suffix.strip("."), "detected_entities": []})
    return inventory
