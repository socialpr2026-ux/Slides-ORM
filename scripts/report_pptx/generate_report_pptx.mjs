#!/usr/bin/env node
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

function parseArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) continue;
    const key = token.slice(2);
    const next = argv[index + 1];
    if (!next || next.startsWith("--")) {
      args[key] = true;
    } else {
      args[key] = next;
      index += 1;
    }
  }
  return args;
}

function usage() {
  return `Usage:
node scripts/report_pptx/generate_report_pptx.mjs \\
  --template "reference.pptx" \\
  --raw "raw_export.xlsx" \\
  --orm "orm_report.xlsx" \\
  --analytics "analytics_report.xlsx" \\
  --summary "summary.docx" \\
  --brand "Энтеролактис" \\
  --period "апрель 2026" \\
  --output "outputs/report.pptx"

Optional:
  --engine python-v0.3.13                   Use current Python pipeline for slides 1-12.
  --pipeline-root "path/to/package"          Override ai_presentation_orm_v0_3_13 path.
  --campaign-materials-override 35           Optional slide 05 campaign materials override.
  --campaign-month-sheet "Май 2026"          Optional ORM month sheet for slides 02, 05, 08-10.
  --ratings-sheet "Рейтинги"                 Optional ORM ratings sheet for slide 11.
  --ratings-current-period "Апрель 2026"     Optional ratings period label for slide 11.
   --screenshot-backend chatgpt              For Custom GPT: emit/consume ChatGPT screenshot handoff files.
   --screenshots-dir "path/to/screenshots"   Legacy mode only: replace image frames.
   --strict false                            Keep a draft even with QA blockers.
   --source-exclusions "Facebook,Twitter"    Comma-separated sources to exclude from all slides.
   --campaign-matched-analytics 43           Override: how many ORM placements matched in analytics (for organic calculation).`;
}

function requirePath(args, key) {
  const value = args[key];
  if (!value) throw new Error(`Не указан обязательный аргумент --${key}.\n${usage()}`);
  return path.resolve(String(value));
}

function optionalPath(args, key) {
  return args[key] ? path.resolve(String(args[key])) : null;
}

function ensureExists(filePath, label, required = true) {
  if (!filePath) {
    if (required) throw new Error(`Не указан файл ${label}`);
    return;
  }
  if (!fs.existsSync(filePath)) throw new Error(`Файл ${label} не найден: ${filePath}`);
}

function findPython() {
  const candidates = [
    process.env.PYTHON,
    path.join(os.homedir(), ".cache", "codex-runtimes", "codex-primary-runtime", "dependencies", "python", "python.exe"),
    path.join(os.homedir(), ".cache", "codex-runtimes", "codex-primary-runtime", "dependencies", "python", "bin", "python"),
    "python",
    "python3",
  ].filter(Boolean);
  for (const candidate of candidates) {
    const probe = spawnSync(candidate, ["--version"], { encoding: "utf8" });
    if (probe.status === 0) return candidate;
  }
  throw new Error("Python не найден. Нужен Python для устойчивого чтения OOXML XLSX/DOCX/PPTX.");
}

const PY_WORKER = String.raw`
import copy
import io
import json
import math
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
  "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
  "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
  "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
  "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
  "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
}

for prefix, uri in NS.items():
  ET.register_namespace(prefix, uri)

def norm_text(value):
  return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()

def norm_key(value):
  return norm_text(value).replace("\u0451", "\u0435").lower()

def non_empty_count(row):
  return len([cell for cell in row if norm_text(cell)])

def first_non_empty(row):
  for cell in row:
    value = norm_text(cell)
    if value:
      return value
  return ""

def is_total_label(value):
  return norm_key(value) in ("итого", "всего", "total")

def is_context_metric_label(value):
  low = norm_key(value)
  return (
    not low
    or "обсуждение" in low
    or "тематика" in low
    or "последствия" in low
    or "расстройств" in low
    or "пост-ковид" in low
  )

def parse_number(value):
  if value is None:
    return None
  if isinstance(value, (int, float)):
    return float(value)
  raw = str(value).replace("\xa0", " ").replace("\n", " ").strip()
  if not raw:
    return None
  raw = raw.replace(" ", "").replace("%", "").replace(",", ".")
  try:
    return float(raw)
  except Exception:
    return None

def fmt_int(value):
  if value is None:
    return "нет данных"
  try:
    return f"{int(round(float(value))):,}".replace(",", " ")
  except Exception:
    return str(value)

def fmt_pct(value, digits=1):
  if value is None:
    return "нет данных"
  number = float(value)
  if abs(number) <= 1:
    number *= 100
  text = f"{number:.{digits}f}".replace(".", ",")
  if text.endswith(",0"):
    text = text[:-2]
  return f"{text}%"

def period_month(period, capitalized=False):
  month = norm_text(period).split(" ")[0] if norm_text(period) else period
  return month[:1].upper() + month[1:] if capitalized and month else month

def period_year(period):
  match = re.search(r"(20\d{2})", norm_text(period))
  return match.group(1) if match else ""

def previous_period_month(period, capitalized=False):
  months = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
  month = period_month(period).lower()
  if month not in months:
    return period_month(period, capitalized)
  prev = months[(months.index(month) - 1) % len(months)]
  return prev[:1].upper() + prev[1:] if capitalized else prev

def col_to_idx(ref):
  match = re.match(r"([A-Z]+)", ref or "")
  if not match:
    return 0
  out = 0
  for char in match.group(1):
    out = out * 26 + ord(char) - 64
  return out - 1

def idx_to_col(index):
  index += 1
  letters = ""
  while index:
    index, remainder = divmod(index - 1, 26)
    letters = chr(65 + remainder) + letters
  return letters

def read_xlsx(path):
  sheets = {}
  with zipfile.ZipFile(path) as zf:
    names = set(zf.namelist())
    shared = []
    if "xl/sharedStrings.xml" in names:
      root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
      for item in root.findall("m:si", NS):
        shared.append("".join(t.text or "" for t in item.findall(".//m:t", NS)))

    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.get("Id"): rel.get("Target") for rel in rels.findall("pr:Relationship", NS)}
    for sheet in workbook.findall(".//m:sheet", NS):
      sheet_name = sheet.get("name") or "Sheet"
      rid = sheet.get("{%s}id" % NS["r"])
      target = (rel_map.get(rid) or "").lstrip("/")
      sheet_path = "xl/" + target
      if sheet_path not in names:
        sheet_path = "xl/worksheets/" + Path(target).name
      if sheet_path not in names:
        continue

      root = ET.fromstring(zf.read(sheet_path))
      rows = []
      for row in root.findall(".//m:sheetData/m:row", NS):
        values = []
        for cell in row.findall("m:c", NS):
          idx = col_to_idx(cell.get("r"))
          while len(values) <= idx:
            values.append("")
          cell_type = cell.get("t")
          v = cell.find("m:v", NS)
          inline = cell.find("m:is", NS)
          value = ""
          if cell_type == "s" and v is not None and v.text not in (None, ""):
            si = int(float(v.text))
            value = shared[si] if 0 <= si < len(shared) else ""
          elif cell_type == "inlineStr" and inline is not None:
            value = "".join(t.text or "" for t in inline.findall(".//m:t", NS))
          elif v is not None and v.text is not None:
            value = v.text
          values[idx] = value
        while values and values[-1] == "":
          values.pop()
        if any(norm_text(v) for v in values):
          rows.append(values)
      sheets[sheet_name] = rows
  return sheets

def read_docx(path):
  if not path:
    return []
  with zipfile.ZipFile(path) as zf:
    if "word/document.xml" not in zf.namelist():
      return []
    root = ET.fromstring(zf.read("word/document.xml"))
    paragraphs = []
    for para in root.findall(".//w:p", NS):
      text = "".join(t.text or "" for t in para.findall(".//w:t", NS))
      text = norm_text(text)
      if text:
        paragraphs.append(text)
    return paragraphs

def find_sheet(workbook, *needles):
  lower = [(name, name.lower()) for name in workbook.keys()]
  for needle in needles:
    n = needle.lower()
    for original, low in lower:
      if n == low:
        return original, workbook[original]
    for original, low in lower:
      if n in low:
        return original, workbook[original]
  return None, []

def table_header(rows, required, min_non_empty=None):
  required_low = [norm_key(item) for item in required]
  if min_non_empty is None:
    min_non_empty = max(2, len(required))
  for index, row in enumerate(rows):
    if non_empty_count(row) < min_non_empty:
      continue
    low = [norm_key(cell) for cell in row]
    if all(any(req in cell for cell in low) for req in required_low):
      return index, [norm_text(cell) for cell in row]
  return None, []

def value_by_header(row, header, *names):
  low = [norm_key(h) for h in header]
  for name in names:
    target = norm_key(name)
    for index, h in enumerate(low):
      if target == h or target in h:
        return row[index] if index < len(row) else ""
  return ""

def column_index(header, *names):
  low = [norm_key(h) for h in header]
  for name in names:
    target = norm_key(name)
    for index, h in enumerate(low):
      if target == h or target in h:
        return index
  return -1

def pairs_from_cross_table(rows, brand, label_header):
  header_index, header = table_header(rows, [label_header, brand])
  if header_index is None:
    return []
  brand_col = column_index(header, brand)
  if brand_col < 0:
    return []
  out = []
  for row in rows[header_index + 1:]:
    label = norm_text(row[1] if len(row) > 1 and not norm_text(row[0]) else row[0] if row else "")
    if not label:
      continue
    value = parse_number(row[brand_col] if brand_col < len(row) else "")
    if value is not None:
      out.append({"label": label, "value": value})
  return out

def brand_metrics(analytics, brand):
  sheet_name, rows = find_sheet(analytics, "Показатели брендов")
  header_index, header = table_header(rows, ["Объект", "Сообщения", "Аудитория"])
  brands = []
  target = {}
  total_row = {}
  if header_index is not None:
    for row in rows[header_index + 1:]:
      obj = norm_text(value_by_header(row, header, "Объект"))
      messages = parse_number(value_by_header(row, header, "Сообщения"))
      if not obj:
        continue
      if messages is None:
        if brands or total_row:
          break
        continue
      item = {
        "brand": obj,
        "messages": messages,
        "audience": parse_number(value_by_header(row, header, "Аудитория")),
        "views": parse_number(value_by_header(row, header, "Просмотры")),
        "engagement": parse_number(value_by_header(row, header, "Вовлеч")),
        "sovRaw": parse_number(value_by_header(row, header, "Share of voice", "SOV")),
        "source": f"{sheet_name}: Метрики",
      }
      if is_total_label(obj):
        total_row = item
        continue
      brands.append(item)

  total_messages = total_row.get("messages") or sum(b.get("messages") or 0 for b in brands) or None
  for item in brands:
    item["sov"] = (item.get("messages") or 0) / total_messages if total_messages else item.get("sovRaw")
    if norm_key(item["brand"]) == norm_key(brand):
        target = item

  if not brands:
    sheet_name, rows = find_sheet(analytics, "Показатели")
    header_index, header = table_header(rows, ["Объект", "Сообщения", "Аудитория"])
    if header_index is not None:
      for row in rows[header_index + 1:]:
        obj = norm_text(value_by_header(row, header, "Объект"))
        messages = parse_number(value_by_header(row, header, "Сообщения"))
        if not obj:
          continue
        if messages is None:
          if brands or total_row:
            break
          continue
        item = {
          "brand": obj,
          "messages": messages,
          "audience": parse_number(value_by_header(row, header, "Аудитория")),
          "views": parse_number(value_by_header(row, header, "Просмотры")),
          "engagement": parse_number(value_by_header(row, header, "Вовлеч")),
          "sovRaw": parse_number(value_by_header(row, header, "Share of voice", "SOV")),
          "source": f"{sheet_name}: Метрики",
        }
        if is_total_label(obj):
          total_row = item
          continue
        if is_context_metric_label(obj):
          continue
        brands.append(item)
        if norm_key(obj) == norm_key(brand):
          target = item

  if brands:
    total_messages = total_row.get("messages") or sum(b.get("messages") or 0 for b in brands) or total_messages
    for item in brands:
      item["sov"] = (item.get("messages") or 0) / total_messages if total_messages else item.get("sovRaw")
      if norm_key(item["brand"]) == norm_key(brand):
        target = item

  if not target:
    sheet_name, rows = find_sheet(analytics, "Показатели")
    header_index, header = table_header(rows, ["Сообщения", "Аудитория"])
    if header_index is not None:
      for row in rows[header_index + 1:]:
        if is_total_label(first_non_empty(row)):
          messages = parse_number(value_by_header(row, header, "Сообщения"))
          target = {
            "brand": brand,
            "messages": messages,
            "audience": parse_number(value_by_header(row, header, "Аудитория")),
            "views": parse_number(value_by_header(row, header, "Просмотры")),
            "engagement": parse_number(value_by_header(row, header, "Вовлеч")),
            "sov": (messages / total_messages) if messages is not None and total_messages else None,
            "source": f"{sheet_name}: Метрики",
          }
          break
  return brands, target, total_messages

def latest_sov_from_dynamic(analytics):
  sheet_name, rows = find_sheet(analytics, "Динамика SOV")
  values = {}
  period_label = ""
  section_start = None
  for idx, row in enumerate(rows):
    if norm_key(first_non_empty(row)) == "share of voice":
      section_start = idx + 1
      break
  if section_start is None:
    return values, period_label
  header_offset, header = table_header(rows[section_start:], ["Объект"])
  if header_offset is None:
    return values, period_label
  header_index = section_start + header_offset
  month_columns = [(idx, name) for idx, name in enumerate(header) if idx != column_index(header, "Объект") and norm_text(name)]
  if not month_columns:
    return values, period_label
  last_col, period_label = month_columns[-1]
  for row in rows[header_index + 1:]:
    label = norm_text(value_by_header(row, header, "Объект"))
    if not label:
      continue
    value = parse_number(row[last_col] if last_col < len(row) else "")
    if value is None:
      if values:
        break
      continue
    if is_total_label(label):
      continue
    values[norm_key(label)] = value
  return values, period_label

def dynamic_sov(analytics, brand):
  sheet_name, rows = find_sheet(analytics, "Динамика SOV")
  header_index, header = table_header(rows, ["Объект"])
  series = []
  if header_index is None:
    return series
  for row in rows[header_index + 1:]:
    label = norm_text(row[1] if len(row) > 1 and not norm_text(row[0]) else row[0] if row else "")
    if not label:
      continue
    values = []
    for idx, h in enumerate(header):
      if idx == 0 or not h:
        continue
      num = parse_number(row[idx] if idx < len(row) else "")
      if num is not None:
        values.append({"label": h, "value": num})
    if values:
      series.append({"name": label, "values": values})
  return series

def sentiment_data(analytics, brand=""):
  sheet_name, rows = find_sheet(analytics, "Показатели")
  header_index, header = table_header(rows, ["Позитивные", "Нейтральные", "Негативные", "Всего"])
  if header_index is not None:
    for row in rows[header_index + 1:]:
      if is_total_label(first_non_empty(row)):
        totals = {
          "Позитивные": parse_number(value_by_header(row, header, "Позитивные")) or 0,
          "Нейтральные": parse_number(value_by_header(row, header, "Нейтральные")) or 0,
          "Негативные": parse_number(value_by_header(row, header, "Негативные")) or 0,
        }
        total = sum(totals.values())
        return [{"label": k, "value": v, "share": (v / total if total else None)} for k, v in totals.items()]

  sheet_name, rows = find_sheet(analytics, "Тональность")
  header_index, header = table_header(rows, ["Позитив", "Нейтраль", "Негатив"])
  totals = {"Позитивные": 0, "Нейтральные": 0, "Негативные": 0}
  if header_index is not None:
    object_col = column_index(header, "Объект")
    pos = column_index(header, "Позитивные")
    neu = column_index(header, "Нейтральные")
    neg = column_index(header, "Негативные")
    brand_row = None
    if brand and object_col >= 0:
      for row in rows[header_index + 1:]:
        label = norm_text(row[object_col] if object_col < len(row) else "")
        if norm_key(label) == norm_key(brand):
          brand_row = row
          break
    source_rows = [brand_row] if brand_row is not None else rows[header_index + 1:]
    for row in source_rows:
      if row is None or is_total_label(first_non_empty(row)):
        continue
      if parse_number(value_by_header(row, header, "Позитивные")) is None and totals["Позитивные"]:
        break
      if pos >= 0:
        totals["Позитивные"] += parse_number(row[pos] if pos < len(row) else "") or 0
      if neu >= 0:
        totals["Нейтральные"] += parse_number(row[neu] if neu < len(row) else "") or 0
      if neg >= 0:
        totals["Негативные"] += parse_number(row[neg] if neg < len(row) else "") or 0
  total = sum(totals.values())
  return [{"label": k, "value": v, "share": (v / total if total else None)} for k, v in totals.items()]

def initiated_summary(analytics, orm):
  sheet_name, rows = find_sheet(analytics, "Инициированные сообщения")
  header_index, header = table_header(rows, ["Тип размещения", "Ссылка"])
  items = []
  if header_index is not None:
    for row in rows[header_index + 1:]:
      text = norm_text(value_by_header(row, header, "Текст сообщения"))
      url = norm_text(value_by_header(row, header, "Ссылка"))
      platform = norm_text(value_by_header(row, header, "Площадка"))
      post_type = norm_text(value_by_header(row, header, "Тип поста", "Тип размещения"))
      if text or url:
        items.append({"text": text, "url": url, "platform": platform, "type": post_type})
  if not items:
    for candidate_sheet, candidate_rows in orm.items():
      header_index, header = table_header(candidate_rows, ["Тип размещения", "Площадка", "Текст"])
      if header_index is None:
        continue
      for row in candidate_rows[header_index + 1:]:
        post_type = norm_text(value_by_header(row, header, "Тип размещения"))
        platform = norm_text(value_by_header(row, header, "Площадка"))
        text = norm_text(value_by_header(row, header, "Текст", "Текст сообщения"))
        url = norm_text(value_by_header(row, header, "Ссылка на сообщение", "Ссылка"))
        views = parse_number(value_by_header(row, header, "Просмотры"))
        if text or url:
          items.append({"text": text, "url": url, "platform": platform, "type": post_type, "views": views, "sheet": candidate_sheet})
      if items:
        break
  if not items:
    sheet_name, rows = find_sheet(orm, "Публикация")
    for row in rows[1:20]:
      joined = " ".join(norm_text(c) for c in row)
      if "http" in joined:
        items.append({"text": joined[:500], "url": "", "platform": "", "type": ""})
  return items

def build_payload(config):
  brand = config["brand"]
  period = config["period"]
  raw = read_xlsx(config["raw"])
  orm = read_xlsx(config["orm"])
  analytics = read_xlsx(config["analytics"])
  summary = read_docx(config.get("summary"))

  brands, target, total_category_messages = brand_metrics(analytics, brand)
  dynamic_sov_values, dynamic_sov_period = latest_sov_from_dynamic(analytics)
  if dynamic_sov_values:
    for item in brands:
      key = norm_key(item.get("brand"))
      if key in dynamic_sov_values:
        item["sov"] = dynamic_sov_values[key]
        item["sovSource"] = f"Динамика SOV: Share of voice / {dynamic_sov_period}"
    for item in brands:
      if norm_key(item.get("brand")) == norm_key(brand):
        target = item
        break
  sorted_brands = sorted([b for b in brands if b.get("messages") is not None], key=lambda x: x.get("messages") or 0, reverse=True)
  sov_sorted = sorted([b for b in brands if b.get("sov") is not None], key=lambda x: x.get("sov") or 0, reverse=True)
  platform_sheet, platform_rows = find_sheet(analytics, "Площадки")
  type_sheet, type_rows = find_sheet(analytics, "Типы площадок")
  groups_sheet, groups_rows = find_sheet(analytics, "Группы")
  socdem_sheet, socdem_rows = find_sheet(analytics, "Соцдем авторов")
  rating_sheet, rating_rows = find_sheet(analytics, "Оценки")
  positive_reviews_sheet, positive_reviews = find_sheet(analytics, "Отзывы позитив")
  negative_reviews_sheet, negative_reviews = find_sheet(analytics, "Отзывы негатив")
  initiated = initiated_summary(analytics, orm)
  sentiment = sentiment_data(analytics, brand)
  initiated_comments = [item for item in initiated if "комментар" in norm_key(item.get("type"))]
  initiated_reviews = [item for item in initiated if "отзыв" in norm_key(item.get("type"))]
  initiated_views = sum((item.get("views") or 0) for item in initiated if item.get("views") is not None)
  initiated_has_partial_views = any(item.get("views") is None for item in initiated)
  platform_pairs = pairs_from_cross_table(platform_rows, brand, "Площадки")[:8]
  type_pairs = pairs_from_cross_table(type_rows, brand, "Объект")[:6]
  rating_pairs = pairs_from_cross_table(rating_rows, brand, "Оценки")[:8]

  gender_pairs = []
  header_index, header = table_header(socdem_rows, ["Пол", "Сообщения"])
  if header_index is not None:
    for row in socdem_rows[header_index + 1:]:
      label = norm_text(row[1] if len(row) > 1 and not norm_text(row[0]) else row[0] if row else "")
      value = parse_number(value_by_header(row, header, "Сообщения"))
      if label and value is not None:
        gender_pairs.append({"label": label, "value": value})
      if len(gender_pairs) >= 4:
        break

  group_items = []
  header_index, header = table_header(groups_rows, ["Группа", "Сообщения"])
  if header_index is not None:
    msg_col = column_index(header, "Сообщения")
    platform_col = column_index(header, "Площадка")
    group_col = column_index(header, "Группа")
    for row in groups_rows[header_index + 1:]:
      label = norm_text(row[group_col] if 0 <= group_col < len(row) else "")
      value = parse_number(row[msg_col] if 0 <= msg_col < len(row) else "")
      platform = norm_text(row[platform_col] if 0 <= platform_col < len(row) else "")
      if label and value is not None:
        group_items.append({"label": label, "platform": platform, "value": value})
    group_items = sorted(group_items, key=lambda x: x["value"], reverse=True)[:6]

  target_sov = target.get("sov")
  target_rank = None
  for idx, item in enumerate(sov_sorted, 1):
    if item["brand"].lower() == brand.lower():
      target_rank = idx
      break

  top_platform = platform_pairs[0] if platform_pairs else None
  sentiment_total = sum(x["value"] for x in sentiment)
  positive_share = next((x["share"] for x in sentiment if x["label"] == "Позитивные"), None)
  negative_share = next((x["share"] for x in sentiment if x["label"] == "Негативные"), None)
  neutral_share = next((x["share"] for x in sentiment if x["label"] == "Нейтральные"), None)
  seeded_views_text = fmt_int(initiated_views) if initiated_views and not initiated_has_partial_views else "нет подтвержденной суммы"
  rating_total = sum((row.get("value") or 0) for row in rating_pairs)
  rating_five = next((row.get("value") or 0 for row in rating_pairs if "5" in str(row.get("label"))), 0)
  rating_five_share = (rating_five / rating_total) if rating_total else None

  slide_text = {
    "1": {
      "ИНФОРМАЦИОННАЯ КАМПАНИЯ": f"ИНФОРМАЦИОННАЯ КАМПАНИЯ\n{brand.upper()}",
      "Декабрь 2025": period,
    },
    "2": {
      "ПЕРИОД:": f"ПЕРИОД: {period.upper()}",
    },
    "3": {
      "УПОМИНАНИЯ БРЕНДОВ В МАЕ": "УПОМИНАНИЯ БРЕНДОВ В АПРЕЛЕ",
      "Бренды умеренно обсуждаются": (
        f"В брендовом конкурентном поле за апрель зафиксировано {fmt_int(total_category_messages)} сообщений. "
        f"{brand} получил {fmt_int(target.get('messages'))} сообщений и SOV {fmt_pct(target_sov)}, "
        f"заняв {target_rank or 'нет данных'}-е место по объему упоминаний. "
        f"Лидирует {sorted_brands[0]['brand'] if sorted_brands else 'нет данных'} "
        f"({fmt_int(sorted_brands[0]['messages']) if sorted_brands else 'нет данных'} сообщений). "
        "Отдельные тематические объекты не включены в SOV брендов, чтобы не смешивать брендовые и проблемные обсуждения."
      ),
      "Сообщения о брендах": f"Сообщения о брендах, {period}",
      "Аудитория сообщений OTS": f"Аудитория сообщений OTS, {period}",
    },
    "4": {
      "Хронический тонзиллит": (
        "Развитие речи и внимания\n"
        "Утомляемость после нагрузки\n"
        "Восстановление после болезни\n"
        "Назначения невролога\n"
        "Опыт курса и переносимость"
      ),
      "Темы заболеваний ЛОР-органов": (
        "Обсуждения категории сосредоточены в семейно-медицинском контексте: родители обсуждают речь, "
        "концентрацию, утомляемость, восстановление после болезни и назначения невролога. "
        "Для ORM это пространство поддержки через практичные ответы и опыт применения. "
        "Причинность кампании не заявляется: в данных нет валидного разреза «с кампанией / без кампании»."
      ),
      "Контексты обсуждений": "Контексты обсуждений брендов категории",
    },
    "5": {
      "33,4": fmt_pct(target_sov),
      "SoV": f"SoV бренда в брендовом конкурентном поле в апреле",
      "Благодаря работе бренда": (
        f"В апреле {brand} находится в середине конкурентного поля: {fmt_int(target.get('messages'))} сообщений "
        f"и SOV {fmt_pct(target_sov)}. Приоритет — наращивать присутствие в обсуждениях детского развития, "
        "внимания и восстановления после болезни. Данные «без кампании» отсутствуют, поэтому прирост от кампании не заявляется."
      ),
      "В результате информационной кампании": f"SOV {brand} в апреле: {fmt_pct(target_sov)} без причинного вывода о влиянии кампании",
      "+16,6": "н/д",
      "+10,9": "н/д",
      "SOV% брендов": f"SOV% брендов, {period}",
      "Источник данных": f"Источник данных: аналитический Excel и выгрузка мониторинга, {period}.",
    },
    "6": {
      "ВЫВОДЫ ПО АНАЛИЗУ": (
        f"ВЫВОДЫ ПО АНАЛИЗУ:\nУ {brand} {fmt_int(next((x['value'] for x in sentiment if x['label']=='Позитивные'), 0))} позитивных, "
        f"{fmt_int(next((x['value'] for x in sentiment if x['label']=='Нейтральные'), 0))} нейтральных и "
        f"{fmt_int(next((x['value'] for x in sentiment if x['label']=='Негативные'), 0))} негативное упоминание. "
        f"Доля позитива — {fmt_pct(positive_share)}, негатива — {fmt_pct(negative_share)}. "
        "Позитив строится вокруг опыта применения, речи, концентрации и восстановления после нагрузок. "
        "Тональность основана на автоматической разметке Медиалогии и требует ручной проверки спорных сообщений."
      ),
      "ТОНАЛЬНОСТЬ": f"ТОНАЛЬНОСТЬ СООБЩЕНИЙ О БРЕНДАХ КАТЕГОРИИ",
    },
    "7": {
      "Где обсуждают бренды": f"Где обсуждают {brand} и тематику?",
      "Наибольшая доля упоминаний": (
        f"Обсуждение {brand} сосредоточено в соцсетях и пользовательских сообществах: "
        f"{top_platform['label'] if top_platform else 'ключевая площадка'} — {fmt_int(top_platform['value']) if top_platform else 'нет данных'} сообщений. "
        "Основной вес дают живые родительские обсуждения, а не медийные публикации."
      ),
      "Источник данных": f"Источник данных: сервис мониторинга Медиалогия SM, 1–30 апреля 2026 года. Исключен спам и нерелевантные упоминания.",
    },
    "8": {
      "В декабре опубликовано": (
        f"В апреле в ORM-таблице учтено {fmt_int(len(initiated))} материалов: "
        f"{fmt_int(len(initiated_comments))} комментариев и {fmt_int(len(initiated_reviews))} отзывов. "
        f"Просмотры: {seeded_views_text}. Если часть строк без просмотров, суммарный охват не заявляется как подтвержденный KPI."
      ),
    },
    "9": {
      "Примеры СООБЩЕНИЙ": "ПРИМЕРЫ СООБЩЕНИЙ О БРЕНДЕ",
      "В декабре было зафиксировано": (
        f"В апреле в ORM-таблице учтено {fmt_int(len(initiated_comments))} комментариев о бренде {brand} "
        "в сообществах, форумах, чатах и материалах в топе выдачи. Для клиентской версии нужны реальные скриншоты источников."
      ),
    },
    "10": {
      "Примеры СООБЩЕНИЙ": "ПРИМЕРЫ ОТЗЫВОВ О БРЕНДЕ",
      "В декабре был опубликован": (
        f"В апреле в ORM-таблице учтено {fmt_int(len(initiated_reviews))} отзывов о бренде {brand}. "
        "В отзывах авторы описывают опыт курса, концентрацию, речь, утомляемость и восстановление после болезни. "
        "Для клиентской версии нужны реальные скриншоты отзывов."
      ),
    },
    "11": {
      "РЕЙТИНГИ Пикторид": f"РЕЙТИНГИ {brand} В ОТЗЫВАХ",
      "Детализация рейтингов": (
        "Рейтинги рассчитаны по сравнительному отчету Медиалогии, лист «Оценки». "
        "В локальных файлах нет стартовой динамики по карточкам, поэтому слайд показывает текущую структуру оценок за апрель."
      ),
      "Количество отзывов": f"Количество отзывов на площадках, {period}",
      "Тональность площадок": f"Тональность площадок с отзывами, {period}",
      "На момент старта": (
        f"В сравнительном отчете по {brand} за апрель найдено {fmt_int(rating_total)} оцененных отзывов, "
        f"из них {fmt_int(rating_five)} с оценкой 5. Доля оценок 5 — {fmt_pct(rating_five_share)}. "
        "Динамика к старту кампании не рассчитывается без исторического листа рейтингов."
      ),
    },
    "12": {
      "выводы": "выводы",
      "Бренды категории умеренно обсуждаются": "\n".join(summary[:4]) if summary else (
        f"{brand} получил {fmt_int(target.get('messages'))} брендовых сообщений в апреле и SOV {fmt_pct(target_sov)} внутри брендового конкурентного поля. "
        f"Бренд находится на {target_rank or 'нет данных'}-м месте: видимость есть, но лидеры категории существенно крупнее по объему обсуждений.\n\n"
        "Инфополе категории формируется вокруг родительских и медицинских сценариев: развитие речи, концентрация, утомляемость, восстановление после болезни, назначения невролога.\n\n"
        f"Тональность {brand} преимущественно позитивно-нейтральная: позитив {fmt_pct(positive_share)}, нейтральные {fmt_pct(neutral_share)}, негатив {fmt_pct(negative_share)}. "
        "Спорные сообщения о назначениях, составе, цене и эффективности требуют ручной проверки.\n\n"
        f"В ORM-таблице за апрель учтено {fmt_int(len(initiated))} материалов: {fmt_int(len(initiated_comments))} комментариев и {fmt_int(len(initiated_reviews))} отзывов. "
        "Просмотры заполнены не по всем строкам, поэтому охват посева не заявляется как подтвержденный KPI.\n\n"
        f"По отзывам видно {fmt_int(rating_total)} оцененных карточек {brand}, из них {fmt_int(rating_five)} с оценкой 5. "
        "Динамику рейтингов от старта кампании нужно подтверждать через Google Sheet или отдельный исторический лист."
      ),
    },
  }

  charts = {
    "3": [
      [{"name": "Сообщения", "points": [{"label": b["brand"], "value": b.get("messages") or 0} for b in sorted_brands[:8]]}],
      [{"name": "Аудитория", "points": [{"label": b["brand"], "value": b.get("audience") or 0} for b in sorted(brands, key=lambda x: x.get("audience") or 0, reverse=True)[:8]]}],
    ],
    "4": [
      [{"name": "SOV", "points": [{"label": b["brand"], "value": (b.get("sov") or 0) * 100 if abs(b.get("sov") or 0) <= 1 else b.get("sov") or 0} for b in sov_sorted[:8]]}],
      [{"name": "Сообщения", "points": [{"label": b["brand"], "value": b.get("messages") or 0} for b in sorted_brands[:6]]}],
    ],
    "5": [
      [{"name": "Типы площадок", "points": type_pairs[:6]}],
      [{"name": "Топ источников", "points": platform_pairs[:8]}],
    ],
    "6": [
      [{"name": "Посев", "points": [{"label": x["type"] or x["platform"] or "Размещение", "value": 1} for x in initiated[:8]]}],
    ],
    "9": [
      [{"name": "Тональность", "points": [{"label": x["label"], "value": x["value"]} for x in sentiment]}],
      [{"name": "Доля", "points": [{"label": x["label"], "value": (x["share"] or 0) * 100} for x in sentiment]}],
    ],
    "10": [
      [{"name": "Пол", "points": gender_pairs or [{"label": "нет данных", "value": 0}]}],
      [{"name": "Авторы", "points": gender_pairs or [{"label": "нет данных", "value": 0}]}],
    ],
    "11": [
      [{"name": "Оценки", "points": rating_pairs or [{"label": "нет данных", "value": 0}]}],
      [{"name": "Отзывы", "points": [
        {"label": "Позитивные", "value": max(0, len(positive_reviews) - 2)},
        {"label": "Негативные", "value": max(0, len(negative_reviews) - 2)},
      ]}],
      [{"name": "Тональность", "points": [{"label": x["label"], "value": x["value"]} for x in sentiment]}],
    ],
  }

  return {
    "brand": brand,
    "period": period,
    "inputs": {
      "raw": {"path": config["raw"], "sheets": list(raw.keys())},
      "orm": {"path": config["orm"], "sheets": list(orm.keys())},
      "analytics": {"path": config["analytics"], "sheets": list(analytics.keys())},
      "summary": {"path": config.get("summary"), "paragraphs": summary[:8]},
    },
    "metrics": {
      "brand": target,
      "brandRankBySov": target_rank,
      "categoryMessages": total_category_messages,
      "platforms": platform_pairs,
      "platformTypes": type_pairs,
      "groups": group_items,
      "sentiment": sentiment,
      "gender": gender_pairs,
      "ratings": rating_pairs,
      "initiatedCount": len(initiated),
    },
    "slideText": slide_text,
    "charts": charts,
  }

def pptx_inventory(path):
  with zipfile.ZipFile(path) as zf:
    names = zf.namelist()
    slides = sorted([n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)], key=lambda n: int(re.search(r"(\d+)", n).group(1)))
    charts = sorted([n for n in names if re.match(r"ppt/charts/chart\d+\.xml$", n)], key=lambda n: int(re.search(r"(\d+)", n).group(1)))
    media = [n for n in names if n.startswith("ppt/media/")]
    slide_texts = {}
    for idx, slide in enumerate(slides, 1):
      root = ET.fromstring(zf.read(slide))
      texts = []
      for shape in root.findall(".//p:sp", NS):
        text = "".join(t.text or "" for t in shape.findall(".//a:t", NS))
        if norm_text(text):
          texts.append(norm_text(text))
      for table in root.findall(".//a:tbl", NS):
        text = " ".join(norm_text(t.text) for t in table.findall(".//a:t", NS) if norm_text(t.text))
        if text:
          texts.append(norm_text(text))
      slide_texts[str(idx)] = texts
    return {"slides": len(slides), "charts": len(charts), "media": len(media), "slideTexts": slide_texts}

def rel_targets(zf, rels_path):
  if rels_path not in zf.namelist():
    return {}
  root = ET.fromstring(zf.read(rels_path))
  out = {}
  base = str(Path(rels_path).parent.parent).replace("\\", "/")
  for rel in root.findall("pr:Relationship", NS):
    target = rel.get("Target") or ""
    if target.startswith("../"):
      full = str(Path(base, target).as_posix())
      while "/../" in full:
        full = re.sub(r"[^/]+/\.\./", "", full, count=1)
      target = full
    elif not target.startswith("/"):
      target = str(Path(base, target).as_posix())
    else:
      target = target.lstrip("/")
    out[rel.get("Id")] = {"type": rel.get("Type") or "", "target": target}
  return out

def replace_shape_text(root, replacements, brand, period, log, slide_no):
  for shape in root.findall(".//p:sp", NS):
    runs = shape.findall(".//a:t", NS)
    if not runs:
      continue
    original = "".join(r.text or "" for r in runs)
    normalized = norm_text(original)
    new_text = None
    for needle, replacement in sorted(replacements.items(), key=lambda item: len(item[0]), reverse=True):
      if needle in normalized:
        new_text = replacement
        break
    if new_text is None:
      new_text = original
      new_text = re.sub(r"ПИКТОРИД®|ПИКТОРИД|Пикторид®|Пикторид", brand, new_text)
      new_text = re.sub(r"Декабрь 2025|декабрь 2025|декабре 2025|декабря 2025|1 – 31 декабря 2025 года", period, new_text)
      if new_text == original:
        continue
    runs[0].text = new_text
    for extra in runs[1:]:
      extra.text = ""
    log.append({"slide": slide_no, "type": "text", "from": normalized[:120], "to": norm_text(new_text)[:160]})

def replace_inline_text(root, brand, period, log, slide_no):
  month = period_month(period)
  month_cap = period_month(period, True)
  prev_month = previous_period_month(period)
  prev_month_cap = previous_period_month(period, True)
  year = period_year(period)
  replacements = [
    (r"ПИКТОРИД®|ПИКТОРИД|Пикторид®|Пикторид", brand),
    (r"Декабрь 2025|декабрь 2025", period),
    (r"1\s*[–-]\s*31 декабря 2025 года", period),
    (r"Ноябрь", prev_month_cap),
    (r"ноябрь|ноябре|ноября", prev_month),
    (r"Декабрь", month_cap),
    (r"декабрь", month),
    (r"декабре", month),
    (r"декабря", month),
  ]
  if year:
    replacements.append((r"2025", year))
  for node in root.findall(".//a:t", NS):
    original = node.text or ""
    new_text = original
    for pattern, replacement in replacements:
      new_text = re.sub(pattern, replacement, new_text)
    if new_text != original:
      node.text = new_text
      log.append({"slide": slide_no, "type": "inline-text", "from": norm_text(original)[:120], "to": norm_text(new_text)[:160]})

def remove_creation_ids(root):
  for parent in root.findall(".//a:extLst", NS):
    for ext in list(parent):
      if ext.find("a16:creationId", NS) is not None:
        parent.remove(ext)

def clear_pictures_if_missing(root, slide_no, screenshots, log):
  if slide_no not in (7, 8):
    return
  tree = root.find(".//p:spTree", NS)
  if tree is None:
    return
  pics = [child for child in list(tree) if child.tag == "{%s}pic" % NS["p"]]
  if screenshots:
    return
  for pic in pics:
    tree.remove(pic)
  if pics:
    log.append({
      "slide": slide_no,
      "type": "image",
      "status": "draft",
      "message": f"Удалены старые изображения эталона: {len(pics)}. Для клиентской версии нужен --screenshots-dir.",
    })

def set_cache_points(cache, points, value_key):
  if cache is None:
    return
  for child in list(cache):
    if child.tag in ("{%s}pt" % NS["c"], "{%s}ptCount" % NS["c"]):
      cache.remove(child)
  count = ET.SubElement(cache, "{%s}ptCount" % NS["c"])
  count.set("val", str(len(points)))
  for idx, point in enumerate(points):
    pt = ET.SubElement(cache, "{%s}pt" % NS["c"])
    pt.set("idx", str(idx))
    v = ET.SubElement(pt, "{%s}v" % NS["c"])
    v.text = str(point[value_key])

def ensure_cache(parent_path, ser, cache_name):
  node = ser.find(parent_path, NS)
  if node is None:
    return None
  cache = node.find("c:" + cache_name, NS)
  if cache is None:
    cache = ET.SubElement(node, "{%s}%s" % (NS["c"], cache_name))
  return cache

def set_ref_formula(ref_node, formula):
  if ref_node is None:
    return
  formula_node = ref_node.find("c:f", NS)
  if formula_node is None:
    formula_node = ET.Element("{%s}f" % NS["c"])
    ref_node.insert(0, formula_node)
  formula_node.text = formula

def update_chart_xml(xml_bytes, series_payload, log, chart_name):
  root = ET.fromstring(xml_bytes)
  series_nodes = root.findall(".//c:ser", NS)
  if not series_nodes or not series_payload:
    return xml_bytes
  parent_map = {child: parent for parent in root.iter() for child in parent}
  for idx, ser in enumerate(list(series_nodes)):
    if idx >= len(series_payload):
      parent = parent_map.get(ser)
      if parent is not None:
        parent.remove(ser)
      continue
    payload = series_payload[idx]
    points = payload.get("points") or [{"label": "нет данных", "value": 0}]
    points = [{"label": norm_text(p.get("label")) or "нет данных", "value": parse_number(p.get("value")) or 0} for p in points]
    tx_ref = ser.find("c:tx/c:strRef", NS)
    if tx_ref is not None:
      set_ref_formula(tx_ref, "'Лист1'!$" + idx_to_col(idx + 1) + "$1")
    tx_cache = ensure_cache("c:tx/c:strRef", ser, "strCache")
    if tx_cache is not None:
      set_cache_points(tx_cache, [{"label": payload.get("name") or "Данные"}], "label")
    cat_ref = ser.find("c:cat/c:strRef", NS)
    if cat_ref is not None:
      set_ref_formula(cat_ref, "'Лист1'!$A$2:$A$" + str(len(points) + 1))
    val_ref = ser.find("c:val/c:numRef", NS)
    if val_ref is not None:
      value_col = idx_to_col(idx + 1)
      set_ref_formula(val_ref, "'Лист1'!$" + value_col + "$2:$" + value_col + "$" + str(len(points) + 1))
    cat_cache = ensure_cache("c:cat/c:strRef", ser, "strCache")
    if cat_cache is None:
      cat_cache = ensure_cache("c:cat/c:strLit", ser, "strCache")
    val_cache = ensure_cache("c:val/c:numRef", ser, "numCache")
    if val_cache is None:
      val_cache = ensure_cache("c:val/c:numLit", ser, "numCache")
    if cat_cache is not None:
      set_cache_points(cat_cache, points, "label")
    if val_cache is not None:
      set_cache_points(val_cache, points, "value")
  log.append({"type": "chart", "chart": chart_name, "series": len(series_payload)})
  return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def chart_paths_for_slide(zf, slide_no):
  rels_path = f"ppt/slides/_rels/slide{slide_no}.xml.rels"
  rels = rel_targets(zf, rels_path)
  out = []
  for item in rels.values():
    if item["type"].endswith("/chart"):
      out.append(item["target"])
  return out

def embedded_workbook_for_chart(zf, chart_path):
  rels_path = f"ppt/charts/_rels/{Path(chart_path).name}.rels"
  rels = rel_targets(zf, rels_path)
  root = ET.fromstring(zf.read(chart_path))
  external = root.find(".//c:externalData", NS)
  if external is None:
    return None
  rid = external.get("{%s}id" % NS["r"])
  target = rels.get(rid, {}).get("target")
  return target if target and target.startswith("ppt/embeddings/") else None

def normalize_embedded_text(data, brand, period):
  try:
    text = data.decode("utf-8")
  except Exception:
    return data
  month = period_month(period)
  month_cap = period_month(period, True)
  prev_month = previous_period_month(period)
  prev_month_cap = previous_period_month(period, True)
  year = period_year(period)
  replacements = [
    (r"ПИКТОРИД®|ПИКТОРИД|Пикторид®|Пикторид", brand),
    (r"Декабрь 2025|декабрь 2025", period),
    (r"Ноябрь", prev_month_cap),
    (r"ноябрь|ноябре|ноября", prev_month),
    (r"Декабрь", month_cap),
    (r"декабрь|декабре|декабря", month),
  ]
  if year:
    replacements.append((r"2025", year))
  for pattern, replacement in replacements:
    text = re.sub(pattern, replacement, text)
  return text.encode("utf-8")

def write_inline_cell(row_node, row_index, col_index, value):
  cell = ET.SubElement(row_node, "{%s}c" % NS["m"])
  cell.set("r", f"{idx_to_col(col_index)}{row_index}")
  cell.set("t", "inlineStr")
  inline = ET.SubElement(cell, "{%s}is" % NS["m"])
  text = ET.SubElement(inline, "{%s}t" % NS["m"])
  text.text = norm_text(value)
  return cell

def write_number_cell(row_node, row_index, col_index, value):
  cell = ET.SubElement(row_node, "{%s}c" % NS["m"])
  cell.set("r", f"{idx_to_col(col_index)}{row_index}")
  number = ET.SubElement(cell, "{%s}v" % NS["m"])
  number.text = str(parse_number(value) or 0)
  return cell

def update_embedded_workbook(xlsx_bytes, series_payload, brand, period, log, workbook_name):
  if not series_payload:
    return xlsx_bytes
  with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zin:
    names = zin.namelist()
    workbook = ET.fromstring(zin.read("xl/workbook.xml"))
    rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.get("Id"): rel.get("Target") for rel in rels.findall("pr:Relationship", NS)}
    first_sheet = workbook.find(".//m:sheet", NS)
    if first_sheet is None:
      return xlsx_bytes
    rid = first_sheet.get("{%s}id" % NS["r"])
    target = (rel_map.get(rid) or "").lstrip("/")
    sheet_path = "xl/" + target
    if sheet_path not in names:
      sheet_path = "xl/worksheets/" + Path(target).name
    if sheet_path not in names:
      return xlsx_bytes

    series = []
    for item in series_payload:
      points = item.get("points") or [{"label": "нет данных", "value": 0}]
      series.append({
        "name": norm_text(item.get("name")) or "Данные",
        "points": [{"label": norm_text(p.get("label")) or "нет данных", "value": parse_number(p.get("value")) or 0} for p in points],
      })
    max_points = max(len(item["points"]) for item in series)

    sheet_root = ET.fromstring(zin.read(sheet_path))
    dimension = sheet_root.find("m:dimension", NS)
    if dimension is not None:
      dimension.set("ref", f"A1:{idx_to_col(len(series))}{max_points + 1}")
    sheet_data = sheet_root.find("m:sheetData", NS)
    if sheet_data is None:
      sheet_data = ET.SubElement(sheet_root, "{%s}sheetData" % NS["m"])
    for child in list(sheet_data):
      sheet_data.remove(child)

    header = ET.SubElement(sheet_data, "{%s}row" % NS["m"])
    header.set("r", "1")
    write_inline_cell(header, 1, 0, "Источник")
    for idx, item in enumerate(series, start=1):
      write_inline_cell(header, 1, idx, item["name"])

    for point_index in range(max_points):
      row_index = point_index + 2
      row_node = ET.SubElement(sheet_data, "{%s}row" % NS["m"])
      row_node.set("r", str(row_index))
      label = series[0]["points"][point_index]["label"] if point_index < len(series[0]["points"]) else "нет данных"
      write_inline_cell(row_node, row_index, 0, label)
      for series_index, item in enumerate(series, start=1):
        value = item["points"][point_index]["value"] if point_index < len(item["points"]) else 0
        write_number_cell(row_node, row_index, series_index, value)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
      for name in names:
        data = zin.read(name)
        if name == sheet_path:
          data = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
        elif name in ("xl/sharedStrings.xml", "xl/styles.xml"):
          data = normalize_embedded_text(data, brand, period)
        zout.writestr(name, data)
  log.append({"type": "embedded-workbook", "workbook": workbook_name, "series": len(series), "points": max_points})
  return out.getvalue()

def edit_pptx(template, output, payload, screenshots_dir):
  output = Path(output)
  output.parent.mkdir(parents=True, exist_ok=True)
  screenshots = []
  if screenshots_dir:
    screenshots = [p for p in sorted(Path(screenshots_dir).glob("*")) if p.suffix.lower() in (".png", ".jpg", ".jpeg")]
  log = []
  chart_update_log = []
  workbook_update_log = []
  old_brand_hits = []
  with zipfile.ZipFile(template) as zin, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
    names = zin.namelist()
    slide_chart_map = {}
    for slide_no in range(1, 13):
      slide_chart_map[str(slide_no)] = chart_paths_for_slide(zin, slide_no)
    chart_payload_by_path = {}
    workbook_payload_by_path = {}
    for slide_no, chart_payloads in payload["charts"].items():
      for idx, chart_path in enumerate(slide_chart_map.get(str(slide_no), [])):
        if idx < len(chart_payloads):
          chart_payload_by_path[chart_path] = chart_payloads[idx]
          workbook_path = embedded_workbook_for_chart(zin, chart_path)
          if workbook_path:
            workbook_payload_by_path[workbook_path] = chart_payloads[idx]

    for name in names:
      data = zin.read(name)
      slide_match = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
      if slide_match:
        slide_no = int(slide_match.group(1))
        root = ET.fromstring(data)
        remove_creation_ids(root)
        replace_shape_text(
          root,
          payload["slideText"].get(str(slide_no), {}),
          payload["brand"],
          payload["period"],
          log,
          slide_no,
        )
        replace_inline_text(root, payload["brand"], payload["period"], log, slide_no)
        clear_pictures_if_missing(root, slide_no, screenshots, log)
        data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
      elif name in chart_payload_by_path:
        data = update_chart_xml(data, chart_payload_by_path[name], chart_update_log, name)
      elif name in workbook_payload_by_path:
        data = update_embedded_workbook(data, workbook_payload_by_path[name], payload["brand"], payload["period"], workbook_update_log, name)
      zout.writestr(name, data)

  inventory = pptx_inventory(output)
  for slide, texts in inventory["slideTexts"].items():
    for text in texts:
      if "Пикторид" in text or "ПИКТОРИД" in text:
        old_brand_hits.append({"slide": slide, "text": text[:200]})
  return {
    "output": str(output),
    "outputBytes": output.stat().st_size,
    "inventory": inventory,
    "textUpdates": log,
    "chartUpdates": chart_update_log,
    "workbookUpdates": workbook_update_log,
    "oldBrandTextHits": old_brand_hits,
    "screenshotsProvided": len(screenshots),
  }

def main():
  config = json.loads(os.environ.get("REPORT_PPTX_CONFIG") or sys.stdin.read())
  payload = build_payload(config)
  before = pptx_inventory(config["template"])
  result = edit_pptx(config["template"], config["output"], payload, config.get("screenshotsDir"))
  blockers = []
  warnings = []
  if before["slides"] != 12 or result["inventory"]["slides"] != 12:
    blockers.append("В шаблоне или результате не 12 слайдов.")
  if before["charts"] != result["inventory"]["charts"]:
    blockers.append(f"Количество нативных графиков изменилось: было {before['charts']}, стало {result['inventory']['charts']}.")
  if before["media"] != result["inventory"]["media"] and not config.get("screenshotsDir"):
    warnings.append("Медиа-слоты изменились из-за удаления старых скриншотов эталона на слайдах 7-8.")
  if result["oldBrandTextHits"]:
    blockers.append("В текстовом слое остались упоминания Пикторид.")
  if not config.get("screenshotsDir"):
    blockers.append("Слайды 7-8 не клиентски готовы: локальные скриншоты не переданы через --screenshots-dir.")
    warnings.append("Слайды 7-8 сохранены как internal draft без старых изображений эталона.")

  log = {
    "status": "blocked" if blockers and config.get("strict", True) else "draft" if (blockers or warnings) else "ready_for_pm_review",
    "mode": "native_pptx_clone_edit",
    "note": "Слайды клонируются из эталонного PPTX; графики остаются нативными PPTX-объектами. Выводы сформированы из входных данных, не скопированы из эталона.",
    "blockers": blockers,
    "warnings": warnings,
    "templateInventory": before,
    "result": result,
    "payload": payload,
  }
  log_path = Path(config["output"]).with_suffix(".generation-log.json")
  log_path.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
  print(json.dumps({"output": config["output"], "log": str(log_path), "status": log["status"], "blockers": blockers, "warnings": warnings}, ensure_ascii=False, indent=2))
  if blockers and config.get("strict", True):
    sys.exit(2)

main()
`;

function q(value) {
  return JSON.stringify(String(value ?? ""));
}

function currentRepoRoot() {
  return path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
}

function buildPythonPipelineYaml(config, args, buildDir) {
  const sourceSystem = String(args["source-system"] || "Brand Analytics");
  const campaignOverride = args["campaign-materials-override"] === undefined
    ? ""
    : String(args["campaign-materials-override"]);
  const campaignMonthSheet = args["campaign-month-sheet"] === undefined
    ? ""
    : String(args["campaign-month-sheet"]);
  const ratingsSheet = args["ratings-sheet"] === undefined
    ? "Рейтинги"
    : String(args["ratings-sheet"]);
  const ratingsCurrentPeriod = args["ratings-current-period"] === undefined
    ? ""
    : String(args["ratings-current-period"]);
  const screenshotBackend = args["screenshot-backend"] === undefined
    ? ""
    : String(args["screenshot-backend"]);
  const sourceExclusions = args["source-exclusions"] === undefined
    ? ""
    : String(args["source-exclusions"]);
  const campaignMatchedAnalytics = args["campaign-matched-analytics"] === undefined
    ? ""
    : String(args["campaign-matched-analytics"]);
  const rawForSlide04 = config.raw || config.analytics;
  const inputFiles = [
    config.analytics,
    config.raw,
    config.orm,
    config.summary,
    config.template,
  ].filter(Boolean);
  const uniqueInputFiles = [...new Set(inputFiles)];
  return [
    "project:",
    `  brand: ${q(config.brand)}`,
    `  client: ""`,
    `  period: ${q(config.period)}`,
    `  category: ""`,
    `  report_type: "monthly_orm_report"`,
    "",
    "paths:",
    `  template_pptx: ${q(config.template)}`,
    `  input_dir: ${q(path.dirname(config.analytics || config.raw || config.orm))}`,
    `  output_dir: ${q(buildDir)}`,
    "",
    "inputs:",
    "  files:",
    ...uniqueInputFiles.map((filePath) => `    - ${q(filePath)}`),
    "  exclude_patterns:",
    `    - "~$*"`,
    "",
    "processing:",
    "  max_sample_rows_per_sheet: 15",
    "  max_sample_cols_per_sheet: 18",
    "  detect_file_roles: true",
    "  analyze_pptx_template: true",
    "  generate_missing_data_report: true",
    "  generate_slide_readiness: true",
    "",
    "rules:",
    "  do_not_copy_template_conclusions: true",
    "  do_not_invent_data: true",
    "  require_sources_for_claims: true",
    `  insight_formula: "факт -> интерпретация -> действие -> оговорка"`,
    `  source_exclusions: ${sourceExclusions ? JSON.stringify(sourceExclusions.split(",").map((s) => s.trim())) : "[]"}`,
    "",
    "build:",
    "  first4_combined_test: true",
    "  slide01_cover_test: true",
    "  slide02_plan_fact_test: true",
    `  slide02_media_plan_excel: ${q(config.orm)}`,
    `  slide02_fact_excel: ${q(config.orm)}`,
    `  slide02_fact_month_sheet: ${q(campaignMonthSheet)}`,
    `  slide02_fact_source_system: ${q("проектная таблица ORM")}`,
    "  slide03_brand_mentions_test: true",
    `  slide03_analytics_excel: ${q(config.analytics)}`,
    `  slide03_analytical_context_excel: ${q(config.analytics)}`,
    `  slide03_source_system: ${q(sourceSystem)}`,
    "  slide04_problem_field_test: true",
    `  slide04_analytical_excel: ""`,
    `  slide04_raw_excel: ${q(rawForSlide04)}`,
    `  slide04_report_excel: ""`,
    `  slide04_source_system: ${q(sourceSystem)}`,
    "  slide05_sov_test: true",
    `  slide05_analytics_excel: ${q(config.analytics)}`,
    `  slide05_project_orm_excel: ${q(config.orm)}`,
    `  slide05_campaign_month_sheet: ${q(campaignMonthSheet)}`,
    `  slide05_campaign_materials_override: ${campaignOverride ? q(campaignOverride) : "null"}`,
    `  slide05_source_system: ${q(sourceSystem)}`,
    "  slide06_tonality_test: true",
    `  slide06_analytics_excel: ${q(config.analytics)}`,
    `  slide06_source_system: ${q(sourceSystem)}`,
    "  slide07_sources_test: true",
    `  slide07_analytics_excel: ${q(config.analytics)}`,
    `  slide07_source_system: ${q(sourceSystem)}`,
    "  slide08_seeding_metrics_test: true",
    `  slide08_project_orm_excel: ${q(config.orm)}`,
    `  slide08_month_sheet: ${q(campaignMonthSheet || "Май")}`,
    `  slide08_source_system: ${q("проектная таблица ORM")}`,
    "  slide09_message_examples_test: true",
    `  slide09_project_orm_excel: ${q(config.orm)}`,
    `  slide09_month_sheet: ${q(campaignMonthSheet || "Май")}`,
    `  slide09_source_system: ${q("проектная таблица ORM")}`,
    `  slide09_10_screenshot_backend: ${q(screenshotBackend)}`,
    "  slide10_review_examples_test: true",
    `  slide10_project_orm_excel: ${q(config.orm)}`,
    `  slide10_month_sheet: ${q(campaignMonthSheet || "Май")}`,
    `  slide10_source_system: ${q("проектная таблица ORM")}`,
    "  slide11_seeding_metrics_test: true",
    `  slide11_ratings_excel: ${q(config.orm)}`,
    `  slide11_ratings_sheet: ${q(ratingsSheet)}`,
    `  slide11_current_period: ${q(ratingsCurrentPeriod)}`,
    `  slide11_placement_month_sheet: ${q(campaignMonthSheet || "Май")}`,
    `  slide11_source_system: ${q("проектная таблица ORM")}`,
    "  slide12_final_conclusions_test: true",
    `  campaign_matched_in_analytics_override: ${campaignMatchedAnalytics ? q(campaignMatchedAnalytics) : "null"}`,
    "",
  ].join("\n");
}

function readJsonIfExists(filePath) {
  if (!fs.existsSync(filePath)) return null;
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function collectPythonPipelineQa(buildDir, slidesReady) {
  const blockers = [];
  const warnings = [];
  const bySlide = [];
  for (let slide = 1; slide <= 12; slide += 1) {
    const id = String(slide).padStart(2, "0");
    const qa = readJsonIfExists(path.join(buildDir, `slide_${id}_qa.json`));
    const result = readJsonIfExists(path.join(buildDir, `slide_${id}_build_result.json`));
    const slideBlockers = [
      ...(qa?.blockers || []),
      ...(result?.blockers || []),
      ...(result?.qa_blockers || []),
    ];
    const slideWarnings = [
      ...(qa?.warnings || []),
      ...(result?.warnings || []),
      ...(result?.qa_warnings || []),
    ];
    const chartUpdates = result?.chart_updates || [];
    const chartImageOutputs = result?.chart_image_outputs || [];
    if (
      (Array.isArray(chartImageOutputs) && chartImageOutputs.length) ||
      (Array.isArray(chartUpdates) && chartUpdates.some((item) => String(item?.type || "").includes("image_fallback")))
    ) {
      slideWarnings.push("Uses rendered chart image fallback; native chart parts are preserved, but editable chart data update is not confirmed.");
    }
    if (qa || result) {
      bySlide.push({
        slide,
        status: qa?.status || result?.status || result?.qa_status || "not_reported",
        blockers: slideBlockers,
        warnings: slideWarnings,
      });
    }
    for (const blocker of slideBlockers) blockers.push(`Slide ${id}: ${blocker}`);
    for (const warning of slideWarnings) warnings.push(`Slide ${id}: ${warning}`);
  }

  const combined = readJsonIfExists(path.join(buildDir, `slides_${slidesReady.replace("-", "_")}_combined_result.json`));
  for (const blocker of combined?.qa_blockers || combined?.blockers || []) {
    blockers.push(String(blocker));
  }
  for (const warning of combined?.qa_warnings || combined?.warnings || []) {
    warnings.push(String(warning));
  }

  const crossSlideQa = readJsonIfExists(path.join(buildDir, "cross_slide_qa.json"));
  if (crossSlideQa) {
    for (const blocker of crossSlideQa.blockers || []) {
      blockers.push(`cross-slide: ${String(blocker)}`);
    }
    for (const warning of crossSlideQa.warnings || []) {
      warnings.push(`cross-slide: ${String(warning)}`);
    }
  }

  return {
    status: blockers.length ? "blocked" : warnings.length ? "draft" : "ready_for_pm_review",
    blockers: [...new Set(blockers)],
    warnings: [...new Set(warnings)],
    bySlide,
  };
}

function runPythonPipeline(config, args) {
  const pipelineRoot = args["pipeline-root"]
    ? path.resolve(String(args["pipeline-root"]))
    : path.join(currentRepoRoot(), "ai_presentation_orm_v0_3_13");
  const runPipeline = path.join(pipelineRoot, "run_pipeline.py");
  ensureExists(runPipeline, "python pipeline run_pipeline.py");

  const outputPath = path.resolve(config.output);
  const buildDir = path.join(path.dirname(outputPath), `${path.basename(outputPath, path.extname(outputPath))}_ai_orm_build`);
  fs.mkdirSync(buildDir, { recursive: true });
  const generatedConfig = path.join(buildDir, "config.generated.ai_orm_manager.yaml");
  fs.writeFileSync(generatedConfig, buildPythonPipelineYaml(config, args, buildDir), "utf8");

  const python = findPython();
  const pythonPath = process.env.PYTHONPATH || "";
  const result = spawnSync(python, [runPipeline, "--config", generatedConfig], {
    cwd: pipelineRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      PYTHONIOENCODING: "utf-8",
      PYTHONUTF8: "1",
      PYTHONPATH: pythonPath + (pythonPath ? ";" : "") + pipelineRoot,
      PIP_REQUIRE_VIRTUALENV: "false",
    },
    maxBuffer: 1024 * 1024 * 80,
  });

  if (result.stdout) process.stdout.write(result.stdout);
  if (result.stderr) process.stderr.write(result.stderr);
  if (result.status !== 0) {
    throw new Error(`Python pipeline завершился со статусом ${result.status}. Сборка: ${buildDir}`);
  }

  const producedCandidates = [
    "draft_slides_01_12_combined_test.pptx",
    "draft_slides_01_11_combined_test.pptx",
    "draft_slides_01_08_combined_test.pptx",
    "draft_slides_01_07_combined_test.pptx",
    "draft_slides_01_06_combined_test.pptx",
    "draft_slides_01_05_combined_test.pptx",
    "draft_slides_01_04_combined_test.pptx",
  ].map((fileName) => path.join(buildDir, fileName));
  const produced = producedCandidates.find((candidate) => fs.existsSync(candidate));
  const producedMatch = produced ? path.basename(produced).match(/draft_slides_(\d+_\d+)_combined_test\.pptx$/) : null;
  const slidesReady = producedMatch ? producedMatch[1].replace("_", "-") : "unknown";
  ensureExists(produced, "python pipeline output");
  fs.copyFileSync(produced, outputPath);
  const qa = collectPythonPipelineQa(buildDir, slidesReady);

  const crossSlideQaData = readJsonIfExists(path.join(buildDir, "cross_slide_qa.json"));
  const bridgeLog = {
    status: qa.status,
    mode: "ai_presentation_orm_v0_3_13_python_pipeline",
    output: outputPath,
    buildDir,
    generatedConfig,
    sourceDeck: produced,
    slidesReady,
    blockers: qa.blockers,
    warnings: qa.warnings,
    qaBySlide: qa.bySlide,
    crossSlideQa: crossSlideQaData || null,
  };
  const logPath = outputPath.replace(/\.pptx$/i, ".generation-log.json");
  fs.writeFileSync(logPath, JSON.stringify(bridgeLog, null, 2), "utf8");
  console.log(JSON.stringify({ output: outputPath, log: logPath, status: bridgeLog.status, buildDir }, null, 2));
  if (qa.blockers.length && config.strict) {
    throw new Error(`Python pipeline собрал draft, но QA заблокировал клиентскую готовность. Блокеры: ${qa.blockers.join(" | ")}`);
  }
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log(usage());
    return;
  }

  const config = {
    template: requirePath(args, "template"),
    raw: requirePath(args, "raw"),
    orm: requirePath(args, "orm"),
    analytics: requirePath(args, "analytics"),
    summary: optionalPath(args, "summary"),
    screenshotsDir: optionalPath(args, "screenshots-dir"),
    brand: String(args.brand || "").trim(),
    period: String(args.period || "").trim(),
    output: requirePath(args, "output"),
    strict: args.strict === undefined ? true : String(args.strict).toLowerCase() !== "false",
  };

  if (!config.brand) throw new Error(`Не указан --brand.\n${usage()}`);
  if (!config.period) throw new Error(`Не указан --period.\n${usage()}`);
  ensureExists(config.template, "template");
  ensureExists(config.raw, "raw");
  ensureExists(config.orm, "orm");
  ensureExists(config.analytics, "analytics");
  ensureExists(config.summary, "summary", false);
  if (config.screenshotsDir && !fs.existsSync(config.screenshotsDir)) {
    throw new Error(`Папка screenshots-dir не найдена: ${config.screenshotsDir}`);
  }

  const engine = String(args.engine || "legacy").toLowerCase();
  if (["python", "python-v0.3.13", "ai-orm", "ai-orm-v0.3.13"].includes(engine)) {
    runPythonPipeline(config, args);
    return;
  }

  const python = findPython();
  const result = spawnSync(python, ["-"], {
    input: PY_WORKER,
    encoding: "utf8",
    env: {
      ...process.env,
      PYTHONIOENCODING: "utf-8",
      PYTHONUTF8: "1",
      REPORT_PPTX_CONFIG: JSON.stringify(config),
    },
    maxBuffer: 1024 * 1024 * 80,
  });

  if (result.stdout) process.stdout.write(result.stdout);
  if (result.stderr) process.stderr.write(result.stderr);
  if (result.status !== 0) {
    throw new Error(`Генерация завершилась со статусом ${result.status}. Проверьте generation-log.json рядом с output.`);
  }
}

try {
  main();
} catch (error) {
  console.error(error.message || String(error));
  process.exit(1);
}
