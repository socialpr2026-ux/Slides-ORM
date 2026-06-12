from __future__ import annotations

from pathlib import Path
import json


def write_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def write_all_reports(
    output_dir: Path,
    config: dict,
    file_roles: dict,
    template_map: dict,
    data_inventory: dict,
    slide_readiness: dict,
    missing_data: dict,
) -> None:
    write_json(output_dir / "file_roles.json", file_roles)
    write_json(output_dir / "template_map.json", template_map)
    write_json(output_dir / "data_inventory.json", data_inventory)
    write_json(output_dir / "slide_readiness_report.json", slide_readiness)
    write_json(output_dir / "missing_data_report.json", missing_data)

    (output_dir / "file_inventory.md").write_text(_file_inventory_md(file_roles, template_map, data_inventory), encoding="utf-8")
    (output_dir / "slide_readiness_report.md").write_text(_slide_readiness_md(slide_readiness), encoding="utf-8")
    (output_dir / "missing_data_report.md").write_text(_missing_data_md(missing_data), encoding="utf-8")
    write_json(output_dir / "pipeline_summary.json", {
        "project": config.get("project", {}),
        "template": template_map.get("name"),
        "slide_count": template_map.get("slide_count"),
        "input_file_count": len(file_roles.get("files", [])),
        "available_entities": slide_readiness.get("available_entities", []),
        "readiness_status_counts": slide_readiness.get("status_counts", {}),
        "missing_data_count": len(missing_data.get("items", [])),
    })


def _file_inventory_md(file_roles: dict, template_map: dict, data_inventory: dict) -> str:
    lines = [
        "# File inventory",
        "",
        "## Template PPTX",
        "",
        f"- File: `{template_map.get('name')}`",
        f"- Slides: {template_map.get('slide_count')}",
        f"- Size: {template_map.get('slide_width_cm')} × {template_map.get('slide_height_cm')} cm",
        f"- Charts: {template_map.get('summary', {}).get('total_charts')}",
        f"- Tables: {template_map.get('summary', {}).get('total_tables')}",
        "",
        "## Input file roles",
        "",
        "| File | Role guess | Notes |",
        "|---|---|---|",
    ]
    for f in file_roles.get("files", []):
        lines.append(f"| `{f.get('name')}` | {f.get('role_guess')} | {f.get('sample_preview','')[:120]} |")

    lines += ["", "## Data inventory", ""]
    for f in data_inventory.get("files", []):
        lines += [f"### `{f.get('name')}`", ""]
        lines.append(f"- Type: {f.get('type')}")
        lines.append(f"- Detected entities: {', '.join(f.get('detected_entities', [])) or 'none'}")
        if f.get("sheets"):
            lines.append("")
            lines.append("| Sheet | Range | Entities |")
            lines.append("|---|---|---|")
            for sh in f.get("sheets", []):
                lines.append(f"| {sh.get('name')} | {sh.get('range')} | {', '.join(sh.get('detected_entities', []))} |")
        lines.append("")
    return "\n".join(lines)


def _slide_readiness_md(slide_readiness: dict) -> str:
    lines = [
        "# Slide readiness report",
        "",
        f"Available entities: {', '.join(slide_readiness.get('available_entities', []))}",
        "",
        "| Slide | Role | Status | Missing entities | Note |",
        "|---:|---|---|---|---|",
    ]
    for slide in slide_readiness.get("slides", []):
        lines.append(
            f"| {slide.get('slide_number')} | {slide.get('role')} | {slide.get('status')} | "
            f"{', '.join(slide.get('missing_entities', []))} | {slide.get('note')} |"
        )
    return "\n".join(lines)


def _missing_data_md(missing_data: dict) -> str:
    lines = [
        "# Missing data report",
        "",
        "| Slide | Role | Block | Missing | Where to get | Impact |",
        "|---:|---|---|---|---|---|",
    ]
    for item in missing_data.get("items", []):
        lines.append(
            f"| {item.get('slide_number')} | {item.get('role')} | {item.get('block')} | "
            f"{item.get('missing')} | {item.get('where_to_get')} | {item.get('impact_if_missing')} |"
        )
    return "\n".join(lines)
