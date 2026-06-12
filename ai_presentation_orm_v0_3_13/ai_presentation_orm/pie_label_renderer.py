from __future__ import annotations

import math
from pathlib import Path
from typing import Any

from .chart_style import CHART_COMPACT_PX, CHART_LABEL_PX


def _fmt_int(value: int | float) -> str:
    return f"{int(round(float(value))):,}".replace(",", " ")


def _fmt_pct(value: float, digits: int = 1) -> str:
    return f"{float(value) * 100:.{digits}f}%".replace(".", ",")


def _boxes_overlap(a: tuple[int, int, int, int], b: tuple[int, int, int, int], *, gap: int = 0) -> bool:
    return not (a[2] + gap <= b[0] or b[2] + gap <= a[0] or a[3] + gap <= b[1] or b[3] + gap <= a[1])


def _count_overlaps(boxes: list[tuple[int, int, int, int]]) -> int:
    count = 0
    for idx, box in enumerate(boxes):
        for other in boxes[idx + 1:]:
            if _boxes_overlap(box, other, gap=4):
                count += 1
    return count


def _count_out_of_bounds(boxes: list[tuple[int, int, int, int]], width: int, height: int) -> int:
    return sum(1 for x1, y1, x2, y2 in boxes if x1 < 0 or y1 < 0 or x2 > width or y2 > height)


def _box_intersects_circle(box: tuple[int, int, int, int], center: tuple[int, int], radius: int) -> bool:
    x1, y1, x2, y2 = box
    cx, cy = center
    closest_x = min(max(cx, x1), x2)
    closest_y = min(max(cy, y1), y2)
    return (closest_x - cx) ** 2 + (closest_y - cy) ** 2 < radius ** 2


def _count_chart_overlaps(boxes: list[tuple[int, int, int, int]], center: tuple[int, int], radius: int) -> int:
    return sum(1 for box in boxes if _box_intersects_circle(box, center, radius))


def _adjust_column(items: list[dict[str, Any]], *, min_y: int, max_y: int, gap: int) -> None:
    if not items:
        return
    items.sort(key=lambda item: item["natural_y"])
    for item in items:
        item["top"] = max(min_y, min(item["natural_y"] - item["height"] // 2, max_y - item["height"]))
    for idx in range(1, len(items)):
        prev = items[idx - 1]
        item = items[idx]
        item["top"] = max(item["top"], prev["top"] + prev["height"] + gap)
    overflow = items[-1]["top"] + items[-1]["height"] - max_y
    if overflow > 0:
        items[-1]["top"] -= overflow
        for idx in range(len(items) - 2, -1, -1):
            next_item = items[idx + 1]
            item = items[idx]
            item["top"] = min(item["top"], next_item["top"] - item["height"] - gap)
        underflow = min_y - items[0]["top"]
        if underflow > 0:
            for item in items:
                item["top"] += underflow


def render_pie_base(
    *,
    items: list[dict[str, Any]],
    out_path: Path,
    width: int,
    height: int,
    center: tuple[int, int],
    radius: int,
    inner_radius: int = 0,
    start_angle: float = -90.0,
) -> dict[str, Any]:
    from PIL import Image, ImageDraw

    if not items:
        items = [{"label": "нет данных", "value": 1, "share": 1, "color": "#D9D9D9"}]
    total = sum(max(float(item.get("value", 0) or 0), 0) for item in items) or 1
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    cx, cy = center
    bbox = [cx - radius, cy - radius, cx + radius, cy + radius]
    angle = start_angle
    sector_count = 0
    for item in items:
        value = max(float(item.get("value", 0) or 0), 0)
        extent = 360 * value / total
        color = str(item.get("color") or "#BFBFBF")
        if extent > 0:
            draw.pieslice(bbox, start=angle, end=angle + extent, fill=color)
            sector_count += 1
        angle += extent
    if inner_radius > 0:
        inner_radius = min(inner_radius, max(1, radius - 1))
        draw.ellipse([cx - inner_radius, cy - inner_radius, cx + inner_radius, cy + inner_radius], fill="white")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path)
    return {"path": str(out_path), "diagnostics": {"sector_count": sector_count}}


def render_labeled_pie(
    *,
    items: list[dict[str, Any]],
    out_path: Path,
    width: int,
    height: int,
    center: tuple[int, int],
    radius: int,
    inner_radius: int = 0,
    start_angle: float = -90.0,
    label_font_size: int = CHART_LABEL_PX,
    value_font_size: int = CHART_LABEL_PX,
    min_label_gap: int = 16,
    side_padding: int = 48,
    label_y_bounds: tuple[int, int] | None = None,
    connector_color: str = "#BFBFBF",
    _allow_font_fallback: bool = True,
) -> dict[str, Any]:
    from PIL import Image, ImageDraw, ImageFont

    if _allow_font_fallback and (label_font_size > CHART_COMPACT_PX or value_font_size > CHART_COMPACT_PX):
        result = render_labeled_pie(
            items=items,
            out_path=out_path,
            width=width,
            height=height,
            center=center,
            radius=radius,
            inner_radius=inner_radius,
            start_angle=start_angle,
            label_font_size=label_font_size,
            value_font_size=value_font_size,
            min_label_gap=min_label_gap,
            side_padding=side_padding,
            label_y_bounds=label_y_bounds,
            connector_color=connector_color,
            _allow_font_fallback=False,
        )
        diagnostics = result.get("diagnostics") or {}
        if (
            diagnostics.get("overlap_count", 0) == 0
            and diagnostics.get("out_of_bounds_count", 0) == 0
            and diagnostics.get("color_mismatch_count", 0) == 0
            and diagnostics.get("label_chart_overlap_count", 0) == 0
        ):
            return result
        return render_labeled_pie(
            items=items,
            out_path=out_path,
            width=width,
            height=height,
            center=center,
            radius=radius,
            inner_radius=inner_radius,
            start_angle=start_angle,
            label_font_size=CHART_COMPACT_PX,
            value_font_size=CHART_COMPACT_PX,
            min_label_gap=min_label_gap,
            side_padding=side_padding,
            label_y_bounds=label_y_bounds,
            connector_color=connector_color,
            _allow_font_fallback=False,
        )

    def font(size: int, *, bold: bool = False):
        try:
            return ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf", size=size)
        except OSError:
            return ImageFont.load_default()

    def text_size(draw: ImageDraw.ImageDraw, text: str, font_obj) -> tuple[int, int]:
        bbox = draw.textbbox((0, 0), text, font=font_obj)
        return int(bbox[2] - bbox[0]), int(bbox[3] - bbox[1])

    if not items:
        items = [{"label": "нет данных", "value": 1, "share": 1, "color": "#D9D9D9"}]
    total = sum(max(float(item.get("value", 0) or 0), 0) for item in items) or 1
    label_font = font(label_font_size)
    label_bold_font = font(label_font_size, bold=True)
    value_font = font(value_font_size)
    line_h = int(label_font_size * 1.05)
    marker = max(11, int(label_font_size * 0.34))
    marker_gap = max(8, int(label_font_size * 0.28))
    chart_padding = max(10, int(label_font_size * 0.22))

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    cx, cy = center

    angle = start_angle
    label_items: list[dict[str, Any]] = []
    segments: list[dict[str, Any]] = []
    for item in items:
        value = max(float(item.get("value", 0) or 0), 0)
        extent = 360 * value / total
        color = str(item.get("color") or "#BFBFBF")
        mid_angle = angle + extent / 2
        theta = math.radians(mid_angle)
        label = str(item.get("label") or "").strip()
        value_label = _fmt_int(value)
        share_label = _fmt_pct(float(item.get("share", value / total) or 0))
        lines = [(label, label_bold_font if bool(item.get("highlight")) else label_font), (value_label, value_font), (share_label, value_font)]
        widths = [text_size(draw, text, font_obj)[0] for text, font_obj in lines]
        max_w = max(widths or [0])
        height_box = line_h * len(lines)
        side = "right" if math.cos(theta) >= 0 else "left"
        natural_y = int(cy + radius * 1.14 * math.sin(theta))
        label_items.append({
            "source": item,
            "theta": theta,
            "side": side,
            "natural_y": natural_y,
            "height": height_box,
            "max_w": max_w,
            "lines": lines,
            "highlight": bool(item.get("highlight")),
            "color": color,
        })
        segments.append({"start": angle, "end": angle + extent, "color": color})
        angle += extent

    if not _allow_font_fallback and label_font_size <= CHART_COMPACT_PX and value_font_size <= CHART_COMPACT_PX:
        left_required = max(
            (side_padding + marker + marker_gap + int(item["max_w"]) for item in label_items if item["side"] == "left"),
            default=0,
        )
        right_required = max(
            (side_padding + marker + marker_gap + int(item["max_w"]) for item in label_items if item["side"] == "right"),
            default=0,
        )
        max_left_radius = cx - left_required - chart_padding if left_required else radius
        max_right_radius = width - right_required - chart_padding - cx if right_required else radius
        radius = max(1, min(radius, int(max_left_radius), int(max_right_radius)))

    bbox = [cx - radius, cy - radius, cx + radius, cy + radius]
    for segment in segments:
        draw.pieslice(bbox, start=segment["start"], end=segment["end"], fill=segment["color"])

    if inner_radius > 0:
        inner_radius = min(inner_radius, max(1, radius - 1))
        draw.ellipse([cx - inner_radius, cy - inner_radius, cx + inner_radius, cy + inner_radius], fill="white")

    min_y, max_y = label_y_bounds or (side_padding, height - side_padding)
    left_items = [item for item in label_items if item["side"] == "left"]
    right_items = [item for item in label_items if item["side"] == "right"]
    _adjust_column(left_items, min_y=min_y, max_y=max_y, gap=min_label_gap)
    _adjust_column(right_items, min_y=min_y, max_y=max_y, gap=min_label_gap)

    boxes: list[tuple[int, int, int, int]] = []
    text_boxes: list[tuple[int, int, int, int]] = []
    color_mismatch_count = 0
    for item in label_items:
        top = int(item["top"])
        if item["side"] == "left":
            marker_x = side_padding
            text_x = marker_x + marker + marker_gap
            text_anchor = "la"
            x1 = marker_x
            text_x1 = text_x
            x2 = text_x + item["max_w"]
            text_x2 = x2
        else:
            text_x = width - side_padding
            marker_x = int(text_x - item["max_w"] - marker_gap - marker)
            text_anchor = "ra"
            x1 = marker_x
            text_x1 = text_x - item["max_w"]
            x2 = text_x
            text_x2 = text_x
        marker_y = top + int((line_h - marker) / 2)
        marker_center = (marker_x + marker // 2, marker_y + marker // 2)
        sector_edge = (
            int(cx + radius * 1.02 * math.cos(item["theta"])),
            int(cy + radius * 1.02 * math.sin(item["theta"])),
        )
        draw.line([sector_edge, marker_center], fill=connector_color, width=2)
        draw.rectangle([marker_x, marker_y, marker_x + marker, marker_y + marker], fill=item["color"])
        for idx, (text, font_obj) in enumerate(item["lines"]):
            draw.text(
                (text_x, top + idx * line_h),
                text,
                fill="#000000",
                font=font_obj,
                anchor=text_anchor,
                stroke_width=0,
            )
        boxes.append((int(x1), int(top), int(x2), int(top + item["height"])))
        text_boxes.append((int(text_x1), int(top), int(text_x2), int(top + item["height"])))
        if str(item["source"].get("color") or "") != item["color"]:
            color_mismatch_count += 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path)
    diagnostics = {
        "label_count": len(label_items),
        "sector_count": len(items),
        "overlap_count": _count_overlaps(boxes),
        "out_of_bounds_count": _count_out_of_bounds(boxes, width, height),
        "label_chart_overlap_count": _count_chart_overlaps(text_boxes, center, radius + chart_padding),
        "color_mismatch_count": color_mismatch_count,
        "font_size_px": label_font_size,
        "chart_radius": radius,
        "label_boxes": [list(box) for box in boxes],
        "text_boxes": [list(box) for box in text_boxes],
    }
    return {"path": str(out_path), "diagnostics": diagnostics}
