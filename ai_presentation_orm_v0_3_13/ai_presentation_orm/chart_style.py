from __future__ import annotations

from .style_rules import CHART_FONT_SIZE, FONT_FAMILY

CHART_FONT = FONT_FAMILY
CHART_LABEL_SIZE = CHART_FONT_SIZE
CHART_COMPACT_SIZE = CHART_FONT_SIZE


def chart_pt_to_px(points: float, *, dpi: int = 300) -> int:
    return max(1, int(round(float(points) * dpi / 72)))


CHART_LABEL_PX = chart_pt_to_px(CHART_LABEL_SIZE)
CHART_COMPACT_PX = chart_pt_to_px(CHART_COMPACT_SIZE)
