"""Resolve DrawingML / VML colors to RGB and compare them."""
from __future__ import annotations

import colorsys
import re
from typing import Optional

from lxml import etree

from .package import NS, local_name

RGB = tuple[int, int, int]

# DrawingML preset colors (a:prstClr) and VML color names likely to be used for markup
NAMED = {
    "red": "FF0000", "darkred": "8B0000", "crimson": "DC143C", "firebrick": "B22222",
    "orangered": "FF4500", "tomato": "FF6347", "maroon": "800000", "magenta": "FF00FF",
    "fuchsia": "FF00FF", "pink": "FFC0CB", "hotpink": "FF69B4", "orange": "FFA500",
    "yellow": "FFFF00", "blue": "0000FF", "green": "008000", "lime": "00FF00",
    "black": "000000", "white": "FFFFFF", "gray": "808080", "grey": "808080",
}


def parse_hex(value: str) -> Optional[RGB]:
    value = value.strip().lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    if not re.fullmatch(r"[0-9a-fA-F]{6}", value):
        return None
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def to_hex(rgb: RGB) -> str:
    return "%02X%02X%02X" % rgb


def _clamp(x: float) -> float:
    return max(0.0, min(1.0, x))


def _apply_modifiers(rgb: RGB, color_el: etree._Element) -> RGB:
    """Approximate lumMod/lumOff/shade/tint, the modifiers Word uses for theme color variants."""
    r, g, b = (c / 255.0 for c in rgb)
    for mod in color_el:
        name = local_name(mod)
        try:
            val = int(mod.get("val", "0")) / 100000.0
        except ValueError:
            continue
        if name in ("lumMod", "lumOff"):
            h, l, s = colorsys.rgb_to_hls(r, g, b)
            l = _clamp(l * val) if name == "lumMod" else _clamp(l + val)
            r, g, b = colorsys.hls_to_rgb(h, l, s)
        elif name == "shade":
            r, g, b = r * val, g * val, b * val
        elif name == "tint":
            r, g, b = (1 - (1 - c) * val for c in (r, g, b))
    return tuple(int(round(_clamp(c) * 255)) for c in (r, g, b))  # type: ignore[return-value]


def resolve_dml_color(parent: Optional[etree._Element], theme: dict[str, str]) -> Optional[RGB]:
    """Color of the first color child of parent (a:solidFill, a:lnRef, ...)."""
    if parent is None:
        return None
    for el in parent:
        name = local_name(el)
        base: Optional[str] = None
        if name == "srgbClr":
            base = el.get("val")
        elif name == "schemeClr":
            base = theme.get(el.get("val", ""))
        elif name == "prstClr":
            base = NAMED.get(el.get("val", "").lower())
        elif name == "sysClr":
            base = el.get("lastClr")
        elif name == "scrgbClr":
            try:
                rgb = tuple(int(round(int(el.get(k, "0")) / 100000.0 * 255)) for k in ("r", "g", "b"))
                return _apply_modifiers(rgb, el)  # type: ignore[arg-type]
            except ValueError:
                return None
        else:
            continue
        rgb = parse_hex(base) if base else None
        return _apply_modifiers(rgb, el) if rgb else None
    return None


def resolve_vml_color(value: Optional[str]) -> Optional[RGB]:
    """VML colors: '#C00000', 'red', '#c00000 [3204]', 'window' ..."""
    if not value:
        return None
    token = value.strip().split()[0].lower()
    if token.startswith("#"):
        return parse_hex(token)
    if token in NAMED:
        return parse_hex(NAMED[token])
    return None


def distance(a: RGB, b: RGB) -> float:
    return sum((x - y) ** 2 for x, y in zip(a, b)) ** 0.5


def is_red_hue(rgb: RGB, max_hue_deg: float = 20.0, min_sat: float = 0.6, min_val: float = 0.45) -> bool:
    h, s, v = colorsys.rgb_to_hsv(*(c / 255.0 for c in rgb))
    hue = h * 360
    return (hue <= max_hue_deg or hue >= 360 - max_hue_deg) and s >= min_sat and v >= min_val


def is_reddish(rgb: RGB) -> bool:
    """Wider net (red, orange-red, pink, magenta-ish) used only for low-confidence reporting."""
    h, s, v = colorsys.rgb_to_hsv(*(c / 255.0 for c in rgb))
    hue = h * 360
    return (hue <= 40 or hue >= 300) and s >= 0.35 and v >= 0.35
