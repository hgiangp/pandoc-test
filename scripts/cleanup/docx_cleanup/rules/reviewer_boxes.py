"""ReviewerBoxes: red (rounded) rectangles that a reader drew around content to highlight it.

They carry no information, but they make the shape rasterizer render empty red boxes or
merge them into (and split) real figures. A shape is a reviewer box when ALL hold:

  geometry   rectangle / rounded rectangle (param `geometries`)
  fill       none, or nearly transparent (param `maxFillOpacity`)
  color      line color close to one of `colors`, or a clear red hue (`matchRedHue`)
  text       no text inside
  connected  no connector attached (a real diagram box usually has arrows attached)

Confidence: high   = all hold, stand-alone shape                     -> removed in Apply mode
            medium = all hold, but inside a group / drawing canvas   -> removed only if includeInGroups
            low    = reddish line and exactly one condition fails    -> reported only (for tuning)
"""
from __future__ import annotations

from typing import Any, Optional

from lxml import etree

from ..colors import (distance, is_red_hue, is_reddish, parse_hex, resolve_dml_color,
                      resolve_vml_color, to_hex)
from ..model import Finding, Rule
from ..package import NS, DocxPackage, ShapeRef, local_name, q

EMU_PER_PT = 12700


class ReviewerBoxes(Rule):
    id = "ReviewerBoxes"
    description = "Red highlight rectangles drawn by reviewers"
    defaults: dict[str, Any] = {
        "colors": ["FF0000", "C00000", "E60012"],
        "colorTolerance": 60,
        "matchRedHue": True,
        "geometries": ["rect", "roundRect"],
        "maxFillOpacity": 0.1,
        "includeInGroups": False,
    }

    def __init__(self, params: Optional[dict[str, Any]] = None):
        super().__init__(params)
        self._targets = [c for c in (parse_hex(x) for x in self.params["colors"]) if c]
        if len(self._targets) != len(self.params["colors"]):
            raise ValueError(f"{self.id}: invalid color in {self.params['colors']}")

    # ------------------------------------------------------------ detect

    def detect(self, pkg: DocxPackage) -> list[Finding]:
        theme = pkg.theme_colors()
        connected: dict[etree._Element, set[str]] = {}
        findings = []
        for ref in pkg.iter_shapes():
            if ref.kind == "dml":
                feats = self._dml_features(ref.element, theme)
            else:
                feats = self._vml_features(ref.element)
            if feats is None:
                continue
            # connectors only exist inside a canvas/group (Word cannot connect floating shapes)
            container = pkg.container_element(ref)
            if container is not None and container not in connected:
                connected[container] = pkg.connected_ids(container)
            feats["connected"] = (container is not None and bool(ref.shape_id)
                                  and ref.shape_id in connected[container])
            feats["container"] = ref.container

            finding = self._classify(ref, feats)
            if finding is not None:
                finding.location = pkg.paragraph_text(ref.part, ref.paragraph_index)[:80]
                findings.append(finding)
        return findings

    def _classify(self, ref: ShapeRef, f: dict[str, Any]) -> Optional[Finding]:
        rgb = parse_hex(f["lineColor"]) if f.get("lineColor") else None
        if rgb is None or not f["hasLine"]:
            return None
        matched = any(distance(rgb, t) <= self.params["colorTolerance"] for t in self._targets) \
            or (self.params["matchRedHue"] and is_red_hue(rgb))
        if not matched and not is_reddish(rgb):
            return None

        checks = {
            "geometry": f["geometry"] in self.params["geometries"],
            "fill": f["fillOpacity"] <= self.params["maxFillOpacity"],
            "color": matched,
            "text": not f["hasText"],
            "connected": not f["connected"],
        }
        failed = [k for k, ok in checks.items() if not ok]
        if not failed:
            confidence = "high" if ref.container == "top" else "medium"
            reason = "all conditions met" + ("" if ref.container == "top" else f" (inside {ref.container})")
        elif len(failed) == 1:
            confidence = "low"
            reason = f"fails: {failed[0]}"
        else:
            return None

        return Finding(rule_id=self.id, part=ref.part, paragraph_index=ref.paragraph_index,
                       shape_id=ref.shape_id, shape_name=ref.shape_name, container=ref.container,
                       confidence=confidence, reason=reason, features=f, ref=ref)

    # ------------------------------------------------------------ apply

    def apply(self, pkg: DocxPackage, findings: list[Finding]) -> None:
        for fd in findings:
            remove = fd.confidence == "high" or (fd.confidence == "medium" and self.params["includeInGroups"])
            if remove and fd.ref is not None:
                pkg.remove_shape(fd.ref)
                fd.action = "removed"
            elif fd.confidence == "medium":
                fd.action = "reported (set includeInGroups to remove)"

    # ------------------------------------------------------------ features

    @staticmethod
    def _dml_features(wsp: etree._Element, theme: dict[str, str]) -> Optional[dict[str, Any]]:
        sp_pr = wsp.find("wps:spPr", NS)
        if sp_pr is None:
            return None
        style = wsp.find("wps:style", NS)

        prst = sp_pr.find("a:prstGeom", NS)
        geometry = prst.get("prst", "") if prst is not None else (
            "custom" if sp_pr.find("a:custGeom", NS) is not None else "")

        # fill: explicit spPr fill wins, otherwise the style's fillRef (idx 0 = no fill)
        fill_opacity = 0.0
        fill_kind = "none"
        explicit = [c for c in sp_pr if local_name(c) in
                    ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")]
        if explicit:
            kind = local_name(explicit[0])
            if kind == "solidFill":
                fill_kind = "solid"
                fill_opacity = _dml_alpha(explicit[0])
            elif kind != "noFill":
                fill_kind = kind
                fill_opacity = 1.0
        elif style is not None:
            fill_ref = style.find("a:fillRef", NS)
            if fill_ref is not None and fill_ref.get("idx", "0") != "0":
                fill_kind = "style"
                fill_opacity = _dml_alpha(fill_ref)

        # line: explicit a:ln wins, otherwise the style's lnRef (idx 0 = no line)
        ln = sp_pr.find("a:ln", NS)
        has_line, color, width_pt, dash = False, None, 0.75, "solid"
        if ln is not None and ln.find("a:noFill", NS) is not None:
            has_line = False
        else:
            ln_fill = ln.find("a:solidFill", NS) if ln is not None else None
            if ln_fill is not None:
                has_line, color = True, resolve_dml_color(ln_fill, theme)
            elif style is not None:
                ln_ref = style.find("a:lnRef", NS)
                if ln_ref is not None and ln_ref.get("idx", "0") != "0":
                    has_line, color = True, resolve_dml_color(ln_ref, theme)
            if ln is not None:
                if ln.get("w"):
                    width_pt = round(int(ln.get("w")) / EMU_PER_PT, 2)
                d = ln.find("a:prstDash", NS)
                if d is not None:
                    dash = d.get("val", "solid")

        text = "".join(t.text or "" for t in wsp.iter(q("w:t"))).strip()
        return {
            "kind": "dml",
            "geometry": geometry,
            "fill": fill_kind,
            "fillOpacity": round(fill_opacity, 2),
            "hasLine": has_line,
            "lineColor": to_hex(color) if color else "",
            "lineWidthPt": width_pt,
            "dash": dash,
            "hasText": bool(text),
        }

    @staticmethod
    def _vml_features(el: etree._Element) -> Optional[dict[str, Any]]:
        name = local_name(el)
        geometry = {"rect": "rect", "roundrect": "roundRect", "oval": "ellipse"}.get(name)
        if geometry is None:
            return None

        fill_el = el.find("v:fill", NS)
        filled = el.get("filled", "t").lower() not in ("f", "false")
        opacity = 1.0
        if fill_el is not None:
            if fill_el.get("on", "t").lower() in ("f", "false"):
                filled = False
            opacity = _vml_fraction(fill_el.get("opacity"), 1.0)
        fill_opacity = opacity if filled else 0.0

        stroke_el = el.find("v:stroke", NS)
        has_line = el.get("stroked", "t").lower() not in ("f", "false")
        color_val = el.get("strokecolor", "black")
        dash = "solid"
        if stroke_el is not None:
            if stroke_el.get("on", "t").lower() in ("f", "false"):
                has_line = False
            color_val = stroke_el.get("color", color_val)
            dash = stroke_el.get("dashstyle", "solid")
        color = resolve_vml_color(color_val)

        text = "".join(t.text or "" for t in el.iter(q("w:t"))).strip()
        return {
            "kind": "vml",
            "geometry": geometry,
            "fill": "solid" if filled else "none",
            "fillOpacity": round(fill_opacity, 2),
            "hasLine": has_line,
            "lineColor": to_hex(color) if color else "",
            "lineWidthPt": _vml_points(el.get("strokeweight"), 0.75),
            "dash": dash,
            "hasText": bool(text),
        }


def _dml_alpha(parent: etree._Element) -> float:
    """Opacity (0..1) of the color inside a fill element; 1.0 when no a:alpha is set."""
    alpha = parent.find(".//a:alpha", NS)
    if alpha is None:
        return 1.0
    try:
        return int(alpha.get("val", "100000")) / 100000.0
    except ValueError:
        return 1.0


def _vml_fraction(value: Optional[str], default: float) -> float:
    """VML opacity: '0.5', '50%' or '32768f' (fixed point 1/65536)."""
    if not value:
        return default
    v = value.strip().lower()
    try:
        if v.endswith("f"):
            return int(v[:-1]) / 65536.0
        if v.endswith("%"):
            return float(v[:-1]) / 100.0
        return float(v)
    except ValueError:
        return default


def _vml_points(value: Optional[str], default: float) -> float:
    if not value:
        return default
    v = value.strip().lower()
    try:
        if v.endswith("pt"):
            return float(v[:-2])
        if v.endswith("px"):
            return round(float(v[:-2]) * 0.75, 2)
        return round(int(v) / EMU_PER_PT, 2)  # bare number = EMU
    except ValueError:
        return default
