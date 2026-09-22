"""UnwrapLayoutTables: single-column tables used only to lay out a figure/table with its caption.

NS documents put a figure (or a data table) together with its caption inside a one-cell
table. pandoc then renders the whole block as a grid table, so images and tables appear
inside a box and the caption is not next to the figure. Unwrapping moves the cell content
out, in order, and drops the wrapper.

A table is a layout wrapper when:

  columns    every row has exactly one cell
  rows       at most `maxRows` rows
  content    it holds a drawing / picture / OLE object or a nested table
  caption    one of its own paragraphs is a caption (Caption style or SEQ field)

Confidence: high   = all hold                              -> unwrapped in Apply mode
            medium = content but no caption                -> unwrapped only if unwrapWithoutCaption
            low    = content but more rows than maxRows,
                     or text only (e.g. a "Note" box)      -> reported only
Tables with more than one column (e.g. figures side by side) are never touched.
"""
from __future__ import annotations

from typing import Any

from ..model import Finding, Rule
from ..package import NS, DocxPackage, has_ancestor, paragraph_plain_text, q

CONTENT_TAGS = ("w:drawing", "w:pict", "w:object")


class UnwrapLayoutTables(Rule):
    id = "UnwrapLayoutTables"
    description = "Single-column tables wrapping a figure/table and its caption"
    defaults: dict[str, Any] = {
        "maxRows": 4,
        "unwrapWithoutCaption": False,
    }

    def detect(self, pkg: DocxPackage) -> list[Finding]:
        findings = []
        for part, tbl, para_idx in pkg.iter_tables():
            rows = tbl.findall("w:tr", NS)
            cells = [row.findall("w:tc", NS) for row in rows]
            if not rows or any(len(c) != 1 for c in cells):
                continue                                   # multi-column: never a wrapper
            own_blocks = [b for row in cells for b in row[0] if b.tag != q("w:tcPr")]
            own_paras = [b for b in own_blocks if b.tag == q("w:p")]
            nested = [b for b in own_blocks if b.tag == q("w:tbl")]
            drawings = sum(1 for p in own_paras for tag in CONTENT_TAGS for el in p.iter(q(tag))
                           if not has_ancestor(el, "mc:Fallback"))
            captions = [p for p in own_paras if pkg.is_caption(p)]

            feats = {
                "rows": len(rows),
                "captions": len(captions),
                "drawings": drawings,
                "nestedTables": len(nested),
                "nestedInTable": bool(tbl.xpath("ancestor::w:tbl", namespaces=NS)),
            }
            has_content = drawings > 0 or len(nested) > 0
            too_many_rows = len(rows) > self.params["maxRows"]

            if has_content and captions and not too_many_rows:
                confidence, reason = "high", "caption + figure/table"
            elif has_content and not too_many_rows:
                confidence, reason = "medium", "figure/table without caption"
            elif has_content:
                confidence, reason = "low", f"fails: rows ({len(rows)} > maxRows)"
            elif not too_many_rows and any(paragraph_plain_text(p) for p in own_paras):
                confidence, reason = "low", "text only (box / note?)"
            else:
                continue

            label = captions[0] if captions else next((p for p in own_paras if paragraph_plain_text(p)), None)
            findings.append(Finding(
                rule_id=self.id, part=part, paragraph_index=para_idx, shape_id="",
                shape_name=f"table {len(rows)}x1", container="table", confidence=confidence,
                reason=reason, features=feats, target=tbl,
                location=paragraph_plain_text(label)[:80] if label is not None else ""))
        return findings

    def apply(self, pkg: DocxPackage, findings: list[Finding]) -> None:
        for fd in findings:
            unwrap = fd.confidence == "high" or (
                fd.confidence == "medium" and self.params["unwrapWithoutCaption"])
            if unwrap:
                pkg.unwrap_table(fd.part, fd.target)
                fd.action = "unwrapped"
            elif fd.confidence == "medium":
                fd.action = "reported (set unwrapWithoutCaption to unwrap)"
