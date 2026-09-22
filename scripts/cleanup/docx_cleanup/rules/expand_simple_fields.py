"""ExpandSimpleFields: rewrite <w:fldSimple> as a complex field (begin/instr/separate/result/end).

Word stores a field either as a complex field (w:fldChar runs) or as a compact w:fldSimple.
pandoc (checked up to 3.11) drops the result of every w:fldSimple, and for SEQ also the
caption label in front of it:

    Table {SEQ 1-3} Example      ->  "Example"
    Fig. {STYLEREF}-{SEQ} Ex.    ->  "Fig. - Ex."
    shown in {REF _Ref1}         ->  "shown in "            (cross reference lost)

The complex form carries exactly the same instruction and result, and pandoc reads it
correctly (REF becomes a link). This is a pandoc compatibility fix, not a data cleanup: it
belongs to the "pandoc" profile that runs right before pandoc, for every document.
"""
from __future__ import annotations

import copy
import re
from typing import Any

from lxml import etree

from ..model import Finding, Rule
from ..package import NS, DocxPackage, has_ancestor, paragraph_plain_text, q

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


class ExpandSimpleFields(Rule):
    id = "ExpandSimpleFields"
    description = "w:fldSimple -> complex field, so pandoc keeps field results"
    defaults: dict[str, Any] = {}

    def detect(self, pkg: DocxPackage) -> list[Finding]:
        findings = []
        for part in pkg.story_parts:
            root = pkg.xml(part)
            paragraphs = {p: i for i, p in enumerate(root.iter(q("w:p")))}
            for fld in root.iter(q("w:fldSimple")):
                if has_ancestor(fld, "mc:Fallback"):
                    continue          # pandoc reads mc:Choice; Word regenerates the fallback
                instr = fld.get(q("w:instr"), "")
                keyword = (instr.split() or ["?"])[0].upper()
                para = next(fld.iterancestors(q("w:p")), None)
                findings.append(Finding(
                    rule_id=self.id, part=part,
                    paragraph_index=paragraphs.get(para, -1) if para is not None else -1,
                    shape_id="", shape_name="", container="field", confidence="high",
                    reason=f"fldSimple {keyword}",
                    features={"instr": instr.strip(), "result": _text(fld)},
                    location=paragraph_plain_text(para)[:80] if para is not None else "",
                    target=fld))
        return findings

    def apply(self, pkg: DocxPackage, findings: list[Finding]) -> None:
        for fd in findings:
            _expand(fd.target)
            pkg.mark_dirty(fd.part)
            fd.action = "expanded"


def _text(el: etree._Element) -> str:
    return re.sub(r"\s+", " ", "".join(t.text or "" for t in el.iter(q("w:t")))).strip()


def _run(rpr: etree._Element | None, child: etree._Element) -> etree._Element:
    r = etree.Element(q("w:r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    r.append(child)
    return r


def _expand(fld: etree._Element) -> None:
    parent = fld.getparent()
    if parent is None:
        return
    pos = parent.index(fld)
    first_run = fld.find("w:r", NS)
    rpr = first_run.find("w:rPr", NS) if first_run is not None else None

    begin = etree.Element(q("w:fldChar"))
    begin.set(q("w:fldCharType"), "begin")
    for attr in ("fldLock", "dirty"):
        if fld.get(q(f"w:{attr}")) is not None:
            begin.set(q(f"w:{attr}"), fld.get(q(f"w:{attr}")))
    instr = etree.Element(q("w:instrText"))
    instr.set(XML_SPACE, "preserve")
    instr.text = fld.get(q("w:instr"), "")
    separate = etree.Element(q("w:fldChar"))
    separate.set(q("w:fldCharType"), "separate")
    end = etree.Element(q("w:fldChar"))
    end.set(q("w:fldCharType"), "end")

    result = [c for c in fld if c.tag != q("w:fldData")]
    new = [_run(rpr, begin), _run(rpr, instr), _run(rpr, separate), *result, _run(rpr, end)]
    parent.remove(fld)
    for offset, el in enumerate(new):
        parent.insert(pos + offset, el)
