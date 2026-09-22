"""Read, inspect and safely modify the XML parts of a .docx package."""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Optional

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Parts of the document that can contain drawings (the "stories")
STORY_PART = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$")

# VML elements that are shapes (v:shapetype is only a template)
VML_SHAPES = {"shape", "rect", "roundrect", "oval", "line", "polyline", "arc", "curve", "group", "image"}


def q(tag: str) -> str:
    """'w:p' -> '{namespace}p'"""
    prefix, local = tag.split(":")
    return "{%s}%s" % (NS[prefix], local)


def local_name(el: etree._Element) -> str:
    return etree.QName(el).localname


def has_ancestor(el: etree._Element, tag: str) -> bool:
    return el.xpath("boolean(ancestor::%s)" % tag, namespaces=NS)


@dataclass
class ShapeRef:
    """A single shape found in a story part.

    kind:      'dml' (wps:wsp) or 'vml' (v:rect, v:roundrect, ...)
    container: 'top' (the shape is the whole drawing), 'group' or 'canvas'
    """

    part: str
    element: etree._Element
    kind: str
    container: str
    paragraph_index: int
    shape_id: str
    shape_name: str


class DocxPackage:
    def __init__(self, path: Path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as zf:
            self._infos = zf.infolist()
            self._data = {i.filename: zf.read(i.filename) for i in self._infos}
        self._trees: dict[str, etree._Element] = {}
        self._dirty: set[str] = set()
        self._theme: Optional[dict[str, str]] = None
        self._caption_styles: Optional[set[str]] = None

    # ------------------------------------------------------------ parts

    @property
    def story_parts(self) -> list[str]:
        return [n for n in self._data if STORY_PART.match(n)]

    def xml(self, part: str) -> etree._Element:
        if part not in self._trees:
            self._trees[part] = etree.fromstring(self._data[part])
        return self._trees[part]

    def mark_dirty(self, part: str) -> None:
        self._dirty.add(part)

    @property
    def dirty_parts(self) -> set[str]:
        return set(self._dirty)

    def save(self, out: Path) -> None:
        """Write the package; untouched parts are copied byte for byte."""
        out = Path(out)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
            for info in self._infos:
                name = info.filename
                if name in self._dirty:
                    data = etree.tostring(self._trees[name], xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
                else:
                    data = self._data[name]
                zf.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
        validate_package(out, self._dirty)

    # ------------------------------------------------------------ theme

    def theme_colors(self) -> dict[str, str]:
        """Scheme color name -> RRGGBB from word/theme/theme1.xml (aliases tx1/bg1/... included)."""
        if self._theme is not None:
            return self._theme
        colors: dict[str, str] = {}
        name = next((n for n in self._data if re.match(r"^word/theme/theme\d*\.xml$", n)), None)
        if name:
            root = self.xml(name)
            scheme = root.find(".//a:clrScheme", NS)
            if scheme is not None:
                for child in scheme:
                    srgb = child.find("a:srgbClr", NS)
                    sys_ = child.find("a:sysClr", NS)
                    if srgb is not None:
                        colors[local_name(child)] = srgb.get("val", "").upper()
                    elif sys_ is not None:
                        colors[local_name(child)] = sys_.get("lastClr", "").upper()
        for alias, target in (("tx1", "dk1"), ("bg1", "lt1"), ("tx2", "dk2"), ("bg2", "lt2")):
            if target in colors:
                colors[alias] = colors[target]
        self._theme = colors
        return colors

    # ------------------------------------------------------------ styles

    def caption_style_ids(self) -> set[str]:
        """Ids of the built-in Caption style and of every style based on it.

        Matched by style *name*: localized Word writes ids like 'a3', but the name of a
        built-in style is always the English 'caption'.
        """
        if self._caption_styles is not None:
            return self._caption_styles
        ids: set[str] = set()
        if "word/styles.xml" in self._data:
            based_on: dict[str, str] = {}
            for st in self.xml("word/styles.xml").iter(q("w:style")):
                sid = st.get(q("w:styleId"), "")
                name = st.find("w:name", NS)
                if name is not None and name.get(q("w:val"), "").lower() == "caption":
                    ids.add(sid)
                base = st.find("w:basedOn", NS)
                if base is not None:
                    based_on[sid] = base.get(q("w:val"), "")
            changed = True
            while changed:
                changed = False
                for sid, base in based_on.items():
                    if base in ids and sid not in ids:
                        ids.add(sid)
                        changed = True
        self._caption_styles = ids
        return ids

    def is_caption(self, p: etree._Element) -> bool:
        """Caption style (or derived), or a paragraph holding a SEQ field (Insert Caption)."""
        style = p.find("w:pPr/w:pStyle", NS)
        if style is not None and style.get(q("w:val"), "") in self.caption_style_ids():
            return True
        for instr in p.iter(q("w:instrText")):
            if re.search(r"\bSEQ\b", instr.text or ""):
                return True
        for fld in p.iter(q("w:fldSimple")):
            if re.search(r"\bSEQ\b", fld.get(q("w:instr"), "")):
                return True
        return False

    # ------------------------------------------------------------ discovery

    def iter_tables(self) -> Iterator[tuple[str, etree._Element, int]]:
        """(part, w:tbl, paragraph index of its first paragraph), outer tables first."""
        for part in self.story_parts:
            root = self.xml(part)
            para_index = _paragraph_index(root)
            for tbl in root.iter(q("w:tbl")):
                if has_ancestor(tbl, "w:txbxContent") or has_ancestor(tbl, "mc:Fallback"):
                    continue
                first = tbl.find(".//w:p", NS)
                yield part, tbl, para_index.get(first, -1) if first is not None else -1


    def iter_shapes(self) -> Iterator[ShapeRef]:
        """All DrawingML (wps:wsp) and VML shapes, ignoring mc:Fallback copies."""
        for part in self.story_parts:
            root = self.xml(part)
            para_index = _paragraph_index(root)

            for wsp in root.iter(q("wps:wsp")):
                if has_ancestor(wsp, "mc:Fallback"):
                    continue
                if has_ancestor(wsp, "wpc:wpc"):
                    container = "canvas"
                elif has_ancestor(wsp, "wpg:wgp"):
                    container = "group"
                else:
                    container = "top"
                sid, sname = _dml_id(wsp, container)
                yield ShapeRef(part, wsp, "dml", container, _para_of(wsp, para_index), sid, sname)

            for el in root.iter("{%s}*" % NS["v"]):
                if local_name(el) not in VML_SHAPES or local_name(el) == "group":
                    continue
                if has_ancestor(el, "mc:Fallback"):
                    continue
                container = "group" if has_ancestor(el, "v:group") else "top"
                yield ShapeRef(part, el, "vml", container, _para_of(el, para_index),
                               el.get("id", ""), el.get(q("o:spid"), "") or el.get("id", ""))

    @staticmethod
    def container_element(ref: ShapeRef) -> Optional[etree._Element]:
        """The drawing canvas, or outermost group, that holds a grouped shape."""
        if ref.container == "canvas":
            return _ancestor(ref.element, "wpc:wpc")
        if ref.container == "group":
            tag = "wpg:wgp" if ref.kind == "dml" else "v:group"
            found = ref.element.xpath("ancestor::%s" % tag, namespaces=NS)
            return found[0] if found else None
        return None

    @staticmethod
    def connected_ids(container: etree._Element) -> set[str]:
        """Ids of the shapes a connector is attached to, within one canvas/group.

        Connector ids are only unique inside their container, so never compare across them.
        """
        ids = set()
        for tag in ("a:stCxn", "a:endCxn"):
            for el in container.iter(q(tag)):
                if el.get("id"):
                    ids.add(el.get("id"))
        return ids

    def paragraph_text(self, part: str, paragraph_index: int, lookahead: int = 30) -> str:
        """Text of the paragraph, or of the next non-empty paragraph (often the caption)."""
        paras = _body_paragraphs(self.xml(part))
        for p in paras[paragraph_index: paragraph_index + lookahead]:
            text = paragraph_plain_text(p)
            if text:
                return text
        return ""

    # ------------------------------------------------------------ removal

    def remove_shape(self, ref: ShapeRef) -> None:
        """Remove a shape without leaving an invalid package behind.

        top-level DrawingML: the whole mc:AlternateContent (Choice + Fallback) or w:drawing
        shape in a group/canvas: the child only; the mc:Fallback copy of the container is
        dropped because it can no longer be kept in sync (Word regenerates it on save)
        top-level VML: the whole w:pict
        """
        el = ref.element
        if ref.kind == "dml":
            if ref.container == "top":
                drawing = _ancestor(el, "w:drawing")
                alt = _ancestor(drawing, "mc:AlternateContent") if drawing is not None else None
                _remove_and_prune(alt if alt is not None else drawing)
            else:
                container_tag = "wpc:wpc" if ref.container == "canvas" else "wpg:wgp"
                container = _ancestor(el, container_tag)
                el.getparent().remove(el)
                _drop_fallback(container)
                if container is not None and not _has_shape_children(container):
                    self.remove_drawing_of(container)
        else:
            if ref.container == "top":
                pict = _ancestor(el, "w:pict")
                if pict is not None and len([c for c in pict if local_name(c) in VML_SHAPES]) <= 1:
                    alt = _ancestor(pict, "mc:AlternateContent")
                    _remove_and_prune(alt if alt is not None else pict)
                else:
                    el.getparent().remove(el)
            else:
                el.getparent().remove(el)
        self.mark_dirty(ref.part)

    def unwrap_table(self, part: str, tbl: etree._Element) -> None:
        """Replace a single-column table by the content of its cells, in order.

        Paragraphs, nested tables and bookmarks move out as they are; only the table, row
        and cell wrappers (and their properties) disappear. A cell always ends with a
        paragraph, so the moved content never leaves two tables directly adjacent.
        """
        parent = tbl.getparent()
        pos = parent.index(tbl)
        moved: list[etree._Element] = []
        for child in tbl:
            if child.tag in (q("w:tblPr"), q("w:tblGrid")):
                continue
            if child.tag != q("w:tr"):
                moved.append(child)          # e.g. bookmarkStart/End between rows
                continue
            for row_child in child:
                if row_child.tag in (q("w:trPr"), q("w:tblPrEx")):
                    continue
                if row_child.tag != q("w:tc"):
                    moved.append(row_child)
                    continue
                moved.extend(c for c in row_child if c.tag != q("w:tcPr"))
        parent.remove(tbl)
        for offset, el in enumerate(moved):
            parent.insert(pos + offset, el)
        self.mark_dirty(part)

    def remove_drawing_of(self, el: etree._Element) -> None:
        drawing = _ancestor(el, "w:drawing")
        if drawing is None:
            return
        alt = _ancestor(drawing, "mc:AlternateContent")
        _remove_and_prune(alt if alt is not None else drawing)


# ---------------------------------------------------------------- helpers

def _ancestor(el: Optional[etree._Element], tag: str) -> Optional[etree._Element]:
    if el is None:
        return None
    found = el.xpath("ancestor::%s[1]" % tag, namespaces=NS)
    return found[0] if found else None


def _remove_and_prune(el: Optional[etree._Element]) -> None:
    """Remove el; also remove its w:r if nothing but run properties is left."""
    if el is None or el.getparent() is None:
        return
    parent = el.getparent()
    parent.remove(el)
    if parent.tag == q("w:r") and all(c.tag == q("w:rPr") for c in parent):
        grand = parent.getparent()
        if grand is not None:
            grand.remove(parent)


def _drop_fallback(container: Optional[etree._Element]) -> None:
    alt = _ancestor(container, "mc:AlternateContent")
    if alt is None:
        return
    for fb in alt.findall("mc:Fallback", NS):
        alt.remove(fb)


def _has_shape_children(container: etree._Element) -> bool:
    for tag in ("wps:wsp", "pic:pic", "wpg:grpSp", "wpg:wgp"):
        if container.find(".//" + tag, NS) is not None:
            return True
    return False


def _dml_id(wsp: etree._Element, container: str) -> tuple[str, str]:
    c_nv = wsp.find("wps:cNvPr", NS)
    if c_nv is not None:
        return c_nv.get("id", ""), c_nv.get("name", "")
    if container == "top":
        inline_or_anchor = wsp.xpath("ancestor::wp:anchor[1] | ancestor::wp:inline[1]", namespaces=NS)
        if inline_or_anchor:
            doc_pr = inline_or_anchor[0].find("wp:docPr", NS)
            if doc_pr is not None:
                return doc_pr.get("id", ""), doc_pr.get("name", "")
    return "", ""


def _body_paragraphs(root: etree._Element) -> list[etree._Element]:
    """Paragraphs of the story itself, not the ones inside text boxes."""
    return [p for p in root.iter(q("w:p")) if not has_ancestor(p, "w:txbxContent")]


def _paragraph_index(root: etree._Element) -> dict[etree._Element, int]:
    return {p: i for i, p in enumerate(_body_paragraphs(root))}


def _para_of(el: etree._Element, index: dict[etree._Element, int]) -> int:
    for p in el.iterancestors(q("w:p")):
        if p in index:
            return index[p]
    return -1


def paragraph_plain_text(p: etree._Element) -> str:
    texts = [t.text or "" for t in p.iter(q("w:t")) if not has_ancestor(t, "w:txbxContent")]
    return re.sub(r"\s+", " ", "".join(texts)).strip()


def validate_package(path: Path, parts: set[str]) -> None:
    """The written file must be a readable zip whose modified parts are well-formed XML."""
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad:
            raise ValueError(f"Corrupt zip entry after save: {bad}")
        for name in parts:
            etree.fromstring(zf.read(name))
