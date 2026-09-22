"""Build minimal .docx files with specific drawings, so tests need no binary fixtures."""
from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

NSDECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

# Localized Word: built-in Caption has id "a3" but English name "caption";
# "FigCaption" is a custom style based on it.
STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'}>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="a3"><w:name w:val="caption"/><w:basedOn w:val="a"/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FigCaption"><w:name w:val="Fig Caption"/><w:basedOn w:val="a3"/></w:style>
</w:styles>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"""


def theme_xml(accent2: str = "ED7D31") -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements><a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="{accent2}"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>"""


# ---------------------------------------------------------------- shape snippets

def sp_pr(geom: str = "roundRect", fill: str = "<a:noFill/>", line: str = "") -> str:
    return f'<wps:spPr><a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>{fill}{line}</wps:spPr>'


def ln(color: str = "FF0000", w: int = 28575, scheme: bool = False) -> str:
    clr = f'<a:schemeClr val="{color}"/>' if scheme else f'<a:srgbClr val="{color}"/>'
    return f'<a:ln w="{w}"><a:solidFill>{clr}</a:solidFill></a:ln>'


# red rounded rectangle without fill: a reviewer box
RED_BOX_SPPR = sp_pr("roundRect", "<a:noFill/>", ln("FF0000"))


def solid(color: str, alpha: int | None = None) -> str:
    a = f'<a:alpha val="{alpha}"/>' if alpha is not None else ""
    return f'<a:solidFill><a:srgbClr val="{color}">{a}</a:srgbClr></a:solidFill>'


def txbx(text: str) -> str:
    return f"<wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>"


def wsp(sppr: str, text: str = "", cnv_id: int | None = None, extra: str = "") -> str:
    cnv = f'<wps:cNvPr id="{cnv_id}" name="Shape {cnv_id}"/>' if cnv_id is not None else ""
    return f"<wps:wsp>{cnv}<wps:cNvSpPr/>{sppr}{extra}{txbx(text) if text else ''}<wps:bodyPr/></wps:wsp>"


def connector(cnv_id: int, start: int, end: int) -> str:
    return (f'<wps:wsp><wps:cNvPr id="{cnv_id}" name="Connector {cnv_id}"/>'
            f'<wps:cNvCnPr><a:stCxn id="{start}" idx="0"/><a:endCxn id="{end}" idx="0"/></wps:cNvCnPr>'
            f'{sp_pr("straightConnector1", "", ln("000000"))}<wps:bodyPr/></wps:wsp>')


def anchored(graphic_uri: str, inner: str, doc_id: int, fallback: str = '<w:pict><v:roundrect strokecolor="red" filled="f"/></w:pict>') -> str:
    """A floating drawing wrapped in mc:AlternateContent, as Word writes it."""
    return (
        "<mc:AlternateContent><mc:Choice Requires=\"wps\"><w:drawing>"
        f'<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" '
        f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/><wp:extent cx="1000000" cy="500000"/>'
        f'<wp:docPr id="{doc_id}" name="Drawing {doc_id}"/>'
        f'<a:graphic><a:graphicData uri="{graphic_uri}">{inner}</a:graphicData></a:graphic>'
        "</wp:anchor></w:drawing></mc:Choice>"
        f"<mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"
    )


URI_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
URI_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
URI_WPC = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"


def top_shape(sppr: str, text: str = "", doc_id: int = 1) -> str:
    return anchored(URI_WPS, wsp(sppr, text), doc_id)


def group(children: str, doc_id: int = 10) -> str:
    return anchored(URI_WPG, f"<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr/>{children}</wpg:wgp>", doc_id)


def canvas(children: str, doc_id: int = 20) -> str:
    return anchored(URI_WPC, f"<wpc:wpc><wpc:bg/><wpc:whole/>{children}</wpc:wpc>", doc_id)


def para(drawing: str = "", text: str = "") -> str:
    run_d = f"<w:r>{drawing}</w:r>" if drawing else ""
    run_t = f"<w:r><w:t>{text}</w:t></w:r>" if text else ""
    return f"<w:p>{run_d}{run_t}</w:p>"


def caption(text: str, style: str = "a3", seq: bool = True) -> str:
    """Caption paragraph as Insert Caption writes it: label + SEQ field + text."""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    field = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
             '<w:r><w:instrText xml:space="preserve"> SEQ Figure \\* ARABIC </w:instrText></w:r>'
             '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>30</w:t></w:r>'
             '<w:r><w:fldChar w:fldCharType="end"/></w:r>') if seq else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">Fig. </w:t></w:r>{field}<w:r><w:t xml:space="preserve"> {text}</w:t></w:r></w:p>'


def table(*rows: str, cols: int = 1) -> str:
    """rows: the XML content of each cell (one cell per row unless cols > 1)."""
    trs = "".join("<w:tr>" + f"<w:tc><w:tcPr/>{r}</w:tc>" * cols + "</w:tr>" for r in rows)
    grid = "<w:gridCol/>" * cols
    return f'<w:tbl><w:tblPr/><w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl>'


def inline_picture() -> str:
    return ('<w:p><w:r><w:drawing><wp:inline><wp:docPr id="99" name="Picture 99"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')


@pytest.fixture
def make_docx(tmp_path):
    counter = {"n": 0}

    def _make(*paragraphs: str, accent2: str = "ED7D31") -> Path:
        counter["n"] += 1
        path = tmp_path / f"sample{counter['n']}.docx"
        body = "".join(paragraphs)
        doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               f"<w:document {NSDECL}><w:body>{body}<w:sectPr/></w:body></w:document>")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", CONTENT_TYPES)
            zf.writestr("_rels/.rels", ROOT_RELS)
            zf.writestr("word/document.xml", doc)
            zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
            zf.writestr("word/theme/theme1.xml", theme_xml(accent2))
            zf.writestr("word/styles.xml", STYLES)
        return path

    return _make
