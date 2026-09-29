from __future__ import annotations

import os
import shutil
import subprocess
import zipfile

import pytest
from conftest import para
from lxml import etree

from docx_cleanup.__main__ import load_config, run
from docx_cleanup.package import NS, DocxPackage
from docx_cleanup.rules.expand_simple_fields import ExpandSimpleFields

W = NS["w"]


def simple(instr: str, result: str, attrs: str = "") -> str:
    return (f'<w:fldSimple w:instr=" {instr} "{attrs}>'
            f'<w:r><w:rPr><w:b/></w:rPr><w:t>{result}</w:t></w:r></w:fldSimple>')


def caption_simple() -> str:
    return ('<w:p><w:r><w:t xml:space="preserve">Table </w:t></w:r>'
            + simple("STYLEREF 1 \\s", "1") + '<w:r><w:noBreakHyphen/></w:r>'
            + simple("SEQ Table \\* ARABIC \\s 1", "3")
            + '<w:r><w:t xml:space="preserve"> Example</w:t></w:r></w:p>')


def apply(path, out):
    pkg = DocxPackage(path)
    rule = ExpandSimpleFields()
    found = rule.detect(pkg)
    rule.apply(pkg, found)
    pkg.save(out)
    return found


def paragraph(path, index=0) -> etree._Element:
    with zipfile.ZipFile(path) as zf:
        return etree.fromstring(zf.read("word/document.xml")).findall(".//w:p", NS)[index]


def test_detects_every_simple_field(make_docx):
    f = ExpandSimpleFields().detect(DocxPackage(make_docx(caption_simple())))
    assert [(x.reason, x.features["result"]) for x in f] == [("fldSimple STYLEREF", "1"), ("fldSimple SEQ", "3")]
    assert f[0].location == "Table 13 Example"


def test_expands_to_complex_field_with_same_instruction_and_result(make_docx, tmp_path):
    out = tmp_path / "out.docx"
    assert [x.action for x in apply(make_docx(caption_simple()), out)] == ["expanded", "expanded"]
    p = paragraph(out)
    assert p.find(".//w:fldSimple", NS) is None
    kinds = []
    for r in p.findall("w:r", NS):
        fc = r.find("w:fldChar", NS)
        it = r.find("w:instrText", NS)
        t = r.find("w:t", NS)
        kinds.append(fc.get(f"{{{W}}}fldCharType") if fc is not None
                     else f"instr:{it.text.strip()}" if it is not None
                     else f"t:{t.text}" if t is not None else "other")
    assert kinds == ["t:Table ",
                     "begin", "instr:STYLEREF 1 \\s", "separate", "t:1", "end",
                     "other",
                     "begin", "instr:SEQ Table \\* ARABIC \\s 1", "separate", "t:3", "end",
                     "t: Example"]
    # run formatting of the result is reused for the field-code runs
    assert all(r.find("w:rPr/w:b", NS) is not None for r in p.findall("w:r", NS)[1:6])


def test_keeps_lock_and_dirty_flags(make_docx, tmp_path):
    src = make_docx('<w:p>' + simple("SEQ Figure", "2", ' w:fldLock="1" w:dirty="true"') + '</w:p>')
    out = tmp_path / "out.docx"
    apply(src, out)
    begin = paragraph(out).find(".//w:fldChar", NS)
    assert begin.get(f"{{{W}}}fldLock") == "1" and begin.get(f"{{{W}}}dirty") == "true"


def test_simple_field_inside_hyperlink_and_nested(make_docx, tmp_path):
    nested = simple("REF _Ref1 \\h", "Table 1-2").replace(
        '<w:r><w:rPr><w:b/></w:rPr><w:t>Table 1-2</w:t></w:r>', simple("QUOTE x", "Table 1-2"))
    src = make_docx('<w:p><w:hyperlink w:anchor="_Ref1">' + nested + '</w:hyperlink></w:p>')
    out = tmp_path / "out.docx"
    assert len(apply(src, out)) == 2
    p = paragraph(out)
    assert p.find(".//w:fldSimple", NS) is None
    assert len(p.findall("w:hyperlink/w:r/w:fldChar", NS)) == 6        # 2 fields x begin/separate/end
    assert "".join(t.text for t in p.iter(f"{{{W}}}t")) == "Table 1-2"


def test_profile_pandoc_is_valid(make_docx, tmp_path):
    out = tmp_path / "out.docx"
    findings = run(make_docx(caption_simple()), out, load_config("pandoc", None))
    assert [f.action for f in findings] == ["expanded", "expanded"]


def _pandoc() -> str | None:
    return os.environ.get("PANDOC") or shutil.which("pandoc")


@pytest.mark.skipif(_pandoc() is None, reason="pandoc not installed (set PANDOC=/path/to/pandoc)")
def test_pandoc_keeps_caption_numbers_and_cross_references(make_docx, tmp_path):
    body = (caption_simple()
            + '<w:p><w:r><w:t xml:space="preserve">shown in </w:t></w:r>'
            + simple("REF _Ref1 \\h", "Table 1-3") + '</w:p>'
            + '<w:p><w:bookmarkStart w:id="1" w:name="_Ref1"/><w:bookmarkEnd w:id="1"/></w:p>')
    src = make_docx(body)
    out = tmp_path / "out.docx"
    apply(src, out)

    def md(path):
        return subprocess.run([_pandoc(), "-f", "docx", "-t", "gfm", "--wrap=none", str(path)],
                              capture_output=True, text=True, check=True).stdout

    before, after = (md(p).replace("**", "") for p in (src, out))   # fixture results are bold
    assert "Table 1‑3 Example" not in before and "shown in Table" not in before
    assert "Table 1‑3 Example" in after
    # pandoc 3.11 turns REF into a link; older versions keep it as plain text
    assert "shown in [Table 1-3](#_Ref1)" in after or "shown in Table 1-3" in after
