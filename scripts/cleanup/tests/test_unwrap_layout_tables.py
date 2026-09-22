from __future__ import annotations

import zipfile

from conftest import (RED_BOX_SPPR, caption, inline_picture, para, table, top_shape)
from lxml import etree

from docx_cleanup.__main__ import load_config, run
from docx_cleanup.package import NS, DocxPackage, paragraph_plain_text
from docx_cleanup.rules.unwrap_layout_tables import UnwrapLayoutTables

DATA_TABLE = table("<w:p><w:r><w:t>No.</w:t></w:r></w:p>", "<w:p><w:r><w:t>1.</w:t></w:r></w:p>", cols=2)


def detect(path, **params):
    return UnwrapLayoutTables(params).detect(DocxPackage(path))


def apply(path, out, **params):
    pkg = DocxPackage(path)
    rule = UnwrapLayoutTables(params)
    found = rule.detect(pkg)
    rule.apply(pkg, found)
    pkg.save(out)
    return found


def body(path) -> etree._Element:
    with zipfile.ZipFile(path) as zf:
        return etree.fromstring(zf.read("word/document.xml")).find("w:body", NS)


def block_kinds(path) -> list[str]:
    """Top-level blocks of the body: 'tbl' or the paragraph text ('img' for a picture)."""
    out = []
    for b in body(path):
        if b.tag.endswith("}tbl"):
            out.append("tbl")
        elif b.tag.endswith("}p"):
            if b.find(".//w:drawing", NS) is not None:
                out.append("img")
            else:
                out.append(paragraph_plain_text(b))
    return out


# ---------------------------------------------------------------- detection

def test_figure_with_caption_in_one_cell_is_high(make_docx):
    src = make_docx(para(text="Before"), table(inline_picture() + caption("Judgement state")), para(text="After"))
    f = detect(src)
    assert [(x.confidence, x.reason) for x in f] == [("high", "caption + figure/table")]
    assert f[0].location == "Fig. 30 Judgement state"


def test_caption_and_figure_in_two_rows_is_high(make_docx):
    src = make_docx(table(inline_picture(), caption("Transition")))
    assert [x.confidence for x in detect(src)] == ["high"]


def test_data_table_wrapped_with_its_caption_is_high(make_docx):
    src = make_docx(table(caption("List of input hardwire signals") + DATA_TABLE + "<w:p/>"))
    f = detect(src)
    assert [x.confidence for x in f] == ["high"]          # the 2-column data table itself is not a candidate
    assert f[0].features["nestedTables"] == 1


def test_caption_found_by_derived_style_or_seq_field(make_docx):
    by_style = make_docx(table(inline_picture() + caption("x", style="FigCaption", seq=False)))
    by_seq = make_docx(table(inline_picture() + caption("x", style="", seq=True)))
    no_caption = make_docx(table(inline_picture() + caption("x", style="", seq=False)))
    assert [x.confidence for x in detect(by_style)] == ["high"]
    assert [x.confidence for x in detect(by_seq)] == ["high"]
    assert [x.confidence for x in detect(no_caption)] == ["medium"]


def test_floating_shape_counts_as_figure(make_docx):
    src = make_docx(table(para(top_shape(RED_BOX_SPPR)) + caption("Fig")))
    f = detect(src)
    assert [x.confidence for x in f] == ["high"]
    assert f[0].features["drawings"] == 1                  # mc:Fallback copy not counted


def test_text_only_box_is_low(make_docx):
    src = make_docx(table("<w:p><w:r><w:t>Note: do not reset.</w:t></w:r></w:p>"))
    assert [(x.confidence, x.reason) for x in detect(src)] == [("low", "text only (box / note?)")]


def test_long_single_column_table_is_low_or_ignored(make_docx):
    rows = [inline_picture() + caption("c")] + ["<w:p><w:r><w:t>row</w:t></w:r></w:p>"] * 5
    assert [x.reason for x in detect(make_docx(table(*rows)))] == ["fails: rows (6 > maxRows)"]
    text_rows = ["<w:p><w:r><w:t>item</w:t></w:r></w:p>"] * 6
    assert detect(make_docx(table(*text_rows))) == []


def test_multi_column_tables_are_never_candidates(make_docx):
    src = make_docx(table(inline_picture() + caption("a"), cols=2))
    assert detect(src) == []


# ---------------------------------------------------------------- unwrap

def test_unwrap_keeps_content_in_order(make_docx, tmp_path):
    src = make_docx(para(text="Before"), table(inline_picture() + caption("Judgement state")), para(text="After"))
    out = tmp_path / "out.docx"
    assert [x.action for x in apply(src, out)] == ["unwrapped"]
    assert block_kinds(out) == ["Before", "img", "Fig. 30 Judgement state", "After"]


def test_unwrap_two_rows_and_nested_data_table(make_docx, tmp_path):
    src = make_docx(table(caption("List of signals"), DATA_TABLE + "<w:p/>"))
    out = tmp_path / "out.docx"
    apply(src, out)
    assert block_kinds(out) == ["Fig. 30 List of signals", "tbl", ""]
    assert len(body(out).find("w:tbl", NS).findall("w:tr/w:tc", NS)) == 4   # data table intact


def test_bookmark_of_caption_survives(make_docx, tmp_path):
    cap = caption("x").replace("</w:pPr>", '</w:pPr><w:bookmarkStart w:id="1" w:name="_Ref136255150"/>', 1)
    cap = cap.replace("</w:p>", '<w:bookmarkEnd w:id="1"/></w:p>')
    src = make_docx(table(inline_picture() + cap))
    out = tmp_path / "out.docx"
    apply(src, out)
    names = [b.get(f"{{{NS['w']}}}name") for b in body(out).iter(f"{{{NS['w']}}}bookmarkStart")]
    assert names == ["_Ref136255150"]


def test_medium_needs_opt_in(make_docx, tmp_path):
    src = make_docx(table(inline_picture()))
    out1, out2 = tmp_path / "a.docx", tmp_path / "b.docx"
    assert [x.action for x in apply(src, out1)] == ["reported (set unwrapWithoutCaption to unwrap)"]
    assert block_kinds(out1) == ["tbl"]
    assert [x.action for x in apply(src, out2, unwrapWithoutCaption=True)] == ["unwrapped"]
    assert block_kinds(out2) == ["img"]


def test_nested_wrappers_are_both_unwrapped(make_docx, tmp_path):
    inner = table(inline_picture() + caption("inner"))
    src = make_docx(table(inner + caption("outer")))
    out = tmp_path / "out.docx"
    assert [x.action for x in apply(src, out)] == ["unwrapped", "unwrapped"]
    assert block_kinds(out) == ["img", "Fig. 30 inner", "Fig. 30 outer"]


def test_ns_profile_runs_both_rules_in_order(make_docx, tmp_path):
    # the red box sits inside the wrapper: it is removed first, then the wrapper is unwrapped
    src = make_docx(table(para(top_shape(RED_BOX_SPPR), "Important") + inline_picture() + caption("Fig")))
    out = tmp_path / "out.docx"
    findings = run(src, out, load_config("ns", None))
    assert [(f.rule_id, f.action) for f in findings] == [("ReviewerBoxes", "removed"),
                                                        ("UnwrapLayoutTables", "unwrapped")]
    assert block_kinds(out) == ["Important", "img", "Fig. 30 Fig"]


def test_cli_summary_counts_unwrapped(make_docx, tmp_path, capsys):
    src = make_docx(table(inline_picture() + caption("x")))
    run(src, tmp_path / "out.docx", load_config("ns", None))
    assert "UnwrapLayoutTables [Apply]: 1 found, 1 changed" in capsys.readouterr().out
