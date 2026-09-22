from __future__ import annotations

import json
import zipfile

from conftest import (canvas, connector, group, ln, para, solid, sp_pr, top_shape, wsp)
from lxml import etree

from docx_cleanup.__main__ import main, run
from docx_cleanup.package import NS, DocxPackage
from docx_cleanup.rules.reviewer_boxes import ReviewerBoxes

RED_BOX = sp_pr("roundRect", "<a:noFill/>", ln("FF0000"))


def detect(path, **params):
    return ReviewerBoxes(params).detect(DocxPackage(path))


def apply(path, out, **params):
    pkg = DocxPackage(path)
    rule = ReviewerBoxes(params)
    found = rule.detect(pkg)
    rule.apply(pkg, found)
    pkg.save(out)
    return found


def document(path) -> etree._Element:
    with zipfile.ZipFile(path) as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def count(root, xpath) -> int:
    return int(root.xpath(f"count({xpath})", namespaces=NS))


# ---------------------------------------------------------------- detection

def test_red_rounded_box_around_text_is_high(make_docx):
    f = detect(make_docx(para(top_shape(RED_BOX), "Important text")))
    assert [x.confidence for x in f] == ["high"]
    assert f[0].location == "Important text"
    assert f[0].features["lineColor"] == "FF0000"


def test_plain_rectangle_and_dark_red_match(make_docx):
    f = detect(make_docx(para(top_shape(sp_pr("rect", "<a:noFill/>", ln("C00000"))), "x")))
    assert [x.confidence for x in f] == ["high"]


def test_box_with_text_is_only_low(make_docx):
    f = detect(make_docx(para(top_shape(RED_BOX, text="Error state"))))
    assert [(x.confidence, x.reason) for x in f] == [("low", "fails: text")]


def test_filled_box_is_only_low(make_docx):
    f = detect(make_docx(para(top_shape(sp_pr("roundRect", solid("FFFFFF"), ln("FF0000"))))))
    assert [(x.confidence, x.reason) for x in f] == [("low", "fails: fill")]


def test_nearly_transparent_fill_counts_as_no_fill(make_docx):
    f = detect(make_docx(para(top_shape(sp_pr("roundRect", solid("FF0000", alpha=5000), ln("FF0000"))))))
    assert [x.confidence for x in f] == ["high"]


def test_ellipse_is_only_low(make_docx):
    f = detect(make_docx(para(top_shape(sp_pr("ellipse", "<a:noFill/>", ln("FF0000"))))))
    assert [(x.confidence, x.reason) for x in f] == [("low", "fails: geometry")]


def test_blue_and_black_boxes_are_ignored(make_docx):
    f = detect(make_docx(para(top_shape(sp_pr("rect", "<a:noFill/>", ln("0000FF")), doc_id=1)),
                         para(top_shape(sp_pr("rect", "<a:noFill/>", ln("000000")), doc_id=2))))
    assert f == []


def test_theme_color_is_resolved(make_docx):
    box = sp_pr("roundRect", "<a:noFill/>", ln("accent2", scheme=True))
    assert [x.confidence for x in detect(make_docx(para(top_shape(box)), accent2="FF0000"))] == ["high"]
    # default Office accent2 is orange: reddish, but not red enough -> low
    assert [x.reason for x in detect(make_docx(para(top_shape(box))))] == ["fails: color"]


def test_line_from_shape_style(make_docx):
    style = ('<wps:style><a:lnRef idx="2"><a:srgbClr val="FF0000"/></a:lnRef>'
             '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
             '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
             '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>')
    shape = wsp('<wps:spPr><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></wps:spPr>', extra=style)
    from conftest import URI_WPS, anchored
    f = detect(make_docx(para(anchored(URI_WPS, shape, 1))))
    assert [x.confidence for x in f] == ["high"]


def test_fallback_copy_is_not_counted_twice(make_docx):
    # the mc:Fallback of each drawing holds a red VML roundrect as well
    assert len(detect(make_docx(para(top_shape(RED_BOX))))) == 1


def test_box_in_group_is_medium(make_docx):
    children = wsp(RED_BOX, cnv_id=11) + wsp(sp_pr("rect", solid("DDDDDD"), ln("000000")), "Content", cnv_id=12)
    f = detect(make_docx(para(group(children))))
    assert [(x.confidence, x.container) for x in f] == [("medium", "group")]


def test_box_with_connector_in_canvas_is_diagram_content(make_docx):
    children = (wsp(RED_BOX, cnv_id=21) + wsp(sp_pr("rect", "<a:noFill/>", ln("000000")), cnv_id=23)
                + connector(22, 21, 23))
    f = detect(make_docx(para(canvas(children))))
    assert [(x.confidence, x.reason) for x in f] == [("low", "fails: connected")]


def test_connector_ids_are_scoped_to_their_canvas(make_docx):
    # canvas A connects shape id 21; the red box in canvas B also has id 21
    a = wsp(sp_pr("rect", "<a:noFill/>", ln("000000")), cnv_id=21) + connector(22, 21, 21)
    b = wsp(RED_BOX, cnv_id=21)
    f = detect(make_docx(para(canvas(a, doc_id=30)), para(canvas(b, doc_id=31))))
    assert [x.confidence for x in f] == ["medium"]


def test_vml_roundrect(make_docx):
    vml = '<w:pict><v:roundrect id="r1" strokecolor="red" strokeweight="2.25pt" filled="f"/></w:pict>'
    f = detect(make_docx(para(vml, "Legacy text")))
    assert [x.confidence for x in f] == ["high"]
    assert f[0].features["lineWidthPt"] == 2.25


def test_vml_rect_filled_by_default(make_docx):
    vml = '<w:pict><v:rect id="r1" strokecolor="#FF0000"/></w:pict>'  # VML default: filled white
    assert [x.reason for x in detect(make_docx(para(vml)))] == ["fails: fill"]


# ---------------------------------------------------------------- removal

def test_apply_removes_top_level_box_and_keeps_text(make_docx, tmp_path):
    src = make_docx(para(top_shape(RED_BOX), "Important text"))
    out = tmp_path / "out.docx"
    f = apply(src, out)
    assert [x.action for x in f] == ["removed"]
    root = document(out)
    assert count(root, "//mc:AlternateContent") == 0
    assert count(root, "//w:drawing") == 0
    assert count(root, "//w:pict") == 0
    assert "".join(root.itertext()) == "Important text"


def test_apply_leaves_low_confidence_alone(make_docx, tmp_path):
    src = make_docx(para(top_shape(RED_BOX, text="Error state")))
    out = tmp_path / "out.docx"
    assert [x.action for x in apply(src, out)] == ["reported"]
    assert count(document(out), "//wps:wsp") == 1


def test_group_box_kept_unless_include_in_groups(make_docx, tmp_path):
    children = wsp(RED_BOX, cnv_id=11) + wsp(sp_pr("rect", solid("DDDDDD"), ln("000000")), "Content", cnv_id=12)
    src = make_docx(para(group(children)))

    out1 = tmp_path / "keep.docx"
    assert [x.action for x in apply(src, out1)] == ["reported (set includeInGroups to remove)"]
    assert count(document(out1), "//wpg:wgp/wps:wsp") == 2

    out2 = tmp_path / "remove.docx"
    assert [x.action for x in apply(src, out2, includeInGroups=True)] == ["removed"]
    root = document(out2)
    assert count(root, "//wpg:wgp/wps:wsp") == 1
    assert count(root, "//mc:Fallback") == 0          # stale VML copy dropped
    assert count(root, "//mc:Choice") == 1


def test_group_with_only_boxes_is_removed_entirely(make_docx, tmp_path):
    src = make_docx(para(group(wsp(RED_BOX, cnv_id=11) + wsp(RED_BOX, cnv_id=12)), "Text"))
    out = tmp_path / "out.docx"
    apply(src, out, includeInGroups=True)
    root = document(out)
    assert count(root, "//w:drawing") == 0
    assert count(root, "//mc:AlternateContent") == 0


def test_apply_removes_vml_box(make_docx, tmp_path):
    src = make_docx(para('<w:pict><v:roundrect id="r1" strokecolor="red" filled="f"/></w:pict>', "Text"))
    out = tmp_path / "out.docx"
    apply(src, out)
    root = document(out)
    assert count(root, "//w:pict") == 0
    assert count(root, "//w:r") == 1


# ---------------------------------------------------------------- package / CLI

def test_untouched_parts_are_byte_identical(make_docx, tmp_path):
    src = make_docx(para(top_shape(RED_BOX), "Text"))
    out = tmp_path / "out.docx"
    apply(src, out)
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(out) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name != "word/document.xml":
                assert a.read(name) == b.read(name), name


def _config(tmp_path, **rule):
    path = tmp_path / "cfg.json"
    path.write_text(json.dumps({"version": 1, "rules": {"ReviewerBoxes": rule}}), encoding="utf-8")
    return str(path)


def test_cli_report_mode_changes_nothing(make_docx, tmp_path):
    src = make_docx(para(top_shape(RED_BOX), "Text"))
    out = tmp_path / "out.docx"
    assert main([str(src), "-o", str(out), "--config", _config(tmp_path, mode="Apply"), "--mode", "Report"]) == 0
    assert out.read_bytes() == src.read_bytes()
    manifest = (tmp_path / "out.cleanup-manifest.csv").read_text(encoding="utf-8-sig")
    assert "ReviewerBoxes,reported,high" in manifest


def test_cli_off_rule_is_not_enabled_by_mode_override(make_docx, tmp_path):
    src = make_docx(para(top_shape(RED_BOX), "Text"))
    out = tmp_path / "out.docx"
    assert main([str(src), "-o", str(out), "--config", _config(tmp_path, mode="Off"), "--mode", "Apply"]) == 0
    assert out.read_bytes() == src.read_bytes()


def test_cli_rejects_unknown_parameter(make_docx, tmp_path, capsys):
    src = make_docx(para(top_shape(RED_BOX)))
    cfg = _config(tmp_path, mode="Apply", colour=["FF0000"])
    assert main([str(src), "-o", str(tmp_path / "o.docx"), "--config", cfg]) == 1
    assert "unknown parameter" in capsys.readouterr().err


def test_shipped_profiles_are_valid(make_docx, tmp_path):
    src = make_docx(para(top_shape(RED_BOX), "Text"))
    for profile in ("default", "ns"):
        assert main([str(src), "-o", str(tmp_path / f"{profile}.docx"), "--profile", profile]) == 0
    assert (tmp_path / "default.docx").read_bytes() == src.read_bytes()
    assert count(document(tmp_path / "ns.docx"), "//wps:wsp") == 0


def test_run_refuses_to_overwrite_input(make_docx):
    src = make_docx(para("", "Text"))
    try:
        run(src, src, {"rules": {}})
    except ValueError as exc:
        assert "differ" in str(exc)
    else:
        raise AssertionError("expected ValueError")
