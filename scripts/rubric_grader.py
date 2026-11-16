"""Weighted JSON Rubric Grader, Rubric Scaffolder, and Batch Exam Scorer."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from ooxml_patch import PT_PER_CM, audit_ooxml
from ref_diff import _norm_formula, _vals_equal, compare_documents
from wps_env import infer_kind, inspect_container, wps_session


@dataclass
class GradeContext:
    file_path: str
    kind: str
    xml: dict[str, Any]
    app: Any = None
    doc: Any = None


HandlerFn = Callable[[GradeContext, dict[str, Any]], tuple[bool, str]]
CHECK_HANDLERS: dict[str, HandlerFn] = {}


def register_check(name: str) -> Callable[[HandlerFn], HandlerFn]:
    def _wrap(fn: HandlerFn) -> HandlerFn:
        CHECK_HANDLERS[name] = fn
        return fn
    return _wrap


def _close(a: float, b: float, tol: float = 0.5) -> bool:
    return abs(float(a) - float(b)) <= tol


def _cm_close(pt_val: float, cm_val: float, tol_cm: float = 0.05) -> bool:
    return abs(float(pt_val) / PT_PER_CM - float(cm_val)) <= tol_cm


# ---------------------------------------------------------------------------
# Word Check Handlers
# ---------------------------------------------------------------------------

@register_check("word_page_setup")
def _chk_word_page_setup(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    sec_idx = int(p.get("section", 1))
    tol_cm = float(p.get("tol_cm", 0.05))
    details: list[str] = []

    if ctx.doc is not None:
        if sec_idx > int(ctx.doc.Sections.Count):
            return False, f"section {sec_idx} > count {ctx.doc.Sections.Count}"
        if "sections_count" in p and int(ctx.doc.Sections.Count) != int(p["sections_count"]):
            return False, f"sections_count={ctx.doc.Sections.Count} != {p['sections_count']}"
        ps = ctx.doc.Sections(sec_idx).PageSetup
        checks = []
        if "orientation" in p:
            checks.append(("orient", int(ps.Orientation) == int(p["orientation"]), int(ps.Orientation)))
        if "section_start" in p:
            checks.append(("secStart", int(ps.SectionStart) == int(p["section_start"]), int(ps.SectionStart)))
        if "gutter_pos" in p:
            checks.append(("gutterPos", int(ps.GutterPos) == int(p["gutter_pos"]), int(ps.GutterPos)))
        if "page_width_cm" in p:
            checks.append(("pageW", _cm_close(ps.PageWidth, p["page_width_cm"], tol_cm), round(ps.PageWidth / PT_PER_CM, 2)))
        if "page_height_cm" in p:
            checks.append(("pageH", _cm_close(ps.PageHeight, p["page_height_cm"], tol_cm), round(ps.PageHeight / PT_PER_CM, 2)))
        if "margins_cm" in p:
            m = p["margins_cm"]
            vals = [ps.TopMargin, ps.BottomMargin, ps.LeftMargin, ps.RightMargin]
            exp = [m["top"], m["bottom"], m["left"], m["right"]] if isinstance(m, dict) else list(m)
            ok_m = all(_cm_close(v, e, tol_cm) for v, e in zip(vals, exp))
            checks.append(("margins_cm", ok_m, [round(v / PT_PER_CM, 2) for v in vals]))
        failed = [f"{k}={v}" for k, ok, v in checks if not ok]
        return not failed, ("OK" if not failed else f"failed: {', '.join(failed)}")

    # XML fallback
    secs = ctx.xml.get("sections", [])
    if "sections_count" in p and len(secs) != int(p["sections_count"]):
        return False, f"xml sections_count={len(secs)} != {p['sections_count']}"
    if sec_idx > len(secs):
        return False, f"xml section {sec_idx} missing"
    sec = secs[sec_idx - 1]
    pgMar = sec.get("pgMar", {})
    if "margins_cm" in p:
        m = p["margins_cm"]
        exp = [m["top"], m["bottom"], m["left"], m["right"]] if isinstance(m, dict) else list(m)
        keys = ("top", "bottom", "left", "right")
        got = [int(pgMar.get(k, 0)) / 20.0 / PT_PER_CM for k in keys]
        if not all(abs(g - e) <= tol_cm for g, e in zip(got, exp)):
            return False, f"xml margins_cm={[round(x, 2) for x in got]} != {exp}"
    return True, "OK (XML)"


@register_check("word_style")
def _chk_word_style(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        return False, "COM required for word_style"
    sid = p.get("style_id")
    st = ctx.doc.Styles(int(sid) if sid is not None else str(p["style_name"]))
    pf = st.ParagraphFormat
    f = st.Font
    fails = []
    if "font_size" in p and not _close(f.Size, p["font_size"], 0.1):
        fails.append(f"size={f.Size}")
    if "font_name" in p and p["font_name"] not in (f.Name, f.NameFarEast, f.NameAscii):
        fails.append(f"font={f.Name}/{f.NameFarEast}")
    if "bold" in p and bool(f.Bold) != bool(p["bold"]):
        fails.append(f"bold={f.Bold}")
    if "align" in p and int(pf.Alignment) != int(p["align"]):
        fails.append(f"align={pf.Alignment}")
    if "space_before" in p and not _close(pf.SpaceBefore, p["space_before"], 0.2):
        fails.append(f"before={pf.SpaceBefore}")
    if "space_after" in p and not _close(pf.SpaceAfter, p["space_after"], 0.2):
        fails.append(f"after={pf.SpaceAfter}")
    if "line_unit_before" in p and not _close(pf.LineUnitBefore, p["line_unit_before"], 0.05):
        fails.append(f"unitBefore={pf.LineUnitBefore}")
    if "line_unit_after" in p and not _close(pf.LineUnitAfter, p["line_unit_after"], 0.05):
        fails.append(f"unitAfter={pf.LineUnitAfter}")
    if "line_spacing_rule" in p and int(pf.LineSpacingRule) != int(p["line_spacing_rule"]):
        fails.append(f"lineRule={pf.LineSpacingRule}")
    if "bottom_border" in p:
        bd = pf.Borders(-3)
        bb = p["bottom_border"]
        if bb.get("has_border") and int(bd.LineStyle) == 0:
            fails.append("no_bottom_border")
        if "line_width" in bb and int(bd.LineWidth) != int(bb["line_width"]):
            fails.append(f"bdrW={bd.LineWidth}")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("word_paragraph_match")
def _chk_word_para(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        return False, "COM required for word_paragraph_match"
    target = None
    if "index" in p:
        idx = int(p["index"])
        if 1 <= idx <= int(ctx.doc.Paragraphs.Count):
            target = ctx.doc.Paragraphs(idx)
    else:
        sub = str(p.get("text_contains", ""))
        for i in range(1, int(ctx.doc.Paragraphs.Count) + 1):
            cand = ctx.doc.Paragraphs(i)
            if sub in str(cand.Range.Text):
                target = cand
                break
    if target is None:
        return False, "target paragraph not found"

    pf = target.Format
    rf = target.Range.Font
    fails = []
    if "style_name" in p and p["style_name"] not in str(target.Style.NameLocal):
        fails.append(f"style={target.Style.NameLocal}")
    if "list_string" in p and str(target.Range.ListFormat.ListString) != str(p["list_string"]):
        fails.append(f"list={target.Range.ListFormat.ListString!r}")
    if "align" in p and int(pf.Alignment) != int(p["align"]):
        fails.append(f"align={pf.Alignment}")
    if "font_size" in p and not _close(rf.Size, p["font_size"], 0.1):
        fails.append(f"fontSize={rf.Size}")
    if "c_indent" in p and not _close(pf.CharacterUnitFirstLineIndent, p["c_indent"], 0.05):
        fails.append(f"cIndent={pf.CharacterUnitFirstLineIndent}")
    if "line_spacing_rule" in p and int(pf.LineSpacingRule) != int(p["line_spacing_rule"]):
        fails.append(f"lineRule={pf.LineSpacingRule}")
    if "line_spacing" in p and not _close(pf.LineSpacing, p["line_spacing"], 0.2):
        fails.append(f"lineSpacing={pf.LineSpacing}")
    if "unit_before" in p and not _close(pf.LineUnitBefore, p["unit_before"], 0.05):
        fails.append(f"unitBefore={pf.LineUnitBefore}")
    if "unit_after" in p and not _close(pf.LineUnitAfter, p["unit_after"], 0.05):
        fails.append(f"unitAfter={pf.LineUnitAfter}")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("word_text_absent")
def _chk_word_text_absent(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        return False, "COM required for word_text_absent"
    txt = str(ctx.doc.Content.Text)
    forbidden = list(p.get("forbidden", []))
    found = [s for s in forbidden if s in txt]
    return not found, ("OK" if not found else f"found forbidden={found!r}")


@register_check("word_header_watermark")
def _chk_word_hdr_wm(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    sec_idx = int(p.get("section", 1))
    fails = []
    if ctx.doc is not None:
        hdr = ctx.doc.Sections(sec_idx).Headers(1)
        if "link_to_previous" in p and bool(hdr.LinkToPrevious) != bool(p["link_to_previous"]):
            fails.append(f"LinkToPrevious={hdr.LinkToPrevious}")
        if "has_field_substr" in p:
            codes = [str(f.Code.Text) for f in hdr.Range.Fields]
            if not any(p["has_field_substr"].upper() in c.upper() for c in codes):
                fails.append(f"fields={codes}")
        if "shape_count" in p and int(hdr.Range.ShapeRange.Count) != int(p["shape_count"]):
            fails.append(f"shapeCount={hdr.Range.ShapeRange.Count}")
        if "watermark_text" in p:
            ok_wm = any(
                int(s.Type) == 15 and p["watermark_text"] in str(s.TextEffect.Text)
                for s in hdr.Shapes
            )
            if not ok_wm:
                fails.append("text_watermark_missing")
        if p.get("picture_no_washout") and hdr.Range.ShapeRange.Count > 0:
            sh = hdr.Range.ShapeRange(1)
            if int(sh.PictureFormat.ColorType) == 4:
                fails.append("picture_still_washout")
        if "bottom_border" in p:
            bd = hdr.Range.Paragraphs(1).Borders(-3)
            if p["bottom_border"].get("has_border") and int(bd.LineStyle) == 0:
                fails.append("header_no_bottom_border")

    secs_xml = ctx.xml.get("sections", [])
    if p.get("behind_doc") and sec_idx <= len(secs_xml):
        hdrs = secs_xml[sec_idx - 1].get("headers", [])
        if not any("1" in h.get("behind_doc", []) for h in hdrs):
            fails.append("xml_missing_behindDoc=1")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("word_numbering_xml")
def _chk_word_numbering(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    ilvl = int(p.get("ilvl", 0))
    num_fmt_prefix = str(p.get("num_fmt_prefix", ""))
    lvl_text = p.get("lvl_text")
    lvl_restart = p.get("lvl_restart")
    for ab in ctx.xml.get("numbering_abstracts", []):
        for lv in ab.get("levels", []):
            if lv["ilvl"] != ilvl:
                continue
            nf = str(lv.get("numFmt") or "")
            if num_fmt_prefix and not nf.startswith(num_fmt_prefix):
                continue
            if lvl_text is not None and lv.get("lvlText") != lvl_text:
                continue
            if lvl_restart is not None and lv.get("lvlRestart") not in (None, str(lvl_restart)):
                continue
            return True, f"matched {nf}/{lv.get('lvlText')}"
    return False, f"no numbering level matching ilvl={ilvl}, fmt={num_fmt_prefix}, text={lvl_text}"


@register_check("word_table")
def _chk_word_table(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    t_idx = int(p.get("table_index", 1))
    fails = []
    if ctx.doc is not None:
        if t_idx > int(ctx.doc.Tables.Count):
            return False, f"table {t_idx} missing"
        t = ctx.doc.Tables(t_idx)
        if "rows" in p and int(t.Rows.Count) != int(p["rows"]):
            fails.append(f"rows={t.Rows.Count}")
        if "cols" in p and int(t.Columns.Count) != int(p["cols"]):
            fails.append(f"cols={t.Columns.Count}")
        if "align" in p and int(t.Rows.Alignment) != int(p["align"]):
            fails.append(f"align={t.Rows.Alignment}")
        if "heading_repeat" in p and int(t.Rows(1).HeadingFormat) != (-1 if p["heading_repeat"] else 0):
            fails.append(f"headingRepeat={t.Rows(1).HeadingFormat}")
        for cc in p.get("cell_checks", []):
            cell = t.Cell(int(cc["row"]), int(cc["col"]))
            txt = str(cell.Range.Text).rstrip("\r\x07")
            if "text_contains" in cc and cc["text_contains"] not in txt:
                fails.append(f"cell({cc['row']},{cc['col']}) text={txt!r}")
            if "bg_color_bgr" in cc and int(cell.Shading.BackgroundPatternColor) != int(cc["bg_color_bgr"]):
                fails.append(f"cell({cc['row']},{cc['col']}) bg={cell.Shading.BackgroundPatternColor}")
            if "font_name" in cc and cell.Range.Characters.Count > 1:
                fn = cell.Range.Characters(1).Font.Name
                if fn != cc["font_name"]:
                    fails.append(f"cell({cc['row']},{cc['col']}) font={fn}")
    if p.get("has_wps_diagonals") and not ctx.xml.get("has_wps_diagonals"):
        fails.append("xml_missing_wpsCustomData:diagonals")
    if "theme_fill" in p and p["theme_fill"] not in ctx.xml.get("theme_fills", []):
        fails.append(f"xml_missing_themeFill={p['theme_fill']}")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("word_effects_xml")
def _chk_word_effects(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    fails = []
    for k in ("has_w14_textFill", "has_w14_glow", "has_w14_shadow", "has_w14_reflection"):
        if k in p and bool(ctx.xml.get(k)) != bool(p[k]):
            fails.append(f"{k}={ctx.xml.get(k)}")
    if "glow_sat_mod" in p and str(p["glow_sat_mod"]) not in ctx.xml.get("glow_sat_mods", []):
        fails.append(f"satMod={ctx.xml.get('glow_sat_mods')}")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("word_chart")
def _chk_word_chart(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    idx = int(p.get("chart_index", 1))
    charts = ctx.xml.get("charts", [])
    if idx > len(charts):
        return False, f"chart{idx}.xml missing"
    cxml = charts[idx - 1]
    fails = []
    if "series_count" in p and int(cxml["series_count"]) != int(p["series_count"]):
        fails.append(f"series={cxml['series_count']}")
    if "pt_count" in p and cxml["series"]:
        if int(cxml["series"][0]["val_pt_count"] or 0) != int(p["pt_count"]):
            fails.append(f"ptCount={cxml['series'][0]['val_pt_count']}")
    if "title" in p and cxml.get("title") != p["title"]:
        fails.append(f"xmlTitle={cxml.get('title')!r}")
    if ctx.doc is not None:
        try:
            ctx.doc.Repaginate()
        except Exception:
            pass
        pool = [ctx.doc.Shapes(i) for i in range(1, int(ctx.doc.Shapes.Count) + 1) if ctx.doc.Shapes(i).HasChart]
        if pool:
            sh = pool[0]
            if "chart_type" in p and int(sh.Chart.ChartType) != int(p["chart_type"]):
                fails.append(f"chartType={sh.Chart.ChartType}")
            if "wrap_type" in p and int(sh.WrapFormat.Type) != int(p["wrap_type"]):
                fails.append(f"wrap={sh.WrapFormat.Type}")
            if "width_cm" in p and not _cm_close(sh.Width, p["width_cm"], 0.15):
                fails.append(f"w_cm={sh.Width / PT_PER_CM:.2f}")
            if "height_cm" in p and not _cm_close(sh.Height, p["height_cm"], 0.15):
                fails.append(f"h_cm={sh.Height / PT_PER_CM:.2f}")
            if "page_number" in p and int(sh.Anchor.Information(3)) != int(p["page_number"]):
                fails.append(f"page={sh.Anchor.Information(3)}")
    return not fails, ("OK" if not fails else ", ".join(fails))


# ---------------------------------------------------------------------------
# Excel Check Handlers
# ---------------------------------------------------------------------------

@register_check("excel_sheets")
def _chk_excel_sheets(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    names = [s["name"] for s in ctx.xml.get("sheets", [])]
    if "exact_names" in p and names != list(p["exact_names"]):
        return False, f"sheets={names} != {p['exact_names']}"
    missing = [n for n in p.get("must_contain", []) if n not in names]
    return not missing, ("OK" if not missing else f"missing={missing}")


@register_check("excel_range_formula")
def _chk_excel_formula(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        return False, "COM required for excel_range_formula"
    ws = ctx.doc.Sheets(p["sheet"])
    raw = ws.Range(p["range"]).Formula
    grid = raw if isinstance(raw, tuple) else ((raw,),)
    if grid and not isinstance(grid[0], tuple):
        grid = (grid,)
    sub = _norm_formula(p.get("formula_contains", ""))
    exact = _norm_formula(p.get("formula_exact", ""))
    bad = []
    for ri, row in enumerate(grid, 1):
        for ci, cell_f in enumerate(row, 1):
            nf = _norm_formula(cell_f)
            if exact and nf != exact:
                bad.append((ri, ci, nf))
            elif sub and sub not in nf:
                bad.append((ri, ci, nf))
    return not bad, ("OK" if not bad else f"mismatches={len(bad)} sample={bad[:2]}")


@register_check("excel_range_values")
def _chk_excel_values(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        return False, "COM required for excel_range_values"
    ws = ctx.doc.Sheets(p["sheet"])
    tol = float(p.get("tol", 1e-4))
    bad = []
    for addr, exp in p.get("cells", {}).items():
        got = ws.Range(addr).Value
        if not _vals_equal(got, exp, tol=tol):
            bad.append(f"{addr}: {got!r} != {exp!r}")
    return not bad, ("OK" if not bad else "; ".join(bad[:5]))


@register_check("excel_number_format")
def _chk_excel_numfmt(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        return False, "COM required for excel_number_format"
    ws = ctx.doc.Sheets(p["sheet"])
    nf = str(ws.Range(p["range"]).NumberFormat or "")
    if "contains" in p:
        ok = str(p["contains"]) in nf
        return ok, f"NumberFormat={nf!r}"
    if "exact" in p:
        ok = nf == str(p["exact"])
        return ok, f"NumberFormat={nf!r}"
    return True, f"NumberFormat={nf!r}"


@register_check("excel_conditional_formatting")
def _chk_excel_cf(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    sname = p["sheet"]
    exp_sqref = str(p["sqref"]).replace("$", "")
    exp_fml = _norm_formula(p["formula_norm"]).lstrip("=")
    ws_xml = next((w for w in ctx.xml.get("worksheets", []) if w.get("name") == sname), None)
    if ws_xml is None:
        return False, f"sheet {sname!r} not found in XML"
    xml_ok = False
    for cf in ws_xml.get("conditional_formatting", []):
        if str(cf.get("sqref", "")).replace("$", "") == exp_sqref:
            for r in cf.get("rules", []):
                if any(_norm_formula(f).lstrip("=") == exp_fml for f in r.get("formulas", [])):
                    xml_ok = True
    if not xml_ok:
        return False, f"XML CF mismatch: {ws_xml.get('conditional_formatting')}"

    if ctx.doc is not None and "interior_color_bgr" in p:
        fcs = ctx.doc.Sheets(sname).Cells.FormatConditions
        ok_col = any(
            int(fcs(i).Interior.Color) == int(p["interior_color_bgr"])
            for i in range(1, int(fcs.Count) + 1)
        )
        if not ok_col:
            return False, "interior_color_bgr mismatch"
    return True, "OK"


@register_check("excel_sort_state")
def _chk_excel_sort(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    ws_xml = next((w for w in ctx.xml.get("worksheets", []) if w.get("name") == p["sheet"]), None)
    if ws_xml is None:
        return False, f"sheet {p['sheet']!r} not found"
    ss = ws_xml.get("sort_state")
    if not ss:
        return False, "sortState missing in XML"
    if "ref" in p and ss.get("ref") != p["ref"]:
        return False, f"sortState ref={ss.get('ref')} != {p['ref']}"
    if "conditions_count" in p and len(ss.get("conditions", [])) != int(p["conditions_count"]):
        return False, f"conditions={len(ss.get('conditions', []))}"
    return True, f"sortState={ss}"


@register_check("excel_pivot_table")
def _chk_excel_pivot(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        pts = ctx.xml.get("pivot_tables", [])
        ok = bool(pts)
        return ok, f"xml_pivots={pts}"
    ws = ctx.doc.Sheets(p["sheet"])
    if int(ws.PivotTables().Count) == 0:
        return False, "no PivotTable on sheet"
    pt = ws.PivotTables(1)
    rows = [str(f.Name) for f in pt.RowFields]
    datas = [str(f.Name) for f in pt.DataFields]
    fails = []
    if "row_fields" in p and rows != list(p["row_fields"]):
        fails.append(f"rowFields={rows}")
    if "data_fields_count" in p and len(datas) != int(p["data_fields_count"]):
        fails.append(f"dataFields={datas}")
    return not fails, ("OK" if not fails else ", ".join(fails))


# ---------------------------------------------------------------------------
# PPT Check Handlers
# ---------------------------------------------------------------------------

@register_check("ppt_page_setup")
def _chk_ppt_page_setup(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    fails = []
    if ctx.doc is not None:
        ps = ctx.doc.PageSetup
        if "slide_size" in p and int(ps.SlideSize) != int(p["slide_size"]):
            fails.append(f"SlideSize={ps.SlideSize}")
        if "width_pt" in p and not _close(ps.SlideWidth, p["width_pt"], 0.5):
            fails.append(f"W={ps.SlideWidth}")
        if "height_pt" in p and not _close(ps.SlideHeight, p["height_pt"], 0.5):
            fails.append(f"H={ps.SlideHeight}")
    if "slides_count" in p and int(ctx.xml.get("slides_count", 0)) != int(p["slides_count"]):
        fails.append(f"slides={ctx.xml.get('slides_count')}")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("ppt_print_options")
def _chk_ppt_print(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    prn = ctx.xml.get("prnPr", {})
    fails = []
    for k in ("prnWhat", "clrMode", "hiddenSlides"):
        if k in p and prn.get(k) != str(p[k]):
            fails.append(f"{k}={prn.get(k)!r}")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("ppt_master_background")
def _chk_ppt_bg(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        return False, "COM required for ppt_master_background"
    fails = []
    if "master_fill_type" in p and int(ctx.doc.SlideMaster.Background.Fill.Type) != int(p["master_fill_type"]):
        fails.append(f"MasterFill={ctx.doc.SlideMaster.Background.Fill.Type}")
    if p.get("all_follow_master"):
        not_following = [i for i in range(1, int(ctx.doc.Slides.Count) + 1) if int(ctx.doc.Slides(i).FollowMasterBackground) != -1]
        if not_following:
            fails.append(f"notFollow={not_following}")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("ppt_slide_transition")
def _chk_ppt_transition(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        return False, "COM required for ppt_slide_transition"
    total = int(ctx.doc.Slides.Count)
    targets = list(range(1, total + 1)) if p.get("slides", "all") == "all" else [int(x) for x in p["slides"]]
    fails = []
    for si in targets:
        tr = ctx.doc.Slides(si).SlideShowTransition
        if "entry_effect" in p and int(tr.EntryEffect) != int(p["entry_effect"]):
            fails.append(f"S{si}.eff={tr.EntryEffect}")
        if "advance_on_time" in p and int(tr.AdvanceOnTime) != (-1 if p["advance_on_time"] else 0):
            fails.append(f"S{si}.adv={tr.AdvanceOnTime}")
        if "advance_time_range" in p:
            lo, hi = p["advance_time_range"]
            if not (lo <= float(tr.AdvanceTime) <= hi):
                fails.append(f"S{si}.time={tr.AdvanceTime}")
        if "hidden" in p and int(tr.Hidden) != (-1 if p["hidden"] else 0):
            fails.append(f"S{si}.hidden={tr.Hidden}")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("ppt_shape_match")
def _chk_ppt_shape(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    si = int(p["slide"])
    slides_xml = ctx.xml.get("slides", [])
    s_xml = slides_xml[si - 1] if si <= len(slides_xml) else {}
    fails = []

    if ctx.doc is not None:
        sl = ctx.doc.Slides(si)
        W, H = float(ctx.doc.PageSetup.SlideWidth), float(ctx.doc.PageSetup.SlideHeight)
        matched = None
        for j in range(1, int(sl.Shapes.Count) + 1):
            sp = sl.Shapes(j)
            if "type" in p and int(sp.Type) != int(p["type"]):
                continue
            if "name" in p and str(sp.Name) != str(p["name"]):
                continue
            if p.get("has_smartart") and not sp.HasSmartArt:
                continue
            if p.get("has_chart") and not sp.HasChart:
                continue
            if p.get("has_table") and not sp.HasTable:
                continue
            if "text_contains" in p:
                txt = str(sp.TextFrame.TextRange.Text) if (sp.HasTextFrame and sp.TextFrame.HasText) else ""
                if p["text_contains"] not in txt:
                    continue
            matched = sp
            break
        if matched is None:
            return False, f"shape not found on slide {si}"
        if p.get("right_aligned_to_slide") and not _close(matched.Left + matched.Width, W, 1.0):
            fails.append(f"R={matched.Left + matched.Width:.1f}/{W}")
        if p.get("top_aligned_to_slide") and not _close(matched.Top, 0.0, 1.0):
            fails.append(f"T={matched.Top:.1f}")
        if p.get("centered_h") and not _close(matched.Left, (W - matched.Width) / 2.0, 1.0):
            fails.append(f"L={matched.Left:.1f}")
        if p.get("centered_v") and not _close(matched.Top, (H - matched.Height) / 2.0, 1.0):
            fails.append(f"T={matched.Top:.1f}")
        if "align" in p and int(matched.TextFrame.TextRange.ParagraphFormat.Alignment) != int(p["align"]):
            fails.append(f"align={matched.TextFrame.TextRange.ParagraphFormat.Alignment}")
        if "smallcaps" in p and int(matched.TextFrame2.TextRange.Font.Smallcaps) != (-1 if p["smallcaps"] else 0):
            fails.append(f"smallcaps={matched.TextFrame2.TextRange.Font.Smallcaps}")
        if "font_fill_pattern" in p:
            f2 = matched.TextFrame2.TextRange.Font
            if int(f2.Fill.Type) != 2 or int(f2.Fill.Pattern) != int(p["font_fill_pattern"]):
                fails.append(f"pattern={f2.Fill.Type}/{getattr(f2.Fill, 'Pattern', None)}")
        if p.get("shadow_visible") and int(matched.TextFrame2.TextRange.Font.Shadow.Visible) != -1:
            fails.append("shadow_not_visible")
        if "smartart_nodes" in p:
            nodes = [str(n.TextFrame2.TextRange.Text).strip() for n in matched.SmartArt.AllNodes]
            if nodes != list(p["smartart_nodes"]):
                fails.append(f"smartartNodes={nodes}")

    if "body_pr_anchor" in p or "body_pr_anchor_ctr" in p:
        sub = str(p.get("text_contains", p.get("name", "")))
        bp_match = next(
            (x for x in s_xml.get("shapes_body_pr", []) if sub in (x.get("text") or "") or sub == (x.get("name") or "")),
            None,
        )
        if bp_match is None:
            fails.append("xml_bodyPr_shape_not_found")
        else:
            if "body_pr_anchor" in p and bp_match.get("anchor") != p["body_pr_anchor"]:
                fails.append(f"xml_anchor={bp_match.get('anchor')!r}")
            if "body_pr_anchor_ctr" in p and bp_match.get("anchorCtr") != str(p["body_pr_anchor_ctr"]):
                fails.append(f"xml_anchorCtr={bp_match.get('anchorCtr')!r}")

    if p.get("hmerge_all_aligned") and int(s_xml.get("hmerge_missing_algn", 0)) > 0:
        fails.append(f"hmerge_missing_algn={s_xml.get('hmerge_missing_algn')}")

    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("ppt_align_distribute")
def _chk_ppt_align_dist(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is None:
        return False, "COM required for ppt_align_distribute"
    sl = ctx.doc.Slides(int(p["slide"]))
    tol = float(p.get("tol_pt", 0.8))
    sps = [
        sl.Shapes(i)
        for i in range(1, int(sl.Shapes.Count) + 1)
        if ("shape_type" not in p or int(sl.Shapes(i).Type) == int(p["shape_type"]))
    ]
    if len(sps) < 2:
        return False, f"need >=2 shapes, got {len(sps)}"
    sps.sort(key=lambda s: float(s.Left))
    fails = []
    if p.get("top_aligned"):
        t0 = float(sps[0].Top)
        if not all(abs(float(s.Top) - t0) <= tol for s in sps):
            fails.append(f"tops={[round(float(s.Top), 1) for s in sps]}")
    if p.get("left_aligned"):
        l0 = float(sps[0].Left)
        if not all(abs(float(s.Left) - l0) <= tol for s in sps):
            fails.append(f"lefts={[round(float(s.Left), 1) for s in sps]}")
    if p.get("h_distributed") and len(sps) >= 3:
        gaps = [float(sps[i + 1].Left) - (float(sps[i].Left) + float(sps[i].Width)) for i in range(len(sps) - 1)]
        if not all(abs(g - gaps[0]) <= tol for g in gaps):
            fails.append(f"gaps={[round(g, 2) for g in gaps]}")
    return not fails, ("OK" if not fails else ", ".join(fails))


@register_check("ppt_chart")
def _chk_ppt_chart(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    idx = int(p.get("chart_index", 1))
    charts = ctx.xml.get("charts", [])
    if idx > len(charts):
        return False, f"ppt chart{idx}.xml missing"
    cxml = charts[idx - 1]
    fails = []
    if "series_count" in p and int(cxml["series_count"]) != int(p["series_count"]):
        fails.append(f"series={cxml['series_count']}")
    if "pt_count" in p and cxml["series"]:
        if int(cxml["series"][0]["val_pt_count"] or 0) != int(p["pt_count"]):
            fails.append(f"ptCount={cxml['series'][0]['val_pt_count']}")
    if ctx.doc is not None and "slide" in p:
        sl = ctx.doc.Slides(int(p["slide"]))
        ch_sps = [sl.Shapes(i) for i in range(1, int(sl.Shapes.Count) + 1) if sl.Shapes(i).HasChart]
        if not ch_sps:
            fails.append("no_chart_on_slide")
        else:
            ch = ch_sps[0].Chart
            if "chart_type" in p and int(ch.ChartType) != int(p["chart_type"]):
                fails.append(f"chartType={ch.ChartType}")
            if "has_title" in p and bool(ch.HasTitle) != bool(p["has_title"]):
                fails.append(f"hasTitle={ch.HasTitle}")
            if p.get("no_gridlines"):
                if ch.Axes(1).HasMajorGridlines or ch.Axes(2).HasMajorGridlines:
                    fails.append("has_gridlines")
    return not fails, ("OK" if not fails else ", ".join(fails))


# ---------------------------------------------------------------------------
# Grader Engine, Rubric Scaffolder & Batch Runner
# ---------------------------------------------------------------------------

def grade_with_rubric(
    candidate_path: str,
    rubric: dict[str, Any],
    *,
    use_com: bool = True,
) -> dict[str, Any]:
    """Evaluate a candidate file against a weighted JSON rubric."""
    t0 = time.time()
    info = inspect_container(candidate_path)
    kind = rubric.get("kind") or info["kind"]
    xml_data = audit_ooxml(info["path"]) if info["container"] == "ooxml" else {"ooxml_supported": False}
    rules = rubric.get("rules", [])
    results: list[dict[str, Any]] = []

    def _run_rules(app: Any, doc: Any) -> None:
        ctx = GradeContext(file_path=info["path"], kind=kind, xml=xml_data, app=app, doc=doc)
        if doc is not None and kind == "word":
            try:
                doc.Repaginate()
            except Exception:
                pass
        for rule in rules:
            rtype = rule["type"]
            pts = float(rule.get("points", 1.0))
            handler = CHECK_HANDLERS.get(rtype)
            if handler is None:
                ok, detail = False, f"Unknown rule type: {rtype}"
            else:
                try:
                    ok, detail = handler(ctx, rule.get("params", {}))
                except Exception as exc:
                    ok, detail = False, f"Exception: {exc}"
            results.append({
                "id": rule.get("id", ""),
                "name": rule.get("name", rtype),
                "type": rtype,
                "passed": bool(ok),
                "earned": pts if ok else 0.0,
                "points": pts,
                "detail": detail,
            })

    if use_com:
        with wps_session(kind, info["path"], read_only=True, with_window=False) as (app, doc, _):
            _run_rules(app, doc)
    else:
        _run_rules(None, None)

    earned = sum(r["earned"] for r in results)
    max_pts = sum(r["points"] for r in results)
    return {
        "title": rubric.get("title", Path(candidate_path).name),
        "file": info["path"],
        "kind": kind,
        "elapsed_sec": round(time.time() - t0, 3),
        "earned_points": round(earned, 2),
        "max_points": round(max_pts, 2),
        "score_pct": round(100.0 * earned / max_pts, 1) if max_pts else 0.0,
        "all_passed": bool(results) and all(r["passed"] for r in results),
        "results": results,
    }


def scaffold_rubric(reference_path: str, total_points: float = 20.0) -> dict[str, Any]:
    """Generate a starter JSON rubric by inspecting a reference answer file."""
    info = inspect_container(reference_path)
    kind = info["kind"]
    xml_data = audit_ooxml(info["path"]) if info["container"] == "ooxml" else {}
    rules: list[dict[str, Any]] = []

    with wps_session(kind, info["path"], read_only=True, with_window=False) as (app, doc, _):
        if kind == "word":
            ps = doc.Sections(1).PageSetup
            rules.append({
                "id": "w_page",
                "name": "页面设置（边距/方向/节数）",
                "type": "word_page_setup",
                "params": {
                    "section": 1,
                    "sections_count": int(doc.Sections.Count),
                    "orientation": int(ps.Orientation),
                    "margins_cm": [round(float(x) / PT_PER_CM, 2) for x in (ps.TopMargin, ps.BottomMargin, ps.LeftMargin, ps.RightMargin)],
                },
            })
            if int(doc.Tables.Count) > 0:
                t = doc.Tables(1)
                rules.append({
                    "id": "w_tbl",
                    "name": "表格行列与对齐",
                    "type": "word_table",
                    "params": {
                        "table_index": 1,
                        "rows": int(t.Rows.Count),
                        "cols": int(t.Columns.Count),
                        "align": int(t.Rows.Alignment),
                    },
                })
            for idx, c in enumerate(xml_data.get("charts", []), 1):
                rules.append({
                    "id": f"w_chart_{idx}",
                    "name": f"内嵌图表{idx}数据点数",
                    "type": "word_chart",
                    "params": {
                        "chart_index": idx,
                        "series_count": c["series_count"],
                        "pt_count": c["series"][0]["val_pt_count"] if c["series"] else None,
                    },
                })

        elif kind == "excel":
            snames = [s["name"] for s in xml_data.get("sheets", [])]
            rules.append({"id": "e_sheets", "name": "工作表结构", "type": "excel_sheets", "params": {"exact_names": snames}})
            for ws_x in xml_data.get("worksheets", []):
                sn = ws_x["name"]
                for idx, cf in enumerate(ws_x.get("conditional_formatting", []), 1):
                    fml = cf["rules"][0]["formulas"][0] if cf.get("rules") and cf["rules"][0].get("formulas") else ""
                    if fml:
                        rules.append({
                            "id": f"e_cf_{sn}_{idx}",
                            "name": f"[{sn}] 条件格式规则",
                            "type": "excel_conditional_formatting",
                            "params": {"sheet": sn, "sqref": cf["sqref"], "formula_norm": fml},
                        })
                if ws_x.get("sort_state"):
                    rules.append({
                        "id": f"e_sort_{sn}",
                        "name": f"[{sn}] 多字段排序状态",
                        "type": "excel_sort_state",
                        "params": {"sheet": sn, "ref": ws_x["sort_state"]["ref"]},
                    })

        elif kind == "ppt":
            ps = doc.PageSetup
            rules.append({
                "id": "p_setup",
                "name": "幻灯片尺寸与总页数",
                "type": "ppt_page_setup",
                "params": {
                    "slide_size": int(ps.SlideSize),
                    "width_pt": round(float(ps.SlideWidth), 1),
                    "height_pt": round(float(ps.SlideHeight), 1),
                    "slides_count": int(doc.Slides.Count),
                },
            })
            if xml_data.get("prnPr"):
                rules.append({
                    "id": "p_print",
                    "name": "打印选项配置",
                    "type": "ppt_print_options",
                    "params": dict(xml_data["prnPr"]),
                })
            tr1 = doc.Slides(1).SlideShowTransition
            if int(tr1.EntryEffect) != 0:
                rules.append({
                    "id": "p_trans",
                    "name": "幻灯片切换与计时",
                    "type": "ppt_slide_transition",
                    "params": {
                        "slides": "all",
                        "entry_effect": int(tr1.EntryEffect),
                        "advance_on_time": int(tr1.AdvanceOnTime) == -1,
                    },
                })
            for idx, c in enumerate(xml_data.get("charts", []), 1):
                rules.append({
                    "id": f"p_chart_{idx}",
                    "name": f"图表{idx}系列与数据点数",
                    "type": "ppt_chart",
                    "params": {
                        "chart_index": idx,
                        "series_count": c["series_count"],
                        "pt_count": c["series"][0]["val_pt_count"] if c["series"] else None,
                    },
                })

    if rules:
        per = round(total_points / len(rules), 2)
        for r in rules[:-1]:
            r["points"] = per
        rules[-1]["points"] = round(total_points - per * (len(rules) - 1), 2)

    return {
        "title": Path(reference_path).stem,
        "kind": kind,
        "rules": rules,
    }


def batch_grade(candidate_dir: str, ref_or_rubric_dir: str, *, use_com: bool = True) -> dict[str, Any]:
    """Batch-grade all matching exam files against rubrics (*.rubric.json) or reference files."""
    t0 = time.time()
    c_root = Path(os.path.abspath(candidate_dir))
    r_root = Path(os.path.abspath(ref_or_rubric_dir))

    cand_files = {
        p.relative_to(c_root): p
        for p in sorted(c_root.rglob("*"))
        if p.is_file() and p.suffix.lower() in (".docx", ".wps", ".xlsx", ".et", ".pptx", ".dps") and not p.name.startswith("~$")
    }
    ref_files = [p for p in sorted(r_root.rglob("*")) if p.is_file() and not p.name.startswith("~$")]

    reports: list[dict[str, Any]] = []
    # Match by relative path, stem+suffix, or rubric json stem
    for rel, cpath in cand_files.items():
        rubric_match = next(
            (rf for rf in ref_files if rf.name == f"{cpath.stem}.rubric.json" or rf.name == f"{cpath.name}.rubric.json"),
            None,
        )
        if rubric_match:
            rubric = json.loads(rubric_match.read_text(encoding="utf-8"))
            rep = grade_with_rubric(str(cpath), rubric, use_com=use_com)
            reports.append({"mode": "rubric", "file": str(rel), **rep})
            continue

        ref_match = next(
            (rf for rf in ref_files if rf.name == cpath.name and rf.parent.name.replace("参考答案", "") == cpath.parent.name.replace("作答文件", "")),
            None,
        ) or next((rf for rf in ref_files if rf.name == cpath.name), None)
        if ref_match:
            suite = compare_documents(str(cpath), str(ref_match), use_com=use_com)
            d = suite.to_dict()
            reports.append({
                "mode": "ref_diff",
                "file": str(rel),
                "earned_points": float(d["passed"]),
                "max_points": float(d["total"]),
                "score_pct": d["score_pct"],
                "all_passed": d["all_passed"],
                "details": d["checks"],
            })

    total_earned = sum(r["earned_points"] for r in reports)
    total_max = sum(r["max_points"] for r in reports)
    return {
        "candidate_dir": str(c_root),
        "reference_dir": str(r_root),
        "elapsed_sec": round(time.time() - t0, 3),
        "matched_files": len(reports),
        "earned_points": round(total_earned, 2),
        "max_points": round(total_max, 2),
        "score_pct": round(100.0 * total_earned / total_max, 1) if total_max else 0.0,
        "all_passed": bool(reports) and all(r["all_passed"] for r in reports),
        "reports": reports,
    }


def _print_grade_report(rep: dict[str, Any]) -> None:
    print(f"=== RUBRIC GRADE ({rep['kind'].upper()}): {rep['title']} ===")
    for r in rep["results"]:
        tag = "PASS" if r["passed"] else "FAIL"
        suffix = f" ({r['detail']})" if r["detail"] else ""
        print(f"[{tag}] [{r['earned']}/{r['points']}分] {r['name']}{suffix}")
    status = "ALL PASS" if rep["all_passed"] else "HAS FAIL"
    print(
        f"DONE in {rep['elapsed_sec']:.2f}s | {status} | "
        f"{rep['earned_points']}/{rep['max_points']} 分 ({rep['score_pct']}%)"
    )


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="Weighted JSON Rubric Grader & Batch Exam Scorer.")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_grade = sub.add_parser("grade", help="Grade one candidate file with a JSON rubric.")
    p_grade.add_argument("candidate")
    p_grade.add_argument("rubric")
    p_grade.add_argument("--no-com", action="store_true")
    p_grade.add_argument("--json", action="store_true")

    p_init = sub.add_parser("init", help="Scaffold a starter JSON rubric from a reference answer.")
    p_init.add_argument("reference")
    p_init.add_argument("-o", "--output", required=True)
    p_init.add_argument("--total-points", type=float, default=20.0)

    p_batch = sub.add_parser("batch", help="Batch-grade a candidate folder against rubrics or references.")
    p_batch.add_argument("candidate_dir")
    p_batch.add_argument("ref_or_rubric_dir")
    p_batch.add_argument("--no-com", action="store_true")
    p_batch.add_argument("--json", action="store_true")

    args = parser.parse_args()
    if args.cmd == "grade":
        rubric = json.loads(Path(args.rubric).read_text(encoding="utf-8"))
        rep = grade_with_rubric(args.candidate, rubric, use_com=not args.no_com)
        if args.json:
            print(json.dumps(rep, ensure_ascii=False, indent=2))
        else:
            _print_grade_report(rep)
        sys.exit(0 if rep["all_passed"] else 1)

    if args.cmd == "init":
        rubric = scaffold_rubric(args.reference, total_points=args.total_points)
        Path(args.output).write_text(json.dumps(rubric, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Rubric scaffolded -> {args.output} ({len(rubric['rules'])} rules)")
        sys.exit(0)

    if args.cmd == "batch":
        rep = batch_grade(args.candidate_dir, args.ref_or_rubric_dir, use_com=not args.no_com)
        if args.json:
            print(json.dumps(rep, ensure_ascii=False, indent=2))
        else:
            print(f"=== BATCH SCOREBOARD ({rep['matched_files']} files) ===")
            for item in rep["reports"]:
                tag = "PASS" if item["all_passed"] else "FAIL"
                print(f"[{tag}] {item['file']} ({item['mode']}) -> {item['earned_points']}/{item['max_points']} ({item['score_pct']}%)")
            print(f"TOTAL: {rep['earned_points']}/{rep['max_points']} ({rep['score_pct']}%) in {rep['elapsed_sec']:.2f}s")
        sys.exit(0 if rep["all_passed"] else 1)


if __name__ == "__main__":
    main()
