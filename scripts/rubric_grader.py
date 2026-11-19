"""Weighted JSON Rubric Grader, Rubric Scaffolder, and Batch Exam Scorer."""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

try:
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None  # type: ignore

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from ooxml_patch import PT_PER_CM, audit_ooxml, read_excel_static
from ref_diff import _norm_formula, _vals_equal, compare_documents
from wps_env import WPSBatchSession, infer_kind, inspect_container, wps_batch_session, wps_session

SUPPORTED_EXAM_EXTS = (
    ".docx", ".doc", ".wps",
    ".xlsx", ".xls", ".et",
    ".pptx", ".ppt", ".dps",
)


@dataclass
class GradeContext:
    file_path: str
    kind: str
    xml: dict[str, Any]
    app: Any = None
    doc: Any = None
    _excel_static_cache: dict[str, Any] | None = None

    @property
    def excel_static(self) -> dict[str, Any]:
        if self._excel_static_cache is None:
            if self.kind == "excel" and self.xml.get("ooxml_supported", True):
                self._excel_static_cache = read_excel_static(self.file_path)
            else:
                self._excel_static_cache = {}
        return self._excel_static_cache


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


def _col_letters_to_idx(col_str: str) -> int:
    res = 0
    for ch in col_str.upper():
        res = res * 26 + (ord(ch) - ord("A") + 1)
    return res


def _parse_cell_ref(cell_ref: str) -> tuple[int, int]:
    cleaned = cell_ref.strip().split("!")[-1].replace("$", "")
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", cleaned)
    if not m:
        raise ValueError(f"Invalid Excel cell reference: {cell_ref!r}")
    return int(m.group(2)), _col_letters_to_idx(m.group(1))


def _parse_cell_range(range_str: str) -> tuple[int, int, int, int]:
    cleaned = range_str.strip().split("!")[-1]
    if ":" in cleaned:
        left, right = cleaned.split(":", 1)
        r1, c1 = _parse_cell_ref(left)
        r2, c2 = _parse_cell_ref(right)
        return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)
    r, c = _parse_cell_ref(cleaned)
    return r, c, r, c


def _require_ooxml_or_com(ctx: GradeContext, rule_name: str) -> str | None:
    if not ctx.xml.get("ooxml_supported", True):
        return f"COM required for non-OOXML file ({rule_name})"
    return None


# ---------------------------------------------------------------------------
# Word Check Handlers
# ---------------------------------------------------------------------------

@register_check("word_page_setup")
def _chk_word_page_setup(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    sec_idx = int(p.get("section", 1))
    tol_cm = float(p.get("tol_cm", 0.05))

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

    err = _require_ooxml_or_com(ctx, "word_page_setup")
    if err:
        return False, err

    secs = ctx.xml.get("sections", [])
    if "sections_count" in p and len(secs) != int(p["sections_count"]):
        return False, f"xml sections_count={len(secs)} != {p['sections_count']}"
    if sec_idx > len(secs):
        return False, f"xml section {sec_idx} missing"
    sec = secs[sec_idx - 1]
    fails = []
    if "orientation" in p and int(sec.get("orientation", 0)) != int(p["orientation"]):
        fails.append(f"orient={sec.get('orientation')}")
    if "section_start" in p and int(sec.get("section_start", 2)) != int(p["section_start"]):
        fails.append(f"secStart={sec.get('section_start')}")
    if "gutter_pos" in p and int(sec.get("gutter_pos", 0)) != int(p["gutter_pos"]):
        fails.append(f"gutterPos={sec.get('gutter_pos')}")
    if "page_width_cm" in p:
        pw = sec.get("page_width_cm")
        if pw is None or abs(float(pw) - float(p["page_width_cm"])) > tol_cm:
            fails.append(f"pageW={pw}")
    if "page_height_cm" in p:
        ph = sec.get("page_height_cm")
        if ph is None or abs(float(ph) - float(p["page_height_cm"])) > tol_cm:
            fails.append(f"pageH={ph}")
    if "margins_cm" in p:
        m = p["margins_cm"]
        exp = [m["top"], m["bottom"], m["left"], m["right"]] if isinstance(m, dict) else list(m)
        got = sec.get("margins_cm")
        if not got or not all(abs(float(g) - float(e)) <= tol_cm for g, e in zip(got, exp)):
            fails.append(f"margins_cm={got}")
    return not fails, ("OK (XML)" if not fails else f"xml failed: {', '.join(fails)}")


@register_check("word_style")
def _chk_word_style(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is not None:
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

    err = _require_ooxml_or_com(ctx, "word_style")
    if err:
        return False, err

    styles_bundle = ctx.xml.get("styles", {})
    by_id = styles_bundle.get("by_id", {})
    aliases = styles_bundle.get("aliases", {})
    key = str(p["style_id"]) if p.get("style_id") is not None else str(p.get("style_name", ""))
    resolved_id = aliases.get(key.lower(), key)
    st_xml = by_id.get(resolved_id)
    if st_xml is None:
        return False, f"xml style {key!r} not found"

    fails = []
    if "font_size" in p:
        fsz = st_xml.get("font_size")
        if fsz is None or not _close(fsz, p["font_size"], 0.1):
            fails.append(f"size={fsz}")
    if "font_name" in p and p["font_name"] not in st_xml.get("font_names", []):
        fails.append(f"font={st_xml.get('font_names')}")
    if "bold" in p and bool(st_xml.get("bold")) != bool(p["bold"]):
        fails.append(f"bold={st_xml.get('bold')}")
    if "align" in p:
        al = st_xml.get("align")
        if al is None or int(al) != int(p["align"]):
            fails.append(f"align={al}")
    if "space_before" in p:
        sb = st_xml.get("space_before")
        if sb is None or not _close(sb, p["space_before"], 0.2):
            fails.append(f"before={sb}")
    if "space_after" in p:
        sa = st_xml.get("space_after")
        if sa is None or not _close(sa, p["space_after"], 0.2):
            fails.append(f"after={sa}")
    if "line_unit_before" in p:
        ub = st_xml.get("line_unit_before")
        if ub is None or not _close(ub, p["line_unit_before"], 0.05):
            fails.append(f"unitBefore={ub}")
    if "line_unit_after" in p:
        ua = st_xml.get("line_unit_after")
        if ua is None or not _close(ua, p["line_unit_after"], 0.05):
            fails.append(f"unitAfter={ua}")
    if "line_spacing_rule" in p:
        lr = st_xml.get("line_spacing_rule")
        if lr is None or int(lr) != int(p["line_spacing_rule"]):
            fails.append(f"lineRule={lr}")
    if "bottom_border" in p:
        bb = p["bottom_border"]
        b_xml = st_xml.get("bottom_border") or {}
        if bb.get("has_border") and not b_xml.get("has_border"):
            fails.append("no_bottom_border")
        if "line_width" in bb and int(b_xml.get("line_width", 0)) != int(bb["line_width"]):
            fails.append(f"bdrW={b_xml.get('line_width')}")
    return not fails, ("OK (XML)" if not fails else f"xml: {', '.join(fails)}")


@register_check("word_paragraph_match")
def _chk_word_para(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is not None:
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

    err = _require_ooxml_or_com(ctx, "word_paragraph_match")
    if err:
        return False, err

    paras = ctx.xml.get("paragraphs", [])
    target_x = None
    if "index" in p:
        idx = int(p["index"])
        if 1 <= idx <= len(paras):
            target_x = paras[idx - 1]
    else:
        sub = str(p.get("text_contains", ""))
        target_x = next((cand for cand in paras if sub in cand.get("full_text", cand.get("text", ""))), None)
    if target_x is None:
        return False, "xml target paragraph not found"

    fails = []
    if "style_name" in p:
        sn = str(p["style_name"]).lower()
        got_s = f"{target_x.get('style', '')} {target_x.get('style_id', '')}".lower()
        alias_sid = str(ctx.xml.get("styles", {}).get("aliases", {}).get(sn, "")).lower()
        if sn not in got_s and (not alias_sid or alias_sid != str(target_x.get("style_id", "")).lower()):
            fails.append(f"style={target_x.get('style')}/{target_x.get('style_id')}")
    if "list_string" in p and str(p["list_string"]):
        if target_x.get("num_id") is None and not str(target_x.get("full_text", "")).startswith(str(p["list_string"])):
            fails.append("xml_missing_numPr")
    if "align" in p:
        al = target_x.get("align")
        if al is None or int(al) != int(p["align"]):
            fails.append(f"align={al}")
    if "font_size" in p:
        fsz = target_x.get("font_size")
        if fsz is None or not _close(fsz, p["font_size"], 0.1):
            fails.append(f"fontSize={fsz}")
    if "c_indent" in p:
        ci = target_x.get("c_indent")
        if ci is None or not _close(ci, p["c_indent"], 0.05):
            fails.append(f"cIndent={ci}")
    if "line_spacing_rule" in p:
        lr = target_x.get("line_rule")
        if lr is None or int(lr) != int(p["line_spacing_rule"]):
            fails.append(f"lineRule={lr}")
    if "line_spacing" in p:
        ls = target_x.get("line_spacing")
        if ls is None or not _close(ls, p["line_spacing"], 0.2):
            fails.append(f"lineSpacing={ls}")
    if "unit_before" in p:
        ub = target_x.get("unit_before")
        if ub is None or not _close(ub, p["unit_before"], 0.05):
            fails.append(f"unitBefore={ub}")
    if "unit_after" in p:
        ua = target_x.get("unit_after")
        if ua is None or not _close(ua, p["unit_after"], 0.05):
            fails.append(f"unitAfter={ua}")
    return not fails, ("OK (XML)" if not fails else f"xml: {', '.join(fails)}")


@register_check("word_text_absent")
def _chk_word_text_absent(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    forbidden = list(p.get("forbidden", []))
    if ctx.doc is not None:
        txt = str(ctx.doc.Content.Text)
        found = [s for s in forbidden if s in txt]
        return not found, ("OK" if not found else f"found forbidden={found!r}")

    err = _require_ooxml_or_com(ctx, "word_text_absent")
    if err:
        return False, err
    txt = str(ctx.xml.get("full_text", ""))
    found = [s for s in forbidden if s in txt]
    return not found, ("OK (XML)" if not found else f"xml found forbidden={found!r}")


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
    else:
        err = _require_ooxml_or_com(ctx, "word_header_watermark")
        if err:
            return False, err
        secs_xml = ctx.xml.get("sections", [])
        if sec_idx > len(secs_xml):
            return False, f"xml section {sec_idx} missing"
        sec = secs_xml[sec_idx - 1]
        hdrs = sec.get("headers", [])
        if "link_to_previous" in p and bool(sec.get("header_link_to_previous")) != bool(p["link_to_previous"]):
            fails.append(f"xml LinkToPrevious={sec.get('header_link_to_previous')}")
        if "has_field_substr" in p:
            all_instr = [ins for h in hdrs for ins in h.get("instr_text", [])]
            if not any(p["has_field_substr"].upper() in c.upper() for c in all_instr):
                fails.append(f"xml fields={all_instr}")
        if "shape_count" in p:
            cnt = sum(int(h.get("shape_count", h.get("drawings", 0))) for h in hdrs)
            if cnt != int(p["shape_count"]):
                fails.append(f"xml shapeCount={cnt}")
        if "watermark_text" in p:
            all_wm = [t for h in hdrs for t in h.get("wm_texts", [])]
            if not any(p["watermark_text"] in t for t in all_wm):
                fails.append("xml text_watermark_missing")
        if p.get("picture_no_washout") and any(h.get("picture_washout") for h in hdrs):
            fails.append("xml picture_still_washout")
        if "bottom_border" in p and p["bottom_border"].get("has_border"):
            if not any(h.get("has_bottom_border") for h in hdrs):
                fails.append("xml header_no_bottom_border")

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
            if lvl_restart is not None:
                raw_lr = lv.get("lvlRestart")
                effective_lr = str(raw_lr) if raw_lr is not None else (str(ilvl) if ilvl > 0 else "0")
                if effective_lr != str(lvl_restart):
                    continue
            return True, f"matched {nf}/{lv.get('lvlText')}"
    return False, f"no numbering level matching ilvl={ilvl}, fmt={num_fmt_prefix}, text={lvl_text}, restart={lvl_restart}"


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
    else:
        err = _require_ooxml_or_com(ctx, "word_table")
        if err:
            return False, err
        tables = ctx.xml.get("tables", [])
        if t_idx > len(tables):
            return False, f"xml table {t_idx} missing"
        tx = tables[t_idx - 1]
        if "rows" in p and int(tx["rows"]) != int(p["rows"]):
            fails.append(f"rows={tx['rows']}")
        if "cols" in p and int(tx["cols"]) != int(p["cols"]):
            fails.append(f"cols={tx['cols']}")
        if "align" in p and int(tx["align"]) != int(p["align"]):
            fails.append(f"align={tx['align']}")
        if "heading_repeat" in p and bool(tx["heading_repeat"]) != bool(p["heading_repeat"]):
            fails.append(f"headingRepeat={tx['heading_repeat']}")
        for cc in p.get("cell_checks", []):
            coord = f"{int(cc['row'])},{int(cc['col'])}"
            cell = tx["cells"].get(coord)
            if cell is None:
                fails.append(f"cell({coord}) missing")
                continue
            if "text_contains" in cc and cc["text_contains"] not in cell.get("text", ""):
                fails.append(f"cell({coord}) text={cell.get('text')!r}")
            if "bg_color_bgr" in cc and cell.get("bg_color_bgr") != int(cc["bg_color_bgr"]):
                fails.append(f"cell({coord}) bg={cell.get('bg_color_bgr')}")
            if "font_name" in cc and cell.get("font_name") != cc["font_name"]:
                fails.append(f"cell({coord}) font={cell.get('font_name')}")

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
    elif "chart_type" in p and cxml.get("chart_type_code") is not None:
        if int(cxml["chart_type_code"]) != int(p["chart_type"]):
            fails.append(f"xmlChartType={cxml['chart_type_code']}")
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
    sub = _norm_formula(p.get("formula_contains", ""))
    exact = _norm_formula(p.get("formula_exact", ""))

    if ctx.doc is not None:
        ws = ctx.doc.Sheets(p["sheet"])
        raw = ws.Range(p["range"]).Formula
        grid = raw if isinstance(raw, tuple) else ((raw,),)
        if grid and not isinstance(grid[0], tuple):
            grid = (grid,)
        bad = []
        for ri, row in enumerate(grid, 1):
            for ci, cell_f in enumerate(row, 1):
                nf = _norm_formula(cell_f)
                if exact and nf != exact:
                    bad.append((ri, ci, nf))
                elif sub and sub not in nf:
                    bad.append((ri, ci, nf))
        return not bad, ("OK" if not bad else f"mismatches={len(bad)} sample={bad[:2]}")

    err = _require_ooxml_or_com(ctx, "excel_range_formula")
    if err:
        return False, err
    st = ctx.excel_static
    ws_st = st.get("sheets", {}).get(p["sheet"])
    if ws_st is None:
        return False, f"sheet {p['sheet']!r} not found"
    r1, c1, r2, c2 = _parse_cell_range(p["range"])
    bad = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            nf = _norm_formula(ws_st["formulas"].get((r, c), ""))
            if exact and nf != exact:
                bad.append((r - r1 + 1, c - c1 + 1, nf))
            elif sub and sub not in nf:
                bad.append((r - r1 + 1, c - c1 + 1, nf))
    return not bad, ("OK (static)" if not bad else f"mismatches={len(bad)} sample={bad[:2]}")


@register_check("excel_range_values")
def _chk_excel_values(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    tol = float(p.get("tol", 1e-4))
    if ctx.doc is not None:
        ws = ctx.doc.Sheets(p["sheet"])
        bad = []
        for addr, exp in p.get("cells", {}).items():
            got = ws.Range(addr).Value
            if not _vals_equal(got, exp, tol=tol):
                bad.append(f"{addr}: {got!r} != {exp!r}")
        return not bad, ("OK" if not bad else "; ".join(bad[:5]))

    err = _require_ooxml_or_com(ctx, "excel_range_values")
    if err:
        return False, err
    st = ctx.excel_static
    ws_st = st.get("sheets", {}).get(p["sheet"])
    if ws_st is None:
        return False, f"sheet {p['sheet']!r} not found"
    bad = []
    for addr, exp in p.get("cells", {}).items():
        coord = _parse_cell_ref(addr)
        got = ws_st["values"].get(coord)
        fml = ws_st["formulas"].get(coord)
        if got is None and isinstance(fml, str) and fml.startswith("=") and exp is not None:
            bad.append(f"{addr}: cached value missing (None) for {fml!r}; run with WPS COM")
        elif not _vals_equal(got, exp, tol=tol):
            bad.append(f"{addr}: {got!r} != {exp!r}")
    return not bad, ("OK (static)" if not bad else "; ".join(bad[:5]))


@register_check("excel_number_format")
def _chk_excel_numfmt(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is not None:
        ws = ctx.doc.Sheets(p["sheet"])
        nf = str(ws.Range(p["range"]).NumberFormat or "")
    else:
        err = _require_ooxml_or_com(ctx, "excel_number_format")
        if err:
            return False, err
        ws_st = ctx.excel_static.get("sheets", {}).get(p["sheet"])
        if ws_st is None:
            return False, f"sheet {p['sheet']!r} not found"
        r1, c1, _, _ = _parse_cell_range(p["range"])
        nf = str(ws_st["number_formats"].get((r1, c1), "General"))

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
    matched_rules: list[dict[str, Any]] = []
    for cf in ws_xml.get("conditional_formatting", []):
        if str(cf.get("sqref", "")).replace("$", "") == exp_sqref:
            for r in cf.get("rules", []):
                if any(_norm_formula(f).lstrip("=") == exp_fml for f in r.get("formulas", [])):
                    xml_ok = True
                    matched_rules.append(r)
    if not xml_ok:
        return False, f"XML CF mismatch: {ws_xml.get('conditional_formatting')}"

    if "interior_color_bgr" in p:
        if ctx.doc is not None:
            fcs = ctx.doc.Sheets(sname).Cells.FormatConditions
            ok_col = any(
                int(fcs(i).Interior.Color) == int(p["interior_color_bgr"])
                for i in range(1, int(fcs.Count) + 1)
            )
            if not ok_col:
                return False, "interior_color_bgr mismatch"
        else:
            xml_cols = [r.get("dxf_color_bgr") for r in matched_rules if r.get("dxf_color_bgr") is not None]
            if xml_cols and int(p["interior_color_bgr"]) not in xml_cols:
                return False, f"xml interior_color_bgr={xml_cols} != {p['interior_color_bgr']}"
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
        err = _require_ooxml_or_com(ctx, "excel_pivot_table")
        if err:
            return False, err
        pts = ctx.xml.get("pivot_tables", [])
        sname = p.get("sheet")
        scoped = [pt for pt in pts if not sname or pt.get("sheet") in (None, sname)]
        if not scoped:
            return False, f"no xml PivotTable on sheet {sname!r}"
        pt = scoped[0]
        fails = []
        if "row_fields" in p and pt.get("row_fields") != list(p["row_fields"]):
            fails.append(f"rowFields={pt.get('row_fields')}")
        if "data_fields_count" in p and len(pt.get("data_fields", [])) != int(p["data_fields_count"]):
            fails.append(f"dataFields={pt.get('data_fields')}")
        return not fails, (f"OK (XML): {pt.get('name')}" if not fails else f"xml: {', '.join(fails)}")

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
    else:
        err = _require_ooxml_or_com(ctx, "ppt_page_setup")
        if err:
            return False, err
        if "slide_size" in p and int(ctx.xml.get("slide_size", 0)) != int(p["slide_size"]):
            fails.append(f"xmlSlideSize={ctx.xml.get('slide_size')}")
        if "width_pt" in p:
            w = ctx.xml.get("width_pt")
            if w is None or not _close(w, p["width_pt"], 0.5):
                fails.append(f"xmlW={w}")
        if "height_pt" in p:
            h = ctx.xml.get("height_pt")
            if h is None or not _close(h, p["height_pt"], 0.5):
                fails.append(f"xmlH={h}")

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
    fails = []
    if ctx.doc is not None:
        if "master_fill_type" in p and int(ctx.doc.SlideMaster.Background.Fill.Type) != int(p["master_fill_type"]):
            fails.append(f"MasterFill={ctx.doc.SlideMaster.Background.Fill.Type}")
        if p.get("all_follow_master"):
            not_following = [i for i in range(1, int(ctx.doc.Slides.Count) + 1) if int(ctx.doc.Slides(i).FollowMasterBackground) != -1]
            if not_following:
                fails.append(f"notFollow={not_following}")
        return not fails, ("OK" if not fails else ", ".join(fails))

    err = _require_ooxml_or_com(ctx, "ppt_master_background")
    if err:
        return False, err
    if "master_fill_type" in p:
        mft = ctx.xml.get("master_bg_fill_type")
        if mft is None or int(mft) != int(p["master_fill_type"]):
            fails.append(f"xmlMasterFill={mft}")
    if p.get("all_follow_master"):
        not_following = [
            idx for idx, s in enumerate(ctx.xml.get("slides", []), 1)
            if not s.get("follow_master_bg", True)
        ]
        if not_following:
            fails.append(f"xmlNotFollow={not_following}")
    return not fails, ("OK (XML)" if not fails else ", ".join(fails))


@register_check("ppt_slide_transition")
def _chk_ppt_transition(ctx: GradeContext, p: dict[str, Any]) -> tuple[bool, str]:
    if ctx.doc is not None:
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

    err = _require_ooxml_or_com(ctx, "ppt_slide_transition")
    if err:
        return False, err
    slides = ctx.xml.get("slides", [])
    targets = list(range(1, len(slides) + 1)) if p.get("slides", "all") == "all" else [int(x) for x in p["slides"]]
    fails = []
    for si in targets:
        if si > len(slides):
            fails.append(f"S{si}.missing")
            continue
        sx = slides[si - 1]
        tx = sx.get("transition", {})
        if "entry_effect" in p:
            got_eff = int(tx.get("entry_effect", 0))
            if got_eff != int(p["entry_effect"]):
                fails.append(f"S{si}.eff={got_eff}({tx.get('effect_tag')})")
        if "advance_on_time" in p and bool(tx.get("advance_on_time")) != bool(p["advance_on_time"]):
            fails.append(f"S{si}.adv={tx.get('advance_on_time')}")
        if "advance_time_range" in p:
            lo, hi = p["advance_time_range"]
            t_val = float(tx.get("advance_time", 0.0))
            if not (lo <= t_val <= hi):
                fails.append(f"S{si}.time={t_val}")
        if "hidden" in p and bool(sx.get("hidden")) != bool(p["hidden"]):
            fails.append(f"S{si}.hidden={sx.get('hidden')}")
    return not fails, ("OK (XML)" if not fails else f"xml: {', '.join(fails)}")


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
    else:
        err = _require_ooxml_or_com(ctx, "ppt_shape_match")
        if err:
            return False, err
        W = float(ctx.xml.get("width_pt") or 960.0)
        H = float(ctx.xml.get("height_pt") or 540.0)
        matched_x = None
        for sx in s_xml.get("shapes", []):
            if "type" in p and int(sx.get("type", 0)) != int(p["type"]):
                continue
            if "name" in p and str(sx.get("name") or "") != str(p["name"]):
                continue
            if p.get("has_smartart") and not sx.get("has_smartart"):
                continue
            if p.get("has_chart") and not sx.get("has_chart"):
                continue
            if p.get("has_table") and not sx.get("has_table"):
                continue
            if "text_contains" in p and p["text_contains"] not in (sx.get("text") or ""):
                continue
            matched_x = sx
            break
        if matched_x is None:
            return False, f"xml shape not found on slide {si}"
        if p.get("right_aligned_to_slide") and not _close(matched_x["left"] + matched_x["w"], W, 1.0):
            fails.append(f"R={matched_x['left'] + matched_x['w']:.1f}/{W}")
        if p.get("top_aligned_to_slide") and not _close(matched_x["top"], 0.0, 1.0):
            fails.append(f"T={matched_x['top']:.1f}")
        if p.get("centered_h") and not _close(matched_x["left"], (W - matched_x["w"]) / 2.0, 1.0):
            fails.append(f"L={matched_x['left']:.1f}")
        if p.get("centered_v") and not _close(matched_x["top"], (H - matched_x["h"]) / 2.0, 1.0):
            fails.append(f"T={matched_x['top']:.1f}")
        if "align" in p and matched_x.get("align") != int(p["align"]):
            fails.append(f"align={matched_x.get('align')}")
        if "smallcaps" in p and bool(matched_x.get("smallcaps")) != bool(p["smallcaps"]):
            fails.append(f"smallcaps={matched_x.get('smallcaps')}")
        if "font_fill_pattern" in p and matched_x.get("font_fill_pattern") != int(p["font_fill_pattern"]):
            fails.append(f"pattern={matched_x.get('font_fill_pattern')}")
        if p.get("shadow_visible") and not matched_x.get("shadow_visible"):
            fails.append("shadow_not_visible")
        if "smartart_nodes" in p and matched_x.get("smartart_nodes") != list(p["smartart_nodes"]):
            fails.append(f"smartartNodes={matched_x.get('smartart_nodes')}")

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
    tol = float(p.get("tol_pt", 0.8))
    if ctx.doc is not None:
        sl = ctx.doc.Slides(int(p["slide"]))
        items = [
            {
                "left": float(sl.Shapes(i).Left),
                "top": float(sl.Shapes(i).Top),
                "w": float(sl.Shapes(i).Width),
            }
            for i in range(1, int(sl.Shapes.Count) + 1)
            if ("shape_type" not in p or int(sl.Shapes(i).Type) == int(p["shape_type"]))
        ]
    else:
        err = _require_ooxml_or_com(ctx, "ppt_align_distribute")
        if err:
            return False, err
        slides = ctx.xml.get("slides", [])
        si = int(p["slide"])
        if si > len(slides):
            return False, f"xml slide {si} missing"
        items = [
            {"left": float(sx["left"]), "top": float(sx["top"]), "w": float(sx["w"])}
            for sx in slides[si - 1].get("shapes", [])
            if ("shape_type" not in p or int(sx.get("type", 0)) == int(p["shape_type"]))
        ]

    if len(items) < 2:
        return False, f"need >=2 shapes, got {len(items)}"
    items.sort(key=lambda s: s["left"])
    fails = []
    if p.get("top_aligned"):
        t0 = items[0]["top"]
        if not all(abs(s["top"] - t0) <= tol for s in items):
            fails.append(f"tops={[round(s['top'], 1) for s in items]}")
    if p.get("left_aligned"):
        l0 = items[0]["left"]
        if not all(abs(s["left"] - l0) <= tol for s in items):
            fails.append(f"lefts={[round(s['left'], 1) for s in items]}")
    if p.get("h_distributed") and len(items) >= 3:
        gaps = [items[i + 1]["left"] - (items[i]["left"] + items[i]["w"]) for i in range(len(items) - 1)]
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
    else:
        if "chart_type" in p and cxml.get("chart_type_code") is not None:
            if int(cxml["chart_type_code"]) != int(p["chart_type"]):
                fails.append(f"xmlChartType={cxml['chart_type_code']}")
        if "has_title" in p and bool(cxml.get("title") is not None) != bool(p["has_title"]):
            fails.append(f"xmlHasTitle={cxml.get('title') is not None}")
        if p.get("no_gridlines") and cxml.get("has_gridlines"):
            fails.append("xml_has_gridlines")
    return not fails, ("OK" if not fails else ", ".join(fails))


# ---------------------------------------------------------------------------
# Grader Engine, Rubric Scaffolder & Batch Runner
# ---------------------------------------------------------------------------

def grade_with_rubric(
    candidate_path: str,
    rubric: dict[str, Any],
    *,
    use_com: bool = True,
    session: WPSBatchSession | None = None,
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
        ctx_mgr = (
            session.open(info["path"], read_only=True, with_window=False)
            if session is not None
            else wps_session(kind, info["path"], read_only=True, with_window=False)
        )
        with ctx_mgr as (app, doc, _):
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


def _coord_str(r: int, c: int) -> str:
    res = []
    n = c
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res.append(chr(65 + rem))
    return f"{''.join(reversed(res))}{r}"


def scaffold_rubric(
    reference_path: str,
    total_points: float = 20.0,
    *,
    use_com: bool = True,
) -> dict[str, Any]:
    """Generate a starter JSON rubric by inspecting a reference answer file (COM or static XML)."""
    info = inspect_container(reference_path)
    kind = info["kind"]
    xml_data = audit_ooxml(info["path"]) if info["container"] == "ooxml" else {}
    rules: list[dict[str, Any]] = []

    def _build_word_rules(doc: Any | None) -> None:
        if doc is not None:
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
            for sid, label in ((-1, "正文"), (-2, "标题 1"), (-3, "标题 2")):
                try:
                    st = doc.Styles(sid)
                    rules.append({
                        "id": f"w_style_{abs(sid)}",
                        "name": f"样式 [{label}] 字体与对齐",
                        "type": "word_style",
                        "params": {
                            "style_id": sid,
                            "font_size": round(float(st.Font.Size), 1),
                            "bold": bool(st.Font.Bold),
                            "align": int(st.ParagraphFormat.Alignment),
                        },
                    })
                except Exception:
                    pass
            if int(doc.Tables.Count) > 0:
                t = doc.Tables(1)
                tbl_params: dict[str, Any] = {
                    "table_index": 1,
                    "rows": int(t.Rows.Count),
                    "cols": int(t.Columns.Count),
                    "align": int(t.Rows.Alignment),
                }
                if xml_data.get("has_wps_diagonals"):
                    tbl_params["has_wps_diagonals"] = True
                if xml_data.get("theme_fills"):
                    tbl_params["theme_fill"] = xml_data["theme_fills"][0]
                rules.append({"id": "w_tbl", "name": "表格行列、对齐与表头特征", "type": "word_table", "params": tbl_params})
        elif xml_data.get("sections"):
            sec1 = xml_data["sections"][0]
            page_params: dict[str, Any] = {
                "section": 1,
                "sections_count": xml_data["sections_count"],
                "orientation": sec1.get("orientation", 0),
            }
            if sec1.get("margins_cm"):
                page_params["margins_cm"] = [round(x, 2) for x in sec1["margins_cm"]]
            rules.append({"id": "w_page", "name": "页面设置（边距/方向/节数）", "type": "word_page_setup", "params": page_params})
            styles_bundle = xml_data.get("styles", {})
            for sid, label in (("-1", "正文"), ("-2", "标题 1")):
                mapped = styles_bundle.get("aliases", {}).get(sid)
                st = styles_bundle.get("by_id", {}).get(mapped or "")
                if st and st.get("font_size") is not None:
                    sp: dict[str, Any] = {"style_id": int(sid), "font_size": st["font_size"]}
                    if st.get("align") is not None:
                        sp["align"] = st["align"]
                    rules.append({"id": f"w_style_{abs(int(sid))}", "name": f"样式 [{label}] 字体与对齐", "type": "word_style", "params": sp})
            for idx, tx in enumerate(xml_data.get("tables", [])[:1], 1):
                tp: dict[str, Any] = {"table_index": idx, "rows": tx["rows"], "cols": tx["cols"], "align": tx["align"]}
                if xml_data.get("has_wps_diagonals"):
                    tp["has_wps_diagonals"] = True
                if xml_data.get("theme_fills"):
                    tp["theme_fill"] = xml_data["theme_fills"][0]
                rules.append({"id": f"w_tbl_{idx}", "name": f"表格{idx}结构", "type": "word_table", "params": tp})

        if xml_data.get("numbering_abstracts"):
            ab = xml_data["numbering_abstracts"][0]
            if ab.get("levels"):
                lv0 = ab["levels"][0]
                rules.append({
                    "id": "w_num",
                    "name": "编号/多级列表格式",
                    "type": "word_numbering_xml",
                    "params": {
                        "ilvl": lv0["ilvl"],
                        "num_fmt_prefix": (lv0.get("numFmt") or "")[:6],
                        "lvl_text": lv0.get("lvlText"),
                    },
                })
        eff_params = {
            k: True for k in ("has_w14_textFill", "has_w14_glow", "has_w14_shadow", "has_w14_reflection")
            if xml_data.get(k)
        }
        if eff_params:
            rules.append({"id": "w_fx", "name": "文字艺术特效 (w14)", "type": "word_effects_xml", "params": eff_params})
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

    def _build_excel_rules(wb: Any | None) -> None:
        snames = [s["name"] for s in xml_data.get("sheets", [])]
        if not snames and wb is not None:
            snames = [str(wb.Sheets(i).Name) for i in range(1, int(wb.Sheets.Count) + 1)]
        rules.append({"id": "e_sheets", "name": "工作表结构", "type": "excel_sheets", "params": {"exact_names": snames}})

        st = read_excel_static(info["path"]) if info["container"] == "ooxml" else {}
        for sn in snames:
            ws_st = st.get("sheets", {}).get(sn, {})
            fmls = [(rc, fv) for rc, fv in sorted(ws_st.get("formulas", {}).items()) if isinstance(fv, str) and fv.startswith("=")]
            if fmls:
                (r0, c0), fv0 = fmls[0]
                fn_m = re.match(r"=\s*([A-Za-z_][A-Za-z0-9_]*)\(", fv0)
                f_sub = f"={fn_m.group(1)}(" if fn_m else fv0.split(")")[0]
                rules.append({
                    "id": f"e_fml_{sn}",
                    "name": f"[{sn}] 关键公式检查 ({_coord_str(r0, c0)})",
                    "type": "excel_range_formula",
                    "params": {"sheet": sn, "range": _coord_str(r0, c0), "formula_contains": f_sub},
                })
            val_items = [(_coord_str(r, c), v) for (r, c), v in sorted(ws_st.get("values", {}).items()) if (r, c) in ws_st.get("formulas", {}) and isinstance(ws_st["formulas"][(r, c)], str) and ws_st["formulas"][(r, c)].startswith("=")]
            if val_items:
                rules.append({
                    "id": f"e_val_{sn}",
                    "name": f"[{sn}] 公式计算结果校验",
                    "type": "excel_range_values",
                    "params": {"sheet": sn, "cells": dict(val_items[:3])},
                })

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
        for idx, pt in enumerate(xml_data.get("pivot_tables", []), 1):
            p_params: dict[str, Any] = {"sheet": pt.get("sheet") or (snames[0] if snames else "")}
            if pt.get("row_fields"):
                p_params["row_fields"] = pt["row_fields"]
            if pt.get("data_fields"):
                p_params["data_fields_count"] = len(pt["data_fields"])
            rules.append({"id": f"e_pivot_{idx}", "name": f"数据透视表{idx}字段", "type": "excel_pivot_table", "params": p_params})

    def _build_ppt_rules(prs: Any | None) -> None:
        if prs is not None:
            ps = prs.PageSetup
            rules.append({
                "id": "p_setup",
                "name": "幻灯片尺寸与总页数",
                "type": "ppt_page_setup",
                "params": {
                    "slide_size": int(ps.SlideSize),
                    "width_pt": round(float(ps.SlideWidth), 1),
                    "height_pt": round(float(ps.SlideHeight), 1),
                    "slides_count": int(prs.Slides.Count),
                },
            })
            m_bg = int(prs.SlideMaster.Background.Fill.Type)
            if m_bg != 0:
                rules.append({
                    "id": "p_bg",
                    "name": "母版背景与全页跟随",
                    "type": "ppt_master_background",
                    "params": {
                        "master_fill_type": m_bg,
                        "all_follow_master": all(int(prs.Slides(i).FollowMasterBackground) == -1 for i in range(1, int(prs.Slides.Count) + 1)),
                    },
                })
            tr1 = prs.Slides(1).SlideShowTransition
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
        else:
            rules.append({
                "id": "p_setup",
                "name": "幻灯片尺寸与总页数",
                "type": "ppt_page_setup",
                "params": {
                    "slide_size": xml_data.get("slide_size", 7),
                    "width_pt": xml_data.get("width_pt"),
                    "height_pt": xml_data.get("height_pt"),
                    "slides_count": xml_data.get("slides_count", 0),
                },
            })
            if xml_data.get("master_bg_fill_type"):
                rules.append({
                    "id": "p_bg",
                    "name": "母版背景与全页跟随",
                    "type": "ppt_master_background",
                    "params": {
                        "master_fill_type": xml_data["master_bg_fill_type"],
                        "all_follow_master": all(s.get("follow_master_bg", True) for s in xml_data.get("slides", [])),
                    },
                })
            s1 = (xml_data.get("slides") or [{}])[0]
            tx1 = s1.get("transition", {})
            if tx1.get("present") and tx1.get("entry_effect"):
                rules.append({
                    "id": "p_trans",
                    "name": "幻灯片切换与计时",
                    "type": "ppt_slide_transition",
                    "params": {
                        "slides": "all",
                        "entry_effect": tx1["entry_effect"],
                        "advance_on_time": bool(tx1.get("advance_on_time")),
                    },
                })

        if xml_data.get("prnPr"):
            rules.append({
                "id": "p_print",
                "name": "打印选项配置",
                "type": "ppt_print_options",
                "params": dict(xml_data["prnPr"]),
            })
        anchored_added = 0
        for s_idx, sx in enumerate(xml_data.get("slides", []), 1):
            for bp in sx.get("shapes_body_pr", []):
                if bp.get("anchor") and bp.get("anchorCtr") is not None and (bp.get("text") or bp.get("name")):
                    sp_p: dict[str, Any] = {
                        "slide": s_idx,
                        "body_pr_anchor": bp["anchor"],
                        "body_pr_anchor_ctr": bp["anchorCtr"],
                    }
                    if bp.get("text"):
                        sp_p["text_contains"] = bp["text"][:20]
                    else:
                        sp_p["name"] = bp["name"]
                    rules.append({"id": f"p_anchor_{s_idx}_{anchored_added + 1}", "name": f"第{s_idx}页文本框锚定", "type": "ppt_shape_match", "params": sp_p})
                    anchored_added += 1
                    if anchored_added >= 2:
                        break
            if anchored_added >= 2:
                break
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

    if use_com:
        with wps_session(kind, info["path"], read_only=True, with_window=False) as (_, doc, _):
            if kind == "word":
                _build_word_rules(doc)
            elif kind == "excel":
                _build_excel_rules(doc)
            elif kind == "ppt":
                _build_ppt_rules(doc)
    else:
        if kind == "word":
            _build_word_rules(None)
        elif kind == "excel":
            _build_excel_rules(None)
        elif kind == "ppt":
            _build_ppt_rules(None)

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


def _resolve_batch_target(
    rel: Path,
    cpath: Path,
    ckind: str,
    r_root: Path,
    ref_files: list[Path],
    single_rubrics_by_kind: dict[str, Path],
    single_refs_by_kind: dict[str, Path],
) -> tuple[str, Path] | None:
    if r_root.is_file():
        if r_root.name.endswith(".json"):
            rub = json.loads(r_root.read_text(encoding="utf-8"))
            rkind = rub.get("kind")
            if not rkind or rkind == ckind:
                return ("rubric", r_root)
            return None
        if infer_kind(str(r_root)) == ckind:
            return ("ref_diff", r_root)
        return None

    # 1. Exact relative rubric or stem/name rubric
    rel_rubric = next(
        (
            cand
            for cand in (
                r_root / rel.with_suffix(".rubric.json"),
                r_root / rel.with_name(f"{cpath.name}.rubric.json"),
            )
            if cand.is_file()
        ),
        None,
    )
    if rel_rubric:
        return ("rubric", rel_rubric)

    rubric_match = next(
        (rf for rf in ref_files if rf.name == f"{cpath.stem}.rubric.json" or rf.name == f"{cpath.name}.rubric.json"),
        None,
    )
    if rubric_match:
        return ("rubric", rubric_match)

    # 2. Exact relative reference or normalized mirror directory
    norm_parts = [part.replace("作答文件", "参考答案") for part in rel.parts]
    mirror_rel = Path(*norm_parts) if norm_parts else rel
    for cand_ref in (r_root / rel, r_root / mirror_rel):
        if cand_ref.is_file():
            return ("ref_diff", cand_ref)

    # 3. Same filename across subdirectories
    ref_match = next(
        (
            rf for rf in ref_files
            if rf.name == cpath.name
            and rf.parent.name.replace("参考答案", "") == cpath.parent.name.replace("作答文件", "")
        ),
        None,
    ) or next((rf for rf in ref_files if rf.name == cpath.name and not rf.name.endswith(".json")), None)
    if ref_match:
        return ("ref_diff", ref_match)

    # 4. Single rubric or single reference document per kind (e.g. N students vs 1 standard answer/rubric)
    if ckind in single_rubrics_by_kind:
        return ("rubric", single_rubrics_by_kind[ckind])
    if ckind in single_refs_by_kind:
        return ("ref_diff", single_refs_by_kind[ckind])
    return None


def batch_grade(
    candidate_dir: str,
    ref_or_rubric_dir: str,
    *,
    use_com: bool = True,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> dict[str, Any]:
    """Batch-grade candidate files against a rubric file, reference file, or directory."""
    t0 = time.time()
    c_root = Path(os.path.abspath(candidate_dir))
    r_root = Path(os.path.abspath(ref_or_rubric_dir))

    cand_files = {
        p.relative_to(c_root): p
        for p in sorted(c_root.rglob("*"))
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXAM_EXTS and not p.name.startswith("~$")
    }
    ref_files = [
        p for p in sorted(r_root.rglob("*"))
        if p.is_file() and not p.name.startswith("~$")
    ] if r_root.is_dir() else []

    # Pre-index single rubric / single reference per kind for N-students-vs-1-answer workflows
    rubrics_by_kind: dict[str, list[Path]] = {}
    refs_by_kind: dict[str, list[Path]] = {}
    for rf in ref_files:
        if rf.name.endswith(".rubric.json") or rf.suffix.lower() == ".json":
            try:
                rjson = json.loads(rf.read_text(encoding="utf-8"))
                rk = rjson.get("kind")
                if rk:
                    rubrics_by_kind.setdefault(rk, []).append(rf)
            except Exception:
                pass
        elif rf.suffix.lower() in SUPPORTED_EXAM_EXTS:
            refs_by_kind.setdefault(infer_kind(str(rf)), []).append(rf)

    single_rubrics_by_kind = {k: v[0] for k, v in rubrics_by_kind.items() if len(v) == 1}
    single_refs_by_kind = {k: v[0] for k, v in refs_by_kind.items() if len(v) == 1}

    matched_tasks: list[tuple[Path, Path, str, str, Path]] = []
    unmatched_files: list[str] = []
    for rel, cpath in cand_files.items():
        ckind = infer_kind(str(cpath))
        target = _resolve_batch_target(
            rel, cpath, ckind, r_root, ref_files, single_rubrics_by_kind, single_refs_by_kind
        )
        if target is None:
            unmatched_files.append(str(rel))
        else:
            mode, target_path = target
            matched_tasks.append((rel, cpath, ckind, mode, target_path))

    # Group by kind so each WPS app launches at most once across the batch
    reports_by_rel: dict[str, dict[str, Any]] = {}
    rubric_cache: dict[str, dict[str, Any]] = {}
    ref_com_cache: dict[str, dict[str, Any]] = {}
    total_tasks = len(matched_tasks)
    completed = 0

    tasks_by_kind: dict[str, list[tuple[Path, Path, str, Path]]] = {}
    for rel, cpath, ckind, mode, tpath in matched_tasks:
        tasks_by_kind.setdefault(ckind, []).append((rel, cpath, mode, tpath))

    for kind, kind_tasks in tasks_by_kind.items():
        def _process_one(rel: Path, cpath: Path, mode: str, tpath: Path, sess: WPSBatchSession | None) -> None:
            nonlocal completed
            if mode == "rubric":
                tkey = str(tpath)
                if tkey not in rubric_cache:
                    rubric_cache[tkey] = json.loads(tpath.read_text(encoding="utf-8"))
                rep = grade_with_rubric(str(cpath), rubric_cache[tkey], use_com=use_com, session=sess)
                reports_by_rel[str(rel)] = {"mode": "rubric", "file": str(rel), **rep}
            else:
                suite = compare_documents(
                    str(cpath), str(tpath),
                    use_com=use_com, session=sess, ref_com_cache=ref_com_cache,
                )
                d = suite.to_dict()
                reports_by_rel[str(rel)] = {
                    "mode": "ref_diff",
                    "file": str(rel),
                    "kind": kind,
                    "elapsed_sec": d["elapsed_sec"],
                    "earned_points": float(d["passed"]),
                    "max_points": float(d["total"]),
                    "score_pct": d["score_pct"],
                    "all_passed": d["all_passed"],
                    "details": d["checks"],
                }
            completed += 1
            if progress_callback is not None:
                progress_callback(completed, total_tasks, str(rel))

        if use_com:
            with wps_batch_session(kind, with_window=False) as sess:
                for rel, cpath, mode, tpath in kind_tasks:
                    _process_one(rel, cpath, mode, tpath, sess)
        else:
            for rel, cpath, mode, tpath in kind_tasks:
                _process_one(rel, cpath, mode, tpath, None)

    reports = [reports_by_rel[str(rel)] for rel, *_ in matched_tasks if str(rel) in reports_by_rel]
    total_earned = sum(r["earned_points"] for r in reports)
    total_max = sum(r["max_points"] for r in reports)
    return {
        "candidate_dir": str(c_root),
        "reference_dir": str(r_root),
        "elapsed_sec": round(time.time() - t0, 3),
        "matched_files": len(reports),
        "unmatched_files": unmatched_files,
        "earned_points": round(total_earned, 2),
        "max_points": round(total_max, 2),
        "score_pct": round(100.0 * total_earned / total_max, 1) if total_max else 0.0,
        "all_passed": bool(reports) and all(r["all_passed"] for r in reports),
        "reports": reports,
    }


def _collect_failed_items(item: dict[str, Any]) -> list[dict[str, Any]]:
    if item.get("mode") == "rubric":
        return [
            {
                "name": r.get("name", ""),
                "category": r.get("type", "rubric"),
                "points": f"0/{r.get('points', 0)}",
                "detail": r.get("detail", ""),
                "hint": "",
            }
            for r in item.get("results", [])
            if not r.get("passed")
        ]
    return [
        {
            "name": c.get("name", ""),
            "category": c.get("category", "diff"),
            "points": "0/1",
            "detail": c.get("detail", ""),
            "hint": c.get("hint", ""),
        }
        for c in item.get("details", [])
        if not c.get("passed")
    ]


def export_batch_report(report: dict[str, Any], output_path: str) -> str:
    """Export a batch grading report to UTF-8 BOM CSV or multi-sheet XLSX."""
    out = Path(os.path.abspath(output_path))
    out.parent.mkdir(parents=True, exist_ok=True)
    ext = out.suffix.lower()

    if ext == ".csv":
        with open(out, "w", encoding="utf-8-sig", newline="") as fp:
            writer = csv.writer(fp)
            writer.writerow(["文件", "模式", "题型", "得分", "满分", "得分率(%)", "全项通过", "耗时(s)", "失分项"])
            for item in report.get("reports", []):
                fails = _collect_failed_items(item)
                fail_str = "; ".join(f"{f['name']}({f['detail']})" for f in fails)
                writer.writerow([
                    item.get("file", ""),
                    item.get("mode", ""),
                    item.get("kind", ""),
                    item.get("earned_points", 0),
                    item.get("max_points", 0),
                    item.get("score_pct", 0),
                    "PASS" if item.get("all_passed") else "FAIL",
                    item.get("elapsed_sec", ""),
                    fail_str,
                ])
        return str(out)

    if ext == ".xlsx":
        if openpyxl is None:
            raise RuntimeError("openpyxl is required to export .xlsx score reports.")
        wb = openpyxl.Workbook()
        ws_sum = wb.active
        ws_sum.title = "成绩汇总"
        ws_sum.append(["文件", "模式", "题型", "得分", "满分", "得分率(%)", "全项通过", "耗时(s)", "失分项数"])
        for item in report.get("reports", []):
            fails = _collect_failed_items(item)
            ws_sum.append([
                item.get("file", ""),
                item.get("mode", ""),
                item.get("kind", ""),
                item.get("earned_points", 0),
                item.get("max_points", 0),
                item.get("score_pct", 0),
                "PASS" if item.get("all_passed") else "FAIL",
                item.get("elapsed_sec", 0),
                len(fails),
            ])
        ws_sum.append([
            f"合计 ({report.get('matched_files', 0)} 份)",
            "",
            "",
            report.get("earned_points", 0),
            report.get("max_points", 0),
            report.get("score_pct", 0),
            "PASS" if report.get("all_passed") else "FAIL",
            report.get("elapsed_sec", 0),
            "",
        ])

        ws_fail = wb.create_sheet("失分明细")
        ws_fail.append(["文件", "失分项", "类别", "分值", "实测详情", "修复建议"])
        for item in report.get("reports", []):
            for f in _collect_failed_items(item):
                ws_fail.append([
                    item.get("file", ""),
                    f["name"],
                    f["category"],
                    f["points"],
                    f["detail"],
                    f["hint"],
                ])

        if report.get("unmatched_files"):
            ws_un = wb.create_sheet("未匹配文件")
            ws_un.append(["未匹配作答文件"])
            for uf in report["unmatched_files"]:
                ws_un.append([uf])

        wb.save(out)
        wb.close()
        return str(out)

    raise ValueError(f"Unsupported export format {ext!r}; use .csv or .xlsx")


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
    p_init.add_argument("--no-com", action="store_true")

    p_batch = sub.add_parser("batch", help="Batch-grade a candidate folder against rubrics or references.")
    p_batch.add_argument("candidate_dir")
    p_batch.add_argument("ref_or_rubric_dir")
    p_batch.add_argument("--no-com", action="store_true")
    p_batch.add_argument("--json", action="store_true")
    p_batch.add_argument("--export", help="Optional .csv or .xlsx output path for score report.")

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
        rubric = scaffold_rubric(args.reference, total_points=args.total_points, use_com=not args.no_com)
        Path(args.output).write_text(json.dumps(rubric, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Rubric scaffolded -> {args.output} ({len(rubric['rules'])} rules)")
        sys.exit(0)

    if args.cmd == "batch":
        rep = batch_grade(args.candidate_dir, args.ref_or_rubric_dir, use_com=not args.no_com)
        if args.export:
            exported = export_batch_report(rep, args.export)
            if not args.json:
                print(f"Exported report -> {exported}")
        if args.json:
            print(json.dumps(rep, ensure_ascii=False, indent=2))
        else:
            print(f"=== BATCH SCOREBOARD ({rep['matched_files']} files) ===")
            for item in rep["reports"]:
                tag = "PASS" if item["all_passed"] else "FAIL"
                print(f"[{tag}] {item['file']} ({item['mode']}) -> {item['earned_points']}/{item['max_points']} ({item['score_pct']}%)")
            if rep.get("unmatched_files"):
                print(f"UNMATCHED ({len(rep['unmatched_files'])}): {rep['unmatched_files']}")
            print(f"TOTAL: {rep['earned_points']}/{rep['max_points']} ({rep['score_pct']}%) in {rep['elapsed_sec']:.2f}s")
        sys.exit(0 if rep["all_passed"] else 1)


if __name__ == "__main__":
    main()
