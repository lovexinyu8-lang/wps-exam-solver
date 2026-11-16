"""Grading-standard dual-track auditor & 'Candidate vs Reference Answer' comparator."""

from __future__ import annotations

import json
import os
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from ooxml_patch import NS, audit_ooxml
from wps_env import infer_kind, inspect_container, wps_session


class AuditSuite:
    """Structured PASS/FAIL assertion collector for exam verification & grading."""

    def __init__(self, title: str = "EXAM AUDIT") -> None:
        self.title = title
        self.t0 = time.time()
        self.items: list[dict[str, Any]] = []

    def check(self, name: str, cond: bool, detail: Any = "", *, category: str = "general") -> bool:
        passed = bool(cond)
        self.items.append({
            "name": name,
            "category": category,
            "passed": passed,
            "detail": "" if detail is None else str(detail),
        })
        return passed

    @property
    def passed_count(self) -> int:
        return sum(1 for it in self.items if it["passed"])

    @property
    def total_count(self) -> int:
        return len(self.items)

    @property
    def all_passed(self) -> bool:
        return bool(self.items) and all(it["passed"] for it in self.items)

    def to_dict(self) -> dict[str, Any]:
        total = self.total_count
        passed = self.passed_count
        return {
            "title": self.title,
            "elapsed_sec": round(time.time() - self.t0, 3),
            "passed": passed,
            "total": total,
            "score_pct": round(100.0 * passed / total, 1) if total else 0.0,
            "all_passed": self.all_passed,
            "checks": self.items,
        }

    def format_report(self) -> str:
        lines = [f"=== {self.title} ==="]
        for it in self.items:
            tag = "PASS" if it["passed"] else "FAIL"
            suffix = f" ({it['detail']})" if it["detail"] else ""
            lines.append(f"[{tag}] {it['name']}{suffix}")
        d = self.to_dict()
        status = "ALL PASS" if d["all_passed"] else "HAS FAIL"
        lines.append(
            f"DONE in {d['elapsed_sec']:.2f}s | {status} ({d['passed']}/{d['total']}, {d['score_pct']}%)"
        )
        return "\n".join(lines)


def _norm_formula(f: Any) -> str:
    if f is None:
        return ""
    s = str(f).strip().upper()
    s = re.sub(r"'([^']+)'!", r"\1!", s)
    return s


def _vals_equal(a: Any, b: Any, tol: float = 1e-4) -> bool:
    if a is None and b is None:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) <= tol
    return str(a).strip() == str(b).strip()


# ---------------------------------------------------------------------------
# Word Snapshot & Diff
# ---------------------------------------------------------------------------

def _snapshot_word_com(file_path: str) -> dict[str, Any]:
    with wps_session("word", file_path, read_only=True) as (app, doc, _):
        try:
            doc.Repaginate()
        except Exception:
            pass

        sections = []
        for si in range(1, int(doc.Sections.Count) + 1):
            sec = doc.Sections(si)
            ps = sec.PageSetup
            hdr = sec.Headers(1)
            ftr = sec.Footers(1)
            hdr_shapes = []
            for idx in range(1, int(hdr.Range.ShapeRange.Count) + 1):
                sh = hdr.Range.ShapeRange(idx)
                hdr_shapes.append({
                    "type": int(sh.Type),
                    "name": str(sh.Name),
                    "w": round(float(sh.Width), 1),
                    "h": round(float(sh.Height), 1),
                    "rot": round(float(sh.Rotation), 1),
                })
            sections.append({
                "section": si,
                "section_start": int(ps.SectionStart),
                "orientation": int(ps.Orientation),
                "page_w": round(float(ps.PageWidth), 1),
                "page_h": round(float(ps.PageHeight), 1),
                "margins": [
                    round(float(ps.TopMargin), 1),
                    round(float(ps.BottomMargin), 1),
                    round(float(ps.LeftMargin), 1),
                    round(float(ps.RightMargin), 1),
                ],
                "gutter": round(float(ps.Gutter), 1),
                "gutter_pos": int(ps.GutterPos),
                "hdr_link": bool(hdr.LinkToPrevious),
                "hdr_fields": [str(f.Code.Text).strip() for f in hdr.Range.Fields],
                "hdr_shapes": hdr_shapes,
                "ftr_link": bool(ftr.LinkToPrevious),
            })

        styles = {}
        for sid in (-2, -3, -4, -1):  # Heading 1/2/3, Normal
            try:
                st = doc.Styles(sid)
                pf = st.ParagraphFormat
                styles[str(sid)] = {
                    "name": str(st.NameLocal),
                    "size": float(st.Font.Size),
                    "bold": int(st.Font.Bold),
                    "align": int(pf.Alignment),
                    "before": round(float(pf.SpaceBefore), 2),
                    "after": round(float(pf.SpaceAfter), 2),
                    "unit_before": round(float(pf.LineUnitBefore), 2),
                    "unit_after": round(float(pf.LineUnitAfter), 2),
                    "line_rule": int(pf.LineSpacingRule),
                    "line_spacing": round(float(pf.LineSpacing), 2),
                    "bottom_bdr": (int(pf.Borders(-3).LineStyle), int(pf.Borders(-3).LineWidth)),
                }
            except Exception:
                pass

        paras = []
        for i in range(1, int(doc.Paragraphs.Count) + 1):
            p = doc.Paragraphs(i)
            txt = str(p.Range.Text).rstrip("\r\x07\x0c")
            if not txt:
                continue
            pf = p.Format
            try:
                sname = str(p.Style.NameLocal)
            except Exception:
                sname = ""
            paras.append({
                "text": txt[:60],
                "style": sname,
                "list": str(p.Range.ListFormat.ListString),
                "align": int(pf.Alignment),
                "c_indent": round(float(pf.CharacterUnitFirstLineIndent), 2),
                "line_rule": int(pf.LineSpacingRule),
                "line_spacing": round(float(pf.LineSpacing), 2),
                "unit_before": round(float(pf.LineUnitBefore), 2),
                "unit_after": round(float(pf.LineUnitAfter), 2),
                "font_size": float(p.Range.Font.Size),
            })

        tables = []
        for ti in range(1, int(doc.Tables.Count) + 1):
            t = doc.Tables(ti)
            rc, cc = int(t.Rows.Count), int(t.Columns.Count)
            sample_cells = {}
            for r in (1, 2, rc):
                if r <= rc:
                    for c in (1, min(2, cc), cc):
                        try:
                            cell = t.Cell(r, c)
                            sample_cells[f"{r},{c}"] = {
                                "text": str(cell.Range.Text).rstrip("\r\x07")[:30],
                                "bg": int(cell.Shading.BackgroundPatternColor),
                                "bdr_bottom": (int(cell.Borders(-3).LineStyle), int(cell.Borders(-3).LineWidth)),
                            }
                        except Exception:
                            pass
            try:
                head_fmt = int(t.Rows(1).HeadingFormat)
            except Exception:
                head_fmt = None
            tables.append({
                "rows": rc,
                "cols": cc,
                "align": int(t.Rows.Alignment),
                "heading_repeat": head_fmt,
                "sample_cells": sample_cells,
            })

        shapes = []
        for si in range(1, int(doc.Shapes.Count) + 1):
            sh = doc.Shapes(si)
            item: dict[str, Any] = {
                "type": int(sh.Type),
                "w": round(float(sh.Width), 1),
                "h": round(float(sh.Height), 1),
                "wrap": int(sh.WrapFormat.Type),
                "has_chart": bool(sh.HasChart),
            }
            if sh.HasChart:
                item["chart_type"] = int(sh.Chart.ChartType)
                item["chart_title"] = str(sh.Chart.ChartTitle.Text) if sh.Chart.HasTitle else None
            shapes.append(item)

        inlines = []
        for ii in range(1, int(doc.InlineShapes.Count) + 1):
            ish = doc.InlineShapes(ii)
            item = {
                "type": int(ish.Type),
                "w": round(float(ish.Width), 1),
                "h": round(float(ish.Height), 1),
                "has_chart": bool(ish.HasChart),
            }
            if ish.HasChart:
                item["chart_type"] = int(ish.Chart.ChartType)
            inlines.append(item)

        return {
            "sections": sections,
            "styles": styles,
            "paragraphs": paras,
            "tables": tables,
            "shapes": shapes,
            "inlines": inlines,
            "comments_count": int(doc.Comments.Count),
            "bg_visible": int(doc.Background.Fill.Visible),
            "bg_rgb": int(doc.Background.Fill.ForeColor.RGB) if int(doc.Background.Fill.Visible) == -1 else None,
        }


def _diff_word(suite: AuditSuite, cand_path: str, ref_path: str, use_com: bool) -> None:
    c_xml = audit_ooxml(cand_path)
    r_xml = audit_ooxml(ref_path)

    if c_xml.get("ooxml_supported", True) and r_xml.get("ooxml_supported", True):
        suite.check(
            "XML: 节(sectPr)数量一致",
            c_xml["sections_count"] == r_xml["sections_count"],
            f"cand={c_xml['sections_count']}, ref={r_xml['sections_count']}",
            category="word_xml",
        )
        for idx, (cs, rs) in enumerate(zip(c_xml["sections"], r_xml["sections"]), 1):
            suite.check(
                f"XML: 第{idx}节页面尺寸与边距(pgSz/pgMar)",
                cs["pgSz"] == rs["pgSz"] and cs["pgMar"] == rs["pgMar"],
                f"cand={cs['pgSz']}/{cs['pgMar']} vs ref={rs['pgSz']}/{rs['pgMar']}",
                category="word_xml",
            )
            c_hdr = [(h.get("type"), h.get("drawings", 0), h.get("behind_doc", []), h.get("instr_text", [])) for h in cs["headers"]]
            r_hdr = [(h.get("type"), h.get("drawings", 0), h.get("behind_doc", []), h.get("instr_text", [])) for h in rs["headers"]]
            suite.check(
                f"XML: 第{idx}节页眉图片/域结构",
                c_hdr == r_hdr,
                f"cand={c_hdr} vs ref={r_hdr}",
                category="word_xml",
            )

        c_nums = sorted([(l["ilvl"], l["numFmt"], l["lvlText"], l["lvlRestart"]) for a in c_xml["numbering_abstracts"] for l in a["levels"]])
        r_nums = sorted([(l["ilvl"], l["numFmt"], l["lvlText"], l["lvlRestart"]) for a in r_xml["numbering_abstracts"] for l in a["levels"]])
        if r_nums:
            suite.check(
                "XML: 编号/多级列表定义 (numFmt/lvlText/lvlRestart)",
                all(x in c_nums for x in r_nums),
                f"missing={[x for x in r_nums if x not in c_nums][:5]}",
                category="word_xml",
            )

        for key in ("has_w14_textFill", "has_w14_glow", "has_w14_shadow", "has_w14_reflection", "has_wps_diagonals"):
            if r_xml[key] or c_xml[key]:
                suite.check(f"XML: 特效标记 {key}", c_xml[key] == r_xml[key], f"cand={c_xml[key]}, ref={r_xml[key]}", category="word_xml")

        if r_xml["glow_sat_mods"] or c_xml["glow_sat_mods"]:
            suite.check("XML: 发光饱和度 w14:satMod", c_xml["glow_sat_mods"] == r_xml["glow_sat_mods"], f"cand={c_xml['glow_sat_mods']} vs ref={r_xml['glow_sat_mods']}", category="word_xml")
        if r_xml["theme_fills"] or c_xml["theme_fills"]:
            suite.check("XML: 主题底纹 w:themeFill", c_xml["theme_fills"] == r_xml["theme_fills"], f"cand={len(c_xml['theme_fills'])} vs ref={len(r_xml['theme_fills'])}", category="word_xml")

        suite.check("XML: 内嵌图表数量", len(c_xml["charts"]) == len(r_xml["charts"]), f"cand={len(c_xml['charts'])}, ref={len(r_xml['charts'])}", category="word_xml")
        for idx, (cc, rc) in enumerate(zip(c_xml["charts"], r_xml["charts"]), 1):
            c_sig = (cc["chart_types"], cc["series_count"], [(s["cat_pt_count"], s["val_pt_count"], s["sample_cats"], s["sample_vals"]) for s in cc["series"]], cc["title"])
            r_sig = (rc["chart_types"], rc["series_count"], [(s["cat_pt_count"], s["val_pt_count"], s["sample_cats"], s["sample_vals"]) for s in rc["series"]], rc["title"])
            suite.check(f"XML: 图表{idx}类型/系列/数据点/标题", c_sig == r_sig, f"cand={c_sig} vs ref={r_sig}", category="word_xml")

    if not use_com:
        return

    c_com = _snapshot_word_com(cand_path)
    r_com = _snapshot_word_com(ref_path)
    suite.check("COM: 样式定义 (标题1/2/3/正文)", c_com["styles"] == r_com["styles"], f"diff_keys={[k for k in r_com['styles'] if c_com['styles'].get(k) != r_com['styles'][k]]}", category="word_com")
    suite.check("COM: 非空段落数量", len(c_com["paragraphs"]) == len(r_com["paragraphs"]), f"cand={len(c_com['paragraphs'])}, ref={len(r_com['paragraphs'])}", category="word_com")
    para_diffs = [
        (i + 1, cp, rp)
        for i, (cp, rp) in enumerate(zip(c_com["paragraphs"], r_com["paragraphs"]))
        if cp != rp
    ]
    suite.check("COM: 段落文本/样式/编号/缩进/行距一致性", not para_diffs, f"mismatches={len(para_diffs)} first={para_diffs[:2]}", category="word_com")
    suite.check("COM: 表格结构/底纹/边框", c_com["tables"] == r_com["tables"], f"cand={len(c_com['tables'])} vs ref={len(r_com['tables'])}", category="word_com")
    suite.check("COM: 浮动图形与内嵌图形", c_com["shapes"] == r_com["shapes"] and c_com["inlines"] == r_com["inlines"], f"shapes={len(c_com['shapes'])}/{len(r_com['shapes'])}, inlines={len(c_com['inlines'])}/{len(r_com['inlines'])}", category="word_com")
    suite.check("COM: 批注与页面背景", c_com["comments_count"] == r_com["comments_count"] and c_com["bg_visible"] == r_com["bg_visible"], f"comments={c_com['comments_count']}/{r_com['comments_count']}, bg={c_com['bg_visible']}/{r_com['bg_visible']}", category="word_com")


# ---------------------------------------------------------------------------
# Excel Snapshot & Diff
# ---------------------------------------------------------------------------

def _col_letter(col_idx: int) -> str:
    res = []
    n = col_idx
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res.append(chr(65 + rem))
    return "".join(reversed(res))


def _cell_addr(row_idx: int, col_idx: int) -> str:
    return f"{_col_letter(col_idx)}{row_idx}"


def _as_grid(raw: Any) -> tuple[tuple[Any, ...], ...]:
    if not isinstance(raw, tuple):
        return ((raw,),)
    if raw and not isinstance(raw[0], tuple):
        return (raw,)
    return raw


def _snapshot_excel_com(file_path: str) -> dict[str, Any]:
    with wps_session("excel", file_path, read_only=True) as (app, wb, _):
        sheets = {}
        sheet_names = [str(wb.Sheets(i).Name) for i in range(1, int(wb.Sheets.Count) + 1)]
        for sname in sheet_names:
            ws = wb.Sheets(sname)
            ur = ws.UsedRange
            addr = str(ur.Address)
            row_start = int(ur.Row)
            col_start = int(ur.Column)
            vals = _as_grid(ur.Value)
            fmls = _as_grid(ur.Formula)
            nfmt = ur.NumberFormat
            fcs = []
            for i in range(1, int(ws.Cells.FormatConditions.Count) + 1):
                fc = ws.Cells.FormatConditions(i)
                try:
                    color = int(fc.Interior.Color)
                except Exception:
                    color = None
                fcs.append({
                    "type": int(fc.Type),
                    "applies_to": str(fc.AppliesTo.Address),
                    "formula1": _norm_formula(getattr(fc, "Formula1", "")),
                    "color": color,
                })
            pts = []
            for i in range(1, int(ws.PivotTables().Count) + 1):
                pt = ws.PivotTables(i)
                pts.append({
                    "name": str(pt.Name),
                    "row_fields": [str(f.Name) for f in pt.RowFields],
                    "data_fields": [str(f.Name) for f in pt.DataFields],
                    "table_range": str(pt.TableRange1.Address),
                })
            charts = []
            for i in range(1, int(ws.ChartObjects().Count) + 1):
                ch = ws.ChartObjects(i).Chart
                charts.append({
                    "type": int(ch.ChartType),
                    "style": int(ch.ChartStyle),
                    "has_title": bool(ch.HasTitle),
                    "title": str(ch.ChartTitle.Text) if ch.HasTitle else None,
                })
            sheets[sname] = {
                "address": addr,
                "row_start": row_start,
                "col_start": col_start,
                "values": vals,
                "formulas": fmls,
                "number_format": nfmt if isinstance(nfmt, str) else None,
                "format_conditions": fcs,
                "pivot_tables": pts,
                "charts": charts,
            }
        return {"sheet_names": sheet_names, "sheets": sheets}


def _diff_excel(suite: AuditSuite, cand_path: str, ref_path: str, use_com: bool) -> None:
    c_xml = audit_ooxml(cand_path)
    r_xml = audit_ooxml(ref_path)

    if c_xml.get("ooxml_supported", True) and r_xml.get("ooxml_supported", True):
        c_names = [s["name"] for s in c_xml["sheets"]]
        r_names = [s["name"] for s in r_xml["sheets"]]
        suite.check("XML: 工作表名称与顺序", c_names == r_names, f"cand={c_names} vs ref={r_names}", category="excel_xml")

        for cs, rs in zip(c_xml["worksheets"], r_xml["worksheets"]):
            sname = rs["name"]
            c_cf = [(x["sqref"], [(r["type"], [_norm_formula(f) for f in r["formulas"]]) for r in x["rules"]]) for x in cs["conditional_formatting"]]
            r_cf = [(x["sqref"], [(r["type"], [_norm_formula(f) for f in r["formulas"]]) for r in x["rules"]]) for x in rs["conditional_formatting"]]
            if r_cf or c_cf:
                suite.check(f"XML: [{sname}] 条件格式 sqref 与公式", c_cf == r_cf, f"cand={c_cf} vs ref={r_cf}", category="excel_xml")
            if rs["sort_state"] or cs["sort_state"]:
                suite.check(f"XML: [{sname}] 排序状态 sortState", cs["sort_state"] == rs["sort_state"], f"cand={cs['sort_state']} vs ref={rs['sort_state']}", category="excel_xml")
            if rs["auto_filter"] or cs["auto_filter"]:
                suite.check(f"XML: [{sname}] 自动筛选 autoFilter", cs["auto_filter"] == rs["auto_filter"], f"cand={cs['auto_filter']} vs ref={rs['auto_filter']}", category="excel_xml")

        suite.check("XML: 透视表定义数量", len(c_xml["pivot_tables"]) == len(r_xml["pivot_tables"]), f"cand={len(c_xml['pivot_tables'])} vs ref={len(r_xml['pivot_tables'])}", category="excel_xml")
        suite.check("XML: 图表数量与类型", len(c_xml["charts"]) == len(r_xml["charts"]), f"cand={len(c_xml['charts'])} vs ref={len(r_xml['charts'])}", category="excel_xml")

    if not use_com:
        return

    c_com = _snapshot_excel_com(cand_path)
    r_com = _snapshot_excel_com(ref_path)
    for sname in r_com["sheet_names"]:
        if sname not in c_com["sheets"]:
            suite.check(f"COM: 工作表 [{sname}] 存在", False, "missing in candidate", category="excel_com")
            continue
        cs = c_com["sheets"][sname]
        rs = r_com["sheets"][sname]
        suite.check(f"COM: [{sname}] UsedRange 范围", cs["address"] == rs["address"], f"cand={cs['address']} vs ref={rs['address']}", category="excel_com")

        c_map = {
            (cs["row_start"] + ri, cs["col_start"] + ci): (cs["values"][ri][ci], str(cs["formulas"][ri][ci] or ""))
            for ri in range(len(cs["values"]))
            for ci in range(len(cs["values"][ri]))
        }
        r_map = {
            (rs["row_start"] + ri, rs["col_start"] + ci): (rs["values"][ri][ci], str(rs["formulas"][ri][ci] or ""))
            for ri in range(len(rs["values"]))
            for ci in range(len(rs["values"][ri]))
        }
        val_diffs = []
        fml_diffs = []
        for coord in sorted(set(c_map) | set(r_map)):
            cv_cell, cf_cell = c_map.get(coord, (None, ""))
            rv_cell, rf_cell = r_map.get(coord, (None, ""))
            if not _vals_equal(cv_cell, rv_cell):
                val_diffs.append((_cell_addr(*coord), cv_cell, rv_cell))
            if rf_cell.startswith("=") or cf_cell.startswith("="):
                if _norm_formula(cf_cell) != _norm_formula(rf_cell):
                    fml_diffs.append((_cell_addr(*coord), cf_cell, rf_cell))

        suite.check(f"COM: [{sname}] 单元格计算值一致性", not val_diffs, f"mismatches={len(val_diffs)} sample={val_diffs[:3]}", category="excel_com")
        suite.check(f"COM: [{sname}] 公式结构一致性", not fml_diffs, f"mismatches={len(fml_diffs)} sample={fml_diffs[:3]}", category="excel_com")
        if rs["format_conditions"] or cs["format_conditions"]:
            suite.check(f"COM: [{sname}] 条件格式对象属性", cs["format_conditions"] == rs["format_conditions"], f"cand={cs['format_conditions']} vs ref={rs['format_conditions']}", category="excel_com")
        if rs["pivot_tables"] or cs["pivot_tables"]:
            suite.check(f"COM: [{sname}] 数据透视表字段与范围", cs["pivot_tables"] == rs["pivot_tables"], f"cand={cs['pivot_tables']} vs ref={rs['pivot_tables']}", category="excel_com")
        if rs["charts"] or cs["charts"]:
            suite.check(f"COM: [{sname}] 图表类型与标题", cs["charts"] == rs["charts"], f"cand={cs['charts']} vs ref={rs['charts']}", category="excel_com")


# ---------------------------------------------------------------------------
# PPT Snapshot & Diff
# ---------------------------------------------------------------------------

def _compare_ppt_shapes(c_shapes: list[dict[str, Any]], r_shapes: list[dict[str, Any]], tol: float = 0.5) -> list[str]:
    if len(c_shapes) != len(r_shapes):
        return [f"shape count {len(c_shapes)} != {len(r_shapes)}"]
    diffs: list[str] = []
    for idx, (cs, rs) in enumerate(zip(c_shapes, r_shapes), 1):
        for key in ("type", "text", "has_table", "has_chart", "chart_type", "smartart_nodes", "align", "smallcaps", "font_fill_type"):
            if cs.get(key) != rs.get(key):
                diffs.append(f"#{idx}.{key}: {cs.get(key)!r} != {rs.get(key)!r}")
        for dim in ("left", "top", "w", "h"):
            if abs(float(cs[dim]) - float(rs[dim])) > tol:
                diffs.append(f"#{idx}.{dim}: {cs[dim]} != {rs[dim]}")
    return diffs


def _snapshot_ppt_com(file_path: str) -> dict[str, Any]:
    with wps_session("ppt", file_path, read_only=True, with_window=False) as (_, prs, _):
        ps = prs.PageSetup
        slides = []
        for si in range(1, int(prs.Slides.Count) + 1):
            sl = prs.Slides(si)
            tr = sl.SlideShowTransition
            shapes = []
            for j in range(1, int(sl.Shapes.Count) + 1):
                sp = sl.Shapes(j)
                txt = ""
                align = None
                smallcaps = None
                font_fill_type = None
                try:
                    if sp.HasTextFrame and sp.TextFrame.HasText:
                        txt = str(sp.TextFrame.TextRange.Text).strip()
                        align = int(sp.TextFrame.TextRange.ParagraphFormat.Alignment)
                except Exception:
                    pass
                try:
                    if getattr(sp, "HasTextFrame", False) and sp.TextFrame2.HasText:
                        f2 = sp.TextFrame2.TextRange.Font
                        smallcaps = int(f2.Smallcaps)
                        font_fill_type = int(f2.Fill.Type)
                except Exception:
                    pass
                sa_nodes = []
                if sp.HasSmartArt:
                    sa_nodes = [
                        str(sp.SmartArt.AllNodes(k).TextFrame2.TextRange.Text).strip()
                        for k in range(1, int(sp.SmartArt.AllNodes.Count) + 1)
                    ]
                shapes.append({
                    "type": int(sp.Type),
                    "left": round(float(sp.Left), 1),
                    "top": round(float(sp.Top), 1),
                    "w": round(float(sp.Width), 1),
                    "h": round(float(sp.Height), 1),
                    "text": txt[:40],
                    "align": align,
                    "smallcaps": smallcaps,
                    "font_fill_type": font_fill_type,
                    "has_table": bool(sp.HasTable),
                    "has_chart": bool(sp.HasChart),
                    "chart_type": int(sp.Chart.ChartType) if sp.HasChart else None,
                    "smartart_nodes": sa_nodes,
                })
            slides.append({
                "slide": si,
                "follow_master_bg": int(sl.FollowMasterBackground),
                "bg_fill_type": int(sl.Background.Fill.Type),
                "hidden": int(tr.Hidden),
                "entry_effect": int(tr.EntryEffect),
                "advance_on_time": int(tr.AdvanceOnTime),
                "advance_time": round(float(tr.AdvanceTime), 2),
                "comments": [str(sl.Comments(k).Text) for k in range(1, int(sl.Comments.Count) + 1)],
                "shapes": sorted(shapes, key=lambda x: (x["type"], x["text"], x["left"], x["top"])),
            })
        return {
            "slide_size": int(ps.SlideSize),
            "slide_w": round(float(ps.SlideWidth), 1),
            "slide_h": round(float(ps.SlideHeight), 1),
            "master_bg_type": int(prs.SlideMaster.Background.Fill.Type),
            "slides": slides,
        }


def _diff_ppt(suite: AuditSuite, cand_path: str, ref_path: str, use_com: bool) -> None:
    c_xml = audit_ooxml(cand_path)
    r_xml = audit_ooxml(ref_path)

    if c_xml.get("ooxml_supported", True) and r_xml.get("ooxml_supported", True):
        suite.check("XML: 幻灯片页面尺寸 sldSz", c_xml["sldSz"] == r_xml["sldSz"], f"cand={c_xml['sldSz']} vs ref={r_xml['sldSz']}", category="ppt_xml")
        if r_xml["prnPr"] or c_xml["prnPr"]:
            suite.check("XML: 打印设置 presProps/prnPr", c_xml["prnPr"] == r_xml["prnPr"], f"cand={c_xml['prnPr']} vs ref={r_xml['prnPr']}", category="ppt_xml")
        suite.check("XML: 幻灯片页数", c_xml["slides_count"] == r_xml["slides_count"], f"cand={c_xml['slides_count']} vs ref={r_xml['slides_count']}", category="ppt_xml")
        for cs, rs in zip(c_xml["slides"], r_xml["slides"]):
            sname = Path(rs["file"]).stem
            suite.check(
                f"XML: {sname} 隐藏/切换/动画/hMerge对齐",
                cs["hidden"] == rs["hidden"]
                and cs["has_transition"] == rs["has_transition"]
                and cs["has_timing"] == rs["has_timing"]
                and cs["hmerge_missing_algn"] == 0,
                f"cand={cs} vs ref={rs}",
                category="ppt_xml",
            )
            suite.check(
                f"XML: {sname} bodyPr 锚定(anchor/anchorCtr)",
                cs["body_prs"] == rs["body_prs"],
                f"cand={cs['body_prs']} vs ref={rs['body_prs']}",
                category="ppt_xml",
            )
        suite.check("XML: 图表数量与数据点数", len(c_xml["charts"]) == len(r_xml["charts"]), f"cand={len(c_xml['charts'])} vs ref={len(r_xml['charts'])}", category="ppt_xml")
        for idx, (cc, rc) in enumerate(zip(c_xml["charts"], r_xml["charts"]), 1):
            suite.check(f"XML: 图表{idx}系列与数据点", cc["series"] == rc["series"] and cc["chart_types"] == rc["chart_types"], f"cand={cc} vs ref={rc}", category="ppt_xml")

    if not use_com:
        return

    c_com = _snapshot_ppt_com(cand_path)
    r_com = _snapshot_ppt_com(ref_path)
    suite.check("COM: 幻灯片宽高与母版背景", abs(c_com["slide_w"] - r_com["slide_w"]) < 0.5 and abs(c_com["slide_h"] - r_com["slide_h"]) < 0.5 and c_com["master_bg_type"] == r_com["master_bg_type"], f"cand=({c_com['slide_w']}x{c_com['slide_h']}, bg={c_com['master_bg_type']}) vs ref=({r_com['slide_w']}x{r_com['slide_h']}, bg={r_com['master_bg_type']})", category="ppt_com")
    for cs, rs in zip(c_com["slides"], r_com["slides"]):
        si = rs["slide"]
        suite.check(
            f"COM: 第{si}页切换/计时/隐藏/批注",
            cs["hidden"] == rs["hidden"]
            and cs["entry_effect"] == rs["entry_effect"]
            and cs["advance_on_time"] == rs["advance_on_time"]
            and cs["comments"] == rs["comments"]
            and cs["follow_master_bg"] == rs["follow_master_bg"],
            f"cand_eff={cs['entry_effect']} vs ref_eff={rs['entry_effect']}",
            category="ppt_com",
        )
        shape_diffs = _compare_ppt_shapes(cs["shapes"], rs["shapes"])
        suite.check(
            f"COM: 第{si}页图形结构(位置/文本/SmartArt/图表)",
            not shape_diffs,
            "; ".join(shape_diffs[:4]) if shape_diffs else f"cnt={len(rs['shapes'])}",
            category="ppt_com",
        )


def compare_documents(
    candidate_path: str,
    reference_path: str,
    *,
    use_com: bool = True,
) -> AuditSuite:
    """Compare candidate solution against a reference answer using COM + OOXML dual-track grading."""
    c_info = inspect_container(candidate_path)
    r_info = inspect_container(reference_path)
    if c_info["kind"] != r_info["kind"]:
        raise ValueError(f"Mismatched document types: {c_info['kind']} vs {r_info['kind']}")

    kind = c_info["kind"]
    suite = AuditSuite(f"REF DIFF ({kind.upper()}): {Path(candidate_path).name}")
    if kind == "word":
        _diff_word(suite, c_info["path"], r_info["path"], use_com=use_com)
    elif kind == "excel":
        _diff_excel(suite, c_info["path"], r_info["path"], use_com=use_com)
    elif kind == "ppt":
        _diff_ppt(suite, c_info["path"], r_info["path"], use_com=use_com)
    else:
        raise ValueError(f"Unsupported kind: {kind!r}")
    return suite


def main() -> None:
    import argparse

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="Compare candidate answer file against reference answer (COM + OOXML)."
    )
    parser.add_argument("candidate", help="Candidate answer file path.")
    parser.add_argument("reference", help="Reference answer file path.")
    parser.add_argument("--no-com", action="store_true", help="Run pure OOXML comparison without WPS COM.")
    parser.add_argument("--json", action="store_true", help="Output structured JSON report.")
    args = parser.parse_args()

    suite = compare_documents(args.candidate, args.reference, use_com=not args.no_com)
    if args.json:
        print(json.dumps(suite.to_dict(), ensure_ascii=False, indent=2))
    else:
        print(suite.format_report())
    sys.exit(0 if suite.all_passed else 1)


if __name__ == "__main__":
    main()
