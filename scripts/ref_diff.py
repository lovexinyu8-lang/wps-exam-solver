"""Grading-standard dual-track auditor & 'Candidate vs Reference Answer' comparator."""

from __future__ import annotations

import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from ooxml_patch import audit_ooxml, read_excel_static
from wps_env import WPSBatchSession, infer_kind, inspect_container, wps_batch_session, wps_session


class AuditSuite:
    """Structured PASS/FAIL assertion collector for exam verification & grading."""

    def __init__(self, title: str = "EXAM AUDIT") -> None:
        self.title = title
        self.t0 = time.time()
        self.items: list[dict[str, Any]] = []

    def check(
        self,
        name: str,
        cond: bool,
        detail: Any = "",
        *,
        category: str = "general",
        location: str = "",
        expected: Any = None,
        actual: Any = None,
        hint: str = "",
    ) -> bool:
        passed = bool(cond)
        entry: dict[str, Any] = {
            "name": name,
            "category": category,
            "passed": passed,
            "detail": "" if detail is None else str(detail),
        }
        if location:
            entry["location"] = str(location)
        if expected is not None:
            entry["expected"] = expected
        if actual is not None:
            entry["actual"] = actual
        if hint:
            entry["hint"] = str(hint)
        self.items.append(entry)
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
            loc = f"@{it['location']} " if it.get("location") else ""
            suffix = f" ({it['detail']})" if it["detail"] else ""
            hint = f" -> 建议: {it['hint']}" if (not it["passed"] and it.get("hint")) else ""
            lines.append(f"[{tag}] {loc}{it['name']}{suffix}{hint}")
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

def _extract_word_snapshot(doc: Any) -> dict[str, Any]:
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


def _snapshot_word_com(file_path: str, session: WPSBatchSession | None = None) -> dict[str, Any]:
    ctx_mgr = session.open(file_path, read_only=True) if session is not None else wps_session("word", file_path, read_only=True)
    with ctx_mgr as (_, doc, _):
        return _extract_word_snapshot(doc)


def _diff_word_tables_detail(c_tables: list[dict[str, Any]], r_tables: list[dict[str, Any]]) -> list[str]:
    if len(c_tables) != len(r_tables):
        return [f"table count {len(c_tables)} != {len(r_tables)}"]
    diffs: list[str] = []
    for idx, (ct, rt) in enumerate(zip(c_tables, r_tables), 1):
        for k in ("rows", "cols", "align", "heading_repeat"):
            if ct.get(k) != rt.get(k):
                diffs.append(f"T{idx}.{k}: {ct.get(k)!r} != {rt.get(k)!r}")
        c_cells = ct.get("sample_cells", ct.get("cells", {}))
        r_cells = rt.get("sample_cells", rt.get("cells", {}))
        for coord, r_val in sorted(r_cells.items()):
            c_val = c_cells.get(coord)
            if c_val != r_val:
                diffs.append(f"T{idx}({coord}): {c_val} != {r_val}")
                if len(diffs) >= 6:
                    return diffs
    return diffs


def _diff_word(
    suite: AuditSuite,
    cand_path: str,
    ref_path: str,
    use_com: bool,
    *,
    session: WPSBatchSession | None = None,
    ref_com_cache: dict[str, dict[str, Any]] | None = None,
) -> None:
    c_xml = audit_ooxml(cand_path)
    r_xml = audit_ooxml(ref_path)

    xml_ok = c_xml.get("ooxml_supported", True) and r_xml.get("ooxml_supported", True)
    if xml_ok:
        suite.check(
            "XML: 节(sectPr)数量一致",
            c_xml["sections_count"] == r_xml["sections_count"],
            f"cand={c_xml['sections_count']}, ref={r_xml['sections_count']}",
            category="word_xml",
            location="document.xml",
            expected=r_xml["sections_count"],
            actual=c_xml["sections_count"],
            hint="核对是否正确插入下一页分节符 (wdSectionBreakNextPage)",
        )
        for idx, (cs, rs) in enumerate(zip(c_xml["sections"], r_xml["sections"]), 1):
            suite.check(
                f"XML: 第{idx}节页面尺寸与边距(pgSz/pgMar)",
                cs["pgSz"] == rs["pgSz"] and cs["pgMar"] == rs["pgMar"],
                f"cand={cs['pgSz']}/{cs['pgMar']} vs ref={rs['pgSz']}/{rs['pgMar']}",
                category="word_xml",
                location=f"section[{idx}]",
                expected={"pgSz": rs["pgSz"], "pgMar": rs["pgMar"]},
                actual={"pgSz": cs["pgSz"], "pgMar": cs["pgMar"]},
                hint="核对页面方向、垂直文字方向设置顺序及装订线位置",
            )
            c_hdr = [(h.get("type"), h.get("shape_count", h.get("drawings", 0)), h.get("behind_doc", []), h.get("instr_text", [])) for h in cs["headers"]]
            r_hdr = [(h.get("type"), h.get("shape_count", h.get("drawings", 0)), h.get("behind_doc", []), h.get("instr_text", [])) for h in rs["headers"]]
            suite.check(
                f"XML: 第{idx}节页眉图片/域结构",
                c_hdr == r_hdr,
                f"cand={c_hdr} vs ref={r_hdr}",
                category="word_xml",
                location=f"section[{idx}].header",
                expected=r_hdr,
                actual=c_hdr,
                hint="页眉图片水印需设 WrapFormat.Type=5 (behindDoc='1') 并传 Anchor=header.Range；先插日期域再插水印",
            )

        c_nums = sorted([(l["ilvl"], l["numFmt"], l["lvlText"], l["lvlRestart"]) for a in c_xml["numbering_abstracts"] for l in a["levels"]])
        r_nums = sorted([(l["ilvl"], l["numFmt"], l["lvlText"], l["lvlRestart"]) for a in r_xml["numbering_abstracts"] for l in a["levels"]])
        if r_nums:
            missing_nums = [x for x in r_nums if x not in c_nums]
            suite.check(
                "XML: 编号/多级列表定义 (numFmt/lvlText/lvlRestart)",
                not missing_nums,
                f"missing={missing_nums[:5]}",
                category="word_xml",
                location="word/numbering.xml",
                expected=r_nums,
                actual=c_nums,
                hint="使用 ListTemplates.Add(True) 绑定内建样式并实测校验 chineseCounting 枚举值",
            )

        effect_hints = {
            "has_w14_textFill": "请调用 ooxml_patch.patch_word_grad_fill() 注入文字渐变并用 WPS 重存",
            "has_w14_glow": "请调用 ooxml_patch.patch_word_glow() 注入发光节点并用 WPS 重存",
            "has_wps_diagonals": "请调用 ooxml_patch.patch_word_double_diagonal_header() 注入双斜线表头并用 WPS 重存",
        }
        for key in ("has_w14_textFill", "has_w14_glow", "has_w14_shadow", "has_w14_reflection", "has_wps_diagonals"):
            if r_xml[key] or c_xml[key]:
                suite.check(
                    f"XML: 特效标记 {key}",
                    c_xml[key] == r_xml[key],
                    f"cand={c_xml[key]}, ref={r_xml[key]}",
                    category="word_xml",
                    location="word/document.xml",
                    expected=r_xml[key],
                    actual=c_xml[key],
                    hint=effect_hints.get(key, ""),
                )

        if r_xml["glow_sat_mods"] or c_xml["glow_sat_mods"]:
            suite.check(
                "XML: 发光饱和度 w14:satMod",
                c_xml["glow_sat_mods"] == r_xml["glow_sat_mods"],
                f"cand={c_xml['glow_sat_mods']} vs ref={r_xml['glow_sat_mods']}",
                category="word_xml",
                location="word/document.xml",
                expected=r_xml["glow_sat_mods"],
                actual=c_xml["glow_sat_mods"],
                hint="请调用 patch_word_glow() 补齐 w14:satMod 并经 WPS 重存",
            )
        if r_xml["theme_fills"] or c_xml["theme_fills"]:
            suite.check(
                "XML: 主题底纹 w:themeFill",
                c_xml["theme_fills"] == r_xml["theme_fills"],
                f"cand={len(c_xml['theme_fills'])} vs ref={len(r_xml['theme_fills'])}",
                category="word_xml",
                location="word/document.xml",
                expected=r_xml["theme_fills"],
                actual=c_xml["theme_fills"],
                hint="请调用 patch_word_table_theme_fill() 补全 w:themeFill 主题槽位",
            )

        suite.check(
            "XML: 内嵌图表数量",
            len(c_xml["charts"]) == len(r_xml["charts"]),
            f"cand={len(c_xml['charts'])}, ref={len(r_xml['charts'])}",
            category="word_xml",
            location="word/charts",
            expected=len(r_xml["charts"]),
            actual=len(c_xml["charts"]),
        )
        max_charts = max(len(c_xml["charts"]), len(r_xml["charts"]))
        for idx in range(max_charts):
            if idx >= len(c_xml["charts"]) or idx >= len(r_xml["charts"]):
                suite.check(
                    f"XML: 图表{idx + 1}存在性",
                    False,
                    f"cand_charts={len(c_xml['charts'])}, ref_charts={len(r_xml['charts'])}",
                    category="word_xml",
                    location=f"word/charts/chart{idx + 1}.xml",
                )
                continue
            cc, rc = c_xml["charts"][idx], r_xml["charts"][idx]
            c_sig = (cc["chart_types"], cc["series_count"], [(s["cat_pt_count"], s["val_pt_count"], s["sample_cats"], s["sample_vals"]) for s in cc["series"]], cc["title"])
            r_sig = (rc["chart_types"], rc["series_count"], [(s["cat_pt_count"], s["val_pt_count"], s["sample_cats"], s["sample_vals"]) for s in rc["series"]], rc["title"])
            suite.check(
                f"XML: 图表{idx + 1}类型/系列/数据点/标题",
                c_sig == r_sig,
                f"cand={c_sig} vs ref={r_sig}",
                category="word_xml",
                location=rc["file"],
                expected=r_sig,
                actual=c_sig,
                hint="图表默认仅 4 行模板，超过 4 行需在 ChartData.Activate() 后用 Rows(5).Insert() 扩行后再填数",
            )
    elif not use_com:
        suite.check(
            "XML: 容器格式支持纯静态比对",
            False,
            f"cand={c_xml.get('container')}, ref={r_xml.get('container')}",
            category="word_xml",
            hint="旧版二进制 OLE2 (.doc/.wps) 需移除 --no-com 参数使用 WPS COM 对比",
        )
        return

    if not use_com:
        c_paras = c_xml.get("paragraphs", [])
        r_paras = r_xml.get("paragraphs", [])
        suite.check(
            "XML: 非空段落数量",
            len(c_paras) == len(r_paras),
            f"cand={len(c_paras)}, ref={len(r_paras)}",
            category="word_xml",
            location="word/document.xml",
            expected=len(r_paras),
            actual=len(c_paras),
        )
        p_diffs = []
        for i, (cp, rp) in enumerate(zip(c_paras, r_paras), 1):
            for k in ("full_text", "style", "align", "c_indent", "line_rule", "line_spacing", "font_size"):
                if cp.get(k) != rp.get(k):
                    p_diffs.append((i, k, cp.get(k), rp.get(k)))
                    break
        suite.check(
            "XML: 段落文本/样式/缩进/行距一致性",
            not p_diffs,
            f"mismatches={len(p_diffs)} first={p_diffs[:2]}",
            category="word_xml",
            location="word/document.xml",
            expected=p_diffs[0][3] if p_diffs else None,
            actual=p_diffs[0][2] if p_diffs else None,
        )
        tbl_diffs = _diff_word_tables_detail(c_xml.get("tables", []), r_xml.get("tables", []))
        suite.check(
            "XML: 表格结构与单元格内容",
            not tbl_diffs,
            "; ".join(tbl_diffs) if tbl_diffs else f"cnt={len(r_xml.get('tables', []))}",
            category="word_xml",
            location="word/document.xml#tables",
        )
        return

    if session is not None:
        c_com = _snapshot_word_com(cand_path, session=session)
        if ref_com_cache is not None and ref_path in ref_com_cache:
            r_com = ref_com_cache[ref_path]
        else:
            r_com = _snapshot_word_com(ref_path, session=session)
            if ref_com_cache is not None:
                ref_com_cache[ref_path] = r_com
    else:
        with wps_batch_session("word") as sess:
            c_com = _snapshot_word_com(cand_path, session=sess)
            r_com = _snapshot_word_com(ref_path, session=sess)

    diff_style_keys = [k for k in r_com["styles"] if c_com["styles"].get(k) != r_com["styles"][k]]
    suite.check(
        "COM: 样式定义 (标题1/2/3/正文)",
        not diff_style_keys,
        f"diff_keys={diff_style_keys}",
        category="word_com",
        location="Styles",
        expected={k: r_com["styles"][k] for k in diff_style_keys} if diff_style_keys else None,
        actual={k: c_com["styles"].get(k) for k in diff_style_keys} if diff_style_keys else None,
        hint="请使用内建枚举 ID doc.Styles(-1/-2/-3/-4) 统一设置字体、段前段后行数与下边框",
    )
    suite.check(
        "COM: 非空段落数量",
        len(c_com["paragraphs"]) == len(r_com["paragraphs"]),
        f"cand={len(c_com['paragraphs'])}, ref={len(r_com['paragraphs'])}",
        category="word_com",
        expected=len(r_com["paragraphs"]),
        actual=len(c_com["paragraphs"]),
    )
    para_diffs = [
        (i + 1, cp, rp)
        for i, (cp, rp) in enumerate(zip(c_com["paragraphs"], r_com["paragraphs"]))
        if cp != rp
    ]
    suite.check(
        "COM: 段落文本/样式/编号/缩进/行距一致性",
        not para_diffs,
        f"mismatches={len(para_diffs)} first={para_diffs[:2]}",
        category="word_com",
        location=f"Paragraph({para_diffs[0][0]})" if para_diffs else "",
        expected=para_diffs[0][2] if para_diffs else None,
        actual=para_diffs[0][1] if para_diffs else None,
    )
    tbl_diffs = _diff_word_tables_detail(c_com["tables"], r_com["tables"])
    suite.check(
        "COM: 表格结构/底纹/边框",
        not tbl_diffs,
        "; ".join(tbl_diffs) if tbl_diffs else f"cnt={len(r_com['tables'])}",
        category="word_com",
        location="Tables",
        expected=len(r_com["tables"]),
        actual=len(c_com["tables"]),
    )
    suite.check(
        "COM: 浮动图形与内嵌图形",
        c_com["shapes"] == r_com["shapes"] and c_com["inlines"] == r_com["inlines"],
        f"shapes={len(c_com['shapes'])}/{len(r_com['shapes'])}, inlines={len(c_com['inlines'])}/{len(r_com['inlines'])}",
        category="word_com",
        location="Shapes/InlineShapes",
        expected={"shapes": r_com["shapes"], "inlines": r_com["inlines"]},
        actual={"shapes": c_com["shapes"], "inlines": c_com["inlines"]},
    )
    suite.check(
        "COM: 批注与页面背景",
        c_com["comments_count"] == r_com["comments_count"] and c_com["bg_visible"] == r_com["bg_visible"],
        f"comments={c_com['comments_count']}/{r_com['comments_count']}, bg={c_com['bg_visible']}/{r_com['bg_visible']}",
        category="word_com",
        hint="设置页面背景颜色必须先调用 doc.Background.Fill.Solid() 再设颜色",
    )


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


def _extract_excel_snapshot(wb: Any) -> dict[str, Any]:
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


def _snapshot_excel_com(file_path: str, session: WPSBatchSession | None = None) -> dict[str, Any]:
    ctx_mgr = session.open(file_path, read_only=True) if session is not None else wps_session("excel", file_path, read_only=True)
    with ctx_mgr as (_, wb, _):
        return _extract_excel_snapshot(wb)


def _diff_excel(
    suite: AuditSuite,
    cand_path: str,
    ref_path: str,
    use_com: bool,
    *,
    session: WPSBatchSession | None = None,
    ref_com_cache: dict[str, dict[str, Any]] | None = None,
) -> None:
    c_xml = audit_ooxml(cand_path)
    r_xml = audit_ooxml(ref_path)

    xml_ok = c_xml.get("ooxml_supported", True) and r_xml.get("ooxml_supported", True)
    if xml_ok:
        c_names = [s["name"] for s in c_xml["sheets"]]
        r_names = [s["name"] for s in r_xml["sheets"]]
        suite.check(
            "XML: 工作表名称与顺序",
            c_names == r_names,
            f"cand={c_names} vs ref={r_names}",
            category="excel_xml",
            location="xl/workbook.xml",
            expected=r_names,
            actual=c_names,
        )

        c_ws_by_name = {w["name"]: w for w in c_xml["worksheets"]}
        for rs in r_xml["worksheets"]:
            sname = rs["name"]
            cs = c_ws_by_name.get(sname)
            if cs is None:
                suite.check(f"XML: 工作表 [{sname}] 存在", False, "missing in candidate", category="excel_xml", location=sname)
                continue
            c_cf = [(x["sqref"], [(r["type"], [_norm_formula(f) for f in r["formulas"]]) for r in x["rules"]]) for x in cs["conditional_formatting"]]
            r_cf = [(x["sqref"], [(r["type"], [_norm_formula(f) for f in r["formulas"]]) for r in x["rules"]]) for x in rs["conditional_formatting"]]
            if r_cf or c_cf:
                suite.check(
                    f"XML: [{sname}] 条件格式 sqref 与公式",
                    c_cf == r_cf,
                    f"cand={c_cf} vs ref={r_cf}",
                    category="excel_xml",
                    location=sname,
                    expected=r_cf,
                    actual=c_cf,
                    hint="设置条件格式前必须先 ws.Activate() 且 rng.Select() 将 ActiveCell 锚定至选区左上角，防止公式行号偏移",
                )
            if rs["sort_state"] or cs["sort_state"]:
                suite.check(
                    f"XML: [{sname}] 排序状态 sortState",
                    cs["sort_state"] == rs["sort_state"],
                    f"cand={cs['sort_state']} vs ref={rs['sort_state']}",
                    category="excel_xml",
                    location=sname,
                    expected=rs["sort_state"],
                    actual=cs["sort_state"],
                    hint="必须调用原生 ws.Sort 排序以生成 <sortState>，严禁用 Python 排好序直接覆盖值",
                )
            if rs["auto_filter"] or cs["auto_filter"]:
                suite.check(
                    f"XML: [{sname}] 自动筛选 autoFilter",
                    cs["auto_filter"] == rs["auto_filter"],
                    f"cand={cs['auto_filter']} vs ref={rs['auto_filter']}",
                    category="excel_xml",
                    location=sname,
                    expected=rs["auto_filter"],
                    actual=cs["auto_filter"],
                )

        suite.check(
            "XML: 透视表定义数量",
            len(c_xml["pivot_tables"]) == len(r_xml["pivot_tables"]),
            f"cand={len(c_xml['pivot_tables'])} vs ref={len(r_xml['pivot_tables'])}",
            category="excel_xml",
            expected=len(r_xml["pivot_tables"]),
            actual=len(c_xml["pivot_tables"]),
        )
        suite.check(
            "XML: 图表数量与类型",
            len(c_xml["charts"]) == len(r_xml["charts"]),
            f"cand={len(c_xml['charts'])} vs ref={len(r_xml['charts'])}",
            category="excel_xml",
            expected=len(r_xml["charts"]),
            actual=len(c_xml["charts"]),
        )
    elif not use_com:
        suite.check(
            "XML: 容器格式支持纯静态比对",
            False,
            f"cand={c_xml.get('container')}, ref={r_xml.get('container')}",
            category="excel_xml",
            hint="旧版二进制 OLE2 (.xls/.et) 需移除 --no-com 参数使用 WPS COM 对比",
        )
        return

    if not use_com:
        c_st = read_excel_static(cand_path)
        r_st = read_excel_static(ref_path)
        for sname in r_st.get("sheet_names", []):
            if sname not in c_st.get("sheets", {}):
                continue
            cs = c_st["sheets"][sname]
            rs = r_st["sheets"][sname]
            all_coords = sorted(set(cs["formulas"]) | set(rs["formulas"]))
            fml_diffs = []
            val_diffs = []
            for coord in all_coords:
                cf_raw = cs["formulas"].get(coord)
                rf_raw = rs["formulas"].get(coord)
                cf_str = str(cf_raw or "")
                rf_str = str(rf_raw or "")
                if rf_str.startswith("=") or cf_str.startswith("="):
                    if _norm_formula(cf_str) != _norm_formula(rf_str):
                        fml_diffs.append((_cell_addr(*coord), cf_str, rf_str))
                cv = cs["values"].get(coord)
                rv = rs["values"].get(coord)
                if not _vals_equal(cv, rv):
                    val_diffs.append((_cell_addr(*coord), cv, rv))
            suite.check(
                f"STATIC: [{sname}] 公式结构一致性",
                not fml_diffs,
                f"mismatches={len(fml_diffs)} sample={fml_diffs[:3]}",
                category="excel_static",
                location=sname,
                expected=fml_diffs[0][2] if fml_diffs else None,
                actual=fml_diffs[0][1] if fml_diffs else None,
                hint="请检查公式函数名与绝对/相对引用是否与参考答案一致",
            )
            suite.check(
                f"STATIC: [{sname}] 单元格缓存值一致性",
                not val_diffs,
                f"mismatches={len(val_diffs)} sample={val_diffs[:3]}",
                category="excel_static",
                location=sname,
                expected=val_diffs[0][2] if val_diffs else None,
                actual=val_diffs[0][1] if val_diffs else None,
                hint="若公式单元格缓存值为 None，请用 WPS 打开触发 Calculate() 保存后再检",
            )
        return

    if session is not None:
        c_com = _snapshot_excel_com(cand_path, session=session)
        if ref_com_cache is not None and ref_path in ref_com_cache:
            r_com = ref_com_cache[ref_path]
        else:
            r_com = _snapshot_excel_com(ref_path, session=session)
            if ref_com_cache is not None:
                ref_com_cache[ref_path] = r_com
    else:
        with wps_batch_session("excel") as sess:
            c_com = _snapshot_excel_com(cand_path, session=sess)
            r_com = _snapshot_excel_com(ref_path, session=sess)

    for sname in r_com["sheet_names"]:
        if sname not in c_com["sheets"]:
            suite.check(f"COM: 工作表 [{sname}] 存在", False, "missing in candidate", category="excel_com", location=sname)
            continue
        cs = c_com["sheets"][sname]
        rs = r_com["sheets"][sname]
        suite.check(
            f"COM: [{sname}] UsedRange 范围",
            cs["address"] == rs["address"],
            f"cand={cs['address']} vs ref={rs['address']}",
            category="excel_com",
            location=sname,
            expected=rs["address"],
            actual=cs["address"],
        )

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

        suite.check(
            f"COM: [{sname}] 单元格计算值一致性",
            not val_diffs,
            f"mismatches={len(val_diffs)} sample={val_diffs[:3]}",
            category="excel_com",
            location=f"{sname}!{val_diffs[0][0]}" if val_diffs else sname,
            expected=val_diffs[0][2] if val_diffs else None,
            actual=val_diffs[0][1] if val_diffs else None,
        )
        suite.check(
            f"COM: [{sname}] 公式结构一致性",
            not fml_diffs,
            f"mismatches={len(fml_diffs)} sample={fml_diffs[:3]}",
            category="excel_com",
            location=f"{sname}!{fml_diffs[0][0]}" if fml_diffs else sname,
            expected=fml_diffs[0][2] if fml_diffs else None,
            actual=fml_diffs[0][1] if fml_diffs else None,
        )
        if rs["format_conditions"] or cs["format_conditions"]:
            suite.check(
                f"COM: [{sname}] 条件格式对象属性",
                cs["format_conditions"] == rs["format_conditions"],
                f"cand={cs['format_conditions']} vs ref={rs['format_conditions']}",
                category="excel_com",
                location=sname,
                expected=rs["format_conditions"],
                actual=cs["format_conditions"],
            )
        if rs["pivot_tables"] or cs["pivot_tables"]:
            suite.check(
                f"COM: [{sname}] 数据透视表字段与范围",
                cs["pivot_tables"] == rs["pivot_tables"],
                f"cand={cs['pivot_tables']} vs ref={rs['pivot_tables']}",
                category="excel_com",
                location=sname,
                expected=rs["pivot_tables"],
                actual=cs["pivot_tables"],
            )
        if rs["charts"] or cs["charts"]:
            suite.check(
                f"COM: [{sname}] 图表类型与标题",
                cs["charts"] == rs["charts"],
                f"cand={cs['charts']} vs ref={rs['charts']}",
                category="excel_com",
                location=sname,
                expected=rs["charts"],
                actual=cs["charts"],
            )


# ---------------------------------------------------------------------------
# PPT Snapshot & Diff
# ---------------------------------------------------------------------------

def _compare_ppt_shapes(
    c_shapes: list[dict[str, Any]],
    r_shapes: list[dict[str, Any]],
    tol: float = 0.5,
    *,
    keys: tuple[str, ...] = ("type", "text", "has_table", "has_chart", "chart_type", "smartart_nodes", "align", "smallcaps", "font_fill_type"),
) -> list[str]:
    if len(c_shapes) != len(r_shapes):
        return [f"shape count {len(c_shapes)} != {len(r_shapes)}"]
    diffs: list[str] = []
    for idx, (cs, rs) in enumerate(zip(c_shapes, r_shapes), 1):
        for key in keys:
            if cs.get(key) != rs.get(key):
                diffs.append(f"#{idx}.{key}: {cs.get(key)!r} != {rs.get(key)!r}")
        for dim in ("left", "top", "w", "h"):
            if abs(float(cs[dim]) - float(rs[dim])) > tol:
                diffs.append(f"#{idx}.{dim}: {cs[dim]} != {rs[dim]}")
    return diffs


def _extract_ppt_snapshot(prs: Any) -> dict[str, Any]:
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


def _snapshot_ppt_com(file_path: str, session: WPSBatchSession | None = None) -> dict[str, Any]:
    ctx_mgr = session.open(file_path, read_only=True, with_window=False) if session is not None else wps_session("ppt", file_path, read_only=True, with_window=False)
    with ctx_mgr as (_, prs, _):
        return _extract_ppt_snapshot(prs)


def _diff_ppt(
    suite: AuditSuite,
    cand_path: str,
    ref_path: str,
    use_com: bool,
    *,
    session: WPSBatchSession | None = None,
    ref_com_cache: dict[str, dict[str, Any]] | None = None,
) -> None:
    c_xml = audit_ooxml(cand_path)
    r_xml = audit_ooxml(ref_path)

    xml_ok = c_xml.get("ooxml_supported", True) and r_xml.get("ooxml_supported", True)
    if xml_ok:
        suite.check(
            "XML: 幻灯片页面尺寸 sldSz",
            c_xml["sldSz"] == r_xml["sldSz"],
            f"cand={c_xml['sldSz']} vs ref={r_xml['sldSz']}",
            category="ppt_xml",
            location="ppt/presentation.xml",
            expected=r_xml["sldSz"],
            actual=c_xml["sldSz"],
            hint="16:9 宽屏需设 SlideSize=7, SlideWidth=960, SlideHeight=540（仅设 SlideSize=15 会缩成 720x405）",
        )
        if r_xml["prnPr"] or c_xml["prnPr"]:
            suite.check(
                "XML: 打印设置 presProps/prnPr",
                c_xml["prnPr"] == r_xml["prnPr"],
                f"cand={c_xml['prnPr']} vs ref={r_xml['prnPr']}",
                category="ppt_xml",
                location="ppt/presProps.xml",
                expected=r_xml["prnPr"],
                actual=c_xml["prnPr"],
                hint="Kwpp 保存会丢失 PrintOptions，请调用 patch_ppt_print_options() 并用 WPS 重存",
            )
        suite.check(
            "XML: 幻灯片页数",
            c_xml["slides_count"] == r_xml["slides_count"],
            f"cand={c_xml['slides_count']} vs ref={r_xml['slides_count']}",
            category="ppt_xml",
            location="ppt/slides",
            expected=r_xml["slides_count"],
            actual=c_xml["slides_count"],
        )
        for cs, rs in zip(c_xml["slides"], r_xml["slides"]):
            sname = Path(rs["file"]).stem
            suite.check(
                f"XML: {sname} 隐藏/切换/动画/hMerge对齐",
                cs["hidden"] == rs["hidden"]
                and cs["has_transition"] == rs["has_transition"]
                and cs["has_timing"] == rs["has_timing"]
                and cs["hmerge_missing_algn"] == 0,
                f"cand={cs['hidden']}/{cs['has_transition']}/{cs['has_timing']}/hMergeErr={cs['hmerge_missing_algn']} vs ref={rs['hidden']}/{rs['has_transition']}/{rs['has_timing']}",
                category="ppt_xml",
                location=rs["file"],
                hint="PPT 表格合并后若丢失隐藏 hMerge='1' 居中对齐，请调用 patch_ppt_hmerge_align()",
            )
            suite.check(
                f"XML: {sname} bodyPr 锚定(anchor/anchorCtr)",
                cs["body_prs"] == rs["body_prs"],
                f"cand={cs['body_prs']} vs ref={rs['body_prs']}",
                category="ppt_xml",
                location=rs["file"],
                expected=rs["body_prs"],
                actual=cs["body_prs"],
                hint="顶端对齐必须用 TextFrame2 同时设置 VerticalAnchor=1 与 HorizontalAnchor=1 以写出 anchorCtr='0'",
            )
        suite.check(
            "XML: 图表数量与数据点数",
            len(c_xml["charts"]) == len(r_xml["charts"]),
            f"cand={len(c_xml['charts'])} vs ref={len(r_xml['charts'])}",
            category="ppt_xml",
            location="ppt/charts",
            expected=len(r_xml["charts"]),
            actual=len(c_xml["charts"]),
        )
        max_charts = max(len(c_xml["charts"]), len(r_xml["charts"]))
        for idx in range(max_charts):
            if idx >= len(c_xml["charts"]) or idx >= len(r_xml["charts"]):
                suite.check(
                    f"XML: 图表{idx + 1}存在性",
                    False,
                    f"cand_charts={len(c_xml['charts'])}, ref_charts={len(r_xml['charts'])}",
                    category="ppt_xml",
                    location=f"ppt/charts/chart{idx + 1}.xml",
                )
                continue
            cc, rc = c_xml["charts"][idx], r_xml["charts"][idx]
            suite.check(
                f"XML: 图表{idx + 1}系列与数据点",
                cc["series"] == rc["series"] and cc["chart_types"] == rc["chart_types"],
                f"cand={cc} vs ref={rc}",
                category="ppt_xml",
                location=rc["file"],
                expected=rc,
                actual=cc,
                hint="PPT 图表超过 4 行数据必须先在 ChartData.Activate() 后循环 Rows(5).Insert() 扩行",
            )
    elif not use_com:
        suite.check(
            "XML: 容器格式支持纯静态比对",
            False,
            f"cand={c_xml.get('container')}, ref={r_xml.get('container')}",
            category="ppt_xml",
            hint="旧版二进制 OLE2 (.ppt/.dps) 需移除 --no-com 参数使用 WPS COM 对比",
        )
        return

    if not use_com:
        suite.check(
            "XML: 母版背景填充类型",
            c_xml.get("master_bg_fill_type") == r_xml.get("master_bg_fill_type"),
            f"cand={c_xml.get('master_bg_fill_type')} vs ref={r_xml.get('master_bg_fill_type')}",
            category="ppt_xml",
            location="ppt/slideMasters/slideMaster1.xml",
            expected=r_xml.get("master_bg_fill_type"),
            actual=c_xml.get("master_bg_fill_type"),
        )
        for idx, (cs, rs) in enumerate(zip(c_xml["slides"], r_xml["slides"]), 1):
            suite.check(
                f"XML: 第{idx}页切换与跟随母版",
                cs.get("follow_master_bg") == rs.get("follow_master_bg")
                and cs.get("transition") == rs.get("transition"),
                f"cand_bg={cs.get('follow_master_bg')}, trans={cs.get('transition')} vs ref_bg={rs.get('follow_master_bg')}, trans={rs.get('transition')}",
                category="ppt_xml",
                location=f"slide{idx}",
            )
            c_s = sorted(cs.get("shapes", []), key=lambda x: (x["type"], x["text"][:40], x["left"], x["top"]))
            r_s = sorted(rs.get("shapes", []), key=lambda x: (x["type"], x["text"][:40], x["left"], x["top"]))
            c_s_norm = [{**x, "text": x["text"][:40]} for x in c_s]
            r_s_norm = [{**x, "text": x["text"][:40]} for x in r_s]
            sdiffs = _compare_ppt_shapes(
                c_s_norm,
                r_s_norm,
                keys=("type", "text", "has_table", "has_chart", "smartart_nodes", "align", "smallcaps", "font_fill_pattern"),
            )
            suite.check(
                f"XML: 第{idx}页图形坐标与文本样式",
                not sdiffs,
                "; ".join(sdiffs[:4]) if sdiffs else f"cnt={len(r_s)}",
                category="ppt_xml",
                location=f"slide{idx}",
                hint="无头模式下禁止用 ShapeRange.Align/Distribute，请直接用几何坐标赋值 Left/Top",
            )
        return

    if session is not None:
        c_com = _snapshot_ppt_com(cand_path, session=session)
        if ref_com_cache is not None and ref_path in ref_com_cache:
            r_com = ref_com_cache[ref_path]
        else:
            r_com = _snapshot_ppt_com(ref_path, session=session)
            if ref_com_cache is not None:
                ref_com_cache[ref_path] = r_com
    else:
        with wps_batch_session("ppt", with_window=False) as sess:
            c_com = _snapshot_ppt_com(cand_path, session=sess)
            r_com = _snapshot_ppt_com(ref_path, session=sess)

    suite.check(
        "COM: 幻灯片宽高与母版背景",
        abs(c_com["slide_w"] - r_com["slide_w"]) < 0.5
        and abs(c_com["slide_h"] - r_com["slide_h"]) < 0.5
        and c_com["master_bg_type"] == r_com["master_bg_type"],
        f"cand=({c_com['slide_w']}x{c_com['slide_h']}, bg={c_com['master_bg_type']}) vs ref=({r_com['slide_w']}x{r_com['slide_h']}, bg={r_com['master_bg_type']})",
        category="ppt_com",
        location="PageSetup/SlideMaster",
    )
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
            location=f"Slide({si})",
            hint="第1页常默认 FollowMasterBackground=0，改母版背景后需遍历所有页设 FollowMasterBackground=-1",
        )
        shape_diffs = _compare_ppt_shapes(cs["shapes"], rs["shapes"])
        suite.check(
            f"COM: 第{si}页图形结构(位置/文本/SmartArt/图表)",
            not shape_diffs,
            "; ".join(shape_diffs[:4]) if shape_diffs else f"cnt={len(rs['shapes'])}",
            category="ppt_com",
            location=f"Slide({si}).Shapes",
        )


def compare_documents(
    candidate_path: str,
    reference_path: str,
    *,
    use_com: bool = True,
    session: WPSBatchSession | None = None,
    ref_com_cache: dict[str, dict[str, Any]] | None = None,
) -> AuditSuite:
    """Compare candidate solution against a reference answer using COM + OOXML dual-track grading."""
    c_info = inspect_container(candidate_path)
    r_info = inspect_container(reference_path)
    if c_info["kind"] != r_info["kind"]:
        raise ValueError(f"Mismatched document types: {c_info['kind']} vs {r_info['kind']}")

    kind = c_info["kind"]
    suite = AuditSuite(f"REF DIFF ({kind.upper()}): {Path(candidate_path).name}")
    if kind == "word":
        _diff_word(suite, c_info["path"], r_info["path"], use_com=use_com, session=session, ref_com_cache=ref_com_cache)
    elif kind == "excel":
        _diff_excel(suite, c_info["path"], r_info["path"], use_com=use_com, session=session, ref_com_cache=ref_com_cache)
    elif kind == "ppt":
        _diff_ppt(suite, c_info["path"], r_info["path"], use_com=use_com, session=session, ref_com_cache=ref_com_cache)
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
