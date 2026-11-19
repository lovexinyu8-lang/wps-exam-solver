"""Unit tests for scripts/rubric_grader.py and scripts/ref_diff.py (--no-com parity, 23 rules, batch & export)."""

from __future__ import annotations

import csv
import json
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

TESTS_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = TESTS_DIR.parent / "scripts"
for p in (str(TESTS_DIR), str(SCRIPTS_DIR)):
    if p not in sys.path:
        sys.path.insert(0, p)

import ooxml_patch
import ref_diff
import rubric_grader
from fixtures import (
    build_excel_reordered_fixture,
    build_ole2_dummy,
    build_ppt_fixture,
    build_word_fixture,
)


class TestRubricAndDiff(unittest.TestCase):
    def test_all_9_word_rules_pass_and_fail_no_com(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td_p = Path(td)
            w_ok = build_word_fixture(td_p / "ok.docx")
            ooxml_patch.patch_word_grad_fill(str(w_ok), "主标题文本", [(0, "accent2", None, None)])
            ooxml_patch.patch_word_glow(str(w_ok), "表头发光目标", sat_mod=175000)
            ooxml_patch.patch_word_table_theme_fill(str(w_ok), table_index=0, row_index=0, theme_fill="accent2", fill_hex="#ED7D31")
            ooxml_patch.patch_word_double_diagonal_header(str(w_ok))

            rubric = {
                "title": "Word 9 Rules",
                "kind": "word",
                "rules": [
                    {"id": "w1", "type": "word_page_setup", "params": {"section": 1, "sections_count": 1, "orientation": 0, "margins_cm": [2.5, 2.5, 2.5, 2.5]}},
                    {"id": "w2", "type": "word_style", "params": {"style_id": -2, "font_size": 16.0, "bold": True, "align": 1, "bottom_border": {"has_border": True, "line_width": 12}}},
                    {"id": "w3", "type": "word_paragraph_match", "params": {"text_contains": "正文内容段落", "style_name": "Normal", "align": 3, "font_size": 12.0, "c_indent": 2.0, "unit_before": 0.5}},
                    {"id": "w4", "type": "word_text_absent", "params": {"forbidden": ["请删除占位符"]}},
                    {"id": "w5", "type": "word_header_watermark", "params": {"section": 1, "link_to_previous": False, "has_field_substr": "TIME", "shape_count": 2, "watermark_text": "机密水印", "picture_no_washout": True, "bottom_border": {"has_border": True}, "behind_doc": True}},
                    {"id": "w6", "type": "word_numbering_xml", "params": {"ilvl": 1, "num_fmt_prefix": "chineseCounting", "lvl_text": "（%2）", "lvl_restart": 0}},
                    {"id": "w7", "type": "word_table", "params": {"table_index": 1, "rows": 2, "cols": 2, "align": 1, "heading_repeat": True, "has_wps_diagonals": True, "theme_fill": "accent2", "cell_checks": [{"row": 1, "col": 1, "bg_color_bgr": ooxml_patch.rgb_to_bgr_int(0xED, 0x7D, 0x31)}]}},
                    {"id": "w8", "type": "word_effects_xml", "params": {"has_w14_textFill": True, "has_w14_glow": True, "glow_sat_mod": 175000}},
                    {"id": "w9", "type": "word_chart", "params": {"chart_index": 1, "series_count": 1, "pt_count": 5, "title": "销售趋势", "chart_type": 51}},
                ],
            }
            rep_ok = rubric_grader.grade_with_rubric(str(w_ok), rubric, use_com=False)
            self.assertTrue(rep_ok["all_passed"], rep_ok["results"])

            # Explicit failure checks including lvlRestart=0 bug fix
            w_bad_restart = build_word_fixture(td_p / "bad_restart.docx", variant="fail_restart")
            rep_restart = rubric_grader.grade_with_rubric(str(w_bad_restart), {"kind": "word", "rules": [rubric["rules"][5]]}, use_com=False)
            self.assertFalse(rep_restart["all_passed"])

            w_bad_hdr = build_word_fixture(td_p / "bad_hdr.docx", variant="fail_header")
            rep_hdr = rubric_grader.grade_with_rubric(str(w_bad_hdr), {"kind": "word", "rules": [rubric["rules"][4]]}, use_com=False)
            self.assertFalse(rep_hdr["all_passed"])

            w_bad_txt = build_word_fixture(td_p / "bad_txt.docx", variant="forbidden_text")
            rep_txt = rubric_grader.grade_with_rubric(str(w_bad_txt), {"kind": "word", "rules": [rubric["rules"][3]]}, use_com=False)
            self.assertFalse(rep_txt["all_passed"])

    def test_all_7_excel_rules_pass_and_fail_no_com(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td_p = Path(td)
            x_ok = build_excel_reordered_fixture(td_p / "ok.xlsx")
            rubric = {
                "title": "Excel 7 Rules",
                "kind": "excel",
                "rules": [
                    {"id": "e1", "type": "excel_sheets", "params": {"exact_names": ["Alpha", "Beta"]}},
                    {"id": "e2", "type": "excel_range_formula", "params": {"sheet": "Alpha", "range": "B2", "formula_contains": "VLOOKUP("}},
                    {"id": "e3", "type": "excel_range_values", "params": {"sheet": "Alpha", "cells": {"B2": 100}}},
                    {"id": "e4", "type": "excel_number_format", "params": {"sheet": "Beta", "range": "A2:A3", "contains": "¥"}},
                    {"id": "e5", "type": "excel_conditional_formatting", "params": {"sheet": "Alpha", "sqref": "A2:H10", "formula_norm": "WEEKDAY($B2,2)>5", "interior_color_bgr": ooxml_patch.rgb_to_bgr_int(255, 192, 0)}},
                    {"id": "e6", "type": "excel_sort_state", "params": {"sheet": "Beta", "ref": "A1:A3", "conditions_count": 1}},
                    {"id": "e7", "type": "excel_pivot_table", "params": {"sheet": "Alpha", "row_fields": ["客户级别"], "data_fields_count": 1}},
                ],
            }
            rep_ok = rubric_grader.grade_with_rubric(str(x_ok), rubric, use_com=False)
            self.assertTrue(rep_ok["all_passed"], rep_ok["results"])

            x_bad_cf = build_excel_reordered_fixture(td_p / "bad_cf.xlsx", variant="fail_cf")
            rep_cf = rubric_grader.grade_with_rubric(str(x_bad_cf), {"kind": "excel", "rules": [rubric["rules"][4]]}, use_com=False)
            self.assertFalse(rep_cf["all_passed"])

            x_bad_fml = build_excel_reordered_fixture(td_p / "bad_fml.xlsx", variant="fail_formula")
            rep_fml = rubric_grader.grade_with_rubric(str(x_bad_fml), {"kind": "excel", "rules": [rubric["rules"][1]]}, use_com=False)
            self.assertFalse(rep_fml["all_passed"])

            x_bad_val = build_excel_reordered_fixture(td_p / "bad_val.xlsx", variant="fail_val")
            rep_val = rubric_grader.grade_with_rubric(str(x_bad_val), {"kind": "excel", "rules": [rubric["rules"][2]]}, use_com=False)
            self.assertFalse(rep_val["all_passed"])

            x_no_cache = build_excel_reordered_fixture(td_p / "no_cache.xlsx", variant="no_cached_val")
            rep_nc = rubric_grader.grade_with_rubric(str(x_no_cache), {"kind": "excel", "rules": [rubric["rules"][2]]}, use_com=False)
            self.assertFalse(rep_nc["all_passed"])
            self.assertIn("cached value missing", rep_nc["results"][0]["detail"])

    def test_all_7_ppt_rules_pass_and_fail_no_com(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td_p = Path(td)
            p_ok = build_ppt_fixture(td_p / "ok.pptx")
            rubric = {
                "title": "PPT 7 Rules",
                "kind": "ppt",
                "rules": [
                    {"id": "p1", "type": "ppt_page_setup", "params": {"slide_size": 7, "width_pt": 960.0, "height_pt": 540.0, "slides_count": 1}},
                    {"id": "p2", "type": "ppt_print_options", "params": {"prnWhat": "handouts6", "clrMode": "bw", "hiddenSlides": "1"}},
                    {"id": "p3", "type": "ppt_master_background", "params": {"master_fill_type": 6, "all_follow_master": True}},
                    {"id": "p4", "type": "ppt_slide_transition", "params": {"slides": [1], "entry_effect": 3948, "advance_on_time": True, "advance_time_range": [3.0, 10.0], "hidden": False}},
                    {"id": "p5", "type": "ppt_shape_match", "params": {"slide": 1, "name": "TitleBox", "right_aligned_to_slide": True, "top_aligned_to_slide": True, "align": 2, "smallcaps": True, "font_fill_pattern": 16, "shadow_visible": True, "body_pr_anchor": "t", "body_pr_anchor_ctr": "0", "hmerge_all_aligned": True}},
                    {"id": "p6", "type": "ppt_align_distribute", "params": {"slide": 1, "shape_type": 1, "top_aligned": True, "h_distributed": True}},
                    {"id": "p7", "type": "ppt_chart", "params": {"chart_index": 1, "series_count": 1, "pt_count": 5, "chart_type": 51, "has_title": True, "no_gridlines": True}},
                ],
            }
            rep_ok = rubric_grader.grade_with_rubric(str(p_ok), rubric, use_com=False)
            self.assertTrue(rep_ok["all_passed"], rep_ok["results"])

            # SmartArt node assertion
            sa_rule = {"id": "p_sa", "type": "ppt_shape_match", "params": {"slide": 1, "has_smartart": True, "smartart_nodes": ["步骤一", "步骤二"]}}
            self.assertTrue(rubric_grader.grade_with_rubric(str(p_ok), {"kind": "ppt", "rules": [sa_rule]}, use_com=False)["all_passed"])

            # Failure cases
            for var, ridx in (("fail_size", 0), ("fail_print", 1), ("fail_master_bg", 2), ("fail_anchor", 4), ("fail_align", 5)):
                p_bad = build_ppt_fixture(td_p / f"{var}.pptx", variant=var)
                rep_bad = rubric_grader.grade_with_rubric(str(p_bad), {"kind": "ppt", "rules": [rubric["rules"][ridx]]}, use_com=False)
                self.assertFalse(rep_bad["all_passed"], f"{var} should fail rule {ridx}")

    def test_ref_diff_structured_diagnostics_and_chart_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td_p = Path(td)
            w_ref = build_word_fixture(td_p / "ref.docx", variant="extra_chart")
            w_cand = build_word_fixture(td_p / "cand.docx", variant="fail_header")
            suite = ref_diff.compare_documents(str(w_cand), str(w_ref), use_com=False)
            self.assertFalse(suite.all_passed)
            d = suite.to_dict()
            # Chart 2 missing must be reported explicitly, not truncated by zip()
            chart2_fail = next((c for c in d["checks"] if "图表2" in c["name"] and not c["passed"]), None)
            self.assertIsNotNone(chart2_fail)
            hdr_fail = next((c for c in d["checks"] if "页眉图片" in c["name"] and not c["passed"]), None)
            self.assertIsNotNone(hdr_fail)
            self.assertIn("location", hdr_fail)
            self.assertIn("hint", hdr_fail)
            self.assertIn("建议:", suite.format_report())

            # OLE2 non-com produces explicit unsupported FAIL check
            ole = build_ole2_dummy(td_p / "cand.wps")
            ole_ref = build_ole2_dummy(td_p / "ref.wps")
            ole_suite = ref_diff.compare_documents(str(ole), str(ole_ref), use_com=False)
            self.assertFalse(ole_suite.all_passed)
            self.assertGreaterEqual(ole_suite.total_count, 1)

    def test_scaffold_rubric_no_com_across_all_kinds(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td_p = Path(td)
            w_ref = build_word_fixture(td_p / "w.docx")
            x_ref = build_excel_reordered_fixture(td_p / "x.xlsx")
            p_ref = build_ppt_fixture(td_p / "p.pptx")

            for path, expected_types in (
                (w_ref, {"word_page_setup", "word_style", "word_numbering_xml", "word_table", "word_chart"}),
                (x_ref, {"excel_sheets", "excel_range_formula", "excel_range_values", "excel_conditional_formatting", "excel_sort_state", "excel_pivot_table"}),
                (p_ref, {"ppt_page_setup", "ppt_print_options", "ppt_master_background", "ppt_slide_transition", "ppt_shape_match", "ppt_chart"}),
            ):
                rub = rubric_grader.scaffold_rubric(str(path), total_points=20.0, use_com=False)
                self.assertAlmostEqual(sum(r["points"] for r in rub["rules"]), 20.0, places=2)
                found_types = {r["type"] for r in rub["rules"]}
                self.assertTrue(expected_types.issubset(found_types), f"{path.name}: missing {expected_types - found_types}")
                # Self-grading with generated rubric should pass 100%
                rep = rubric_grader.grade_with_rubric(str(path), rub, use_com=False)
                self.assertTrue(rep["all_passed"], rep["results"])

    def test_batch_grade_single_ref_mirror_unmatched_and_export(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td_p = Path(td)
            cand_dir = td_p / "作答目录"
            cand_dir.mkdir()
            build_ppt_fixture(cand_dir / "202401_张三.pptx")
            build_ppt_fixture(cand_dir / "202402_李四.pptx", variant="fail_print")
            build_word_fixture(cand_dir / "未匹配_文档.docx")

            # Mode 1: Single reference file inside folder (N students vs 1 reference)
            ref_dir = td_p / "标准答案"
            ref_dir.mkdir()
            p_ref = build_ppt_fixture(ref_dir / "标准答案.pptx")

            progress_events = []
            rep = rubric_grader.batch_grade(
                str(cand_dir),
                str(ref_dir),
                use_com=False,
                progress_callback=lambda cur, tot, name: progress_events.append((cur, tot, name)),
            )
            self.assertEqual(rep["matched_files"], 2)
            self.assertEqual(rep["unmatched_files"], ["未匹配_文档.docx"])
            self.assertFalse(rep["all_passed"])
            self.assertEqual(progress_events[-1][0], 2)

            # Mode 2: Single rubric json passed directly as file
            rub_file = td_p / "only_ppt.rubric.json"
            rub_file.write_text(json.dumps(rubric_grader.scaffold_rubric(str(p_ref), use_com=False), ensure_ascii=False), encoding="utf-8")
            rep_single_rub = rubric_grader.batch_grade(str(cand_dir), str(rub_file), use_com=False)
            self.assertEqual(rep_single_rub["matched_files"], 2)

            # Mode 2b: Mirror directory matching (作答文件 -> 参考答案)
            mirror_cand = td_p / "mirror_cand" / "作答文件"
            mirror_ref = td_p / "mirror_ref" / "参考答案"
            mirror_cand.mkdir(parents=True)
            mirror_ref.mkdir(parents=True)
            build_excel_reordered_fixture(mirror_cand / "表格题.xlsx")
            build_excel_reordered_fixture(mirror_ref / "表格题.xlsx")
            rep_mirror = rubric_grader.batch_grade(str(mirror_cand.parent), str(mirror_ref.parent), use_com=False)
            self.assertEqual(rep_mirror["matched_files"], 1)
            self.assertTrue(rep_mirror["all_passed"])

            # Mode 3: CSV & XLSX export
            csv_out = td_p / "scores.csv"
            xlsx_out = td_p / "scores.xlsx"
            rubric_grader.export_batch_report(rep, str(csv_out))
            rubric_grader.export_batch_report(rep, str(xlsx_out))

            with open(csv_out, encoding="utf-8-sig", newline="") as fp:
                rows = list(csv.reader(fp))
            self.assertEqual(len(rows), 3)  # header + 2 files
            self.assertEqual(rows[1][0], "202401_张三.pptx")
            self.assertEqual(rows[1][6], "PASS")
            self.assertEqual(rows[2][6], "FAIL")

            wb = openpyxl.load_workbook(xlsx_out)
            try:
                self.assertIn("成绩汇总", wb.sheetnames)
                self.assertIn("失分明细", wb.sheetnames)
                self.assertIn("未匹配文件", wb.sheetnames)
                self.assertEqual(wb["成绩汇总"].max_row, 4)  # header + 2 + total
                self.assertGreaterEqual(wb["失分明细"].max_row, 2)
            finally:
                wb.close()


if __name__ == "__main__":
    unittest.main()
