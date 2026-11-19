"""Unit tests for scripts/ooxml_patch.py."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

TESTS_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = TESTS_DIR.parent / "scripts"
for p in (str(TESTS_DIR), str(SCRIPTS_DIR)):
    if p not in sys.path:
        sys.path.insert(0, p)

import ooxml_patch
from fixtures import (
    build_excel_reordered_fixture,
    build_ole2_dummy,
    build_ppt_fixture,
    build_word_fixture,
)


class TestOOXMLPatch(unittest.TestCase):
    def test_excel_reordered_sheet_mapping_via_rels(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            xlsx = build_excel_reordered_fixture(Path(td) / "reordered.xlsx")
            audit = ooxml_patch.audit_excel_xml(str(xlsx))

            self.assertEqual([s["name"] for s in audit["sheets"]], ["Alpha", "Beta"])
            ws_by_name = {w["name"]: w for w in audit["worksheets"]}
            # Alpha -> sheet2.xml (conditional_formatting, no sort_state)
            self.assertEqual(ws_by_name["Alpha"]["file"], "xl/worksheets/sheet2.xml")
            self.assertEqual(len(ws_by_name["Alpha"]["conditional_formatting"]), 1)
            self.assertIsNone(ws_by_name["Alpha"]["sort_state"])
            # Beta -> sheet1.xml (sort_state, no conditional_formatting)
            self.assertEqual(ws_by_name["Beta"]["file"], "xl/worksheets/sheet1.xml")
            self.assertEqual(ws_by_name["Beta"]["sort_state"]["ref"], "A1:A3")
            self.assertEqual(ws_by_name["Beta"]["conditional_formatting"], [])
            # Pivot cache + sheet rels resolution
            self.assertEqual(len(audit["pivot_tables"]), 1)
            self.assertEqual(audit["pivot_tables"][0]["sheet"], "Alpha")
            self.assertEqual(audit["pivot_tables"][0]["row_fields"], ["客户级别"])

    def test_word_w14_patchers_cover_body_table_and_textbox(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            docx_path = build_word_fixture(Path(td) / "word.docx")

            # Body + Textbox grad fill
            cnt_grad_body = ooxml_patch.patch_word_grad_fill(
                str(docx_path),
                "主标题文本",
                [(0, "accent2", 25000, 75000), (100000, "accent2", 85000, None)],
            )
            cnt_grad_box = ooxml_patch.patch_word_grad_fill(
                str(docx_path),
                "文本框渐变目标",
                [(0, "accent1", None, None), (100000, "accent1", 70000, None)],
            )
            self.assertEqual(cnt_grad_body, 1)
            self.assertEqual(cnt_grad_box, 1)

            # Table cell glow
            cnt_glow_tbl = ooxml_patch.patch_word_glow(
                str(docx_path),
                "表头发光目标",
                radius_pt=12.0,
                sat_mod=175000,
            )
            self.assertEqual(cnt_glow_tbl, 1)

            # Theme fill & double diagonal header idempotency
            patched_cells = ooxml_patch.patch_word_table_theme_fill(
                str(docx_path), table_index=0, row_index=0, theme_fill="accent2", fill_hex="#ED7D31"
            )
            self.assertEqual(patched_cells, 2)
            ooxml_patch.patch_word_double_diagonal_header(str(docx_path))
            ooxml_patch.patch_word_double_diagonal_header(str(docx_path))  # idempotent

            audit = ooxml_patch.audit_word_xml(str(docx_path))
            self.assertTrue(audit["has_w14_textFill"])
            self.assertTrue(audit["has_w14_glow"])
            self.assertIn("175000", audit["glow_sat_mods"])
            self.assertEqual(audit["theme_fills"], ["accent2", "accent2"])
            self.assertTrue(audit["has_wps_diagonals"])

            with zipfile.ZipFile(docx_path, "r") as z:
                doc_xml = z.read("word/document.xml").decode("utf-8")
            self.assertEqual(doc_xml.count("wpsCustomData:diagonalParaType"), 1)

    def test_ppt_patchers_print_options_and_hmerge(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            pptx_path = build_ppt_fixture(Path(td) / "deck.pptx", variant="fail_hmerge")
            # Also strip print options first via variant
            mod_p = ooxml_patch.patch_ppt_print_options(
                str(pptx_path), prn_what="handouts4", clr_mode="gray", hidden_slides=False
            )
            self.assertIn("ppt/presProps.xml", mod_p)
            mod_h = ooxml_patch.patch_ppt_hmerge_align(str(pptx_path), slide_number=1, algn="ctr")
            self.assertIn("ppt/slides/slide1.xml", mod_h)

            audit = ooxml_patch.audit_ppt_xml(str(pptx_path))
            self.assertEqual(audit["prnPr"]["prnWhat"], "handouts4")
            self.assertEqual(audit["prnPr"]["clrMode"], "gray")
            self.assertEqual(audit["prnPr"]["hiddenSlides"], "0")
            self.assertEqual(audit["slides"][0]["hmerge_missing_algn"], 0)

    def test_non_zip_ole2_safety(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            ole_path = build_ole2_dummy(Path(td) / "legacy.wps")
            audit = ooxml_patch.audit_ooxml(str(ole_path))
            self.assertFalse(audit["ooxml_supported"])
            self.assertEqual(audit["container"], "ole2_or_binary")
            with self.assertRaises(ValueError):
                ooxml_patch.rewrite_zip(str(ole_path), {"a": lambda b: b})
            self.assertEqual(ooxml_patch.read_excel_static(str(ole_path)), {})


if __name__ == "__main__":
    unittest.main()
