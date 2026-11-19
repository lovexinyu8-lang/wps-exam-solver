"""Unit tests for scripts/wps_env.py (non-COM helpers & COMWorkerQueue)."""

from __future__ import annotations

import os
import sys
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock

TESTS_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = TESTS_DIR.parent / "scripts"
for p in (str(TESTS_DIR), str(SCRIPTS_DIR)):
    if p not in sys.path:
        sys.path.insert(0, p)

import wps_env
from fixtures import build_ole2_dummy, build_word_fixture


class TestWpsEnvUnit(unittest.TestCase):
    def test_infer_kind(self) -> None:
        self.assertEqual(wps_env.infer_kind("a.docx"), "word")
        self.assertEqual(wps_env.infer_kind("b.wps"), "word")
        self.assertEqual(wps_env.infer_kind("c.et"), "excel")
        self.assertEqual(wps_env.infer_kind("d.xlsx"), "excel")
        self.assertEqual(wps_env.infer_kind("e.pptx"), "ppt")
        self.assertEqual(wps_env.infer_kind("f.dps"), "ppt")
        with self.assertRaises(ValueError):
            wps_env.infer_kind("bad.txt")

    def test_inspect_container_and_lock_files(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td_p = Path(td)
            docx_p = build_word_fixture(td_p / "sample.docx")
            lock_p = td_p / "~$sample.docx"
            lock_p.write_text("lock", encoding="utf-8")

            info = wps_env.inspect_container(str(docx_p))
            self.assertEqual(info["container"], "ooxml")
            self.assertEqual(info["kind"], "word")
            self.assertEqual(info["lock_files"], [str(lock_p)])

            ole_p = build_ole2_dummy(td_p / "legacy.wps")
            self.assertEqual(wps_env.inspect_container(str(ole_p))["container"], "ole2")

            unk_p = td_p / "weird.rtf"
            unk_p.write_bytes(b"{\\rtf1 hello}")
            self.assertEqual(wps_env.inspect_container(str(unk_p))["container"], "unknown")

            with self.assertRaises(FileNotFoundError):
                wps_env.inspect_container(str(td_p / "missing.docx"))

    def test_safe_backup_and_restore(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            target = Path(td) / "file.docx"
            target.write_bytes(b"ORIGINAL")
            bak = wps_env.safe_backup(str(target))
            self.assertTrue(os.path.exists(bak))
            target.write_bytes(b"CORRUPTED")
            wps_env.restore_backup(bak, str(target))
            self.assertEqual(target.read_bytes(), b"ORIGINAL")

    def test_check_write_lock_error_message(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            target = Path(td) / "locked.docx"
            target.write_bytes(b"DATA")
            (Path(td) / "~$locked.docx").write_text("x", encoding="utf-8")
            with mock.patch("builtins.open", side_effect=PermissionError("locked")):
                with self.assertRaises(PermissionError) as ctx:
                    wps_env.check_write_lock(str(target))
                self.assertIn("~$locked.docx", str(ctx.exception))

    def test_com_worker_queue_serial_order_and_single_thread(self) -> None:
        observed_threads: list[int] = []
        order: list[int] = []

        with wps_env.COMWorkerQueue() as q:
            futs = []
            for idx in range(5):
                def _task(i: int = idx) -> int:
                    observed_threads.append(threading.get_ident())
                    order.append(i)
                    return i * 2

                futs.append(q.submit(_task))
            results = [f.result(timeout=5) for f in futs]

            def _boom() -> None:
                raise ZeroDivisionError("boom")

            err_fut = q.submit(_boom)
            with self.assertRaises(ZeroDivisionError):
                err_fut.result(timeout=5)

        self.assertEqual(results, [0, 2, 4, 6, 8])
        self.assertEqual(order, [0, 1, 2, 3, 4])
        self.assertEqual(len(set(observed_threads)), 1)
        self.assertNotEqual(observed_threads[0], threading.get_ident())
        with self.assertRaises(RuntimeError):
            q.submit(lambda: 1)

    def test_com_thread_scope_runs_cleanly(self) -> None:
        with wps_env.com_thread_scope():
            pass


if __name__ == "__main__":
    unittest.main()
