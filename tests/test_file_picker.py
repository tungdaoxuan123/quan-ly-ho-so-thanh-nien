import tests
import subprocess
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import patch

from quan_ly_ho_so.workbook.manager import native_workbook_dialog


class WindowsFilePickerTests(unittest.TestCase):
    def test_windows_picker_from_web_worker_preserves_unicode_path(self):
        selected = "C:/Hồ sơ thanh niên/Danh sách 2027.xlsx"
        with patch("quan_ly_ho_so.workbook.manager.sys.platform", "win32"), \
             patch("quan_ly_ho_so.workbook.manager.subprocess.CREATE_NO_WINDOW", 0x08000000, create=True), \
             patch("quan_ly_ho_so.workbook.manager.subprocess.run") as run:
            run.return_value.stdout = selected
            with ThreadPoolExecutor(max_workers=1) as worker:
                self.assertEqual(worker.submit(native_workbook_dialog).result(), Path(selected))
                self.assertEqual(worker.submit(native_workbook_dialog).result(), Path(selected))
            self.assertEqual(run.call_count, 2)
            args, kwargs = run.call_args
            self.assertEqual(args[0][-1], "--choose-workbook")
            self.assertTrue(args[0][-2].endswith("run.py"))
            self.assertEqual(kwargs["encoding"], "utf-8-sig")
            self.assertTrue(kwargs["check"])

    def test_cancel_does_not_select_a_workbook(self):
        with patch("quan_ly_ho_so.workbook.manager.sys.platform", "win32"), \
             patch("quan_ly_ho_so.workbook.manager.subprocess.CREATE_NO_WINDOW", 0x08000000, create=True), \
             patch("quan_ly_ho_so.workbook.manager.subprocess.run") as run:
            run.return_value.stdout = ""
            self.assertIsNone(native_workbook_dialog())

    def test_packaged_picker_relaunches_exe_without_python_install(self):
        with patch("quan_ly_ho_so.workbook.manager.sys.platform", "win32"), \
             patch("quan_ly_ho_so.workbook.manager.sys.frozen", True, create=True), \
             patch("quan_ly_ho_so.workbook.manager.subprocess.CREATE_NO_WINDOW", 0x08000000, create=True), \
             patch("quan_ly_ho_so.workbook.manager.subprocess.run") as run:
            run.return_value.stdout = ""
            native_workbook_dialog()
            self.assertEqual(run.call_args.args[0], [__import__("sys").executable, "--choose-workbook"])

    def test_dialog_failure_is_handled(self):
        with patch("quan_ly_ho_so.workbook.manager.sys.platform", "win32"), \
             patch("quan_ly_ho_so.workbook.manager.subprocess.CREATE_NO_WINDOW", 0x08000000, create=True), \
             patch("quan_ly_ho_so.workbook.manager.subprocess.run", side_effect=subprocess.CalledProcessError(1, "powershell")), \
             self.assertLogs(level="ERROR"):
            self.assertIsNone(native_workbook_dialog())
