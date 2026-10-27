import contextlib
import io
import os
from pathlib import Path
import tempfile
import unittest
from unittest import mock

import hwpx_edit
from pdf_export import save_pdf


class FakeHwp:
    def __init__(self, mode=0x20000, fail=False, invalid=False):
        self.mode = mode
        self.fail = fail
        self.invalid = invalid

    def SetMessageBoxMode(self, mode):
        previous = self.mode
        nibble = mode & 0xF0000
        if nibble:
            self.mode = (self.mode & ~0xF0000) | (0 if nibble == 0xF0000 else nibble)
        return previous

    def SaveAs(self, path, format, arg):
        assert self.mode & 0xF0000 == 0x10000
        assert not os.path.exists(path)
        if self.fail:
            raise RuntimeError("export failed")
        Path(path).write_bytes(b"not a PDF" if self.invalid else b"%PDF-1.7\nfixture")
        return True


class PdfExportTests(unittest.TestCase):
    def test_success_restores_mode_and_replaces_only_after_export(self):
        for mode in (0, 0x20000, 0x20010):
            with self.subTest(mode=mode), tempfile.TemporaryDirectory() as tmp:
                output = Path(tmp) / "result.pdf"
                output.write_bytes(b"old PDF")
                hwp = FakeHwp(mode)
                save_pdf(hwp, output)
                self.assertEqual(hwp.mode, mode)
                self.assertTrue(output.read_bytes().startswith(b"%PDF-"))
                self.assertEqual(list(Path(tmp).iterdir()), [output])

    def test_failed_export_preserves_existing_pdf_and_restores_mode(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "result.pdf"
            output.write_bytes(b"old PDF")
            hwp = FakeHwp(fail=True)
            with self.assertRaises(RuntimeError):
                save_pdf(hwp, output)
            self.assertEqual(output.read_bytes(), b"old PDF")
            self.assertEqual(hwp.mode, 0x20000)

    def test_non_pdf_output_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "result.pdf"
            hwp = FakeHwp(invalid=True)
            with self.assertRaises(RuntimeError):
                save_pdf(hwp, output)
            self.assertFalse(output.exists())
            self.assertEqual(hwp.mode, 0x20000)


class OpenableHwp(FakeHwp):
    def Open(self, path, fmt, arg):
        return True


class PdfBackendRouting(unittest.TestCase):
    """--to-pdf는 모든 OS에서 rhwp가 기본이고 COM은 --pdf-backend com으로만 탄다."""

    def run_pdf(self, platform, backend):
        out, err = io.StringIO(), io.StringIO()
        with tempfile.TemporaryDirectory() as tmp, \
                mock.patch.object(hwpx_edit.sys, "platform", platform), \
                mock.patch.object(hwpx_edit, "cmd_to_pdf_rhwp") as rhwp, \
                mock.patch.object(hwpx_edit, "create_hwp_com",
                                  return_value=(None, OpenableHwp())) as com, \
                mock.patch.object(hwpx_edit, "close_hwp_com"), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = 0
            try:
                hwpx_edit.cmd_to_pdf(os.path.join(tmp, "a.hwpx"), os.path.join(tmp, "a.pdf"),
                                     backend=backend)
            except SystemExit as exc:
                code = exc.code
        return code, rhwp, com, out.getvalue(), err.getvalue()

    def test_default_is_rhwp_even_on_windows(self):
        for platform in ("win32", "darwin", "linux"):
            with self.subTest(platform=platform):
                code, rhwp, com, _, _ = self.run_pdf(platform, "rhwp")
                self.assertEqual(code, 0)
                rhwp.assert_called_once()
                com.assert_not_called()

    def test_cli_without_flag_uses_rhwp_on_windows(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "a.hwpx")
            Path(src).write_bytes(b"PK")
            with mock.patch.object(hwpx_edit.sys, "platform", "win32"), \
                    mock.patch.object(hwpx_edit.sys, "argv", ["hwpx_edit.py", src, "--to-pdf"]), \
                    mock.patch.object(hwpx_edit, "cmd_to_pdf_rhwp") as rhwp, \
                    mock.patch.object(hwpx_edit, "create_hwp_com") as com, \
                    contextlib.redirect_stdout(io.StringIO()):
                hwpx_edit.main()
        rhwp.assert_called_once()
        com.assert_not_called()

    def test_com_backend_on_windows_uses_hancom(self):
        code, rhwp, com, out, _ = self.run_pdf("win32", "com")
        self.assertEqual(code, 0)
        com.assert_called_once()
        rhwp.assert_not_called()
        self.assertEqual(out.strip().splitlines()[-1], "engine=com")

    def test_com_backend_off_windows_is_refused(self):
        code, rhwp, com, _, err = self.run_pdf("darwin", "com")
        self.assertEqual(code, 1)
        self.assertIn("Windows", err)
        com.assert_not_called()
        rhwp.assert_not_called()


if __name__ == "__main__":
    unittest.main()
