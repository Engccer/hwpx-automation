import os
from pathlib import Path
import tempfile
import unittest

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


if __name__ == "__main__":
    unittest.main()
