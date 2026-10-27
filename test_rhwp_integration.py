"""실제 rhwp 바이너리로 hwpx_edit.py의 rhwp 축을 돈다(RHWP_INTEGRATION=1일 때만).

표본은 tests/fixtures/의 익명 양식이다(python-hwpx로 만든 가상 문서를 rhwp로 HWP 왕복).
"""
import contextlib
import inspect
import io
import os
import subprocess
import sys
from pathlib import Path
import shutil
import tempfile
import unittest

import hwpx_edit
import hwpx_gate

FIXTURES = Path(__file__).resolve().parent / "tests" / "fixtures"


def quiet_call(fn, *args, **kwargs):
    out = io.StringIO()
    with contextlib.redirect_stdout(out):
        res = fn(*args, **kwargs)
    return res, out.getvalue()


@unittest.skipUnless(os.environ.get("RHWP_INTEGRATION"), "RHWP_INTEGRATION=1일 때 실제 rhwp로 실행")
class FromHwpIntegration(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)
        self.hwp = self.dir / "sample-form.hwp"
        shutil.copyfile(FIXTURES / "sample-form.hwp", self.hwp)

    def test_fixture_converts_with_full_recall_and_valid_schema(self):
        res, out = quiet_call(hwpx_edit.hwp_to_hwpx, str(self.hwp))
        expected = "passed" if hwpx_edit.find_java() else "unverified"
        self.assertEqual((res.engine, res.status), ("rhwp", expected), out)
        ok, report = hwpx_gate.validate(res.path)
        self.assertTrue(ok, report)
        from hwpx_tomd import convert
        md = convert(res.path)
        self.assertEqual((md.recall, md.char_recall), (1.0, 1.0))
        self.assertIn("홍길동", md.markdown)


@unittest.skipUnless(os.environ.get("RHWP_INTEGRATION"), "RHWP_INTEGRATION=1일 때 실제 rhwp로 실행")
class PdfIntegration(unittest.TestCase):
    """--to-pdf 기본 rhwp: 표본 HWPX와 Pandoc 산출물 모두 빈 네모 0."""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)

    def assert_clean_pdf(self, hwpx):
        from rhwp_pdf import save_pdf
        report = save_pdf(hwpx, self.dir / "out.pdf")
        self.assertEqual(report["missing"], [])
        self.assertGreaterEqual(report["pages"], 1)

    def test_fixture_hwpx_renders_without_missing_glyphs(self):
        hwp = self.dir / "sample-form.hwp"
        shutil.copyfile(FIXTURES / "sample-form.hwp", hwp)
        res, _ = quiet_call(hwpx_edit.hwp_to_hwpx, str(hwp))
        self.assert_clean_pdf(res.path)

    def test_pandoc_output_renders_with_rhwp(self):
        try:
            import pypandoc_hwpx  # noqa: F401
        except ImportError:
            self.skipTest("pypandoc-hwpx 없음")
        import subprocess
        import sys
        md, hwpx = self.dir / "t.md", self.dir / "t.hwpx"
        md.write_text("# 시험 문서\n\n가상의 안내문입니다.\n\n| 항목 | 내용 |\n|---|---|\n| 장소 | 강당 |\n",
                      encoding="utf-8")
        script = Path(__file__).resolve().parent / "convert" / "hwpx_convert.py"
        subprocess.run([sys.executable, str(script), str(md), "-o", str(hwpx)],
                       check=True, capture_output=True)
        self.assert_clean_pdf(hwpx)


HAS_ALLOW_RHWP = "allow_rhwp" in inspect.signature(hwpx_gate.gate).parameters
EDIT = Path(__file__).resolve().parent / "hwpx_edit.py"


@unittest.skipUnless(os.environ.get("RHWP_INTEGRATION"), "RHWP_INTEGRATION=1일 때 실제 rhwp로 실행")
class GridIntegration(unittest.TestCase):
    """표본 HWP → 변환 → --grid-set-cell → --tables → 게이트, 그리고 --to-hwp 왕복."""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)
        hwp = self.dir / "sample-form.hwp"
        shutil.copyfile(FIXTURES / "sample-form.hwp", hwp)
        self.hwpx = Path(quiet_call(hwpx_edit.hwp_to_hwpx, str(hwp))[0].path)

    def cli(self, *argv):
        return subprocess.run([sys.executable, str(EDIT), *map(str, argv)], capture_output=True,
                              text=True, env=dict(os.environ, PYTHONIOENCODING="utf-8"))

    @unittest.skipUnless(HAS_ALLOW_RHWP, "hwpx_gate.gate(allow_rhwp=)는 세션 sign의 M3 통합 뒤")
    def test_grid_set_cell_then_tables_shows_value_and_gate_passes(self):
        out = self.dir / "filled.hwpx"
        r = self.cli(self.hwpx, "--grid-set-cell", "0,1,1", "김철수", "-o", out)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(r.stdout.strip().splitlines()[-1], "engine=rhwp status=passed")
        r = self.cli(out, "--tables")
        self.assertIn("[0,1,1] 김철수", r.stdout)

    def test_to_hwp_round_trips_through_jar(self):
        out = self.dir / "dist.hwp"
        r = self.cli(self.hwpx, "--to-hwp", "-o", out)
        self.assertEqual(r.returncode, 0, r.stderr)
        expected = "passed" if hwpx_edit.find_java() else "unverified"
        self.assertEqual(r.stdout.strip().splitlines()[-1], f"engine=rhwp status={expected}")
        self.assertTrue(out.read_bytes().startswith(b"\xd0\xcf\x11\xe0"))  # CFB(HWP5)


if __name__ == "__main__":
    unittest.main()
