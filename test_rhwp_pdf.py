import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zlib

from rhwp_pdf import embedded_fonts, missing_glyphs, save_pdf

HERE = Path(__file__).resolve().parent


def fixture_pdf(compress=False):
    """rhwp export-pdf 출력 형태를 흉내 낸 최소 PDF: LastResort 1종 + 일반 글꼴 1종."""
    cmap = (b"/CIDInit /ProcSet findresource begin\n"
            b"1 beginbfchar\n<0001> <2705>\nendbfchar\n"
            b"1 beginbfchar\n<0004> <D83DDCCC>\nendbfchar\n"
            b"1 beginbfrange\n<0002> <0003> <25B6>\nendbfrange\nend\n")
    head = b"<< /Length %d >>" % len(cmap)
    if compress:
        cmap = zlib.compress(cmap)
        head = b"<< /Length %d /Filter /FlateDecode >>" % len(cmap)
    return (b"%PDF-1.7\n"
            b"5 0 obj\n<<\n  /Type /Font\n  /Subtype /Type0\n"
            b"  /BaseFont /SLMAXX+LastResort-Identity-H\n  /Encoding /Identity-H\n"
            b"  /DescendantFonts [9 0 R]\n  /ToUnicode 6 0 R\n>>\nendobj\n"
            b"6 0 obj\n" + head + b"\nstream\n" + cmap + b"\nendstream\nendobj\n"
            b"9 0 obj\n<<\n  /Type /Font\n  /Subtype /CIDFontType0\n"
            b"  /BaseFont /SLMAXX+LastResort\n  /FontDescriptor 10 0 R\n>>\nendobj\n"
            b"7 0 obj\n<<\n  /Type /Font\n  /Subtype /Type0\n"
            b"  /BaseFont /DDLMFA+AppleSDGothicNeo-Regular-Identity-H\n>>\nendobj\n"
            b"8 0 obj\n<<\n  /Type /FontDescriptor\n  /FontName /DDLMFA+AppleSDGothicNeo-Regular\n>>\nendobj\n"
            b"%%EOF\n")


FAKE_RHWP = """#!{python}
import os, sys
mode = os.environ["FAKE_RHWP_MODE"]
out = sys.argv[sys.argv.index("-o") + 1]
if mode == "fail":
    sys.stderr.write("오류: 문서 로드 실패\\n")
    sys.exit(1)
if mode == "password" and sys.stdin.readline().strip() != "s3cret":
    sys.exit(1)
if mode == "nonpdf":
    open(out, "wb").write(b"not a PDF")
    sys.exit(0)
data = open(os.environ["FAKE_RHWP_PDF"], "rb").read()
open(out, "wb").write(data)
sys.stderr.write("LAYOUT_OVERFLOW: page=0, sec=0, col=0, para=9, overflow=4.6px\\n")
"""


class GlyphScanTests(unittest.TestCase):
    def test_lastresort_characters_are_listed(self):
        for compress in (False, True):
            with self.subTest(compress=compress):
                self.assertEqual(missing_glyphs(fixture_pdf(compress)),
                                 ["▶", "▷", "✅", "\U0001f4cc"])

    def test_lastresort_without_tounicode_is_still_flagged(self):
        pdf = fixture_pdf().replace(b"/ToUnicode 6 0 R", b"")
        self.assertEqual(missing_glyphs(pdf), ["�"])

    def test_multi_character_mapping_is_split_into_code_points(self):
        pdf = fixture_pdf().replace(b"<0004> <D83DDCCC>", b"<0004> <00660069>")
        self.assertEqual(missing_glyphs(pdf), ["f", "i", "\u25b6", "\u25b7", "\u2705"])

    def test_clean_pdf_has_no_missing_glyphs(self):
        clean = fixture_pdf().replace(b"LastResort", b"NanumGothic")
        self.assertEqual(missing_glyphs(clean), [])

    def test_embedded_fonts_drop_subset_prefix_and_encoding(self):
        self.assertEqual(embedded_fonts(fixture_pdf()),
                         ["AppleSDGothicNeo-Regular", "LastResort"])


@unittest.skipIf(sys.platform == "win32", "가짜 rhwp 실행 파일이 shebang에 의존")
class NonWindowsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        tmp = Path(self.tmp.name)
        bin_dir = tmp / "bin"
        bin_dir.mkdir()
        fake = bin_dir / "rhwp"
        fake.write_text(FAKE_RHWP.format(python=sys.executable), encoding="utf-8")
        fake.chmod(0o755)
        (tmp / "fixture.pdf").write_bytes(fixture_pdf())
        self.src = tmp / "in.hwpx"
        self.src.write_bytes(b"PK")
        self.out = tmp / "out" / "result.pdf"
        self.env = {"PATH": f"{bin_dir}{os.pathsep}{os.environ['PATH']}",
                    "FAKE_RHWP_PDF": str(tmp / "fixture.pdf")}
        self.saved = {k: os.environ.get(k) for k in [*self.env, "FAKE_RHWP_MODE"]}
        os.environ.update(self.env)

    def tearDown(self):
        for k, v in self.saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        self.tmp.cleanup()

    def test_success_reports_fonts_overflow_and_missing_glyphs(self):
        os.environ["FAKE_RHWP_MODE"] = "ok"
        report = save_pdf(self.src, self.out)
        self.assertTrue(self.out.read_bytes().startswith(b"%PDF-"))
        self.assertEqual(report["overflow"], 1)
        self.assertIn("✅", report["missing"])
        self.assertIn("AppleSDGothicNeo-Regular", report["fonts"])
        self.assertEqual(list(self.out.parent.iterdir()), [self.out])

    def test_failed_export_preserves_existing_pdf(self):
        os.environ["FAKE_RHWP_MODE"] = "fail"
        self.out.parent.mkdir()
        self.out.write_bytes(b"old PDF")
        with self.assertRaisesRegex(RuntimeError, "문서 로드 실패"):
            save_pdf(self.src, self.out)
        self.assertEqual(self.out.read_bytes(), b"old PDF")
        self.assertEqual(list(self.out.parent.iterdir()), [self.out])

    def test_scan_error_preserves_existing_pdf(self):
        import rhwp_pdf
        os.environ["FAKE_RHWP_MODE"] = "ok"
        self.out.parent.mkdir()
        self.out.write_bytes(b"old PDF")
        original = rhwp_pdf.missing_glyphs
        rhwp_pdf.missing_glyphs = lambda pdf: (_ for _ in ()).throw(ValueError("bad cmap"))
        try:
            with self.assertRaises(ValueError):
                save_pdf(self.src, self.out)
        finally:
            rhwp_pdf.missing_glyphs = original
        self.assertEqual(self.out.read_bytes(), b"old PDF")
        self.assertEqual(list(self.out.parent.iterdir()), [self.out])

    def test_non_pdf_output_is_rejected(self):
        os.environ["FAKE_RHWP_MODE"] = "nonpdf"
        with self.assertRaises(RuntimeError):
            save_pdf(self.src, self.out)
        self.assertFalse(self.out.exists())

    def test_password_goes_through_stdin(self):
        os.environ["FAKE_RHWP_MODE"] = "password"
        save_pdf(self.src, self.out, password="s3cret")
        self.assertTrue(self.out.exists())

    def test_missing_rhwp_explains_install(self):
        os.environ["PATH"] = str(Path(self.tmp.name) / "empty")
        with self.assertRaisesRegex(RuntimeError, "rhwp"):
            save_pdf(self.src, self.out)

    def test_cli_exits_2_and_keeps_pdf_when_glyphs_are_missing(self):
        env = dict(os.environ, FAKE_RHWP_MODE="ok", PYTHONIOENCODING="utf-8")
        r = subprocess.run([sys.executable, str(HERE / "hwpx_edit.py"), str(self.src),
                            "--to-pdf", "-o", str(self.out)],
                           capture_output=True, text=True, env=env)
        self.assertEqual(r.returncode, 2, r.stderr)
        self.assertIn("U+2705", r.stderr)
        self.assertTrue(self.out.exists())

    def test_check_env_accepts_path_java_for_shell_wrapper(self):
        fake_java = Path(self.tmp.name) / "bin" / "java"
        fake_java.write_text('#!/bin/sh\necho \'openjdk version "21.0.4"\' >&2\n', encoding="utf-8")
        fake_java.chmod(0o755)
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        env.pop("JAVA_HOME", None)
        r = subprocess.run([sys.executable, str(HERE / "hwpx_edit.py"), "--check-env"],
                           capture_output=True, text=True, env=env)
        ready = r.stdout.split("바로 사용 가능:", 1)[1].splitlines()[0]
        self.assertIn("HWP→HWPX 변환", ready)

    def test_check_env_reports_rhwp_for_pdf(self):
        r = subprocess.run([sys.executable, str(HERE / "hwpx_edit.py"), "--check-env"],
                           capture_output=True, text=True, env=dict(os.environ, PYTHONIOENCODING="utf-8"))
        tier4 = r.stdout.split("[Tier 4]", 1)[1]
        self.assertIn("[O] rhwp", tier4)

    def test_check_env_without_rhwp_shows_install_hint(self):
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PATH="/usr/bin:/bin")
        r = subprocess.run([sys.executable, str(HERE / "hwpx_edit.py"), "--check-env"],
                           capture_output=True, text=True, env=env)
        self.assertIn("[설치] rhwp", r.stdout)
        self.assertNotIn("PDF 변환(rhwp)", r.stdout.split("바로 사용 가능:", 1)[1].splitlines()[0])

    def test_check_env_ignores_cwd_bin_java_when_java_home_unset(self):
        cwd = Path(self.tmp.name) / "cwd"
        (cwd / "bin").mkdir(parents=True)
        decoy = cwd / "bin" / "java"
        decoy.write_text("#!/bin/sh\necho decoy >&2\n", encoding="utf-8")
        decoy.chmod(0o755)
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PATH="/nonexistent")
        env.pop("JAVA_HOME", None)
        r = subprocess.run([sys.executable, str(HERE / "hwpx_edit.py"), "--check-env"],
                           capture_output=True, text=True, env=env, cwd=cwd)
        self.assertIn("[설치] JDK", r.stdout)

    def test_cli_succeeds_on_clean_render(self):
        fixture = Path(os.environ["FAKE_RHWP_PDF"])
        fixture.write_bytes(fixture_pdf().replace(b"LastResort", b"NanumGothic"))
        env = dict(os.environ, FAKE_RHWP_MODE="ok", PYTHONIOENCODING="utf-8")
        r = subprocess.run([sys.executable, str(HERE / "hwpx_edit.py"), str(self.src),
                            "--to-pdf", "-o", str(self.out)],
                           capture_output=True, text=True, env=env)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("rhwp", r.stdout)


if __name__ == "__main__":
    unittest.main()
