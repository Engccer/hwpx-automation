"""문서(SKILL.md·reference)가 약속한 동작을 코드가 지키는지 보는 시험.

모두 임시 디렉터리에서 돌고, 한컴 COM·Pandoc·네트워크를 부르지 않는다
(COM·Pandoc 호출부는 가짜 객체로 바꾼다).
"""
import contextlib
import io
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock
import zipfile

from lxml import etree

try:
    import pymupdf  # noqa: F401
    _HAS_PYMUPDF = True
except ImportError:
    _HAS_PYMUPDF = False

import hwpx_com
import hwpx_edit

sys.path.insert(0, str(Path(__file__).resolve().parent / "convert"))
import hwpx_convert  # noqa: E402

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS = {"hp": HP}


def make_table_hwpx(path, rows=4, cols=2):
    """python-hwpx로 rows×cols 표 하나가 든 실제 HWPX를 만든다."""
    from hwpx.document import HwpxDocument
    doc = HwpxDocument.new()
    table = doc.add_table(rows, cols)
    for r in range(rows):
        for c in range(cols):
            table.set_cell_text(r, c, f"R{r}C{c}")
    doc.save_to_path(str(path))


def read_table(path):
    with zipfile.ZipFile(path) as zf:
        name = [n for n in zf.namelist() if "section" in n.lower() and n.endswith(".xml")][0]
        root = etree.fromstring(zf.read(name))
    return root.find(".//hp:tbl", NS)


def rewrite_section(path, fn):
    with zipfile.ZipFile(path) as zf:
        items = [(i, zf.read(i.filename)) for i in zf.infolist()]
    name = [i.filename for i, _ in items if "section" in i.filename.lower() and i.filename.endswith(".xml")][0]
    with zipfile.ZipFile(path, "w") as zf:
        for info, data in items:
            if info.filename == name:
                root = etree.fromstring(data)
                fn(root)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zf.writestr(info, data)


def quiet(fn, *args, **kwargs):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        try:
            fn(*args, **kwargs)
            code = 0
        except SystemExit as exc:
            code = exc.code if isinstance(exc.code, int) else 1
    return code, out.getvalue(), err.getvalue()


class DeleteRowsKeepsTableConsistent(unittest.TestCase):
    """SKILL.md 구조적 편집 필수 규칙 2·3: rowAddr 순차, rowCnt == 실제 tr 수."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.src = Path(self.tmp.name) / "t.hwpx"
        make_table_hwpx(self.src)
        self.out = Path(self.tmp.name) / "out.hwpx"

    def tearDown(self):
        self.tmp.cleanup()

    def assert_consistent(self, tbl):
        rows = tbl.findall("hp:tr", NS)
        self.assertEqual(int(tbl.get("rowCnt")), len(rows))
        for i, tr in enumerate(rows):
            for addr in tr.findall("hp:tc/hp:cellAddr", NS):
                self.assertEqual(int(addr.get("rowAddr")), i)

    def test_delete_middle_rows(self):
        code, _, err = quiet(hwpx_edit.cmd_delete_rows, str(self.src), 0, [1, 2], str(self.out))
        self.assertEqual(code, 0, err)
        tbl = read_table(self.out)
        self.assert_consistent(tbl)
        texts = [t.text for t in tbl.iter(f"{{{HP}}}t")]
        self.assertEqual(texts, ["R0C0", "R0C1", "R3C0", "R3C1"])

    def test_delete_empty_trailing_rows(self):
        def blank_last_two(root):
            rows = root.find(".//hp:tbl", NS).findall("hp:tr", NS)
            for tr in rows[2:]:
                for t in tr.iter(f"{{{HP}}}t"):
                    t.text = ""
        rewrite_section(self.src, blank_last_two)
        code, _, err = quiet(hwpx_edit.cmd_delete_empty_rows, str(self.src), 0, str(self.out))
        self.assertEqual(code, 0, err)
        tbl = read_table(self.out)
        self.assertEqual(len(tbl.findall("hp:tr", NS)), 2)
        self.assert_consistent(tbl)

    def test_refuses_row_inside_vertical_merge(self):
        def merge_col0_rows_1_2(root):
            rows = root.find(".//hp:tbl", NS).findall("hp:tr", NS)
            anchor = rows[1].findall("hp:tc", NS)[0]
            anchor.find("hp:cellSpan", NS).set("rowSpan", "2")
            rows[2].remove(rows[2].findall("hp:tc", NS)[0])
        rewrite_section(self.src, merge_col0_rows_1_2)
        for victims in ([2], [1]):
            with self.subTest(victims=victims):
                code, _, err = quiet(hwpx_edit.cmd_delete_rows, str(self.src), 0, victims, str(self.out))
                self.assertNotEqual(code, 0)
                self.assertIn("병합", err)
                self.assertFalse(self.out.exists())


    def test_empty_rows_refuse_vertical_merge(self):
        def merge_and_blank(root):
            rows = root.find(".//hp:tbl", NS).findall("hp:tr", NS)
            anchor = rows[2].findall("hp:tc", NS)[0]
            anchor.find("hp:cellSpan", NS).set("rowSpan", "2")
            rows[3].remove(rows[3].findall("hp:tc", NS)[0])
            for t in rows[3].iter(f"{{{HP}}}t"):
                t.text = ""
        rewrite_section(self.src, merge_and_blank)
        code, _, err = quiet(hwpx_edit.cmd_delete_empty_rows, str(self.src), 0, str(self.out))
        self.assertNotEqual(code, 0)
        self.assertIn("병합", err)
        self.assertFalse(self.out.exists())


class DeleteRowsBesideMerges(unittest.TestCase):
    """병합과 무관한 행은 병합이 있는 표에서도 지워진다."""

    def run_case(self, merge, victims, remaining):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "t.hwpx"; out = Path(tmp) / "o.hwpx"
            make_table_hwpx(src, rows=5, cols=2)

            def apply(root):
                rows = root.find(".//hp:tbl", NS).findall("hp:tr", NS)
                start, span, col = merge
                cell = rows[start].findall("hp:tc", NS)[col]
                sp = cell.find("hp:cellSpan", NS)
                sp.set("rowSpan" if span[0] == "r" else "colSpan", str(span[1]))
                if span[0] == "r":
                    for r in range(start + 1, start + span[1]):
                        rows[r].remove(rows[r].findall("hp:tc", NS)[col])
                else:
                    rows[start].remove(rows[start].findall("hp:tc", NS)[col + 1])
            rewrite_section(src, apply)
            code, _, err = quiet(hwpx_edit.cmd_delete_rows, str(src), 0, victims, str(out))
            self.assertEqual(code, 0, err)
            tbl = read_table(out)
            rows = tbl.findall("hp:tr", NS)
            self.assertEqual(int(tbl.get("rowCnt")), len(rows))
            firsts = []
            for i, tr in enumerate(rows):
                for addr in tr.findall("hp:tc/hp:cellAddr", NS):
                    self.assertEqual(int(addr.get("rowAddr")), i)
                firsts.append(next(tr.iter(f"{{{HP}}}t")).text)
            self.assertEqual(firsts, remaining)

    def test_horizontal_merge_only(self):
        self.run_case((1, ("c", 2), 0), [1], ["R0C0", "R2C0", "R3C0", "R4C0"])

    def test_rows_outside_vertical_merge(self):
        self.run_case((0, ("r", 2), 0), [3], ["R0C0", "R1C1", "R2C0", "R4C0"])
        self.run_case((3, ("r", 2), 0), [0], ["R1C0", "R2C0", "R3C0", "R4C1"])
        self.run_case((1, ("r", 2), 0), [3], ["R0C0", "R1C0", "R2C1", "R4C0"])


class WindowsJavaDetection(unittest.TestCase):
    """--check-env Tier 2(Windows)는 hwp2hwpx.bat과 같은 순서로 java를 찾는다."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def fake_jdk(self, name):
        exe = self.root / name / "bin" / "java.exe"
        exe.parent.mkdir(parents=True)
        exe.write_text("")
        return exe

    def test_java_home_env(self):
        exe = self.fake_jdk("jdk")
        found = hwpx_edit.find_windows_java(
            {"JAVA_HOME": str(exe.parent.parent)}, adoptium_glob=str(self.root / "none*"),
            which=lambda _: None)
        self.assertEqual(found, str(exe))

    def test_adoptium_then_path(self):
        exe = self.fake_jdk("jdk-21.0.4")
        found = hwpx_edit.find_windows_java(
            {}, adoptium_glob=str(self.root / "jdk-21*"), which=lambda _: "C:/other/java.exe")
        self.assertEqual(found, str(exe))
        found = hwpx_edit.find_windows_java(
            {}, adoptium_glob=str(self.root / "none*"), which=lambda _: "C:/other/java.exe")
        self.assertEqual(found, "C:/other/java.exe")

    def test_adoptium_picks_highest_like_bat(self):
        self.fake_jdk("jdk-21.0.1")
        newer = self.fake_jdk("jdk-21.0.4")
        found = hwpx_edit.find_windows_java(
            {}, adoptium_glob=str(self.root / "jdk-21*"), which=lambda _: None)
        self.assertEqual(found, str(newer))

    def test_missing(self):
        self.assertIsNone(hwpx_edit.find_windows_java(
            {"JAVA_HOME": str(self.root / "nope")}, adoptium_glob=str(self.root / "none*"),
            which=lambda _: None))

    def test_check_env_reports_ready_on_windows(self):
        exe = self.fake_jdk("jdk")
        env = dict(os.environ, JAVA_HOME=str(exe.parent.parent))
        with mock.patch.object(hwpx_edit.sys, "platform", "win32"), \
                mock.patch.object(hwpx_edit.shutil, "which", lambda *a, **k: None), \
                mock.patch.dict(hwpx_edit.os.environ, env, clear=True):
            _, out, _ = quiet(hwpx_edit.cmd_check_env)
        self.assertIn("[O] JDK", out)
        self.assertNotIn("JAVA_HOME을 수정", out)


class NormalizeRefusesZeroCharInput(unittest.TestCase):
    """--normalize 본문 자수 보존 검증은 COM이 입력을 0자로 읽는 경우(Pandoc 생성물)를 막는다."""

    class FakeHwp:
        def __init__(self, texts):
            self.texts = list(texts)
            self.saved = []
            self.quit_called = False

        def save_as(self, path, format="HWPX"):
            self.saved.append(path)
            Path(path).write_bytes(b"PK")
            return True

        def quit(self):
            self.quit_called = True

    def run_normalize(self, texts):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        src = Path(tmp.name) / "in.hwpx"
        src.write_bytes(b"PK")
        fake = self.FakeHwp(texts)
        with mock.patch.object(hwpx_com, "create_hwp", return_value=fake), \
                mock.patch.object(hwpx_com, "open_for_com"), \
                mock.patch.object(hwpx_com, "extract_full_text", side_effect=lambda _: fake.texts.pop(0)):
            code, out, err = quiet(hwpx_com.cmd_normalize, str(src), str(Path(tmp.name) / "o.hwpx"))
        return fake, code, out, err

    def test_zero_char_input_is_not_saved(self):
        fake, code, out, err = self.run_normalize(["", ""])
        self.assertNotEqual(code, 0)
        self.assertEqual(fake.saved, [])
        self.assertTrue(fake.quit_called)
        self.assertNotIn("정규화 완료", out)

    def test_whitespace_only_input_is_not_saved(self):
        fake, code, _, _ = self.run_normalize(["\r\n", "\r\n"])
        self.assertNotEqual(code, 0)
        self.assertEqual(fake.saved, [])

    def test_whitespace_only_output_fails(self):
        fake, code, out, _ = self.run_normalize(["본문", " \r\n"])
        self.assertNotEqual(code, 0)
        self.assertNotIn("정규화 완료", out)

    def test_normal_input_passes(self):
        fake, code, out, _ = self.run_normalize(["본문", "본문"])
        self.assertEqual(code, 0)
        self.assertEqual(len(fake.saved), 1)
        self.assertIn("정규화 완료", out)


class ConvertDocxWithQuoteFix(unittest.TestCase):
    """hwpx_convert.py는 DOCX 입력도 기본 설정(따옴표 보호 켬)으로 변환한다."""

    def test_binary_input_is_passed_through(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "doc.docx"
            src.write_bytes(b"PK\x03\x04\xff\xfe\x00binary")
            calls = []
            fake = mock.Mock()
            fake.convert_to_hwpx.side_effect = lambda i, o, r: calls.append(i)
            with mock.patch.object(hwpx_convert, "PandocToHwpx", fake):
                code, _, err = quiet(lambda: hwpx_convert.convert_file(
                    str(src), str(Path(tmp) / "o.hwpx"), None, "hwpx"))
            self.assertEqual(calls, [str(src)], err)

    def test_markdown_quotes_still_protected(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "a.md"
            src.write_text('그는 "안녕"이라고 했다.', encoding="utf-8")
            seen = []
            fake = mock.Mock()
            fake.convert_to_hwpx.side_effect = lambda i, o, r: seen.append(Path(i).read_text(encoding="utf-8"))
            with mock.patch.object(hwpx_convert, "PandocToHwpx", fake), \
                    mock.patch.object(hwpx_convert, "_restore_quotes_in_hwpx"):
                quiet(lambda: hwpx_convert.convert_file(str(src), str(Path(tmp) / "o.hwpx"), None, "hwpx"))
            self.assertEqual(len(seen), 1)
            self.assertNotIn('"', seen[0])


class FindReplaceRoundTrip(unittest.TestCase):
    """hwpx_edit.py --find/--replace는 설치된 python-hwpx 세대와 무관하게 치환 결과를
    파일에 저장해야 한다(6.x에서 저장 API 개명으로 메모리 치환만 되고 파일은 그대로
    남던 회귀의 방지. 이 시험은 설치된 세대 하나에서만 돈다)."""

    def test_replacement_reaches_output_file(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        src = Path(tmp.name) / "t.hwpx"
        make_table_hwpx(src)
        out = Path(tmp.name) / "o.hwpx"
        code, so, se = quiet(hwpx_edit.cmd_find_replace, str(src), "R1C1", "홍길동", str(out))
        self.assertEqual(code, 0, se)
        self.assertTrue(out.exists(), "치환 결과 파일이 없다")
        texts = [t.text or "" for t in read_table(out).iter("{%s}t" % HP)]
        self.assertIn("홍길동", texts)
        self.assertNotIn("R1C1", texts)
        self.assertIn("1건 치환", so)


class SignPdfImageGate(unittest.TestCase):
    """hwpx_sign.py --pdf: 검증 PDF에 이미지 객체가 하나도 없으면 서명이 렌더되지
    않은 문서이므로 [OK]가 아니라 실패로 끝나야 한다."""

    def _pdf(self, path, with_image):
        import pymupdf
        doc = pymupdf.open()
        page = doc.new_page()
        page.insert_text((72, 72), "( sign )")
        if with_image:
            pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 20, 10), False)
            pix.clear_with(0)
            page.insert_image(pymupdf.Rect(100, 100, 150, 130), pixmap=pix)
        doc.save(str(path))
        doc.close()

    @unittest.skipUnless(_HAS_PYMUPDF, "PyMuPDF 없음")
    def test_count_pdf_images(self):
        import hwpx_sign
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        a, b = Path(tmp.name) / "a.pdf", Path(tmp.name) / "b.pdf"
        self._pdf(a, False)
        self._pdf(b, True)
        self.assertEqual(hwpx_sign.count_pdf_images(str(a)), 0)
        self.assertEqual(hwpx_sign.count_pdf_images(str(b)), 1)

    def _run_main(self, image_count):
        import hwpx_sign
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        doc = Path(tmp.name) / "d.hwpx"
        doc.write_bytes(b"PK")
        img = Path(tmp.name) / "s.png"
        img.write_bytes(b"png")
        argv = ["hwpx_sign.py", str(doc), "--image", str(img), "--pdf"]
        with mock.patch.object(sys, "argv", argv), \
                mock.patch.object(hwpx_sign, "com_insert"), \
                mock.patch.object(hwpx_sign, "to_floating", return_value=(0, 0)), \
                mock.patch.object(hwpx_sign.subprocess, "run", return_value=mock.Mock(returncode=0)), \
                mock.patch.object(hwpx_sign, "count_pdf_images", return_value=image_count):
            out, err = io.StringIO(), io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                code = hwpx_sign.main()
        return code, out.getvalue(), err.getvalue()

    def test_pdf_without_image_fails(self):
        code, out, err = self._run_main(0)
        self.assertEqual(code, 1)
        self.assertIn("[ERROR]", err)
        self.assertNotIn("[OK] 검증 PDF", out)

    def test_pdf_with_image_passes(self):
        code, out, err = self._run_main(1)
        self.assertEqual(code, 0)
        self.assertIn("[OK] 검증 PDF", out)

    def test_missing_pymupdf_is_reported_not_fatal(self):
        code, out, err = self._run_main(None)
        self.assertEqual(code, 0)
        self.assertIn("PyMuPDF", out + err)

    def test_locked_output_file_gives_taskkill_hint(self):
        import hwpx_sign
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        doc = Path(tmp.name) / "d.hwpx"
        doc.write_bytes(b"PK")
        img = Path(tmp.name) / "s.png"
        img.write_bytes(b"png")
        argv = ["hwpx_sign.py", str(doc), "--image", str(img)]
        with mock.patch.object(sys, "argv", argv), \
                mock.patch("shutil.copyfile", side_effect=PermissionError(13, "locked")), \
                mock.patch.object(hwpx_sign, "com_insert") as com:
            out, err = io.StringIO(), io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                code = hwpx_sign.main()
        self.assertEqual(code, 1)
        self.assertIn("taskkill", err.getvalue())
        com.assert_not_called()


if __name__ == "__main__":
    unittest.main()
