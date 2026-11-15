"""--to-md --underline 회귀 시험 (hwpx-tomd 0.5.0 밑줄 옵션 연결).

익명 표본 sample-fields.hwpx에 밑줄 charPr를 하나 더하고 마지막 문단을
'작성일 <밑줄 글> 서명 <밑줄 빈칸>'으로 바꾼 합성 HWPX를 쓴다.
"""
import contextlib
import inspect
import io
import os
from pathlib import Path
import re
import tempfile
import unittest
import zipfile

import hwpx_edit

try:
    from hwpx_tomd import convert
    HAS_UNDERLINE = "underline" in inspect.signature(convert).parameters
except ImportError:
    HAS_UNDERLINE = False

FIXTURE = Path(__file__).parent / "tests" / "fixtures" / "sample-fields.hwpx"
LAST_RUN = re.compile(r'<hp:run charPrIDRef="0"><hp:t>[^<]*2026[^<]*</hp:t></hp:run>')


def make_underlined(dest):
    with zipfile.ZipFile(FIXTURE) as src, zipfile.ZipFile(dest, "w") as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "Contents/header.xml":
                h = data.decode("utf-8")
                base = re.search(r'<hh:charPr id="0".*?</hh:charPr>', h, re.S).group(0)
                ul = base.replace('id="0"', 'id="7"', 1).replace(
                    '<hh:underline type="NONE"', '<hh:underline type="BOTTOM"')
                h = h.replace('<hh:charProperties itemCnt="7">', '<hh:charProperties itemCnt="8">')
                h = h.replace("</hh:charProperties>", ul + "</hh:charProperties>")
                data = h.encode("utf-8")
            elif item.filename == "Contents/section0.xml":
                s = data.decode("utf-8")
                runs = ('<hp:run charPrIDRef="0"><hp:t>작성일 </hp:t></hp:run>'
                        '<hp:run charPrIDRef="7"><hp:t>가상학교</hp:t></hp:run>'
                        '<hp:run charPrIDRef="0"><hp:t> 서명 </hp:t></hp:run>'
                        '<hp:run charPrIDRef="7"><hp:t>      </hp:t></hp:run>')
                s, n = LAST_RUN.subn(runs, s)
                assert n == 1
                data = s.encode("utf-8")
            out.writestr(item, data)


@unittest.skipUnless(HAS_UNDERLINE, "hwpx-tomd 0.5.0 이상 필요")
class UnderlineToMdTests(unittest.TestCase):
    def to_md(self, path, **kw):
        out_md = str(path) + (".ul" if kw.get("underline") else "") + ".md"
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            hwpx_edit.cmd_to_md(str(path), out_md, **kw)
        return Path(out_md).read_text(encoding="utf-8")

    def test_underline_flag_marks_text_and_blank(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc = Path(tmp) / "밑줄 표본.hwpx"
            make_underlined(doc)
            md = self.to_md(doc, underline=True)
            self.assertIn("<u>가상학교</u>", md)
            self.assertRegex(md, r"서명 (\\_)+")

    def test_default_output_has_no_underline_marks(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc = Path(tmp) / "밑줄 표본.hwpx"
            make_underlined(doc)
            md = self.to_md(doc)
            self.assertNotIn("<u>", md)
            self.assertNotIn("\\_", md)
            self.assertIn("작성일 가상학교 서명", md)

    def test_document_without_underline_is_unchanged_by_flag(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc = Path(tmp) / "표본.hwpx"
            doc.write_bytes(FIXTURE.read_bytes())
            self.assertEqual(self.to_md(doc), self.to_md(doc, underline=True))


if __name__ == "__main__":
    unittest.main()
