"""--find/--replace: 0건은 경고·exit 1, run 경계 걸침 진단, --allow-zero, --cross-run (이슈 #11)."""
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile

from hwpx.document import HwpxDocument

HERE = Path(__file__).resolve().parent
CLI = HERE / 'hwpx_edit.py'

CROSS_PARA = (
    '<hp:p id="77" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
    '<hp:run charPrIDRef="0"><hp:t>    다) 기대효과: </hp:t></hp:run>'
    '<hp:run charPrIDRef="1"><hp:t>AAA를 통해 BBB를 안내하고 CCC를 홍보할</hp:t></hp:run>'
    '<hp:run charPrIDRef="0"><hp:t> 수 있으며 DDD 활동을 통해 EEE 신장</hp:t></hp:run>'
    '<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000"'
    ' baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/></hp:linesegarray></hp:p>'
)
SAME_STYLE_PARA = (
    '<hp:p id="78" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
    '<hp:run charPrIDRef="0"><hp:t>같은 모양 앞</hp:t></hp:run>'
    '<hp:run charPrIDRef="0"><hp:t>뒤 이어짐</hp:t></hp:run></hp:p>'
)


def make_doc(path, *, table_text=None):
    """python-hwpx 새 문서의 자리표 문단을 run 경계 걸침 문단 둘로 바꾼다."""
    doc = HwpxDocument.new()
    doc.add_paragraph('PLACEHOLDER')
    if table_text is not None:
        doc.add_table(1, 1).cell(0, 0).text = table_text
    base = Path(path).with_suffix('.base.hwpx')
    doc.save_to_path(str(base))
    with zipfile.ZipFile(base) as zf:
        names = zf.namelist()
        files = {n: zf.read(n) for n in names}
    sec = files['Contents/section0.xml'].decode('utf-8')
    sec, n = re.subn(r'<hp:p [^>]*><hp:run charPrIDRef="0"><hp:t>PLACEHOLDER</hp:t></hp:run></hp:p>',
                     lambda _m: CROSS_PARA + SAME_STYLE_PARA, sec)
    assert n == 1
    files['Contents/section0.xml'] = sec.encode('utf-8')
    with zipfile.ZipFile(path, 'w') as out:
        for name in names:
            out.writestr(name, files[name],
                         zipfile.ZIP_STORED if name == 'mimetype' else zipfile.ZIP_DEFLATED)


def section(path):
    with zipfile.ZipFile(path) as zf:
        return zf.read('Contents/section0.xml').decode('utf-8')


def para(xml, pid):
    return re.search(rf'<hp:p id="{pid}".*?</hp:p>', xml, re.S).group(0)


class FindReplaceTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.src = self.dir / 'src.hwpx'
        self.out = self.dir / 'out.hwpx'

    def tearDown(self):
        self._tmp.cleanup()

    def run_cli(self, *args):
        return subprocess.run([sys.executable, str(CLI), str(self.src), *args, '-o', str(self.out)],
                              capture_output=True, text=True, encoding='utf-8')

    def test_within_one_run_replaces(self):
        make_doc(self.src)
        r = self.run_cli('--find', 'CCC를 홍보할', '--replace', 'Y')
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn('1건 치환 완료', r.stdout)
        self.assertIn('BBB를 안내하고 Y</hp:t>', section(self.out))

    def test_cross_run_zero_warns_diagnoses_and_exits_nonzero_without_output(self):
        make_doc(self.src)
        r = self.run_cli('--find', 'CCC를 홍보할 수 있으며', '--replace', 'X')
        self.assertEqual(r.returncode, 1)
        self.assertFalse(self.out.exists())
        self.assertIn('치환이 0건', r.stderr)
        self.assertIn('run 2~3에 걸쳐 있음 (charPrIDRef 1/0)', r.stderr)
        self.assertIn('--cross-run', r.stderr)
        self.assertIn('출력 파일을 만들지 않았습니다', r.stderr)

    def test_same_char_shape_split_is_also_not_matched(self):
        # 기본 치환의 실제 범위: 글자 모양이 같아도 run이 나뉘면 못 찾는다(문서에 적은 범위).
        make_doc(self.src)
        r = self.run_cli('--find', '앞뒤', '--replace', 'X')
        self.assertEqual(r.returncode, 1)
        self.assertIn('run 1~2에 걸쳐 있음 (charPrIDRef 0)', r.stderr)

    def test_not_found_anywhere(self):
        make_doc(self.src)
        r = self.run_cli('--find', '없는 말', '--replace', 'X')
        self.assertEqual(r.returncode, 1)
        self.assertIn('문서에서 찾지 못했습니다', r.stderr)
        self.assertFalse(self.out.exists())

    def test_allow_zero_saves_and_exits_zero(self):
        make_doc(self.src)
        r = self.run_cli('--find', '없는 말', '--replace', 'X', '--allow-zero')
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn('치환이 0건', r.stderr)
        self.assertTrue(self.out.exists())
        self.assertIn('engine=lxml', r.stdout.strip().splitlines()[-1])

    def test_cross_run_replaces_keeps_outside_text_and_drops_lineseg(self):
        make_doc(self.src)
        r = self.run_cli('--find', 'CCC를 홍보할 수 있으며', '--replace', 'X', '--cross-run')
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn('1건 치환 완료', r.stdout)
        p = para(section(self.out), 77)
        self.assertIn('<hp:run charPrIDRef="1"><hp:t>AAA를 통해 BBB를 안내하고 X</hp:t></hp:run>', p)
        self.assertIn('<hp:run charPrIDRef="0"><hp:t> DDD 활동을 통해 EEE 신장</hp:t></hp:run>', p)
        self.assertNotIn('linesegarray', p)

    def test_cross_run_does_not_double_replace_table_cell(self):
        # everywhere가 표 칸까지 바꾸므로 lxml 2차 표 패스를 건너뛰어야 한다(A→AB가 ABB가 되지 않게).
        make_doc(self.src, table_text='표 안 A')
        r = self.run_cli('--find', '표 안 A', '--replace', '표 안 AB', '--cross-run')
        self.assertEqual(r.returncode, 0, r.stderr)
        xml = section(self.out)
        self.assertIn('표 안 AB', xml)
        self.assertNotIn('표 안 ABB', xml)


if __name__ == '__main__':
    unittest.main()
