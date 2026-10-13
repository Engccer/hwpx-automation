"""셀 전체 입력은 inline 제어 요소 뒤 기존 본문까지 대체해야 한다."""
import contextlib
import hashlib
import io
from pathlib import Path
import tempfile
import unittest
import zipfile

from lxml import etree

import hwpx_edit

HP = hwpx_edit.NS['hp']

class SetCellFullReplacementTests(unittest.TestCase):
    def test_child_tail_and_other_runs_replaced_preserving_styles(self):
        for control in ('tab', 'lineBreak', 'fwSpace'):
            with self.subTest(control=control), tempfile.TemporaryDirectory() as temp:
                source = Path(temp) / 'source.hwpx'
                output = Path(temp) / 'result.hwpx'
                xml = f'''<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="{HP}">
                <hp:p><hp:run><hp:tbl><hp:tr><hp:tc borderFillIDRef="7">
                <hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/>
                <hp:subList><hp:p paraPrIDRef="3"><hp:run charPrIDRef="4">
                <hp:t>기존 앞<hp:{control}/>기존 뒤</hp:t></hp:run>
                <hp:run charPrIDRef="5"><hp:t>다른 런<hp:tab/>다른 꼬리</hp:t></hp:run>
                </hp:p></hp:subList></hp:tc><hp:tc><hp:subList><hp:p><hp:run><hp:t>인접 셀 보존</hp:t></hp:run></hp:p></hp:subList></hp:tc>
                </hp:tr></hp:tbl></hp:run></hp:p></hs:sec>'''
                with zipfile.ZipFile(source, 'w') as archive:
                    archive.writestr('mimetype', 'application/hwp+zip')
                    archive.writestr('Contents/section0.xml', xml.encode())
                original = hashlib.sha256(source.read_bytes()).hexdigest()
                with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                    hwpx_edit.cmd_set_cell(str(source), 0, 0, 0, '새 내용', output=str(output))
                with zipfile.ZipFile(output) as archive:
                    root = etree.fromstring(archive.read('Contents/section0.xml'))
                cells = root.findall('.//hp:tc', hwpx_edit.NS)
                self.assertEqual(hwpx_edit.get_cell_text(cells[0]), '새 내용')
                self.assertEqual(hwpx_edit.get_cell_text(cells[1]), '인접 셀 보존')
                texts = cells[0].findall('.//hp:t', hwpx_edit.NS)
                self.assertEqual([hwpx_edit.t_full_text(t) for t in texts], ['새 내용', ''])
                self.assertTrue(all(len(t) == 0 for t in texts))
                self.assertEqual(cells[0].get('borderFillIDRef'), '7')
                self.assertEqual(cells[0].find('.//hp:p', hwpx_edit.NS).get('paraPrIDRef'), '3')
                self.assertEqual([r.get('charPrIDRef') for r in cells[0].findall('.//hp:run', hwpx_edit.NS)], ['4', '5'])
                self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), original)

if __name__ == '__main__':
    unittest.main()
