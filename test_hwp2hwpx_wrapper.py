"""convert/hwp2hwpx.sh: 두 번째 인자가 폴더면 그 안에 <입력 이름>.hwpx로 저장한다(이슈 #15).

가짜 java(JAVA_HOME/bin/java)가 마지막 인자 경로에 파일을 만들어 래퍼가 넘긴 출력 경로만 본다.
"""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

WRAPPER = Path(__file__).resolve().parent / 'convert' / 'hwp2hwpx.sh'
FAKE_JAVA = '#!/bin/sh\nfor last; do :; done\nprintf converted > "$last"\n'


@unittest.skipIf(sys.platform.startswith('win'), 'POSIX 셸 래퍼')
class WrapperOutputTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        java = self.dir / 'jdk' / 'bin' / 'java'
        java.parent.mkdir(parents=True)
        java.write_text(FAKE_JAVA, encoding='utf-8')
        java.chmod(0o755)
        self.env = dict(os.environ, JAVA_HOME=str(self.dir / 'jdk'))
        self.src = self.dir / '양식 샘플.hwp'
        self.src.write_bytes(b'hwp')

    def tearDown(self):
        self._tmp.cleanup()

    def run_wrapper(self, *args):
        return subprocess.run(['sh', str(WRAPPER), str(self.src), *args],
                              env=self.env, capture_output=True, text=True)

    def test_folder_output_saves_named_file_inside(self):
        out = self.dir / 'out dir'
        out.mkdir()
        for arg in (str(out), str(out) + '/'):
            with self.subTest(arg=arg):
                r = self.run_wrapper(arg)
                self.assertEqual(r.returncode, 0, r.stderr)
                self.assertEqual((out / '양식 샘플.hwpx').read_text(), 'converted')
                (out / '양식 샘플.hwpx').unlink()

    def test_file_output_is_used_as_is(self):
        out = self.dir / 'result.hwpx'
        r = self.run_wrapper(str(out))
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertTrue(out.is_file())

    def test_default_output_goes_to_work_folder(self):
        r = self.run_wrapper()
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertTrue((self.dir / '_work-hwpx-automation' / '양식 샘플.hwpx').is_file())


if __name__ == '__main__':
    unittest.main()
