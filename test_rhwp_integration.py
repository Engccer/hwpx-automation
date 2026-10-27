"""실제 rhwp 바이너리로 hwpx_edit.py의 rhwp 축을 돈다(RHWP_INTEGRATION=1일 때만).

표본은 tests/fixtures/의 익명 양식이다(python-hwpx로 만든 가상 문서를 rhwp로 HWP 왕복).
"""
import contextlib
import io
import os
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


if __name__ == "__main__":
    unittest.main()
