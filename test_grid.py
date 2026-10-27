"""M5 rhwp 축(--tables·--fields·--fill-fields·--grid-set-cell·--to-hwp)을 가짜 rhwp로 본다.

플래그가 rhwp 인자로 맞게 옮겨지는지, .hwp 입력이 변환을 먼저 거치는지, rhwp 오류와
게이트 실패가 산출물 없이 전달되는지를 확인한다. 실제 바이너리는 test_rhwp_integration.py.
"""
import contextlib
import io
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock

import hwpx_edit
import hwpx_gate
from test_doc_promises import make_table_hwpx

FAKE_RHWP = r'''#!{python}
import json, os, shutil, sys
if sys.argv[1:] == ["--version"]:
    print("rhwp v0.8.7"); sys.exit(0)
args = sys.argv[1:]
with open(os.environ["FAKE_RHWP_LOG"], "a", encoding="utf-8") as fh:
    fh.write(json.dumps(args, ensure_ascii=False) + "\n")
mode = os.environ.get("FAKE_RHWP_MODE", "ok")
def out_of(flag="-o"):
    return args[args.index(flag) + 1]
def emit(**kw):
    print(json.dumps(dict(schemaVersion="1.0", **kw), ensure_ascii=False))
cmd = args[0] if args[0] != "edit" else "edit " + args[1]
if cmd == "export-hwpx":
    shutil.copyfile(os.environ["FAKE_RHWP_SRC"], args[2])
    emit(output=args[2], format="hwpx", verify=dict(identical=True))
elif cmd == "export-tables":
    emit(tables=[dict(index=0, section=0, paragraph=3, rows=2, cols=2, cells=[
        dict(row=0, col=0, rowSpan=1, colSpan=2, text="제목"),
        dict(row=1, col=0, rowSpan=1, colSpan=1, text="성명"),
        dict(row=1, col=1, rowSpan=1, colSpan=1, text="")])])
elif cmd == "fields":
    emit(fieldCount=1, fields=[dict(name="성명", value="", guide="이름을 쓰세요")])
elif cmd == "edit fill-fields":
    missing = ["학교"] if mode == "notfound" else []
    if "--dry-run" in args:
        emit(filledCount=1, notFound=missing, ambiguous=[], dryRun=True)
    else:
        shutil.copyfile(args[2], out_of())
        emit(filledCount=1, notFound=[], ambiguous=[], output=out_of())
elif cmd == "edit set-cell":
    if mode == "nested":
        sys.stderr.write("오류: 본문 최상위 표 5 번이 없습니다 (최상위 표 1개; 중첩 표는 v1 범위 밖).\n")
        sys.exit(1)
    shutil.copyfile(args[2], out_of())
    emit(oldText="", newText=args[args.index("--text") + 1], output=out_of(), overflow=[])
elif cmd == "convert":
    open(args[2], "wb").write(b"HWP Document File")
    emit(output=args[2], format="hwp5", verify=dict(identical=True))
else:
    sys.exit(2)
'''


@unittest.skipIf(sys.platform == "win32", "가짜 rhwp 실행 파일이 shebang에 의존")
class RhwpGridCommands(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)
        self.doc = self.dir / "doc.hwpx"
        make_table_hwpx(self.doc)
        self.hwp = self.dir / "form.hwp"
        self.hwp.write_bytes(b"HWP Document File")
        fake = self.dir / "rhwp"
        fake.write_text(FAKE_RHWP.replace("{python}", sys.executable), encoding="utf-8")
        fake.chmod(0o755)
        self.log = self.dir / "rhwp.log"
        env = {"HWPX_RHWP_BIN": str(fake), "FAKE_RHWP_SRC": str(self.doc),
               "FAKE_RHWP_LOG": str(self.log), "FAKE_RHWP_MODE": "ok"}
        for p in (mock.patch.dict(os.environ, env),
                  mock.patch.object(hwpx_edit, "find_java", return_value=None),
                  mock.patch.object(hwpx_gate, "validate", return_value=(True, "ok"))):
            p.start()
            self.addCleanup(p.stop)
        self.gate = mock.patch.object(hwpx_gate, "gate", return_value=hwpx_gate.GateReport())
        self.gate_mock = self.gate.start()
        self.addCleanup(self.gate.stop)

    def cli(self, *argv):
        out, err = io.StringIO(), io.StringIO()
        code = 0
        with mock.patch.object(sys, "argv", ["hwpx_edit.py", *map(str, argv)]), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                hwpx_edit.main()
            except SystemExit as exc:
                code = exc.code
        return code, out.getvalue(), err.getvalue()

    def calls(self):
        if not self.log.exists():
            return []
        return [json.loads(l) for l in self.log.read_text(encoding="utf-8").splitlines()]

    def test_tables_prints_grid_coordinates(self):
        code, out, err = self.cli(self.doc, "--tables")
        self.assertEqual(code, 0, err)
        self.assertEqual(self.calls()[0], ["export-tables", str(self.doc), "--json"])
        self.assertIn("[0,0,0] (병합 1×2) 제목", out)
        self.assertIn("[0,1,1]", out)
        self.assertIn("--set-cell", out)

    def test_grid_set_cell_maps_flags_and_gates_with_rhwp_allowance(self):
        out_path = self.dir / "o.hwpx"
        code, out, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "홍길동", "-o", out_path)
        self.assertEqual(code, 0, err)
        call = self.calls()[0]
        self.assertEqual(call[:3], ["edit", "set-cell", str(self.doc)])
        for flag, val in (("--table", "0"), ("--row", "1"), ("--col", "1"), ("--text", "홍길동")):
            self.assertEqual(call[call.index(flag) + 1], val)
        self.assertTrue(out_path.exists())
        _, kwargs = self.gate_mock.call_args
        self.assertEqual((kwargs["allow"], kwargs["allow_rhwp"], kwargs["expect_text_change"]),
                         (hwpx_gate.is_text_run_key, True, True))
        self.assertEqual(out.strip().splitlines()[-1], "engine=rhwp status=passed")

    def test_hwp_input_is_converted_before_grid_edit(self):
        code, out, err = self.cli(self.hwp, "--grid-set-cell", "0,1,1", "x")
        self.assertEqual(code, 0, err)
        self.assertEqual([c[0] for c in self.calls()], ["export-hwpx", "edit"])
        converted = self.dir / "_work-hwpx-automation" / "form.hwpx"
        self.assertEqual(self.calls()[1][2], str(converted))
        self.assertTrue((self.dir / "_work-hwpx-automation" / "form_편집.hwpx").exists())

    def test_nested_table_error_is_passed_through_with_set_cell_hint(self):
        os.environ["FAKE_RHWP_MODE"] = "nested"
        out_path = self.dir / "o.hwpx"
        code, _, err = self.cli(self.doc, "--grid-set-cell", "5,0,0", "x", "-o", out_path)
        self.assertEqual(code, 1)
        self.assertIn("중첩 표는 v1 범위 밖", err)
        self.assertIn("--set-cell", err)
        self.assertFalse(out_path.exists())

    def test_gate_failure_keeps_output_unpublished(self):
        failed = hwpx_gate.GateReport()
        failed.fail("사라진 요소 tbl×1")
        self.gate_mock.return_value = failed
        out_path = self.dir / "o.hwpx"
        code, _, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "x", "-o", out_path)
        self.assertEqual(code, 1)
        self.assertIn("게이트 실패", err)
        self.assertFalse(out_path.exists())
        self.assertTrue((self.dir / "o.rejected.hwpx").exists())

    def test_fill_fields_refuses_unknown_names_before_writing(self):
        os.environ["FAKE_RHWP_MODE"] = "notfound"
        out_path = self.dir / "o.hwpx"
        code, _, err = self.cli(self.doc, "--fill-fields", '{"학교":"x"}', "-o", out_path)
        self.assertEqual(code, 1)
        self.assertIn("학교", err)
        self.assertEqual(len(self.calls()), 1)
        self.assertIn("--dry-run", self.calls()[0])
        self.assertFalse(out_path.exists())

    def test_fill_fields_passes_data_and_writes(self):
        out_path = self.dir / "o.hwpx"
        code, out, err = self.cli(self.doc, "--fill-fields", '{"성명":"홍길동"}', "-o", out_path)
        self.assertEqual(code, 0, err)
        write = self.calls()[1]
        self.assertEqual(write[write.index("--data") + 1], '{"성명":"홍길동"}')
        self.assertNotIn("--dry-run", write)
        self.assertTrue(out_path.exists())
        self.assertEqual(out.strip().splitlines()[-1], "engine=rhwp status=passed")

    def test_fields_lists_names(self):
        code, out, _ = self.cli(self.doc, "--fields")
        self.assertEqual(code, 0)
        self.assertIn("성명", out)
        self.assertIn("이름을 쓰세요", out)

    def test_to_hwp_says_hwpx_is_canonical_and_is_unverified_without_jdk(self):
        out_path = self.dir / "o.hwp"
        code, out, err = self.cli(self.doc, "--to-hwp", "-o", out_path)
        self.assertEqual(code, 0, err)
        self.assertEqual(self.calls()[0][:2], ["convert", str(self.doc)])
        self.assertIn("--verify-pages", self.calls()[0])
        self.assertIn("정본은 HWPX", out)
        self.assertEqual(out.strip().splitlines()[-1], "engine=rhwp status=unverified")
        self.assertTrue(out_path.exists())

    def test_to_hwp_roundtrip_loss_fails_with_com_hint(self):
        lossy = self.dir / "lossy.hwpx"
        make_table_hwpx(lossy, rows=2)
        out_path = self.dir / "o.hwp"
        with mock.patch.object(hwpx_edit, "find_java", return_value="java"), \
                mock.patch.object(hwpx_edit, "_run_jar",
                                  side_effect=lambda s, o: Path(o).write_bytes(lossy.read_bytes())):
            code, _, err = self.cli(self.doc, "--to-hwp", "-o", out_path)
        self.assertEqual(code, 1)
        self.assertIn("왕복 대조 실패", err)
        self.assertIn("다른 이름 저장", err)
        self.assertFalse(out_path.exists())

    def test_to_hwp_rejects_hwp_input_and_non_hwp_output(self):
        code, _, err = self.cli(self.hwp, "--to-hwp")
        self.assertEqual(code, 1)
        self.assertIn("HWPX 입력만", err)
        code, _, err = self.cli(self.doc, "--to-hwp", "-o", self.dir / "x.hwpx")
        self.assertEqual(code, 1)


if __name__ == "__main__":
    unittest.main()
