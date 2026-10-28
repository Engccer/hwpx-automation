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
from test_doc_promises import make_table_hwpx, rewrite_section


def section_text(path):
    import zipfile
    with zipfile.ZipFile(path) as zf:
        return zf.read("Contents/section0.xml").decode("utf-8")

FAKE_RHWP = r'''#!{python}
import json, os, re, shutil, sys, zipfile
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
def rewrite(src, dst, change):
    # rhwp처럼 모든 hp:p id를 0부터 다시 매기고, change로 section 글자를 바꿔 쓴다
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"Contents/section\d+\.xml", item.filename):
                text = data.decode("utf-8")
                n = iter(range(10 ** 6))
                text = re.sub(r'(<hp:p\b[^>]*?\bid=")\d+(")', lambda m: m.group(1) + str(next(n)) + m.group(2), text)
                text = change(text)
                if mode == "blank_other":   # 대상 밖 글자 공백화
                    text = text.replace("꼬리 문단", "", 1).replace("<hp:t>항목</hp:t>", "<hp:t></hp:t>", 1)
                if mode == "drop_para":     # 대상 밖 문단 삭제(문단 수 감소)
                    text = re.sub(r'<hp:p [^>]*>(?:(?!<hp:p ).)*?(꼬리 문단|작성일)(?:(?!<hp:p ).)*?</hp:p>', "", text, count=1, flags=re.S)
                if mode == "secpr":         # 대상 밖 구역 속성 소실
                    text = re.sub(r'(<hp:secPr\b[^>]*?) textDirection="[^"]*"', r"\1", text, count=1)
                data = text.encode("utf-8")
            zout.writestr(item, data)
cmd = args[0] if args[0] != "edit" else "edit " + args[1]
if cmd == "export-hwpx":
    shutil.copyfile(os.environ["FAKE_RHWP_SRC"], args[2])
    emit(output=args[2], format="hwpx", verify=dict(identical=True))
elif cmd == "export-tables":
    emit(tables=[dict(index=0, section=0, paragraph=3, rows=2, cols=2, cells=[
        dict(row=0, col=0, rowSpan=1, colSpan=2, text="제목"),
        dict(row=1, col=0, rowSpan=1, colSpan=1, text="성명"),
        dict(row=1, col=1, rowSpan=1, colSpan=1, text="첫줄\n  [0,9,9] 가짜", nested=[dict(rows=1)])])])
elif cmd == "fields":
    emit(fieldCount=1, fields=[dict(name="성명", value="", guide="이름을 쓰세요")],
         textSecurity=dict(status="suspicious" if mode == "suspicious" else "clean"))
elif cmd == "edit fill-fields":
    missing = ["학교"] if mode == "notfound" else []
    confusable = [dict(name="성명", lookalikes=["셩명"])] if mode == "confusable" else []
    if "--dry-run" in args:
        emit(filledCount=1, notFound=missing, ambiguous=[], confusable=confusable, dryRun=True)
    else:
        data = json.loads(args[args.index("--data") + 1])
        (key, value), = data.items()
        k = int(re.fullmatch(r".*\[(\d+)\]", key).group(1)) if "[" in key else 0
        def fill(text):
            spots = list(re.finditer(r'(name="성명"[^>]*?) dirty="0"((?:(?!</hp:ctrl>).)*</hp:ctrl><hp:t>)([^<]*)(</hp:t>)', text, re.S))
            m = spots[k]
            return text[:m.start()] + m.group(1) + ' dirty="1"' + m.group(2) + value + m.group(4) + text[m.end():]
        rewrite(args[2], out_of(), fill)
        emit(filledCount=0 if mode == "zero_filled" else 1, notFound=[], ambiguous=[], output=out_of())
elif cmd == "edit set-cell":
    if mode == "nested":
        sys.stderr.write("오류: 본문 최상위 표 5 번이 없습니다 (최상위 표 1개; 중첩 표는 v1 범위 밖).\n")
        sys.exit(1)
    row, col = args[args.index("--row") + 1], args[args.index("--col") + 1]
    if mode == "wrong_cell":
        row, col = "1", "0"
    new = args[args.index("--text") + 1]
    rewrite(args[2], out_of(), lambda t: t.replace(f"<hp:t>R{row}C{col}</hp:t>", f"<hp:t>{new}</hp:t>", 1))
    emit(oldText=f"R{args[args.index('--row') + 1]}C{args[args.index('--col') + 1]}", newText=new,
         output=out_of(), overflow=[])
elif cmd == "convert":
    open(args[2], "wb").write(b"HWP Document File")
    emit(output=args[2], format="hwp5", verify=dict(identical=True))
else:
    sys.exit(2)
'''

FIELDS = Path(__file__).resolve().parent / "tests" / "fixtures" / "sample-fields.hwpx"


def make_form(path):
    # 머리 문단 · 2×2 표(R0C0…) · 꼬리 문단. 격자 편집 대상 밖 손실을 볼 문단이 있다
    from hwpx.document import HwpxDocument
    doc = HwpxDocument.new()
    doc.add_paragraph("머리 문단")
    table = doc.add_table(2, 2)
    for r in range(2):
        for c in range(2):
            table.set_cell_text(r, c, f"R{r}C{c}")
    doc.add_paragraph("꼬리 문단")
    doc.save_to_path(str(path))


@unittest.skipIf(sys.platform == "win32", "가짜 rhwp 실행 파일이 shebang에 의존")
class RhwpGridCommands(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)
        self.doc = self.dir / "doc.hwpx"
        make_form(self.doc)
        self.form = self.dir / "fields.hwpx"
        self.form.write_bytes(FIELDS.read_bytes())
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

    def patch_gate(self, **kw):
        """게이트 자체를 바꿔 끼우는 시험만 쓴다(나머지는 실제 게이트가 가짜 rhwp 산출물을 본다)."""
        patcher = mock.patch.object(hwpx_gate, "gate", **kw)
        self.addCleanup(patcher.stop)
        return patcher.start()

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
        self.assertIn("<hp:t>홍길동</hp:t>", section_text(out_path))
        self.assertIn("p id 재부여", out)
        self.assertEqual(out.strip().splitlines()[-1], "engine=rhwp status=passed")

    def test_grid_set_cell_rejects_changes_outside_target(self):
        for mode, needle in (("blank_other", "대상 밖"), ("drop_para", "대상 밖"), ("secpr", "대상 밖"),
                             ("wrong_cell", "대상 밖")):
            with self.subTest(mode=mode):
                os.environ["FAKE_RHWP_MODE"] = mode
                out_path = self.dir / f"{mode}.hwpx"
                code, _, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "홍길동", "-o", out_path)
                self.assertEqual(code, 1, err)
                self.assertIn(needle, err)
                self.assertFalse(out_path.exists())
                self.assertEqual(len(list(self.dir.glob(f"{mode}.rejected-*.hwpx"))), 1)

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

    def test_gate_failure_keeps_previous_output(self):
        failed = hwpx_gate.GateReport()
        failed.fail("사라진 요소 tbl×1")
        self.patch_gate(return_value=failed)
        out_path = self.dir / "o.hwpx"
        out_path.write_bytes(b"previous")
        code, _, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "x", "-o", out_path)
        self.assertEqual(code, 1)
        self.assertIn("게이트 실패", err)
        self.assertEqual(out_path.read_bytes(), b"previous")
        self.assertIn("거부된 산출물 보존", err)
        self.assertEqual(len(list(self.dir.glob("o.rejected-*.hwpx"))), 1)

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
        code, out, err = self.cli(self.form, "--fill-fields", '{"성명[1]":"김철수"}', "-o", out_path)
        self.assertEqual(code, 0, err)
        write = self.calls()[1]
        self.assertEqual(write[write.index("--data") + 1], '{"성명[1]":"김철수"}')
        self.assertNotIn("--dry-run", write)
        self.assertIn("<hp:t>김철수</hp:t>", section_text(out_path))
        self.assertIn("<hp:t>값2</hp:t>", section_text(out_path))
        self.assertEqual(out.strip().splitlines()[-1], "engine=rhwp status=passed")

    def test_fill_fields_rejects_changes_outside_target(self):
        for mode in ("blank_other", "drop_para"):
            with self.subTest(mode=mode):
                os.environ["FAKE_RHWP_MODE"] = mode
                out_path = self.dir / f"{mode}.hwpx"
                code, _, err = self.cli(self.form, "--fill-fields", '{"성명[1]":"김철수"}', "-o", out_path)
                self.assertEqual(code, 1, err)
                self.assertIn("대상 밖", err)
                self.assertFalse(out_path.exists())

    def test_fill_fields_wrong_occurrence_is_rejected(self):
        # 사용자는 첫째(성명[0])를 원했는데 산출물은 둘째가 바뀌었다고 보는 경우와 같다
        out_path = self.dir / "o.hwpx"
        with mock.patch.object(hwpx_edit, "_field_edits",
                               return_value=[hwpx_gate.FieldEdit("성명", "김철수", 0)]):
            code, _, err = self.cli(self.form, "--fill-fields", '{"성명[1]":"김철수"}', "-o", out_path)
        self.assertEqual(code, 1)
        self.assertFalse(out_path.exists())

    def test_fill_fields_write_envelope_mismatch_is_checked_before_publishing(self):
        os.environ["FAKE_RHWP_MODE"] = "zero_filled"
        out_path = self.dir / "o.hwpx"
        out_path.write_bytes(b"previous")
        code, _, err = self.cli(self.form, "--fill-fields", '{"성명[1]":"김철수"}', "-o", out_path)
        self.assertEqual(code, 1)
        self.assertIn("사전 점검과 다릅니다", err)
        self.assertEqual(out_path.read_bytes(), b"previous")
        self.assertEqual(len(list(self.dir.glob("o.rejected-*.hwpx"))), 1)

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
        self.assertEqual(len(list(self.dir.glob("o.rejected-*.hwp"))), 1)

    def test_to_hwp_rejects_hwp_input_and_non_hwp_output(self):
        code, _, err = self.cli(self.hwp, "--to-hwp")
        self.assertEqual(code, 1)
        self.assertIn("HWPX 입력만", err)
        self.assertEqual(self.calls(), [])
        code, _, err = self.cli(self.doc, "--to-hwp", "-o", self.dir / "x.hwpx")
        self.assertEqual(code, 1)


    def test_to_hwp_cannot_mix_with_edit_commands(self):
        code, _, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "x", "--to-hwp", "-o", self.dir / "b.hwp")
        self.assertEqual(code, 1)
        self.assertIn("함께 쓸 수 없습니다", err)
        self.assertEqual(self.calls(), [])

    def test_to_hwp_jar_failure_with_jdk_is_rejected(self):
        out_path = self.dir / "o.hwp"
        with mock.patch.object(hwpx_edit, "find_java", return_value="java"), \
                mock.patch.object(hwpx_edit, "_run_jar", side_effect=RuntimeError("JAR 변환 실패(exit 1)")):
            code, _, err = self.cli(self.doc, "--to-hwp", "-o", out_path)
        self.assertEqual(code, 1)
        self.assertIn("JAR 변환 실패", err)
        self.assertFalse(out_path.exists())
        self.assertEqual(len(list(self.dir.glob("o.rejected-*.hwp"))), 1)

    def test_to_hwp_chart_loss_fails_in_cross_check(self):
        def add_chart(root):
            p = root.find(".//{http://www.hancom.co.kr/hwpml/2011/paragraph}p")
            etree.SubElement(p, "{http://www.hancom.co.kr/hwpml/2011/paragraph}chart")
        from lxml import etree
        plain = self.dir / "plain.hwpx"
        plain.write_bytes(self.doc.read_bytes())
        rewrite_section(self.doc, add_chart)
        out_path = self.dir / "o.hwp"
        with mock.patch.object(hwpx_edit, "find_java", return_value="java"), \
                mock.patch.object(hwpx_edit, "_run_jar",
                                  side_effect=lambda s, o: Path(o).write_bytes(plain.read_bytes())):
            code, _, err = self.cli(self.doc, "--to-hwp", "-o", out_path)
        self.assertEqual(code, 1)
        self.assertIn("chart 1→0", err)
        self.assertFalse(out_path.exists())

    def test_fill_fields_refuses_confusable_names(self):
        os.environ["FAKE_RHWP_MODE"] = "confusable"
        code, _, err = self.cli(self.doc, "--fill-fields", '{"성명":"x"}', "-o", self.dir / "o.hwpx")
        self.assertEqual(code, 1)
        self.assertIn("닮은 이름", err)
        self.assertEqual(len(self.calls()), 1)

    def test_rhwp_edits_gate_by_position_without_callbacks(self):
        gate = self.patch_gate(return_value=hwpx_gate.GateReport())
        self.cli(self.form, "--fill-fields", '{"성명[1]":"x"}', "-o", self.dir / "f.hwpx")
        self.cli(self.doc, "--grid-set-cell", "0,1,1", "y", "-o", self.dir / "g.hwpx")
        (_, f_kw), (_, g_kw) = gate.call_args_list
        self.assertEqual(f_kw, dict(allow_rhwp=True, edits=[hwpx_gate.FieldEdit("성명", "x", 1)]))
        self.assertEqual(g_kw, dict(allow_rhwp=True, edits=[hwpx_gate.CellEdit(0, 1, 1, "y", old_text="R1C1")]))

    def test_default_output_is_not_overwritten_but_explicit_output_is(self):
        code, _, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "x")
        self.assertEqual(code, 0, err)
        default = self.dir / "_work-hwpx-automation" / "doc.hwpx"
        first = default.read_bytes()
        code, _, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "y")
        self.assertEqual(code, 1)
        self.assertIn("-o로 경로를 지정", err)
        self.assertEqual(default.read_bytes(), first)
        explicit = self.dir / "o.hwpx"
        explicit.write_bytes(b"old")
        code, _, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "y", "-o", explicit)
        self.assertEqual(code, 0, err)
        self.assertIn("<hp:t>y</hp:t>", section_text(explicit))

    def test_fields_maps_args_and_warns_on_text_security(self):
        os.environ["FAKE_RHWP_MODE"] = "suspicious"
        code, _, err = self.cli(self.doc, "--fields")
        self.assertEqual(code, 0)
        self.assertEqual(self.calls()[0], ["fields", str(self.doc), "--json"])
        self.assertIn("textSecurity=suspicious", err)

    def test_tables_marks_nested_and_keeps_document_text_on_one_line(self):
        _, out, _ = self.cli(self.doc, "--tables")
        self.assertIn("(중첩 표 있음", out)
        self.assertNotIn("\n  [0,9,9]", out)

    def test_gate_without_rhwp_allowance_refuses_before_running(self):
        def old_gate(original, produced, *, allow=None, expect_text_change=False):
            raise AssertionError("불리면 안 된다")
        with mock.patch.object(hwpx_gate, "gate", old_gate):
            code, _, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "x", "-o", self.dir / "o.hwpx")
        self.assertEqual(code, 1)
        self.assertIn("allow_rhwp", err)
        self.assertEqual(self.calls(), [])

    def test_gate_exception_is_reported_and_rejected(self):
        self.patch_gate(side_effect=ValueError("깨진 XML"))
        out_path = self.dir / "o.hwpx"
        code, _, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "x", "-o", out_path)
        self.assertEqual(code, 1)
        self.assertIn("게이트 검사 중 오류", err)
        self.assertNotIn("Traceback", err)
        self.assertFalse(out_path.exists())
        self.assertEqual(len(list(self.dir.glob("o.rejected-*.hwpx"))), 1)

    def test_allowed_rhwp_loss_is_shown_to_user(self):
        rep = hwpx_gate.GateReport()
        rep.allowed[("Contents/header.xml", "switch", ())] = (3, 0)
        self.patch_gate(return_value=rep)
        code, out, err = self.cli(self.doc, "--grid-set-cell", "0,1,1", "x", "-o", self.dir / "o.hwpx")
        self.assertEqual(code, 0, err)
        self.assertIn("허용된 손실(rhwp 재직렬화)", out)
        self.assertEqual(out.strip().splitlines()[-1], "engine=rhwp status=passed")

    def test_work_folder_input_does_not_overwrite_itself(self):
        work = self.dir / "_work-hwpx-automation"
        work.mkdir()
        src = work / "doc.hwpx"
        src.write_bytes(self.doc.read_bytes())
        code, out, err = self.cli(src, "--grid-set-cell", "0,1,1", "x")
        self.assertEqual(code, 0, err)
        self.assertTrue((work / "doc_편집.hwpx").exists())
        self.assertEqual(src.read_bytes(), self.doc.read_bytes())


if __name__ == "__main__":
    unittest.main()
