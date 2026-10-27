import io
import json
import os
import re
import sys
import tempfile
import unittest
import zipfile
from collections import namedtuple
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock

import hwpx_gate
import hwpx_sign

# 가짜 rhwp: FAKE_SIGN_PLAN(JSON 파일)의 응답을 명령별로 돌려주고 인자를 로그에 남긴다.
FAKE_RHWP = """#!{python}
import json, os, sys
args = sys.argv[1:]
if args == ["--version"]:
    print("rhwp v0.8.7"); sys.exit(0)
plan = json.load(open(os.environ["FAKE_SIGN_PLAN"], encoding="utf-8"))
with open(plan["log"], "a", encoding="utf-8") as f:
    f.write(json.dumps(args, ensure_ascii=False) + "\\n")
cmd = args[0]
if cmd == "search":
    print(json.dumps(dict(plan["search"], schemaVersion="1.0"), ensure_ascii=False))
elif cmd == "export-render-tree":
    page = int(args[args.index("-p") + 1]); folder = args[args.index("-o") + 1]
    name = plan.get("tree_name") or "render_tree_%03d.json" % (page + 1)
    json.dump(plan["tree"], open(os.path.join(folder, name), "w", encoding="utf-8"), ensure_ascii=False)
elif cmd == "measure-width":
    if "widths" not in plan:
        sys.stderr.write("오류: 알 수 없는 명령\\n"); sys.exit(1)
    texts = args[args.index("--") + 1:]
    print("font=함초롬바탕 size=10pt (13.333px) ratio=100%")
    print("text\\tchars\\twidth_px\\tper_char_px")
    for t in texts:
        w = plan["widths"].get(t, len(t) * 10.0)
        print("%s\\t%d\\t%.3f\\t%.4f" % (t, len(t), w, w / max(len(t), 1)))
elif cmd == "edit":
    out = args[args.index("-o") + 1]
    open(out, "wb").write(b"PK-fake")
    print(json.dumps({"schemaVersion": "1.0", "output": out, "outputFormat": "hwpx",
                      "page": int(args[args.index("--page") + 1]), "overflow": [], "binDataId": 1}))
"""

ANCHOR_LINE = "성 명 : 홍길동 (서명)"


def tree(runs, body=(113.4, 132.3, 566.9, 876.8)):
    """render-tree 모양: Page > Body > Column > TextLine > TextRun. runs = [(pi, text, x, y, w, h)]."""
    lines = {}
    for pi, text, x, y, w, h in runs:
        lines.setdefault((pi, y), []).append(
            {"type": "TextRun", "pi": pi, "text": text, "bbox": {"x": x, "y": y, "w": w, "h": h}})
    bx, by, bw, bh = body
    col = {"type": "Column", "col": 0, "bbox": {"x": bx, "y": by, "w": bw, "h": bh},
           "children": [{"type": "TextLine", "pi": pi, "bbox": {"x": bx, "y": y, "w": bw, "h": 13.3},
                         "children": rs} for (pi, y), rs in lines.items()]}
    return {"type": "Page", "bbox": {"x": 0, "y": 0, "w": 793.7, "h": 1122.5},
            "children": [{"type": "Body", "bbox": {"x": bx, "y": by, "w": bw, "h": bh}, "children": [col]}]}


def match(paragraph=11, offset=10, page=0, **extra):
    return dict({"section": 0, "paragraph": paragraph, "page": page, "charOffset": offset,
                 "length": 4, "text": ANCHOR_LINE, "context": ANCHOR_LINE}, **extra)


@unittest.skipIf(sys.platform == "win32", "가짜 rhwp 실행 파일이 shebang에 의존")
class RhwpSignTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        fake = self.dir / "rhwp"
        fake.write_text(FAKE_RHWP.replace("{python}", sys.executable), encoding="utf-8")
        fake.chmod(0o755)
        self.log = self.dir / "argv.log"
        self.plan_path = self.dir / "plan.json"
        self.doc = self.dir / "동의서.hwpx"
        self.doc.write_bytes(b"PK-doc")
        self.img = self.dir / "sig.png"
        self.img.write_bytes(b"\x89PNG")
        env = {"HWPX_RHWP_BIN": str(fake), "FAKE_SIGN_PLAN": str(self.plan_path)}
        patches = [mock.patch.dict(os.environ, env),
                   # 산출물 후처리·게이트는 실제 HWPX가 필요하다. 여기서는 좌표·거부 판단만 본다.
                   mock.patch.object(hwpx_sign, "float_over_anchor"),
                   mock.patch.object(hwpx_sign.hwpx_gate, "gate", return_value=hwpx_gate.GateReport())]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        self.plan(search=[match()], runs=[(11, ANCHOR_LINE, 113.4, 366.9, 130.0, 13.3)],
                  widths={"성 명 : 홍길동 ": 95.6, ANCHOR_LINE: 130.0})

    def tearDown(self):
        self.tmp.cleanup()

    def plan(self, *, search, runs, widths=None, tree_name=None):
        plan = {"log": str(self.log), "tree": tree(runs),
                "search": {"matches": search, "matchCount": len(search), "totalMatchCount": len(search)}}
        if widths is not None:
            plan["widths"] = widths
        if tree_name:
            plan["tree_name"] = tree_name
        self.plan_path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")

    def sign(self, *extra):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            rc = hwpx_sign.main([str(self.doc), "--image", str(self.img), "--height-mm", "10", *extra])
        return rc, out.getvalue(), err.getvalue()

    def calls(self, cmd):
        lines = self.log.read_text(encoding="utf-8").splitlines() if self.log.exists() else []
        return [a for a in map(json.loads, lines) if a[0] == cmd]

    def insert_args(self):
        args = self.calls("edit")[0]
        return {k: int(args[args.index(k) + 1]) for k in ("--page", "--x", "--y", "--width", "--height")}

    def test_mid_run_anchor_is_measured_and_centered(self):
        rc, out, err = self.sign()
        self.assertEqual(rc, 0, err)
        # 앵커 x = 113.4 + 130.0 × 95.6/130.0 = 209.0px, 폭 = 243.4 - 209.0 = 34.4px
        w, h = round(20 * hwpx_sign.MM), round(10 * hwpx_sign.MM)
        self.assertEqual(self.insert_args(), {
            "--page": 0, "--width": w, "--height": h,
            "--x": round((209.0 + 34.4 / 2) * 75 - w / 2),
            "--y": round((366.9 + 13.3 / 2) * 75 - h / 2)})
        self.assertIn("measure-width", err)
        self.assertEqual(out.strip().splitlines()[-1], "engine=rhwp")
        self.assertTrue((self.dir / "동의서_서명.hwpx").exists())

    def test_whole_run_anchor_needs_no_measurement(self):
        self.plan(search=[match(offset=10)],
                  runs=[(11, "성 명 : 홍길동 ", 113.4, 366.9, 95.6, 13.3), (11, "(서명)", 209.0, 366.9, 34.4, 13.3)])
        rc, _, err = self.sign()
        self.assertEqual(rc, 0, err)
        self.assertFalse(self.calls("measure-width"))
        self.assertEqual(self.insert_args()["--x"], round((209.0 + 17.2) * 75 - round(20 * hwpx_sign.MM) / 2))

    def test_unmeasurable_anchor_is_estimated_and_blocked(self):
        self.plan(search=[match()], runs=[(11, ANCHOR_LINE, 113.4, 366.9, 130.0, 13.3)])  # measure-width 실패
        rc, _, err = self.sign()
        self.assertEqual(rc, 1)
        self.assertIn("--allow-estimated", err)
        self.assertFalse(self.calls("edit"))
        rc, _, err = self.sign("--allow-estimated")
        self.assertEqual(rc, 0, err)
        # 글자 수 비례: 앵커는 14자 중 10번째부터 4자
        x0, x1 = 113.4 + 130.0 * 10 / 14, 113.4 + 130.0
        self.assertEqual(self.insert_args()["--x"], round((x0 + x1) / 2 * 75 - round(20 * hwpx_sign.MM) / 2))

    def test_zero_and_multiple_matches_are_refused(self):
        self.plan(search=[], runs=[])
        rc, _, err = self.sign()
        self.assertEqual(rc, 1)
        self.assertIn("찾지 못했습니다", err)
        self.plan(search=[match(paragraph=5, offset=0), match()],
                  runs=[(11, ANCHOR_LINE, 113.4, 366.9, 130.0, 13.3)],
                  widths={"성 명 : 홍길동 ": 95.6, ANCHOR_LINE: 130.0})
        rc, _, err = self.sign()
        self.assertEqual(rc, 1)
        self.assertIn("--occurrence", err)
        self.assertFalse(self.calls("edit"))
        rc, _, err = self.sign("--occurrence", "3")
        self.assertEqual(rc, 1)
        self.assertIn("2곳뿐", err)
        rc, _, err = self.sign("--occurrence", "2")
        self.assertEqual(rc, 0, err)   # 1번째(문단 5)는 render-tree에 run이 없어 2번째가 쓰였다
        self.assertEqual(self.calls("export-render-tree")[-1][-3:-2], ["0"])

    def test_table_anchor_points_to_com(self):
        self.plan(search=[match(cell={"cell": 0, "control": 0, "paragraph": 0})], runs=[])
        rc, _, err = self.sign()
        self.assertEqual(rc, 1)
        self.assertIn("--backend com", err)

    def test_render_tree_page_mismatch_stops(self):
        self.plan(search=[match()], runs=[(11, ANCHOR_LINE, 113.4, 366.9, 130.0, 13.3)],
                  tree_name="render_tree_002.json")
        rc, _, err = self.sign()
        self.assertEqual(rc, 1)
        self.assertIn("쪽 번호", err)

    def test_mm_options_convert_to_hwpunit(self):
        rc, _, err = self.sign("--width-mm", "30", "--height-mm", "12", "--gap-mm", "7", "--vert-adjust", "-500")
        self.assertEqual(rc, 0, err)
        a = self.insert_args()
        self.assertEqual((a["--width"], a["--height"]), (8504, 3402))   # 30mm, 12mm × 283.46
        body_right = (113.4 + 566.9) * 75
        self.assertEqual(a["--x"], round(body_right) - 8504 - 1984)      # 우측 정렬, 여백 7mm
        self.assertEqual(a["--y"], round((366.9 + 13.3 / 2) * 75 - 3402 / 2) - 500)

    def test_hwp_input_goes_through_converter(self):
        hwp = self.dir / "동의서.hwp"
        hwp.write_bytes(b"HWP")
        Conv = namedtuple("ConvertResult", "path engine status messages")
        with mock.patch.object(hwpx_sign, "_convert_hwp",
                               return_value=Conv(str(self.doc), "rhwp", "passed", ["변환 완료"])) as conv:
            out, err = io.StringIO(), io.StringIO()
            with redirect_stdout(out), redirect_stderr(err):
                rc = hwpx_sign.main([str(hwp), "--image", str(self.img), "--height-mm", "10"])
        self.assertEqual(rc, 0, err.getvalue())
        conv.assert_called_once_with(str(hwp))
        self.assertEqual(self.calls("search")[0][1], str(self.doc))
        self.assertTrue((self.dir / "동의서_서명.hwpx").exists())

    def test_gate_failure_keeps_rejected_output(self):
        failed = hwpx_gate.GateReport()
        failed.fail("사라진 요소 lineseg×1")
        with mock.patch.object(hwpx_sign.hwpx_gate, "gate", return_value=failed):
            rc, out, err = self.sign()
        self.assertEqual(rc, 1)
        self.assertIn("--backend com", err)
        self.assertTrue((self.dir / "_work-hwpx-automation" / "동의서_서명.hwpx").exists())
        self.assertFalse((self.dir / "동의서_서명.hwpx").exists())
        self.assertNotIn("engine=", out)

    def test_hwp_input_without_converter_exits_1(self):
        hwp = self.dir / "동의서.hwp"
        hwp.write_bytes(b"HWP")
        with mock.patch.dict(sys.modules, {"hwpx_edit": None}):
            out, err = io.StringIO(), io.StringIO()
            with redirect_stdout(out), redirect_stderr(err):
                rc = hwpx_sign.main([str(hwp), "--image", str(self.img), "--height-mm", "10"])
        self.assertEqual(rc, 1)
        self.assertIn("--from-hwp", err.getvalue())

    def test_pdf_without_pymupdf_counts_svg_images(self):
        def pdf(svg_counts):
            out, err = io.StringIO(), io.StringIO()
            with mock.patch.object(hwpx_sign.subprocess, "run", return_value=mock.Mock(returncode=0)) as run, \
                    mock.patch.object(hwpx_sign, "count_pdf_images", return_value=None), \
                    mock.patch.object(hwpx_sign, "svg_image_count", side_effect=svg_counts), \
                    redirect_stdout(out), redirect_stderr(err):
                rc = hwpx_sign.make_pdf(str(self.dir / "동의서_서명.hwpx"), svg_check=(str(self.doc), 0))
            self.assertEqual(run.call_args[0][0][-2:], ["--pdf-backend", "rhwp"])
            return rc, out.getvalue(), err.getvalue()
        rc, out, _ = pdf([0, 1])
        self.assertEqual(rc, 0)
        self.assertIn("그림 0→1개", out)
        rc, _, err = pdf([1, 1])
        self.assertEqual(rc, 1)
        self.assertIn("늘지 않았습니다", err)

    def test_convert_without_edit_m2_explains(self):
        with mock.patch.dict(sys.modules, {"hwpx_edit": None}):
            with self.assertRaises(hwpx_sign.SignError) as cm:
                hwpx_sign._convert_hwp(str(self.dir / "a.hwp"))
        self.assertIn("--from-hwp", str(cm.exception))


def make_consent_hwpx(path):
    """익명 동의서 표본. python-hwpx Skeleton을 한컴 저장본처럼 다듬는다.

    Skeleton 고유의 paraPr textDir 속성과 빈 <hp:t/>는 rhwp가 지우므로 미리 뺀다.
    남는 rhwp 왕복 차이는 한컴 저장 공문과 같다(OUTLINE 8~10 switch, 문단 id 재부여).
    """
    from hwpx import HwpxDocument
    doc = HwpxDocument.new()
    doc.add_paragraph("개인정보 수집·이용 동의서")
    for i in range(8):
        doc.add_paragraph(f"본 동의서의 {i + 1}번째 조항 내용입니다. 수집 항목과 이용 목적을 확인합니다.")
    doc.add_paragraph("2026년 10월 28일")
    doc.add_paragraph(ANCHOR_LINE)
    doc.save_to_path(path)
    fix = {"Contents/header.xml": lambda d: d.replace(b' textDir="LTR"', b""),
           "Contents/section0.xml": lambda d: d.replace(b"<hp:t/>", b"")}
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = fix.get(item.filename, lambda d: d)(zin.read(item.filename))
            zout.writestr(item, data)
    os.replace(tmp, path)


@unittest.skipUnless(os.environ.get("RHWP_INTEGRATION"), "RHWP_INTEGRATION=1일 때 실제 rhwp로 실행")
class RhwpSignIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.doc = self.dir / "consent-sign.hwpx"
        make_consent_hwpx(self.doc)
        from PIL import Image
        Image.new("RGB", (300, 150), "white").save(self.dir / "sig.png")

    def tearDown(self):
        self.tmp.cleanup()

    def test_signature_lands_on_anchor(self):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            rc = hwpx_sign.main([str(self.doc), "--image", str(self.dir / "sig.png"), "--width-mm", "20"])
        self.assertEqual(rc, 0, err.getvalue())
        result = self.dir / "consent-sign_서명.hwpx"
        self.assertIn("허용된 손실(rhwp 재직렬화)", out.getvalue())
        rep = hwpx_gate.gate(self.doc, result, allow_rhwp=True, allow_bindata_added=1)
        self.assertEqual(rep.status, "passed", rep.summary())
        self.assertEqual({k[1] for k in rep.allowed}, {"switch", "case", "default", "heading", "p"})

        _, box = hwpx_sign.locate(str(self.doc), "(서명)", None)
        with tempfile.TemporaryDirectory() as td:
            hwpx_sign.rhwp_cli.run(["export-svg", str(result), "-p", "0", "-o", td])
            svg = next(Path(td).glob("*.svg")).read_text(encoding="utf-8")
        images = re.findall(r'<image x="([\d.]+)" y="([\d.]+)" width="([\d.]+)" height="([\d.]+)"', svg)
        self.assertEqual(len(images), 1)
        x, y, w, h = map(float, images[0])
        self.assertAlmostEqual(x + w / 2, box.x + box.w / 2, delta=0.5)
        self.assertAlmostEqual(y + h / 2, box.y + box.h / 2, delta=0.5)


    def test_hwp_input_converts_and_keeps_empty_runs(self):
        # rhwp가 변환한 HWPX는 첫 문단에 빈 <hp:t>가 있고, insert-image가 그 자리를 그림으로 바꿔 쓴다.
        hwp = self.dir / "sample-form.hwp"
        hwp.write_bytes((Path(__file__).parent / "tests" / "fixtures" / "sample-form.hwp").read_bytes())
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            rc = hwpx_sign.main([str(hwp), "--image", str(self.dir / "sig.png"), "--anchor", "작성일"])
        self.assertEqual(rc, 0, err.getvalue())
        self.assertTrue((self.dir / "sample-form_서명.hwpx").exists())
        self.assertEqual(out.getvalue().strip().splitlines()[-1], "engine=rhwp")

if __name__ == "__main__":
    unittest.main()
