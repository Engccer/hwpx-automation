import os
import re
import tempfile
import unittest
import zipfile
from pathlib import Path

import hwpx_gate


def make_hwpx(path, text="첫 문단 내용입니다 2026", pic=False):
    """python-hwpx Skeleton으로 작은 HWPX를 만든다. pic=True면 BinData 그림 하나를 넣는다."""
    from hwpx import HwpxDocument
    doc = HwpxDocument.new()
    doc.add_paragraph(text)
    doc.save_to_path(path)
    if pic:
        inject_bindata(path, "BinData/image1.png", b"\x89PNG fake")


def inject_bindata(path, name, data):
    rewrite(path, add={name: data})


def rewrite(path, *, add=None, drop=None, transform=None):
    """zip을 다시 써서 항목을 추가·삭제·변형한다(원본 보존 규칙은 시험이라 무시)."""
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            if drop and item.filename in drop:
                continue
            data = zin.read(item.filename)
            if transform and item.filename in transform:
                data = transform[item.filename](data)
            zout.writestr(item, data)
        for name, data in (add or {}).items():
            zout.writestr(name, data)
    os.replace(tmp, path)


class GateTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.orig = self.dir / "orig.hwpx"
        make_hwpx(self.orig, pic=True)
        self.prod = self.dir / "prod.hwpx"
        self.prod.write_bytes(self.orig.read_bytes())

    def tearDown(self):
        self.tmp.cleanup()

    def section_name(self):
        return next(n for n in hwpx_gate.read_parts(self.orig) if re.fullmatch(r"Contents/section\d+\.xml", n))

    def test_identical_copy_passes(self):
        rep = hwpx_gate.gate(self.orig, self.prod)
        self.assertEqual(rep.status, "passed", rep.summary())
        self.assertFalse(rep.lost)

    def test_removed_element_is_lost(self):
        def drop_lineseg(data):
            return re.sub(rb"<hp:linesegarray>.*?</hp:linesegarray>", b"", data, count=1, flags=re.S)
        rewrite(self.prod, transform={self.section_name(): drop_lineseg})
        rep = hwpx_gate.gate(self.orig, self.prod)
        self.assertEqual(rep.status, "failed")
        self.assertTrue(any(k[1] in ("lineseg", "linesegarray") for k in rep.lost), rep.lost)

    def test_allowed_change_is_not_lost(self):
        def drop_lineseg(data):
            return re.sub(rb"<hp:linesegarray>.*?</hp:linesegarray>", b"", data, count=1, flags=re.S)
        rewrite(self.prod, transform={self.section_name(): drop_lineseg})
        rep = hwpx_gate.gate(self.orig, self.prod, allow=hwpx_gate.is_text_run_key)
        self.assertNotIn("사라진 요소", " ".join(rep.messages))

    def test_attribute_change_is_detected(self):
        header = "Contents/header.xml"
        def bump(data):
            return data.replace(b'textColor="#000000"', b'textColor="#FF0000"', 1)
        before = hwpx_gate.read_parts(self.orig)[header]
        if b'textColor="#000000"' not in before:
            self.skipTest("Skeleton에 textColor 속성이 없음")
        rewrite(self.prod, transform={header: bump})
        rep = hwpx_gate.gate(self.orig, self.prod)
        self.assertEqual(rep.status, "failed")
        self.assertTrue(any("textColor" in str(k) for k in rep.lost))

    def test_bindata_change_and_removal_fail(self):
        rewrite(self.prod, transform={"BinData/image1.png": lambda d: d + b"x"})
        rep = hwpx_gate.gate(self.orig, self.prod)
        self.assertEqual(rep.bindata_changed, ["BinData/image1.png"])
        self.assertEqual(rep.status, "failed")
        self.prod.write_bytes(self.orig.read_bytes())
        rewrite(self.prod, drop={"BinData/image1.png"})
        rep = hwpx_gate.gate(self.orig, self.prod, require_census=False)
        self.assertEqual(rep.bindata_removed, ["BinData/image1.png"])
        self.assertEqual(rep.status, "failed")

    def test_bindata_addition_respects_allowance(self):
        inject_bindata(self.prod, "BinData/image2.png", b"\x89PNG two")
        self.assertEqual(hwpx_gate.gate(self.orig, self.prod).status, "failed")
        self.assertEqual(hwpx_gate.gate(self.orig, self.prod, allow_bindata_added=1).status, "passed")

    def test_text_loss_is_detected(self):
        def cut(data):
            return data.replace("2026".encode("utf-8"), b"", 1)
        rewrite(self.prod, transform={self.section_name(): cut})
        rep = hwpx_gate.gate(self.orig, self.prod, allow=lambda k: True)
        self.assertEqual(rep.status, "failed")
        self.assertEqual(sum(rep.text_missing.values()), 4)

    def test_text_change_can_be_expected(self):
        def cut(data):
            return data.replace("2026".encode("utf-8"), b"2027", 1)
        rewrite(self.prod, transform={self.section_name(): cut})
        rep = hwpx_gate.gate(self.orig, self.prod, allow=lambda k: True, expect_text_change=True)
        self.assertEqual(rep.status, "passed", rep.summary())

    def test_missing_validator_means_unverified(self):
        original = hwpx_gate.shutil.which
        hwpx_gate.shutil.which = lambda name: None if name == "hwpx-validate" else original(name)
        try:
            rep = hwpx_gate.gate(self.orig, self.prod)
        finally:
            hwpx_gate.shutil.which = original
        self.assertEqual(rep.status, "unverified")
        self.assertIsNone(rep.validate_ok)

    def collapse_outline_switch(self, data, used):
        """rhwp가 하는 대로 paraPr 하나의 OUTLINE switch를 default 갈래로 접는다.
        used=True면 본문이 쓰는 paraPr(id 0), False면 안 쓰는 OUTLINE 8~10단계 paraPr."""
        text = data.decode("utf-8")
        for m in re.finditer(r'<hh:paraPr id="(\d+)".*?</hh:paraPr>', text, re.S):
            sw = re.search(r'<hp:switch><hp:case [^>]*><hh:heading [^>]*/></hp:case>'
                           r'<hp:default>(<hh:heading [^>]*/>)</hp:default></hp:switch>', m.group(0))
            if (m.group(1) == "0") == used and sw:
                block = m.group(0).replace(sw.group(0), sw.group(1))
                return (text[:m.start()] + block + text[m.end():]).encode("utf-8")
        self.skipTest("Skeleton header에 OUTLINE switch를 가진 paraPr가 없음")

    def test_rhwp_known_losses_are_allowed(self):
        header = "Contents/header.xml"
        renumber = iter(range(1, 1000))
        rewrite(self.prod, transform={
            header: lambda d: self.collapse_outline_switch(d, used=False),
            self.section_name(): lambda d: re.sub(rb'<hp:p id="\d+"', lambda m: b'<hp:p id="%d"' % next(renumber), d)})
        self.assertEqual(hwpx_gate.gate(self.orig, self.prod).status, "failed")
        rep = hwpx_gate.gate(self.orig, self.prod, allow_rhwp=True)
        self.assertEqual(rep.status, "passed", rep.summary())
        self.assertFalse(rep.lost)
        self.assertEqual({k[1] for k in rep.allowed}, {"switch", "case", "default", "heading", "p"})
        self.assertIn("허용된 손실(rhwp 재직렬화): ", rep.summary())
        self.assertIn("p id 재부여×", hwpx_gate.describe_lost(rep))

    def test_rhwp_losses_outside_allowlist_still_fail(self):
        def drop_lineseg(data):
            return re.sub(rb"<hp:linesegarray>.*?</hp:linesegarray>", b"", data, count=1, flags=re.S)
        rewrite(self.prod, transform={
            "Contents/header.xml": lambda d: self.collapse_outline_switch(d, used=True),
            self.section_name(): drop_lineseg})
        rep = hwpx_gate.gate(self.orig, self.prod, allow_rhwp=True)
        self.assertEqual(rep.status, "failed")
        lost_tags = {(k[0], k[1]) for k in rep.lost}
        self.assertIn(("Contents/header.xml", "switch"), lost_tags)
        self.assertTrue(any(t in ("lineseg", "linesegarray") for _, t in lost_tags), lost_tags)
        self.assertFalse(rep.allowed)

    def test_cross_check_reports_table_and_text_difference(self):
        other = self.dir / "other.hwpx"
        make_hwpx(other, text="첫 문단 내용입니다 2026", pic=True)
        self.assertEqual(hwpx_gate.cross_check(self.orig, other).status, "passed")
        make_hwpx(other, text="다른 글자", pic=False)
        rep = hwpx_gate.cross_check(self.orig, other)
        self.assertEqual(rep.status, "failed")
        self.assertTrue(any("BinData" in m for m in rep.messages))
        self.assertTrue(any("글자" in m for m in rep.messages))


    def test_cross_check_compares_bindata_content(self):
        other = self.dir / "other.hwpx"
        other.write_bytes(self.orig.read_bytes())
        rewrite(other, transform={"BinData/image1.png": lambda d: d + b"x"})
        rep = hwpx_gate.cross_check(self.orig, other)
        self.assertEqual(rep.status, "failed")
        self.assertTrue(any("해시 불일치 1건" in m for m in rep.messages), rep.messages)


class CensusTests(unittest.TestCase):
    def test_census_distinguishes_attributes(self):
        a = hwpx_gate.element_census(b'<r><hp:t xmlns:hp="u" a="1"/><hp:t xmlns:hp="u" a="2"/></r>')
        b = hwpx_gate.element_census(b'<r><hp:t xmlns:hp="u" a="1"/><hp:t xmlns:hp="u" a="1"/></r>')
        diff = hwpx_gate.compare_census(a, b)
        self.assertEqual(len(diff), 2)
        self.assertEqual(hwpx_gate.tag_census(b'<r><x/><x/><y/></r>'), {"r": 1, "x": 2, "y": 1})


@unittest.skipUnless(os.environ.get("RHWP_INTEGRATION"), "RHWP_INTEGRATION=1일 때 실제 rhwp로 실행")
class RhwpIntegrationTests(unittest.TestCase):
    """실제 바이너리 계약 확인. 표본은 python-hwpx Skeleton으로 즉석 생성한다(개인 문서 금지)."""

    def setUp(self):
        import rhwp_cli
        self.rhwp_cli = rhwp_cli
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.src = self.dir / "src.hwpx"
        make_hwpx(self.src)

    def tearDown(self):
        self.tmp.cleanup()

    def test_export_hwpx_envelope_and_gate_shape(self):
        out = self.dir / "rt.hwpx"
        r = self.rhwp_cli.run(["export-hwpx", str(self.src), self.rhwp_cli.OUTPUT, "--verify"],
                              output=out, required_fields=("output", "format", "verify"))
        self.assertEqual(r.envelope["format"], "hwpx")
        self.assertTrue(r.envelope["verify"]["identical"])
        rep = hwpx_gate.gate(self.src, out)
        # rhwp 자가 검증은 통과하지만 독립 게이트는 header.xml 손실(switch·paraPr)을 본다.
        self.assertIn(rep.status, ("passed", "failed"))
        self.assertEqual(rep.validate_ok, True)
        self.assertFalse(rep.text_missing)

    def test_export_pdf_through_adapter(self):
        from rhwp_pdf import save_pdf
        report = save_pdf(self.src, self.dir / "out.pdf")
        self.assertEqual(report["pages"], 1)
        self.assertTrue((self.dir / "out.pdf").read_bytes().startswith(b"%PDF-"))


if __name__ == "__main__":
    unittest.main()
