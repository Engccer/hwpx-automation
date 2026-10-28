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
        rep = hwpx_gate.gate(self.orig, self.prod, allow=lambda k: k[0] in ("lineseg", "linesegarray"))
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
        rewrite(self.prod, transform={self.section_name(): drop_lineseg})
        rep = hwpx_gate.gate(self.orig, self.prod, allow_rhwp=True)
        self.assertEqual(rep.status, "failed")
        self.assertTrue(any(k[1] in ("lineseg", "linesegarray") for k in rep.lost), rep.lost)
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

    def test_cross_check_counts_embedded_objects_and_chart_parts(self):
        def add(tag):
            return lambda d: d.replace(b"</hp:run>", f"<hp:{tag}/></hp:run>".encode(), 1)
        for tag in ("chart", "ole", "equation"):
            with self.subTest(tag=tag):
                other = self.dir / f"{tag}.hwpx"
                other.write_bytes(self.orig.read_bytes())
                rewrite(other, transform={self.section_name(): add(tag)})
                rep = hwpx_gate.cross_check(other, self.orig)
                self.assertEqual(rep.status, "failed")
                self.assertIn(f"{tag} 1→0", " ".join(rep.messages))
        a, b = self.dir / "a.hwpx", self.dir / "b.hwpx"
        for path, data in ((a, b"<c:chart>1</c:chart>"), (b, b"<c:chart>2</c:chart>")):
            path.write_bytes(self.orig.read_bytes())
            rewrite(path, add={"Chart/chart1.xml": data})
        rep = hwpx_gate.cross_check(a, b)
        self.assertTrue(any("Chart/ 파트 해시 불일치" in m for m in rep.messages), rep.messages)


NS = 'xmlns:hh="h" xmlns:hp="p"'


def header(*para_prs):
    return f'<hh:head {NS}><hh:paraProperties>{"".join(para_prs)}</hh:paraProperties></hh:head>'.encode()


def outline_para_pr(pid, collapsed=False):
    body = ('<hh:heading type="NONE" level="0"/>' if collapsed else
            '<hp:switch><hp:case hp:required-namespace="n"><hh:heading type="OUTLINE" level="8"/></hp:case>'
            '<hp:default><hh:heading type="NONE" level="0"/></hp:default></hp:switch>')
    return f'<hh:paraPr id="{pid}">{body}</hh:paraPr>'


class SplitRhwpKnownTests(unittest.TestCase):
    """allow_rhwp 허용 목록의 경계. 실측한 rhwp 왕복 모양을 작은 XML로 재현한다."""

    def test_header_collapse_allowed_only_in_unused_parapr(self):
        before = header(outline_para_pr(0), outline_para_pr(16))
        unused = header(outline_para_pr(0), outline_para_pr(16, collapsed=True))
        diff, known = hwpx_gate.split_rhwp_known("Contents/header.xml", before, unused, {"0"})
        self.assertFalse([k for k, (a, b) in diff.items() if b < a])
        self.assertEqual({k[0] for k in known}, {"switch", "case", "default", "heading"})
        used = header(outline_para_pr(0, collapsed=True), outline_para_pr(16))
        diff, known = hwpx_gate.split_rhwp_known("Contents/header.xml", before, used, {"0"})
        self.assertIn("switch", {k[0] for k, (a, b) in diff.items() if b < a})
        self.assertFalse(known)

    def test_section_renumber_and_shadow_none(self):
        sec = lambda ids, shadow: (f'<hs:sec xmlns:hs="s" {NS}>' + "".join(
            f'<hp:p id="{i}" paraPrIDRef="{r}"/>' for i, r in ids) + shadow + "</hs:sec>").encode()
        none = '<hp:shadow type="NONE" alpha="0"/>'
        before = sec([("77", 0), ("91", 1)], none)
        diff, known = hwpx_gate.split_rhwp_known("Contents/section0.xml", before,
                                                 sec([("0", 0), ("1", 1)], ""), set())
        self.assertFalse(diff)
        self.assertEqual(sorted(k[0] for k in known), ["p", "p", "shadow"])
        # 재부여와 함께 다른 속성이 바뀌면 허용하지 않는다
        diff, known = hwpx_gate.split_rhwp_known("Contents/section0.xml", before,
                                                 sec([("0", 0), ("1", 2)], none), set())
        self.assertIn("p", {k[0] for k, (a, b) in diff.items() if b < a})
        self.assertFalse(known)
        # 그림자가 생기는 변경(NONE→DROP)은 허용하지 않는다
        diff, known = hwpx_gate.split_rhwp_known("Contents/section0.xml", before,
                                                 sec([("77", 0), ("91", 1)], '<hp:shadow type="DROP" alpha="0"/>'), set())
        self.assertIn("shadow", {k[0] for k, (a, b) in diff.items() if b < a})
        self.assertFalse(known)


SEC = 'xmlns:hs="s" xmlns:hp="p"'


def run(text, cp="0", extra=""):
    return f'<hp:run charPrIDRef="{cp}">{extra}<hp:t>{text}</hp:t></hp:run>'


def para(pid, *runs, ppr="0"):
    return f'<hp:p id="{pid}" paraPrIDRef="{ppr}">{"".join(runs)}<hp:linesegarray><hp:lineseg vertpos="0"/></hp:linesegarray></hp:p>'


def cell(r, c, text, pid):
    return (f'<hp:tc><hp:subList id="" vertAlign="CENTER">{para(pid, run(text))}</hp:subList>'
            f'<hp:cellAddr colAddr="{c}" rowAddr="{r}"/><hp:cellSz width="100" height="10"/></hp:tc>')


def form(cells=("성명", "", "학년", ""), head="머리 문단", tail="꼬리 문단", ids=(1, 2, 3, 4, 5, 6, 7),
         sec_pr='<hp:secPr textDirection="HORIZONTAL"/>', tail_runs=None):
    """머리 문단(secPr) · 2×2 표 · 꼬리 문단."""
    tbl = ('<hp:tbl rowCnt="2" colCnt="2">' + "".join(
        f'<hp:tr>{cell(r, 0, cells[2 * r], ids[2 + 2 * r])}{cell(r, 1, cells[2 * r + 1], ids[3 + 2 * r])}</hp:tr>'
        for r in range(2)) + "</hp:tbl>")
    body = (para(ids[0], run(head, extra=sec_pr)) + para(ids[1], f'<hp:run charPrIDRef="0">{tbl}<hp:t/></hp:run>')
            + para(ids[6], *(tail_runs or [run(tail)])))
    return {"Contents/section0.xml": f"<hs:sec {SEC}>{body}</hs:sec>".encode()}


def field_para(pid, before="앞 ", value="옛값", after=" 뒤", name="성명", fid="7", other=""):
    begin = (f'<hp:ctrl><hp:fieldBegin id="{fid}" type="CLICK_HERE" name="{name}" editable="1" dirty="0" '
             f'zorder="-1" fieldid="{fid}0"/></hp:ctrl>')
    end = f'<hp:ctrl><hp:fieldEnd beginIDRef="{fid}" fieldid="{fid}0"/></hp:ctrl>'
    return para(pid, f'<hp:run charPrIDRef="0"><hp:t>{before}</hp:t>{begin}<hp:t>{value}</hp:t>{end}'
                     f'<hp:t>{after}</hp:t></hp:run>' + other)


def fields_doc(**kw):
    rest = para(9, run("다른 문단"))
    return {"Contents/section0.xml": f"<hs:sec {SEC}>{field_para(1, **kw)}{rest}</hs:sec>".encode()}


class EditCheck:
    EDIT = [hwpx_gate.CellEdit(0, 0, 1, "홍길동")]

    def after(self, **kw):
        return form(cells=("성명", "홍길동", "학년", ""), **kw)

    def check(self, before, after, edits):
        rep = hwpx_gate.GateReport()
        hwpx_gate.compare_edit(before, after, edits, rep)
        return rep

    def assert_fails(self, before, after, edits, needle):
        rep = self.check(before, after, edits)
        self.assertEqual(rep.status, "failed", rep.messages)
        self.assertIn(needle, " / ".join(rep.messages))


class EditCompareTests(EditCheck, unittest.TestCase):
    """rhwp 편집 산출물의 위치별 대조(edits). 대상 밖은 p id 말고 아무것도 바뀌면 안 된다."""

    def test_cell_edit_with_renumbered_ids_passes(self):
        rep = self.check(form(), form(cells=("성명", "홍길동", "학년", ""), ids=range(7)),
                         [hwpx_gate.CellEdit(0, 0, 1, "홍길동", old_text="")])
        self.assertEqual(rep.status, "passed", rep.messages)
        self.assertEqual({k[1] for k in rep.allowed}, {"p"})

    def test_cell_edit_rejects_blanked_text_outside_target(self):
        self.assert_fails(form(), form(cells=("", "홍길동", "학년", "")),
                          [hwpx_gate.CellEdit(0, 0, 1, "홍길동")], "대상 밖")
        self.assert_fails(form(), form(cells=("성명", "홍길동", "학년", ""), tail=""),
                          [hwpx_gate.CellEdit(0, 0, 1, "홍길동")], "대상 밖")

    def test_cell_edit_rejects_paragraph_loss_and_secpr_change(self):
        after = form(cells=("성명", "홍길동", "학년", ""))
        dropped = {k: v.replace(para(7, run("꼬리 문단")).encode(), b"") for k, v in after.items()}
        self.assert_fails(form(), dropped, [hwpx_gate.CellEdit(0, 0, 1, "홍길동")], "대상 밖")
        self.assert_fails(form(), form(cells=("성명", "홍길동", "학년", ""), sec_pr='<hp:secPr/>'),
                          [hwpx_gate.CellEdit(0, 0, 1, "홍길동")], "대상 밖")

    def test_cell_edit_rejects_swapped_char_shapes_outside_target(self):
        before = form(tail_runs=[run("가", "1"), run("나", "2")])
        after = form(cells=("성명", "홍길동", "학년", ""), tail_runs=[run("가", "2"), run("나", "1")])
        self.assert_fails(before, after, [hwpx_gate.CellEdit(0, 0, 1, "홍길동")], "대상 밖")

    def test_cell_edit_landing_in_another_empty_cell_is_rejected(self):
        # 요청은 (0,1)인데 rhwp가 같은 빈 칸 (1,1)에 썼다: 대상 칸 글자 불일치 + 대상 밖 변화
        rep = self.check(form(), form(cells=("성명", "", "학년", "홍길동")),
                         [hwpx_gate.CellEdit(0, 0, 1, "홍길동", old_text="")])
        self.assertEqual(rep.status, "failed")

    def test_cell_edit_checks_old_text_and_inner_objects(self):
        self.assert_fails(form(), form(cells=("성명", "홍길동", "학년", "")),
                          [hwpx_gate.CellEdit(0, 0, 1, "홍길동", old_text="다른 글자")], "옛 글자")
        pic = '<hp:pic id="1"><hp:img binaryItemIDRef="image1"/></hp:pic>'
        sec = form()["Contents/section0.xml"]
        before = {"Contents/section0.xml": sec.replace(b'<hp:p id="4" paraPrIDRef="0"><hp:run charPrIDRef="0">',
                                                       f'<hp:p id="4" paraPrIDRef="0"><hp:run charPrIDRef="0">{pic}'.encode())}
        self.assertIn(b"hp:pic", before["Contents/section0.xml"])
        self.assert_fails(before, form(cells=("성명", "홍길동", "학년", "")),
                          [hwpx_gate.CellEdit(0, 0, 1, "홍길동")], "글자 밖 요소")

    def test_missing_target_is_refused(self):
        self.assert_fails(form(), form(), [hwpx_gate.CellEdit(3, 0, 0, "x")], "대상 확정 실패")
        self.assert_fails(form(), form(), [hwpx_gate.CellEdit(0, 5, 0, "x")], "대상 확정 실패")

    def test_field_fill_passes_and_ignores_dirty_flag(self):
        after = {k: v.replace(b'dirty="0" zorder="-1" ', b'dirty="1" ') for k, v in fields_doc(value="홍길동").items()}
        rep = self.check(fields_doc(), after, [hwpx_gate.FieldEdit("성명", "홍길동")])
        self.assertEqual(rep.status, "passed", rep.messages)

    def test_field_fill_rejects_text_moved_into_field(self):
        # 앞뒤 글자를 비우고 값 자리에 몰아 넣으면 문단 글자열은 같아도 거부한다
        self.assert_fails(fields_doc(before="앞", value="옛", after="뒤"),
                          fields_doc(before="", value="앞새뒤", after=""),
                          [hwpx_gate.FieldEdit("성명", "앞새뒤")], "대상 밖")

    def test_field_fill_rejects_wrong_value_and_other_paragraph_change(self):
        self.assert_fails(fields_doc(), fields_doc(value="김철수"), [hwpx_gate.FieldEdit("성명", "홍길동")], "요청값")
        after = {k: v.replace("다른 문단".encode(), b"") for k, v in fields_doc(value="홍길동").items()}
        self.assert_fails(fields_doc(), after, [hwpx_gate.FieldEdit("성명", "홍길동")], "대상 밖")
        self.assert_fails(fields_doc(), fields_doc(), [hwpx_gate.FieldEdit("학교", "x")], "대상 확정 실패")

    def test_image_insert_allows_exactly_one_picture_in_anchor_paragraph(self):
        pic = '<hp:pic id="9"><hp:pos horzOffset="1"/><hp:img binaryItemIDRef="image9"/></hp:pic>'
        before = form()
        sec = before["Contents/section0.xml"]
        anchored = {"Contents/section0.xml": sec.replace(b'<hp:run charPrIDRef="0"><hp:t>\xea\xbc\xac',
                                                          f'<hp:run charPrIDRef="0">{pic}<hp:t>꼬'.encode())}
        self.assertNotEqual(anchored["Contents/section0.xml"], sec)
        ok = self.check(before, anchored, [hwpx_gate.ImageInsert(0, 2)])
        self.assertEqual(ok.status, "passed", ok.messages)
        self.assert_fails(before, anchored, [hwpx_gate.ImageInsert(0, 0)], "그림 하나만")
        blanked = {k: v.replace("머리 문단".encode(), b"") for k, v in anchored.items()}
        self.assert_fails(before, blanked, [hwpx_gate.ImageInsert(0, 2)], "그림 하나만")


class CanonicalTokenTests(EditCheck, unittest.TestCase):
    """rhwp가 한컴 문서를 다시 쓸 때 생기는 포장 차이(실측)는 통과하고 글자·모양 차이는 거부한다."""

    def test_run_and_t_boundaries_empty_t_and_crlf_are_ignored(self):
        before = form(tail_runs=[run("꼬리 문단"), '<hp:run charPrIDRef="0"/>'])
        split = [run("꼬리"), run(" 문단", extra="<hp:t/>")]
        rep = self.check(before, self.after(tail_runs=split), self.EDIT)
        self.assertEqual(rep.status, "passed", rep.messages)
        crlf = form(tail_runs=[run("가&#13;\n나")])
        rep = self.check(crlf, form(cells=("성명", "홍길동", "학년", ""), tail_runs=[run("가\n나")]), self.EDIT)
        self.assertEqual(rep.status, "passed", rep.messages)

    def test_char_shape_change_of_some_letters_is_rejected(self):
        before = form(tail_runs=[run("꼬리 문단")])
        self.assert_fails(before, self.after(tail_runs=[run("꼬리"), run(" 문단", "3")]), self.EDIT, "대상 밖")

    def test_picture_child_order_is_ignored(self):
        pic = lambda a, b: f'<hp:pic id="1">{a}{b}</hp:pic>'
        img, rect = '<hp:img binaryItemIDRef="image1"/>', '<hp:imgRect/>'
        before = form(tail_runs=[run("꼬리", extra=pic(img, rect))])
        rep = self.check(before, self.after(tail_runs=[run("꼬리", extra=pic(rect, img))]), self.EDIT)
        self.assertEqual(rep.status, "passed", rep.messages)
        self.assert_fails(before, self.after(tail_runs=[run("꼬리", extra=pic(rect, ""))]), self.EDIT, "대상 밖")

    def test_nested_table_text_inside_target_cell_is_kept(self):
        nested = '<hp:tbl rowCnt="1" colCnt="1"><hp:tr><hp:tc><hp:subList>' + para(50, run("안쪽")) + \
                 '</hp:subList><hp:cellAddr colAddr="0" rowAddr="0"/></hp:tc></hp:tr></hp:tbl>'
        sec = form()["Contents/section0.xml"]
        mark = b'<hp:cellAddr colAddr="1" rowAddr="0"/>'
        before = {"Contents/section0.xml": sec.replace(b"</hp:subList>" + mark,
                                                       para(60, run("", extra=nested)).encode() + b"</hp:subList>" + mark)}
        self.assertNotEqual(before["Contents/section0.xml"], sec)
        lost = {k: v.replace(b"<hp:t>\xec\x95\x88\xec\xaa\xbd</hp:t>", b"<hp:t></hp:t>")
                for k, v in before.items()}
        filled = {k: v.replace(b'<hp:p id="4" paraPrIDRef="0"><hp:run charPrIDRef="0"><hp:t></hp:t>',
                               '<hp:p id="4" paraPrIDRef="0"><hp:run charPrIDRef="0"><hp:t>홍길동</hp:t>'.encode())
                  for k, v in lost.items()}
        self.assertIn("홍길동".encode(), filled["Contents/section0.xml"])
        self.assert_fails(before, filled, self.EDIT, "글자 밖 요소")


class NormalizationCounterexampleTests(EditCheck, unittest.TestCase):
    """정규화 항목마다 여전히 실패해야 하는 반례."""

    def test_letters_moved_between_runs_are_rejected(self):
        before = form(tail_runs=[run("가", "1"), run("나", "2")])
        self.assert_fails(before, self.after(tail_runs=[run("가나", "1"), run("", "2")]), self.EDIT, "대상 밖")

    def test_field_deletion_is_rejected(self):
        sec = fields_doc()["Contents/section0.xml"]
        plain = re.sub(rb"<hp:ctrl>.*?</hp:ctrl>", b"", sec)
        self.assertNotEqual(plain, sec)
        self.assert_fails({"Contents/section0.xml": sec}, {"Contents/section0.xml": plain}, [], "대상 밖")

    def test_picture_deletion_is_rejected(self):
        pic = '<hp:pic id="1"><hp:img binaryItemIDRef="image1"/></hp:pic>'
        before = form(tail_runs=[run("꼬리", extra=pic)])
        self.assert_fails(before, self.after(tail_runs=[run("꼬리")]), self.EDIT, "대상 밖")

    def test_char_shape_change_after_field_in_same_run_is_rejected(self):
        # fieldEnd와 뒤 글자가 한 run에 있을 때 그 run의 모양만 바꿔도 거부한다
        def doc(cp):
            begin = ('<hp:ctrl><hp:fieldBegin id="7" type="CLICK_HERE" name="성명" dirty="0" fieldid="70"/>'
                     '</hp:ctrl>')
            end = '<hp:ctrl><hp:fieldEnd beginIDRef="7" fieldid="70"/></hp:ctrl>'
            body = para(1, f'<hp:run charPrIDRef="0"><hp:t>앞 </hp:t>{begin}<hp:t>옛값</hp:t></hp:run>'
                           f'<hp:run charPrIDRef="{cp}">{end}<hp:t> 뒤</hp:t></hp:run>')
            return {"Contents/section0.xml": f"<hs:sec {SEC}>{body}</hs:sec>".encode()}
        same = {k: v.replace("옛값".encode(), "홍길동".encode()) for k, v in doc("0").items()}
        rep = self.check(doc("0"), same, [hwpx_gate.FieldEdit("성명", "홍길동")])
        self.assertEqual(rep.status, "passed", rep.messages)
        changed = {k: v.replace("옛값".encode(), "홍길동".encode()) for k, v in doc("5").items()}
        self.assert_fails(doc("0"), changed, [hwpx_gate.FieldEdit("성명", "홍길동")], "대상 밖")

    def test_textdir_rtl_loss_is_rejected(self):
        ppr = lambda d: f'<hh:paraPr id="0"{d}><hh:align horizontal="LEFT"/></hh:paraPr>'
        for value, ok in (("LTR", True), ("AUTO", True), ("RTL", False)):
            diff, known = hwpx_gate.split_rhwp_known("Contents/header.xml", header(ppr(f' textDir="{value}"')),
                                                     header(ppr("")), {"0"})
            lost = [k for k, (a, b) in diff.items() if b < a]
            self.assertEqual(not lost, ok, (value, diff))


class PackageGateTests(unittest.TestCase):
    """section 밖 파트: 다른 파트·content.hpf·header·바탕쪽 손실은 편집 대조(edits)에서 거부한다."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.orig = self.dir / "orig.hwpx"
        make_hwpx(self.orig)
        rewrite(self.orig, add={"Chart/chart0.xml": b"<c/>", "Contents/masterpage0.xml":
                                f'<hm:masterPage xmlns:hm="m" {NS}><hp:p id="1"><hp:run charPrIDRef="0">'
                                '<hp:t>대외비</hp:t></hp:run></hp:p></hm:masterPage>'.encode()})
        self.prod = self.dir / "prod.hwpx"

    def tearDown(self):
        self.tmp.cleanup()

    def gate(self, **kw):
        self.prod.write_bytes(self.orig.read_bytes())
        rewrite(self.prod, **kw)
        return hwpx_gate.gate(self.orig, self.prod, allow_rhwp=True, edits=[])

    def assert_fails(self, needle, **kw):
        rep = self.gate(**kw)
        self.assertEqual(rep.status, "failed", rep.messages)
        self.assertIn(needle, " / ".join(rep.messages))

    def assert_not_failed(self, **kw):
        rep = self.gate(**kw)
        self.assertNotEqual(rep.status, "failed", rep.messages)
        return rep

    def test_other_parts_must_survive(self):
        self.assert_fails("파트 사라짐", drop={"Chart/chart0.xml"})
        self.assert_fails("파트 바뀜", transform={"settings.xml": lambda d: d + b" "})
        self.assert_not_failed(transform={"Preview/PrvText.txt": lambda d: b"x"},
                               add={"META-INF/container.rdf.bak": b"x"})

    def test_content_hpf_metadata_manifest_and_spine(self):
        hpf = "Contents/content.hpf"
        self.assert_fails("메타데이터", transform={hpf: lambda d: d.replace(b"synthetic-fixture-author", b"")})
        self.assert_fails("매니페스트", transform={hpf: lambda d: re.sub(rb'<opf:item id="section0"[^>]*/>', b"", d)})
        two = lambda d: d.replace(b'<opf:itemref idref="section0"', b'<opf:itemref idref="x"/><opf:itemref idref="section0"')
        rewrite(self.orig, transform={hpf: two})
        self.assert_fails("구역 순서", transform={hpf: lambda d: d.replace(
            b'<opf:itemref idref="x"/><opf:itemref idref="section0"', b'<opf:itemref idref="section0"')})
        # rhwp가 바꾸는 것: 선언 공백·hashkey·linear·항목 순서
        self.assert_not_failed(transform={hpf: lambda d: d.replace(b'standalone="yes" ?>', b'standalone="yes"?>')
                                          .replace(b"</opf:manifest>", b'<opf:item id="n" href="BinData/n.png" media-type="image/png"/></opf:manifest>')})

    def test_header_swaps_text_and_unused_parapr_changes(self):
        h = "Contents/header.xml"
        def swap_fonts(d):
            a = re.search(rb'(<hh:charPr id="0".*?)(<hh:fontRef [^>]*/>)', d, re.S)
            b = re.search(rb'(<hh:charPr id="1".*?)(<hh:fontRef [^>]*/>)', d, re.S)
            self.assertNotEqual(a.group(2), b.group(2))
            d = d[:b.start(2)] + a.group(2) + d[b.end(2):]
            return d[:a.start(2)] + b.group(2) + d[a.end(2):]
        self.assert_fails("사라진 요소", transform={h: swap_fonts})
        self.assert_fails("사라진 요소", transform={h: lambda d: re.sub(rb"(<hh:paraHead[^>]*>)[^<]+", rb"\1", d)})
        # 본문이 쓰지 않는 paraPr라도 개요 갈래 접힘 말고 다른 변화(정렬)는 거부한다
        def realign(d):
            ids = re.findall(rb'<hh:paraPr id="(\d+)"', d)
            last = ids[-1]
            return re.sub(rb'(<hh:paraPr id="' + last + rb'".*?<hh:align horizontal=")\w+', rb"\1RIGHT", d, count=1, flags=re.S)
        self.assert_fails("사라진 요소", transform={h: realign})
        rep = self.assert_not_failed(transform={h: lambda d: d.replace(b' textDir="LTR"', b"")})
        self.assertIn("textDir LTR·AUTO 소실×20", rep.allowed_line())

    def test_master_page_text_is_compared(self):
        self.assert_fails("masterpage0.xml", transform={"Contents/masterpage0.xml": lambda d: d.replace("대외비".encode(), b"")})


class UsedParaPrTests(unittest.TestCase):
    def test_single_quotes_and_spaces_count_as_used(self):
        parts = {"Contents/section0.xml": f"<hs:sec {SEC}><hp:p paraPrIDRef = '7'/></hs:sec>".encode(),
                 "Contents/header.xml": b"<h/>"}
        self.assertEqual(hwpx_gate.used_parapr_ids(parts), {"7"})

    def test_unparsable_part_means_unknown_and_no_allowance(self):
        parts = {"Contents/section0.xml": b"<broken", "Contents/header.xml": b"<h/>"}
        self.assertIsNone(hwpx_gate.used_parapr_ids(parts))
        before = header(outline_para_pr(16))
        after = header(outline_para_pr(16, collapsed=True))
        diff, known = hwpx_gate.split_rhwp_known("Contents/header.xml", before, after, None)
        self.assertIn("switch", {k[0] for k, (a, b) in diff.items() if b < a})
        self.assertFalse(known)


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
