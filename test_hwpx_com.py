"""hwpx_com.py 단위 시험 (가짜 Hwp, 한컴 없이 돈다)."""
import contextlib
import io
import os
from pathlib import Path
import tempfile
import unittest
from unittest import mock

import hwpx_com


def stats(pages, tables=0, shapes=0, text="본문"):
    return {"pages": pages, "tables": tables, "shapes": shapes, "text": text}


class FakeHwp:
    """open한 경로를 기억하고, doc_stats 대역이 그 경로의 통계를 돌려준다."""

    def __init__(self, table, merged):
        self.table = table          # 원본 경로 → 통계
        self.merged = merged        # 합본 통계
        self.current = None
        self.inserted = []
        self.saved = None
        self.quit_called = False

    def open(self, path, fmt, arg):
        self.current = path
        return True

    @property
    def hwp(self):
        return self

    def MovePos(self, move_id, para, pos):
        assert move_id == 3
        return True

    def insert_file(self, path, **kw):
        assert kw["keep_section"] == 1
        self.inserted.append(path)
        return True

    def save_as(self, path, format):
        Path(path).write_bytes(b"PK fake")
        self.saved = path
        return True

    def clear(self, option):
        pass

    def quit(self):
        self.quit_called = True

    def stats(self, _hwp):
        if self.current == self.saved:
            return self.merged
        return self.table[self.current]


class CheckMergeTests(unittest.TestCase):
    def test_pass_when_sums_and_text_match(self):
        parts = [stats(2, 1, 1, "가 나\r\n"), stats(1, 0, 0, "다")]
        self.assertEqual(hwpx_com.check_merge(parts, stats(3, 1, 1, "가나\r\n\r\n다")), [])

    def test_extra_blank_page_fails(self):
        # BreakPage를 함께 넣으면 문서마다 빈 쪽이 끼는 사고(이슈 #14)
        parts = [stats(2), stats(2)]
        problems = hwpx_com.check_merge(parts, stats(3 + 2, text="본문본문"))
        self.assertTrue(any("쪽수" in p for p in problems))

    def test_lost_picture_and_reordered_text_fail(self):
        parts = [stats(1, shapes=2, text="첫"), stats(1, shapes=1, text="둘")]
        problems = hwpx_com.check_merge(parts, stats(2, shapes=2, text="둘첫"))
        self.assertEqual(len(problems), 2)


class AppendTests(unittest.TestCase):
    def run_append(self, tmp, merged, output=None):
        files = [os.path.join(tmp, n) for n in ("a.hwpx", "b 사진.hwpx", "c.hwp")]
        for f in files:
            Path(f).write_bytes(b"x")
        table = {f: stats(i + 1, text=f"글{i}") for i, f in enumerate(files)}
        fake = FakeHwp(table, merged)
        out, err = io.StringIO(), io.StringIO()
        with mock.patch.object(hwpx_com, "create_hwp", return_value=fake), \
                mock.patch.object(hwpx_com, "doc_stats", side_effect=fake.stats), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                hwpx_com.cmd_append(files[0], files[1:], output)
                code = 0
            except SystemExit as exc:
                code = exc.code
        return fake, files, code, out.getvalue(), err.getvalue()

    def test_default_output_and_insert_order(self):
        with tempfile.TemporaryDirectory() as tmp:
            fake, files, code, out, _ = self.run_append(tmp, stats(6, text="글0글1글2"))
            self.assertEqual(code, 0)
            self.assertEqual(fake.inserted, files[1:])
            expected = os.path.join(tmp, "_work-hwpx-automation", "a_합본.hwpx")
            self.assertEqual(fake.saved, expected)
            self.assertTrue(out.rstrip().endswith("engine=com"))
            self.assertTrue(fake.quit_called)
            # 기본 출력이 이미 있으면 덮지 않는다
            _, _, code, _, _ = self.run_append(tmp, stats(6, text="글0글1글2"))
            self.assertIn("이미 있습니다", str(code))

    def test_mismatch_moves_output_to_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            out_path = os.path.join(tmp, "합본.hwpx")
            _, _, code, _, err = self.run_append(tmp, stats(9, text="글0글1글2"), out_path)
            self.assertEqual(code, 1)
            self.assertFalse(os.path.exists(out_path))
            self.assertTrue(os.path.exists(os.path.join(tmp, "합본.rejected.hwpx")))
            self.assertIn("쪽수 9 (원본 합 6)", err)

    def test_empty_body_input_refused(self):
        with tempfile.TemporaryDirectory() as tmp:
            files = [os.path.join(tmp, n) for n in ("a.hwpx", "pandoc.hwpx")]
            for f in files:
                Path(f).write_bytes(b"x")
            fake = FakeHwp({files[0]: stats(1), files[1]: stats(1, text="")}, stats(2))
            with mock.patch.object(hwpx_com, "create_hwp", return_value=fake), \
                    mock.patch.object(hwpx_com, "doc_stats", side_effect=fake.stats):
                with self.assertRaises(SystemExit) as cm:
                    hwpx_com.cmd_append(files[0], files[1:], os.path.join(tmp, "o.hwpx"))
            self.assertIn("0자", str(cm.exception.code))
            self.assertIsNone(fake.saved)

    def test_in_place_output_refused(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "a.hwpx")
            b = os.path.join(tmp, "b.hwpx")
            for f in (a, b):
                Path(f).write_bytes(b"x")
            with self.assertRaises(SystemExit) as cm:
                hwpx_com.cmd_append(a, [b], b)
            self.assertIn("in-place", str(cm.exception.code))


if __name__ == "__main__":
    unittest.main()
