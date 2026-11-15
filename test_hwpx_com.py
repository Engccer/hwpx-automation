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


class FakePrintSet:
    def __init__(self, accept_printer=True):
        self.HSet = object()
        self._accept = accept_printer

    def __setattr__(self, key, value):
        if key == "PrinterName" and not self.__dict__.get("_accept", True):
            value = "다른 프린터"  # 한글이 이름을 무시하는 경우
        object.__setattr__(self, key, value)


class FakePrintHwp:
    """HPrint 대역. 실제 스풀러·한컴 없이 인쇄 흐름만 본다."""

    def __init__(self, log, pages=2, ok=True, accept_printer=True):
        self.log, self.PageCount, self.ok = log, pages, ok
        self.pset = FakePrintSet(accept_printer)
        outer = self

        class Params:
            HPrint = outer.pset

        class Action:
            def GetDefault(self, name, hset):
                outer.log.append(("GetDefault", name))

            def Execute(self, name, hset):
                params = {k: v for k, v in vars(outer.pset).items() if not k.startswith("_")}
                outer.log.append(("Execute", name, params))
                return outer.ok

        self.HParameterSet, self.HAction = Params(), Action()

    @property
    def hwp(self):
        return self

    def open(self, path, fmt, arg):
        self.log.append(("open", os.path.basename(path)))
        return True

    def clear(self, option):
        pass

    def quit(self):
        self.log.append(("quit",))


class FakeWatcher:
    def __init__(self, lines):
        self.lines = lines

    def communicate(self, timeout=None):
        return "\n".join(self.lines), ""

    def poll(self):
        return 0


PRINTER = {"Name": "가상 프린터", "Default": True, "PortName": "IP_1", "DuplexingMode": "OneSided",
           "Collate": True, "DefaultPrinter": "가상 프린터"}


class QueueTests(unittest.TestCase):
    def test_spooling_observation_is_not_final(self):
        jobs, end = hwpx_com.parse_queue_watch(["JOB 7 6 0 Spooling", "EMPTY"])
        self.assertEqual(jobs, {"7": {"pages": None}})
        status, desc = hwpx_com.queue_verdict(jobs, end, 6, 2)
        self.assertEqual(status, "unverified")

    def test_pages_times_copies_pass(self):
        lines = ["JOB 7 3 0 Spooling", "JOB 7 12 0 Paused", "JOB 7 12 4 Printing", "EMPTY"]
        jobs, end = hwpx_com.parse_queue_watch(lines)
        self.assertEqual(hwpx_com.queue_verdict(jobs, end, 6, 2)[0], "passed")

    def test_one_copy_pages_is_unverified_not_passed(self):
        # 드라이버가 부수를 맡는 경우와 한 부만 나간 경우가 구별되지 않는다
        jobs, end = hwpx_com.parse_queue_watch(["JOB 7 6 0 Paused", "EMPTY"])
        status, desc = hwpx_com.queue_verdict(jobs, end, 6, 2)
        self.assertEqual(status, "unverified")
        self.assertIn("2부", desc)
        self.assertEqual(hwpx_com.queue_verdict(jobs, end, 6, 1)[0], "passed")

    def test_wrong_pages_mismatch_and_nothing_seen_unverified(self):
        jobs, end = hwpx_com.parse_queue_watch(["JOB 7 9 0 Printing", "EMPTY"])
        self.assertEqual(hwpx_com.queue_verdict(jobs, end, 6, 2)[0], "mismatch")
        self.assertEqual(hwpx_com.queue_verdict({}, "NONE", 6, 2)[0], "unverified")

    def test_sheet_count(self):
        self.assertEqual(hwpx_com.sheet_count(2, 30, "TwoSidedLongEdge"), 30)
        self.assertEqual(hwpx_com.sheet_count(3, 2, "TwoSidedShortEdge"), 4)
        self.assertEqual(hwpx_com.sheet_count(3, 2, "OneSided"), 6)


class PrintTests(unittest.TestCase):
    def run_print(self, info=PRINTER, watch=("JOB 1 4 0 Printing", "EMPTY"), ok=True,
                  default_after=None, watcher=None, accept_printer=True, **kw):
        log = []
        fake = FakePrintHwp(log, pages=2, ok=ok, accept_printer=accept_printer)

        def set_duplex(name, mode):
            log.append(("duplex", mode))

        def restore_default(original):
            log.append(("restore_default", original))
            return default_after

        def start_watch(*args):
            log.append(("watch",))
            return watcher or FakeWatcher(list(watch))

        out, err = io.StringIO(), io.StringIO()
        with tempfile.TemporaryDirectory() as tmp:
            doc = os.path.join(tmp, "가정 통신문.hwpx")
            Path(doc).write_bytes(b"x")
            with contextlib.ExitStack() as stack:
                for name, kw2 in (("printer_info", {"return_value": dict(info)}),
                                  ("set_duplex", {"side_effect": set_duplex}),
                                  ("restore_default_printer", {"side_effect": restore_default}),
                                  ("start_queue_watch", {"side_effect": start_watch}),
                                  ("create_hwp", {"return_value": fake})):
                    stack.enter_context(mock.patch.object(hwpx_com, name, **kw2))
                stack.enter_context(contextlib.redirect_stdout(out))
                stack.enter_context(contextlib.redirect_stderr(err))
                try:
                    hwpx_com.cmd_print(doc, **kw)
                    code = 0
                except SystemExit as exc:
                    code = exc.code
        return log, code, out.getvalue(), err.getvalue()

    def test_dry_changes_nothing_and_reports_plan(self):
        log, code, out, _ = self.run_print(copies=30, duplex="long", dry=True)
        self.assertEqual(code, 0)
        kinds = [e[0] for e in log]
        self.assertNotIn("duplex", kinds)
        self.assertNotIn("Execute", kinds)
        self.assertNotIn("restore_default", kinds)
        self.assertIn("프린터=가상 프린터, 부수=30, 양면=TwoSidedLongEdge(지금 OneSided", out)
        self.assertIn("쪽수=2, 용지=30장", out)
        self.assertTrue(out.rstrip().endswith("engine=com status=dry"))
        self.assertEqual(log[0], ("open", "print_input.hwpx"))  # ASCII 임시 사본을 연다

    def test_duplex_set_before_print_and_restored_after(self):
        log, code, out, _ = self.run_print(copies=2, duplex="long",
                                           watch=("JOB 1 4 0 Printing", "EMPTY"))
        self.assertEqual(code, 0)
        kinds = [e[0] for e in log]
        self.assertLess(kinds.index("duplex"), kinds.index("Execute"))
        self.assertEqual([e for e in log if e[0] == "duplex"],
                         [("duplex", "TwoSidedLongEdge"), ("duplex", "OneSided")])
        params = next(e[2] for e in log if e[0] == "Execute")
        self.assertEqual((params["PrinterName"], params["NumCopy"], params["Collate"], params["Range"]),
                         ("가상 프린터", 2, 1, 0))
        self.assertEqual(params["PrintToFile"], 0)  # 지난 실행 값이 남아도 종이로
        self.assertIn("양면 설정 복원: OneSided", out)
        self.assertTrue(out.rstrip().endswith("engine=com status=passed"))

    def test_same_duplex_is_not_touched(self):
        log, _, out, _ = self.run_print(duplex="none")
        self.assertNotIn("duplex", [e[0] for e in log])
        self.assertIn("프린터 설정 그대로", out)

    def test_execute_failure_still_restores(self):
        log, code, _, _ = self.run_print(duplex="short", ok=False)
        self.assertIn("Execute=False", str(code))
        self.assertEqual(log[-2:], [("duplex", "OneSided"), ("restore_default", "가상 프린터")])

    def test_mismatch_exits_1_and_default_printer_restored(self):
        log, code, out, _ = self.run_print(watch=("JOB 1 5 0 Printing", "EMPTY"),
                                           default_after="가상 프린터")
        self.assertEqual(code, 1)
        self.assertIn("기본 프린터 복원: 가상 프린터", out)
        self.assertTrue(out.rstrip().endswith("engine=com status=mismatch"))

    def test_port_prompt_printer_needs_file(self):
        info = dict(PRINTER, PortName="PORTPROMPT:")
        log, code, _, _ = self.run_print(info=info)
        self.assertIn("--print-to-file", str(code))
        self.assertEqual(log, [])
        log, code, _, _ = self.run_print(info=info, to_file="out.pdf")
        params = next(e[2] for e in log if e[0] == "Execute")
        self.assertEqual(params["PrintToFile"], 1)
        self.assertTrue(params["filename"].endswith("out.pdf"))

    def test_printer_name_not_accepted_aborts_before_execute(self):
        log, code, _, _ = self.run_print(duplex="long", accept_printer=False)
        self.assertIn("받지 않았습니다", str(code))
        kinds = [e[0] for e in log]
        self.assertNotIn("Execute", kinds)
        self.assertNotIn("restore_default", kinds)
        self.assertEqual(log[-1], ("duplex", "OneSided"))

    def test_watch_timeout_still_prints_result_line(self):
        import subprocess

        class SlowWatcher:
            killed = False

            def communicate(self, timeout=None):
                if timeout is not None:
                    raise subprocess.TimeoutExpired("powershell", timeout)
                return "JOB 1 2 0 Printing", ""

            def kill(self):
                SlowWatcher.killed = True

            def poll(self):
                return 0

        log, code, out, _ = self.run_print(watcher=SlowWatcher())
        self.assertEqual(code, 0)
        self.assertTrue(SlowWatcher.killed)
        self.assertIn("비지 않음", out)
        self.assertTrue(out.rstrip().endswith("engine=com status=passed"))

    def test_duplex_set_failure_does_not_claim_restore(self):
        def fail(name, mode):
            raise RuntimeError("권한 없음")
        log = []
        err = io.StringIO()
        with tempfile.TemporaryDirectory() as tmp:
            doc = os.path.join(tmp, "a.hwpx")
            Path(doc).write_bytes(b"x")
            with mock.patch.object(hwpx_com, "printer_info", return_value=dict(PRINTER)), \
                    mock.patch.object(hwpx_com, "set_duplex", side_effect=fail), \
                    mock.patch.object(hwpx_com, "create_hwp") as create, \
                    contextlib.redirect_stderr(err):
                with self.assertRaises(SystemExit) as cm:
                    hwpx_com.cmd_print(doc, duplex="long")
        self.assertIn("양면 설정을 바꾸지 못해", str(cm.exception.code))
        self.assertNotIn("되돌리지 못했습니다", err.getvalue())
        create.assert_not_called()

    def test_print_to_file_must_differ_from_input(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc = os.path.join(tmp, "a.hwpx")
            Path(doc).write_bytes(b"x")
            with self.assertRaises(SystemExit):
                hwpx_com.cmd_print(doc, to_file=doc.upper())

    def test_copies_must_be_positive(self):
        with self.assertRaises(SystemExit):
            hwpx_com.cmd_print("x.hwpx", copies=0)


if __name__ == "__main__":
    unittest.main()
