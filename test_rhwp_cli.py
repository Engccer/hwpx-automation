import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

import rhwp_cli
from rhwp_cli import (OUTPUT, RhwpContractError, RhwpNotInstalled, RhwpRuntimeError,
                      RhwpUsageError, RhwpVerifyError, RhwpVersionError, run)

FAKE_RHWP = """#!{python}
import json, os, sys
mode = os.environ.get("FAKE_RHWP_MODE", "ok")
if sys.argv[1:] == ["--version"]:
    print("rhwp " + os.environ.get("FAKE_RHWP_VERSION", "v0.8.7")); sys.exit(0)
out = None
for i, a in enumerate(sys.argv):
    if a.endswith("staged.hwpx") or a.endswith("staged.pdf"):
        out = a
if mode == "password":
    if sys.stdin.readline().strip() != "s3cret":
        sys.stderr.write("암호 불일치\\n"); sys.exit(1)
    mode = "ok"
if mode == "fail":
    sys.stderr.write("오류: 문서 로드 실패\\n"); sys.exit(1)
if mode == "usage":
    sys.stderr.write("사용법: ...\\n"); sys.exit(2)
if mode == "verify":
    open(out, "wb").write(b"PK-bad")
    print(json.dumps({"schemaVersion": "1.0", "verify": {"identical": False}}))
    sys.exit(3)
if mode == "pages":
    sys.exit(4)
if mode == "nojson":
    open(out, "wb").write(b"PK"); print("done"); sys.exit(0)
if mode == "oldschema":
    open(out, "wb").write(b"PK"); print(json.dumps({"schemaVersion": "0.9"})); sys.exit(0)
if mode == "nooutput":
    print(json.dumps({"schemaVersion": "1.0", "output": "x"})); sys.exit(0)
if out:
    open(out, "wb").write(b"PK-ok")
env = {"schemaVersion": "1.0", "output": out, "format": "hwpx", "changedPages": [0],
       "untrustedFields": ["oldText"], "oldText": "문서에서 온 글자"}
print("로그 한 줄"); print(json.dumps(env, ensure_ascii=False))
"""


@unittest.skipIf(sys.platform == "win32", "가짜 rhwp 실행 파일이 shebang에 의존")
class RunTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        tmp = Path(self.tmp.name)
        self.fake = tmp / "rhwp"
        self.fake.write_text(FAKE_RHWP.replace("{python}", sys.executable), encoding="utf-8")
        self.fake.chmod(0o755)
        self.out = tmp / "out" / "result.hwpx"
        self.saved = {k: os.environ.get(k) for k in ("HWPX_RHWP_BIN", "FAKE_RHWP_MODE", "FAKE_RHWP_VERSION")}
        os.environ["HWPX_RHWP_BIN"] = str(self.fake)
        os.environ.pop("FAKE_RHWP_MODE", None)
        os.environ.pop("FAKE_RHWP_VERSION", None)

    def tearDown(self):
        for k, v in self.saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        self.tmp.cleanup()

    def test_success_moves_output_and_parses_envelope(self):
        r = run(["export-hwpx", "in.hwp", OUTPUT], output=self.out, required_fields=("output", "format"))
        self.assertTrue(r.ok)
        self.assertEqual(r.output, str(self.out))
        self.assertEqual(self.out.read_bytes(), b"PK-ok")
        self.assertEqual(r.envelope["format"], "hwpx")
        self.assertEqual(rhwp_cli.untrusted(r.envelope, "oldText"), "문서에서 온 글자")
        self.assertEqual(sorted(p.name for p in self.out.parent.iterdir()), ["result.hwpx"])

    def test_runtime_failure_preserves_existing_output(self):
        os.environ["FAKE_RHWP_MODE"] = "fail"
        self.out.parent.mkdir()
        self.out.write_bytes(b"old")
        with self.assertRaisesRegex(RhwpRuntimeError, "문서 로드 실패"):
            run(["export-hwpx", "in.hwp", OUTPUT], output=self.out)
        self.assertEqual(self.out.read_bytes(), b"old")

    def test_usage_error_is_our_bug(self):
        os.environ["FAKE_RHWP_MODE"] = "usage"
        with self.assertRaises(RhwpUsageError):
            run(["edit", "set-cell", "x.hwpx", OUTPUT], output=self.out)

    def test_verify_failure_keeps_rejected_copy(self):
        os.environ["FAKE_RHWP_MODE"] = "verify"
        with self.assertRaises(RhwpVerifyError) as cm:
            run(["export-hwpx", "in.hwp", OUTPUT, "--verify"], output=self.out)
        self.assertFalse(self.out.exists())
        self.assertTrue(cm.exception.rejected_path.endswith("result.rejected.hwpx"))
        self.assertEqual(Path(cm.exception.rejected_path).read_bytes(), b"PK-bad")

    def test_pages_mismatch_is_verify_error(self):
        os.environ["FAKE_RHWP_MODE"] = "pages"
        with self.assertRaisesRegex(RhwpVerifyError, "쪽수"):
            run(["export-hwpx", "in.hwp", OUTPUT, "--verify-pages"], output=self.out)

    def test_missing_envelope_is_contract_error(self):
        os.environ["FAKE_RHWP_MODE"] = "nojson"
        with self.assertRaisesRegex(RhwpContractError, "JSON 봉투"):
            run(["export-hwpx", "in.hwp", OUTPUT], output=self.out)
        self.assertFalse(self.out.exists())

    def test_schema_mismatch_is_contract_error(self):
        os.environ["FAKE_RHWP_MODE"] = "oldschema"
        with self.assertRaisesRegex(RhwpContractError, "schemaVersion"):
            run(["export-hwpx", "in.hwp", OUTPUT], output=self.out)

    def test_missing_required_field_is_contract_error(self):
        with self.assertRaisesRegex(RhwpContractError, "필수 필드"):
            run(["export-hwpx", "in.hwp", OUTPUT], output=self.out, required_fields=("verify",))

    def test_exit_zero_without_output_is_contract_error(self):
        os.environ["FAKE_RHWP_MODE"] = "nooutput"
        with self.assertRaisesRegex(RhwpContractError, "산출 파일"):
            run(["export-hwpx", "in.hwp", OUTPUT], output=self.out)

    def test_verify_callback_blocks_publication(self):
        def reject(path):
            raise ValueError("게이트 실패")
        with self.assertRaises(ValueError):
            run(["export-hwpx", "in.hwp", OUTPUT], output=self.out, verify=reject)
        self.assertFalse(self.out.exists())

    def test_password_goes_through_stdin(self):
        os.environ["FAKE_RHWP_MODE"] = "password"
        r = run(["export-hwpx", "in.hwp", OUTPUT], output=self.out, password="s3cret")
        self.assertTrue(r.ok)

    def test_old_version_is_rejected(self):
        os.environ["FAKE_RHWP_VERSION"] = "v0.8.6"
        with self.assertRaisesRegex(RhwpVersionError, "0.8.7"):
            run(["info", "x.hwpx"])

    def test_newer_version_only_warns(self):
        os.environ["FAKE_RHWP_VERSION"] = "v0.9.3"
        r = run(["export-hwpx", "in.hwp", OUTPUT], output=self.out)
        self.assertTrue(r.ok)
        self.assertTrue(any("새 버전" in w for w in r.warnings))

    def test_missing_binary_explains_install(self):
        os.environ["HWPX_RHWP_BIN"] = str(self.fake) + ".missing"
        with self.assertRaisesRegex(RhwpNotInstalled, "releases"):
            run(["info", "x.hwpx"])

    def test_output_without_placeholder_is_usage_error(self):
        with self.assertRaises(RhwpUsageError):
            run(["export-hwpx", "in.hwp", "out.hwpx"], output=self.out)

    def test_run_without_output_returns_envelope(self):
        r = run(["info", "x.hwpx"])
        self.assertIsNone(r.output)
        self.assertEqual(r.envelope["schemaVersion"], "1.0")


class ParseTests(unittest.TestCase):
    def test_parse_version(self):
        self.assertEqual(rhwp_cli.parse_version("rhwp v0.8.7 - HWP 파일 뷰어"), (0, 8, 7))
        self.assertIsNone(rhwp_cli.parse_version("garbage"))

    def test_parse_envelope_with_trailing_log(self):
        self.assertEqual(rhwp_cli.parse_envelope('x\n{"a": 1}\n경고 줄'), {"a": 1})
        self.assertIsNone(rhwp_cli.parse_envelope("no json"))


if __name__ == "__main__":
    unittest.main()
