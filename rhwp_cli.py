"""rhwp CLI 어댑터: 이 저장소에서 rhwp를 부르는 유일한 입구.

rhwp(https://github.com/edwardkim/rhwp)는 HWP·HWPX를 한컴 없이 읽고 고치고 그리는
Rust 단일 바이너리다. 파이썬 바인딩이 없어 서브프로세스로만 쓴다. 이 모듈은
바이너리 탐색, 버전·계약 검사, `--json` 봉투 파싱, exit code 매핑, "임시 폴더에
쓰고 검증을 통과한 뒤 제자리로 옮기기"를 한곳에 모은다. 다른 모듈은 rhwp 명령
문자열을 직접 조립하지 않고 :func:`run`만 부른다.

계약 가정(v0.8.7 실측): 모든 명령이 `--json`으로 `schemaVersion: "1.0"` 봉투를
stdout에 내고, exit 0 성공 · 1 런타임 실패 · 2 사용법 오류 · 3 `--verify` 불일치 ·
4 `--verify-pages` 쪽수 불일치다. 봉투의 `untrustedFields`에 적힌 값은 문서에서
온 글자이므로 사람에게 보여 주되 분기 조건으로 쓰지 않는다.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field

MIN_VERSION = (0, 8, 7)  # export-hwpx 위치 인자·edit insert-image·--json 봉투를 실측한 버전
TESTED_MAX = (0, 8, 7)   # 이 위로는 경고만 낸다. 상한 차단은 2~4주 릴리스마다 도구를 멈추게 한다
SCHEMA_VERSION = "1.0"
OUTPUT = "<OUTPUT>"  # run()의 args에서 산출 파일 자리를 표시하는 자리표
ENV_BIN = "HWPX_RHWP_BIN"

INSTALL_HINT = ("rhwp CLI가 필요합니다. https://github.com/edwardkim/rhwp/releases 에서 "
                "플랫폼용 바이너리를 받아 SHA256SUMS.txt로 확인한 뒤 PATH에 두세요.")


class RhwpError(RuntimeError):
    """rhwp 호출 실패의 공통 부모."""


class RhwpNotInstalled(RhwpError):
    pass


class RhwpVersionError(RhwpError):
    pass


class RhwpUsageError(RhwpError):
    """exit 2. rhwp 인자를 잘못 조립한 우리 쪽 결함으로 취급한다."""


class RhwpRuntimeError(RhwpError):
    """exit 1 또는 알 수 없는 코드."""


class RhwpVerifyError(RhwpError):
    """exit 3·4. rhwp 자가 검증이 저장본 불일치·쪽수 불일치를 보고했다."""

    def __init__(self, message, rejected_path=None):
        super().__init__(message)
        self.rejected_path = rejected_path


class RhwpContractError(RhwpError):
    """exit 0인데 봉투·산출물이 계약과 다르다. 산출물은 게시하지 않는다."""


@dataclass
class RhwpResult:
    exit_code: int
    envelope: dict | None
    stdout: str
    stderr: str
    output: str | None = None
    warnings: list = field(default_factory=list)

    @property
    def ok(self):
        return self.exit_code == 0


def find_rhwp():
    """환경변수 HWPX_RHWP_BIN이 있으면 그 경로, 없으면 PATH의 rhwp. 없으면 None."""
    override = os.environ.get(ENV_BIN)
    if override:
        return override if os.access(override, os.X_OK) else None
    return shutil.which("rhwp")


def parse_version(text):
    m = re.search(r"v?(\d+)\.(\d+)\.(\d+)", text or "")
    return tuple(int(x) for x in m.groups()) if m else None


def version(rhwp=None):
    rhwp = rhwp or find_rhwp()
    if not rhwp:
        raise RhwpNotInstalled(INSTALL_HINT)
    r = subprocess.run([rhwp, "--version"], capture_output=True, text=True, timeout=30)
    ver = parse_version(r.stdout or r.stderr)
    if ver is None:
        raise RhwpVersionError(f"rhwp 버전을 읽을 수 없습니다: {(r.stdout or r.stderr).strip()[:200]}")
    return ver


def check_version(rhwp=None):
    """MIN_VERSION 미만이면 예외, TESTED_MAX 초과면 경고 문자열 목록을 돌려준다."""
    ver = version(rhwp)
    if ver < MIN_VERSION:
        raise RhwpVersionError(
            f"rhwp {fmt(ver)}는 너무 낮습니다. {fmt(MIN_VERSION)} 이상이 필요합니다. {INSTALL_HINT}")
    if ver > TESTED_MAX:
        return [f"rhwp {fmt(ver)}는 실측한 {fmt(TESTED_MAX)}보다 새 버전입니다. "
                "봉투 계약 검사로 어긋남을 잡지만 결과를 한 번 더 확인하세요."]
    return []


def fmt(ver):
    return "v" + ".".join(str(x) for x in ver)


def parse_envelope(stdout):
    """stdout에서 JSON 봉투를 꺼낸다. 앞뒤에 로그 줄이 섞여도 첫 '{'부터 읽는다."""
    text = stdout or ""
    start = text.find("{")
    if start < 0:
        return None
    try:
        return json.loads(text[start:])
    except json.JSONDecodeError:
        # 봉투 뒤에 로그가 붙은 경우: 마지막 '}'까지로 다시 시도
        end = text.rfind("}")
        if end > start:
            try:
                return json.loads(text[start:end + 1])
            except json.JSONDecodeError:
                return None
        return None


def run(args, *, output=None, password=None, required_fields=(), verify=None,
        timeout=300, expect_json=True, keep_rejected=True):
    """rhwp를 한 번 실행한다.

    args: rhwp 뒤에 붙는 인자 목록. 산출 파일 자리는 OUTPUT 자리표로 적는다.
          `--json`은 여기서 붙이므로 적지 않는다.
    output: 최종 산출 파일 경로. 주면 임시 폴더에 먼저 쓰고, 성공·계약 검사·verify를
            모두 통과한 뒤 os.replace로 옮긴다. 실패하면 기존 output은 건드리지 않는다.
    password: 암호 문서. `--password-stdin`으로 넘긴다(프로세스 목록 노출 방지).
    required_fields: exit 0 봉투에 반드시 있어야 하는 키. 없으면 RhwpContractError.
    verify: callable(staged_path) -> None. 예외를 던지면 산출물을 게시하지 않고 그대로 전파한다.
    keep_rejected: exit 3·4 산출물을 `<output>.rejected<ext>`로 남겨 원인 조사에 쓴다.
    """
    rhwp = find_rhwp()
    if not rhwp:
        raise RhwpNotInstalled(INSTALL_HINT)
    warnings = check_version(rhwp)

    cmd = [rhwp]
    staged = None
    stage_dir = None
    if output is not None:
        output = os.path.abspath(output)
        os.makedirs(os.path.dirname(output), exist_ok=True)
        stage_dir = tempfile.TemporaryDirectory(prefix="rhwp-", dir=os.path.dirname(output))
        staged = os.path.join(stage_dir.name, "staged" + os.path.splitext(output)[1])
        if OUTPUT not in args:
            stage_dir.cleanup()
            raise RhwpUsageError("output을 주면 args에 OUTPUT 자리표가 있어야 합니다.")
    cmd += [staged if a == OUTPUT else a for a in args]
    if expect_json and "--json" not in cmd:
        cmd.append("--json")
    if password:
        cmd.append("--password-stdin")

    try:
        r = subprocess.run(cmd, input=f"{password}\n" if password else None,
                           capture_output=True, text=True, timeout=timeout)
        envelope = parse_envelope(r.stdout) if expect_json else None
        result = RhwpResult(r.returncode, envelope, r.stdout, r.stderr, warnings=warnings)
        tail = (r.stderr or r.stdout).strip()[-500:]

        if r.returncode == 2:
            raise RhwpUsageError(f"rhwp 사용법 오류(exit 2): {tail}")
        if r.returncode in (3, 4):
            rejected = None
            if staged and keep_rejected and os.path.exists(staged):
                root, ext = os.path.splitext(output)
                rejected = f"{root}.rejected{ext}"
                os.replace(staged, rejected)
            what = "저장본 불일치(--verify)" if r.returncode == 3 else "쪽수 불일치(--verify-pages)"
            raise RhwpVerifyError(f"rhwp 자가 검증 실패, {what}: {tail}", rejected_path=rejected)
        if r.returncode != 0:
            raise RhwpRuntimeError(f"rhwp 실패(exit {r.returncode}): {tail}")

        if expect_json:
            if envelope is None:
                raise RhwpContractError(f"rhwp가 exit 0인데 JSON 봉투를 내지 않았습니다: {tail}")
            if envelope.get("schemaVersion") != SCHEMA_VERSION:
                raise RhwpContractError(
                    f"rhwp 봉투 schemaVersion이 {envelope.get('schemaVersion')!r}입니다. "
                    f"{SCHEMA_VERSION!r}만 지원합니다.")
            missing = [k for k in required_fields if k not in envelope]
            if missing:
                raise RhwpContractError(f"rhwp 봉투에 필수 필드가 없습니다: {missing}")
        if staged is not None:
            if not os.path.exists(staged) or os.path.getsize(staged) == 0:
                raise RhwpContractError("rhwp가 exit 0인데 산출 파일을 만들지 않았습니다.")
            if verify is not None:
                verify(staged)
            os.replace(staged, output)
            result.output = output
        return result
    finally:
        if stage_dir is not None:
            stage_dir.cleanup()


def untrusted(envelope, key):
    """봉투에서 문서 유래 값을 꺼낸다. 표시 전용이며 분기 조건으로 쓰지 않는다."""
    return (envelope or {}).get(key)
