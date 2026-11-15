#!/usr/bin/env python3
"""hwpx_com.py - 한컴 COM 네이티브 HWPX 파이프라인 CLI (pyhwpx 기반)

XML/Pandoc 파이프라인(hwpx_edit.py 편집 명령, convert/hwpx_convert.py)과
별도로 운용하는 두 번째 백엔드. 두 파이프라인을 섞지 않기 위해 파일을 분리했다.

핵심 규칙 (reference/warnings-com.md):
- 이 스크립트가 생성한 HWPX는 COM·XML 양쪽 파이프라인에서 유효 (2026-06-10 검증)
- Pandoc(hwpx_convert.py)이 생성한 HWPX를 COM으로 열어 같은 파일에 저장하면
  본문이 전부 소실된다. 그래서 이 스크립트는 입력 파일을 절대 in-place로
  덮어쓰지 않고 항상 별도 출력 파일에 저장한다.
- pyhwpx의 get_text_file()은 기본 option이 'saveblock:true'(선택 블록만)라서
  선택이 없으면 None을 반환한다. 전체 본문은 GetTextFile(format, "")로 추출.

요구 사항: Windows + 한컴오피스 + pyhwpx(pywin32 포함). 진단: --diagnose
"""

import argparse
import os
import re
import sys

BODY_HEIGHT_PT = 10          # 본문 글자 크기 (바탕글 기본값)
BYPRODUCT_DIR = "_work-hwpx-automation"
HEADING_HEIGHT_PT = {1: 16, 2: 14, 3: 12}

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def create_hwp(visible=False):
    """pyhwpx Hwp 인스턴스를 생성 (새 문서, 보안 모듈 자동 등록)."""
    try:
        from pyhwpx import Hwp
    except ImportError as exc:
        raise SystemExit(
            "pyhwpx가 필요합니다. `pip install pyhwpx` 후 다시 실행하세요."
        ) from exc
    return Hwp(new=True, visible=visible, register_module=True)


def get_output_path(filepath, suffix, ext):
    """<입력>_<suffix>.<ext> 형식의 기본 출력 경로."""
    base = os.path.splitext(filepath)[0]
    return f"{base}_{suffix}{ext}"


def open_for_com(hwp, filepath, password=None):
    """COM으로 파일을 열고 실패 시 예외. HWP/HWPX 포맷 자동 판별."""
    abs_input = os.path.abspath(filepath)
    if not os.path.exists(abs_input):
        raise SystemExit(f"오류: 파일이 없습니다: {abs_input}")
    fmt = "HWP" if abs_input.lower().endswith(".hwp") else "HWPX"
    arg = f"password:{password}" if password else ""
    if not hwp.open(abs_input, fmt, arg):
        raise SystemExit(f"오류: 한컴 COM이 파일을 열지 못했습니다: {abs_input}")


def extract_full_text(hwp):
    """전체 본문 텍스트. pyhwpx 래퍼의 saveblock 함정을 우회해 저수준 호출."""
    return hwp.hwp.GetTextFile("TEXT", "") or ""


# ---------------------------------------------------------------------------
# Markdown 부분집합 파서 (--from-md)
# ---------------------------------------------------------------------------

TABLE_SEP_RE = re.compile(r"^\s*\|?[\s:\-|]+\|?\s*$")


def parse_md_blocks(md_text):
    """MD를 블록 리스트로 변환.

    지원: 제목(#/##/###), 본문 단락, 불릿(-/*), 파이프 표, **굵게** 인라인.
    그 외 문법은 일반 텍스트로 취급한다.
    """
    blocks = []
    lines = md_text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        m = re.match(r"^(#{1,3})\s+(.*)$", stripped)
        if m:
            blocks.append(("heading", len(m.group(1)), m.group(2).strip()))
            i += 1
            continue

        if stripped.startswith("|") and i + 1 < len(lines) and TABLE_SEP_RE.match(lines[i + 1].strip()) and lines[i + 1].strip().startswith(("|", ":", "-")):
            rows = [split_table_row(stripped)]
            i += 2  # 헤더 + 구분선 소비
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(split_table_row(lines[i].strip()))
                i += 1
            cols = max(len(r) for r in rows)
            rows = [r + [""] * (cols - len(r)) for r in rows]
            blocks.append(("table", rows))
            continue

        m = re.match(r"^[-*]\s+(.*)$", stripped)
        if m:
            blocks.append(("bullet", m.group(1).strip()))
            i += 1
            continue

        blocks.append(("para", stripped))
        i += 1
    return blocks


def split_table_row(line):
    return [c.strip() for c in line.strip().strip("|").split("|")]


def insert_inline(hwp, text):
    """**굵게** 인라인 토글을 처리하며 텍스트 삽입."""
    parts = text.split("**")
    for idx, part in enumerate(parts):
        if not part:
            continue
        bold = idx % 2 == 1 and len(parts) >= 3
        if bold:
            hwp.set_font(Bold=True)
        hwp.insert_text(part)
        if bold:
            hwp.set_font(Bold=False)


def render_blocks(hwp, blocks):
    """파싱된 블록을 COM으로 순차 렌더링 (문서를 위에서 아래로 구축)."""
    for block in blocks:
        kind = block[0]
        if kind == "heading":
            _, level, text = block
            hwp.set_font(Bold=True, Height=HEADING_HEIGHT_PT[level])
            hwp.insert_text(text)
            hwp.set_font(Bold=False, Height=BODY_HEIGHT_PT)
            hwp.BreakPara()
        elif kind == "para":
            insert_inline(hwp, block[1])
            hwp.BreakPara()
        elif kind == "bullet":
            hwp.insert_text("• ")
            insert_inline(hwp, block[1])
            hwp.BreakPara()
        elif kind == "table":
            rows = block[1]
            hwp.create_table(rows=len(rows), cols=len(rows[0]), treat_as_char=True, header=True)
            flat = [cell for row in rows for cell in row]
            for idx, cell in enumerate(flat):
                if cell:
                    insert_inline(hwp, cell)
                if idx < len(flat) - 1:
                    hwp.TableRightCell()
            # 문서를 순차 구축하므로 표가 현재 문서 끝 → MoveDocEnd로 표 탈출
            hwp.MoveDocEnd()
            hwp.BreakPara()


# ---------------------------------------------------------------------------
# 명령
# ---------------------------------------------------------------------------

def cmd_diagnose():
    """pyhwpx 기반 COM 진단 (기동·종료까지 실측)."""
    try:
        import pyhwpx
        print(f"- pyhwpx: OK ({getattr(pyhwpx, '__version__', '버전 불명')})")
    except ImportError as exc:
        print(f"- pyhwpx: 오류 ({exc})")
        return 1
    hwp = None
    try:
        hwp = create_hwp(visible=False)
        print("- HWPFrame.HwpObject: OK (visible=False 기동)")
        print(f"- 한컴 버전: {hwp.Version}")
        return 0
    except Exception as exc:
        print(f"- HWPFrame.HwpObject: 오류 ({exc})")
        return 1
    finally:
        if hwp is not None:
            hwp.quit()


def cmd_from_md(output_hwpx, md_path):
    """Markdown 부분집합 → COM 네이티브 HWPX 생성."""
    if not os.path.exists(md_path):
        raise SystemExit(f"오류: 입력 MD가 없습니다: {md_path}")
    with open(md_path, encoding="utf-8") as f:
        md_text = f.read()
    blocks = parse_md_blocks(md_text)
    if not blocks:
        raise SystemExit("오류: 입력 MD에서 렌더링할 블록을 찾지 못했습니다.")

    abs_output = os.path.abspath(output_hwpx)
    os.makedirs(os.path.dirname(abs_output) or ".", exist_ok=True)

    hwp = create_hwp(visible=False)
    try:
        render_blocks(hwp, blocks)
        if not hwp.save_as(abs_output, format="HWPX"):
            raise SystemExit(f"오류: HWPX 저장 실패: {abs_output}")
    finally:
        hwp.quit()

    size = os.path.getsize(abs_output)
    n_tables = sum(1 for b in blocks if b[0] == "table")
    print(f"생성 완료: {abs_output} ({size:,}바이트, 블록 {len(blocks)}개, 표 {n_tables}개)")


def cmd_insert_image(filepath, image_path, output=None, width=0, height=0, password=None):
    """기존 HWP/HWPX 문서 끝에 이미지 삽입 후 '별도 파일'로 저장.

    in-place 저장은 Pandoc HWPX 본문 소실 사고를 막기 위해 지원하지 않는다.
    """
    abs_image = os.path.abspath(image_path)
    if not os.path.exists(abs_image):
        raise SystemExit(f"오류: 이미지가 없습니다: {abs_image}")
    if output is None:
        output = get_output_path(filepath, "img", ".hwpx")
    abs_output = os.path.abspath(output)
    if abs_output == os.path.abspath(filepath):
        raise SystemExit("오류: 입력과 같은 경로로는 저장할 수 없습니다 (in-place 금지, warnings-com.md 4번).")

    sizeoption = 1 if (width or height) else 0
    hwp = create_hwp(visible=False)
    try:
        open_for_com(hwp, filepath, password)
        hwp.MoveDocEnd()
        hwp.insert_picture(abs_image, sizeoption=sizeoption, width=width, height=height)
        if not hwp.save_as(abs_output, format="HWPX"):
            raise SystemExit(f"오류: HWPX 저장 실패: {abs_output}")
    finally:
        hwp.quit()
    print(f"이미지 삽입 완료: {abs_output}")


def cmd_get_text(filepath, password=None):
    """COM 기준 본문 텍스트 출력. Pandoc HWPX 호환성 점검에도 사용 (0자면 비호환)."""
    hwp = create_hwp(visible=False)
    try:
        open_for_com(hwp, filepath, password)
        text = extract_full_text(hwp)
    finally:
        hwp.quit()
    print(f"[COM 본문 길이: {len(text)}자]")
    if text:
        print(text)
    else:
        print("경고: COM이 본문을 0자로 읽었습니다. Pandoc 생성 HWPX일 가능성이 큽니다 (warnings-com.md 4번).", file=sys.stderr)


def cmd_normalize(filepath, output=None, password=None):
    """COM으로 열어 '별도 파일'로 재저장 (한컴 정규화).

    XML 파이프라인(python-hwpx 등)으로 생성·편집한 HWPX를 한컴이 직접 다시
    저장하게 하여 lineseg·미리보기(PrvText/PrvImage)·내부 캐시를 모두 한컴
    네이티브 상태로 재계산한다. 인쇄·배포 직전 마무리 단계로 사용.
    in-place 저장은 지원하지 않는다 (warnings-com.md 4번 사고 방지).
    """
    if output is None:
        output = get_output_path(filepath, "norm", ".hwpx")
    abs_output = os.path.abspath(output)
    if abs_output == os.path.abspath(filepath):
        raise SystemExit("오류: 입력과 같은 경로로는 저장할 수 없습니다 (in-place 금지).")
    os.makedirs(os.path.dirname(abs_output) or ".", exist_ok=True)

    # 같은 프로세스에서 Hwp 인스턴스를 두 번 만들면 COM 재초기화 segfault/연결
    # 오류가 난다(주의사항 2번). 검증까지 한 인스턴스로 처리한다.
    hwp = create_hwp(visible=False)
    try:
        open_for_com(hwp, filepath, password)
        text_before = extract_full_text(hwp)
        n_before = len(text_before)
        if not text_before.strip():
            # COM이 본문을 0자로 읽으면 전후 비교가 "0자 → 0자"로 통과해 버린다.
            # 이 경우(대개 Pandoc 생성물)는 저장하지 않는다(warnings-com.md 4번).
            raise SystemExit("오류: COM이 입력 본문을 0자로 읽었습니다. 입력이 COM 비호환"
                             "(예: Pandoc 생성물)이라 정규화하지 않습니다 (warnings-com.md 4번).")
        if not hwp.save_as(abs_output, format="HWPX"):
            raise SystemExit(f"오류: HWPX 저장 실패: {abs_output}")
        # save_as 후 현재 문서가 출력본이므로 같은 인스턴스에서 재열어 검증
        open_for_com(hwp, abs_output)
        text_after = extract_full_text(hwp)
        n_after = len(text_after)
    finally:
        hwp.quit()
    if not text_after.strip():
        raise SystemExit("오류: 정규화 후 본문이 0자입니다. 입력이 COM 비호환(예: Pandoc 생성물)일 수 있습니다.")
    print(f"정규화 완료: {abs_output} (본문 {n_before}자 → {n_after}자)")


def cmd_from_hwp(filepath, output=None, password=None):
    """HWP(바이너리) → HWPX 변환 (한컴 COM SaveAs).

    convert/hwp2hwpx.bat(JDK 21)이 정본이고, 이 명령은 JDK가 없는 Windows에서
    한컴오피스만으로 변환하는 폴백이다(이슈 #3). 변환 후 본문 글자 수를 재열어
    검증한다. 연속 호출 시 잔류 Hwp.exe가 다음 호출을 수 분간 멈추게 할 수
    있으므로 배치 변환 전에는 잔류 프로세스를 정리하라(warnings-com.md 9번).
    """
    if not filepath.lower().endswith(".hwp"):
        raise SystemExit("오류: --from-hwp 입력은 .hwp 파일이어야 합니다.")
    defaulted = output is None
    if defaulted:
        output = os.path.splitext(filepath)[0] + ".hwpx"
    abs_output = os.path.abspath(output)
    if not abs_output.lower().endswith(".hwpx"):
        raise SystemExit("오류: 출력 파일은 .hwpx 확장자여야 합니다 (저장 포맷이 HWPX 고정).")
    if abs_output == os.path.abspath(filepath):
        raise SystemExit("오류: 입력과 같은 경로로는 저장할 수 없습니다.")
    if defaulted and os.path.exists(abs_output):
        raise SystemExit(
            f"오류: 기본 출력 경로에 파일이 이미 있습니다: {abs_output}\n"
            "  덮어쓰려면 -o로 경로를 명시하세요."
        )
    os.makedirs(os.path.dirname(abs_output) or ".", exist_ok=True)

    hwp = create_hwp(visible=False)
    try:
        open_for_com(hwp, filepath, password)
        n_before = len(extract_full_text(hwp))
        if not hwp.save_as(abs_output, format="HWPX"):
            raise SystemExit(f"오류: HWPX 저장 실패: {abs_output}")
        # save_as 후 현재 문서가 출력본 → 같은 인스턴스에서 재열어 검증
        open_for_com(hwp, abs_output)
        n_after = len(extract_full_text(hwp))
    finally:
        hwp.quit()
    if n_before and not n_after:
        raise SystemExit("오류: 변환 후 본문이 0자입니다. 변환 결과를 신뢰하지 마세요.")
    print(f"변환 완료: {abs_output} (본문 {n_before}자 → {n_after}자)")
    print("권장: hwpx_edit.py의 --to-md 자가검증(recall)으로 변환 품질을 확인하세요.")


def doc_stats(hwp):
    """현재 문서의 쪽수·표 수·그리기 개체(그림 포함) 수·본문."""
    tables = shapes = 0
    ctrl = hwp.hwp.HeadCtrl
    while ctrl:
        if ctrl.CtrlID == "tbl":
            tables += 1
        elif ctrl.CtrlID == "gso":
            shapes += 1
        ctrl = ctrl.Next
    return {"pages": hwp.PageCount, "tables": tables, "shapes": shapes,
            "text": extract_full_text(hwp)}


def _squash(text):
    return re.sub(r"\s+", "", text)


def check_merge(parts, merged):
    """합본 통계가 원본 합과 같은지. 어긋난 항목 설명 목록(빈 목록이면 통과)."""
    problems = []
    for key, label in (("pages", "쪽수"), ("tables", "표"), ("shapes", "그림·그리기 개체")):
        want = sum(p[key] for p in parts)
        if merged[key] != want:
            problems.append(f"{label} {merged[key]} (원본 합 {want})")
    if _squash(merged["text"]) != "".join(_squash(p["text"]) for p in parts):
        problems.append("본문 글자(공백 제외)가 원본을 차례로 이은 것과 다름")
    return problems


def cmd_append(filepath, others, output=None, password=None):
    """첫 문서 끝에 나머지 문서를 차례로 이어 붙여 HWPX 한 파일로 저장 (이슈 #14).

    InsertFile(KeepSection=1)이 문서마다 구역을 새로 열어 쪽을 넘기므로
    BreakPage를 넣지 않는다(넣으면 빈 쪽이 낀다). 저장 뒤 합본을 다시 열어
    쪽수·표·그리기 개체 수가 원본 합과 같고 본문이 원본을 차례로 이은 것과
    같은지 대조하고, 어긋나면 `<출력>.rejected.hwpx`로 옮기고 exit 1.
    """
    files = [os.path.abspath(f) for f in [filepath, *others]]
    for f in files:
        if not os.path.exists(f):
            raise SystemExit(f"오류: 파일이 없습니다: {f}")
    defaulted = output is None
    if defaulted:
        stem = os.path.splitext(os.path.basename(filepath))[0]
        output = os.path.join(os.path.dirname(files[0]), BYPRODUCT_DIR, f"{stem}_합본.hwpx")
    abs_output = os.path.abspath(output)
    if not abs_output.lower().endswith(".hwpx"):
        raise SystemExit("오류: 출력 파일은 .hwpx 확장자여야 합니다.")
    if abs_output in files:
        raise SystemExit("오류: 입력과 같은 경로로는 저장할 수 없습니다 (in-place 금지).")
    if defaulted and os.path.exists(abs_output):
        raise SystemExit(f"오류: 기본 출력 경로에 파일이 이미 있습니다: {abs_output}\n"
                         "  덮어쓰려면 -o로 경로를 명시하세요.")
    os.makedirs(os.path.dirname(abs_output), exist_ok=True)

    hwp = create_hwp(visible=False)
    try:
        parts = []
        for f in files:
            open_for_com(hwp, f, password)
            stats = doc_stats(hwp)
            if not stats["text"].strip():
                raise SystemExit(f"오류: COM이 본문을 0자로 읽었습니다: {f}\n"
                                 "  Pandoc 생성물은 합치지 않습니다 (warnings-com.md 4번).")
            parts.append(stats)
        open_for_com(hwp, files[0], password)
        for f in files[1:]:
            # MovePos(3)=파일 끝. MoveDocEnd(HAction.Run)는 한컴 판에 따라 무시된다(warnings-com.md 16번)
            hwp.hwp.MovePos(3, 0, 0)
            if not hwp.insert_file(f, keep_section=1, keep_charshape=1,
                                   keep_parashape=1, keep_style=1):
                raise SystemExit(f"오류: 한컴 COM이 이어 붙이지 못했습니다: {f}")
        if not hwp.save_as(abs_output, format="HWPX"):
            raise SystemExit(f"오류: HWPX 저장 실패: {abs_output}")
        open_for_com(hwp, abs_output)
        merged = doc_stats(hwp)
    finally:
        hwp.clear(1)
        hwp.quit()

    for i, (f, p) in enumerate(zip(files, parts), 1):
        print(f"  {i}. {os.path.basename(f)}: {p['pages']}쪽, 표 {p['tables']}, 그림·그리기 개체 {p['shapes']}")
    problems = check_merge(parts, merged)
    if problems:
        rejected = os.path.splitext(abs_output)[0] + ".rejected.hwpx"
        os.replace(abs_output, rejected)
        print("오류: 합본 대조 실패: " + "; ".join(problems), file=sys.stderr)
        print(f"  합본은 {rejected}에 남겼습니다. 결과물로 쓰지 마세요.", file=sys.stderr)
        sys.exit(1)
    print(f"합치기 완료: {abs_output} ({len(files)}개 문서, {merged['pages']}쪽, "
          f"표 {merged['tables']}, 그림·그리기 개체 {merged['shapes']}, 원본 합과 일치)")
    print("engine=com")


# ---------------------------------------------------------------------------
# 종이 인쇄 (이슈 #16)
# ---------------------------------------------------------------------------

DUPLEX_MODES = {"long": "TwoSidedLongEdge", "short": "TwoSidedShortEdge", "none": "OneSided"}
QUEUE_WATCH_SEC = 300        # 대기열이 빌 때까지 기다리는 최대 시간
QUEUE_GRACE_SEC = 5          # 인쇄 명령이 끝난 뒤 작업이 안 보이면 이만큼 더 기다린다

PRINTER_INFO_PS = r"""
$all = @(Get-CimInstance Win32_Printer)
$p = if ($env:HWPX_PRINTER) { $all | Where-Object Name -eq $env:HWPX_PRINTER } else { $all | Where-Object Default }
if (-not $p) { "{}"; exit 0 }
$c = Get-PrintConfiguration -PrinterName $p.Name
[pscustomobject]@{ Name = $p.Name; Default = [bool]$p.Default; PortName = $p.PortName;
  DuplexingMode = [string]$c.DuplexingMode; Collate = [bool]$c.Collate;
  DefaultPrinter = [string]($all | Where-Object Default | Select-Object -First 1 -ExpandProperty Name)
} | ConvertTo-Json -Compress
"""

# Windows "기본 프린터 자동 관리"는 마지막으로 쓴 프린터를 기본으로 바꾼다(2026-11-15 실측)
DEFAULT_PRINTER_PS = r"""
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Get-CimInstance Win32_Printer | Where-Object Default | Select-Object -First 1 -ExpandProperty Name
"""
SET_DEFAULT_PRINTER_PS = r"""
$p = Get-CimInstance Win32_Printer | Where-Object Name -eq $env:HWPX_PRINTER
$r = Invoke-CimMethod -InputObject $p -MethodName SetDefaultPrinter
if ($r.ReturnValue -ne 0) { throw "SetDefaultPrinter=$($r.ReturnValue)" }
"""

SET_DUPLEX_PS = ("Set-PrintConfiguration -PrinterName $env:HWPX_PRINTER "
                 "-DuplexingMode $env:HWPX_DUPLEX -ErrorAction Stop")

QUEUE_WATCH_PS = r"""
$p = $env:HWPX_PRINTER
$pre = @(Get-PrintJob -PrinterName $p | ForEach-Object Id)
"READY"
$deadline = (Get-Date).AddSeconds([int]$env:HWPX_WATCH_SEC)
$seen = 0; $doneAt = $null
while ((Get-Date) -lt $deadline) {
  $jobs = @(Get-PrintJob -PrinterName $p | Where-Object { $pre -notcontains $_.Id })
  foreach ($j in $jobs) { "JOB {0} {1} {2} {3}" -f $j.Id, $j.TotalPages, $j.PagesPrinted, $j.JobStatus; $seen++ }
  if ($seen -gt 0 -and $jobs.Count -eq 0) { "EMPTY"; exit 0 }
  if (Test-Path -LiteralPath $env:HWPX_WATCH_DONE) {
    if (-not $doneAt) { $doneAt = Get-Date }
    if ($seen -eq 0 -and ((Get-Date) - $doneAt).TotalSeconds -gt [int]$env:HWPX_WATCH_GRACE) { "NONE"; exit 0 }
  }
  Start-Sleep -Milliseconds 200
}
"TIMEOUT"
"""


def _ps_env(**values):
    env = os.environ.copy()
    env.update({f"HWPX_{k.upper()}": str(v) for k, v in values.items()})
    return env


def run_ps(script, **values):
    """PowerShell 한 번 실행. 값은 환경변수로 넘겨 따옴표 이스케이프를 피한다."""
    import subprocess
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                       capture_output=True, text=True, encoding="utf-8", errors="replace",
                       env=_ps_env(**values))
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout).strip())
    return r.stdout


def printer_info(name=None):
    """프린터 이름·기본 여부·포트·양면·한 부씩. name이 없으면 기본 프린터."""
    import json
    out = run_ps("[Console]::OutputEncoding = [Text.Encoding]::UTF8\n" + PRINTER_INFO_PS,
                 printer=name or "").strip()
    info = json.loads(out or "{}")
    if not info.get("Name"):
        raise SystemExit(f"오류: 프린터를 찾지 못했습니다: {name or '(기본 프린터 없음)'}")
    return info


def set_duplex(name, mode):
    run_ps(SET_DUPLEX_PS, printer=name, duplex=mode)


def restore_default_printer(original):
    """인쇄 뒤 Windows가 기본 프린터를 바꿨으면 되돌린다. 되돌린 이름 또는 None."""
    if not original or run_ps(DEFAULT_PRINTER_PS).strip() == original:
        return None
    run_ps(SET_DEFAULT_PRINTER_PS, printer=original)
    return original


def start_queue_watch(name, done_flag):
    """새 인쇄 작업을 지켜보는 PowerShell을 띄우고 기준 대기열을 잡을 때까지 기다린다."""
    import subprocess
    proc = subprocess.Popen(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", QUEUE_WATCH_PS],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8",
        errors="replace",
        env=_ps_env(printer=name, watch_sec=QUEUE_WATCH_SEC, watch_done=done_flag,
                    watch_grace=QUEUE_GRACE_SEC))
    first = proc.stdout.readline().strip()
    if first != "READY":
        proc.kill()
        raise RuntimeError(f"대기열 감시를 시작하지 못했습니다: {first or proc.stderr.read().strip()}")
    return proc


def parse_queue_watch(lines):
    """감시 출력 → (작업별 {"pages": 스풀이 끝난 관측의 최대 TotalPages 또는 None}, 끝 상태).

    스풀 중(JobStatus에 Spooling)인 작업의 TotalPages는 아직 다 차지 않은 값이라
    최종값으로 쓰지 않는다(Microsoft Print to PDF 실측: 스풀 중 6, 최종 12).
    """
    jobs, end = {}, "TIMEOUT"
    for line in lines:
        parts = line.split(maxsplit=4)
        if parts and parts[0] == "JOB" and len(parts) >= 3:
            job = jobs.setdefault(parts[1], {"pages": None})
            status = parts[4] if len(parts) > 4 else ""
            if "Spooling" not in status:
                job["pages"] = max(job["pages"] or 0, int(parts[2] or 0))
        elif parts and parts[0] in ("EMPTY", "NONE", "TIMEOUT"):
            end = parts[0]
    return jobs, end


def queue_verdict(jobs, end, pages, copies):
    """대기열 대조 → (status, 설명). status: passed | mismatch | unverified.

    TotalPages는 한글이 부수만큼 찍어 보내면 쪽수 × 부수(이슈 #16 실측 프린터),
    드라이버가 부수를 맡으면 한 부 쪽수다(Microsoft Print to PDF 실측: 6쪽 2부에
    TotalPages 6, 출력 PDF 12쪽). 둘 다 맞는 값으로 본다.
    """
    expected = pages * copies
    if not jobs:
        return "unverified", ("대기열에서 작업을 보지 못했습니다(너무 빨리 지나갔거나 들어가지 않음). "
                              "프린터 출력으로 확인하세요")
    if any(j["pages"] is None for j in jobs.values()):
        return "unverified", (f"작업 {len(jobs)}건이 스풀 중에 대기열을 떠나 최종 쪽수를 보지 못했습니다"
                              f"(기대 {pages}쪽 × {copies}부). 프린터 출력으로 확인하세요")
    total = sum(j["pages"] for j in jobs.values())
    if copies > 1 and total == pages:
        expected = pages
        how = f"한 부 쪽수({pages})와 같음, 부수({copies})는 프린터 드라이버가 처리"
    else:
        how = f"기대 {expected} = {pages}쪽 × {copies}부"
    tail = {"EMPTY": "대기열 비움 확인",
            "TIMEOUT": f"{QUEUE_WATCH_SEC}초 안에 대기열이 비지 않음(인쇄 중일 수 있음)"}.get(end, "")
    desc = f"작업 {len(jobs)}건, TotalPages {total} ({how}). {tail}".rstrip(". ")
    return ("passed" if total == expected else "mismatch"), desc


def sheet_count(pages, copies, duplex_mode):
    per_copy = -(-pages // 2) if duplex_mode.startswith("TwoSided") else pages
    return per_copy * copies


def cmd_print(filepath, printer=None, copies=1, duplex=None, dry=False,
              to_file=None, password=None):
    """HWP/HWPX를 프린터로 인쇄 (한컴 COM HPrint).

    양면은 HPrint에 항목이 없어 프린터 드라이버 기본값(DuplexingMode)을 따른다.
    --duplex가 현재 값과 다르면 인쇄 전에 바꾸고 끝나면 되돌린다. 실행 전에
    프린터·부수·양면·쪽수·장수를 한 줄로 찍고, 인쇄 뒤에는 대기열의 TotalPages를
    쪽수 × 부수와 대조한다. --dry는 인쇄와 설정 변경 없이 계획 줄만 찍는다.
    """
    import shutil
    import tempfile
    if copies < 1:
        raise SystemExit("오류: --copies는 1 이상이어야 합니다.")
    src = os.path.abspath(filepath)
    if not os.path.exists(src):
        raise SystemExit(f"오류: 파일이 없습니다: {src}")
    info = printer_info(printer)
    name = info["Name"]
    if info.get("PortName") == "PORTPROMPT:" and not (dry or to_file):
        raise SystemExit(f"오류: '{name}'은 저장 위치를 묻는 창을 띄우는 프린터라 자동 인쇄가 멈춥니다. "
                         "--print-to-file로 출력 파일을 주세요.")
    current = info["DuplexingMode"]
    target = DUPLEX_MODES[duplex] if duplex else current
    change = target != current

    # 한글·공백·클라우드 경로를 COM에 직접 넘기지 않는다(이슈 #16 실측)
    work = tempfile.mkdtemp(prefix="hwpx_print_")
    staged = os.path.join(work, "print_input" + os.path.splitext(src)[1].lower())
    shutil.copyfile(src, staged)
    restored = default_restored = None
    printed = False
    hwp = watcher = None
    try:
        if change and not dry:
            try:
                set_duplex(name, target)
            except RuntimeError as exc:
                raise SystemExit(f"오류: 양면 설정을 바꾸지 못해 인쇄하지 않았습니다: {exc}")
        hwp = create_hwp(visible=False)
        open_for_com(hwp, staged, password)
        pages = hwp.PageCount
        sheets = sheet_count(pages, copies, target)
        duplex_desc = target + (f"(지금 {current}, 인쇄 뒤 되돌림)" if change else "(프린터 설정 그대로)")
        print(f"인쇄 계획: 프린터={name}, 부수={copies}, 양면={duplex_desc}, "
              f"쪽수={pages}, 용지={sheets}장" + (f", 파일로={to_file}" if to_file else ""))
        if dry:
            print("--dry: 인쇄하지 않았고 프린터 설정도 바꾸지 않았습니다.")
            print("engine=com status=dry")
            return
        done_flag = os.path.join(work, "done")
        watcher = start_queue_watch(name, done_flag)
        h = hwp.hwp
        pset = h.HParameterSet.HPrint
        h.HAction.GetDefault("Print", pset.HSet)
        pset.PrinterName = name
        pset.Range = 0
        pset.NumCopy = copies
        pset.Collate = 1
        pset.PrintMethod = 0
        pset.Device = 0
        if to_file:
            pset.PrintToFile = 1
            pset.filename = os.path.abspath(to_file)
        printed = True
        ok = h.HAction.Execute("Print", pset.HSet)
        open(done_flag, "w").close()
        out, _ = watcher.communicate(timeout=QUEUE_WATCH_SEC + 30)
        if not ok:
            raise SystemExit("오류: 한컴 COM 인쇄 명령이 실패했습니다(Execute=False).")
        jobs, end = parse_queue_watch(out.splitlines())
        status, desc = queue_verdict(jobs, end, pages, copies)
    finally:
        if watcher is not None and watcher.poll() is None:
            watcher.kill()
        if hwp is not None:
            hwp.clear(1)
            hwp.quit()
        if change and not dry:
            try:
                set_duplex(name, current)
                restored = current
            except RuntimeError as exc:
                print(f"경고: 양면 설정을 되돌리지 못했습니다({exc}). 직접 되돌리세요: "
                      f"Set-PrintConfiguration -PrinterName '{name}' -DuplexingMode {current}",
                      file=sys.stderr)
        if printed:
            try:
                default_restored = restore_default_printer(info.get("DefaultPrinter"))
            except RuntimeError as exc:
                print(f"경고: Windows가 바꾼 기본 프린터를 되돌리지 못했습니다({exc}). "
                      f"원래 기본 프린터: {info.get('DefaultPrinter')}", file=sys.stderr)
        shutil.rmtree(work, ignore_errors=True)
    if restored:
        print(f"양면 설정 복원: {restored}")
    if default_restored:
        print(f"기본 프린터 복원: {default_restored} (Windows가 인쇄 뒤 바꾼 것을 되돌림)")
    print(f"대기열 대조: {desc}")
    print(f"engine=com status={status}")
    if status == "mismatch":
        sys.exit(1)


def cmd_to_pdf(filepath, output=None, password=None):
    """HWP/HWPX → PDF (한컴 COM SaveAs)."""
    if output is None:
        output = os.path.splitext(filepath)[0] + ".pdf"
    abs_output = os.path.abspath(output)
    os.makedirs(os.path.dirname(abs_output) or ".", exist_ok=True)

    hwp = create_hwp(visible=False)
    try:
        open_for_com(hwp, filepath, password)
        from pdf_export import save_pdf
        save_pdf(hwp.hwp, abs_output)
    finally:
        hwp.quit()

    if not os.path.exists(abs_output) or os.path.getsize(abs_output) == 0:
        raise SystemExit(f"오류: PDF 출력 파일이 생성되지 않았습니다: {abs_output}")
    print(f"PDF 저장: {abs_output}")


def main():
    parser = argparse.ArgumentParser(
        description="한컴 COM 네이티브 HWPX 파이프라인 (pyhwpx 기반)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  %(prog)s --diagnose                               # pyhwpx + COM 진단
  %(prog)s output.hwpx --from-md input.md           # MD → COM 네이티브 HWPX 생성
  %(prog)s doc.hwpx --insert-image stamp.png        # 문서 끝에 이미지 삽입 (doc_img.hwpx)
  %(prog)s doc.hwpx --insert-image s.png --image-width 40 --image-height 20
  %(prog)s doc.hwpx --get-text                      # COM 기준 본문 추출 (호환성 점검)
  %(prog)s doc.hwpx --normalize -o final.hwpx       # 한컴 재저장 정규화 (인쇄 직전 마무리)
  %(prog)s doc.hwp --from-hwp                       # HWP → HWPX 변환 (JDK 없는 기기 폴백)
  %(prog)s doc.hwpx --to-pdf                        # PDF 저장
  %(prog)s a.hwpx --append b.hwp c.hwpx -o 합본.hwpx  # 여러 문서를 한 파일로 (쪽수·개체 대조)
  %(prog)s doc.hwpx --print --copies 30 --duplex long --dry   # 인쇄 계획만 (먼저 사용자에게 보고)
  %(prog)s doc.hwpx --print --copies 30 --duplex long         # 인쇄 + 대기열 쪽수 대조
주의: 입력 파일은 절대 in-place로 덮어쓰지 않는다. Pandoc 생성 HWPX는
      COM이 본문을 0자로 읽으므로(--get-text로 확인) 이 파이프라인에 넣지 말 것.
        """,
    )
    parser.add_argument("file", nargs="?", help="대상 파일 (--from-md에서는 출력 HWPX 경로)")
    parser.add_argument("--diagnose", action="store_true", help="pyhwpx + 한컴 COM 진단")
    parser.add_argument("--from-md", metavar="MD", help="Markdown 부분집합으로 새 HWPX 생성")
    parser.add_argument("--insert-image", metavar="IMG", help="문서 끝에 이미지 삽입 (별도 파일 저장)")
    parser.add_argument("--image-width", type=int, default=0, help="이미지 너비(mm), --insert-image와 함께")
    parser.add_argument("--image-height", type=int, default=0, help="이미지 높이(mm), --insert-image와 함께")
    parser.add_argument("--get-text", action="store_true", help="COM 기준 본문 텍스트 출력")
    parser.add_argument("--normalize", action="store_true",
                        help="COM으로 열어 별도 HWPX로 재저장 (lineseg·미리보기 한컴 재계산)")
    parser.add_argument("--from-hwp", action="store_true",
                        help="HWP → HWPX 변환 (hwp2hwpx.bat의 JDK 없는 기기용 폴백)")
    parser.add_argument("--to-pdf", action="store_true", help="PDF로 저장")
    parser.add_argument("--append", nargs="+", metavar="FILE",
                        help="대상 문서 끝에 FILE들을 차례로 이어 붙여 HWPX 한 파일로 저장 "
                             "(기본 출력: _work-hwpx-automation/<대상>_합본.hwpx)")
    parser.add_argument("--print", dest="do_print", action="store_true",
                        help="프린터로 인쇄 (실행 전 계획 한 줄, 뒤에 대기열 쪽수 대조)")
    parser.add_argument("--printer", help="--print 프린터 이름 (생략하면 Windows 기본 프린터)")
    parser.add_argument("--copies", type=int, default=1, help="--print 부수 (기본 1)")
    parser.add_argument("--duplex", choices=sorted(DUPLEX_MODES),
                        help="--print 양면: long(긴 쪽 넘김)·short(짧은 쪽)·none(단면). "
                             "생략하면 프린터 설정 그대로. 다르면 바꾸고 인쇄 뒤 되돌린다")
    parser.add_argument("--dry", action="store_true",
                        help="--print 계획(프린터·부수·양면·쪽수·용지 장수)만 찍고 인쇄·설정 변경 안 함")
    parser.add_argument("--print-to-file", metavar="PATH",
                        help="--print 결과를 종이 대신 파일로 (Microsoft Print to PDF 등 확인용)")
    parser.add_argument("--password", help="암호화된 문서의 열기 암호")
    parser.add_argument("-o", "--output", help="출력 파일 경로")

    args = parser.parse_args()

    if args.diagnose:
        sys.exit(cmd_diagnose())

    if args.file is None:
        parser.error("대상 파일 경로가 필요합니다 (--diagnose 제외).")

    if args.from_md:
        cmd_from_md(args.file, args.from_md)
    elif args.insert_image:
        cmd_insert_image(args.file, args.insert_image, args.output,
                         args.image_width, args.image_height, args.password)
    elif args.normalize:
        cmd_normalize(args.file, args.output, args.password)
    elif args.from_hwp:
        cmd_from_hwp(args.file, args.output, args.password)
    elif args.get_text:
        cmd_get_text(args.file, args.password)
    elif args.to_pdf:
        cmd_to_pdf(args.file, args.output, args.password)
    elif args.append:
        cmd_append(args.file, args.append, args.output, args.password)
    elif args.do_print:
        cmd_print(args.file, args.printer, args.copies, args.duplex, args.dry,
                  args.print_to_file, args.password)
    else:
        parser.error("동작 플래그가 필요합니다: --from-md / --insert-image / --from-hwp / --get-text / --to-pdf / --append / --print / --diagnose")


if __name__ == "__main__":
    main()
