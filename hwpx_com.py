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


def cmd_to_pdf(filepath, output=None, password=None):
    """HWP/HWPX → PDF (한컴 COM SaveAs)."""
    if output is None:
        output = os.path.splitext(filepath)[0] + ".pdf"
    abs_output = os.path.abspath(output)
    os.makedirs(os.path.dirname(abs_output) or ".", exist_ok=True)

    hwp = create_hwp(visible=False)
    try:
        open_for_com(hwp, filepath, password)
        hwp.save_as(abs_output, format="PDF")
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
  %(prog)s doc.hwpx --to-pdf                        # PDF 저장
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
    parser.add_argument("--to-pdf", action="store_true", help="PDF로 저장")
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
    elif args.get_text:
        cmd_get_text(args.file, args.password)
    elif args.to_pdf:
        cmd_to_pdf(args.file, args.output, args.password)
    else:
        parser.error("동작 플래그가 필요합니다: --from-md / --insert-image / --get-text / --to-pdf / --diagnose")


if __name__ == "__main__":
    main()
