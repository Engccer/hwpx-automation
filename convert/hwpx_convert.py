#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hwpx_convert.py - Unified Document Converter
다양한 문서 형식을 HWPX 또는 HTML로 변환하는 통합 도구입니다.

Supported input formats: MD, DOCX, HTML, RST, TEX, TXT (Pandoc 지원 형식)
Output formats: HWPX, HTML

Usage:
    # 단일 파일 변환
    ./hwpx_convert.py document.md -o output.hwpx
    ./hwpx_convert.py document.md -o output.html --format html

    # 템플릿 사용
    ./hwpx_convert.py document.docx -o output.hwpx --reference-doc template.hwpx

    # 일괄 변환
    ./hwpx_convert.py ./docs/ -o ./output/ --batch
    ./hwpx_convert.py ./docs/ -o ./output/ --batch --recursive --verbose
"""

import argparse
import glob
import os
import sys

from pypandoc_hwpx.PandocToHwpx import PandocToHwpx
from pypandoc_hwpx.PandocToHtml import PandocToHtml


# 지원하는 입력 확장자
SUPPORTED_EXTENSIONS = {
    '.md', '.markdown',  # Markdown
    '.docx',             # Word
    '.html', '.htm',     # HTML
    '.rst',              # reStructuredText
    '.tex',              # LaTeX
    '.txt',              # Plain text
}


def get_default_reference():
    """기본 blank.hwpx 템플릿 경로 반환"""
    import pypandoc_hwpx
    pkg_dir = os.path.dirname(os.path.abspath(pypandoc_hwpx.__file__))
    return os.path.join(pkg_dir, "blank.hwpx")


def find_input_files(input_dir, recursive=False):
    """디렉토리에서 변환 가능한 파일 찾기"""
    files = []
    for ext in SUPPORTED_EXTENSIONS:
        if recursive:
            pattern = os.path.join(input_dir, '**', f'*{ext}')
            files.extend(glob.glob(pattern, recursive=True))
        else:
            pattern = os.path.join(input_dir, f'*{ext}')
            files.extend(glob.glob(pattern))
    return sorted(set(files))


def convert_file(input_path, output_path, ref_doc, output_format, verbose=False):
    """단일 파일 변환"""
    if verbose:
        print(f"  변환: {input_path} -> {output_path}")

    try:
        if output_format == 'hwpx':
            PandocToHwpx.convert_to_hwpx(input_path, output_path, ref_doc)
        else:  # html
            PandocToHtml.convert_to_html(input_path, output_path)
        return True
    except Exception as e:
        print(f"  오류 ({input_path}): {e}", file=sys.stderr)
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Unified Document to HWPX/HTML Converter (문서 통합 변환 도구)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # 단일 파일 변환
  %(prog)s document.md -o output.hwpx
  %(prog)s document.docx -o output.hwpx --reference-doc template.hwpx
  %(prog)s document.md -o output.html --format html

  # 일괄 변환
  %(prog)s ./docs/ -o ./output/ --batch
  %(prog)s ./docs/ -o ./output/ --batch --recursive --verbose

  # 테스트 (실제 변환 없이)
  %(prog)s ./docs/ -o ./output/ --batch --dry-run

Supported input formats: MD, DOCX, HTML, RST, TEX, TXT
        """
    )

    parser.add_argument("input", nargs='+', help="입력 파일 또는 디렉토리")
    parser.add_argument("-o", "--output", required=True, help="출력 파일 또는 디렉토리")
    parser.add_argument("--reference-doc", help="스타일 참조용 HWPX 템플릿")
    parser.add_argument("--format", choices=['hwpx', 'html'], default='hwpx',
                        help="출력 형식 (기본: hwpx)")
    parser.add_argument("--batch", action="store_true",
                        help="일괄 변환 모드: 디렉토리 내 모든 파일 변환")
    parser.add_argument("--recursive", action="store_true",
                        help="하위 디렉토리 포함 (--batch와 함께 사용)")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="상세 출력")
    parser.add_argument("--dry-run", action="store_true",
                        help="변환 없이 미리보기만 표시")

    args = parser.parse_args()

    # 참조 문서 설정
    ref_doc = args.reference_doc or get_default_reference()
    if args.format == 'hwpx' and not os.path.exists(ref_doc):
        print(f"오류: 참조 템플릿을 찾을 수 없습니다: {ref_doc}", file=sys.stderr)
        sys.exit(1)

    # 출력 확장자 결정
    out_ext = '.hwpx' if args.format == 'hwpx' else '.html'

    # 변환할 파일 목록 수집
    files_to_convert = []

    if args.batch:
        # 일괄 변환 모드
        input_dir = args.input[0]
        if not os.path.isdir(input_dir):
            print(f"오류: '{input_dir}'은(는) 디렉토리가 아닙니다", file=sys.stderr)
            sys.exit(1)

        input_files = find_input_files(input_dir, args.recursive)
        if not input_files:
            print(f"경고: 변환할 파일이 없습니다 ({input_dir})")
            sys.exit(0)

        output_dir = args.output
        if not args.dry_run:
            os.makedirs(output_dir, exist_ok=True)

        for input_file in input_files:
            # 상대 경로 구조 유지
            rel_path = os.path.relpath(input_file, input_dir)
            base_name = os.path.splitext(rel_path)[0]
            output_file = os.path.join(output_dir, base_name + out_ext)
            files_to_convert.append((input_file, output_file))

    else:
        # 단일/다중 파일 모드
        for input_file in args.input:
            if not os.path.exists(input_file):
                print(f"경고: 파일을 찾을 수 없습니다: {input_file}", file=sys.stderr)
                continue

            if os.path.isdir(input_file):
                print(f"경고: '{input_file}'은(는) 디렉토리입니다. --batch 옵션을 사용하세요.",
                      file=sys.stderr)
                continue

            # 단일 파일이면서 출력이 디렉토리가 아닌 경우
            if len(args.input) == 1 and not args.output.endswith(os.sep):
                output_file = args.output
                if not output_file.endswith(out_ext):
                    output_file = os.path.splitext(output_file)[0] + out_ext
            else:
                # 다중 파일이거나 출력이 디렉토리인 경우
                output_dir = args.output
                if not args.dry_run:
                    os.makedirs(output_dir, exist_ok=True)
                base_name = os.path.splitext(os.path.basename(input_file))[0]
                output_file = os.path.join(output_dir, base_name + out_ext)

            files_to_convert.append((input_file, output_file))

    if not files_to_convert:
        print("변환할 파일이 없습니다.")
        sys.exit(0)

    # 변환 실행
    success_count = 0
    fail_count = 0

    if args.verbose or len(files_to_convert) > 1:
        print(f"\n총 {len(files_to_convert)}개 파일 변환 시작 (형식: {args.format.upper()})\n")

    for input_file, output_file in files_to_convert:
        if args.dry_run:
            print(f"[DRY-RUN] {input_file} -> {output_file}")
            success_count += 1
        else:
            # 출력 디렉토리 생성
            output_dir = os.path.dirname(output_file)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)

            if convert_file(input_file, output_file, ref_doc, args.format, args.verbose):
                success_count += 1
            else:
                fail_count += 1

    # 결과 요약
    if args.verbose or len(files_to_convert) > 1 or args.dry_run:
        print(f"\n변환 완료: 성공 {success_count}개, 실패 {fail_count}개")

    sys.exit(0 if fail_count == 0 else 1)


if __name__ == "__main__":
    main()
