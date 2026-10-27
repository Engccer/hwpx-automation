#!/usr/bin/env python3
"""
hwpx_edit.py - HWPX 파일 편집 유틸리티

사용법:
  python hwpx_edit.py --check-env                            # 기능 계층별 의존성·런타임 점검 (첫 실행 시)
  python hwpx_edit.py <파일.hwp> --from-hwp                   # HWP → HWPX (rhwp, JAR 교차 대조. --engine jar 폴백)
  python hwpx_edit.py <파일.hwpx> --info                       # 표 구조 출력
  python hwpx_edit.py <파일.hwpx> --to-md                      # HWPX → Markdown 변환 (hwpx-tomd 엔진, API 불필요)
  python hwpx_edit.py --diagnose-com                         # 한컴 COM 자동화 진단
  python hwpx_edit.py <파일.hwpx> --to-pdf                    # HWPX/HWP → PDF (rhwp. --pdf-backend com은 Windows 한컴)
  python hwpx_edit.py <파일.hwpx> --find "이전" --replace "이후"  # 텍스트 치환
  python hwpx_edit.py <파일.hwpx> --set-cell 0,1,0 "텍스트"     # 표 셀에 텍스트 입력
  python hwpx_edit.py <파일.hwpx> --split-cell 0,1,0            # 병합 셀 분할

의존성:
  - 편집 명령(--find/--set-cell 등): python-hwpx, lxml
  - 변환(--to-md): hwpx-tomd 패키지 (pip install hwpx-tomd). 변환 엔진은 이 패키지가
    단일 소스로 관리하며, 본 파일은 출력 경로·콘솔 메시지 등 CLI 래퍼만 담당한다.
  - PDF 변환(--to-pdf): rhwp CLI(rhwp_pdf.py). --pdf-backend com은 Windows 한컴오피스 + pywin32
  - HWP 변환(--from-hwp, .hwp 입력): rhwp CLI. JDK 21 + 번들 JAR은 --engine jar 폴백과 교차 대조
"""

import argparse
import contextlib
import copy
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter, namedtuple
from io import BytesIO

import rhwp_cli

# lxml 미설치 기기에서도 --check-env(환경 진단)가 동작해야 하므로 여기서 죽지 않는다.
# 파일 명령의 실제 가드는 main()에서 설치 안내와 함께 수행한다.
try:
    from lxml import etree
except ImportError:
    etree = None

try:
    import winreg
except ImportError:
    winreg = None


# HWPX 네임스페이스
NS = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'ha': 'http://www.hancom.co.kr/hwpml/2011/app',
    'config': 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
}


# 도구가 자동 저장하는 출력 폴더(원본 비파괴용). 결과물/부산물 구분은 도구가 하지
# 않으며, 무엇이 최종 결과물인지는 호출하는 워크플로우(SKILL.md 절차)가 판단해
# 작업 폴더로 승격한다. 폴더명은 출처(hwpx-automation) + 성격(work=중간 작업물)을
# 드러낸다.
BYPRODUCT_DIR = '_work-hwpx-automation'


def get_output_path(filepath, ext=None):
    """입력 파일 기준 BYPRODUCT_DIR 폴더에 출력 경로 생성. ext가 None이면 원래 확장자 유지.
    이미 BYPRODUCT_DIR 안에 있으면 같은 폴더에 출력 (중첩 방지)."""
    input_dir = os.path.dirname(os.path.abspath(filepath))
    if os.path.basename(input_dir) == BYPRODUCT_DIR:
        output_dir = input_dir
    else:
        output_dir = os.path.join(input_dir, BYPRODUCT_DIR)
    os.makedirs(output_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(filepath))[0]
    if ext:
        return os.path.join(output_dir, f"{base}{ext}")
    else:
        return os.path.join(output_dir, os.path.basename(filepath))


def is_encrypted_hwpx(filepath):
    """HWPX 파일이 AES 등으로 암호화되어 있는지 확인.

    META-INF/manifest.xml의 <odf:encryption-data> 존재 여부로 판단.
    암호화된 경우 Contents/section0.xml이 암호문이어서 XML 파싱이 실패한다.
    """
    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            if 'META-INF/manifest.xml' not in zf.namelist():
                return False
            manifest = zf.read('META-INF/manifest.xml')
            return b'encryption-data' in manifest
    except (zipfile.BadZipFile, KeyError):
        return False


ENCRYPTION_HINT = (
    "이 HWPX 파일은 암호화되어 있습니다 (AES-256-CBC).\n"
    "XML 직접 파싱으로 처리할 수 없으므로 한컴 COM으로 암호를 먼저 제거해야 합니다.\n"
    "\n"
    "해결: 한컴오피스 설치된 Windows에서 아래 스크립트로 복호화하세요.\n"
    "  hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')\n"
    "  hwp.Open(abs_in, 'HWPX', 'password:<비밀번호>')\n"
    "  act = hwp.CreateAction('FilePasswordChange')\n"
    "  pset = act.CreateSet(); act.GetDefault(pset)\n"
    "  pset.SetItem('String', ''); pset.SetItem('Ask', 0)\n"
    "  pset.SetItem('ReadString', ''); pset.SetItem('WriteString', '')\n"
    "  pset.SetItem('RWAsk', 0); act.Execute(pset)\n"
    "  hwp.SaveAs(abs_out, 'HWPX', ''); hwp.Quit()\n"
    "상세: reference/encrypted-hwpx.md 참조"
)


# XML 1.0 Char 프로덕션에서 배제된 제어문자(탭 0x09·개행 0x0A·복귀 0x0D 제외).
ILLEGAL_XML_BYTES = re.compile(rb"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def sanitize_xml_parts(all_files):
    """HWPX의 XML 파트에서 XML 1.0 불법 제어문자를 제거한다(제자리 수정).

    hwp2hwpx 변환본은 하이퍼링크 필드의 <hp:stringParam name="Command"> 값에
    HWP 바이너리의 NUL 패딩을 그대로 써 넣는다. NUL 등은 XML 1.0에서
    이스케이프로도 표현할 수 없는 불법 문자라 lxml이 그 지점에서 파싱을
    중단한다("Premature end of data").

    정제는 파싱 직전 바이트가 아니라 **all_files 딕셔너리에** 적용한다. 트리만
    고치면 재직렬화되지 않는 파트(header.xml·section1.xml·--add-preview
    산출물)에 불법 바이트가 남아 hwpx-validate-package가 malformed XML을 낸다.

    BinData는 건드리지 않는다(이미지 바이너리에는 제어 바이트가 정상적으로
    수만 개 들어 있다). 탭·개행·복귀는 XML 1.0 합법이라 보존한다.

    Returns:
        제거한 총 바이트 수(0이면 아무것도 바꾸지 않았다).
    """
    total = 0
    for name in list(all_files):
        if name.startswith('BinData/') or not name.endswith(('.xml', '.hpf', '.rdf')):
            continue
        data = all_files[name]
        if not ILLEGAL_XML_BYTES.search(data):
            continue
        found = ILLEGAL_XML_BYTES.findall(data)
        all_files[name] = ILLEGAL_XML_BYTES.sub(b'', data)
        total += len(found)
        kinds = ', '.join(f"0x{v[0]:02X} {n}개" for v, n in sorted(Counter(found).items()))
        print(f"sanitize: {name}에서 XML 1.0 불법 제어문자 {len(found)}바이트 제거 ({kinds})")
    return total


def open_hwpx(filepath):
    """HWPX 파일을 열어 section0.xml의 etree와 원본 ZIP 데이터를 반환"""
    if is_encrypted_hwpx(filepath):
        print(f"오류: {ENCRYPTION_HINT}", file=sys.stderr)
        sys.exit(1)

    with zipfile.ZipFile(filepath, 'r') as zf:
        # section 파일 찾기
        section_files = [n for n in zf.namelist() if 'section' in n.lower() and n.endswith('.xml')]
        if not section_files:
            print("오류: section XML 파일을 찾을 수 없습니다.", file=sys.stderr)
            sys.exit(1)

        section_path = section_files[0]
        all_files = {}
        for name in zf.namelist():
            all_files[name] = zf.read(name)

    # 불법 제어문자는 여기서 한 번만 걷어낸다. all_files를 고치므로 파싱뿐
    # 아니라 재패키징되는 산출물까지 정상 XML이 된다.
    sanitize_xml_parts(all_files)

    try:
        root = etree.fromstring(all_files[section_path])
    except etree.XMLSyntaxError as exc:
        print(f"오류: {section_path} 파싱 실패: {exc}", file=sys.stderr)
        sys.exit(1)
    return root, all_files, section_path


def sanitize_header(all_files):
    """header.xml의 borderFill 문제를 자동 수정.

    HWP→HWPX 변환 시 문단 테두리(hh:border)가 참조하는 borderFill에
    faceColor="#000000"이 포함되어 문단 배경이 검게 렌더링되는 문제를 수정한다.
    문단 테두리용 borderFill의 fillBrush에서 faceColor가 검은색이면 "none"으로 변경.
    """
    header_path = 'Contents/header.xml'
    if header_path not in all_files:
        return 0

    header_root = etree.fromstring(all_files[header_path])

    # 1) 문단 스타일(paraPr)의 border가 참조하는 borderFillIDRef 수집
    para_border_bf_ids = set()
    for border_el in header_root.findall('.//' + '{http://www.hancom.co.kr/hwpml/2011/head}border'):
        bf_ref = border_el.get('borderFillIDRef')
        if bf_ref:
            para_border_bf_ids.add(bf_ref)

    if not para_border_bf_ids:
        return 0

    # 2) 해당 borderFill의 winBrush faceColor="#000000" → "none"
    fixed = 0
    for bf_el in header_root.findall('.//' + '{http://www.hancom.co.kr/hwpml/2011/head}borderFill'):
        bf_id = bf_el.get('id')
        if bf_id not in para_border_bf_ids:
            continue
        for wb in bf_el.findall('.//' + '{http://www.hancom.co.kr/hwpml/2011/core}winBrush'):
            fc = wb.get('faceColor')
            if fc and fc.lower() in ('#000000', '000000'):
                wb.set('faceColor', 'none')
                fixed += 1

    if fixed > 0:
        all_files[header_path] = serialize_xml(header_root)
        print(f"sanitize: 문단 테두리 배경색 {fixed}건 수정 (faceColor → none)")

    return fixed


def fix_empty_cells(root):
    """표 셀에 <hp:p>가 없으면 기본 빈 문단을 자동 삽입.

    Pandoc 기반 hwpx_convert.py 등으로 MD의 빈 표 셀(` | | `)이 HWPX로 변환될 때
    <hp:subList>만 있고 내부 <hp:p>가 없는 구조가 생성됨. 한글 엔진은 이를 만나면
    약 15초 로딩 후 안전 종료함(XSD 스키마는 통과하므로 hwpx-validate로는 탐지 불가).
    이 함수는 빈 셀의 subList 안에 minimal <hp:p><hp:run/></hp:p>을 삽입해 한글이
    정상 렌더링하도록 보정함. 중첩 표(subList 안 <hp:tbl>)가 있는 셀은 건너뜀.
    """
    HP_NS = NS['hp']
    fixed = 0
    for tc in root.findall('.//hp:tc', NS):
        sub_list = tc.find('hp:subList', NS)
        if sub_list is None:
            continue
        # 셀 내용이 이미 존재하면(p 또는 중첩 tbl) 건너뜀
        if sub_list.find('hp:p', NS) is not None:
            continue
        if sub_list.find('hp:tbl', NS) is not None:
            continue
        # 빈 셀: minimal 문단 삽입
        p = etree.SubElement(sub_list, f'{{{HP_NS}}}p')
        p.set('id', '0')
        p.set('paraPrIDRef', '0')
        p.set('styleIDRef', '0')
        p.set('pageBreak', '0')
        p.set('columnBreak', '0')
        p.set('merged', '0')
        run = etree.SubElement(p, f'{{{HP_NS}}}run')
        run.set('charPrIDRef', '0')
        fixed += 1
    return fixed


def serialize_xml(elem):
    """HWPX 파트용 XML 직렬화.

    standalone="yes" 선언이 필수다. 빠지면 한글은 그대로 열지만
    hwpx-validate-package가 section/header 파트에 대해
    'missing XML declaration with standalone="yes"' ERROR를 낸다
    (XSD 검증인 hwpx-validate와 --to-md는 통과하므로 자동 검증에 안 잡힘).
    """
    return etree.tostring(elem, xml_declaration=True, encoding='UTF-8',
                          standalone=True)


def zip_layout(filepath):
    """원본 ZIP의 항목별 압축 방식 맵을 반환(재패키징 시 복제용). 실패 시 빈 dict."""
    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            return {i.filename: i.compress_type for i in zf.infolist()}
    except Exception:
        return {}


def write_hwpx_zip(output, all_files, layout=None):
    """HWPX 규격에 맞게 ZIP을 재패키징한다.

    - mimetype은 **첫 항목 + ZIP_STORED(무압축)**여야 한다(OPC 계열 규칙).
      일괄 ZIP_DEFLATED로 쓰면 hwpx-validate-package가 'must use ZIP_STORED'
      ERROR를 내며, 편집 전에는 통과하던 파일이 편집만으로 검증에 실패한다.
    - 나머지 항목은 원본이 쓰던 압축 방식을 그대로 복제한다(한컴이 STORED로
      두는 version.xml·Preview/PrvImage.png 등 보존). 원본에 없던 새 항목은
      ZIP_DEFLATED.
    """
    layout = layout or {}
    names = list(all_files)
    if 'mimetype' in all_files:
        names = ['mimetype'] + [n for n in names if n != 'mimetype']

    tmp = output + '.tmp'
    with zipfile.ZipFile(tmp, 'w') as zf:
        for name in names:
            compress = (zipfile.ZIP_STORED if name == 'mimetype'
                        else layout.get(name, zipfile.ZIP_DEFLATED))
            zf.writestr(name, all_files[name], compress_type=compress)
    os.replace(tmp, output)


def save_hwpx(filepath, root, all_files, section_path, output=None):
    """수정된 section XML을 HWPX 파일에 저장"""
    if output is None:
        output = get_output_path(filepath)

    # 저장 전 빈 셀 자동 보정 (Pandoc 변환물 등의 한글 크래시 방지)
    empty_cells = fix_empty_cells(root)
    if empty_cells > 0:
        print(f"fix_empty_cells: 빈 셀 {empty_cells}개에 기본 문단 삽입")

    all_files[section_path] = serialize_xml(root)

    # 저장 전 header.xml 자동 sanitize
    sanitize_header(all_files)

    write_hwpx_zip(output, all_files, zip_layout(filepath))

    print(f"저장: {output}")


def find_tables(root):
    """section XML에서 모든 <hp:tbl> 요소를 찾아 반환"""
    return root.findall('.//hp:tbl', NS)


def get_table_rows(table):
    """테이블에서 행(hp:tr) 목록 반환"""
    return table.findall('hp:tr', NS)


def get_row_cells(row):
    """행에서 셀(hp:tc) 목록 반환"""
    return row.findall('hp:tc', NS)


def localname(tag):
    """네임스페이스를 제거한 태그 로컬명"""
    return tag.split('}')[-1] if '}' in tag else tag


def t_full_text(t_elem):
    """<hp:t>의 전체 텍스트를 추출.

    핵심: t_elem.text만 읽으면 <hp:t> 내부의 <hp:tab>/<hp:lineBreak> 등
    자식 요소의 tail 텍스트를 통째로 잃는다(객관식 선택지 ②③⑤가 tab.tail에
    들어있는 경우 등). itertext로 전체를 모으되 tab/lineBreak는 공백으로
    치환해 단어 경계를 보존한다.
    (2026-06-05 재검증으로 확인·수정된 핵심 결함.)
    """
    parts = [t_elem.text or '']
    for child in t_elem:
        if localname(child.tag) in ('tab', 'lineBreak', 'br'):
            parts.append(' ')
        else:
            parts.append(''.join(child.itertext()))
        parts.append(child.tail or '')
    return ''.join(parts)


def get_cell_text(cell):
    """셀의 텍스트 내용 추출 (모든 <hp:t> 요소 결합, tail 포함)"""
    texts = [t_full_text(t) for t in cell.findall('.//hp:t', NS)]
    return ' '.join(t for t in texts if t)


def get_cell_span(cell):
    """셀의 rowSpan, colSpan 값 추출.

    주의: span은 cellAddr가 아니라 별도 <hp:cellSpan> 요소에 있다. cellAddr에는
    colAddr/rowAddr(논리 위치)만 있다. 과거 구현은 cellAddr.get('rowSpan')을
    읽어 항상 (1,1)을 반환했고, 그 결과 가로·세로 병합이 모두 무시되어 표 정렬이
    깨졌다(2026-06-05 이질 문서 검증으로 발견·수정).
    """
    span = cell.find('hp:cellSpan', NS)
    if span is None:
        return 1, 1
    return int(span.get('rowSpan', '1')), int(span.get('colSpan', '1'))


def cmd_to_md(filepath, output=None, cell_br=False, merge_fill=False):
    """HWPX를 Markdown으로 변환 (hwpx-tomd 패키지 엔진 호출, 외부 API 불필요).

    변환 엔진(글상자 reading-order 수집, <hp:t> tail 보존, 표 cellAddr/cellSpan
    그리드 배치, 자가검증)은 독립 패키지 hwpx-tomd가 단일 소스로 보유한다. 이
    함수는 출력 경로(BYPRODUCT_DIR/.md)와 콘솔 메시지 등 기존 CLI 동작만 감싼다.
    설치: pip install hwpx-tomd
      (로컬 개발: pip install -e path/to/hwpx-tomd)

    cell_br=True이면 표 셀 내부 문단을 <br>로 구분한다(긴 지문이 셀 안에 있는
    고사지·보고서에 권장). merge_fill=True이면 병합으로 덮인 칸을 시작 칸 값으로
    채운다(행 단위 파싱·LLM 입력용; 기본은 GFM 정렬 보존).
    """
    try:
        from hwpx_tomd import convert, HwpxEncryptedError, HwpxError
    except ImportError:
        print(
            "오류: --to-md에는 hwpx-tomd 패키지가 필요합니다.\n"
            "  설치: pip install hwpx-tomd\n"
            "  (로컬 개발: pip install -e path/to/hwpx-tomd)\n"
            "변환 엔진은 hwpx-tomd 패키지가 단일 소스로 관리합니다(코드 분기 방지).",
            file=sys.stderr,
        )
        sys.exit(3)

    try:
        result = convert(filepath, cell_br=cell_br, merge_fill=merge_fill)
    except HwpxEncryptedError as e:
        print(f"오류: {e}", file=sys.stderr)
        sys.exit(1)
    except HwpxError as e:
        print(f"오류: {e}", file=sys.stderr)
        sys.exit(1)

    if output is None:
        output = get_output_path(filepath, '.md')
    with open(output, 'w', encoding='utf-8') as f:
        f.write(result.markdown)

    print(f"변환 완료: {output}")
    print(f"크기: {len(result.markdown)} 글자")
    for w in result.warnings:
        print(f"경고: {w}", file=sys.stderr)
    if not result.warnings:
        print(f"자가검증 recall: 단어 {result.recall:.1%} · 글자 {result.char_recall:.1%}")


def cmd_info(filepath):
    """표 구조 정보 출력"""
    root, _, _ = open_hwpx(filepath)
    tables = find_tables(root)

    if not tables:
        print("표를 찾을 수 없습니다.")
        return

    print(f"총 {len(tables)}개의 표 발견\n")

    for t_idx, table in enumerate(tables):
        rows = get_table_rows(table)
        print(f"=== 표 {t_idx} ===")
        print(f"  행 수: {len(rows)}")

        for r_idx, row in enumerate(rows):
            cells = get_row_cells(row)
            cell_infos = []
            for c_idx, cell in enumerate(cells):
                text = get_cell_text(cell)
                row_span, col_span = get_cell_span(cell)
                preview = text[:20] + '...' if len(text) > 20 else text
                span_info = ''
                if row_span > 1 or col_span > 1:
                    span_info = f' [병합:{row_span}x{col_span}]'
                cell_infos.append(f'[{c_idx}] "{preview}"{span_info}')
            print(f"  행 {r_idx} ({len(cells)}셀): {' | '.join(cell_infos)}")
        print()


def sanitize_illegal_xml_file(filepath):
    """XML 파트에 불법 제어문자가 있으면 정제한 임시 사본 경로를 반환한다.

    python-hwpx(HwpxDocument.open)는 자체 파서로 section XML을 읽으므로
    open_hwpx의 정제를 거치지 않는다. 열기 전에 정제본을 만들어 그 파일을
    넘긴다. 정제할 것이 없으면 None을 반환한다. 임시본은 시스템 temp에
    만든다(원본 폴더에 두면 동기화 폴더 오염·읽기 전용 폴더 실패·중단 시
    문서 사본 잔존 문제가 생긴다). 호출자는 반환된 파일을 반드시 삭제한다.
    """
    with zipfile.ZipFile(filepath, 'r') as zf:
        data = {n: zf.read(n) for n in zf.namelist()}

    if sanitize_xml_parts(data) == 0:
        return None

    tmp_dir = tempfile.mkdtemp(prefix='hwpx_edit_')
    tmp_path = os.path.join(tmp_dir, os.path.basename(filepath))
    write_hwpx_zip(tmp_path, data, zip_layout(filepath))
    return tmp_path


@contextlib.contextmanager
def sanitized_source(filepath):
    """불법 제어문자가 있으면 정제 임시본 경로를, 없으면 원본 경로를 내준다.

    임시본은 블록을 벗어날 때(오류·조기 종료 포함) 반드시 삭제된다.
    """
    tmp = sanitize_illegal_xml_file(filepath)
    try:
        yield tmp or filepath
    finally:
        if tmp:
            shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)


def cmd_find_replace(filepath, find_text, replace_text, output=None):
    """텍스트 치환: python-hwpx(본문 런, 서식 보존) + lxml(표 셀 등 누락분 보완)"""
    from hwpx.document import HwpxDocument

    # 암호화 문서는 여기서 막는다. 2차 패스가 부르는 open_hwpx도 같은 검사를 하지만,
    # 그때는 python-hwpx가 다시 쓴 중간 파일을 보게 되고 그 파일의 manifest에는
    # encryption-data가 남지 않아 검사를 통과해 버린다(→ 암호문을 파싱하다 XMLSyntaxError).
    if is_encrypted_hwpx(filepath):
        print(f"오류: {ENCRYPTION_HINT}", file=sys.stderr)
        sys.exit(1)

    # 1차: python-hwpx (본문 런 대상, 서식 보존, 런 분할 텍스트도 처리)
    #
    # API 이력: 6.0에서 replace_text_in_runs → doc.text.replace로 이동(7.0 제거 예정),
    # save → save_to_path로 개명(6.x에 save 자체가 없음). 구 API를 그대로 쓰면
    # 6.x에서 AttributeError로 --find/--replace가 전면 실패한다. 신 API 우선 +
    # 구 API 폴백으로 5.x~7.x를 모두 지원한다.
    # python-hwpx는 자체 파서를 쓰므로 open_hwpx의 정제를 거치지 않는다. 불법
    # 제어문자가 있으면 정제본을 만들어 그 파일을 넘긴다(정상 파일은 무비용).
    with sanitized_source(filepath) as src:
        try:
            doc = HwpxDocument.open(src)
        except Exception as e:
            # hwp2hwpx(JAR) 변환물은 container.xml이 Preview/PrvText.txt를 선언하면서도
            # 그 항목을 만들지 않아 python-hwpx가 열기를 거부한다(한글·--to-md는 정상).
            if 'Preview/PrvText.txt' in str(e):
                # --add-preview는 in-place가 아니라 byproduct 경로에 쓴다. 같은 경로로
                # 안내하면 사용자가 명령을 그대로 따라도 원본은 그대로여서 무한 반복에
                # 걸린다. 실제 산출 경로와 그 파일로 재실행하라는 것까지 알려준다.
                fixed = get_output_path(filepath)
                print(f"오류: {e}\n"
                      f"  hwp2hwpx 변환물의 알려진 결함입니다. 두 단계로 보정하세요:\n"
                      f"    python hwpx_edit.py \"{filepath}\" --add-preview\n"
                      f"    python hwpx_edit.py \"{fixed}\" --find ... --replace ...\n"
                      f"  (--add-preview는 원본을 고치지 않고 위 경로에 보정본을 만듭니다)",
                      file=sys.stderr)
                sys.exit(1)
            raise

        # isinstance(str) 배제: 구버전에서 doc.text가 본문 문자열이면 str.replace가
        # 잡혀 문서를 고치지 않고 문자열만 반환하는 조용한 실패가 된다.
        text_ns = getattr(doc, 'text', None)
        if not isinstance(text_ns, str) and hasattr(text_ns, 'replace'):
            count_runs = text_ns.replace(find_text, replace_text)
        else:
            count_runs = doc.replace_text_in_runs(find_text, replace_text)
        save_path = output if output else get_output_path(filepath)
        save_doc = getattr(doc, 'save_to_path', None) or getattr(doc, 'save')
        save_doc(save_path)

    # 2차: lxml으로 1차에서 누락된 <hp:t> 요소만 추가 치환 (표 셀 등)
    # 1차가 이미 치환한 텍스트에는 find_text가 남아있지 않으므로,
    # 여기서 매칭되는 것은 python-hwpx가 접근하지 못한 요소(표 셀 등)뿐임.
    # 단, python-hwpx가 런 구조를 변경하면서 원본 텍스트가 잔존할 수 있으므로
    # replace_text 내에 find_text가 포함되는 경우의 무한 치환도 방지.
    root, all_files, section_path = open_hwpx(save_path)
    count_xml = 0

    # 표 내부 <hp:t>만 대상으로 제한하여 본문 런과의 중복 치환 방지
    tables = root.findall('.//hp:tbl', NS)
    for tbl in tables:
        for t_elem in tbl.findall('.//hp:t', NS):
            if t_elem.text and find_text in t_elem.text:
                t_elem.text = t_elem.text.replace(find_text, replace_text)
                count_xml += 1
    # 표 치환이 0건이어도 **항상** save_hwpx를 태운다. 예전처럼 count_xml > 0에
    # 걸어 두면 본문에만 매칭된 흔한 경우에 python-hwpx 산출물이 그대로 나가고
    # sanitize_header(검은 배경)·fix_empty_cells·ZIP 레이아웃 복제가 통째로
    # 건너뛰어진다. 같은 명령이 "표에 걸렸는가"에 따라 다른 품질을 내면 안 되고,
    # CLAUDE.md의 "편집은 반드시 save_hwpx()를 통해 저장" 규칙과도 어긋난다.
    save_hwpx(save_path, root, all_files, section_path, output=save_path)

    total = count_runs + count_xml
    print(f"'{find_text}' → '{replace_text}': {total}건 치환 완료 (본문 {count_runs} + 표 {count_xml})")

    # 방금 채운 텍스트가 "한 줄로 입력"(SQUEEZE) 문단에서 과압축되는지 검사
    if total > 0:
        _warn_squeeze_after_edit(root, all_files, replace_text)


def cmd_set_cell(filepath, table_idx, row_idx, col_idx, text, output=None):
    """표 셀의 텍스트를 설정"""
    root, all_files, section_path = open_hwpx(filepath)
    tables = find_tables(root)

    if table_idx >= len(tables):
        print(f"오류: 표 {table_idx}이(가) 없습니다. (총 {len(tables)}개)", file=sys.stderr)
        sys.exit(1)

    rows = get_table_rows(tables[table_idx])
    if row_idx >= len(rows):
        print(f"오류: 행 {row_idx}이(가) 없습니다. (총 {len(rows)}개)", file=sys.stderr)
        sys.exit(1)

    cells = get_row_cells(rows[row_idx])
    if col_idx >= len(cells):
        print(f"오류: 셀 {col_idx}이(가) 없습니다. (총 {len(cells)}개)", file=sys.stderr)
        sys.exit(1)

    cell = cells[col_idx]

    # 셀 안의 첫 번째 <hp:t> 요소를 찾아서 텍스트 설정
    t_elems = cell.findall('.//hp:t', NS)
    if t_elems:
        # hp:t의 tab/lineBreak 같은 자식 뒤 tail에도 기존 본문이 남는다.
        # text만 바꾸면 그 본문이 새 값 뒤에 붙으므로 모든 텍스트 내용을
        # 비운다. 문단·run·hp:t 자체와 서식 속성은 그대로 보존한다.
        for t in t_elems:
            for child in list(t):
                t.remove(child)
            t.text = ''
        # 첫 번째 <hp:t>에 새 값을 넣고 나머지는 빈 텍스트로 유지한다.
        t_elems[0].text = text
    else:
        # <hp:t> 요소가 없으면 첫 번째 run에 추가
        runs = cell.findall('.//hp:run', NS)
        if runs:
            t_new = etree.SubElement(runs[0], '{http://www.hancom.co.kr/hwpml/2011/paragraph}t')
            t_new.text = text
        else:
            print(f"경고: 표{table_idx} 행{row_idx} 셀{col_idx}에 run 요소가 없어 텍스트를 삽입할 수 없습니다.", file=sys.stderr)
            sys.exit(1)

    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"표{table_idx} 행{row_idx} 셀{col_idx} ← \"{text}\"")

    # 방금 채운 텍스트가 "한 줄로 입력"(SQUEEZE) 문단에서 과압축되는지 검사
    _warn_squeeze_after_edit(root, all_files, text)


def _table_grid_metrics(table):
    """표 전체를 훑어 colAddr별 단일 열 너비, rowAddr별 단일 행 높이를 수집.

    병합 해제로 새 셀을 만들 때 cellSz를 채울 기준값. span이 1인 셀만 표본으로
    쓴다(병합 셀의 cellSz는 여러 칸의 합이라 기준이 못 됨).
    """
    col_w = {}
    row_h = {}
    for tr in get_table_rows(table):
        for tc in get_row_cells(tr):
            addr = tc.find('hp:cellAddr', NS)
            sz = tc.find('hp:cellSz', NS)
            if addr is None or sz is None:
                continue
            row_span, col_span = get_cell_span(tc)
            c = _to_int(addr.get('colAddr'), -1)
            r = _to_int(addr.get('rowAddr'), -1)
            if col_span == 1 and c >= 0 and c not in col_w and sz.get('width'):
                col_w[c] = sz.get('width')
            if row_span == 1 and r >= 0 and r not in row_h and sz.get('height'):
                row_h[r] = sz.get('height')
    return col_w, row_h


def _make_unmerged_cell(anchor, col_addr, row_addr, width, height, same_row):
    """병합으로 덮여 있던 자리에 넣을 새 <hp:tc> 생성.

    앵커 셀을 deepcopy해 테두리·여백·문단/글자 스타일을 보존하고, 본문 내용
    (텍스트·개체·중첩 표)만 비운다. linesegarray는 제거해 한글이 재계산하게 한다.
    """
    new_tc = copy.deepcopy(anchor)

    addr = new_tc.find('hp:cellAddr', NS)
    addr.set('colAddr', str(col_addr))
    addr.set('rowAddr', str(row_addr))

    span = new_tc.find('hp:cellSpan', NS)
    if span is not None:
        span.set('rowSpan', '1')
        span.set('colSpan', '1')

    sz = new_tc.find('hp:cellSz', NS)
    if sz is not None:
        if width:
            sz.set('width', width)
        if height:
            sz.set('height', height)

    # 다른 행으로 들어가는 복제본은 머리행 표식을 지운다 (repeatHeader용 header="1")
    if not same_row and 'header' in new_tc.attrib:
        del new_tc.attrib['header']

    # 본문 비우기: 첫 문단의 골격(paraPr/charPr 참조)만 남기고 내용 제거
    sub = new_tc.find('hp:subList', NS)
    if sub is not None:
        first_p = sub.find('hp:p', NS)
        for child in list(sub):
            sub.remove(child)
        if first_p is not None:
            runs = first_p.findall('hp:run', NS)
            for extra in runs[1:]:
                first_p.remove(extra)
            if runs:
                for child in list(runs[0]):
                    runs[0].remove(child)
            for lsa in first_p.findall('hp:linesegarray', NS):
                first_p.remove(lsa)
            sub.append(first_p)
    return new_tc


def cmd_split_cell(filepath, table_idx, row_idx, col_idx, output=None):
    """병합 셀을 완전히 해제: cellSpan을 1x1로 되돌리고, 병합이 덮고 있던
    칸마다 새 <hp:tc>를 생성해 표 그리드를 복원한다.

    주의: span은 <hp:cellAddr>(논리 좌표)가 아니라 별도 <hp:cellSpan>에 있다.
    과거 구현은 cellAddr에서 읽고 써서 (1) 항상 1x1로 읽혀 실제 변화가 없었고
    (2) cellAddr에 스키마상 무의미한 rowSpan/colSpan 속성만 남겼다(이슈 #2,
    get_cell_span의 2026-06-05 수정과 같은 계열의 쓰기 경로 누락).

    또한 span만 1로 바꿔서는 해제가 완성되지 않는다: rowSpan=N인 셀이 덮는
    N-1개 행에는 해당 열의 <hp:tc>가 아예 없으므로 행마다 새로 만들어야 한다
    (colSpan도 동일). 오염된 결과물은 hwpx-validate·--to-md 자가검증을 모두
    통과하므로(reference/warnings-editing.md) 결과 확인은 --info로 할 것.
    """
    root, all_files, section_path = open_hwpx(filepath)
    tables = find_tables(root)

    if table_idx >= len(tables):
        print(f"오류: 표 {table_idx}이(가) 없습니다.", file=sys.stderr)
        sys.exit(1)

    table = tables[table_idx]
    rows = get_table_rows(table)
    if row_idx >= len(rows):
        print(f"오류: 행 {row_idx}이(가) 없습니다.", file=sys.stderr)
        sys.exit(1)

    cells = get_row_cells(rows[row_idx])
    if col_idx >= len(cells):
        print(f"오류: 셀 {col_idx}이(가) 없습니다.", file=sys.stderr)
        sys.exit(1)

    cell = cells[col_idx]
    cell_addr = cell.find('hp:cellAddr', NS)
    if cell_addr is None:
        print("오류: cellAddr 요소가 없습니다.", file=sys.stderr)
        sys.exit(1)

    row_span, col_span = get_cell_span(cell)
    if row_span == 1 and col_span == 1:
        print(f"경고: 표{table_idx} 행{row_idx} 셀{col_idx}은 병합 셀이 아닙니다"
              " (1x1). 변경 없이 종료합니다.", file=sys.stderr)
        sys.exit(1)

    anchor_col = _to_int(cell_addr.get('colAddr'), col_idx)
    anchor_row = _to_int(cell_addr.get('rowAddr'), row_idx)
    col_w, row_h = _table_grid_metrics(table)

    # 단일 칸(1x1) 표본이 없을 때의 크기 폴백: 병합 합계를 span 수로 균등 분할
    # (예: 표 전체가 하나의 병합 셀인 극단 케이스)
    sz = cell.find('hp:cellSz', NS)
    fallback_w = fallback_h = None
    if sz is not None:
        if sz.get('width') and col_span > 1:
            fallback_w = str(_to_int(sz.get('width'), 0) // col_span)
        if sz.get('height') and row_span > 1:
            fallback_h = str(_to_int(sz.get('height'), 0) // row_span)

    # 앵커 셀: span 1x1 + 크기를 단일 칸 기준으로 축소
    span_el = cell.find('hp:cellSpan', NS)
    span_el.set('rowSpan', '1')
    span_el.set('colSpan', '1')
    if sz is not None:
        if row_span > 1 and (row_h.get(anchor_row) or fallback_h):
            sz.set('height', row_h.get(anchor_row) or fallback_h)
        if col_span > 1 and (col_w.get(anchor_col) or fallback_w):
            sz.set('width', col_w.get(anchor_col) or fallback_w)
    # 크기가 바뀐 앵커의 레이아웃 캐시 제거 → 한글이 재계산 (구조적 편집 불변식)
    anchor_sub = cell.find('hp:subList', NS)
    if anchor_sub is not None:
        for p_el in anchor_sub.findall('hp:p', NS):
            for lsa in p_el.findall('hp:linesegarray', NS):
                p_el.remove(lsa)
    # 과거 버그가 cellAddr에 남긴 스키마 외 속성 정리
    for junk in ('rowSpan', 'colSpan'):
        cell_addr.attrib.pop(junk, None)

    # 행 요소를 rowAddr로 찾기 위한 매핑 (tr 순서 == rowAddr가 일반적이나 방어)
    row_by_addr = {}
    for i, tr in enumerate(rows):
        tcs = get_row_cells(tr)
        addr = tcs[0].find('hp:cellAddr', NS) if tcs else None
        key = _to_int(addr.get('rowAddr'), i) if addr is not None else i
        row_by_addr.setdefault(key, tr)

    created = 0
    for r in range(anchor_row, anchor_row + row_span):
        target_tr = row_by_addr.get(r)
        if target_tr is None:
            print(f"경고: rowAddr={r} 행을 찾지 못해 건너뜀", file=sys.stderr)
            continue
        for c in range(anchor_col, anchor_col + col_span):
            if r == anchor_row and c == anchor_col:
                continue
            new_tc = _make_unmerged_cell(
                cell, c, r, col_w.get(c) or fallback_w, row_h.get(r) or fallback_h,
                same_row=(r == anchor_row))
            # colAddr 순서를 유지해 삽입
            inserted = False
            for sibling in get_row_cells(target_tr):
                s_addr = sibling.find('hp:cellAddr', NS)
                if s_addr is not None and _to_int(s_addr.get('colAddr'), -1) > c:
                    sibling.addprevious(new_tc)
                    inserted = True
                    break
            if not inserted:
                target_tr.append(new_tc)
            created += 1

    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"표{table_idx} 행{row_idx} 셀{col_idx}: 병합 해제 "
          f"({row_span}x{col_span} → 1x1, 새 셀 {created}개 생성)")


def cmd_delete_after(filepath, marker_text, output=None):
    """특정 텍스트가 포함된 요소부터 이후의 모든 최상위 요소를 삭제"""
    root, all_files, section_path = open_hwpx(filepath)
    HP_T = f"{{{NS['hp']}}}t"

    elements_to_remove = []
    removing = False
    for child in list(root):
        if not removing:
            for t_elem in child.iter(HP_T):
                if t_elem.text and marker_text in t_elem.text:
                    removing = True
                    break
        if removing:
            elements_to_remove.append(child)

    if not elements_to_remove:
        print(f"'{marker_text}'를 포함하는 요소를 찾지 못했습니다.", file=sys.stderr)
        sys.exit(1)

    for elem in elements_to_remove:
        root.remove(elem)

    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"'{marker_text}' 이후 {len(elements_to_remove)}개 요소 삭제 완료")


def _rows_in_vertical_merge(rows, indices):
    """indices 중 세로 병합(rowSpan>1)의 시작 행이거나 그 병합에 덮인 행의 목록."""
    hit = set()
    for r, row in enumerate(rows):
        for cell in get_row_cells(row):
            row_span, _ = get_cell_span(cell)
            if row_span > 1:
                hit.update(i for i in indices if r <= i < r + row_span)
    return sorted(hit)


def _renumber_rows(table):
    """행 삭제 뒤 rowCnt와 각 셀의 rowAddr를 실제 행 순서에 맞춘다
    (필수 규칙: rowCnt == 실제 tr 수, rowAddr 순차)."""
    rows = get_table_rows(table)
    table.set('rowCnt', str(len(rows)))
    for i, row in enumerate(rows):
        for addr in row.findall('hp:tc/hp:cellAddr', NS):
            addr.set('rowAddr', str(i))


def _refuse_merged_rows(table_idx, rows, indices):
    merged = _rows_in_vertical_merge(rows, indices)
    if merged:
        print(f"오류: 표 {table_idx}의 행 {merged}이(가) 세로 병합에 걸려 있어 지우면 표가 깨집니다. "
              f"먼저 --split-cell로 병합을 풀거나 한글에서 지우세요.", file=sys.stderr)
        sys.exit(1)


def cmd_delete_empty_rows(filepath, table_idx, output=None):
    """테이블 끝의 빈 행들을 삭제"""
    root, all_files, section_path = open_hwpx(filepath)
    tables = find_tables(root)

    if table_idx >= len(tables):
        print(f"오류: 표 {table_idx}이(가) 없습니다. (총 {len(tables)}개)", file=sys.stderr)
        sys.exit(1)

    table = tables[table_idx]
    rows = get_table_rows(table)
    empty = []

    for i in range(len(rows) - 1, 0, -1):
        row = rows[i]
        is_empty = True
        for t_elem in row.findall('.//hp:t', NS):
            if t_elem.text and t_elem.text.strip():
                is_empty = False
                break
        if is_empty:
            empty.append(i)
        else:
            break

    if not empty:
        print(f"표 {table_idx}: 빈 행이 없습니다.")
        return
    _refuse_merged_rows(table_idx, rows, empty)
    for i in empty:
        table.remove(rows[i])
    removed_count = len(empty)
    _renumber_rows(table)

    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"표 {table_idx}: 끝에서 {removed_count}개 빈 행 삭제 완료")


def cmd_trim_cell_text(filepath, table_idx, row_idx, col_idx, output=None):
    """셀 텍스트에서 첫 번째 줄만 유지 (\\n 이후 삭제)"""
    root, all_files, section_path = open_hwpx(filepath)
    tables = find_tables(root)

    if table_idx >= len(tables):
        print(f"오류: 표 {table_idx}이(가) 없습니다.", file=sys.stderr)
        sys.exit(1)

    rows = get_table_rows(tables[table_idx])
    if row_idx >= len(rows):
        print(f"오류: 행 {row_idx}이(가) 없습니다.", file=sys.stderr)
        sys.exit(1)

    cells = get_row_cells(rows[row_idx])
    if col_idx >= len(cells):
        print(f"오류: 셀 {col_idx}이(가) 없습니다.", file=sys.stderr)
        sys.exit(1)

    cell = cells[col_idx]
    trimmed = False

    for t_elem in cell.findall('.//hp:t', NS):
        if t_elem.text and '\n' in t_elem.text:
            old_text = t_elem.text
            t_elem.text = t_elem.text.split('\n', 1)[0]
            trimmed = True

    if not trimmed:
        print(f"표{table_idx} 행{row_idx} 셀{col_idx}: 줄바꿈이 없어 변경 없음.")
        return

    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"표{table_idx} 행{row_idx} 셀{col_idx}: 첫 줄만 유지")


def cmd_delete_rows(filepath, table_idx, row_indices, output=None):
    """테이블에서 특정 행들을 삭제 (인덱스 목록)"""
    root, all_files, section_path = open_hwpx(filepath)
    tables = find_tables(root)

    if table_idx >= len(tables):
        print(f"오류: 표 {table_idx}이(가) 없습니다.", file=sys.stderr)
        sys.exit(1)

    table = tables[table_idx]
    rows = get_table_rows(table)
    targets = []

    for idx in sorted(set(row_indices), reverse=True):
        if idx >= len(rows):
            print(f"경고: 행 {idx}이(가) 없습니다. 건너뜁니다.", file=sys.stderr)
            continue
        targets.append(idx)

    if not targets:
        print("삭제할 행이 없습니다.")
        return
    _refuse_merged_rows(table_idx, rows, targets)
    for idx in targets:
        table.remove(rows[idx])
    removed = targets
    _renumber_rows(table)

    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"표 {table_idx}: 행 {removed} 삭제 완료")


def cmd_sanitize(filepath, output=None):
    """header.xml의 borderFill 문제를 수정하여 저장 (독립 실행용)"""
    root, all_files, section_path = open_hwpx(filepath)
    fixed = sanitize_header(all_files)
    if fixed == 0:
        print("수정할 항목이 없습니다.")
        return
    if output is None:
        output = get_output_path(filepath)
    write_hwpx_zip(output, all_files, zip_layout(filepath))
    print(f"저장: {output}")


def cmd_fix_empty_cells(filepath, output=None):
    """표 셀에 빠진 <hp:p>를 채워 한글 크래시 문제 수정 (Pandoc 변환물 후처리용)"""
    root, all_files, section_path = open_hwpx(filepath)
    fixed = fix_empty_cells(root)
    if fixed == 0:
        print("빈 셀이 없습니다 (추가 수정 불필요).")
        return
    # save_hwpx가 다시 fix_empty_cells를 호출하지만, 이미 보정되어 0건 반환(문제 없음)
    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"빈 셀 {fixed}개에 기본 문단 삽입 완료")


# ── "한 줄로 입력"(lineWrap="SQUEEZE") 과압축 감지·보정 ─────────────────────
# 양식(템플릿)의 제목 문단에 문단 모양 "한 줄로 입력"(paraPr breakSetting
# lineWrap="SQUEEZE")이 걸려 있으면, 긴 텍스트를 채웠을 때 한글이 줄바꿈 대신
# 장평을 무제한 압축해 글자가 겹쳐 뭉개진다(짧은 제목 전제로 디자인된 양식에
# 긴 제목을 채우는 자동화 워크플로우의 대표 함정, 2026-07-06 보도자료 실측).
# 아래 도구는 SQUEEZE 문단 중 추정 자연 폭이 가용 폭을 초과하는 것만 골라
# lineWrap="BREAK"로 전환한다(같은 paraPr을 쓰는 짧은 라벨은 건드리지 않도록
# paraPr을 복제해 해당 문단만 재지정).

# 추정 자연 폭이 가용 폭의 이 배율을 넘으면 과압축으로 판정.
# "한 줄로 입력"의 정상 용도(살짝 조임)는 남기고 뭉개지는 경우만 잡는다.
SQUEEZE_FLAG_RATIO = 1.1


def _find_header_path(all_files):
    """ZIP 내 header.xml 경로를 반환 (통상 Contents/header.xml)."""
    for name in all_files:
        if name.endswith('header.xml'):
            return name
    return None


def _to_int(value, default):
    """XML 속성값을 int로 안전 변환 (외부 생성 HWPX의 비정상 속성값 방어)."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _char_width_units(ch):
    """글자 폭 근사 계수: 전각(한글·한자·전각기호)=1.0, 반각(ASCII 등)=0.55."""
    return 1.0 if ord(ch) >= 0x1100 else 0.55


def _estimate_para_width(p_elem, char_metrics):
    """문단의 자연(비압축) 렌더링 폭을 HWPUNIT으로 근사.

    charPr height(1pt=100 HWPUNIT)를 전각 글자 폭으로 보고 반각은 0.55배.
    charPr의 장평(ratio %)·자간(spacing %)도 반영한다. 표·개체가 섞인 run은
    텍스트만 계산한다.
    """
    total = 0.0
    text_parts = []
    for run in p_elem.findall('hp:run', NS):
        height, ratio, spacing = char_metrics.get(
            run.get('charPrIDRef', '0'), (1000, 100, 0))
        per_char = height * (ratio / 100.0) * (1.0 + spacing / 100.0)
        for t_elem in run.findall('hp:t', NS):
            s = t_full_text(t_elem)
            text_parts.append(s)
            total += sum(_char_width_units(ch) for ch in s) * per_char
    return total, ''.join(text_parts)


def _available_width(p_elem, root):
    """문단이 놓인 컨테이너(셀·글상자·본문)의 가용 줄 폭을 HWPUNIT으로 근사."""
    node = p_elem.getparent()
    while node is not None:
        name = localname(node.tag)
        if name == 'tc':
            cell_sz = node.find('hp:cellSz', NS)
            margin = node.find('hp:cellMargin', NS)
            pad = 1020
            if margin is not None:
                pad = (_to_int(margin.get('left'), 510)
                       + _to_int(margin.get('right'), 510))
            if cell_sz is not None:
                width = _to_int(cell_sz.get('width'), 0)
                if width > 0:
                    return max(width - pad, 1000)
        elif name == 'drawText':
            width = _to_int(node.get('lastWidth'), 0)
            if width <= 0:
                # lastWidth(레이아웃 캐시)가 없으면 도형 자체의 크기로 폴백
                shape = node.getparent()
                if shape is not None:
                    sz = shape.find('hp:sz', NS)
                    if sz is not None:
                        width = _to_int(sz.get('width'), 0)
            margin = node.find('hp:textMargin', NS)
            pad = 566
            if margin is not None:
                pad = (_to_int(margin.get('left'), 283)
                       + _to_int(margin.get('right'), 283))
            if width > 0:
                return max(width - pad, 1000)
        node = node.getparent()

    # 본문: 용지 폭 - 좌우 여백 - 제본 여백
    page_pr = root.find('.//hp:pagePr', NS)
    if page_pr is not None:
        width = _to_int(page_pr.get('width'), 59528)
        margin = page_pr.find('hp:margin', NS)
        left = _to_int(margin.get('left'), 8504) if margin is not None else 8504
        right = _to_int(margin.get('right'), 8504) if margin is not None else 8504
        gutter = _to_int(margin.get('gutter'), 0) if margin is not None else 0
        return max(width - left - right - gutter, 1000)
    return 42520  # A4 기본 여백 근사 폴백


def find_squeeze_overflows(root, all_files):
    """lineWrap="SQUEEZE" 문단 중 텍스트가 가용 폭을 초과해 과압축될 문단 목록.

    반환: [{'p_elem', 'para_pr_id', 'text', 'est_width', 'avail_width'}]
    """
    header_path = _find_header_path(all_files)
    if header_path is None:
        return []
    header_root = etree.fromstring(all_files[header_path])

    squeeze_ids = set()
    for para_pr in header_root.findall('.//hh:paraPr', NS):
        pr_id = para_pr.get('id')
        if pr_id is None:
            continue
        bs = para_pr.find('hh:breakSetting', NS)
        if bs is not None and bs.get('lineWrap') == 'SQUEEZE':
            squeeze_ids.add(pr_id)
    if not squeeze_ids:
        return []

    # charPr id → (height, ratio %, spacing %) : 폭 추정에 장평·자간 반영
    char_metrics = {}
    for char_pr in header_root.findall('.//hh:charPr', NS):
        cp_id = char_pr.get('id')
        if cp_id is None:
            continue
        height = _to_int(char_pr.get('height'), 1000)
        ratio_el = char_pr.find('hh:ratio', NS)
        spacing_el = char_pr.find('hh:spacing', NS)
        ratio = _to_int(ratio_el.get('hangul'), 100) if ratio_el is not None else 100
        spacing = _to_int(spacing_el.get('hangul'), 0) if spacing_el is not None else 0
        char_metrics[cp_id] = (height, ratio, spacing)

    issues = []
    for p_elem in root.iter('{%s}p' % NS['hp']):
        para_pr_id = p_elem.get('paraPrIDRef')
        if para_pr_id is None or para_pr_id not in squeeze_ids:
            continue
        est_width, text = _estimate_para_width(p_elem, char_metrics)
        if not text.strip():
            continue
        avail = _available_width(p_elem, root)
        if est_width > avail * SQUEEZE_FLAG_RATIO:
            issues.append({
                'p_elem': p_elem,
                'para_pr_id': para_pr_id,
                'text': text,
                'est_width': int(est_width),
                'avail_width': avail,
            })
    return issues


def _warn_squeeze_after_edit(root, all_files, inserted_text):
    """편집 직후, 방금 넣은 텍스트가 SQUEEZE 문단에서 과압축되는지 검사해 경고.

    파일을 다시 읽지 않고 편집에 사용한 메모리 상태(root, all_files)를 재사용한다.
    """
    if not inserted_text or not inserted_text.strip():
        return
    try:
        issues = find_squeeze_overflows(root, all_files)
    except Exception:
        return  # 경고는 부가 기능: 검사 실패가 편집 성공을 가리면 안 됨
    for issue in issues:
        if inserted_text in issue['text']:
            preview = issue['text'][:40]
            print(
                f"경고: '{preview}…' 문단에 \"한 줄로 입력\"(lineWrap=SQUEEZE)이 걸려 있어 "
                f"긴 텍스트가 한 줄로 과압축됩니다(추정 폭 {issue['est_width']} > "
                f"가용 {issue['avail_width']} HWPUNIT). "
                f"`--fix-squeeze`로 자연 줄바꿈으로 전환하세요.",
                file=sys.stderr,
            )


def cmd_list_squeeze(filepath):
    """"한 줄로 입력" 과압축 문단을 나열만 한다 (파일 변경 없음)."""
    root, all_files, _ = open_hwpx(filepath)
    issues = find_squeeze_overflows(root, all_files)
    if not issues:
        print('과압축되는 "한 줄로 입력"(SQUEEZE) 문단이 없습니다.')
        return
    print(f'과압축 SQUEEZE 문단 {len(issues)}개:')
    for i, issue in enumerate(issues):
        preview = issue['text'][:50] + ('…' if len(issue['text']) > 50 else '')
        print(f"  [{i}] paraPr={issue['para_pr_id']} "
              f"추정 {issue['est_width']} / 가용 {issue['avail_width']} HWPUNIT | {preview}")
    print('`--fix-squeeze`로 해당 문단만 자연 줄바꿈(BREAK)으로 전환할 수 있습니다.')


def cmd_fix_squeeze(filepath, output=None):
    """과압축되는 "한 줄로 입력" 문단을 자연 줄바꿈으로 전환.

    같은 paraPr을 공유하는 짧은 라벨(의도된 디자인)은 보존하기 위해, 원본
    paraPr을 수정하지 않고 lineWrap="BREAK"인 복제 paraPr을 새 id로 추가한 뒤
    과압축 문단만 재지정한다. 재지정 문단의 linesegarray(레이아웃 캐시)는
    제거해 한글이 다시 계산하게 한다.
    """
    root, all_files, section_path = open_hwpx(filepath)
    issues = find_squeeze_overflows(root, all_files)
    if not issues:
        print('과압축되는 "한 줄로 입력"(SQUEEZE) 문단이 없습니다. 변경 없음.')
        return

    header_path = _find_header_path(all_files)
    header_root = etree.fromstring(all_files[header_path])
    para_props = header_root.find('.//hh:paraProperties', NS)
    if para_props is None:
        print("오류: header.xml에 paraProperties가 없어 보정할 수 없습니다.", file=sys.stderr)
        sys.exit(1)
    all_para_prs = para_props.findall('hh:paraPr', NS)
    max_id = max((_to_int(pp.get('id'), 0) for pp in all_para_prs), default=0)

    def _find_squeeze_source(pr_id):
        """복제 원본 paraPr 선택. id가 중복된 비정상 header에서도 SQUEEZE가
        걸린 항목을 우선해 엉뚱한 paraPr을 복제하는 일을 막는다."""
        matches = [pp for pp in para_props.findall('hh:paraPr', NS)
                   if pp.get('id') == pr_id]
        for pp in matches:
            bs = pp.find('hh:breakSetting', NS)
            if bs is not None and bs.get('lineWrap') == 'SQUEEZE':
                return pp
        return matches[0] if matches else None

    clone_map = {}  # 원본 paraPr id → BREAK 복제본 id
    fixed = 0
    skipped = 0
    for issue in issues:
        old_id = issue['para_pr_id']
        if old_id not in clone_map:
            src = _find_squeeze_source(old_id)
            if src is None:
                skipped += 1
                print(f"경고: paraPr id={old_id}를 header에서 찾지 못해 건너뜀",
                      file=sys.stderr)
                continue
            new_pr = copy.deepcopy(src)
            max_id += 1
            new_pr.set('id', str(max_id))
            bs = new_pr.find('hh:breakSetting', NS)
            if bs is not None:
                bs.set('lineWrap', 'BREAK')
            para_props.append(new_pr)
            clone_map[old_id] = str(max_id)
        p_elem = issue['p_elem']
        p_elem.set('paraPrIDRef', clone_map[old_id])
        # 레이아웃 캐시 제거 → 한글이 열 때 재계산
        for lsa in p_elem.findall('hp:linesegarray', NS):
            p_elem.remove(lsa)
        fixed += 1
        preview = issue['text'][:50] + ('…' if len(issue['text']) > 50 else '')
        print(f"fix-squeeze: paraPr {old_id} → {clone_map[old_id]} (BREAK) | {preview}")

    if fixed == 0:
        print("오류: 전환할 수 있는 문단이 없습니다 (header paraPr 조회 실패).",
              file=sys.stderr)
        sys.exit(1)

    para_props.set('itemCnt', str(len(para_props.findall('hh:paraPr', NS))))
    all_files[header_path] = serialize_xml(header_root)
    save_hwpx(filepath, root, all_files, section_path, output)
    msg = f'과압축 SQUEEZE 문단 {fixed}개를 자연 줄바꿈으로 전환 완료'
    if skipped:
        msg += f' (건너뜀 {skipped}개)'
    print(msg)


def cmd_add_preview(filepath, output=None):
    """container.xml이 선언했지만 누락된 Preview/PrvText.txt를 section 텍스트로
    생성해 주입한다.

    hwp2hwpx.bat(hwpxlib) 변환물은 container.xml의 <ocf:rootfiles>가
    Preview/PrvText.txt를 rootfile로 선언하면서도 ZIP에 Preview/ 항목을 만들지
    않는다. 한글은 정상적으로 열고 --to-md recall도 100%지만, python-hwpx의
    hwpx-validate는 'Root content Preview/PrvText.txt is missing'으로 실패한다.
    변환물을 정본·배포·검증 대상으로 쓸 때만 보정하면 된다(읽기·편집엔 무해).

    section XML은 재직렬화하지 않고 원본 바이트를 그대로 보존한다(최소 침습).
    mimetype을 첫 항목·ZIP_STORED로 재패키징한다(hwpx 규칙).
    """
    if is_encrypted_hwpx(filepath):
        print(f"오류: {ENCRYPTION_HINT}", file=sys.stderr)
        sys.exit(1)

    with zipfile.ZipFile(filepath, 'r') as zf:
        names = zf.namelist()
        data = {n: zf.read(n) for n in names}

    sanitize_xml_parts(data)

    if 'Preview/PrvText.txt' in data:
        print("Preview/PrvText.txt가 이미 있습니다. 변경 없음.")
        return

    # section 텍스트로 미리보기 본문 생성 (재직렬화 없이 파싱만)
    section_files = [n for n in names
                     if 'section' in n.lower() and n.endswith('.xml')]
    lines = []
    if section_files:
        sroot = etree.fromstring(data[section_files[0]])
        for t_elem in sroot.findall('.//hp:t', NS):
            s = t_full_text(t_elem)
            if s and s.strip():
                lines.append(s.strip())
    prv = '\n'.join(lines)

    if output is None:
        output = get_output_path(filepath)

    # mimetype 첫 항목·ZIP_STORED, 나머지는 원본 바이트·압축 방식 그대로 + PrvText 추가
    data['Preview/PrvText.txt'] = prv.encode('utf-8')
    write_hwpx_zip(output, data, zip_layout(filepath))
    print(f"add-preview: Preview/PrvText.txt 생성 ({len(prv)}자) → {output}")


def get_hwp_automation_module_path():
    """한컴 자동화 보안 모듈 레지스트리 경로를 반환."""
    if winreg is None:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\HNC\HwpAutomation\Modules") as key:
            value, _ = winreg.QueryValueEx(key, "FilePathCheckerModule")
            return value
    except OSError:
        return None


def create_hwp_com():
    """HWPFrame.HwpObject COM 객체를 생성하고 필요한 모듈을 등록."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("pywin32가 필요합니다. `pip install pywin32` 후 다시 실행하세요.") from exc

    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        module_path = get_hwp_automation_module_path()
        if module_path and os.path.exists(module_path):
            # 보안 팝업 없이 로컬 파일을 열고 저장하기 위한 한컴 공식 샘플 모듈.
            try:
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except Exception as exc:
                print(f"경고: FilePathCheckerModule 등록 실패: {exc}", file=sys.stderr)
        return pythoncom, hwp
    except Exception:
        pythoncom.CoUninitialize()
        raise


def close_hwp_com(pythoncom, hwp):
    """COM 객체와 COM apartment를 정리."""
    if hwp is not None:
        try:
            hwp.Quit()
        except Exception:
            pass
    if pythoncom is not None:
        pythoncom.CoUninitialize()


def cmd_diagnose_com():
    """한컴 COM 기반 작업 가능 여부를 점검."""
    print("한컴 COM 진단")
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
        print("- pywin32: OK")
    except ImportError as exc:
        print(f"- pywin32: 오류 ({exc})")
        return 1

    module_path = get_hwp_automation_module_path()
    if module_path:
        status = "OK" if os.path.exists(module_path) else "경고: 파일 없음"
        print(f"- FilePathCheckerModule: {status}")
        print(f"  {module_path}")
    else:
        print("- FilePathCheckerModule: 경고: 레지스트리 등록 없음")

    pythoncom_mod = None
    hwp = None
    try:
        pythoncom_mod, hwp = create_hwp_com()
        try:
            version = hwp.Version
        except Exception:
            version = "(버전 확인 실패)"
        print("- HWPFrame.HwpObject: OK")
        print(f"- 한컴 버전: {version}")
        return 0
    except Exception as exc:
        print(f"- HWPFrame.HwpObject: 오류 ({exc})")
        return 1
    finally:
        close_hwp_com(pythoncom_mod, hwp)


ADOPTIUM_JDK21_GLOB = r"C:\Program Files\Eclipse Adoptium\jdk-21*"


def find_windows_java(environ, adoptium_glob=ADOPTIUM_JDK21_GLOB, which=None):
    """hwp2hwpx.bat과 같은 순서(JAVA_HOME → Adoptium jdk-21* → PATH)로 java.exe를 찾는다.
    Adoptium이 여럿이면 bat의 for 루프처럼 이름순 마지막(높은 버전)을 쓴다."""
    import glob as _glob
    home = environ.get("JAVA_HOME")
    if home:
        exe = os.path.join(home, "bin", "java.exe")
        if os.path.exists(exe):
            return exe
    for d in sorted(_glob.glob(adoptium_glob), reverse=True):
        exe = os.path.join(d, "bin", "java.exe")
        if os.path.exists(exe):
            return exe
    return (which or shutil.which)("java")


def find_java():
    """hwp2hwpx 래퍼가 실제로 쓸 java. sh는 JAVA_HOME → PATH, bat은 find_windows_java 순서."""
    if sys.platform == "win32":
        return find_windows_java(os.environ)
    home = os.environ.get("JAVA_HOME")
    exe = os.path.join(home, "bin", "java") if home else None
    return exe if exe and os.access(exe, os.X_OK) else shutil.which("java")


# ---------- HWP → HWPX 변환 (rhwp 기본, JAR 폴백) ----------

ConvertResult = namedtuple("ConvertResult", "path engine status messages")
HWP_ENGINES = ("rhwp", "jar")
CONVERT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "convert")


def _jar_script():
    return os.path.join(CONVERT_DIR, "hwp2hwpx.bat" if sys.platform == "win32" else "hwp2hwpx.sh")


def _run_jar(src, out):
    """번들 hwp2hwpx JAR로 변환한다(PrvText 보정 전 원시 산출물)."""
    script = _jar_script()
    cmd = [script, src, out] if sys.platform == "win32" else ["sh", script, src, out]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    if r.returncode != 0 or not os.path.exists(out) or os.path.getsize(out) == 0:
        tail = (r.stderr or r.stdout).strip()[-300:]
        raise RuntimeError(f"JAR 변환 실패(exit {r.returncode}): {tail}")


def _rhwp_export_hwpx(src, out, verify=None):
    return rhwp_cli.run(["export-hwpx", os.path.abspath(src), rhwp_cli.OUTPUT, "--verify"],
                        output=out, required_fields=("output", "format", "verify"), verify=verify)


def convert_fallback_hint(src, dest, failed_engine):
    """변환 실패 시 사람이 그대로 따라 칠 폴백 명령(어느 폴더에서든 되게 절대 경로)."""
    edit = f'python "{os.path.abspath(__file__)}"'
    com = f'  한컴 COM: python "{os.path.join(os.path.dirname(CONVERT_DIR), "hwpx_com.py")}" "{src}" --from-hwp'
    if failed_engine == "jar":
        lines = [f'  rhwp: {edit} "{src}" --from-hwp']
        if sys.platform == "win32":
            lines.append(com)
        return "폴백:\n" + "\n".join(lines)
    script = _jar_script() if sys.platform == "win32" else f'bash "{_jar_script()}"'
    hint = ("폴백(JAR, JDK 21 필요):\n"
            f'  {edit} "{src}" --from-hwp --engine jar\n'
            "  또는 손으로 두 단계:\n"
            f'    {script} "{src}" "{dest}"\n'
            f'    {edit} "{dest}" --add-preview -o "{dest}"')
    if sys.platform == "win32":
        hint += "\n폴백(JDK 없는 Windows + 한컴오피스):\n" + com
    return hint


def _independent_check(staged, src, engine, messages):
    """다른 엔진 변환물과 교차 대조하고 hwpx-validate를 돌린다.
    불일치·검증 실패는 RuntimeError, 대조 불가는 "unverified"를 돌려준다."""
    import hwpx_gate

    status = "passed"
    ok, out = hwpx_gate.validate(staged)
    if ok is None:
        status = "unverified"
        messages.append(out)
    elif not ok:
        raise RuntimeError("hwpx-validate 실패: " + out.splitlines()[-1][:160])

    other = "jar" if engine == "rhwp" else "rhwp"
    with tempfile.TemporaryDirectory(prefix="hwp-ref-") as td:
        ref = os.path.join(td, "ref.hwpx")
        if other == "jar" and not find_java():
            messages.append("JDK 없음: JAR 교차 대조 생략(unverified)")
            return "unverified"
        try:
            _run_jar(src, ref) if other == "jar" else _rhwp_export_hwpx(src, ref)
        except (RuntimeError, OSError, subprocess.SubprocessError) as exc:
            messages.append(f"{other} 교차 대조용 변환 실패, 대조 생략(unverified): {exc}")
            return "unverified"
        if not zipfile.is_zipfile(ref):
            messages.append(f"{other} 교차 대조용 변환물이 HWPX가 아님, 대조 생략(unverified)")
            return "unverified"
        try:
            rep = hwpx_gate.cross_check(staged, ref)
        except Exception as exc:  # 깨진 zip·XML: 산출물을 믿을 근거가 없으니 거부
            raise RuntimeError(f"교차 대조 중 산출물을 읽지 못함: {type(exc).__name__}: {exc}") from exc
    if rep.status == "failed":
        raise RuntimeError(f"{engine}·{other} 교차 대조 실패: {rep.summary()}")
    messages.append(f"교차 대조({other}): {rep.summary()}")
    return status if rep.status == "passed" else "unverified"


def hwp_to_hwpx(src, dest=None, *, engine="rhwp"):
    """HWP를 HWPX로 변환한다. 정본 형식은 HWPX이고 `.hwp`는 이 입구에서만 만난다.

    dest가 None이면 입력 옆 `_work-hwpx-automation/<stem>.hwpx`.
    산출물은 hwpx-validate와 다른 엔진 변환물과의 교차 대조(hwpx_gate.cross_check)를
    통과해야 dest에 놓인다. 대조 불가(JDK·hwpx-tomd 없음)면 status="unverified".
    실패하면 RuntimeError(rhwp_cli.RhwpError 포함)이고 메시지에 폴백 명령이 들어 있다.
    거부된 산출물은 `<dest stem>.rejected.hwpx`로 남는다.
    stdout 마지막에 `engine=<rhwp|jar> status=<passed|unverified>`를 찍는다.
    """
    if engine not in HWP_ENGINES:
        raise ValueError(f"engine은 {HWP_ENGINES} 중 하나: {engine!r}")
    src = os.path.abspath(src)
    dest = os.path.abspath(dest or get_output_path(src, ".hwpx"))
    if not dest.lower().endswith(".hwpx") or os.path.normcase(dest) == os.path.normcase(src):
        raise ValueError(f"변환 출력은 원본과 다른 .hwpx 경로여야 합니다: {dest}")
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    messages = []
    state = {}

    def check(staged):
        try:
            state["status"] = _independent_check(staged, src, engine, messages)
        except RuntimeError:
            root, ext = os.path.splitext(dest)
            shutil.copyfile(staged, f"{root}.rejected{ext}")
            messages.append(f"거부된 산출물 보존: {root}.rejected{ext}")
            raise

    try:
        if engine == "rhwp":
            _rhwp_export_hwpx(src, dest, verify=check)
        else:
            with tempfile.TemporaryDirectory(prefix="jar-", dir=os.path.dirname(dest)) as td:
                raw, staged = os.path.join(td, "raw.hwpx"), os.path.join(td, "staged.hwpx")
                _run_jar(src, raw)
                with open(os.devnull, "w") as quiet, contextlib.redirect_stdout(quiet):
                    cmd_add_preview(raw, staged)  # JAR 산출물의 PrvText 누락 보정
                if os.path.exists(staged):
                    messages.append("JAR 산출물에 Preview/PrvText.txt 보정(--add-preview)")
                else:
                    shutil.copyfile(raw, staged)
                check(staged)
                os.replace(staged, dest)
    except (RuntimeError, OSError, subprocess.SubprocessError) as exc:
        detail = [str(exc)] + [m for m in messages if m.startswith("거부된")]
        rejected = getattr(exc, "rejected_path", None)
        if rejected:
            detail.append(f"거부된 산출물 보존: {rejected}")
        raise RuntimeError("\n".join(detail) + "\n" + convert_fallback_hint(src, dest, engine)) from exc

    status = state.get("status", "unverified")
    for m in messages:
        print(m)
    print(f"변환: {src} → {dest}")
    print(f"engine={engine} status={status}")
    return ConvertResult(dest, engine, status, messages)


def convert_hwp_or_exit(src, dest=None, engine="rhwp"):
    try:
        return hwp_to_hwpx(src, dest, engine=engine)
    except (RuntimeError, ValueError) as exc:
        print(f"오류: HWP→HWPX 변환 실패 (engine={engine})\n{exc}", file=sys.stderr)
        sys.exit(1)


def cmd_check_env():
    """기능 계층(tier)별 의존성·런타임 준비 상태를 한 번에 리포트.

    처음 클론한 사용자가 어느 워크플로우까지 바로 쓸 수 있는지, 무엇을 더
    설치해야 하는지 한눈에 확인한다. 읽기 전용이며 아무것도 설치·변경하지 않는다.
    이 스킬은 API 키를 쓰지 않으므로(전부 로컬 도구) 점검 대상은 pip 패키지와
    시스템 런타임(JDK, Pandoc, 한컴오피스 COM)이다. COM 계층의 상세 진단(보안모듈
    DLL·레지스트리·한컴 기동)은 중복하지 않고 --diagnose-com으로 위임한다.
    """
    import glob as _glob
    import importlib.util

    try:
        sys.stdout.reconfigure(encoding="utf-8")  # Windows cp949 안전
    except Exception:
        pass

    def installed(module):
        try:
            return importlib.util.find_spec(module) is not None
        except (ImportError, ModuleNotFoundError, ValueError):
            return False

    here = os.path.dirname(os.path.abspath(__file__))
    convert_dir = os.path.join(here, "convert")
    ready = []  # 바로 사용 가능한 tier 이름

    print("hwpx-automation 환경 점검")
    print("=" * 48)

    # Tier 1: 읽기·편집 (필수, pip만)
    print("\n[Tier 1] 읽기·편집 (필수 · pip)")
    tier1 = [
        ("python-hwpx", "hwpx", "편집 명령(--find/--set-cell 등)"),
        ("lxml", "lxml", "XML 파이프라인 전반"),
        ("hwpx-tomd", "hwpx_tomd", "--to-md 변환 엔진"),
    ]
    t1_ok = True
    for pip_name, module, use in tier1:
        if installed(module):
            print(f"  [O] {pip_name:<14} {use}")
        else:
            print(f"  [설치] {pip_name:<12} pip install {pip_name}  ({use})")
            t1_ok = False
    if t1_ok:
        ready.append("읽기·편집")

    # Tier 2: HWP→HWPX 변환·PDF·이미지 (rhwp, 모든 OS 기본 엔진)
    print("\n[Tier 2] HWP→HWPX 변환·PDF·이미지 (rhwp · 모든 OS 기본 엔진)")
    rhwp = rhwp_cli.find_rhwp()
    if rhwp:
        try:
            notes = rhwp_cli.check_version(rhwp)
            print(f"  [O] rhwp           {rhwp_cli.fmt(rhwp_cli.version(rhwp))}  {rhwp}")
            for note in notes:
                print(f"  [경고] {note}")
            ready.append("HWP→HWPX 변환·PDF(rhwp)")
        except rhwp_cli.RhwpError as exc:
            print(f"  [경고] rhwp         {exc}")
    else:
        print("  [설치] rhwp         https://github.com/edwardkim/rhwp/releases 바이너리를 PATH에 "
              f"(또는 {rhwp_cli.ENV_BIN}), v{'.'.join(map(str, rhwp_cli.MIN_VERSION))} 이상")

    # Tier 3: MD/DOCX/HTML → HWPX 변환 (pip + Pandoc)
    print("\n[Tier 3] MD/DOCX/HTML→HWPX 변환 (pip + Pandoc)")
    t3_ok = True
    if installed("pypandoc_hwpx"):
        print(f"  [O] pypandoc-hwpx  convert/hwpx_convert.py 변환")
    else:
        print(f"  [설치] pypandoc-hwpx  pip install pypandoc-hwpx")
        t3_ok = False
    pandoc = shutil.which("pandoc")
    if pandoc:
        print(f"  [O] pandoc         {pandoc}")
    else:
        print(f"  [경고] pandoc       Pandoc 바이너리 없음 (pypandoc-hwpx가 번들 제공할 수도 있음)")
        print(f"          필요 시 설치: https://pandoc.org/installing.html")
    if t3_ok:
        ready.append("MD→HWPX 변환")

    # Tier 4: 폴백 (JDK 21 + 번들 JAR, Windows 한컴 COM). 사용자가 고를 때만 쓴다.
    print("\n[Tier 4] 폴백 (JDK 21 + 번들 JAR · Windows 한컴 COM, 명시할 때만)")
    jar_ok = True

    def java_version(exe):
        try:
            out = subprocess.run([exe, "-version"], capture_output=True, text=True, timeout=10)
            lines = (out.stderr or out.stdout or "").splitlines()
            return lines[0].strip() if lines else ""
        except Exception:
            return ""

    found_java = find_java()  # 점검과 실제 실행(hwp2hwpx.sh/.bat)이 어긋나지 않게 같은 탐색
    if found_java:
        ver = java_version(found_java)
        print(f"  [O] JDK          {found_java}{('  ·  ' + ver) if ver else ''}")
    else:
        print(f"  [설치] JDK        JDK 21 설치(또는 JAVA_HOME 지정): https://adoptium.net/ (Temurin 21)")
        jar_ok = False
    for name, pattern in [
        ("hwp2hwpx", os.path.join(convert_dir, "hwp2hwpx*.jar")),
        ("hwplib", os.path.join(convert_dir, "lib", "hwplib*.jar")),
        ("hwpxlib", os.path.join(convert_dir, "lib", "hwpxlib*.jar")),
    ]:
        if _glob.glob(pattern):
            print(f"  [O] {name:<14} JAR 번들됨")
        else:
            print(f"  [경고] {name:<12} JAR 없음: {pattern}")
            jar_ok = False
    if jar_ok:
        print("      --engine jar 변환과 rhwp 변환의 교차 대조에 쓰인다")
        ready.append("JAR 폴백")
    if sys.platform != "win32":
        print("  [-] 한컴 COM 폴백(--pdf-backend com·서명·한컴 정규화)은 Windows + 한컴오피스 전용")
    else:
        if installed("win32com"):
            print(f"  [O] pywin32        COM 폴백 의존성(--pdf-backend com, hwpx_sign.py 등)")
            ready.append("한컴 COM 폴백")
        else:
            print(f"  [설치] pywin32      pip install pywin32")
        print("  [폴백] JDK 없이 한컴으로 HWP 변환: python hwpx_com.py 파일.hwp --from-hwp"
              "  (사전 확인: python hwpx_com.py --diagnose)")
        if installed("pyhwpx"):
            print(f"  [O] pyhwpx         hwpx_com.py 네이티브 파이프라인(선택)")
        else:
            print(f"  [선택] pyhwpx       pip install pyhwpx (hwpx_com.py 사용 시에만)")
        print("  상세(보안모듈 DLL·레지스트리·한컴 기동): python hwpx_edit.py --diagnose-com")

    print("\n" + "=" * 48)
    print("바로 사용 가능: " + (", ".join(ready) if ready else "(없음): 최소 Tier 1을 먼저 설치"))
    print("핵심 한 번에 설치: pip install -r requirements.txt")
    print("API 키는 필요 없음(전부 로컬 도구). Tier 2~4는 필요한 워크플로우에서만 설치하면 된다.")
    return 0


PDF_BACKENDS = ("rhwp", "com")


def cmd_to_pdf(filepath, output=None, password=None, backend="rhwp"):
    """HWP/HWPX를 PDF로 저장. 모든 OS에서 rhwp가 기본이고, com(Windows 한컴오피스)은
    한컴과 같은 조판이 필요할 때 명시하는 폴백이다."""
    if output is None:
        output = get_output_path(filepath, ".pdf")

    if backend == "rhwp":
        cmd_to_pdf_rhwp(filepath, output, password)
        return
    if sys.platform != "win32":
        print("오류: --pdf-backend com은 Windows + 한컴오피스 전용입니다. "
              "이 OS에서는 기본값(rhwp)을 쓰세요.", file=sys.stderr)
        sys.exit(1)

    abs_input = os.path.abspath(filepath)
    abs_output = os.path.abspath(output)
    os.makedirs(os.path.dirname(abs_output), exist_ok=True)

    ext = os.path.splitext(filepath)[1].lower()
    fmt = "HWP" if ext == ".hwp" else "HWPX"
    open_arg = f"password:{password}" if password else ""

    pythoncom_mod = None
    hwp = None
    try:
        pythoncom_mod, hwp = create_hwp_com()
        opened = hwp.Open(abs_input, fmt, open_arg)
        if not opened:
            raise RuntimeError(f"한컴 COM이 파일을 열지 못했습니다: {abs_input}")
        from pdf_export import save_pdf
        save_pdf(hwp, abs_output)
    except Exception as exc:
        print(f"오류: PDF 저장 실패: {exc}", file=sys.stderr)
        sys.exit(1)
    finally:
        close_hwp_com(pythoncom_mod, hwp)

    if not os.path.exists(abs_output) or os.path.getsize(abs_output) == 0:
        print(f"오류: PDF 출력 파일이 생성되지 않았습니다: {abs_output}", file=sys.stderr)
        sys.exit(1)

    print(f"PDF 저장: {abs_output}")
    print("engine=com")


def cmd_to_pdf_rhwp(filepath, output, password=None):
    """rhwp 렌더링 PDF. 글꼴에 없는 문자가 빈 네모로 찍혔으면 PDF는 남기고 exit 2."""
    from rhwp_pdf import save_pdf
    try:
        report = save_pdf(filepath, output, password)
    except Exception as exc:
        why = "암호 문서는" if password else "rhwp로 안 되면"
        print(f"오류: PDF 저장 실패: {exc}\n  폴백: {why} Windows 한컴오피스에서 --pdf-backend com",
              file=sys.stderr)
        sys.exit(1)

    print(f"PDF 저장: {os.path.abspath(output)}")
    print("  렌더러: rhwp(한컴 아님) · 사용 글꼴: " + ", ".join(report["fonts"]))
    print("engine=rhwp")
    if report["overflow"]:
        print(f"경고: 쪽 하단 넘침 {report['overflow']}건. 쪽 끝 줄이 잘리지 않았는지 렌더를 눈으로 확인하세요.",
              file=sys.stderr)
    if report["missing"]:
        chars = " ".join(f"{c}(U+{ord(c):04X})" for c in report["missing"])
        print(f"오류: 글꼴에 없는 문자가 빈 네모로 찍혔습니다: {chars}. "
              "배포 전에 해당 문자를 바꾸거나 Windows 한컴(--pdf-backend com)으로 변환하세요.", file=sys.stderr)
        sys.exit(2)


# ---------- rhwp 축: 격자 좌표·누름틀·역변환 (HWPX 위에서만) ----------

GRID_HINT = ("격자 좌표는 --tables로 먼저 본다. 원본 보존이 중요한 공문은 기존 "
             "--set-cell T,R,C(lxml, 물리 셀 인덱스)가 1순위다.")
TO_HWP_COM_HINT = "그림·차트가 빠졌으면 Windows 한컴오피스에서 HWPX를 열어 HWP로 다른 이름 저장한다."


def _rhwp_or_exit(fn, *args, **kwargs):
    try:
        return fn(*args, **kwargs)
    except rhwp_cli.RhwpError as exc:
        print(f"오류: {exc}", file=sys.stderr)
        sys.exit(1)


def _keep_rejected(staged, output):
    root, ext = os.path.splitext(output)
    rejected = f"{root}.rejected{ext}"
    shutil.copyfile(staged, rejected)
    return rejected


def _rhwp_edit_output(filepath, output):
    """rhwp 편집의 기본 출력. 입력과 같은 자리가 되면(_work 안 입력) _편집을 붙여
    다음 실행의 비교 기준(원본)을 지키고, 결과는 항상 .hwpx다."""
    output = os.path.abspath(output or get_output_path(filepath))
    if os.path.normcase(output) == os.path.normcase(os.path.abspath(filepath)):
        output = os.path.splitext(output)[0] + "_편집.hwpx"
    return output


def _gated(src, output, report):
    """rhwp 편집 산출물의 게시 전 검증. 게이트 실패·검사 오류면 .rejected로 남기고 예외."""
    import inspect
    import hwpx_gate

    params = inspect.signature(hwpx_gate.gate).parameters.values()
    if not any(p.name == "allow_rhwp" or p.kind is p.VAR_KEYWORD for p in params):
        raise rhwp_cli.RhwpError("이 hwpx_gate.py에는 rhwp 재직렬화 허용 목록(allow_rhwp)이 없어 "
                                 "rhwp 편집 산출물을 판정할 수 없습니다. 저장소를 최신으로 받으세요.")

    def verify(staged):
        try:
            rep = hwpx_gate.gate(src, staged, allow=hwpx_gate.is_text_run_key, allow_rhwp=True,
                                 expect_text_change=True)
        except Exception as exc:
            rejected = _keep_rejected(staged, output)
            raise rhwp_cli.RhwpError(f"게이트 검사 중 오류: {type(exc).__name__}: {exc}\n"
                                     f"거부된 산출물 보존: {rejected}") from exc
        report["gate"] = rep
        if rep.status == "failed":
            rejected = _keep_rejected(staged, output)
            detail = hwpx_gate.describe_lost(rep)
            head = rep.summary().splitlines()[0]  # 허용된 손실 줄은 describe_lost 끝에 이미 있다
            raise rhwp_cli.RhwpError(f"{head}\n{detail}\n거부된 산출물 보존: {rejected}")
    return verify


def _print_gate(report):
    rep = report["gate"]
    print(rep.summary())
    print(f"engine=rhwp status={'passed' if rep.passed else 'unverified'}")


def _one_line(text):
    """문서 유래 글자를 한 줄로(개행이 가짜 좌표 줄을 만들지 않게)."""
    return str(text or "").replace("\r", " ").replace("\n", " ")


def cmd_tables(filepath):
    """rhwp export-tables 격자를 사람이 읽는 형태로 출력(변경 없음)."""
    r = _rhwp_or_exit(rhwp_cli.run, ["export-tables", os.path.abspath(filepath)],
                      required_fields=("tables",))
    tables = r.envelope["tables"]
    if not tables:
        print("표가 없습니다.")
    for t in tables:
        print(f"표 {t.get('index')} (구역 {t.get('section')}, 문단 {t.get('paragraph')}): "
              f"{t.get('rows')}행 × {t.get('cols')}열")
        for c in t.get("cells", []):
            note = ""
            if c.get("rowSpan", 1) > 1 or c.get("colSpan", 1) > 1:
                note += f" (병합 {c.get('rowSpan', 1)}×{c.get('colSpan', 1)})"
            if c.get("nested"):
                note += " (중첩 표 있음: 안쪽 표는 --set-cell로)"
            print(f"  [{t.get('index')},{c.get('row')},{c.get('col')}]{note} {_one_line(c.get('text'))}")
    print(f"--grid-set-cell T,R,C 좌표는 위 [표,행,열]이다. {GRID_HINT}")


def cmd_fields(filepath):
    """rhwp fields: 누름틀 이름·현재값·안내문(변경 없음)."""
    r = _rhwp_or_exit(rhwp_cli.run, ["fields", os.path.abspath(filepath)], required_fields=("fields",))
    fields = r.envelope["fields"]
    if not fields:
        print("누름틀이 없습니다.")
    for f in fields:
        line = f"  {_one_line(f.get('name'))} = {_one_line(f.get('value'))!r}"
        if f.get("guide"):
            line += f"  (안내: {_one_line(f['guide'])})"
        print(line)
    print(f"누름틀 {len(fields)}개")
    security = (r.envelope.get("textSecurity") or {}).get("status")
    if security and security != "clean":
        print(f"경고: 누름틀 이름에 숨은 문자·닮은 문자 신호가 있습니다(textSecurity={security}). "
              "rhwp fields <파일> --json으로 확인하세요.", file=sys.stderr)


def cmd_fill_fields(filepath, data, output=None):
    """rhwp edit fill-fields. 없는·모호한·닮은 이름이 있으면 쓰지 않는다(exit 0이어도 거짓 성공 금지)."""
    src = os.path.abspath(filepath)
    output = _rhwp_edit_output(filepath, output)
    args = ["edit", "fill-fields", src, "--data", data]
    keys = ("filledCount", "notFound", "ambiguous")
    plan = _rhwp_or_exit(rhwp_cli.run, args + ["--dry-run"], required_fields=keys)

    def refused(env):
        return env["notFound"] or env["ambiguous"] or env.get("confusable") or not env["filledCount"]

    if refused(plan.envelope):
        env = plan.envelope
        print(f"오류: 누름틀을 채우지 않았습니다. 없는 이름 {env['notFound']}, 모호한 이름 {env['ambiguous']}, "
              f"닮은 이름 {env.get('confusable') or []}. 이름은 --fields로 확인하고, 같은 이름은 "
              "\"이름[k]\"로 지목한다.", file=sys.stderr)
        sys.exit(1)
    report = {}
    r = _rhwp_or_exit(rhwp_cli.run, args + ["-o", rhwp_cli.OUTPUT], output=output,
                      required_fields=keys + ("output",), verify=_gated(src, output, report))
    if refused(r.envelope) or r.envelope["filledCount"] != plan.envelope["filledCount"]:
        root, ext = os.path.splitext(output)
        os.replace(output, f"{root}.rejected{ext}")
        print(f"오류: 쓰기 결과가 사전 점검과 다릅니다. 산출물을 {root}.rejected{ext}로 옮겼습니다.",
              file=sys.stderr)
        sys.exit(1)
    print(f"누름틀 {r.envelope['filledCount']}개 채움 → {output}")
    _print_gate(report)


def cmd_grid_set_cell(filepath, table_idx, row_idx, col_idx, text, output=None):
    """rhwp edit set-cell(격자 좌표). 중첩 표·병합 칸 오류는 rhwp 메시지를 그대로 전한다."""
    src = os.path.abspath(filepath)
    output = _rhwp_edit_output(filepath, output)
    report = {}
    try:
        r = rhwp_cli.run(["edit", "set-cell", src, "--table", str(table_idx), "--row", str(row_idx),
                          "--col", str(col_idx), "--text", text, "-o", rhwp_cli.OUTPUT],
                         output=output, required_fields=("oldText", "newText", "output"),
                         verify=_gated(src, output, report))
    except rhwp_cli.RhwpError as exc:
        print(f"오류: {exc}\n  {GRID_HINT}", file=sys.stderr)
        sys.exit(1)
    old = _one_line(rhwp_cli.untrusted(r.envelope, "oldText"))
    print(f"[{table_idx},{row_idx},{col_idx}] {old!r} → {r.envelope['newText']!r} 저장: {output}")
    if r.envelope.get("overflow"):
        print("경고: 값이 칸 폭을 넘칩니다. --to-pdf로 렌더를 확인하세요.", file=sys.stderr)
    _print_gate(report)


def _embedded_counts(path):
    """cross_check가 세지 않는 차트·OLE·수식 수와 Chart/ 파트 수."""
    import hwpx_gate
    parts = hwpx_gate.read_parts(path)
    counts = Counter()
    for b in hwpx_gate.xml_parts(parts).values():
        census = hwpx_gate.tag_census(b)
        for tag in ("chart", "ole", "equation"):
            counts[tag] += census.get(tag, 0)
    counts["Chart/ 파트"] = sum(1 for n in parts if n.startswith("Chart/"))
    return counts


def cmd_to_hwp(filepath, output=None):
    """배포용 HWPX → HWP 내보내기. 정본은 HWPX이며 이 경로로 편집하지 않는다.
    산출 .hwp를 JAR로 다시 HWPX로 바꿔 원본과 대조한다(JDK 없으면 unverified)."""
    import hwpx_gate

    src = os.path.abspath(filepath)
    output = os.path.abspath(output or get_output_path(filepath, ".hwp"))
    if not output.lower().endswith(".hwp"):
        print(f"오류: --to-hwp 출력은 .hwp여야 합니다: {output}", file=sys.stderr)
        sys.exit(1)
    state = {"status": "unverified", "messages": []}

    def reject(staged, reason):
        rejected = _keep_rejected(staged, output)
        raise rhwp_cli.RhwpError(f"JAR 왕복 대조 실패: {reason}\n거부된 산출물 보존: {rejected}\n  {TO_HWP_COM_HINT}")

    def verify(staged):
        if not find_java():
            state["messages"].append("JDK 없음: JAR 왕복 대조 생략(unverified)")
            return
        with tempfile.TemporaryDirectory(prefix="hwp-back-") as td:
            back = os.path.join(td, "back.hwpx")
            try:
                _run_jar(staged, back)  # JDK가 있는데 못 읽으면 독립 파서가 거부한 HWP다
                rep = hwpx_gate.cross_check(src, back)
                before, after = _embedded_counts(src), _embedded_counts(back)
            except Exception as exc:
                reject(staged, f"{type(exc).__name__}: {exc}")
        lost = [f"{k} {before[k]}→{after[k]}" for k in before if after[k] < before[k]]
        if rep.status == "failed" or lost:
            reject(staged, " / ".join(([rep.summary()] if rep.status == "failed" else []) + lost))
        state["status"] = rep.status
        state["messages"].append(f"JAR 왕복 대조: {rep.summary()}")

    try:
        rhwp_cli.run(["convert", src, rhwp_cli.OUTPUT, "--verify", "--verify-pages"], output=output,
                     required_fields=("output", "format", "verify"), verify=verify)
    except rhwp_cli.RhwpError as exc:
        print(f"오류: HWP 내보내기 실패: {exc}\n  {TO_HWP_COM_HINT}", file=sys.stderr)
        sys.exit(1)
    for m in state["messages"]:
        print(m)
    print(f"HWP 내보내기: {output}")
    print("  배포용 사본이다. 정본은 HWPX이며 이후 편집은 원본 HWPX에서 한다.")
    print(f"engine=rhwp status={state['status']}")


def cmd_remove_text(filepath, search_text, output=None):
    """특정 텍스트를 포함하는 <hp:t> 요소를 완전히 제거"""
    root, all_files, section_path = open_hwpx(filepath)
    removed_count = 0

    for t_elem in root.findall('.//hp:t', NS):
        if t_elem.text and search_text in t_elem.text:
            parent = t_elem.getparent()
            if parent is not None:
                parent.remove(t_elem)
                removed_count += 1

    if removed_count == 0:
        print(f"'{search_text}'를 포함하는 텍스트 요소를 찾지 못했습니다.")
        return

    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"'{search_text}' 포함 {removed_count}개 텍스트 요소 제거 완료")


def parse_cell_ref(cell_ref):
    """'표,행,셀' 형식의 문자열을 파싱"""
    parts = cell_ref.split(',')
    if len(parts) != 3:
        print(f"오류: 셀 참조 형식이 잘못되었습니다. '표,행,셀' 형식 사용 (예: 0,1,0)", file=sys.stderr)
        sys.exit(1)
    try:
        return int(parts[0]), int(parts[1]), int(parts[2])
    except ValueError:
        print(f"오류: 셀 참조는 숫자여야 합니다.", file=sys.stderr)
        sys.exit(1)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # Windows cp949 안전 (cmd_check_env와 동일 패턴)
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(
        description='HWPX 파일 편집 유틸리티',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  %(prog)s --check-env                    # 기능 계층별 의존성·런타임 점검
  %(prog)s doc.hwp --from-hwp            # HWP → HWPX (rhwp, JAR 교차 대조)
  %(prog)s doc.hwp --from-hwp --engine jar
  %(prog)s doc.hwpx --info
  %(prog)s doc.hwpx --to-md
  %(prog)s doc.hwpx --to-md -o output.md
  %(prog)s doc.hwpx --find "이전" --replace "이후"
  %(prog)s doc.hwpx --set-cell 0,1,0 "새 텍스트"
  %(prog)s doc.hwpx --split-cell 0,1,0
  %(prog)s doc.hwpx --tables              # rhwp 격자 좌표 확인
  %(prog)s doc.hwpx --grid-set-cell 0,4,3 "홍길동"
  %(prog)s doc.hwpx --fields
  %(prog)s doc.hwpx --fill-fields '{"성명":"홍길동"}'
  %(prog)s doc.hwpx --to-hwp              # 배포용 HWP 사본
  %(prog)s doc.hwpx --delete-after "<서식 2>"
  %(prog)s doc.hwpx --delete-empty-rows 1
  %(prog)s doc.hwpx --trim-cell 0,1,0
  %(prog)s doc.hwpx --delete-rows 1 11,12
  %(prog)s doc.hwpx --remove-text "2027. 2. 28."
  %(prog)s doc.hwpx --sanitize
  %(prog)s doc.hwpx --fix-empty-cells     # Pandoc 변환 후 한글 크래시 방지
  %(prog)s doc.hwpx --list-squeeze        # "한 줄로 입력" 과압축 문단 나열
  %(prog)s doc.hwpx --fix-squeeze         # 과압축 문단을 자연 줄바꿈으로 전환
  %(prog)s --diagnose-com                 # 한컴 COM 자동화 진단
  %(prog)s doc.hwpx --to-pdf              # PDF 저장 (rhwp)
  %(prog)s doc.hwpx --to-pdf --pdf-backend com   # Windows 한컴 COM PDF
  %(prog)s doc.hwpx --set-cell 0,1,0 "텍스트" -o output.hwpx
        """)

    parser.add_argument('file', nargs='?',
                        help='HWPX 파일 경로 (.hwp를 주면 먼저 HWPX로 변환해 그 위에서 동작)')
    parser.add_argument('--from-hwp', action='store_true',
                        help='HWP → HWPX 변환만 수행 (기본 rhwp + JAR 교차 대조, -o로 출력 경로)')
    parser.add_argument('--engine', choices=HWP_ENGINES, default='rhwp',
                        help='HWP → HWPX 변환 엔진 (기본 rhwp, jar는 명시 폴백). '
                             '--from-hwp와 .hwp 입력의 자동 변환에 적용')
    parser.add_argument('--info', action='store_true', help='표 구조 정보 출력')
    parser.add_argument('--to-md', action='store_true',
                        help='HWPX → Markdown 변환 (XML 직접 파싱, API 불필요)')
    parser.add_argument('--to-pdf', action='store_true',
                        help='HWPX/HWP → PDF 저장 (기본 rhwp, 모든 OS. '
                             '글꼴에 없는 문자가 빈 네모로 찍히면 exit 2)')
    parser.add_argument('--pdf-backend', choices=PDF_BACKENDS, default='rhwp',
                        help='--to-pdf 렌더러 (기본 rhwp, com은 Windows 한컴오피스 명시 폴백)')
    parser.add_argument('--check-env', action='store_true',
                        help='기능 계층(tier)별 의존성·런타임 준비 상태 점검 (첫 실행·의존성 의심 시, 읽기 전용)')
    parser.add_argument('--diagnose-com', action='store_true',
                        help='한컴오피스 COM 자동화 사용 가능 여부 진단')
    parser.add_argument('--password',
                        help='--to-pdf에서 암호화된 HWPX/HWP를 열 때 사용할 비밀번호')
    parser.add_argument('--find', help='찾을 텍스트')
    parser.add_argument('--replace', help='바꿀 텍스트')
    parser.add_argument('--set-cell', nargs=2, metavar=('TABLE,ROW,COL', 'TEXT'),
                        help='표 셀에 텍스트 입력 (예: --set-cell 0,1,0 "텍스트")')
    parser.add_argument('--tables', action='store_true',
                        help='rhwp 격자 좌표로 표 보기 (--grid-set-cell 좌표 확인용, 변경 없음)')
    parser.add_argument('--fields', action='store_true', help='누름틀 이름·현재값 보기 (rhwp, 변경 없음)')
    parser.add_argument('--fill-fields', metavar='JSON|@파일',
                        help='누름틀 채우기 (rhwp, 예: \'{"성명":"홍길동"}\'). 없는·모호한 이름이 있으면 쓰지 않음')
    parser.add_argument('--grid-set-cell', nargs=2, metavar=('TABLE,ROW,COL', 'TEXT'),
                        help='rhwp 격자 좌표(--tables)로 셀 입력. 원본 보존이 중요한 공문은 --set-cell이 1순위')
    parser.add_argument('--to-hwp', action='store_true',
                        help='배포용 HWPX → HWP 내보내기 (rhwp, JAR 왕복 대조). 정본은 HWPX')
    parser.add_argument('--split-cell', metavar='TABLE,ROW,COL',
                        help='병합 셀 분할 (예: --split-cell 0,1,0)')
    parser.add_argument('--delete-after', metavar='TEXT',
                        help='해당 텍스트 포함 요소부터 이후 전부 삭제')
    parser.add_argument('--delete-empty-rows', metavar='TABLE', type=int,
                        help='테이블 끝의 빈 행 삭제 (예: --delete-empty-rows 1)')
    parser.add_argument('--trim-cell', metavar='TABLE,ROW,COL',
                        help='셀 텍스트의 첫 줄만 유지 (예: --trim-cell 0,1,0)')
    parser.add_argument('--delete-rows', nargs=2, metavar=('TABLE', 'ROW_INDICES'),
                        help='테이블에서 특정 행 삭제 (예: --delete-rows 1 11,12)')
    parser.add_argument('--remove-text', metavar='TEXT',
                        help='해당 텍스트를 포함하는 <hp:t> 요소를 완전히 제거')
    parser.add_argument('--sanitize', action='store_true',
                        help='header.xml의 문단 배경색 등 알려진 렌더링 문제 자동 수정')
    parser.add_argument('--fix-empty-cells', action='store_true',
                        help='표의 빈 셀에 기본 문단 삽입 (Pandoc/hwpx_convert.py 변환 후 한글 크래시 방지 필수)')
    parser.add_argument('--add-preview', action='store_true',
                        help='누락된 Preview/PrvText.txt를 section 텍스트로 생성·주입 '
                             '(hwp2hwpx 변환물의 hwpx-validate 실패 보정. section 무변형)')
    parser.add_argument('--list-squeeze', action='store_true',
                        help='"한 줄로 입력"(lineWrap=SQUEEZE)으로 과압축되는 문단 나열 (변경 없음)')
    parser.add_argument('--fix-squeeze', action='store_true',
                        help='과압축되는 "한 줄로 입력" 문단을 자연 줄바꿈(BREAK)으로 전환 '
                             '(양식에 긴 제목을 채워 글자가 겹쳐 뭉개질 때)')
    parser.add_argument('--cell-br', action='store_true',
                        help='--to-md와 함께 사용: 표 셀 내부 문단을 <br>로 구분 '
                             '(긴 지문이 셀 안에 있는 고사지·보고서에 권장)')
    parser.add_argument('--merge-fill', action='store_true',
                        help='--to-md와 함께 사용: 병합으로 덮인 칸을 시작 칸 값으로 채움 '
                             '(행 단위 파싱·LLM 입력용; 기본은 GFM 정렬 보존)')
    parser.add_argument('-o', '--output',
                        help='출력 파일 경로 (기본: _work-hwpx-automation/ 폴더). '
                             '최종 결과물을 작업 폴더에 바로 두려면 -o로 지정')

    args = parser.parse_args()

    if args.check_env:
        sys.exit(cmd_check_env())

    if args.diagnose_com:
        sys.exit(cmd_diagnose_com())

    if etree is None:
        print(
            "오류: lxml이 설치되어 있지 않습니다.\n"
            "  설치: pip install -r requirements.txt  (python-hwpx, lxml, hwpx-tomd)\n"
            "  전체 환경 점검: python hwpx_edit.py --check-env",
            file=sys.stderr,
        )
        sys.exit(1)

    if not args.file:
        parser.print_help()
        sys.exit(1)

    if not os.path.exists(args.file):
        print(f"오류: 파일을 찾을 수 없습니다: {args.file}", file=sys.stderr)
        sys.exit(1)

    # --to-hwp는 다른 동작과 섞지 않는다(.hwp 출력이 편집 경로로 새지 않게)
    if args.to_hwp:
        others = [a for a in ('info', 'to_md', 'to_pdf', 'tables', 'fields', 'fill_fields',
                              'grid_set_cell', 'find', 'set_cell', 'split_cell', 'delete_after',
                              'delete_rows', 'trim_cell', 'remove_text', 'sanitize', 'fix_empty_cells',
                              'add_preview', 'list_squeeze', 'fix_squeeze', 'from_hwp')
                  if getattr(args, a) not in (None, False)]
        if others or args.delete_empty_rows is not None:
            print("오류: --to-hwp는 다른 동작과 함께 쓸 수 없습니다.", file=sys.stderr)
            sys.exit(1)
        if args.file.lower().endswith('.hwp'):
            print("오류: --to-hwp는 HWPX 입력만 받습니다(이미 .hwp입니다).", file=sys.stderr)
            sys.exit(1)

    if args.output and args.output.lower().endswith('.hwp') and not args.to_hwp:
        print("오류: 결과물은 HWPX로 저장합니다. -o에 .hwp를 줄 수 없습니다.", file=sys.stderr)
        sys.exit(1)

    # -o 의 부모 폴더가 없으면 저장 단계(임시 파일 생성)에서 FileNotFoundError로 죽는다
    if args.output:
        os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)

    if args.from_hwp:
        if not args.file.lower().endswith('.hwp'):
            print("오류: --from-hwp는 .hwp 입력만 받습니다.", file=sys.stderr)
            sys.exit(1)
        convert_hwp_or_exit(args.file, args.output, args.engine)
        return

    # 정본 형식은 HWPX다. .hwp는 입구에서 변환하고 그 HWPX 위에서 동작한다.
    # PDF는 렌더링만 하므로 원본을 그대로 넘긴다.
    # 편집 결과는 변환물과 다른 이름(<이름>_편집.hwpx)에 둔다. 같은 자리에 두면 같은
    # .hwp로 다시 부를 때 입구 변환이 편집 결과를 원본 변환물로 덮는다.
    if args.file.lower().endswith('.hwp') and not args.to_pdf:
        args.file = convert_hwp_or_exit(args.file, engine=args.engine).path
        writes_hwpx = not (args.info or args.to_md or args.list_squeeze or args.tables or args.fields)
        if writes_hwpx and not args.output:
            args.output = os.path.splitext(args.file)[0] + "_편집.hwpx"

    if args.info:
        cmd_info(args.file)
    elif args.to_md:
        cmd_to_md(args.file, args.output, cell_br=args.cell_br, merge_fill=args.merge_fill)
    elif args.to_pdf:
        cmd_to_pdf(args.file, args.output, password=args.password, backend=args.pdf_backend)
    elif args.tables:
        cmd_tables(args.file)
        return
    elif args.fields:
        cmd_fields(args.file)
        return
    elif args.fill_fields is not None:
        cmd_fill_fields(args.file, args.fill_fields, args.output)
        return
    elif args.grid_set_cell:
        t, r, c = parse_cell_ref(args.grid_set_cell[0])
        cmd_grid_set_cell(args.file, t, r, c, args.grid_set_cell[1], args.output)
        return
    elif args.to_hwp:
        cmd_to_hwp(args.file, args.output)
        return
    elif args.find is not None and args.replace is not None:
        cmd_find_replace(args.file, args.find, args.replace, args.output)
    elif args.set_cell:
        t, r, c = parse_cell_ref(args.set_cell[0])
        cmd_set_cell(args.file, t, r, c, args.set_cell[1], args.output)
    elif args.split_cell:
        t, r, c = parse_cell_ref(args.split_cell)
        cmd_split_cell(args.file, t, r, c, args.output)
    elif args.delete_after:
        cmd_delete_after(args.file, args.delete_after, args.output)
    elif args.delete_empty_rows is not None:
        cmd_delete_empty_rows(args.file, args.delete_empty_rows, args.output)
    elif args.trim_cell:
        t, r, c = parse_cell_ref(args.trim_cell)
        cmd_trim_cell_text(args.file, t, r, c, args.output)
    elif args.delete_rows:
        table_idx = int(args.delete_rows[0])
        row_indices = [int(x) for x in args.delete_rows[1].split(',')]
        cmd_delete_rows(args.file, table_idx, row_indices, args.output)
    elif args.remove_text:
        cmd_remove_text(args.file, args.remove_text, args.output)
    elif args.sanitize:
        cmd_sanitize(args.file, args.output)
    elif args.fix_empty_cells:
        cmd_fix_empty_cells(args.file, args.output)
    elif args.list_squeeze:
        cmd_list_squeeze(args.file)
    elif args.fix_squeeze:
        cmd_fix_squeeze(args.file, args.output)
    elif args.add_preview:
        cmd_add_preview(args.file, args.output)
    else:
        parser.print_help()
        sys.exit(1)

    if not (args.info or args.to_md or args.to_pdf or args.list_squeeze):
        print("engine=lxml")


if __name__ == '__main__':
    main()
