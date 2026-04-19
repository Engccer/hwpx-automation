#!/usr/bin/env python3
"""
hwpx_edit.py - HWPX 파일 편집 유틸리티

사용법:
  python hwpx_edit.py <파일.hwpx> --info                       # 표 구조 출력
  python hwpx_edit.py <파일.hwpx> --to-md                      # HWPX → Markdown 변환 (XML 직접 파싱, API 불필요)
  python hwpx_edit.py <파일.hwpx> --find "이전" --replace "이후"  # 텍스트 치환
  python hwpx_edit.py <파일.hwpx> --set-cell 0,1,0 "텍스트"     # 표 셀에 텍스트 입력
  python hwpx_edit.py <파일.hwpx> --split-cell 0,1,0            # 병합 셀 분할

의존성: python-hwpx, lxml
"""

import argparse
import copy
import os
import sys
import zipfile
from io import BytesIO
from lxml import etree


# HWPX 네임스페이스
NS = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'ha': 'http://www.hancom.co.kr/hwpml/2011/app',
    'config': 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
}


def get_output_path(filepath, ext=None):
    """입력 파일 기준 _output/ 폴더에 출력 경로 생성. ext가 None이면 원래 확장자 유지.
    이미 _output/ 안에 있으면 같은 폴더에 출력 (중첩 방지)."""
    input_dir = os.path.dirname(os.path.abspath(filepath))
    if os.path.basename(input_dir) == '_output':
        output_dir = input_dir
    else:
        output_dir = os.path.join(input_dir, '_output')
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
        section_xml = zf.read(section_path)
        all_files = {}
        for name in zf.namelist():
            all_files[name] = zf.read(name)

    root = etree.fromstring(section_xml)
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
        all_files[header_path] = etree.tostring(
            header_root, xml_declaration=True, encoding='UTF-8')
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
        # 빈 셀 — minimal 문단 삽입
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


def save_hwpx(filepath, root, all_files, section_path, output=None):
    """수정된 section XML을 HWPX 파일에 저장"""
    if output is None:
        output = get_output_path(filepath)

    # 저장 전 빈 셀 자동 보정 (Pandoc 변환물 등의 한글 크래시 방지)
    empty_cells = fix_empty_cells(root)
    if empty_cells > 0:
        print(f"fix_empty_cells: 빈 셀 {empty_cells}개에 기본 문단 삽입")

    all_files[section_path] = etree.tostring(root, xml_declaration=True, encoding='UTF-8')

    # 저장 전 header.xml 자동 sanitize
    sanitize_header(all_files)

    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in all_files.items():
            zf.writestr(name, data)

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


def get_cell_text(cell):
    """셀의 텍스트 내용 추출 (모든 <hp:t> 요소 결합)"""
    texts = []
    for t_elem in cell.findall('.//hp:t', NS):
        if t_elem.text:
            texts.append(t_elem.text)
    return ' '.join(texts)


def get_cell_paragraph_texts(cell):
    """셀 내부의 문단(<hp:p>)별 텍스트 리스트 반환.

    get_cell_text는 모든 <hp:t>를 공백 하나로 합쳐 문단 경계를 잃지만,
    이 함수는 각 문단을 별도 문자열로 유지한다. 긴 지문이 표 셀 안에
    있는 고사 원안·보고서 등에서 문단 구조를 보존할 때 사용한다.
    중첩 표(하위 <hp:tbl>)를 가진 문단은 건너뛴다 (중복 방지).
    """
    texts = []
    sub_list = cell.find('hp:subList', NS)
    if sub_list is None:
        return texts
    for para in sub_list.findall('hp:p', NS):
        if para.find('.//hp:tbl', NS) is not None:
            continue
        parts = []
        for t_elem in para.findall('.//hp:t', NS):
            if t_elem.text:
                parts.append(t_elem.text)
        text = ''.join(parts).strip()
        if text:
            texts.append(text)
    return texts


def get_cell_span(cell):
    """셀의 cellAddr에서 rowSpan, colSpan 값 추출"""
    cell_addr = cell.find('hp:cellAddr', NS)
    if cell_addr is None:
        return 1, 1
    row_span = int(cell_addr.get('rowSpan', '1'))
    col_span = int(cell_addr.get('colSpan', '1'))
    return row_span, col_span


def get_para_direct_text(para):
    """단락에서 직접 텍스트만 추출 (중첩 표 제외)"""
    texts = []
    for run in para.findall('hp:run', NS):
        for t_elem in run.findall('hp:t', NS):
            if t_elem.text:
                texts.append(t_elem.text)
    return ''.join(texts).strip()


def parse_table_to_rows(table, cell_br=False):
    """표를 파싱하여 행별 셀 데이터 리스트 반환. [(텍스트, colSpan), ...]

    cell_br=True이면 셀 내부 문단(<hp:p>)을 <br>로 구분한다.
    False이면 기존 동작(모든 <hp:t>를 공백 하나로 합침)을 유지한다.
    """
    rows_data = []
    for row in get_table_rows(table):
        row_data = []
        for cell in get_row_cells(row):
            if cell_br:
                paras = get_cell_paragraph_texts(cell)
                cell_text = '<br>'.join(
                    p.replace('|', '\\|').replace('\n', ' ') for p in paras
                )
            else:
                cell_text = get_cell_text(cell).replace('\n', ' ').replace('|', '\\|')
            _, col_span = get_cell_span(cell)
            row_data.append((cell_text, col_span))
        rows_data.append(row_data)
    return rows_data


def table_to_markdown(rows_data):
    """파싱된 표 데이터를 Markdown 표 형식으로 변환"""
    if not rows_data:
        return ''

    max_cols = max(sum(cs for _, cs in row) for row in rows_data)
    if max_cols == 0:
        return ''

    lines = []
    for row_idx, row in enumerate(rows_data):
        expanded = []
        for text, col_span in row:
            expanded.append(text)
            for _ in range(col_span - 1):
                expanded.append('')
        while len(expanded) < max_cols:
            expanded.append('')

        line = '| ' + ' | '.join(expanded[:max_cols]) + ' |'
        lines.append(line)

        if row_idx == 0:
            sep = '| ' + ' | '.join(['---'] * max_cols) + ' |'
            lines.append(sep)

    return '\n'.join(lines)


def cmd_to_md(filepath, output=None, cell_br=False):
    """HWPX를 XML 직접 파싱하여 Markdown으로 변환 (API 불필요).

    cell_br=True이면 표 셀 내부 문단을 <br>로 구분한다 (고사지·보고서처럼
    긴 지문이 셀 안에 있는 문서에서 문단 구조 보존).
    """
    if is_encrypted_hwpx(filepath):
        print(f"오류: {ENCRYPTION_HINT}", file=sys.stderr)
        sys.exit(1)

    with zipfile.ZipFile(filepath, 'r') as zf:
        section_files = sorted([
            n for n in zf.namelist()
            if 'section' in n.lower() and n.endswith('.xml') and 'Contents/' in n
        ])
        if not section_files:
            print("오류: section XML 파일을 찾을 수 없습니다.", file=sys.stderr)
            sys.exit(1)

        all_lines = []
        for section_file in section_files:
            root = etree.fromstring(zf.read(section_file))
            for child in root:
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'p':
                    direct_text = get_para_direct_text(child)
                    if direct_text:
                        all_lines.append(direct_text)
                    for tbl in child.findall('.//hp:tbl', NS):
                        rows_data = parse_table_to_rows(tbl, cell_br=cell_br)
                        md_table = table_to_markdown(rows_data)
                        if md_table:
                            all_lines.append('')
                            all_lines.append(md_table)
                            all_lines.append('')

    # 연속 빈 줄 정리
    cleaned = []
    prev_blank = False
    for line in all_lines:
        if line.strip() == '':
            if not prev_blank:
                cleaned.append('')
            prev_blank = True
        else:
            cleaned.append(line)
            prev_blank = False

    md_content = '\n'.join(cleaned)

    if output is None:
        output = get_output_path(filepath, '.md')

    with open(output, 'w', encoding='utf-8') as f:
        f.write(md_content)

    print(f"변환 완료: {output}")
    print(f"크기: {len(md_content)} 글자")


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


def cmd_find_replace(filepath, find_text, replace_text, output=None):
    """텍스트 치환 — python-hwpx(본문 런, 서식 보존) + lxml(표 셀 등 누락분 보완)"""
    from hwpx.document import HwpxDocument

    # 1차: python-hwpx (본문 런 대상, 서식 보존, 런 분할 텍스트도 처리)
    doc = HwpxDocument.open(filepath)
    count_runs = doc.replace_text_in_runs(find_text, replace_text)
    save_path = output if output else get_output_path(filepath)
    doc.save(save_path)

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
    if count_xml > 0:
        save_hwpx(save_path, root, all_files, section_path, output=save_path)

    total = count_runs + count_xml
    print(f"'{find_text}' → '{replace_text}': {total}건 치환 완료 (본문 {count_runs} + 표 {count_xml})")


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
        # 첫 번째 <hp:t>에 텍스트 설정, 나머지 비우기
        t_elems[0].text = text
        for t in t_elems[1:]:
            t.text = ''
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


def cmd_split_cell(filepath, table_idx, row_idx, col_idx, output=None):
    """병합된 셀의 rowSpan/colSpan을 1로 설정하여 분할"""
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
    cell_addr = cell.find('hp:cellAddr', NS)

    if cell_addr is None:
        print("경고: cellAddr 요소가 없습니다.", file=sys.stderr)
        sys.exit(1)

    old_row_span = cell_addr.get('rowSpan', '1')
    old_col_span = cell_addr.get('colSpan', '1')

    cell_addr.set('rowSpan', '1')
    cell_addr.set('colSpan', '1')

    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"표{table_idx} 행{row_idx} 셀{col_idx}: 병합 해제 ({old_row_span}x{old_col_span} → 1x1)")


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


def cmd_delete_empty_rows(filepath, table_idx, output=None):
    """테이블 끝의 빈 행들을 삭제"""
    root, all_files, section_path = open_hwpx(filepath)
    tables = find_tables(root)

    if table_idx >= len(tables):
        print(f"오류: 표 {table_idx}이(가) 없습니다. (총 {len(tables)}개)", file=sys.stderr)
        sys.exit(1)

    table = tables[table_idx]
    rows = get_table_rows(table)
    removed_count = 0

    for i in range(len(rows) - 1, 0, -1):
        row = rows[i]
        is_empty = True
        for t_elem in row.findall('.//hp:t', NS):
            if t_elem.text and t_elem.text.strip():
                is_empty = False
                break
        if is_empty:
            table.remove(row)
            removed_count += 1
        else:
            break

    if removed_count == 0:
        print(f"표 {table_idx}: 빈 행이 없습니다.")
        return

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
    removed = []

    for idx in sorted(row_indices, reverse=True):
        if idx >= len(rows):
            print(f"경고: 행 {idx}이(가) 없습니다. 건너뜁니다.", file=sys.stderr)
            continue
        table.remove(rows[idx])
        removed.append(idx)

    if not removed:
        print("삭제할 행이 없습니다.")
        return

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
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in all_files.items():
            zf.writestr(name, data)
    print(f"저장: {output}")


def cmd_fix_empty_cells(filepath, output=None):
    """표 셀에 빠진 <hp:p>를 채워 한글 크래시 문제 수정 (Pandoc 변환물 후처리용)"""
    root, all_files, section_path = open_hwpx(filepath)
    fixed = fix_empty_cells(root)
    if fixed == 0:
        print("빈 셀이 없습니다 (추가 수정 불필요).")
        return
    # save_hwpx가 다시 fix_empty_cells를 호출하지만, 이미 보정되어 0건 반환 — 문제 없음
    save_hwpx(filepath, root, all_files, section_path, output)
    print(f"빈 셀 {fixed}개에 기본 문단 삽입 완료")


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
    parser = argparse.ArgumentParser(
        description='HWPX 파일 편집 유틸리티',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  %(prog)s doc.hwpx --info
  %(prog)s doc.hwpx --to-md
  %(prog)s doc.hwpx --to-md -o output.md
  %(prog)s doc.hwpx --find "이전" --replace "이후"
  %(prog)s doc.hwpx --set-cell 0,1,0 "새 텍스트"
  %(prog)s doc.hwpx --split-cell 0,1,0
  %(prog)s doc.hwpx --delete-after "<서식 2>"
  %(prog)s doc.hwpx --delete-empty-rows 1
  %(prog)s doc.hwpx --trim-cell 0,1,0
  %(prog)s doc.hwpx --delete-rows 1 11,12
  %(prog)s doc.hwpx --remove-text "2027. 2. 28."
  %(prog)s doc.hwpx --sanitize
  %(prog)s doc.hwpx --fix-empty-cells     # Pandoc 변환 후 한글 크래시 방지
  %(prog)s doc.hwpx --set-cell 0,1,0 "텍스트" -o output.hwpx
        """)

    parser.add_argument('file', help='HWPX 파일 경로')
    parser.add_argument('--info', action='store_true', help='표 구조 정보 출력')
    parser.add_argument('--to-md', action='store_true',
                        help='HWPX → Markdown 변환 (XML 직접 파싱, API 불필요)')
    parser.add_argument('--find', help='찾을 텍스트')
    parser.add_argument('--replace', help='바꿀 텍스트')
    parser.add_argument('--set-cell', nargs=2, metavar=('TABLE,ROW,COL', 'TEXT'),
                        help='표 셀에 텍스트 입력 (예: --set-cell 0,1,0 "텍스트")')
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
    parser.add_argument('--cell-br', action='store_true',
                        help='--to-md와 함께 사용: 표 셀 내부 문단을 <br>로 구분 '
                             '(긴 지문이 셀 안에 있는 고사지·보고서에 권장)')
    parser.add_argument('-o', '--output', help='출력 파일 경로 (기본: _output/ 폴더)')

    args = parser.parse_args()

    if not os.path.exists(args.file):
        print(f"오류: 파일을 찾을 수 없습니다: {args.file}", file=sys.stderr)
        sys.exit(1)

    if args.info:
        cmd_info(args.file)
    elif args.to_md:
        cmd_to_md(args.file, args.output, cell_br=args.cell_br)
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
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == '__main__':
    main()
