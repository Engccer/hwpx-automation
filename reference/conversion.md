# MD → HWPX 변환 및 스타일 후처리

## hwpx_convert.py 기본 변환

### 사전 요구사항

```bash
pip install pypandoc-hwpx   # pypandoc + HWPX 변환 지원
```

- 미설치 시 `ModuleNotFoundError: No module named 'pypandoc_hwpx'` 발생
- 내부적으로 Pandoc 사용 (시스템에 Pandoc 설치 필요)

### 기본 사용법

```bash
python <스킬디렉토리>/convert/hwpx_convert.py <입력.md> -o <출력.hwpx>
```

### 알려진 제한사항

| 요소 | 변환 결과 | 해결 방법 |
|------|----------|----------|
| 제목 (`#`, `##`, `###`, `####`) | 정상 변환 (스타일 ID 자동 매핑) | - |
| 표 (마크다운 테이블) | 정상 변환 (borderFill=3, 테두리 있음) | - |
| 볼드/이탤릭 | 정상 변환 | - |
| 목록 (`-`, `1.`) | 정상 변환 (paraPrIDRef에 들여쓰기 적용) | - |
| **빈 표 셀 (`\| \|`)** | **`<hp:p>` 누락 → 한글 크래시** | `fix_empty_cells()` 필수 (아래 참조) |
| **blockquote (`>`)** | **내용 전체 누락** | 전처리 필요 (아래 참조) |
| **따옴표 안 텍스트 (`"…"`, `'…'`, `"…"`, `'…'`)** | **내용 전체 누락 (스마트·ASCII 따옴표 모두 해당)** | PUA 마커 전처리+후처리 필요 (아래 참조) |
| 각주 (`[^N]`) | 정상 변환 (페이지 하단 각주) | 일부 각주가 문서 끝 미주로 남을 수 있음 |
| 수평선 (`---`) | 변환되지 않음 | - |
| 인라인 코드, 코드 블록 | 일반 텍스트로 변환 | - |

### Blockquote 전처리 패턴

hwpx_convert.py가 `>` blockquote 내용을 완전히 누락시키므로, 변환 전 마커로 치환해야 한다.

```python
# -*- coding: utf-8 -*-
"""MD 전처리: blockquote를 마커 텍스트로 변환"""

QUOTE_MARKER = "\u3010\uc778\uc6a9\u3011"  # 【인용】

def preprocess_md(input_path, output_path):
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    modified = []
    for line in lines:
        if line.startswith('> '):
            modified.append(QUOTE_MARKER + line[2:])
        else:
            modified.append(line)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(modified)
```

변환 후 후처리 스크립트에서 `【인용】` 마커를 찾아 스타일을 적용하고 마커를 제거한다.

### 따옴표 전처리/후처리 패턴

Pandoc의 HWPX writer는 모든 종류의 따옴표(`"…"`, `'…'`, `"…"`, `'…'`) 안의 텍스트를 통째로 누락시킨다. 한국어 문서에서 대화체·인용·강조에 따옴표가 빈번하게 사용되므로 **blockquote 누락만큼 심각한 데이터 손실**이다.

**해결**: 전처리에서 따옴표를 유니코드 사설 영역(PUA) 마커로 치환 → Pandoc 변환 → 후처리에서 마커를 스마트 따옴표로 복원.

```python
import re

# PUA 마커 정의 (U+FFF0~FFF3, Pandoc이 건드리지 않음)
QUOT_OPEN   = '\uFFF0'  # 여는 큰따옴표 마커
QUOT_CLOSE  = '\uFFF1'  # 닫는 큰따옴표 마커
SQUOT_OPEN  = '\uFFF2'  # 여는 작은따옴표 마커
SQUOT_CLOSE = '\uFFF3'  # 닫는 작은따옴표 마커

def preprocess_quotes(text):
    """변환 전: 따옴표 → PUA 마커 치환"""
    # 스마트 따옴표
    text = text.replace('\u201c', QUOT_OPEN).replace('\u201d', QUOT_CLOSE)
    text = text.replace('\u2018', SQUOT_OPEN).replace('\u2019', SQUOT_CLOSE)
    # ASCII 큰따옴표 (쌍 단위, 내용 있는 경우만)
    text = re.sub(r'"([^"\n]+?)"', lambda m: QUOT_OPEN + m.group(1) + QUOT_CLOSE, text)
    # ASCII 작은따옴표 (쌍 단위, 내용 있는 경우만 — 1글자도 포함)
    text = re.sub(r"'([^'\n]+?)'", lambda m: SQUOT_OPEN + m.group(1) + SQUOT_CLOSE, text)
    return text

def postprocess_quotes(xml_text):
    """변환 후: PUA 마커 → 스마트 따옴표 복원 (section0.xml에 적용)"""
    xml_text = xml_text.replace(QUOT_OPEN,   '\u201c')  # \u201c = "
    xml_text = xml_text.replace(QUOT_CLOSE,  '\u201d')  # \u201d = "
    xml_text = xml_text.replace(SQUOT_OPEN,  '\u2018')  # \u2018 = '
    xml_text = xml_text.replace(SQUOT_CLOSE, '\u2019')  # \u2019 = '
    return xml_text
```

**주의 — `re.sub` replacement string의 역참조 함정**:

```python
# ❌ 잘못된 예시 — \1이 그룹 역참조가 아닌 SOH 제어문자(U+0001)로 해석됨
re.sub(r'"([^"]+)"', '\u300e\\1\u300f', text)

# ✅ 올바른 예시 — lambda로 그룹 참조
re.sub(r'"([^"]+)"', lambda m: '\u300e' + m.group(1) + '\u300f', text)
```

비-raw 문자열에서 `\uXXXX`가 유니코드 문자로 먼저 해석된 후 `\\1`이 `\1`이 되는데, `re.sub`이 이를 octal escape(U+0001)로 처리한다. **반드시 lambda 사용.**

### HWPX 변환 결과 텍스트 검증 시 주의사항

**`hwpx_edit.py --to-md`의 맹점**: 추출 도구 자체도 따옴표 안 텍스트를 누락시킬 수 있다. `--to-md` 출력만으로 검증하면 따옴표 누락이 검증에서도 걸러지지 않는다.

**`<hp:t>` join 방식으로 검증**: Pandoc은 특수 문자 경계에서 텍스트를 별도 `<hp:run>`으로 분리하므로, 단순 substring 검색이 실패할 수 있다. 정확한 검증은 모든 `<hp:t>` 텍스트를 join:

```python
import zipfile, re

def extract_all_text(hwpx_path):
    """section0.xml의 모든 <hp:t> 텍스트를 join하여 반환 (검증용)"""
    with zipfile.ZipFile(hwpx_path, 'r') as z:
        section = z.read('Contents/section0.xml').decode('utf-8')
    return ''.join(re.findall(r'<hp:t>(.*?)</hp:t>', section))

# 사용 예
full_text = extract_all_text('output.hwpx')
assert '본' in full_text       # 1글자 따옴표 내 텍스트 확인
assert '어떻게' in full_text
```

## 스타일 후처리 — header.xml 수정

변환된 HWPX의 `Contents/header.xml`에는 기본 스타일만 포함되어 있다.
표 배경색, 굵은 글씨, 이탤릭 등을 추가하려면 `borderFill`과 `charPr` 항목을 추가해야 한다.

### borderFill 추가 (표 배경색)

```python
import zipfile, re

def add_border_fills(header_xml):
    """header.xml에 배경색 borderFill 항목 추가"""
    m = re.search(r'<hh:borderFills itemCnt="(\d+)">', header_xml)
    old_cnt = int(m.group(1))
    header_xml = header_xml.replace(
        f'<hh:borderFills itemCnt="{old_cnt}">',
        f'<hh:borderFills itemCnt="{old_cnt + 2}">'
    )

    # 테두리 + 배경색 borderFill 템플릿
    def make_border_fill(bf_id, face_color):
        return f'''<hh:borderFill id="{bf_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
        <hh:slash type="NONE" Crooked="0" isCounter="0" />
        <hh:backSlash type="NONE" Crooked="0" isCounter="0" />
        <hh:leftBorder type="SOLID" width="0.12 mm" color="#000000" />
        <hh:rightBorder type="SOLID" width="0.12 mm" color="#000000" />
        <hh:topBorder type="SOLID" width="0.12 mm" color="#000000" />
        <hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000" />
        <hh:diagonal type="SOLID" width="0.1 mm" color="#000000" />
        <hc:fillBrush>
          <hc:winBrush faceColor="{face_color}" hatchColor="#000000" alpha="0" />
        </hc:fillBrush>
    </hh:borderFill>'''

    header_bf = make_border_fill(old_cnt + 1, "#E6E6E6")  # 표 헤더 행
    summary_bf = make_border_fill(old_cnt + 2, "#F2F2F2")  # 합계/요약 행

    header_xml = header_xml.replace(
        '</hh:borderFills>',
        header_bf + '\n' + summary_bf + '\n</hh:borderFills>'
    )
    return header_xml, old_cnt + 1, old_cnt + 2  # header_bf_id, summary_bf_id
```

### charPr 추가 (굵은/이탤릭 글자)

```python
def add_char_property(header_xml, height=1100, bold=False, italic=False,
                      text_color="#000000"):
    """header.xml에 새 charPr 추가. 반환: (수정된 xml, 새 charPr ID)"""
    m = re.search(r'<hh:charProperties itemCnt="(\d+)">', header_xml)
    old_cnt = int(m.group(1))
    header_xml = header_xml.replace(
        f'<hh:charProperties itemCnt="{old_cnt}">',
        f'<hh:charProperties itemCnt="{old_cnt + 1}">'
    )

    attrs = f'id="{old_cnt}" height="{height}" textColor="{text_color}"'
    attrs += ' shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE"'
    attrs += ' borderFillIDRef="2"'
    if bold:
        attrs += ' bold="1"'
    if italic:
        attrs += ' italic="1"'

    new_char = f'''<hh:charPr {attrs}>
        <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0" />
        <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100" />
        <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0" />
        <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100" />
        <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0" />
        <hh:underline type="NONE" shape="SOLID" color="#000000" />
        <hh:strikeout shape="NONE" color="#000000" />
        <hh:outline type="NONE" />
        <hh:shadow type="NONE" color="#C0C0C0" offsetX="5" offsetY="5" />
      </hh:charPr>'''

    header_xml = header_xml.replace(
        '</hh:charProperties>',
        new_char + '\n</hh:charProperties>'
    )
    return header_xml, old_cnt
```

### 본문 글꼴 크기 변경

```python
def adjust_font_size(header_xml, char_pr_id=0, new_height=1100):
    """charPr의 height 변경 (1100 = 11pt, 1000 = 10pt)"""
    header_xml = header_xml.replace(
        f'<hh:charPr id="{char_pr_id}" height="1000"',
        f'<hh:charPr id="{char_pr_id}" height="{new_height}"'
    )
    return header_xml
```

## 스타일 후처리 — section0.xml 수정

### 장 제목 박스 (표 박스 래퍼) 삽입

장 제목 등 단락을 1행 1열 배경색 표로 감싸는 패턴.

**⚠️ 핵심 주의사항 (2026-04 실측 확인)**

1. **`<hp:ctrl>` 래퍼 금지**: 표를 단락 내부에 삽입할 때 `<hp:run><hp:ctrl><hp:tbl>…</hp:tbl></hp:ctrl></hp:run>` 구조는 XSD 스키마는 통과하지만 한컴 COM의 `Open()`에서 RPC 크래시(`-2147023170 — 원격 프로시저 호출 실패`) 발생. Pandoc 원본 표 구조와 동일하게 **`<hp:run>` 직계**에 `<hp:tbl>`을 두어야 함.

   ```xml
   <!-- ❌ 크래시 발생 -->
   <hp:p><hp:run><hp:ctrl><hp:tbl …>…</hp:tbl></hp:ctrl></hp:run></hp:p>

   <!-- ✅ 올바른 구조 (Pandoc 원본과 동일) -->
   <hp:p><hp:run><hp:tbl …>…</hp:tbl></hp:run></hp:p>
   ```

2. **`<hp:subList id="">` 빈 id 금지**: 셀 내부 subList의 `id` 속성은 고유한 숫자(정수 문자열)를 반드시 부여해야 함. 빈 문자열은 COM Open 크래시의 원인.

3. **Pandoc 출력의 heading styleIDRef 매핑** (기본값, 2026-04 실측):
   - `#` (h1) → styleIDRef="2"
   - `##` (h2) → styleIDRef="3"
   - `###` (h3) → styleIDRef="4"
   - `####` (h4) → styleIDRef="5"

   **주의**: 중집위 회의자료 등 다른 템플릿에서는 `##`→4로 매핑되기도 함. 새 문서 작업 시 실제 분포를 먼저 확인할 것:

   ```bash
   python -c "import zipfile, re; s = zipfile.ZipFile('doc.hwpx').read('Contents/section0.xml').decode(); from collections import Counter; print(Counter(re.findall(r'styleIDRef=\"(\d+)\"', s)))"
   ```

**장 제목 박스 템플릿 (검증 완료)**:

```python
CHAPTER_BOX_TEMPLATE = (
    '<hp:p paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
    '<hp:run charPrIDRef="0">'
    '<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" '
    'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" '
    'rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="{bf_id}" noAdjust="0">'
    '<hp:sz width="45000" widthRelTo="ABSOLUTE" height="1200" heightRelTo="ABSOLUTE" protect="0"/>'
    '<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
    'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" '
    'horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
    '<hp:outMargin left="0" right="0" top="200" bottom="600"/>'
    '<hp:inMargin left="510" right="510" top="200" bottom="200"/>'
    '<hp:tr>'
    '<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" '
    'borderFillIDRef="{bf_id}">'
    '<hp:subList id="{sub_id}" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" '
    'linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
    '<hp:p paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
    '<hp:run charPrIDRef="{char_id}"><hp:t>{title}</hp:t></hp:run>'
    '</hp:p>'
    '</hp:subList>'
    '<hp:cellAddr colAddr="0" rowAddr="0"/>'
    '<hp:cellSpan colSpan="1" rowSpan="1"/>'
    '<hp:cellSz width="45000" height="1200"/>'
    '<hp:cellMargin left="510" right="510" top="200" bottom="200"/>'
    '</hp:tc></hp:tr></hp:tbl></hp:run></hp:p>'
)


def wrap_chapter_headings(section_xml, chapter_bf_id, chapter_char_id,
                          target_style_id="3"):
    """대상 heading 문단을 배경색 표 박스로 치환. target_style_id는 문서에 맞춰 지정 (## 은 보통 3)."""
    pattern = re.compile(
        rf'<hp:p([^>]*styleIDRef="{target_style_id}"[^>]*)>(.*?)</hp:p>',
        re.DOTALL,
    )
    counter = {'n': 0}

    def esc(s):
        return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    def repl(m):
        texts = re.findall(r'<hp:t[^>]*>([^<]*)</hp:t>', m.group(2))
        title = ''.join(texts).strip()
        if not title:
            return m.group(0)
        counter['n'] += 1
        return CHAPTER_BOX_TEMPLATE.format(
            tbl_id=f"99000{counter['n']:03d}",
            sub_id=f"98000{counter['n']:03d}",
            bf_id=chapter_bf_id,
            char_id=chapter_char_id,
            title=esc(title),
        )

    return pattern.sub(repl, section_xml)
```

검증 경로: `hwpx-validate` 통과 + 한컴 COM `Open()` 통과 + 한글 수동 오픈 정상. 2026-04 장교조 교육감 후보자 정책제안서 5종(MD→HWPX) 재작업에서 확인.

### 표 헤더/합계 행 스타일링

```python
def style_table_rows(section_xml, header_bf_id, summary_bf_id, bold_char_id,
                     summary_table_indices=None):
    """표의 첫 행(헤더)에 배경색+굵은 글씨, 특정 표의 마지막 행(합계)에 배경색+굵은 글씨 적용.

    Args:
        summary_table_indices: 합계 행이 있는 표의 인덱스 집합 (예: {0, 1, 3})
    """
    if summary_table_indices is None:
        summary_table_indices = set()

    tables = list(re.finditer(r'<hp:tbl\b[^>]*>', section_xml))

    for tbl_idx, tbl_match in enumerate(tables):
        tbl_start = tbl_match.start()
        tbl_end = section_xml.find('</hp:tbl>', tbl_start) + len('</hp:tbl>')
        tbl_xml = section_xml[tbl_start:tbl_end]

        rows = list(re.finditer(r'<hp:tr\b', tbl_xml))
        if not rows:
            continue

        # 헤더 행 (첫 번째 행)
        r0_start = rows[0].start()
        r0_end = rows[1].start() if len(rows) > 1 else tbl_xml.find('</hp:tr>', r0_start) + 7
        header_row = tbl_xml[r0_start:r0_end]
        styled = re.sub(r'borderFillIDRef="3"', f'borderFillIDRef="{header_bf_id}"', header_row)
        styled = re.sub(r'charPrIDRef="0"', f'charPrIDRef="{bold_char_id}"', styled)
        tbl_xml = tbl_xml[:r0_start] + styled + tbl_xml[r0_end:]

        # 합계 행 (마지막 행)
        if tbl_idx in summary_table_indices:
            rows = list(re.finditer(r'<hp:tr\b', tbl_xml))  # 재계산
            rn_start = rows[-1].start()
            rn_end = tbl_xml.find('</hp:tr>', rn_start) + 7
            last_row = tbl_xml[rn_start:rn_end]
            styled = re.sub(r'borderFillIDRef="3"', f'borderFillIDRef="{summary_bf_id}"', last_row)
            styled = re.sub(r'charPrIDRef="0"', f'charPrIDRef="{bold_char_id}"', styled)
            tbl_xml = tbl_xml[:rn_start] + styled + tbl_xml[rn_end:]

        section_xml = section_xml[:tbl_start] + tbl_xml + section_xml[tbl_end:]
        # 이후 표 위치 재계산
        tables = list(re.finditer(r'<hp:tbl\b[^>]*>', section_xml))

    return section_xml
```

### 표 열 너비 조정 (가독성)

**문제**: Pandoc HWPX 기본 변환은 모든 표의 열을 **균등 너비**로 생성한다. 내용 길이와 무관하게 일률 너비가 적용되어, 짧은 헤더(예: `순위`, `응답률`)와 긴 헤더(예: `개선 과제`)가 같은 폭을 차지하며 여백만 남고 내용이 눈에 들어오지 않는다(2026-04-20 소통실 피드백).

**해결**: 표별로 열 너비 비율을 배정하는 후처리 함수를 적용한다. 총 너비(기본 45000 HWPUNIT)는 유지하고, `colAddr` 기준으로 각 셀의 `<hp:cellSz width="..."/>`를 덮어쓴다.

```python
import re

TABLE_TOTAL_W = 45000  # hp:sz width 기본값 (A4 본문폭)

def apply_column_widths(section_xml: str, widths_by_table: dict,
                        skip_single_col: bool = True) -> str:
    """표 열 너비를 내용 기반 비율로 조정.

    Args:
        widths_by_table: {content_table_index: [col0_width, col1_width, ...]}
                         content_table_index는 1열짜리 장 제목 박스를 제외한
                         실제 내용 표(colCnt > 1)의 0-based 순번.
                         각 너비 리스트의 합은 TABLE_TOTAL_W(45000)가 되어야 함.
        skip_single_col: 1열 박스 표(장 제목 등)는 건드리지 않음 (기본 True).

    예시:
        apply_column_widths(section, {
            0: [13500, 31500],          # 항목/내용 30:70
            1: [31500, 13500],          # 지표/수치 70:30
            2: [6000, 30000, 9000],     # 순위/과제/응답률 13:67:20
        })
    """
    content_idx = 0
    out, cursor = [], 0
    for m in re.finditer(r'<hp:tbl\b[^>]*colCnt="(\d+)"[^>]*>', section_xml):
        col_cnt = int(m.group(1))
        tbl_start = m.start()
        tbl_end = section_xml.find('</hp:tbl>', tbl_start) + len('</hp:tbl>')
        out.append(section_xml[cursor:tbl_start])
        tbl = section_xml[tbl_start:tbl_end]

        if skip_single_col and col_cnt == 1:
            out.append(tbl)
        elif col_cnt > 1:
            widths = widths_by_table.get(content_idx)
            content_idx += 1
            if widths and len(widths) == col_cnt:
                assert sum(widths) == TABLE_TOTAL_W, (
                    f"표 {content_idx - 1} 너비 합계 {sum(widths)} != {TABLE_TOTAL_W}")
                def _repl(cell_m):
                    cell = cell_m.group(0)
                    ca = re.search(r'<hp:cellAddr colAddr="(\d+)"', cell)
                    if not ca:
                        return cell
                    col = int(ca.group(1))
                    return re.sub(
                        r'(<hp:cellSz width=")\d+(")',
                        rf'\g<1>{widths[col]}\g<2>',
                        cell, count=1,
                    )
                tbl = re.sub(r'<hp:tc\b.*?</hp:tc>', _repl, tbl, flags=re.DOTALL)
            out.append(tbl)
        else:
            out.append(tbl)
        cursor = tbl_end
    out.append(section_xml[cursor:])
    return ''.join(out)
```

**너비 배정 가이드라인** (내용 비례):

| 열 유형 | 권장 비율 | 예 (총 45000 기준) |
|:---|:---:|:---|
| 순위·번호 (1~2자리) | 12~15% | 6000 |
| 짧은 라벨 (지표·항목·주제 등) | 25~35% | 13500 |
| 수치·퍼센트 | 18~25% | 9000 ~ 11000 |
| 긴 설명·과제명 | 60~70% | 30000 ~ 31500 |
| 2열 표 (라벨/값) | 30:70 또는 70:30 | 13500/31500 |
| 3열 표 (순위/과제/수치) | 13:67:20 내외 | 6000/30000/9000 |

**원칙**:
1. **열의 합은 `TABLE_TOTAL_W`(45000)**. 분모 유지 안 하면 `<hp:sz>` 와 불일치해 한글이 재계산.
2. **모든 행의 cellSz를 같은 colAddr끼리 일치시킬 것**. 일부 행만 바꾸면 한글이 첫 행 기준으로 표시해 나머지 행이 찌그러짐.
3. **1열 박스 표(장 제목)는 건드리지 않음**. `skip_single_col=True` 기본값 유지.
4. **검증**: 적용 후 `hwpx-validate` 통과 + COM `Open/SaveAs` 왕복 확인.

**적용 사례**: 2026-04-20 장교조 근로지원인 설문조사 붙임 요약 HWPX (표 4개, 2·3열 혼재). 조사 개요 30:70, 8대 지표 70:30, TOP 5 13:67:20, 주제별 언급 70:30 적용 후 가독성 확인.

### 인용문 (blockquote) 스타일링

```python
QUOTE_MARKER = "\u3010\uc778\uc6a9\u3011"  # 【인용】

def style_blockquotes(section_xml, quote_para_id, italic_char_id):
    """【인용】 마커가 있는 문단에 좌측 여백 + 이탤릭 적용 후 마커 제거"""
    styled_count = 0
    start_search = 0
    while True:
        idx = section_xml.find(QUOTE_MARKER, start_search)
        if idx < 0:
            break

        p_start = section_xml.rfind('<hp:p ', 0, idx)
        p_end = section_xml.find('</hp:p>', idx) + len('</hp:p>')
        if p_start < 0 or p_end < len('</hp:p>'):
            start_search = idx + 1
            continue

        para = section_xml[p_start:p_end]
        new_para = para.replace(QUOTE_MARKER, '')
        new_para = re.sub(r'paraPrIDRef="\d+"', f'paraPrIDRef="{quote_para_id}"', new_para, count=1)
        new_para = re.sub(r'charPrIDRef="\d+"', f'charPrIDRef="{italic_char_id}"', new_para)

        section_xml = section_xml[:p_start] + new_para + section_xml[p_end:]
        styled_count += 1
        start_search = p_start + len(new_para)

    return section_xml, styled_count
```

인용문용 paraPr (좌측 여백) 추가:

```python
def add_quote_para_property(header_xml):
    """좌측 여백이 있는 paraPr 추가 (인용문용)"""
    m = re.search(r'<hh:paraProperties itemCnt="(\d+)">', header_xml)
    old_cnt = int(m.group(1))
    header_xml = header_xml.replace(
        f'<hh:paraProperties itemCnt="{old_cnt}">',
        f'<hh:paraProperties itemCnt="{old_cnt + 1}">'
    )

    # paraPr 0을 복제하여 좌측 여백 추가
    m0 = re.search(r'<hh:paraPr id="0".*?</hh:paraPr>', header_xml, re.DOTALL)
    quote_para = m0.group(0)
    quote_para = re.sub(r'id="0"', f'id="{old_cnt}"', quote_para, count=1)
    quote_para = quote_para.replace(
        '<hc:intent value="0" unit="HWPUNIT" />',
        '<hc:intent value="0" unit="HWPUNIT" />\n'
        '              <hc:left value="2000" unit="HWPUNIT" />'
    )
    header_xml = header_xml.replace(
        '</hh:paraProperties>',
        quote_para + '\n</hh:paraProperties>'
    )
    return header_xml, old_cnt
```

## 전체 변환 워크플로우 예시

```python
# -*- coding: utf-8 -*-
"""MD → 스타일 적용 HWPX 변환 전체 흐름"""
import zipfile, re, os, subprocess, sys

INPUT_MD = "report.md"
OUTPUT_HWPX = "report.hwpx"
TEMP_MD = "_temp_for_hwpx.md"
CONVERTER = "C:/Users/pc/Windows-Projects/tools/hwpx-automation/convert/hwpx_convert.py"

# 1. 전처리: 따옴표 마커 치환 + blockquote 마커 삽입
# (preprocess_md 안에서 preprocess_quotes 호출 필수)
preprocess_md(INPUT_MD, TEMP_MD)

# 2. 기본 변환
subprocess.run([sys.executable, CONVERTER, TEMP_MD, "-o", OUTPUT_HWPX], check=True)
os.remove(TEMP_MD)

# 3. 표 스타일 후처리 (header.xml + section0.xml)
with zipfile.ZipFile(OUTPUT_HWPX, 'r') as z:
    files = {n: z.read(n) for n in z.namelist() if not n.endswith('/')}
    file_list = [n for n in z.namelist()]

header_xml = files['Contents/header.xml'].decode('utf-8')
section_xml = files['Contents/section0.xml'].decode('utf-8')

header_xml, hdr_bf, sum_bf = add_border_fills(header_xml)
header_xml, bold_id = add_char_property(header_xml, height=1000, bold=True)
header_xml, italic_id = add_char_property(header_xml, height=1100, italic=True, text_color="#555555")
header_xml, quote_para_id = add_quote_para_property(header_xml)
header_xml = adjust_font_size(header_xml, char_pr_id=0, new_height=1100)

section_xml = style_table_rows(section_xml, hdr_bf, sum_bf, bold_id, {0, 1, 3})
section_xml, cnt = style_blockquotes(section_xml, quote_para_id, italic_id)
# 따옴표 마커 → 스마트 따옴표 복원
section_xml = postprocess_quotes(section_xml)

section_xml = re.sub(r'<hp:linesegarray>.*?</hp:linesegarray>', '', section_xml, flags=re.DOTALL)

# 빈 셀 수정 (필수! 빈 셀이 있으면 한글 크래시)
from lxml import etree
section_root = etree.fromstring(section_xml.encode('utf-8'))
fix_empty_cells(section_root)
section_xml = etree.tostring(section_root, encoding='unicode')

files['Contents/header.xml'] = header_xml.encode('utf-8')
files['Contents/section0.xml'] = section_xml.encode('utf-8')

with zipfile.ZipFile(OUTPUT_HWPX, 'w', zipfile.ZIP_DEFLATED) as z:
    for name in file_list:
        if name.endswith('/'):
            continue
        ct = zipfile.ZIP_STORED if name == 'mimetype' else zipfile.ZIP_DEFLATED
        z.writestr(name, files[name], compress_type=ct)

# 4. python-hwpx API로 머리글/바닥글 추가
from hwpx.document import HwpxDocument
doc = HwpxDocument.open(OUTPUT_HWPX)
doc.set_header_text("보고서 제목", page_type="BOTH")
doc.set_footer_text("- \u0002 -", page_type="BOTH")  # \u0002 = 자동 페이지 번호
doc.save_to_path(OUTPUT_HWPX)

# 5. 검증
subprocess.run(["hwpx-validate", OUTPUT_HWPX], check=True)
```

## 필수 후처리: 빈 셀 수정

hwpx_convert.py는 마크다운 표의 빈 셀(`| |`)을 `<hp:subList>` 안에 `<hp:p>` 없이 생성한다.
한글은 모든 셀에 최소 하나의 `<hp:p>`가 있어야 하므로, 빈 셀이 하나라도 있으면 **파일이 열리다가 바로 닫히는 크래시**가 발생한다.

**변환 후 반드시 이 수정을 적용해야 한다.**

```python
from lxml import etree

HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph'

def fix_empty_cells(section_root):
    """빈 셀의 <hp:subList>에 빈 <hp:p> 추가. 한글 크래시 방지."""
    fixed = 0
    for sublist in section_root.iter(f'{{{HP}}}subList'):
        paras = sublist.findall(f'{{{HP}}}p')
        if len(paras) == 0:
            p = etree.SubElement(sublist, f'{{{HP}}}p')
            p.set('paraPrIDRef', '0')
            p.set('styleIDRef', '0')
            p.set('pageBreak', '0')
            p.set('columnBreak', '0')
            p.set('merged', '0')
            run = etree.SubElement(p, f'{{{HP}}}run')
            run.set('charPrIDRef', '0')
            t = etree.SubElement(run, f'{{{HP}}}t')
            t.text = ''
            fixed += 1
    return fixed
```

**증상**: 파일이 열리다가 즉시 닫힘 (오류 메시지 없음)
**원인**: `<hp:subList>` 내부에 `<hp:p>` 자식 요소가 0개
**해결**: 빈 `<hp:p><hp:run><hp:t></hp:t></hp:run></hp:p>` 삽입

## 주의사항

### 한국어 텍스트 검색 시 인코딩 함정

Windows 터미널(cp949)에서 `python -c "..."` 로 한국어 문자열을 전달하면 cp949로 인코딩된다.
HWPX 내부 XML은 UTF-8이므로 바이트 불일치로 검색 실패.

**해결 방법**:
1. Python 스크립트를 `.py` 파일로 작성 (UTF-8 저장) + `# -*- coding: utf-8 -*-`
2. 또는 `python -X utf8 -c "..."` 플래그 사용
3. 또는 유니코드 이스케이프 사용: `"\uc778\uc6a9"` (= "인용")

### header.xml 스타일 구조

```
header.xml
├── fontfaces     — 글꼴 정의 (맑은 고딕 등)
├── borderFills   — 테두리 + 배경색 정의 (id로 참조)
├── charProperties — 글자 속성 (크기, 색상, 굵기, 기울임)
├── paraProperties — 문단 속성 (여백, 들여쓰기, 줄간격)
└── styles         — 스타일 정의 (paraPr + charPr 조합)
```

- `borderFillIDRef`: `<hp:tc>` 셀의 테두리/배경 참조
- `charPrIDRef`: `<hp:run>` 텍스트의 글자 속성 참조
- `paraPrIDRef`: `<hp:p>` 문단의 문단 속성 참조

### hwpx_convert.py 생성 기본값

| 요소 | 기본값 |
|------|--------|
| 본문 글꼴 | 맑은 고딕 10pt (charPr id=0, height=1000) |
| 표 셀 borderFill | id=3 (실선 테두리, 배경 없음) |
| 제목 스타일 | Heading 1~9 자동 매핑 (## → Heading 2 등) |
| 용지 | A4, 좌우 72mm, 상 42.55mm, 하 49.6mm |
| 각주 | 문서 끝 텍스트로 변환 (HWPX 각주 요소 아님) |
| **줄간격** | **180%** (`lineSpacing type="PERCENT" value="180"`) — 160% 등으로 변경 시 후처리 필수 |

### Pandoc HWPX lineSpacing 주의사항

Pandoc HWPX writer가 생성하는 `lineSpacing` 태그는 **네임스페이스 접두사가 없다**:
- Pandoc 출력: `lineSpacing type="PERCENT" value="180" unit="HWPUNIT"`
- 한컴 COM 출력: `<hc:lineSpacing type="percent" value="160"/>`

줄간격을 후처리로 변경할 때 `hc:lineSpacing` 패턴이 아닌 `lineSpacing` 패턴으로 검색해야 한다:

```python
# ✅ 올바른 패턴 (네임스페이스 무관)
header_xml = re.sub(r'(lineSpacing[^/]*?)value="180"', r'\1value="160"', header_xml)

# ❌ 잘못된 패턴 (Pandoc HWPX에서 매칭 안 됨)
header_xml = re.sub(r'(<hc:lineSpacing[^/]*?)value="\d+"', r'\1value="160"', header_xml)
```

## 주의사항

1. **hwpx_convert.py blockquote 누락**: `>` blockquote 내용이 변환 시 완전히 누락됨. 반드시 전처리로 마커 치환 필요
2. **hwpx_convert.py 따옴표 안 텍스트 누락**: `"…"`, `'…'`, `"…"`, `'…'` 안의 텍스트가 통째로 누락됨. PUA 마커 전처리+후처리 필수 (위 "따옴표 전처리/후처리 패턴" 참조)
3. **hwpx_convert.py 각주**: `[^N]` 각주가 HWPX 각주 요소가 아닌 문서 끝 일반 텍스트로 변환됨
4. **한국어 검색 인코딩**: Windows 터미널(cp949)에서 `python -c` 한국어 → HWPX UTF-8 불일치. `.py` 파일 또는 유니코드 이스케이프 사용
5. **`re.sub` 역참조 함정**: 비-raw 문자열에서 `'\uXXXX\\1\uYYYY'` 형태의 replacement를 사용하면 `\1`이 그룹 역참조가 아닌 SOH 제어문자(U+0001)로 해석됨. 반드시 `lambda m: ... + m.group(1) + ...` 사용
6. **변환 결과 검증**: `hwpx_edit.py --to-md`도 따옴표 안 텍스트를 누락시킬 수 있으므로 검증에 부적합. `<hp:t>` 텍스트를 join하여 검증 (위 "HWPX 변환 결과 텍스트 검증" 참조)
7. **변환 스크립트 보존**: MD→HWPX 변환 시 생성한 Python 스크립트는 삭제하지 않고 프로젝트 폴더에 보관한다 (향후 재변환·수정 용도). 사용자가 명시적으로 삭제를 요청한 경우에만 삭제
8. **`pypandoc.convert_file()` 입력 경로 대괄호 함정**: pypandoc은 내부에서 `glob.glob(str(source))`를 `glob.escape` 없이 호출하므로, 파일명/경로의 `[`, `]`, `*`, `?`가 glob 문자 클래스로 오인되어 매칭 실패 → `WindowsPath` fallback → `TypeError: 'WindowsPath' object is not iterable`. 한글과 무관한 라이브러리 자체 버그. 대표 재현 파일명 패턴: `[붙임]`, `[공고]`, `[수정]`, `2. [안건]` 등.
   **해결**: MD→HWPX 변환 시 입력 MD를 항상 `tempfile.TemporaryDirectory()` 안에 ASCII-safe 이름(`input.md`)으로 쓴 뒤 `hwpx_convert.py`에 넘긴다. 출력 경로는 대괄호가 있어도 안전 (`--output=`은 glob 대상이 아님). (2026-04-20 검증)
