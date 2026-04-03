# MD → HWPX 생성 (Build-from-scratch 방식)

python-hwpx API로 문서 구조를 처음부터 빌드하고, XML 후처리로 세밀한 스타일링을 적용하는 방식.
Pandoc 방식(`hwpx_convert.py`)의 한계를 완전히 해결한다.

## Pandoc 방식 대비 장점

| 항목 | Pandoc 방식 | Build-from-scratch |
|------|------------|-------------------|
| Blockquote | 내용 누락 (마커 전처리 필요) | 좌측 컬러바 + 배경색 직접 적용 |
| 각주 `[^N]` | 문서 끝 일반 텍스트 | 위첨자 숫자 + 별도 각주 섹션 |
| 표 스타일 | 헤더/합계 배경색만 후처리 | 헤더/교대행/합계행/볼드셀/기본폰트 모두 제어 |
| 글꼴 | 변환기 기본값 (맑은 고딕 10pt) | fontface 추가, 크기/색상/굵기 완전 제어 |
| 인라인 마크다운 | pandoc이 처리 | `parse_inline()`으로 직접 처리 필요 |
| 표지/머리글/바닥글 | API로 추가 가능 | API로 추가 가능 |

## 5단계 워크플로우

### 1단계: MD 파싱

마크다운을 블록 리스트로 변환. 블록 타입: `heading`, `paragraph`, `table`, `blockquote`, `list`, `note`, `hr`

```python
def parse_markdown(md):
    """마크다운 → 블록 리스트."""
    lines = md.split("\n")
    blocks = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1; continue
        # heading
        m = re.match(r'^(#{1,4})\s+(.+)', line)
        if m:
            blocks.append({"type": "heading", "level": len(m.group(1)), "text": m.group(2)})
            i += 1; continue
        # table (| ... | 행이 연속)
        if line.strip().startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(c.replace("-","").replace(":","").strip() == "" for c in cells):
                    rows.append(cells)
                i += 1
            blocks.append({"type": "table", "rows": rows})
            continue
        # blockquote
        if line.startswith("> "):
            bq_lines = []
            while i < len(lines) and (lines[i].startswith("> ") or lines[i].startswith(">")):
                bq_lines.append(lines[i].lstrip("> ").rstrip())
                i += 1
            blocks.append({"type": "blockquote", "lines": bq_lines})
            continue
        # list
        m = re.match(r'^(\s*)[•\-\*]\s+(.+)', line)
        if m:
            items = []
            while i < len(lines) and re.match(r'^\s*[•\-\*]\s+', lines[i]):
                items.append({"level": 0, "text": re.sub(r'^\s*[•\-\*]\s+', '', lines[i])})
                i += 1
            blocks.append({"type": "list", "items": items})
            continue
        # paragraph (기본)
        blocks.append({"type": "paragraph", "text": line})
        i += 1
    return blocks
```

인라인 마크다운 처리:

```python
def parse_inline(text):
    """인라인 마크다운 → (text, bold, footnote, code) 튜플 리스트."""
    parts = []
    regex = re.compile(r"(\*\*[^*]+\*\*)|(\[\^\d+\])|(`[^`]+`)")
    last = 0
    for m in regex.finditer(text):
        if m.start() > last:
            parts.append({"text": text[last:m.start()]})
        if m.group(1):
            parts.append({"text": m.group(1)[2:-2], "bold": True})
        elif m.group(2):
            parts.append({"text": m.group(2), "footnote": True})
        elif m.group(3):
            parts.append({"text": m.group(3)[1:-1], "code": True})
        last = m.end()
    if last < len(text):
        parts.append({"text": text[last:]})
    return parts if parts else [{"text": text}]
```

### 2단계: 스타일 주입 (header.xml)

빈 HWPX 템플릿(`HwpxDocument.new()`)의 `header.xml`에 커스텀 스타일을 추가.

**빈 템플릿 기본 카운트**:
- `charPr`: ID 0-6 (7개)
- `borderFill`: ID 1-2 (2개)
- `paraPr`: ID 0-19 (20개)
- `fontface`: 각 언어별 ID 0(함초롬돋움), 1(함초롬바탕) — 2개

#### 폰트 추가 (맑은 고딕)

빈 템플릿은 함초롬돋움/바탕만 포함. 맑은 고딕 등을 사용하려면 fontface에 추가:

```python
MALGUN = "맑은 고딕"
malgun_font_xml = (
    f'<hh:font id="2" face="{MALGUN}" type="TTF" isEmbedded="0">'
    '<hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4" '
    'contrast="0" strokeVariation="1" armStyle="1" letterform="1" '
    'midline="1" xHeight="1"/></hh:font>'
)
# 7개 언어 모두에 추가 (HANGUL, LATIN, HANJA, JAPANESE, OTHER, SYMBOL, USER)
for lang in ["HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"]:
    header_xml = header_xml.replace(
        f'<hh:fontface lang="{lang}" fontCnt="2">',
        f'<hh:fontface lang="{lang}" fontCnt="3">'
    )
    # 해당 lang 블록의 </hh:fontface> 직전에 삽입
    idx = header_xml.find(f'<hh:fontface lang="{lang}" fontCnt="3">')
    close_idx = header_xml.find("</hh:fontface>", idx)
    header_xml = header_xml[:close_idx] + malgun_font_xml + header_xml[close_idx:]
```

이후 charPr의 `fontRef`에서 `hangul="2" latin="2" ...`로 참조.

#### borderFill 추가

```python
def make_border_fill(bf_id, face_color="#FFFFFF", left_border=None, bottom_border=None):
    lb = f'type="SOLID" width="{left_border[0]}" color="{left_border[1]}"' if left_border else 'type="SOLID" width="0.12 mm" color="#CCCCCC"'
    bb = f'type="SOLID" width="{bottom_border[0]}" color="{bottom_border[1]}"' if bottom_border else 'type="SOLID" width="0.12 mm" color="#CCCCCC"'
    return f"""<hh:borderFill id="{bf_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
      <hh:slash type="NONE" Crooked="0" isCounter="0"/>
      <hh:backSlash type="NONE" Crooked="0" isCounter="0"/>
      <hh:leftBorder {lb}/>
      <hh:rightBorder type="SOLID" width="0.12 mm" color="#CCCCCC"/>
      <hh:topBorder type="SOLID" width="0.12 mm" color="#CCCCCC"/>
      <hh:bottomBorder {bb}/>
      <hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>
      <hc:fillBrush>
        <hc:winBrush faceColor="{face_color}" hatchColor="#000000" alpha="0"/>
      </hc:fillBrush>
    </hh:borderFill>"""
```

일반적인 borderFill 세트:
- **표 헤더**: 진한 배경 (`#1F4E79`)
- **교대 행**: 연한 배경 (`#F2F7FB`)
- **인용문 배경**: 좌측 컬러바 + 배경색 (`left_border=("0.4 mm", "#4472C4")`, `face_color="#F5F5F5"`)
- **H2 하단선**: 하단 보더 (`bottom_border=("0.3 mm", "#1F4E79")`)

#### charPr 추가

```python
def make_char_pr(pr_id, height=1100, text_color="#000000", bold=False, italic=False,
                 border_fill_ref=2, superscript=False):
    attrs = f'id="{pr_id}" height="{height}" textColor="{text_color}"'
    attrs += f' shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE"'
    attrs += f' borderFillIDRef="{border_fill_ref}"'
    if bold: attrs += ' bold="1"'
    if italic: attrs += ' italic="1"'
    if superscript: attrs += ' supscript="SUPERSCRIPT"'
    return f"""<hh:charPr {attrs}>
      <hh:fontRef hangul="2" latin="2" hanja="2" japanese="2" other="2" symbol="2" user="2"/>
      ... (ratio/spacing/relSz/offset/underline/strikeout/outline/shadow)
    </hh:charPr>"""
```

**주의**: `fontRef`의 숫자는 fontface ID. 맑은 고딕을 ID=2로 추가했으면 `hangul="2"`.

#### paraPr 추가

```python
def make_para_pr(pr_id, left_indent=0, line_spacing=160, space_before=0,
                 space_after=0, align="JUSTIFY", border_fill_ref=2):
    return f"""<hh:paraPr id="{pr_id}" align="{align}" ...>
      <hh:lineSpacing type="PERCENT" value="{line_spacing}" unit="HWPUNIT"/>
      <hh:margin>
        <hc:left value="{left_indent}" unit="HWPUNIT"/>
        <hc:prev value="{space_before}" unit="HWPUNIT"/>
        <hc:next value="{space_after}" unit="HWPUNIT"/>
      </hh:margin>
      <hh:border borderFillIDRef="{border_fill_ref}">...</hh:border>
    </hh:paraPr>"""
```

### 3단계: python-hwpx API로 빌드

```python
doc = HwpxDocument.new()
# ... 스타일 주입 (2단계) ...

# 표지
doc.add_paragraph("보고서 제목", para_pr_id_ref=S["pp_cover_title"],
                  char_pr_id_ref=S["cp_cover_title"])

# 본문
for block in blocks:
    if block["type"] == "paragraph":
        p = doc.add_paragraph("", para_pr_id_ref=S["pp_body"], char_pr_id_ref=S["cp_body"])
        add_inline_runs(p, block["text"], S, S["cp_body"])
    elif block["type"] == "table":
        tbl = doc.add_table(nrows, ncols, width=TABLE_WIDTH,
                           border_fill_id_ref=S["bf_table_cell"],
                           char_pr_id_ref=S["cp_tbl_body"],
                           para_pr_id_ref=S["pp_table"])
        for ri, row in enumerate(rows):
            for ci, cell_text in enumerate(row):
                # 마크다운 기호 제거 후 삽입
                clean = re.sub(r"\*\*([^*]+)\*\*", r"\1", cell_text)
                clean = re.sub(r"\[\^\d+\]", "", clean)
                tbl.set_cell_text(ri, ci, clean.strip())
    # ... heading, blockquote, list 등

# 머리글/바닥글
doc.set_header_text("보고서 제목", page_type="BOTH")
doc.set_footer_text("- \u0002 -", page_type="BOTH")

doc.save_to_path(output_path)
```

**인라인 run 추가 함수**:

```python
def add_inline_runs(para, text, S, default_cp):
    parts = parse_inline(text)
    for part in parts:
        if part.get("footnote"):
            num = re.sub(r"[\[\]^]", "", part["text"])  # [^1] → 1
            para.add_run(num, char_pr_id_ref=S["cp_footnote_ref"])
        elif part.get("bold"):
            para.add_run(part["text"], char_pr_id_ref=S["cp_bold"])
        else:
            para.add_run(part["text"], char_pr_id_ref=default_cp)
```

### 4단계: XML 후처리

python-hwpx API만으로는 표 헤더/교대행/볼드셀 스타일링이 불가능하므로 XML 후처리 필요.

#### 표 스타일링 (`style_tables_xml`)

```python
def style_tables_xml(section_xml, S, summary_indices):
    tables = list(re.finditer(r"<hp:tbl\b[^>]*>", section_xml))
    tbl_idx = 0
    while tbl_idx < len(tables):  # ★ for 루프 사용 금지! while 필수
        tbl_match = tables[tbl_idx]
        tbl_start = tbl_match.start()
        tbl_end = section_xml.find("</hp:tbl>", tbl_start) + len("</hp:tbl>")
        tbl_xml = section_xml[tbl_start:tbl_end]

        # 1. 헤더 행: borderFill → bf_table_header, charPr → cp_tbl_header
        # 2. 데이터 행: charPrIDRef="0" → cp_tbl_body (기본 폰트 교체)
        # 3. 교대 행: borderFill → bf_alt_row
        # 4. 합계 행: borderFill → bf_summary_row, charPr → cp_tbl_bold
        # 5. 볼드 셀: 특정 셀의 charPr → cp_tbl_bold

        section_xml = section_xml[:tbl_start] + tbl_xml + section_xml[tbl_end:]
        tables = list(re.finditer(r"<hp:tbl\b[^>]*>", section_xml))  # ★ 매번 재계산
        tbl_idx += 1
    return section_xml
```

#### lineseg 제거 + 빈 셀 수정

```python
section_xml = re.sub(r"<hp:linesegarray>.*?</hp:linesegarray>", "", section_xml, flags=re.DOTALL)
# 빈 셀에 최소 <hp:p> 추가 (fix_empty_cells)
```

### 5단계: 저장

```python
doc.save_to_path(output_path)
# 또는 ZIP 수동 리패키징 (스타일 주입 후):
#   mimetype → ZIP_STORED (첫 항목), 나머지 → ZIP_DEFLATED
```

## precompute_styles 패턴

스타일 ID를 빌드 시점에 미리 계산하여, `inject_styles`와 `build_body`에서 동일한 ID를 사용:

```python
def precompute_styles():
    S = {}
    # borderFill
    bf_base = 2  # 빈 템플릿 기본 borderFill 수
    S["bf_table_header"] = bf_base + 1  # 3
    S["bf_alt_row"] = bf_base + 2       # 4
    S["bf_quote_bg"] = bf_base + 3      # 5
    # ...

    # charPr
    cp = 7  # 빈 템플릿 기본 charPr 수
    for name in ["cp_body", "cp_h1", "cp_h2", "cp_h3", "cp_h4",
                  "cp_tbl_header", "cp_tbl_body", "cp_tbl_bold", ...]:
        S[name] = cp
        cp += 1

    # paraPr
    pp = 20  # 빈 템플릿 기본 paraPr 수
    for name in ["pp_body", "pp_h1", "pp_h2", ...]:
        S[name] = pp
        pp += 1
    return S
```

**핵심**: `inject_styles`에서 추가하는 순서와 `precompute_styles`의 열거 순서가 반드시 일치해야 함.

## 디자인 토큰 (DOCX 스크립트와 공유 가능)

```python
COLOR_PRIMARY = "#1F4E79"     # 진파랑 (제목, 표 헤더)
COLOR_SECONDARY = "#2E75B6"   # 중파랑 (H3, 각주 번호)
COLOR_ALT_ROW = "#F2F7FB"     # 연파랑 (교대 행)
COLOR_QUOTE_BG = "#F5F5F5"    # 연회색 (인용문 배경)
COLOR_GRAY = "#666666"        # 회색 (인용문 출처)
COLOR_WHITE = "#FFFFFF"
COLOR_BLACK = "#000000"
COLOR_DARK = "#333333"        # 진회색 (H4)

# HWPX height 단위: 1pt = 100 hwpunit
BODY_HT = 1100   # 11pt
SMALL_HT = 900   # 9pt
H1_HT = 1600     # 16pt
H2_HT = 1400     # 14pt
H3_HT = 1200     # 12pt

# 표 너비 (A4, 양쪽 1인치 여백: 약 159.2mm = 45300 hwpunit)
TABLE_WIDTH = 45300
```

## 주요 함정 (Gotchas)

### 1. lineSpacing 단위

```python
# ❌ 잘못됨 — 16000% 줄간격, 수백 쪽 생성
<hh:lineSpacing type="PERCENT" value="{line_spacing * 100}" unit="HWPUNIT"/>

# ✅ 올바름 — 160% 줄간격
<hh:lineSpacing type="PERCENT" value="{line_spacing}" unit="HWPUNIT"/>
```

HWPML PERCENT 타입은 정수 그대로 사용. `value="160"` = 160%.

### 2. fontRef ID

빈 템플릿 기본 폰트:
- ID 0 = 함초롬돋움
- ID 1 = 함초롬바탕

맑은 고딕을 사용하려면 **fontface에 ID=2로 추가** 후 charPr의 fontRef에서 `hangul="2"` 참조.
fontRef를 `"0"`이나 `"1"`로 두면 함초롬 폰트가 적용됨.

### 3. set_cell_text의 기본 charPr

`tbl.set_cell_text(ri, ci, text)` → 생성되는 run의 `charPrIDRef="0"` (템플릿 기본값).
`add_table(char_pr_id_ref=S["cp_tbl_body"])` 파라미터는 초기 빈 셀에만 적용되고,
`set_cell_text`가 새로 만드는 run에는 적용되지 않음.

**해결**: XML 후처리에서 데이터 행의 `charPrIDRef="0"`을 `cp_tbl_body`로 일괄 교체.

### 4. style_tables_xml 루프 버그

```python
# ❌ 잘못됨 — section_xml 수정 후 이전 위치 사용
for tbl_idx, tbl_match in enumerate(tables):
    # tbl_match.start()가 이전 section_xml 기준이라 위치 어긋남

# ✅ 올바름 — while 루프 + 매번 재계산
tbl_idx = 0
while tbl_idx < len(tables):
    tbl_match = tables[tbl_idx]
    # ... 처리 ...
    tables = list(re.finditer(r"<hp:tbl\b[^>]*>", section_xml))
    tbl_idx += 1
```

### 5. 마크다운 기호가 표 셀에 남는 문제

`set_cell_text`는 원시 텍스트를 삽입하므로 `**미흡**`, `[^3]` 등이 그대로 남음.

**해결**: 삽입 전 정규식으로 제거:
```python
clean = re.sub(r"\*\*([^*]+)\*\*", r"\1", cell_text)  # **bold** → bold
clean = re.sub(r"\[\^\d+\]", "", clean)                # [^1] → 제거
```

볼드가 필요한 셀은 위치를 기록해두고 XML 후처리에서 `charPrIDRef`를 `cp_tbl_bold`로 변경.

### 6. 인용문 출처/제목/노트의 각주 참조

`doc.add_paragraph(text, ...)` 형태로 raw 텍스트를 넣으면 `[^N]`이 그대로 남음.
모든 블록에서 `add_inline_runs()`를 사용하여 인라인 마크다운을 처리해야 함.

### 7. ensure_run_style 호환성

python-hwpx의 `ensure_run_style()` 메서드는 내부적으로 `xml.etree.ElementTree.SubElement`를 사용하지만,
실제 문서 요소는 `lxml.etree._Element`이므로 **TypeError 발생**.
→ charPr/paraPr는 API가 아닌 XML 문자열 주입 방식으로 처리해야 함.

## 실제 적용 사례

DPI 보고서 HWPX 생성 스크립트 (`generate_hwpx.py`):
- 입력: 534행 마크다운 (7장 + 14개 각주, 8개 표)
- 출력: 42KB HWPX (659 단락, 8 표, 22쪽)
- 스타일: 21 charPr, 7 borderFill, 16 paraPr, 맑은 고딕 폰트
- 소요: 스크립트 작성 ~2시간, 이후 재생성 <5초

## 주의사항

1. **build-from-scratch lineSpacing 단위**: HWPML PERCENT 타입은 정수 그대로 사용 (160% = `value="160"`). `* 100` 하면 16000%가 되어 수백 쪽 문서 생성
2. **build-from-scratch fontRef**: 빈 템플릿의 기본 폰트(함초롬돋움/바탕)가 아닌 맑은 고딕 등을 사용하려면 fontface에 새 폰트를 추가하고 charPr의 fontRef ID를 변경해야 함
3. **build-from-scratch set_cell_text**: `set_cell_text()`로 생성된 run은 charPrIDRef="0" (템플릿 기본값)을 사용. XML 후처리에서 해당 셀의 charPrIDRef를 커스텀 ID로 교체 필요
4. **style_tables_xml 루프**: `for` 루프에서 section_xml을 수정하면 이후 표 위치가 달라지므로, 반드시 `while` 루프 + 매 반복 tables 리스트 재계산 사용
