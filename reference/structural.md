# 구조적 편집 (행/표/단락 추가) — raw lxml 전용

> python-hwpx API에 `add_row()`, `insert_row()`가 없으므로, 표에 행을 추가/복제하는 경우에만 raw lxml + regex를 사용한다.

## `<hp:linesegarray>` — 텍스트 변경 시 손상의 핵심 원인

각 단락(`<hp:p>`)에 포함된 `<hp:linesegarray>`는 텍스트 줄바꿈 위치를 `textpos` 속성으로 추적함.

```xml
<hp:linesegarray>
  <hp:lineseg textpos="0" vertpos="0" .../>   <!-- 1줄: 0~4자 -->
  <hp:lineseg textpos="5" vertpos="1600" .../> <!-- 2줄: 5~13자 -->
  <hp:lineseg textpos="14" vertpos="3200" .../> <!-- 3줄: 14자~ -->
</hp:linesegarray>
```

**텍스트 길이가 변경되면** `textpos` 불일치 → 한글이 "손상/변조"로 판단. 증상:
- "문서가 손상되었거나 변조되었을 가능성이 있습니다" 경고 → 열기 거부
- "파일이 손상되었습니다" → 열기 불가

**해결**: 텍스트가 변경된 모든 단락에서 `<hp:linesegarray>`를 제거. 한글이 열 때 자동 재계산.

**기술적 근거**: openhwp Rust 소스에서 확인 — `line_segments: Option<LineSegmentArray>` (KS X 6101:2024 표준). `Option`이므로 생략이 스키마적으로 유효. 어떤 공개 프로젝트도 `textpos` 재계산을 구현하지 않으며, "제거 후 한글이 재생성" 접근이 현존하는 최선의 방법.

```python
import re

def strip_lineseg(xml_fragment):
    """변경된 영역의 linesegarray를 제거. 한글이 열 때 자동 재계산함."""
    return re.sub(
        r'<hp:linesegarray>.*?</hp:linesegarray>',
        '', xml_fragment, flags=re.DOTALL
    )
```

**python-hwpx API 사용 시**: lineseg 제거가 자동 처리됨. 내부에서 `_clear_paragraph_layout_cache()` 호출.
**raw lxml/regex 사용 시**: 수동으로 `strip_lineseg()` 호출 필수.

## 필수 체크리스트

1. **`linesegarray` 제거**: 텍스트가 변경된 모든 `<hp:p>`에서 제거
2. **`rowAddr` 순차**: 새 행의 `<hp:cellAddr rowAddr="N"/>`이 중복 없이 순차적. 중복 → 무한 로딩
3. **`rowCnt` 일치**: `<hp:tbl rowCnt="N">`이 실제 `<hp:tr>` 수와 일치. 불일치 → "파일이 손상되었습니다"
4. **표 `id` 고유화**: 새 표의 `id` 속성이 기존과 중복 금지

## 되는 것 vs 안 되는 것

| 작업 | 결과 | 조건 |
|------|------|------|
| 기존 `<hp:t>` 텍스트 변경 (길이 변화) | **성공** | linesegarray 제거 필수 |
| 표에 행 추가 (복제) | **성공** | rowAddr 순차 + rowCnt 일치 + lineseg 제거 |
| 새 단락 `<hp:p>` 추가 (복제) | **성공** | lineseg 제거 |
| 새 표 `<hp:tbl>` 추가 (복제) | **성공** | id 고유화 + rowAddr 순차 + rowCnt 일치 + lineseg 제거 |
| rowCnt와 실제 행 수 불일치 | **실패** | 한글이 검증함 |

## 코드 패턴

### 네임스페이스

```python
NS = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
}
```

### 표에 행 추가

```python
import zipfile, re

def add_table_row(xml, table_id, cell_texts):
    """기존 표의 마지막 행을 복제하여 새 행을 추가하고 텍스트를 설정."""
    tbl_match = re.search(rf'<hp:tbl[^>]*id="{table_id}"[^>]*>', xml)
    tbl_start = tbl_match.start()
    tbl_end = xml.find('</hp:tbl>', tbl_start) + len('</hp:tbl>')
    tbl = xml[tbl_start:tbl_end]

    lte = tbl.rfind('</hp:tr>') + len('</hp:tr>')
    lts = tbl.rfind('<hp:tr>', 0, lte)
    last_row = tbl[lts:lte]

    last_addr = max(int(x) for x in re.findall(r'rowAddr="(\d+)"', last_row))
    new_row = last_row.replace(
        f'rowAddr="{last_addr}"', f'rowAddr="{last_addr + 1}"'
    )

    cells = list(re.finditer(r'<hp:tc[^>]*>.*?</hp:tc>', new_row, re.DOTALL))
    offset = 0
    for i, cm in enumerate(cells):
        if i >= len(cell_texts):
            break
        old = cm.group(0)
        tc = [0]
        def cr(m, txt=cell_texts[i]):
            tc[0] += 1
            return f'<hp:t>{txt}</hp:t>' if tc[0] == 1 else '<hp:t></hp:t>'
        new = re.sub(r'<hp:t>.*?</hp:t>|<hp:t/>', cr, old, flags=re.DOTALL)
        s, e = cm.start() + offset, cm.end() + offset
        new_row = new_row[:s] + new + new_row[e:]
        offset += len(new) - len(old)

    new_row = re.sub(
        r'<hp:linesegarray>.*?</hp:linesegarray>', '', new_row, flags=re.DOTALL
    )

    ip = tbl_start + lte
    xml = xml[:ip] + new_row + xml[ip:]
    old_cnt = re.search(r'rowCnt="(\d+)"', xml[tbl_start:tbl_start+500]).group(1)
    xml = xml[:tbl_start] + xml[tbl_start:].replace(
        f'rowCnt="{old_cnt}"', f'rowCnt="{int(old_cnt)+1}"', 1
    )
    return xml
```

### 단락 복제 + 텍스트 변경

```python
def clone_paragraph(xml, search_text, new_text, insert_before_text):
    """기존 단락을 복제하여 텍스트를 변경하고 지정 위치에 삽입."""
    src_idx = xml.find(search_text)
    src_p_start = xml.rfind('<hp:p ', 0, src_idx)
    src_p_end = xml.find('</hp:p>', src_idx) + len('</hp:p>')
    src_p = xml[src_p_start:src_p_end]

    idx = [0]
    def replacer(m):
        if idx[0] == 0:
            idx[0] += 1
            return f'<hp:t>{new_text}</hp:t>'
        idx[0] += 1
        return '<hp:t></hp:t>'
    new_p = re.sub(r'<hp:t>.*?</hp:t>|<hp:t/>', replacer, src_p, flags=re.DOTALL)
    new_p = re.sub(
        r'<hp:linesegarray>.*?</hp:linesegarray>', '', new_p, flags=re.DOTALL
    )

    target_idx = xml.find(insert_before_text)
    target_p_start = xml.rfind('<hp:p ', 0, target_idx)
    return xml[:target_p_start] + new_p + xml[target_p_start:]
```

### ZIP 패키징 주의사항

```python
with zipfile.ZipFile(output, 'w') as zf:
    zf.writestr('mimetype', b'application/hwp+zip',
                compress_type=zipfile.ZIP_STORED)  # 첫 번째, STORED
    for name in file_order:
        if name == 'mimetype':
            continue
        ct = original_compress_types.get(name, zipfile.ZIP_DEFLATED)
        zf.writestr(name, all_files[name], compress_type=ct)
```

- `mimetype`이 DEFLATED → "변조" 경고. 반드시 `ZIP_STORED`
- `mimetype`은 ZIP에서 **첫 번째** 항목
- 원본 압축 방식 보존: `{n: zf.getinfo(n).compress_type for n in zf.namelist()}`

### XML 직렬화 주의사항

```python
xml_bytes = etree.tostring(root, encoding='UTF-8')
xml_bytes = xml_bytes.replace(
    b"<?xml version='1.0' encoding='UTF-8'?>",
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
)
```

텍스트만 변경하는 경우 이 불일치가 문제를 일으키지는 않음.
