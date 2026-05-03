# python-hwpx API 패턴 (v2.9.1)

## 기본 구조

```python
from hwpx.document import HwpxDocument

doc = HwpxDocument.open("template.hwpx")
# ... 편집 ...
doc.save_to_path("output.hwpx")

# 저장 시 자동 검증
doc.save_to_path("output.hwpx", validate_on_save=True)
```

## 텍스트 치환 (서식 보존, lineseg 자동 처리)

```python
doc = HwpxDocument.open("template.hwpx")
count = doc.replace_text_in_runs("{{이름}}", "홍길동")
# 스타일 필터링: 특정 색상/밑줄이 있는 텍스트만 치환
count = doc.replace_text_in_runs("{{날짜}}", "2026-03-10", text_color="#FF0000")
doc.save_to_path("output.hwpx")
```

## 테이블 API (lineseg 자동 처리)

```python
doc = HwpxDocument.open("template.hwpx")

# 표 찾기
from hwpx.oxml.object_finder import ObjectFinder
finder = ObjectFinder(doc)
tables = finder.find_all(tag="tbl")
table = tables[0]

# 셀 텍스트 설정 (logical=True: 병합 셀 논리 좌표 사용)
table.set_cell_text(1, 0, "텍스트", logical=True)

# 셀 직접 접근
cell = table.cell(1, 0)
cell.text = "새 텍스트"           # getter/setter, lineseg 자동 제거
cell.add_paragraph("추가 문단")   # 셀 안에 문단 추가

# 셀 맵 (병합 포함, 2D 그리드)
cell_map = table.get_cell_map()   # list[list[HwpxTableGridPosition]]

# 행/열 정보
print(table.row_count, table.column_count)

# 셀 병합/분할
table.split_merged_cell(1, 0)

# 새 표 생성 (문서에 추가)
doc.add_table(3, 4)  # 3행 4열
```

## 머리글/바닥글

```python
doc.set_header_text("문서 제목", page_type="BOTH")  # BOTH, ODD, EVEN
doc.set_footer_text("- {} -".format("페이지"), page_type="BOTH")
doc.remove_header()  # 제거
```

## 섹션/문단

```python
# 문단 추가 (이전 문단 서식 자동 상속)
doc.add_paragraph("새 문단 텍스트", inherit_style=True)

# 섹션 추가
doc.add_section(after=0)  # 첫 번째 섹션 뒤에 추가
```

## 서식 검색

```python
# 밑줄이 있는 런 찾기
underlined = doc.find_runs_by_style(underline_type="BOTTOM")
for run in underlined:
    print(run.text)

# 특정 색상 텍스트 찾기
red_runs = doc.find_runs_by_style(text_color="#FF0000")
```

## 프로그래밍 방식 검증

```python
report = doc.validate()
if not report.ok:
    for issue in report.issues:
        print(f"{issue.part_name}: {issue.message} (line {issue.line})")
```
