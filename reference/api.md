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
count = doc.text.replace("{{이름}}", "홍길동")
# 스타일 필터링: 특정 색상/밑줄이 있는 텍스트만 치환
count = doc.text.replace("{{날짜}}", "2026-03-10", text_color="#FF0000")
doc.save_to_path("output.hwpx")
```

> **API 이력(6.0)**: `doc.replace_text_in_runs`는 `doc.text.replace`로 이동했고(구 이름은 7.0에서 제거),
> `doc.save`는 `doc.save_to_path`로 개명됐다(6.x에는 `save` 자체가 없어 호출 시 AttributeError).
> 구버전도 지원해야 하면 `getattr(doc, "save_to_path", None) or doc.save` 식으로 폴백한다.
>
> **표 셀은 잡히지 않는다**: `doc.text.replace`도 `replace_text_in_runs`도 **표 셀 안 텍스트는 0건**이다
> (6.0.2 실측). 표까지 치환하려면 lxml로 `hp:tbl` 하위 `hp:t`를 따로 순회해야 한다
> (`hwpx_edit.py --find/--replace`가 쓰는 2단 구성).

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

## python-hwpx API (v2.9.1): 주요 패턴

### 기본 구조

```python
from hwpx.document import HwpxDocument

doc = HwpxDocument.open("template.hwpx")
# ... 편집 ...
doc.save_to_path("output.hwpx")
```

### 텍스트 치환 (서식 보존, lineseg 자동 처리)

```python
doc = HwpxDocument.open("template.hwpx")
count = doc.text.replace("{{이름}}", "홍길동")   # 6.0 이전 이름: replace_text_in_runs (7.0에서 제거)
doc.save_to_path("output.hwpx")                  # 6.0 이전 이름: save (6.x에는 없음)
```

> **표 셀은 이 API로 잡히지 않는다**(6.0.2 실측, 0건). 표까지 치환하려면 lxml로 `hp:tbl` 하위 `hp:t`를 따로 순회한다 — `hwpx_edit.py --find/--replace`가 이미 그 2단 구성으로 처리하므로 직접 스크립트를 짜기 전에 CLI를 먼저 쓸 것.

### 테이블 API (lineseg 자동 처리)

```python
doc = HwpxDocument.open("template.hwpx")

# 표 찾기: ObjectFinder 사용
from hwpx.oxml.object_finder import ObjectFinder
finder = ObjectFinder(doc)
tables = finder.find_all(tag="tbl")

# 셀 텍스트 설정 (병합 셀도 자동 처리)
table = tables[0]
table.set_cell_text(1, 0, "텍스트", logical=True)

# 셀 직접 접근
cell = table.cell(1, 0)
cell.text = "새 텍스트"           # getter/setter
cell.add_paragraph("추가 문단")   # 셀 안에 문단 추가

# 셀 문단 정렬: add_table(para_pr_id_ref=)는 셀에 적용 안 됨(셀 기본 paraPr=0=CENTER).
# 문단별로 직접 지정. 정렬값은 header.xml <hh:align horizontal="LEFT|CENTER|JUSTIFY|RIGHT"> 확인
cell.paragraphs[0].para_pr_id_ref = 3            # 첫 문단(예: 3=JUSTIFY 양쪽정렬)
cell.add_paragraph("", para_pr_id_ref=3)         # 추가 문단

# 셀 맵 (병합 포함, 논리 좌표 → 물리 좌표)
cell_map = table.get_cell_map()

# 셀 병합/분할
table.split_merged_cell(1, 0)

# 새 표 생성
doc.add_table(3, 4)  # 3행 4열
```

### 부분 서식: 밑줄·이탤릭·볼드 (run 단위)

문단 일부만 서식을 주려면 텍스트 대신 run으로 추가한다. `add_run`이 charPr를 자동 생성·재사용하므로 `ensure_run_style`을 따로 부르지 않아도 된다.

```python
p = doc.add_paragraph("")
p.add_run("일치하지 ")
p.add_run("않는", bold=True, underline=True)   # 볼드+밑줄(부정 발문·어법·지문 밑줄)
p.add_run(" 것은?")
p.add_run("Sunflowers", italic=True)            # 이탤릭(작품·매체 제목)
```

표 셀도 동일: `cell.paragraphs[0].add_run(...)` / `cell.add_paragraph("").add_run(...)`.

### 머리글/바닥글

```python
doc.set_header_text("문서 제목", page_type="BOTH")
doc.set_footer_text("페이지 번호", page_type="BOTH")
```

### 검증 (편집 후)

```python
# 프로그래밍 방식 검증
report = doc.validate()
if not report.ok:
    for issue in report.issues:
        print(f"{issue.part_name}: {issue.message}")

# 저장 시 자동 검증
doc.save_to_path("output.hwpx", validate_on_save=True)
```

### 서식 검색

```python
# 밑줄이 있는 런 찾기
underlined = doc.find_runs_by_style(underline_type="BOTTOM")
for run in underlined:
    print(run.text)
```

> 전체 API 패턴 (색상 필터링, 섹션/문단 추가 등)은 `reference/api.md` 참조.

### v2.9.0 신규 API (2026.4)

#### 테이블 자동화 (양식 채우기에 유용)

```python
# 문서 내 모든 표의 메타데이터 조회
table_map = doc.get_table_map()

# 라벨 텍스트로 셀 탐색 (공백·대소문자·콜론 정규화 지원)
result = doc.find_cell_by_label("성명", direction="right")

# 경로 기반 일괄 채우기 ("라벨 > 방향 > ..." 형식)
doc.fill_by_path({"성명 > right": "홍길동", "생년월일 > right": "1990-01-01"})
```

#### 이미지·도형

```python
# 이미지 임베딩 (ZIP에 추가, manifest ID 반환, hp:pic 요소 생성은 별도)
item_id = doc.add_image(open("logo.png", "rb").read(), "png")
doc.list_images()       # 임베딩된 이미지 메타데이터 조회
doc.remove_image(id)    # 이미지 제거

# 도형 삽입 (HWPUNIT 단위, 7200 per inch)
doc.add_line(start_x=0, start_y=0, end_x=14400, end_y=0)
doc.add_rectangle(width=14400, height=7200, fill_color="#E6E6E6")
# 텍스트(안내문·빈 작성 줄)를 담은 테두리 박스(학생 양식 편지칸·구획 박스)는
# add_rectangle이 아니라 raw hp:rect로 만든다 → reference/conversion.md make_rect_box
```

#### 기타

```python
doc.add_footnote("각주 텍스트")           # 각주
doc.add_endnote("미주 텍스트")            # 미주
doc.add_bookmark("bookmark_name")         # 북마크
doc.add_hyperlink("https://...", "표시 텍스트")  # 하이퍼링크
doc.add_memo_with_anchor("메모 텍스트")   # 메모 + 앵커 자동 연결

doc.export_html()       # HTML 내보내기
doc.export_markdown()   # Markdown 내보내기
doc.export_text()       # 텍스트 내보내기

doc.remove_paragraph(0) # 문단 삭제 (인덱스 또는 객체)
doc.remove_section(0)   # 섹션 삭제
doc.set_columns(2)      # 다단 설정

# 스타일 자동 생성/재사용
cp_id = doc.ensure_run_style(bold=True, italic=False)
```

> **lxml 6.x 호환성**: python-hwpx 2.9.1의 `requires lxml<6` 제약이 유지되지만, 실측상 lxml 6.1.0에서도 정상 동작한다 (2026-05-04 `HwpxDocument.open()` + `hwpx_edit.py --to-md` 재검증 완료). pip 경고만 발생하므로, lxml 6.x가 이미 설치돼 있다면 핀 경고를 무시하고 그대로 써도 된다.
