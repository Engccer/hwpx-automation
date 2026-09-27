# python-hwpx API 패턴

CLI(`hwpx_edit.py`)로 안 되는 편집을 python-hwpx 스크립트로 직접 짤 때 읽는다. 머신마다 설치본 판이 다르다(3.x·6.x). 판마다 이름이 다른 API는 아래에 적었고, 설치본에서 `dir(HwpxDocument)`로 확인한다.

## 목차

- [기본 구조](#기본-구조)
- [텍스트 치환](#텍스트-치환-서식-보존-lineseg-자동-처리)
- [테이블 API](#테이블-api-lineseg-자동-처리)
- [부분 서식](#부분-서식-밑줄이탤릭볼드-run-단위)
- [머리글/바닥글](#머리글바닥글)
- [섹션/문단](#섹션문단)
- [서식 검색](#서식-검색)
- [검증](#검증)
- [양식 채우기·이미지·도형·기타](#양식-채우기이미지도형기타)
- [lxml 6.x](#lxml-6x)

## 기본 구조

```python
from hwpx.document import HwpxDocument

doc = HwpxDocument.open("template.hwpx")
# ... 편집 ...
doc.save_to_path("output.hwpx")
```

## 텍스트 치환 (서식 보존, lineseg 자동 처리)

```python
doc = HwpxDocument.open("template.hwpx")
count = doc.text.replace("{{이름}}", "홍길동")          # 6.x
# count = doc.replace_text_in_runs("{{이름}}", "홍길동")  # 3.x (6.x에도 남아 있으나 7.0 제거 예정)
# 스타일 필터링(6.x): 특정 색상/밑줄이 있는 텍스트만 치환
count = doc.text.replace("{{날짜}}", "2026-03-10", text_color="#FF0000")
doc.save_to_path("output.hwpx")
```

> **API 이름(6.0)**: `doc.replace_text_in_runs`는 `doc.text.replace`로 이동했고(구 이름은 7.0에서 제거 예정), 3.x에는 `doc.text`가 없다.
> `doc.save`는 `doc.save_to_path`로 개명됐다. 여러 판을 지원해야 하면 `hwpx_edit.py`의 `cmd_find_replace`처럼 신 이름 → 구 이름 순으로 폴백한다.
>
> **표 셀은 잡히지 않는다**: `doc.text.replace`도 `replace_text_in_runs`도 **표 셀 안 텍스트는 0건**이다.
> 표까지 치환하려면 lxml로 `hp:tbl` 하위 `hp:t`를 따로 순회해야 한다.
> `hwpx_edit.py --find/--replace`가 이미 그 2단 구성으로 처리하므로 직접 스크립트를 짜기 전에 CLI를 먼저 쓸 것.

## 테이블 API (lineseg 자동 처리)

```python
doc = HwpxDocument.open("template.hwpx")

# 표 찾기: 문단마다 담긴 표를 모은다
tables = [t for p in doc.paragraphs for t in p.tables]
table = tables[0]

# 셀 텍스트 설정 (logical=True: 병합 셀 논리 좌표 사용)
table.set_cell_text(1, 0, "텍스트", logical=True)

# 셀 직접 접근
cell = table.cell(1, 0)
cell.text = "새 텍스트"           # getter/setter, lineseg 자동 제거
cell.add_paragraph("추가 문단")   # 셀 안에 문단 추가

# 셀 문단 정렬: add_table(para_pr_id_ref=)는 셀에 적용 안 됨(셀 기본 paraPr=0=CENTER).
# 문단별로 직접 지정. 정렬값은 header.xml <hh:align horizontal="LEFT|CENTER|JUSTIFY|RIGHT"> 확인
cell.paragraphs[0].para_pr_id_ref = 3            # 첫 문단(예: 3=JUSTIFY 양쪽정렬)
cell.add_paragraph("", para_pr_id_ref=3)         # 추가 문단

# 셀 맵 (병합 포함, 2D 그리드)
cell_map = table.get_cell_map()   # list[list[HwpxTableGridPosition]]

# 행/열 정보
print(table.row_count, table.column_count)

# 셀 병합/분할
table.split_merged_cell(1, 0)

# 새 표 생성 (문서에 추가)
doc.add_table(3, 4)  # 3행 4열
```

`hwpx.tools.object_finder.ObjectFinder`는 파일 경로를 받아 raw 요소를 찾는 도구라, 여기서 얻은 결과에는 `set_cell_text`가 없다.

## 부분 서식: 밑줄·이탤릭·볼드 (run 단위)

문단 일부만 서식을 주려면 텍스트 대신 run으로 추가한다. `add_run`이 charPr를 자동 생성·재사용하므로 서식 ID를 따로 만들지 않아도 된다.

```python
p = doc.add_paragraph("")
p.add_run("일치하지 ")
p.add_run("않는", bold=True, underline=True)   # 볼드+밑줄(부정 발문·어법·지문 밑줄)
p.add_run(" 것은?")
p.add_run("Sunflowers", italic=True)            # 이탤릭(작품·매체 제목)
```

표 셀도 동일: `cell.paragraphs[0].add_run(...)` / `cell.add_paragraph("").add_run(...)`.

charPr ID가 따로 필요하면 `doc.ensure_run_style(bold=True)`(6.x는 deprecated, `doc.styles.ensure_run`)로 만든다.

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

doc.remove_paragraph(0) # 문단 삭제 (인덱스 또는 객체)
doc.remove_section(0)   # 섹션 삭제
doc.set_columns(2)      # 다단 설정
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

## 검증

```python
report = doc.validate()
if not report.ok:
    for issue in report.issues:
        print(f"{issue.part_name}: {issue.message} (line {issue.line})")
```

저장할 때마다 검증하려면 문서를 만들 때 `validate_on_save=True`를 준다(`save_to_path`의 인자가 아니다).

## 양식 채우기·이미지·도형·기타

```python
# 문서 내 모든 표의 메타데이터 조회
table_map = doc.get_table_map()

# 라벨 텍스트로 셀 탐색 (공백·대소문자·콜론 정규화 지원)
result = doc.find_cell_by_label("성명", direction="right")

# 경로 기반 일괄 채우기 ("라벨 > 방향 > ..." 형식)
doc.fill_by_path({"성명 > right": "홍길동", "생년월일 > right": "1990-01-01"})

# 이미지 임베딩 (ZIP에 추가, manifest ID 반환, hp:pic 요소 생성은 별도)
item_id = doc.add_image(open("logo.png", "rb").read(), "png")
doc.list_images()       # 임베딩된 이미지 메타데이터 조회
doc.remove_image(id)    # 이미지 제거

# 도형 삽입 (HWPUNIT 단위, 7200 per inch)
doc.add_line(start_x=0, start_y=0, end_x=14400, end_y=0)
doc.add_rectangle(width=14400, height=7200, fill_color="#E6E6E6")
# 텍스트(안내문·빈 작성 줄)를 담은 테두리 박스(학생 양식 편지칸·구획 박스)는
# add_rectangle이 아니라 raw hp:rect로 만든다 → reference/conversion.md make_rect_box

doc.add_footnote("각주 텍스트")           # 각주
doc.add_endnote("미주 텍스트")            # 미주
doc.add_bookmark("bookmark_name")         # 북마크
doc.add_hyperlink("https://...", "표시 텍스트")  # 하이퍼링크
doc.add_memo_with_anchor("메모 텍스트")   # 메모 + 앵커 자동 연결

doc.export_html()       # HTML 내보내기
doc.export_markdown()   # Markdown 내보내기
doc.export_text()       # 텍스트 내보내기
```

`fill_by_path`로 채운 셀은 양식 예시의 글자모양(회색·이탤릭)을 그대로 물려받는다(`reference/warnings-editing.md` 14번).

## lxml 6.x

python-hwpx의 메타데이터가 옛 판에서 `lxml<6`을 요구해 pip 경고가 뜰 수 있지만, lxml 6.x에서도 `HwpxDocument.open()`과 `hwpx_edit.py --to-md`는 정상 동작한다. lxml 6.x가 이미 설치돼 있으면 경고를 무시하고 그대로 쓴다. → 사례
