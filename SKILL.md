---
name: hwpx-automation
description: "HWP/HWPX/PDF 문서 읽기, 변환, 편집을 위한 통합 워크플로우. HWP 또는 HWPX 파일을 다룰 때 사용. 트리거: (1) HWP/HWPX/PDF 파일 읽기/파싱 요청 (2) HWP에서 HWPX 변환 요청 (3) HWPX 문서 편집(텍스트 치환, 표 셀 채우기, 양식 작성) (4) 한글 문서 템플릿 기반 자동화 작업 (5) HWPX 구조적 편집(행/표/단락 추가) (6) HWPX에 이미지 삽입 (7) HWPX→PDF 변환 (8) 한컴 COM 자동화"
---

# HWP/HWPX 작업 자동화 스킬

## 도구 위치

이 SKILL.md와 같은 디렉토리에 모든 도구가 포함되어 있다:
- **hwpx_edit.py**: 이 디렉토리의 `hwpx_edit.py`
- **hwp2hwpx.bat**: 이 디렉토리의 `convert/hwp2hwpx.bat`
- **python-hwpx CLI**: `pip install python-hwpx` (v2.9.0+) — `hwpx-validate`, `hwpx-page-guard` 등

실행 시 이 스킬 디렉토리의 절대경로를 사용한다. 예: `python <스킬디렉토리>/hwpx_edit.py`

## 의사결정 트리

```
HWP/HWPX 작업 요청
├── 읽기 (내용 파악)
│   ├── HWPX 파일
│   │   ├── 암호화 감지(META-INF/manifest.xml에 encryption-data 존재)
│   │   │   → 한컴 COM으로 먼저 암호 제거 (reference/encrypted-hwpx.md 참조)
│   │   ├── 표 셀 안에 긴 지문이 있는 문서(고사지·보고서)
│   │   │   → hwpx_edit.py --to-md --cell-br (셀 내부 문단을 <br>로 구분)
│   │   └── 그 외 → hwpx_edit.py --to-md (XML 직접 파싱, 무료, 정확)
│   ├── HWP 파일 → kordoc (HWP 5.x 바이너리 직접 파싱, 로컬, API 불필요)
│   │     npx kordoc <파일.hwp>
│   └── PDF 파일 → kordoc (빠른 읽기) 또는 /docparse (고품질 퓨전)
│         npx kordoc <파일.pdf>
├── 편집 (기존 문서 수정)
│   ├── HWP 파일 → 먼저 HWPX로 변환 → 이후 HWPX 편집
│   └── HWPX 파일
│       ├── 단순 텍스트 치환 → hwpx_edit.py --find/--replace
│       ├── 표 셀 채우기 → hwpx_edit.py --set-cell 또는 python-hwpx Table API
│       ├── 섹션/서식 삭제 → hwpx_edit.py --delete-after
│       ├── 빈 행 삭제 → hwpx_edit.py --delete-empty-rows
│       ├── 셀 텍스트 정리 → hwpx_edit.py --trim-cell
│       ├── 행 삭제 → hwpx_edit.py --delete-rows
│       ├── 텍스트 요소 제거 → hwpx_edit.py --remove-text
│       ├── 검은 배경 수정 → hwpx_edit.py --sanitize (save_hwpx 시 자동 적용)
│       ├── 빈 셀 hp:p 보정 → hwpx_edit.py --fix-empty-cells (save_hwpx 시 자동; Pandoc/hwpx_convert.py 변환물에 반드시 실행)
│       ├── 표/문단/머리글 추가·수정 → python-hwpx API (lineseg 자동 처리)
│       ├── 구조적 편집 (행 추가/복제) → Python + regex (아래 구조적 편집 규칙 필수)
│       └── 복잡한 편집 → python-hwpx + lxml 직접 사용 (reference/api.md + reference/structural.md 참조)
│   ※ 편집 후 검증: hwpx-validate + hwpx-page-guard
└── 새로 생성 (마크다운 등에서)
    ├── 간단한 문서 (스타일 최소, 빠른 변환)
    │   → hwpx_convert.py + 스타일 후처리 (아래 "MD → HWPX 변환 (Pandoc)" 참조)
    └── 보고서급 문서 (표/인용문/각주/커스텀 스타일 필요)
        → python-hwpx build-from-scratch (아래 "MD → HWPX 생성 (Build-from-scratch)" 참조)
```

## 도구별 용도 선택

| 작업 | 도구 | lineseg 처리 |
|------|------|-------------|
| MD/HTML → HWPX 기본 변환 | `hwpx_convert.py` (`pip install pypandoc-hwpx`) | N/A |
| 변환 후 스타일 후처리 | raw XML (header.xml + section0.xml) | **수동 제거 필수** |
| 빠른 셀 채우기, 텍스트 치환, 행 삭제 | `hwpx_edit.py` CLI | 자동 |
| 표 생성, 셀 병합/분할, 머리글/바닥글 | python-hwpx API | 자동 |
| 표에 행 추가/복제, 표 복제 | raw lxml + regex | **수동 제거 필수** |
| MD → HWPX 직접 빌드 (보고서급) | python-hwpx API + raw XML | **수동 제거 필수** |
| 무결성 검증, 쪽수 드리프트 감지 | python-hwpx CLI | N/A |

## hwpx_edit.py 사용법

```bash
# HWPX → Markdown 변환 (XML 직접 파싱, API 불필요, 오프라인 사용 가능)
python hwpx_edit.py <파일.hwpx> --to-md [-o output.md]

# 표 셀 안에 긴 지문(일기·본문)이 있는 문서는 --cell-br 권장
#  - 기본(--to-md만): 셀 내 모든 <hp:t>를 공백 하나로 합침 (문단 경계 손실)
#  - --cell-br:      셀 내 <hp:p> 문단을 <br>로 구분 (고사지·보고서 권장)
python hwpx_edit.py <파일.hwpx> --to-md --cell-br [-o output.md]

# 암호화된 HWPX (AES-256-CBC)는 hwpx_edit.py가 자동 감지하고 오류 메시지로
# 해제 절차 안내한다. 해제는 한컴 COM으로 FilePasswordChange 액션 사용.
# 상세: reference/encrypted-hwpx.md

# 표 구조 확인 (표 개수, 행/열 수, 셀 내용 미리보기)
python hwpx_edit.py <파일.hwpx> --info

# 텍스트 치환 (본문 + 표 셀 모두 처리)
python hwpx_edit.py <파일.hwpx> --find "이전" --replace "이후"

# 표 셀에 텍스트 입력 (표번호,행번호,셀번호 — 0부터 시작)
python hwpx_edit.py <파일.hwpx> --set-cell 0,1,0 "텍스트"

# 병합 셀 분할
python hwpx_edit.py <파일.hwpx> --split-cell 0,1,0

# 특정 텍스트 이후 모든 요소 삭제 (서식 분리 등)
python hwpx_edit.py <파일.hwpx> --delete-after "<서식 2>"

# 테이블 끝의 빈 행 삭제
python hwpx_edit.py <파일.hwpx> --delete-empty-rows 1

# 셀 텍스트의 첫 줄만 유지 (줄바꿈 이후 삭제)
python hwpx_edit.py <파일.hwpx> --trim-cell 1,1,1

# 특정 행 삭제 (쉼표로 행 인덱스 구분)
python hwpx_edit.py <파일.hwpx> --delete-rows 1 11,12

# 특정 텍스트를 포함하는 <hp:t> 요소 완전히 제거
python hwpx_edit.py <파일.hwpx> --remove-text "2027. 2. 28."

# HWP→HWPX 변환 후 검은 배경 문제 수정
python hwpx_edit.py <파일.hwpx> --sanitize

# 빈 표 셀에 기본 hp:p 삽입 (Pandoc/hwpx_convert.py 변환 후 한글이 15초 로딩 후
# 닫히는 문제 해결. MD 표의 빈 셀 ` | | `이 <hp:subList>만 있고 <hp:p>가 없는
# 구조로 생성되는 것이 원인. XSD 스키마는 통과하므로 hwpx-validate로는 탐지
# 불가. hwpx_edit.py 내부 save_hwpx 경로에서는 자동 적용되지만, 외부 도구로
# 만든 HWPX는 이 명령을 단독 실행해야 함)
python hwpx_edit.py <파일.hwpx> --fix-empty-cells

# 별도 파일로 저장
python hwpx_edit.py <파일.hwpx> --set-cell 0,1,0 "텍스트" -o output.hwpx
```

## kordoc 사용법 (HWP/HWPX/PDF 읽기)

```bash
# HWP → Markdown (HWP 5.x 바이너리 직접 파싱, API 불필요, 오프라인)
npx kordoc <파일.hwp>

# HWPX → Markdown
npx kordoc <파일.hwpx>

# PDF → Markdown (표 감지 포함)
npx kordoc <파일.pdf>

# JSON 구조화 출력 (블록+메타데이터)
npx kordoc <파일.hwp> --format json

# 파일로 저장
npx kordoc <파일.hwp> -o output.md

# 특정 페이지만 파싱
npx kordoc <파일.pdf> --pages 1-5

# 배치 변환
npx kordoc *.pdf -d ./converted/
```

> kordoc는 npm 패키지. HWP 5.x 바이너리를 직접 파싱하므로 한컴오피스 불필요.
> HWPX 읽기는 hwpx_edit.py --to-md와 kordoc 모두 가능. 편집 워크플로우 연계 시 hwpx_edit.py 권장.
> PDF 고품질 파싱은 /docparse 스킬 사용 (다중 파서 퓨전).

## python-hwpx CLI 도구 (v2.9.0+)

```bash
# XSD 스키마 검증 — 편집 후 무결성 확인
hwpx-validate <파일.hwpx>

# ZIP/OPC 패키지 구조 검증 (mimetype, container.xml, manifest)
hwpx-validate-package <파일.hwpx>

# 레퍼런스 대비 페이지 드리프트 감지 — 양식 편집 시 필수
hwpx-page-guard --reference <원본.hwpx> --output <결과.hwpx>

# 문서 구조 심층 분석 (폰트, 스타일 ID, 표 구조 등)
hwpx-analyze-template <파일.hwpx> [--extract-dir <경로>] [--json]

# HWPX 내용 추출 (텍스트/마크다운)
hwpx-text-extract <파일.hwpx> [--format markdown] [--output <파일>]

# HWPX 언팩/리팩
hwpx-unpack <파일.hwpx> <디렉토리> [--pretty-xml]
hwpx-pack <디렉토리> <출력.hwpx>
```

## HWP → HWPX 변환

```bash
convert/hwp2hwpx.bat <입력.hwp> [출력.hwpx]
```

- 출력 파일 미지정 시 같은 폴더에 `.hwpx` 확장자로 생성
- 서식 100% 보존 (Java 기반, hwplib + hwpxlib)
- 요구사항: JDK 21 (`C:/Program Files/Eclipse Adoptium/jdk-21.0.10.7-hotspot`)

## 핵심 워크플로우: MD → HWPX 변환 (Pandoc 방식)

> 간단한 문서에 적합. 복잡한 보고서는 아래 "Build-from-scratch 방식" 참조.

1. **사전 요구사항 확인**: `pip install pypandoc-hwpx` (미설치 시 ModuleNotFoundError)
2. **전처리** (필수!): (a) 따옴표(`"…"`, `'…'`, `"…"`, `'…'`) → PUA 마커(U+FFF0~3) 치환 (Pandoc HWPX writer가 따옴표 안 텍스트를 통째 누락시킴) 또는 `「…」` 한글 괄호로 대체, (b) `> ` blockquote → `【인용】` 마커 치환 또는 일반 문단으로 변환. **둘 다 필수.**
3. **기본 변환**: `python hwpx_convert.py <입력.md> -o <출력.hwpx>`
4. **빈 셀 수정** (필수!): `python hwpx_edit.py <출력.hwpx> --fix-empty-cells` — MD 표의 빈 셀(` | | `)이 HWPX에서 `<hp:subList>`만 있고 `<hp:p>`가 없는 상태로 변환됨. 한글 엔진이 이를 만나면 약 15초 로딩 후 안전 종료함(XSD 스키마는 통과하므로 `hwpx-validate`로는 탐지 불가). `hwpx_convert.py`는 이 보정을 자동 적용하지 않으므로 **변환 직후 반드시 실행**.
5. **스타일 후처리** (raw XML 편집):
   - header.xml: borderFill 추가 (표 배경색), charPr 추가 (굵은/이탤릭), 본문 글꼴 크기 변경
   - section0.xml: PUA 마커 → 스마트 따옴표 복원(`postprocess_quotes`), 표 헤더 행 배경색, 합계 행 배경색, 인용문 좌측 여백+이탤릭, `【인용】` 마커 제거
   - lineseg 전체 제거 (`re.sub(r'<hp:linesegarray>.*?</hp:linesegarray>', '', xml, flags=re.DOTALL)`)
6. **머리글/바닥글**: python-hwpx API (`doc.set_header_text()`, `doc.set_footer_text()`)
7. **검증**: `hwpx-validate <출력.hwpx>`

**주의**: 스타일 후처리에서 한국어 텍스트 검색 시, `python -c "..."` 터미널 명령은 cp949 인코딩 문제 발생. 반드시 `.py` 파일(UTF-8)로 작성하거나 유니코드 이스케이프(`"\uc778\uc6a9"` = "인용") 사용.

> 상세 코드 패턴 (borderFill/charPr 추가, 표 행 스타일링, 인용문 스타일링 등)은 `reference/conversion.md` 참조.

## 핵심 워크플로우: MD → HWPX 생성 (Build-from-scratch 방식)

> 보고서급 문서에 적합. python-hwpx API로 문서 구조를 처음부터 빌드하고, XML 후처리로 세밀한 스타일링을 적용한다.
> Pandoc 방식의 한계 (따옴표 안 텍스트 누락, blockquote 누락, 각주 미변환, 제한적 스타일)를 완전히 해결.

### 언제 사용하는가

| 조건 | Pandoc 방식 | Build-from-scratch |
|------|------------|-------------------|
| 표 스타일 (헤더 배경, 교대 행, 볼드 셀) | 후처리 필요 | 직접 제어 |
| 따옴표 안 텍스트 (`"…"`, `'…'`) | 누락 → PUA 마커 전처리+후처리 필수 | 직접 제어 (문제 없음) |
| 인용문 (blockquote) | 누락 → 마커 전처리 필요 | 좌측 컬러바 + 배경색 직접 적용 |
| 각주 | 문서 끝 일반 텍스트 | 위첨자 + 각주 섹션 분리 |
| 커스텀 글꼴 | 제한적 | fontface 직접 추가 |
| 구현 비용 | 낮음 | 높음 (프로젝트별 스크립트 작성) |

### 5단계 워크플로우

```
[1] MD 파싱 → 블록 리스트 (heading, paragraph, table, blockquote, list, footnote)
[2] 빈 HWPX 템플릿 생성 → header.xml에 커스텀 스타일 주입
[3] python-hwpx API로 문서 빌드 (표지, 본문, 각주 섹션)
[4] XML 후처리 (표 스타일링, lineseg 제거, 빈 셀 수정)
[5] ZIP 리패키징 및 저장
```

### 핵심 패턴 요약

**1단계 — MD 파싱**: `parse_markdown()` + `parse_inline()`으로 블록/인라인 구조 추출

**2단계 — 스타일 주입** (`inject_styles`):
- `header.xml`의 `borderFills`, `charProperties`, `paraProperties`에 커스텀 항목 추가
- 빈 템플릿 기본 카운트: charPr 0-6 (7개), borderFill 1-2 (2개), paraPr 0-19 (20개)
- 새 항목 ID = 기본 카운트 + 순서 (결정론적 계산)

**3단계 — API 빌드**: `doc.add_paragraph()`, `doc.add_table()`, `tbl.set_cell_text()`

**4단계 — XML 후처리**: 표 헤더/교대행/합계행/볼드셀 스타일링, lineseg 전체 제거

**5단계 — 저장**: `doc.save_to_path()` 또는 ZIP 수동 리패키징

> 상세 코드 패턴, 주요 함정, 디자인 토큰 등은 `reference/build-from-scratch.md` 참조.

## 핵심 워크플로우: 양식 채우기

1. **원본이 HWP이면 변환**: `convert/hwp2hwpx.bat input.hwp output.hwpx`
2. **내용 파악**: `hwpx_edit.py output.hwpx --to-md`
3. **표 구조 확인**: `hwpx_edit.py output.hwpx --info`
4. **(선택) 심층 분석**: `hwpx-analyze-template output.hwpx --json` (스타일 ID 맵 필요 시)
5. **데이터 채우기**: `--set-cell`로 개별 셀 또는 Python 스크립트로 일괄 처리
6. **무결성 검증**: `hwpx-validate result.hwpx`
7. **쪽수 드리프트 감지**: `hwpx-page-guard -r output.hwpx -o result.hwpx`
8. **내용 확인**: `--to-md`로 최종 텍스트 검증

## python-hwpx API (v2.9.0) — 주요 패턴

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
count = doc.replace_text_in_runs("{{이름}}", "홍길동")
doc.save_to_path("output.hwpx")
```

### 테이블 API (lineseg 자동 처리)

```python
doc = HwpxDocument.open("template.hwpx")

# 표 찾기 — ObjectFinder 사용
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

# 셀 맵 (병합 포함, 논리 좌표 → 물리 좌표)
cell_map = table.get_cell_map()

# 셀 병합/분할
table.split_merged_cell(1, 0)

# 새 표 생성
doc.add_table(3, 4)  # 3행 4열
```

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
# 이미지 임베딩 (ZIP에 추가, manifest ID 반환 — hp:pic 요소 생성은 별도)
item_id = doc.add_image(open("logo.png", "rb").read(), "png")
doc.list_images()       # 임베딩된 이미지 메타데이터 조회
doc.remove_image(id)    # 이미지 제거

# 도형 삽입 (HWPUNIT 단위, 7200 per inch)
doc.add_line(start_x=0, start_y=0, end_x=14400, end_y=0)
doc.add_rectangle(width=14400, height=7200, fill_color="#E6E6E6")
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

> **lxml 6.0.4 호환성**: python-hwpx의 `requires lxml<6` 제약이 있으나 실제 동작은 정상 (2026.4 확인). pip 경고만 발생.

## 구조적 편집 (행/표/단락 추가) — raw lxml 필요

> python-hwpx API에 `add_row()`, `insert_row()`가 **없으므로**, 표에 행을 추가/복제하는 경우에만 raw lxml + regex를 사용한다.
> 단순 셀 편집, 텍스트 치환, 표 생성은 위의 API를 사용하라.

### 3가지 필수 규칙

1. **`<hp:linesegarray>` 제거**: 텍스트가 변경된 모든 `<hp:p>`에서 제거. 한글이 열 때 자동 재계산.
   - python-hwpx API 사용 시 **자동 처리됨** (내부에서 `_clear_paragraph_layout_cache()` 호출)
   - raw lxml/regex로 직접 XML을 편집할 때만 수동 제거 필요
   - openhwp Rust 소스로 검증: `line_segments: Option<LineSegmentArray>` — 생략이 스키마적으로 유효

2. **`rowAddr` 순차**: 새 행의 `<hp:cellAddr rowAddr="N"/>`이 기존 행과 중복 없이 순차적. 중복 → 무한 로딩.

3. **`rowCnt` 일치**: `<hp:tbl rowCnt="N">`이 실제 `<hp:tr>` 수와 일치. 불일치 → "파일이 손상되었습니다".

4. **표 `id` 고유화**: 새 표 복제 시 `id` 중복 금지.

> 상세 코드 패턴 (행 추가, 단락 복제, ZIP 패키징, XML 직렬화 등)은 `reference/structural.md` 참조.

## HWPX 파일 구조

```
document.hwpx (ZIP)
├── mimetype                    # "application/hwp+zip" (첫 항목, ZIP_STORED)
├── META-INF/
│   ├── container.xml           # 루트 파일 지정
│   ├── container.rdf           # 메타데이터
│   └── manifest.xml            # 파일 목록
├── Contents/
│   ├── content.hpf             # 콘텐츠 매니페스트
│   ├── header.xml              # 문서 설정 (글꼴, 스타일, borderFill 등)
│   └── section0.xml            # 본문 내용 ★ 편집 대상
├── Preview/
│   ├── PrvImage.png            # 미리보기 이미지
│   └── PrvText.txt             # 미리보기 텍스트
├── settings.xml                # 편집 설정
└── version.xml                 # 버전 정보
```

> XML 요소 구조, 네임스페이스 딕셔너리 등 상세 구조는 `reference/format.md` 참조.

## 한컴 COM 자동화 (Windows 전용, 한컴오피스 필수)

한컴오피스가 설치된 Windows 환경에서 `HWPFrame.HwpObject` COM을 통해 HWPX를 조작할 수 있다.
XML 직접 편집으로 불가능한 작업(이미지 삽입, PDF 변환)에 사용.

### 보안모듈 (팝업 제거)

한컴 COM으로 파일을 열거나 저장할 때 "접근 허용" 보안 대화상자가 반복 표시된다. 이를 제거하려면:

1. **DLL**: `<repo>/vendor/FilePathCheckerModuleExample.dll` — 저장소에 커밋되지 않음 (한컴 독점 라이선스). 설치: `vendor/README.md` 참조
2. **레지스트리**: `HKCU\SOFTWARE\HNC\HwpAutomation\Modules` → 문자열 값 `FilePathCheckerModule` = DLL 전체 경로
3. **코드**: `hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')`

> 다운로드 출처: `https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/보안모듈(Automation).zip` → 압축 해제 후 `vendor/`에 배치 → 레지스트리 등록

> **경로 일치 주의**: 레지스트리에 등록된 DLL 경로와 실제 DLL 파일 경로가 일치해야 한다. 레지스트리는 남아 있지만 DLL이 다른 위치에만 있으면(예: 프로젝트 이동 후) 보안 팝업이 계속 뜨며 `RegisterModule` 호출도 무시된다. 증상: 스크립트가 매 Open/SaveAs마다 멈춤. 진단: `reg query "HKCU\SOFTWARE\HNC\HwpAutomation\Modules"`로 경로 확인 후 실제 DLL 존재 여부 검증.

### 암호화된 HWPX 해제 (비밀번호 필요)

AES-256-CBC로 암호화된 HWPX는 `hwpx_edit.py`로 직접 파싱할 수 없다. `Open()`의 세 번째 인자로 비밀번호를 넘긴 뒤 `FilePasswordChange` 액션으로 암호를 제거해야 한다. 단순 `SaveAs`만 하면 암호화된 채로 저장된다.

```python
hwp.Open(abs_in, 'HWPX', f'password:{PASSWORD}')
act = hwp.CreateAction('FilePasswordChange')
pset = act.CreateSet()
act.GetDefault(pset)
pset.SetItem('String', '')       # 새 암호: 빈 문자열 = 제거
pset.SetItem('Ask', 0)           # 0 = 대화상자 없이 적용
pset.SetItem('ReadString', '')
pset.SetItem('WriteString', '')
pset.SetItem('RWAsk', 0)
act.Execute(pset)
hwp.SaveAs(abs_out, 'HWPX', '')
```

> 감지: `META-INF/manifest.xml`에 `<odf:encryption-data>` 항목 존재 여부로 확인. `hwpx_edit.py`는 `is_encrypted_hwpx()`로 자동 감지하고 친절한 오류 메시지 출력.
> 상세: `reference/encrypted-hwpx.md` 참조

### 기본 패턴

```python
import win32com.client as win32
import pythoncom, os, time

pythoncom.CoInitialize()
hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.XHwpWindows.Item(0).Visible = False
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')

hwp.Open(os.path.abspath('input.hwpx'), 'HWPX', '')
# ... 작업 ...
hwp.SaveAs(os.path.abspath('output.hwpx'), 'HWPX', '')
hwp.SaveAs(os.path.abspath('output.pdf'), 'PDF', '')
hwp.Quit()
pythoncom.CoUninitialize()
```

### HWPX → PDF 변환

```python
hwp.Open(abs_path, 'HWPX', '')
hwp.SaveAs(abs_pdf_path, 'PDF', '')
```

### 이미지 삽입

```python
# 커서를 원하는 위치로 이동
hwp.HAction.Run('MoveDocEnd')

# 이미지 삽입 (Width/Height 단위: mm)
ctrl = hwp.InsertPicture(abs_img_path, True, 1, False, False, 0, 23, 15)
# params: path, Embedded, sizeoption(1=지정크기), Reverse, Watermark, Effect, Width_mm, Height_mm
```

### 이미지 위치/배치 변경 — XML 후처리 방식 (권장)

COM의 `ShapeObjDialog`로 속성 변경 시 `TreatAsChar`, `TextWrap` 등이 정상 반영되지 않는 경우가 많다.
**권장 워크플로우**: COM으로 이미지 삽입 → HWPX 저장 → XML 후처리로 `<hp:pic>` 속성 변경 → 다시 COM으로 PDF 저장.

```python
# XML에서 hp:pic의 pos/sz/textWrap 속성을 직접 수정
# 수동 완성본의 <hp:pic> 요소를 분석하여 좌표/크기를 복사하는 것이 가장 정확
import re
section_xml = re.sub(
    r'(<hp:pic[^>]*?)textWrap="[^"]*"',
    r'\1textWrap="IN_FRONT_OF_TEXT"', section_xml)
section_xml = re.sub(
    r'(<hp:pic[^>]*>.*?)<hp:pos [^/]*/>(.*?</hp:pic>)',
    r'\1<hp:pos treatAsChar="0" ... vertRelTo="PAPER" horzRelTo="PAPER" '
    r'vertOffset="NNN" horzOffset="NNN"/>\2',
    section_xml, flags=re.DOTALL)
```

### 텍스트 검색

```python
fr = hwp.HParameterSet.HFindReplace
hwp.HAction.GetDefault('RepeatFind', fr.HSet)
fr.FindString = '검색어'
fr.MatchCase = 1
fr.ReplaceMode = 0  # find only
hwp.HAction.Execute('RepeatFind', fr.HSet)
# 이후 커서 이동: hwp.HAction.Run('MoveRight') 등
```

### 주요 주의사항

1. **COM 프로세스 잔류**: 에러 발생 시 `hwp.Quit()`이 호출되지 않아 `Hwp.exe`가 잔류. 다음 실행 전 `taskkill /F /IM Hwp.exe` 필요
2. **COM 재초기화 금지**: 동일 프로세스에서 `CoUninitialize()` 후 `CoInitialize()`를 다시 호출하면 segfault 발생 가능. COM 작업을 2단계로 나눠야 하면 **별도 Python 스크립트**로 분리
3. **HParameterSet 속성명**: COM 타입 라이브러리에 따라 속성명이 다름. `dir(obj)` 또는 `_prop_map_put_`으로 확인. 예: `IgnoreCase` → `MatchCase`, `FileName` → `filename`
4. **InsertPicture 반환값**: 성공 시 `IDHwpCtrlCode` COM 객체, 실패 시 `None`/`False`
5. **절대 경로 필수**: COM API에 전달하는 파일 경로는 반드시 `os.path.abspath()` 사용
6. **`Close` 메서드 없음**: `hwp.Close()`는 존재하지 않음. 문서를 닫으려면 `hwp.Quit()` (앱 종료)

## 상세 레퍼런스

| 주제 | 파일 | 내용 |
|------|------|------|
| python-hwpx API 전체 | `reference/api.md` | 색상 필터 치환, 섹션/문단 추가, 서식 검색, 검증 등 전체 패턴 |
| MD → HWPX 변환 (Pandoc) | `reference/conversion.md` | hwpx_convert.py 사용법, blockquote 전처리, 스타일 후처리 코드 (borderFill/charPr/표 스타일링/인용문) |
| MD → HWPX 생성 (Build) | `reference/build-from-scratch.md` | 5단계 워크플로우, 디자인 토큰, precompute_styles 패턴, 주요 함정 (lineSpacing/fontRef/set_cell_text) |
| 구조적 편집 코드 | `reference/structural.md` | lineseg 원리, 행 추가/단락 복제 코드, ZIP 패키징, XML 직렬화 |
| HWPX 파일 구조 | `reference/format.md` | ZIP 내부 구조, section0.xml 요소, 네임스페이스 딕셔너리 |
| 스킬 업데이트 | `reference/update-checklist.md` | python-hwpx 업데이트, GitHub 인사이트 수집, 반영 기준 |
| 암호화 HWPX 해제 | `reference/encrypted-hwpx.md` | 암호화 감지, 한컴 COM 암호 해제 전체 스크립트, 주의사항 |
| 한컴 COM 액션 테이블 | `reference/action-table.md` | 모든 HAction ID와 대응 ParameterSet ID (공식, 2025.04). Gemini 파싱, "한글"→"호글" 오인식 있으나 API명 검색에 무영향 |
| 한컴 COM API 가이드 | `reference/hwp-automation.md` | IHwpObject 전체 API — 프로퍼티, 메서드, 이벤트, ParameterSet 상세 (공식, 2025.04). Gemini 파싱, 다이어그램은 텍스트 설명으로 대체됨 |
| 한컴 COM 파라미터셋 | `reference/parameterset-table.md` | 149개 ParameterSet 필드/타입/기본값 (공식, 2025.04). Mistral 파싱, 워터마크 환각 잔여 가능, heading이 평문 처리된 경우 있음 |

## 주의사항

워크플로우별 주의사항은 각 reference 파일에 포함:
- **XML 편집 전반**: `reference/warnings-editing.md` — find/replace 동작, set-cell 체이닝, 서식 보존, PrvText.txt 누락, manifest self-closing 등
- **MD → HWPX 변환**: `reference/conversion.md` 하단 — blockquote 누락, 각주 변환, 인코딩, 스크립트 보존
- **Build-from-scratch**: `reference/build-from-scratch.md` 하단 — lineSpacing 단위, fontRef, set_cell_text, style_tables_xml 루프
- **한컴 COM 자동화**: `reference/warnings-com.md` — 이미지 XML 삽입 불가, 양식+이미지 워크플로우, PDF 변환
