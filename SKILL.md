---
name: hwpx-automation
description: "HWP/HWPX 문서 읽기, 변환, 편집을 위한 통합 워크플로우. HWP 또는 HWPX 파일을 다룰 때 사용. HWP 파일은 모두 HWPX로 변환 후 처리한다. 트리거: (1) HWP/HWPX 파일 읽기/파싱 요청 (2) HWP→HWPX 변환 요청 (3) HWPX 문서 편집(텍스트 치환, 표 셀 채우기, 양식 작성) (4) 한글 문서 템플릿 기반 자동화 작업 (5) HWPX 구조적 편집(행/표/단락 추가) (6) HWPX에 이미지 삽입 (7) HWPX→PDF 변환 (8) 한컴 COM 자동화 (9) HWPX 서명란에 서명·도장 이미지 삽입(signature/seal/도장 삽입, 동의서·계약서·서약서 서명)"
---

# HWP/HWPX 작업 자동화 스킬

## 도구 위치

이 SKILL.md와 같은 디렉토리에 모든 도구가 포함되어 있다:
- **hwpx_edit.py**: 이 디렉토리의 `hwpx_edit.py` (편집 명령 + `--to-md` CLI 래퍼)
- **hwpx-tomd**: `--to-md` 변환 엔진을 단일 소스로 보유한 독립 패키지. `pip install hwpx-tomd` (PyPI·GitHub `Engccer/hwpx-tomd` 공개. 엔진 자체를 수정할 때만 로컬 editable: `pip install -e path/to/hwpx-tomd`). 라이브러리로도 직접 사용 가능(`from hwpx_tomd import to_markdown, convert`). 변환 로직은 이 패키지에만 있고 `hwpx_edit.py`는 호출만 한다(코드 분기 방지).
- **hwpx_convert.py**: 이 디렉토리의 `convert/hwpx_convert.py` (MD/DOCX/HTML/RST/TEX/TXT → HWPX 변환, `pip install pypandoc-hwpx` 필요)
- **hwpx_com.py**: 이 디렉토리의 `hwpx_com.py` (한컴 COM 네이티브 파이프라인, pyhwpx 기반, Windows + 한컴오피스 전용, `pip install pyhwpx` 필요). MD→HWPX 생성·이미지 삽입·본문 추출·한컴 재저장 정규화(--normalize)·PDF 변환. XML/Pandoc 파이프라인과 **분리 운용**(아래 "한컴 COM 자동화" 참조)
- **hwpx_sign.py**: 이 디렉토리의 `hwpx_sign.py` (서명란(기준 텍스트 "(서명)" 등)에 서명/도장 이미지를 삽입하는 전용 도구, Windows + 한컴오피스 필요). COM으로 이미지를 넣어 BinData를 확보한 뒤 XML 후처리로 floating PAPER 절대좌표 전환·앵커 위 문단 이동·lineseg 제거를 자동 수행한다(아래 "서명/도장 이미지 삽입" 참조)
- **hwp2hwpx.bat**: 이 디렉토리의 `convert/hwp2hwpx.bat`
- **python-hwpx CLI**: `pip install python-hwpx` (v2.9.0+): `hwpx-validate`, `hwpx-page-guard` 등

실행 시 이 스킬 디렉토리의 절대경로를 사용한다. 예: `python <스킬디렉토리>/hwpx_edit.py`

## Step 0: 환경 점검 (첫 실행 또는 의존성 의심 시)

```bash
python <스킬디렉토리>/hwpx_edit.py --check-env
```

기능 계층(tier)별로 무엇이 바로 되는지 한 번에 출력한다(읽기 전용, 아무것도 설치하지 않음). 이 스킬은 API 키를 쓰지 않으므로(전부 로컬 도구) 점검 대상은 pip 패키지와 시스템 런타임이다:

- **Tier 1 (읽기·편집, 필수)**: `python-hwpx`·`lxml`·`hwpx-tomd`. `pip install -r requirements.txt` 한 줄이면 충족하며, 대부분의 작업(`--to-md`·텍스트 치환·셀 편집)은 여기까지면 된다.
- **Tier 2 (HWP→HWPX)**: JDK 21 + 번들 JAR. JDK 경로는 `convert/hwp2hwpx.bat`의 `JAVA_HOME`을 단일 소스로 점검하며, 경로가 다르면 그 한 줄만 고치면 된다.
- **Tier 3 (MD/DOCX/HTML→HWPX)**: `pypandoc-hwpx`(+ Pandoc). Pandoc은 `pypandoc-hwpx`가 번들 제공할 수 있어 경고만 떠도 변환이 동작할 수 있다.
- **Tier 4 (PDF·이미지·서명)**: Windows + 한컴오피스 COM(`pywin32`, 선택적 `pyhwpx`). 보안모듈 DLL·레지스트리·한컴 기동까지의 상세 진단은 `--diagnose-com`으로 위임한다.

처음 클론한 환경이거나 특정 워크플로우가 의존성·런타임 누락으로 실패하면 먼저 실행해 무엇을 설치할지 확인한다. 환경이 준비된 뒤에는 매 작업마다 실행할 필요가 없다.

## 의사결정 트리

```
HWP/HWPX 작업 요청
├── 읽기 (내용 파악)
│   ├── HWP 파일 → 먼저 HWPX로 변환 (convert/hwp2hwpx.bat) → 이후 HWPX 읽기 단계로
│   │   ※ HWP 직접 파싱 도구(예: kordoc)는 표/텍스트 박스가 복잡한 출판사
│   │     워크시트·고사지에서 바이너리 잔여 문자 leak, 행 누락, 셀 내용 손실이
│   │     발생하므로 사용하지 않는다 (2026-04-30 검증)
│   └── HWPX 파일
│       ├── 암호화 감지(META-INF/manifest.xml에 encryption-data 존재)
│       │   → 한컴 COM으로 먼저 암호 제거 (reference/encrypted-hwpx.md 참조)
│       ├── 표 셀 안에 긴 지문이 있는 문서(고사지·보고서)
│       │   → hwpx_edit.py --to-md --cell-br (셀 내부 문단을 <br>로 구분)
│       └── 그 외 → hwpx_edit.py --to-md (XML 직접 파싱, 무료, 정확)
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
│       ├── 긴 제목이 한 줄로 겹쳐 뭉개짐(양식의 "한 줄로 입력" 문단) → hwpx_edit.py --fix-squeeze
│       ├── 빈 셀 hp:p 보정 → hwpx_edit.py --fix-empty-cells (save_hwpx 시 자동; Pandoc/hwpx_convert.py 변환물에 반드시 실행)
│       ├── 표/문단/머리글 추가·수정 → python-hwpx API (lineseg 자동 처리)
│       ├── 서명란에 서명/도장 이미지 삽입 → hwpx_sign.py (아래 "서명/도장 이미지 삽입")
│       ├── 구조적 편집 (행 추가/복제) → Python + regex (아래 구조적 편집 규칙 필수)
│       └── 복잡한 편집 → python-hwpx + lxml 직접 사용 (reference/api.md + reference/structural.md 참조)
│   ※ 편집 후 검증: hwpx-validate + hwpx-page-guard
└── 새로 생성 (마크다운 등에서)
    ├── 간단한 문서 (스타일 최소, 빠른 변환)
    │   → hwpx_convert.py + 스타일 후처리 (아래 "MD → HWPX 변환 (Pandoc)" 참조)
    ├── 보고서급 문서 (표/인용문/각주/커스텀 스타일 필요)
    │   → python-hwpx build-from-scratch (아래 "MD → HWPX 생성 (Build-from-scratch)" 참조)
    └── 한컴 충실도 최우선 + 이후 COM 후속 작업(이미지 삽입·PDF) 예정
        → hwpx_com.py --from-md (COM 네이티브 생성, Windows + 한컴오피스 필요)
        ※ COM 생성물만 COM 재편집이 안전. Pandoc 생성물은 COM에 넣지 말 것
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
| 서명란에 서명/도장 이미지 삽입 | `hwpx_sign.py` (COM+XML) | **자동 제거** |
| 무결성 검증, 쪽수 드리프트 감지 | python-hwpx CLI | N/A |

## 출력·정리 규칙 (결과물 vs 부산물)

**원칙**: 도구(`hwpx_edit.py`, `hwp2hwpx.bat`)는 "원본을 덮어쓰지 않는다"만 책임진다. 무엇이 최종 결과물이고 무엇이 부산물인지는 도구가 알 수 없으므로(같은 `--to-md`도 워크플로우에 따라 결과물이거나 부산물이다), **호출하는 워크플로우가 판단**한다.

- **도구 기본 출력**: `-o` 미지정 시 입력 폴더의 `_work-hwpx-automation/`에 저장한다(원본 비파괴).
- **작업 마무리**: 작업 종료 시 **최종 결과물은 작업 폴더로** 옮기고(또는 처음부터 `-o`로 작업 폴더를 지정), 그 전까지의 **중간 부산물은 `_work-hwpx-automation/`에 잔류**시킨다. 작업 폴더에는 원본과 최종 결과물만 남아 깔끔하게 유지된다.

### 시나리오별 결과물/부산물

| 시작점 | 작업 | 최종 결과물(작업 폴더) | 부산물(`_work-hwpx-automation/`) |
|--------|------|----------------------|----------------------------------|
| HWP | 읽기 →md | `<이름>.md` | `<이름>.hwpx`(변환 중간물) |
| HWP | 편집 →HWPX | `<이름>.hwpx`(편집본) | 변환·단계별 중간 HWPX |
| HWPX | 읽기 →md | `<이름>.md` | (없음) |
| HWPX | 편집 | 편집본 HWPX | 단계별 중간본 |
| HWPX/HWP | →PDF | `<이름>.pdf` | (PDF가 최종이면 없음) |
| MD | →HWPX | `<이름>.hwpx` | 빈셀보정 전 중간본 |
| 없음 | build-from-scratch | `<이름>.hwpx` | 템플릿 중간물 |

> 도구를 SKILL 워크플로우 없이 순수 CLI로 단독 사용하면 결과물이 `_work-hwpx-automation/`에 생긴다. 이때는 `-o`로 작업 폴더를 직접 지정하면 된다.

## hwpx_edit.py 사용법

```bash
# HWPX → Markdown 변환 (hwpx-tomd 엔진, 외부 API 불필요, 오프라인 사용 가능)
#  - reading-order 재귀 순회: 글상자(drawText)·그리기 개체 내부 본문까지 수집
#  - <hp:t> 내부 tail 보존: <tab>/<lineBreak>로 구분된 선택지 ②③⑤ 등 누락 없음
#  - 자가검증 3종(엔진이 수행): 단어 recall + 글자 멀티셋 recall(char_recall) +
#    객관식 마커 보존 가드(①②③ 등이 줄면 임계값 무관 경고). 깨끗하면 recall을
#    stdout에 "단어 X% · 글자 Y%"로, 누락 의심 시 경고를 stderr에 출력
#  - 본문 이미지가 있으면 이미지 내 텍스트 누락 가능성을 stderr 경고로 고지
python hwpx_edit.py <파일.hwpx> --to-md [-o output.md]

# 표 셀 안에 긴 지문(일기·본문)이 있는 문서는 --cell-br 권장
#  - 기본(--to-md만): 셀 내 문단을 공백으로 합침
#  - --cell-br:      셀 내 <hp:p> 문단을 <br>로 구분 (고사지·보고서 권장)
python hwpx_edit.py <파일.hwpx> --to-md --cell-br [-o output.md]

# 병합으로 덮인 칸을 시작 칸 값으로 채우기 (행 단위 파싱·LLM 입력용)
#  - 기본은 GFM 정렬 보존(병합 칸은 빈 칸). --merge-fill은 시작 칸 값 복제
python hwpx_edit.py <파일.hwpx> --to-md --merge-fill [-o output.md]

# 암호화된 HWPX (AES-256-CBC)는 hwpx_edit.py가 자동 감지하고 오류 메시지로
# 해제 절차 안내한다. 해제는 한컴 COM으로 FilePasswordChange 액션 사용.
# 상세: reference/encrypted-hwpx.md

# 표 구조 확인 (표 개수, 행/열 수, 셀 내용 미리보기)
python hwpx_edit.py <파일.hwpx> --info

# 텍스트 치환 (본문 + 표 셀 모두 처리)
python hwpx_edit.py <파일.hwpx> --find "이전" --replace "이후"

# 표 셀에 텍스트 입력 (표번호,행번호,셀번호, 0부터 시작)
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

# "한 줄로 입력" 과압축 감지·보정 (양식 채우기 후 긴 제목이 한 줄로 겹쳐 뭉개질 때.
# 사람이 만든 양식의 제목 문단에 문단 모양 "한 줄로 입력"(paraPr breakSetting
# lineWrap="SQUEEZE")이 걸려 있으면 긴 텍스트를 채웠을 때 한글이 줄바꿈 대신
# 장평을 무제한 압축해 글자가 겹친다. hwpx-validate·--to-md로는 탐지 불가,
# 렌더링(PDF)에서만 드러남. --fix-squeeze는 lineWrap="BREAK" 복제 paraPr을 새 id로
# 추가해 과압축 문단만 재지정한다(같은 paraPr을 쓰는 짧은 라벨 디자인은 보존).
# --find/--replace·--set-cell도 채운 텍스트가 과압축되면 stderr로 자동 경고한다)
python hwpx_edit.py <파일.hwpx> --list-squeeze                # 감지만 (변경 없음)
python hwpx_edit.py <파일.hwpx> --fix-squeeze [-o output.hwpx]  # 자연 줄바꿈 전환

# 누락된 Preview/PrvText.txt 생성·주입 (hwp2hwpx 변환물의 hwpx-validate 실패 보정.
# hwpxlib 변환물은 container.xml이 Preview/PrvText.txt를 rootfile로 선언하면서도
# ZIP에 Preview/ 항목을 안 만들어 hwpx-validate가 'Root content ... missing'으로
# 실패한다. 한글은 정상 열림·--to-md recall 100%라 읽기·편집엔 무해하므로, 변환물을
# 정본·배포·검증 대상으로 쓸 때만 실행. section XML은 재직렬화하지 않고 원본 바이트를
# 보존하며 PrvText만 추가한다. -o 미지정 시 _work-hwpx-automation/에 출력)
python hwpx_edit.py <파일.hwpx> --add-preview [-o output.hwpx]

# 별도 파일로 저장
python hwpx_edit.py <파일.hwpx> --set-cell 0,1,0 "텍스트" -o output.hwpx
```

> **`--to-md` 보장과 한계** (2026-06-06 재검증, 엔진=hwpx-tomd 패키지): **텍스트
> 완전성은 보장**한다. 실문서 33종(워크시트·고사 원안·교육과정·평가계획·체크리스트
> 등)에서 원본 `<hp:t>` 대비 글자 멀티셋 손실 0·객관식 마커 손실 0으로 문자·마커
> 단위 완벽 보존을 입증했다(상용 파서 Upstage와 동급 또는 초과). 수정된 결함:
> ① `<hp:t>` tail에 든 선택지 ②③⑤ 누락, ② 글상자(drawText) 본문 대량 누락,
> ③ cellSpan 오독으로 rowSpan·colSpan이 모두 무시되던 표 정렬 붕괴. 표는 이제
> cellAddr/cellSpan 기반 그리드 배치로 세로·가로 병합 시에도 열 정렬을 유지한다.
> 변환 후 자가검증 3종(단어 recall + 글자 멀티셋 recall + 마커 보존 가드)이 조용한
> 누락을 막는다.
> **단 레이아웃은 근사**다: 글상자는 XML anchor 위치(문서 순서)에 삽입되어 시각적
> 배치와 다를 수 있고(예: 빈칸이 지시문보다 먼저), 중첩표(셀 안의 표)는 텍스트로
> 평탄화된다. **이미지 내 텍스트(제목·도표·캡션)는 추출 범위 밖**이며, 본문에
> 이미지가 있으면 stderr 경고로 고지한다(필요 시 OCR 파서 병용). 암호화(AES)
> 배포본은 파싱 불가하며 자동 감지해 안내한다. 시각적 배치 재현이 중요하거나
> 이미지 경고·recall 경고가 나면 OCR·레이아웃 인식 파서(Upstage 등 상용 문서
> 파서)로 교차검증할 것.

## HWP 읽기 (HWPX 변환 경유)

HWP 바이너리 직접 파싱(예: kordoc)은 다음과 같은 손실이 발생하므로 사용하지 않는다:
- 출판사 워크시트·고사지처럼 **중첩 표·텍스트 박스가 있는 문서**에서 바이너리 잔여 문자 leak (수십 자~수백 자)
- 매칭 표나 다열 표에서 **행 통째 누락**, 셀 내용 손실
- 변경 추적(track changes) 처리가 거칠어 원본·수정본 텍스트가 그대로 concatenate

따라서 **HWP는 먼저 HWPX로 변환한 뒤** `hwpx_edit.py --to-md`로 읽는다:

```bash
# 1. HWP → HWPX 변환 (서식 100% 보존)
convert/hwp2hwpx.bat <파일.hwp> <파일.hwpx>

# 2. HWPX → Markdown
python hwpx_edit.py <파일.hwpx> --to-md -o output.md

# 표 셀에 긴 지문이 있는 경우 (고사지·보고서)
python hwpx_edit.py <파일.hwpx> --to-md --cell-br -o output.md
```

### 폴백: hwp2hwpx가 실패하는 HWP (pyhwp 경유)

일부 HWP는 hwp2hwpx(hwplib)가 `java.util.EmptyStackException`으로 죽는다
(`ForInlineControl.fieldEnd` — 표 셀 안 필드 컨트롤의 시작/끝 짝이 안 맞는 문서.
공공기관 안내문에서 실측, 2026-07-06). 이때는 pyhwp로 우회한다:

```bash
pip install pyhwp   # hwp5txt / hwp5proc 제공

# hwp5txt는 표 내용을 <표> 플레이스홀더로 뭉개므로, 본문이 표 안에 있는
# 공문·안내문은 반드시 hwp5proc xml 경유로 셀 텍스트까지 걷는다:
hwp5proc xml <파일.hwp> > full.xml
python convert/hwp_xml_to_md.py full.xml output.md
```

- `convert/hwp_xml_to_md.py`: pyhwp XML에서 표 구조(행='|', 셀 내 문단='/')를
  보존해 텍스트를 추출하는 폴백 파서. 각 텍스트를 정확히 한 번만 출력하도록
  `TableControl` 경계에서 가지치기한다(문단·중첩 표 중복 방지).
- 한계: 서식·이미지 미보존(검색·내용 파악용). 서식 보존이 필요하면 한컴 COM
  (Windows)으로 HWPX 재저장 후 정식 경로로.

> **PDF 읽기**: 이 스킬의 범위가 아니다. 단순 PDF는 일반 텍스트 추출 도구로 읽고, 표·복합 레이아웃은 별도의 PDF/문서 파서(다중 파서 퓨전 도구 등)를 사용한다.

## python-hwpx CLI 도구 (v2.9.0+)

```bash
# XSD 스키마 검증: 편집 후 무결성 확인
hwpx-validate <파일.hwpx>

# ZIP/OPC 패키지 구조 검증 (mimetype, container.xml, manifest)
hwpx-validate-package <파일.hwpx>

# 레퍼런스 대비 페이지 드리프트 감지: 양식 편집 시 필수
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

- 출력 파일 미지정 시 입력 폴더의 `_work-hwpx-automation/` 하위에 `.hwpx`로 생성(원본 비파괴). HWPX가 이후 단계의 입력일 뿐이면 부산물이므로 거기 두고, HWP→HWPX 변환 자체가 목적이면 작업 폴더로 옮긴다(아래 "작업 마무리" 참조). 출력 경로를 직접 지정하려면 두 번째 인자로 명시
- 서식 100% 보존 (Java 기반, hwplib + hwpxlib)
- 요구사항: JDK 21 (`C:/Program Files/Eclipse Adoptium/jdk-21.0.10.7-hotspot`)
- 입력 경로에 cp949 외 문자(en-dash, em-dash 등)가 있어도 내부에서 `%TEMP%`로 staging해 처리
- ⚠ **Git Bash에서 호출 금지**: `cmd.exe /c`를 거치는 Git Bash에서 한글·공백 경로 인자를 넘기면 cp949 이중 셸 해석으로 깨져(`Exit code 2` + 깨진 바이트) Usage 분기로 빠진다. bat 내부 `%TEMP%` staging은 JVM argv 문제만 막을 뿐 그 앞단 cmd.exe 인자 전달 깨짐은 못 막는다. **PowerShell에서 직접 호출**하거나, 정 Bash가 필요하면 입력을 ASCII 경로 임시 폴더에 복사한 뒤 `java -cp "<convert>/hwp2hwpx-1.0.0.jar;<convert>/lib/hwplib-*.jar;<convert>/lib/hwpxlib-*.jar;<convert>" Hwp2HwpxCLI in.hwp out.hwpx`를 직접 실행한다
- ⚠ **변환물은 `hwpx-validate`가 `Preview/PrvText.txt` 누락으로 실패한다**(container.xml은 선언, ZIP엔 Preview/ 없음). 한글은 정상 열림·`--to-md` recall 100%라 읽기·편집엔 무해. 변환물을 **정본·배포·검증 대상**으로 쓸 때만 `python hwpx_edit.py <파일.hwpx> --add-preview`로 보정한다(section 무변형, 교훈 10).

## 핵심 워크플로우: MD → HWPX 변환 (Pandoc 방식)

> 간단한 문서에 적합. 복잡한 보고서는 아래 "Build-from-scratch 방식" 참조.

1. **사전 요구사항 확인**: `pip install pypandoc-hwpx` (미설치 시 ModuleNotFoundError)
2. **전처리**: (a) 따옴표(`"…"`, `'…'`, `"…"`, `'…'`)는 **`hwpx_convert.py`가 자동 보호한다**(변환 단계에 내장, 2026-06-09 도구 반영). Pandoc HWPX writer가 따옴표 쌍 안 텍스트를 통째 누락시키는 버그를 변환 전 PUA(U+E000~) 치환 → 변환 후 Contents/*.xml에서 원복으로 우회한다(아포스트로피도 안전 왕복). 끄려면 `--no-quote-fix`. **수동 PUA 전처리 더는 불필요.** (b) `> ` blockquote → `【인용】` 마커 치환 또는 일반 문단으로 변환은 **여전히 수동 필수**.
3. **기본 변환**: `python convert/hwpx_convert.py <입력.md> -o <출력.hwpx>` (따옴표 보호 자동 적용)
4. **빈 셀 수정** (필수!): `python hwpx_edit.py <출력.hwpx> --fix-empty-cells`: MD 표의 빈 셀(` | | `)이 HWPX에서 `<hp:subList>`만 있고 `<hp:p>`가 없는 상태로 변환됨. 한글 엔진이 이를 만나면 약 15초 로딩 후 안전 종료함(XSD 스키마는 통과하므로 `hwpx-validate`로는 탐지 불가). `hwpx_convert.py`는 이 보정을 자동 적용하지 않으므로 **변환 직후 반드시 실행**.
5. **스타일 후처리** (raw XML 편집):
   - header.xml: borderFill 추가 (표 배경색), charPr 추가 (굵은/이탤릭), 본문 글꼴 크기 변경
   - section0.xml: PUA 마커 → 스마트 따옴표 복원(`postprocess_quotes`), 표 헤더 행 배경색, 합계 행 배경색, 인용문 좌측 여백+이탤릭, `【인용】` 마커 제거
   - lineseg 전체 제거 (`re.sub(r'<hp:linesegarray>.*?</hp:linesegarray>', '', xml, flags=re.DOTALL)`)
6. **머리글/바닥글**: python-hwpx API (`doc.set_header_text()`, `doc.set_footer_text()`)
7. **검증**: `hwpx-validate <출력.hwpx>`
8. **인쇄용 문서면 디자인 보정** (안내문·가정통신문·고사지·**학생 작성용 양식** 등): 기본 변환물은 균일 180% 줄간격 + 큰 섹션 제목으로 페이지 수가 부풀고 표가 경계에서 잘린다. 줄간격 용도별 차등화(읽는 곳은 조이고 쓰는 곳은 200~215%로 넓힘), 섹션 제목 위계 압축, 용지 위·헤더 여백 축소, 논리 구획 앞 `pageBreak="1"` 삽입, 핵심 평가어 굵게+밑줄 강조를 적용. 학생이 손으로 쓸 작성칸·내용 구획은 표가 아니라 **`hp:rect` 작성칸/구획 박스**로 만든다(재사용 함수 `make_rect_box` 제공) → `reference/conversion.md`의 "인쇄용 배포 문서 디자인" 참조.

**주의**: 스타일 후처리에서 한국어 텍스트 검색 시, `python -c "..."` 터미널 명령은 cp949 인코딩 문제 발생. 반드시 `.py` 파일(UTF-8)로 작성하거나 유니코드 이스케이프(`"\uc778\uc6a9"` = "인용") 사용.

> 상세 코드 패턴 (borderFill/charPr 추가, 표 행 스타일링, 인용문 스타일링 등)은 `reference/conversion.md` 참조.

## 핵심 워크플로우: MD → HWPX 생성 (Build-from-scratch 방식)

> 보고서급 문서에 적합. python-hwpx API로 문서 구조를 처음부터 빌드하고, XML 후처리로 세밀한 스타일링을 적용한다.
> Pandoc 방식의 한계 (따옴표 안 텍스트 누락, blockquote 누락, 각주 미변환, 제한적 스타일)를 완전히 해결.

### 언제 사용하는가

| 조건 | Pandoc 방식 | Build-from-scratch |
|------|------------|-------------------|
| 표 스타일 (헤더 배경, 교대 행, 볼드 셀) | 후처리 필요 | 직접 제어 |
| 따옴표 안 텍스트 (`"…"`, `'…'`) | `hwpx_convert.py` 자동 보호(내장) | 직접 제어 (문제 없음) |
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

**1단계 MD 파싱**: `parse_markdown()` + `parse_inline()`으로 블록/인라인 구조 추출

**2단계 스타일 주입** (`inject_styles`):
- `header.xml`의 `borderFills`, `charProperties`, `paraProperties`에 커스텀 항목 추가
- 빈 템플릿 기본 카운트: charPr 0-6 (7개), borderFill 1-2 (2개), paraPr 0-19 (20개)
- 새 항목 ID = 기본 카운트 + 순서 (결정론적 계산)

**3단계 API 빌드**: `doc.add_paragraph()`, `doc.add_table()`, `tbl.set_cell_text()`

**4단계 XML 후처리**: 표 헤더/교대행/합계행/볼드셀 스타일링, lineseg 전체 제거

**5단계 저장**: `doc.save_to_path()` 또는 ZIP 수동 리패키징

> 상세 코드 패턴, 주요 함정, 디자인 토큰 등은 `reference/build-from-scratch.md` 참조.

## 핵심 워크플로우: 양식 채우기

1. **원본이 HWP이면 변환**: `convert/hwp2hwpx.bat input.hwp output.hwpx`
2. **내용 파악**: `hwpx_edit.py output.hwpx --to-md`
3. **표 구조 확인**: `hwpx_edit.py output.hwpx --info`
4. **(선택) 심층 분석**: `hwpx-analyze-template output.hwpx --json` (스타일 ID 맵 필요 시)
5. **데이터 채우기**: `--set-cell`로 개별 셀 또는 Python 스크립트로 일괄 처리
6. **"한 줄로 입력" 과압축 확인**: 제목·긴 문장을 채웠다면 `--list-squeeze`로 확인, 걸리면 `--fix-squeeze`로 자연 줄바꿈 전환 (양식 제목 문단의 lineWrap="SQUEEZE"가 긴 텍스트를 한 줄로 욱여넣어 글자가 겹침. `--find/--replace`·`--set-cell`이 자동 경고하지만, Python 스크립트로 일괄 채운 경우엔 수동 확인 필요. 상세: reference/warnings-editing.md 13번)
7. **무결성 검증**: `hwpx-validate result.hwpx`
8. **쪽수 드리프트 감지**: `hwpx-page-guard -r output.hwpx -o result.hwpx`
9. **내용 확인**: `--to-md`로 최종 텍스트 검증. 서식이 중요한 문서는 `--to-pdf`로 렌더링 육안 확인(과압축·겹침은 텍스트 검증으로는 안 잡힌다)

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
count = doc.replace_text_in_runs("{{이름}}", "홍길동")
doc.save_to_path("output.hwpx")
```

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

## 구조적 편집 (행/표/단락 추가): raw lxml 필요

> python-hwpx API에 `add_row()`, `insert_row()`가 **없으므로**, 표에 행을 추가/복제하는 경우에만 raw lxml + regex를 사용한다.
> 단순 셀 편집, 텍스트 치환, 표 생성은 위의 API를 사용하라.

### 3가지 필수 규칙

1. **`<hp:linesegarray>` 제거**: 텍스트가 변경된 모든 `<hp:p>`에서 제거. 한글이 열 때 자동 재계산.
   - python-hwpx API 사용 시 **자동 처리됨** (내부에서 `_clear_paragraph_layout_cache()` 호출)
   - raw lxml/regex로 직접 XML을 편집할 때만 수동 제거 필요
   - openhwp Rust 소스로 검증: `line_segments: Option<LineSegmentArray>`(생략이 스키마적으로 유효)

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

## 서명/도장 이미지 삽입 (hwpx_sign.py)

동의서·계약서·서약서 등의 서명란("(서명)" 같은 기준 텍스트)에 자필 서명/도장 이미지를 넣는 작업. **단순 COM `InsertPicture`로는 안 된다**: 다음 3대 함정 때문에 30분씩 걸리던 작업을 `hwpx_sign.py`가 한 번에 처리한다.

**왜 단순 삽입이 실패하는가 (도구가 자동 회피하는 함정)**:
1. COM `InsertPicture`는 그림을 floating(`textWrap="TOP_AND_BOTTOM"`) + `curSz=0,0`으로 넣는다. **서명란이 페이지 하단에 있으면 큰 그림이 인라인으로 그 줄에 못 들어가 다음 페이지로 밀린다.**
2. floating + `vertRelTo="PAPER"` 절대좌표를 줘도 **앵커 문단이 페이지 끝이면 개체가 다음 페이지에 그려진다**(표시 페이지는 앵커 문단의 페이지를 따름). → 앵커를 페이지 내 **위쪽 문단**으로 옮겨야 한다.
3. 텍스트가 바뀐 문단의 `<hp:linesegarray>`(줄 레이아웃 캐시)를 안 지우면 **한컴 COM이 문서 열기를 거부**한다. 치명적으로 `hwpx-validate`(XSD)·`--to-md`(recall)는 모두 통과 → **자동 검증으로 못 잡고 PDF 변환(COM Open)에서만 드러남.**

`hwpx_sign.py`는 (a) COM으로 anchor 자리에 이미지를 넣어 BinData 확보 → (b) XML 후처리로 floating PAPER 절대좌표 전환(크기/curSz 보정) → (c) 앵커를 anchor 문단의 직전 문단으로 이동 → (d) 세로 좌표를 페이지 여백+lineseg vertpos로 자동 계산 → (e) 변경 문단 lineseg 제거를 모두 자동 수행한다.

```bash
# 기본: 서명란 "(서명)"에 서명 이미지를 우측 정렬·서명줄 세로중앙으로 삽입
python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(서명)" --pdf

# 위치 미세조정(한 번 --pdf로 보고 어긋나면 두 값만 바꿔 재실행)
python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(서명)" \
  --width-mm 22 --horz-offset 44000 --vert-adjust -800 --pdf -o 최종.hwpx

# 서명란에 세로 여유가 충분하면 floating 없이 단순 인라인
python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(서명)" --inline
```

옵션: `--width-mm`(기본 20, 세로는 종횡비 자동) · `--horz-offset`(HWPUNIT, 미지정 시 우측정렬) · `--vert-adjust`(+아래/−위) · `--gap-mm`(우측 여백, 기본 7) · `--inline` · `-o`(기본 `<입력>_서명.hwpx`) · `--pdf`(검증 PDF 동시 생성).

**최단 워크플로우**:
1. (필요 시) 이름·날짜 등 텍스트 먼저 채우기: `hwpx_edit.py --find/--replace` 또는 `--set-cell`.
2. 서명 삽입 + 검증: `hwpx_sign.py … --anchor "(서명)" --pdf`. 생성된 PDF를 육안 확인.
3. 위치가 어긋나면 `--horz-offset`/`--vert-adjust`만 바꿔 재실행(좌표는 HWPUNIT, 7200/inch).

**실전 팁·함정**:
- **서명이 기존 텍스트와 겹칠 때**: 서명란 텍스트가 줄 우측에 몰린 양식(예: 선행 공백이 많아 "성 명 : 홍길동"이 우측 끝에 붙는 동의서)에서는 우측정렬 서명이 이름 위를 덮는다. → ① `--horz-offset`으로 서명을 더 우측/이름 밖으로 밀거나, ② 삽입 전에 `hwpx_edit.py`로 서명란 텍스트의 **선행 공백을 줄여 이름을 왼쪽으로 당긴** 뒤 서명을 넣는다.
- **이미지 배경**: 투명/흰 배경 PNG를 쓴다. 회색·불투명 배경 이미지는 서명란에 박스가 비치므로 배경 제거 후 사용.
- **검증은 반드시 PDF로**: `hwpx-validate`/`--to-md`는 lineseg 손상(함정 3)을 못 잡는다. `--pdf`(COM Open)로만 위치·열림을 확인한다.
- **COM 프로세스 잔류**: 연속 실행 중 PDF 변환이 "문서 열기 실패"로 죽으면 `taskkill /F /IM Hwp.exe` 후 재시도.
- **서명 이미지 경로**: 이 저장소에는 서명 이미지를 포함하지 않는다. 사용할 서명/도장 이미지 파일 경로는 호출 시 `--image`로 직접 지정한다(투명/흰 배경 PNG 권장).

## 한컴 COM 자동화 (Windows 전용, 한컴오피스 필수)

한컴오피스가 설치된 Windows 환경에서 `HWPFrame.HwpObject` COM을 통해 HWPX를 조작할 수 있다.
XML 직접 편집으로 불가능한 작업(이미지 삽입, PDF 변환)과 COM 네이티브 문서 생성에 사용.

### 두 파이프라인 분리 원칙 (2026-06-10 신설)

이 스킬은 두 백엔드를 가지며 **섞지 않는다**:

| 파이프라인 | 도구 | 환경 | 적합 작업 |
|-----------|------|------|----------|
| XML/Pandoc | hwpx_edit.py, hwpx_convert.py, python-hwpx | 크로스플랫폼, 한컴 불필요 | 기존 문서 편집, 대량 처리, 서버 |
| COM 네이티브 | **hwpx_com.py** (pyhwpx) | Windows + 한컴오피스 | 신규 생성+COM 후속 작업, 이미지 삽입, PDF |

- **COM 생성 HWPX → XML 파이프라인 읽기·편집**: 안전 (recall 100% 검증, 2026-06-10)
- **Pandoc 생성 HWPX → COM**: 금지 (본문 0자 인식, 같은 파일 저장 시 내용 전체 소실, `reference/warnings-com.md` 4번)
- hwpx_com.py는 입력 파일을 절대 in-place로 덮어쓰지 않는다 (위 사고의 구조적 방지)

### hwpx_com.py 사용법 (pyhwpx 기반)

```bash
# pyhwpx + COM 진단 (기동·종료까지 실측)
python hwpx_com.py --diagnose

# Markdown 부분집합 → COM 네이티브 HWPX 생성
# 지원: 제목 #/##/### (굵게 16/14/12pt), 본문 단락, **굵게** 인라인, 불릿(-/*), 파이프 표
python hwpx_com.py output.hwpx --from-md input.md

# 문서 끝에 이미지 삽입 (별도 파일 <입력>_img.hwpx로 저장, mm 단위 크기 지정 가능)
python hwpx_com.py doc.hwpx --insert-image stamp.png [--image-width 40 --image-height 20] [-o out.hwpx]

# COM 기준 본문 텍스트 추출: Pandoc HWPX 호환성 점검에도 사용 (0자면 비호환)
python hwpx_com.py doc.hwpx --get-text

# 한컴 재저장 정규화: XML 파이프라인(python-hwpx 등) 생성·편집물을 한컴이 직접
# 재저장해 lineseg·미리보기(PrvText/PrvImage)·내부 캐시를 네이티브로 재계산.
# 인쇄·배포 직전 마무리 표준 단계. 본문 자수 보존 자동 검증, in-place 금지.
python hwpx_com.py doc.hwpx --normalize [-o final.hwpx]

# PDF 저장 (hwpx_edit.py --to-pdf와 동일 기능, COM 작업 연속 시 이쪽 사용)
python hwpx_com.py doc.hwpx --to-pdf [-o out.pdf] [--password "암호"]
```

> **pyhwpx 함정**: `get_text_file()` 래퍼는 기본 option이 `saveblock:true`(선택 블록만)라서
> 선택이 없으면 None을 반환한다. 전체 본문은 저수준 `hwp.hwp.GetTextFile("TEXT", "")` 호출
> (hwpx_com.py `--get-text`가 이 방식). 상세: `reference/warnings-com.md` 7번

### 빠른 진단 및 PDF 변환

```bash
# pywin32, 보안 모듈 레지스트리, HWPFrame.HwpObject 생성 가능 여부 확인
python hwpx_edit.py --diagnose-com

# 기본 출력: 원본 폴더의 _work-hwpx-automation/<파일명>.pdf (원본 비파괴)
# PDF가 최종 결과물이면 작업 폴더로 옮기거나 -o로 작업 폴더를 직접 지정
python hwpx_edit.py <파일.hwpx> --to-pdf

# 출력 경로 지정
python hwpx_edit.py <파일.hwpx> --to-pdf -o <파일.pdf>

# 암호화된 HWPX/HWP를 PDF로 저장
python hwpx_edit.py <파일.hwpx> --to-pdf --password "<비밀번호>" -o <파일.pdf>
```

### 보안모듈 (팝업 제거)

한컴 COM으로 파일을 열거나 저장할 때 "접근 허용" 보안 대화상자가 반복 표시된다. 이를 제거하려면:

1. **DLL**: `<repo>/vendor/FilePathCheckerModuleExample.dll`(저장소에 커밋되지 않음, 한컴 독점 라이선스). 설치: `vendor/README.md` 참조
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

```bash
python hwpx_edit.py <파일.hwpx> --to-pdf [-o output.pdf]
```

### 이미지 삽입

```python
# 커서를 원하는 위치로 이동
hwp.HAction.Run('MoveDocEnd')

# 이미지 삽입 (Width/Height 단위: mm)
ctrl = hwp.InsertPicture(abs_img_path, True, 1, False, False, 0, 23, 15)
# params: path, Embedded, sizeoption(1=지정크기), Reverse, Watermark, Effect, Width_mm, Height_mm
```

### 이미지 위치/배치 변경: XML 후처리 방식 (권장)

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
| 스킬 업데이트 (외부 의존성) | `reference/update-checklist.md` | python-hwpx 업데이트, GitHub 인사이트 수집, 반영 기준 |
| 암호화 HWPX 해제 | `reference/encrypted-hwpx.md` | 암호화 감지, 한컴 COM 암호 해제 전체 스크립트, 주의사항 |
| 한컴 COM 액션 테이블 | `reference/action-table.md` | 모든 HAction ID와 대응 ParameterSet ID (공식, 2025.04). Gemini 파싱, "한글"→"호글" 오인식 있으나 API명 검색에 무영향 |
| 한컴 COM API 가이드 | `reference/hwp-automation.md` | IHwpObject 전체 API: 프로퍼티, 메서드, 이벤트, ParameterSet 상세 (공식, 2025.04). Gemini 파싱, 다이어그램은 텍스트 설명으로 대체됨 |
| 한컴 COM 파라미터셋 | `reference/parameterset-table.md` | 149개 ParameterSet 필드/타입/기본값 (공식, 2025.04). Mistral 파싱, 워터마크 환각 잔여 가능, heading이 평문 처리된 경우 있음 |

## 주의사항

워크플로우별 주의사항은 각 reference 파일에 포함:
- **XML 편집 전반**: `reference/warnings-editing.md`: find/replace 동작, set-cell 체이닝, 서식 보존, PrvText.txt 누락, manifest self-closing 등
- **MD → HWPX 변환**: `reference/conversion.md` 하단: blockquote 누락, 각주 변환, 인코딩, 스크립트 보존
- **Build-from-scratch**: `reference/build-from-scratch.md` 하단: lineSpacing 단위, fontRef, set_cell_text, style_tables_xml 루프
- **한컴 COM 자동화**: `reference/warnings-com.md`: 이미지 XML 삽입 불가, 양식+이미지 워크플로우, PDF 변환

## 개선 반영 (함정·패턴 환류)

이 스킬은 다양한 실제 hwpx 작업(파싱·편집·생성·변환)에서 반복 사용되며 검증·진화하는 GitHub 공개 자산(Engccer/hwpx-automation)이다. 완성도 요구가 높은 작업일수록 검증 가치가 크다. 실사용 중 발견한, 일반화 가치가 있는 함정·우회·패턴은 SKILL.md 본문(의사결정 트리·도구 용도 표·주의사항)이나 해당 `reference/*.md`로 직접 반영한다.

**반영 트리거** (하나라도 해당하면 본문 또는 reference로 반영):
- 도구(`hwpx_edit.py`·python-hwpx)로 안 돼서 직접 zipfile·lxml·COM으로 우회한 경우
- 예상과 다른 동작(파싱 누락, 치환 실패, 구조 깨짐, 무한 로딩 등)
- 새 문서 유형에서의 실측 결과
- 검증(`hwpx-validate` / `--to-md` self-recall)으로 잡아낸 결함

**반영 경로**:
1. 같은 함정이 반복되거나 데이터 손실·문서 손상을 막는 패턴이면, SKILL.md 본문(의사결정 트리·도구 용도 표·주의사항)이나 해당 `reference/*.md`로 반영한다.
2. 우회가 반복되면 `hwpx_edit.py`에 옵션·기능으로 흡수할지 검토한다(예: hp:t 분할 치환, 문단째 제거).
3. **변환 자체의 결함**(파싱 누락·표 정렬·마커 손실 등 `--to-md`/self-recall로 드러나는 출력 오류)은 `hwpx_edit.py`가 아니라 **변환 엔진 `hwpx-tomd`** 소관이다(`hwpx_edit.py --to-md`는 이 엔진을 호출만 한다). 따라서 변환 정확도 문제는 이 스킬이 아니라 같은 엔진을 고쳐야 한다. 엔진은 PyPI·GitHub(Engccer/hwpx-tomd)로 공개돼 있고 **GitHub가 단일 진실 원천(SSoT)**이다. 개선 경로는 **누가 실행하느냐**로 갈린다(엔진 repo의 `CONTRIBUTING.md`가 정본):
   - **유지보수자(엔진을 editable git 체크아웃으로 보유, `pip install -e`)**: `hwpx-tomd`의 `tests/test_hwpx_tomd.py`에 **실패하는 회귀 테스트를 먼저 추가** → `core.py` 수정 → 버전(`_version.py`) bump → `git push` → PyPI publish. editable이라 수정이 양쪽 스킬에 즉시 라이브 전파된다.
   - **다운스트림 사용자(`pip install hwpx-tomd`로 설치)**: site-packages를 직접 고치면 재설치 때 사라지고 upstream에도 반영되지 않으며 push 권한도 없다. 대신 **최소 재현 HWPX(또는 그 구조)와 기대·실제 마크다운을 첨부해 `github.com/Engccer/hwpx-tomd`에 이슈를 열거나, 실패 테스트+`core.py` 패치로 PR**을 보낸다(엔진 개선이 반영되는 유일한 경로).

   (편집·생성·CLI 함정은 1·2번대로 이 스킬에 남기고, **변환 함정만 엔진으로** 보낸다.)

> **구분**: 외부 의존성 최신화(python-hwpx·hwplib 버전, GitHub 인사이트)는 `update-checklist.md`가 담당한다. 변환 함정은 위 3번 기준으로 `hwpx-tomd` 엔진으로 라우팅한다.
