---
name: hwpx-automation
description: "HWP/HWPX 문서 읽기, 변환, 편집을 위한 통합 워크플로우. HWP 또는 HWPX 파일을 다룰 때 사용. HWP 파일은 모두 HWPX로 변환 후 처리한다. 트리거: (1) HWP/HWPX 파일 읽기/파싱 요청 (2) HWP→HWPX 변환 요청 (3) HWPX 문서 편집(텍스트 치환, 표 셀 채우기, 양식 작성) (4) 한글 문서 템플릿 기반 자동화 작업 (5) HWPX 구조적 편집(행/표/단락 추가) (6) HWPX에 이미지 삽입 (7) HWPX→PDF 변환 (8) 한컴 COM 자동화 (9) HWPX 서명란에 서명·도장 이미지 삽입(signature/seal/도장 삽입, 동의서·계약서·서약서 서명)"
metadata:
    version: "1.3.0"
---

# HWP/HWPX 작업 자동화 스킬

## 도구 위치

이 SKILL.md와 같은 디렉토리에 모든 도구가 포함되어 있다:
- **hwpx_edit.py**: 이 디렉토리의 `hwpx_edit.py` (편집 명령 + `--to-md` CLI 래퍼)
- **hwpx-tomd**: `--to-md` 변환 엔진을 단일 소스로 보유한 독립 패키지. `pip install hwpx-tomd` (PyPI·GitHub `Engccer/hwpx-tomd` 공개. 엔진 자체를 수정할 때만 로컬 editable: `pip install -e path/to/hwpx-tomd`). 라이브러리로도 직접 사용 가능(`from hwpx_tomd import to_markdown, convert`). 변환 로직은 이 패키지에만 있고 `hwpx_edit.py`는 호출만 한다(코드 분기 방지).
- **hwpx_convert.py**: 이 디렉토리의 `convert/hwpx_convert.py` (MD/DOCX/HTML/RST/TEX/TXT → HWPX 변환, `pip install pypandoc-hwpx` 필요)
- **hwpx_com.py**: 이 디렉토리의 `hwpx_com.py` (한컴 COM 네이티브 파이프라인, pyhwpx 기반, Windows + 한컴오피스 전용, `pip install pyhwpx` 필요). MD→HWPX 생성·HWP→HWPX 변환·이미지 삽입·본문 추출·한컴 재저장 정규화(--normalize)·PDF 변환 (아래 "한컴 COM 자동화" 참조)
- **PDF 변경 추적 경고 자동 처리**: 두 CLI의 `--to-pdf`는 `pdf_export.py`를 공유한다. PDF 저장 구간에서만 변경 추적 형식 경고에 저장으로 응답하고 원래 메시지 모드를 복원한다. 원본 이력은 수정하지 않는다. 상세와 적용 범위는 `reference/warnings-com.md` 15번 참조.
- **rhwp_pdf.py**: Windows가 아닌 OS에서 `hwpx_edit.py --to-pdf`가 쓰는 PDF 백엔드. 오픈소스 조판 엔진 rhwp CLI를 호출하고, 글꼴에 없어 빈 네모로 찍힌 문자를 PDF에서 찾아 보고한다(아래 "macOS·Linux에서 PDF 변환 (rhwp)" 참조)
- **hwpx_sign.py**: 이 디렉토리의 `hwpx_sign.py` (서명란(기준 텍스트 "(서명)" 등)에 서명/도장 이미지를 삽입하는 전용 도구, Windows + 한컴오피스 필요. 아래 "서명/도장 이미지 삽입" 참조)
- **hwp2hwpx**: 이 디렉토리의 `convert/hwp2hwpx.bat`(Windows) 또는 `convert/hwp2hwpx.sh`(macOS/Linux)
- **python-hwpx CLI**: `pip install python-hwpx`에 포함: `hwpx-validate`, `hwpx-page-guard` 등. 머신마다 설치본 버전이 다르다(3.x·6.x). `hwpx_edit.py`는 양쪽을 지원하고, API를 직접 부를 때만 이름 차이를 확인한다(`reference/api.md`)

실행 시 이 스킬 디렉토리의 절대경로를 사용한다. 예: `python <스킬디렉토리>/hwpx_edit.py`

## Step 0: 환경 점검 (첫 실행 또는 의존성 의심 시)

```bash
python <스킬디렉토리>/hwpx_edit.py --check-env
```

기능 계층(tier)별로 무엇이 바로 되는지 한 번에 출력한다(읽기 전용, 아무것도 설치하지 않음). 이 스킬은 API 키를 쓰지 않으므로(전부 로컬 도구) 점검 대상은 pip 패키지와 시스템 런타임이다:

- **Tier 1 (읽기·편집, 필수)**: `python-hwpx`·`lxml`·`hwpx-tomd`. `pip install -r requirements.txt` 한 줄이면 충족하며, 대부분의 작업(`--to-md`·텍스트 치환·셀 편집)은 여기까지면 된다.
- **Tier 2 (HWP→HWPX)**: JDK 21 + 번들 JAR. 래퍼는 Windows `convert/hwp2hwpx.bat`, macOS/Linux `convert/hwp2hwpx.sh`이며 `JAVA_HOME` → (Windows는 Adoptium 표준 설치 →) PATH 순으로 java를 자동 탐색한다. `--check-env`도 같은 순서로 찾는다.
- **Tier 3 (MD/DOCX/HTML→HWPX)**: `pypandoc-hwpx`(+ Pandoc). Pandoc은 `pypandoc-hwpx`가 번들 제공할 수 있어 경고만 떠도 변환이 동작할 수 있다.
- **Tier 4 (PDF·이미지·서명)**: Windows + 한컴오피스 COM(`pywin32`, 선택적 `pyhwpx`). 보안모듈 DLL·레지스트리·한컴 기동까지의 상세 진단은 `--diagnose-com`으로 위임한다. macOS·Linux는 PDF만 `rhwp` CLI로 되고 이미지 삽입·서명·한컴 정규화는 안 된다.

처음 클론한 환경이거나 특정 워크플로우가 의존성·런타임 누락으로 실패하면 먼저 실행해 무엇을 설치할지 확인한다. 환경이 준비된 뒤에는 매 작업마다 실행할 필요가 없다.

## 의사결정 트리

```
HWP/HWPX 작업 요청
├── 읽기 (내용 파악)
│   ├── HWP 파일 → 먼저 HWPX로 변환 (아래 "HWP → HWPX 변환": 셸 규칙·JDK 없는 Windows 폴백)
│   │   → 이후 HWPX 읽기 단계로
│   │   ※ HWP 직접 파싱 도구(예: kordoc)는 사용하지 않는다(같은 절 참조)
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
│     (hwp2hwpx 변환물은 PrvText 누락으로 hwpx-validate가 실패한다. 아래 "HWP → HWPX 변환")
├── PDF 만들기 → hwpx_edit.py --to-pdf
│   (Windows는 한컴 COM, macOS·Linux는 rhwp. 아래 "macOS·Linux에서 PDF 변환 (rhwp)")
│   ※ Pandoc(hwpx_convert.py) 변환물은 Windows COM으로 PDF를 내지 않는다(아래 줄).
│     한컴 PDF가 필요하면 hwpx_com.py --from-md로 COM에서 생성한다
└── 새로 생성 (마크다운 등에서)
    ├── 간단한 문서 (스타일 최소, 빠른 변환)
    │   → hwpx_convert.py + 스타일 후처리 (아래 "MD → HWPX 변환 (Pandoc)" 참조)
    ├── 보고서급 문서 (표 스타일/인용문/커스텀 스타일 필요)
    │   → python-hwpx build-from-scratch (아래 "MD → HWPX 생성 (Build-from-scratch)" 참조)
    └── 한컴 충실도 최우선 + 이후 COM 후속 작업(이미지 삽입·정규화·PDF) 예정
        → hwpx_com.py --from-md (COM 네이티브 생성, Windows + 한컴오피스 필요)
        ※ COM은 Pandoc 생성물의 본문을 0자로 읽는다. Pandoc 생성물은 COM으로 처리하지
          않는다(--normalize·--insert-image·Windows --to-pdf·hwpx_sign.py 모두. --get-text
          호환성 점검만 예외, reference/warnings-com.md 4번). COM 생성물은 XML 파이프라인에서
          읽고 편집해도 안전하다
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
| HWPX/HWP → PDF | `hwpx_edit.py --to-pdf` (Windows 한컴 COM · 그 밖의 OS rhwp) | N/A |
| 무결성 검증, 쪽수 드리프트 감지 | python-hwpx CLI | N/A |

`--delete-rows`·`--delete-empty-rows`는 행을 지운 뒤 표의 `rowCnt`와 `rowAddr`를 다시 맞춘다(아래 "구조적 편집"의 필수 규칙 2·3). 지울 행이 세로 병합에 걸려 있으면 지우지 않고 멈추므로, 먼저 `--split-cell`로 병합을 푼다.

## 출력·정리 규칙 (결과물 vs 부산물)

**원칙**: 도구는 "원본을 덮어쓰지 않는다"만 책임진다. 무엇이 최종 결과물이고 무엇이 부산물인지는 도구가 알 수 없으므로(같은 `--to-md`도 워크플로우에 따라 결과물이거나 부산물이다), **호출하는 워크플로우가 판단**한다.

- **도구 기본 출력**: `hwpx_edit.py`와 hwp2hwpx 래퍼(`.bat`·`.sh`)는 `-o`(래퍼는 두 번째 인자) 미지정 시 입력 폴더의 `_work-hwpx-automation/`에 저장한다(원본 비파괴). 입력이 이미 `_work-hwpx-automation/` 안에 있으면 같은 폴더의 같은 이름, 즉 그 파일 자리에 쓴다.
- **입력 옆에 쓰는 도구**: `hwpx_com.py`는 `<입력>.pdf`·`<입력>_img.hwpx`·`<입력>_norm.hwpx`, `--from-hwp`는 `<입력>.hwpx`, `hwpx_sign.py`는 `<입력>_서명.hwpx`를 입력 파일 옆에 만든다. 부산물로 둘 것이면 `-o`로 `_work-hwpx-automation/`을 지정한다.
- **여러 편집을 이어 할 때**: 편집 명령은 매번 입력 파일을 새로 읽는다. 원본에 `--set-cell` 등을 여러 번 부르면 마지막 호출만 남으므로, 두 번째 호출부터는 앞 호출의 결과(`_work-hwpx-automation/<이름>.hwpx`)를 입력으로 준다. 셀이 많으면 스크립트 한 번으로 처리한다.
- **작업 마무리**: 작업 종료 시 **최종 결과물은 작업 폴더로** 옮기고(또는 처음부터 `-o`로 작업 폴더를 지정), 그 전까지의 **중간 부산물은 `_work-hwpx-automation/`에 잔류**시킨다. 작업 폴더에는 원본과 최종 결과물만 남아 깔끔하게 유지된다. 최종본 이름이 원본과 같으면 옮길 때 원본을 덮으므로 새 이름을 준다.
- **비교 기준 남기기**: `_work-hwpx-automation/` 안 파일을 입력으로 주면 그 자리를 덮으므로, 편집 전 파일을 `hwpx-page-guard --reference`의 기준으로 쓰려면 첫 편집에 `-o`로 새 이름을 준다.

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
# 두 번째 셀부터는 앞 결과를 입력으로 (위 "여러 편집을 이어 할 때")
python hwpx_edit.py <폴더>/_work-hwpx-automation/<파일.hwpx> --set-cell 0,1,1 "텍스트"

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

# 빈 표 셀에 기본 hp:p 삽입 (Pandoc/hwpx_convert.py 변환물은 한글이 15초 로딩 후
# 닫힌다. hwpx-validate로는 탐지 불가. hwpx_edit.py의 저장 경로에서는 자동 적용되지만,
# 외부 도구로 만든 HWPX는 이 명령을 단독 실행해야 함. 원인: reference/conversion.md
# "필수 후처리: 빈 셀 수정")
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

> **`--to-md` 보장과 한계** (엔진=hwpx-tomd 패키지): **텍스트 완전성은 보장**한다.
> 표 밖 본문 문단·글상자·각주·따옴표 안 글자까지 원본 `<hp:t>`의 글자와 객관식
> 마커를 빠짐없이 옮긴다. 표는 cellAddr/cellSpan 기반 그리드 배치로 세로·가로 병합
> 시에도 열 정렬을 유지한다. 변환 후 자가검증 3종(단어 recall + 글자 멀티셋 recall +
> 마커 보존 가드)이 조용한 누락을 막는다. → 사례
> **단 레이아웃은 근사**다: 글상자는 XML anchor 위치(문서 순서)에 삽입되어 시각적
> 배치와 다를 수 있고(예: 빈칸이 지시문보다 먼저), 중첩표(셀 안의 표)는 텍스트로
> 평탄화된다. **이미지 내 텍스트(제목·도표·캡션)는 추출 범위 밖**이며, 본문에
> 이미지가 있으면 stderr 경고로 고지한다(필요 시 OCR 파서 병용). 암호화(AES)
> 배포본은 파싱 불가하며 자동 감지해 안내한다. 시각적 배치 재현이 중요하거나
> 이미지 경고·recall 경고가 나면 OCR·레이아웃 인식 파서(Upstage 등 상용 문서
> 파서)로 교차검증할 것. 변환 결과가 틀리면 이 스킬이 아니라 엔진 `hwpx-tomd`를
> 고친다(아래 "개선 반영").

> **PDF 읽기**: 이 스킬의 범위가 아니다. 단순 PDF는 일반 텍스트 추출 도구로 읽고, 표·복합 레이아웃은 별도의 PDF/문서 파서(다중 파서 퓨전 도구 등)를 사용한다.

## python-hwpx CLI 도구

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

HWP는 먼저 HWPX로 변환한 뒤 `hwpx_edit.py --to-md`로 읽거나 편집한다. HWP 바이너리 직접 파싱(예: kordoc)은 사용하지 않는다: 중첩 표·텍스트 박스가 있는 문서(출판사 워크시트·고사지)에서 바이너리 잔여 문자가 새고(수십~수백 자), 매칭 표·다열 표의 행이 통째 누락되며, 변경 추적이 원본·수정본 텍스트를 이어 붙인다. → 사례

```bash
convert/hwp2hwpx.bat <입력.hwp> [출력.hwpx]        # Windows (PowerShell에서 호출)
bash convert/hwp2hwpx.sh <입력.hwp> [출력.hwpx]    # macOS/Linux
```

- 출력 파일 미지정 시 입력 폴더의 `_work-hwpx-automation/` 하위에 `.hwpx`로 생성(원본 비파괴). HWPX가 이후 단계의 입력일 뿐이면 부산물이므로 거기 두고, HWP→HWPX 변환 자체가 목적이면 작업 폴더로 옮긴다(위 "작업 마무리" 참조). 출력 경로를 직접 지정하려면 두 번째 인자로 명시
- 서식 100% 보존 (Java 기반, hwplib + hwpxlib)
- 요구사항: JDK 21. 두 래퍼 모두 `JAVA_HOME` → (Windows는 Adoptium 표준 설치 →) PATH 순으로 java를 자동 탐색
- **JDK 없는 Windows 폴백**: 한컴오피스가 있으면 `python hwpx_com.py <입력.hwp> --from-hwp [-o 출력.hwpx]`로 COM 변환(변환 후 본문 글자 수 재검증 내장). JDK 설치 전까지의 우회이며, 연속·배치 호출 전에는 잔류 `Hwp.exe`를 정리할 것(`reference/warnings-com.md` 9번)
- Windows(.bat): 입력 경로에 cp949 외 문자(en-dash, em-dash 등)가 있어도 내부에서 `%TEMP%`로 staging해 처리. macOS/Linux(.sh)는 JVM이 UTF-8 argv를 그대로 받으므로 staging 불필요
- ⚠ **Git Bash에서 호출 금지**: Windows에서 Claude Code의 Bash 도구는 Git Bash이므로 이 셸에서는 래퍼 대신 아래 절차(ASCII 경로로 복사한 뒤 `java -cp`)를 쓴다. `cmd.exe /c`를 거치는 Git Bash에서 한글·공백 경로 인자를 넘기면 cp949 이중 셸 해석으로 깨져(`Exit code 2` + 깨진 바이트) Usage 분기로 빠진다. bat 내부 `%TEMP%` staging은 JVM argv 문제만 막을 뿐 그 앞단 cmd.exe 인자 전달 깨짐은 못 막는다. **PowerShell에서 직접 호출**하거나, 정 Bash가 필요하면 입력을 ASCII 경로 임시 폴더에 복사한 뒤 `java -cp "<convert>/*;<convert>/lib/*;<convert>" Hwp2HwpxCLI in.hwp out.hwpx`를 직접 실행한다(`<convert>`는 이 스킬의 `convert` 폴더. Java 클래스패스는 `폴더/*`만 JAR 와일드카드로 받는다)
- ⚠ **변환물은 `hwpx-validate`가 `Preview/PrvText.txt` 누락으로 실패한다**(container.xml은 선언, ZIP엔 Preview/ 없음). 한글은 정상 열림·`--to-md` recall 100%라 읽기·편집엔 무해. 변환물을 **정본·배포·검증 대상**으로 쓸 때만 `python hwpx_edit.py <파일.hwpx> --add-preview`로 보정한다(section 무변형). `--find/--replace`는 이 누락으로 여는 것을 거부하며, 그때 출력하는 두 단계(`--add-preview` 후 보정본으로 재실행)를 따른다.

### 번들 JAR 출처·폴백 (hwp2hwpx가 예외로 죽을 때)

hwp2hwpx가 `EmptyStackException`·`IndexOutOfBoundsException` 등 예외로 죽거나 번들 JAR을 재빌드할 때 `reference/hwp-conversion.md`를 읽는다.

## 핵심 워크플로우: MD → HWPX 변환 (Pandoc 방식)

MD 등을 `hwpx_convert.py`로 변환할 때 `reference/conversion.md`의 같은 제목 절을 읽는다.

## 핵심 워크플로우: MD → HWPX 생성 (Build-from-scratch 방식)

보고서급 문서를 python-hwpx로 처음부터 만들 때 `reference/build-from-scratch.md`의 같은 제목 절을 읽는다.

## 핵심 워크플로우: 양식 채우기

1. **원본이 HWP이면 변환**: 위 "HWP → HWPX 변환"(Windows는 PowerShell에서 bat, 또는 `java -cp` 직접)
2. **내용 파악**: `hwpx_edit.py output.hwpx --to-md`
3. **표 구조 확인**: `hwpx_edit.py output.hwpx --info`
4. **(선택) 심층 분석**: `hwpx-analyze-template output.hwpx --json` (스타일 ID 맵 필요 시)
5. **데이터 채우기**: `--set-cell`로 개별 셀(두 번째부터 앞 결과를 입력으로, 위 "여러 편집을 이어 할 때") 또는 Python 스크립트로 일괄 처리
6. **"한 줄로 입력" 과압축 확인**: 제목·긴 문장을 채웠다면 `--list-squeeze`로 확인, 걸리면 `--fix-squeeze`로 자연 줄바꿈 전환 (양식 제목 문단의 lineWrap="SQUEEZE"가 긴 텍스트를 한 줄로 욱여넣어 글자가 겹침. `--find/--replace`·`--set-cell`이 자동 경고하지만, Python 스크립트로 일괄 채운 경우엔 수동 확인 필요. 상세: reference/warnings-editing.md 13번)
7. **무결성 검증**: `hwpx-validate result.hwpx` (HWP 변환물이면 PrvText 누락 실패는 무해, 위 "HWP → HWPX 변환")
8. **쪽수 드리프트 감지**: `hwpx-page-guard --reference output.hwpx --output result.hwpx`
9. **내용 확인**: `--to-md`로 최종 텍스트 검증. 서식이 중요한 문서는 `--to-pdf`로 렌더링 육안 확인(과압축·겹침은 텍스트 검증으로는 안 잡힌다)

## python-hwpx API: 주요 패턴

CLI로 안 되는 편집을 python-hwpx 스크립트로 직접 짤 때 `reference/api.md`를 읽는다.

## 구조적 편집 (행/표/단락 추가): raw lxml 필요

표에 행을 추가·복제하거나 raw XML로 문단을 고칠 때 `reference/structural.md`를 읽는다. 텍스트를 바꾼 문단의 `<hp:linesegarray>` 제거, `rowAddr` 순차, `rowCnt` 일치, 표 `id` 고유가 필수 규칙이다.

## HWPX 파일 구조

ZIP 내부 구조·네임스페이스가 필요할 때 `reference/format.md`를 읽는다.

## 서명/도장 이미지 삽입 (hwpx_sign.py)

서명란에 서명·도장 이미지를 넣을 때 `reference/signing.md`를 읽는다.

## macOS·Linux에서 PDF 변환 (rhwp)

Windows가 아닌 OS에서 `--to-pdf`를 쓰거나 rhwp 설치·출력 해석(`글꼴에 없는 문자가 빈 네모로 찍혔습니다`, exit 2)이 필요할 때 `reference/rhwp-pdf.md`를 읽는다. exit 2가 난 PDF는 배포하지 않는다.

## 한컴 COM 자동화 (Windows 전용, 한컴오피스 필수)

Windows + 한컴오피스에서 COM(`hwpx_com.py`·보안모듈·암호 해제·이미지 삽입)을 쓸 때 `reference/com.md`를 읽는다. SSH 세션에서는 COM 서버가 뜨지 않는다(같은 문서 "원격 실행").

## 상세 레퍼런스

규칙 끝의 "→ 사례"는 `reference/cases.md`의 같은 제목 절(근거가 된 실측·경위)을 가리킨다. 작업에는 필요 없다.

| 주제 | 파일 | 내용 |
|------|------|------|
| python-hwpx API 전체 | `reference/api.md` | 텍스트 치환, 표 API, 부분 서식, 머리글/바닥글, 서식 검색, 검증, 이미지·도형 등 |
| MD → HWPX 변환 (Pandoc) | `reference/conversion.md` | 변환 워크플로우, hwpx_convert.py 사용법, blockquote 전처리, 스타일 후처리 코드 (borderFill/charPr/표 스타일링/인용문), 인쇄용 디자인 |
| MD → HWPX 생성 (Build) | `reference/build-from-scratch.md` | 5단계 워크플로우, 디자인 토큰, precompute_styles 패턴, 주요 함정 (lineSpacing/fontRef/set_cell_text) |
| 구조적 편집 코드 | `reference/structural.md` | lineseg 원리, 필수 규칙, 행 추가/단락 복제 코드, ZIP 패키징, XML 직렬화 |
| HWPX 파일 구조 | `reference/format.md` | ZIP 내부 구조, section0.xml 요소, 네임스페이스 딕셔너리 |
| HWP 변환 폴백 | `reference/hwp-conversion.md` | 번들 JAR 출처·패치, hwp2hwpx 예외별 폴백(한컴 COM·pyhwp) |
| 서명/도장 삽입 | `reference/signing.md` | hwpx_sign.py 옵션, 앵커 대체 vs 겹쳐 찍기, 검증, 함정 |
| rhwp PDF | `reference/rhwp-pdf.md` | macOS·Linux PDF 설치·충실도·출력 해석 |
| 한컴 COM 자동화 | `reference/com.md` | 파이프라인 분리, hwpx_com.py, 보안모듈, 원격 실행, 기본 패턴, 이미지 삽입 |
| 스킬 업데이트·개선 반영 | `reference/update-checklist.md` | python-hwpx 업데이트, GitHub 인사이트 수집, 함정 환류 경로 |
| 암호화 HWPX 해제 | `reference/encrypted-hwpx.md` | 암호화 감지, 한컴 COM 암호 해제 전체 스크립트, 주의사항 |
| 한컴 COM 액션 테이블 | `reference/action-table.md` | 모든 HAction ID와 대응 ParameterSet ID (공식, 2025.04). Gemini 파싱, "한글"→"호글" 오인식 있으나 API명 검색에 무영향 |
| 한컴 COM API 가이드 | `reference/hwp-automation.md` | IHwpObject 전체 API: 프로퍼티, 메서드, 이벤트, ParameterSet 상세 (공식, 2025.04). Gemini 파싱, 다이어그램은 텍스트 설명으로 대체됨 |
| 한컴 COM 파라미터셋 | `reference/parameterset-table.md` | 149개 ParameterSet 필드/타입/기본값 (공식, 2025.04). Mistral 파싱, 워터마크 환각 잔여 가능, heading이 평문 처리된 경우 있음 |

> 위 표의 마지막 3개 파일(`action-table.md`, `hwp-automation.md`, `parameterset-table.md`)은 한컴 공식 문서의 파싱본이라 저작권 고려로 저장소에 포함하지 않는다. 필요 시 README "참조 문서" 절의 안내에 따라 한컴 개발자 센터 아카이브의 원본 PDF를 내려받아 직접 파싱해 재생성한다.

## 주의사항

워크플로우별 주의사항은 각 reference 파일에 포함:
- **XML 편집 전반**: `reference/warnings-editing.md`: find/replace 동작, set-cell 이어 부르기, 서식 보존, PrvText.txt 누락, manifest self-closing, ZIP·직렬화 규칙 등
- **MD → HWPX 변환**: `reference/conversion.md` 하단 "주의사항": blockquote 누락, 빈 표 셀, 인코딩, 스크립트 보존
- **Build-from-scratch**: `reference/build-from-scratch.md` "주요 함정 (Gotchas)": lineSpacing 단위, fontRef, set_cell_text, style_tables_xml 루프
- **한컴 COM 자동화**: `reference/warnings-com.md`: 이미지 XML 삽입 불가, 양식+이미지 워크플로우, PDF 변환

## 개선 반영 (함정·패턴 환류)

도구로 안 돼서 우회했거나 예상과 다른 동작·새 함정을 만났을 때 `reference/update-checklist.md`의 같은 제목 절을 읽는다.
