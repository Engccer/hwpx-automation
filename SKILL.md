---
name: hwpx-automation
description: "HWP/HWPX 문서 읽기, 변환, 편집을 위한 통합 워크플로우. HWP 또는 HWPX 파일을 다룰 때 사용. HWP 파일은 모두 HWPX로 변환 후 처리한다. 트리거: (1) HWP/HWPX 파일 읽기/파싱 요청 (2) HWP→HWPX 변환 요청 (3) HWPX 문서 편집(텍스트 치환, 표 셀 채우기, 양식 작성) (4) 한글 문서 템플릿 기반 자동화 작업 (5) HWPX 구조적 편집(행/표/단락 추가) (6) HWPX에 이미지 삽입 (7) HWPX→PDF 변환 (8) 한컴 COM 자동화 (9) HWPX 서명란에 서명·도장 이미지 삽입(signature/seal/도장 삽입, 동의서·계약서·서약서 서명)"
metadata:
    version: "1.2.0"
---

# HWP/HWPX 작업 자동화 스킬

## 도구 위치

이 SKILL.md와 같은 디렉토리에 모든 도구가 포함되어 있다:
- **hwpx_edit.py**: 이 디렉토리의 `hwpx_edit.py` (편집 명령 + `--to-md` CLI 래퍼)
- **hwpx-tomd**: `--to-md` 변환 엔진을 단일 소스로 보유한 독립 패키지. `pip install hwpx-tomd` (PyPI·GitHub `Engccer/hwpx-tomd` 공개. 엔진 자체를 수정할 때만 로컬 editable: `pip install -e path/to/hwpx-tomd`). 라이브러리로도 직접 사용 가능(`from hwpx_tomd import to_markdown, convert`). 변환 로직은 이 패키지에만 있고 `hwpx_edit.py`는 호출만 한다(코드 분기 방지).
- **hwpx_convert.py**: 이 디렉토리의 `convert/hwpx_convert.py` (MD/DOCX/HTML/RST/TEX/TXT → HWPX 변환, `pip install pypandoc-hwpx` 필요)
- **hwpx_com.py**: 이 디렉토리의 `hwpx_com.py` (한컴 COM 네이티브 파이프라인, pyhwpx 기반, Windows + 한컴오피스 전용, `pip install pyhwpx` 필요). MD→HWPX 생성·이미지 삽입·본문 추출·한컴 재저장 정규화(--normalize)·PDF 변환. XML/Pandoc 파이프라인과 **분리 운용**(아래 "한컴 COM 자동화" 참조)
- **PDF 변경 추적 경고 자동 처리**: 두 CLI의 `--to-pdf`는 `pdf_export.py`를 공유한다. PDF 저장 구간에서만 변경 추적 형식 경고에 저장으로 응답하고 원래 메시지 모드를 복원한다. 원본 이력은 수정하지 않는다. 상세와 적용 범위는 `reference/warnings-com.md` 15번 참조.
- **rhwp_pdf.py**: Windows가 아닌 OS에서 `hwpx_edit.py --to-pdf`가 쓰는 PDF 백엔드. 오픈소스 조판 엔진 rhwp CLI를 호출하고, 글꼴에 없어 빈 네모로 찍힌 문자를 PDF에서 찾아 보고한다(아래 "macOS·Linux에서 PDF 변환 (rhwp)" 참조)
- **hwpx_sign.py**: 이 디렉토리의 `hwpx_sign.py` (서명란(기준 텍스트 "(서명)" 등)에 서명/도장 이미지를 삽입하는 전용 도구, Windows + 한컴오피스 필요). COM으로 이미지를 넣어 BinData를 확보한 뒤 XML 후처리로 floating PAPER 절대좌표 전환·앵커 위 문단 이동·lineseg 제거를 자동 수행한다(아래 "서명/도장 이미지 삽입" 참조)
- **hwp2hwpx**: 이 디렉토리의 `convert/hwp2hwpx.bat`(Windows) 또는 `convert/hwp2hwpx.sh`(macOS/Linux)
- **python-hwpx CLI**: `pip install python-hwpx` (v2.9.0+): `hwpx-validate`, `hwpx-page-guard` 등

실행 시 이 스킬 디렉토리의 절대경로를 사용한다. 예: `python <스킬디렉토리>/hwpx_edit.py`

## Step 0: 환경 점검 (첫 실행 또는 의존성 의심 시)

```bash
python <스킬디렉토리>/hwpx_edit.py --check-env
```

기능 계층(tier)별로 무엇이 바로 되는지 한 번에 출력한다(읽기 전용, 아무것도 설치하지 않음). 이 스킬은 API 키를 쓰지 않으므로(전부 로컬 도구) 점검 대상은 pip 패키지와 시스템 런타임이다:

- **Tier 1 (읽기·편집, 필수)**: `python-hwpx`·`lxml`·`hwpx-tomd`. `pip install -r requirements.txt` 한 줄이면 충족하며, 대부분의 작업(`--to-md`·텍스트 치환·셀 편집)은 여기까지면 된다.
- **Tier 2 (HWP→HWPX)**: JDK 21 + 번들 JAR. 래퍼는 Windows `convert/hwp2hwpx.bat`, macOS/Linux `convert/hwp2hwpx.sh`이며 둘 다 `JAVA_HOME` → PATH 순으로 java를 자동 탐색한다.
- **Tier 3 (MD/DOCX/HTML→HWPX)**: `pypandoc-hwpx`(+ Pandoc). Pandoc은 `pypandoc-hwpx`가 번들 제공할 수 있어 경고만 떠도 변환이 동작할 수 있다.
- **Tier 4 (PDF·이미지·서명)**: Windows + 한컴오피스 COM(`pywin32`, 선택적 `pyhwpx`). 보안모듈 DLL·레지스트리·한컴 기동까지의 상세 진단은 `--diagnose-com`으로 위임한다. macOS·Linux는 PDF만 `rhwp` CLI로 되고 이미지 삽입·서명·한컴 정규화는 안 된다.

처음 클론한 환경이거나 특정 워크플로우가 의존성·런타임 누락으로 실패하면 먼저 실행해 무엇을 설치할지 확인한다. 환경이 준비된 뒤에는 매 작업마다 실행할 필요가 없다.

## 의사결정 트리

```
HWP/HWPX 작업 요청
├── 읽기 (내용 파악)
│   ├── HWP 파일 → 먼저 HWPX로 변환 (convert/hwp2hwpx.bat, macOS/Linux는 .sh) → 이후 HWPX 읽기 단계로
│   │   ※ JDK 없는 Windows + 한컴오피스 있음 → python hwpx_com.py 파일.hwp --from-hwp
│   │     (COM 폴백. JDK 설치 전까지의 우회, 아래 "HWP → HWPX 변환" 참조)
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
├── PDF 만들기 → hwpx_edit.py --to-pdf
│   (Windows는 한컴 COM, macOS·Linux는 rhwp. 아래 "macOS·Linux에서 PDF 변환 (rhwp)")
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
| HWPX/HWP → PDF | `hwpx_edit.py --to-pdf` (Windows 한컴 COM · 그 밖의 OS rhwp) | N/A |
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
# 1. HWP → HWPX 변환 (서식 100% 보존. Windows는 .bat, macOS/Linux는 .sh)
convert/hwp2hwpx.bat <파일.hwp> <파일.hwpx>
bash convert/hwp2hwpx.sh <파일.hwp> <파일.hwpx>

# 2. HWPX → Markdown
python hwpx_edit.py <파일.hwpx> --to-md -o output.md

# 표 셀에 긴 지문이 있는 경우 (고사지·보고서)
python hwpx_edit.py <파일.hwpx> --to-md --cell-br -o output.md
```

### 번들 JAR 출처·폴백 (hwp2hwpx가 예외로 죽을 때)

hwp2hwpx가 `EmptyStackException`·`IndexOutOfBoundsException` 등 예외로 죽거나 번들 JAR을 재빌드할 때 `reference/hwp-conversion.md`를 읽는다.


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
convert/hwp2hwpx.bat <입력.hwp> [출력.hwpx]        # Windows (PowerShell에서 호출)
bash convert/hwp2hwpx.sh <입력.hwp> [출력.hwpx]    # macOS/Linux
```

- 출력 파일 미지정 시 입력 폴더의 `_work-hwpx-automation/` 하위에 `.hwpx`로 생성(원본 비파괴). HWPX가 이후 단계의 입력일 뿐이면 부산물이므로 거기 두고, HWP→HWPX 변환 자체가 목적이면 작업 폴더로 옮긴다(아래 "작업 마무리" 참조). 출력 경로를 직접 지정하려면 두 번째 인자로 명시
- 서식 100% 보존 (Java 기반, hwplib + hwpxlib)
- 요구사항: JDK 21. 두 래퍼 모두 `JAVA_HOME` → (Windows는 Adoptium 표준 설치 →) PATH 순으로 java를 자동 탐색
- **JDK 없는 Windows 폴백**: 한컴오피스가 있으면 `python hwpx_com.py <입력.hwp> --from-hwp [-o 출력.hwpx]`로 COM 변환(변환 후 본문 글자 수 재검증 내장, `--to-md` recall 100% 실측). JDK 설치 전까지의 우회이며, 연속·배치 호출 전에는 잔류 `Hwp.exe`를 정리할 것(`reference/warnings-com.md` 9번)
- Windows(.bat): 입력 경로에 cp949 외 문자(en-dash, em-dash 등)가 있어도 내부에서 `%TEMP%`로 staging해 처리. macOS/Linux(.sh)는 JVM이 UTF-8 argv를 그대로 받으므로 staging 불필요
- ⚠ **Git Bash에서 호출 금지**: `cmd.exe /c`를 거치는 Git Bash에서 한글·공백 경로 인자를 넘기면 cp949 이중 셸 해석으로 깨져(`Exit code 2` + 깨진 바이트) Usage 분기로 빠진다. bat 내부 `%TEMP%` staging은 JVM argv 문제만 막을 뿐 그 앞단 cmd.exe 인자 전달 깨짐은 못 막는다. **PowerShell에서 직접 호출**하거나, 정 Bash가 필요하면 입력을 ASCII 경로 임시 폴더에 복사한 뒤 `java -cp "<convert>/hwp2hwpx-1.0.0.jar;<convert>/lib/hwplib-*.jar;<convert>/lib/hwpxlib-*.jar;<convert>" Hwp2HwpxCLI in.hwp out.hwpx`를 직접 실행한다
- ⚠ **변환물은 `hwpx-validate`가 `Preview/PrvText.txt` 누락으로 실패한다**(container.xml은 선언, ZIP엔 Preview/ 없음). 한글은 정상 열림·`--to-md` recall 100%라 읽기·편집엔 무해. 변환물을 **정본·배포·검증 대상**으로 쓸 때만 `python hwpx_edit.py <파일.hwpx> --add-preview`로 보정한다(section 무변형, 교훈 10).

## 핵심 워크플로우: MD → HWPX 변환 (Pandoc 방식)

MD 등을 `hwpx_convert.py`로 변환할 때 `reference/conversion.md`의 같은 제목 절을 읽는다.

## 핵심 워크플로우: MD → HWPX 생성 (Build-from-scratch 방식)

보고서급 문서를 python-hwpx로 처음부터 만들 때 `reference/build-from-scratch.md`의 같은 제목 절을 읽는다.

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

CLI로 안 되는 편집을 python-hwpx 스크립트로 직접 짤 때 `reference/api.md`를 읽는다.

## 구조적 편집 (행/표/단락 추가): raw lxml 필요

표에 행을 추가·복제하거나 raw XML로 문단을 고칠 때 `reference/structural.md`를 읽는다.

## HWPX 파일 구조

ZIP 내부 구조·네임스페이스가 필요할 때 `reference/format.md`를 읽는다.

## 서명/도장 이미지 삽입 (hwpx_sign.py)

서명란에 서명·도장 이미지를 넣을 때 `reference/signing.md`를 읽는다.

## macOS·Linux에서 PDF 변환 (rhwp)

Windows가 아닌 OS에서 `--to-pdf`를 쓰거나 rhwp 설치·출력 해석(`글꼴에 없는 문자가 빈 네모로 찍혔습니다`, exit 2)이 필요할 때 `reference/rhwp-pdf.md`를 읽는다.

## 한컴 COM 자동화 (Windows 전용, 한컴오피스 필수)

Windows + 한컴오피스에서 COM(`hwpx_com.py`·보안모듈·암호 해제·이미지 삽입)을 쓸 때 `reference/com.md`를 읽는다.

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

> 위 표의 마지막 3개 파일(`action-table.md`, `hwp-automation.md`, `parameterset-table.md`)은 한컴 공식 문서의 파싱본이라 저작권 고려로 저장소에 포함하지 않는다. 필요 시 README "참조 문서" 절의 안내에 따라 한컴 개발자 센터 아카이브의 원본 PDF를 내려받아 직접 파싱해 재생성한다.

## 주의사항

워크플로우별 주의사항은 각 reference 파일에 포함:
- **XML 편집 전반**: `reference/warnings-editing.md`: find/replace 동작, set-cell 체이닝, 서식 보존, PrvText.txt 누락, manifest self-closing 등
- **MD → HWPX 변환**: `reference/conversion.md` 하단: blockquote 누락, 각주 변환, 인코딩, 스크립트 보존
- **Build-from-scratch**: `reference/build-from-scratch.md` 하단: lineSpacing 단위, fontRef, set_cell_text, style_tables_xml 루프
- **한컴 COM 자동화**: `reference/warnings-com.md`: 이미지 XML 삽입 불가, 양식+이미지 워크플로우, PDF 변환

## 개선 반영 (함정·패턴 환류)

도구로 안 돼서 우회했거나 예상과 다른 동작·새 함정을 만났을 때 `reference/update-checklist.md`의 같은 제목 절을 읽는다.
