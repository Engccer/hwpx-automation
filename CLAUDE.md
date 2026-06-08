# hwpx-automation

HWP/HWPX 문서 읽기, 변환, 편집을 위한 CLI 도구 + Claude Code 스킬

## 프로젝트 구조

```
hwpx-automation/
├── hwpx_edit.py          # HWPX 읽기/편집 통합 CLI
├── SKILL.md              # Claude Code 스킬 정의 (의사결정 트리 + 사용법)
├── convert/
│   ├── hwp2hwpx.bat      # HWP→HWPX 변환 (Windows, JDK 21 필요)
│   ├── hwp2hwpx-1.0.0.jar
│   ├── Hwp2HwpxCLI.java  # 변환기 소스
│   └── lib/              # hwplib-1.1.10.jar, hwpxlib-1.0.8.jar
├── reference/            # 상세 레퍼런스 (API, 구조적 편집, 파일 형식 등)
├── vendor/               # 서드파티 바이너리 (한컴 FilePathChecker DLL 등, .gitignore 처리)
│   └── README.md         # DLL 다운로드·레지스트리 등록 절차 + 라이선스 고지
├── requirements.txt      # python-hwpx, lxml (편집), hwpx-tomd (--to-md)
├── LICENSE               # MIT (vendor/ 하위 바이너리는 제외)
├── NOTICE                # 서드파티 라이선스 고지 (번들 JAR + vendor/ DLL)
└── README.md
```

## 핵심 도구

### hwpx_edit.py

HWPX 파일의 읽기와 편집을 모두 처리하는 통합 CLI. `--to-md` 변환 엔진은 독립
패키지 `hwpx-tomd`가 단일 소스로 보유하며(코드 분기 방지), 이 파일은 호출 래퍼만
담당한다. 편집 명령은 이 파일이 직접 구현한다.

```bash
# 읽기
python hwpx_edit.py <파일.hwpx> --to-md              # HWPX → Markdown (hwpx-tomd 엔진)
python hwpx_edit.py <파일.hwpx> --info                # 표 구조 확인

# 편집
python hwpx_edit.py <파일.hwpx> --find "A" --replace "B"  # 텍스트 치환
python hwpx_edit.py <파일.hwpx> --set-cell 0,1,0 "텍스트"  # 표 셀 입력
python hwpx_edit.py <파일.hwpx> --split-cell 0,1,0         # 병합 셀 분할
python hwpx_edit.py <파일.hwpx> --delete-after "텍스트"     # 특정 위치 이후 삭제
python hwpx_edit.py <파일.hwpx> --delete-empty-rows 0      # 빈 행 삭제
python hwpx_edit.py <파일.hwpx> --delete-rows 0,3,5        # 특정 행 삭제
python hwpx_edit.py <파일.hwpx> --sanitize                 # 검은 배경 등 수정

# 한컴 COM 자동화 (Windows + 한컴오피스)
python hwpx_edit.py --diagnose-com                         # COM 사용 가능 여부 진단
python hwpx_edit.py <파일.hwpx> --to-pdf                    # HWPX/HWP → PDF
```

### convert/hwp2hwpx.bat

HWP(한컴 바이너리) → HWPX(Open XML) 변환. 서식 100% 보존. JDK 21 필요.

```bash
convert/hwp2hwpx.bat input.hwp [output.hwpx]
```

## 의존성

- **Python**: `python-hwpx` (비상업 라이선스), `lxml` (BSD-3-Clause), 편집 명령에 사용
- **hwpx-tomd** (MIT): `--to-md` 변환 엔진. `pip install hwpx-tomd` (PyPI·GitHub `Engccer/hwpx-tomd` 공개). 라이브러리로도 직접 사용 가능
- **Java**: JDK 21 (HWP→HWPX 변환 시에만)
- **Windows + 한컴오피스 + pywin32**: `--to-pdf`, `--diagnose-com` 등 COM 자동화 기능에 사용
- **번들된 JAR**: hwplib, hwpxlib, hwp2hwpx (모두 Apache-2.0)

## 스킬 사용

`SKILL.md`에 의사결정 트리와 전체 사용법이 정의되어 있다. 복잡한 작업 시 `reference/` 폴더의 상세 문서를 참조:

| 문서 | 내용 |
|------|------|
| `reference/api.md` | python-hwpx API 레퍼런스 |
| `reference/structural.md` | 구조적 편집 (행/표/단락 추가) 가이드 |
| `reference/format.md` | HWPX 파일 형식 상세 |
| `reference/build-from-scratch.md` | Markdown → HWPX 직접 빌드 |
| `reference/conversion.md` | HWP→HWPX 변환 상세 |
| `reference/warnings-editing.md` | 편집 시 주의사항 |
| `reference/warnings-com.md` | 한컴 COM 자동화 주의사항 |
| `reference/update-checklist.md` | 업데이트 체크리스트 |

## 코드 컨벤션

- hwpx_edit.py에 새 기능 추가 시 기존 CLI 인수 패턴(`--동작`, `--동작 인수`)을 따른다
- 편집 기능은 반드시 `save_hwpx()` 함수를 통해 저장 (sanitize 자동 적용)
- 구조적 편집(raw XML) 후에는 `lineseg` 수동 제거 필수
