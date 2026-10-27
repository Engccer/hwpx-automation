# hwpx-automation

HWP/HWPX 문서 읽기, 변환, 편집을 위한 CLI 도구 + 코딩 에이전트 스킬

## 프로젝트 구조

```
hwpx-automation/
├── hwpx_edit.py          # HWPX 읽기/편집 통합 CLI (XML 파이프라인)
├── hwpx_com.py           # 한컴 COM 네이티브 파이프라인 CLI (pyhwpx, Windows 전용)
├── hwpx_sign.py          # 서명란에 서명/도장 이미지 삽입 (COM 삽입 + XML floating 후처리)
├── rhwp_cli.py           # rhwp 호출의 유일한 입구 (탐색·버전·JSON 봉투·exit code·임시 폴더→검증→교체)
├── hwpx_gate.py          # rhwp 산출물 독립 손실 게이트 (요소 census·BinData 해시·글자 대조·hwpx-validate)
├── rhwp_pdf.py           # macOS·Linux --to-pdf 백엔드 (rhwp_cli 위에서 빈 네모 문자 검출)
├── pdf_export.py         # Windows --to-pdf 공통 (변경 추적 경고 자동 응답)
├── test_*.py             # 단위 시험 (rhwp_cli·hwpx_gate·rhwp_pdf·pdf_export, 문서가 약속한 동작: test_doc_promises)
├── SKILL.md              # 에이전트 스킬 정의 (의사결정 트리 + 사용법)
├── convert/
│   ├── hwp2hwpx.bat      # HWP→HWPX 변환 (Windows, JDK 21 필요)
│   ├── hwp2hwpx.sh       # HWP→HWPX 변환 (macOS/Linux, JDK 21 필요)
│   ├── hwp2hwpx-1.0.0-edc0527p1.jar  # 상류 커밋 + 자체 패치 (파일명 해시가 버전)
│   ├── Hwp2HwpxCLI.java  # 변환기 소스
│   ├── hwpx_convert.py   # MD/DOCX/HTML 등 → HWPX (Pandoc)
│   ├── hwp_xml_to_md.py  # pyhwp XML → MD 폴백 파서
│   ├── patches/          # JAR 자체 패치 + 재빌드 절차
│   └── lib/              # hwplib-1.1.11.jar, hwpxlib-1.0.9.jar
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
python hwpx_edit.py <파일.hwp> --from-hwp             # HWP → HWPX (rhwp, JAR 교차 대조. --engine jar 폴백)

# 편집
python hwpx_edit.py <파일.hwpx> --find "A" --replace "B"  # 텍스트 치환
python hwpx_edit.py <파일.hwpx> --set-cell 0,1,0 "텍스트"  # 표 셀 입력
python hwpx_edit.py <파일.hwpx> --split-cell 0,1,0         # 병합 셀 분할
python hwpx_edit.py <파일.hwpx> --delete-after "텍스트"     # 특정 위치 이후 삭제
python hwpx_edit.py <파일.hwpx> --delete-empty-rows 0      # 빈 행 삭제
python hwpx_edit.py <파일.hwpx> --delete-rows 0,3,5        # 특정 행 삭제
python hwpx_edit.py <파일.hwpx> --sanitize                 # 검은 배경 등 수정
python hwpx_edit.py <파일.hwpx> --list-squeeze             # "한 줄로 입력" 과압축 문단 나열
python hwpx_edit.py <파일.hwpx> --fix-squeeze              # 과압축 문단을 자연 줄바꿈으로 전환

# 한컴 COM 자동화 (Windows + 한컴오피스)
python hwpx_edit.py --diagnose-com                         # COM 사용 가능 여부 진단
python hwpx_edit.py <파일.hwpx> --to-pdf                    # HWPX/HWP → PDF (rhwp, --pdf-backend com은 Windows 한컴)
```

### hwpx_com.py

한컴 COM(실제 한글 프로그램)을 구동하는 **두 번째 백엔드**. pyhwpx 기반이며
XML/Pandoc 파이프라인과 섞지 않기 위해 파일을 분리했다(Windows + 한컴오피스 전용).
이 스크립트가 생성한 HWPX는 COM·XML 양쪽 파이프라인에서 유효하지만,
Pandoc 생성 HWPX를 입력으로 넣으면 안 된다(`--get-text` 호환성 점검만 예외, `reference/warnings-com.md` 4번).
입력 파일을 절대 in-place로 덮어쓰지 않는다.

```bash
python hwpx_com.py --diagnose                              # pyhwpx + COM 진단
python hwpx_com.py output.hwpx --from-md input.md           # MD → COM 네이티브 HWPX 생성
python hwpx_com.py <파일.hwp> --from-hwp                    # HWP → HWPX 변환 (JDK 없는 기기 폴백)
python hwpx_com.py <파일.hwpx> --insert-image img.png       # 문서 끝 이미지 삽입 (별도 파일 저장)
python hwpx_com.py <파일.hwpx> --normalize                  # 한컴 재저장 정규화 (python-hwpx 산출물, 별도 파일)
python hwpx_com.py <파일.hwpx> --get-text                   # COM 기준 본문 추출 (호환성 점검)
python hwpx_com.py <파일.hwpx> --to-pdf                     # PDF 저장
```

### hwpx_sign.py

동의서·계약서 등의 서명란(기준 텍스트 `(서명)`·`(인)`)에 자필 서명/도장 이미지를
삽입한다. 기본 백엔드는 rhwp(모든 OS): `search`로 앵커를 찾고 `export-render-tree`의
글자 상자 중심에 겹쳐 `edit insert-image`로 넣은 뒤, 그림을 앵커 문단 맨 앞·글 앞으로
옮기고 `hwpx_gate.gate(allow_rhwp=True)`와 앵커 재측정을 통과해야 결과물로 둔다.
결과는 항상 `<입력>_서명.hwpx`, 출력 끝줄은 `engine=`. 단위 상수(`MM`·`PX`)는 이 파일에 있다.

`--backend com`(Windows + 한컴오피스)은 COM으로 BinData만 확보한 뒤 XML 후처리로
floating PAPER 절대좌표 전환·앵커 위 문단 이동·lineseg 제거를 수행한다. 좌표는
**편집 전 원본**의 lineseg에서 읽는다(`reference/warnings-com.md` 10번).

```bash
python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(서명)" --pdf
python hwpx_sign.py 동의서.hwpx --image 서명.png --backend com --keep-anchor --horz-offset 27510 --pdf
```

상세 워크플로우·함정은 `reference/signing.md`.

### convert/hwp2hwpx.bat · hwp2hwpx.sh

HWP(한컴 바이너리) → HWPX(Open XML) 변환. 서식 100% 보존. JDK 21 필요.

```bash
convert/hwp2hwpx.bat input.hwp [output.hwpx]        # Windows
bash convert/hwp2hwpx.sh input.hwp [output.hwpx]    # macOS/Linux
```

## 의존성

- **Python**: `python-hwpx` (비상업 라이선스), `lxml` (BSD-3-Clause), 편집 명령에 사용
- **hwpx-tomd** (MIT): `--to-md` 변환 엔진. `pip install hwpx-tomd` (PyPI·GitHub `Engccer/hwpx-tomd` 공개). 라이브러리로도 직접 사용 가능
- **Java**: JDK 21 (`--engine jar` 폴백과 rhwp 변환물의 교차 대조에만. 없으면 변환은 `unverified`)
- **Windows + 한컴오피스 + pywin32**: `--to-pdf --pdf-backend com`, `--diagnose-com` 등 COM 자동화 기능에 사용
- **rhwp CLI** (MIT, 별도 설치, v0.8.7 이상): HWP→HWPX 변환과 `--to-pdf`의 기본 엔진(모든 OS). GitHub Releases 바이너리를 PATH에 두거나 `HWPX_RHWP_BIN`으로 경로를 준다
- **pyhwpx** (선택): `hwpx_com.py` COM 네이티브 파이프라인 전용. `pip install pyhwpx` (1.7.2 검증, 2026-06-10)
- **번들된 JAR**: hwplib, hwpxlib, hwp2hwpx (모두 Apache-2.0)

## 스킬 사용

`SKILL.md`에 의사결정 트리와 전체 사용법이 정의되어 있다. 복잡한 작업 시 `reference/` 폴더의 상세 문서를 참조:

| 문서 | 내용 |
|------|------|
| `reference/api.md` | python-hwpx API 레퍼런스 |
| `reference/structural.md` | 구조적 편집 (행/표/단락 추가) 가이드 |
| `reference/format.md` | HWPX 파일 형식 상세 |
| `reference/build-from-scratch.md` | Markdown → HWPX 직접 빌드 |
| `reference/conversion.md` | MD → HWPX 변환(Pandoc) 워크플로우·스타일 후처리 |
| `reference/hwp-conversion.md` | HWP→HWPX 번들 JAR 출처·폴백 |
| `reference/encrypted-hwpx.md` | 암호화 HWPX 감지·해제 |
| `reference/com.md` | 한컴 COM 자동화 (hwpx_com.py, 보안모듈, 원격 실행) |
| `reference/signing.md` | 서명/도장 이미지 삽입 |
| `reference/rhwp-pdf.md` | macOS·Linux PDF (rhwp) |
| `reference/cases.md` | 규칙의 근거가 된 실측·사고 경위 |
| `reference/warnings-editing.md` | 편집 시 주의사항 |
| `reference/warnings-com.md` | 한컴 COM 자동화 주의사항 |
| `reference/update-checklist.md` | 업데이트 체크리스트 |

## 배포

외부 작업 공간에 배포한 사본은 이 저장소의 `origin/master`를 기준으로 한다. 작업 공간의 `sync_agent_docs.py`가 실행될 때마다 `_external_skills.json` 매니페스트에 지정한 대상 디렉터리에 최신 커밋의 추적 파일을 풀어 넣는다(`.gitignore` 파일은 보존). 배포한 사본은 직접 수정하지 않고, 이 저장소에 커밋·push하면 다음 동기화에서 반영된다.

## 코드 컨벤션

- hwpx_edit.py에 새 기능 추가 시 기존 CLI 인수 패턴(`--동작`, `--동작 인수`)을 따른다
- 편집 기능은 반드시 `save_hwpx()` 함수를 통해 저장 (sanitize 자동 적용). 저장 경로에 조건을 걸지 않는다
- HWPX를 쓰는 새 경로는 `zipfile.ZipFile(..., 'w')`를 직접 쓰지 말고 `write_hwpx_zip()`을, XML 직렬화는 `serialize_xml()`을 쓴다. HWPX를 읽는 새 경로는 `zipfile`로 읽은 직후 `sanitize_xml_parts()`를 태운다(`reference/warnings-editing.md` 15·18번)
- 구조적 편집(raw XML) 후에는 `lineseg` 수동 제거 필수
- rhwp는 `rhwp_cli.run()`으로만 부른다(명령 문자열 직접 조립 금지). 산출 파일 자리는 `rhwp_cli.OUTPUT` 자리표로 적고, 게시 전 검증은 `verify=` 콜백으로 끼운다. 봉투 필드명은 `output`·`format`·`pageCount`·`verify`(실측 v0.8.7)
- rhwp가 쓴 HWPX는 전체 재직렬화라 원본에 없던 손실이 생긴다(header.xml의 OUTLINE 7~9단계 `hp:switch` 블록과 미사용 paraPr이 사라진다, rhwp `--verify`는 이를 못 본다). 산출물은 `hwpx_gate.gate()`를 통과한 뒤에만 결과물로 삼고, 원본 XML이 없는 HWP 변환은 JAR 변환물과 `cross_check()`로 대조한다
- 정본 형식은 HWPX다. `.hwp`는 입구에서 변환하고 결과물은 HWPX로 저장한다. `.hwp`를 `.hwp`로 저장하는 편집 경로는 만들지 않는다
