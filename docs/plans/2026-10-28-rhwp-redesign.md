# hwpx-automation: rhwp 중심 재설계 플랜

## Context

이음오피스(ieumoffice.com) 조사에서 한글 편집 엔진 rhwp(edwardkim/rhwp, Rust 단일 바이너리, MIT, v0.8.7 2026-10-06)가 HWP·HWPX를 OS 무관하게 읽고·고치고·그리는 것을 확인했다. hwpx-automation은 지금 rhwp를 macOS·Linux PDF 백엔드(`rhwp_pdf.py`)로만 쓰고, 나머지는 세 의존성에 나눠 맡긴다.

- HWP→HWPX: JDK 21 + hwp2hwpx JAR. PrvText 누락으로 `hwpx-validate`·`HwpxDocument.open`·`--find/--replace`가 실패해 `--add-preview` 보정이 필요하다.
- 이미지·서명·암호 해제·한컴 정규화·한컴 PDF: Windows + 한컴오피스 COM 전용. 맥에서는 불가능하고 SSH 세션에서는 COM 서버가 뜨지 않아 `schtasks` 우회가 필요하다.
- HWPX 미세 편집: python-hwpx + lxml. 원본 XML을 바이트 단위로 보존한다.

위원장 결정(2026-10-28): **축별 분담**. rhwp 1순위 축은 HWP→HWPX 변환, PDF, 이미지·서명 삽입, 누름틀, 표 격자 좌표·검색·쪽 좌표. 기존 lxml 1순위 축은 HWPX 미세 편집(원본 보존). 읽기(`--to-md`)는 hwpx-tomd 유지. JAR·한컴 COM은 폴백으로 남긴다.

위원장 추가 결정(같은 날): **정본 형식은 HWPX다. `.hwp`는 어떤 명령에서든 먼저 HWPX로 변환해 작업하고 결과물도 HWPX로 저장한다.** 근거는 둘이다. (1) HWPX만 XML이라 lxml 미세 편집·`hwpx-validate`·hwpx-tomd recall·태그 대조 같은 독립 검증이 가능하고, `.hwp`를 rhwp로 직접 고치면 rhwp가 rhwp를 채점하는 구조가 된다. (2) 공문서가 HWPX로 가는 정책 방향과 맞다. 상대가 `.hwp`를 요구할 때만 마지막 단계에서 rhwp `convert`로 `.hwp`를 만들고, 보관 정본은 HWPX다.

## 실측 근거 (2026-10-28, 이 맥, rhwp v0.8.7, 표본 HWP 3건·HWPX 2건)

| 축 | 측정 | 결과 |
|---|---|---|
| HWP→HWPX | `rhwp export-hwpx` vs JAR | rhwp 0.01~0.02초, JAR 0.1~0.2초. rhwp 산출물은 PrvText·PrvImage 포함, `hwpx-validate` 통과(1건 `langID` 경고만), `HwpxDocument.open` 성공. JAR 산출물은 PrvText 누락으로 validate·open 실패. hwpx-tomd 본문 단어 집합은 3건 모두 완전 일치 |
| 읽기 | `export-markdown` vs hwpx-tomd | rhwp는 쪽 단위 파일이고 쪽·줄 경계에서 단어 14개가 쪼개짐. hwpx-tomd 유지 |
| HWPX 편집 | `edit replace-text` 35건 | validate 통과, tomd recall 100%. 그러나 전체 재직렬화: header.xml `hp:switch` 23→20(미사용 paraPr의 OUTLINE heading 3개 소실), section0 `hp:t` 300→369. rhwp 자체 `--verify`·`hwpx-roundtrip`은 diff 0으로 PASS여서 **이 손실을 못 본다**. 기존 lxml 도구는 header.xml 바이트 동일 |
| 표 셀 | `edit set-cell` | HWP·HWPX 성공. 좌표는 export-tables 격자(병합 반영). 기존 `--set-cell T,R,C`는 물리 tc 인덱스라 같은 셀이 다른 좌표다. 중첩 표는 "v1 범위 밖"으로 거부 |
| 이미지 | `edit insert-image` | HWP·HWPX 양쪽 성공, export-svg에 `<image>` 생성. 단위 HWPUNIT(1/7200인치). `search --json`이 section·paragraph·page·charOffset을, `export-render-tree`가 TextRun bbox(px, 96dpi, ×75=HWPUNIT)를 돌려줘 "(서명)" 위 좌표 계산 가능 |
| PDF | `export-pdf` 21쪽 | 3.3초. 기존 `reference/rhwp-pdf.md` 규칙 그대로 |
| Pandoc 산출물 | `export-text`, `export-hwpx` | COM은 본문 0자로 읽지만 rhwp는 전문을 읽고 재저장도 된다(PrvText 추가). 한컴에서 열리는지는 Windows 미실측 |
| 환경 | 설치 | 맥 `~/.local/bin/rhwp` v0.8.7. Windows hunyong-samsung 미설치. PyPI 바인딩 없음(subprocess가 유일한 통합) |
| 상류 | 릴리스·이슈 | 6개월 25회 릴리스(2~4주 간격), 0.x. switch·chart 저장 손실 이슈 #3551·#2197·#3546은 closed. 모든 명령이 `--json`(schemaVersion 1.0)·`--dry-run`, exit 0/1/2/3(verify)/4(pages). `rhwp edit <sub> --help`는 exit 2, `-h`를 써야 함 |

## 설계 원칙

1. **rhwp 호출은 어댑터 하나로만.** 새 모듈 `rhwp_cli.py`가 바이너리 탐색·버전 확인·`--json` 봉투 파싱·exit code 매핑·임시 폴더→검증→`os.replace`를 맡고, 다른 파일은 이 모듈만 부른다. `rhwp_pdf.py`의 `save_pdf()`가 가진 패턴(shutil.which, 임시 폴더, 헤더 검증, `--password-stdin`)을 여기로 끌어올린다.
2. **rhwp가 쓴 파일은 독립 손실 게이트를 거친다.** rhwp 자가 검증은 자기 IR 밖의 손실을 못 보므로, 산출물마다 (a) hwpx-tomd recall(글자·단어 100%) (b) header.xml·section0.xml 태그 census 대조(사라진 요소 이름·수 보고) (c) `hwpx-validate`를 묶은 `hwpx_gate.py`를 통과해야 결과물로 인정한다. 변환·이미지·HWP 편집 축에서 필수, PDF는 기존 빈 네모 검사 유지.
3. **폴백은 자동 전환이 아니라 명시 안내다.** rhwp가 exit 1·3·4나 게이트 실패를 내면 결과를 `_work-hwpx-automation/`에 남기고 어느 폴백(JAR·COM)을 어떤 명령으로 부를지 stderr에 적는다. 조용한 자동 전환은 어느 엔진이 썼는지 알 수 없게 만든다.
4. **좌표 체계는 섞지 않는다.** 기존 `--set-cell T,R,C`(물리 tc 인덱스, lxml)는 그대로 두고, rhwp 격자 좌표 명령은 `--grid-*` 접두로 분리한다. 안내문에 "격자 좌표는 `--tables`로 먼저 본다"를 적는다.
5. **`.hwp`는 입구에서만 만난다.** 모든 명령은 `.hwp` 입력을 받으면 먼저 M2의 변환을 거쳐 HWPX 위에서 동작한다. `.hwp`를 `.hwp`로 저장하는 편집 경로는 만들지 않는다. 역방향은 `--to-hwp`(rhwp `convert`) 하나만 두고, rhwp가 경고하는 대로 그림·차트가 빠질 수 있으니 결과 안내에 "정본은 HWPX" 문장을 넣는다.
6. **버전 핀.** 어댑터가 `rhwp --version`을 읽어 0.8.7 미만이면 거부한다(이유: `export-hwpx` 위치 인자·`edit insert-image`·`--json` 봉투 계약을 이 버전에서 실측했다). 상한은 두지 않되 1.0 전까지는 CHANGELOG에 실측 버전을 적는다.

## 마일스톤

각 마일스톤은 독립 커밋·push 단위다(main-direct 저장소, 자율성 헌장). 순서 근거: 어댑터와 게이트가 모든 축의 전제라 먼저, 그다음 통증이 큰 순서(JAR의 PrvText·JDK 의존 → 맥에서 불가능했던 서명 → PDF 라우팅 → 신규 기능).

### M1. 어댑터와 손실 게이트

- 신규 `rhwp_cli.py`
  - `find_rhwp()` → 경로 또는 `None`(기존 `rhwp_pdf.INSTALL_HINT` 이관).
  - `version()` → 튜플, `MIN_VERSION = (0, 8, 7)`.
  - `run(args, *, input_path, output_path=None, password=None, timeout=120)` → `RhwpResult(ok, exit_code, envelope: dict|None, stderr)`. `--json`을 항상 붙이고 stdout을 파싱한다. 출력 파일은 임시 폴더에 쓰고 성공·게이트 통과 뒤 `os.replace`. `untrustedFields`에 든 값은 그대로 사람에게 보여 주되 분기 조건으로 쓰지 않는다.
  - exit 매핑: 0 성공, 1 `RhwpRuntimeError`, 2 `RhwpUsageError`(우리 버그로 취급), 3·4 `RhwpVerifyError`(산출물 보존 후 보고).
- 신규 `hwpx_gate.py`
  - `tag_census(xml_bytes) -> Counter`, `compare_census(before, after) -> {tag: (a, b)}`.
  - `tomd_recall(path)` → hwpx-tomd `convert()`의 `recall`·`char_recall`.
  - `validate(path)` → `hwpx-validate` 서브프로세스 결과.
  - `gate(original_path, produced_path, *, require_census=True) -> GateReport(passed, lost_tags, recall, validate_ok, messages)`. HWP 원본에는 census를 적용할 수 없으니(XML이 없음) `require_census=False`로 recall·validate만 본다.
- `rhwp_pdf.py`는 `rhwp_cli.run`을 쓰도록 바꾸고 빈 네모 검사(`missing_glyphs`)와 `LAYOUT_OVERFLOW` 집계는 그대로 둔다.
- 테스트: 신규 `test_rhwp_cli.py`(가짜 rhwp 스크립트로 성공·exit 1/2/3/4·버전 미달·미설치·봉투 파싱 실패), `test_hwpx_gate.py`(census 차이 검출, recall 미달, validate 실패를 각각 하나씩). `test_rhwp_pdf.py`는 통과 유지.
- 검증: `python -m unittest` 전체 통과. 실제 바이너리가 있으면 `RHWP_INTEGRATION=1`로 켜지는 통합 테스트가 표본 HWPX 1건에 대해 게이트 통과.

### M2. HWP→HWPX를 rhwp 기본으로, JAR은 폴백

- `hwpx_edit.py`에 `--from-hwp <입력.hwp>` 플래그(또는 입력이 `.hwp`이면 `--to-md`·편집 전에 자동 변환). 구현 `cmd_from_hwp()`: `rhwp export-hwpx <in> <out> --verify --json` → `hwpx_gate.gate(require_census=False)` → 성공 시 `_work-hwpx-automation/<이름>.hwpx`.
- 실패 시 안내: `bash convert/hwp2hwpx.sh` 또는 `.bat` 명령과 `--add-preview` 후속을 그대로 출력. JAR 래퍼·패치·`--add-preview`·`--sanitize`는 삭제하지 않는다(EmptyStackException 같은 rhwp 미실측 문서군의 폴백).
- `--check-env` Tier 재정의: Tier 1 읽기·편집(python-hwpx·lxml·hwpx-tomd), Tier 2 변환·PDF·이미지(**rhwp**, 버전 표시), Tier 3 MD→HWPX(pypandoc-hwpx), Tier 4 폴백(JDK+JAR, Windows 한컴 COM). `cmd_check_env()`·`test_rhwp_pdf.py`의 check-env 테스트를 함께 고친다.
- SKILL.md 259행 "HWP 바이너리 직접 파싱 금지"에 예외를 명시: rhwp는 허용하며 근거는 실측(단어 집합 일치·PrvText 포함·validate 통과). kordoc 금지는 유지.
- 테스트: `test_doc_promises.py`에 `FromHwpRoutesToRhwpThenJar`(가짜 rhwp 성공 → 결과 경로, 가짜 rhwp exit 3 → JAR 안내 문구 포함).
- 검증: 표본 HWP 3건에서 rhwp 변환물 `--to-md` recall 100%, `hwpx-validate` 통과, 변환 시간 로그.

### M3. 이미지·서명을 맥·Windows 공통 rhwp 경로로

- `hwpx_sign.py`에 백엔드 분기 `--backend rhwp|com`(기본 `rhwp`; Windows + COM 가능 환경에서도 rhwp 기본, `com`은 옵트인).
- rhwp 경로 `rhwp_sign()`:
  1. `rhwp search <doc> <앵커> --json` → page·paragraph·charOffset. 매치 0건·2건 이상은 기존 규칙대로 exit 1(거짓 OK 금지, `--occurrence`로 지정).
  2. `rhwp export-render-tree <doc> -p <page>` → 앵커 TextRun bbox(px). 앵커 글자가 run 중간이면 run 폭을 글자 수로 비례 분할해 x를 추정하고 WARN을 남긴다.
  3. px×75 → HWPUNIT, `--width-mm`·`--height-mm`·`--gap-mm`·`--vert-adjust`(기존 옵션 이름 유지, mm→HWPUNIT는 1mm = 283.46 HWPUNIT) 적용. "(서명)" 위에 겹쳐 찍는 기본값은 앵커 bbox 중심.
  4. `rhwp edit insert-image <doc> --image <png> --page --x --y --width --height -o <임시>` → `hwpx_gate.gate()` → `<입력>_서명.<같은 확장자>`.
  5. `--pdf`는 `hwpx_edit.py --to-pdf`(M4 뒤에는 rhwp 기본)로 만들고 기존 `count_pdf_images()` 또는 `rhwp export-svg -p <page>`의 `<image>` 수로 확인(pymupdf 없는 환경 대안).
- COM 경로(`com_insert`·`to_floating`·`anchor_geometry`)는 그대로 두고 `--backend com`에서만 쓴다. `reference/signing.md`·`evals/scenarios.md` T4("Mac 사용자, 원격 Windows")를 rhwp 경로로 다시 쓴다.
- 테스트: `test_sign_rhwp.py`(가짜 rhwp가 search·render-tree·insert-image 봉투를 순서대로 돌려줄 때 좌표 계산값, 매치 0·2건 거부, run 중간 앵커 WARN, mm 변환).
- 검증: 표본 동의서.hwp(실제 "(서명)" 앵커 있음)에 시험 PNG를 넣고 export-svg로 위치 확인. 한컴에서 열리는지는 M6 Windows 게이트.

### M4. PDF 라우팅 통일과 Pandoc 산출물 갭 해소

- `hwpx_edit.py --to-pdf`: 모든 OS에서 rhwp 기본, `--pdf-backend com`(Windows 한컴 동일 조판이 꼭 필요할 때) 옵트인. `cmd_to_pdf()` 분기와 `test_rhwp_pdf.py`·`test_pdf_export.py` 기대값 수정.
- Pandoc(`hwpx_convert.py`) 산출물의 PDF는 rhwp로 낸다고 SKILL.md 의사결정 트리에 명시(기존 "문서가 정하지 않았다" 갭 해소). COM `--to-pdf`에는 Pandoc 산출물 입력 금지가 그대로 남는다.
- 검증: `test.md`→HWPX→rhwp PDF 1건, 기존 표본 HWPX PDF 1건에서 빈 네모 0.

### M5. 누름틀·격자 좌표·역변환 (신규 기능, HWPX 위에서만)

- `hwpx_edit.py` 신규 플래그. 입력이 `.hwp`이면 M2 변환을 먼저 거친 HWPX에 적용하고, 산출물은 항상 HWPX:
  - `--tables`: `rhwp export-tables --json`을 사람이 읽는 격자 표로 출력(좌표 확인용, 변경 없음).
  - `--fields` / `--fill-fields <JSON|@파일>`: `rhwp fields`·`rhwp edit fill-fields`.
  - `--grid-set-cell T,R,C TEXT`: `rhwp edit set-cell`(격자 좌표). 중첩 표 거부 메시지는 그대로 전달하고 기존 `--set-cell`(lxml, 물리 인덱스)을 안내. 원본 보존이 중요한 공문은 기존 `--set-cell`이 1순위라는 문장을 도움말에 둔다.
  - `--to-hwp`: `rhwp convert <in.hwpx> <out.hwp> --verify --verify-pages`. 상대가 `.hwp`를 요구할 때만 쓰는 비상구. 결과 안내에 "정본은 HWPX, 그림·차트 손실 가능" 문장을 넣고 변환 전후 `rhwp export-text` 글자 수를 대조한다.
- `--fill-fields`·`--grid-set-cell` 산출물은 `rhwp_cli.run` + `hwpx_gate.gate`(census 포함) 필수.
- 테스트: 가짜 rhwp로 플래그→인자 매핑, `.hwp` 입력 시 변환 선행, 중첩 표 오류 전달, `--to-hwp` 안내 문구.
- 검증: 평가서.hwp → 변환 → `--grid-set-cell 0,4,3`로 성명 칸 변경 → `--tables`로 확인 → 게이트 통과.

### M6. Windows 설치와 실기 게이트 (자리: 로컬·위원장)

- hunyong-samsung에 rhwp v0.8.7 windows x86_64 zip 설치: 맥에서 받아 `SHA256SUMS.txt` 대조 → `scp` → `%USERPROFILE%\.local\bin`에 풀고 사용자 PATH 등록(`setx`는 영구 변경이라 실행 전 명령을 보고). SSH 기본 셸이 cmd인 점을 지킨다(`reference/environment.md`).
- 한컴 2024 실기 확인(위원장 또는 원격 데스크톱 세션): rhwp 변환 HWPX 3건, rhwp 서명 삽입본(HWP·HWPX), Pandoc 재저장본이 한글에서 열리고 서명 위치가 "(서명)" 위에 있는지. 열리지 않는 문서군이 있으면 그 축은 COM 폴백을 기본으로 되돌리고 CLAUDE.md에 사유를 적는다.
- `~/.claude/reference/environment.md`에 Windows rhwp 설치 위치 한 줄.

### M7. 문서 분배와 버전

- SKILL.md: 머리에 "정본 형식은 HWPX. `.hwp`는 입구에서 변환하고 결과물은 HWPX" 원칙 문장. 의사결정 트리(HWP 입력 → rhwp 변환, 서명 → rhwp, PDF → rhwp, HWPX 미세 편집 → lxml, `.hwp` 요구 시 → `--to-hwp`), "도구별 용도 선택" 표에 엔진 열 추가, Tier 표, `reference/rhwp-pdf.md`를 `reference/rhwp.md`로 확장(설치·버전·봉투·게이트·좌표 단위·폴백 조건). k-skill `rhwp-edit`·`rhwp-advanced`·`hwp`는 수정하지 않고 "같은 엔진의 WASM·안내형 래퍼, HWP5 출력만" 한 줄 라우팅.
- CLAUDE.md: 의존성 절(rhwp 최소 버전·설치), 함정(전체 재직렬화·중첩 표·좌표 체계 분리·자가 검증 한계). 날짜 있는 서사는 CHANGELOG로.
- PROGRESS.md 현재 상태 갱신, CHANGELOG.md에 마일스톤별 항목, `metadata.version` 1.4.0 → 1.5.0(동작 변경 minor; M1 커밋에서 올리고 이후 유지).
- `python sync_agent_docs.py`로 AGENTS.md 재생성. 마일스톤마다 commit·push, 마지막에 `doc-audit`.

## 폴백 규칙

| 축 | 1순위 | 폴백 | 전환 조건 |
|---|---|---|---|
| HWP→HWPX | rhwp export-hwpx | JAR hwp2hwpx + `--add-preview`, Windows는 COM `--from-hwp` | rhwp exit 1/3/4 또는 게이트 실패 |
| PDF | rhwp export-pdf | Windows 한컴 COM(`--pdf-backend com`) | 한컴 동일 조판이 요구될 때, 빈 네모 exit 2 |
| 이미지·서명 | rhwp insert-image | `hwpx_sign.py --backend com` | 앵커 미해결, 게이트 실패, M6 실기에서 한컴 비호환 발견 |
| HWPX 미세 편집 | lxml(`--find/--replace`·`--set-cell`·행 삭제) | rhwp `--grid-*` | 사용자가 격자 좌표를 원할 때만 |
| 누름틀·격자 셀(HWPX) | rhwp | 기존 `--set-cell`(lxml) | 중첩 표, 원본 보존 우선 문서 |
| HWPX→HWP 역변환 | rhwp convert(`--to-hwp`) | Windows 한컴 COM 다른 이름 저장 | 그림·차트가 빠졌을 때 |
| 암호 해제 | 한컴 COM(기존) | rhwp `--password-stdin`은 PDF·변환에만 | 암호 HWPX 표본으로 실측하기 전까지 유지 |
| 한컴 재저장 정규화 | 한컴 COM(기존) | 없음 | |

## 하지 않는 것

- hwpx-tomd 대체, 기존 lxml 편집 경로·JAR·COM 코드 삭제, k-skill 스킬 수정, MCP 서버(`rhwp mcp-serve`) 등록(별도 판정), `scaffold`로 Pandoc 대체(별도 실측 뒤 판정), `cargo install`(릴리스 바이너리만).

## 리뷰 게이트 판정

새 판정 계층(손실 게이트)과 외부 도구 계약 가정(rhwp JSON 봉투·exit code)을 담으므로 설계 단계 적대적 리뷰 대상이다. M1 착수 전 codex `adversarial-review`(실패 신호면 `model: fable` 서브에이전트)로 이 플랜을 검토한다. 구현 중 묶음별 spec-compliance 서브에이전트, M6 실기 확인은 리뷰로 대체할 수 없는 게이트다.

## 전체 검증

1. `python -m unittest` 전체 통과(가짜 rhwp 기반).
2. `RHWP_INTEGRATION=1 python -m unittest`로 실제 바이너리 통합 테스트(표본은 scratchpad에서 `tests/fixtures/`로 익명 표본 1~2건을 새로 만들어 넣는다. 개인 문서는 저장소에 넣지 않는다).
3. 시나리오 재연: 표본 HWP 변환→`--to-md`, 동의서 서명 삽입→export-svg, 정책.hwpx PDF, 평가서.hwp 변환→격자 셀 변경→`--tables`, 그 결과 `--to-hwp` 역변환 글자 수 대조.
4. Windows 실기(M6)에서 한컴 2024 열기 확인 뒤에만 SKILL.md에서 "rhwp 기본"을 확정 문장으로 둔다. 그 전까지는 "실기 확인 대기" 상태를 PROGRESS.md에 적는다.
5. `evals/scenarios.md` T1~T5를 개정 문서만 읽는 시험자에게 돌려 라우팅이 달라졌는지 확인한다.
