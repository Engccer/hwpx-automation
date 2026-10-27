# rhwp 재설계 M2~M7 병렬 세션 계획

기준 플랜: `docs/plans/2026-10-28-rhwp-redesign.md`. M1은 cc08c7e로 끝났다. 코디네이터는 이 문서를 쓴 세션이며 코드를 만지지 않는다.

## §1 마일스톤·확정 판정·모델 배정

| 세션 | 마일스톤 | 소유 핵심 | 모델·노력 |
|---|---|---|---|
| `edit` | M2 → M4 → M5 (차례로, 마일스톤마다 리뷰·commit·push) | `hwpx_edit.py` | opus · medium |
| `sign` | M3 | `hwpx_sign.py`, `hwpx_gate.py` | opus · medium |
| 코디네이터 | M6 (Windows 설치만) | 저장소 밖 | 이 세션 |
| `docs` (2차 웨이브) | M7 문서 분배 + doc-audit | SKILL.md·CLAUDE.md·README·PROGRESS·CHANGELOG·BACKLOG·reference/ | opus · medium |

위원장 확정 판정(2026-10-28, 플랜 원칙 7의 대체 결정과 같다):

- 실기 검증은 위원장 실사용으로 한다. **옵트인 단계 없이 rhwp가 기본값**이고 JAR·COM은 명시 폴백이다. `--engine jar`, `--backend com`, `--pdf-backend com`은 사용자가 고를 때만 쓴다.
- rhwp 재직렬화의 header.xml 손실(OUTLINE 7~9단계 `hp:switch`, 미사용 paraPr)은 **게이트 허용 목록으로 통과시키되 보고에 표기**한다. 그 밖의 요소 손실은 여전히 실패다.
- M6은 hunyong-samsung 설치까지만. `setx` 실행 전 명령을 보고한다. 한컴 2024 호환은 위원장이 실사용에서 본다.
- 세션 `edit`가 세 마일스톤을 이어 하는 것은 위원장 지정이다(스킬의 "세션은 마일스톤 하나" 규칙보다 우선). 완화: 마일스톤마다 push를 끝내고 보고 파일을 남기므로, 문맥이 커져 세션을 닫게 되면 새 창이 보고 파일로 잇는다.
- 설계 단계 적대적 리뷰는 플랜 단계에서 끝났다(플랜 "리뷰 게이트 판정" 절). 구현 중 리뷰는 마일스톤마다 spec-compliance + code-quality 서브에이전트(`model: opus`) 1회씩이고, 계층 선택이 갈리는 자리(예: M3 좌표 산출 방식)만 codex `adversarial-review` 1회를 더 쓴다.

## §2 파일 소유권 지도

관측 SHA: 33ab881 (origin/master). 술어는 "이 마일스톤이 그 파일을 고치는가"다.

| 파일 | `edit` | `sign` | `docs` | 비고 |
|---|---|---|---|---|
| `hwpx_edit.py` | 소유 | 읽기 | 읽기 | M2 `hwp_to_hwpx()`·`--from-hwp`·`--engine`·check-env Tier, M4 `--pdf-backend`, M5 `--tables`·`--fields`·`--fill-fields`·`--grid-set-cell`·`--to-hwp` |
| `rhwp_pdf.py`, `pdf_export.py` | 소유 | 금지 | 읽기 | M4 |
| `test_doc_promises.py`, `test_rhwp_pdf.py`, `test_pdf_export.py` | 소유 | 금지 | 금지 | |
| 신규 `test_from_hwp.py`, `test_grid.py` 등 | 소유 | 금지 | 금지 | 파일 이름은 edit가 정한다 |
| `tests/fixtures/` (익명 표본) | 소유 | 읽기 | 읽기 | 개인 문서 금지 |
| `hwpx_sign.py` | 금지 | 소유 | 읽기 | |
| `hwpx_gate.py`, `test_hwpx_gate.py` | **금지** | 소유 | 읽기 | 허용 목록 계약은 아래 §2-1 ② |
| 신규 `test_sign_rhwp.py` | 금지 | 소유 | 금지 | |
| `reference/signing.md`, `evals/scenarios.md` T4 | 금지 | 소유 | 읽기 | |
| `reference/hwp-conversion.md`, `reference/rhwp-pdf.md`, `reference/conversion.md` | 소유 | 금지 | 읽기·M7 확장 | M7이 `rhwp-pdf.md`를 `rhwp.md`로 확장 |
| `SKILL.md` | 절 "HWP → HWPX 변환"·"macOS·Linux에서 PDF 변환"·"hwpx_edit.py 사용법"·"Step 0" 한정 | 절 "서명/도장 이미지 삽입" 한정 | 전체(M7) | 의사결정 트리·도구별 용도 표·Tier 표 전면 개정은 M7. edit·sign은 자기 명령의 행동 문장만 고친다 |
| `CLAUDE.md` | 절 `hwpx_edit.py`·의존성 한 줄 | 절 `hwpx_sign.py` | 전체(M7) | |
| `CHANGELOG.md`, `PROGRESS.md`, `docs/BACKLOG.md` | 자기 항목만 | 자기 항목만 | 전체 | rebase 뒤 `comm -23` 대조 필수 |
| `rhwp_cli.py`, `test_rhwp_cli.py` | 금지 | 금지 | 금지 | M1 산출. 부족하면 코디네이터에 보고(어느 필드·어떤 봉투) |
| `hwpx_com.py`, `convert/*` | 금지 | 금지 | 금지 | 폴백은 손대지 않는다 |
| `AGENTS.md` | 금지 | 금지 | 재생성(`python ../sync_agent_docs.py`) | 생성물 |

### §2-1 경계 인터페이스 (병렬 전 고정)

① **HWP 입구 변환 함수** (`edit`가 M2에서 구현, `sign`이 M3에서 호출)

```python
# hwpx_edit.py
ConvertResult = namedtuple("ConvertResult", "path engine status messages")
#   path: 산출 HWPX 경로(str). dest가 None이면 입력 옆 `_work-hwpx-automation/<stem>.hwpx`
#   engine: "rhwp" | "jar"
#   status: "passed" | "unverified"   (unverified = JDK 없어 cross_check 생략, 또는 hwpx-tomd 없음)
#   messages: list[str]  (사람에게 보여 줄 줄. 교차 대조 요약·전환 사유)
def hwp_to_hwpx(src, dest=None, *, engine="rhwp") -> ConvertResult
#   실패: rhwp_cli.RhwpError 계열 또는 RuntimeError를 올린다. 메시지에 폴백 명령(convert/hwp2hwpx.sh|.bat + --add-preview)을 넣는다.
#   stdout에 `engine=<rhwp|jar> status=<passed|unverified>` 한 줄을 찍는다.
```

`sign`은 `.hwp` 입력을 받으면 `hwpx_sign._convert_hwp(src)` 래퍼 안에서 `from hwpx_edit import hwp_to_hwpx`를 지연 import해 부른다. `edit`가 통합되기 전에는 ImportError를 잡아 "`.hwp`는 `python hwpx_edit.py --from-hwp`로 먼저 변환(M2)"로 exit 1한다. 테스트는 래퍼를 patch한다.

② **게이트 허용 목록** (`sign`이 M3에서 구현, `edit`가 M5에서 호출)

```python
# hwpx_gate.py
def gate(original, produced, *, allow=None, allow_bindata_added=0,
         expect_text_change=False, require_census=True, require_bindata=True,
         allow_rhwp=False) -> GateReport
#   allow_rhwp=True: Contents/header.xml의 census 손실 중 rhwp 재직렬화로 알려진 것
#     (OUTLINE 7~9단계 hp:switch/case/default 블록과 그 하위, 미사용 paraPr과 그 하위)만
#     실패 대신 GateReport.allowed(dict, lost와 같은 키 형식)에 넣는다.
#     section*.xml 손실과 그 밖의 header.xml 손실은 여전히 fail.
#   GateReport.summary()와 describe_lost()는 allowed가 비어 있지 않으면
#     "허용된 손실(rhwp 재직렬화): switch×3, paraPr×2 ..." 한 줄을 반드시 낸다(보고 표기 결정).
#   기존 allow 콜백·is_text_run_key는 그대로. 기존 시그니처 호출은 동작이 바뀌지 않는다.
```

`edit`의 M5(`--fill-fields`·`--grid-set-cell`)는 `gate(..., allow=is_text_run_key, allow_rhwp=True, expect_text_change=True)`로 부른다. M5 착수 시점에 `sign`이 아직 통합되지 않았으면 `allow_rhwp`를 쓰는 코드는 쓰되 그 테스트는 `sign` 통합 뒤 rebase하고 돌린다. `edit`는 `hwpx_gate.py`를 고치지 않는다.

③ **좌표·단위**: HWPUNIT = 1/7200인치, 1mm = 283.46 HWPUNIT, render-tree px(96dpi) × 75 = HWPUNIT. `sign`이 `hwpx_sign.py`에 상수로 둔다. `edit`는 쓰지 않는다.

④ **엔진 보고 문장**: 산출물을 내는 모든 명령은 stdout 마지막에 `engine=<rhwp|jar|com|lxml>`와 전환 사유(폴백이면)를 찍는다(플랜 원칙 3). 양쪽 세션 공통.

## §3 git 격리 (main-direct, 브랜치 이름은 `master`)

```bash
# 코디네이터가 미리 만든다(prepare-worktrees.sh는 origin/main 고정이라 손으로)
git -C ~/Mac-Projects/hwpx-automation worktree add ~/hwpx-automation-wt/<name> -b feat/<name> origin/master
# 세션: 자기 브랜치에만, pathspec 커밋(git add -A 금지)
# 통합(마일스톤마다):
git fetch origin && base=$(git rev-parse origin/master) && echo "$base" > ~/hwpx-automation-wt/<name>-reports/base.sha && git rebase "$base"
comm -23 <(git show "${base}:CHANGELOG.md" | sort) <(sort CHANGELOG.md)   # PROGRESS.md, docs/BACKLOG.md, SKILL.md, CLAUDE.md도
python3 -m unittest && RHWP_INTEGRATION=1 python3 -m unittest
git push origin feat/<name>:master && git -C ~/Mac-Projects/hwpx-automation pull --ff-only
```

- `--force` 금지. `git stash` 금지(바로 커밋). 리뷰 중 rebase·커밋 금지.
- 게이트는 2초짜리 단위 시험이라 `gate-lock.py` 불필요. Python 패키지는 전역 설치본을 그대로 쓴다(`npm install` 없음). rhwp는 `~/.local/bin/rhwp` v0.8.7.
- 기준선: `~/.claude/parallel-sessions/hwpx-automation/baseline-<base 12자리>.log` (75건 통과, skipped 2).
- 마일스톤 push 뒤 메인 체크아웃은 세션이 `pull --ff-only`로 맞춘다.

## §4 웨이브

| 웨이브 | 세션 | 조건 |
|---|---|---|
| 1 | `edit`(M2→M4→M5), `sign`(M3) | 동시 착수. 코디네이터는 그 사이 M6 설치 |
| 2 | `docs`(M7 + doc-audit) | `edit`·`sign` 통합 완료 뒤 새 창 |

위원장 판정 대기: 없음(판정은 위에서 확정). 최종 순서는 위원장 결정으로 남긴다.

## §5 착수 프롬프트·보고 파일

- 프롬프트 사본: `~/.claude/parallel-sessions/hwpx-automation/<name>.prompt.txt`
- 보고 디렉터리: `~/hwpx-automation-wt/<name>-reports/` (`<단계>-<YYYYmmddHHMM>.md`, 리뷰는 `review-<주제>-<시각>.md`, 판정 요청은 `judgment-<시각>.md`)
- 코디네이터 주소: 이 세션(`ListAgents` 기준 `hwpx-automation-*`), 프롬프트에 `이름 [ref]`로 적는다.

## 정정·종료 기록

- 2026-10-28 07:30, 9b7e7d7 기준. §2-1 ② 정정: 허용 목록은 header.xml 개요 블록에 그치지 않는다. 세션 sign이 실측하고 코디네이터가 한컴 저장 HWPX(문단 16,671개)로 재현했다: rhwp 재직렬화는 모든 문서에서 section*.xml의 `hp:p id`를 0,1,2…로 다시 매기고(문단 수·내용은 그대로) `hp:shadow type="NONE"`을 지운다(32→22). `allow_rhwp=True`는 이 둘도 allowed로 보낸다. 조건: p id는 문단 수가 같고 id를 뺀 속성 묶음이 같을 때만, shadow는 `type="NONE"`일 때만. 둘 다 "허용된 손실(rhwp 재직렬화)" 줄에 건수로 보고한다. 개요 블록 허용은 본문이 쓰지 않는 paraPr 안의 손실로 좁힌다(본문이 쓰는 paraPr의 개요 단계 손실은 실패). 기존 시그니처 호출의 동작은 바뀌지 않는다.
- M6 완료(9b7e7d7): hunyong-samsung `C:\Users\pc\.local\bin\rhwp.exe` v0.8.7, 해시 대조, `setx` 불필요.
- M2 통합(718bc46, 세션 edit): 단위 91건. 리뷰 8번(cross_check BinData 해시 미대조)은 hwpx_gate 소유자인 sign의 M3 범위로 넘겼다(07:40 전파).
- 07:50 소유권 예외 승인: sign이 M3 커밋에 edit 소유 test_doc_promises.py의 SignPdfImageGate argv 두 곳에 `--backend com`을 넣는다(기본 백엔드 전환과 같은 커밋이어야 시험이 깨지지 않는다). edit에 전파.
- 08:00 판정(edit judgment-202610280735): `hwpx_sign.py --pdf`의 하위 호출은 sign이 M3에서 `--backend com`이면 `--pdf-backend com`을 붙인다(소유권대로). edit의 `reference/com.md` 수정(`--to-pdf` 예시에 `--pdf-backend com`) 자진 신고 수용.
- M4 통합(d018158, 세션 edit): 단위 98건. `--to-pdf` 모든 OS rhwp 기본, `--pdf-backend com` Windows 폴백. 암호 HWPX의 rhwp PDF 해제는 미실측(M7 인계 목록에 있음).
- M3 통합(038b798, 세션 sign): 단위 116건. 허용 목록 술어 실측(본문이 안 쓰는 paraPr 안 개요 블록·p id 재부여·shadow NONE 소실만), cross_check BinData 해시 대조, `--backend com`이면 `--pdf-backend com`. sign 창 닫음(08:20).

## M7 인계 목록 (edit·sign 보고에서 모음, docs 세션이 처리)

- CLAUDE.md 구조도 `rhwp_pdf.py`("macOS·Linux")·`pdf_export.py`("Windows --to-pdf 공통") 설명과 색인 "macOS·Linux PDF" 줄을 "모든 OS rhwp 기본, Windows COM은 `--pdf-backend com`"으로.
- SKILL.md 절 제목 "macOS·Linux에서 PDF 변환 (rhwp)"과 상세 레퍼런스 색인 줄. 의사결정 트리 72행 "hwp2hwpx 변환물은 PrvText 누락…"은 `--engine jar`·래퍼 직접 호출에만 해당.
- `evals/scenarios.md` 39행: Pandoc 산출물 PDF 갭이 메워졌으니 기대 답 수정.
- 암호 문서: rhwp `export-pdf`의 암호 해제는 미실측. 실패 안내에 "암호 문서는 Windows `--pdf-backend com`"이 들어 있다. BACKLOG 감시 항목으로.
- 표본 함정(CLAUDE.md 함정 후보): python-hwpx로 갓 만든 HWPX를 rhwp convert로 만든 HWP는 lineseg가 없어 export-hwpx `--verify`가 exit 3. 한 번 더 왕복한 HWP는 차이 0.
- AGENTS.md 재생성(`python ../sync_agent_docs.py`).
- (sign 보고) 아직 "서명은 Windows+한컴 전용"이라고 말하는 자리: `SKILL.md` 19행(도구 위치)·36행 근처(Tier 표)·105행(도구별 용도 표 `hwpx_sign.py (COM+XML)`), `README.md` 75행, `CLAUDE.md` 11행(구조 주석). 새 기본 동작·T4 채점과 맞춘다.
- (sign 보고) `reference/rhwp.md`에 옮길 rhwp 계약 실측: `search`의 `page`는 문단 시작 쪽, 표 셀 매치의 `cell` 키와 컨테이너 기준 `pi`, `export-render-tree`·`measure-width`는 `--json` 없음(exit 2, `rhwp_cli.run(expect_json=False)`), render-tree 파일명 `render_tree_<쪽+1 세 자리>.json`, `measure-width`는 TSV, `edit insert-image` 봉투 필드(`output, outputFormat, page, x, y, width, height, overflow, binDataId, changedPages`)와 "그 쪽 첫 문단에 SQUARE로 붙여 빈 문단이면 줄이 접힌다" 함정(`float_over_anchor`가 처리).
- (sign 보고) 위원장 실사용 확인 권고: 한컴 2024에서 rhwp 서명본 열기, 앵커 문단이 여러 줄인 표본 하나 포함. PROGRESS "실기 확인 대기" 문장에 넣는다.
