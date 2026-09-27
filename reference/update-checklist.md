# hwpx 스킬 업데이트 체크리스트

사용자가 "hwpx 스킬 업데이트" 또는 "hwpx 업데이트 확인"을 요청하면 이 체크리스트를 따른다.

## 1단계: python-hwpx 라이브러리 업데이트 (필수)

```bash
# 현재 버전 확인
pip show python-hwpx

# 최신 버전으로 업데이트
pip install --upgrade python-hwpx

# 업데이트 후 버전 확인
pip show python-hwpx
```

### 릴리즈 노트 확인

```bash
# PyPI 또는 GitHub에서 최신 릴리즈 노트 확인
gh api repos/airmang/python-hwpx/releases --jq '.[0:3] | .[] | "\(.tag_name): \(.name)\n\(.body)\n---"'
```

### 반영 판단

새 릴리즈에 아래 항목이 있으면 SKILL.md 또는 reference/ 파일 업데이트 필요:

| 변경 유형 | 반영 대상 |
|-----------|----------|
| 새 CLI 도구 추가 | SKILL.md의 "python-hwpx CLI 도구" 섹션 |
| 새 API 메서드 (테이블, 문단 등) | reference/api.md |
| 기존 API 변경/삭제 | SKILL.md + reference/api.md |
| 검증/무결성 관련 변경 | SKILL.md의 워크플로우 섹션 |
| 버그 수정만 | requirements.txt 버전만 업데이트 |

반영 후 requirements.txt의 버전 핀도 갱신:
```
python-hwpx>=새버전
```

## 2단계: GitHub 인사이트 수집 (선택)

관련 저장소의 최근 활동을 확인하여 채택 가치가 있는 패턴이 있는지 검토한다.

### 모니터링 대상 저장소

| 저장소 | 관심 포인트 | 확인 명령 |
|--------|-----------|----------|
| **Canine89/hwpxskill** | 새 검증 패턴, 멀티에이전트 지원 | `gh api repos/Canine89/hwpxskill/commits --jq '.[0:3]\|.[].commit.message'` |
| **openhwp/openhwp** | HWPX 쓰기/변환 개선, 새 포맷 지원 | `gh api repos/openhwp/openhwp/releases --jq '.[0:3]\|.[].tag_name'` |
| **jkf87/hwp-mcp** | MCP 패턴 참고 | `gh api repos/jkf87/hwp-mcp/commits --jq '.[0:3]\|.[].commit.message'` |
| **neolord0/hwplib** | hwp2hwpx 변환기 기반 라이브러리 | `gh api repos/neolord0/hwplib/releases --jq '.[0:3]\|.[].tag_name'` |

### 인사이트 반영 기준

- **즉시 반영**: 우리 스킬의 버그를 수정하거나 안정성을 높이는 패턴
- **검토 후 반영**: 새 기능 추가 (사용자에게 제안 후 결정)
- **참고만**: 아키텍처 인사이트, 향후 로드맵 참고

## 3단계: hwpx_edit.py 동작 확인

```bash
# CLI 기본 동작 확인
python <스킬디렉토리>/hwpx_edit.py --help

# python-hwpx CLI 동작 확인
hwpx-validate --help
hwpx-page-guard --help
```

## 업데이트 이력

| 날짜 | python-hwpx 버전 | 주요 변경 |
|------|-----------------|----------|
| 2026-03-10 | v2.8.2 | Table API, 검증 CLI, 머리글/바닥글, 서식 검색 추가. SKILL.md/GUIDE.md 전면 개편 |
| 2026-05-04 | v2.9.1 | `requires lxml<6` 핀 신규 도입(실측상 lxml 6.x에서도 정상 동작 검증). SKILL.md lxml 호환성 노트 + reference/api.md 버전 라벨 갱신 |

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
