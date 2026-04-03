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
