# hwpx-automation

HWP/HWPX 문서 읽기, 변환, 편집을 위한 CLI 도구 + Claude Code 스킬

## 기능

- **HWPX → Markdown 변환**: XML 직접 파싱, API 불필요, 무료
- **HWP → HWPX 변환**: Java 기반, 서식 100% 보존
- **HWPX 편집**: 텍스트 치환, 표 셀 채우기, 병합 셀 분할
- **Claude Code 스킬**: `SKILL.md`를 통한 AI 에이전트 자동화 워크플로우

## 빠른 시작

```bash
# 의존성 설치
pip install python-hwpx lxml

# HWPX → Markdown 변환
python hwpx_edit.py document.hwpx --to-md

# 문서 구조 확인 (표 목록, 셀 구조)
python hwpx_edit.py document.hwpx --info

# 텍스트 치환
python hwpx_edit.py document.hwpx --find "이전 텍스트" --replace "새 텍스트"

# 표 셀에 텍스트 입력 (표0, 행1, 셀0)
python hwpx_edit.py document.hwpx --set-cell 0,1,0 "내용"

# 병합 셀 분할
python hwpx_edit.py document.hwpx --split-cell 0,1,0
```

## HWP → HWPX 변환

```bash
# JDK 21 필요
convert/hwp2hwpx.bat input.hwp output.hwpx
```

## Claude Code 스킬로 사용

이 저장소를 Claude Code의 스킬 디렉토리에 배치하면, HWP/HWPX 관련 작업 시 자동으로 `SKILL.md`의 의사결정 트리를 따라 최적의 도구를 선택합니다.

```bash
# 스킬 디렉토리에 심링크 (예시)
ln -s /path/to/hwpx-automation ~/.claude/skills/hwpx-automation
```

## 시스템 요구사항

- Python 3.10+
- JDK 21 (HWP → HWPX 변환 시)
- `python-hwpx`, `lxml`

## 프로젝트 구조

```
hwpx-automation/
├── hwpx_edit.py          # HWPX 읽기/편집 CLI
├── SKILL.md              # Claude Code 스킬 정의
├── convert/
│   ├── hwp2hwpx.bat      # HWP→HWPX 변환 (Windows)
│   ├── hwp2hwpx-1.0.0.jar
│   └── lib/              # hwplib, hwpxlib
├── reference/            # 상세 레퍼런스 문서
│   ├── api.md
│   ├── structural.md
│   ├── format.md
│   └── ...
└── requirements.txt
```

## 참조 문서

`reference/` 폴더에 상세한 API 문서, 구조적 편집 가이드, 파일 형식 설명이 있습니다.

일부 레퍼런스 파일(action-table.md, hwp-automation.md, parameterset-table.md)은 [한컴 개발자 센터 아카이브](https://github.com/hancom-io/devcenter-archive)의 PDF를 파싱한 것으로, 저작권 고려로 `.gitignore`에서 제외되어 있습니다. 필요 시 해당 PDF를 직접 다운로드하여 파싱하세요.

## License

[MIT](LICENSE)
