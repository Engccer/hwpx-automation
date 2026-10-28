# hwpx-automation

HWP/HWPX 문서 읽기, 변환, 편집을 위한 CLI 도구 + 코딩 에이전트 스킬

## 기능

- **HWPX → Markdown 변환**: XML 직접 파싱, API 불필요, 무료
- **HWP → HWPX 변환**: 오픈소스 엔진 rhwp로 변환하고, JDK가 있으면 번들 JAR 변환물과 교차 대조
- **HWPX 편집**: 텍스트 치환, 표 셀 채우기, 행 삭제(원본 바이트 보존), 누름틀 채우기, 병합 셀 분할
- **서명·도장 삽입**: 서명란의 "(서명)" 위에 서명 이미지를 겹쳐 찍기
- **PDF 변환**: rhwp(모든 OS, 글꼴에 없는 문자 자동 검출). 한컴과 같은 조판이 필요하면 Windows 한컴오피스
- **배포용 HWP 사본**: 상대가 `.hwp`를 요구할 때 HWPX에서 내보내기(정본은 HWPX)
- **에이전트 스킬**: `SKILL.md`를 통한 AI 에이전트 자동화 워크플로우

## 설치

- Python 3.10+, `pip install -r requirements.txt` (python-hwpx, lxml, hwpx-tomd)
- [rhwp](https://github.com/edwardkim/rhwp/releases) CLI v0.8.7 이상: 변환·PDF·서명의 기본 엔진. 릴리스 바이너리를 PATH에 둔다
- 선택: JDK 21(HWP 변환 교차 대조·`--to-hwp` 왕복 대조·JAR 폴백), Windows + 한컴오피스 + `pywin32`(한컴 COM 폴백)

```bash
python hwpx_edit.py --check-env   # 기능 계층별로 무엇이 바로 되는지 확인 (읽기 전용)
```

## 실행

```bash
python hwpx_edit.py document.hwpx --to-md                          # Markdown으로 읽기
python hwpx_edit.py document.hwp --from-hwp                        # HWP → HWPX
python hwpx_edit.py document.hwpx --find "이전" --replace "이후"    # 텍스트 치환
python hwpx_edit.py document.hwpx --set-cell 0,1,0 "내용"           # 표 셀 입력 (표0, 행1, 셀0)
python hwpx_edit.py document.hwpx --to-pdf -o document.pdf         # PDF
python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(서명)" --pdf   # 서명 삽입
```

결과는 원본을 덮지 않고 입력 폴더의 `_work-hwpx-automation/`에 생긴다(`-o`로 지정 가능). 서명본은 입력 옆 `<입력>_서명.hwpx`. 전체 명령은 `python hwpx_edit.py -h`.

## 에이전트 스킬로 사용

이 저장소를 스킬(`SKILL.md`)을 인식하는 코딩 에이전트의 스킬 디렉토리(Claude Code는 `~/.claude/skills/`)에 배치하면, HWP/HWPX 관련 작업 시 자동으로 `SKILL.md`의 의사결정 트리를 따라 최적의 도구를 선택합니다. <!-- sanitize: allow 설치 대상 안내 -->

```bash
ln -s /path/to/hwpx-automation ~/.claude/skills/hwpx-automation  # sanitize: allow 설치 예시
```

## 문서

- `SKILL.md`: 의사결정 트리와 사용법
- `reference/`: 상세 레퍼런스(rhwp 엔진, 서명, 변환, 한컴 COM, python-hwpx API 등)
- `CLAUDE.md`: 개발 규칙과 함정

일부 레퍼런스 파일(action-table.md, hwp-automation.md, parameterset-table.md)은 [한컴 개발자 센터 아카이브](https://github.com/hancom-io/devcenter-archive)의 PDF를 파싱한 것으로, 저작권 고려로 `.gitignore`에서 제외되어 있습니다. 필요 시 해당 PDF를 직접 다운로드하여 파싱하세요.

## License

이 프로젝트의 코드는 [MIT](LICENSE) 라이선스로 배포됩니다.

단, 이 프로젝트는 다음 서드파티 라이브러리에 의존하며, 각각 별도의 라이선스가 적용됩니다:

| 라이브러리 | 라이선스 | 비고 |
|-----------|----------|------|
| [python-hwpx](https://github.com/airmang/python-hwpx) | **Non-Commercial** | 상업적 사용 불가 |
| [lxml](https://lxml.de/) | BSD-3-Clause | |
| [hwplib](https://github.com/neolord0/hwplib) | Apache-2.0 | JAR 포함 |
| [hwpxlib](https://github.com/neolord0/hwpxlib) | Apache-2.0 | JAR 포함 |
| [hwp2hwpx](https://github.com/neolord0/hwp2hwpx) | Apache-2.0 | JAR 포함 |
| [rhwp](https://github.com/edwardkim/rhwp) | MIT | 별도 설치 CLI (변환·PDF·서명 기본 엔진) |

> **주의**: `python-hwpx`는 비상업적 라이선스입니다. 이 프로젝트를 상업적 목적으로 사용하려면 `python-hwpx`의 라이선스 조건을 별도로 확인하세요.

## 관련 프로젝트

시각장애 사용자를 위한 에이전트 스킬 번들 [skills-for-the-blind](https://github.com/Engccer/skills-for-the-blind)의 멤버 스킬이다.
