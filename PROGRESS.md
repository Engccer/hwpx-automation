# 진행 상태

## 현재

- rhwp 중심 재설계(`docs/plans/2026-10-28-rhwp-redesign.md`) M1~M7 완료. rhwp가 HWP→HWPX 변환·PDF·서명·누름틀·격자 좌표 셀·`--to-hwp`의 기본 엔진이고, JAR·한컴 COM은 명시 폴백, HWPX 미세 편집은 원본 바이트를 지키는 python-hwpx·lxml 경로, 읽기는 hwpx-tomd다. 단위·통합 시험 140건 통과(`RHWP_INTEGRATION=1` 포함).
- 1차 웨이브 codex 횡단 리뷰 판정 block(high 5·medium 4, 보고 `~/hwpx-automation-wt/coordinator-reports/review-cross-wave1.md`). rhwp 편집 게이트가 대상 밖 문단·텍스트 손실을 통과시키는 등 결함이 master에 남아 있고 fix 세션이 수정 중이다(BACKLOG "codex 횡단 리뷰 block 수정").
- rhwp v0.8.7 설치: 맥 `~/.local/bin/rhwp`, Windows hunyong-samsung `C:\Users\pc\.local\bin\rhwp.exe`.
- HWP 변환 번들(JAR 폴백): edc0527+p1 / hwplib 1.1.11 / hwpxlib 1.0.9.

## 실기 확인 대기 (위원장 실사용)

한컴 2024에서 열리고 내용·위치가 맞는지 실사용 중에 확인한다. 열리지 않는 문서군이 나오면 그 축만 COM 폴백을 기본으로 되돌리고 CLAUDE.md에 사유를 적는다.

- rhwp 변환 HWPX(`--from-hwp`)
- rhwp 서명본(`hwpx_sign.py`): 서명이 "(서명)" 위에 있는지. 앵커 문단이 여러 줄인 표본을 하나 포함한다(앵커 문단의 lineseg를 그대로 두므로 한컴이 캐시를 어긋난 것으로 볼 수 있다)
- Pandoc 산출물을 rhwp로 재저장한 HWPX

## 다음 단계

- 열린 항목과 감시 항목은 `docs/BACKLOG.md`.

## 미결 결정

- 없음.
