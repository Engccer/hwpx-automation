# 백로그

## 착수 순서

| 묶음 | 항목 | 자리 | 순서 |
|---|---|---|---|
| 변환 | **M2 HWP→HWPX rhwp 기본과 JAR 교차 대조** | 로컬 | 선행 |
| 서명 | **M3 서명 삽입 rhwp 백엔드** | 로컬 | 병행 |
| PDF | **M4 PDF 백엔드 플래그** | 로컬 | 변환 뒤 |
| 신규 | **M5 누름틀·격자 좌표·역변환** | 로컬 | PDF 뒤 |
| 문서 | **M7 문서 분배** | 로컬 | 신규·서명 뒤 |

M2·M4·M5는 모두 `hwpx_edit.py`를 고치므로 차례로 둔다. M3은 `hwpx_sign.py`·`hwpx_gate.py`를 고쳐 따로 간다. 병렬 세션 계획은 `docs/plans/2026-10-28-rhwp-parallel-plan.md`.

## 열린 항목

- [ ] **M2 HWP→HWPX rhwp 기본과 JAR 교차 대조**: `--engine rhwp|jar`(기본 rhwp), `cross_check`, `--check-env` Tier 재정의, SKILL.md 259행 예외. 플랜 M2.
- [ ] **M3 서명 삽입 rhwp 백엔드**: `--backend rhwp|com`(기본 rhwp), search→render-tree→measure-width→insert-image, 결과는 항상 `.hwpx`. 플랜 M3. header.xml 허용 손실은 게이트 허용 목록 + 보고 표기로 확정(2026-10-28).
- [ ] **M4 PDF 백엔드 플래그**: `--pdf-backend rhwp|com`, Pandoc 산출물 PDF 라우팅 명시. 플랜 M4.
- [ ] **M5 누름틀·격자 좌표·역변환**: `--tables`·`--fields`·`--fill-fields`·`--grid-set-cell`·`--to-hwp`(배포용, JAR 왕복 대조). 플랜 M5.
- [ ] **M7 문서 분배**: SKILL.md 원칙 문장·의사결정 트리·Tier 표, `reference/rhwp.md`, doc-audit. 플랜 M7.
