# 백로그

## 착수 순서

| 묶음 | 항목 | 자리 | 순서 |
|---|---|---|---|
| 변환 | **M2 HWP→HWPX rhwp 기본과 JAR 교차 대조** | 로컬 | 완료 |
| 서명 | **M3 서명 삽입 rhwp 백엔드** | 로컬 | 완료 |
| PDF | **M4 PDF 백엔드 플래그** | 로컬 | 완료 |
| 신규 | **M5 누름틀·격자 좌표·역변환** | 로컬 | 완료 |
| 문서 | **M7 문서 분배** | 로컬 | 신규·서명 뒤 |

M2·M4·M5는 모두 `hwpx_edit.py`를 고치므로 차례로 둔다. 병렬 세션 계획은 `docs/plans/2026-10-28-rhwp-parallel-plan.md`.

## 열린 항목

- [x] **M2 HWP→HWPX rhwp 기본과 JAR 교차 대조**: `--from-hwp`·`--engine rhwp|jar`(기본 rhwp)·`hwp_to_hwpx()`, `cross_check`, `--check-env` Tier 재정의, SKILL.md 예외 문장. 2026-10-28 완료.
- [x] **M3 서명 삽입 rhwp 백엔드**: `--backend rhwp|com`(기본 rhwp), search→render-tree→measure-width→insert-image→앵커 문단 이동, 결과는 항상 `.hwpx`, 게이트 `allow_rhwp`. 2026-10-28 완료.
- [x] **M4 PDF 백엔드 플래그**: `--pdf-backend rhwp|com`(기본 rhwp 모든 OS), Pandoc 산출물 PDF는 rhwp. 2026-10-28 완료.
- [x] **M5 누름틀·격자 좌표·역변환**: `--tables`·`--fields`·`--fill-fields`·`--grid-set-cell`·`--to-hwp`(배포용, JAR 왕복 대조). 2026-10-28 완료. `--grid-set-cell` 실기 통합 시험은 `hwpx_gate.gate(allow_rhwp=)`(M3) 통합 뒤에 돈다.
- [ ] **M7 문서 분배**: SKILL.md 원칙 문장·의사결정 트리·Tier 표, `reference/rhwp.md`, doc-audit. 플랜 M7.
