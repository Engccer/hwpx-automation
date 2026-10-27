# 백로그

## 착수 순서

| 묶음 | 항목 | 자리 | 순서 |
|---|---|---|---|
| 변환 | **M2 HWP→HWPX rhwp 옵트인과 JAR 교차 대조** | 로컬 | 선행 |
| 서명 | **M3 서명 삽입 rhwp 백엔드** | 로컬 | 병행 |
| PDF | **M4 PDF 백엔드 플래그** | 로컬 | 변환 뒤 |
| 신규 | **M5 누름틀·격자 좌표·역변환** | 로컬 | PDF 뒤 |
| 실기 | **M6 Windows rhwp 설치와 한컴 실기 게이트** | 로컬·위원장 | 병행 |
| 문서 | **M7 문서 분배** | 로컬 | 신규 뒤 |

M2·M4·M5는 모두 `hwpx_edit.py`를 고치므로 차례로 둔다. M3은 `hwpx_sign.py`만 고쳐 따로 간다.

## 열린 항목

- [ ] **M2 HWP→HWPX rhwp 옵트인과 JAR 교차 대조**: `--engine rhwp|jar`(기본 jar), `cross_check`, `--check-env` Tier 재정의, SKILL.md 259행 예외. 플랜 M2.
- [ ] **M3 서명 삽입 rhwp 백엔드**: `--backend rhwp|com`(기본 com), search→render-tree→measure-width→insert-image, 결과는 항상 `.hwpx`. 플랜 M3. 선행 판정: header.xml 허용 손실(PROGRESS 미결).
- [ ] **M4 PDF 백엔드 플래그**: `--pdf-backend rhwp|com`, Pandoc 산출물 PDF 라우팅 명시. 플랜 M4.
- [ ] **M5 누름틀·격자 좌표·역변환**: `--tables`·`--fields`·`--fill-fields`·`--grid-set-cell`·`--to-hwp`(배포용, JAR 왕복 대조). 플랜 M5.
- [ ] **M6 Windows rhwp 설치와 한컴 실기 게이트**: hunyong-samsung 설치(`setx` 전 보고), 축별 한컴 2024 열기·위치 확인. 플랜 M6.
- [ ] **M7 문서 분배**: SKILL.md 원칙 문장·의사결정 트리·Tier 표, `reference/rhwp.md`, doc-audit. 플랜 M7.
