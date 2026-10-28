# 백로그

## 착수 순서

| 묶음 | 항목 | 자리 | 순서 |
|---|---|---|---|
| fix 웨이브 | **codex 횡단 리뷰 block 수정**, **cross_check에 chart·ole·equation 수 대조 추가**, **--check-env 머리줄을 문서와 맞춤** | 로컬 | 선행 |

## 열린 항목

- [ ] **codex 횡단 리뷰 block 수정**: 1차 웨이브 횡단 리뷰(`~/hwpx-automation-wt/coordinator-reports/review-cross-wave1.md`) 9건. high: 편집 허용 콜백이 대상 밖 문단·구역 손실 통과, 대상 밖 텍스트 소실 미검사, `used_parapr_ids` 표기 의존 오분류, 입구 교차 대조의 수식·차트·OLE 공백(아래 항목), 누름틀 결과를 게시 뒤 검사. medium: 서명 게이트 일반 예외 시 거부 산출물 미보존, 고정 이름 재실행 덮어쓰기, `engine=` 끝줄 계약 불완전, 다른 글꼴 폭을 measured로 인정. fix 세션(브랜치 `feat/fix`)이 수정하고 통합 때 이 항목과 SKILL.md·CLAUDE.md 해당 문장을 닫는다.
- [ ] **cross_check에 chart·ole·equation 수 대조 추가**: `hwpx_gate.cross_check`는 표·셀·그림·BinData만 센다. `--to-hwp`는 `hwpx_edit._embedded_counts()`로 따로 막았지만 HWP→HWPX 변환의 rhwp·JAR 교차 대조에는 차트·OLE·수식이 빠져도 통과하는 빈틈이 있다. `_embedded_counts()`를 `cross_check`로 옮기고 `Chart/` 파트 내용 해시도 대조해, 입구(`_independent_check`)와 `--to-hwp`가 같은 보존 계약을 쓰게 한다.
- [ ] **--check-env 머리줄을 문서와 맞춤**: `cmd_check_env()`의 Tier 4 머리줄·주석은 "명시할 때만"이지만 JDK는 있으면 HWP 변환 교차 대조와 `--to-hwp` 왕복 대조에 자동으로 쓰인다(SKILL.md Tier 표는 "폴백, JDK는 있으면 대조에 자동 사용"). Tier 2 머리줄에는 누름틀·격자 셀·`--to-hwp`가 없다. 같은 김에 `rhwp_pdf.py` 모듈 docstring("macOS·Linux에서")과 `hwpx_sign.py` docstring("결과물은 항상")도 맞춘다.

## 감시 항목 (조건이 생기면 처리)

- **누름틀 실표본 통합 시험**: `--fields`·`--fill-fields`는 가짜 rhwp로만 시험했다. 누름틀이 든 익명 표본이 생기면 `test_rhwp_integration.py`에 실제 rhwp 시험을 추가한다.
- **암호 HWPX의 rhwp PDF 해제 실측**: `--to-pdf --password`가 rhwp `--password-stdin`으로 암호 문서를 푸는지 표본이 없어 확인하지 못했다. 실패 안내는 Windows `--pdf-backend com`을 가리킨다. 암호 표본이 생기면 재고 `reference/rhwp.md` 폴백 표를 맞춘다.
- **한컴 2024 실사용 확인 결과 반영**: 위원장 실사용 확인(PROGRESS "실기 확인 대기")에서 열리지 않거나 위치가 어긋나는 문서군이 나오면 그 축을 COM 폴백 기본으로 되돌리고 CLAUDE.md에 사유를 적는다.
