# 백로그

## 착수 순서

| 묶음 | 항목 | 자리 | 순서 |
|---|---|---|---|
| 게이트 | **cross_check에 chart·ole·equation 수 대조 추가** | 로컬 | 병행 |

## 열린 항목

- [ ] **cross_check에 chart·ole·equation 수 대조 추가**: `hwpx_gate.cross_check`는 표·셀·그림·BinData만 센다. `--to-hwp`는 `hwpx_edit._embedded_counts()`로 따로 막았지만 HWP→HWPX 변환의 rhwp·JAR 교차 대조에는 차트·OLE·수식이 빠져도 통과하는 빈틈이 있다. `_embedded_counts()`를 `cross_check`로 옮기고 `--to-hwp`는 그것을 쓰게 한다.

## 감시 항목 (조건이 생기면 처리)

- **누름틀 실표본 통합 시험**: `--fields`·`--fill-fields`는 가짜 rhwp로만 시험했다. 누름틀이 든 익명 표본이 생기면 `test_rhwp_integration.py`에 실제 rhwp 시험을 추가한다.
- **암호 HWPX의 rhwp PDF 해제 실측**: `--to-pdf --password`가 rhwp `--password-stdin`으로 암호 문서를 푸는지 표본이 없어 확인하지 못했다. 실패 안내는 Windows `--pdf-backend com`을 가리킨다. 암호 표본이 생기면 재고 `reference/rhwp.md` 폴백 표를 맞춘다.
- **한컴 2024 실사용 확인 결과 반영**: 위원장 실사용 확인(PROGRESS "실기 확인 대기")에서 열리지 않거나 위치가 어긋나는 문서군이 나오면 그 축을 COM 폴백 기본으로 되돌리고 CLAUDE.md에 사유를 적는다.
