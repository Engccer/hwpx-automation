# 변경 이력

## 2026-10-28

- M6: Windows hunyong-samsung에 rhwp v0.8.7(`C:\Users\pc\.local\bin\rhwp.exe`, SHA256 릴리스 대조)을 설치했다. 한컴 2024 호환은 위원장 실사용에서 확인한다.
- rhwp 중심 재설계 M1: `rhwp_cli.py`(어댑터)와 `hwpx_gate.py`(독립 손실 게이트)를 추가하고 `rhwp_pdf.py`를 어댑터 위로 옮겼다. 봉투·exit code 계약 검사, 0.8.7 하한, 임시 폴더→검증→교체를 한곳에 모았다. 플랜 `docs/plans/2026-10-28-rhwp-redesign.md`, codex 적대적 리뷰 반영.
- 실측: rhwp HWPX 왕복은 자가 검증 PASS지만 게이트가 header.xml `hp:switch`·paraPr 손실을 잡는다. HWP 3건의 rhwp·JAR 변환물 교차 대조는 글자·표·셀·그림·BinData 수 일치.
- 버전 1.4.0 → 1.5.0.
- M2: HWP→HWPX 기본 엔진을 rhwp로 바꿨다. `hwpx_edit.py --from-hwp`(`--engine jar` 폴백)와 `hwp_to_hwpx()`를 추가하고, 변환물은 `hwpx-validate`와 다른 엔진 변환물 교차 대조를 통과해야 놓인다. 실패하면 `.rejected.hwpx`를 남기고 JAR 폴백 명령을 안내한다. `.hwp` 입력은 `--to-pdf`를 뺀 모든 명령의 입구에서 변환한다. `--check-env` Tier를 rhwp 기본·JAR·COM 폴백 구조로 재정의했다. 익명 표본 `tests/fixtures/sample-form.hwp`와 통합 시험 `test_rhwp_integration.py`를 추가했다.
- M4: `--to-pdf`를 모든 OS에서 rhwp 기본으로 바꾸고 `--pdf-backend rhwp|com`을 추가했다. `com`은 Windows가 아니면 exit 1로 거부한다. 산출 뒤 stdout 마지막 줄에 `engine=rhwp|com`을 찍는다. Pandoc 산출물 PDF는 rhwp로 낸다고 SKILL.md에 적었다. 표본 HWPX와 Pandoc 산출물 rhwp PDF 빈 네모 0을 통합 시험으로 고정했다.
- M3: `hwpx_sign.py`에 `--backend rhwp|com`(기본 rhwp, 모든 OS)을 두었다. rhwp 경로는 search→render-tree 글자 상자(run 중간은 measure-width)→앵커 중심 겹쳐 찍기→insert-image→그림을 앵커 문단 맨 앞·글 앞으로 옮김→게이트·앵커 재측정. 결과는 늘 `<입력>_서명.hwpx`, `.hwp`는 `hwp_to_hwpx`로 입구 변환, 출력 끝줄 `engine=`. 표 안 앵커·게이트 실패는 `--backend com` 안내.
- `hwpx_gate.gate(allow_rhwp=True)`: 본문이 쓰지 않는 paraPr 안 손실(OUTLINE 8~10단계 switch), 문단 id 재부여, `shadow type="NONE"`을 `GateReport.allowed`로 보내고 "허용된 손실(rhwp 재직렬화)" 줄로 보고한다. `cross_check`는 BinData 내용 해시도 대조한다.
- 실측: rhwp insert-image는 그림을 쪽 첫 문단에 어울림으로 붙여, 첫 문단이 빈 문서는 rhwp 조판이 한 줄 올라간다(앵커 문단으로 옮겨 해결). 한컴 저장 공문 2건은 게이트 통과, 표·글상자 많은 학교 양식 3건은 표 셀 속성·대각선 테두리 손실로 게이트가 막는다. HWP 3건의 rhwp·JAR 변환물 BinData 170개는 바이트가 같다.
- M5: rhwp 축 신규 명령 `--tables`(격자 좌표 보기)·`--grid-set-cell`(격자 좌표 셀 입력)·`--fields`·`--fill-fields`(없는·모호한 이름은 dry-run에서 거부)·`--to-hwp`(배포용 HWP, JAR 왕복 cross_check, 실패 시 한컴 다른 이름 저장 안내)를 추가했다. 편집 산출물은 `hwpx_gate.gate(allow=is_text_run_key, allow_rhwp=True, expect_text_change=True)`를 통과해야 놓이고, 실패하면 `.rejected.hwpx`로 남는다. `.hwp` 입력은 변환을 먼저 거친다. 단위 시험 `test_grid.py`.

## 2026-10-26

- HWP 변환 번들을 상류 edc0527과 hwplib 1.1.11로 갱신했다. 컨트롤 범위 검사 패치는 유지한다.
- Windows에서 HWP 40건의 변환·XML 본문·표 수를 이전 번들과 대조하고 기존 문제 문서의 한컴 COM 변환 결과와 Markdown 일치를 확인했다.
- Windows·Unix 래퍼의 JAR 참조와 번들 라이선스 버전 표기를 맞췄다.
