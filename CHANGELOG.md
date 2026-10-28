# 변경 이력

## 2026-10-28

rhwp 중심 재설계. 플랜 `docs/plans/2026-10-28-rhwp-redesign.md`, 병렬 계획 `docs/plans/2026-10-28-rhwp-parallel-plan.md`.

- 위원장 결정: 축별 옵트인 단계 없이 rhwp를 기본값으로 두고 JAR·COM은 명시 폴백으로 남긴다. 실기 검증은 위원장 실사용으로 한다. rhwp 재직렬화의 알려진 손실은 게이트 허용 목록으로 통과시키고 보고에 표기한다.

- M1 어댑터·게이트: `rhwp_cli.py`(유일한 rhwp 입구: 봉투·exit code 계약 검사, 0.8.7 하한, 임시 폴더→검증→교체)와 `hwpx_gate.py`(요소 census·BinData 해시·글자·`hwpx-validate` 독립 손실 게이트)를 추가하고 `rhwp_pdf.py`를 어댑터 위로 옮겼다. codex 적대적 리뷰 반영. 버전 1.4.0 → 1.5.0.
  - 실측: rhwp HWPX 왕복은 자가 검증 PASS지만 게이트가 header.xml `hp:switch`·paraPr 손실을 잡는다. HWP 3건의 rhwp·JAR 변환물은 글자·표·셀·그림·BinData 수가 같다.
- M2 HWP→HWPX rhwp 기본: `hwpx_edit.py --from-hwp`(`--engine jar` 폴백)와 `hwp_to_hwpx()`. 변환물은 `hwpx-validate`와 JAR 변환물 교차 대조를 통과해야 놓이고(JDK 없으면 unverified), 실패하면 `.rejected.hwpx`와 JAR 폴백 명령을 남긴다. `.hwp` 입력은 `--to-pdf`를 뺀 모든 명령의 입구에서 변환한다. `--check-env` Tier를 rhwp 기본·JAR·COM 폴백 구조로 재정의. 익명 표본 `tests/fixtures/sample-form.hwp`, 통합 시험 `test_rhwp_integration.py`.
- M3 서명 삽입 rhwp 백엔드: `hwpx_sign.py --backend rhwp|com`(기본 rhwp, 모든 OS). search→render-tree 글자 상자(run 중간은 measure-width)→앵커 중심 겹쳐 찍기→insert-image→그림을 앵커 문단 맨 앞·글 앞으로→게이트·앵커 재측정. 결과는 늘 `<입력>_서명.hwpx`, `.hwp`는 `hwp_to_hwpx`로 입구 변환, 출력 끝줄 `engine=`. 표 안 앵커·게이트 실패는 `--backend com` 안내.
  - `hwpx_gate.gate(allow_rhwp=True)`: 본문이 쓰지 않는 paraPr 안 손실(OUTLINE 8~10단계 switch), 문단 id 재부여, `shadow type="NONE"` 소실을 `GateReport.allowed`로 보내고 "허용된 손실(rhwp 재직렬화)" 줄로 보고한다. `cross_check`는 BinData 내용 해시도 대조한다.
  - 실측: rhwp insert-image는 그림을 쪽 첫 문단에 어울림으로 붙여, 첫 문단이 빈 문서는 rhwp 조판이 한 줄 올라간다(앵커 문단으로 옮겨 해결). 한컴 저장 공문 2건은 게이트 통과, 표·글상자 많은 학교 양식 3건은 표 셀 속성·대각선 테두리 손실로 막힌다. HWP 3건의 rhwp·JAR 변환물 BinData 170개는 바이트가 같다.
- M4 PDF 백엔드: `--to-pdf`를 모든 OS rhwp 기본으로, `--pdf-backend rhwp|com` 추가(`com`은 Windows 밖에서 exit 1). stdout 마지막 줄 `engine=rhwp|com`. Pandoc 산출물 PDF는 rhwp. 표본 HWPX와 Pandoc 산출물 rhwp PDF 빈 네모 0을 통합 시험으로 고정.
- M5 누름틀·격자 좌표·역변환: `--tables`·`--grid-set-cell`·`--fields`·`--fill-fields`(없는·모호한·닮은 이름은 dry-run에서 거부)·`--to-hwp`(배포용 HWP, JAR 왕복 `cross_check`와 차트·OLE·수식 수 대조, 실패 시 한컴 다른 이름 저장 안내). 편집 산출물은 `gate(allow=is_text_run_key, allow_rhwp=True, expect_text_change=True)`를 통과해야 놓이고(실패하면 `.rejected.hwpx`) 허용된 손실 줄은 stdout에 보인다. `.hwp` 입력은 변환을 먼저 거친다. 단위 시험 `test_grid.py`.
- M6: Windows hunyong-samsung에 rhwp v0.8.7(`C:\Users\pc\.local\bin\rhwp.exe`, SHA256 릴리스 대조) 설치. 한컴 2024 호환은 위원장 실사용에서 확인한다.
- M7 문서 분배: `reference/rhwp-pdf.md`를 `reference/rhwp.md`로 확장(설치·호출 계약·게이트·좌표 단위·명령별 실측 계약·폴백 조건·PDF). SKILL.md 정본 형식 원칙·의사결정 트리·도구 표 엔진 열·Tier 표·완료 보고 문장, CLAUDE.md rhwp 함정 절, README 축약, evals 기대 답 갱신.
- fix 웨이브(codex 횡단 리뷰 block 9건): 판정 계층을 바꿨다. rhwp 편집 산출물은 census+허용 콜백(`is_text_run_key`, 제거) 대신 `gate(allow_rhwp=True, edits=[...])`로 `section*.xml`을 순서 보존 토큰열로 위치별 대조한다. 대상(`CellEdit`·`FieldEdit`·`ImageInsert`)은 원본 XML과 요청값으로 정하고 칸 안쪽·누름틀 값 자리·삽입 그림만 가린다(codex 적대적 설계 리뷰 반영: 문단 census는 서식 맞바꿈을, 문단 단위 허용은 누름틀 앞뒤 글자 이동을 못 막는다). 허용 목록에 `fieldBegin zorder="-1"` 소실(실측) 추가, paraPr 사용 판정은 XML 속성 파싱.
  - `rhwp_cli.run`: `verify(staged, envelope)`로 봉투 사후조건을 게시 전에 검사, 어떤 verify 예외·봉투 계약 오류든 산출물을 `<stem>.rejected-<시각>`으로 보존. rhwp 기본 출력(입구 변환물·`_편집`·`--to-hwp`·`_서명`)은 있으면 덮지 않는다. `--to-md`는 `engine=hwpx-tomd`, 명시 폴백은 `reason=사용자 지정`, 서명은 `engine=rhwp status=`.
  - `cross_check`가 chart·ole·equation 수와 `Chart/` 파트 내용 해시를 대조하고 HWP 입구와 `--to-hwp`가 함께 쓴다(`_embedded_counts` 제거). `--check-env` Tier 2·4 머리줄, `rhwp_pdf.py`·`hwpx_sign.py` docstring을 문서와 맞췄다.
  - 서명: run 중간 앵커는 함초롬바탕·자간 0·보통 굵기일 때만 measured(`measure-width --font`는 모르는 글꼴도 대체 폭을 낸다, 실측). 새 게이트가 `float_over_anchor`의 빈 run `<hp:t/>` 무조건 복원을 잡아 원본 모양대로 고쳤다.
  - 최종 검토(fable)가 section 밖 구멍을 찾아 막았다: `Chart/`·`DocHistory/`·`Scripts/` 등 그 밖 파트는 이름·바이트 대조, `content.hpf`는 메타데이터·매니페스트·구역 순서 대조, 바탕쪽은 토큰 대조, header census는 소속 조상·글자를 키에 넣고 미사용 paraPr 허용을 개요 갈래 접힘으로 좁혔다. 칸 안 중첩 표는 글자까지 비교.
  - 한컴 저장 실문서 12건 rhwp 왕복으로 재측정: 위치별 대조가 run·t 경계 재분할·빈 run·빈 t·`\r\n`·그림 안 자식 순서·필드 zorder 때문에 12건 모두 거부하던 것을 정규화해 10건 통과, 남은 2건은 진짜 변화(그림 thumbnailBinIDRef 소실, 필드 끝 위치 이동). paraPr `textDir` LTR·AUTO 소실 허용으로 python-hwpx 산출물도 rhwp 편집이 된다(함정 항목 삭제).
  - 시험: 가짜 rhwp가 실제로 바꾼 산출물을 내고 대상 밖 공백화·문단 삭제·secPr 소실·다른 칸·다른 글꼴을 거부. 누름틀 익명 표본 `tests/fixtures/sample-fields.hwpx`로 실호출 통합 시험(`이름[k]`는 0부터, 실측). 리뷰(opus) 반영: 같은 `.hwp` 재실행은 원본보다 새 변환물을 재사용(다시 변환하지 않음), JAR 입구의 검증 예외도 거부 산출물 보존, 누름틀 값은 한 줄 글자·숫자만(쓰기 전 거부). 정규화로 같게 본 차이도 종류별 건수로 허용 줄에 보고하고 항목마다 거부 반례 시험을 둔다. 단위·통합 184건.

## 2026-10-26

- HWP 변환 번들을 상류 edc0527과 hwplib 1.1.11로 갱신했다. 컨트롤 범위 검사 패치는 유지한다.
- Windows에서 HWP 40건의 변환·XML 본문·표 수를 이전 번들과 대조하고 기존 문제 문서의 한컴 COM 변환 결과와 Markdown 일치를 확인했다.
- Windows·Unix 래퍼의 JAR 참조와 번들 라이선스 버전 표기를 맞췄다.
