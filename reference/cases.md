# 사례

규칙의 근거가 된 실측·사고 경위. 규칙 끝의 "→ 사례"가 여기 같은 제목 절을 가리킨다. 규칙 자체는 SKILL.md와 각 참고 문서에 있다. 날짜와 수치는 그때의 관찰이며, 판이 바뀌면 다시 재야 한다.

## 목차

- [SKILL.md](#skillmd)
- [reference/hwp-conversion.md](#referencehwp-conversionmd)
- [reference/conversion.md](#referenceconversionmd)
- [reference/build-from-scratch.md](#referencebuild-from-scratchmd)
- [reference/api.md](#referenceapimd)
- [reference/structural.md](#referencestructuralmd)
- [reference/signing.md](#referencesigningmd)
- [reference/rhwp-pdf.md](#referencerhwp-pdfmd)
- [reference/com.md](#referencecommd)
- [reference/warnings-com.md](#referencewarnings-commd)
- [reference/warnings-editing.md](#referencewarnings-editingmd)

## SKILL.md

### `--to-md` 보장과 한계

- 2026-06-06 재검증(엔진=hwpx-tomd 패키지): 실문서 33종(워크시트·고사 원안·교육과정·평가계획·체크리스트 등)에서 원본 `<hp:t>` 대비 글자 멀티셋 손실 0·객관식 마커 손실 0으로 문자·마커 단위 완벽 보존을 입증했다(상용 파서 Upstage와 동급 또는 초과).
- 그때 수정된 결함: ① `<hp:t>` tail에 든 선택지 ②③⑤ 누락, ② 글상자(drawText) 본문 대량 누락, ③ cellSpan 오독으로 rowSpan·colSpan이 모두 무시되던 표 정렬 붕괴.
- 2026-09-27 스킬 감사: Pandoc 변환물 fixture(표 밖 제목·본문 문단, 따옴표, 각주, 표)를 `--to-md`로 읽어 표 밖 본문·각주·따옴표 안 글자가 모두 나오고 recall 100%임을 확인했다. `reference/warnings-editing.md` 10번과 `reference/conversion.md`의 "표만 추출"·"따옴표 누락" 서술은 옛 엔진의 것이었다.

### HWP → HWPX 변환

- kordoc 판정(2026-04-30 검증): 표/텍스트 박스가 복잡한 출판사 워크시트·고사지에서 바이너리 잔여 문자 leak, 행 누락, 셀 내용 손실이 발생했다.
- `--from-hwp`(COM 변환) 결과의 `--to-md` recall 100% 실측. hwp2hwpx 변환물은 PrvText가 없어도 `--to-md` recall 100%였다.
- 2026-09-27 스킬 감사: 문서의 수동 클래스패스 `-cp "<convert>/hwp2hwpx-1.0.0.jar;<convert>/lib/hwplib-*.jar;…"`는 JAR 이름이 바뀐 뒤 없는 파일을 가리켰고, `hwplib-*.jar` 같은 부분 와일드카드는 Java가 확장하지 않아 `kr.dogfoot.hwplib` 클래스를 찾지 못했다. `폴더/*`로는 찾았다.

## reference/hwp-conversion.md

### 번들 JAR의 출처·패치

- 상류 커밋 `c9d8a27`은 2026-08-24 커밋이다.
- 패치 계기: 554쪽 연구보고서 실측(2026-08-27)에서 `ForChars.extendControl`이 죽었다. 본문 무손실 근거: 같은 HWP를 한컴 COM으로 변환한 HWPX와 `--to-md` 결과가 글자·표 행 단위로 완전히 일치함을 확인.

### 폴백 1: hwp2hwpx가 죽는 HWP (한컴 COM 경유, 서식 보존)

- `EmptyStackException`(`ForInlineControl.fieldEnd`): 공공기관 안내문에서 실측(2026-07-06).
- `IndexOutOfBoundsException`(`ForChars.extendControl`): 2026-08-27.
- 원격 `schtasks /IT` 실행: 2026-08-27 실측, 554쪽 문서 약 1분.

### 폴백 2: pyhwp 경유 (텍스트만)

- 문자 서식 복원: 2026-08-27, 델파이 조사지의 삭제·변경 표시 복원에 사용.

## reference/conversion.md

### 따옴표 보호 (자동)

- 2026-06-09 도구 반영: 그 전에는 수동 PUA 전처리(U+FFF0~FFF3 마커 치환)·후처리(`postprocess_quotes`로 스마트 따옴표 복원)를 했다. U+FFF0~FFF3은 사설 영역이 아니라 미할당 문자였고, 후처리는 ASCII 따옴표를 스마트 따옴표로 바꿔 원문을 변형했다. 내장 방식은 U+E000~(사설 영역)을 쓰고 원형 그대로 복원한다.
- 2026-09-27 스킬 감사: pypandoc-hwpx 0.1.1 + `hwpx_convert.py`로 fixture를 변환해 따옴표 안 글자 보존, `> ` 인용 블록 누락, 빈 표 칸 subList에 `hp:p` 없음, 각주는 `hp:footNote` 요소, `#`→styleIDRef 2·`##`→3, 여백 left/right 7200·top 4255·bottom 4960·header 4250 HWPUNIT, 표 폭 45000을 확인했다. 옛 서술의 "좌우 72mm, 상 42.55mm, 하 49.6mm"는 HWPUNIT을 100으로 나눈 오기였고, "각주는 문서 끝 텍스트"는 이 판에서 맞지 않았다.

### 장 제목 박스 (표 박스 래퍼) 삽입

- `<hp:ctrl>` 래퍼 크래시·빈 subList id·heading 매핑은 2026-04 실측 확인. 다른 템플릿(중집위 회의자료)에서는 `##`→4 매핑을 봤다.
- 검증 경로(`hwpx-validate` + COM `Open()` + 한글 수동 오픈)는 다건 MD→HWPX 변환 작업에서 확인.

### 표 헤더/합계 행 스타일링

- 2026-09-27 스킬 감사: 예시 `style_table_rows`가 `for tbl_idx, tbl_match in enumerate(tables)` 안에서 `tables`를 다시 계산했다. 반복은 첫 목록을 계속 돌므로, `borderFillIDRef="3"`을 두 자리 id로 바꿔 길이가 달라지면 두 번째 표부터 옛 위치로 잘라 붙인다. build-from-scratch "4. style_tables_xml 루프 버그"와 같은 결함이라 while 루프로 고쳤다.

### 제목 글자색이 파란색으로 나온다

- 2026-07-29에 추가된 함정.

### 인쇄용 배포 문서 디자인

- 실측 교훈 2026-06-01, 2026-06-05 보강. 2026-06-05 성취경험 글쓰기 양식·예시문 비교로 3개 축과 rect 박스 패턴을 추가했다.

### 학생 작성용 양식: `hp:rect` 작성칸/구획 박스

- 2026-06-05. rect 박스 실측: 성취경험 글쓰기 양식(rect 1개, 편지칸)·예시문(rect 2개, 영어/해석 박스).

### `pypandoc.convert_file()` 입력 경로 대괄호 함정

- 2026-04-20 검증.

## reference/build-from-scratch.md

### 핵심 워크플로우: MD → HWPX 생성 (Build-from-scratch 방식)

- 옛 서술은 Pandoc 방식의 한계로 "따옴표 안 텍스트 누락·각주 미변환"을 들었다. 따옴표는 `hwpx_convert.py`가 자동 보호하고, 각주는 pypandoc-hwpx 0.1.1에서 `hp:footNote` 요소로 변환된다(위 conversion.md 절).
- 실제 적용 사례: DPI 보고서 HWPX 생성 스크립트(`generate_hwpx.py`). 입력 534행 마크다운(7장 + 14개 각주, 8개 표), 출력 42KB HWPX(659 단락, 8 표, 22쪽), 스타일 21 charPr·7 borderFill·16 paraPr·맑은 고딕, 스크립트 작성 약 2시간, 이후 재생성 5초 미만.

### 7. ensure_run_style 호환성

- 2026-09-27 스킬 감사: python-hwpx 3.1.0(Mac)·6.3.0(Windows)에서 `HwpxDocument.new()` 문서에 `ensure_run_style(bold=True)`가 정상 동작했다(6.3.0은 `doc.styles.ensure_run`으로 옮긴다는 DeprecationWarning). 열어 둔 기존 문서에서는 다시 재지 않았다.

## reference/api.md

### lxml 6.x

- 2026-05-04: python-hwpx 2.9.1이 `requires lxml<6`을 도입했고, lxml 6.1.0에서 `HwpxDocument.open()` + `hwpx_edit.py --to-md` 재검증 완료.
- 2026-09-27 스킬 감사: 설치본은 Mac python-hwpx 3.1.0(메타데이터 `lxml<7`), Windows 6.3.0. `save_to_path`에 `validate_on_save` 인자가 없고(생성자 인자), `hwpx.oxml.object_finder`는 없으며(`hwpx.tools.object_finder`), 표 객체는 `[t for p in doc.paragraphs for t in p.tables]`로 양쪽에서 얻어짐을 메모리 안 실험으로 확인했다. 옛 예시(`ObjectFinder(doc).find_all(tag="tbl")[0].set_cell_text(...)`, `save_to_path(..., validate_on_save=True)`)는 TypeError였다.
- 표 셀 0건: `doc.text.replace`·`replace_text_in_runs` 모두 6.0.2에서 실측.

## reference/structural.md

### 표에 행을 추가·확장한 뒤 쪽 나눔 (`pageBreak` × `treatAsChar`)

- 내역 19행 문서 실측, 2026-08-10 이슈 #4.

## reference/signing.md

### 서명/도장 이미지 삽입 (hwpx_sign.py)

- 도구 이전에는 다음 함정 때문에 서명 삽입에 30분씩 걸렸다.
- 함정 4(COM이 lineseg를 빼고 저장)는 한컴 13.0 실측. 비가시 창이라 재계산도 하지 않는다.

### 선행 공백은 명백한 사전 신호다

- 2026-07-15 대학 제출 동의서에서 선행 공백 61칸을 흘려보내 이름이 서명에 가려진 실사고. 사용자가 공백 61→39칸으로 당겨 수정. 당시 우측정렬 서명은 horzOffset ≈ 161mm, 폭 22mm였다(horzOffset은 용지 여백에서 계산되고 기본 폭은 20mm).

### 검증은 반드시 PDF로, 그것도 눈으로

- 2026-07-15 실사고: 텍스트 추출로 "성 명 : 홍길동"이 보여 통과로 판단했으나 실제로는 서명 이미지가 그 이름을 덮고 있었다.

### `hp:pic`이 `hp:run` 밖에 놓이면 그림이 조용히 사라진다

- 2026-08-24 서약서 실측, 수정 완료. 앵커 이동 대상인 직전 문단이 빈 문단이면 그 run이 self-closing이라 `</hp:run>` 앞 삽입이 실패하고, 예전 폴백은 pic을 `</hp:p>` 앞(=run 밖)에 넣었다.

## reference/rhwp-pdf.md

### 충실도

- rhwp v0.8.6, 2026-09-17, 한컴 PDF가 함께 있는 성명·답변서 4건 대조: 쪽수 4/4 일치, 글자 누락 0, 줄 단위 일치 96.7~100%. 글꼴 대체 목록은 맥 실측. 쪽 하단 넘침 경고가 떠도 쪽수는 그대로였다.
- v0.8.6에서 `--font-path`·`RHWP_FONT_PATH`·Noto Emoji 사용자 설치가 빈 네모 대체를 바꾸지 못한 것도 실측.

## reference/com.md

### 두 파이프라인 분리 원칙

- 2026-06-10 신설. COM 생성 HWPX → XML 파이프라인 읽기·편집은 recall 100%로 검증(2026-06-10).

### 원격 실행 (SSH)

- 2026-08-27 실측, 554쪽 문서 약 1분.

## reference/warnings-com.md

### 4. COM은 Pandoc 생성 HWPX 본문을 인식하지 못함

- 2026.4 확인. 본문 인라인 삽입 후 별도 파일 SaveAs 보존은 2026.3.31 세션에서 확인, 2026.4.16 재현 성공.

### 6·7. COM 생성 HWPX 호환성·`get_text_file()` 래퍼

- 6번은 2026-06-10 검증, 7번은 2026-06-10 발견.

### 9. 연속 COM 호출 hang

- 2026-08-10 이슈 #3 관찰: 두 번째 호출이 3분 이상 무응답한 사례. 잔류 `Hwp.exe`를 강제 종료한 뒤에는 정상 동작했다(원인 미확정, 잔류 프로세스 추정).

### 10. COM은 자기가 건드린 문단의 lineseg를 저장 시 빼고 쓴다

- 2026-08-12 한컴 13.0.0.3903 실측. `hwpx_sign.py`의 floating 배치가 이 이유로 전면 실패했고, 두 규칙(원본에서 읽기·추정)으로 해결했다. 추정식은 한컴 저장물 2종에서 정확히 일치.

### 11. 글상자·표 안의 lineseg는 개체 기준 상대좌표

- 실측: 서명이 서명줄보다 100mm 위, 가로는 페이지 밖.

### 12. 한컴은 저장하면서 BinData 번호를 재부여한다

- 2026-08-12 실측.

### 15. 변경 추적 문서의 PDF 저장 확인창

- 변경 추적 HWP와 HWPX에서 무인 완료, 수동 저장 PDF와 5쪽 텍스트·렌더 일치를 검증했다.

## reference/warnings-editing.md

### 5. 병합 셀 주의 (span의 위치)

- 과거 `--split-cell` 버그, 이슈 #2로 2026-08-10 수정. 읽기 경로 `get_cell_span`은 2026-06-05에 같은 버그를 먼저 수정했었다.

### 13·14. "한 줄로 입력" 과압축, 기입 예시 글자모양 상속

- 13번은 2026-07-06 보도자료 양식 실측, 14번은 2026-08-10 이슈 #4 실측.

### 15. HWPX를 쓸 때 패키지 규격

- 2026-08-12 이슈 #7 실측, 수정 완료: 과거 `save_hwpx()`는 모든 ZIP 항목을 `ZIP_DEFLATED`로 쓰고 `etree.tostring`에 `standalone=True`를 주지 않았다.

### 16. python-hwpx API 이름은 판마다 다르다

- 2026-08-12 이슈 #6 실측, 수정 완료: `cmd_find_replace`가 `doc.replace_text_in_runs()`와 `doc.save()`를 써서 6.x에서 저장 직전에 `AttributeError`로 죽었다(출력 파일도 생성되지 않음).

### 17. 편집 명령의 저장 경로에 조건을 걸지 말 것

- 2026-08-12 리뷰 지적, 수정 완료: `cmd_find_replace`가 표 치환이 1건 이상일 때만 `save_hwpx()`를 타도록 돼 있어, 본문에만 매칭된 흔한 경우에는 python-hwpx 산출물이 그대로 나갔다.

### 18. hwp2hwpx 변환본의 XML 1.0 불법 제어문자로 파싱 실패

- 2026-08-20 실측, 수정 완료: `--info`·`--set-cell` 등 lxml 경로와 `--find/--replace`의 python-hwpx 경로가 모두 raw 트레이스백으로 죽었다.
