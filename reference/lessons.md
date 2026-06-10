# hwpx-automation 실전 교훈 로그 (lessons)

이 스킬을 실제 작업(파싱·편집·생성·변환)에 쓰며 발견한 함정·우회·실측을 누적한다. SKILL.md "자가 진화 (실사용 교훈 누적)" 섹션의 누적 루프를 따른다. 일반화 가치가 생기면 SKILL.md 본문(의사결정 트리·도구 용도 표·주의사항)이나 해당 `reference/*.md`로 **승격**하고, 여기에는 "승격됨"으로 표시해 둔다.

> 외부 의존성(python-hwpx·hwplib 버전, GitHub 인사이트) 최신화는 본 로그가 아니라 `update-checklist.md` 소관이다. 본 로그는 **우리 도구의 실사용 함정·우회·실측**만 다룬다.

## 항목 형식

날짜 · 작업 맥락 · 증상 · 원인 · 해결/우회 · (가능하면) 도구 개선 제안 · 상태(기록 / 승격 / 도구 반영)

---

## 2026-06-08 · 고사 원안 인쇄 양식(hwpx) 정본화

작업 맥락: 교무부 2026 고사 양식 hwpx를 1학년 영어 빈 양식으로 개조(헤더 치환 + 가이드 30항목·표 10개 본문 제거). 과년도·신규 양식 다수를 `--to-md`로 변환·비교. 스케일이 크고 인쇄 직결이라 완성도 요구가 높은 첫 실전 검증 사례.

### 교훈 1 — hp:t는 서식 강조로 쪼개진다: 부분 문자열 치환 실패

- **증상**: `--find/--replace`와 `xml.replace`로 헤더 "학년 학년 과목"·"2 학년" 치환이 0건.
- **원인**: 한 셀의 표시 텍스트가 강조(charPr 20/21 교차)로 여러 `<hp:run><hp:t>`로 분할되어, 연속 문자열이 태그 경계로 끊겨 매칭되지 않음.
- **해결**: ① 셀 통짜 교체는 `--set-cell r,c,n "값"`이 정답(분할 무시하고 셀 내용 재작성). ② 부분 치환이 꼭 필요하면 run 시퀀스(`<hp:run charPrIDRef="21"><hp:t>학년</hp:t></hp:run>…`)를 정규식으로 잡아 단일 run으로 교체.
- **도구 개선 제안**: hp:t 분할을 인지해 "셀/문단의 표시 텍스트" 단위로 치환하는 옵션(가칭 `--replace-display "old" "new"`)이 있으면 양식 헤더 편집이 쉬워짐.
- **상태**: 기록 (도구 미반영)

### 교훈 2 — `--delete-after`로 양식 본문만 비우기: 유효, secPr 보존

- **맥락**: 가이드 30항목 + 표 10개를 제거하고 헤더·배점표·안내문만 남겨야 함.
- **결과**: `--delete-after "남길 마지막 문구"`로 113개 요소 제거, `hwpx-validate` 통과, 섹션 속성(secPr)·페이지 설정 보존. "특정 지점 이후 본문 비우기"에 안전.
- **상태**: 기록 (의사결정 트리/주의사항 승격 후보)

### 교훈 3 — `--remove-text`는 hp:t만 제거(빈 문단 잔존)

- **증상**: 안내 문단을 `--remove-text`로 지우면 hp:t는 사라지나 빈 `<hp:p>`(빈 줄)가 남음.
- **해결**: 문단째 제거는 hp:p 단위 정규식(`<hp:p …>(?:(?!</hp:p>).)*?marker.*?</hp:p>`, DOTALL)으로 직접 편집.
- **도구 개선 제안**: `--remove-para "텍스트"`(텍스트를 포함한 문단 통째 제거) 옵션.
- **상태**: 기록

### 교훈 4 — 직접 zipfile 재패키징 시 mimetype 규칙

- **핵심**: section0.xml을 직접 편집·재저장할 때 `mimetype`을 **첫 항목·`ZIP_STORED`**로 써야 한다. 어기면 한글에서 열리지 않거나 손상 경고.
- **상태**: 기록 (`format.md` 승격 후보)

### 교훈 5 — 한글 CLI 인자: 정상, 단 '-' 시작 주의

- **실측**: PowerShell → `hwpx_edit.py`의 한글 인자(`--delete-after`/`--set-cell`/`--remove-text`)는 정상 작동(Windows argv가 wide-char).
- **함정**: '-'로 시작하는 검색어(예: `"-서답형 …"`)는 argparse가 옵션으로 오인 → 따옴표로도 충돌. '-' 없는 부분 문자열로 우회.
- **상태**: 기록

### 검증 메모

`--to-md`(hwpx-tomd) self-recall + `hwpx-validate`를 각 편집 단계마다 돌려 무손실·무결성을 확인했다. 양식 개조처럼 비파괴가 중요한 작업의 표준 절차로 둔다.

---

## 2026-06-09 · 다단(2단) 양식 채우기: 자동 변환 금지

작업 맥락: 2026 1학년 영어 기말 원안(md)을 2단 인쇄 hwpx로 변환. 양식은 교무부 2026 공식 양식(헤더표 + 배점표 + 빈 2단 본문, `--delete-after`로 본문만 비운 상태).

### 교훈 6 — 다단 양식은 Pandoc/build-from-scratch로 새로 찍으면 단(段)이 날아간다 → open-fill만

- **증상/위험**: 2단 양식을 두고도 `hwpx_convert.py`(Pandoc)나 build-from-scratch로 **새 문서를 생성**해 본문을 만들면 결과물이 단단(1단)이 된다.
- **원인**: `section0.xml`의 `<hp:colPr ... colCount="2" ...>`(다단)은 **secPr에 박힌 섹션 속성**이다. 새 문서 생성 경로는 기본이 단단이라, `set_columns(2)`를 명시 호출하지 않는 한 colPr이 안 붙는다. 거리·운으로 못 막는 함정.
- **해결**: 다단을 유지해야 하는 양식 채우기는 **반드시 `HwpxDocument.open(양식.hwpx)`로 기존 파일을 열어** 표·문단·글상자만 추가/치환한다. `--delete-after`로 본문만 비운 양식(교훈 2)은 secPr·colPr이 보존되므로 그 위에 채우면 2단이 유지된다(실측: 양식 colPr `type="NEWSPAPER" colCount="2" sameGap="1420"≈5mm`).
- **도구 개선 제안**: `--info` 출력에 secPr 단 수(colCount)를 노출하면 "이 양식이 2단인지" 즉시 확인 가능.
- **상태**: 기록 (의사결정 트리 "양식 채우기" 분기 승격 후보). 고사문항출제 스킬 `레이아웃-조판-예산.md` "위험 3지점 #1"로 동시 반영됨.

### 교훈 7 — add_table(para_pr_id_ref=)는 셀 문단 정렬에 적용되지 않는다

- **증상**: `doc.add_table(..., para_pr_id_ref=3)`으로 표 셀 정렬을 지정해도, 생성된 셀 문단은 기본 paraPr 0으로 만들어진다(고사 2단 지문 박스가 의도(JUSTIFY)와 달리 CENTER로 렌더).
- **해결**: 셀 문단별로 직접 지정한다 — 첫 문단 `cell.paragraphs[0].para_pr_id_ref = N`, 추가 문단 `cell.add_paragraph('', para_pr_id_ref=N)`. 정렬값은 header.xml `<hh:paraPr><hh:align horizontal="LEFT|CENTER|JUSTIFY|RIGHT"/>`로 확인. 양식에 LEFT가 없으면 JUSTIFY(영문 본문 표준)를 쓴다.
- **상태**: 승격됨 (SKILL.md "테이블 API" 셀 문단 정렬)

### 교훈 8 — 부분 서식(밑줄·이탤릭·볼드)은 para.add_run(...)로

- **API**: `p = doc.add_paragraph(''); p.add_run(seg, bold=True, underline=True, italic=True)`. add_run이 charPr를 자동 생성/재사용하므로 `ensure_run_style`을 따로 부르지 않아도 된다. 표 셀도 `cell.paragraphs[0].add_run(...)` / `cell.add_paragraph('').add_run(...)`로 동일.
- **활용(고사 원안)**: 텍스트의 마커를 파싱해 토큰별 run으로 — `<u>…</u>`→볼드+밑줄(부정 발문·어법 밑줄·지문 내 밑줄·우리말 영작), `*…*`→이탤릭(작품명), `**…**`→볼드(문항 번호·배점). 밑줄 답란(`____`)이 든 줄은 양쪽정렬에서 글자 사이가 벌어지므로 그 줄만 셀 기본 정렬로 둔다.
- **상태**: 승격됨 (SKILL.md "부분 서식 — 밑줄·이탤릭·볼드 (run 단위)")

### 교훈 9 — Pandoc 변환 표의 열폭·행높이·셀 다문단은 raw lxml 후처리

- **맥락**: 서답형 답안지(Pandoc 변환)는 6열 균등(답란 좁음)·셀 1문단이라 학생 작성 공간이 부족.
- **해결**: section0.xml을 lxml로 열어 답란표(헤더 텍스트로 식별)의 `<hp:cellSz width/height>`를 colAddr/rowAddr별로 재배분하고, 답란 셀 `<hp:p>`를 deepcopy로 복제해 "(1)"·빈·"(2)"·빈 다문단으로 만든다. 텍스트 바꾼 문단은 `<hp:linesegarray>` 제거. mimetype을 첫 항목·ZIP_STORED로 재패키징(교훈 4).
- **상태**: 승격됨 (conversion.md "표 열 너비 조정" 도입부)

---

## 2026-06-09 · hwp2hwpx 변환물의 Preview/PrvText.txt 누락

작업 맥락: 교무부 2026 고사 양식 HWP(서답형 채점기준표·답안지 등)를 `hwp2hwpx.bat`로 변환해 고사문항출제 스킬의 정본 양식으로 채택. 채택 전 `hwpx-validate`로 무결성 확인.

### 교훈 10 — hwp2hwpx 출력은 container.xml이 PrvText.txt를 선언하지만 파일이 없어 hwpx-validate 실패

- **증상**: `hwp2hwpx.bat` 변환물 전부가 `hwpx-validate`에서 `HwpxStructureError: Root content 'Preview/PrvText.txt' declared in container.xml is missing.`로 실패. 단 **한글에서는 정상 열림**, `--to-md`도 recall 100%(읽기·파싱엔 무해).
- **원인**: 변환물 `META-INF/container.xml`의 `<ocf:rootfiles>`가 `Preview/PrvText.txt`(+미선언이지만 통상 PrvImage.png)를 rootfile로 선언하는데, ZIP에 `Preview/` 항목 자체가 없다(엔트리 8개: mimetype·version·manifest·container·content.hpf·header·section0·settings). `manifest.xml`은 빈 껍데기라 교차 선언도 없음. python-hwpx 검증기는 container rootfile 존재를 강제해 실패.
- **해결(우회)**: `section0.xml`의 `<hp:t>` 텍스트를 모아 `Preview/PrvText.txt`를 만들어 ZIP에 주입(mimetype 첫 항목·`ZIP_STORED` 유지, 교훈 4). 주입 후 3종 모두 `hwpx-validate` 통과 + recall 100% 유지. (대안: container.xml에서 PrvText rootfile 선언 줄을 제거해도 일관성 회복. 실제 한글 hwpx는 PrvText를 가지므로 주입이 더 충실.)
- **적용 범위 메모**: 이 누락은 hwp2hwpx 출력 **전반의 특성**이다(같은 배치로 만든 고사 양식 hwpx 12종 모두 해당). 읽기·편집·인쇄엔 무해하므로 **검증 통과가 필요한 정본/배포 산출물에서만** 보정하면 된다.
- **상태**: **도구 반영됨** (`hwpx_edit.py --add-preview`, 2026-06-09). section XML 재직렬화 없이 `<hp:t>` 텍스트로 `Preview/PrvText.txt`를 생성해 mimetype 첫 항목·ZIP_STORED로 재패키징하며 주입한다. 멱등(이미 있으면 변경 없음). SKILL.md "hwpx_edit.py 사용법"·"HWP → HWPX 변환" 주의에 반영. 남은 후보: `hwp2hwpx.bat`(hwpxlib) 단계에서 PrvText/PrvImage를 애초에 생성하도록 보강(자바 측).

---

## 2026-06-09 · Pandoc 변환의 따옴표 쌍 안 텍스트 누락 (실제 제출 산출물 손상)

작업 맥락: 고사 서답형 채점기준표(md→hwpx, Pandoc 경로)를 제출 산출물로 생성. 출제교사 날인행 추가차 재검토하다 본문 누락을 발견.

### 교훈 11 — Pandoc HWPX writer가 따옴표 "쌍" 안 텍스트를 통째 누락 → recall로 못 잡힘

- **증상**: 채점기준표 hwpx에서 `우리말 "일어나서 ~ 즐길"에 따라 'get up early → enjoy ~' 순서` → 변환물은 `우리말 에 따라 순서`로, 따옴표 쌍 안 텍스트가 통째로 사라짐. `'This morning, Mia was late for school.'`도 소실. 이미 폴더에 제출 예정으로 놓여 있던 산출물이 조용히 손상돼 있었다.
- **원인**: Pandoc HWPX writer(pypandoc-hwpx)가 따옴표 쌍을 만나면 그 안 텍스트를 출력에서 누락시킨다. SKILL.md가 "전처리 필수"로 경고했으나 수동 전처리를 빠뜨리면 그대로 통과한다. **`--to-md` recall은 hwpx→md 방향만 검증**하므로(원본 md→hwpx 손실은 보지 않음) recall 100%로도 이 손실을 못 잡는다. 아포스트로피(’ in doesn't)는 쌍이 아니라 보존됐다(혼동 주의).
- **해결(도구 반영)**: `hwpx_convert.py`에 **따옴표 보호를 기본 내장**. 변환 전 6종 따옴표(“”‘’"')를 고유 PUA(U+E000~E005)로 치환해 Pandoc이 쌍으로 인식하지 못하게 하고, 변환 후 결과 hwpx의 `Contents/*.xml`에서 PUA를 원래 따옴표로 정확히 원복한다(mimetype 첫 항목·ZIP_STORED 유지, 교훈 4). 입력에 따옴표가 없으면 무동작, `--no-quote-fix`로 비활성화. 재생성 후 따옴표 12종 구절 전부 보존·recall 100%·validate 통과 확인.
- **검증 보강 메모**: recall이 100%여도 **md→hwpx 방향 손실은 별도 확인**이 필요하다. 제출/정본 산출물은 변환 후 소스 md의 따옴표 쌍 구절이 hwpx 텍스트에 모두 있는지 대조한다(주석 `<!-- -->` 제외 후 검사). 같은 점검으로 고사 원안.hwpx는 12개 구절 누락 0(양식 채우기/API 경로라 무손상)으로 확인.
- **상태**: **도구 반영됨** (`hwpx_convert.py` 따옴표 자동 보호, 2026-06-09). SKILL.md "MD → HWPX 변환(Pandoc 방식)" 전처리 단계·비교표에 반영. 남은 후보: ① `_restore_quotes_in_hwpx` 후처리를 `--to-md` 자가검증처럼 "소스 대비 따옴표 구절 보존" 가드로 옵션화, ② `--add-preview -o` 무변경 시에도 출력 파일을 생성하도록 보완(이번에 cg3 미생성).

---

## 2026-06-10 · COM 네이티브 파이프라인 신설 (hwpx_com.py)

작업 맥락: "프로덕션급 HWPX 생성·편집" 요구에 대한 구조적 답으로, 기존에 진단·PDF로만 좁혀 두었던 COM 분기를 pyhwpx 기반의 두 번째 백엔드로 정식 확장. 기존 hwpx_edit.py에 플래그를 얹지 않고 별도 스크립트로 분리해 두 파이프라인 혼용을 파일 수준에서 차단.

### 교훈 12 — 호환성은 단방향: COM 생성물은 양쪽에서 유효, Pandoc 생성물만 COM 금지

- **실측**: pyhwpx 1.7.2로 "새 문서 + 텍스트 + 2x3 표 + SaveAs(HWPX)" 왕복 테스트. ① COM 재오픈 시 본문 전체 정상(0자 문제 없음) ② hwpx-tomd 변환 recall 단어·글자 100%, 표 구조 보존. MD 부분집합(제목 3단계·굵게 인라인·불릿·파이프 표) 렌더링도 recall 100%.
- **결론**: 금지 방향은 "Pandoc 생성물 → COM"(기존 교훈) 하나뿐. COM 생성물은 XML 파이프라인에서 자유롭게 읽기·편집 가능. hwpx_com.py는 사고 방지를 위해 입력 in-place 저장을 코드에서 금지(별도 출력 강제).
- **상태**: **도구 반영됨** (`hwpx_com.py` 신설: --diagnose/--from-md/--insert-image/--get-text/--to-pdf). SKILL.md "두 파이프라인 분리 원칙"·의사결정 트리·warnings-com.md 6번에 반영.

### 교훈 13 — pyhwpx get_text_file()은 기본이 saveblock:true (선택 없으면 None)

- **증상**: COM으로 연 문서에서 `hwp.get_text_file()`이 None 반환 → `len()` 크래시.
- **원인**: 래퍼 기본 인자가 `option='saveblock:true'`(선택 블록만). 선택이 없으면 빈 결과.
- **해결**: 전체 본문은 저수준 `hwp.hwp.GetTextFile("TEXT", "")` 직접 호출. hwpx_com.py --get-text가 이 방식.
- **상태**: **도구 반영됨** + warnings-com.md 7번 승격.
