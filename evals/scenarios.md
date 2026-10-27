# hwpx-automation 시험 시나리오
<!-- sanitize: allow-file 시험 전제가 검증 환경(Claude Code)을 명시한다 -->

skill-audit 6·8단계 전후 시험용. 시험자에게는 "시나리오" 절만 주고 "채점 항목"은 주지 않는다. T1~T4는 한 시험자, T5(평범한 하루)는 별도 시험자에게 준다.

공통 전제(시험자에게 준다): 스킬 문서만 읽고 명령은 실행하지 않는 사고 실험으로 답한다. 공식 한컴 규격 사본 3개(`reference/action-table.md`·`hwp-automation.md`·`parameterset-table.md`)는 저장소에 없고 사용자 로컬에만 있다. 에이전트는 Claude Code이며 Windows에서는 Bash 도구가 Git Bash, macOS에서는 zsh다.

## 시나리오

### T1. 가정통신문 MD → 배포용 HWPX·PDF (Windows + 한컴오피스)

학교 가정통신문 `notice.md`가 있다. 본문에 큰따옴표 인용("…")이 많고, 표 하나에 빈 칸이 몇 개 있고, `> ` 인용 블록이 한 개 있다. 스타일은 기본이면 된다. 사용자는 인쇄·배포용 HWPX와 PDF를 원한다. Windows PC에 한컴오피스와 pyhwpx가 설치돼 있다. 어떤 도구·순서로 만들고, 따옴표·빈 칸·인용 블록을 각각 어떻게 처리하며, 배포 직전 한컴 재저장 정규화(`--normalize`)와 PDF는 어떻게 하는가?

### T2. 받은 HWP 신청서 양식 채우기 (Windows, Claude Code)

`D:\업무\2026 신청서.hwp`(한글 파일명·공백 포함)를 받았다. 표의 성명·소속·연락처 칸 세 곳을 채우고, 결과 HWPX를 작업 폴더에 남기려 한다. JDK 21과 rhwp v0.8.7이 설치돼 있다. 변환은 어떤 명령으로 하며, 무엇을 완료 보고에 적는가? rhwp 변환이 실패하면 어떻게 하는가? JAR 래퍼(`hwp2hwpx.bat`)를 직접 불러 만든 변환물에 `hwpx-validate`를 돌렸더니 `Preview/PrvText.txt` 관련 오류가 났다면 어떻게 판단하고 무엇을 하는가? 세 칸을 채울 때 명령을 어떻게 이어 부르는가? 결과물과 부산물은 각각 어디에 남는가?

### T3. 고사지 HWPX 읽기 (macOS)

중간고사 원안 `exam.hwpx`(한글에서 저장)를 마크다운으로 읽어 문항을 검토하려 한다. 문서 제목·과목·출제교사는 표 밖 본문 문단에 있고, 지문은 표 셀 안에 여러 문단으로 들어 있다. 그림 몇 개에 도표 글자가 있다. 어떤 명령을 쓰는가? 표 밖 제목·출제교사가 결과에 나오는가? 이미지 경고가 뜨면 어떻게 하는가?

### T4. 동의서에 서명 넣고 PDF로 (macOS 사용자, 원격 Windows 있음)

사용자는 Mac에서 작업 중이고 `동의서.hwpx`의 "성 명 : ______ (서명)" 칸에 자필 서명 PNG를 넣어 PDF로 제출하려 한다. 이름 줄 앞에 선행 공백이 50칸쯤 있다. 한국 관행대로 "(서명)" 글자 위에 겹쳐 찍기를 원한다. Mac에 rhwp가 설치돼 있고, Windows PC(한컴오피스 있음)에 SSH로 접속할 수도 있다. 어디서 어떤 명령으로 하는가? 결과 HWPX는 어디에 생기는가? 검증은 어떻게 하는가? 게이트가 실패하면 어떻게 하는가?

### T5. 평범한 하루: HWPX 공문 날짜 고치고 표 채우기 (macOS)

한글로 저장된 `회의록.hwpx`(암호 없음)에서 본문과 표에 여러 번 나오는 날짜 "2026. 9. 1."을 "2026. 10. 1."로 바꾸고, 첫 번째 표의 2행 1열·2행 2열·3행 1열 세 칸에 이름을 넣은 뒤 결과를 확인하려 한다. 어떤 명령을 어떤 순서로 부르는가? 결과 파일은 어디에 생기는가? 결과를 어떻게 확인하는가? 답과 함께 **읽은 파일 목록**(파일마다 전부/일부)을 적는다.

## 채점 항목

### T1 (8)
1. 도구: `convert/hwpx_convert.py`(Pandoc) + 후처리. `hwpx_com.py --from-md`를 고르면 그 근거가 "COM 후속 작업 예정"이어야 한다(둘 다 정답 가능. Pandoc을 고르면 아래 2~8로 채점).
2. 따옴표: `hwpx_convert.py`가 자동 보호하므로 수동 PUA 전처리·`postprocess_quotes` 후처리를 하지 않는다.
3. 인용 블록 `> `: 변환 전 수동 처리(마커 치환 또는 일반 문단)가 필요하다.
4. 빈 칸: 변환 직후 `hwpx_edit.py --fix-empty-cells`를 반드시 실행한다(`hwpx_convert.py`는 자동 적용 안 함).
5. 스타일 후처리 등 raw XML 편집을 했으면 lineseg를 제거한다.
6. `--normalize`를 Pandoc 변환물에 쓰지 않는다(COM이 Pandoc 본문을 0자로 읽음 → 빈 문서. 도구 가드도 이 경우를 못 잡는다는 점까지 짚으면 가점 메모).
7. PDF: `hwpx_edit.py --to-pdf`(기본 rhwp, Windows에서도)로 낸다. Pandoc 변환물을 `--pdf-backend com`·`hwpx_com.py --to-pdf`에 넣지 않는다(COM이 본문을 0자로 읽음). 한컴 PDF가 꼭 필요하면 `hwpx_com.py --from-md`로 COM에서 생성한 문서로 낸다.
8. 최종 결과물을 작업 폴더로 두고 중간본은 `_work-hwpx-automation/`에 남긴다.

### T2 (8)
1. 변환: `python hwpx_edit.py "D:\업무\2026 신청서.hwp" --from-hwp`(또는 `.hwp`를 편집 명령에 바로 준다). Python 진입이라 Git Bash에서도 된다. 래퍼 `.bat`를 Git Bash에서 직접 부르지 않는다.
2. 출력 끝줄 `engine=rhwp status=<passed|unverified>`를 완료 보고에 옮긴다. `unverified`를 "검증됨"이라 하지 않는다.
3. rhwp 실패(exit 1)면 `.rejected.hwpx`가 남고 stderr의 JAR 폴백(`--engine jar`)을 따른다. 폴백으로 전환했으면 완료 보고에 엔진과 사유를 쓴다. 조용히 다른 도구로 넘어가지 않는다.
4. PrvText 누락은 JAR 변환물의 알려진 무해 결함으로 판단하고, 정본·배포·검증 대상이면 `--add-preview`로 보정한다(`--engine jar`는 보정을 포함한다. 수동 `touch Preview/PrvText.txt`·XML 직접 편집+`hwpx-pack` 우회는 낡은 조치).
5. 래퍼를 Git Bash에서 굳이 부른다면 PowerShell을 거치거나 ASCII 임시 경로 복사 뒤 `java -cp`를 직접 부르고, 실제 JAR 이름 `hwp2hwpx-1.0.0-edc0527p1.jar`(또는 와일드카드)를 쓴다. 언급이 없으면 0, `hwp2hwpx-1.0.0.jar`를 쓰면 오답.
6. 세 칸: `--set-cell`을 원본에 반복 호출하면 마지막만 남는다 → 두 번째부터 앞 결과를 입력으로 쓰거나 스크립트 한 번으로 채운다.
7. 결과물은 작업 폴더로 옮기거나 `-o`로 지정하고, 변환 HWPX·중간본은 `_work-hwpx-automation/`에 남긴다.
8. 양식 편집 뒤 `--list-squeeze`(과압축)·`hwpx-page-guard`(쪽수) 확인을 든다.

### T3 (5)
1. `hwpx_edit.py exam.hwpx --to-md --cell-br`.
2. 표 밖 본문(제목·출제교사)도 결과에 나온다(텍스트 완전성). "표만 추출된다"고 답하면 낡은 사실 추종.
3. 이미지 속 글자는 추출 범위 밖 → 이미지 경고가 뜨면 OCR·레이아웃 파서로 교차검증.
4. recall 경고(단어·글자·마커)가 뜨면 누락 의심으로 다룬다.
5. 변환 결과가 틀리면 엔진 `hwpx-tomd` 소관이다(이 스킬의 hwpx_edit.py를 고치지 않는다). 시나리오가 묻지 않으므로 언급이 없으면 0, 틀리게 말하면 오답으로 센다.

### T4 (8)
1. Mac에서 바로 한다: `hwpx_sign.py`의 기본 백엔드는 rhwp라 모든 OS에서 돈다. Windows로 옮기거나 SSH·`schtasks`로 COM을 부르지 않는다.
2. 명령: `python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(서명)" --pdf`. 겹쳐 찍기가 rhwp 경로의 기본이라 `--keep-anchor`가 필요 없다(붙여도 무해).
3. 선행 공백 50칸은 사전 신호 → 삽입 전에 `hwpx_edit.py`로 공백을 줄여 이름을 왼쪽으로 당긴다(선제 적용).
4. `--pdf`로 검증 PDF를 만들고 서명 구역을 확대해 렌더 이미지를 육안 확인한다(텍스트 추출로 판단하지 않는다). 맥 PDF는 rhwp 조판이라 한컴과 다를 수 있음을 안다.
5. 결과 HWPX 기본 위치는 입력 옆 `동의서_서명.hwpx`(`_work-hwpx-automation/` 아님). 완료 보고에 출력 끝줄의 `engine=rhwp`를 적는다.
6. 게이트 실패면 산출물이 `_work-hwpx-automation/`에 남고, 표 안 앵커는 삽입 전에 멈춘다. 그때만 Windows에서 `--backend com`(SSH면 `schtasks /IT`)으로 넣는다. 조용히 다른 엔진으로 넘어가지 않는다.
7. "허용된 손실(rhwp 재직렬화)" 줄을 실패로 오해하지 않고, 완료 보고에 그 줄을 그대로 옮긴다.
8. 겹침을 결함으로 보고 "고칠까요"라고 되묻지 않는다.

### T5 (6)
1. 날짜: `hwpx_edit.py 회의록.hwpx --find "2026. 9. 1." --replace "2026. 10. 1."`(본문+표 한 번에).
2. 셀 좌표는 0부터: 2행 1열 = `0,1,0`, 2행 2열 = `0,1,1`, 3행 1열 = `0,2,0`.
3. 이어 부르기: 두 번째 명령부터 앞 결과(`_work-hwpx-automation/회의록.hwpx`)를 입력으로 쓰거나 스크립트 한 번으로 처리한다(원본에 반복 호출하면 마지막만 남음).
4. 결과는 `_work-hwpx-automation/`에 생기고, 최종본은 작업 폴더로 옮기거나 `-o`로 지정한다.
5. 확인: `--to-md`(또는 `--info`)로 텍스트 확인, `hwpx-validate`. 긴 텍스트를 넣었으면 과압축 경고 확인.
6. 읽은 파일: SKILL.md 외에 조건부 참고 문서(COM·변환·build·서명 관련)를 열지 않는다(효율 지표, 정답 수에는 넣지 않고 따로 기록).
