# 서명/도장 이미지 삽입 (hwpx_sign.py)

동의서·계약서·서약서 등의 서명란("(서명)" 같은 기준 텍스트)에 자필 서명/도장 이미지를 넣는 작업. 기본 백엔드는 **rhwp**이고 macOS·Linux·Windows 어디서나 돈다(rhwp CLI v0.8.7 이상). 한컴 COM 경로는 `--backend com`으로 고를 때만 쓴다(Windows + 한컴오피스).

결과물은 기본 입력 옆 `<입력>_서명.hwpx`다(이미 있으면 덮지 않고 exit 1, `-o`로 다른 `.hwpx` 지정 가능). `.hwp`를 넣으면 입구에서 `hwpx_edit.hwp_to_hwpx`로 HWPX로 바꾼 뒤 작업한다. 출력 끝줄은 `engine=rhwp status=<passed|unverified>` 또는 `engine=com reason=사용자 지정`이며, 완료 보고에 그대로 적는다.

```bash
# 기본: "(서명)" 글자 위에 겹쳐 찍기(rhwp, 모든 OS)
python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(서명)" --pdf

# 앵커가 여러 곳이면 N번째를 고른다(첫 실행이 위치 목록을 보여 준다)
python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(인)" --occurrence 2 --pdf

# 위치 미세조정(HWPUNIT, 1mm = 283.46)
python hwpx_sign.py 동의서.hwpx --image 서명.png --width-mm 22 --vert-adjust -800 --pdf

# 본문 우측 끝에서 7mm 안쪽으로 우측 정렬
python hwpx_sign.py 동의서.hwpx --image 서명.png --gap-mm 7 --pdf
```

옵션: `--anchor`(기본 `(서명)`) · `--occurrence N`(1부터) · `--width-mm`(기본 20) · `--height-mm`(미지정 시 종횡비. Pillow가 없으면 폭의 0.65배) · `--horz-offset`(종이 왼쪽 기준 HWPUNIT, 절대 가로 위치) · `--vert-adjust`(HWPUNIT, +아래/−위) · `--gap-mm`(주면 본문 우측 정렬) · `--allow-estimated` · `--backend rhwp|com` · `-o`(`.hwpx`만) · `--pdf`.

## rhwp 경로가 하는 일

1. `rhwp search`로 앵커의 쪽·문단·글자 위치를 찾는다. **0건이나 2건 이상이면 멈춘다**(2건 이상은 `--occurrence`로 고른다). 표 안 앵커는 아직 지원하지 않아 `--backend com`을 안내한다.
2. `rhwp export-render-tree`에서 같은 문단의 본문 글자 묶음(TextRun) 상자를 읽어 앵커 글자의 화면 상자를 구한다. 앵커가 묶음 중간이면 `rhwp measure-width`로 앞부분 폭을 재고 WARN을 낸다. 그 묶음이 함초롬바탕(장평 고르게, 자간 0, 보통 굵기)이 아니거나 판정하지 못하거나 폭을 재지 못하면 글자 수 비례로 추정하는데(estimated), 이때는 `--allow-estimated` 없이는 넣지 않는다.
3. 기본 배치는 **앵커 상자 중심에 서명 중심을 겹친다**(한국 결재 관행: "(서명)"·"(인)" 표시를 지우지 않고 그 위에 찍는다. 글자와 서명이 포개져 보이는 것이 정상이며 결함이 아니다. 겹침을 보고 "고칠까요"라고 되묻지 않는다). 세로는 앵커 줄의 중앙이다.
4. `rhwp edit insert-image`로 넣은 뒤, 그림을 **앵커 문단 맨 앞**으로 옮기고 글 앞으로(`IN_FRONT_OF_TEXT`, 겹침 허용) 놓는다. rhwp는 그림을 그 쪽 첫 문단에 어울림으로 붙이는데, 첫 문단이 빈 문단이면 rhwp 조판이 그 줄을 접어 아래 줄이 모두 올라가 서명이 앵커에서 벗어나고, 어울림은 한컴에서 앵커 줄 글자를 그림 밖으로 밀어낸다.
5. 산출물은 `hwpx_gate`(rhwp 허용 목록 포함)를 통과하고, **앵커를 다시 재서 넣기 전과 같은 자리(0.5px 이내)일 때만** 결과물이 된다. 실패하면 산출물을 `_work-hwpx-automation/`에 남기고 `--backend com`을 안내한다. 조용히 COM으로 넘어가지 않는다.

**게이트 보고 읽기**: 성공 출력에 "허용된 손실(rhwp 재직렬화): p id 재부여×N, switch×3 …" 줄이 붙는다. rhwp가 문서를 통째로 다시 쓰며 생기는 알려진 차이(본문이 쓰지 않는 paraPr의 개요 8~10단계 블록, 문단 id 재번호, 그림자 없음 표시)이며 화면에는 영향이 없다. 표·글상자가 많은 학교 양식처럼 그 밖의 손실(표 셀 속성·대각선 테두리·줄 배치 캐시 등)이 나오면 게이트가 막는다. 그런 문서는 `--backend com`으로 넣는다.

**검증은 PDF를 눈으로**: `--pdf`는 `hwpx_edit.py --to-pdf`(rhwp)로 검증 PDF를 만들고 그림 수를 센다(PyMuPDF. 없거나 PDF 변환이 실패하면 `rhwp export-svg`로 그 쪽의 그림이 원본보다 늘었는지). 그림 수만으로 위치는 모른다. **서명 구역을 확대해 렌더 이미지를 육안 확인**한다(PDF 텍스트 추출은 그림이 글자를 덮어도 글자를 그대로 읽어 겹침을 못 잡는다). rhwp 경로의 PDF는 rhwp가 그린 것이라 한컴 화면과 조판이 다를 수 있다. 제출 전 한컴에서 한 번 열어 보는 것이 가장 확실하다.

## 실전 팁·함정

- **서명이 기존 텍스트와 겹칠 때**: 서명란 텍스트가 줄 우측에 몰린 양식(예: 선행 공백이 많아 "성 명 : 홍길동"이 우측 끝에 붙는 동의서)에서는 서명이 이름 위를 덮는다. **선행 공백은 명백한 사전 신호다**: 서명 삽입 전에 앵커 문단의 선행 공백 수를 세어, 라벨이 이미 우측(예: 공백 40칸+)에 몰려 있으면 삽입 전에 `hwpx_edit.py --find/--replace`로 공백을 줄여 이름을 왼쪽으로 당긴다("일단 넣고 PDF 보고 판단"이 아니다). 그 뒤에도 겹치면 `--horz-offset`으로 서명을 이름 밖으로 민다. → 사례
- **이미지 배경**: 투명/흰 배경 PNG를 쓴다. 회색·불투명 배경 이미지는 서명란에 박스가 비치므로 배경 제거 후 사용.
- **앵커 표시를 지우고 넣어야 할 때**: rhwp 경로는 앵커 글자를 늘 남긴다. 표시를 지운 자리에 넣는 양식은 `--backend com`(기본이 앵커 대체)을 쓴다.
- **줄바꿈에 걸친 앵커·쪽을 넘는 문단의 앵커**는 자리를 하나로 정할 수 없어 멈춘다. 앵커를 더 짧게 주거나 `--backend com`을 쓴다.
- **서명 이미지 경로**: 이 저장소에는 서명 이미지를 포함하지 않는다. 사용할 서명/도장 이미지 파일 경로는 호출 시 `--image`로 직접 지정한다.

## COM 경로 (`--backend com`, Windows + 한컴오피스)

표 안 앵커, rhwp 게이트 실패, 한컴에서만 맞는 조판이 필요할 때 쓴다. 원격 Windows에 SSH로 붙어 돌릴 때는 `reference/com.md` "원격 실행"을 따른다.

```bash
python hwpx_sign.py 동의서.hwpx --image 서명.png --backend com --anchor "(서명)" --pdf
python hwpx_sign.py 동의서.hwpx --image 서명.png --backend com --keep-anchor --horz-offset 27510 --pdf
python hwpx_sign.py 동의서.hwpx --image 서명.png --backend com --inline
```

COM 경로의 `--pdf`는 `--pdf-backend com`으로 한컴이 직접 PDF를 만든다(한컴이 문서를 실제로 여는지까지 확인된다). 기본값도 rhwp와 다르다: 앵커 텍스트를 **지우고** 그 자리에 넣고(`--keep-anchor`로 겹쳐 찍기), 가로는 본문 우측 끝에서 `--gap-mm`(기본 7) 안쪽으로 우측 정렬한다. 표시 위에 정확히 얹으려면 출력된 계산값에서 시작해 `--horz-offset`으로 맞춘다(글자 단위 x는 lineseg에 없어 자동 계산할 수 없다).

**단순 COM `InsertPicture`로는 안 된다** (도구가 자동 회피하는 함정):
1. COM `InsertPicture`는 그림을 floating(`textWrap="TOP_AND_BOTTOM"`) + `curSz=0,0`으로 넣는다. **서명란이 페이지 하단에 있으면 큰 그림이 인라인으로 그 줄에 못 들어가 다음 페이지로 밀린다.**
2. floating + `vertRelTo="PAPER"` 절대좌표를 줘도 **앵커 문단이 페이지 끝이면 개체가 다음 페이지에 그려진다**(표시 페이지는 앵커 문단의 페이지를 따름). → 앵커를 페이지 내 **위쪽 문단**으로 옮겨야 한다.
3. 텍스트가 바뀐 문단의 `<hp:linesegarray>`(줄 레이아웃 캐시)를 안 지우면 **한컴 COM이 문서 열기를 거부**한다. `hwpx-validate`(XSD)·`--to-md`(recall)는 모두 통과 → **자동 검증으로 못 잡고 PDF 변환(COM Open)에서만 드러남.**
4. 반대 방향의 함정: **COM이 그림을 넣은 문단은 저장 시 `<hp:linesegarray>`를 아예 빼고 쓴다**(원리: `reference/warnings-com.md` 10번). 도구는 **편집 전 원본**에서 anchor 문단의 lineseg를 읽고, 그 문단에도 캐시가 없으면 직전 줄에서 `vertpos + vertsize + spacing`으로 한 줄 아래를 추정한다(추정 시 WARN 출력).

도구는 (a) COM으로 anchor 자리에 이미지를 넣어 BinData 확보 → (b) XML 후처리로 floating PAPER 절대좌표 전환 → (c) 앵커를 anchor 문단 바로 앞의 `<hp:p>`로 이동(직전 문단 안에 표·글상자가 있으면 그 안의 마지막 문단이 골라져 그림이 다른 쪽·위치에 그려질 수 있다. `--pdf`로 확인하고 어긋나면 `--horz-offset`/`--vert-adjust`나 `--inline`) → (d) 세로 좌표를 페이지 여백+원본 lineseg vertpos로 계산 → (e) 변경 문단 lineseg 제거를 수행한다.

COM 경로 함정:
- **서명란이 글상자·표 안에 있으면 자동 좌표가 어긋난다**: 그 안의 lineseg `vertpos`/`horzpos`는 개체 기준 상대좌표다. 도구가 WARN을 내며 `--horz-offset`/`--vert-adjust`로 보정한다(개체 기준 오프셋은 문서마다 상수). 공문 양식은 본문 전체가 글상자 한 개인 경우가 흔하니 WARN이 뜨면 반드시 PDF를 육안 확인.
- **서명이 서명줄이 아니라 윗줄에 찍히는 경우**: 서명란이 `<hp:lineBreak/>`로 소속·직급·성명을 한 문단에 쌓은 양식이면 자동 세로 좌표가 첫 줄에 맞는다. 줄 간격만큼 `--vert-adjust`로 내린다(예: 3줄 서명란에서 약 +3600).
- **`hp:pic`이 `hp:run` 밖에 놓이면 그림이 조용히 사라진다**: 한글은 오류 없이 그림을 그리지 않고 `hwpx-validate`·`hwpx-validate-package`는 통과한다. `--pdf`가 PDF 이미지 객체 수 0이면 `[ERROR]`·종료 코드 1(PyMuPDF 없으면 `pdfimages -list <PDF>`로 확인). 도구는 빈 문단의 self-closing run을 열린 태그로 바꿔 그 안에 넣고, run이 없으면 만들어 감싼다. → 사례
- **COM 프로세스 잔류**: 연속 실행 중 PDF 변환이 "문서 열기 실패"로 죽거나 출력 파일 쓰기가 `PermissionError`로 막히면(직전 Hwp.exe가 파일을 잠금) `reference/warnings-com.md` 9번대로 정리한 뒤 재시도.
