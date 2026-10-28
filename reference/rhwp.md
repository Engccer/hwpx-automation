# rhwp 엔진

[rhwp](https://github.com/edwardkim/rhwp)는 HWP·HWPX를 한컴 없이 읽고 고치고 그리는 Rust 단일 바이너리(MIT)다. 이 스킬에서 HWP→HWPX 변환, PDF, 서명 삽입, 누름틀, 격자 좌표 셀, 배포용 HWP 내보내기의 기본 엔진이고 모든 OS에서 돈다. HWPX 미세 편집(`--find/--replace`·`--set-cell`·행 삭제)은 원본 바이트를 지키는 lxml이 1순위이며 rhwp를 쓰지 않는다.

## 설치·버전

- Releases에서 플랫폼 바이너리(예: `rhwp-v0.8.7-macos-aarch64.tar.gz`, Windows는 x86_64 zip)와 `SHA256SUMS.txt`를 받아 체크섬을 확인한 뒤 PATH(예: `~/.local/bin`, Windows `%USERPROFILE%\.local\bin`)에 둔다. PATH 밖에 두면 환경변수 `HWPX_RHWP_BIN`에 바이너리 경로를 준다(PATH보다 우선).
- `python hwpx_edit.py --check-env`의 Tier 2에 `[O] rhwp`와 버전이 뜨면 준비된 것이다.
- **하한 v0.8.7**: 그 미만이면 모든 호출이 `RhwpVersionError`로 거부된다(`export-hwpx` 위치 인자·`edit insert-image`·`--json` 봉투 계약을 이 버전에서 실측했다).
- **실측 상한 `TESTED_MAX`**(`rhwp_cli.py`, 지금 v0.8.7)보다 새 버전은 막지 않고 경고만 낸다. 봉투 계약 검사가 어긋남을 잡지만, 새 버전으로 올린 뒤 첫 결과는 한 번 더 확인한다. 실측을 마친 버전으로 `TESTED_MAX`를 올린다.
- PyPI 바인딩이 없어 서브프로세스가 유일한 통합 방법이다. `cargo install`은 쓰지 않는다(릴리스 바이너리만).

## 호출 계약 (`rhwp_cli.run`)

rhwp는 `rhwp_cli.run()`으로만 부른다. 명령 문자열을 직접 조립하지 않는다.

- **봉투**: `--json`을 붙여 stdout의 JSON 봉투를 읽는다. `schemaVersion`이 `"1.0"`이 아니거나, 명령별 필수 필드(`required_fields`)가 없거나, exit 0인데 산출 파일이 없거나 비어 있으면 `RhwpContractError`로 실패하고 산출물을 게시하지 않는다.
- **산출 파일**: 인자 목록에서 산출 자리를 `rhwp_cli.OUTPUT` 자리표로 적고 `output=`을 주면, 대상 폴더 안 임시 폴더에 먼저 쓰고 성공·계약 검사·`verify=` 콜백을 모두 통과한 뒤에만 `os.replace`로 옮긴다. 실패하면 기존 파일을 건드리지 않는다. 콜백은 `verify(staged, envelope)`이며(인자 하나짜리 콜백도 받는다) 봉투 사후조건(채운 수가 사전 점검과 같은지 등)도 게시 전에 여기서 본다. exit 3·4, 봉투 계약 오류, verify의 어떤 예외든 산출물을 `<출력 stem>.rejected-<YYYYmmddHHMMSS><확장자>`로 남기고(같은 초에 겹치면 `-2`) 예외의 `rejected_path`와 메시지 끝에 적는다. RhwpError가 아닌 예외는 `RhwpVerifyError`로 감싼다.
- **암호**: `password=`는 `--password-stdin`으로 넘긴다(프로세스 목록에 노출되지 않게).
- **문서 유래 값**: 봉투의 `untrustedFields`에 든 값은 사람에게 보여 주되 분기 조건으로 쓰지 않는다(`rhwp_cli.untrusted()`).
- **`--json`이 없는 명령**: `export-render-tree`·`measure-width`는 `--json`을 주면 exit 2다. `expect_json=False`로 부른다.

| rhwp exit | 예외 | 뜻·처리 |
|---|---|---|
| 0 | (없음) | 성공. 봉투·산출물 계약 검사를 거친다 |
| 1 | `RhwpRuntimeError` | 런타임 실패(알 수 없는 코드도 여기) |
| 2 | `RhwpUsageError` | 사용법 오류. 우리 쪽 인자 조립 결함으로 본다 |
| 3 | `RhwpVerifyError` | `--verify` 저장본 불일치. 산출물은 `<출력 stem>.rejected-<시각><확장자>`로 남는다 |
| 4 | `RhwpVerifyError` | `--verify-pages` 쪽수 불일치. 산출물 보존은 3과 같다 |

rhwp 원시 exit code는 `RhwpResult.exit_code`에 그대로 남고 래퍼 오류와 섞이지 않는다.

명령별 필수 봉투 필드(v0.8.7 실측): `export-hwpx`·`convert`는 `output, format, verify`, `export-pdf`는 `output, format, pageCount`, `edit set-cell`은 `oldText, newText, output`, `edit fill-fields`는 `filledCount, notFound, ambiguous`(쓰기는 `output` 추가), `export-tables`는 `tables`, `fields`는 `fields`, `edit insert-image`는 `output, page, overflow`, `search`는 `matches, totalMatchCount`, `export-svg`는 `pages`.

## 독립 손실 게이트 (`hwpx_gate.py`)

rhwp는 문서를 자기 내부 구조로 읽어 **통째로 다시 써서** 저장한다. 내부 구조가 담지 않는 것은 조용히 사라지고, rhwp 자신의 `--verify`·`hwpx-roundtrip`은 그 손실을 보지 못한다. 그래서 rhwp가 쓴 HWPX는 `hwpx_gate.gate()`를 통과한 뒤에만 결과물이 된다. → 사례

검사 항목:
- **요소 census**: `Contents/*.xml`에서 (태그, 정렬된 속성 묶음) 단위로 개수를 세어 사라지거나 바뀐 요소를 찾는다. `header.xml`은 키에 id(글꼴 묶음은 lang)를 가진 조상 경로와 요소 글자를 넣어, 글자 모양·테두리 사이에서 자식이 옮겨 가거나 번호 서식(`paraHead`) 글자가 바뀐 것도 잡는다. 목록 머리의 `itemCnt`는 자식에서 나오는 값이라 보지 않는다. 편집 대조에서는 원본에 있던 항목(같은 소속 경로) 아래에 새 요소가 붙으면 실패다. rhwp가 글자 모양에 채우는 "없음" 기본값(그림자·밑줄·외곽선 `type="NONE"`, 취소선 `shape="NONE"`)만 허용하고 센다. 새 id 항목(칸 글자용 새 charPr 등)의 추가는 보고만 한다.
- **편집 대조 (`edits=`)**: rhwp 편집 산출물은 `section*.xml`을 census 대신 문서 순서의 토큰열(시작 태그와 속성, 글자, 끝 태그)로 펴서 원본과 위치별로 대조한다. 대상은 산출물이 아니라 원본 XML과 요청값으로 정한다. `CellEdit(표, 행, 열, 새값, old_text)`는 문서 순서의 최상위 표에서 `cellAddr`가 맞는 칸의 `subList` 안쪽만, `FieldEdit(이름, 값, k)`는 그 CLICK_HERE 누름틀의 fieldBegin을 품은 ctrl 끝부터 fieldEnd를 품은 ctrl 시작까지만, `ImageInsert(구역, 문단)`은 그 최상위 문단에 더해진 그림 하나만 가린다. 가린 자리 밖은 위치·부모·순서·속성·글자가 모두 같아야 하고, 가린 자리 안은 글자 밖 요소가 같고 글자가 요청값과 같아야 한다(칸 안의 중첩 표는 글자까지 같아야 한다). 비교 전에 rhwp가 한컴 문서를 다시 쓸 때 바꾸는 포장(실측)을 정규화한다: run·`<hp:t>` 경계는 보지 않고 글자 모양이 바뀌는 자리와 글자만 보며, 빈 run·빈 t, `\r\n`과 `\n`, 그림(`pic`) 안 자식 순서, 필드 `zorder`·빈 `metaTag`는 같게 본다. 바탕쪽 같은 그 밖의 본문 파트도 같은 토큰열로 완전히 같아야 한다.
- **패키지 파트 (`edits=`)**: `Contents/*.xml`·`BinData/`·`content.hpf` 밖 파트(`Chart/`·`DocHistory/`·`Scripts/`·`settings.xml` 등)는 이름과 바이트가 같아야 한다. `Preview/`와 `META-INF/container.xml`만 rhwp가 다시 만들어도 된다. `content.hpf`는 메타데이터(제목·작성자·날짜), 매니페스트 항목(`hashkey` 밖 속성 전부), 구역 순서(spine)를 대조한다. 항목 순서·`hashkey`·`linear`는 rhwp가 바꾸므로 보지 않는다. `allow=` 콜백과 `expect_text_change`는 rhwp 편집에 쓰지 않는다.
- **BinData**: 파일 집합과 SHA-256. 그림·OLE·차트 변형을 잡는다. 그림 삽입은 `allow_bindata_added=1`.
- **글자**: hwpx-tomd로 두 문서의 글자 멀티셋을 비교한다. `edits=`를 주면 위치별 대조가 글자까지 보므로 이 대조는 하지 않는다.
- **스키마**: `hwpx-validate`.

상태는 `passed`·`failed`·`unverified` 셋이다. hwpx-tomd나 `hwpx-validate`가 없어 못 본 항목이 있으면 `passed`가 아니라 `unverified`이며, "검증됨"이라고 보고하지 않는다.

**허용 목록 (`allow_rhwp=True`)**: rhwp 재직렬화로 알려진 손실 다섯 가지만 실패 대신 `GateReport.allowed`로 보낸다.
1. `header.xml`에서 **본문이 한 번도 참조하지 않는 paraPr**(header 밖 모든 XML의 `paraPrIDRef` 속성을 XML로 파싱해 판정. 파싱하지 못한 파트가 있으면 이 허용을 적용하지 않는다) 안의 `switch`·`case`·`default`와 `heading`(type NONE, 또는 OUTLINE level 7~9) 감소. 실측된 모양은 개요 8~10단계(`level` 7~9)의 `hp:switch/case/default` 블록이 `heading type=NONE`으로 접히는 것이다. 스타일만 참조하는 paraPr도 미사용으로 보는 것은 개요 8~10단계 paraPr가 스타일에서만 참조되기 때문이다. 미사용 paraPr라도 그 밖의 변화(정렬 등)와 본문이 쓰는 paraPr의 같은 손실은 실패다.
2. 본문 파트의 `hp:p id` 재부여. 문단 `id`를 빼고 센 census가 완전히 같을 때(문단 수·나머지 속성이 같을 때)만이다.
3. 본문 파트의 `hp:shadow type="NONE"` 소실. 그림자 변화가 전부 이것일 때만이며, NONE→DROP처럼 그림자가 새로 생기는 변경은 실패다.
4. `header.xml` paraPr의 `textDir` LTR·AUTO 소실(한컴 저장 문서 12건 중 6건, 한글은 어느 쪽이든 왼쪽부터 쓴다). RTL 소실은 실패다.
5. 편집 대조에서 모든 `fieldBegin`의 `zorder` 소실(채우지 않은 누름틀 포함). 채운 누름틀의 `dirty` 변화는 대상 자리의 정상 변화로 본다.

그 밖의 손실(본문 섹션 요소, 표 셀 속성, 대각선 테두리, 줄 배치 캐시 등)은 모두 실패다. 허용 구역 안에서 늘어난 요소는 `gained`로 보고만 한다.

**보고 줄**: `allowed`가 비어 있지 않으면 `summary()` 둘째 줄과 `describe_lost()` 끝에 `허용된 손실(rhwp 재직렬화): p id 재부여×N, switch×3, …`이 붙는다(따로 쓰려면 `allowed_line()`). 비교에서 정규화로 같게 본 차이도 종류별 건수로 이 줄에 나온다: `run·t 경계 재분할(문단)`, `줄바꿈 \r\n→\n`, `그림 안 자식 순서`, `필드 빈 metaTag`, `fieldBegin zorder`, `paraPr textDir LTR·AUTO 소실`, `content.hpf 항목 순서`·`hashkey`·`linear`, `itemCnt 갱신`, `<요소> 없음 기본값 채움`. 대상 칸 안의 문단·글자 모양 변화와 문단 쪼개기는 실패가 아니지만 `칸 안 문단·글자 모양 변화`로 센다. 화면에는 영향이 없는 차이다. 이 줄이 찍히면 완료 보고에 그대로 옮겨 적는다.

**교차 대조 (`cross_check`)**: 원본 XML이 없는 HWP 변환은 같은 HWP를 다른 엔진(JAR)으로 바꾼 HWPX와 대조한다. 글자 멀티셋, 표(`tbl`)·셀(`tc`)·그림(`pic`)·차트(`chart`)·OLE(`ole`)·수식(`equation`) 수, BinData와 `Chart/` 파트의 수와 내용 해시 멀티셋을 본다. HWP 입구 변환과 `--to-hwp` 왕복이 이 함수 하나를 쓴다. JDK가 없어 대조 상대를 못 만들면 `unverified`다.

## 좌표·단위

- HWPUNIT = 1/7200인치. 1mm = 283.46 HWPUNIT(`hwpx_sign.MM`).
- `export-render-tree`의 px는 96dpi이고 px × 75 = HWPUNIT(`hwpx_sign.PX`).
- 쪽 번호는 `search`·`export-render-tree`·`insert-image` 모두 0부터다. 사람에게 보일 때만 +1한다.
- **표 좌표 체계는 둘이고 섞지 않는다.** `--set-cell T,R,C`(lxml)는 물리 `hp:tc` 인덱스, `--grid-set-cell T,R,C`(rhwp)는 병합을 반영한 격자 좌표다. 같은 셀이 다른 좌표가 되므로 격자 좌표는 `--tables`로 먼저 본다.

## 명령별 계약 (v0.8.7 실측)

- **`search <문서> <글자>`**: 매치마다 구역·문단·글자 위치와 `page`를 준다. `page`는 매치가 있는 쪽이 아니라 **문단이 시작하는 쪽**이다. 쪽을 넘는 문단 안 앵커는 화면 위치를 확정할 수 없다. 표 셀 안 매치에는 `cell` 키가 붙는다.
- **`export-render-tree <문서> -p <쪽> -o <폴더>`**: 파일 이름은 `render_tree_<쪽+1, 세 자리>.json`이다. 셀·글상자 안 TextRun은 `pi`를 그 컨테이너 기준으로 매겨 본문 문단 번호와 겹치므로, 본문 앵커는 `Body>Column>TextLine` 바로 아래 run만 본다. 글자 크기·글꼴 정보는 없다.
- **`measure-width`**: TSV를 낸다. render-tree에 글꼴 정보가 없으므로 run 전체 폭에 대한 비율로 환산해 run 중간 글자의 x를 구한다. `--font`를 받지만 모르는 글꼴도 오류 없이 대체 폭을 돌려주므로(실측) 그 글꼴 치수를 rhwp가 아는지 확인할 수 없다. 그래서 잘리는 글자 묶음이 header.xml의 charPr→fontRef로 함초롬바탕(기본값)이고 언어 사이 장평이 같고 자간 0·보통 굵기일 때만 `measured`로 보고, 판정하지 못하거나 다른 글꼴이면 `estimated`(`--allow-estimated` 없으면 중단)다.
- **`edit insert-image`**: 봉투 필드는 `output, outputFormat, page, x, y, width, height, overflow, binDataId, changedPages`다. 그림을 **그 쪽 첫 문단**에 `textWrap="SQUARE"`·`allowOverlap="0"`으로 붙인다. 첫 문단이 빈 문단이면 rhwp 조판이 그 줄을 접어 아래 줄이 모두 올라가고, 어울림 배치는 한컴에서 앵커 줄 글자를 밀어낸다. `hwpx_sign.float_over_anchor`가 그림을 앵커 문단 맨 앞(구역 정의·컨트롤 뒤)으로 옮기고 `IN_FRONT_OF_TEXT`·`allowOverlap="1"`로 놓으며, 빈 run이 남으면 `<hp:t/>`를 되돌린다. 원본의 같은 run에 `<hp:t/>`가 없었으면 빈 run 그대로 둔다. 넣은 뒤 앵커를 다시 재서 0.5px 안에서 같을 때만 결과물로 둔다. → 사례
- **`edit set-cell`**: 중첩 표와 일부 병합 칸은 "v1 범위 밖"으로 거부한다. 그 메시지를 그대로 전하고 `--set-cell`(lxml)을 안내한다.
- **`edit fill-fields`**: exit 0이어도 없는 이름·모호한 이름·닮은 이름(`confusable`)이 있거나 채운 수가 0일 수 있다. `--dry-run`으로 먼저 점검해 그런 경우 쓰지 않는다. 같은 이름이 여럿이면 `"이름[k]"`(k는 0부터)로 지목한다. 값은 한 줄 글자나 숫자만 받는다(줄바꿈·탭·`true`·`null`은 쓰기 전에 거부). 쓰기 봉투가 사전 점검과 다르면(채운 수·거부 목록) 게시 전에 거부한다.
- **`convert <in.hwpx> <out.hwp> --verify --verify-pages`**: 그림·차트가 빠질 수 있다고 rhwp가 경고한다. `--to-hwp`는 산출 HWP를 JAR로 되돌린 HWPX를 원본과 `cross_check`한다. 차트·OLE·수식 수나 `Chart/` 파트 내용이 하나라도 다르거나, JDK가 있는데 JAR가 못 읽으면 실패다.

## 폴백 조건

폴백은 자동 전환이 아니다. rhwp가 실패하면 산출물을 `.rejected`나 `_work-hwpx-automation/`에 남기고 stderr에 폴백 명령을 적는다. 폴백으로 전환했으면 완료 보고에 엔진과 사유를 쓴다.

| 축 | 기본 | 폴백 | 전환 조건 |
|---|---|---|---|
| HWP→HWPX | rhwp `export-hwpx` + JAR 교차 대조 | `--engine jar`(JAR + PrvText 보정), rhwp가 실패했는데 JDK가 없거나 JAR도 실패한 Windows는 `hwpx_com.py --from-hwp` | rhwp exit 1·3·4, `hwpx-validate` 실패, 교차 대조 불일치 |
| PDF | rhwp `export-pdf` | Windows 한컴 `--pdf-backend com` | 한컴과 같은 조판이 필요할 때, 빈 네모 exit 2, rhwp가 암호 문서를 못 열 때 |
| 서명·도장 | rhwp `insert-image`(`hwpx_sign.py`) | `hwpx_sign.py --backend com` | 표 안 앵커, 앵커 위치 미확정, 게이트 실패, 앵커 표시를 지우고 넣는 양식 |
| HWPX 미세 편집 | lxml(`--find/--replace`·`--set-cell`·행 삭제) | rhwp `--grid-set-cell` | 사용자가 격자 좌표를 원할 때만 |
| 누름틀·격자 셀 | rhwp(`--fill-fields`·`--grid-set-cell`) | `--set-cell`(lxml) | 중첩 표, 게이트 실패, 원본 보존이 우선인 문서 |
| HWPX→HWP | rhwp `convert`(`--to-hwp`) | Windows 한컴 다른 이름 저장 | 왕복 대조 실패(그림·표·글자·차트 감소) |
| 암호 해제 | 한컴 COM | 없음 | rhwp `--password-stdin`은 PDF 변환에만 쓴다 |
| 한컴 재저장 정규화 | 한컴 COM(`hwpx_com.py --normalize`) | 없음 | |

## PDF (`--to-pdf`)

`hwpx_edit.py --to-pdf`는 모든 OS에서 rhwp `export-pdf`를 부른다(`rhwp_pdf.py`). 한컴 없이 한글 조판 규칙으로 쪽·줄을 나눈다. 한컴과 똑같은 조판이 필요하면 Windows + 한컴오피스에서 `--pdf-backend com`을 명시한다(그 밖의 OS에서는 exit 1로 거부).

```bash
python hwpx_edit.py <파일.hwpx> --to-pdf -o <파일.pdf>                     # rhwp, engine=rhwp
python hwpx_edit.py <파일.hwpx> --to-pdf --pdf-backend com -o <파일.pdf>   # Windows 한컴, engine=com reason=사용자 지정
python hwpx_edit.py <파일.hwpx> --to-pdf --password <암호>                  # 암호 문서(--password-stdin으로 전달)
```

`.hwp`는 변환하지 않고 원본을 그대로 렌더링한다. Pandoc(`hwpx_convert.py`) 산출물도 rhwp로 PDF를 낸다. COM은 Pandoc 산출물의 본문을 0자로 읽으므로 `--pdf-backend com`에 넣지 않는다.

**충실도**: 대조한 문서에서는 한컴 PDF와 쪽수·글자가 같았고, 남은 차이는 세 가지다. → 사례
- 글꼴이 이 컴퓨터 글꼴로 대체된다(맥: 맑은 고딕·HY헤드라인M → Apple SD Gothic Neo, 명조 → 나눔명조).
- 한 문단에서 줄 끝 단어 하나가 다음 줄로 넘어가는 수준의 줄바꿈 차이가 가끔 생긴다.
- 제목 옆 짧은 라벨의 세로 위치가 조금 다르다.

**출력 해석**:
- 표준출력 `사용 글꼴:` 목록으로 대체 여부를 확인한다. 그다음 줄이 `engine=rhwp`다.
- stderr `경고: 쪽 하단 넘침 N건`: 조판이 쪽 끝을 몇 px 넘었다는 뜻이다. 쪽수는 대개 그대로지만 쪽 끝 줄을 렌더로 확인한다.
- **exit 2 + stderr `글꼴에 없는 문자가 빈 네모로 찍혔습니다: ✅(U+2705)`**: PDF는 남지만 배포하면 안 된다. 이모지·특수기호가 설치 글꼴에 없어 LastResort(빈 네모)로 찍힌 것이다. `--font-path`·`RHWP_FONT_PATH`·Noto Emoji 사용자 설치로는 이 대체가 바뀌지 않았다(→ 사례). 해당 문자를 바꾸거나 Windows 한컴(`--pdf-backend com`)으로 변환한다.
- rhwp 실패(exit 1)면 stderr의 폴백 줄을 따른다. 암호 문서를 rhwp가 못 열면 Windows `--pdf-backend com`이다.

## 함정

- **rhwp v0.8.7이 진짜로 바꾸는 문서는 거부된다.** 한컴 저장 문서 12건 실측에서 그림의 `thumbnailBinIDRef` 소실과 필드 끝(`fieldEnd`) 위치 이동(필드 범위가 뒤 글자까지 넓어짐)이 1건씩 나왔다. 새 판에서 다시 잰다(BACKLOG 감시 항목). 게이트 실패 메시지가 가리키는 자리를 확인하고 `--set-cell`(lxml)이나 `--backend com`을 쓴다.
- **표본 HWP를 만들 때**: python-hwpx로 갓 만든 HWPX를 rhwp `convert`로 바꾼 HWP는 lineseg가 없어 `export-hwpx --verify`가 exit 3이다. 한 번 더 왕복한 HWP는 차이가 0이다. 시험 표본은 왕복을 한 번 더 거친다.
- **rhwp 자가 검증 통과는 무손실을 뜻하지 않는다.** `--verify`·`hwpx-roundtrip`이 차이 0이어도 header.xml 요소가 줄어 있을 수 있다. 판정은 늘 `hwpx_gate`로 한다.
- **rhwp가 읽기 엔진은 아니다.** `export-markdown`은 쪽 단위 파일이고 쪽·줄 경계에서 단어가 쪼개진다. 읽기는 hwpx-tomd(`--to-md`)다.
- **k-skill `rhwp-edit`·`rhwp-advanced`·`hwp`**는 같은 엔진의 WASM·안내형 래퍼이고 HWP5 출력만 낸다. HWPX 정본·게이트가 필요한 작업은 이 스킬로 한다.
