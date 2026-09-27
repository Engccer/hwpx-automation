# HWP → HWPX 변환: 번들 JAR과 폴백

기본 변환 명령과 셸 규칙은 SKILL.md "HWP → HWPX 변환" 절에 있다. 이 문서는 번들 JAR을 다시 빌드하거나 hwp2hwpx가 예외로 죽을 때 읽는다.

## 번들 JAR의 출처·패치

`convert/hwp2hwpx-1.0.0-c9d8a27p1.jar`는 상류 `neolord0/hwp2hwpx` 커밋 `c9d8a27`(이슈 #5 컨트롤 불일치 수정·표 쪽 나눔 매핑 수정 포함)에 자체 패치 1건(`convert/patches/0001-extendControl-range-guard.patch`)을 더해 빌드한 것이다. 상류는 버전 번호를 1.0.0에 고정한 채 커밋만 쌓으므로 **파일명의 커밋 해시가 곧 버전**이다. 의존은 `lib/hwplib-1.1.10.jar`·`lib/hwpxlib-1.0.9.jar`(Maven Central). 재빌드 절차는 `convert/patches/README.md`.

- 패치 내용: 머리글·바닥글(마스터 페이지) 문단에서 **컨트롤 문자 수가 컨트롤 데이터 수보다 많으면** `ForChars.extendControl`이 `IndexOutOfBoundsException(Index 1 out of bounds for length 1)`으로 죽는다. 상류 #5 수정은 null 체크뿐이라 이 경우를 못 잡는다. 범위 밖 인덱스는 건너뛴다(본문은 무손실). → 사례
- 상류가 갱신되면 패치가 여전히 필요한지 먼저 확인하고, 필요하면 같은 절차로 재빌드해 파일명 해시를 바꾼다.

## 폴백 1: hwp2hwpx가 죽는 HWP (한컴 COM 경유, 서식 보존)

hwp2hwpx가 예외로 죽는 문서는 Windows 한컴오피스 COM으로 HWPX를 만들면 정식 경로(`--to-md`)를 그대로 탈 수 있다: `python hwpx_com.py <파일.hwp> --from-hwp -o <파일.hwpx>`. 원격(SSH)에서 돌릴 때는 `reference/com.md` "원격 실행"을 따른다.

예외 유형별 원인:

| 예외 | 위치 | 원인 |
|---|---|---|
| `EmptyStackException` | `ForInlineControl.fieldEnd` | 표 셀 안 필드 컨트롤의 시작/끝 짝 불일치 |
| `IndexOutOfBoundsException` | `ForChars.extendControl` | 컨트롤 문자 수 > 컨트롤 데이터 수, 주로 머리글·바닥글 문단. 현 번들 JAR은 패치돼 통과 |

## 폴백 2: pyhwp 경유 (텍스트만)

한컴 COM을 쓸 수 없을 때는 pyhwp로 우회한다:

```bash
pip install pyhwp   # hwp5txt / hwp5proc 제공

# hwp5txt는 표 내용을 <표> 플레이스홀더로 뭉개므로, 본문이 표 안에 있는
# 공문·안내문은 반드시 hwp5proc xml 경유로 셀 텍스트까지 걷는다:
hwp5proc xml <파일.hwp> > full.xml
python convert/hwp_xml_to_md.py full.xml output.md
```

- `convert/hwp_xml_to_md.py`: pyhwp XML에서 표 구조(행='|', 셀 내 문단='/')를
  보존해 텍스트를 추출하는 폴백 파서. 각 텍스트를 정확히 한 번만 출력하도록
  `TableControl` 경계에서 가지치기한다(문단·중첩 표 중복 방지).
- 한계: 서식·이미지 미보존(검색·내용 파악용). 서식 보존이 필요하면 폴백 1(한컴 COM)로.
- 단, **취소선·글자색 같은 문자 서식은 어느 마크다운 경로에서도 나오지 않는다**(`--to-md`·kordoc 모두). 필요하면 `hwp5proc xml`의 `CharShape` 속성(`underline="line_through"`, `text-color`)을 문단 텍스트에 대응시켜 복원한다. → 사례
