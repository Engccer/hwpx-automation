# 주의사항: XML 편집 전반

1. **`--find/--replace` 2단계 치환**: (1) python-hwpx로 본문 런 치환 (서식 보존, 런 분할 텍스트 지원) → (2) lxml로 표 셀 내부만 추가 치환. 영역이 분리되어 중복 치환 없음. **주의**: `--find/--replace`는 내부적으로 `HwpxDocument.open()`을 사용하여 구조 검증이 엄격함 (PrvText.txt 필수). `--set-cell`은 lxml 직접 파싱이라 검증 없이 통과. HWP→HWPX 변환 파일은 PrvText 누락으로 `--find/--replace`가 거부되며, 그때 출력하는 두 단계(`--add-preview` 후 보정본으로 재실행)를 따른다(11번)
2. **편집 후 검증**: `hwpx-validate`로 무결성 확인, 양식 편집 시 `hwpx-page-guard`로 쪽수 드리프트 감지
3. **자동 출력 폴더와 `--set-cell` 다중 호출**: `-o` 미지정 시 입력 파일과 같은 디렉터리의 `_work-hwpx-automation/` 폴더에 저장 (원본 비파괴). **주의**: 편집 명령은 매번 입력 파일을 새로 읽는다. 원본을 입력으로 `--set-cell`을 여러 번 부르면(`&&` 체이닝 포함) 같은 출력 자리를 덮어써 **마지막 호출만 남는다**. 두 번째 호출부터는 앞 호출의 결과를 입력으로 준다(입력이 `_work-hwpx-automation/` 안에 있으면 `-o` 없이도 그 파일 자리에 쌓인다). 셀이 많으면 스크립트 한 번으로 처리한다
4. **서식 보존 원칙**: 가능한 한 `<hp:t>` 요소의 텍스트만 변경하고, XML 구조는 건드리지 않음
5. **병합 셀 주의**: `rowSpan`/`colSpan`을 변경하면 표 레이아웃이 깨질 수 있음. 한글에서 확인 필요
    - **span의 위치**: 좌표는 `<hp:cellAddr colAddr rowAddr>`, 병합 범위는 별도 `<hp:cellSpan colSpan rowSpan>`이다. cellAddr에서 span을 읽고 쓰면 항상 1x1로 읽히고 스키마 외 속성만 남는다. → 사례
    - **병합 해제는 span=1만으로 완성되지 않는다**: `rowSpan=N` 셀이 덮는 N-1개 행에는 해당 열의 `<hp:tc>`가 아예 없다. 해제하려면 행마다 새 tc를 생성해 `cellAddr` 좌표·`cellSz`를 채워야 한다(`colSpan`도 동일). 현재 `--split-cell`이 이를 자동 수행(앵커 deepcopy로 서식 보존 + 본문 비움 + `linesegarray` 제거)
    - **검증 사각지대**: span이 오염된 결과물도 `hwpx-validate`(XSD)와 `--to-md` 자가검증 recall을 모두 통과한다(lineseg 손상과 같은 계열). 병합 편집 후에는 `--info`의 행별 셀 수·`[병합:NxM]` 표시로 확인하라
6. **다중 섹션 문서**: `section1.xml` 등 존재 가능
7. **네임스페이스 필수**: `hp`, `hs`, `hh`, `hc` (`reference/format.md` 참조)
8. **HWP→HWPX 변환 후 검은 배경**: `--sanitize` 또는 `save_hwpx()` 자동 수정
9. **HWPML 2016 → 2011**: python-hwpx v2.8+는 네임스페이스 자동 변환 지원
10. **`--to-md`는 표 밖 본문도 추출한다**: 문서 제목, 지도교사, 머리글 등 표 바깥 텍스트도 나온다(엔진 hwpx-tomd, SKILL.md "`--to-md` 보장과 한계"). 추출 범위 밖은 이미지 속 글자다
11. **HWP→HWPX 변환 후 `PrvText.txt` 누락**: hwp2hwpx 변환 결과에는 `Preview/PrvText.txt`가 없다(container.xml은 선언). python-hwpx API(`HwpxDocument.open()`)·`hwpx-pack`·`hwpx-validate`가 이 누락으로 실패한다. 한글·`--to-md`·`--set-cell`은 정상이다. 정본·배포·검증 대상이거나 python-hwpx로 열어야 하면 `python hwpx_edit.py <파일.hwpx> --add-preview`로 보정한다(가장 간단하다. 손으로 할 때는 ZIP 안에 `Preview/PrvText.txt`가 들어가야 한다)
12. **manifest.xml self-closing 태그 주의**: HWP→HWPX 변환 결과의 manifest.xml이 `<odf:manifest .../>` (self-closing) 형태일 수 있음. `replace('</odf:manifest>', ...)` 패턴이 실패하므로, self-closing 여부를 먼저 확인하고 처리
13. **양식의 "한 줄로 입력" 문단에 긴 텍스트 채우면 글자가 겹쳐 뭉개짐**: 사람이 디자인한 양식의 제목 문단에는 문단 모양 "한 줄로 입력"(`paraPr`의 `<hh:breakSetting ... lineWrap="SQUEEZE"/>`)이 걸려 있는 경우가 있다. 짧은 제목 전제의 디자인이라, 자동화로 **긴 제목·문장을 채우면 한글이 줄바꿈 대신 장평을 무제한 압축**해 글자가 서로 겹쳐 판독 불가가 된다(음수 자간 charPr과 결합하면 더 심함). `hwpx-validate`(XSD)와 `--to-md` recall로는 잡히지 않고 **렌더링(PDF·인쇄)에서만 드러난다**. → 사례
    - **감지(`--list-squeeze`)**: SQUEEZE 문단 중 추정 자연 폭(charPr height·장평·자간 반영 근사)이 컨테이너 가용 폭(셀·글상자·본문)의 1.1배를 넘는 문단만 나열.
    - **보정(`--fix-squeeze`)**: 원본 paraPr은 남기고 `lineWrap="BREAK"` 복제 paraPr을 새 id로 추가해 과압축 문단만 재지정(같은 paraPr을 공유하는 짧은 라벨의 의도된 디자인은 보존). 재지정 문단의 `linesegarray`는 제거해 한글이 재계산하게 함.
    - **자동 경고**: `--find/--replace`·`--set-cell`이 텍스트를 채운 직후 그 텍스트가 SQUEEZE 문단에서 과압축되면 stderr로 경고하고 `--fix-squeeze`를 안내한다.
    - **워크플로우 규칙**: 양식 채우기에서 제목·긴 문장을 채운 뒤에는 `--list-squeeze`로 확인하거나 PDF로 육안 검증하라. 자동 생성 문서의 제목은 폰트 축소·자간 압축으로 한 줄에 욱여넣지 말고 자연 줄바꿈(2줄)을 허용하는 것이 기본이다.
14. **양식의 기입 예시(placeholder) 글자모양이 그대로 상속됨**: 관공서 배포 양식은 기입 예시를 **회색(`textColor="#808080"`)·이탤릭** charPr로 넣어 둔다. 이 셀을 `--set-cell`이나 python-hwpx `fill_by_path`로 채우면 텍스트만 바뀌고 `charPrIDRef`는 그대로라 **제출본이 회색 이탤릭으로 인쇄된다**. 13번(SQUEEZE)과 같은 계열: `hwpx-validate`·`--to-md` recall로는 안 잡히고 렌더링에서만 드러난다. → 사례
    - **감지**: 채울 셀의 `<hp:run charPrIDRef>`가 가리키는 `header.xml`의 charPr에 회색 `textColor`나 `<hh:italic/>`이 있는지 확인. 예시 행·기입 칸·라벨이 각각 다른 charPr id를 쓰므로 셀마다 봐야 한다.
    - **보정 패턴** (양식의 글꼴·크기·자간은 보존하고 서식만 정상화, `--fix-squeeze`의 복제 paraPr 패턴과 동형):
      ```python
      new = copy.deepcopy(charpr_by_id[src_id])   # 예시용 charPr 복제
      new.set('id', str(next_id))
      new.set('textColor', '#000000')
      for it in new.findall(QH('italic')):        # 이탤릭 제거
          new.remove(it)
      charProperties.append(new)
      charProperties.set('itemCnt', ...)          # itemCnt 갱신 필수
      # 채운 셀의 <hp:run charPrIDRef>만 새 id로 교체
      ```
    - **워크플로우 규칙**: 양식 채우기 후 제출 전 PDF 렌더링으로 글자색·기울임을 육안 확인하라. Pandoc 변환물의 파란 제목 함정(`reference/conversion.md`)과 같은 "텍스트 검증으로는 안 잡히는" 부류다.
15. **HWPX를 쓸 때 패키지 규격 — mimetype 압축·`standalone="yes"`**: 모든 ZIP 항목을 `ZIP_DEFLATED`로 쓰거나 `etree.tostring`에 `standalone=True`를 주지 않으면 **편집 전에는 통과하던 문서가 편집만으로** `hwpx-validate-package`에서 ERROR 2건(`mimetype: must use ZIP_STORED`, `missing XML declaration with standalone="yes"`)을 낸다. 한글은 정상적으로 열고 `hwpx-validate`(XSD)·`--to-md` recall도 통과하므로 **자동 검증으로는 드러나지 않는다** — 13·14번과 같은 계열이다. 공문서 제출본처럼 규격 검증을 통과해야 하는 산출물에서 문제가 된다. → 사례
    - **현재 구현**: `write_hwpx_zip()`이 mimetype을 **첫 항목 + ZIP_STORED**로 쓰고, 나머지는 `zip_layout()`이 읽어 둔 **원본의 항목별 압축 방식을 복제**한다(한컴이 STORED로 두는 `version.xml`·`Preview/PrvImage.png`·`BinData/*` 보존). XML 직렬화는 `serialize_xml()` 한 곳으로 모아 `standalone="yes"`를 보장한다(section뿐 아니라 **header.xml도 검증 대상**이므로 sanitize·fix-squeeze 경로도 이 함수를 쓴다).
    - **직접 짜는 스크립트도 같다**: `reference/conversion.md`·`build-from-scratch.md`·`structural.md`의 재패키징 예시처럼 직접 ZIP을 쓸 때는 mimetype을 첫 항목·`ZIP_STORED`로 두고 나머지는 원본 항목의 압축 방식을 따르며, XML은 `standalone="yes"` 선언과 함께 직렬화한다. `hwpx_edit.py` 안의 새 경로는 `write_hwpx_zip()`·`serialize_xml()`을 쓴다(CLAUDE.md 코드 컨벤션).
    - **검증 습관**: 편집 명령을 추가·수정했으면 산출물에 `hwpx-validate-package`를 돌려 **원본과 같은 상태(WARN만)로 유지되는지** 확인한다. XSD 통과만으로는 이 부류를 못 잡는다.
16. **python-hwpx API 이름은 판마다 다르다**: 6.0에서 `doc.replace_text_in_runs()`가 `doc.text.replace`로 옮겨졌고(7.0 제거 예정) 3.x에는 `doc.text`가 없다. `cmd_find_replace`는 `doc.text.replace` → 없으면 `replace_text_in_runs`, `save_to_path` → 없으면 `save` 순으로 폴백한다. 한쪽 이름만 쓰면 다른 판에서 `AttributeError`로 전면 실패한다. → 사례
    - **표 셀 2차 lxml 패스는 그대로 유지해야 한다**: `doc.text.replace`도 `replace_text_in_runs`도 **표 셀 안 텍스트는 0건**이다. 엔진 버전이 올라가도 재실측 전까지 2차 패스를 걷어내지 말 것.
17. **편집 명령의 저장 경로에 조건을 걸지 말 것**: 저장이 조건부면 조건에 걸리지 않은 경우 `sanitize_header`(검은 배경)·`fix_empty_cells`·ZIP 레이아웃 복제가 통째로 건너뛰어져, 같은 명령이 입력에 따라 다른 품질을 낸다. `cmd_find_replace`는 표 치환이 0건이어도 항상 `save_hwpx()`를 탄다. **새 편집 명령을 만들 때도 저장은 무조건 `save_hwpx()` 한 경로로 모을 것**(15번의 규칙과 한 쌍). → 사례
18. **hwp2hwpx 변환본의 XML 1.0 불법 제어문자로 파싱 실패**: hwp2hwpx는 HWP 하이퍼링크 필드의 `Command` 문자열을 NUL 패딩째로 `<hp:stringParam>`에 써 넣는다. NUL은 XML 1.0에서 이스케이프로도 표현할 수 없어 lxml이 `Premature end of data in tag stringParam`으로 파싱을 중단한다. → 사례
    - **현재 구현**: HWPX를 여는 모든 경로가 적재 직후 `sanitize_xml_parts()`로 XML 파트(`.xml`·`.hpf`·`.rdf`, `BinData/` 제외)를 한 번 검사해 `[\x00-\x08\x0B\x0C\x0E-\x1F]`를 제거한다. 발견 시 `sanitize:` 한 줄(파일별 제거 바이트 수·코드값)을 남긴다. 탭·개행·복귀는 XML 1.0 합법이라 보존한다.
    - **파싱만이 아니라 산출물도 정제해야 한다**: 정제한 트리를 파싱에만 쓰고 `all_files`에 되쓰지 않으면, 재직렬화되지 않는 파트(header.xml·section1.xml·`--add-preview` 산출물)에 불법 바이트가 그대로 남아 `hwpx-validate-package`가 `malformed XML`을 낸다. 그래서 정제는 트리가 아니라 **바이트 딕셔너리(`all_files`)에** 적용한다.
    - **python-hwpx 경로는 별도 처리**: `HwpxDocument.open()`은 자체 파서라 위 경로를 타지 않는다. `sanitized_source()`가 정제 임시본(시스템 temp)을 만들어 넘기고 블록을 벗어날 때(오류·`sys.exit` 포함) 삭제한다.
    - **`--to-md`는 엔진 소관**: hwpx-tomd 0.2.1 이상이 같은 방어를 갖는다(이슈 hwpx-tomd#1, `requirements.txt` 핀 `>=0.2.1`).
