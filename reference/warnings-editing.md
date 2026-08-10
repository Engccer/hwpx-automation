# 주의사항: XML 편집 전반

1. **`--find/--replace` 2단계 치환**: (1) python-hwpx로 본문 런 치환 (서식 보존, 런 분할 텍스트 지원) → (2) lxml로 표 셀 내부만 추가 치환. 영역이 분리되어 중복 치환 없음. **주의**: `--find/--replace`는 내부적으로 `HwpxDocument.open()`을 사용하여 구조 검증이 엄격함 (PrvText.txt 필수). `--set-cell`은 lxml 직접 파싱이라 검증 없이 통과. HWP→HWPX 변환 파일은 `--find/--replace`가 실패할 수 있으므로, 본문 텍스트 치환은 XML 직접 편집 + `hwpx-pack` 리팩으로 우회
2. **편집 후 검증**: `hwpx-validate`로 무결성 확인, 양식 편집 시 `hwpx-page-guard`로 쪽수 드리프트 감지
3. **자동 출력 폴더와 `--set-cell` 다중 호출**: `-o` 미지정 시 입력 파일과 같은 디렉터리의 `_work-hwpx-automation/` 폴더에 저장 (원본 비파괴). **주의**: `--set-cell`을 여러 번 체이닝(`&&`)하면 매번 원본에서 읽어 `_work-hwpx-automation/`에 덮어쓰므로 **마지막 호출만 남는다**. 여러 셀을 채울 때는 Python 스크립트로 순차 호출하되 반드시 `-o`로 동일 파일을 지정하라
4. **서식 보존 원칙**: 가능한 한 `<hp:t>` 요소의 텍스트만 변경하고, XML 구조는 건드리지 않음
5. **병합 셀 주의**: `rowSpan`/`colSpan`을 변경하면 표 레이아웃이 깨질 수 있음. 한글에서 확인 필요
    - **span의 위치**: 좌표는 `<hp:cellAddr colAddr rowAddr>`, 병합 범위는 별도 `<hp:cellSpan colSpan rowSpan>`이다. cellAddr에서 span을 읽고 쓰면 항상 1x1로 읽히고 스키마 외 속성만 남는다(과거 `--split-cell` 버그, 이슈 #2로 2026-08-10 수정. 읽기 경로 `get_cell_span`은 2026-06-05에 같은 버그를 먼저 수정했었음)
    - **병합 해제는 span=1만으로 완성되지 않는다**: `rowSpan=N` 셀이 덮는 N-1개 행에는 해당 열의 `<hp:tc>`가 아예 없다. 해제하려면 행마다 새 tc를 생성해 `cellAddr` 좌표·`cellSz`를 채워야 한다(`colSpan`도 동일). 현재 `--split-cell`이 이를 자동 수행(앵커 deepcopy로 서식 보존 + 본문 비움 + `linesegarray` 제거)
    - **검증 사각지대**: span이 오염된 결과물도 `hwpx-validate`(XSD)와 `--to-md` 자가검증 recall을 모두 통과한다(lineseg 손상과 같은 계열). 병합 편집 후에는 `--info`의 행별 셀 수·`[병합:NxM]` 표시로 확인하라
6. **다중 섹션 문서**: `section1.xml` 등 존재 가능
7. **네임스페이스 필수**: `hp`, `hs`, `hh`, `hc` (`reference/format.md` 참조)
8. **HWP→HWPX 변환 후 검은 배경**: `--sanitize` 또는 `save_hwpx()` 자동 수정
9. **HWPML 2016 → 2011**: python-hwpx v2.8+는 네임스페이스 자동 변환 지원
10. **`--to-md`는 표만 추출**: 문서 제목, 지도교사, 머리글 등 **표 바깥 본문 텍스트는 `--to-md`에 나타나지 않는다**. 양식 편집 시 `--to-md` + `grep "<hp:t>" section0.xml` 병행으로 표 밖 텍스트를 반드시 확인
11. **HWP→HWPX 변환 후 `PrvText.txt` 누락**: `hwp2hwpx.bat` 변환 결과에 `Preview/PrvText.txt`가 없는 경우가 있음. python-hwpx API (`HwpxDocument.open()`)와 `hwpx-pack`이 이 파일 누락 시 실패하므로, 변환 직후 `mkdir -p Preview && touch Preview/PrvText.txt`로 빈 파일을 생성해 두라
12. **manifest.xml self-closing 태그 주의**: HWP→HWPX 변환 결과의 manifest.xml이 `<odf:manifest .../>` (self-closing) 형태일 수 있음. `replace('</odf:manifest>', ...)` 패턴이 실패하므로, self-closing 여부를 먼저 확인하고 처리
13. **양식의 "한 줄로 입력" 문단에 긴 텍스트 채우면 글자가 겹쳐 뭉개짐** (2026-07-06 보도자료 양식 실측): 사람이 디자인한 양식의 제목 문단에는 문단 모양 "한 줄로 입력"(`paraPr`의 `<hh:breakSetting ... lineWrap="SQUEEZE"/>`)이 걸려 있는 경우가 있다. 짧은 제목 전제의 디자인이라, 자동화로 **긴 제목·문장을 채우면 한글이 줄바꿈 대신 장평을 무제한 압축**해 글자가 서로 겹쳐 판독 불가가 된다(음수 자간 charPr과 결합하면 더 심함). `hwpx-validate`(XSD)와 `--to-md` recall로는 잡히지 않고 **렌더링(PDF·인쇄)에서만 드러난다**.
    - **감지(`--list-squeeze`)**: SQUEEZE 문단 중 추정 자연 폭(charPr height·장평·자간 반영 근사)이 컨테이너 가용 폭(셀·글상자·본문)의 1.1배를 넘는 문단만 나열.
    - **보정(`--fix-squeeze`)**: 원본 paraPr은 남기고 `lineWrap="BREAK"` 복제 paraPr을 새 id로 추가해 과압축 문단만 재지정(같은 paraPr을 공유하는 짧은 라벨의 의도된 디자인은 보존). 재지정 문단의 `linesegarray`는 제거해 한글이 재계산하게 함.
    - **자동 경고**: `--find/--replace`·`--set-cell`이 텍스트를 채운 직후 그 텍스트가 SQUEEZE 문단에서 과압축되면 stderr로 경고하고 `--fix-squeeze`를 안내한다.
    - **워크플로우 규칙**: 양식 채우기에서 제목·긴 문장을 채운 뒤에는 `--list-squeeze`로 확인하거나 PDF로 육안 검증하라. 자동 생성 문서의 제목은 폰트 축소·자간 압축으로 한 줄에 욱여넣지 말고 자연 줄바꿈(2줄)을 허용하는 것이 기본이다.
14. **양식의 기입 예시(placeholder) 글자모양이 그대로 상속됨** (2026-08-10 이슈 #4 실측): 관공서 배포 양식은 기입 예시를 **회색(`textColor="#808080"`)·이탤릭** charPr로 넣어 둔다. 이 셀을 `--set-cell`이나 python-hwpx `fill_by_path`로 채우면 텍스트만 바뀌고 `charPrIDRef`는 그대로라 **제출본이 회색 이탤릭으로 인쇄된다**. 13번(SQUEEZE)과 같은 계열: `hwpx-validate`·`--to-md` recall로는 안 잡히고 렌더링에서만 드러난다.
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
