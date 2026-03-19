# 주의사항: XML 편집 전반

1. **`--find/--replace` 2단계 치환**: (1) python-hwpx로 본문 런 치환 (서식 보존, 런 분할 텍스트 지원) → (2) lxml로 표 셀 내부만 추가 치환. 영역이 분리되어 중복 치환 없음. **주의**: `--find/--replace`는 내부적으로 `HwpxDocument.open()`을 사용하여 구조 검증이 엄격함 (PrvText.txt 필수). `--set-cell`은 lxml 직접 파싱이라 검증 없이 통과. HWP→HWPX 변환 파일은 `--find/--replace`가 실패할 수 있으므로, 본문 텍스트 치환은 XML 직접 편집 + `hwpx-pack` 리팩으로 우회
2. **편집 후 검증**: `hwpx-validate`로 무결성 확인, 양식 편집 시 `hwpx-page-guard`로 쪽수 드리프트 감지
3. **자동 출력 폴더와 `--set-cell` 다중 호출**: `-o` 미지정 시 입력 파일과 같은 디렉터리의 `_output/` 폴더에 저장 (원본 보존). **주의**: `--set-cell`을 여러 번 체이닝(`&&`)하면 매번 원본에서 읽어 `_output/`에 덮어쓰므로 **마지막 호출만 남는다**. 여러 셀을 채울 때는 Python 스크립트로 순차 호출하되 반드시 `-o`로 동일 파일을 지정하라
4. **서식 보존 원칙**: 가능한 한 `<hp:t>` 요소의 텍스트만 변경하고, XML 구조는 건드리지 않음
5. **병합 셀 주의**: `rowSpan`/`colSpan`을 변경하면 표 레이아웃이 깨질 수 있음. 한글에서 확인 필요
6. **다중 섹션 문서**: `section1.xml` 등 존재 가능
7. **네임스페이스 필수**: `hp`, `hs`, `hh`, `hc` (`reference/format.md` 참조)
8. **HWP→HWPX 변환 후 검은 배경**: `--sanitize` 또는 `save_hwpx()` 자동 수정
9. **HWPML 2016 → 2011**: python-hwpx v2.8+는 네임스페이스 자동 변환 지원
10. **`--to-md`는 표만 추출**: 문서 제목, 지도교사, 머리글 등 **표 바깥 본문 텍스트는 `--to-md`에 나타나지 않는다**. 양식 편집 시 `--to-md` + `grep "<hp:t>" section0.xml` 병행으로 표 밖 텍스트를 반드시 확인
11. **HWP→HWPX 변환 후 `PrvText.txt` 누락**: `hwp2hwpx.bat` 변환 결과에 `Preview/PrvText.txt`가 없는 경우가 있음. python-hwpx API (`HwpxDocument.open()`)와 `hwpx-pack`이 이 파일 누락 시 실패하므로, 변환 직후 `mkdir -p Preview && touch Preview/PrvText.txt`로 빈 파일을 생성해 두라
12. **manifest.xml self-closing 태그 주의**: HWP→HWPX 변환 결과의 manifest.xml이 `<odf:manifest .../>` (self-closing) 형태일 수 있음. `replace('</odf:manifest>', ...)` 패턴이 실패하므로, self-closing 여부를 먼저 확인하고 처리
