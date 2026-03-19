# 주의사항: 한컴 COM 자동화

1. **HWPX 이미지 삽입은 XML 직접 편집 불가**: `<hp:pic>` 요소를 XML로 직접 구성해도 한글이 인식하지 않음 (XSD 검증은 통과하지만 렌더링 안 됨). 반드시 **한컴 COM `hwp.InsertPicture()`로 삽입** → XML 후처리로 위치/배치 조정
2. **양식 채우기 + 이미지 삽입 워크플로우**: (1) hwpx_edit.py로 셀 채우기 (2) XML 직접 편집으로 날짜/텍스트 수정 (3) COM으로 이미지 삽입 → HWPX 저장 (4) XML 후처리로 이미지 위치 조정 (5) COM으로 PDF 저장. 3~5단계를 **별도 Python 스크립트**로 분리 (COM segfault 방지)
3. **HWPX→PDF 변환**: `simple-hwp2pdf`는 실제로 MS Word COM 기반이라 HWPX 미지원. 한컴오피스 COM `hwp.SaveAs(path, 'PDF', '')`가 가장 확실
