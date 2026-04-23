# 주의사항: 한컴 COM 자동화

1. **HWPX 이미지 삽입은 XML 직접 편집 불가**: `<hp:pic>` 요소를 XML로 직접 구성하면 `hc:` 네임스페이스 참조 불일치로 **문서 손상** (한글에서 "손상된 파일" 오류). XSD 검증을 통과해도 한글이 열 때 실패함. 반드시 **한컴 COM `hwp.InsertPicture()`로 삽입** → XML 후처리로 위치/배치 조정
2. **양식 채우기 + 이미지 삽입 워크플로우**: (1) hwpx_edit.py로 셀 채우기 (2) XML 직접 편집으로 날짜/텍스트 수정 (3) COM으로 이미지 삽입 → HWPX 저장 (4) XML 후처리로 이미지 위치 조정 (5) COM으로 PDF 저장. 3~5단계를 **별도 Python 스크립트**로 분리 (COM segfault 방지)
3. **HWPX→PDF 변환**: `simple-hwp2pdf`는 실제로 MS Word COM 기반이라 HWPX 미지원. 한컴오피스 COM `hwp.SaveAs(path, 'PDF', '')`가 가장 확실
4. **COM은 Pandoc 생성 HWPX 본문을 인식하지 못함** (2026.4 확인): `hwpx_convert.py`(Pandoc)로 생성한 HWPX를 COM으로 열면 **본문 내용이 0자**로 읽힘. COM 저장 시 빈 문서로 덮어씌워 내용 전체 소실. `Visible=False`와 무관하며 사용자 개입과도 무관한 근본적 호환성 문제임.
   - **머리글(헤더) 편집 불가**: COM이 Pandoc HWPX의 secPr/header 구조를 인식 못해 `InsertHeaderFooter` 액션이 빈 머리글만 생성
   - **본문 인라인 삽입은 가능**: `MoveDocBegin` → `InsertPicture` → `BreakPara` → **별도 파일로 SaveAs** 시 내용+이미지 모두 보존됨 (2026.3.31 세션에서 확인, 2026.4.16 재현 성공)
   - **핵심**: 반드시 `SaveAs(다른파일, 'HWPX', '')`로 별도 파일에 저장해야 함. 같은 파일에 SaveAs하면 COM이 Pandoc 내용을 덮어씀
   - 한컴 COM 생성 HWPX(또는 한글에서 직접 저장한 HWPX)에는 이 문제 없음 — Pandoc HWPX에만 해당
5. **보안 모듈 DLL 경로 불일치로 팝업 무한 발생**: 레지스트리 `HKCU\SOFTWARE\HNC\HwpAutomation\Modules`의 `FilePathCheckerModule` 값이 가리키는 경로에 **실제 DLL이 있어야** `RegisterModule`이 유효하다. 경로만 등록되고 파일이 다른 위치에 있으면(프로젝트 이동·재설치 후 흔함) `RegisterModule`은 조용히 성공값을 반환하지만 실제로는 로드 실패해 보안 팝업이 매 Open/SaveAs마다 뜬다. 증상: 스크립트가 반복 멈춤. 진단: `reg query "HKCU\SOFTWARE\HNC\HwpAutomation\Modules"` + `ls <경로>` 둘 다 일치 확인. DLL은 `<repo>/vendor/FilePathCheckerModuleExample.dll`에 배치 (설치: `vendor/README.md`).
6. **암호화된 HWPX는 `SaveAs`만으로 암호가 유지됨**: `Open(path, 'HWPX', 'password:xxx')`로 연 뒤 `SaveAs(path, 'HWPX', '')`하면 원본 암호가 그대로 재적용돼 결과물이 다시 암호화된다. 반드시 `FilePasswordChange` 액션(`String=""`, `Ask=0`)을 먼저 실행해야 평문 HWPX가 나온다. 상세: `reference/encrypted-hwpx.md`
