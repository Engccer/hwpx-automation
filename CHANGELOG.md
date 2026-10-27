# 변경 이력

## 2026-10-28

- M6: Windows hunyong-samsung에 rhwp v0.8.7(`C:\Users\pc\.local\bin\rhwp.exe`, SHA256 릴리스 대조)을 설치했다. 한컴 2024 호환은 위원장 실사용에서 확인한다.
- rhwp 중심 재설계 M1: `rhwp_cli.py`(어댑터)와 `hwpx_gate.py`(독립 손실 게이트)를 추가하고 `rhwp_pdf.py`를 어댑터 위로 옮겼다. 봉투·exit code 계약 검사, 0.8.7 하한, 임시 폴더→검증→교체를 한곳에 모았다. 플랜 `docs/plans/2026-10-28-rhwp-redesign.md`, codex 적대적 리뷰 반영.
- 실측: rhwp HWPX 왕복은 자가 검증 PASS지만 게이트가 header.xml `hp:switch`·paraPr 손실을 잡는다. HWP 3건의 rhwp·JAR 변환물 교차 대조는 글자·표·셀·그림·BinData 수 일치.
- 버전 1.4.0 → 1.5.0.

## 2026-10-26

- HWP 변환 번들을 상류 edc0527과 hwplib 1.1.11로 갱신했다. 컨트롤 범위 검사 패치는 유지한다.
- Windows에서 HWP 40건의 변환·XML 본문·표 수를 이전 번들과 대조하고 기존 문제 문서의 한컴 COM 변환 결과와 Markdown 일치를 확인했다.
- Windows·Unix 래퍼의 JAR 참조와 번들 라이선스 버전 표기를 맞췄다.
