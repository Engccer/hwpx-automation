# 진행 상태

## 현재

- HWP 변환 번들: edc0527+p1 / hwplib 1.1.11 / hwpxlib 1.0.9. Windows 40건 대조 통과.
- rhwp 재설계(`docs/plans/2026-10-28-rhwp-redesign.md`): M1 어댑터·게이트 완료. 단위 75건 통과(`RHWP_INTEGRATION=1` 포함).
- M2 HWP→HWPX rhwp 기본 완료: `--from-hwp`·`--engine`·`hwp_to_hwpx()`, JAR 교차 대조(JDK 없으면 unverified), `.hwp` 입력은 모든 편집·읽기 명령의 입구에서 변환.
- M4 PDF 백엔드 완료: `--to-pdf`는 모든 OS rhwp 기본, `--pdf-backend com`은 Windows 명시 폴백. Pandoc 산출물 PDF도 rhwp.
- M3 서명 삽입 rhwp 백엔드 완료: `hwpx_sign.py --backend rhwp|com`(기본 rhwp, 모든 OS), 게이트 `allow_rhwp` 허용 목록. 한컴에서 열기·위치는 위원장 실사용으로 확인 대기.
- 위원장 결정(2026-10-28): 실기 검증은 위원장이 실사용으로 한다. 축별 옵트인 단계 없이 플랜 원안대로 rhwp를 기본값으로 두고 JAR·COM은 명시 폴백으로 남긴다.

## 다음 단계

- M2~M5 착수 순서와 자리는 `docs/BACKLOG.md`.
- Windows hunyong-samsung rhwp v0.8.7 설치 완료(M6). 한컴 2024 호환은 위원장 실사용에서 확인하고 문제가 보이면 그 축을 COM 폴백 기본으로 되돌린다.

## 미결 결정

- 없음. rhwp 재직렬화의 header.xml 손실(OUTLINE 7~9단계 `hp:switch`·미사용 paraPr)은 게이트 허용 목록으로 통과시키고 보고에 표기한다(2026-10-28).
