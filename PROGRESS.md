# 진행 상태

## 현재

- HWP 변환 번들: edc0527+p1 / hwplib 1.1.11 / hwpxlib 1.0.9. Windows 40건 대조 통과.
- rhwp 재설계(`docs/plans/2026-10-28-rhwp-redesign.md`): M1 어댑터·게이트 완료. 단위 75건 통과(`RHWP_INTEGRATION=1` 포함).
- 위원장 결정(2026-10-28): 실기 검증은 위원장이 실사용으로 한다. 축별 옵트인 단계 없이 플랜 원안대로 rhwp를 기본값으로 두고 JAR·COM은 명시 폴백으로 남긴다.

## 다음 단계

- M2~M5 착수 순서와 자리는 `docs/BACKLOG.md`.
- M6은 Windows hunyong-samsung rhwp 설치만. 한컴 2024 호환은 위원장 실사용에서 확인하고 문제가 보이면 그 축을 COM 폴백 기본으로 되돌린다.

## 미결 결정

- 없음. rhwp 재직렬화의 header.xml 손실(OUTLINE 7~9단계 `hp:switch`·미사용 paraPr)은 게이트 허용 목록으로 통과시키고 보고에 표기한다(2026-10-28).
