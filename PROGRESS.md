# 진행 상태

## 현재

- HWP 변환 번들: edc0527+p1 / hwplib 1.1.11 / hwpxlib 1.0.9. Windows 40건 대조 통과.
- rhwp 재설계(`docs/plans/2026-10-28-rhwp-redesign.md`): M1 어댑터·게이트 완료. 단위 75건 통과(`RHWP_INTEGRATION=1` 포함).
- rhwp 경로는 아직 어느 축에서도 기본값이 아니다. `--to-pdf`(비Windows)만 종전처럼 rhwp.

## 다음 단계

- M2~M5 착수 순서와 자리는 `docs/BACKLOG.md`.
- M6 Windows hunyong-samsung rhwp 설치는 M1 직후 병행. 축별 기본값 승격은 한컴 2024 실기 확인이 선행 조건.

## 미결 결정

- rhwp가 HWPX를 다시 쓸 때 사라지는 header.xml OUTLINE 7~9단계 `hp:switch` 블록·미사용 paraPr을 "허용 손실"로 둘지. M3·M5 산출물이 게이트를 지나려면 이 판정이 필요하다(위원장).
