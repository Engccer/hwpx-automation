"""한컴 COM PDF 내보내기: 변경 추적 형식 경고를 저장 구간에서만 승인."""

import os
import tempfile


def save_pdf(hwp, output):
    """Raw HwpObject를 사용한다. 원본 변경 이력은 수정하지 않는다."""
    output = os.path.abspath(output)
    os.makedirs(os.path.dirname(output), exist_ok=True)
    # 기존 PDF를 덮어쓸 때의 확인창은 자동 응답 범위에 포함하지 않는다.
    with tempfile.TemporaryDirectory(prefix="hwp-pdf-", dir=os.path.dirname(output)) as stage:
        staged = os.path.join(stage, "export.pdf")
        previous = hwp.SetMessageBoxMode(0x10000)
        try:
            if not hwp.SaveAs(staged, "PDF", ""):
                raise RuntimeError("한컴 COM PDF 저장이 취소되거나 실패했습니다.")
        finally:
            # 0은 설정을 바꾸지 않으므로 해당 종류를 먼저 해제한 뒤 복원한다.
            hwp.SetMessageBoxMode(0xF0000)
            if previous & 0xF0000:
                hwp.SetMessageBoxMode(previous & 0xF0000)
        with open(staged, "rb") as result:
            if result.read(5) != b"%PDF-":
                raise RuntimeError("한컴 COM 출력이 PDF 형식이 아닙니다.")
        os.replace(staged, output)
