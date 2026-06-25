# -*- coding: utf-8 -*-
"""hwpx_sign.py: HWPX 문서 서명란에 서명/도장 이미지를 삽입한다.

왜 이 도구가 필요한가 (단순 InsertPicture로 안 되는 이유):
  1) 한컴 COM ``InsertPicture``는 그림을 floating(textWrap=TOP_AND_BOTTOM) +
     ``curSz=0,0`` 으로 넣는다. 서명란이 페이지 하단에 있으면 큰 그림이 인라인
     으로 그 줄에 안 들어가 **다음 페이지로 밀린다**.
  2) floating + ``vertRelTo="PAPER"`` 절대좌표를 줘도, 앵커 문단이 페이지 끝이면
     개체가 **그 다음 페이지에 그려진다**(표시 페이지는 앵커 문단의 페이지를 따름).
     → 앵커를 페이지 내 **위쪽 문단**으로 옮겨야 한다.
  3) 텍스트가 바뀐 문단의 ``<hp:linesegarray>``(줄 레이아웃 캐시)를 제거하지 않으면
     한컴 COM이 **문서 열기를 거부**한다. 치명적으로 ``hwpx-validate``(XSD)와
     ``--to-md``(recall)는 모두 통과해 자동 검증으로 못 잡는다 → PDF 변환에서만 드러남.

해결: COM으로 anchor 텍스트 자리에 이미지를 넣어 BinData만 확보한 뒤, XML 후처리로
  (a) 그림을 floating PAPER 절대좌표로 전환(크기/curSz 보정),
  (b) 앵커를 anchor 문단의 직전(위쪽) 문단으로 이동,
  (c) 세로 좌표를 anchor 문단의 lineseg vertpos + 페이지 여백으로 자동 계산,
  (d) 변경된 문단들의 linesegarray 제거.

사용 예:
  python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(서명)"
  python hwpx_sign.py 동의서.hwpx --image 서명.png --anchor "(서명)" \
         --width-mm 22 --horz-offset 44000 --vert-adjust -800 --pdf

좌표 단위는 HWPUNIT(7200/inch). --horz-offset/--vert-adjust는 미세조정용이며,
미지정 시 우측 정렬 + 서명줄 세로 중앙으로 자동 배치한다. 한 번 --pdf로 위치를
확인하고 어긋나면 두 옵션만 바꿔 재실행하는 것이 최단 경로.
"""
from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import zipfile

MM = 7200 / 25.4  # 1mm in HWPUNIT


def mm2hu(mm: float) -> int:
    return round(mm * MM)


# ---------------------------------------------------------------- COM 단계
def com_insert(doc: str, img: str, anchor: str, w_mm: float, h_mm: float) -> None:
    """anchor 텍스트를 찾아 삭제하고 그 자리에 이미지를 삽입(BinData 확보)."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        try:
            hwp.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

        if not hwp.Open(os.path.abspath(doc), "HWPX", ""):
            raise RuntimeError(f"문서 열기 실패: {doc}")

        hwp.HAction.Run("MoveDocBegin")
        fr = hwp.HParameterSet.HFindReplace
        hwp.HAction.GetDefault("RepeatFind", fr.HSet)
        fr.FindString = anchor
        fr.IgnoreMessage = 1
        fr.Direction = 0
        fr.MatchCase = 1
        fr.ReplaceMode = 0
        if not hwp.HAction.Execute("RepeatFind", fr.HSet):
            hwp.Quit()
            raise RuntimeError(f"anchor 텍스트를 본문에서 찾지 못함: {anchor!r}")

        hwp.HAction.Run("Delete")  # 선택된 anchor 삭제 → 캐럿이 그 자리에
        ctrl = hwp.InsertPicture(os.path.abspath(img), True, 1, False, False, 0, w_mm, h_mm)
        if not ctrl:
            hwp.Quit()
            raise RuntimeError("InsertPicture 실패")

        hwp.SaveAs(os.path.abspath(doc), "HWPX", "")
        hwp.Quit()
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------- XML 단계
_PARA_WITH_PIC = r'<hp:p\b[^>]*>(?:(?!</hp:p>).)*?<hp:pic\b.*?</hp:pic>(?:(?!</hp:p>).)*?</hp:p>'


def _page_margins(xml: str) -> dict:
    m = re.search(r'<hp:margin\b[^>]*/>', xml)
    if not m:
        return dict(top=0, header=0, left=0, right=0)
    blob = m.group(0)
    def g(k):
        mm = re.search(k + r'="(\d+)"', blob)
        return int(mm.group(1)) if mm else 0
    return dict(top=g("top"), header=g("header"), left=g("left"), right=g("right"))


def to_floating(doc: str, out: str, w_hu: int, h_hu: int,
                horz_offset: int | None, vert_adjust: int, gap_hu: int) -> tuple:
    """COM 삽입 결과를 floating PAPER 좌표로 전환 + 앵커 위 문단 이동 + lineseg 제거."""
    zin = zipfile.ZipFile(doc, "r")
    xml = zin.read("Contents/section0.xml").decode("utf-8")

    pm = re.search(_PARA_WITH_PIC, xml, re.DOTALL)
    if not pm:
        zin.close()
        raise RuntimeError("삽입된 그림이 든 문단을 찾지 못함")
    para = pm.group(0)

    # 그림이 든 문단의 줄 위치(vertpos)·본문폭(horzsize)·줄높이(vertsize)
    ls = re.search(r'<hp:lineseg [^>]*vertpos="(\d+)"[^>]*vertsize="(\d+)"[^>]*horzsize="(\d+)"', para)
    if not ls:
        zin.close()
        raise RuntimeError("anchor 문단의 lineseg 정보를 읽지 못함")
    vertpos, vertsize, horzsize = int(ls.group(1)), int(ls.group(2)), int(ls.group(3))

    mg = _page_margins(xml)
    body_top = mg["top"] + mg["header"]
    abs_y = body_top + vertpos
    vert = abs_y + (vertsize - h_hu) // 2 + vert_adjust          # 서명줄 세로 중앙
    if horz_offset is None:
        horz = mg["left"] + horzsize - w_hu - gap_hu             # 우측 정렬
    else:
        horz = horz_offset

    # pic 분리
    picm = re.search(r'<hp:pic\b.*?</hp:pic>', para, re.DOTALL)
    pic = picm.group(0)
    para = para[:picm.start()] + para[picm.end():]
    # anchor 문단 lineseg 제거(텍스트가 바뀐 문단)
    para = re.sub(r'<hp:linesegarray>.*?</hp:linesegarray>', '', para, flags=re.DOTALL)
    xml = xml[:pm.start()] + para + xml[pm.end():]

    # pic 크기/위치 floating 보정
    pic = re.sub(r'<hp:orgSz [^/]*/>', f'<hp:orgSz width="{w_hu}" height="{h_hu}"/>', pic)
    pic = re.sub(r'<hp:curSz [^/]*/>', f'<hp:curSz width="{w_hu}" height="{h_hu}"/>', pic)
    pic = re.sub(r'<hp:sz [^/]*/>',
                 f'<hp:sz width="{w_hu}" widthRelTo="ABSOLUTE" height="{h_hu}" '
                 f'heightRelTo="ABSOLUTE" protect="0"/>', pic)
    pic = re.sub(r'(<hp:pic\b[^>]*?)textWrap="[^"]*"', r'\1textWrap="IN_FRONT_OF_TEXT"', pic)
    new_pos = ('<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="0" '
               'allowOverlap="1" holdAnchorAndSO="0" vertRelTo="PAPER" horzRelTo="PAPER" '
               f'vertAlign="TOP" horzAlign="LEFT" vertOffset="{vert}" horzOffset="{horz}"/>')
    pic = re.sub(r'<hp:pos\b[^>]*/>', lambda m: new_pos, pic, count=1)

    # 앵커를 anchor 문단의 직전(위쪽) 문단으로 이동(floating이 같은 페이지에 그려지도록)
    prevs = list(re.finditer(r'<hp:p\b', xml[:pm.start()]))
    if not prevs:
        zin.close()
        raise RuntimeError("앵커로 쓸 직전 문단이 없음(서명란이 문서 첫 문단)")
    p_start = prevs[-1].start()
    p_end = xml.find("</hp:p>", p_start) + len("</hp:p>")
    dpara = xml[p_start:p_end]
    # 직전 문단의 첫 hp:run 내부 끝(</hp:run>) 앞에 pic 삽입
    dpara2 = re.sub(r'</hp:run>', lambda m: pic + "</hp:run>", dpara, count=1)
    if dpara2 == dpara:  # run이 없으면 문단 끝에
        dpara2 = dpara.replace("</hp:p>", pic + "</hp:p>", 1)
    dpara2 = re.sub(r'<hp:linesegarray>.*?</hp:linesegarray>', '', dpara2, flags=re.DOTALL)
    xml = xml[:p_start] + dpara2 + xml[p_end:]

    # 재패키징
    tmp = out + ".tmp"
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "Contents/section0.xml":
            data = xml.encode("utf-8")
        if item.filename == "mimetype":
            zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
        else:
            zout.writestr(item, data)
    zout.close()
    zin.close()
    os.replace(tmp, out)
    return vert, horz


# ---------------------------------------------------------------- 진입점
def main() -> int:
    ap = argparse.ArgumentParser(description="HWPX 서명란에 서명/도장 이미지 삽입")
    ap.add_argument("doc", help="대상 HWPX")
    ap.add_argument("--image", required=True, help="서명/도장 이미지(투명/흰 배경 PNG 권장)")
    ap.add_argument("--anchor", default="(서명)", help='서명을 놓을 기준 텍스트(기본 "(서명)")')
    ap.add_argument("--width-mm", type=float, default=20.0, help="서명 가로 크기 mm(기본 20)")
    ap.add_argument("--height-mm", type=float, default=None,
                    help="서명 세로 크기 mm(미지정 시 이미지 종횡비로 자동)")
    ap.add_argument("--horz-offset", type=int, default=None,
                    help="종이 좌측 기준 가로 위치 HWPUNIT(미지정 시 우측 정렬)")
    ap.add_argument("--vert-adjust", type=int, default=0,
                    help="세로 미세조정 HWPUNIT(+아래 / -위)")
    ap.add_argument("--gap-mm", type=float, default=7.0,
                    help="우측 정렬 시 본문 우측 끝과의 여백 mm(기본 7)")
    ap.add_argument("--inline", action="store_true",
                    help="floating 전환 없이 COM 인라인 삽입만(서명란에 세로 여유가 충분할 때)")
    ap.add_argument("-o", "--output", default=None, help="출력 경로(기본 입력 옆 _서명.hwpx)")
    ap.add_argument("--pdf", action="store_true", help="삽입 후 검증용 PDF도 생성")
    args = ap.parse_args()

    if not os.path.isfile(args.doc):
        print(f"[ERROR] 문서 없음: {args.doc}", file=sys.stderr); return 1
    if not os.path.isfile(args.image):
        print(f"[ERROR] 이미지 없음: {args.image}", file=sys.stderr); return 1

    # 세로 크기 자동(종횡비)
    h_mm = args.height_mm
    if h_mm is None:
        try:
            from PIL import Image
            iw, ih = Image.open(args.image).size
            h_mm = round(args.width_mm * ih / iw, 2)
        except Exception:
            h_mm = round(args.width_mm * 0.65, 2)  # 폴백 비율

    out = args.output
    if out is None:
        base, ext = os.path.splitext(args.doc)
        out = base + "_서명" + ext

    # 입력을 출력으로 복사 후 그 위에서 작업(원본 비파괴)
    if os.path.abspath(out) != os.path.abspath(args.doc):
        import shutil
        shutil.copyfile(args.doc, out)

    com_insert(out, args.image, args.anchor, args.width_mm, h_mm)

    if not args.inline:
        vert, horz = to_floating(
            out, out, mm2hu(args.width_mm), mm2hu(h_mm),
            args.horz_offset, args.vert_adjust, mm2hu(args.gap_mm))
        print(f"[OK] 서명 삽입(floating): {out}")
        print(f"     크기 {args.width_mm}x{h_mm}mm · vertOffset={vert} horzOffset={horz}")
        print("     위치가 어긋나면 --horz-offset / --vert-adjust 로 재실행하세요.")
    else:
        print(f"[OK] 서명 삽입(inline): {out}")

    if args.pdf:
        pdf = os.path.splitext(out)[0] + ".pdf"
        here = os.path.dirname(os.path.abspath(__file__))
        rc = subprocess.run(
            [sys.executable, os.path.join(here, "hwpx_edit.py"), out, "--to-pdf", "-o", pdf]
        ).returncode
        if rc == 0:
            print(f"[OK] 검증 PDF: {pdf}")
        else:
            print("[WARN] PDF 변환 실패(한컴 COM 잔류 시 taskkill /F /IM Hwp.exe 후 재시도)",
                  file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
