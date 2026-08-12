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
def com_insert(doc: str, img: str, anchor: str, w_mm: float, h_mm: float,
               keep_anchor: bool = False) -> None:
    """anchor 텍스트 자리에 이미지를 삽입(BinData 확보).

    keep_anchor=False(기본)면 anchor 텍스트를 지우고 그 자리에 넣는다.
    keep_anchor=True면 선택만 해제하고 넣어 ``(서명)``·``(인)`` 표시를 남긴다
    (한국 결재 문서 관행: 표시를 지우지 않고 그 위에 겹쳐 찍는다).
    """
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

        if keep_anchor:
            # 선택 해제만 → anchor 텍스트가 남고 캐럿은 그 자리에.
            # 그림은 XML 단계에서 절대좌표 floating으로 다시 배치되므로
            # 캐럿의 정확한 위치는 최종 좌표에 영향을 주지 않는다.
            hwp.HAction.Run("Cancel")
        else:
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


def _lineseg_attrs(tag: str) -> dict:
    """<hp:lineseg .../> 태그 문자열의 정수 속성을 dict로 반환(순서 무관).

    값 패턴에 부호를 포함한다. `(\\d+)`로 두면 음수 spacing·horzpos가 **키째로
    사라져** 호출부의 .get(...,0) 기본값에 조용히 흡수된다.
    """
    return {k: int(v) for k, v in re.findall(r'(\w+)="(-?\d+)"', tag)}


def _first_lineseg(fragment: str) -> dict | None:
    """XML 조각에서 첫 <hp:lineseg>의 속성을 dict로 반환."""
    m = re.search(r'<hp:lineseg\b[^>]*/>', fragment)
    return _lineseg_attrs(m.group(0)) if m else None


# 문단 좌표는 **컨테이너 기준 상대좌표**다. 종이 절대좌표(vertRelTo="PAPER")로
# 환산할 수 없는 컨테이너와, 본문이 아니어서 앵커 후보에서 빼야 하는 컨테이너.
_RELATIVE_CONTAINERS = ("hp:drawText", "hp:tc")
_NON_BODY_CONTAINERS = ("hp:header", "hp:footer", "hp:footnote", "hp:endnote")


def _paragraph_spans(src_xml: str) -> list:
    """(start, end, depth) 문단 구간 목록. 중첩(표 셀·글상자 안 문단) 포함.

    `<hp:p ` 문자열이나 rfind로 문단을 찾으면 **같은 문단 안에서 이미 열렸다
    닫힌 중첩 문단**(인라인 표의 셀 문단 등)이 선택돼, 개체 상대좌표를 종이
    절대좌표로 잘못 쓰게 된다. 여는/닫는 토큰을 스택으로 훑어 실제 포함관계를
    만든다.
    """
    spans, stack = [], []
    for m in re.finditer(r'<hp:p\b|</hp:p>', src_xml):
        if m.group(0) == "</hp:p>":
            if stack:
                start = stack.pop()
                spans.append((start, m.end(), len(stack)))
        else:
            stack.append(m.start())
    return sorted(spans)


def _own_fragment(src_xml: str, span: tuple, spans: list) -> str:
    """문단의 **자기 XML**(중첩 문단 구간을 도려낸 것).

    표 셀·글상자 문단이 앵커 문단 안에 있으면 그 셀의 <hp:linesegarray>가 앞서
    나오므로, 구간 전체에서 첫 lineseg를 집으면 셀 좌표를 문단 좌표로 오인한다.
    """
    start, end, depth = span
    body, cursor = [], start
    for s, e, d in spans:
        if d == depth + 1 and start < s < end:      # 직속 자식 문단은 도려냄
            body.append(src_xml[cursor:s])
            cursor = e
    body.append(src_xml[cursor:end])
    return ''.join(body)


def _own_text(src_xml: str, span: tuple, spans: list) -> str:
    """문단의 **자기 텍스트**(중첩 문단 제외, 태그 제거)."""
    return re.sub(r'<[^>]*>', '', _own_fragment(src_xml, span, spans))


def _open_containers(src_xml: str, pos: int, tags: tuple) -> bool:
    """pos 지점에서 tags 중 하나가 열려 있는지."""
    return any(src_xml.rfind(f"<{t}", 0, pos) > src_xml.rfind(f"</{t}>", 0, pos)
               for t in tags)


def anchor_geometry(src_xml: str, anchor: str) -> dict | None:
    """**편집 전 원본**에서 서명줄(anchor 문단)의 줄 기하 정보를 해석한다.

    왜 원본을 보는가: COM이 그림을 넣은 문단은 레이아웃 캐시가 무효화되어
    한글이 저장 시 그 문단의 <hp:linesegarray>를 **아예 빼고 쓴다**(비가시
    창이라 재계산도 하지 않음, 한컴 13.0 실측). 삽입 후 파일만 보면 서명줄
    좌표를 구할 방법이 없어 floating 배치가 통째로 실패한다. 원본 lineseg는
    그림 때문에 부풀지 않은 **서명줄 자체의 위치·높이**라 겹쳐 찍기에도 맞다.

    반환 dict에 진단 플래그를 함께 싣는다.
      relative=True   좌표가 글상자·표 기준이라 종이 절대좌표로 환산 불가
      estimated=True  앵커 문단에 캐시가 없어 **같은 컨테이너의 직전 형제 문단**
                      에서 한 줄 아래를 추정(vertpos + vertsize + spacing).
    """
    spans = _paragraph_spans(src_xml)
    if not spans:
        return None

    # 앵커가 든 문단 = 앵커 텍스트를 자기 텍스트로 가진 문단 중 가장 깊은 것.
    # 원문 XML 부분문자열이 아니라 태그를 벗긴 자기 텍스트로 찾으므로
    # "(서명)"이 여러 <hp:t> 런으로 쪼개져 있어도 잡힌다.
    # 머리글·꼬리말·각주는 COM의 본문 검색 대상이 아니므로 후보에서 뺀다.
    target = None
    for span in spans:
        if _open_containers(src_xml, span[0], _NON_BODY_CONTAINERS):
            continue
        if anchor in _own_text(src_xml, span, spans):
            target = span
            break
    if target is None:
        return None

    start, end, depth = target
    relative = _open_containers(src_xml, start, _RELATIVE_CONTAINERS)

    ls = _first_lineseg(_own_fragment(src_xml, target, spans))
    if ls:
        return dict(ls, relative=relative)

    # 캐시 없음 → 같은 컨테이너의 직전 형제 문단에서 추정.
    # 문서 전체에서 마지막 lineseg를 집으면 남의 컨테이너(표 셀) 좌표를
    # "직전 줄"이라 부르며 쓰게 된다.
    prev = None
    for span in spans:
        s, e, d = span
        if e <= start and d == depth and not _crosses_container(src_xml, e, start):
            prev = _first_lineseg(_own_fragment(src_xml, span, spans)) or prev
    if not prev or "vertpos" not in prev:
        return None
    return dict(prev,
                vertpos=prev["vertpos"] + prev.get("vertsize", 0) + prev.get("spacing", 0),
                relative=relative, estimated=True)


def _crosses_container(src_xml: str, a: int, b: int) -> bool:
    """a~b 사이에 컨테이너 경계(표·글상자 여닫기)가 있으면 True = 형제가 아님."""
    between = src_xml[a:b]
    return any(f"<{t}" in between or f"</{t}>" in between
               for t in _RELATIVE_CONTAINERS + ("hp:tbl", "hp:rect"))


def _norm(text: str) -> str:
    return re.sub(r'\s+', ' ', text).strip()


def inserted_pic_span(xml: str, src_xml: str | None, anchor: str,
                      img_name: str = "") -> tuple | None:
    """COM이 **방금 넣은** 그림과 그 문단 구간 (pic_start, pic_end, para_span)을 반환.

    단순히 첫 <hp:pic>을 집으면, 머리글 엠블럼·스캔 이미지처럼 **앞부분에 이미 있던
    그림**을 서명으로 오인해 그 그림을 서명 크기로 줄여 서명줄에 옮겨 놓고(진짜
    서명은 인라인으로 방치) 아무 신호 없이 [OK]를 낸다.

    ⚠ binaryItemIDRef로는 판별할 수 없다. 한글은 저장하면서 BinData 번호를
    **재부여**해, 새로 넣은 서명이 image1이 되고 원래 있던 그림이 image2로 밀린다
    (실측). 그래서 판별 순서는 (1) 삽입한 이미지 파일명이 <hp:shapeComment>에
    남는 점 (2) 앵커 문단의 자기 텍스트 일치 순이다.
    """
    pics = list(re.finditer(r'<hp:pic\b.*?</hp:pic>', xml, re.DOTALL))
    if not pics:
        return None

    picked = pics[0] if len(pics) == 1 else None

    if picked is None and img_name:
        named = [p for p in pics if img_name in p.group(0)]
        if len(named) == 1:
            picked = named[0]

    spans = _paragraph_spans(xml)

    def enclosing(pos):
        best = None
        for s, e, d in spans:
            if s <= pos < e and (best is None or d > best[2]):
                best = (s, e, d)
        return best

    if picked is None and src_xml and anchor:
        # 원본 앵커 문단의 자기 텍스트(앵커 제외)와 일치하는 문단의 그림을 고른다.
        # --keep-anchor 여부와 무관하게 성립한다.
        src_spans = _paragraph_spans(src_xml)
        key = ""
        for span in src_spans:
            t = _own_text(src_xml, span, src_spans)
            if anchor in t:
                key = _norm(t.replace(anchor, " "))
                break
        if key:
            hit = []
            for p in pics:
                para = enclosing(p.start())
                if para and key in _norm(_own_text(xml, para, spans)):
                    hit.append(p)
            if len(hit) == 1:
                picked = hit[0]

    if picked is None:
        print("[WARN] 문서에 그림이 여러 개인데 새로 삽입된 서명을 식별하지 "
              "못했습니다. 첫 그림을 서명으로 처리합니다 — 반드시 --pdf로 "
              "확인하세요.", file=sys.stderr)
        picked = pics[0]

    para = enclosing(picked.start())
    if para is None:
        return None
    return picked.start(), picked.end(), para


def to_floating(doc: str, out: str, w_hu: int, h_hu: int,
                horz_offset: int | None, vert_adjust: int, gap_hu: int,
                src_xml: str | None = None, anchor: str = "",
                img_name: str = "") -> tuple:
    """COM 삽입 결과를 floating PAPER 좌표로 전환 + 앵커 위 문단 이동 + lineseg 제거."""
    zin = zipfile.ZipFile(doc, "r")
    xml = zin.read("Contents/section0.xml").decode("utf-8")

    found = inserted_pic_span(xml, src_xml, anchor, img_name)
    if not found:
        zin.close()
        raise RuntimeError("삽입된 그림이 든 문단을 찾지 못함")
    pic_start, pic_end, (para_start, para_end, _) = found
    para = xml[para_start:para_end]
    pic = xml[pic_start:pic_end]

    # 서명줄의 줄 위치(vertpos)·줄높이(vertsize)·본문폭(horzsize).
    # 원본의 anchor 문단을 1순위로 보고, 없으면 삽입 후 문단에 남은 lineseg를 쓴다.
    ls = None
    if src_xml and anchor:
        ls = anchor_geometry(src_xml, anchor)
    if ls is None:
        ls = _first_lineseg(para)
    if ls is None or not {"vertpos", "vertsize", "horzsize"} <= set(ls):
        zin.close()
        raise RuntimeError(
            "서명줄 줄 정보(vertpos·vertsize·horzsize)를 읽지 못했습니다.\n"
            "  원인 후보: 원본에 레이아웃 캐시가 없고 같은 컨테이너에 기준 줄도 없음,\n"
            "  또는 anchor 텍스트가 본문에 없음(머리글·꼬리말은 대상이 아님).\n"
            "  --horz-offset / --vert-adjust로 좌표를 직접 주거나 --inline을 쓰세요.")
    vertpos, vertsize, horzsize = ls["vertpos"], ls["vertsize"], ls["horzsize"]

    # 개체 상대좌표는 종이 절대좌표로 환산할 수 없다. 경고만 하고 진행하면
    # 인쇄 영역 밖에 찍힌 문서가 [OK]·종료코드 0으로 나가므로, 사용자가 좌표를
    # 직접 지정해 책임을 진 경우에만 통과시킨다.
    if ls.get("relative"):
        if horz_offset is None and vert_adjust == 0:
            zin.close()
            raise RuntimeError(
                "서명란이 글상자·표 안에 있어 자동 좌표를 신뢰할 수 없습니다.\n"
                f"  그 개체 기준 원시값: vertpos={vertpos} horzsize={horzsize}\n"
                "  --horz-offset / --vert-adjust로 종이 기준 좌표를 지정하거나\n"
                "  --inline으로 인라인 삽입하세요(개체 기준 오프셋은 문서마다\n"
                "  상수라 한 번 맞추면 재사용됩니다).")
        print("[WARN] 서명란이 글상자·표 안입니다. 자동 계산값은 그 개체 기준이며 "
              "지정하신 오프셋으로만 보정됩니다. --pdf로 반드시 확인하세요.",
              file=sys.stderr)
    if ls.get("estimated"):
        print("[WARN] 서명줄에 레이아웃 캐시가 없어 같은 컨테이너의 직전 줄에서 "
              "추정했습니다. 세로뿐 아니라 **가로 기준폭(horzsize)도 그 줄에서 "
              "물려받으므로**, --pdf로 확인하고 --vert-adjust·--horz-offset "
              "양쪽을 보정하세요.", file=sys.stderr)

    mg = _page_margins(xml)
    body_top = mg["top"] + mg["header"]
    abs_y = body_top + vertpos
    vert = abs_y + (vertsize - h_hu) // 2 + vert_adjust          # 서명줄 세로 중앙
    if horz_offset is None:
        horz = mg["left"] + horzsize - w_hu - gap_hu             # 우측 정렬
    else:
        horz = horz_offset

    # pic을 원래 문단에서 떼어낸다(문단 안에 다른 그림이 있어도 그건 건드리지 않음)
    para = para[:pic_start - para_start] + para[pic_end - para_start:]
    # anchor 문단 lineseg 제거(텍스트가 바뀐 문단)
    para = re.sub(r'<hp:linesegarray>.*?</hp:linesegarray>', '', para, flags=re.DOTALL)
    xml = xml[:para_start] + para + xml[para_end:]

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
    prevs = list(re.finditer(r'<hp:p\b', xml[:para_start]))
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
    ap.add_argument("--keep-anchor", action="store_true",
                    help='anchor 텍스트를 지우지 않고 그 위에 겹쳐 찍는다'
                         '(한국 결재 문서 관행: "(서명)"·"(인)" 표시를 남긴 채 날인). '
                         '가로 위치는 기본 우측 정렬이므로 표시 위에 정확히 얹으려면 '
                         '--horz-offset으로 맞추고 --pdf로 확인한다.')
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

    # 서명줄 좌표 산출용 원본 XML(COM이 앵커 문단의 lineseg를 지우기 전에 확보)
    try:
        with zipfile.ZipFile(args.doc, "r") as zsrc:
            src_xml = zsrc.read("Contents/section0.xml").decode("utf-8")
    except Exception:
        src_xml = None

    # 입력을 출력으로 복사 후 그 위에서 작업(원본 비파괴)
    if os.path.abspath(out) != os.path.abspath(args.doc):
        import shutil
        shutil.copyfile(args.doc, out)

    com_insert(out, args.image, args.anchor, args.width_mm, h_mm,
               keep_anchor=args.keep_anchor)

    # --keep-anchor는 COM의 "Cancel" 액션이 선택을 실제로 해제했는지에 달려 있다.
    # 빌드에 따라 무시되면 InsertPicture가 선택된 앵커를 그대로 대체해, 옵션이
    # 막으려던 바로 그 결과(표시 삭제)가 조용히 나온다. 결과 파일로 확인한다.
    if args.keep_anchor:
        try:
            with zipfile.ZipFile(out, "r") as zchk:
                after = zchk.read("Contents/section0.xml").decode("utf-8")
            if args.anchor not in re.sub(r'<[^>]*>', '', after):
                print(f"[WARN] --keep-anchor를 지정했지만 {args.anchor!r} 표시가 "
                      "결과에 남아 있지 않습니다(설치된 한컴 빌드가 Cancel 액션을 "
                      "무시했을 수 있음). 원본과 대조해 확인하세요.", file=sys.stderr)
        except Exception:
            pass

    if not args.inline:
        try:
            vert, horz = to_floating(
                out, out, mm2hu(args.width_mm), mm2hu(h_mm),
                args.horz_offset, args.vert_adjust, mm2hu(args.gap_mm),
                src_xml=src_xml, anchor=args.anchor,
                img_name=os.path.basename(args.image))
        except RuntimeError as e:
            # 좌표를 신뢰할 수 없을 때의 안내는 트레이스백 없이 그대로 보여준다.
            # COM 삽입까지는 끝난 상태이므로 중간 산출물 경로도 알려준다.
            print(f"[ERROR] {e}\n  (COM 삽입까지 진행된 중간 파일: {out})",
                  file=sys.stderr)
            return 1
        mode = "겹쳐 찍기" if args.keep_anchor else "앵커 대체"
        print(f"[OK] 서명 삽입(floating, {mode}): {out}")
        print(f"     크기 {args.width_mm}x{h_mm}mm · vertOffset={vert} horzOffset={horz}")
        print("     위치가 어긋나면 --horz-offset / --vert-adjust 로 재실행하세요.")
        if args.keep_anchor and args.horz_offset is None:
            print(f"     겹쳐 찍기 기본 가로 위치는 우측 정렬입니다. "
                  f'"{args.anchor}" 표시 위에 정확히 얹으려면 '
                  f"--horz-offset {horz} 부터 조정하세요.")
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
