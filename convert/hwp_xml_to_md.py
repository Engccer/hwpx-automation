#!/usr/bin/env python3
"""pyhwp `hwp5proc xml` 출력에서 표 구조를 보존한 마크다운 텍스트 추출.

hwp2hwpx(Java)가 EmptyStackException으로 실패하는 HWP의 폴백 경로.
각 텍스트는 정확히 한 번만 출력: 문단은 자기 직속 텍스트만(하위 표 제외),
표는 행 단위로 셀 내용을 ' | '로 이어 붙인다.
"""
import sys
import xml.etree.ElementTree as ET

PRUNE = {"TableControl"}  # 이 경계 아래 텍스트는 해당 노드 처리 때만 출력


def direct_text(node):
    """node 아래 Text를 PRUNE 경계에서 잘라 수집."""
    parts = []
    stack = list(node)
    while stack:
        el = stack.pop(0)
        if el.tag in PRUNE:
            continue
        if el.tag == "Text" and el.text:
            parts.append(el.text)
        stack = list(el) + stack
    return " ".join(" ".join(parts).split())


def controls(node):
    """node 아래 PRUNE 컨트롤을 문서 순서로, 중첩 없이 수집."""
    found = []
    stack = list(node)
    while stack:
        el = stack.pop(0)
        if el.tag in PRUNE:
            found.append(el)
            continue
        stack = list(el) + stack
    return found


def render_cell(cell):
    lines = []
    for p in cell.findall("Paragraph"):
        txt = direct_text(p)
        if txt:
            lines.append(txt)
        for ctl in controls(p):
            lines.extend(render_table(ctl))
    return " / ".join(lines)


def render_table(tbl):
    rows = []
    stack = list(tbl)
    trs = []
    while stack:
        el = stack.pop(0)
        if el.tag == "TableRow":
            trs.append(el)
            continue
        if el.tag in PRUNE:
            continue
        stack = list(el) + stack
    for tr in trs:
        cells = [render_cell(c) for c in tr.findall("TableCell")]
        if any(c.strip() for c in cells):
            rows.append("| " + " | ".join(cells) + " |")
    return rows


def render_body(node, out):
    for child in node:
        if child.tag == "Paragraph":
            txt = direct_text(child)
            if txt:
                out.append(txt)
            for ctl in controls(child):
                out.extend(render_table(ctl))
                out.append("")
        elif child.tag in PRUNE:
            out.extend(render_table(child))
            out.append("")
        else:
            render_body(child, out)


def main(xml_path, out_path):
    tree = ET.parse(xml_path)
    out = []
    render_body(tree.getroot(), out)
    lines, prev_blank = [], False
    for ln in out:
        blank = not ln.strip()
        if blank and prev_blank:
            continue
        lines.append(ln)
        prev_blank = blank
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines).strip() + "\n")
    print(f"lines={len(lines)}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
