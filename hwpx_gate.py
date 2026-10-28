"""독립 손실 게이트: rhwp가 쓴 HWPX를 rhwp와 무관한 눈으로 검사한다.

rhwp는 문서를 자기 내부 구조(IR)로 읽어 전체를 다시 써서 저장한다. IR이 담지 않는
것은 조용히 사라지고, rhwp 자신의 `--verify`·`hwpx-roundtrip`은 그 손실을 보지 못한다
(실측 2026-10-28: 치환 1회에 header.xml의 hp:switch 23→20). 그래서 산출물은 원본과
독립적으로 대조한다.

검사 항목
- 요소 census: 모든 Contents/*.xml에서 (태그, 정렬된 속성 묶음) 단위로 개수를 세어
  사라지거나 바뀐 요소를 보고한다. 호출자가 의도한 변화는 allow로 제외한다.
- 편집 대조(edits): rhwp 편집 산출물은 census 대신 section*.xml을 문서 순서의 토큰열로
  펴서 원본과 위치별로 대조한다. 편집 대상 영역(셀 안쪽·누름틀 값·삽입한 그림)만
  가리고, 그 안은 대상 종류별 기대값과 따로 대조한다. 대상 영역은 산출물이 아니라
  원본 XML과 요청값으로 정한다.
- BinData: 파일 집합과 SHA-256. 그림·OLE·차트 변형을 잡는다.
- 텍스트: hwpx-tomd로 두 문서를 마크다운으로 바꿔 글자 멀티셋을 비교한다.
- 스키마: hwpx-validate.

상태는 passed / failed / unverified 셋이다. 검사 도구가 없어 못 본 항목이 있으면
passed가 아니라 unverified다.
"""

from __future__ import annotations

import hashlib
import re
import shutil
import subprocess
import zipfile
from collections import Counter
from dataclasses import dataclass, field

import xml.etree.ElementTree as ET

VOLATILE_ATTRS = set()  # 편집마다 바뀌는 것이 정상인 속성이 발견되면 여기에 적는다


@dataclass
class GateReport:
    status: str = "passed"  # passed | failed | unverified
    lost: dict = field(default_factory=dict)      # census key -> (before, after), after < before
    gained: dict = field(default_factory=dict)    # census key -> (before, after), after > before
    bindata_removed: list = field(default_factory=list)
    bindata_changed: list = field(default_factory=list)
    bindata_added: list = field(default_factory=list)
    text_missing: Counter = field(default_factory=Counter)   # 원본에 있고 산출물에 없는 글자
    text_extra: Counter = field(default_factory=Counter)
    validate_ok: bool | None = None
    messages: list = field(default_factory=list)
    allowed: dict = field(default_factory=dict)   # lost와 같은 키. allow_rhwp가 통과시킨 rhwp 재직렬화 손실

    @property
    def passed(self):
        return self.status == "passed"

    def fail(self, msg):
        self.status = "failed"
        self.messages.append(msg)

    def unverified(self, msg):
        if self.status != "failed":
            self.status = "unverified"
        self.messages.append(msg)

    def summary(self):
        head = {"passed": "게이트 통과", "failed": "게이트 실패", "unverified": "게이트 미검증"}[self.status]
        line = head + ("" if not self.messages else ": " + " / ".join(self.messages))
        return line + ("\n" + self.allowed_line() if self.allowed else "")

    def allowed_line(self):
        tags = Counter()
        for (_, tag, _), (a, b) in self.allowed.items():
            tags["p id 재부여" if tag == "p" else tag] += a - b
        return "허용된 손실(rhwp 재직렬화): " + ", ".join(f"{t}×{n}" for t, n in tags.most_common())


# ---------- 패키지 읽기 ----------

def read_parts(path):
    """HWPX 패키지의 모든 파트. hwp2hwpx JAR 산출물은 XML에 NUL 같은 불법 바이트를
    남기므로 저장소 규칙대로 읽은 직후 sanitize_xml_parts를 태운다."""
    from hwpx_edit import sanitize_xml_parts
    with zipfile.ZipFile(path) as zf:
        parts = {name: zf.read(name) for name in zf.namelist() if not name.endswith("/")}
    sanitize_xml_parts(parts)
    return parts


def xml_parts(parts):
    return {n: b for n, b in parts.items() if n.startswith("Contents/") and n.endswith(".xml")}


def bindata_parts(parts):
    return {n: b for n, b in parts.items() if n.startswith("BinData/")}


# ---------- census ----------

def census_key(el, ignore=()):
    """(태그 로컬명, 정렬된 (속성, 값) 튜플). ignore의 (태그, 속성) 쌍은 뺀다."""
    tag = el.tag.split("}")[-1]
    attrs = tuple(sorted((k.split("}")[-1], v) for k, v in el.attrib.items()
                         if k.split("}")[-1] not in VOLATILE_ATTRS
                         and (tag, k.split("}")[-1]) not in ignore))
    return tag, attrs


def element_census(xml_bytes, ignore=()):
    """(태그 로컬명, 정렬된 (속성, 값) 튜플) -> 개수."""
    census = Counter()
    for _, el in ET.iterparse(_bytes_io(xml_bytes)):
        census[census_key(el, ignore)] += 1
        el.clear()
    return census


def zoned_census(xml_bytes, in_zone):
    """census를 둘로 나눈다: in_zone(el)이 참인 요소와 그 하위(zone), 나머지(rest)."""
    rest, zone = Counter(), Counter()
    depth = 0
    for event, el in ET.iterparse(_bytes_io(xml_bytes), events=("start", "end")):
        if event == "start":
            if depth or in_zone(el):
                depth += 1
            continue
        (zone if depth else rest)[census_key(el)] += 1
        if depth:
            depth -= 1
        el.clear()
    return rest, zone


def _bytes_io(b):
    from io import BytesIO
    return BytesIO(b)


def tag_census(xml_bytes):
    """태그 로컬명 -> 개수(속성 무시). 보고 문구를 짧게 만들 때 쓴다."""
    c = Counter()
    for (tag, _), n in element_census(xml_bytes).items():
        c[tag] += n
    return c


def compare_census(before, after):
    """바뀐 키만 {key: (before, after)}로 돌려준다."""
    return {k: (before.get(k, 0), after.get(k, 0))
            for k in set(before) | set(after) if before.get(k, 0) != after.get(k, 0)}


# ---------- rhwp 재직렬화 허용 목록 ----------
# 실측(2026-10-28, rhwp v0.8.7 insert-image, 한컴 저장 공문 2건·python-hwpx 표본):
#  - header.xml: 본문이 쓰지 않는 paraPr 안의 OUTLINE 8~10단계(level 7~9)
#    hp:switch를 default 갈래(heading type=NONE)로 접는다. 문단이 그 paraPr를
#    쓰지 않으므로 화면에 영향이 없다. 본문이 쓰는 paraPr의 손실은 실패로 둔다.
#  - 본문 파트: 모든 hp:p의 id를 0부터 다시 매긴다(문단 수·내용은 그대로).
#  - 본문 파트: 그림의 hp:shadow type="NONE"(그림자 없음 표시)을 지운다.

def used_parapr_ids(parts):
    """본문 파트(header.xml 밖)의 요소가 참조하는 paraPr id 집합. 파싱하지 못한 파트가
    있으면 None(무엇이 쓰이는지 모르므로 미사용 판정을 하지 않는다)."""
    used = set()
    for name, data in xml_parts(parts).items():
        if name == "Contents/header.xml":
            continue
        try:
            for _, el in ET.iterparse(_bytes_io(data)):
                for k, v in el.attrib.items():
                    if k.split("}")[-1] == "paraPrIDRef":
                        used.add(v.strip())
                el.clear()
        except ET.ParseError:
            return None
    return used


def split_rhwp_known(name, before, after, used_parapr):
    """(실패로 볼 census 차이, 허용할 손실)로 나눈다. 허용 손실은 lost와 같은 형식."""
    if name == "Contents/header.xml" and used_parapr is None:
        return compare_census(element_census(before), element_census(after)), {}
    if name == "Contents/header.xml":
        def unused_parapr(el):
            return el.tag.split("}")[-1] == "paraPr" and el.get("id") not in used_parapr
        a_rest, a_zone = zoned_census(before, unused_parapr)
        b_rest, b_zone = zoned_census(after, unused_parapr)
        zone = compare_census(a_zone, b_zone)
        known = {k: v for k, v in zone.items() if v[1] < v[0]}
        diff = compare_census(a_rest, b_rest)
        diff.update({k: v for k, v in zone.items() if v[1] > v[0]})  # 증가는 gained로 보고만
        return diff, known

    no_pid = {("p", "id")}
    diff = compare_census(element_census(before, no_pid), element_census(after, no_pid))
    # 그림자 변화가 전부 "그림자 없음" 표시의 소실일 때만 허용한다(NONE→DROP 같은 변경은 실패)
    shadows = {k: v for k, v in diff.items() if k[0] == "shadow"}
    known = {}
    if shadows and all(("type", "NONE") in k[1] and b < a for k, (a, b) in shadows.items()):
        known = {k: diff.pop(k) for k in shadows}
    # id를 빼고 센 문단이 완전히 같을 때(수·나머지 속성)만 재부여로 본다
    if not any(k[0] == "p" for k in diff):
        full = compare_census(element_census(before), element_census(after))
        known.update({k: v for k, v in full.items() if k[0] == "p" and v[1] < v[0]})
    return diff, known


# ---------- 편집 대조: 순서 보존 토큰열 ----------
# rhwp 편집(set-cell·fill-fields·insert-image)은 대상 밖을 바꾸지 않아야 한다. 실측(2026-10-28,
# v0.8.7): 대상 밖 section 차이는 p id 재부여뿐이고, 셀 값이 길어 줄이 늘어도 lineseg·cellSz는
# 그대로다. 그래서 대상 밖은 개수가 아니라 위치·부모·순서·속성·글자까지 같아야 통과한다.
# 대상 영역은 산출물이 아니라 원본 XML과 요청값으로 정한다(산출물에서도 같은 규칙으로 찾는다).

@dataclass(frozen=True)
class CellEdit:
    """rhwp edit set-cell 대상. table은 문서 순서의 최상위 표(다른 표 안에 있지 않은 tbl) 번호,
    row·col은 cellAddr. old_text를 주면 원본 칸 글자와 교차 확인한다(공백 무시)."""
    table: int
    row: int
    col: int
    text: str
    old_text: str | None = None


@dataclass(frozen=True)
class FieldEdit:
    """rhwp edit fill-fields 대상. 이름이 같은 CLICK_HERE 누름틀 중 occurrence번째(0부터),
    None이면 그 이름 전부."""
    name: str
    value: str
    occurrence: int | None = None


@dataclass(frozen=True)
class ImageInsert:
    """서명 그림 삽입 대상. section의 최상위 문단 paragraph에 그림(pic) 하나만 더해진다."""
    section: int
    paragraph: int


class EditTargetError(ValueError):
    """편집 대상을 하나로 정하지 못했다."""


FREE_IN_CELL = {"p", "run", "t", "linesegarray", "lineseg"}  # 칸 내용 교체로 바뀌어도 되는 요소
FREE_IN_FIELD = {"run", "t"}                                  # 누름틀 값 자리에서 바뀌어도 되는 요소
FIELD_VOLATILE = {"dirty"}  # rhwp가 채운 누름틀에서 바꾸는 fieldBegin 속성(실측)
# rhwp 재직렬화가 모든 fieldBegin에서 지우는 zorder="-1"(실측, 채우지 않은 누름틀 포함)


def _local(tag):
    return tag.split("}")[-1]


def section_number(name):
    m = re.fullmatch(r"Contents/section(\d+)\.xml", name)
    return int(m.group(1)) if m else None


class _Section:
    """section 트리 하나와 부모 지도."""

    def __init__(self, name, data):
        self.name = name
        self.number = section_number(name)
        self.root = ET.fromstring(data)
        self.parent = {c: p for p in self.root.iter() for c in p}

    def ancestors(self, el):
        while el in self.parent:
            el = self.parent[el]
            yield el

    def nearest(self, el, tag):
        return next((a for a in self.ancestors(el) if _local(a.tag) == tag), None)


def _sections(parts):
    names = sorted((n for n in xml_parts(parts) if section_number(n) is not None), key=section_number)
    return [_Section(n, parts[n]) for n in names]


def _find_cell(sections, edit):
    """(구역, 칸의 subList)."""
    tables = [(sec, el) for sec in sections for el in sec.root.iter()
              if _local(el.tag) == "tbl" and sec.nearest(el, "tbl") is None]
    if not 0 <= edit.table < len(tables):
        raise EditTargetError(f"최상위 표 {edit.table}번이 없습니다(최상위 표 {len(tables)}개).")
    sec, tbl = tables[edit.table]
    hits = []
    for tc in tbl.iter():
        if _local(tc.tag) != "tc" or sec.nearest(tc, "tbl") is not tbl:
            continue
        addr = next((c for c in tc if _local(c.tag) == "cellAddr"), None)
        if addr is not None and (addr.get("rowAddr"), addr.get("colAddr")) == (str(edit.row), str(edit.col)):
            hits.append(tc)
    subs = [c for tc in hits for c in tc if _local(c.tag) == "subList"]
    if len(hits) != 1 or len(subs) != 1:
        raise EditTargetError(f"표 {edit.table}의 칸 ({edit.row},{edit.col})을 하나로 정하지 못했습니다"
                              f"(일치 {len(hits)}개).")
    return sec, subs[0]


def _find_fields(sections, edit):
    """[(구역, fieldBegin을 품은 ctrl, fieldEnd를 품은 ctrl, 문단)]."""
    found = []
    for sec in sections:
        ends = {e.get("beginIDRef"): e for e in sec.root.iter() if _local(e.tag) == "fieldEnd"}
        for fb in sec.root.iter():
            if (_local(fb.tag) != "fieldBegin" or fb.get("type") != "CLICK_HERE"
                    or fb.get("name") != edit.name):
                continue
            fe = ends.get(fb.get("id"))
            p = sec.nearest(fb, "p")
            if fe is None or p is None or sec.nearest(fe, "p") is not p:
                raise EditTargetError(f"누름틀 {edit.name!r}의 끝(fieldEnd)이 같은 문단에 없습니다.")
            c_begin, c_end = sec.parent.get(fb), sec.parent.get(fe)
            if c_begin is None or c_end is None or _local(c_begin.tag) != "ctrl" or _local(c_end.tag) != "ctrl":
                raise EditTargetError(f"누름틀 {edit.name!r}의 시작·끝이 ctrl 안에 있지 않습니다.")
            found.append((sec, c_begin, c_end, p))
    if edit.occurrence is not None:
        if not 0 <= edit.occurrence < len(found):
            raise EditTargetError(f"누름틀 {edit.name!r}의 {edit.occurrence}번째가 없습니다({len(found)}개).")
        found = [found[edit.occurrence]]
    if not found:
        raise EditTargetError(f"누름틀 {edit.name!r}이 없습니다.")
    return found


def _targets(sections, edits):
    """편집마다 [(구역 번호, 종류, 기대값 정보, 가릴 범위 지정)]. 종류별 범위 지정은 _tokens가 푼다."""
    out = []
    for e in edits:
        if isinstance(e, CellEdit):
            sec, sub = _find_cell(sections, e)
            out.append((sec.number, ("cell", sub), e))
        elif isinstance(e, FieldEdit):
            out += [(sec.number, ("field", cb, ce, p), e) for sec, cb, ce, p in _find_fields(sections, e)]
        elif isinstance(e, ImageInsert):
            sec = next((s for s in sections if s.number == e.section), None)
            top = [p for p in sec.root if _local(p.tag) == "p"] if sec is not None else []
            if not 0 <= e.paragraph < len(top):
                raise EditTargetError(f"구역 {e.section}의 문단 {e.paragraph}이 없습니다.")
            out.append((sec.number, ("image", top[e.paragraph]), e))
        else:
            raise TypeError(f"알 수 없는 편집 대상: {e!r}")
    return out


def _tokens(sec, marks, drop=None):
    """sec를 (시작 태그+속성 | 글자 | 끝 태그) 토큰열로 편다. marks[i]의 자리는 ("mask", i)
    하나로 바꾸고 원래 토큰열은 regions[i]에 둔다. drop 요소(삽입한 그림)는 통째로 뺀다.
    rhwp 재직렬화 허용분(p id, shadow type=NONE, fieldBegin zorder="-1")은 여기서 정규화하고
    shadow·zorder 수를 센다."""
    field_begins = {fb for m in marks if m[0] == "field" for fb in m[1]}
    tokens, spans, seen = [], {}, Counter()

    def attrs(el):
        a = {_local(k): v for k, v in el.attrib.items()}
        tag = _local(el.tag)
        if tag == "p":
            a.pop("id", None)
        if tag == "fieldBegin" and a.get("zorder") == "-1":
            del a["zorder"]
            seen["zorder"] += 1
        if el in field_begins:
            for k in FIELD_VOLATILE:
                a.pop(k, None)
        return tuple(sorted(a.items()))

    def walk(el, in_t):
        tag = _local(el.tag)
        if el is drop:
            return
        if tag == "shadow" and el.get("type") == "NONE":
            seen["shadow"] += 1
            return
        start = len(tokens)
        tokens.append(("<", tag, attrs(el)))
        in_t = in_t or tag == "t"
        if el.text and (in_t or el.text.strip()):
            tokens.append(("txt", el.text))
        for child in el:
            walk(child, in_t)
            if child.tail and (in_t or child.tail.strip()):
                tokens.append(("txt", child.tail))
        tokens.append((">", tag))
        spans[el] = (start, len(tokens))

    walk(sec.root, False)
    ranges = []
    for i, m in enumerate(marks):
        if m[0] == "cell":  # subList 태그는 비교하고 안쪽만 가린다
            a, b = spans[m[1]]
            ranges.append((a + 1, b - 1, i))
        elif m[0] == "field":  # fieldBegin을 품은 ctrl 끝 ~ fieldEnd를 품은 ctrl 시작
            a, b = spans[m[2]][1], spans[m[3]][0]
            if a > b:
                raise EditTargetError("누름틀의 시작과 끝 순서가 뒤집혔습니다.")
            ranges.append((a, b, i))
    ranges.sort()
    if any(x[1] > y[0] for x, y in zip(ranges, ranges[1:])):
        raise EditTargetError("편집 대상 자리가 서로 겹칩니다.")
    out, regions, pos = [], {}, 0
    for a, b, i in ranges:
        out += tokens[pos:a]
        out.append(("mask", i))
        regions[i] = tokens[a:b]
        pos = b
    return out + tokens[pos:], regions, seen


def _text_of(tokens):
    return "".join(t[1] for t in tokens if t[0] == "txt")


def _skeleton(tokens, free):
    """free 태그와 글자를 뺀 구조 골격."""
    return [t for t in tokens if t[0] != "txt" and t[1] not in free]


def _marks(targets, number):
    """구역 number의 가림 지정. field는 (종류, fieldBegin 묶음, 시작 ctrl, 끝 ctrl)."""
    out = []
    for sec_no, spec, _ in targets:
        if sec_no != number:
            continue
        if spec[0] == "cell":
            out.append(("cell", spec[1]))
        elif spec[0] == "field":
            fbs = tuple(c for c in spec[1] if _local(c.tag) == "fieldBegin")
            out.append(("field", fbs, spec[1], spec[2]))
    return out


def compare_edit(src_parts, dst_parts, edits, rep):
    """section*.xml을 편집 대상만 가리고 위치별로 대조한다. 실패는 rep.fail, 허용분은 rep.allowed."""
    src, dst = _sections(src_parts), _sections(dst_parts)
    if [s.name for s in src] != [s.name for s in dst]:
        rep.fail(f"구역 파트가 다름: {[s.name for s in src]} → {[s.name for s in dst]}")
        return
    try:
        t_src, t_dst = _targets(src, edits), _targets(dst, edits)
    except EditTargetError as exc:
        rep.fail(f"편집 대상 확정 실패: {exc}")
        return
    if [(n, s[0]) for n, s, _ in t_src] != [(n, s[0]) for n, s, _ in t_dst]:
        rep.fail("편집 대상의 수·자리가 원본과 산출물에서 다름")
        return

    normalized = Counter()
    for s, d in zip(src, dst):
        ms, md = _marks(t_src, s.number), _marks(t_dst, d.number)
        images = [spec[1] for n, spec, _ in t_dst if n == d.number and spec[0] == "image"]
        try:
            a, a_regions, a_shadow = _tokens(s, ms)
            b, b_regions, b_shadow = _tokens(d, md)
            if images:  # 앵커 문단의 그림 하나를 뺐을 때 원본과 같아야 한다
                for pic in (el for p in images for el in p.iter() if _local(el.tag) == "pic"):
                    trial = _tokens(d, md, drop=pic)
                    if trial[0] == a:
                        b, b_regions, b_shadow = trial
                        break
                else:
                    rep.fail(f"{s.name}: 앵커 문단에 그림 하나만 더해진 모양이 아님 " + _first_diff(a, b))
                    continue
        except EditTargetError as exc:
            rep.fail(f"편집 대상 확정 실패: {exc}")
            return
        normalized.update(a_shadow)
        normalized.subtract(b_shadow)
        if a != b:
            rep.fail(f"{s.name}: 편집 대상 밖이 달라짐 " + _first_diff(a, b))
            continue
        ids = [[p.get("id") for p in x.root.iter() if _local(p.tag) == "p"] for x in (s, d)]
        renumbered = sum(i != j for i, j in zip(*ids))
        if renumbered:
            rep.allowed[(s.name, "p", (("id", "*"),))] = (renumbered, 0)
        edits_here = [e for n, spec, e in t_src if n == s.number and spec[0] in ("cell", "field")]
        for i, e in enumerate(edits_here):
            before, after = a_regions[i], b_regions[i]
            if isinstance(e, CellEdit):
                if _skeleton(before, FREE_IN_CELL) != _skeleton(after, FREE_IN_CELL):
                    rep.fail(f"{s.name}: 표 {e.table} 칸 ({e.row},{e.col}) 안의 글자 밖 요소가 달라짐 "
                             + _first_diff(_skeleton(before, FREE_IN_CELL), _skeleton(after, FREE_IN_CELL)))
                if _text_of(after) != e.text:
                    rep.fail(f"{s.name}: 표 {e.table} 칸 ({e.row},{e.col}) 글자가 요청값과 다름"
                             f"({_text_of(after)[:40]!r} ≠ {e.text[:40]!r})")
                if e.old_text is not None and \
                        re.sub(r"\s", "", _text_of(before)) != re.sub(r"\s", "", e.old_text):
                    rep.fail(f"{s.name}: 원본 표 {e.table} 칸 ({e.row},{e.col}) 글자가 rhwp가 읽은 옛 글자와 "
                             "다름(좌표 해석 불일치)")
            else:
                if _skeleton(before, FREE_IN_FIELD) != _skeleton(after, FREE_IN_FIELD):
                    rep.fail(f"{s.name}: 누름틀 {e.name!r} 값 자리의 글자 밖 요소가 달라짐")
                if _text_of(after) != e.value:
                    rep.fail(f"{s.name}: 누름틀 {e.name!r} 값이 요청값과 다름"
                             f"({_text_of(after)[:40]!r} ≠ {e.value[:40]!r})")
    for what, key in (("shadow", ("shadow", (("type", "NONE"),))),
                      ("zorder", ("fieldBegin zorder", (("zorder", "-1"),)))):
        n = normalized[what]
        if n > 0:
            rep.allowed[("Contents/section*.xml",) + key] = (n, 0)
        elif n < 0:
            rep.fail(f"{key[0]} 표시 {-n}개가 새로 생김")


def _first_diff(x, y):
    n = next((i for i, (p, q) in enumerate(zip(x, y)) if p != q), min(len(x), len(y)))

    def show(seq):
        return " ".join(_tok(t) for t in seq[max(n - 1, 0):n + 3])[:200]
    return f"(토큰 {n}: 원본 {show(x)} / 산출물 {show(y)})"


def _tok(t):
    if t[0] == "<":
        return "<" + t[1] + "".join(f' {k}="{v}"' for k, v in t[2][:3]) + ">"
    if t[0] == ">":
        return "</" + t[1] + ">"
    if t[0] == "txt":
        return repr(t[1][:20])
    return f"[편집 자리 {t[1]}]"


# ---------- BinData ----------

def bindata_digest(parts):
    return {n: hashlib.sha256(b).hexdigest() for n, b in bindata_parts(parts).items()}


# ---------- 텍스트 ----------

def chars_of(text):
    return Counter(ch for ch in text if ch.isalnum())


def markdown_chars(path):
    """hwpx-tomd로 변환한 마크다운의 글자 멀티셋. 패키지가 없으면 None."""
    try:
        from hwpx_tomd import convert
    except ImportError:
        return None
    return chars_of(convert(path).markdown)


# ---------- 스키마 ----------

def validate(path):
    """hwpx-validate 결과 (ok|None, 출력). 도구가 없으면 (None, 안내)."""
    exe = shutil.which("hwpx-validate")
    if not exe:
        return None, "hwpx-validate 없음 (pip install python-hwpx)"
    r = subprocess.run([exe, str(path)], capture_output=True, text=True, timeout=120)
    out = (r.stdout + r.stderr).strip()
    return r.returncode == 0 and "Error" not in out and "error" not in out.lower().split("warning")[0], out


# ---------- 게이트 ----------

def gate(original, produced, *, allow=None, allow_bindata_added=0,
         expect_text_change=False, require_census=True, require_bindata=True,
         allow_rhwp=False, edits=None):
    """원본 HWPX 대비 산출 HWPX의 손실을 검사한다.

    allow: callable((tag, attrs)) -> bool. True면 그 census 변화는 의도된 것으로 본다.
    allow_bindata_added: 새로 들어와도 되는 BinData 파일 수(그림 삽입은 1).
    expect_text_change: 치환·채우기처럼 글자가 달라지는 편집이면 True(텍스트 대조 생략).
    allow_rhwp: rhwp 재직렬화로 알려진 손실(split_rhwp_known)을 실패 대신 allowed에 담는다.
    edits: rhwp 편집 대상 목록(CellEdit·FieldEdit·ImageInsert). 주면 section*.xml은 census 대신
           compare_edit로 위치별 대조하고(글자 포함), 마크다운 글자 대조는 그 대조가 대신한다.
           rhwp 편집 산출물은 allow·expect_text_change 없이 이것만 쓴다.
    """
    rep = GateReport()
    src = read_parts(original)
    dst = read_parts(produced)

    if edits is not None:
        compare_edit(src, dst, edits, rep)

    if require_census:
        used = used_parapr_ids(src) if allow_rhwp else None
        if allow_rhwp and used is None:
            rep.messages.append("본문 파트를 읽지 못해 paraPr 사용 여부를 모름(header 허용 목록 미적용)")
        for name in sorted(set(xml_parts(src)) | set(xml_parts(dst))):
            if name not in dst:
                rep.fail(f"{name}이 산출물에 없음")
                continue
            if name not in src:
                rep.gained[(name, ())] = (0, 1)
                continue
            if edits is not None and section_number(name) is not None:
                continue  # compare_edit가 위치별로 봤다
            if allow_rhwp:
                diff, known = split_rhwp_known(name, src[name], dst[name], used)
                rep.allowed.update({(name,) + k: v for k, v in known.items()})
            else:
                diff = compare_census(element_census(src[name]), element_census(dst[name]))
            for key, (a, b) in diff.items():
                if allow and allow(key):
                    continue
                if b < a:
                    rep.lost[(name,) + key] = (a, b)
                else:
                    rep.gained[(name,) + key] = (a, b)
        if rep.lost:
            tags = Counter(k[1] for k in rep.lost)
            rep.fail("사라진 요소 " + ", ".join(f"{t}×{n}" for t, n in tags.most_common(8)))
    elif edits is not None:
        rep.fail("편집 대조(edits)는 census 검사와 함께만 쓴다")

    if require_bindata:
        before, after = bindata_digest(src), bindata_digest(dst)
        rep.bindata_removed = sorted(set(before) - set(after))
        rep.bindata_added = sorted(set(after) - set(before))
        rep.bindata_changed = sorted(n for n in set(before) & set(after) if before[n] != after[n])
        if rep.bindata_removed or rep.bindata_changed:
            rep.fail(f"BinData 변형: 삭제 {rep.bindata_removed} 변경 {rep.bindata_changed}")
        if len(rep.bindata_added) > allow_bindata_added:
            rep.fail(f"BinData 추가가 허용({allow_bindata_added})을 넘음: {rep.bindata_added}")

    if not expect_text_change and edits is None:
        a, b = markdown_chars(original), markdown_chars(produced)
        if a is None or b is None:
            rep.unverified("hwpx-tomd 없음, 텍스트 대조 생략")
        else:
            rep.text_missing = a - b
            rep.text_extra = b - a
            if rep.text_missing:
                rep.fail(f"글자 {sum(rep.text_missing.values())}개 누락: "
                         + "".join(sorted(rep.text_missing))[:40])

    ok, out = validate(produced)
    rep.validate_ok = ok
    if ok is None:
        rep.unverified(out)
    elif not ok:
        rep.fail("hwpx-validate 실패: " + out.splitlines()[-1][:160])

    if rep.status == "passed":
        rep.messages.append("요소·BinData·글자·스키마 이상 없음")
    return rep


def cross_check(a_path, b_path):
    """같은 원본에서 다른 엔진이 만든 두 HWPX를 대조한다(HWP 변환의 독립 기준).

    글자 멀티셋, 표·셀·그림·차트·OLE·수식 수, BinData와 Chart/ 파트의 수·내용 해시를 본다.
    어느 쪽이 맞는지는 말하지 않고 다른 점만 보고한다. HWP 입구 변환과 --to-hwp 왕복이
    이 함수 하나로 객체 보존을 판정한다.
    """
    rep = GateReport()
    pa, pb = read_parts(a_path), read_parts(b_path)

    def count(parts, tag):
        return sum(tag_census(b).get(tag, 0) for b in xml_parts(parts).values())

    for tag, label in (("tbl", "표"), ("tc", "셀"), ("pic", "그림"),
                       ("chart", "chart"), ("ole", "ole"), ("equation", "equation")):
        ca, cb = count(pa, tag), count(pb, tag)
        if ca != cb:
            rep.fail(f"{label} 수 불일치 {ca} vs {cb}" if label != tag else f"{tag} {ca}→{cb}")
    # 파일 이름은 엔진마다 달리 매길 수 있어 내용 해시의 멀티셋으로 본다.
    # 실측(2026-10-28, HWP 3건·BinData 170개): rhwp·JAR 변환물의 바이트가 모두 같다.
    for prefix, label in (("BinData/", "BinData"), ("Chart/", "Chart/ 파트")):
        da = [hashlib.sha256(b).hexdigest() for n, b in pa.items() if n.startswith(prefix)]
        db = [hashlib.sha256(b).hexdigest() for n, b in pb.items() if n.startswith(prefix)]
        if len(da) != len(db):
            rep.fail(f"{label} 수 불일치 {len(da)} vs {len(db)}")
        ha, hb = Counter(da), Counter(db)
        if ha != hb:
            rep.fail(f"{label} 해시 불일치 {max(sum((ha - hb).values()), sum((hb - ha).values()))}건")
    ta, tb = markdown_chars(a_path), markdown_chars(b_path)
    if ta is None or tb is None:
        rep.unverified("hwpx-tomd 없음, 텍스트 대조 생략")
    else:
        rep.text_missing, rep.text_extra = ta - tb, tb - ta
        if rep.text_missing or rep.text_extra:
            rep.fail(f"글자 불일치: 한쪽에만 {sum(rep.text_missing.values())}·"
                     f"{sum(rep.text_extra.values())}개")
    if rep.status == "passed":
        rep.messages.append("두 엔진의 글자·표·셀·그림·차트·OLE·수식 수와 BinData·Chart 내용 일치")
    return rep


def describe_lost(rep, limit=10):
    """사람이 읽을 손실 목록. (파일, 태그, 속성) 순."""
    lines = []
    for (name, tag, attrs), (a, b) in sorted(rep.lost.items())[:limit]:
        attr_txt = " ".join(f'{k}="{v}"' for k, v in attrs[:3])
        lines.append(f"  {name}: <{tag} {attr_txt}> {a}→{b}")
    if len(rep.lost) > limit:
        lines.append(f"  … 외 {len(rep.lost) - limit}건")
    if rep.allowed:
        lines.append("  " + rep.allowed_line())
    return "\n".join(lines)
