"""독립 손실 게이트: rhwp가 쓴 HWPX를 rhwp와 무관한 눈으로 검사한다.

rhwp는 문서를 자기 내부 구조(IR)로 읽어 전체를 다시 써서 저장한다. IR이 담지 않는
것은 조용히 사라지고, rhwp 자신의 `--verify`·`hwpx-roundtrip`은 그 손실을 보지 못한다
(실측 2026-10-28: 치환 1회에 header.xml의 hp:switch 23→20). 그래서 산출물은 원본과
독립적으로 대조한다.

검사 항목
- 요소 census: 모든 Contents/*.xml에서 (태그, 정렬된 속성 묶음) 단위로 개수를 세어
  사라지거나 바뀐 요소를 보고한다. 호출자가 의도한 변화는 allow로 제외한다.
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
    """본문 파트(header.xml 밖)의 문단이 참조하는 paraPr id 집합."""
    used = set()
    for name, data in xml_parts(parts).items():
        if name != "Contents/header.xml":
            used.update(m.decode() for m in re.findall(rb'paraPrIDRef="(\d+)"', data))
    return used


def split_rhwp_known(name, before, after, used_parapr):
    """(실패로 볼 census 차이, 허용할 손실)로 나눈다. 허용 손실은 lost와 같은 형식."""
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
         allow_rhwp=False):
    """원본 HWPX 대비 산출 HWPX의 손실을 검사한다.

    allow: callable((tag, attrs)) -> bool. True면 그 census 변화는 의도된 것으로 본다.
    allow_bindata_added: 새로 들어와도 되는 BinData 파일 수(그림 삽입은 1).
    expect_text_change: 치환·채우기처럼 글자가 달라지는 편집이면 True(텍스트 대조 생략).
    allow_rhwp: rhwp 재직렬화로 알려진 손실(split_rhwp_known)을 실패 대신 allowed에 담는다.
    """
    rep = GateReport()
    src = read_parts(original)
    dst = read_parts(produced)

    if require_census:
        used = used_parapr_ids(src) if allow_rhwp else None
        for name in sorted(set(xml_parts(src)) | set(xml_parts(dst))):
            if name not in dst:
                rep.fail(f"{name}이 산출물에 없음")
                continue
            if name not in src:
                rep.gained[(name, ())] = (0, 1)
                continue
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

    if require_bindata:
        before, after = bindata_digest(src), bindata_digest(dst)
        rep.bindata_removed = sorted(set(before) - set(after))
        rep.bindata_added = sorted(set(after) - set(before))
        rep.bindata_changed = sorted(n for n in set(before) & set(after) if before[n] != after[n])
        if rep.bindata_removed or rep.bindata_changed:
            rep.fail(f"BinData 변형: 삭제 {rep.bindata_removed} 변경 {rep.bindata_changed}")
        if len(rep.bindata_added) > allow_bindata_added:
            rep.fail(f"BinData 추가가 허용({allow_bindata_added})을 넘음: {rep.bindata_added}")

    if not expect_text_change:
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

    글자 멀티셋, 표(tbl) 수, 셀(tc) 수, BinData 수와 내용 해시를 본다. 어느 쪽이 맞는지는
    말하지 않고 다른 점만 보고한다.
    """
    rep = GateReport()
    pa, pb = read_parts(a_path), read_parts(b_path)

    def count(parts, tag):
        return sum(tag_census(b).get(tag, 0) for b in xml_parts(parts).values())

    for tag, label in (("tbl", "표"), ("tc", "셀"), ("pic", "그림")):
        ca, cb = count(pa, tag), count(pb, tag)
        if ca != cb:
            rep.fail(f"{label} 수 불일치 {ca} vs {cb}")
    na, nb = len(bindata_parts(pa)), len(bindata_parts(pb))
    if na != nb:
        rep.fail(f"BinData 수 불일치 {na} vs {nb}")
    # 파일 이름은 엔진마다 달리 매길 수 있어 내용 해시의 멀티셋으로 본다.
    # 실측(2026-10-28, HWP 3건·BinData 170개): rhwp·JAR 변환물의 바이트가 모두 같다.
    ha, hb = Counter(bindata_digest(pa).values()), Counter(bindata_digest(pb).values())
    if ha != hb:
        rep.fail(f"BinData 해시 불일치 {max(sum((ha - hb).values()), sum((hb - ha).values()))}건")
    ta, tb = markdown_chars(a_path), markdown_chars(b_path)
    if ta is None or tb is None:
        rep.unverified("hwpx-tomd 없음, 텍스트 대조 생략")
    else:
        rep.text_missing, rep.text_extra = ta - tb, tb - ta
        if rep.text_missing or rep.text_extra:
            rep.fail(f"글자 불일치: 한쪽에만 {sum(rep.text_missing.values())}·"
                     f"{sum(rep.text_extra.values())}개")
    if rep.status == "passed":
        rep.messages.append("두 엔진의 글자·표·셀·그림 수와 BinData 내용 일치")
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


def is_text_run_key(key):
    """치환·채우기에서 정상적으로 바뀌는 글자 런·줄 배치 요소인지."""
    return key[0] in {"t", "run", "lineseg", "linesegarray", "p"} or re.fullmatch(r"ctrl|secPr", key[0]) is not None
