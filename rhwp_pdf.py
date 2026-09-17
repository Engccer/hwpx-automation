"""rhwp PDF 내보내기: 한컴 COM이 없는 macOS·Linux에서 HWP/HWPX → PDF.

rhwp(https://github.com/edwardkim/rhwp)는 한컴 없이 HWP/HWPX를 조판·렌더링하는
오픈소스 엔진이다. 결과는 한컴 PDF와 쪽수·줄 나눔이 거의 같지만 글꼴은 이 컴퓨터에
설치된 글꼴로 대체된다. 대체 글꼴에도 없는 문자는 LastResort(빈 네모)로 찍히므로
저장 후 그 문자를 찾아 보고한다.
"""

import os
import re
import shutil
import subprocess
import tempfile
import zlib

INSTALL_HINT = ("rhwp CLI가 필요합니다. https://github.com/edwardkim/rhwp/releases 에서 "
                "플랫폼용 바이너리를 받아 SHA256SUMS.txt로 확인한 뒤 PATH에 두세요.")


def _objects(pdf):
    return {int(num): body for num, body in re.findall(rb"(?m)^(\d+) 0 obj\b(.*?)\bendobj", pdf, re.S)}


def _stream(body):
    m = re.search(rb"stream\r?\n(.*?)\r?\nendstream", body, re.S)
    if not m:
        return b""
    data = m.group(1)
    return zlib.decompress(data) if b"/FlateDecode" in body[:m.start()] else data


def _utf16(hexstr):
    return bytes.fromhex(hexstr.decode()).decode("utf-16-be", errors="replace")


def _cmap_chars(cmap):
    chars = set()
    for block in re.findall(rb"beginbfchar(.*?)endbfchar", cmap, re.S):
        for _, dst in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
            chars.update(_utf16(dst))  # 합자·ZWJ 시퀀스는 여러 글자로 매핑될 수 있다
    for block in re.findall(rb"beginbfrange(.*?)endbfrange", cmap, re.S):
        for lo, hi, dst in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
            base = ord(_utf16(dst)[-1])
            chars.update(chr(base + i) for i in range(int(hi, 16) - int(lo, 16) + 1))
    return chars


def missing_glyphs(pdf):
    """LastResort 글꼴로 그려진(=글꼴에 글리프가 없는) 문자를 정렬해 돌려준다."""
    objs = _objects(pdf)
    lastresort = [body for body in objs.values() if re.search(rb"/BaseFont\s*/\S*LastResort", body)]
    chars = set()
    # 문자 대응표(ToUnicode)는 상위 Type0 글꼴에만 있고 하위 CID 글꼴에는 없다.
    for body in lastresort:
        ref = re.search(rb"/ToUnicode\s+(\d+)\s+0\s+R", body)
        if ref:
            chars |= _cmap_chars(_stream(objs.get(int(ref.group(1)), b"")))
    return sorted(chars or ({"�"} if lastresort else set()))


def embedded_fonts(pdf):
    """PDF에 실제로 임베드된 글꼴 이름(서브셋 접두어·인코딩 접미어 제거)."""
    names = re.findall(rb"/(?:BaseFont|FontName)\s*/(?:[A-Z]{6}\+)?([^\s/<>\[\]()]+)", pdf)
    return sorted({n.decode("latin-1").removesuffix("-Identity-H") for n in names})


def save_pdf(src, output, password=None):
    """rhwp로 PDF를 만든다. 실패하면 기존 output을 건드리지 않는다.

    Returns: {"missing": 빈 네모로 찍힌 문자, "fonts": 임베드 글꼴, "overflow": 쪽 하단 넘침 경고 수}
    """
    rhwp = shutil.which("rhwp")
    if not rhwp:
        raise RuntimeError(INSTALL_HINT)
    output = os.path.abspath(output)
    os.makedirs(os.path.dirname(output), exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="rhwp-pdf-", dir=os.path.dirname(output)) as stage:
        staged = os.path.join(stage, "export.pdf")
        cmd = [rhwp, "export-pdf", os.path.abspath(src), "-o", staged]
        if password:
            cmd.append("--password-stdin")
        r = subprocess.run(cmd, input=f"{password}\n" if password else None,
                           capture_output=True, text=True)
        if r.returncode != 0 or not os.path.exists(staged):
            raise RuntimeError(f"rhwp PDF 변환 실패(exit {r.returncode}): {r.stderr.strip()[-500:]}")
        with open(staged, "rb") as fh:
            pdf = fh.read()
        if not pdf.startswith(b"%PDF-"):
            raise RuntimeError("rhwp 출력이 PDF 형식이 아닙니다.")
        # 분석이 실패해도 기존 output을 덮어쓰지 않도록 교체 전에 끝낸다.
        report = {"missing": missing_glyphs(pdf), "fonts": embedded_fonts(pdf),
                  "overflow": r.stderr.count("LAYOUT_OVERFLOW:")}
        os.replace(staged, output)
    return report
