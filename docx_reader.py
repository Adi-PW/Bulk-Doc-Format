"""
Linearise a .docx into an ordered list of blocks the parser can reason about.

Why not just ``paragraph.text``?  Because exam docs rely on things python-docx
does not surface on its own:

* Word auto-numbering ("1.", "(a)", "I.") lives in numbering.xml, not in the text.
* Soft line breaks (Shift+Enter) pack a question, its options and the answer
  into a single paragraph.
* Runs inside hyperlinks / tracked insertions are skipped by ``paragraph.runs``.

Every block is either a ``Line`` (one visual line of text with formatted runs)
or a ``Table`` (grid of cells, each cell holding Lines).
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.oxml.ns import qn

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V_NS = "urn:schemas-microsoft-com:vml"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

_ZERO_WIDTH = dict.fromkeys(map(ord, "\u200b\u200c\u200d\ufeff\u2060\u00ad"), None)


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #
@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    superscript: bool = False
    subscript: bool = False
    image: Optional[bytes] = None       # raw image bytes when the run is a picture
    image_ext: str = ""
    image_cx: int = 0                   # EMU
    image_cy: int = 0

    def same_format(self, other: "Run") -> bool:
        return (self.image is None and other.image is None and
                (self.bold, self.italic, self.underline, self.superscript, self.subscript) ==
                (other.bold, other.italic, other.underline, other.superscript, other.subscript))


@dataclass
class Line:
    runs: List[Run] = field(default_factory=list)
    auto_number: str = ""               # rendered Word numbering, e.g. "1." / "(a)" / "•"
    kind: str = "line"
    heading: bool = False               # paragraph uses a Word Heading/Title style

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs if r.image is None)

    @property
    def has_image(self) -> bool:
        return any(r.image is not None for r in self.runs)

    def is_blank(self) -> bool:
        return not self.text.strip() and not self.has_image


@dataclass
class Cell:
    lines: List[Line] = field(default_factory=list)
    grid_span: int = 1
    v_merge: str = ""                   # "", "restart" or "continue"

    @property
    def text(self) -> str:
        return "\n".join(l.text for l in self.lines)


@dataclass
class Table:
    rows: List[List[Cell]] = field(default_factory=list)
    kind: str = "table"

    @property
    def text(self) -> str:
        return "\n".join(" | ".join(c.text for c in row) for row in self.rows)

    def is_blank(self) -> bool:
        return all(not c.text.strip() and not any(l.has_image for l in c.lines)
                   for row in self.rows for c in row)

    @property
    def n_cols(self) -> int:
        return max((sum(c.grid_span for c in row) for row in self.rows), default=0)

    def flatten(self) -> List[Line]:
        """Row-major list of every non-blank line in the table."""
        out: List[Line] = []
        for row in self.rows:
            for c in row:
                if c.v_merge == "continue":
                    continue
                out.extend(l for l in c.lines if not l.is_blank())
        return out


Block = object  # Line | Table


# --------------------------------------------------------------------------- #
# Numbering
# --------------------------------------------------------------------------- #
def _to_roman(n: int) -> str:
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out or "0"


def _to_letters(n: int) -> str:
    # Word style: a..z, aa..zz, aaa..
    if n <= 0:
        return "0"
    letter = chr(ord("a") + (n - 1) % 26)
    return letter * ((n - 1) // 26 + 1)


def _fmt_number(n: int, fmt: str) -> str:
    if fmt in ("lowerLetter",):
        return _to_letters(n)
    if fmt in ("upperLetter",):
        return _to_letters(n).upper()
    if fmt == "lowerRoman":
        return _to_roman(n).lower()
    if fmt == "upperRoman":
        return _to_roman(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "none":
        return ""
    return str(n)  # decimal + anything exotic


@dataclass
class _Level:
    start: int = 1
    fmt: str = "decimal"
    text: str = "%1."
    restart: Optional[int] = None
    is_lgl: bool = False


class Numbering:
    """Minimal but faithful renderer of Word list numbering."""

    def __init__(self, doc):
        self.abstract: Dict[str, Dict[int, _Level]] = {}
        self.num_to_abs: Dict[str, str] = {}
        self.num_overrides: Dict[str, Dict[int, Tuple[Optional[int], Optional[_Level]]]] = {}
        self.counters: Dict[str, Dict[int, int]] = {}
        self.started: Dict[str, set] = {}
        self.style_numpr: Dict[str, Tuple[Optional[str], Optional[int]]] = {}
        self.style_based: Dict[str, str] = {}
        try:
            part = doc.part.numbering_part
            root = part.element
        except Exception:
            root = None
        if root is not None:
            for an in root.findall(qn("w:abstractNum")):
                aid = an.get(qn("w:abstractNumId"))
                self.abstract[aid] = {int(l.get(qn("w:ilvl"), "0")): self._parse_lvl(l)
                                      for l in an.findall(qn("w:lvl"))}
                # numStyleLink indirection is rare in exam docs; ignored safely.
            for num in root.findall(qn("w:num")):
                nid = num.get(qn("w:numId"))
                an = num.find(qn("w:abstractNumId"))
                self.num_to_abs[nid] = an.get(qn("w:val")) if an is not None else ""
                ov: Dict[int, Tuple[Optional[int], Optional[_Level]]] = {}
                for lo in num.findall(qn("w:lvlOverride")):
                    ilvl = int(lo.get(qn("w:ilvl"), "0"))
                    so = lo.find(qn("w:startOverride"))
                    lv = lo.find(qn("w:lvl"))
                    ov[ilvl] = (int(so.get(qn("w:val"))) if so is not None else None,
                                self._parse_lvl(lv) if lv is not None else None)
                self.num_overrides[nid] = ov
        # paragraph styles can carry numbering too
        try:
            sroot = doc.styles.element
            for st in sroot.findall(qn("w:style")):
                sid = st.get(qn("w:styleId"))
                based = st.find(qn("w:basedOn"))
                if based is not None:
                    self.style_based[sid] = based.get(qn("w:val"))
                ppr = st.find(qn("w:pPr"))
                if ppr is not None:
                    np_ = ppr.find(qn("w:numPr"))
                    if np_ is not None:
                        nid = np_.find(qn("w:numId"))
                        il = np_.find(qn("w:ilvl"))
                        self.style_numpr[sid] = (nid.get(qn("w:val")) if nid is not None else None,
                                                 int(il.get(qn("w:val"))) if il is not None else None)
        except Exception:
            pass

    @staticmethod
    def _parse_lvl(l) -> _Level:
        def val(tag, default=None):
            el = l.find(qn(tag))
            return el.get(qn("w:val")) if el is not None else default
        lv = _Level()
        lv.start = int(val("w:start", "1") or 1)
        lv.fmt = val("w:numFmt", "decimal") or "decimal"
        lv.text = val("w:lvlText", "") if l.find(qn("w:lvlText")) is not None else ""
        r = val("w:lvlRestart")
        lv.restart = int(r) if r is not None else None
        lv.is_lgl = l.find(qn("w:isLgl")) is not None
        return lv

    def _level(self, nid: str, ilvl: int) -> Optional[_Level]:
        ov = self.num_overrides.get(nid, {}).get(ilvl)
        if ov and ov[1] is not None:
            return ov[1]
        return self.abstract.get(self.num_to_abs.get(nid, ""), {}).get(ilvl)

    def _numpr_for(self, p) -> Tuple[Optional[str], int]:
        ppr = p.find(qn("w:pPr"))
        nid, ilvl = None, None
        if ppr is not None:
            np_ = ppr.find(qn("w:numPr"))
            if np_ is not None:
                n = np_.find(qn("w:numId"))
                i = np_.find(qn("w:ilvl"))
                nid = n.get(qn("w:val")) if n is not None else None
                ilvl = int(i.get(qn("w:val"))) if i is not None else None
            if nid is None:
                ps = ppr.find(qn("w:pStyle"))
                sid = ps.get(qn("w:val")) if ps is not None else None
                seen = set()
                while sid and sid not in seen:
                    seen.add(sid)
                    if sid in self.style_numpr:
                        snid, silvl = self.style_numpr[sid]
                        nid = snid
                        if ilvl is None:
                            ilvl = silvl
                        break
                    sid = self.style_based.get(sid)
        return nid, (ilvl or 0)

    def render(self, p) -> str:
        nid, ilvl = self._numpr_for(p)
        if not nid or nid == "0" or nid not in self.num_to_abs:
            return ""
        lvl = self._level(nid, ilvl)
        if lvl is None:
            return ""
        # Counters are kept per list instance (numId), which is how Word and
        # LibreOffice render these docs (verified against LibreOffice output).
        ctr = self.counters.setdefault(nid, {})
        started = self.started.setdefault(nid, set())
        if ilvl not in started:
            so = self.num_overrides.get(nid, {}).get(ilvl, (None, None))[0]
            ctr[ilvl] = (so if so is not None else lvl.start) - 1
            started.add(ilvl)
        ctr[ilvl] = ctr.get(ilvl, lvl.start - 1) + 1
        # reset deeper levels
        for deeper in list(ctr.keys()):
            if deeper > ilvl:
                dl = self._level(nid, deeper)
                if dl is None or dl.restart is None or dl.restart > ilvl:
                    ctr.pop(deeper, None)
                    started.discard(deeper)
        if lvl.fmt == "bullet":
            return "•"
        text = lvl.text or ""

        def sub(m):
            k = int(m.group(1)) - 1
            kl = self._level(nid, k) or _Level()
            n = ctr.get(k, kl.start)
            fmt = "decimal" if (lvl.is_lgl and kl.fmt != "decimal") else kl.fmt
            return _fmt_number(n, fmt)

        return re.sub(r"%(\d)", sub, text).strip()


# --------------------------------------------------------------------------- #
# Run-level formatting (direct + character style + paragraph style)
# --------------------------------------------------------------------------- #
class _Styles:
    def __init__(self, doc):
        self.rpr: Dict[str, Dict[str, bool]] = {}
        self.based: Dict[str, str] = {}
        self.heading_ids: set = set()
        try:
            for st in doc.styles.element.findall(qn("w:style")):
                sid = st.get(qn("w:styleId"))
                b = st.find(qn("w:basedOn"))
                if b is not None:
                    self.based[sid] = b.get(qn("w:val"))
                nm = st.find(qn("w:name"))
                name = (nm.get(qn("w:val")) if nm is not None else "") or ""
                ppr = st.find(qn("w:pPr"))
                outline = ppr.find(qn("w:outlineLvl")) if ppr is not None else None
                if re.match(r"^(heading\s*\d|title|subtitle)$", name.strip(), re.I) or (
                        outline is not None and outline.get(qn("w:val"), "9") not in ("9",)):
                    self.heading_ids.add(sid)
                rpr = st.find(qn("w:rPr"))
                if rpr is not None:
                    self.rpr[sid] = _read_rpr(rpr)
        except Exception:
            pass

    def is_heading(self, sid: Optional[str]) -> bool:
        seen = set()
        while sid and sid not in seen:
            if sid in self.heading_ids:
                return True
            seen.add(sid)
            sid = self.based.get(sid)
        return False

    def resolve(self, sid: Optional[str]) -> Dict[str, bool]:
        chain, seen = [], set()
        while sid and sid not in seen:
            seen.add(sid)
            chain.append(sid)
            sid = self.based.get(sid)
        out: Dict[str, bool] = {}
        for s in reversed(chain):
            out.update(self.rpr.get(s, {}))
        return out


def _on(el) -> Optional[bool]:
    if el is None:
        return None
    v = el.get(qn("w:val"))
    return v not in ("0", "false", "off", "none")


def _read_rpr(rpr) -> Dict[str, bool]:
    out: Dict[str, bool] = {}
    for key, tag in (("bold", "w:b"), ("italic", "w:i")):
        v = _on(rpr.find(qn(tag)))
        if v is not None:
            out[key] = v
    u = rpr.find(qn("w:u"))
    if u is not None:
        out["underline"] = (u.get(qn("w:val")) or "single") not in ("none", "0")
    va = rpr.find(qn("w:vertAlign"))
    if va is not None:
        out["superscript"] = va.get(qn("w:val")) == "superscript"
        out["subscript"] = va.get(qn("w:val")) == "subscript"
    return out


# --------------------------------------------------------------------------- #
# Reader
# --------------------------------------------------------------------------- #
class DocxReader:
    def __init__(self, data: bytes):
        self.doc = Document(io.BytesIO(data))
        self.numbering = Numbering(self.doc)
        self.styles = _Styles(self.doc)
        self._img_cache: Dict[str, Tuple[bytes, str]] = {}

    # ---- public ---------------------------------------------------------- #
    def blocks(self) -> List[Block]:
        body = self.doc.element.body
        out: List[Block] = []
        self._walk_container(body, out)
        return out

    # ---- traversal ------------------------------------------------------- #
    def _walk_container(self, container, out: List[Block]):
        for child in container.iterchildren():
            tag = child.tag
            if tag == qn("w:p"):
                out.extend(self._paragraph_lines(child))
            elif tag == qn("w:tbl"):
                out.append(self._table(child))
            elif tag == qn("w:sdt"):
                content = child.find(qn("w:sdtContent"))
                if content is not None:
                    self._walk_container(content, out)
            elif tag == f"{{{MC_NS}}}AlternateContent":
                choice = child.find(f"{{{MC_NS}}}Choice")
                if choice is not None:
                    self._walk_container(choice, out)

    def _table(self, tbl) -> Table:
        t = Table()
        for tr in tbl.findall(qn("w:tr")):
            row: List[Cell] = []
            for tc in tr.findall(qn("w:tc")):
                cell = Cell()
                tcpr = tc.find(qn("w:tcPr"))
                if tcpr is not None:
                    gs = tcpr.find(qn("w:gridSpan"))
                    if gs is not None:
                        cell.grid_span = int(gs.get(qn("w:val"), "1"))
                    vm = tcpr.find(qn("w:vMerge"))
                    if vm is not None:
                        cell.v_merge = vm.get(qn("w:val")) or "continue"
                blocks: List[Block] = []
                self._walk_container(tc, blocks)
                for b in blocks:
                    if isinstance(b, Line):
                        cell.lines.append(b)
                    elif isinstance(b, Table):        # nested table -> flatten
                        cell.lines.extend(b.flatten())
                row.append(cell)
            t.rows.append(row)
        return t

    def _paragraph_lines(self, p) -> List[Line]:
        num = self.numbering.render(p)
        ppr = p.find(qn("w:pPr"))
        pstyle = None
        if ppr is not None:
            ps = ppr.find(qn("w:pStyle"))
            pstyle = ps.get(qn("w:val")) if ps is not None else None
        base = self.styles.resolve(pstyle)
        is_heading = self.styles.is_heading(pstyle)
        if ppr is not None and ppr.find(qn("w:outlineLvl")) is not None:
            is_heading = ppr.find(qn("w:outlineLvl")).get(qn("w:val"), "9") != "9"

        lines: List[Line] = [Line()]
        for r in self._iter_runs(p):
            fmt = dict(base)
            rpr = r.find(qn("w:rPr"))
            if rpr is not None:
                rs = rpr.find(qn("w:rStyle"))
                if rs is not None:
                    fmt.update(self.styles.resolve(rs.get(qn("w:val"))))
                fmt.update(_read_rpr(rpr))
            for child in r.iterchildren():
                tag = child.tag
                if tag == qn("w:t"):
                    self._append(lines[-1], child.text or "", fmt)
                elif tag == qn("w:tab"):
                    self._append(lines[-1], "\t", fmt)
                elif tag in (qn("w:br"), qn("w:cr")):
                    lines.append(Line())
                elif tag == qn("w:noBreakHyphen"):
                    self._append(lines[-1], "-", fmt)
                elif tag == qn("w:sym"):
                    ch = child.get(qn("w:char"))
                    try:
                        c = chr(int(ch, 16))
                        if 0xF000 <= ord(c) <= 0xF0FF:
                            c = chr(ord(c) - 0xF000)
                        self._append(lines[-1], c, fmt)
                    except Exception:
                        pass
                elif tag in (qn("w:drawing"), qn("w:pict")) or tag == f"{{{MC_NS}}}AlternateContent":
                    img = self._image_from(child)
                    if img:
                        lines[-1].runs.append(img)
        # inline math (m:oMath) outside runs -> plain text
        for m in p.iter(f"{{{M_NS}}}oMath"):
            txt = "".join(t.text or "" for t in m.iter(f"{{{M_NS}}}t"))
            if txt:
                self._append(lines[-1], txt, {})
        for ln in lines:
            for run in ln.runs:
                if run.image is None:
                    run.text = run.text.translate(_ZERO_WIDTH).replace("\u00a0", " ")
        if is_heading:
            for ln in lines:
                ln.heading = True
        if num:
            first = next((l for l in lines if not l.is_blank()), lines[0])
            first.auto_number = num
            first.runs.insert(0, Run(text=num + " "))
        return lines

    def _iter_runs(self, p):
        """Document-order runs, skipping deletions, text boxes and fallbacks."""
        skip_tags = {qn("w:del"), qn("w:moveFrom"), qn("w:txbxContent"),
                     f"{{{MC_NS}}}Fallback", qn("w:pPr"), qn("w:rPr")}

        def walk(el):
            for child in el.iterchildren():
                if child.tag in skip_tags:
                    continue
                if child.tag == qn("w:r"):
                    yield child
                elif child.tag == f"{{{M_NS}}}oMath" or child.tag == f"{{{M_NS}}}oMathPara":
                    continue  # handled separately
                else:
                    yield from walk(child)

        yield from walk(p)

    @staticmethod
    def _append(line: Line, text: str, fmt: Dict[str, bool]):
        if not text:
            return
        run = Run(text=text, bold=fmt.get("bold", False), italic=fmt.get("italic", False),
                  underline=fmt.get("underline", False), superscript=fmt.get("superscript", False),
                  subscript=fmt.get("subscript", False))
        if line.runs and line.runs[-1].same_format(run):
            line.runs[-1].text += text
        else:
            line.runs.append(run)

    def _image_from(self, el) -> Optional[Run]:
        rid, cx, cy = None, 0, 0
        blip = next(el.iter(f"{{{A_NS}}}blip"), None)
        if blip is not None:
            rid = blip.get(f"{{{R_NS}}}embed") or blip.get(f"{{{R_NS}}}link")
            ext = next(el.iter(f"{{{WP_NS}}}extent"), None)
            if ext is not None:
                cx, cy = int(ext.get("cx", "0")), int(ext.get("cy", "0"))
        else:
            imd = next(el.iter(f"{{{V_NS}}}imagedata"), None)
            if imd is not None:
                rid = imd.get(f"{{{R_NS}}}id")
        if not rid:
            return None
        try:
            if rid not in self._img_cache:
                part = self.doc.part.related_parts[rid]
                ct = getattr(part, "content_type", "") or ""
                ext = ct.split("/")[-1].replace("jpeg", "jpg").replace("x-emf", "emf").replace("x-wmf", "wmf")
                self._img_cache[rid] = (part.blob, ext)
            blob, ext = self._img_cache[rid]
        except Exception:
            return None
        if ext not in ("png", "jpg", "gif", "bmp", "tiff"):
            return None  # vector formats (emf/wmf) cannot be re-embedded reliably
        return Run(image=blob, image_ext=ext, image_cx=cx, image_cy=cy)


def read_blocks(data: bytes) -> List[Block]:
    return DocxReader(data).blocks()
