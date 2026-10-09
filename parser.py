"""
Turns the linear block stream from ``docx_reader`` into structured questions.

Segmentation is a small state machine:

    pre  --(valid question start)-->  body  --(answer line)-->  solution
     ^                                                            |
     +------------------(next valid question start)---------------+

Inside a question, options are found by scanning *backwards* for the last
complete marker run  ... (d) (c) (b) (a).  Everything above option (a) is the
question body, so statements labelled (a)/(b)/(c) above the real options stay
in the body.

Bilingual layer (English + Hindi in one document)
-------------------------------------------------
A question body may hold two versions: question + options in one language,
then question (numbered or not) + options in the other.  The split point is
the first line written in the *other* script that comes right after a complete
option run.  The solution is split at a Hindi label (स्पष्टीकरण / व्याख्या) or,
failing that, at the first line in the other script.

Language is decided by script only: Devanagari -> Hindi, otherwise English.
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

try:  # package layout (qbg/...) or flat layout (all files next to app.py)
    from .docx_reader import Line, Run, Table, read_blocks
except ImportError:
    from docx_reader import Line, Run, Table, read_blocks

# --------------------------------------------------------------------------- #
# Patterns
# --------------------------------------------------------------------------- #
_DEV_DIGITS = str.maketrans("०१२३४५६७८९", "0123456789")
_HI_LETTERS = "कखगघङ"          # Hindi option letters (क)=a (ख)=b (ग)=c (घ)=d (ङ)=e

# Question start: "1.", "1)", "1:", "Q1.", "Q.1", "Question 1:", "प्रश्न 1.", "13 'Admission'"
QSTART_RE = re.compile(
    r"""^\s*[\[(]?\s*
        (?:(?:[Qq](?:ues(?:tion)?|UES(?:TION)?)?|प्रश्न)\s*(?:[Nn][Oo]\.?|सं\.?|संख्या)?\s*[.:\-]?\s*)?
        (?P<num>[0-9०-९]{1,3})(?![0-9०-९])
        (?:\s*[.):।]\s*|\s+(?=[A-Z'"\u2018\u201c(\[\u0900-\u097F]))
    """,
    re.X,
)

# Option marker: "(a)" "a)" "a." "(A)" "A)" "A." "(क)" "क)" and "©" (Word's autocorrect of "(c)")
_BULLET = r"(?:[•\u2022\u25aa\u25cf\u25e6\u2023\u2043*\-]\s*)?"   # "• (a) ..." in a bullet list
OPT_LETTER_RE = re.compile(
    r"^\s*" + _BULLET + r"(?:\(\s*(?P<l1>[a-hA-Hक-ङ]|©)\s*\)|(?P<l2>[a-hA-Hक-ङ])\s*[.)]|(?P<l3>©))\s*")
# Word list that kept counting from an earlier list: options shown as e. f. g. h. / i. j. k. l.
OPT_ANY_LETTER_RE = re.compile(r"^\s*" + _BULLET + r"\(?\s*(?P<l>[a-zA-Z])\s*[.)]\s*")
# Numeric options: "(1)" "1)" "1." (only used when no letter options exist)
OPT_NUM_RE = re.compile(r"^\s*" + _BULLET + r"(?:\(\s*(?P<n1>[1-8१-८])\s*\)|(?P<n2>[1-8१-८])\s*[.)])\s*")

ANSWER_RE = re.compile(
    r"""^\s*(?:
            (?:(?i:correct|right)\s+)?(?i:ans(?:wer)?|key)\s*(?:is\s*)?[.:\-\u2013\u2014=]*
          | (?i:correct\s+(?:option|choice))\s*(?:is\s*)?[.:\-\u2013\u2014=]*
          | (?:सही\s+)?उत्तर\s*[:\-\u2013\u2014=]+
        )\s*[:\-\u2013\u2014]?\s*
        (?:Option\s*|विकल्प\s*)?
        [\[(]?\s*(?P<ans>[a-hA-H1-8क-ङ१-८]|©)\s*[\])]?
        (?![A-Za-z0-9\u0900-\u097F])
        [\s.):\-\u2013\u2014]*
    """,
    re.X,
)

_HI_EXPL = r"(?:स्पष्टीकरण|व्याख्या|हल|समाधान)\s*(?:[:\-\u2013\u2014.]+|$)"
EXPL_RE = re.compile(
    r"^\s*(?:(?:Explanation|Explaination|Explanations|Expl\.?|Solution|Sol\.?|Rationale)"
    r"\s*[:\-\u2013\u2014.]*|" + _HI_EXPL + r")\s*",
    re.I,
)
HI_EXPL_RE = re.compile(r"^\s*" + _HI_EXPL + r"\s*")
EN_EXPL_RE = re.compile(
    r"^\s*(?:Explanation|Explaination|Explanations|Expl\.?|Solution|Sol\.?|Rationale)\b", re.I)

# an explanation label somewhere inside the answer line ("Ans. (b) Zero FIR Explanation: ...")
INLINE_EXPL_RE = re.compile(
    r"(?:^|(?<=\s))(?:Explanation|Explaination|Solution|स्पष्टीकरण|व्याख्या)\s*[:\-\u2013\u2014]", re.I)

OPTIONS_LABEL_RE = re.compile(
    r"^\s*(?:Options?|Choices?|Alternatives?|विकल्प)\s*[:\-\u2013\u2014]?\s*$", re.I)

# lines that look like statement / list items (used for unnumbered fallback)
LISTISH_RE = re.compile(
    r"^\s*(?:\(?[a-zA-Z]\)|[a-zA-Z]\.|\(?[ivxIVX]{1,5}[.)]|\(?\d{1,2}[.)]|[•\-\u2022])\s+"
)

HEADING_RE = re.compile(r"^[A-Z0-9 &,.'()\-/]{3,60}$")

AUTO_EN = "The correct answer is "
AUTO_HI = "सही उत्तर है: "
MIN_SECTION = 3     # a run of >= 3 single-language questions = a language section


def to_index(ch: str) -> int:
    """Answer/option label -> 0-based index ('b' / 'B' / 'ख' / '2' / '२' -> 1, '©' -> 2)."""
    ch = ch.translate(_DEV_DIGITS)
    if ch == "©":
        return 2
    if ch in _HI_LETTERS:
        return _HI_LETTERS.index(ch)
    if ch.isdigit():
        return int(ch) - 1
    return ord(ch.lower()) - ord("a")


def script(text: str) -> str:
    """Language of one line: 'hi' if it has real Devanagari content (>= 3 letters
    and >= 20% of all letters -- Hindi lines often quote English terms such as
    'Falsus in uno'), 'en' if it has Latin letters, '' if it has neither."""
    dev = sum(1 for c in text if "\u0900" <= c <= "\u0963" or "\u0971" <= c <= "\u097F")
    lat = sum(1 for c in text if ("a" <= c <= "z") or ("A" <= c <= "Z"))
    if dev >= 3 and dev >= 0.2 * (dev + lat):
        return "hi"
    if lat:
        return "en"
    return "hi" if dev else ""


def letters_by_script(items) -> Tuple[int, int]:
    dev = lat = 0
    for it in items:
        t = it.text
        dev += sum(1 for c in t if "\u0900" <= c <= "\u0963" or "\u0971" <= c <= "\u097F")
        lat += sum(1 for c in t if ("a" <= c <= "z") or ("A" <= c <= "Z"))
    return dev, lat


def strip_prefix(line: Line, n_chars: int) -> Line:
    """Return a copy of ``line`` with the first ``n_chars`` visible chars removed."""
    new = Line(runs=[], auto_number=line.auto_number)
    remaining = n_chars
    for r in line.runs:
        if r.image is not None:
            new.runs.append(r)
            continue
        if remaining <= 0:
            new.runs.append(copy.copy(r))
            continue
        if len(r.text) <= remaining:
            remaining -= len(r.text)
            continue
        nr = copy.copy(r)
        nr.text = r.text[remaining:]
        remaining = 0
        new.runs.append(nr)
    return lstrip_line(new)


def lstrip_line(line: Line, chars: str = " \t") -> Line:
    while line.runs and line.runs[0].image is None:
        t = line.runs[0].text.lstrip(chars)
        if t:
            line.runs[0] = copy.copy(line.runs[0])
            line.runs[0].text = t
            break
        line.runs.pop(0)
    return line


def rstrip_line(line: Line) -> Line:
    while line.runs and line.runs[-1].image is None:
        t = line.runs[-1].text.rstrip()
        if t:
            line.runs[-1] = copy.copy(line.runs[-1])
            line.runs[-1].text = t
            break
        line.runs.pop()
    return line


def clean_line(line: Line) -> Line:
    """Collapse runs of whitespace/tabs to single spaces and trim."""
    for r in line.runs:
        if r.image is None:
            t = re.sub(r" {3,}", "\t", r.text)
            t = re.sub(r" *\t[ \t]*", "\t", t)
            r.text = re.sub(r" {2,}", " ", t)
    return rstrip_line(lstrip_line(line))


# --------------------------------------------------------------------------- #
# Result model
# --------------------------------------------------------------------------- #
@dataclass
class Question:
    source_number: Optional[int]
    language: str = "en"                                      # en | hi | bi
    body: List[object] = field(default_factory=list)          # English (or the only) version
    options: List[List[Line]] = field(default_factory=list)
    solution: List[object] = field(default_factory=list)
    hi_body: List[object] = field(default_factory=list)       # Hindi version (bilingual only)
    hi_options: List[List[Line]] = field(default_factory=list)
    hi_solution: List[object] = field(default_factory=list)
    correct_index: Optional[int] = None
    answer_raw: str = ""
    auto_solution: List[str] = field(default_factory=list)    # versions whose solution was auto-written
    isolated: bool = False                                    # single-language q inside a bilingual paper
    ok: bool = True
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def label(self) -> str:
        return f"Q.{self.source_number}" if self.source_number is not None else "Q.(unnumbered)"

    @property
    def body_text(self) -> str:
        return "\n".join(b.text for b in (self.body or self.hi_body))

    @property
    def language_name(self) -> str:
        return {"en": "English", "hi": "Hindi", "bi": "English + Hindi"}[self.language]

    @property
    def status(self) -> str:
        if not self.ok:
            return "Failure"
        return "Success - Solution Written Automatically" if self.auto_solution else "Success"

    def remarks(self) -> str:
        msgs = list(self.errors) + list(self.warnings)
        if self.ok and self.auto_solution:
            if self.language == "bi" and len(self.auto_solution) == 1:
                msgs.append(f"{self.auto_solution[0]} solution not in source; written automatically")
            else:
                msgs.append("solution not in source; written automatically")
        return f"Source {self.label}: " + "; ".join(msgs) if msgs else ""


@dataclass
class _Seg:
    number: Optional[int]
    body: List[object] = field(default_factory=list)
    answer_line: Optional[Line] = None
    answer: Optional[str] = None
    answer_tail: Optional[Line] = None
    solution: List[object] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


@dataclass
class _Part:
    stem: List[object] = field(default_factory=list)
    options: List[List[Line]] = field(default_factory=list)
    numeric: bool = False
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Parser
# --------------------------------------------------------------------------- #
class QuestionParser:
    LOOKAHEAD = 600
    START_CANDIDATES = 30

    def __init__(self, blocks: List[object]):
        # Drop blank lines / empty tables up front: they carry no meaning.
        self.items = [b for b in blocks if not b.is_blank()]
        self.duplicate_keys = 0
        self.last_shifted = False

    # ---- line classifiers ----------------------------------------------- #
    @staticmethod
    def qnum(item) -> Optional[Tuple[int, int]]:
        if not isinstance(item, Line):
            return None
        m = QSTART_RE.match(item.text)
        if not m:
            return None
        return int(m.group("num").translate(_DEV_DIGITS)), m.end()

    @staticmethod
    def answer(item) -> Optional[Tuple[str, int]]:
        if not isinstance(item, Line):
            return None
        m = ANSWER_RE.match(item.text)
        if not m:
            return None
        return m.group("ans"), m.end()

    @staticmethod
    def opt_marker(item, numeric: bool = False, shifted: bool = False) -> Optional[Tuple[int, int]]:
        if not isinstance(item, Line):
            return None
        if shifted:     # only for Word auto-lettered lines (typed "i." may be a roman numeral)
            if not item.auto_number:
                return None
            m = OPT_ANY_LETTER_RE.match(item.text)
            return (ord(m.group("l").lower()) - ord("a"), m.end()) if m else None
        if numeric:
            m = OPT_NUM_RE.match(item.text)
            if not m:
                return None
            return to_index(m.group("n1") or m.group("n2")), m.end()
        m = OPT_LETTER_RE.match(item.text)
        if not m:
            return None
        return to_index(m.group("l1") or m.group("l2") or m.group("l3")), m.end()

    # ---- start validation ----------------------------------------------- #
    def _next_number_is(self, j: int, m: int) -> bool:
        """Is the next numbered line after j (before the next answer key) numbered m?"""
        for k in range(j + 1, min(len(self.items), j + self.LOOKAHEAD)):
            it = self.items[k]
            if self.answer(it):
                return False
            q = self.qnum(it)
            if q:
                return q[0] == m
        return False

    def _valid_start(self, i: int, n: int) -> bool:
        """A numbered line is a real question start only if an answer key follows
        before another line with the *same* number appears -- unless that later
        line opens a fresh list (n, n+1, ...: statements or options numbered
        inside the question, e.g. Q.1 whose options are 1. 2. 3. 4.) or is the
        other-language copy of the question, which follows its options."""
        saw_prev = False
        opts_seen = 0
        lang = script(self.items[i].text) if isinstance(self.items[i], Line) else ""
        for j in range(i + 1, min(len(self.items), i + self.LOOKAHEAD)):
            it = self.items[j]
            if self.answer(it):
                return True
            if self.opt_marker(it) or self.opt_marker(it, numeric=True):
                opts_seen += 1
            q = self.qnum(it)
            if q:
                if q[0] == n - 1:
                    saw_prev = True
                elif q[0] == n and not saw_prev and not self._next_number_is(j, n + 1):
                    # same number in the other script after the options = the translation
                    if opts_seen >= 2 and lang and script(it.text) not in ("", lang):
                        continue
                    return False
        return False

    def _valid_restart(self, i: int, expected: int) -> bool:
        """Numbering restarting at 1 (a new section, e.g. the Hindi half of a
        language paper). Rejected if the expected next number shows up before the
        next answer key -- then this '1.' is a list inside an explanation."""
        if not self._valid_start(i, 1):
            return False
        for j in range(i + 1, min(len(self.items), i + self.LOOKAHEAD)):
            it = self.items[j]
            if self.answer(it):
                return True
            q = self.qnum(it)
            if q and expected <= q[0] <= expected + 3:
                return False
        return False

    # ---- segmentation ---------------------------------------------------- #
    def segment(self) -> List[_Seg]:
        """Pick the first-question start that yields the most consistent run.

        Title pages, indexes and instruction lists ("1. Read carefully ...") also
        look like numbered questions. Each numbered line before the first answer
        key is tried as the start; the winner produces the most numbered and
        answered questions, then a well-formed first question, then the shortest
        first question (so an index or instruction list never swallows Q.1)."""
        first_ans = next((i for i, it in enumerate(self.items) if self.answer(it)), None)
        if first_ans is None:
            return []
        cands = [i for i in range(first_ans) if self.qnum(self.items[i])][-self.START_CANDIDATES:]
        if len(cands) <= 1:
            return self._segment(cands[0] if cands else None)
        best, best_score = None, None
        for c in cands:
            segs = self._segment(c)
            if not segs:
                continue
            first = self.build(segs[0])
            score = (sum(1 for s in segs if s.number is not None),
                     sum(1 for s in segs if s.answer is not None),
                     1 if first.ok else 0,
                     1 if first.language == "bi" else 0,   # English+Hindi pair beats its Hindi half
                     -len(segs[0].body))
            if best_score is None or score > best_score:
                best, best_score = segs, score
        return best or []

    def _segment(self, forced_start: Optional[int]) -> List[_Seg]:
        segs: List[_Seg] = []
        phase = "pre"
        expected: Optional[int] = None
        cur: Optional[_Seg] = None
        if forced_start is None:
            return segs

        for i, it in enumerate(self.items):
            if phase == "pre" and i != forced_start:
                continue
            q = self.qnum(it)
            if q and phase in ("pre", "solution"):
                n = q[0]
                in_range = expected is None or expected <= n <= expected + 3
                ok_start = in_range and (phase == "pre" or self._valid_start(i, n))
                restart = (not ok_start and phase == "solution" and n == 1 and expected is not None
                           and expected > 2 and self._valid_restart(i, expected))
                if ok_start or restart:
                    if expected is not None and n != expected and not restart:
                        missing = ", ".join(str(k) for k in range(expected, n))
                        segs and segs[-1].warnings.append(
                            f"next number jumps to {n} (no question numbered {missing} in source)")
                    cur = _Seg(number=n, body=[strip_prefix(it, q[1])])
                    segs.append(cur)
                    phase, expected = "body", n + 1
                    continue

            if phase == "pre":
                continue  # headers, instructions, title page...

            if phase == "body":
                a = self.answer(it)
                if a:
                    cur.answer, cur.answer_line = a[0], it
                    cur.answer_tail = self._answer_tail(it, a[1])
                    phase = "solution"
                    continue
                # answer key missing but the next question clearly began
                if q and expected is not None and expected <= q[0] <= expected + 3 \
                        and self._has_option_run(cur.body) and not self.opt_marker(it) \
                        and not self._continues_list(cur.body, q[0]) \
                        and self._valid_start(i, q[0]):
                    cur = _Seg(number=q[0], body=[strip_prefix(it, q[1])])
                    segs.append(cur)
                    expected = q[0] + 1
                    continue
                cur.body.append(it)
                continue

            # phase == "solution"
            a = self.answer(it)
            if a and isinstance(it, Line):
                new = self._split_unnumbered(cur)
                if new is not None:
                    new.answer, new.answer_line = a[0], it
                    new.answer_tail = self._answer_tail(it, a[1])
                    segs.append(new)
                    cur = new
                    continue
            if isinstance(it, Line) and it.heading:
                continue        # chapter / section titles between questions
            cur.solution.append(it)
        return segs

    @staticmethod
    def _answer_tail(line: Line, end: int) -> Optional[Line]:
        """Explanation written on the answer line itself, if any."""
        tail = strip_prefix(line, end)
        m = INLINE_EXPL_RE.search(tail.text)
        if not m:
            return None
        return strip_prefix(tail, m.start())

    def _continues_list(self, body: List[object], n: int) -> bool:
        """True if the last numbered line already in the body is n-1 (so a line
        numbered n is the next statement of a list, not a new question)."""
        for it in reversed(body[1:]):
            q = self.qnum(it)
            if q:
                return q[0] == n - 1
        return False

    def _has_option_run(self, body: List[object]) -> bool:
        return len(self._find_options(body)[0]) >= 2

    def _split_unnumbered(self, cur: _Seg) -> Optional[_Seg]:
        """An answer key showed up inside a solution: an unnumbered question is
        hiding at the tail of that solution. Peel it off."""
        sol = cur.solution
        idxs, _ = self._find_options(sol)
        if len(idxs) < 2:
            return None
        first_opt = idxs[0]
        k = first_opt - 1
        # statement / list lines directly above the options belong to the question
        while k >= 0 and (isinstance(sol[k], Table) or LISTISH_RE.match(sol[k].text)
                          or (isinstance(sol[k], Line) and sol[k].auto_number)):
            k -= 1
        if k < 0:
            return None
        stem_start = k
        # extend upward over stem-like lines (ending with ? : - or starting with cue words)
        while stem_start - 1 >= 0 and isinstance(sol[stem_start - 1], Line):
            t = sol[stem_start - 1].text.strip()
            if re.search(r"[?:\-\u2013\u2014]$", t) or re.match(
                    r"^(Consider|Which|What|Who|When|Where|Read|Match|Assertion|Reason)\b", t):
                stem_start -= 1
            else:
                break
        stem = sol[stem_start:]
        first = stem[0]
        if isinstance(first, Line):
            first = lstrip_line(strip_prefix(first, 0), " \t.:-)\u2013\u2014")
            stem = [first] + stem[1:]
        cur.solution = sol[:stem_start]
        new = _Seg(number=None, body=stem)
        new.warnings.append("question number missing in source; start of question auto-detected — please verify")
        return new


    def _find_options(self, body: List[object], numeric: bool = False) -> Tuple[List[int], bool]:
        """Indices of option lines (a..), found by scanning back from the end.
        Returns (indices in ascending order, regular_sequence)."""
        seq: List[int] = []
        want: Optional[int] = None
        for idx in range(len(body) - 1, -1, -1):
            m = self.opt_marker(body[idx], numeric)
            if not m:
                continue
            letter = m[0]
            if want is None:
                if letter < 1:          # the last marker must be (b) or later
                    continue
                seq.append(idx)
                want = letter - 1
            elif letter == want:
                seq.append(idx)
                want -= 1
            else:
                continue                # out-of-sequence marker = text inside an option
            if want < 0:
                return list(reversed(seq)), True
        # lenient fallback: last (a) and every marker after it, in order
        a_idx = [i for i in range(len(body)) if (self.opt_marker(body[i], numeric) or (None,))[0] == 0]
        if a_idx:
            start = a_idx[-1]
            out, nxt = [start], 1
            for i in range(start + 1, len(body)):
                m = self.opt_marker(body[i], numeric)
                if m and m[0] == nxt:
                    out.append(i)
                    nxt += 1
            if len(out) >= 2:
                return out, False
        # Word list that kept counting from an earlier list (e. f. g. h.): take the
        # last consecutive auto-lettered run and read it by position.
        if not numeric:
            seq, prev = [], None
            for idx in range(len(body) - 1, -1, -1):
                m = self.opt_marker(body[idx], shifted=True)
                if not m:
                    continue
                if prev is None or m[0] == prev - 1:
                    seq.append(idx)
                    prev = m[0]
                else:
                    break
            if len(seq) >= 2:
                self.last_shifted = True
                return list(reversed(seq)), False
        return [], False

    # ---- one language version: stem + options --------------------------- #
    def _options_any(self, items: List[object]) -> Tuple[List[int], bool, bool]:
        idxs, regular = self._find_options(items)
        if len(idxs) >= 2:
            return idxs, regular, False
        idxs, regular = self._find_options(items, numeric=True)
        return idxs, regular, True

    def _parse_part(self, body: List[object]) -> _Part:
        part = _Part()
        self.last_shifted = False
        idxs, regular, numeric = [], True, False
        for attempt in ("plain", "flatten", "numeric"):
            work = body
            if attempt == "flatten":
                if not any(isinstance(b, Table) for b in body):
                    continue
                work = []
                for b in body:
                    work.extend(b.flatten() if isinstance(b, Table) else [b])
            numeric = attempt == "numeric"
            idxs, regular = self._find_options(work, numeric=numeric)
            if len(idxs) >= 2:
                body = work
                break
        part.numeric = numeric
        if len(idxs) < 2:
            part.errors.append("options (a), (b), ... could not be identified")
            part.stem = [clean_line(copy.deepcopy(b)) if isinstance(b, Line) else b for b in body]
        else:
            prev = body[idxs[0] - 1] if idxs[0] > 0 else None
            if len(idxs) < 4 and prev is not None and (self.opt_marker(prev, numeric)
                                                      or self.opt_marker(prev)):
                labels = "(a), (b), (a), (b)" if len(idxs) == 2 else "repeated letters"
                part.errors.append(f"option labels repeat in source ({labels}) — "
                                   f"fix them to (a), (b), (c), (d) and re-upload")
            if self.last_shifted:
                part.warnings.append("options are lettered from the middle of the alphabet in source "
                                     "(Word list kept counting); read in order as (a), (b), ... — please verify")
            elif not regular:
                part.warnings.append("option labels are out of sequence in source — please verify")
            stem = body[: idxs[0]]
            if stem and isinstance(stem[-1], Line) and OPTIONS_LABEL_RE.match(stem[-1].text):
                stem = stem[:-1]
            part.stem = [clean_line(copy.deepcopy(b)) if isinstance(b, Line) else b for b in stem]
            bounds = idxs + [len(body)]
            for k, start in enumerate(idxs):
                m = self.opt_marker(body[start], numeric) or self.opt_marker(body[start], shifted=True)
                opt_lines = [clean_line(strip_prefix(body[start], m[1]))]
                for extra in body[start + 1: bounds[k + 1]]:
                    if isinstance(extra, Table):
                        opt_lines.extend(clean_line(copy.deepcopy(l)) for l in extra.flatten())
                    else:
                        opt_lines.append(clean_line(copy.deepcopy(extra)))
                opt_lines = [l for l in opt_lines if not l.is_blank()]
                part.options.append(opt_lines)
            if any(not o for o in part.options):
                part.errors.append("one or more options are empty")
        part.stem = [b for b in part.stem if not (isinstance(b, Line) and b.is_blank())]
        if part.stem and isinstance(part.stem[0], Line):
            part.stem[0] = lstrip_line(part.stem[0], " \t.:-)\u2013\u2014।")
        if not part.stem or not any((b.text.strip() if isinstance(b, Line) else True) for b in part.stem):
            part.errors.append("question text is empty")
        return part

    # ---- bilingual split -------------------------------------------------- #
    @staticmethod
    def _first_script(items: List[object]) -> str:
        for it in items:
            s = script(it.text)
            if s:
                return s
        return ""

    def _split_bilingual(self, body: List[object]):
        """(first_version, second_version, first_lang) or None."""
        first_lang = self._first_script(body)
        if not first_lang:
            return None
        other = "hi" if first_lang == "en" else "en"
        for h in range(1, len(body)):
            it = body[h]
            if not isinstance(it, Line) or script(it.text) != other or self.opt_marker(it):
                continue
            idxs, regular, _ = self._options_any(body[:h])
            if len(idxs) < 2 or not regular:
                continue
            between = body[idxs[-1] + 1: h]
            if any(isinstance(t, Line) and script(t.text) == other for t in between):
                continue
            second = self._strip_qnum(body[h:])
            idx2, _, _ = self._options_any(second)
            if len(idx2) >= 2:
                return body[:h], second, first_lang
        return None

    def _strip_qnum(self, items: List[object]) -> List[object]:
        """Drop the repeated question number on the second-language version."""
        if items and isinstance(items[0], Line):
            q = self.qnum(items[0])
            if q:
                return [strip_prefix(items[0], q[1])] + list(items[1:])
        return list(items)

    def _looks_mixed(self, body: List[object], idxs: List[int]) -> bool:
        """Two option runs written in different scripts that could not be split."""
        if len(idxs) < 2:
            return False
        opt_script = self._first_script([body[i] for i in idxs])
        before, _, _ = self._options_any(body[: idxs[0]])
        if len(before) < 2:
            return False
        return self._first_script([body[i] for i in before]) not in ("", opt_script)

    # ---- solution --------------------------------------------------------- #
    def _split_solution(self, items: List[object], first_lang: str):
        """Split a bilingual solution into (english, hindi)."""
        cut = None
        for k, it in enumerate(items):
            if not isinstance(it, Line):
                continue
            if first_lang == "en" and HI_EXPL_RE.match(it.text) and k > 0:
                cut = k
                break
            if first_lang == "hi" and EN_EXPL_RE.match(it.text) and k > 0:
                cut = k
                break
        if cut is None:
            other = "hi" if first_lang == "en" else "en"
            seen_first = False
            for k, it in enumerate(items):
                s = script(it.text)
                if s == first_lang:
                    seen_first = True
                elif s == other and (seen_first or k > 0 or HI_EXPL_RE.match(it.text)):
                    cut = k
                    break
                elif s == other:            # solution starts in the other language
                    cut = k
                    break
        first, second = (items, []) if cut is None else (items[:cut], items[cut:])
        if first_lang == "hi":
            first, second = second, first
        return first, second

    def _clean_solution(self, items: List[object], answer: Optional[str]) -> List[object]:
        out: List[object] = []
        label_done = False
        for it in items:
            if isinstance(it, Line):
                if it.heading:
                    continue
                a = self.answer(it)
                if a and answer is not None and to_index(a[0]) == to_index(answer) \
                        and len(it.text) <= 90:
                    self.duplicate_keys += 1
                    continue        # repeated answer key, e.g. "उत्तर: (b)" after "Answer: b"
                ln = clean_line(copy.deepcopy(it))
                if not label_done:
                    m = EXPL_RE.match(ln.text)
                    if m:
                        ln = clean_line(strip_prefix(ln, m.end()))
                        label_done = True
                    elif ln.text.strip():
                        label_done = True
                if ln.is_blank():
                    continue
                out.append(ln)
            else:
                label_done = True
                out.append(it)
        # trailing section headings like "COMPETITION ACT" belong to no question
        while out and isinstance(out[-1], Line):
            t = out[-1].text.strip()
            if HEADING_RE.match(t) and any(c.isalpha() for c in t) and t.upper() == t \
                    and len(t.split()) <= 8 and not t.endswith("."):
                out.pop()
            else:
                break
        return out

    @staticmethod
    def _auto_solution(option: List[Line], lang: str) -> Line:
        """'The correct answer is <option text>' (Hindi: 'सही उत्तर है: <option text>')."""
        runs = [Run(text=AUTO_HI if lang == "hi" else AUTO_EN)]
        for k, ln in enumerate(option):
            if k:
                runs.append(Run(text=" "))
            runs.extend(copy.deepcopy(r) for r in ln.runs)
        text = "".join(r.text for r in runs if r.image is None).rstrip()
        if text and text[-1] not in ".?!।" and lang == "en":
            runs.append(Run(text="."))
        return Line(runs=runs)

    # ---- build ------------------------------------------------------------ #
    def build(self, seg: _Seg) -> Question:
        qn = Question(source_number=seg.number, warnings=list(seg.warnings))
        body = list(seg.body)
        split = self._split_bilingual(body)
        numeric = False
        if split:
            first, second, first_lang = split
            p1, p2 = self._parse_part(first), self._parse_part(second)
            en, hi = (p1, p2) if first_lang == "en" else (p2, p1)
            qn.language = "bi"
            qn.body, qn.options = en.stem, en.options
            qn.hi_body, qn.hi_options = hi.stem, hi.options
            qn.errors += [f"English: {e}" for e in en.errors] + [f"Hindi: {e}" for e in hi.errors]
            qn.warnings += [f"English: {w}" for w in en.warnings] + [f"Hindi: {w}" for w in hi.warnings]
            if en.options and hi.options and len(en.options) != len(hi.options):
                qn.errors.append(f"English has {len(en.options)} options but Hindi has {len(hi.options)}")
            numeric = en.numeric or hi.numeric
        else:
            first_lang = ""
            part = self._parse_part(body)
            qn.body, qn.options = part.stem, part.options
            qn.errors += part.errors
            qn.warnings += part.warnings
            numeric = part.numeric
            dev, lat = letters_by_script(body)
            qn.language = "hi" if dev > lat else "en"
            idxs, _, _ = self._options_any(body)
            if self._looks_mixed(body, idxs):
                qn.errors.append("contains English and Hindi versions that could not be separated")

        # answer
        if seg.answer is None:
            qn.errors.append("answer key ('Ans.' / 'Answer:' / 'उत्तर:') not found")
        else:
            qn.answer_raw = seg.answer
            idx = to_index(seg.answer)
            if numeric and not seg.answer.translate(_DEV_DIGITS).isdigit() and seg.answer != "©":
                qn.errors.append(f"answer is '{seg.answer}' but the options are numbered 1, 2, 3 ...")
            elif qn.options and 0 <= idx < len(qn.options):
                qn.correct_index = idx
            elif qn.options:
                qn.errors.append(f"answer '{seg.answer}' does not match any of the {len(qn.options)} options")

        # solution
        sol: List[object] = []
        if seg.answer_tail is not None:
            sol.append(seg.answer_tail)
        sol.extend(seg.solution)
        if qn.language == "bi":
            en_s, hi_s = self._split_solution(sol, first_lang)
            qn.solution = self._clean_solution(en_s, seg.answer)
            qn.hi_solution = self._clean_solution(hi_s, seg.answer)
        else:
            qn.solution = self._clean_solution(sol, seg.answer)

        # write missing solutions from the correct option
        k = qn.correct_index
        if not qn.solution:
            if k is not None and qn.options and qn.options[k]:
                qn.solution = [self._auto_solution(qn.options[k], qn.language if qn.language == "hi" else "en")]
                qn.auto_solution.append("English" if qn.language != "hi" else "Hindi")
            else:
                qn.errors.append("explanation / solution not found")
        if qn.language == "bi" and not qn.hi_solution:
            if k is not None and k < len(qn.hi_options) and qn.hi_options[k]:
                qn.hi_solution = [self._auto_solution(qn.hi_options[k], "hi")]
                qn.auto_solution.append("Hindi")
            else:
                qn.errors.append("Hindi explanation / solution not found")
        for tag, opts in (("", qn.options), ("Hindi ", qn.hi_options)):
            norm = [" ".join(" ".join(l.text for l in o).lower().split()) for o in opts]
            dup = sorted({i for i in range(len(norm)) for j in range(len(norm))
                          if i != j and norm[i] and norm[i] == norm[j]})
            if dup:
                labels = " and ".join(f"({chr(97 + i)})" for i in dup)
                qn.warnings.append(f"{tag}options {labels} have identical text in source — please verify")
        qn.ok = not qn.errors
        return qn

    # ---- entry ----------------------------------------------------------- #
    def parse(self) -> List[Question]:
        segs = self.segment()
        self.duplicate_keys = 0
        qs = [self.build(s) for s in segs]
        finalize_languages(qs)
        return qs

    def count_answer_keys(self) -> int:
        return sum(1 for it in self.items if self.answer(it)) - self.duplicate_keys


def _runs(qs: List[Question]):
    i = 0
    while i < len(qs):
        j = i
        while j < len(qs) and qs[j].language == qs[i].language:
            j += 1
        yield i, j
        i = j


def finalize_languages(qs: List[Question]) -> None:
    """In a bilingual paper, a short run (< MIN_SECTION) of single-language
    questions means a missing translation; a longer run is a language section
    (e.g. the English or Hindi half of a language paper) and is fine.
    In a single-language paper, a lone question whose script differs from the
    section around it (e.g. 'translate into Hindi' inside the English section)
    takes the section's language."""
    if not any(q.language == "bi" for q in qs):
        runs = list(_runs(qs))
        for k, (i, j) in enumerate(runs):
            if j - i >= MIN_SECTION:
                continue
            prev_lang = qs[runs[k - 1][0]].language if k > 0 else None
            next_lang = qs[runs[k + 1][0]].language if k + 1 < len(runs) else None
            prev_big = k > 0 and runs[k - 1][1] - runs[k - 1][0] >= MIN_SECTION
            next_big = k + 1 < len(runs) and runs[k + 1][1] - runs[k + 1][0] >= MIN_SECTION
            if prev_big and next_big and prev_lang == next_lang:
                lang = prev_lang
            elif prev_big and not next_big:
                lang = prev_lang
            elif next_big and not prev_big:
                lang = next_lang
            else:
                continue
            for q in qs[i:j]:
                q.language = lang
        return
    i = 0
    while i < len(qs):
        if qs[i].language == "bi":
            i += 1
            continue
        j = i
        while j < len(qs) and qs[j].language == qs[i].language:
            j += 1
        if j - i < MIN_SECTION:
            missing = "Hindi" if qs[i].language == "en" else "English"
            for q in qs[i:j]:
                q.isolated = True
                q.ok = False
                q.errors.append(f"{missing} version not found (the rest of this paper is English + Hindi)")
        i = j


def summarize(qs: List[Question]) -> Dict[str, int]:
    """Counts shown in the app and written to the Excel summary."""
    bi = [q for q in qs if q.language == "bi"]
    iso_en = [q for q in qs if q.isolated and q.language == "en"]
    iso_hi = [q for q in qs if q.isolated and q.language == "hi"]
    return {
        "total": len(qs),
        "ready": sum(q.ok for q in qs),
        "failed": sum(not q.ok for q in qs),
        "bilingual_english": len(bi) + len(iso_en),
        "bilingual_hindi": len(bi) + len(iso_hi),
        "english_only": sum(1 for q in qs if q.language == "en" and not q.isolated),
        "hindi_only": sum(1 for q in qs if q.language == "hi" and not q.isolated),
        "auto_solutions": sum(1 for q in qs if q.ok and q.auto_solution),
    }


def parse_docx(data: bytes) -> Tuple[List[Question], int]:
    """Returns (questions, number_of_answer_keys_seen_in_document)."""
    p = QuestionParser(read_blocks(data))
    qs = p.parse()
    return qs, p.count_answer_keys()
