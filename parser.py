"""
Turns the linear block stream from ``docx_reader`` into structured questions.

Segmentation is a small state machine:

    pre  --(valid question start)-->  body  --(answer line)-->  solution
     ^                                                            |
     +------------------(next valid question start)---------------+

Inside a question, options are found by scanning *backwards* from the answer
line for the last complete marker run  ... (d) (c) (b) (a).  Everything above
option (a) is the question body.  This is what makes Type-3 questions safe:
statements labelled (a)/(b)/(c) sit above the real option run and stay in the
body, because the backward scan stops at the first (a) it meets.
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from .docx_reader import Line, Run, Table, read_blocks

# --------------------------------------------------------------------------- #
# Patterns
# --------------------------------------------------------------------------- #
# Question start: "1.", "1)", "1:", "Q1.", "Q.1", "Question 1:", or "13 'Admission'"
QSTART_RE = re.compile(
    r"""^\s*[\[(]?\s*
        (?:[Qq](?:ues(?:tion)?|UES(?:TION)?)?\s*(?:[Nn][Oo]\.?)?\s*[.:\-]?\s*)?
        (?P<num>\d{1,3})(?!\d)
        (?:\s*[.):]\s*|\s+(?=[A-Z'"\u2018\u201c(\[]))
    """,
    re.X,
)

# Option marker: "(a)", "a)", "a.", "(A)", "A)", "A." -- letter must be a..h
OPT_LETTER_RE = re.compile(r"^\s*(?:\(\s*(?P<l1>[a-hA-H])\s*\)|(?P<l2>[a-hA-H])\s*[.)])\s*")
# Numeric fallback: "(1)", "1)" (used only when no letter options exist)
OPT_NUM_RE = re.compile(r"^\s*(?:\(\s*(?P<n1>[1-8])\s*\)|(?P<n2>[1-8])\s*\))\s*")

ANSWER_RE = re.compile(
    r"""^\s*(?:Correct\s+|Right\s+)?
        (?:[Aa]ns(?:wer)?|ANS(?:WER)?|Key)\s*
        (?:is\s*)?[.:\-\u2013\u2014=]*\s*[:\-\u2013\u2014]?\s*
        (?:Option\s*)?
        [\[(]?\s*(?P<ans>[a-hA-H1-8])\s*[\])]?
        (?![A-Za-z0-9])
        [\s.):\-\u2013\u2014]*
    """,
    re.X,
)

EXPL_RE = re.compile(
    r"^\s*(?:Explanation|Explaination|Explanations|Expl\.?|Solution|Sol\.?|Rationale)\s*[:\-\u2013\u2014.]*\s*",
    re.I,
)

OPTIONS_LABEL_RE = re.compile(r"^\s*(?:Options?|Choices?|Alternatives?)\s*[:\-\u2013\u2014]?\s*$", re.I)

# lines that look like statement / list items (used for unnumbered fallback)
LISTISH_RE = re.compile(
    r"^\s*(?:\(?[a-zA-Z]\)|[a-zA-Z]\.|\(?[ivxIVX]{1,5}[.)]|\(?\d{1,2}[.)]|[•\-\u2022])\s+"
)

HEADING_RE = re.compile(r"^[A-Z0-9 &,.'()\-/]{3,60}$")


# --------------------------------------------------------------------------- #
# Run helpers
# --------------------------------------------------------------------------- #
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
    body: List[object] = field(default_factory=list)          # Line | Table
    options: List[List[Line]] = field(default_factory=list)   # each option = list of Lines
    correct_index: Optional[int] = None
    answer_raw: str = ""
    solution: List[object] = field(default_factory=list)
    ok: bool = True
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def label(self) -> str:
        return f"Q.{self.source_number}" if self.source_number is not None else "Q.(unnumbered)"

    @property
    def body_text(self) -> str:
        return "\n".join(b.text for b in self.body)

    def remarks(self) -> str:
        msgs = self.errors + self.warnings
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


# --------------------------------------------------------------------------- #
# Parser
# --------------------------------------------------------------------------- #
class QuestionParser:
    LOOKAHEAD = 600

    def __init__(self, blocks: List[object]):
        # Drop blank lines / empty tables up front: they carry no meaning.
        self.items = [b for b in blocks if not b.is_blank()]

    # ---- line classifiers ----------------------------------------------- #
    @staticmethod
    def qnum(item) -> Optional[Tuple[int, int]]:
        if not isinstance(item, Line):
            return None
        m = QSTART_RE.match(item.text)
        if not m:
            return None
        # Do not mistake an option line such as "(1) foo" for a question start
        return int(m.group("num")), m.end()

    @staticmethod
    def answer(item) -> Optional[Tuple[str, int]]:
        if not isinstance(item, Line):
            return None
        m = ANSWER_RE.match(item.text)
        if not m:
            return None
        return m.group("ans"), m.end()

    @staticmethod
    def opt_marker(item, numeric: bool = False) -> Optional[Tuple[int, int]]:
        if not isinstance(item, Line):
            return None
        if numeric:
            m = OPT_NUM_RE.match(item.text)
            if not m:
                return None
            return int(m.group("n1") or m.group("n2")) - 1, m.end()
        m = OPT_LETTER_RE.match(item.text)
        if not m:
            return None
        letter = (m.group("l1") or m.group("l2")).lower()
        return ord(letter) - ord("a"), m.end()

    # ---- start validation ----------------------------------------------- #
    def _valid_start(self, i: int, n: int) -> bool:
        """A numbered line is a real question start only if an answer key follows
        before another line with the *same* number appears (unless that later
        line is part of a fresh 1..n list, i.e. statements inside the question)."""
        saw_prev = False
        for j in range(i + 1, min(len(self.items), i + self.LOOKAHEAD)):
            it = self.items[j]
            if self.answer(it):
                return True
            q = self.qnum(it)
            if q:
                if q[0] == n - 1:
                    saw_prev = True
                elif q[0] == n and not saw_prev:
                    return False
        return False

    # ---- segmentation ---------------------------------------------------- #
    def segment(self) -> List[_Seg]:
        """Pick the first-question start that yields the most consistent run.

        Title pages and instruction lists ("1. Read carefully  2. ...") also look
        like numbered questions; trying each candidate that precedes the first
        answer key and keeping the one that produces the most well-formed
        questions (earliest wins ties) separates them reliably."""
        first_ans = next((i for i, it in enumerate(self.items) if self.answer(it)), None)
        if first_ans is None:
            return []
        cands = [i for i in range(first_ans) if self.qnum(self.items[i])]
        cands = [i for i in cands if self._valid_start(i, self.qnum(self.items[i])[0])][:25]
        if len(cands) <= 1:
            return self._segment(cands[0] if cands else None)
        best, best_score = None, None
        for c in cands:
            segs = self._segment(c)
            score = (sum(1 for s in segs if s.number is not None),
                     sum(1 for s in segs if s.answer is not None))
            if best_score is None or score > best_score:
                best, best_score = segs, score
        return best

    def _segment(self, forced_start: Optional[int]) -> List[_Seg]:
        segs: List[_Seg] = []
        phase = "pre"
        expected: Optional[int] = None
        cur: Optional[_Seg] = None
        if forced_start is None:
            return segs

        for i, it in enumerate(self.items):
            q = self.qnum(it)
            if phase == "pre" and i != forced_start:
                continue
            if q and phase in ("pre", "solution"):
                n = q[0]
                in_range = expected is None or expected <= n <= expected + 3
                if in_range and (phase == "pre" or self._valid_start(i, n)):
                    if expected is not None and n != expected:
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
                    tail = strip_prefix(it, a[1])
                    if tail.text.strip() and EXPL_RE.match(tail.text):
                        cur.answer_tail = tail
                    phase = "solution"
                    continue
                # answer key missing but next question clearly began
                if q and expected is not None and expected <= q[0] <= expected + 3 \
                        and self._has_option_run(cur.body) and not self.opt_marker(it) \
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
                    tail = strip_prefix(it, a[1])
                    if tail.text.strip() and EXPL_RE.match(tail.text):
                        new.answer_tail = tail
                    segs.append(new)
                    cur = new
                    continue
            cur.solution.append(it)
        return segs

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

    # ---- options --------------------------------------------------------- #
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
        return [], False

    def build(self, seg: _Seg) -> Question:
        qn = Question(source_number=seg.number, warnings=list(seg.warnings))
        body = list(seg.body)

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

        if len(idxs) < 2:
            qn.ok = False
            qn.errors.append("options (a), (b), ... could not be identified")
            qn.body = [clean_line(b) if isinstance(b, Line) else b for b in body]
        else:
            if not regular:
                qn.warnings.append("option labels are out of sequence in source — please verify")
            stem = body[: idxs[0]]
            if stem and isinstance(stem[-1], Line) and OPTIONS_LABEL_RE.match(stem[-1].text):
                stem = stem[:-1]
            qn.body = [clean_line(copy.deepcopy(b)) if isinstance(b, Line) else b for b in stem]
            bounds = idxs + [len(body)]
            for k, start in enumerate(idxs):
                m = self.opt_marker(body[start], numeric)
                opt_lines = [clean_line(strip_prefix(body[start], m[1]))]
                for extra in body[start + 1: bounds[k + 1]]:
                    if isinstance(extra, Table):
                        opt_lines.extend(clean_line(copy.deepcopy(l)) for l in extra.flatten())
                    else:
                        opt_lines.append(clean_line(copy.deepcopy(extra)))
                opt_lines = [l for l in opt_lines if not l.is_blank()]
                qn.options.append(opt_lines)
            if any(not o for o in qn.options):
                qn.ok = False
                qn.errors.append("one or more options are empty")

        # strip a dangling leading punctuation left after the number ("4.Which" etc.)
        qn.body = [b for b in qn.body if not (isinstance(b, Line) and b.is_blank())]
        if qn.body and isinstance(qn.body[0], Line):
            qn.body[0] = lstrip_line(qn.body[0], " \t.:-)\u2013\u2014")
        if not qn.body or not any((b.text.strip() if isinstance(b, Line) else True) for b in qn.body):
            qn.ok = False
            qn.errors.append("question text is empty")

        # answer
        if seg.answer is None:
            qn.ok = False
            qn.errors.append("answer key ('Ans.' / 'Answer:') not found")
        else:
            qn.answer_raw = seg.answer
            a = seg.answer.lower()
            idx = (int(a) - 1) if a.isdigit() else (ord(a) - ord("a"))
            if qn.options and 0 <= idx < len(qn.options):
                qn.correct_index = idx
            elif qn.options:
                qn.ok = False
                qn.errors.append(f"answer '{seg.answer}' does not match any of the {len(qn.options)} options")

        # solution
        sol: List[object] = []
        if seg.answer_tail is not None:
            sol.append(seg.answer_tail)
        sol.extend(seg.solution)
        sol = self._clean_solution(sol)
        if not sol:
            qn.ok = False
            qn.errors.append("explanation / solution not found")
        qn.solution = sol
        return qn

    @staticmethod
    def _clean_solution(items: List[object]) -> List[object]:
        out: List[object] = []
        label_done = False
        for it in items:
            if isinstance(it, Line):
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

    # ---- entry ----------------------------------------------------------- #
    def parse(self) -> List[Question]:
        return [self.build(s) for s in self.segment()]

    def count_answer_keys(self) -> int:
        return sum(1 for it in self.items if self.answer(it))


def parse_docx(data: bytes) -> Tuple[List[Question], int]:
    """Returns (questions, number_of_answer_keys_seen_in_document)."""
    p = QuestionParser(read_blocks(data))
    return p.parse(), p.count_answer_keys()
