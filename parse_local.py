"""
Offline CLI: parse a Word file and write one formatted .docx per question to a
local folder. No Google account needed. Handy for checking a new doc style.

    python scripts/parse_local.py "paper.docx" out_folder --category Judiciary
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from qbg.builder import QuestionDocBuilder  # noqa: E402
from qbg.parser import parse_docx  # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("docx")
ap.add_argument("out")
ap.add_argument("--category", default="Judiciary")
a = ap.parse_args()

qs, keys = parse_docx(open(a.docx, "rb").read())
os.makedirs(a.out, exist_ok=True)
b = QuestionDocBuilder(a.category)
ok = 0
for i, q in enumerate(qs, 1):
    if q.ok:
        open(os.path.join(a.out, f"Q{i:04d}.docx"), "wb").write(b.build(q))
        ok += 1
    if q.remarks():
        print(f"#{i:<4} {'OK  ' if q.ok else 'FAIL'} {q.remarks()}")
print(f"\n{len(qs)} questions detected, {ok} written, answer keys in file: {keys}")
