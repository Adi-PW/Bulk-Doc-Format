"""Print a human-readable dump of every parsed question (for eyeballing)."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from qbg.parser import parse_docx
from qbg.docx_reader import Table

def show(items, indent="    "):
    for it in items:
        if isinstance(it, Table):
            print(indent + "[TABLE " + " || ".join(" | ".join(c.text.replace("\n"," / ") for c in r) for r in it.rows) + "]")
        else:
            print(indent + it.text)

path = sys.argv[1]
qs, keys = parse_docx(open(path, "rb").read())
print(f"questions={len(qs)} answer_keys_in_doc={keys} ok={sum(q.ok for q in qs)}")
for i, q in enumerate(qs, 1):
    print(f"\n===== #{i} {q.label} ok={q.ok} ans={q.answer_raw}->{q.correct_index} {q.remarks()}")
    print("  Q:"); show(q.body)
    for k, o in enumerate(q.options):
        mark = "*" if k == q.correct_index else " "
        print(f"  {mark}({chr(97+k)}) " + " / ".join(l.text for l in o))
    print("  S:"); show(q.solution)
