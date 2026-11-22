"""Build column.docx from column.md (headings, italics, paragraphs)."""
import re
from docx import Document
from docx.shared import Pt

def add_runs(par, text):
    for part in re.split(r"(\*[^*]+\*)", text):
        if part.startswith("*") and part.endswith("*"):
            par.add_run(part[1:-1]).italic = True
        elif part:
            par.add_run(part)

doc = Document()
doc.styles["Normal"].font.name = "Calibri"
doc.styles["Normal"].font.size = Pt(11)
for block in open("column.md", encoding="utf-8").read().split("\n\n"):
    b = block.strip()
    if not b:
        continue
    if b.startswith("# "):
        doc.add_heading(b[2:], level=1)
    elif b.startswith("## "):
        doc.add_heading(b[3:], level=2)
    else:
        add_runs(doc.add_paragraph(), b)
doc.save("column.docx")
