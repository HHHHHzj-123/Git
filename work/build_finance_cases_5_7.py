from pathlib import Path

from build_finance_cases_2_4 import build_docx, build_markdown, build_xmind
from docx import Document
from docx.shared import Pt
from finance_cases_data_5_7 import CASES


ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")


def apply_long_case_compact_override(docx_path):
    """Named override for episode 7: keep the long case readable without an orphan final page."""
    doc = Document(docx_path)
    for paragraph in doc.paragraphs:
        ppr = paragraph._p.pPr
        if ppr is not None and ppr.numPr is not None:
            paragraph.paragraph_format.space_after = Pt(2)
            paragraph.paragraph_format.line_spacing = 1.20
    doc.save(docx_path)


if __name__ == "__main__":
    for case in CASES:
        md = build_markdown(case)
        docx = build_docx(case)
        if case["episode"] == 7:
            apply_long_case_compact_override(docx)
        xmind, counts = build_xmind(case)
        print(
            f"EP{case['episode']}\tDOCX={docx}\tXMIND={xmind}\t"
            f"TOPICS={counts}\tMD={md.stat().st_size}"
        )
