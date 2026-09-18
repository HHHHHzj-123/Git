from build_finance_cases_2_4 import build_docx, build_markdown, build_xmind
from finance_cases_data_23_25 import CASES


if __name__ == "__main__":
    for case in CASES:
        md = build_markdown(case)
        docx = build_docx(case)
        xmind, counts = build_xmind(case)
        print(
            f"EP{case['episode']}\tDOCX={docx}\tXMIND={xmind}\t"
            f"TOPICS={counts}\tMD={md.stat().st_size}"
        )
