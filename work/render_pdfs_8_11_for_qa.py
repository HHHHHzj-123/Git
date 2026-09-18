from pathlib import Path

import pypdfium2 as pdfium


ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
for episode in (8, 9, 10, 11):
    pdf_path = ROOT / "manual_pdf_8_11" / f"ep{episode}.pdf"
    out_dir = ROOT / f"render_ep{episode}"
    out_dir.mkdir(exist_ok=True)
    pdf = pdfium.PdfDocument(str(pdf_path))
    for index in range(len(pdf)):
        bitmap = pdf[index].render(scale=2.0)
        bitmap.to_pil().save(out_dir / f"page-{index + 1:03d}.png")
    print(f"EP{episode}\tpages={len(pdf)}")
