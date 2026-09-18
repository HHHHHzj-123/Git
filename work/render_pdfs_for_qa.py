from pathlib import Path

import pypdfium2 as pdfium

ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
for pdf_path in sorted((ROOT / "manual_pdf").glob("*.pdf")):
    episode = pdf_path.name.split("第", 1)[1].split("集", 1)[0]
    out_dir = ROOT / f"render_ep{episode}"
    out_dir.mkdir(exist_ok=True)
    pdf = pdfium.PdfDocument(str(pdf_path))
    for index in range(len(pdf)):
        bitmap = pdf[index].render(scale=2.0)
        image = bitmap.to_pil()
        image.save(out_dir / f"page-{index + 1:03d}.png")
    print(f"EP{episode}\tpages={len(pdf)}")
