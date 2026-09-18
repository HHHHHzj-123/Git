from pathlib import Path
import sys
from docx2pdf import convert

src=Path(sys.argv[1]).resolve(); out=Path(sys.argv[2]).resolve(); out.parent.mkdir(parents=True,exist_ok=True)
convert(str(src),str(out))
