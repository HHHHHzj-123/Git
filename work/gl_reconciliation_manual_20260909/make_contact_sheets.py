from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
ROOT=Path(__file__).resolve().parent; qa=ROOT/'qa'
fontp=Path('C:/Windows/Fonts/msyh.ttc'); f=ImageFont.truetype(str(fontp),18) if fontp.exists() else ImageFont.load_default()
def make(paths,out,thumb=(260,368),cols=4):
    rows=(len(paths)+cols-1)//cols; canvas=Image.new('RGB',(cols*thumb[0],rows*(thumb[1]+28)),'#D9D9D9'); d=ImageDraw.Draw(canvas)
    for i,p in enumerate(paths):
        im=Image.open(p).convert('RGB'); im.thumbnail(thumb); x=(i%cols)*thumb[0]+(thumb[0]-im.width)//2; y=(i//cols)*(thumb[1]+28); canvas.paste(im,(x,y)); d.text((i%cols*thumb[0]+6,y+thumb[1]+3),p.stem,fill='black',font=f)
    canvas.save(out)
pages=sorted((qa/'doc-pages').glob('page-*.png'))
for k in range(0,len(pages),12): make(pages[k:k+12],qa/f'doc-contact-{k//12+1}.png')
pages3=sorted((qa/'doc-pages-final3').glob('page-*.png'))
for k in range(0,len(pages3),15): make(pages3[k:k+15],qa/f'doc-final-contact-{k//15+1}.png',thumb=(220,312),cols=5)
books=[p for p in qa.glob('*.png') if not p.name.startswith('doc-contact')]
if books: make(books,qa/'workbook-contact.png',thumb=(420,230),cols=2)
