from PIL import Image,ImageDraw,ImageFont
from pathlib import Path
root=Path(r'C:\Users\HZJ\Desktop\Git\work\ap_supplier_management_20260911')
out=root/'qa'/'contacts';out.mkdir(parents=True,exist_ok=True)
def make(files,name,cols=3,w=420):
    thumbs=[]
    for p in files:
        im=Image.open(p).convert('RGB'); h=int(im.height*w/im.width); im.thumbnail((w,h)); thumbs.append((p,im.copy()))
    rh=max(im.height for _,im in thumbs)+40; rows=(len(thumbs)+cols-1)//cols
    canvas=Image.new('RGB',(cols*w,rows*rh),'white');d=ImageDraw.Draw(canvas)
    for i,(p,im) in enumerate(thumbs):
        x=(i%cols)*w;y=(i//cols)*rh;canvas.paste(im,(x,y+25));d.text((x+5,y+5),p.stem,fill='black')
    canvas.save(out/name)
docs=sorted((root/'qa'/'doc_render').glob('page-*.png'))
for i in range(0,len(docs),8):make(docs[i:i+8],f'doc_{i+1}_{min(i+8,len(docs))}.png')
xs=[p for p in (root/'qa').glob('*.png') if p.parent.name=='qa']
for i in range(0,len(xs),5):make(xs[i:i+5],f'xlsx_{i+1}_{min(i+5,len(xs))}.png',cols=2,w=650)
print('doc_pages',len(docs),'xlsx_sheets',len(xs),'contacts',len(list(out.glob('*.png'))))
