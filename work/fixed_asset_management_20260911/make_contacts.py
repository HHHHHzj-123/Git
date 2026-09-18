from PIL import Image,ImageDraw
from pathlib import Path
root=Path(r'C:\Users\HZJ\Desktop\Git\work\fixed_asset_management_20260911');out=root/'qa'/'contacts';out.mkdir(parents=True,exist_ok=True)
def make(files,name,cols=3,w=420):
 ims=[]
 for p in files:
  im=Image.open(p).convert('RGB');im.thumbnail((w,620));ims.append((p,im.copy()))
 rh=max(x.height for _,x in ims)+34;can=Image.new('RGB',(cols*w,((len(ims)+cols-1)//cols)*rh),'white');d=ImageDraw.Draw(can)
 for i,(p,im) in enumerate(ims):
  x=(i%cols)*w;y=(i//cols)*rh;d.text((x+4,y+4),p.stem,fill='black');can.paste(im,(x,y+22))
 can.save(out/name)
docs=sorted((root/'qa'/'doc_render').glob('page-*.png'))
for i in range(0,len(docs),8):make(docs[i:i+8],f'doc_{i+1}_{min(i+8,len(docs))}.png')
xs=sorted([p for p in (root/'qa').glob('*.png')])
for i in range(0,len(xs),5):make(xs[i:i+5],f'xlsx_{i+1}_{min(i+5,len(xs))}.png',2,650)
print({'doc_pages':len(docs),'xlsx_sheets':len(xs),'contacts':len(list(out.glob('*.png')))})
