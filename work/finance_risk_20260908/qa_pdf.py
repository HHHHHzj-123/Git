from pathlib import Path
import json
from pypdf import PdfReader
import pdfplumber
from PIL import Image, ImageOps, ImageDraw

base=Path(r'C:\Users\HZJ\Desktop\Git\work\finance_risk_20260908')
pdf=base/'output'/'pdf'/'03-会计职业风险实战手册.pdf'
render_dir=base/'tmp'/'pdfs'/'render-v11'
contact_dir=base/'tmp'/'pdfs'/'contacts-v11'
contact_dir.mkdir(parents=True,exist_ok=True)

reader=PdfReader(str(pdf))
def outline_items(items):
    out=[]
    for x in items:
        if isinstance(x,list): out.extend(outline_items(x))
        else: out.append(getattr(x,'title',str(x)))
    return out
outlines=outline_items(reader.outline)
with pdfplumber.open(str(pdf)) as doc:
    texts=[p.extract_text() or '' for p in doc.pages]
    sizes=[(round(float(p.width),2),round(float(p.height),2)) for p in doc.pages]
flat='\n'.join(texts)
links=0
for page in reader.pages:
    for a in page.get('/Annots',[]):
        obj=a.get_object()
        if obj.get('/Subtype')=='/Link': links+=1

pngs=sorted(render_dir.glob('page-*.png'))
thumb_w=240
per_sheet=12
contacts=[]
for start in range(0,len(pngs),per_sheet):
    batch=pngs[start:start+per_sheet]
    thumbs=[]
    for i,p in enumerate(batch,start+1):
        im=Image.open(p).convert('RGB')
        h=max(1,int(im.height*thumb_w/im.width))
        im=im.resize((thumb_w,h))
        canvas=Image.new('RGB',(thumb_w+12,h+32),'white')
        canvas.paste(im,(6,22))
        d=ImageDraw.Draw(canvas); d.text((8,4),f'Page {i}',fill='black')
        thumbs.append(canvas)
    cell_w=max(x.width for x in thumbs); cell_h=max(x.height for x in thumbs)
    sheet=Image.new('RGB',(cell_w*4,cell_h*3),(228,231,235))
    for j,im in enumerate(thumbs): sheet.paste(im,((j%4)*cell_w,(j//4)*cell_h))
    out=contact_dir/f'contact-{start+1:02d}-{start+len(batch):02d}.png';sheet.save(out);contacts.append(str(out))

result={
    'pages':len(reader.pages),'metadata':dict(reader.metadata or {}),
    'outline_count':len(outlines),'first_outlines':outlines[:20],
    'links':links,'extracted_chars':len(flat),'blank_text_pages':[i+1 for i,t in enumerate(texts) if len(t.strip())<20],
    'all_a4_same':len(set(sizes))==1,'page_size':sizes[0] if sizes else None,
    'search_terms':{x:flat.count(x) for x in ['目录','R05','虚开增值税专用发票','R06','补救SOP','C01','S20']},
    'rendered_pngs':len(pngs),'contacts':contacts
}
(base/'pdf_qa.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=True))
