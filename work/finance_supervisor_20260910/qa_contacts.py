from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
R=Path(__file__).resolve().parent;Q=R/'qa_contacts';Q.mkdir(exist_ok=True);font=ImageFont.truetype(r'C:\Windows\Fonts\msyh.ttc',18,index=0)
def num(p):
    try:return int(p.stem.split('-')[-1])
    except:return 0
blank=[];counts={}
for label,folder in [('main',R/'render_main'),('cases',R/'render_cases')]:
    pages=sorted(folder.glob('page-*.png'),key=num);counts[label]=len(pages)
    for i,p in enumerate(pages,1):
        im=Image.open(p).convert('L').resize((180,250));dark=sum(1 for v in im.getdata() if v<235)
        if dark<250:blank.append((label,i,dark))
    for start in range(0,len(pages),16):
        batch=pages[start:start+16];sheet=Image.new('RGB',(1280,1720),'white');dr=ImageDraw.Draw(sheet)
        for j,p in enumerate(batch):
            im=Image.open(p).convert('RGB');im.thumbnail((300,390));x=(j%4)*320+10;y=(j//4)*430+26;sheet.paste(im,(x,y));dr.text((x,y-22),f'{label} {start+j+1}',font=font,fill='black')
        sheet.save(Q/f'{label}_{start+1:03d}_{start+len(batch):03d}.jpg',quality=88)
imgs=sorted((R/'qa_sheets').glob('*.png'))
for start in range(0,len(imgs),12):
    b=imgs[start:start+12];sheet=Image.new('RGB',(1260,1280),'white');dr=ImageDraw.Draw(sheet)
    for j,p in enumerate(b):
        im=Image.open(p).convert('RGB');im.thumbnail((400,280));x=(j%3)*420+10;y=(j//3)*320+28;sheet.paste(im,(x,y));dr.text((x,y-24),p.stem[:40],font=font,fill='black')
    sheet.save(Q/f'xlsx_{start+1:03d}_{start+len(b):03d}.jpg',quality=88)
print({'counts':counts,'xlsx':len(imgs),'blank':blank,'contacts':len(list(Q.glob('*.jpg')))})
