from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
R=Path(__file__).with_name('rendered_docx')
Q=Path(__file__).with_name('qa_contacts'); Q.mkdir(exist_ok=True)
pages=sorted(R.glob('page-*.png'))
font=ImageFont.truetype(r'C:\Windows\Fonts\msyh.ttc',18,index=0)
blank=[]
for i,p in enumerate(pages,1):
    im=Image.open(p).convert('L').resize((180,250)); dark=sum(1 for v in im.getdata() if v<235)
    if dark<300: blank.append((i,dark))
for start in range(0,len(pages),16):
    batch=pages[start:start+16]; sheet=Image.new('RGB',(4*320,4*430),'white'); dr=ImageDraw.Draw(sheet)
    for j,p in enumerate(batch):
        im=Image.open(p).convert('RGB'); im.thumbnail((300,390)); x=(j%4)*320+10; y=(j//4)*430+25
        sheet.paste(im,(x,y)); dr.text((x,y-22),f'Page {start+j+1}',font=font,fill='black')
    sheet.save(Q/f'doc_{start+1:03d}_{start+len(batch):03d}.jpg',quality=88)
# spreadsheet images in groups
imgs=sorted(Path(__file__).with_name('qa_sheets').glob('*.png'))
for start in range(0,len(imgs),12):
    batch=imgs[start:start+12]; sheet=Image.new('RGB',(3*420,4*320),'white'); dr=ImageDraw.Draw(sheet)
    for j,p in enumerate(batch):
        im=Image.open(p).convert('RGB'); im.thumbnail((400,280)); x=(j%3)*420+10; y=(j//3)*320+30
        sheet.paste(im,(x,y)); dr.text((x,y-25),p.stem[:42],font=font,fill='black')
    sheet.save(Q/f'xlsx_{start+1:03d}_{start+len(batch):03d}.jpg',quality=88)
print({'pages':len(pages),'blank_candidates':blank,'excel_renders':len(imgs),'contacts':len(list(Q.glob('*.jpg')))})
