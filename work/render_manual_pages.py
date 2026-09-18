from pathlib import Path
import math
from PIL import Image, ImageDraw
p=Path(r'C:\Users\HZJ\Desktop\Git\work\fpna-output\docx-render\manual.pdf')
out=p.parent/'pages';out.mkdir(exist_ok=True)
files=sorted(out.glob('page-*.png')); thumbs=[]
for i,f in enumerate(files):
    im=Image.open(f).convert('RGB');im.thumbnail((360,500));c=Image.new('RGB',(380,535),'white');c.paste(im,((380-im.width)//2,25));ImageDraw.Draw(c).text((8,5),f'Page {i+1}',fill='black');thumbs.append(c)
for start in range(0,len(thumbs),12):
    batch=thumbs[start:start+12]; sheet=Image.new('RGB',(1520,535*math.ceil(len(batch)/4)),'#cccccc')
    for j,im in enumerate(batch):sheet.paste(im,((j%4)*380,(j//4)*535))
    sheet.save(out/f'contact-{start//12+1:02d}.jpg')
print('pages',len(files),'contacts',math.ceil(len(files)/12))
