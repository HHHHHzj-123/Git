from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
R=Path(__file__).resolve().parent;Q=R/'contacts';Q.mkdir(exist_ok=True);font=ImageFont.truetype(r'C:\Windows\Fonts\msyh.ttc',16,index=0)
def no(p):
 try:return int(p.stem.split('-')[-1])
 except:return 0
counts={};blank=[]
for label in ['interview','cases','onboard']:
 ps=sorted((R/f'render_{label}').glob('page-*.png'),key=no);counts[label]=len(ps)
 for i,p in enumerate(ps,1):
  im=Image.open(p).convert('L').resize((160,220));dark=sum(1 for v in im.getdata() if v<235)
  if dark<220:blank.append((label,i,dark))
 for st in range(0,len(ps),20):
  b=ps[st:st+20];sh=Image.new('RGB',(1250,1600),'white');dr=ImageDraw.Draw(sh)
  for j,p in enumerate(b):
   im=Image.open(p).convert('RGB');im.thumbnail((235,350));x=(j%5)*250+8;y=(j//5)*400+25;sh.paste(im,(x,y));dr.text((x,y-20),f'{label} {st+j+1}',font=font,fill='black')
  sh.save(Q/f'{label}_{st+1:03d}_{st+len(b):03d}.jpg',quality=86)
print({'counts':counts,'blank':blank,'contacts':len(list(Q.glob('*.jpg')))})
