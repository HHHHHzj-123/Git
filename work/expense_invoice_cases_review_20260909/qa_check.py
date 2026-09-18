from pathlib import Path
from PIL import Image,ImageDraw
from zipfile import ZipFile
import xml.etree.ElementTree as ET,json,pdfplumber
R=Path(__file__).parent;Q=R/'qa';ps=sorted((Q/'pages').glob('page-*.png'))
contacts=[]
for k in range(0,len(ps),12):
    c=Image.new('RGB',(1200,1275),'white');d=ImageDraw.Draw(c)
    for i,p in enumerate(ps[k:k+12]):
        im=Image.open(p).convert('RGB');im.thumbnail((290,395));x=(i%4)*300+(300-im.width)//2;y=(i//4)*425+22;c.paste(im,(x,y));d.text(((i%4)*300+6,(i//4)*425+4),p.stem,fill='black')
    out=Q/f'contact_{k//12+1:02d}.jpg';c.save(out,quality=84);contacts.append(str(out))
docx=R/'deliverables'/'企业发票报销与异常支出-分类框架与20案例样稿.docx'
with ZipFile(docx) as z:
    root=ET.fromstring(z.read('word/document.xml'));ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    styles=[]
    for p in root.findall('.//w:p',ns):
        s=p.find('./w:pPr/w:pStyle',ns)
        if s is not None: styles.append(s.get('{%s}val'%ns['w']))
    stats={'paragraphs':len(root.findall('.//w:p',ns)),'tables':len(root.findall('.//w:tbl',ns)),'hyperlinks':len(root.findall('.//w:hyperlink',ns)),'styles':{x:styles.count(x) for x in sorted(set(styles))},'toc':b'TOC' in z.read('word/document.xml')}
with pdfplumber.open(R/'deliverables'/'企业发票报销与异常支出-分类框架与20案例样稿.pdf') as p:
    chars=[len((x.extract_text() or '').strip()) for x in p.pages]
stats.update({'pages':len(chars),'low_text_pages':[i+1 for i,n in enumerate(chars) if n<35],'contacts':contacts})
(Q/'qa.json').write_text(json.dumps(stats,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(stats,ensure_ascii=False,indent=2))
