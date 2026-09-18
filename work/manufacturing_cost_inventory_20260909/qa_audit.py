from pathlib import Path
from PIL import Image, ImageOps, ImageDraw
from zipfile import ZipFile
import json, re, xml.etree.ElementTree as ET
import pdfplumber

ROOT=Path(__file__).parent
QA=ROOT/'qa'
OUT=ROOT/'deliverables'

def contacts(paths, prefix, cols, rows, thumb=(340,480)):
    paths=list(paths)
    pages=[]
    for batch_start in range(0,len(paths),cols*rows):
        batch=paths[batch_start:batch_start+cols*rows]
        canvas=Image.new('RGB',(cols*thumb[0],rows*thumb[1]),'white')
        draw=ImageDraw.Draw(canvas)
        for i,p in enumerate(batch):
            im=Image.open(p).convert('RGB')
            im.thumbnail((thumb[0]-10,thumb[1]-28))
            x=(i%cols)*thumb[0]+(thumb[0]-im.width)//2
            y=(i//cols)*thumb[1]+22
            canvas.paste(im,(x,y))
            draw.text(((i%cols)*thumb[0]+8,(i//cols)*thumb[1]+4),p.stem,fill='black')
        q=QA/f'{prefix}_{len(pages)+1:02d}.jpg'
        canvas.save(q,quality=82)
        pages.append(q)
    return pages

doc_pngs=sorted((QA/'doc_pages').glob('page-*.png'))
doc_contacts=contacts(doc_pngs,'doc_contact',4,4,(300,425))
xls_pngs=sorted([p for p in QA.glob('*.png')])
xls_contacts=contacts(xls_pngs,'xlsx_contact',4,5,(390,255))

docx=OUT/'制造业成本与存货完整核算实操手册.docx'
with ZipFile(docx) as z:
    xml=z.read('word/document.xml')
    root=ET.fromstring(xml)
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    headings=[]
    pstyle_values=[]
    for p in root.findall('.//w:p',ns):
        sty=p.find('./w:pPr/w:pStyle',ns)
        if sty is not None:
            pstyle_values.append(sty.get('{%s}val'%ns['w']) or '')
        if sty is not None and ((sty.get('{%s}val'%ns['w']) or '').startswith('Heading') or (sty.get('{%s}val'%ns['w']) or '') in ('1','2','3')):
            txt=''.join(t.text or '' for t in p.findall('.//w:t',ns))
            headings.append((sty.get('{%s}val'%ns['w']),txt))
    doc_stats={
        'paragraphs':len(root.findall('.//w:p',ns)),
        'tables':len(root.findall('.//w:tbl',ns)),
        'drawings':len(root.findall('.//w:drawing',ns)),
        'headings':len(headings),
        'toc_field':b'TOC' in xml,
        'hyperlinks':len(root.findall('.//w:hyperlink',ns)),
        'heading_sample':headings[:12],
        'pstyle_values':sorted(set(pstyle_values)),
    }

pdf=OUT/'制造业成本与存货完整核算实操手册.pdf'
with pdfplumber.open(pdf) as p:
    chars=[len((x.extract_text() or '').strip()) for x in p.pages]
pdf_stats={'pages':len(chars),'min_chars':min(chars),'low_text_pages':[i+1 for i,n in enumerate(chars) if n<35], 'char_counts':chars}

xmind=OUT/'制造业成本与存货核算全景.xmind'
with ZipFile(xmind) as z:
    content=json.loads(z.read('content.json'))
    def count_topic(t):
        total=1
        for group in (t.get('children') or {}).values():
            if isinstance(group,list):
                total+=sum(count_topic(c) for c in group)
        return total
    xmind_stats={'sheets':len(content),'sheet_titles':[s['title'] for s in content], 'topics':[count_topic(s['rootTopic']) for s in content], 'zip_entries':len(z.namelist())}

xlsx=OUT/'成本与存货月结实操工具.xlsx'
with ZipFile(xlsx) as z:
    wbxml=ET.fromstring(z.read('xl/workbook.xml'))
    nsw={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sheet_names=[s.get('name') for s in wbxml.findall('.//m:sheets/m:sheet',nsw)]
    formula_errors=[]
    for name in z.namelist():
        if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
            raw=z.read(name).decode('utf-8','ignore')
            if any(e in raw for e in ['#VALUE!','#REF!','#DIV/0!','#NAME?']): formula_errors.append(name)
    xlsx_stats={'sheets':len(sheet_names),'sheet_names':sheet_names,'formula_error_xmls':formula_errors,'zip_entries':len(z.namelist())}

report={'docx':doc_stats,'pdf':pdf_stats,'xmind':xmind_stats,'xlsx':xlsx_stats,'contacts':{'doc':[str(x) for x in doc_contacts],'xlsx':[str(x) for x in xls_contacts]}}
(QA/'audit_report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False,indent=2))
