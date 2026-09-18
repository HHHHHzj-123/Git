from PIL import Image, ImageDraw
import glob, os, math
files = glob.glob(r'C:\Users\HZJ\Desktop\Git\work\audit-adjustments-output\preview-*.png')
images=[]
for f in files:
    im=Image.open(f).convert('RGB')
    im.thumbnail((900,500))
    canvas=Image.new('RGB',(920,540),'white')
    canvas.paste(im,(10,30))
    ImageDraw.Draw(canvas).text((10,8),os.path.basename(f),fill='black')
    images.append(canvas)
out=Image.new('RGB',(1840,540*math.ceil(len(images)/2)),'#dddddd')
for i,im in enumerate(images):
    out.paste(im,((i%2)*920,(i//2)*540))
out.save(r'C:\Users\HZJ\Desktop\Git\work\audit-adjustments-output\all-previews.jpg')
