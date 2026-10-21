from pathlib import Path
import fitz,json
from PIL import Image,ImageDraw

r=Path(__file__).resolve().parent
d=fitz.open(r/'3cx_ti_presentation.pdf');assert len(d)==8
out=r/'slides';out.mkdir(exist_ok=True)
contact=Image.new('RGB',(1320,1520),'#e7eceb');draw=ImageDraw.Draw(contact)
for i,p in enumerate(d):
 p.get_pixmap(matrix=fitz.Matrix(1.6,1.6)).save(out/f'{i+1:02}.png')
 im=Image.open(out/f'{i+1:02}.png').convert('RGB');im.thumbnail((640,360))
 x=10+(i%2)*660;y=10+(i//2)*380;contact.paste(im,(x,y));draw.text((x+5,y+363),f'{i+1:02}',fill='#172d38')
contact.save(r/'contact_sheet.png')
print(json.dumps({'pages':len(d),'pngs':8,'pdf_bytes':(r/'3cx_ti_presentation.pdf').stat().st_size}))
