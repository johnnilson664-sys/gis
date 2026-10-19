from pathlib import Path
from PIL import Image,ImageDraw,ImageFont

root=Path('/workspace/work/3cx')
reg='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
bold='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
mono='/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf'
def font(n,b=False,m=False):return ImageFont.truetype(mono if m else bold if b else reg,n)
def text(draw,pos,s,size=30,color='#f3f6fa',b=False,m=False):draw.text(pos,s,font=font(size,b,m),fill=color,spacing=13)
def base():
 im=Image.new('RGB',(1440,810),'#101b29');return im,ImageDraw.Draw(im)

im,d=base()
text(d,(66,45),'КАДР 1 · ОПУБЛИКОВАННЫЙ ИНДИКАТОР · 3CX, 2023',20,'#58d5cd',True)
text(d,(65,108),'Какой именно файл мы проверяем?',42,b=True)
d.rounded_rectangle((65,210,1375,530),radius=20,fill='#1b2b3c')
text(d,(95,242),'ffmpeg.dll',35,b=True,m=True)
text(d,(95,310),'Windows · вариант с подписью 3CX Ltd',25,'#a6b8c9')
text(d,(95,356),'Роль в списке Sophos: изменённый код в точке входа',25,'#a6b8c9')
text(d,(95,425),'SHA-256',18,'#58d5cd',True)
text(d,(95,463),'c485674ee63ec8d4e8fde9800788175a8b02d3f9416d0e763360fff7f8eb4e02',23,m=True)
text(d,(72,590),'Следующий вопрос к своим данным:',27,'#cdf276',True)
text(d,(72,639),'Где присутствовал этот файл и есть ли следы его использования?',25)
text(d,(72,732),'Источник: SophosLabs / IoCs / 3CX IoCs 2023-03.csv. Историческая роль файла.',18,'#a6b8c9')
im.save(root/'demo_source.png')

im,d=base()
text(d,(66,45),'КАДР 2 · СХЕМА ЧТЕНИЯ · ТЕКУЩИХ РЕЗУЛЬТАТОВ НЕТ',20,'#58d5cd',True)
text(d,(65,108),'От поля отчёта — к следующему вопросу',40,b=True)
rows=[('Объект','Совпадает ли полный SHA-256?'),('Дата','Когда анализировали объект?'),('Результаты','Кто и какой ответ дал?'),('Подробности','Тип, подпись, поведение, связи?'),('Возврат внутрь','Что нужно проверить на устройстве?')]
for i,(name,question) in enumerate(rows):
 y=211+i*91
 d.rounded_rectangle((65,y,1375,y+73),radius=10,fill='#1b2b3c')
 text(d,(88,y+20),f'{i+1:02}',22,'#58d5cd',True)
 text(d,(155,y+18),name,26,b=True)
 text(d,(540,y+19),question,25)
text(d,(72,725),'Это учебная схема полей, а не скриншот интерфейса VirusTotal.',22,'#ffb477')
im.save(root/'demo_fields.png')
print('Prepared 2 clearly labelled fallback frames.')
