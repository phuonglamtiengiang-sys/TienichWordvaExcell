"""Original ribbon artwork for Tien Ich Word. GPL-2.0-or-later.
Draw at 4x logical size and export transparent 64px PNGs for high-DPI Office.
Requires Pillow; fonts are read from Windows, not redistributed.
"""
from pathlib import Path
import math
from PIL import Image, ImageDraw, ImageFont

root=Path(__file__).resolve().parent
dest=root/'Assets/icons';dest.mkdir(parents=True,exist_ok=True)
S=4
BLUE='#2563eb';TEAL='#0d9488';PURPLE='#7c3aed';AMBER='#b45309';INK='#243b53';RED='#dc2626'
FONT=Path('C:/Windows/Fonts/seguisb.ttf')
LABELS={
 'mnuUnit':'Đơn vị đọc số','unit_number':'Chỉ đọc số','unit_dong':'Đồng','unit_usd':'Đô la Mỹ',
 'words_vi':'Đọc tiếng Việt','words_en':'Đọc tiếng Anh','text_upper':'CHỮ HOA','text_lower':'chữ thường',
 'text_sentence':'Hoa đầu câu','text_title':'Hoa đầu từ','text_clean':'Làm sạch',
 'btnNormalize':'Chuẩn hóa văn bản','btnEncoding':'Đổi bảng mã','btnDocumentFormat':'Định dạng tài liệu',
 'btnVietQR':'Thanh toán VietQR','mnuCodes':'Tạo mã','code_code128':'Code 128','code_code39':'Code 39','code_ean13':'EAN-13',
 'code_qr':'QR tiếng Việt','btnTranslate':'Dịch vùng chọn','btnTranslationSettings':'Cài đặt dịch','btnHealth':'Kiểm tra add-in','btnGeneralSettings':'Cài đặt add-in','btnAbout':'Giới thiệu'}

class Canvas:
 def __init__(self):self.im=Image.new('RGBA',(32*S,32*S));self.d=ImageDraw.Draw(self.im)
 def rect(self,box,fill=None,outline=None,width=1,r=0):
  box=tuple(round(x*S) for x in box)
  if r:self.d.rounded_rectangle(box,r*S,fill,outline,width=round(width*S))
  else:self.d.rectangle(box,fill,outline,width=round(width*S))
 def line(self,points,fill,width=1.8):self.d.line([(round(x*S),round(y*S)) for x,y in points],fill,width=round(width*S),joint='curve')
 def poly(self,points,fill):self.d.polygon([(round(x*S),round(y*S)) for x,y in points],fill)
 def circle(self,box,fill=None,outline=None,width=1.8):self.d.ellipse(tuple(round(x*S) for x in box),fill,outline,width=round(width*S))
 def text(self,xy,text,size=14,fill=INK):self.d.text((xy[0]*S,xy[1]*S),text,font=ImageFont.truetype(str(FONT),round(size*S)),fill=fill,anchor='mm',stroke_width=0)
 def arrow(self,x,y,right=True,color=BLUE):
  dx=1 if right else -1;self.line([(x-5*dx,y),(x+5*dx,y)],color,1.8);self.line([(x+1*dx,y-4),(x+5*dx,y),(x+1*dx,y+4)],color,1.8)
 def page(self,color=BLUE):
  self.rect((4,2,27,30),'#ffffff',color,1.7,2);self.line([(21,2),(21,8),(27,8)],color,1.5)
 def qr(self,x=3,y=3,size=25,color=INK):
  u=size/9
  for a,b in [(0,0),(6,0),(0,6)]:
   self.rect((x+a*u,y+b*u,x+(a+3)*u,y+(b+3)*u),color)
   self.rect((x+(a+.65)*u,y+(b+.65)*u,x+(a+2.35)*u,y+(b+2.35)*u),'white')
   self.rect((x+(a+1.1)*u,y+(b+1.1)*u,x+(a+1.9)*u,y+(b+1.9)*u),color)
  for a,b in [(4,0),(4,2),(3,4),(5,4),(7,4),(8,5),(4,6),(6,6),(8,7),(5,8),(7,8),(6,3)]:self.rect((x+a*u,y+b*u,x+(a+.85)*u,y+(b+.85)*u),color)

def draw(key):
 c=Canvas()
 if key=='mnuUnit':
  c.circle((2,10,21,29),'#fff7ed',AMBER);c.circle((10,2,29,21),'#fffbeb',AMBER);c.text((19,11),'$',15,AMBER)
 elif key.startswith('unit_'):
  color=TEAL if key=='unit_dong' else PURPLE if key=='unit_usd' else BLUE
  c.circle((3,3,29,29),color);c.text((16,15),{'unit_number':'123','unit_dong':'₫','unit_usd':'$'}[key],13 if key=='unit_number' else 23,'white')
 elif key.startswith('words_'):
  color=TEAL if key=='words_vi' else PURPLE;c.page(color);c.text((15,11),'123',8,color);c.text((15,21),'VI' if key=='words_vi' else 'EN',13,color)
 elif key in ('text_upper','text_lower','text_sentence','text_title'):
  color=BLUE
  if key=='text_title':c.text((11,9),'Aa',13,color);c.text((21,23),'Bb',13,color)
  else:
   c.text((15,15),{'text_upper':'AB','text_lower':'ab','text_sentence':'Aa'}[key],19,color)
   if key=='text_sentence':c.circle((27,23,30,26),TEAL)
   else:
    y=27 if key=='text_upper' else 4;c.line([(5,y),(27,y)],TEAL,2);c.poly([(26,y-4),(30,y),(26,y+4)],TEAL)
 elif key=='text_accent':
  c.text((13,17),'Á',26,TEAL);c.circle((21,2,31,12),RED);c.line([(23,7),(29,7)],'white',1.8)
 elif key=='text_clean':
  c.line([(12,20),(25,4)],TEAL,5);c.poly([(6,15),(17,24),(11,30),(1,22)],TEAL);c.line([(4,22),(11,27)],'white',1.4)
  for x,y in [(6,5),(27,20)]:c.line([(x-3,y),(x+3,y)],AMBER,1.5);c.line([(x,y-3),(x,y+3)],AMBER,1.5)
 elif key=='btnNormalize':
  c.page(TEAL)
  for y,w in [(10,12),(15,9),(20,10)]:c.line([(8,y),(8+w,y)],TEAL,1.7)
  c.circle((18,18,32,32),TEAL);c.line([(21,25),(24,28),(29,22)],'white',1.8)
 elif key=='btnEncoding':
  c.rect((1,3,15,20),'#f5f3ff',PURPLE,1.5,2);c.text((8,11),'Â',12,PURPLE)
  c.rect((17,13,31,30),'#eff6ff',BLUE,1.5,2);c.text((24,21),'U',12,BLUE);c.arrow(23,6,True,TEAL);c.arrow(8,26,False,TEAL)
 elif key=='btnDocumentFormat':
  c.page(PURPLE);c.line([(9,7),(9,26)],'#c4b5fd',1.5);c.line([(22,10),(22,26)],'#c4b5fd',1.5)
  for y,x in [(12,12),(17,11),(22,11)]:c.line([(x,y),(20,y)],PURPLE,1.7)
  c.rect((0,26,30,31),'#fef3c7',AMBER,1)
  for x in range(3,29,4):c.line([(x,27),(x,29.5)],AMBER,1)
 elif key=='btnVietQR':
  c.qr(1,2,25,TEAL);c.circle((17,16,32,31),'white',TEAL,1.5);c.text((24.5,23),'₫',12,TEAL)
 elif key=='mnuCodes' or key.startswith('code_'):
  if key=='code_qr':c.qr(3,3,26,BLUE)
  else:
   bars=[1,2,1,1,3,1,2,1,2];x=3
   for w in bars:c.rect((x,4,x+w*.65,22),INK);x+=w*.65+1
   if key=='mnuCodes':c.qr(16,15,16,BLUE)
   else:c.text((16,28),{'code_code128':'128','code_code39':'39','code_ean13':'13'}[key],8,BLUE)
 elif key=='btnTranslate':
  c.rect((1,2,21,22),BLUE,r=3);c.poly([(6,20),(6,27),(12,20)],BLUE);c.text((11,11),'A',15,'white')
  c.rect((13,13,31,30),'white',TEAL,1.8,3);c.text((22,21),'Đ',13,TEAL)
 elif key=='btnTranslationSettings':
  for i in range(8):
   a=i*math.pi/4;c.line([(16+9*math.cos(a),15+9*math.sin(a)),(16+13*math.cos(a),15+13*math.sin(a))],PURPLE,5)
  c.circle((6,5,26,25),PURPLE);c.circle((12,11,20,19),'white');c.circle((20,20,32,32),TEAL);c.text((26,25),'A',9,'white')
 elif key=='btnGeneralSettings':
  for y,x in [(7,11),(16,22),(25,15)]:
   c.line([(3,y),(29,y)],BLUE,2.5);c.circle((x-3,y-3,x+3,y+3),'white',BLUE,1.6)
 elif key=='btnHealth':
  c.page(TEAL);c.line([(8,16),(12,16),(15,9),(18,23),(21,16),(25,16)],TEAL,2)
 elif key=='btnAbout':
  c.circle((2,2,30,30),BLUE);c.circle((14,7,18,11),'white');c.line([(16,14),(16,24)],'white',3)
 else:raise ValueError(key)
 return c.im.resize((64,64),Image.Resampling.LANCZOS)

for key in LABELS:draw(key).save(dest/(key+'.png'))
sheet=Image.new('RGB',(900,math.ceil(len(LABELS)/6)*128),'#f8fafc');d=ImageDraw.Draw(sheet)
for i,(key,label) in enumerate(LABELS.items()):
 x=(i%6)*150;y=(i//6)*128;icon=draw(key);sheet.paste(icon,(x+43,y+12),icon)
 d.text((x+75,y+90),label,font=ImageFont.truetype(str(FONT),12),anchor='mm',fill=INK)
sheet.save(root/'Assets/ribbon-icons-preview.png')
print('Drew',len(LABELS),'original transparent icons')
