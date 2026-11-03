"""Original Excel ribbon artwork, using the visual style of the Word icon set.
Embeds PNG resources with standard OPC image relationships, without VBA changes.
"""
from pathlib import Path
import math
import xml.etree.ElementTree as E
from PIL import Image, ImageDraw, ImageFont
from word.draw_ribbon_icons import Canvas,draw,BLUE,TEAL,PURPLE,AMBER,INK,RED,FONT

ROOT=Path(__file__).resolve().parent
FOLDER=ROOT/'excel_icons'
SHARED={
 'mnuNumberUnit':'mnuUnit','unitNumber':'unit_number','unitDong':'unit_dong','unitUSD':'unit_usd',
 'btnNumVN':'words_vi','btnNumEN':'words_en','btnXoaDau':'text_accent','btnInHoa':'text_upper',
 'btnInThuong':'text_lower','btnSentenceCase':'text_sentence','btnDauMoiTu':'text_title',
 'btnClean':'text_clean','btnVietQR':'btnVietQR','mnuBarcodes':'mnuCodes','btnCode128':'code_code128',
 'btnCode39':'code_code39','btnEAN13':'code_ean13','btnQRCode':'code_qr',
 'btnTranslateSelection':'btnTranslate','btnTranslationSettings':'btnTranslationSettings','btnAbout':'btnAbout'}
EXTRA={'btnViewRates','btnTachTen','mnuDoiCase','btnBankList','btnClearBarcode','btnCheckCCCD',
       'btnLookupMST','btnConvertAddress','btnCompareTables','btnExportPDF','btnPrintSelection','btnUndoAddin'}

def grid(c,box,color,mark=None):
 x,y,w,h=box;c.rect((x,y,x+w,y+h),'white',color,1.4,1)
 for a in range(1,3):c.line([(x+a*w/3,y),(x+a*w/3,y+h)],color,.8)
 for b in range(1,3):c.line([(x,y+b*h/3),(x+w,y+b*h/3)],color,.8)
 if mark:c.rect((x+w/3+1,y+h/3+1,x+2*w/3-1,y+2*h/3-1),mark)

def artwork(key):
 if key in SHARED:return draw(SHARED[key])
 c=Canvas()
 if key=='btnViewRates':
  c.rect((2,3,30,29),'white',AMBER,1.5,2);c.text((9,12),'$',13,AMBER);c.text((23,21),'₫',13,TEAL);c.arrow(23,9,True,TEAL);c.arrow(9,23,False,AMBER)
 elif key=='btnTachTen':
  c.circle((12,2,20,10),BLUE);c.rect((8,12,24,16),BLUE,r=2)
  c.line([(16,17),(16,21),(6,21),(6,24)],TEAL,1.4);c.line([(16,21),(26,21),(26,24)],TEAL,1.4)
  c.rect((1,25,13,31),'#eff6ff',BLUE,1.2,1);c.rect((19,25,31,31),'#f0fdfa',TEAL,1.2,1)
 elif key=='mnuDoiCase':
  c.text((10,11),'A',22,BLUE);c.text((24,18),'a',16,TEAL);c.arrow(11,27,True,PURPLE)
 elif key=='btnBankList':
  c.poly([(2,9),(14,2),(26,9)],BLUE);c.line([(2,11),(26,11)],BLUE,2)
  for x in [6,14,22]:c.line([(x,13),(x,23)],BLUE,2.6)
  c.line([(2,25),(20,25)],BLUE,2);c.circle((19,19,28,28),'white',TEAL,1.8);c.line([(27,27),(31,31)],TEAL,2.2)
 elif key=='btnClearBarcode':
  for x,w in [(2,1),(5,2),(9,1),(12,2),(16,1),(19,2),(23,1)]:c.rect((x,3,x+w,23),INK)
  c.circle((16,16,32,32),'white',RED,1.6);c.line([(21,21),(27,27)],RED,2);c.line([(21,27),(27,21)],RED,2)
 elif key=='btnCheckCCCD':
  c.rect((1,5,29,27),'white',TEAL,1.7,3);c.circle((5,10,12,17),TEAL);c.rect((4,19,14,23),TEAL,r=2)
  for y in [11,16,21]:c.line([(17,y),(25,y)],TEAL,1.6)
  c.circle((20,19,32,31),TEAL);c.line([(23,25),(25,27),(29,22)],'white',1.5)
 elif key=='btnLookupMST':
  c.page(PURPLE);c.text((13,14),'#',19,PURPLE);c.circle((18,18,28,28),'white',TEAL,1.8);c.line([(27,27),(31,31)],TEAL,2.2)
 elif key=='btnConvertAddress':
  c.circle((2,2,17,17),TEAL);c.poly([(3,12),(16,12),(9.5,25)],TEAL);c.circle((6,6,13,13),'white')
  c.line([(10,28),(25,28),(25,13)],PURPLE,1.8);c.poly([(21,16),(25,10),(29,16)],PURPLE)
 elif key=='btnCompareTables':
  grid(c,(1,3,17,19),BLUE);grid(c,(14,12,17,19),TEAL,'#fbbf24');c.arrow(23,5,True,PURPLE)
 elif key=='btnExportPDF':
  c.page(RED);c.rect((1,13,31,25),RED,r=2);c.text((16,19),'PDF',10,'white');c.line([(10,29),(24,29)],RED,1)
 elif key=='btnPrintSelection':
  c.rect((7,2,25,12),'white',TEAL,1.5,1);c.rect((2,10,30,25),TEAL,r=3)
  c.circle((24,13,27,16),'#a7f3d0');grid(c,(7,20,18,11),TEAL)
 elif key=='btnUndoAddin':
  points=[]
  for i in range(41):
   a=math.radians(-170+i*6);points.append((17+11*math.cos(a),16+11*math.sin(a)))
  c.line(points,PURPLE,3);c.poly([(2,6),(2,17),(12,14)],PURPLE)
 else:raise KeyError(key)
 return c.im.resize((64,64),Image.Resampling.LANCZOS)

def apply_excel_icons(contents):
 ns='http://schemas.microsoft.com/office/2009/07/customui'
 relns='http://schemas.openxmlformats.org/package/2006/relationships'
 ct='http://schemas.openxmlformats.org/package/2006/content-types'
 E.register_namespace('',ns)
 tree=E.fromstring(contents['customUI/customUI14.xml'])
 relpath='customUI/_rels/customUI14.xml.rels'
 relationships=E.fromstring(contents[relpath]) if relpath in contents else E.Element('{'+relns+'}Relationships')
 controls=[e for e in tree.iter() if e.tag.split('}')[-1] in ('button','menu','toggleButton')]
 ids={e.get('id') for e in controls}
 assert ids==set(SHARED)|EXTRA,('Unmapped controls',ids^(set(SHARED)|EXTRA))
 FOLDER.mkdir(exist_ok=True);(FOLDER/'icons').mkdir(exist_ok=True)
 for control in controls:
  key=control.get('id');rid='img_'+key
  icon=artwork(key);path=FOLDER/'icons'/f'{key}.png';icon.save(path)
  contents['customUI/images/'+path.name]=path.read_bytes()
  control.attrib.pop('imageMso',None);control.attrib.pop('getImage',None);control.set('image',rid)
  for old in list(relationships):
   if old.get('Id')==rid:relationships.remove(old)
  E.SubElement(relationships,'{'+relns+'}Relationship',Id=rid,Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',Target='images/'+path.name)
 content_types=E.fromstring(contents['[Content_Types].xml'])
 if not any(e.get('Extension','').lower()=='png' for e in content_types):E.SubElement(content_types,'{'+ct+'}Default',Extension='png',ContentType='image/png')
 contents['customUI/customUI14.xml']=E.tostring(tree,encoding='utf-8',xml_declaration=True)
 contents[relpath]=E.tostring(relationships,encoding='utf-8',xml_declaration=True)
 contents['[Content_Types].xml']=E.tostring(content_types,encoding='utf-8',xml_declaration=True)
 sheet=Image.new('RGB',(1050,math.ceil(len(controls)/7)*126),'#f8fafc');d=ImageDraw.Draw(sheet)
 for i,control in enumerate(controls):
  x=i%7*150;y=i//7*126;icon=artwork(control.get('id'));sheet.paste(icon,(x+43,y+10),icon)
  label=control.get('label','');font=ImageFont.truetype(str(FONT),11)
  if len(label)>21:
   words=label.split();half=len(words)//2;d.text((x+75,y+88),' '.join(words[:half]),font=font,anchor='mm',fill=INK);d.text((x+75,y+103),' '.join(words[half:]),font=font,anchor='mm',fill=INK)
  else:d.text((x+75,y+91),label,font=font,anchor='mm',fill=INK)
 sheet.save(FOLDER/'preview.png')
 return len(controls)
