from pathlib import Path
import subprocess, xml.etree.ElementTree as E
root=Path(__file__).resolve().parent.parent
subprocess.run([__import__('sys').executable,str(root/'branding/build_logo.py')],check=True,cwd=root)
subprocess.run([__import__('sys').executable,str(root/'word/generate_unikey_tables.py')],check=True,cwd=root)
ns='http://schemas.microsoft.com/office/2009/07/customui'
E.register_namespace('',ns)
def el(element,parent=None,**args):
    return E.Element('{'+ns+'}'+element,args) if parent is None else E.SubElement(parent,'{'+ns+'}'+element,args)
ui=el('customUI',onLoad='OnRibbonLoad',loadImage='LoadRibbonImage');r=el('ribbon',ui);tabs=el('tabs',r);tab=el('tab',tabs,id='tabTienIchWord',label='Tiện ích Word')
def group(id,label): return el('group',tab,id=id,label=label)
def button(g,id,label,image,action,help,tag=None):
    args=dict(id=id,label=label,imageMso=image,size='large',onAction=action,screentip=label,supertip=help)
    if g.tag == '{'+ns+'}menu': args.pop('size')
    if tag: args['tag']=tag
    return el('button',g,**args)
g=group('grpWords','Đổi số thành chữ')
m=el('menu',g,id='mnuUnit',label='Đơn vị đọc số',size='large',imageMso='EquationInsertGallery',screentip='Chọn đơn vị đọc số',supertip='Chọn Chỉ đọc số, Đồng hoặc Đô la Mỹ. Nhớ riêng cho Word; không có tỷ giá hay quy đổi tiền.')
for code,label in [('number','Chỉ đọc số'),('dong','Đồng'),('usd','Đô la Mỹ')]: el('toggleButton',m,id='unit_'+code,label=label,tag=code,onAction='OnUnit',getPressed='GetUnitPressed',screentip=label,supertip='Chọn đơn vị '+label+'. Chỉ đổi cách đọc tên tiền, không quy đổi tỷ giá.')
for code,label in [('vi','Tiếng Việt'),('en','Tiếng Anh')]: button(g,'words_'+code,label,'EquationInsertGallery','OnNumber','Bôi đen một số nguyên rồi bấm. Chèn phần chữ trong ngoặc ngay sau số gốc. Hỗ trợ ±999.999.999.999.999. Dấu chấm/phẩy/khoảng trắng chỉ để ngăn nhóm ba chữ số; không nhận số thập phân. Tiếng Anh chỉ viết hoa đầu câu. Ctrl+Z để hoàn tác.',code)
g=group('grpText','Xử lý văn bản')
for code,label in [('upper','CHỮ HOA'),('lower','chữ thường'),('sentence','Hoa đầu câu'),('title','Hoa đầu từ'),('clean','Làm sạch')]: button(g,'text_'+code,label,'ChangeCase' if code!='clean' else 'ClearFormatting','OnText','Chọn đoạn chữ thuần trong nội dung chính rồi bấm '+label+'. Giữ định dạng từng ký tự khi đổi chữ. Làm sạch giữ ngắt đoạn và tab, bỏ khoảng trắng thừa/ẩn. Không xử lý vùng gồm trường, ảnh, điều khiển hoặc dấu kết thúc ô bảng. Ctrl+Z để hoàn tác.',code)
button(g,'btnNormalize','Chuẩn hóa văn bản','ClearFormatting','OnNormalize','Chọn tối đa 100.000 ký tự rồi mở form. Tùy chọn khoảng trắng, dấu câu, dòng trống và nối dòng PDF. Có xem trước; nối dòng PDF cần kiểm tra trước khi áp dụng. Ctrl+Z hoàn tác.')
button(g,'btnEncoding','Đổi bảng mã','ChangeCaseGallery','OnEncoding','Chọn văn bản lỗi font, chọn nguồn TCVN3 hoặc VNI Windows rồi kiểm tra xem trước. Chuyển sang Unicode và phông Times New Roman. Nguồn .VnTimeH chọn tùy chọn chữ HOA. Không tự đoán bảng mã. Ctrl+Z hoàn tác.')
button(g,'btnDocumentFormat','Định dạng tài liệu','Font','OnDocumentFormat','Chọn phông, cỡ chữ, giãn dòng và thụt dòng đầu. Áp dụng vùng chọn hoặc toàn bộ nội dung chính. Đổi lề là tùy chọn cho tất cả section. Không đổi khổ giấy. Ctrl+Z hoàn tác.')
g=group('grpPayment','Thanh toán VietQR')
button(g,'btnVietQR','Thanh toán VietQR','PictureInsertFromFile','OnVietQR','Mở form, chọn ngân hàng, nhập tài khoản; có thể lưu dùng lần sau. Số tiền và nội dung là tùy chọn. Để trống số tiền để người chuyển tự nhập. Dữ liệu gửi img.vietqr.io khi bấm Tạo QR. Nhúng ảnh tại con trỏ/cuối vùng chọn, không thay chữ. Cần Internet; Ctrl+Z để gỡ ảnh vừa chèn.')
g=group('grpCodes','Mã vạch và QR')
m=el('menu',g,id='mnuCodes',label='Tạo mã',size='large',imageMso='PictureInsertFromFile',screentip='Tạo mã vạch hoặc QR',supertip='Đặt con trỏ hoặc chọn nội dung, chọn loại mã rồi kiểm tra trong form. Chèn PNG độ phân giải cao. Cần Internet; nội dung gửi dịch vụ tạo mã. Ctrl+Z để hoàn tác.')
for code,label in [('code128','Code 128'),('code39','Code 39'),('ean13','EAN-13'),('qr','QR tiếng Việt')]: button(m,'code_'+code,label,'PictureInsertFromFile','OnCode','Đặt con trỏ và nhập nội dung tạo '+label+'. Code 128 nhận ASCII; Code 39 nhận chữ hoa/số; EAN-13 kiểm tra mã; QR nhận tiếng Việt tối đa 2.000 byte UTF-8. Mã vạch gửi bwipjs-api.metafloor.com, QR gửi api.qrserver.com. Không tự ép mã vạch quá rộng làm mờ; rút ngắn hoặc dùng QR.',code)
g=group('grpTranslate','Dịch ngôn ngữ')
button(g,'btnTranslate','Dịch vùng chọn','Translate','OnTranslate','Chọn tối đa 20.000 ký tự rồi bấm. Xác nhận gửi tới dịch vụ trong Cài đặt dịch. MyMemory hoặc Google web; Tự động dự phòng khi MyMemory hết hạn mức. Bản dịch chèn dưới đoạn gốc. Có Dừng; lỗi/hủy không chèn dở. Cần Internet; Ctrl+Z để hoàn tác.')
button(g,'btnTranslationSettings','Cài đặt dịch','PropertySheet','OnTranslationSettings','Chọn ngôn ngữ nguồn/đích, MyMemory/Google web/Tự động và email MyMemory tùy chọn, rồi Lưu mặc định. Nhớ riêng cho Word. Lưu chưa dịch; chọn đoạn văn và dùng Dịch vùng chọn.')
g=group('grpAbout','Thông tin')
button(g,'btnHealth','Kiểm tra add-in','Info','OnHealth','Xem phiên bản, trạng thái kết nối Word và xuất nhật ký lỗi. Nhật ký mới không ghi nội dung tài liệu. Nếu gặp lỗi nạp add-in, dùng bộ cài để kiểm tra cài đặt.')
button(g,'btnGeneralSettings','Cài đặt add-in','PropertySheet','OnGeneralSettings','Bật hoặc tắt cảnh báo vùng chọn lớn; đặt ngưỡng 100–100.000 ký tự. Mặc định 1.000. Cảnh báo xuất hiện trước xử lý văn bản, đổi bảng mã, đọc số, tạo mã và dịch. Có thể chọn Hủy để không thực hiện.')
button(g,'btnAbout','Giới thiệu','Info','OnAbout','Xem tác giả Hồ Phương Lâm, email, phiên bản Word và mã mời tác giả một ly cà phê. Nhấp đúp mã để phóng to. Mã nhúng sẵn, không cần Internet.')
icons=[]
for control in ui.iter():
    if control.tag.split('}')[-1] in ('button','menu','toggleButton'):
        icon_id=control.attrib['id']
        control.attrib.pop('imageMso',None)
        control.set('image',icon_id)
        icon_path=root/'word/Assets/icons'/f'{icon_id}.png'
        assert icon_path.is_file(),f'Missing icon: {icon_id}'
        icons.append(icon_path)
(root/'word/ribbon.xml').write_bytes(E.tostring(ui,encoding='utf-8',xml_declaration=True))
subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-File',str(root/'tests/test_word_ribbon_schema.ps1'),'-RibbonPath',str(root/'word/ribbon.xml'),'-SchemaPath',str(root/'word/Assets/customui14.xsd')],check=True,cwd=root)
out=root/'output/Word-2.30.4';out.mkdir(exist_ok=True)
subprocess.run([__import__('sys').executable,str(root/'word/build_installer_icon.py')],check=True,cwd=root)
csc=r'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
common=[csc,'/nologo','/platform:anycpu','/r:System.Drawing.dll','/r:System.Windows.Forms.dll']
subprocess.run(common+['/target:library',f'/out:{out / "TienIchWord.dll"}','/r:System.Core.dll','/r:Microsoft.CSharp.dll','/r:System.Web.Extensions.dll',
    f'/resource:{root / "word/ribbon.xml"},TienIchWord.ribbon.xml',f'/resource:{root / "word/Assets/banks.tsv"},TienIchWord.banks.tsv',f'/resource:{root / "about/support-momo-original.png"},TienIchWord.support-momo.png',f'/resource:{root / "output/branding/TienIchWord-AboutHeader.png"},TienIchWord.word-header.png']+
    [f'/resource:{path},TienIchWord.icons.{path.stem}.png' for path in icons]+
    [str(root/'word'/n) for n in ('WordUtilities.cs','Forms.cs','Registration.cs','TextUtilities.cs','EditingForms.cs','UniKeyDecoder.cs','UniKeyTables.generated.cs','RibbonImages.cs','OfficeImage.cs','Reliability.cs')],check=True,cwd=root)
subprocess.run([__import__('sys').executable,str(root/'word/package_word.py'),'--source-only'],check=True,cwd=root)
installer=out/'CaiDat-TienIchWord-2.30.4-Setup.exe'
subprocess.run(common+['/target:winexe',f'/win32icon:{out / "TienIchWord.ico"}',f'/out:{installer}',f'/resource:{out / "TienIchWord.dll"},TienIchWord.Addin.dll',f'/resource:{out / "MaNguon-Word-v2.30.4.zip"},TienIchWord.Source.zip',f'/resource:{root / "word/THAM-KHAO.md"},TienIchWord.References.md',str(root/'word/Setup.cs')],check=True,cwd=root)
subprocess.run([__import__('sys').executable,str(root/'word/package_word.py')],check=True,cwd=root)
print('Built Word 2.30.4; groups:',len(tab),'custom icons:',len(icons))
