from pathlib import Path
import zipfile, re, sys, xml.etree.ElementTree as ET
from ribbon_tips import apply_hover_help
from excel_icons import apply_excel_icons

root=Path(__file__).parent
path=Path(sys.argv[1]).resolve() if len(sys.argv)>1 else root/'output/TienIchExcel-v2.27.xlam'
with zipfile.ZipFile(path) as z:
    contents={n:z.read(n) for n in z.namelist()}
with zipfile.ZipFile(root/'barcode.xlam') as original:
    xml=original.read('customUI/customUI14.xml').decode('utf-8')
xml=xml.replace('<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">','<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_Load">')
menu='''<menu id="mnuNumberUnit" label="Đơn vị đọc số" size="large" imageMso="EquationInsertGallery" screentip="Chọn đơn vị cho DocSo và DocSoEN">
<toggleButton id="unitNumber" label="Chỉ đọc số" tag="number" onAction="Ribbon_SetUnit" getPressed="Ribbon_UnitPressed"/>
<toggleButton id="unitDong" label="Đọc đồng" tag="dong" onAction="Ribbon_SetUnit" getPressed="Ribbon_UnitPressed"/>
<toggleButton id="unitUSD" label="Đọc đô la Mỹ" tag="usd" onAction="Ribbon_SetUnit" getPressed="Ribbon_UnitPressed"/>
</menu>'''
xml=xml.replace('<group id="grpDoiSo" label="Đổi số thành chữ">','<group id="grpDoiSo" label="Đổi số thành chữ">'+menu)
xml=xml.replace('<group id="grpAbout" label="Thông tin">','''<group id="grpCompare" label="Đối chiếu dữ liệu">
<button id="btnCompareTables" label="Đối chiếu hai bảng" size="large" imageMso="FindDialog" onAction="Ribbon_CompareTables" screentip="Đối chiếu hai bảng theo mã" supertip="Chọn bảng A/B, cột mã và các cột cần so sánh. Xuất báo cáo riêng: chỉ có ở A, chỉ có ở B, cùng mã khác thông tin; tách mã trống/trùng để kiểm tra."/>
</group><group id="grpAbout" label="Thông tin">''')
xml=xml.replace('Xóa toàn bộ các ảnh mã vạch và QR Code nằm trong vùng ô được quét chọn.','Chỉ xóa ảnh mã do phiên bản 2 tạo trong các ô đang chọn. Giữ nguyên ảnh và đối tượng khác.')
xml=xml.replace('Quét chọn cột MST, tự động điền Tên công ty và Địa chỉ vào 2 cột bên cạnh dưới dạng text tĩnh.','Tra MST qua Invoy, dự phòng Esgoo, không cần key. Invoy giới hạn công khai 100 lượt/ngày. Hai cột kết quả phải trống; dòng lỗi để trống để có thể thử lại.')
xml=xml.replace('Tách &amp; phân tích CCCD','Phân tích cấu trúc CCCD')
xml=xml.replace('Tỉnh/TP sang 3 cột bên cạnh.','Tên tỉnh theo bảng mã CCCD của bản gốc sang 3 cột bên cạnh. Không xác minh danh tính hoặc tỉnh hiện tại.')
xml=xml.replace('Hoặc dùng công thức: =DocSo(ô_chứa_số)','Công thức: =DocSo(ô_chứa_số). Đơn vị theo cài đặt; dùng đối số thứ hai để đặt riêng.')
xml=xml.replace('Hoặc dùng công thức: =DocSoEN(ô_chứa_số)','Công thức: =DocSoEN(ô_chứa_số). Đơn vị theo cài đặt; dùng đối số thứ hai để đặt riêng.')
xml=xml.replace('<group id="grpAbout" label="Thông tin">','''<group id="grpAddress" label="Địa chỉ hành chính">
<button id="btnConvertAddress" label="Địa chỉ cũ → mới" size="large" imageMso="FindDialog" onAction="Ribbon_ConvertAddress" screentip="Đổi địa chỉ cũ sang mới — dùng nút hoặc công thức" supertip="Chọn cột địa chỉ cũ rồi bấm nút để xuất báo cáo riêng. Các phần xã/phường, quận/huyện, tỉnh/thành cần cách nhau bằng dấu phẩy.&#10;&#10;Công thức trong ô Excel (A2 chứa địa chỉ cũ):&#10;=DiaChiMoi(A2) — trả địa chỉ mới.&#10;=DiaChiMoiTrangThai(A2) — trạng thái hoặc lý do chưa chuyển được.&#10;=DiaChiMoiGoiY(A2) — các đơn vị mới có thể tương ứng.&#10;&#10;Nếu chia tách, trùng tên hoặc chưa khớp, địa chỉ mới để trống; xem trạng thái và gợi ý. Dùng ngoại tuyến, không cần API key. Dữ liệu 06/10/2026, nguồn tinhthanhpho.com."/>
</group><group id="grpAbout" label="Thông tin">''')
xml=xml.replace('Nếu chia tách, trùng tên hoặc chưa khớp, địa chỉ mới để trống; xem trạng thái và gợi ý. Dùng ngoại tuyến, không cần API key. Dữ liệu 06/10/2026, nguồn tinhthanhpho.com.', 'Có bổ sung lịch sử sáp nhập xã năm 2024 và trước tháng 7/2025. Tự nối các lần sáp nhập khi đủ căn cứ.&#10;Nếu chia tách, trùng tên qua các thời kỳ hoặc thiếu mắt xích lịch sử, địa chỉ mới để trống; xem trạng thái và gợi ý. Dùng ngoại tuyến, không cần API key. Nguồn: Cục Thống kê, NQ1191 và tinhthanhpho.com (bản 2026.10.8).')
xml=xml.replace('<group id="grpAbout" label="Thông tin">', '''<group id="grpUndo" label="Hoàn tác"><button id="btnUndoAddin" label="Hoàn tác" size="large" imageMso="Undo" onAction="Ribbon_Undo" screentip="Hoàn tác thao tác sửa ô của add-in" supertip="Lưu tối đa 20 thao tác trong phiên Excel: đổi chữ, làm sạch, tách tên, đọc số, CCCD và tra MST. Chặn nếu ô kết quả đã bị sửa tiếp. Không áp dụng cho QR/barcode, báo cáo mới hoặc cài đặt. Đóng Excel sẽ mất lịch sử."/></group><group id="grpAbout" label="Thông tin">''')
xml=xml.replace('<group id="grpAbout" label="Thông tin">', '''<group id="grpTranslate" label="Dịch ngôn ngữ"><button id="btnTranslateHelp" label="Công thức dịch" size="large" imageMso="Translate" onAction="Ribbon_TranslateHelp" screentip="Dịch văn bản bằng công thức trong ô" supertip="=Dich(A2,&quot;vi&quot;,&quot;en&quot;) — Việt sang Anh. Đảo en/vi để dịch ngược lại. Xem lỗi với DichTrangThai, cùng đối số. Cần Internet; nội dung gửi đến MyMemory. Không cần key; giới hạn miễn phí 5.000 ký tự/ngày và 500 byte UTF-8/yêu cầu. Dịch máy cần kiểm tra lại."/></group><group id="grpAbout" label="Thông tin">''')
xml=xml.replace('<group id="grpTranslate" label="Dịch ngôn ngữ">', '''<group id="grpTranslate" label="Dịch ngôn ngữ"><button id="btnTranslateSelection" label="Dịch ô đã chọn" size="large" imageMso="Translate" onAction="Ribbon_TranslateSelection" screentip="Dịch trực tiếp, không nhập công thức" supertip="Chọn một cột văn bản, chọn ngôn ngữ trong form rồi bấm Dịch. Ghi bản dịch vào cột trống bên phải, giữ nguyên nguồn và hỗ trợ Hoàn tác. Tối đa 100 ô, tổng 5.000 ký tự/lần; tự chia đoạn dài. Nội dung gửi đến MyMemory; hạn mức miễn phí vẫn là 5.000 ký tự/ngày."/>''')
xml=xml.replace('giới hạn miễn phí 5.000 ký tự/ngày và 500 byte UTF-8/yêu cầu.', 'tối đa 5.000 ký tự/ô, tự chia các yêu cầu không quá 500 byte UTF-8; hạn mức miễn phí vẫn 5.000 ký tự/ngày.')
xml=xml.replace('đọc số, CCCD và tra MST.', 'đọc số, CCCD, tra MST và dịch trực tiếp.')
tree=ET.fromstring(xml)
namespace='http://schemas.microsoft.com/office/2009/07/customui'
ET.register_namespace('',namespace)
for group in tree.iter('{'+namespace+'}group'):
    if group.get('id') == 'grpTranslate':
        for button in list(group):
            if button.get('id') == 'btnTranslateHelp':
                group.remove(button)
            elif button.get('id') == 'btnTranslateSelection':
                button.set('label', 'Dịch ngôn ngữ')
                button.set('screentip', 'Chọn ô rồi bấm để dịch ngay')
                button.set('supertip', '1. Chọn các ô văn bản trong một cột, không chọn tiêu đề.\n2. Để trống cột ngay bên phải.\n3. Bấm nút này: bản dịch được ghi sang cột bên phải, giữ nguyên văn bản gốc.\nDùng ngôn ngữ và cách dịch đã lưu trong Cài đặt dịch. Có thể Dừng hoặc Hoàn tác. Mỗi lần tối đa 100 ô, tổng 5.000 ký tự. Cần Internet; nội dung gửi MyMemory hoặc Google web theo lựa chọn.')
        ET.SubElement(group, '{'+namespace+'}button', {'id':'btnTranslationSettings','label':'Cài đặt dịch','size':'large','imageMso':'PropertySheet','onAction':'Ribbon_TranslationSettings','screentip':'Chọn tiếng cần dịch và cách dịch cho lần sau','supertip':'Chọn văn bản đang là tiếng gì, muốn dịch sang tiếng gì và cách dịch: MyMemory, Google web hoặc Tự động. Email là tùy chọn của MyMemory: 5.000 ký tự/ngày khi để trống, 50.000 khi có email hợp lệ theo dịch vụ. Bấm Lưu cài đặt rồi Đóng; lưu chưa bắt đầu dịch. Google web không cần email, có thể bị giới hạn.'})
        assert len(list(group)) == 2, 'Translation group must have translate and settings buttons'
        examples = '\n\nDùng công thức trong ô khác (A2 chứa văn bản):\n=Dich(A2) — theo ngôn ngữ đã lưu.\n=Dich(A2,"vi","en") — Việt sang Anh.\n=Dich(A2,"en","vi") — Anh sang Việt.\n=DichTrangThai(A2,"vi","en") — xem lý do lỗi.\nXem thêm: Cài đặt dịch → Cách dùng công thức.'
        for button in group:
            button.set('supertip', button.get('supertip','') + examples)
groups={element.get('id'):element for element in tree.iter('{'+namespace+'}group')}
address_group=groups['grpAddress']
lookup_group=groups['grpTraCuu']
for button in list(address_group):
    address_group.remove(button)
    lookup_group.append(button)
for parent in tree.iter():
    if address_group in list(parent):
        parent.remove(address_group)
        break
for group in tree.iter('{'+namespace+'}group'):
    if group.get('id') == 'grpCompare':
        group.set('label', 'Công cụ dữ liệu')
        ET.SubElement(group, '{'+namespace+'}button', {'id':'btnExportPDF','label':'Xuất PDF vùng chọn','size':'large','imageMso':'FileSaveAsPdfOrXps','onAction':'Ribbon_ExportSelectionPDF','screentip':'Xuất vùng ô đang chọn thành tệp PDF','supertip':'1. Quét chọn một vùng ô liên tục cần xuất.\n2. Bấm nút và chọn tên tệp, nơi lưu.\n3. PDF được tạo và mở để kiểm tra.\nChỉ xuất vùng chọn trên sheet hiện tại, không kèm các sheet khác. Vừa 1 trang theo chiều ngang; vùng dài chia nhiều trang. Giữ hướng giấy, khổ giấy, lề, đầu/chân trang hiện tại. Tạm bỏ lặp hàng/cột tiêu đề; tự khôi phục thiết lập in sau khi xuất. Dòng/cột ẩn hoặc bị lọc không được in. Có xác nhận trước khi ghi đè tệp. Không áp dụng Hoàn tác cho tệp PDF đã xuất.'})
for group in tree.iter('{'+namespace+'}group'):
    if group.get('id') == 'grpCompare':
        ET.SubElement(group, '{'+namespace+'}button', {'id':'btnPrintSelection','label':'In vùng chọn','size':'large','imageMso':'PrintPreviewAndPrint','onAction':'Ribbon_PrintSelection','screentip':'Xem trước rồi in đúng vùng ô đang chọn','supertip':'1. Chọn riêng một sheet và quét một vùng ô liên tục.\n2. Bấm nút để mở bản xem trước của Excel; chưa gửi lệnh in.\n3. Kiểm tra các trang, bấm Print/In trong bản xem trước để chọn máy in và in; hoặc đóng để hủy.\nVừa 1 trang theo chiều ngang, vùng dài chia nhiều trang. Giữ hướng giấy, khổ giấy, lề và đầu/chân trang hiện tại; muốn đổi hãy đặt trong Page Layout trước khi bấm. Tạm bỏ lặp hàng/cột tiêu đề ngoài vùng chọn. Dòng/cột ẩn hoặc bị lọc không được in. Đóng bản xem trước sẽ khôi phục vùng in và tỷ lệ cũ. Cần máy in hoặc trình điều khiển in hoạt động.'})
for group in tree.iter('{'+namespace+'}group'):
    if group.get('id') == 'grpDoiSo':
        ET.SubElement(group, '{'+namespace+'}button', {'id':'btnViewRates','label':'Xem tỷ giá','size':'large','imageMso':'AccountingFormat','onAction':'Ribbon_ViewRates','screentip':'Mở bảng tỷ giá Vietcombank và sao chép vào Excel','supertip':'Bấm để mở form và tự tải tỷ giá, không cần chọn ô trước. Xem các cột Mã, Tên ngoại tệ, Mua tiền mặt, Mua chuyển khoản, Bán; đơn vị VND cho 1 ngoại tệ.\nChọn dòng, giữ Ctrl/Shift để chọn nhiều dòng rồi bấm Sao chép dòng chọn hoặc Ctrl+C. Bấm Sao chép toàn bảng để chép tất cả. Đóng form, chọn ô đích rồi Ctrl+V để dán thành 5 cột.\nCó thể kèm nguồn và thời điểm công bố/tải khi sao chép. Nguồn chính thức: Vietcombank; tỷ giá tham khảo. Cập nhật tối đa 1 lần/5 phút trong phiên Excel; nút Cập nhật trong thời gian này dùng lại bảng đã tải. Dấu - là không có giá. Tiện ích chỉ xem/sao chép, không tự sửa ô hoặc chuyển đổi số tiền.'})
apply_hover_help(tree, namespace)
for element in tree.iter():
    if element.get('onAction') == 'Ribbon_LookupMST':
        element.set('screentip', 'Tra MST qua Thongtincongty.vn, dự phòng MaSoThue')
        element.set('supertip', 'Chọn một cột MST, để trống 2 cột bên phải. Kết quả: Tên doanh nghiệp | Địa chỉ. Nguồn chính: Địa chỉ theo CQT trên Thongtincongty.vn; dự phòng: Địa chỉ Thuế trên MaSoThue. Không cần key. Chỉ nhận đúng MST và mã chi nhánh, lấy nguyên văn; không dùng gợi ý sáp nhập hoặc tự ghép xã/phường. Thông báo cuối ghi nguồn thực tế. Hai nguồn lỗi/bị chặn: để trống và báo lý do. Dữ liệu có thể chậm cập nhật; đối chiếu chính thức tại tracuunnt.gdt.gov.vn. Tối đa 100 MST/lần, cần Internet. Có thể Hoàn tác.')
xml=ET.tostring(tree,encoding='unicode')
callbacks=set()
for element in tree.iter():
    for key in ('onLoad','onAction','getPressed'):
        if key in element.attrib: callbacks.add(element.attrib[key])
source='\n'.join(p.read_text(encoding='utf-8-sig') for p in (root/'src').glob('*.txt'))
for callback in callbacks:
    assert re.search(r'Public Sub '+re.escape(callback)+r'\b',source),callback
contents['customUI/customUI14.xml']=xml.encode('utf-8')
workbook=contents['xl/workbook.xml'].decode('utf-8')
workbook=re.sub(r'(<workbookPr\b[^>]*?)(/?>)',lambda m: re.sub(r'\sautoCompressPictures="[^"]*"','',m[1])+' autoCompressPictures="0"'+m[2],workbook,count=1)
contents['xl/workbook.xml']=workbook.encode('utf-8')
apply_excel_icons(contents)
temporary=path.with_suffix('.tmp')
with zipfile.ZipFile(temporary,'w',zipfile.ZIP_DEFLATED) as z:
    for name,data in contents.items(): z.writestr(name,data)
temporary.replace(path)
print('Ribbon XML verified; callbacks:',len(callbacks))
