"""Hover instructions shown by Excel for buttons, menus and choices."""
TIPS = {
 'mnuNumberUnit': 'Bấm mũi tên rồi chọn Chỉ đọc số, Đọc đồng hoặc Đọc đô la Mỹ. Lựa chọn được nhớ, dùng cho các lần đổi số và công thức DocSo/DocSoEN không ghi đơn vị riêng.',
 'unitNumber': 'Bấm để chỉ đọc số, không thêm tên tiền. Ví dụ 125 → một trăm hai mươi lăm. Lựa chọn áp dụng cho đơn vị đọc số mặc định.',
 'unitDong': 'Bấm để đọc kèm đơn vị đồng. Ví dụ 125 → một trăm hai mươi lăm đồng. Lựa chọn áp dụng cho đơn vị đọc số mặc định.',
 'unitUSD': 'Bấm để đọc kèm đơn vị đô la Mỹ. Kết quả đọc tiếng Anh dùng US dollars. Lựa chọn áp dụng cho đơn vị đọc số mặc định.',
 'mnuDoiCase': 'Chọn các ô văn bản, bấm mũi tên rồi chọn kiểu viết hoa hoặc viết thường. Kết quả đổi ngay trong các ô đã chọn; các ô có công thức được giữ nguyên. Có thể Hoàn tác.',
 'btnInHoa': 'Chọn các ô văn bản rồi bấm để đổi ngay trong ô. Ví dụ Nguyễn Văn An → NGUYỄN VĂN AN.\nCông thức trong ô khác: =ToUpperUnicode(A2). Có thể Hoàn tác.',
 'btnInThuong': 'Chọn các ô văn bản rồi bấm để đổi ngay trong ô. Ví dụ NGUYỄN VĂN AN → nguyễn văn an.\nCông thức trong ô khác: =ToLowerUnicode(A2). Có thể Hoàn tác.',
 'btnSentenceCase': 'Chọn các ô văn bản rồi bấm nút để viết hoa chữ đầu câu ngay trong ô.\nCông thức trong ô khác: =ToSentenceCase(A2). Có thể Hoàn tác.',
 'btnDauMoiTu': 'Chọn các ô văn bản rồi bấm để viết hoa chữ đầu mỗi từ. Ví dụ nguyễn văn an → Nguyễn Văn An.\nCông thức trong ô khác: =ToTitleCase(A2). Có thể Hoàn tác.',
 'btnXoaDau': 'Chọn các ô văn bản rồi bấm để bỏ dấu ngay trong ô. Ví dụ Nguyễn Văn An → Nguyen Van An.\nCông thức để giữ nguyên ô gốc: =BoDau(A2). Có thể Hoàn tác.',
 'btnClean': 'Chọn các ô văn bản rồi bấm để xóa khoảng trắng thừa và xử lý khoảng trắng ẩn ngay trong ô.\nCông thức để giữ nguyên ô gốc: =LamSachText(A2). Có thể Hoàn tác.',
 'btnTachTen': 'Chọn một cột họ tên, để trống hai cột ngay bên phải rồi bấm nút. Kết quả lần lượt là Họ và tên đệm, Tên.\nVí dụ Nguyễn Văn An → Nguyễn Văn | An.\nCông thức trong ô khác: =TachHoDem(A2) và =TachTen(A2). Có thể Hoàn tác.',
 'mnuBarcodes': 'Chọn một cột mã/văn bản, để trống cột bên phải, bấm mũi tên rồi chọn loại mã. Ảnh được chèn bên phải. Cần Internet; không hỗ trợ Hoàn tác ảnh.',
 'btnCode128': 'Chọn một cột mã sản phẩm, để trống cột bên phải rồi bấm nút. Nhận chữ/số và ký hiệu ASCII, không dùng tiếng Việt có dấu. Ảnh chèn bên phải. Cần Internet; không hỗ trợ Hoàn tác ảnh.',
 'btnCode39': 'Chọn một cột mã, để trống cột bên phải rồi bấm nút. Dùng chữ A–Z, số và ký hiệu Code 39 cho phép; chữ thường đổi sang hoa. Ảnh chèn bên phải. Cần Internet; không hỗ trợ Hoàn tác ảnh.',
 'btnEAN13': 'Chọn một cột mã gồm 12 chữ số hoặc 13 chữ số có số kiểm tra đúng; để trống cột bên phải rồi bấm nút. Nên lưu mã dạng Text để giữ số 0 đầu. Ảnh chèn bên phải. Cần Internet; không hỗ trợ Hoàn tác ảnh.',
 'btnQRCode': 'Chọn một cột văn bản hoặc đường dẫn, để trống cột bên phải rồi bấm nút. Hỗ trợ tiếng Việt có dấu; ảnh QR chèn bên phải. Cần Internet; nội dung gửi tới dịch vụ tạo QR. Không hỗ trợ Hoàn tác ảnh.',
 'btnVietQR': 'Bấm để mở form, chọn ngân hàng và nhập số tài khoản. Có thể nhập số tiền và nội dung chuyển khoản, rồi bấm Tạo mã QR. Để trống số tiền nếu muốn người chuyển tự nhập. Ngân hàng và tài khoản được nhớ cho lần sau. Cần Internet; không hỗ trợ Hoàn tác ảnh.',
 'btnClearBarcode': 'Chọn vùng ô chứa ảnh mã rồi bấm nút. Chỉ xóa ảnh mã do add-in tạo có ô neo trong vùng chọn; giữ nguyên ảnh khác. Không hỗ trợ Hoàn tác thao tác xóa ảnh.',
 'btnAbout': 'Bấm để xem tác giả Hồ Phương Lâm, phiên bản add-in và mã QR mời tác giả một ly cà phê. Có thể phóng to mã QR trong cửa sổ giới thiệu.',
 'btnCompareTables': 'Bấm để mở form.\nSo từng ô và tô màu: chọn bảng A/B, chuyển sheet rồi quét khối. Bấm Tô màu ô khác: ô giống nhau bỏ màu nền, ô khác màu vàng, phần dư màu cam. Tối đa 100.000 ô/vùng. Màu cũ trong vùng được thay theo kết quả so sánh; bấm Hoàn tác để khôi phục.\nGhép theo cột mã: chọn bảng có tiêu đề và cột mã để xuất báo cáo riêng. Có thể chọn hai sheet hoặc hai file đang mở.',
 'btnUndoAddin': 'Bấm để quay lại thao tác gần nhất của add-in, tối đa 20 thao tác trong phiên Excel. Hỗ trợ sửa ô và gỡ màu đối chiếu. Khi gỡ màu, màu nền/quy tắc màu gốc được giữ nguyên. Không hỗ trợ ảnh, báo cáo mới hoặc cài đặt. Giữ file gốc đang mở; đóng Excel sẽ mất lịch sử.',
}

def apply_hover_help(tree, namespace):
    count = 0
    for element in tree.iter():
        if element.tag.rsplit('}', 1)[-1] not in ('button', 'menu', 'toggleButton'):
            continue
        if element.get('id') in TIPS:
            element.set('supertip', TIPS[element.get('id')])
        if element.get('id') == 'btnCompareTables':
            element.set('screentip', 'Đối chiếu hai bảng và tô màu ô khác')
        if not element.get('screentip'):
            element.set('screentip', element.get('label'))
        assert element.get('supertip'), element.get('id')
        assert len(element.get('screentip')) <= 128, element.get('id')
        assert len(element.get('supertip')) <= 1024, element.get('id')
        count += 1
    print('Verified hover instructions for', count, 'buttons/menu choices')
