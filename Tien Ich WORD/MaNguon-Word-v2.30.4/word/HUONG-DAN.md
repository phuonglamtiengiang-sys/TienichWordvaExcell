# Tiện ích Word Việt Nam 2.30.4

Tác giả: Hồ Phương Lâm — hophuonglam@outlook.com

## Bộ biểu tượng mới

Phiên bản 1.3 thay toàn bộ biểu tượng trên thanh công cụ và trong menu bằng 24 icon tự vẽ, đồng nhất nét và màu. Icon nền trong suốt 64 px được nhúng trong DLL để hiển thị sắc nét ở các mức thu phóng thông dụng. Không phụ thuộc tên biểu tượng có sẵn của từng bản Office và không tải hình từ Internet.

VI/EN phân biệt hai cách đọc số; AB/ab/Aa phân biệt đổi kiểu chữ; chổi dùng làm sạch, tờ giấy có dấu kiểm dùng chuẩn hóa, hai bảng chữ có mũi tên dùng đổi bảng mã, tờ giấy có thước dùng định dạng tài liệu. VietQR dùng mã QR kèm ký hiệu đồng; mã vạch và QR có biểu tượng riêng; dịch dùng hai ô chữ, cài đặt dịch dùng bánh răng. Các hình mã vạch/QR trên nút là hình minh họa, không phải mã để quét.

Rê chuột vào từng nút để đọc hướng dẫn chi tiết. Tên nút và chức năng được giữ nguyên.

## Bộ chuyển mã UniKey

Từ phiên bản 1.2, TCVN3 và VNI Windows được giải mã bằng phần chuyển C# từ UniKey VnConv/UVConverter 1.1.3b của Phạm Kim Long. Bảng dữ liệu được tạo trực tiếp từ mã nguồn, không dùng bảng gõ tay của bản 1.1. Không cần cài hoặc chạy ứng dụng UniKey. Đây là bản tích hợp C# cho Word, không phải bản UniKey chính thức và không gồm toàn bộ các bảng mã của UVConverter.

Nếu có dấu VNI đứng riêng, công cụ báo lỗi và chưa sửa tài liệu. Các ký tự ngoài bảng byte cũ được giữ lại; vẫn nên chọn riêng từng vùng đúng bảng mã và kiểm tra xem trước. Ctrl+Z hoàn tác như trước.

Gói phát hành có `LICENSE-GPL-2.0.txt`, `THIRD-PARTY-NOTICE.md` và `MaNguon-Word-v2.30.4.zip`, gồm mã nguồn Word, mã nguồn UniKey gốc và hướng dẫn dựng lại. Mã nguồn Word tích hợp được cung cấp theo GPL-2.0-or-later. Nguồn, phiên bản và các thay đổi được ghi trong thông báo đi kèm.

Danh sách nguồn và dịch vụ được tham khảo (UniKey, Microsoft Office, VietQR, BWIP-JS, QRServer, MyMemory và Google Translate web) có trong `THAM-KHAO.md` và gói mã nguồn. Các dịch vụ web là API được gọi khi sử dụng tính năng, không phải mã nguồn được nhúng; xem ghi chú quyền và giấy phép trong các tệp đi kèm.

## Ba công cụ mới trong nhóm Xử lý văn bản

**Chuẩn hóa văn bản:** bôi đen tối đa 100.000 ký tự, bấm nút rồi chọn sửa khoảng trắng, khoảng trắng quanh dấu câu, bỏ dòng trống hoặc nối dòng PDF. Luôn có bản xem trước trước khi Áp dụng. Bỏ dòng trống và nối dòng PDF mặc định tắt. Nối các ngắt dòng mềm; chỉ nối ngắt đoạn khi phía trước là chữ/số/dấu phẩy và phía sau bắt đầu ngay bằng chữ thường. Không thể tự nhận diện chính xác cấu trúc PDF, nên cần kiểm tra xem trước. Dấu chấm được giữ nguyên để tránh sửa nhầm số, dấu viết tắt và địa chỉ web. Tab được giữ, trừ khi nằm trên dòng trống bạn chọn xóa.

**Đổi bảng mã:** bôi đen văn bản lỗi font, chọn nguồn TCVN3 (ABC) hoặc VNI Windows, xem trước rồi áp dụng. Đích là Unicode với phông Times New Roman. Đối với TCVN3 dùng phông chữ hoa như .VnTimeH, bật tùy chọn chữ HOA. Không tự đoán bảng mã; không chọn chuyển mã đối với đoạn Unicode đang đúng hoặc trộn nhiều bảng mã. Nếu có nhiều bảng mã, xử lý riêng từng vùng. Các phần không đổi giữ định dạng; một ký tự VNI ghép lấy định dạng ở vị trí thay thế. Cỡ chữ, đậm/nghiêng không bị đặt lại toàn vùng.

**Định dạng tài liệu:** chọn vùng hoặc toàn bộ nội dung chính, chọn phông/cỡ chữ, giãn dòng theo bội số và thụt dòng đầu theo cm. Phông lấy từ danh sách đang cài trên máy. Giãn dòng và thụt đầu dòng áp dụng cho cả các đoạn giao với vùng chọn. Đổi lề mặc định tắt; khi bật, áp dụng cho tất cả section trong tài liệu, kể cả khi chỉ định dạng chữ ở vùng chọn. Giữ khổ giấy, căn đoạn và đậm/nghiêng. Đây là thiết lập định dạng trực tiếp, không phải mẫu bắt buộc theo quy định pháp luật.

Cả ba công cụ chạy ngoại tuyến và hỗ trợ một lần Ctrl+Z để hoàn tác một lượt áp dụng. Chuẩn hóa văn bản/chuyển mã chỉ nhận chữ trong nội dung chính, không gồm trường, ảnh, điều khiển nội dung hoặc nhiều ô bảng.

## Cài đặt hoặc sửa lỗi không bật được

1. Giải nén thư mục `TienIchWord-v2.30.4` từ gói phát hành.
2. Lưu tài liệu và đóng tất cả cửa sổ Microsoft Word.
3. Chạy `CaiDat-TienIchWord.exe`, bấm **Cài đặt / cập nhật**. Trình cài đã nhúng sẵn DLL add-in và mã nguồn.
4. Mở lại Word và tìm tab **Tiện ích Word**.

Bản 1.0.1 sửa giao tiếp COM lúc khởi động. Bộ cài cập nhật đúng thư mục `%LOCALAPPDATA%\TienIchWord\Addin`, đăng ký cho tài khoản hiện tại và sao lưu/khôi phục riêng mục vô hiệu hóa có tên `TienIchWord.dll`. Không thay đổi bảo mật macro hay các add-in khác.

Nếu tab chưa xuất hiện: File → Options → Add-ins → Manage: COM Add-ins → Go → chọn **Tiện ích Word Việt Nam**. Không dùng nút Browse để chọn trực tiếp DLL; hãy dùng bộ cài. Nếu vẫn lỗi, gửi ảnh cửa sổ COM Add-ins và dòng Load Behavior.

## Sử dụng

- **Đổi số thành chữ:** chọn số nguyên, chọn đơn vị Chỉ đọc số/Đồng/Đô la Mỹ, rồi Tiếng Việt hoặc Tiếng Anh. Chèn phần chữ trong ngoặc sau số gốc. Tiếng Anh chỉ viết hoa đầu câu. Tối đa ±999.999.999.999.999; chưa hỗ trợ số thập phân. Không có xem tỷ giá hoặc quy đổi tiền.
- **Xử lý văn bản:** chọn chữ rồi đổi hoa/thường, hoa đầu câu/đầu từ hoặc làm sạch khoảng trắng. Tối đa 100.000 ký tự mỗi lần.
- **Thanh toán VietQR:** chọn ngân hàng, nhập tài khoản; có thể lưu dùng lần sau. Để trống số tiền nếu muốn người chuyển tự nhập. Kiểm tra tài khoản và nội dung trước khi dùng QR.
- **Mã vạch và QR:** Code 128, Code 39, EAN-13 và QR tiếng Việt. Ảnh nhúng vào tài liệu; QR tối đa 2.000 byte UTF-8, mã vạch tối đa 80 ký tự và có kiểm tra chiều rộng.
- **Dịch ngôn ngữ:** chọn ngôn ngữ và dịch vụ trong Cài đặt dịch, chọn tối đa 20.000 ký tự rồi Dịch vùng chọn. Bản dịch chèn dưới đoạn gốc. MyMemory, Google web hoặc tự động dự phòng khi MyMemory hết hạn mức. Có thể nhập email MyMemory; hạn mức thực tế do nhà cung cấp quyết định. Word không dùng công thức ô như Excel.
- **Giới thiệu:** thông tin tác giả, lời mời một ly cà phê và mã ủng hộ; nhấp đúp mã để phóng to.

Rê chuột trên từng biểu tượng để xem hướng dẫn. Ctrl+Z hoàn tác thay đổi văn bản hoặc ảnh vừa chèn. Chức năng xử lý chữ chỉ hỗ trợ nội dung chính; từ chối vùng bảo vệ, trường, ảnh, điều khiển hoặc dấu kết thúc ô bảng để tránh làm hỏng tài liệu.

## Kết nối và dữ liệu

Tạo VietQR gửi thông tin QR tới `img.vietqr.io`; mã vạch tới `bwipjs-api.metafloor.com`; QR thường tới `api.qrserver.com`. Dịch gửi văn bản được chọn tới MyMemory/Google web sau xác nhận. Các chức năng này cần Internet và phụ thuộc dịch vụ bên ngoài. Thiết lập lưu riêng tại `%LOCALAPPDATA%\TienIchWord`, không dùng chung cấu hình Excel.

Dành cho Microsoft Word trên Windows với .NET Framework 4.x. Đã kiểm tra kết nối trên Word Windows 64-bit tại máy phát triển; chưa kiểm chứng Word 32-bit. Không dành cho Word trên web, macOS hoặc WPS. Bản Excel hiện có không bị sửa đổi.

Để gỡ: đóng Word, chạy bộ cài và bấm **Gỡ khỏi Word**. Cấu hình cá nhân được giữ lại.

## Bản vá đọc ảnh VietQR 2.30.4

Sửa thêm lỗi đổi CHỮ HOA/chữ thường khi chọn nhiều đoạn có đối tượng nổi. Dùng chức năng đổi kiểu chữ của Word để giữ đúng vị trí, định dạng và hỗ trợ Ctrl+Z. Đã tái hiện lỗi bằng bản cũ và kiểm tra bản sửa trên bản sao tài liệu thực tế.

Dịch vụ VietQR có thể trả SVG cho đường dẫn .png. Bản vá nhận diện định dạng thật và nhúng SVG trực tiếp bằng Word, giữ đầy đủ trình bày của dịch vụ. Giữ form VietQR cũ; không cần TienIchVietQR.exe hay công cụ thiết kế bên ngoài. Đã kiểm tra Word trên máy có hỗ trợ SVG và QR trong DOCX lưu lại giải mã đúng. Các bản Office cũ không hỗ trợ SVG cần nâng cấp để tạo loại ảnh này.

Bản 2.30.4 sửa làm sạch, chuẩn hóa và đổi bảng mã khi vùng chọn có neo hình nổi. Vị trí sửa lấy theo ký tự thực của Word, giữ neo hình và định dạng. Chuyển TCVN3/VNI tiếp tục dùng bảng chuyển UniKey. Ctrl+Z hoàn tác toàn bộ lượt sửa.


Bản 2.30.4 đã loại bỏ nút Bỏ dấu tiếng Việt và xử lý bỏ dấu văn bản theo yêu cầu người dùng.

## Cảnh báo vùng chọn lớn (2.30.4)
Trong nhóm Thông tin, chọn Cài đặt add-in. Bật/tắt cảnh báo và đặt ngưỡng 100–100.000 ký tự (mặc định 1.000). Lưu để nhớ cho lần sau. Khi vượt ngưỡng, chọn Không để hủy hoặc Có để tiếp tục. Áp dụng trước xử lý chữ, làm sạch, chuẩn hóa, chuyển bảng mã, đọc số, tạo mã và dịch. Giới hạn mỗi lần: xử lý 100.000 ký tự, dịch 20.000 ký tự. Không xuất hiện chỉ vì bạn bôi đen văn bản; xuất hiện khi bấm chức năng.
