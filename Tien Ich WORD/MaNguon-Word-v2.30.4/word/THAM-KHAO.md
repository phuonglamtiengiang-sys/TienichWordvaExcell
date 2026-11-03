# Nguồn tham khảo – Tiện ích Word Việt Nam 2.30.4

## Mã nguồn được tích hợp

- **UniKey UVConverter / VnConv 1.1.3b** của Phạm Kim Long: mã nguồn C++ gốc được lưu tại `vendor/unikey/upstream/uvconv/`, gói lưu trữ tại `vendor/unikey/uvconv-src-1.1.3b.zip`. Bộ chuyển C# trong add-in là bản chuyển đổi có chọn lọc phần giải mã TCVN3 và VNI Windows. Bảng dữ liệu được tạo từ `vnconv/data.cpp`. UniKey được phân phối theo GPL-2.0-or-later; giữ thông báo tác quyền gốc trong `vendor/unikey/NOTICE.md`, `vendor/unikey/upstream/uvconv/AUTHORS`, `CREDITS` và `gpl.txt`. Add-in không phải sản phẩm chính thức của UniKey.

## Tài liệu nền tảng tham khảo

- Microsoft Learn – Office/VSTO, Word Object Model, .NET Framework Windows Forms và Ribbon XML: https://learn.microsoft.com/visualstudio/vsto/ ; https://learn.microsoft.com/office/vba/api/overview/word ; https://learn.microsoft.com/openspecs/office_standards/ms-customui/
- Microsoft OfficeDev – lược đồ Ribbon Custom UI được dùng để kiểm tra `word/ribbon.xml`: https://github.com/OfficeDev/office-fluent-ui-command-identifiers

## Dịch vụ web được gọi tùy theo tính năng

Các dịch vụ sau là nguồn API/đầu ra, không phải mã nguồn được nhúng trong add-in. Nội dung do người dùng chọn được gửi tới dịch vụ tương ứng khi yêu cầu tạo mã hoặc dịch được thực hiện.

- **VietQR.io** – ảnh VietQR: https://www.vietqr.io/ ; endpoint ảnh `https://img.vietqr.io/image/`.
- **BWIP-JS / bwip-js-api.metafloor.com** – ảnh mã vạch Code 128, Code 39, EAN-13: https://github.com/metafloor/bwip-js ; https://bwipjs-api.metafloor.com/ .
- **QRServer (goQR.me)** – ảnh QR thường: https://goqr.me/api/ .
- **MyMemory** – dịch máy/dịch bộ nhớ: https://mymemory.translated.net/doc/spec.php . Email MyMemory là tùy chọn và được gửi theo giao thức của nhà cung cấp.
- **Google Translate web endpoint** – phương án dịch web không chính thức, có thể thay đổi hoặc bị giới hạn: `https://translate.googleapis.com/translate_a/single`.

## Mã nguồn của tiện ích

Mã C# của Word, giao diện, icon và cấu hình build nằm trong thư mục `word/`; mã chuyển UniKey và nguồn gốc được lưu cùng tại `word/vendor/unikey/`. Mã tiện ích được cung cấp theo GPL-2.0-or-later, xem `LICENSE-GPL-2.0.txt`. Danh sách băm nguồn UniKey nằm ở `vendor/unikey/SHA256SUMS.txt`; ghi chú các thay đổi nằm ở `vendor/unikey/NOTICE.md`.

Việc liệt kê dịch vụ/tài liệu chỉ nhằm ghi nhận nguồn tham khảo và chỉ ra nơi tính năng lấy dữ liệu. Không hàm ý các nhà cung cấp bảo trợ add-in. API, điều khoản và giới hạn có thể thay đổi; vui lòng xem tài liệu hiện hành của nhà cung cấp.
