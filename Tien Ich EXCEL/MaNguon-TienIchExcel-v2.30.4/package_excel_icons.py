from pathlib import Path
import zipfile

root=Path(__file__).resolve().parent
out=root/'output'
intro='''# Excel 2.28 — Bộ biểu tượng mới

Thay 33 biểu tượng của thanh Tiện ích và các menu bằng icon tự vẽ nền trong suốt, đồng bộ với bản Word. Các chức năng, công thức, bố cục và hướng dẫn khi rê chuột giữ nguyên. Icon được nhúng trực tiếp trong XLAM, không cần thư mục ảnh đi kèm hoặc Internet.

## Cập nhật

1. Lưu các workbook đang làm và đóng Excel.
2. Giải nén gói vào thư mục cố định.
3. Mở Excel → File → Options → Add-ins → Manage: Excel Add-ins → Go.
4. Bỏ chọn bản Tiện ích Excel cũ để không nạp hai phiên bản cùng lúc.
5. Bấm Browse và chọn TienIchExcel-v2.28.xlam trong thư mục đã giải nén, đánh dấu bản mới rồi OK.
6. Mở tab Tiện ích để dùng. Nếu máy yêu cầu quyền chạy macro, thực hiện theo quy trình tin cậy tệp của bạn; bản cập nhật không thay đổi thiết lập bảo mật.

Bản 2.28 chỉ cập nhật giao diện. Mã VBA và các phần workbook khác giữ nguyên từng byte từ 2.27; thông tin phiên bản trong phần lõi/Giới thiệu có thể vẫn là phiên bản trước. Không đưa bộ chuyển mã UniKey của bản Word vào Excel trong lần cập nhật icon này.

Gói có ảnh tổng hợp biểu tượng, kết quả kiểm tra và mã vẽ icon. Mã vẽ dùng Python/Pillow, phông Segoe UI có sẵn trên Windows. Không phân phối phông chữ. Các icon và mã vẽ gốc được cung cấp theo GPL-2.0-or-later cùng văn bản giấy phép; việc này không thay đổi mã VBA trong XLAM.

'''
source=out/'MaNguon-Icon-Excel.zip'
with zipfile.ZipFile(source,'w',zipfile.ZIP_DEFLATED) as z:
 for name in ['excel_icons.py','build_excel_icons.py','package_excel_icons.py','word/draw_ribbon_icons.py','word/LICENSE-GPL-2.0.txt','output/TienIchExcel-v2.27.xlam']:
  z.write(root/name,name)
 z.writestr('README.md','Build: install Pillow, run python build_excel_icons.py from the archive root. Output: output/TienIchExcel-v2.28.xlam. The original 2.27 binary is supplied for an exact icon-only rebuild. Source changes only Ribbon images and their package relationships. See word/LICENSE-GPL-2.0.txt for drawing code/artwork license. No fonts are included.\n')
with zipfile.ZipFile(out/'TienIchExcel-v2.27.zip') as previous,zipfile.ZipFile(out/'TienIchExcel-v2.28.zip','w',zipfile.ZIP_DEFLATED) as z:
 for name in previous.namelist():
  if name.endswith('.xlam'):continue
  if name=='HUONG-DAN.md':
   z.writestr(name,(intro+previous.read(name).decode('utf-8-sig')).encode('utf-8-sig'))
  else:z.writestr(name,previous.read(name))
 z.write(out/'TienIchExcel-v2.28.xlam','TienIchExcel-v2.28.xlam')
 z.write(root/'excel_icons/preview.png','BO-ICON-EXCEL.png')
 z.write(out/'excel-icon-results.txt','KIEM-TRA-ICON.txt')
 z.write(source,source.name)
 z.write(root/'word/LICENSE-GPL-2.0.txt','GIAY-PHEP-MA-VE-ICON.txt')
print(out/'TienIchExcel-v2.28.zip')
