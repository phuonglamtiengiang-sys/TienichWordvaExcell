from pathlib import Path
import zipfile
root=Path(__file__).resolve().parent.parent
src=root/'output/TienIchExcel-v2.28.zip'
dst=root/'output/TienIchExcel-v2.28.1.zip'
with zipfile.ZipFile(src) as old, zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED) as out:
    for item in old.infolist():
        if item.filename.lower().endswith('.xlam') or item.filename.endswith('KIEM-TRA-ICON.txt'): continue
        data=old.read(item.filename)
        if item.filename.endswith('HUONG-DAN.md'):
            data=('Bản vá 2.28.1: nhận đúng ảnh SVG/PNG từ VietQR. Giữ form cũ, không cần EXE phụ.\nĐóng Excel, giải nén gói vào thư mục cố định rồi chọn TienIchExcel-v2.28.1.xlam trong File > Options > Add-ins > Go > Browse.\nNguồn: https://www.vietqr.io/qr-dong\n\n').encode('utf-8')+data
        out.writestr(item.filename,data)
    out.write(root/'output/TienIchExcel-v2.28.1.xlam','TienIchExcel-v2.28.1.xlam')
    for name in ['results.txt','decode-results.txt']:
        out.write(root/'qr_image_fix'/name,'KIEM-TRA-'+name)
    out.write(root/'qr_image_fix/mod_QRImageFormat.vba','MaNguonBanVa/mod_QRImageFormat.bas')
    out.write(root/'src/mod_Core.txt','MaNguonBanVa/mod_Core.txt')
print(dst)
