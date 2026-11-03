from pathlib import Path
p=Path('word/TextUtilities.cs');s=p.read_text(encoding='utf-8-sig');s=s.replace('Vùng sửa cắt giữa một ký tự Unicode. Hãy chọn lại đoạn chữ.");','Vùng sửa cắt giữa một ký tự Unicode. Hãy chọn lại đoạn chữ. "+e.Start+"/"+e.Length+" text="+text.Length+" first="+(parts.Count>0?parts[0].Start:-1));');p.write_text(s,encoding='utf-8-sig')
