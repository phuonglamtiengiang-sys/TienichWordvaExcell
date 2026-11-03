from pathlib import Path
for n in ['WordUtilities.cs','Forms.cs','build_word.py','HUONG-DAN.md','BUILD.md','Setup.cs','package_word.py']:
 p=Path('word')/n;s=p.read_text(encoding='utf-8-sig')
 s=s.replace('1.3.5','1.3.6').replace('20000','100000').replace('20.000','100.000')
 if n=='WordUtilities.cs':s=s.replace('text.Length>5000','text.Length>20000').replace('5.000','20.000')
 elif n in ['Forms.cs','build_word.py','HUONG-DAN.md']:s=s.replace('5.000','20.000')
 if n=='HUONG-DAN.md':s=s.replace('Không tăng giới hạn 100.000 ký tự xử lý/20.000 ký tự dịch.','Giới hạn mỗi lần: xử lý 100.000 ký tự, dịch 20.000 ký tự.')
 p.write_text(s,encoding='utf-8-sig')
p=Path('tests/test_word_icons.ps1');s=p.read_text(encoding='utf-8-sig').replace('1.3.5','1.3.6');p.write_text(s,encoding='utf-8-sig')
