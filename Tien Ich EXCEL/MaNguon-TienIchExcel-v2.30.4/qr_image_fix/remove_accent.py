from pathlib import Path
p=Path('word/WordUtilities.cs');s=p.read_text(encoding='utf-8-sig');s=s.replace('            if(mode=="accent") return StripAccents(s);\n','')
a=s.index('            if(mode=="accent"&&!original.Any');b=s.index('            plan.Add(edits);ApplyPlan(range,plan,false);',a)
s=s[:a]+'''            if(mode!="clean")throw new ArgumentException("Kiểu xử lý văn bản không hợp lệ.");
            var plan=new TextPlan(original);
            var edits=new List<TextEdit>();
            string[] tokens=TextTools.CleanTokens(original);
            for(int i=0;i<original.Length;i++){
                if(tokens[i]!=original[i].ToString())edits.Add(new TextEdit{Start=i,Length=1,Value=tokens[i]});
            }
'''+s[b:];p.write_text(s,encoding='utf-8-sig')
p=Path('word/build_word.py');s=p.read_text(encoding='utf-8-sig').replace("('accent','Bỏ dấu'),",'').replace("code not in ('accent','clean')","code!='clean'");p.write_text(s,encoding='utf-8-sig')
for n in ['build_word.py','package_word.py','WordUtilities.cs','Forms.cs','Setup.cs','HUONG-DAN.md','BUILD.md']:
 p=Path('word')/n;s=p.read_text(encoding='utf-8-sig').replace('1.3.3','1.3.4');p.write_text(s,encoding='utf-8-sig')
p=Path('word/package_word.py');s=p.read_text(encoding='utf-8-sig');s='\n'.join(line for line in s.split('\n') if 'PASS accent removal' not in line and 'PASS decomposed' not in line);s=s.replace('all 24 icons','all 23 icons').replace('PASS 24 distinct','PASS 23 distinct').replace("'test_word_accent_speed.ps1',",'');p.write_text(s,encoding='utf-8-sig')
for n in ['test_word_icons.ps1','test_word_editing.ps1','test_unikey_tables.py','test_unikey_tables.ps1','test_word_text_positions.ps1']:
 p=Path('tests')/n;s=p.read_text(encoding='utf-8-sig').replace('1.3.3','1.3.4').replace('1.3.2','1.3.4')
 if n=='test_word_icons.ps1': s=s.replace('24','23');s=s.replace("$controls.Count -ne 23","$controls.Count -ne 23")
 if n=='test_word_text_positions.ps1':s=s.replace("@('accent','clean','normalize')","@('clean','normalize')")
 p.write_text(s,encoding='utf-8-sig')
p=Path('word/HUONG-DAN.md');s=p.read_text(encoding='utf-8-sig');s='\n'.join(l for l in s.split('\n') if not l.startswith('Bản 1.3.4 tối ưu Bỏ dấu'));s=s.replace(', bỏ dấu hoặc làm sạch', ' hoặc làm sạch').replace('sửa bỏ dấu, làm sạch','sửa làm sạch');s+='\nBản 1.3.4 đã loại bỏ nút Bỏ dấu tiếng Việt và xử lý bỏ dấu văn bản theo yêu cầu người dùng.\n';p.write_text(s,encoding='utf-8-sig')
