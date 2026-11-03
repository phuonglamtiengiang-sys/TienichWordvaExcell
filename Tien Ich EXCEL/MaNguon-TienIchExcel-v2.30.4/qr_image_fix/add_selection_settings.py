from pathlib import Path
p=Path('word/WordUtilities.cs');s=p.read_text(encoding='utf-8-sig').replace('public string Unit =','public bool WarnLargeSelection=true;\n        public int SelectionWarningLimit=1000;\n        public bool ShouldWarn(int count){return WarnLargeSelection&&count>Math.Max(100,Math.Min(20000,SelectionWarningLimit));}\n        public string Unit =')
s=s.replace('"Đã dừng; chưa chèn bản dịch.","Dịch ngôn ngữ"','"Đã hủy thao tác.","Tiện ích Word"')
a=s.index('        public void OnText(object control)');s=s[:a]+'''        object CheckedSelection(WordEngine engine,bool requireText){
            object range=engine.EditableSelection(requireText);int count=((string)((dynamic)range).Text??"").Length;
            if(prefs.ShouldWarn(count)&&MessageBox.Show(Owner(),"Bạn đang chọn "+count.ToString("N0")+" ký tự, vượt ngưỡng cảnh báo "+prefs.SelectionWarningLimit.ToString("N0")+".\\n\\nThao tác trên vùng lớn có thể làm Word chậm. Nên chia thành từng đoạn nhỏ. Bạn có muốn tiếp tục?\\n\\nCó thể đổi ngưỡng trong Cài đặt add-in.","Vùng chọn nhiều văn bản",MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2)!=DialogResult.Yes)throw new OperationCanceledException();
            return range;
        }
        public void OnGeneralSettings(object control){Run(()=>{using(var form=new GeneralSettingsForm(prefs))form.ShowDialog(Owner());});}
''' + s[a:]
a=s.index('        public void OnText(object control)');s=s[:a]+s[a:].replace('engine.EditableSelection(true)','CheckedSelection(engine,true)').replace('engine.EditableSelection(false)','CheckedSelection(engine,false)')
p.write_text(s,encoding='utf-8-sig')
p=Path('word/Forms.cs');s=p.read_text(encoding='utf-8-sig');a=s.index('    public sealed class Bank');s=s[:a]+'''    public sealed class GeneralSettingsForm:UtilityForm {
        public GeneralSettingsForm(Preferences prefs):base("Cài đặt add-in Word",560,290){
            var enabled=new CheckBox{Text="Cảnh báo khi chọn nhiều văn bản",Left=20,Top=20,Width=510,Checked=prefs.WarnLargeSelection};Controls.Add(enabled);
            LabelAt("Hiện cảnh báo khi vùng chọn vượt số ký tự:",20,65,510,28);
            var limit=new NumericUpDown{Left=20,Top=98,Width=200,Minimum=100,Maximum=20000,Increment=100,Value=Math.Max(100,Math.Min(20000,prefs.SelectionWarningLimit)),ThousandsSeparator=true};Controls.Add(limit);
            LabelAt("Mặc định: 1.000 ký tự. Cảnh báo cho phép tiếp tục hoặc hủy. Giới hạn xử lý 20.000 ký tự, dịch 5.000 ký tự vẫn áp dụng. Cài đặt được nhớ cho lần sau.",20,140,515,70);
            ButtonAt("Lưu",20,235,250,(s,e)=>{prefs.WarnLargeSelection=enabled.Checked;prefs.SelectionWarningLimit=(int)limit.Value;try{prefs.Save();Close();}catch(Exception ex){MessageBox.Show(this,ex.Message,"Không lưu được cài đặt");}});
            ButtonAt("Đóng",290,235,250,(s,e)=>Close());
        }
    }
''' + s[a:];p.write_text(s,encoding='utf-8-sig')
p=Path('word/build_word.py');s=p.read_text(encoding='utf-8-sig');s=s.replace("g=group('grpAbout','Thông tin')","g=group('grpAbout','Thông tin')\nbutton(g,'btnGeneralSettings','Cài đặt add-in','PropertySheet','OnGeneralSettings','Bật hoặc tắt cảnh báo vùng chọn lớn; đặt ngưỡng 100–20.000 ký tự. Mặc định 1.000. Cảnh báo xuất hiện trước xử lý văn bản, đổi bảng mã, đọc số, tạo mã và dịch. Có thể chọn Hủy để không thực hiện.')");p.write_text(s,encoding='utf-8-sig')
Path('word/Assets/icons/btnGeneralSettings.png').write_bytes(Path('word/Assets/icons/btnTranslationSettings.png').read_bytes())
for n in ['build_word.py','package_word.py','WordUtilities.cs','Forms.cs','Setup.cs','HUONG-DAN.md','BUILD.md']:
 p=Path('word')/n;s=p.read_text(encoding='utf-8-sig').replace('1.3.4','1.3.5');p.write_text(s,encoding='utf-8-sig')
p=Path('tests/test_word_icons.ps1');s=p.read_text(encoding='utf-8-sig').replace('1.3.4','1.3.5').replace('23','24');p.write_text(s,encoding='utf-8-sig')
p=Path('word/package_word.py');s=p.read_text(encoding='utf-8-sig').replace('23 distinct','24 distinct').replace('all 23 icons','all 24 icons');p.write_text(s,encoding='utf-8-sig')
