from pathlib import Path
p=Path('word/WordUtilities.cs');s=p.read_text(encoding='utf-8-sig')
a=s.index('            string[] tokens=mode=="clean"?TextTools.CleanTokens(original):null;',s.index('public void Transform(object range,string mode)'))
b=s.index('\n        public string Number(',a)
s=s[:a]+'''            var plan=new TextPlan(original);
            var edits=new List<TextEdit>();
            string[] tokens=mode=="clean"?TextTools.CleanTokens(original):null;
            if(mode!="clean"&&mode!="accent")throw new ArgumentException("Kiểu xử lý văn bản không hợp lệ.");
            for(int i=0;i<original.Length;i++){
                string value=mode=="clean"?tokens[i]:TextTools.StripAccents(original[i].ToString());
                if(value!=original[i].ToString())edits.Add(new TextEdit{Start=i,Length=1,Value=value});
            }
            plan.Add(edits);ApplyPlan(range,plan,false);
        }
''' + s[b:];p.write_text(s,encoding='utf-8-sig')
p=Path('word/TextUtilities.cs');s=p.read_text(encoding='utf-8-sig')
s=s.replace('bool changed=false;int start=r.Start;','bool changed=false;')
s=s.replace('foreach(var stage in plan.Stages)foreach(var e in stage.OrderByDescending(e=>e.Start)){dynamic part=r.Document.Range(start+e.Start,start+e.Start+e.Length);part.Text=e.Value;changed=true;}\n                r.SetRange(start,start+plan.Result.Length);','''foreach(var stage in plan.Stages){
                    var cells=new List<TextCell>();var text=new StringBuilder();
                    foreach(dynamic ch in r.Characters){
                        string value=(string)ch.Text??"";
                        if(value.Length==0)continue;
                        cells.Add(new TextCell{Start=text.Length,Length=value.Length,Range=(object)ch.Duplicate});
                        text.Append(value);
                    }
                    if(text.ToString()!=(string)r.Text)throw new InvalidOperationException("Không thể xác định chính xác ký tự trong vùng chọn. Chưa áp dụng thay đổi.");
                    foreach(var e in stage.OrderByDescending(e=>e.Start)){
                        var parts=cells.Where(c=>c.Start<e.Start+e.Length&&c.Start+c.Length>e.Start).ToList();
                        if(parts.Count==0||parts[0].Start!=e.Start||parts.Last().Start+parts.Last().Length!=e.Start+e.Length)
                            throw new InvalidOperationException("Vùng sửa cắt giữa một ký tự Unicode. Hãy chọn lại đoạn chữ.");
                        for(int i=parts.Count-1;i>=0;i--){dynamic part=parts[i].Range;changed=true;part.Text=i==0?e.Value:"";}
                    }
                }''')
s=s.replace('public sealed partial class WordEngine {','public sealed partial class WordEngine {\n        sealed class TextCell {public int Start,Length;public object Range;}')
p.write_text(s,encoding='utf-8-sig')
for name in ['build_word.py','package_word.py','WordUtilities.cs','Forms.cs','Setup.cs','HUONG-DAN.md','BUILD.md']:
 p=Path('word')/name;s=p.read_text(encoding='utf-8-sig').replace('1.3.1','1.3.2');p.write_text(s,encoding='utf-8-sig')
