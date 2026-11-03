from pathlib import Path
p=Path('word/WordUtilities.cs');s=p.read_text(encoding='utf-8-sig');needle='            var plan=new TextPlan(original);';start=s.index(needle,s.index('public void Transform(object range,string mode)'))
s=s[:start]+'''            if(mode=="accent"){
                var replacements=original.Distinct().Select(c=>new {Source=c.ToString(),Value=TextTools.StripAccents(c.ToString())}).Where(x=>x.Source!=x.Value).GroupBy(x=>x.Value).ToList();
                if(replacements.Count==0)return;
                bool screen=app.ScreenUpdating,changed=false;
                try{app.ScreenUpdating=false;Edit(range,"Tiện ích Word - bỏ dấu",()=>{
                    foreach(var group in replacements){
                        dynamic search=r.Duplicate,find=search.Find;
                        find.ClearFormatting();find.Replacement.ClearFormatting();
                        string letters=string.Concat(group.Select(x=>x.Source));
                        // Accent sources contain no Word wildcard metacharacters.
                        string pattern=letters.Length==1?letters:"["+letters+"]";
                        changed=true;
                        find.Execute(FindText:pattern,MatchCase:true,MatchWholeWord:false,MatchWildcards:letters.Length>1,MatchSoundsLike:false,MatchAllWordForms:false,Forward:true,Wrap:0,Format:false,ReplaceWith:group.Key,Replace:2);
                    }
                });}catch{if(changed)r.Document.Undo(1);throw;}finally{app.ScreenUpdating=screen;}
                return;
            }
''' +s[start:];p.write_text(s,encoding='utf-8-sig')
for n in ['WordUtilities.cs','Forms.cs','Setup.cs','build_word.py','package_word.py','HUONG-DAN.md','BUILD.md']:
 p=Path('word')/n;s=p.read_text(encoding='utf-8-sig').replace('1.3.2','1.3.3');p.write_text(s,encoding='utf-8-sig')
p=Path('tests/test_word_text_positions.ps1');s=p.read_text(encoding='utf-8-sig').replace('1.3.2','1.3.3');p.write_text(s,encoding='utf-8-sig')
