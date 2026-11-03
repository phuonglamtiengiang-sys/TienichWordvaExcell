from pathlib import Path
p=Path('word/WordUtilities.cs');s=p.read_text(encoding='utf-8-sig');s=s.replace('''            for(int i=0;i<original.Length;i++){
                string value=mode=="clean"?tokens[i]:TextTools.StripAccents(original[i].ToString());
                if(value!=original[i].ToString())edits.Add(new TextEdit{Start=i,Length=1,Value=value});
            }''','''            if(mode=="accent"){
                var elements=StringInfo.GetTextElementEnumerator(original);
                while(elements.MoveNext()){
                    string source=elements.GetTextElement(),value=TextTools.StripAccents(source);
                    if(value!=source)edits.Add(new TextEdit{Start=elements.ElementIndex,Length=source.Length,Value=value});
                }
            }else for(int i=0;i<original.Length;i++){
                if(tokens[i]!=original[i].ToString())edits.Add(new TextEdit{Start=i,Length=1,Value=tokens[i]});
            }''');p.write_text(s,encoding='utf-8-sig')
