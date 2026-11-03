using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace TienIchWord {
    public sealed class TextEdit {public int Start,Length;public string Value;}
    public sealed class TextPlan {
        public string Original,Result;public readonly List<List<TextEdit>> Stages=new List<List<TextEdit>>();
        public TextPlan(string text){Original=Result=text;}
        public void Replace(string pattern,string replacement){
            var edits=Regex.Matches(Result,pattern).Cast<Match>().Select(m=>new TextEdit{Start=m.Index,Length=m.Length,Value=m.Result(replacement)}).Where(e=>Result.Substring(e.Start,e.Length)!=e.Value).ToList();
            Add(edits);
        }
        public void Add(List<TextEdit> edits){
            if(edits.Count==0)return;var b=new StringBuilder(Result.Length);int position=0;
            foreach(var e in edits.OrderBy(e=>e.Start)){
                if(e.Start<position||e.Length<0||e.Start+e.Length>Result.Length)throw new ArgumentException("Các vùng sửa không hợp lệ.");
                b.Append(Result,position,e.Start-position);b.Append(e.Value);position=e.Start+e.Length;
            }
            b.Append(Result,position,Result.Length-position);Result=b.ToString();Stages.Add(edits);
        }
    }
    public static class VietnameseEncoding {
        public static TextPlan Convert(string text,string source,bool uppercase){return UniKeyDecoder.Convert(text,source,uppercase);}
    }
    public static class TextNormalization {
        public static TextPlan Create(string text,bool spaces,bool punctuation,bool blankLines,bool pdf){
            var p=new TextPlan(text);
            if(pdf){p.Replace(@"\v", " ");p.Replace(@"(?<=[\p{L}\p{N},])\r(?=\p{Ll})"," ");}
            if(spaces){p.Replace("[\u00a0\u202f]"," ");p.Replace("[\u200b\ufeff]","");p.Replace(@" {2,}"," ");p.Replace(@"(?m)(^|\r|\v) +","$1");p.Replace(@" +(?=\r|\v|$)","");}
            if(punctuation){p.Replace(@" +([,;:!?])","$1");p.Replace(@"([,;:!?])(?=\p{L})","$1 ");p.Replace(@"\( +","(");p.Replace(@" +\)",")");}
            if(blankLines)p.Replace(@"\r(?:[ \t]*\r)+","\r");
            return p;
        }
    }
    public sealed class DocumentFormatOptions {
        public string Font="Times New Roman";public float Size=13,Line=1.5f,IndentCm=1,Left=3,Right=2,Top=2,Bottom=2;public bool WholeDocument,Margins;
        public void Validate(){if(string.IsNullOrWhiteSpace(Font)||Size<8||Size>72||Line<1||Line>3||IndentCm<0||IndentCm>5)throw new ArgumentException("Kiểm tra phông, cỡ chữ, giãn dòng và thụt đầu dòng.");foreach(float v in new[]{Left,Right,Top,Bottom})if(v<0.5||v>8)throw new ArgumentException("Lề trang cần từ 0,5 đến 8 cm.");}
    }
    public sealed partial class WordEngine {
        public void FormatDocument(object selection,DocumentFormatOptions options){
            options.Validate();dynamic r=selection,doc=r.Document;
            if((int)r.StoryType!=1||(bool)doc.ReadOnly||(int)doc.ProtectionType!=-1)throw new InvalidOperationException("Chọn nội dung chính trong tài liệu có quyền chỉnh sửa.");
            if(options.WholeDocument)r=doc.Content.Duplicate;
            else if(r.Start==r.End)throw new InvalidOperationException("Hãy chọn các đoạn cần định dạng hoặc chọn Toàn bộ tài liệu.");
            if(options.Margins)foreach(dynamic section in doc.Sections){dynamic ps=section.PageSetup;if((options.Left+options.Right)*28.3464567>(double)ps.PageWidth-72||(options.Top+options.Bottom)*28.3464567>(double)ps.PageHeight-72)throw new InvalidOperationException("Lề quá lớn so với khổ giấy hiện tại.");}
            bool started=false,changed=false;
            try{app.UndoRecord.StartCustomRecord("Tiện ích Word - định dạng tài liệu");started=true;
                r.Font.Name=options.Font;changed=true;r.Font.Size=options.Size;r.ParagraphFormat.LineSpacingRule=5;r.ParagraphFormat.LineSpacing=12*options.Line;r.ParagraphFormat.FirstLineIndent=options.IndentCm*28.3464567f;
                if(options.Margins)foreach(dynamic section in doc.Sections){dynamic ps=section.PageSetup;ps.LeftMargin=options.Left*28.3464567f;ps.RightMargin=options.Right*28.3464567f;ps.TopMargin=options.Top*28.3464567f;ps.BottomMargin=options.Bottom*28.3464567f;}
            }catch{if(started){app.UndoRecord.EndCustomRecord();started=false;}if(changed)doc.Undo(1);throw;}finally{if(started)app.UndoRecord.EndCustomRecord();}
        }
    }
}
