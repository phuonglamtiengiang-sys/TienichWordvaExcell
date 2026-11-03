using System;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;
namespace TienIchWord {
    public sealed class NormalizeForm:UtilityForm {
        readonly WordEngine engine;readonly object target;readonly string original;readonly bool encoding;
        readonly CheckBox spaces,punctuation,blank,pdf,upper;readonly ComboBox source;readonly TextBox preview;readonly Label status;TextPlan plan;
        CheckBox Check(string text,int top,bool value){var c=new CheckBox{Text=text,Left=18,Top=top,Width=700,Height=30,Checked=value};Controls.Add(c);c.CheckedChanged+=(s,e)=>RefreshPreview();return c;}
        public NormalizeForm(WordEngine e,object range,bool convert):base(convert?"Đổi bảng mã tiếng Việt sang Unicode":"Chuẩn hóa văn bản / văn bản từ PDF",760,650){
            engine=e;target=range;original=(string)((dynamic)range).Text;encoding=convert;
            if(convert){
                LabelAt("Chọn đúng bảng mã nguồn. Không tự đoán để tránh đổi nhầm chữ Unicode.",18,15,720,50);
                source=ComboAt(18,68,715);source.Items.AddRange(new object[]{"TCVN3 (ABC) — thường dùng phông .VnTime, .VnArial","VNI Windows — thường dùng phông VNI-Times"});source.SelectedIndex=0;source.SelectedIndexChanged+=(s,x)=>RefreshPreview();
                upper=Check("Nguồn dùng phông TCVN3 chữ HOA (ví dụ .VnTimeH)",110,false);
                LabelAt("Bộ chuyển mã: UniKey UVConverter (bản chuyển C#). Đích: Unicode, phông Times New Roman. Giữ cỡ chữ, chữ đậm/nghiêng tại phần không đổi; ký tự VNI ghép lấy định dạng tại vị trí thay thế. Chỉ xử lý vùng chọn.",18,150,715,90);
            }else{
                spaces=Check("Sửa khoảng trắng thừa và ký tự khoảng trắng ẩn",18,true);
                punctuation=Check("Sửa khoảng trắng quanh dấu phẩy, chấm phẩy, hai chấm, !, ? và ngoặc",56,true);
                blank=Check("Bỏ dòng/đoạn trống trong vùng chọn",94,false);
                pdf=Check("Nối dòng từ PDF: ngắt dòng mềm và đoạn tiếp nối bắt đầu bằng chữ thường",132,false);
                LabelAt("Nối dòng PDF có thể ghép nhầm đoạn. Luôn xem bản xem trước; không tự nối dòng kết thúc bằng dấu chấm hoặc dòng bắt đầu bằng chữ hoa/số. Không sửa dấu chấm để tránh ảnh hưởng số và địa chỉ web.",18,176,715,70);
            }
            LabelAt("Xem trước kết quả (chưa sửa tài liệu)",18,255,715,28);
            preview=new TextBox{Left=18,Top=290,Width=715,Height=245,Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Both,WordWrap=false};Controls.Add(preview);
            status=LabelAt("",18,545,715,45);
            ButtonAt("Áp dụng vào vùng chọn",18,601,470,(s,x)=>Apply());var cancel=ButtonAt("Hủy",504,601,229,(s,x)=>Close());CancelButton=cancel;RefreshPreview();
        }
        void RefreshPreview(){if(preview==null)return;try{
            if(encoding){bool tcvn=source.SelectedIndex==0;upper.Enabled=tcvn;plan=VietnameseEncoding.Convert(original,tcvn?"tcvn":"vni",tcvn&&upper.Checked);}
            else plan=TextNormalization.Create(original,spaces.Checked,punctuation.Checked,blank.Checked,pdf.Checked);
            preview.Text=plan.Result.Replace("\r","\r\n").Replace("\v","\r\n");status.Text=plan.Result==original?(encoding?"Nội dung chưa thay đổi. Chuyển mã vẫn đổi sang phông Unicode.":"Các tùy chọn hiện tại không làm thay đổi nội dung."):"Sẵn sàng áp dụng. Ctrl+Z để hoàn tác toàn bộ lượt này.";
        }catch(Exception ex){status.Text=ex.Message;plan=null;}}
        void Apply(){try{if(plan==null)return;engine.ApplyPlan(target,plan,encoding);DialogResult=DialogResult.OK;Close();}catch(Exception ex){status.Text=ex.Message;}}
    }
    public sealed class DocumentFormatForm:UtilityForm {
        readonly WordEngine engine;readonly object target;readonly ComboBox font;readonly NumericUpDown size,line,indent,left,right,top,bottom;readonly RadioButton whole,selected;readonly CheckBox margins;readonly Label status;
        NumericUpDown Number(string label,int x,int y,decimal value,decimal min,decimal max){LabelAt(label,x,y,310,26);var n=new NumericUpDown{Left=x,Top=y+29,Width=300,Minimum=min,Maximum=max,DecimalPlaces=2,Increment=0.25m,Value=value};Controls.Add(n);return n;}
        public DocumentFormatForm(WordEngine e,object range):base("Chuẩn hóa định dạng tài liệu",700,610){
            engine=e;target=range;
            bool hasSelection=(int)((dynamic)range).End>(int)((dynamic)range).Start;
            selected=new RadioButton{Text="Các đoạn thuộc vùng chọn",Left=18,Top=15,Width=330,Checked=hasSelection};whole=new RadioButton{Text="Toàn bộ nội dung chính",Left=350,Top=15,Width=330,Checked=!hasSelection};Controls.Add(selected);Controls.Add(whole);
            LabelAt("Phông chữ",18,55,640,25);font=ComboAt(18,83,640);using(var installed=new InstalledFontCollection())font.Items.AddRange(installed.Families.Select(f=>f.Name).OrderBy(v=>v).Cast<object>().ToArray());font.SelectedItem="Times New Roman";if(font.SelectedIndex<0)font.SelectedIndex=0;
            size=Number("Cỡ chữ (pt)",18,126,13,8,72);line=Number("Giãn dòng (bội số)",350,126,1.5m,1,3);indent=Number("Thụt dòng đầu (cm)",18,199,1,0,5);
            margins=new CheckBox{Text="Đổi lề TẤT CẢ các section trong tài liệu (đơn vị cm)",Left=18,Top=279,Width=655,Height=30};Controls.Add(margins);
            left=Number("Lề trái",18,318,3,0.5m,8);right=Number("Lề phải",350,318,2,0.5m,8);top=Number("Lề trên",18,388,2,0.5m,8);bottom=Number("Lề dưới",350,388,2,0.5m,8);
            Action enabled=()=>{foreach(var n in new[]{left,right,top,bottom})n.Enabled=margins.Checked;};margins.CheckedChanged+=(s,x)=>enabled();enabled();
            status=LabelAt("Chỉ đổi các thuộc tính trên; giữ chữ đậm/nghiêng, căn đoạn và khổ giấy. Giãn dòng/thụt đầu áp dụng cho cả đoạn giao với vùng chọn. Ctrl+Z để hoàn tác.",18,470,650,85);
            ButtonAt("Áp dụng định dạng",18,560,435,(s,x)=>Apply());var close=ButtonAt("Hủy",470,560,198,(s,x)=>Close());CancelButton=close;
        }
        void Apply(){try{var p=new DocumentFormatOptions{Font=(string)font.SelectedItem,Size=(float)size.Value,Line=(float)line.Value,IndentCm=(float)indent.Value,Left=(float)left.Value,Right=(float)right.Value,Top=(float)top.Value,Bottom=(float)bottom.Value,WholeDocument=whole.Checked,Margins=margins.Checked};engine.FormatDocument(target,p);DialogResult=DialogResult.OK;Close();}catch(Exception ex){status.Text=ex.Message;}}
    }
}
