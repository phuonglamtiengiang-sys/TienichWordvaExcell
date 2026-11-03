using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace TienIchWord {
    public class UtilityForm:Form {
        public UtilityForm(string title,int width,int height){Text=title;ClientSize=new Size(width,height);Font=new Font("Segoe UI",10);AutoScaleDimensions=new SizeF(96,96);AutoScaleMode=AutoScaleMode.Dpi;FormBorderStyle=FormBorderStyle.FixedDialog;MaximizeBox=false;MinimizeBox=false;StartPosition=FormStartPosition.CenterParent;}
        protected Label LabelAt(string text,int x,int y,int width,int height){var c=new Label{Text=text,Left=x,Top=y,Width=width,Height=height};Controls.Add(c);return c;}
        protected TextBox TextAt(int x,int y,int width,string text){var c=new TextBox{Left=x,Top=y,Width=width,Text=text};Controls.Add(c);return c;}
        protected Button ButtonAt(string text,int x,int y,int width,EventHandler click){var c=new Button{Text=text,Left=x,Top=y,Width=width,Height=36};c.Click+=click;Controls.Add(c);return c;}
        protected ComboBox ComboAt(int x,int y,int width){var c=new ComboBox{Left=x,Top=y,Width=width,DropDownStyle=ComboBoxStyle.DropDownList};Controls.Add(c);return c;}
    }
    public sealed class GeneralSettingsForm:UtilityForm {
        public GeneralSettingsForm(Preferences prefs):base("Cài đặt add-in Word",560,290){
            var enabled=new CheckBox{Text="Cảnh báo khi chọn nhiều văn bản",Left=20,Top=20,Width=510,Checked=prefs.WarnLargeSelection};Controls.Add(enabled);
            LabelAt("Hiện cảnh báo khi vùng chọn vượt số ký tự:",20,65,510,28);
            var limit=new NumericUpDown{Left=20,Top=98,Width=200,Minimum=100,Maximum=100000,Increment=100,Value=Math.Max(100,Math.Min(100000,prefs.SelectionWarningLimit)),ThousandsSeparator=true};Controls.Add(limit);
            LabelAt("Mặc định: 1.000 ký tự. Cảnh báo cho phép tiếp tục hoặc hủy. Giới hạn xử lý 100.000 ký tự, dịch 20.000 ký tự vẫn áp dụng. Cài đặt được nhớ cho lần sau.",20,140,515,70);
            ButtonAt("Lưu",20,235,250,(s,e)=>{prefs.WarnLargeSelection=enabled.Checked;prefs.SelectionWarningLimit=(int)limit.Value;try{prefs.Save();Close();}catch(Exception ex){MessageBox.Show(this,ex.Message,"Không lưu được cài đặt");}});
            ButtonAt("Đóng",290,235,250,(s,e)=>Close());
        }
    }
    public sealed class Bank {public string Name{get;set;}public string Bin{get;set;}}
    public sealed class PaymentForm:UtilityForm {
        readonly Preferences prefs;readonly WordEngine engine;readonly object target;readonly ComboBox bank;readonly TextBox account,amount,memo;readonly Label status;readonly Button create;
        public PaymentForm(Preferences p,WordEngine word,object range):base("Thanh toán VietQR — Tiện ích Word",660,430){
            prefs=p;engine=word;target=range;
            LabelAt("1. Ngân hàng",16,22,170,28);bank=ComboAt(190,18,450);
            var banks=Resources.Text("banks.tsv").Split(new[]{'\r','\n'},StringSplitOptions.RemoveEmptyEntries).Select(line=>{var parts=line.Split('\t');return new Bank{Name=parts[0],Bin=parts[1]};}).ToList();
            bank.DisplayMember="Name";bank.ValueMember="Bin";bank.DataSource=banks;int index=banks.FindIndex(b=>b.Bin==p.Bank);bank.SelectedIndex=index>=0?index:0;
            LabelAt("2. Số tài khoản",16,68,170,28);account=TextAt(190,64,450,p.Account);account.MaxLength=19;
            LabelAt("3. Số tiền VND",16,114,170,28);amount=TextAt(190,110,450,"");amount.MaxLength=13;
            LabelAt("4. Nội dung",16,160,170,28);memo=TextAt(190,156,450,"");memo.MaxLength=50;
            LabelAt("Để trống số tiền: QR không ấn định số tiền. Số tiền chỉ dùng chữ số, không dấu phân cách. Nội dung tối đa 50 ký tự; tự bỏ dấu tiếng Việt.",16,202,624,58);
            LabelAt("Khi bấm Tạo QR, ngân hàng, tài khoản, số tiền và nội dung được gửi tới img.vietqr.io để tạo ảnh. Ảnh được nhúng tại cuối vùng chọn/con trỏ, không thay chữ gốc.",16,267,624,60);
            status=LabelAt("Có thể lưu ngân hàng và tài khoản để dùng lại trên máy này.",16,332,624,42);
            ButtonAt("Lưu tài khoản",16,380,170,(s,e)=>{try{SaveAccount();status.Text="Đã lưu ngân hàng và tài khoản cho lần sau.";}catch(Exception ex){status.Text=ex.Message;}});
            create=ButtonAt("Tạo QR và chèn vào Word",200,380,270,(s,e)=>CreateCode());
            var close=ButtonAt("Đóng",484,380,156,(s,e)=>Close());CancelButton=close;
        }
        void SaveAccount(){string bin=(string)bank.SelectedValue;Codes.VietQR(bin,account.Text.Trim(),"","");prefs.Bank=bin;prefs.Account=account.Text.Trim();prefs.Save();}
        void CreateCode(){try{WordEngine.Validate(target,false);string url=Codes.VietQR((string)bank.SelectedValue,account.Text.Trim(),amount.Text.Trim(),memo.Text.Trim());create.Enabled=false;Cursor=Cursors.WaitCursor;status.Text="Đang tạo mã QR...";Refresh();byte[] bytes=Codes.Download(url);engine.InsertPicture(target,bytes,"vietqr");SaveAccount();DialogResult=DialogResult.OK;Close();}catch(Exception ex){status.Text=ex.Message;}finally{create.Enabled=true;Cursor=Cursors.Default;}}
    }
    public sealed class CodeForm:UtilityForm {
        readonly string kind;readonly WordEngine engine;readonly object target;readonly TextBox content;readonly Label status;readonly Button create;
        public CodeForm(string type,WordEngine word,object range):base("Tạo "+type.ToUpperInvariant()+" — Tiện ích Word",650,380){
            kind=type;engine=word;target=range;
            LabelAt("Nội dung cần tạo mã",16,18,610,25);content=TextAt(16,48,618,(string)((dynamic)range).Text??"");content.Multiline=true;content.ScrollBars=ScrollBars.Vertical;content.Height=130;
            LabelAt(type=="qr"?"QR hỗ trợ tiếng Việt, tối đa 2.000 byte UTF-8. Nội dung gửi tới api.qrserver.com khi bấm Tạo mã.":"Code 128 dùng ASCII; Code 39 dùng chữ hoa và số; EAN-13 cần 12 số hoặc 13 số đúng kiểm tra. Nội dung gửi tới bwipjs-api.metafloor.com khi bấm Tạo mã.",16,190,618,68);
            status=LabelAt("Ảnh độ phân giải cao được nhúng vào tài liệu. Mã vạch quá rộng sẽ được báo để bạn rút ngắn nội dung hoặc dùng QR.",16,267,618,55);
            create=ButtonAt("Tạo mã và chèn vào Word",16,330,390,(s,e)=>CreateCode());var close=ButtonAt("Đóng",420,330,214,(s,e)=>Close());CancelButton=close;
        }
        void CreateCode(){try{WordEngine.Validate(target,false);string url=Codes.Url(kind,content.Text);create.Enabled=false;Cursor=Cursors.WaitCursor;status.Text="Đang tải ảnh mã...";Refresh();engine.InsertPicture(target,Codes.Download(url),kind);DialogResult=DialogResult.OK;Close();}catch(Exception ex){status.Text=ex.Message;}finally{create.Enabled=true;Cursor=Cursors.Default;}}
    }
    public sealed class Language {public string Name{get;set;}public string Code{get;set;}}
    public sealed class TranslationSettingsForm:UtilityForm {
        readonly Preferences prefs;readonly ComboBox source,target,provider;readonly TextBox email;readonly Label status;
        static readonly string[] Names={"Tiếng Việt","Tiếng Anh","Tiếng Nhật","Tiếng Hàn","Trung giản thể","Trung phồn thể","Tiếng Pháp","Tiếng Đức","Tiếng Thái","Tiếng Tây Ban Nha","Tiếng Nga","Tiếng Indonesia"};
        ComboBox LanguageAt(int x,string code){var box=ComboAt(x,53,290);box.DisplayMember="Name";box.ValueMember="Code";box.DataSource=Translator.Languages.Select((v,i)=>new Language{Name=Names[i],Code=v}).ToList();box.SelectedValue=Translator.Languages.Contains(code)?code:"vi";return box;}
        public TranslationSettingsForm(Preferences p):base("Cài đặt dịch — Tiện ích Word",640,435){
            prefs=p;LabelAt("Văn bản đang là tiếng",16,20,290,25);LabelAt("Muốn dịch sang",330,20,290,25);source=LanguageAt(16,p.Source);target=LanguageAt(330,p.Target);
            LabelAt("Dịch vụ",16,101,604,25);provider=ComboAt(16,132,604);provider.Items.AddRange(new object[]{"Tự động: MyMemory, hết hạn mức thì dùng Google web","Chỉ MyMemory","Chỉ Google web"});provider.SelectedIndex=p.Provider=="googleweb"?2:p.Provider=="mymemory"?1:0;
            LabelAt("Email MyMemory (tùy chọn)",16,181,604,25);email=TextAt(16,212,604,p.Email);email.MaxLength=254;
            LabelAt("MyMemory: hạn mức miễn phí theo dịch vụ; có thể nhập email hợp lệ để mở rộng hạn mức. Google web không cần email, là kết nối web không chính thức và có thể bị giới hạn.",16,254,604,65);
            status=LabelAt("Chọn đoạn văn rồi bấm Dịch vùng chọn trên Ribbon. Tối đa 20.000 ký tự/lần; bản dịch chèn dưới đoạn gốc. Nội dung được gửi dịch vụ đã chọn sau khi xác nhận.",16,325,604,55);
            ButtonAt("Lưu mặc định",16,388,400,(s,e)=>Save());var close=ButtonAt("Đóng",430,388,190,(s,e)=>Close());CancelButton=close;
        }
        void Save(){try{string address=email.Text.Trim();if(!Translator.ValidEmail(address))throw new ArgumentException("Email chưa đúng định dạng. Có thể để trống.");prefs.Source=(string)source.SelectedValue;prefs.Target=(string)target.SelectedValue;prefs.Provider=provider.SelectedIndex==2?"googleweb":provider.SelectedIndex==1?"mymemory":"auto";prefs.Email=address;prefs.Save();status.Text="Đã lưu. Đóng cửa sổ, chọn đoạn văn rồi bấm Dịch vùng chọn.";}catch(Exception ex){status.Text=ex.Message;}}
    }
    public sealed class ProgressForm:UtilityForm {
        readonly Label status;bool finish;public bool Cancelled{get;private set;}
        public ProgressForm():base("Đang dịch — Tiện ích Word",450,145){status=LabelAt("Đang chuẩn bị dịch...",16,18,418,56);ButtonAt("Dừng",16,91,418,(s,e)=>{Cancelled=true;status.Text="Đang dừng, chờ yêu cầu hiện tại kết thúc...";});FormClosing+=(s,e)=>{if(!finish){Cancelled=true;e.Cancel=true;status.Text="Đang dừng...";}};}
        public void UpdateProgress(int done,int total){status.Text="Đang dịch phần "+done+" / "+total+". Chưa sửa tài liệu.";Refresh();}
        public void Finish(){finish=true;Close();}
    }
    public sealed class AboutForm:UtilityForm {
        public AboutForm():base("Giới thiệu — Tiện ích Word",720,570){
            var brand=new PictureBox{Left=18,Top=7,Width=684,Height=90,SizeMode=PictureBoxSizeMode.Zoom,Image=Resources.Picture("word-header.png")};Controls.Add(brand);
            LabelAt("Phiên bản 2.30.4 — Microsoft Word Windows",20,105,310,28);
            LabelAt("Tác giả: Hồ Phương Lâm\nEmail: hophuonglam@outlook.com",20,139,310,50);
            LabelAt("Chuyển mã: UniKey UVConverter\nPhạm Kim Long — GPL-2.0-or-later",20,194,310,30).Font=new Font("Segoe UI",9);
            LabelAt("Nguồn tham khảo: Microsoft Office/VSTO, UniKey UVConverter, VietQR.io, BWIP-JS, QRServer, MyMemory và Google Translate web. Chi tiết nguồn, giấy phép và cách sử dụng nằm trong THAM-KHAO.md và THIRD-PARTY-NOTICE.md của gói mã nguồn.",20,410,290,100).Font=new Font("Segoe UI",8);
            LabelAt("Nếu bạn dùng Add-in này, xin gửi tặng tôi 1 ly cà phê để tôi có động lực phát triển tiếp.",20,234,290,125).Font=new Font("Segoe UI",14);
            LabelAt("Cảm ơn bạn đã sử dụng và ủng hộ!",20,365,290,36);
            var picture=new PictureBox{Left=340,Top=105,Width=360,Height=395,SizeMode=PictureBoxSizeMode.Zoom,Image=Resources.Picture("support-momo.png")};Controls.Add(picture);
            picture.DoubleClick+=(s,e)=>{using(var large=new UtilityForm("Mã ủng hộ tác giả",650,710)){var view=new PictureBox{Dock=DockStyle.Fill,SizeMode=PictureBoxSizeMode.Zoom,Image=picture.Image};large.Controls.Add(view);large.ShowDialog(this);}};
            LabelAt("Nhấp đúp mã QR để phóng to.",340,500,360,28);
            var close=ButtonAt("Đóng",20,525,290,(s,e)=>Close());CancelButton=close;
            FormClosed+=(s,e)=>{picture.Image.Dispose();brand.Image.Dispose();};
        }
    }
}
