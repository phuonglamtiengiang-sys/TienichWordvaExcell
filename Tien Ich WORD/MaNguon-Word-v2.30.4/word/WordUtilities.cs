using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

[assembly: ComVisible(false)]
[assembly: AssemblyVersion("2.30.4.0")]
[assembly: AssemblyFileVersion("2.30.4.0")]
namespace TienIchWord {
    [ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2 {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode, [In, MarshalAs(UnmanagedType.IDispatch)] object addin, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(int removeMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    }
    [ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility { [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string ribbonId); }

    public sealed class Preferences {
        public bool WarnLargeSelection=true;
        public int SelectionWarningLimit=1000;
        public bool ShouldWarn(int count){return WarnLargeSelection&&count>Math.Max(100,Math.Min(100000,SelectionWarningLimit));}
        public string Unit = "dong", Bank = "970422", Account = "", Source = "vi", Target = "en", Provider = "auto", Email = "";
        public static string Folder { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TienIchWord"); } }
        public static Preferences Load() {
            try { return new JavaScriptSerializer().Deserialize<Preferences>(File.ReadAllText(Path.Combine(Folder,"settings.json"), Encoding.UTF8)) ?? new Preferences(); }
            catch { return new Preferences(); }
        }
        public void Save() {
            Directory.CreateDirectory(Folder);
            string path = Path.Combine(Folder,"settings.json"), temp = path + ".tmp";
            File.WriteAllText(temp, new JavaScriptSerializer().Serialize(this), new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temp, path, null); else File.Move(temp, path);
        }
    }

    public static class TextTools {
        public static string StripAccents(string s) {
            var b = new StringBuilder();
            foreach (char c in s.Normalize(NormalizationForm.FormD))
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) b.Append(c=='đ'?'d':c=='Đ'?'D':c);
            return b.ToString().Normalize(NormalizationForm.FormC);
        }
        public static string Transform(string s, string mode) {
            if(mode=="upper") return s.ToUpperInvariant();
            if(mode=="lower") return s.ToLowerInvariant();
            if(mode=="title") return Regex.Replace(s.ToLowerInvariant(), @"(^|[\s\-])([\p{L}])", m=>m.Groups[1].Value+m.Groups[2].Value.ToUpperInvariant());
            if(mode=="sentence") {
                var b=new StringBuilder(s.ToLowerInvariant()); bool first=true;
                for(int i=0;i<b.Length;i++) {
                    if(first && char.IsLetter(b[i])) { b[i]=char.ToUpperInvariant(b[i]); first=false; }
                    if(b[i]=='.'||b[i]=='!'||b[i]=='?'||b[i]=='\r'||b[i]=='\n'||b[i]=='\v') first=true;
                }
                return b.ToString();
            }
            throw new ArgumentException("Kiểu xử lý văn bản không hợp lệ.");
        }
        public static string[] CleanTokens(string s) {
            var tokens=new string[s.Length]; int last=-1; bool boundary=true;
            for(int i=0;i<s.Length;i++) {
                char c=s[i];
                if(c=='\u200b'||c=='\ufeff') {tokens[i]="";continue;}
                if(c=='\u00a0'||c=='\u202f') c=' ';
                if(c==' ') {
                    tokens[i]=(boundary || (last>=0 && tokens[last]==" "))?"":" ";
                } else {
                    if(c=='\r'||c=='\n'||c=='\v'||c=='\t') {
                        if(last>=0 && tokens[last]==" ") tokens[last]="";
                        boundary=true;
                    } else boundary=false;
                    tokens[i]=c.ToString();
                }
                if(tokens[i].Length>0) last=i;
            }
            if(last>=0 && tokens[last]==" ") tokens[last]="";
            return tokens;
        }
    }

    public static class NumberWords {
        static readonly string[] V={"không","một","hai","ba","bốn","năm","sáu","bảy","tám","chín"};
        static readonly string[] E={"","one","two","three","four","five","six","seven","eight","nine","ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"};
        static readonly string[] Tens={"","","twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"};
        static string ThreeVN(int n,bool first) {
            int h=n/100,t=n%100/10,u=n%10; var b=new StringBuilder();
            if(h>0||!first)b.Append(V[h]+" trăm ");
            if(t>1){b.Append(V[t]+" mươi ");if(u>0)b.Append((u==1?"mốt":u==5?"lăm":V[u])+" ");}
            else if(t==1){b.Append("mười ");if(u>0)b.Append((u==5?"lăm":V[u])+" ");}
            else if(u>0){if(h>0||!first)b.Append("lẻ ");b.Append(V[u]+" ");}
            return b.ToString();
        }
        static string ThreeEN(int n) {
            string s=n>=100?E[n/100]+" hundred ":"";int k=n%100;
            if(k>=20)s+=Tens[k/10]+(k%10>0?" "+E[k%10]:"");else if(k>0)s+=E[k];
            return s.Trim();
        }
        public static long Parse(string input) {
            string s=input.Trim().Replace('\u00a0',' ').Replace('\u202f',' ');
            if(!Regex.IsMatch(s,@"^[+-]?(?:[0-9]+|[0-9]{1,3}(?:[., ][0-9]{3})+)$"))
                throw new ArgumentException("Chỉ đọc số nguyên. Có thể dùng dấu chấm, phẩy hoặc khoảng trắng để ngăn nhóm ba chữ số; không nhận số thập phân.");
            s=s.Replace(".","").Replace(",","").Replace(" ","");long n;
            if(!long.TryParse(s,NumberStyles.AllowLeadingSign,CultureInfo.InvariantCulture,out n)||n<=-1000000000000000L||n>=1000000000000000L)
                throw new ArgumentException("Số nguyên phải nằm trong khoảng ±999.999.999.999.999.");
            return n;
        }
        public static string Read(long n,bool english,string unit) {
            if(n<=-1000000000000000L||n>=1000000000000000L)throw new ArgumentOutOfRangeException("n");
            bool neg=n<0;long value=Math.Abs(n);string s="";
            if(value==0)s=english?"zero":"không";
            else {
                int[] g=new int[5];for(int i=0;i<5;i++){g[i]=(int)(value%1000);value/=1000;}
                string[] units=english?new[]{"","thousand","million","billion","trillion"}:new[]{"","nghìn","triệu","tỷ","nghìn tỷ"};
                bool first=true;for(int i=4;i>=0;i--)if(g[i]>0){s+=(english?ThreeEN(g[i]):ThreeVN(g[i],first))+" "+units[i]+" ";first=false;}
            }
            s=Regex.Replace(s,@"\s+"," ").Trim();if(neg)s=(english?"negative ":"âm ")+s;
            s=char.ToUpperInvariant(s[0])+s.Substring(1);
            if(unit=="dong")s+=english?" dong.":" đồng.";else if(unit=="usd")s+=english?" US dollars.":" đô la Mỹ.";
            return s;
        }
    }

    public static class Codes {
        public static string Encode(string s){return Uri.EscapeDataString(s);}
        public static bool ValidEAN(string s){
            if(!Regex.IsMatch(s,@"^[0-9]{12,13}$"))return false;
            int total=0;for(int i=0;i<12;i++)total+=(s[i]-'0')*(i%2==0?1:3);
            return s.Length==12 || s[12]-'0'==(10-total%10)%10;
        }
        public static string Url(string kind,string text){
            if(string.IsNullOrWhiteSpace(text))throw new ArgumentException("Nhập nội dung cần tạo mã.");
            if(kind=="qr"){
                if(Encoding.UTF8.GetByteCount(text)>2000)throw new ArgumentException("QR tối đa 2.000 byte UTF-8; hãy rút ngắn nội dung.");
                return "https://api.qrserver.com/v1/create-qr-code/?size=900x900&margin=24&format=png&charset-source=UTF-8&data="+Encode(text);
            }
            if(kind=="ean13"&&!ValidEAN(text))throw new ArgumentException("EAN-13 cần 12 chữ số hoặc 13 chữ số có số kiểm tra đúng.");
            if(kind=="code39"){
                text=text.ToUpperInvariant();if(!Regex.IsMatch(text,@"^[0-9A-Z\-. $/+%]+$"))throw new ArgumentException("Code 39 chỉ nhận A-Z, số và ký hiệu - . khoảng trắng $ / + %.");
            } else if(kind=="code128"){
                if(text.Any(c=>c<32||c>126))throw new ArgumentException("Code 128 chỉ nhận ASCII 32–126. Dùng QR cho tiếng Việt.");
            } else if(kind!="ean13")throw new ArgumentException("Loại mã không hợp lệ.");
            if(text.Length>80)throw new ArgumentException("Mã vạch tối đa 80 ký tự; mã dài có thể không vừa trang. Dùng QR nếu cần.");
            return "https://bwipjs-api.metafloor.com/?bcid="+kind+"&text="+Encode(text)+"&scale=6&height=18&includetext&textsize=12&monochrome&backgroundcolor=FFFFFF&paddingwidth=12&paddingheight=6";
        }
        public static string VietQR(string bank,string account,string amount,string memo){
            if(!Regex.IsMatch(bank??"",@"^[0-9]{6}$"))throw new ArgumentException("Chọn ngân hàng hợp lệ.");
            if(!Regex.IsMatch(account??"",@"^[A-Za-z0-9]{1,19}$"))throw new ArgumentException("Số tài khoản cần 1–19 ký tự chữ hoặc số.");
            if(amount!=""&&(!Regex.IsMatch(amount,@"^[0-9]{1,13}$")||decimal.Parse(amount,CultureInfo.InvariantCulture)<=0))throw new ArgumentException("Số tiền là số nguyên VND dương, tối đa 13 chữ số, không dấu phân cách; có thể để trống.");
            if(memo.Length>50)throw new ArgumentException("Nội dung tối đa 50 ký tự.");
            memo=TextTools.StripAccents(memo);if(!Regex.IsMatch(memo,@"^[A-Za-z0-9 ]*$"))throw new ArgumentException("Nội dung chỉ dùng chữ, số và khoảng trắng; dấu tiếng Việt được bỏ.");
            string url="https://img.vietqr.io/image/"+bank+"-"+account+"-"+(amount==""?"compact":"compact2")+".png";
            var q=new List<string>();if(amount!="")q.Add("amount="+amount);if(memo!="")q.Add("addInfo="+Encode(memo));
            return url+(q.Count>0?"?"+string.Join("&",q):"");
        }
        public static byte[] Download(string url){
            byte[] bytes=Network.Get(url);OfficeImage.Read(bytes,url.StartsWith("https://img.vietqr.io/image/",StringComparison.Ordinal));
            return bytes;
        }
    }

    public static class Network {
        public static byte[] Get(string url){
            ServicePointManager.SecurityProtocol|=SecurityProtocolType.Tls12;
            var req=(HttpWebRequest)WebRequest.Create(url);req.Timeout=10000;req.ReadWriteTimeout=10000;req.AutomaticDecompression=DecompressionMethods.GZip|DecompressionMethods.Deflate;
            using(var response=(HttpWebResponse)req.GetResponse())using(var stream=response.GetResponseStream())using(var memory=new MemoryStream()){
                byte[] b=new byte[8192];int n;while((n=stream.Read(b,0,b.Length))>0){memory.Write(b,0,n);if(memory.Length>10000000)throw new IOException("Phản hồi máy chủ quá lớn.");}return memory.ToArray();
            }
        }
    }

    public sealed class TranslationException:Exception { public bool Quota;public TranslationException(string message,bool quota):base(message){Quota=quota;} }
    public static class Translator {
        public static readonly string[] Languages={"vi","en","ja","ko","zh-CN","zh-TW","fr","de","th","es","ru","id"};
        public static bool ValidEmail(string s){return s==""||(s.Length<=254&&!s.Contains("..")&&Regex.IsMatch(s,@"^[^\s@]+@[^\s@.]+(?:\.[^\s@.]+)+$"));}
        public static List<string> Parts(string text){
            var parts=new List<string>();int start=0;
            while(start<text.Length){int end=start,bytes=0,lastSpace=-1;
                while(end<text.Length){int length=char.IsHighSurrogate(text[end])&&end+1<text.Length&&char.IsLowSurrogate(text[end+1])?2:1;int n=Encoding.UTF8.GetByteCount(text.Substring(end,length));if(bytes+n>480)break;bytes+=n;if(char.IsWhiteSpace(text[end]))lastSpace=end;end+=length;}
                if(end<text.Length&&lastSpace>start)end=lastSpace+1;
                if(end==start)throw new ArgumentException("Không chia được văn bản dịch.");parts.Add(text.Substring(start,end-start));start=end;
            }return parts;
        }
        public static string GoogleResponse(string json){
            object[] root=new JavaScriptSerializer().DeserializeObject(json) as object[];if(root==null||root.Length==0)throw new TranslationException("Không đọc được bản dịch Google web.",false);
            var segments=root[0] as object[];if(segments==null)throw new TranslationException("Google web không trả bản dịch.",false);var b=new StringBuilder();
            foreach(object segment in segments){var a=segment as object[];if(a!=null&&a.Length>0&&a[0] is string)b.Append((string)a[0]);}
            if(b.Length==0)throw new TranslationException("Google web không trả bản dịch.",false);return b.ToString();
        }
        public static string MemoryResponse(string json){
            var root=new JavaScriptSerializer().DeserializeObject(json) as Dictionary<string,object>;if(root==null)throw new TranslationException("Không đọc được bản dịch MyMemory.",false);
            object quota,status;root.TryGetValue("quotaFinished",out quota);root.TryGetValue("responseStatus",out status);
            if(Convert.ToString(quota).Equals("true",StringComparison.OrdinalIgnoreCase)||Convert.ToString(status)=="429")throw new TranslationException("MyMemory hết hạn mức hoặc giới hạn lượt gọi.",true);
            if(Convert.ToString(status)!="200")throw new TranslationException("MyMemory từ chối yêu cầu dịch.",false);
            object data;root.TryGetValue("responseData",out data);var d=data as Dictionary<string,object>;object translated;
            if(d==null||!d.TryGetValue("translatedText",out translated)||string.IsNullOrEmpty(Convert.ToString(translated)))throw new TranslationException("MyMemory không trả bản dịch.",false);
            return WebUtility.HtmlDecode(Convert.ToString(translated));
        }
        static string Request(string provider,string text,Preferences p){
            string url=provider=="googleweb"?"https://translate.googleapis.com/translate_a/single?client=gtx&dt=t&sl="+Codes.Encode(p.Source)+"&tl="+Codes.Encode(p.Target)+"&q="+Codes.Encode(text):"https://api.mymemory.translated.net/get?q="+Codes.Encode(text)+"&langpair="+Codes.Encode(p.Source+"|"+p.Target)+(p.Email==""?"":"&de="+Codes.Encode(p.Email));
            try{string json=Encoding.UTF8.GetString(Network.Get(url));return provider=="googleweb"?GoogleResponse(json):MemoryResponse(json);}
            catch(WebException ex){var response=ex.Response as HttpWebResponse;throw new TranslationException("Không kết nối được "+(provider=="googleweb"?"Google web":"MyMemory")+". Thử lại sau hoặc đổi dịch vụ.",response!=null&&(int)response.StatusCode==429);}
        }
        public static string Translate(string text,Preferences p,Func<bool> cancelled,Action<int,int> progress){
            if(text.Length==0||text.Length>20000)throw new ArgumentException("Chọn từ 1 đến 20.000 ký tự để dịch.");
            if(!Languages.Contains(p.Source)||!Languages.Contains(p.Target)||!ValidEmail(p.Email))throw new ArgumentException("Kiểm tra ngôn ngữ và email trong Cài đặt dịch.");
            if(p.Provider!="auto"&&p.Provider!="mymemory"&&p.Provider!="googleweb")throw new ArgumentException("Dịch vụ dịch chưa hợp lệ.");
            if(p.Source==p.Target)return text;
            var result=new StringBuilder();var parts=Parts(text);int done=0;
            foreach(string part in parts){if(cancelled())throw new OperationCanceledException();progress(done,parts.Count);Application.DoEvents();
                if(string.IsNullOrWhiteSpace(part))result.Append(part);
                else if(p.Provider=="googleweb")result.Append(Request("googleweb",part,p));
                else {try{result.Append(Request("mymemory",part,p));}catch(TranslationException ex){if(p.Provider=="auto"&&ex.Quota)result.Append(Request("googleweb",part,p));else throw;}}
                done++;if(cancelled())throw new OperationCanceledException();
            }progress(done,parts.Count);return result.ToString();
        }
    }

    public sealed partial class WordEngine {
        readonly dynamic app;public WordEngine(object application){app=application;}
        public object SelectionRange(){return app.Selection.Range.Duplicate;}
        public static void Validate(object range,bool requireText){
            dynamic r=range,d=r.Document;
            if((int)r.StoryType!=1)throw new InvalidOperationException("Đặt con trỏ hoặc chọn văn bản trong nội dung chính của tài liệu.");
            if((bool)d.ReadOnly||(int)d.ProtectionType!=-1)throw new InvalidOperationException("Tài liệu chỉ đọc hoặc đang được bảo vệ.");
            if((int)r.Fields.Count>0||(int)r.ContentControls.Count>0||(int)r.InlineShapes.Count>0)throw new InvalidOperationException("Chỉ chọn văn bản thuần, không gồm trường, ảnh hoặc điều khiển nội dung.");
            string s=(string)r.Text??"";
            if(s.IndexOf('\a')>=0)throw new InvalidOperationException("Chọn phần chữ bên trong ô bảng, không chọn dấu kết thúc ô hoặc nhiều ô cùng lúc.");
            if(requireText&&(s.Trim().Length==0||s.Length>100000))throw new InvalidOperationException("Chọn một đoạn văn bản từ 1 đến 100.000 ký tự.");
        }
        public object EditableSelection(bool requireText){
            dynamic r=app.Selection.Range.Duplicate;
            while(r.End>r.Start){string s=r.Text as string;if(s!=null&&(s.EndsWith("\r")||s.EndsWith("\a")))r.End=r.End-1;else break;}
            Validate((object)r,requireText);return r;
        }
        void Edit(object range,string title,Action action){
            dynamic r=range;Validate(range,false);bool started=false;
            try{app.UndoRecord.StartCustomRecord(title);started=true;action();}
            finally{if(started)app.UndoRecord.EndCustomRecord();}
        }
        public void Transform(object range,string mode){
            dynamic r=range;Validate(range,true);string original=(string)r.Text;
            // Word positions also count objects which Range.Text may omit.
            // Let Word change case in place rather than indexing a .NET string.
            if(mode=="upper"||mode=="lower"||mode=="title"||mode=="sentence"){
                int wordCase=mode=="upper"?1:mode=="lower"?0:mode=="title"?2:4;
                Edit(range,"Tiện ích Word - đổi chữ",()=>{r.Case=wordCase;});
                return;
            }
            if(mode!="clean")throw new ArgumentException("Kiểu xử lý văn bản không hợp lệ.");
            var plan=new TextPlan(original);
            var edits=new List<TextEdit>();
            string[] tokens=TextTools.CleanTokens(original);
            for(int i=0;i<original.Length;i++){
                if(tokens[i]!=original[i].ToString())edits.Add(new TextEdit{Start=i,Length=1,Value=tokens[i]});
            }
            plan.Add(edits);ApplyPlan(range,plan,false);
        }

        public string Number(object range,bool english,string unit){
            dynamic r=range;Validate(range,true);string words=NumberWords.Read(NumberWords.Parse((string)r.Text),english,unit);
            Edit(range,"Tiện ích Word - đọc số",()=>{dynamic insert=r.Duplicate;insert.Collapse(0);insert.Text=" ("+words+")";});return words;
        }
        public object InsertPicture(object range,byte[] bytes,string kind){
            Validate(range,false);var imageInfo=OfficeImage.Read(bytes,kind=="vietqr");int pixelWidth=imageInfo.Width,pixelHeight=imageInfo.Height;
            dynamic r=range;float available=(float)((double)r.Sections[1].PageSetup.PageWidth-(double)r.Sections[1].PageSetup.LeftMargin-(double)r.Sections[1].PageSetup.RightMargin);
            if((bool)r.Information[12])available=Math.Min(available,(float)r.Cells[1].Width-12);
            bool barcode=kind=="code128"||kind=="code39"||kind=="ean13";
            float width=barcode?pixelWidth/4f:Math.Min(150f,available);
            if(width>available)throw new InvalidOperationException("Mã vạch dài hơn vùng in. Rút ngắn nội dung, đổi trang ngang hoặc dùng QR; không tự ép nhỏ làm mờ mã.");
            string file=Path.Combine(Path.GetTempPath(),"TienIchWord-"+Guid.NewGuid().ToString("N")+(imageInfo.Svg?".svg":".png"));File.WriteAllBytes(file,bytes);object picture=null;bool changed=false;
            try{Edit(range,"Tiện ích Word - chèn mã",()=>{dynamic insert=r.Duplicate;insert.Collapse(0);dynamic pic=r.Document.InlineShapes.AddPicture(file,false,true,insert);changed=true;pic.LockAspectRatio=-1;pic.Width=width;pic.AlternativeText="TienIchWord:1:"+kind;picture=pic;});return picture;}
            catch{if(changed)r.Document.Undo(1);throw;}
            finally{try{File.Delete(file);}catch{}}
        }
        public void AppendTranslation(object range,string expected,string translated){
            dynamic r=range;Validate(range,true);if((string)r.Text!=expected)throw new InvalidOperationException("Văn bản gốc đã thay đổi trong lúc dịch. Chưa chèn kết quả; hãy chọn lại và dịch.");
            Edit(range,"Tiện ích Word - chèn bản dịch",()=>{dynamic insert=r.Paragraphs.Last.Range.Duplicate;insert.End=insert.End-1;insert.Collapse(0);insert.Text="\r"+translated.Replace("\r\n","\r").Replace("\n","\r");});
        }
    }

    [ComVisible(true),Guid("FFB307E9-CC8F-4C5F-9E72-948BE78B018A"),ProgId("HoPhuongLam.WordUtilities"),ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect:IDTExtensibility2,IRibbonExtensibility {
        dynamic app,ribbon;Preferences prefs;bool busy;
        public Connect(){TraceStartup("Constructed");}
        public const string ProgIdValue="HoPhuongLam.WordUtilities", ClassId="{FFB307E9-CC8F-4C5F-9E72-948BE78B018A}";
        public void OnConnection(object application,int connectMode,object addin,ref Array custom){TraceStartup("OnConnection");app=application;prefs=Preferences.Load();try{((dynamic)addin).Object=this;}catch(Exception ex){TraceStartup(ex.ToString());}TraceStartup("Connected");}
        static void TraceStartup(string message){try{Directory.CreateDirectory(Preferences.Folder);File.AppendAllText(Path.Combine(Preferences.Folder,"startup.log"),DateTime.Now.ToString("s")+" "+message+Environment.NewLine);}catch{}}
        public void OnDisconnection(int removeMode,ref Array custom){ribbon=null;app=null;}
        public void OnAddInsUpdate(ref Array custom){}public void OnStartupComplete(ref Array custom){}public void OnBeginShutdown(ref Array custom){}
        public string GetCustomUI(string ribbonId){TraceStartup("GetCustomUI "+ribbonId);return Resources.Text("ribbon.xml");}
        public void OnRibbonLoad(object ui){ribbon=ui;TraceStartup("Ribbon loaded");}
        [return:MarshalAs(UnmanagedType.IDispatch)]
        public object LoadRibbonImage(string imageId){try{return RibbonImages.Load(imageId);}catch(Exception ex){TraceStartup("Image "+imageId+": "+ex.Message);return null;}}
        public string Diagnostics(){return "TienIchWord 2.30.4 | Word connected="+(app!=null);}
        IWin32Window Owner(){try{return new OwnerWindow(new IntPtr((int)app.ActiveWindow.Hwnd));}catch{return null;}}
        void Run(Action action){if(busy){MessageBox.Show(Owner(),"Đang xử lý. Hãy chờ hoàn tất hoặc dừng lượt dịch.","Tiện ích Word");return;}busy=true;try{action();}catch(OperationCanceledException){MessageBox.Show(Owner(),"Đã hủy thao tác.","Tiện ích Word");}catch(Exception ex){HealthLog.Error(ex);MessageBox.Show(Owner(),ex.Message+"\n\nCó thể mở Kiểm tra add-in để xuất nhật ký lỗi.","Tiện ích Word",MessageBoxButtons.OK,MessageBoxIcon.Warning);}finally{busy=false;}}
        WordEngine CreateEngine(){return new WordEngine(app){ShowWorkProgress=true};}
        public void OnHealth(object control){Run(()=>{using(var form=new HealthForm(Diagnostics()))form.ShowDialog(Owner());});}
        object CheckedSelection(WordEngine engine,bool requireText){
            object range=engine.EditableSelection(requireText);int count=((string)((dynamic)range).Text??"").Length;
            if(prefs.ShouldWarn(count)&&MessageBox.Show(Owner(),"Bạn đang chọn "+count.ToString("N0")+" ký tự, vượt ngưỡng cảnh báo "+prefs.SelectionWarningLimit.ToString("N0")+".\n\nThao tác trên vùng lớn có thể làm Word chậm. Nên chia thành từng đoạn nhỏ. Bạn có muốn tiếp tục?\n\nCó thể đổi ngưỡng trong Cài đặt add-in.","Vùng chọn nhiều văn bản",MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2)!=DialogResult.Yes)throw new OperationCanceledException();
            return range;
        }
        public void OnGeneralSettings(object control){Run(()=>{using(var form=new GeneralSettingsForm(prefs))form.ShowDialog(Owner());});}
        public void OnText(object control){Run(()=>{var engine=CreateEngine();engine.Transform(CheckedSelection(engine,true),(string)((dynamic)control).Tag);});}
        public void OnNumber(object control){Run(()=>{var engine=CreateEngine();engine.Number(CheckedSelection(engine,true),(string)((dynamic)control).Tag=="en",prefs.Unit);});}
        public bool GetUnitPressed(object control){return prefs!=null&&prefs.Unit==(string)((dynamic)control).Tag;}
        public void OnUnit(object control,bool pressed){Run(()=>{prefs.Unit=(string)((dynamic)control).Tag;prefs.Save();if(ribbon!=null)ribbon.Invalidate();});}
        public void OnVietQR(object control){Run(()=>{var engine=CreateEngine();object range=CheckedSelection(engine,false);using(var form=new PaymentForm(prefs,engine,range))form.ShowDialog(Owner());});}
        public void OnCode(object control){Run(()=>{var engine=CreateEngine();object range=CheckedSelection(engine,false);using(var form=new CodeForm((string)((dynamic)control).Tag,engine,range))form.ShowDialog(Owner());});}
        public void OnTranslationSettings(object control){Run(()=>{using(var form=new TranslationSettingsForm(prefs))form.ShowDialog(Owner());});}
        public void OnNormalize(object control){Run(()=>{var engine=CreateEngine();using(var form=new NormalizeForm(engine,CheckedSelection(engine,true),false))form.ShowDialog(Owner());});}
        public void OnEncoding(object control){Run(()=>{var engine=CreateEngine();using(var form=new NormalizeForm(engine,CheckedSelection(engine,true),true))form.ShowDialog(Owner());});}
        public void OnDocumentFormat(object control){Run(()=>{using(var form=new DocumentFormatForm(CreateEngine(),SelectionForFormat()))form.ShowDialog(Owner());});}
        object SelectionForFormat(){return app.Selection.Range.Duplicate;}
        public void OnTranslate(object control){Run(()=>{
            var engine=CreateEngine();object range=CheckedSelection(engine,true);string text=(string)((dynamic)range).Text;
            if(text.Length>20000)throw new InvalidOperationException("Chọn tối đa 20.000 ký tự mỗi lần dịch.");
            string provider=prefs.Provider=="auto"?"MyMemory; dự phòng Google web khi hết hạn mức":prefs.Provider=="mymemory"?"MyMemory":"Google web";
            if(MessageBox.Show(Owner(),"Gửi đoạn văn đang chọn ("+text.Length+" ký tự) tới "+provider+" để dịch từ "+prefs.Source+" sang "+prefs.Target+"?\n\nBản dịch được chèn dưới đoạn gốc. Có thể Ctrl+Z sau khi chèn.","Xác nhận dịch ngôn ngữ",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2)!=DialogResult.Yes)return;
            busy=true;try{using(var progress=new ProgressForm()){progress.Show(Owner());string result=Translator.Translate(text,prefs,()=>progress.Cancelled,(a,b)=>progress.UpdateProgress(a,b));progress.Finish();engine.AppendTranslation(range,text,result);}}finally{busy=false;}
        });}
        public void OnAbout(object control){Run(()=>{using(var form=new AboutForm())form.ShowDialog(Owner());});}
    }
    public sealed class OwnerWindow:IWin32Window {public IntPtr Handle{get;private set;}public OwnerWindow(IntPtr handle){Handle=handle;}}
    public static class Resources {
        public static string Text(string name){using(var stream=Assembly.GetExecutingAssembly().GetManifestResourceStream("TienIchWord."+name))using(var reader=new StreamReader(stream,Encoding.UTF8))return reader.ReadToEnd();}
        public static Image Picture(string name){using(var stream=Assembly.GetExecutingAssembly().GetManifestResourceStream("TienIchWord."+name))using(var image=Image.FromStream(stream))return new Bitmap(image);}
    }
}
