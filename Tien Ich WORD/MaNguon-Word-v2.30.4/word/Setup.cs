using System;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
[assembly: AssemblyTitle("Cai dat Tien Ich Word Viet Nam")]
[assembly: AssemblyDescription("Trinh cai dat tien ich Microsoft Word, phien ban 2.30.4")]
[assembly: AssemblyCompany("Ho Phuong Lam")]
[assembly: AssemblyProduct("Tien Ich Word Viet Nam")]
[assembly: AssemblyVersion("2.30.4.0")]
[assembly: AssemblyFileVersion("2.30.4.0")]
[assembly: ComVisible(false)]
class WordSetup:Form {
    readonly Label status;
    static string AddinDirectory{get{return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"TienIchWord","Addin");}}
    static void Extract(string resource,string file){using(Stream input=Assembly.GetExecutingAssembly().GetManifestResourceStream(resource)){if(input==null)throw new InvalidOperationException("Bộ cài thiếu thành phần: "+resource);Directory.CreateDirectory(Path.GetDirectoryName(file));using(var output=new FileStream(file,FileMode.Create,FileAccess.Write,FileShare.None))input.CopyTo(output);}}
    static readonly string ProgId="HoPhuongLam.WordUtilities";
    static readonly string ClassId="{FFB307E9-CC8F-4C5F-9E72-948BE78B018A}";
    static RegistryView[] Views{get{return Environment.Is64BitOperatingSystem?new[]{RegistryView.Registry32,RegistryView.Registry64}:new[]{RegistryView.Registry32};}}
    static void Register(string path){
        string identity=AssemblyName.GetAssemblyName(path).FullName;
        foreach(RegistryView view in Views)using(RegistryKey root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)){
            using(RegistryKey cls=root.CreateSubKey(@"Software\Classes\CLSID\"+ClassId)){
                cls.SetValue("","TienIchWord.Connect");using(RegistryKey id=cls.CreateSubKey("ProgId"))id.SetValue("",ProgId);cls.CreateSubKey("Programmable").Close();
                using(RegistryKey server=cls.CreateSubKey("InprocServer32")){server.SetValue("","mscoree.dll");server.SetValue("ThreadingModel","Both");SetManaged(server,identity,path);using(RegistryKey version=server.CreateSubKey("2.30.4.0"))SetManaged(version,identity,path);}
                cls.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}").Close();
            }
            using(RegistryKey prog=root.CreateSubKey(@"Software\Classes\"+ProgId)){prog.SetValue("","Tiện ích Word Việt Nam");using(RegistryKey cls=prog.CreateSubKey("CLSID"))cls.SetValue("",ClassId);}
            using(RegistryKey addin=root.CreateSubKey(@"Software\Microsoft\Office\Word\Addins\"+ProgId)){addin.SetValue("FriendlyName","Tiện ích Word Việt Nam");addin.SetValue("Description","Đọc số, xử lý văn bản, VietQR, mã vạch/QR, dịch ngôn ngữ");addin.SetValue("LoadBehavior",3,RegistryValueKind.DWord);addin.SetValue("CommandLineSafe",0,RegistryValueKind.DWord);}
        }
    }
    static void SetManaged(RegistryKey key,string identity,string path){key.SetValue("Class","TienIchWord.Connect");key.SetValue("Assembly",identity);key.SetValue("RuntimeVersion","v4.0.30319");key.SetValue("CodeBase",new Uri(path).AbsoluteUri);}
    static void Install(){
        if(Process.GetProcessesByName("WINWORD").Length>0)throw new InvalidOperationException("Hãy lưu tài liệu và thoát hẳn Word, kể cả tiến trình WINWORD.EXE chạy nền, rồi thử lại.");
        Directory.CreateDirectory(AddinDirectory);
        string stage=Path.Combine(Path.GetTempPath(),"TienIchWord-"+Guid.NewGuid().ToString("N")+".dll");
        try{
            Extract("TienIchWord.Addin.dll",stage);
            string destination=Path.Combine(AddinDirectory,"TienIchWord.dll");string pending=destination+".update";
            File.Copy(stage,pending,true);
            if(File.Exists(destination))File.Replace(pending,destination,null);else File.Move(pending,destination);
            Register(destination);
            string data=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"TienIchWord");
            Extract("TienIchWord.Source.zip",Path.Combine(data,"MaNguon-Word-v2.30.4.zip"));
            Extract("TienIchWord.References.md",Path.Combine(data,"THAM-KHAO.md"));
        }finally{try{if(File.Exists(stage))File.Delete(stage);}catch{}}
    }
    static void Uninstall(){
        if(Process.GetProcessesByName("WINWORD").Length>0)throw new InvalidOperationException("Hãy lưu tài liệu và đóng Word trước khi gỡ.");
        foreach(RegistryView view in Views)using(RegistryKey root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)){
            root.DeleteSubKeyTree(@"Software\Microsoft\Office\Word\Addins\"+ProgId,false);
            root.DeleteSubKeyTree(@"Software\Classes\"+ProgId,false);
            root.DeleteSubKeyTree(@"Software\Classes\CLSID\"+ClassId,false);
        }
    }
    WordSetup(){
        Text="Cài đặt Tiện ích Word Việt Nam 2.30.4";ClientSize=new Size(650,370);Font=new Font("Segoe UI",10);FormBorderStyle=FormBorderStyle.FixedDialog;MaximizeBox=false;StartPosition=FormStartPosition.CenterScreen;
        Controls.Add(new Label{Text="TIỆN ÍCH WORD VIỆT NAM",Left=20,Top=20,Width=610,Height=32,Font=new Font("Segoe UI",16,FontStyle.Bold)});
        Controls.Add(new Label{Text="Đổi số thành chữ • Xử lý văn bản • VietQR • Mã vạch/QR\nDịch ngôn ngữ • Giới thiệu tác giả",Left=20,Top=70,Width=610,Height=60});
        Controls.Add(new Label{Text="1. Lưu tài liệu và đóng Word.\n2. Bấm Cài đặt cho tài khoản hiện tại.\n3. Mở Word để dùng tab Tiện ích Word.\n\nDành cho Word Windows 32/64-bit. Không thay đổi thiết lập bảo mật macro. Không cần quyền quản trị thông thường.",Left=20,Top=140,Width=610,Height=115});
        status=new Label{Text="Tác giả: Hồ Phương Lâm — hophuonglam@outlook.com",Left=20,Top=263,Width=610,Height=52};Controls.Add(status);
        var install=new Button{Text="Cài đặt / cập nhật",Left=20,Top=320,Width=245,Height=36};install.Click+=(s,e)=>{try{Install();status.Text="Đã cài. Mở Word và tìm tab Tiện ích Word.\nMã nguồn và tài liệu tham khảo cũng đã được lưu.";}catch(Exception ex){Exception error=ex.InnerException??ex;MessageBox.Show(this,error.Message,"Cài đặt Word",MessageBoxButtons.OK,MessageBoxIcon.Warning);}};Controls.Add(install);
        var remove=new Button{Text="Gỡ khỏi Word",Left=280,Top=320,Width=200,Height=36};remove.Click+=(s,e)=>{try{Uninstall();status.Text="Đã gỡ add-in khỏi Word. Mã nguồn và cấu hình cá nhân được giữ lại.";}catch(Exception ex){Exception error=ex.InnerException??ex;MessageBox.Show(this,error.Message,"Gỡ add-in",MessageBoxButtons.OK,MessageBoxIcon.Warning);}};Controls.Add(remove);
        var close=new Button{Text="Đóng",Left=495,Top=320,Width=135,Height=36};close.Click+=(s,e)=>Close();Controls.Add(close);CancelButton=close;
    }
    [STAThread]static void Main(){Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);Application.Run(new WordSetup());}
}
