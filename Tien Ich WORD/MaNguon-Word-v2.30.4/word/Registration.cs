using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using Microsoft.Win32;
namespace TienIchWord {
    public static class Registration {
        public static string InstallDirectory {get{return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"TienIchWord","Addin");}}
        static RegistryView[] Views {get{return Environment.Is64BitOperatingSystem?new[]{RegistryView.Registry32,RegistryView.Registry64}:new[]{RegistryView.Registry32};}}
        public static void Register(string assemblyPath,string progId,int loadBehavior){
            assemblyPath=Path.GetFullPath(assemblyPath);
            if(!File.Exists(assemblyPath))throw new FileNotFoundException("Không tìm thấy tệp add-in.");
            string assembly=AssemblyName.GetAssemblyName(assemblyPath).FullName;
            foreach(var view in Views)using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)){
                using(var cls=user.CreateSubKey(@"Software\Classes\CLSID\"+Connect.ClassId)){
                    cls.SetValue("",typeof(Connect).FullName);
                    using(var id=cls.CreateSubKey("ProgId"))id.SetValue("",progId);
                    using(var programmable=cls.CreateSubKey("Programmable")){}
                    using(var server=cls.CreateSubKey("InprocServer32")){
                        server.SetValue("","mscoree.dll");server.SetValue("ThreadingModel","Both");
                        SetManaged(server,assembly,assemblyPath);
                        using(var version=server.CreateSubKey(AssemblyName.GetAssemblyName(assemblyPath).Version.ToString()))SetManaged(version,assembly,assemblyPath);
                    }
                    using(var category=cls.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}")){}
                }
                using(var prog=user.CreateSubKey(@"Software\Classes\"+progId)){prog.SetValue("","Tiện ích Word Việt Nam");using(var cls=prog.CreateSubKey("CLSID"))cls.SetValue("",Connect.ClassId);}
                using(var addin=user.CreateSubKey(@"Software\Microsoft\Office\Word\Addins\"+progId)){
                    addin.SetValue("FriendlyName","Tiện ích Word Việt Nam");addin.SetValue("Description","Đọc số, văn bản, VietQR, mã vạch/QR, dịch ngôn ngữ — Hồ Phương Lâm");
                    addin.SetValue("LoadBehavior",loadBehavior,RegistryValueKind.DWord);addin.SetValue("CommandLineSafe",0,RegistryValueKind.DWord);
                }
            }
        }
        static void SetManaged(RegistryKey key,string assembly,string path){key.SetValue("Class",typeof(Connect).FullName);key.SetValue("Assembly",assembly);key.SetValue("RuntimeVersion","v4.0.30319");key.SetValue("CodeBase",new Uri(path).AbsoluteUri);}
        public static void Remove(string progId){
            foreach(var view in Views)using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)){
                user.DeleteSubKeyTree(@"Software\Microsoft\Office\Word\Addins\"+progId,false);
                user.DeleteSubKeyTree(@"Software\Classes\"+progId,false);
                user.DeleteSubKeyTree(@"Software\Classes\CLSID\"+Connect.ClassId,false);
            }
        }
        public static void Install(){
            Install(typeof(Connect).Assembly.Location);
        }
        public static void Install(string sourcePath){
            if(Process.GetProcessesByName("WINWORD").Length>0)throw new InvalidOperationException("Hãy lưu tài liệu và đóng tất cả cửa sổ Word trước khi cài hoặc gỡ.");
            Directory.CreateDirectory(InstallDirectory);
            string destination=Path.Combine(InstallDirectory,"TienIchWord.dll");
            string pending=destination+".update";
            try{
                if(!string.Equals(Path.GetFullPath(sourcePath),Path.GetFullPath(destination),StringComparison.OrdinalIgnoreCase)){
                    File.Copy(sourcePath,pending,true);
                    if(File.Exists(destination))File.Replace(pending,destination,null);
                    else File.Move(pending,destination);
                }
            }catch(IOException ex){
                try{if(File.Exists(pending))File.Delete(pending);}catch{}
                throw new InvalidOperationException("Không thể thay add-in vì TienIchWord.dll vẫn đang được một tiến trình sử dụng. Hãy thoát hẳn Word (kể cả WINWORD.EXE chạy nền) rồi bấm Cài đặt / cập nhật lần nữa.",ex);
            }
            Register(destination,Connect.ProgIdValue,3);
            RestoreOwnDisabledItems();
        }
        public static void RestoreOwnDisabledItems(){
            // Office may persist a disabled item with an unaligned UTF-16 filename.
            // Back up and remove only entries that name this exact add-in DLL.
            foreach(var view in Views)using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
            using(var office=user.OpenSubKey(@"Software\Microsoft\Office")){
                if(office==null)continue;
                foreach(string version in office.GetSubKeyNames()){
                    Version parsed;if(!Version.TryParse(version,out parsed))continue;
                    using(var disabled=user.OpenSubKey(@"Software\Microsoft\Office\"+version+@"\Word\Resiliency\DisabledItems",true)){
                        if(disabled==null)continue;
                        foreach(string name in disabled.GetValueNames()){
                            byte[] data=disabled.GetValue(name) as byte[];if(data==null)continue;
                            string text=Encoding.UTF8.GetString(data)+Encoding.Unicode.GetString(data);
                            if(data.Length>1)text+=Encoding.Unicode.GetString(data,1,data.Length-1);
                            if(text.IndexOf("TienIchWord.dll",StringComparison.OrdinalIgnoreCase)<0)continue;
                            string backup=Path.Combine(Preferences.Folder,"Recovery");Directory.CreateDirectory(backup);
                            File.WriteAllBytes(Path.Combine(backup,version+"-"+view+"-"+name+".bin"),data);
                            disabled.DeleteValue(name,false);
                        }
                    }
                }
            }
        }
        public static void Uninstall(){
            if(Process.GetProcessesByName("WINWORD").Length>0)throw new InvalidOperationException("Hãy lưu tài liệu và đóng tất cả cửa sổ Word trước khi gỡ.");
            Remove(Connect.ProgIdValue);
        }
    }
}
