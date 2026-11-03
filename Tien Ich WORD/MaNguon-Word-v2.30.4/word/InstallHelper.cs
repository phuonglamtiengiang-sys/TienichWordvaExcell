using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
class WordInstallHelper {
    static int Main(string[] args) {
        try {
            if(args.Length!=1 || !File.Exists(args[0])) return 2;
            if(Process.GetProcessesByName("WINWORD").Length>0) return 3;
            string path=Path.GetFullPath(args[0]);
            Assembly addin=Assembly.LoadFrom(path);
            Type type=addin.GetType("TienIchWord.Registration",true);
            type.GetMethod("Install",new[]{typeof(string)}).Invoke(null,new object[]{path});
            return 0;
        } catch { return 10; }
    }
}
