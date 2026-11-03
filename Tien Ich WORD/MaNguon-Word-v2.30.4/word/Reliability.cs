using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace TienIchWord {
    public static class HealthLog {
        public static string PathName {get{return Path.Combine(Preferences.Folder,"health.log");}}
        public static void Error(Exception ex){try{Directory.CreateDirectory(Preferences.Folder);if(File.Exists(PathName)&&new FileInfo(PathName).Length>1048576)File.Move(PathName,PathName+"."+DateTime.Now.Ticks+".old");File.AppendAllText(PathName,DateTime.Now.ToString("s")+" Word 2.30.4 "+ex.GetType().Name+" HRESULT="+ex.HResult.ToString("X8")+Environment.NewLine);}catch{}}
        public static string Report(string state){return state+Environment.NewLine+"Process: "+(Environment.Is64BitProcess?"64-bit":"32-bit")+Environment.NewLine+"CLR: "+Environment.Version+Environment.NewLine+"Log records error codes only; no document text."+Environment.NewLine+(File.Exists(PathName)?File.ReadAllText(PathName):"No error records.");}
    }
    public sealed class HealthForm:UtilityForm {
        public HealthForm(string state):base("Kiểm tra add-in Word",650,390){
            var text=new TextBox{Left=16,Top=18,Width=618,Height=280,Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Both,Text=HealthLog.Report(state)};Controls.Add(text);
            ButtonAt("Xuất nhật ký lỗi",16,330,300,(s,e)=>{using(var file=new SaveFileDialog{Filter="Text (*.txt)|*.txt",FileName="TienIchWord-chan-doan.txt"})if(file.ShowDialog(this)==DialogResult.OK)try{File.WriteAllText(file.FileName,text.Text,Encoding.UTF8);}catch(Exception ex){MessageBox.Show(this,ex.Message);}});
            ButtonAt("Đóng",332,330,300,(s,e)=>Close());
        }
    }
    public sealed class WorkForm:UtilityForm {
        readonly Label status;bool finished;Exception error;public bool Cancelled{get;private set;}
        public WorkForm():base("Đang xử lý văn bản",480,165){status=LabelAt("Đang chuẩn bị...",16,18,448,75);ButtonAt("Dừng và hoàn tác",16,110,448,(s,e)=>{Cancelled=true;status.Text="Đang dừng và khôi phục thay đổi...";});FormClosing+=(s,e)=>{if(!finished){Cancelled=true;e.Cancel=true;}};}
        public void Update(int done,int total,string phase){status.Text=phase+"\n"+done.ToString("N0")+" / "+total.ToString("N0")+". Dừng sẽ hoàn tác lượt này.";status.Refresh();}
        public void Execute(IWin32Window owner,Action action){Shown+=(s,e)=>BeginInvoke((Action)(()=>{try{action();}catch(Exception ex){error=ex;}finally{finished=true;Close();}}));ShowDialog(owner);if(error!=null)throw error;}
    }
    public sealed partial class WordEngine {
        public bool ShowWorkProgress{get;set;}
        public Func<bool> CancelRequested{get;set;}
        public Action<int,int> WorkProgress{get;set;}
        sealed class Segment {public int TextStart,Length,Position;}
        void MapSegments(dynamic document,int start,int end,List<Segment> segments,StringBuilder text,Action<int,int,string> tick){
            tick(text.Length,end-start,"Đang xác định vùng văn bản");
            dynamic part=document.Range(start,end);string value=(string)part.Text??"";
            if(value.Length==0)return;
            if(end-start==value.Length){segments.Add(new Segment{TextStart=text.Length,Length=value.Length,Position=start});text.Append(value);return;}
            if(end-start<=1)throw new InvalidOperationException("Không xác định được vị trí chữ an toàn. Hãy chọn một đoạn nhỏ hơn.");
            int middle=start+(end-start)/2;MapSegments(document,start,middle,segments,text,tick);MapSegments(document,middle,end,segments,text,tick);
        }
        public void ApplyPlan(object range,TextPlan plan,bool unicodeFont){
            dynamic r=range;Validate(range,true);if((string)r.Text!=plan.Original)throw new InvalidOperationException("Vùng chọn đã đổi. Hãy chọn lại.");
            if((bool)r.Document.TrackRevisions||(int)r.Revisions.Count>0)throw new InvalidOperationException("Vùng chọn có Theo dõi thay đổi (Track Changes). Hãy xem xét các thay đổi, rồi tắt theo dõi trước khi làm sạch hoặc đổi bảng mã. Chưa sửa tài liệu.");
            if(plan.Stages.Count==0&&!unicodeFont)return;
            WorkForm progress=null;var timer=Stopwatch.StartNew();bool changed=false;int done=0,total=plan.Stages.Sum(s=>s.Count);
            Action<int,int,string> tick=(n,count,phase)=>{
                if(WorkProgress!=null)WorkProgress(n,count);
                if(CancelRequested!=null&&CancelRequested())throw new OperationCanceledException();
                if(progress!=null&&timer.ElapsedMilliseconds>=100){progress.Update(n,count,phase);Application.DoEvents();timer.Restart();}
                if(progress!=null&&progress.Cancelled)throw new OperationCanceledException();
            };
            Action work=()=>{
                bool screen=app.ScreenUpdating;
                try{app.ScreenUpdating=false;Edit(range,"Tiện ích Word - chuẩn hóa / chuyển mã",()=>{
                    foreach(var stage in plan.Stages){
                        int stageStart=r.Start,stageEnd=r.End,delta=0;string before=(string)r.Text;
                        var segments=new List<Segment>();var text=new StringBuilder();
                        MapSegments(r.Document,stageStart,stageEnd,segments,text,tick);
                        if(text.ToString()!=before)throw new InvalidOperationException("Không xác định được vùng chữ an toàn. Đã hủy lượt sửa.");
                        foreach(var e in stage.OrderByDescending(e=>e.Start)){
                            tick(done,total,"Đang áp dụng thay đổi");
                            var parts=new List<object>();int covered=0;
                            foreach(var segment in segments){
                                int first=Math.Max(e.Start,segment.TextStart),last=Math.Min(e.Start+e.Length,segment.TextStart+segment.Length);
                                if(first>=last)continue;
                                parts.Add((object)r.Document.Range(segment.Position+first-segment.TextStart,segment.Position+last-segment.TextStart));covered+=last-first;
                            }
                            if(covered!=e.Length)throw new InvalidOperationException("Vùng sửa không khớp văn bản. Đã hủy lượt sửa.");
                            for(int i=parts.Count-1;i>=0;i--){dynamic part=parts[i];int oldLength=part.End-part.Start;part.Text=i==0?e.Value:"";changed=true;delta+=(int)part.End-(int)part.Start-oldLength;}
                            done++;
                        }
                        r.SetRange(stageStart,stageEnd+delta);
                    }
                    tick(done,total,"Đang kiểm tra kết quả");
                    if((string)r.Text!=plan.Result)throw new InvalidOperationException("Kết quả không khớp bản xem trước. Đã hủy và hoàn tác.");
                    if(unicodeFont){r.Font.Name="Times New Roman";changed=true;}
                });}catch{if(changed)r.Document.Undo(1);throw;}finally{app.ScreenUpdating=screen;}
            };
            try{
                if(ShowWorkProgress){progress=new WorkForm();IWin32Window owner=null;try{owner=new OwnerWindow(new IntPtr((int)app.ActiveWindow.Hwnd));}catch{}progress.Execute(owner,work);}
                else work();
            }finally{if(progress!=null)progress.Dispose();}
        }
    }
}
