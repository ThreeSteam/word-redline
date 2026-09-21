using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;

public class Observable : INotifyPropertyChanged {
    public event PropertyChangedEventHandler PropertyChanged;
    public void Changed(params string[] names) { foreach(string n in names) { var h=PropertyChanged; if(h!=null) h(this,new PropertyChangedEventArgs(n)); } }
}
public class FileItem {
    public string FullPath {get;set;}
    public string Name {get {return Path.GetFileName(FullPath);}}
    public string DirectoryName {get {return Path.GetDirectoryName(FullPath);}}
}
public class Choice {public int Index {get;set;} public string Label {get;set;} public string Path {get;set;} public override string ToString(){return Label;} }
public class Pair : Observable {
    public string NewPath {get;set;}
    public string NewName {get{return Path.GetFileName(NewPath);}}
    public string OldPath {get;set;}
    public string Output {get;set;}
    bool selected=true; int oldIndex=-1; string status="待配对",detail="",state="pending";
    public Action<Pair> OnChoice;
    public bool Selected {get{return selected;}set{selected=value;Changed("Selected");}}
    public int OldIndex {get{return oldIndex;}set{oldIndex=value;Changed("OldIndex");if(OnChoice!=null)OnChoice(this);}}
    public string Status {get{return status;}}
    public string Detail {get{return detail;}}
    public string BadgeBackground {get{return state=="success"?"#EAF6EF":state=="error"?"#FFF0EF":state=="pending"?"#FFF6E6":"#EEF3FB";}}
    public string BadgeForeground {get{return state=="success"?"#248452":state=="error"?"#CA534A":state=="pending"?"#A77828":"#53749D";}}
    public void SetStatus(string label,string kind,string description=null) {status=label;state=kind;detail=description??label;Changed("Status","Detail","BadgeBackground","BadgeForeground","Output","OldPath");}
}
public static class OutputOptions {
    public static string Validate(string name) {
        if(string.IsNullOrWhiteSpace(name)) return "请输入子文件夹名称，例如 Legal。";
        if(name!=name.Trim() || name.EndsWith(".") || name=="." || name=="..") return "文件夹名称不能是点号，也不能以空格或句点结尾。";
        if(name.IndexOfAny(Path.GetInvalidFileNameChars())>=0 || name.Contains("/") || name.Contains("\\")) return "请输入单个文件夹名称，不要包含路径或 / \\ : 等字符。";
        if(Regex.IsMatch(name,@"^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\.|$)",RegexOptions.IgnoreCase)) return "该名称是 Windows 保留名称，请换一个文件夹名称。";
        if(name.Length>255) return "文件夹名称过长。";
        return null;
    }
}
public class RedlineModel : Observable {
    public ObservableCollection<FileItem> OldFiles {get;private set;}
    public ObservableCollection<FileItem> NewFiles {get;private set;}
    public ObservableCollection<Choice> OldChoices {get;private set;}
    public ObservableCollection<Pair> Pairs {get;private set;}
    public Window Window; DataGrid grid; ListBox oldList,newList;
    readonly string folder=AppDomain.CurrentDomain.BaseDirectory;
    readonly string session=Path.Combine(Path.GetTempPath(),"WordRedline",Guid.NewGuid().ToString("N"));
    readonly JavaScriptSerializer json=new JavaScriptSerializer();
    readonly DispatcherTimer timer=new DispatcherTimer();
    readonly Queue<Pair> queue=new Queue<Pair>();
    Process worker; Pair active; string job; DateTime started; bool busy,subfolder;
    string subfolderName="Legal",message="准备好文档后，就可以开始了。",summary="支持多文件批量配对";
    int total,finished,failed; bool batchSubfolder; string batchFolder; string testDirectory;
    public bool CanEdit {get{return !busy;}}
    public bool CanNameFolder {get{return !busy && subfolder;}}
    public bool CanOpen {get{return grid!=null && grid.SelectedItem is Pair && File.Exists(((Pair)grid.SelectedItem).Output);}}
    public Visibility BusyVisibility {get{return busy?Visibility.Visible:Visibility.Collapsed;}}
    public Visibility ProgressVisibility {get{return busy?Visibility.Visible:Visibility.Collapsed;}}
    public double Progress {get{return total==0?0:100.0*finished/total;}}
    public string Message {get{return message;}set{message=value;Changed("Message");}}
    public string PairSummary {get{return summary;}}
    public bool UseSubfolder {get{return subfolder;}set{subfolder=value;Changed("UseSubfolder","CanNameFolder","SaveHint","SaveHintColor");}}
    public string SubfolderName {get{return subfolderName;}set{subfolderName=value;Changed("SubfolderName","SaveHint","SaveHintColor");}}
    public string SaveHintColor {get{return subfolder && OutputOptions.Validate(subfolderName)!=null?"#C4483C":"#8090A5";}}
    public string SaveHint {get{
        string error=subfolder?OutputOptions.Validate(subfolderName):null;
        if(error!=null)return error;
        return subfolder?"结果位置：新版所在文件夹  /  "+subfolderName+"  /  文件名 - redline.docx    ·    不存在时自动创建":"结果保存在各新版文件旁，命名为「文件名 - redline.docx」；重名自动编号。";
    }}
    public RedlineModel() {OldFiles=new ObservableCollection<FileItem>();NewFiles=new ObservableCollection<FileItem>();OldChoices=new ObservableCollection<Choice>();Pairs=new ObservableCollection<Pair>();timer.Interval=TimeSpan.FromMilliseconds(400);timer.Tick+=(s,e)=>Poll();}
    T Find<T>(string name) where T:FrameworkElement {return (T)Window.FindName(name);}
    public void Wire(Window window) {
        Window=window;Window.DataContext=this;grid=Find<DataGrid>("PairsGrid");oldList=Find<ListBox>("OldList");newList=Find<ListBox>("NewList");
        foreach(string side in new[]{"Old","New"}) {
            bool isOld=side=="Old";
            Find<Button>("Add"+side).Click+=(s,e)=>Choose(isOld);
            Find<Button>("Remove"+side).Click+=(s,e)=>Remove(isOld);
            Border drop=Find<Border>(side+"Drop");
            drop.PreviewDragOver+=(s,e)=> {e.Effects=CanEdit && e.Data.GetDataPresent(DataFormats.FileDrop)?DragDropEffects.Copy:DragDropEffects.None;e.Handled=true;drop.BorderBrush=new SolidColorBrush(Color.FromRgb(37,99,235));};
            drop.DragLeave+=(s,e)=>drop.BorderBrush=new SolidColorBrush(Color.FromRgb(226,231,239));
            drop.PreviewDrop+=(s,e)=> {drop.BorderBrush=new SolidColorBrush(Color.FromRgb(226,231,239));if(CanEdit && e.Data.GetDataPresent(DataFormats.FileDrop))AddFiles((string[])e.Data.GetData(DataFormats.FileDrop),isOld);e.Handled=true;};
        }
        oldList.PreviewKeyDown+=(s,e)=>{if(e.Key==Key.Delete && CanEdit){Remove(true);e.Handled=true;}};
        newList.PreviewKeyDown+=(s,e)=>{if(e.Key==Key.Delete && CanEdit){Remove(false);e.Handled=true;}};
        Find<Button>("Start").Click+=(s,e)=>Begin();Find<Button>("Cancel").Click+=(s,e)=>Cancel();
        Find<Button>("Clear").Click+=(s,e)=>{OldFiles.Clear();NewFiles.Clear();Refresh();};
        Find<Button>("OpenOutput").Click+=(s,e)=>OpenOutput();grid.SelectionChanged+=(s,e)=>Changed("CanOpen");
        grid.MouseDoubleClick+=(s,e)=>{if(CanOpen)OpenOutput();};
        Window.Closing+=(s,e)=>{if(busy)Cancel();};
        using(var stream=Assembly.GetExecutingAssembly().GetManifestResourceStream("logo.png")) {
            var logo=new BitmapImage();logo.BeginInit();logo.CacheOption=BitmapCacheOption.OnLoad;logo.StreamSource=stream;logo.EndInit();logo.Freeze();Find<Image>("BrandIcon").Source=logo;
        }
        using(var stream=Assembly.GetExecutingAssembly().GetManifestResourceStream("app.ico")) {Window.Icon=BitmapFrame.Create(stream,BitmapCreateOptions.None,BitmapCacheOption.OnLoad);}
    }
    void Choose(bool isOld) {var d=new OpenFileDialog {Multiselect=true,Filter="Word 文档|*.docx;*.doc;*.docm"};if(d.ShowDialog(Window)==true)AddFiles(d.FileNames,isOld);}
    void Remove(bool isOld) {var list=isOld?oldList:newList;var files=isOld?OldFiles:NewFiles;foreach(var item in list.SelectedItems.Cast<FileItem>().ToArray())files.Remove(item);Refresh();}
    public void AddFiles(IEnumerable<string> paths,bool isOld) {
        if(busy)return;int ignored=0;var files=isOld?OldFiles:NewFiles;
        foreach(string path in paths) {
            if(!File.Exists(path)||!new[]{".docx",".doc",".docm"}.Contains(Path.GetExtension(path).ToLowerInvariant())||Path.GetFileName(path).StartsWith("~$")){ignored++;continue;}
            string full=Path.GetFullPath(path);if(!files.Any(f=>string.Equals(f.FullPath,full,StringComparison.OrdinalIgnoreCase)))files.Add(new FileItem {FullPath=full});
        }
        Refresh();if(ignored>0)Message+=" 已忽略 "+ignored+" 个不支持的文件。";
    }
    void Refresh() {
        Pairs.Clear();OldChoices.Clear();OldChoices.Add(new Choice {Index=-1,Label="请选择对应旧版"});
        for(int i=0;i<OldFiles.Count;i++)OldChoices.Add(new Choice {Index=i,Label=(i+1)+" · "+OldFiles[i].Name,Path=OldFiles[i].FullPath});
        int[] matches=MatchFiles.Suggest(OldFiles.Select(f=>f.FullPath).ToArray(),NewFiles.Select(f=>f.FullPath).ToArray());
        for(int i=0;i<NewFiles.Count;i++) {
            var p=new Pair {NewPath=NewFiles[i].FullPath,OldIndex=matches[i],OldPath=matches[i]<0?null:OldFiles[matches[i]].FullPath};
            p.SetStatus(matches[i]<0?"待手动配对":"已建议配对",matches[i]<0?"pending":"ready");
            p.OnChoice=row=>{row.OldPath=row.OldIndex<0?null:OldFiles[row.OldIndex].FullPath;row.SetStatus(row.OldIndex<0?"待手动配对":"已调整配对",row.OldIndex<0?"pending":"ready");UpdateSummary();};Pairs.Add(p);
        }
        UpdateSummary();Message="确认新旧对应关系后，点击开始比对。";Changed("CanOpen");
    }
    void UpdateSummary(){summary=Pairs.Count==0?"支持多文件批量配对":Pairs.Count+" 组文档  ·  "+Pairs.Count(p=>p.OldIndex>=0)+" 组已配对";Changed("PairSummary");}
    void SetBusy(bool value){busy=value;Changed("CanEdit","CanNameFolder","BusyVisibility","ProgressVisibility","Progress");}
    public bool Begin() {
        if(busy)return false;grid.CommitEdit(DataGridEditingUnit.Cell,true);grid.CommitEdit(DataGridEditingUnit.Row,true);
        if(UseSubfolder){string error=OutputOptions.Validate(SubfolderName);if(error!=null){Message=error;return false;}}
        var selected=Pairs.Where(p=>p.Selected).ToArray();if(selected.Length==0){Message="请添加文档，并至少选择一组进行比对。";return false;}
        var used=new HashSet<int>();foreach(var p in selected) {
            if(p.OldIndex<0||p.OldIndex>=OldFiles.Count){Message="还有文档未配对，请选择旧版，或取消勾选该行。";return false;}
            if(!used.Add(p.OldIndex)){Message="一份旧版只能对应一份新版，请调整重复配对。";return false;}
            if(string.Equals(OldFiles[p.OldIndex].FullPath,p.NewPath,StringComparison.OrdinalIgnoreCase)){Message="新旧版不能是同一个文件。";return false;}
            if(!File.Exists(OldFiles[p.OldIndex].FullPath)||!File.Exists(p.NewPath)){Message="文件已被移动或删除，请重新添加。";return false;}
        }
        batchSubfolder=UseSubfolder;batchFolder=SubfolderName;total=selected.Length;finished=failed=0;queue.Clear();
        foreach(var p in selected){p.Output=null;p.SetStatus("等待比对","ready");queue.Enqueue(p);}
        Directory.CreateDirectory(session);SetBusy(true);Next();if(busy)timer.Start();return true;
    }
    void Next(){
        if(queue.Count==0){Finish();return;}active=queue.Dequeue();job=Path.Combine(session,Guid.NewGuid().ToString("N")+".json");
        try{
            File.WriteAllText(job,json.Serialize(new {Old=OldFiles[active.OldIndex].FullPath,New=active.NewPath,UseSubfolder=batchSubfolder,SubfolderName=batchFolder}),new UTF8Encoding(true));
            var info=new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),@"WindowsPowerShell\v1.0\powershell.exe"),"-NoProfile -STA -ExecutionPolicy Bypass -File \""+Path.Combine(folder,"CompareWorker.ps1")+"\" -JobFile \""+job+"\"");
            info.UseShellExecute=false;info.CreateNoWindow=true;info.WorkingDirectory=folder;worker=Process.Start(info);started=DateTime.Now;active.SetStatus("启动 Word","ready");
        }catch(Exception ex){Complete("启动失败："+ex.Message,null);}
    }
    Dictionary<string,object> Read(string file){return json.Deserialize<Dictionary<string,object>>(File.ReadAllText(file,Encoding.UTF8));}
    void Poll(){
        if(!busy||worker==null)return;
        try{
            if(!worker.HasExited){
                string phase="启动比对进程";DateTime at=started;
                try{var p=Read(job+".phase");phase=(string)p["Phase"];at=DateTime.Parse((string)p["At"]);}catch{}
                active.SetStatus(phase,"ready");Message="正在比对 "+(finished+1)+" / "+total+"  ·  "+active.NewName+"  ·  "+(int)(DateTime.Now-started).TotalSeconds+" 秒";
                if((DateTime.Now-at).TotalSeconds>(phase=="比较文档"?600:90)){StopOwned();Complete("超时："+phase,null);}return;
            }
            var r=Read(job+".result");Complete(Convert.ToBoolean(r["Success"])?null:Convert.ToString(r["Error"]),r.ContainsKey("Output")?Convert.ToString(r["Output"]):null);
        }catch(Exception ex){StopOwned();Complete(ex.Message,null);}
    }
    void Complete(string error,string output){
        if(worker!=null){worker.Dispose();worker=null;}active.Output=output;active.SetStatus(error==null?"比对完成":"比对失败",error==null?"success":"error",error??output);if(error!=null)failed++;finished++;Changed("CanOpen","Progress");Next();
    }
    void StopOwned(){
        if(worker!=null&&!worker.HasExited){worker.Kill();worker.WaitForExit(2000);}
        try{var o=Read(job+".owner");using(var p=Process.GetProcessById(Convert.ToInt32(o["Id"])))if(p.ProcessName.Equals("WINWORD",StringComparison.OrdinalIgnoreCase)&&p.StartTime.Ticks==Convert.ToInt64(o["Ticks"]))p.Kill();}catch{}
    }
    void Cancel(){timer.Stop();StopOwned();if(worker!=null){worker.Dispose();worker=null;}if(active!=null)active.SetStatus("已停止","pending");foreach(var p in queue)p.SetStatus("未执行","pending");queue.Clear();SetBusy(false);Message="已停止。已生成的结果仍保留在输出文件夹。";}
    void Finish(){
        timer.Stop();SetBusy(false);Message="已完成  ·  成功 "+(finished-failed)+" 组"+(failed>0?"，失败 "+failed+" 组（悬停状态查看原因）":"")+"。结果已保存"+(batchSubfolder?"到各新版目录下的「"+batchFolder+"」。":"在各新版文件旁。");
        var first=Pairs.FirstOrDefault(p=>File.Exists(p.Output));if(first!=null)grid.SelectedItem=first;Changed("CanOpen");
        if(testDirectory!=null){File.WriteAllText(Path.Combine(testDirectory,"batch-result.json"),json.Serialize(new {Total=total,Failed=failed,Outputs=Pairs.Select(p=>p.Output).ToArray()}));Program.ExitCode=failed==0?0:1;Window.Close();}
    }
    void OpenOutput(){var p=grid.SelectedItem as Pair;if(p==null||!File.Exists(p.Output))return;try{Process.Start(new ProcessStartInfo(p.Output){UseShellExecute=true});}catch(Exception ex){Message=ex.Message;}}
    public void Demo(){
        OldFiles.Add(new FileItem {FullPath=@"C:\项目资料\A轮\投资协议_定稿.docx"});OldFiles.Add(new FileItem {FullPath=@"C:\项目资料\A轮\股东协议_定稿.docx"});OldFiles.Add(new FileItem {FullPath=@"C:\项目资料\A轮\公司章程.docx"});
        NewFiles.Add(new FileItem {FullPath=@"C:\项目资料\B轮\投资协议_0920.docx"});NewFiles.Add(new FileItem {FullPath=@"C:\项目资料\B轮\股东协议_修订版.docx"});NewFiles.Add(new FileItem {FullPath=@"C:\项目资料\B轮\公司章程_新版.docx"});Refresh();UseSubfolder=true;
    }
    public void BatchTest(string directory,string mode){
        testDirectory=directory;
        AddFiles(Directory.GetFiles(Path.Combine(directory,"old"),"*.docx"),true);AddFiles(Directory.GetFiles(Path.Combine(directory,"new"),"*.docx"),false);
        if(busy||worker!=null)throw new Exception("Unexpected automatic start");
        File.WriteAllText(Path.Combine(directory,"manual-start-pass.txt"),"PASS");
        int selected=Pairs[0].OldIndex;Pairs[0].OldIndex=-1;if(Begin())throw new Exception("Unpaired row accepted");Pairs[0].OldIndex=selected;
        UseSubfolder=mode!="off";SubfolderName=mode=="default"||mode=="off"?"Legal":mode;
        if(!Begin())throw new Exception(Message);
    }
}
public static class Program {
    public static int ExitCode;
    [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr window,int attr,ref int value,int size);
    [STAThread] public static int Main(string[] args){
        try{
            if(args.Length>0&&args[0]=="/selftest"){SelfTest();return 0;}
            var app=new Application();Window window;
            using(var s=Assembly.GetExecutingAssembly().GetManifestResourceStream("MainWindow.xaml"))window=(Window)XamlReader.Load(s);
            var model=new RedlineModel();model.Wire(window);
            window.SourceInitialized+=(s,e)=>{try{int rounded=2;DwmSetWindowAttribute(new WindowInteropHelper(window).Handle,33,ref rounded,4);}catch{}};
            window.ContentRendered+=(s,e)=>{
                try{
                    if(args.Length>0&&args[0]=="/batchtest"){model.BatchTest(args[1],args.Length>2?args[2]:"off");return;}
                    if(args.Length>0&&(args[0]=="/smoketest"||args[0]=="/preview")){
                        if(args[0]=="/preview")model.Demo();
                        var timer=new DispatcherTimer {Interval=TimeSpan.FromMilliseconds(900)};
                        timer.Tick+=(a,b)=>{timer.Stop();try{if(args.Length>1)Screenshot(window,args[1]);}catch(Exception ex){File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-error.txt"),ex.ToString());ExitCode=1;}window.Close();};timer.Start();
                    }
                }catch(Exception ex){ExitCode=1;File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-error.txt"),ex.ToString());window.Close();}
            };
            app.Run(window);return ExitCode;
        }catch(Exception ex){if(args.Length>0){File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-error.txt"),ex.ToString());}else MessageBox.Show(ex.ToString(),"Word Redline 启动失败");return 1;}
    }
    static void Screenshot(Window window,string file){window.UpdateLayout();var content=(FrameworkElement)window.Content;var bmp=new RenderTargetBitmap((int)content.ActualWidth,(int)content.ActualHeight,96,96,PixelFormats.Pbgra32);var visual=new DrawingVisual();using(var dc=visual.RenderOpen()){dc.DrawRectangle(window.Background,null,new Rect(0,0,content.ActualWidth,content.ActualHeight));dc.DrawRectangle(new VisualBrush(content),null,new Rect(0,0,content.ActualWidth,content.ActualHeight));}bmp.Render(visual);var encoder=new PngBitmapEncoder();encoder.Frames.Add(BitmapFrame.Create(bmp));using(var stream=File.Create(file))encoder.Save(stream);}
    static void Assert(bool ok,string why){if(!ok)throw new Exception(why);}
    static void SelfTest(){
        Assert(MatchFiles.Suggest(new[]{"甲投资协议.docx","甲股东协议.docx","甲章程.docx"},new[]{"乙章程新版.docx","乙增资协议.docx","乙股东协议.docx"}).SequenceEqual(new[]{2,0,1}),"keyword pairing");
        Assert(MatchFiles.Suggest(new[]{"甲投资协议.docx","乙投资协议.docx"},new[]{"丙投资协议.docx"})[0]<0,"ambiguous pairing");
        Assert(MatchFiles.Suggest(new[]{"same.docx"},new[]{"same.docx"})[0]<0,"identical file");
        foreach(var name in new[]{"Legal","法务输出","Legal Review"})Assert(OutputOptions.Validate(name)==null,"Valid folder "+name);
        foreach(var name in new[]{""," ",".","..","../outside","C:\\outside","Legal/test","Legal\\test","CON","con.docx","LPT1","Legal.","Legal ","bad:name"})Assert(OutputOptions.Validate(name)!=null,"Invalid folder "+name);
        var m=new RedlineModel();Assert(!m.UseSubfolder&&m.SubfolderName=="Legal","default options");
        File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"selftest-result.txt"),"PASS: matching, 17 folder-name cases and default output options");
    }
}
