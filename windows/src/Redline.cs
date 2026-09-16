using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

static class MatchFiles
{
    public static string Kind(string path) {
        string s=Path.GetFileNameWithoutExtension(path).ToLowerInvariant();
        if (Regex.IsMatch(s,"股东协议|股東協議|shareholders?")) return "股东协议";
        if (Regex.IsMatch(s,"章程|articles.of.association|constitution")) return "章程";
        if (Regex.IsMatch(s,"股权转让|股份转让|share.purchase|equity.transfer")) return "股权转让";
        if (Regex.IsMatch(s,"投资协议|增资协议|增资认购|investment.agreement|subscription.agreement")) return "投资协议";
        return "";
    }
    public static string Clean(string path) {
        string s=Path.GetFileNameWithoutExtension(path).ToLowerInvariant();
        s=Regex.Replace(s,@"20\d{2}[-_.年]?\d{1,2}[-_.月]?\d{1,2}日?|\bv\d+(\.\d+)*|\d+","");
        s=Regex.Replace(s,"清洁版|修订版|定稿|初稿|草稿|终稿|最终版|签署版|修改版|最新版|新版|旧版|clean|final|draft|revised","");
        return Regex.Replace(s,@"[^\p{L}]+","");
    }
    public static double Score(string a,string b) {
        if (string.Equals(a,b,StringComparison.OrdinalIgnoreCase)) return -1;
        string ka=Kind(a),kb=Kind(b),x=Clean(a),y=Clean(b);
        if(ka!="" && kb!="" && ka!=kb) return -1;
        double sim=0;
        if(x.Length>0 && x==y) sim=1;
        else {
            var aa=new HashSet<string>();var bb=new HashSet<string>();
            for(int i=0;i<x.Length-1;i++) aa.Add(x.Substring(i,2));
            for(int i=0;i<y.Length-1;i++) bb.Add(y.Substring(i,2));
            if(aa.Count+bb.Count>0) sim=2.0*aa.Intersect(bb).Count()/(aa.Count+bb.Count);
        }
        return ka!="" && ka==kb ? .72+.28*sim : sim;
    }
    public static int[] Suggest(IList<string> old,IList<string> newer) {
        var answer=Enumerable.Repeat(-1,newer.Count).ToArray();
        if(old.Count==1 && newer.Count==1 && !string.Equals(old[0],newer[0],StringComparison.OrdinalIgnoreCase)) { answer[0]=0;return answer; }
        for(int n=0;n<newer.Count;n++) {
            var rank=Enumerable.Range(0,old.Count).Select(o=>new { Id=o,Value=Score(old[o],newer[n]) }).OrderByDescending(v=>v.Value).ToArray();
            if(rank.Length==0 || rank[0].Value<.58 || (rank.Length>1 && rank[0].Value-rank[1].Value<.10)) continue;
            int best=rank[0].Id;
            var reverse=Enumerable.Range(0,newer.Count).Select(j=>new {Id=j,Value=Score(old[best],newer[j])}).OrderByDescending(v=>v.Value).ToArray();
            if(reverse[0].Id==n && (reverse.Length==1 || reverse[0].Value-reverse[1].Value>=.10)) answer[n]=best;
        }
        return answer;
    }
}

class RedlineForm:Form
{
    readonly List<string> old=new List<string>(), newer=new List<string>();
    readonly ListBox left=new ListBox(),right=new ListBox();
    readonly DataGridView grid=new DataGridView();
    readonly Label status=new Label();
    readonly Button start=new Button(), cancel=new Button();
    readonly List<Control> inputs=new List<Control>();
    readonly Timer timer=new Timer();
    readonly JavaScriptSerializer json=new JavaScriptSerializer();
    readonly string folder=AppDomain.CurrentDomain.BaseDirectory;
    readonly string session=Path.Combine(Path.GetTempPath(),"WordRedline",Guid.NewGuid().ToString("N"));
    readonly Queue<int> queue=new Queue<int>();
    Process worker; string job; DateTime started; int active=-1,finished,failed,total;
    bool running; readonly string testDir;
    public RedlineForm(string test) {
        testDir=test; Text="Word Redline 2.0 · 批量文档比对"; ClientSize=new Size(1080,760); MinimumSize=new Size(1000,760);
        StartPosition=FormStartPosition.CenterScreen; Font=new Font("Microsoft YaHei UI",10); BackColor=Color.FromArgb(245,247,250);
        Directory.CreateDirectory(session);
        var layout=new TableLayoutPanel { Dock=DockStyle.Fill,Padding=new Padding(24),ColumnCount=1,RowCount=7 };
        foreach(var h in new float[]{44,32,190,36,0,56,48}) layout.RowStyles.Add(new RowStyle(h==0?SizeType.Percent:SizeType.Absolute,h==0?100:h));
        Controls.Add(layout);
        layout.Controls.Add(new Label {Text="先确认配对，再开始比对",Font=new Font(Font.FontFamily,20,FontStyle.Bold),Dock=DockStyle.Fill},0,0);
        layout.Controls.Add(new Label {Text="分别拖入一组旧版和新版 → 检查下方配对 → 点击开始比对",Dock=DockStyle.Fill,ForeColor=Color.DimGray},0,1);
        var columns=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=2};columns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50));columns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50));
        columns.Controls.Add(Side("旧版 · Original",left,old),0,0);columns.Controls.Add(Side("新版 · Revised",right,newer),1,0); layout.Controls.Add(columns,0,2);
        layout.Controls.Add(new Label {Text="配对预览：可下拉修改旧版；不需要的行取消勾选。重名结果自动编号。",Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleLeft},0,3);
        grid.ColumnHeadersHeight=36;grid.RowTemplate.Height=32;grid.Dock=DockStyle.Fill;grid.AllowUserToAddRows=false;grid.AllowUserToDeleteRows=false;grid.RowHeadersVisible=false;grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;grid.BackgroundColor=Color.White;grid.SelectionMode=DataGridViewSelectionMode.FullRowSelect;grid.MultiSelect=false;grid.EditMode=DataGridViewEditMode.EditOnEnter;
        grid.Columns.Add(new DataGridViewCheckBoxColumn {Name="Use",HeaderText="比对",FillWeight=14});
        grid.Columns.Add(new DataGridViewTextBoxColumn {Name="New",HeaderText="新版文件",ReadOnly=true,FillWeight=100});
        grid.Columns.Add(new DataGridViewComboBoxColumn {Name="Old",HeaderText="对应旧版（可调整）",FillWeight=100,FlatStyle=FlatStyle.Flat});
        grid.Columns.Add(new DataGridViewTextBoxColumn {Name="State",HeaderText="配对 / 结果",ReadOnly=true,FillWeight=60});
        grid.Columns.Add(new DataGridViewTextBoxColumn {Name="Output",HeaderText="结果路径",ReadOnly=true,Visible=false});
        grid.CurrentCellDirtyStateChanged+=(s,e)=> { if(grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
        grid.CellValueChanged+=(s,e)=> { if(!running && e.RowIndex>=0 && e.ColumnIndex==2) {grid.Rows[e.RowIndex].Cells[3].Value="手动选择，请核对";int picked=SelectedOld(grid.Rows[e.RowIndex]);grid.Rows[e.RowIndex].Cells[2].ToolTipText=picked>=0?old[picked]:"请选择旧版";grid.Rows[e.RowIndex].DefaultCellStyle.BackColor=Color.White;status.Text="配对已调整。核对后点击开始比对。";} };
        grid.DataError+=(s,e)=> {e.ThrowException=false; status.Text="请选择有效的旧版文件。";};
        grid.CellDoubleClick+=(s,e)=> { if(e.RowIndex>=0 && e.ColumnIndex>=3) OpenResult(e.RowIndex); };
        layout.Controls.Add(grid,0,4);
        status.Dock=DockStyle.Fill;status.ForeColor=Color.FromArgb(30,80,140);status.Text="拖入文件仅生成配对建议，不会自动启动 Word。";status.TextAlign=ContentAlignment.MiddleLeft;layout.Controls.Add(status,0,5);
        var actions=new FlowLayoutPanel {Dock=DockStyle.Fill};
        start.Text="开始比对";start.Size=new Size(140,38);start.BackColor=Color.FromArgb(31,91,160);start.ForeColor=Color.White;start.Click+=(s,e)=>BeginBatch(); actions.Controls.Add(start);
        cancel.Text="取消剩余比对";cancel.Size=new Size(140,38);cancel.Enabled=false;cancel.Click+=(s,e)=>CancelBatch();actions.Controls.Add(cancel);
        var open=Button("打开选中结果",()=> {if(grid.CurrentRow!=null) OpenResult(grid.CurrentRow.Index);});actions.Controls.Add(open);
        var reset=Button("清空 / 下一组",()=>{old.Clear();newer.Clear();RefreshPairs();});inputs.Add(reset);actions.Controls.Add(reset);layout.Controls.Add(actions,0,6);
        timer.Interval=400;timer.Tick+=(s,e)=>Tick();
        FormClosing+=(s,e)=> {if(running) CancelBatch();};
        Shown+=(s,e)=> {
            if(testDir!=null) {
                try {
                    AddFiles(old,Directory.GetFiles(Path.Combine(testDir,"old"),"*.docx"));AddFiles(newer,Directory.GetFiles(Path.Combine(testDir,"new"),"*.docx"));
                    if(worker!=null || running) throw new Exception("Files triggered automatic comparison");
                    File.WriteAllText(Path.Combine(testDir,"manual-start-pass.txt"),"PASS: no comparison before Start");
                    for(int r=0;r<grid.Rows.Count;r++) {
                        int o=SelectedOld(grid.Rows[r]);
                        if(o<0 || MatchFiles.Kind(old[o])!=MatchFiles.Kind(newer[r])) throw new Exception("Incorrect UI pair");
                    }
                    object saved=grid.Rows[0].Cells[2].Value;
                    grid.Rows[0].Cells[2].Value="— 请选择旧版 —";
                    BeginBatch();if(running || worker!=null) throw new Exception("Unmatched row was accepted");
                    if(grid.Rows.Count>1) {
                        grid.Rows[0].Cells[2].Value=grid.Rows[1].Cells[2].Value;
                        BeginBatch();if(running || worker!=null) throw new Exception("Duplicate old file was accepted");
                    }
                    grid.Rows[0].Cells[2].Value=saved;
                    File.WriteAllText(Path.Combine(testDir,"validation-pass.txt"),"PASS: UI suggestions, unpaired blocking, duplicate blocking, manual correction");
                    using(var bitmap=new Bitmap(Width,Height)){DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(testDir,"preview.png"));}
                    BeginBatch();
                    if(!running) {File.WriteAllText(Path.Combine(testDir,"test-error.txt"),status.Text);Close();}
                } catch(Exception ex) {File.WriteAllText(Path.Combine(testDir,"test-error.txt"),ex.ToString());Close();}
            }
        };
    }
    Button Button(string text,Action action) { var b=new Button {Text=text,Size=new Size(140,38)};b.Click+=(s,e)=>action();return b; }
    Control Side(string title,ListBox list,List<string> paths) {
        var p=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=3,Padding=new Padding(0,0,12,0)};p.RowStyles.Add(new RowStyle(SizeType.Absolute,28));p.RowStyles.Add(new RowStyle(SizeType.Percent,100));p.RowStyles.Add(new RowStyle(SizeType.Absolute,48));
        p.Controls.Add(new Label {Text=title+" · 可一次拖入多个文件",Dock=DockStyle.Fill});
        list.Dock=DockStyle.Fill;list.HorizontalScrollbar=true;list.SelectionMode=SelectionMode.MultiExtended;list.AllowDrop=true;inputs.Add(list);
        list.DragEnter+=(s,e)=> {e.Effect=!running && e.Data.GetDataPresent(DataFormats.FileDrop)?DragDropEffects.Copy:DragDropEffects.None;};
        list.DragDrop+=(s,e)=> {if(!running) AddFiles(paths,(string[])e.Data.GetData(DataFormats.FileDrop));};
        list.KeyDown+=(s,e)=>{if(!running && e.KeyCode==Keys.Delete) RemoveSelected(list,paths);};p.Controls.Add(list);
        var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,Margin=new Padding(0)};
        var add=Button("添加文件…",()=> {using(var d=new OpenFileDialog {Multiselect=true,Filter="Word 文档|*.docx;*.doc;*.docm"}) if(d.ShowDialog()==DialogResult.OK) AddFiles(paths,d.FileNames);});
        var remove=Button("移除选中",()=>RemoveSelected(list,paths));inputs.Add(add);inputs.Add(remove);buttons.Controls.Add(add);buttons.Controls.Add(remove);p.Controls.Add(buttons);return p;
    }
    void RemoveSelected(ListBox box,List<string> paths) {foreach(int i in box.SelectedIndices.Cast<int>().OrderByDescending(i=>i).ToArray()) paths.RemoveAt(i);RefreshPairs();}
    void AddFiles(List<string> list,IEnumerable<string> files) {
        int skipped=0;
        foreach(var f in files) {
            string ext=Path.GetExtension(f).ToLowerInvariant();
            if(!File.Exists(f) || !new[]{".docx",".doc",".docm"}.Contains(ext) || Path.GetFileName(f).StartsWith("~$")) {skipped++;continue;}
            string full=Path.GetFullPath(f);if(!list.Contains(full,StringComparer.OrdinalIgnoreCase)) list.Add(full);
        }
        RefreshPairs();if(skipped>0) status.Text+=" 已忽略 "+skipped+" 个非 Word 文件或临时文件。";
    }
    string Choice(int i) {return (i+1)+". "+Path.GetFileName(old[i]);}
    void RefreshPairs() {
        left.Items.Clear();left.Items.AddRange(old.Select((p,i)=>(object)((i+1)+". "+Path.GetFileName(p))).ToArray());
        right.Items.Clear();right.Items.AddRange(newer.Select(p=>(object)Path.GetFileName(p)).ToArray());
        grid.Rows.Clear();var col=(DataGridViewComboBoxColumn)grid.Columns[2];col.Items.Clear();col.Items.Add("— 请选择旧版 —");for(int i=0;i<old.Count;i++) col.Items.Add(Choice(i));
        int[] matches=MatchFiles.Suggest(old,newer);
        for(int i=0;i<newer.Count;i++) {
            int m=matches[i];int r=grid.Rows.Add(true,Path.GetFileName(newer[i]),m<0?"— 请选择旧版 —":Choice(m),m<0?"待手动配对":"建议配对 · 请核对","");
            grid.Rows[r].Cells[1].ToolTipText=newer[i];grid.Rows[r].Cells[2].ToolTipText=m<0?"候选不明确，请手动选择":old[m];
            grid.Rows[r].DefaultCellStyle.BackColor=m<0?Color.FromArgb(255,246,218):Color.White;
        }
        status.Text="旧版 "+old.Count+" 份 / 新版 "+newer.Count+" 份；已建议 "+matches.Count(m=>m>=0)+" 对。确认后点击开始比对。";
    }
    int SelectedOld(DataGridViewRow row) {string value=Convert.ToString(row.Cells[2].Value);for(int i=0;i<old.Count;i++) if(Choice(i)==value) return i;return -1;}
    void BeginBatch() {
        if(running) return;grid.EndEdit();var rows=new List<int>();var used=new HashSet<int>();
        foreach(DataGridViewRow row in grid.Rows) {
            if(!Convert.ToBoolean(row.Cells[0].Value)) continue;
            int o=SelectedOld(row),n=row.Index;
            if(o<0) {status.Text="第 "+(n+1)+" 行尚未配对，请选择旧版或取消该行勾选。";return;}
            if(!used.Add(o)) {status.Text="同一个旧版被分配给多行，请调整为一对一配对。";return;}
            if(string.Equals(old[o],newer[n],StringComparison.OrdinalIgnoreCase)) {status.Text="新旧版不能是同一个文件。";return;}
            if(!File.Exists(old[o])||!File.Exists(newer[n])) {status.Text="文件已被移动或删除，请重新添加。";return;}
            rows.Add(n);
        }
        if(rows.Count==0) {status.Text="请添加文件并至少勾选一组配对。";return;}
        queue.Clear();foreach(int r in rows) {queue.Enqueue(r);grid.Rows[r].Cells[3].Value="等待比对";grid.Rows[r].Cells[4].Value="";}
        finished=failed=0;total=rows.Count;running=true;SetBusy(true);Next();timer.Start();
    }
    void SetBusy(bool busy) {foreach(var c in inputs)c.Enabled=!busy;grid.ReadOnly=busy;start.Enabled=!busy;cancel.Enabled=busy;}
    void Next() {
        if(queue.Count==0) {Finish();return;}
        active=queue.Dequeue();var row=grid.Rows[active];job=Path.Combine(session,Guid.NewGuid().ToString("N")+".json");
        try {
            File.WriteAllText(job,json.Serialize(new {Old=old[SelectedOld(row)],New=newer[active]}),new UTF8Encoding(true));
            var info=new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),@"WindowsPowerShell\v1.0\powershell.exe"),"-NoProfile -STA -ExecutionPolicy Bypass -File \""+Path.Combine(folder,"CompareWorker.ps1")+"\" -JobFile \""+job+"\"");
            info.UseShellExecute=false;info.CreateNoWindow=true;info.WorkingDirectory=folder;worker=Process.Start(info);started=DateTime.Now;row.Cells[3].Value="正在启动 Word";
        } catch(Exception ex) {row.Cells[3].Value="启动失败："+ex.Message;failed++;finished++;worker=null;Next();}
    }
    Dictionary<string,object> Read(string path) {return json.Deserialize<Dictionary<string,object>>(File.ReadAllText(path,Encoding.UTF8));}
    void Tick() {
        if(!running || worker==null)return;
        try {
            if(!worker.HasExited) {
                string phase="启动比对进程";DateTime at=started;
                try {var p=Read(job+".phase");phase=(string)p["Phase"];at=DateTime.Parse((string)p["At"]);}catch{}
                grid.Rows[active].Cells[3].Value=phase;
                status.Text="正在处理 "+(finished+1)+" / "+total+"："+Path.GetFileName(newer[active])+" · "+phase+"（"+(int)(DateTime.Now-started).TotalSeconds+" 秒）";
                if((DateTime.Now-at).TotalSeconds>(phase=="比较文档"?600:90)) {StopOwned();Complete("超时："+phase,null);}
                return;
            }
            var result=Read(job+".result");Complete(Convert.ToBoolean(result["Success"])?null:Convert.ToString(result["Error"]),result.ContainsKey("Output")?Convert.ToString(result["Output"]):null);
        }catch(Exception ex) {StopOwned();Complete("比对失败："+ex.Message,null);}
    }
    void Complete(string error,string output) {
        if(worker!=null){worker.Dispose();worker=null;}
        var row=grid.Rows[active];row.Cells[3].Value=error==null?"完成 · 双击打开结果":error;row.Cells[3].ToolTipText=error??output;row.Cells[4].Value=output??"";
        row.DefaultCellStyle.BackColor=error==null?Color.FromArgb(230,246,237):Color.FromArgb(255,231,228);if(error!=null)failed++;finished++;Next();
    }
    void StopOwned() {
        if(worker!=null && !worker.HasExited){worker.Kill();worker.WaitForExit(2000);}
        try {var p=Read(job+".owner");var w=Process.GetProcessById(Convert.ToInt32(p["Id"]));if(w.ProcessName.Equals("WINWORD",StringComparison.OrdinalIgnoreCase) && w.StartTime.Ticks==Convert.ToInt64(p["Ticks"])) w.Kill();w.Dispose();}catch{}
    }
    void CancelBatch() {
        timer.Stop();StopOwned();if(worker!=null){worker.Dispose();worker=null;}
        if(active>=0)grid.Rows[active].Cells[3].Value="已取消（已有结果保留）";foreach(int r in queue)grid.Rows[r].Cells[3].Value="未执行";queue.Clear();running=false;SetBusy(false);status.Text="已取消剩余任务，已生成的结果文件保留。";
    }
    void Finish() {
        timer.Stop();running=false;SetBusy(false);status.Text="本组完成：成功 "+(finished-failed)+" / "+total+"，失败 "+failed+"。结果保存在各新版文件的原文件夹。";
        if(testDir!=null) {File.WriteAllText(Path.Combine(testDir,"batch-result.json"),json.Serialize(new {Total=total,Failed=failed,Outputs=grid.Rows.Cast<DataGridViewRow>().Select(r=>Convert.ToString(r.Cells[4].Value)).ToArray()}));Close();}
    }
    void OpenResult(int row) {string p=Convert.ToString(grid.Rows[row].Cells[4].Value);if(File.Exists(p)){try{Process.Start(new ProcessStartInfo(p){UseShellExecute=true});}catch(Exception ex){status.Text=ex.Message;}}else status.Text="该行还没有已生成的结果。";}
}

static class Program
{
    [STAThread] static int Main(string[] args) {
        try {
            if(args.Length>0 && args[0]=="/selftest") {SelfTest();return 0;}
            Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);
            var f=new RedlineForm(args.Length>1 && args[0]=="/batchtest"?args[1]:null);
            if(args.Length>0 && args[0]=="/smoketest") {var timer=new Timer {Interval=900};timer.Tick+=(s,e)=>{timer.Stop();if(args.Length>1){using(var b=new Bitmap(f.Width,f.Height)){f.DrawToBitmap(b,new Rectangle(Point.Empty,b.Size));b.Save(args[1]);}}f.Close();};timer.Start();}
            Application.Run(f);return 0;
        } catch(Exception ex) {MessageBox.Show(ex.ToString(),"Word Redline 错误");return 1;}
    }
    static void Assert(bool ok,string why) {if(!ok)throw new Exception(why);}
    static void SelfTest() {
        var old=new[]{"01甲公司投资协议A轮定稿20260706.docx","02甲公司股东协议.docx","03甲公司章程.docx"};
        var newer=new[]{"新公司章程修订版0914.docx","新公司增资协议2026.docx","新公司股东协议最终版.docx"};
        Assert(MatchFiles.Suggest(old,newer).SequenceEqual(new[]{2,0,1}),"keyword matching shuffled order");
        Assert(MatchFiles.Suggest(new[]{"甲投资协议.docx","乙投资协议.docx"},new[]{"丙投资协议.docx"})[0]==-1,"ambiguous pairing must remain unassigned");
        Assert(MatchFiles.Suggest(new[]{"甲旧.docx"},new[]{"乙新.docx"})[0]==0,"single pair");
        Assert(MatchFiles.Suggest(new[]{"同文件.docx"},new[]{"同文件.docx"})[0]==-1,"identical paths rejected");
        Assert(MatchFiles.Score("甲股东协议.docx","甲投资协议.docx")<0,"different document kinds");
        Assert(MatchFiles.Suggest(new[]{"甲章程.docx"},new[]{"甲章程修订版.docx","甲章程最终版.docx"}).All(i=>i<0),"duplicate candidates not reused");
        File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"selftest-result.txt"),"PASS: 6 matching tests");
    }
}
