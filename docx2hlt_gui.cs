// docx2hlt_gui - drag-and-drop window for docx2hlt.py.
//
// Drop .docx files (or folders) on the window, or on this exe's icon; they
// are converted one after another (the structured path drives the NKEditor
// GUI, so never in parallel). The .hlt goes next to each .docx, or into a
// folder chosen in the window (kept per user in %APPDATA%\docx2hlt\gui.ini).
//
// Python: <this folder>\python\python.exe when bundled, otherwise python.exe
// on PATH (the WindowsApps store stub is skipped), then the py launcher.
//
// Build (C# 5 - the .NET Framework compiler that ships with Windows):
//   C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /nologo
//     /target:winexe /codepage:65001 /out:docx2hlt_gui.exe
//     /r:System.Windows.Forms.dll /r:System.Drawing.dll docx2hlt_gui.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

static class Program
{
    [DllImport("user32.dll")]
    static extern bool SetProcessDPIAware();

    [STAThread]
    static void Main(string[] args)
    {
        try { SetProcessDPIAware(); } catch { }
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm(args));
    }
}

class Job
{
    public string Src;
    public string Dst;
    public ListViewItem Item;
    public bool Done;
}

class MainForm : Form
{
    const string NkExe = @"C:\KipoNet\NKEditor\NKEditor.exe";
    const string LayersKey =
        @"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers";
    // docx2hlt.py prints these when the .hlt was made but something needs a look
    const string ReviewFailMark = "검토 파일(.rtf/.compare.docx)을 만들지 못했습니다";
    const string ReviewNoteMark = "비교 문서에 확인할 점이 있습니다";
    const string StaleViewMark = "이미 열려 있던 NKEditor 창";

    readonly string here = AppDomain.CurrentDomain.BaseDirectory;
    readonly string iniPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "docx2hlt", "gui.ini");
    readonly string[] startArgs;

    readonly List<Job> queue = new List<Job>();
    // output path -> the source that wrote it this session, so two same-named
    // .docx from different folders do not overwrite each other in one folder
    readonly Dictionary<string, string> claimed =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    bool running;
    Process proc;
    string pythonExe;
    string pythonPrefix = "";

    ListView list;
    TextBox log;
    RadioButton rbSame, rbCustom;
    TextBox txtOut;
    Button btnBrowse, btnAdd, btnCancel, btnClear, btnOpen;
    ComboBox cmbKind;
    ToolStripStatusLabel status;

    // log line state: docx2hlt.py redraws its progress line with '\r'
    int lineStart;
    bool crPending;
    int lastStream = -1;

    public MainForm(string[] args)
    {
        startArgs = args;
        Text = "docx2hlt - DOCX를 HLT로 변환";
        Font = new Font("맑은 고딕", 9f);
        Size = new Size(960, 680);
        MinimumSize = new Size(640, 460);
        StartPosition = FormStartPosition.CenterScreen;
        BuildUi();
        LoadSettings();
        Shown += delegate { Startup(); };
        FormClosing += OnClosing;
    }

    // ---------------------------------------------------------------- UI

    void BuildUi()
    {
        var root = new TableLayoutPanel();
        root.Dock = DockStyle.Fill;
        root.ColumnCount = 1;
        root.RowCount = 3;
        root.Padding = new Padding(8);
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        Controls.Add(root);

        // row 0: hint + buttons
        var top = new FlowLayoutPanel();
        top.Dock = DockStyle.Fill;
        top.AutoSize = true;
        top.WrapContents = true;
        var hint = new Label();
        hint.Text = "변환할 .docx 파일(또는 폴더)을 이 창에 끌어다 놓으세요.";
        hint.AutoSize = true;
        hint.Font = new Font(Font.FontFamily, 11f, FontStyle.Bold);
        hint.Margin = new Padding(3, 8, 20, 3);
        top.Controls.Add(hint);
        btnAdd = MakeButton("파일 추가...", delegate { AddByDialog(); });
        btnCancel = MakeButton("대기 취소", delegate { CancelPending(); });
        btnClear = MakeButton("끝난 항목 지우기", delegate { ClearFinished(); });
        btnOpen = MakeButton("결과 폴더 열기", delegate { OpenResultFolder(); });
        top.Controls.AddRange(new Control[] { btnAdd, btnCancel, btnClear, btnOpen });
        root.Controls.Add(top, 0, 0);

        // row 1: options
        var opt = new GroupBox();
        opt.Text = "설정";
        opt.Dock = DockStyle.Fill;
        opt.AutoSize = true;
        var grid = new TableLayoutPanel();
        grid.Dock = DockStyle.Fill;
        grid.AutoSize = true;
        grid.ColumnCount = 4;
        grid.RowCount = 3;
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.Controls.Add(MakeLabel("저장 위치:"), 0, 0);
        rbSame = new RadioButton();
        rbSame.Text = "docx 파일과 같은 폴더";
        rbSame.AutoSize = true;
        rbSame.Checked = true;
        grid.Controls.Add(rbSame, 1, 0);
        grid.SetColumnSpan(rbSame, 3);
        rbCustom = new RadioButton();
        rbCustom.Text = "지정한 폴더:";
        rbCustom.AutoSize = true;
        grid.Controls.Add(rbCustom, 1, 1);
        txtOut = new TextBox();
        txtOut.ReadOnly = true;     // set only through the folder dialog
        txtOut.Dock = DockStyle.Fill;
        grid.Controls.Add(txtOut, 2, 1);
        btnBrowse = MakeButton("찾아보기...", delegate { BrowseOut(); });
        grid.Controls.Add(btnBrowse, 3, 1);
        grid.Controls.Add(MakeLabel("문서 종류:"), 0, 2);
        cmbKind = new ComboBox();
        cmbKind.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbKind.Width = 260;
        cmbKind.Items.AddRange(new object[] {
            "자동 판별 (권장)",
            "국내 특허 명세서 (spec)",
            "국문 PCT 명세서 (pct)",
            "영문 PCT 명세서 (pct_en)",
            "의견서 (opinion)",
            "일반 문서 (flat)" });
        cmbKind.SelectedIndex = 0;
        grid.Controls.Add(cmbKind, 1, 2);
        grid.SetColumnSpan(cmbKind, 2);
        opt.Controls.Add(grid);
        root.Controls.Add(opt, 0, 1);
        rbCustom.CheckedChanged += delegate { OutModeChanged(); };

        // row 2: list over log
        var split = new SplitContainer();
        split.Dock = DockStyle.Fill;
        split.Orientation = Orientation.Horizontal;
        root.Controls.Add(split, 0, 2);
        list = new ListView();
        list.Dock = DockStyle.Fill;
        list.View = View.Details;
        list.FullRowSelect = true;
        list.HideSelection = false;
        list.Columns.Add("파일", 280);
        list.Columns.Add("상태", 200);
        list.Columns.Add("결과 파일", 420);
        list.DoubleClick += delegate { RevealSelected(); };
        split.Panel1.Controls.Add(list);
        log = new TextBox();
        log.Dock = DockStyle.Fill;
        log.Multiline = true;
        log.ReadOnly = true;
        log.BackColor = SystemColors.Window;
        log.ScrollBars = ScrollBars.Both;
        log.WordWrap = false;
        log.MaxLength = 0;
        split.Panel2.Controls.Add(log);
        Load += delegate { split.SplitterDistance = split.Height * 2 / 5; };

        var strip = new StatusStrip();
        status = new ToolStripStatusLabel();
        status.Spring = true;
        status.TextAlign = ContentAlignment.MiddleLeft;
        strip.Items.Add(status);
        Controls.Add(strip);

        foreach (Control c in new Control[] { this, list, log, hint, top })
        {
            c.AllowDrop = true;
            c.DragEnter += OnDragEnter;
            c.DragDrop += OnDragDrop;
        }
    }

    Button MakeButton(string text, EventHandler click)
    {
        var b = new Button();
        b.Text = text;
        b.AutoSize = true;
        b.Click += click;
        return b;
    }

    Label MakeLabel(string text)
    {
        var l = new Label();
        l.Text = text;
        l.AutoSize = true;
        l.Anchor = AnchorStyles.Left;
        return l;
    }

    // ---------------------------------------------------------- startup

    void Startup()
    {
        if (IsElevated())
            LogLine("주의: 관리자 권한으로 실행되었습니다. 탐색기에서 끌어다 놓기가 " +
                    "막히므로 일반 권한으로 다시 실행하세요.");
        if (!File.Exists(Path.Combine(here, "docx2hlt.py")))
        {
            LogLine("오류: docx2hlt.py가 이 폴더에 없습니다 - " + here);
            SetStatus("docx2hlt.py 없음 - 폴더 전체를 받았는지 확인하세요");
            return;
        }
        if (File.Exists(NkExe))
            EnsureRunAsInvoker();
        else
            LogLine("주의: NKEditor가 없습니다 (" + NkExe + "). 특허 명세서, PCT, " +
                    "의견서를 변환하려면 NKEditor가 필요합니다.");
        if (Registry.ClassesRoot.OpenSubKey("Word.Application") == null)
            LogLine("주의: Microsoft Word가 없습니다. 수식과 표는 그림 대신 텍스트로 " +
                    "들어가고 비교 문서(.compare.docx)는 만들어지지 않습니다.");

        pythonExe = FindPython(out pythonPrefix);
        if (pythonExe == null)
        {
            LogLine("오류: Python을 찾지 못했습니다. 이 폴더에 python 폴더를 넣거나 " +
                    "Python 3을 설치하세요.");
            SetStatus("Python 없음");
            return;
        }
        string exe = pythonExe, prefix = pythonPrefix;
        new Thread(delegate () { CheckPython(exe, prefix); }) { IsBackground = true }.Start();

        if (startArgs.Length > 0)
            AddPaths(startArgs);
        else
            SetStatus("대기 중 - 파일을 끌어다 놓으세요");
    }

    static bool IsElevated()
    {
        try
        {
            return new WindowsPrincipal(WindowsIdentity.GetCurrent())
                .IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch { return false; }
    }

    // NKEditor's manifest asks for admin; without this the automated launch
    // fails with error 740. Per user (HKCU), no admin needed.
    void EnsureRunAsInvoker()
    {
        try
        {
            using (var k = Registry.CurrentUser.CreateSubKey(LayersKey))
            {
                var cur = k.GetValue(NkExe) as string;
                if (cur != null && cur.IndexOf("RUNASINVOKER",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return;
                var val = string.IsNullOrEmpty(cur) ? "~ RUNASINVOKER"
                                                    : cur.TrimEnd() + " RUNASINVOKER";
                k.SetValue(NkExe, val);
                LogLine("NKEditor를 관리자 권한 없이 실행하도록 설정했습니다 (HKCU AppCompatFlags).");
            }
        }
        catch (Exception e)
        {
            LogLine("주의: NKEditor 실행 권한 설정 실패 - " + e.Message);
        }
    }

    string FindPython(out string prefix)
    {
        prefix = "";
        var bundled = Path.Combine(here, "python", "python.exe");
        if (File.Exists(bundled))
            return bundled;
        var path = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var raw in path.Split(';'))
        {
            var dir = raw.Trim().Trim('"');
            if (dir.Length == 0 ||
                dir.IndexOf("WindowsApps", StringComparison.OrdinalIgnoreCase) >= 0)
                continue;
            try
            {
                var p = Path.Combine(dir, "python.exe");
                if (File.Exists(p))
                    return p;
            }
            catch { }
        }
        try
        {
            var progs = Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData), "Programs", "Python");
            if (Directory.Exists(progs))
            {
                var best = Directory.GetDirectories(progs, "Python3*")
                    .Select(d => Path.Combine(d, "python.exe"))
                    .Where(File.Exists)
                    .OrderByDescending(p => p, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault();
                if (best != null)
                    return best;
            }
        }
        catch { }
        var py = Path.Combine(Environment.GetFolderPath(
            Environment.SpecialFolder.Windows), "py.exe");
        if (File.Exists(py))
        {
            prefix = "-3 ";
            return py;
        }
        return null;
    }

    void CheckPython(string exe, string prefix)
    {
        string msg;
        try
        {
            var psi = NewPsi(exe, prefix + "-c \"import sys; v = sys.version.split()[0]; " +
                "import PIL; print(v, 'Pillow', PIL.__version__)\"");
            using (var p = Process.Start(psi))
            {
                p.StandardInput.Close();
                var outText = p.StandardOutput.ReadToEndAsync();
                var errText = p.StandardError.ReadToEndAsync();
                if (!p.WaitForExit(30000))
                {
                    try { p.Kill(); } catch { }
                    msg = "주의: Python 확인이 30초 안에 끝나지 않았습니다 - " + exe;
                }
                else if (p.ExitCode == 0)
                    msg = "Python: " + exe + " (" + outText.Result.Trim() + ")";
                else if (errText.Result.Contains("PIL"))
                    msg = "오류: Pillow가 없습니다. 다음 명령으로 설치하세요: \"" + exe +
                          "\" " + prefix + "-m pip install Pillow";
                else
                    msg = "오류: Python 실행 실패 - " + exe + " - " + errText.Result.Trim();
            }
        }
        catch (Exception e)
        {
            msg = "오류: Python 실행 실패 - " + exe + " - " + e.Message;
        }
        UI(delegate { LogLine(msg); });
    }

    // ------------------------------------------------------------ input

    void OnDragEnter(object sender, DragEventArgs e)
    {
        e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy : DragDropEffects.None;
    }

    void OnDragDrop(object sender, DragEventArgs e)
    {
        var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
        if (paths != null)
            // the drop source (Explorer) waits until this handler returns
            BeginInvoke(new Action(delegate { AddPaths(paths); }));
    }

    void AddByDialog()
    {
        using (var d = new OpenFileDialog())
        {
            d.Filter = "Word 문서 (*.docx)|*.docx";
            d.Multiselect = true;
            d.Title = "변환할 .docx 파일 선택";
            if (d.ShowDialog(this) == DialogResult.OK)
                AddPaths(d.FileNames);
        }
    }

    void AddPaths(IEnumerable<string> paths)
    {
        int added = 0;
        foreach (var p in paths)
        {
            if (Directory.Exists(p))
            {
                // top level only: a case folder's subfolders hold old rounds
                foreach (var f in Directory.GetFiles(p, "*.docx").OrderBy(x => x))
                    if (Enqueue(f)) added++;
            }
            else if (File.Exists(p))
            {
                if (Enqueue(p)) added++;
            }
        }
        if (added == 0)
            SetStatus("추가된 .docx 파일이 없습니다");
        StartNext();
    }

    bool Enqueue(string path)
    {
        var name = Path.GetFileName(path);
        if (!path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            LogLine("건너뜀 (.docx가 아님): " + name);
            return false;
        }
        if (name.StartsWith("~$"))
            return false;   // Word's lock file of an open document
        var full = Path.GetFullPath(path);
        if (queue.Any(j => !j.Done &&
                string.Equals(j.Src, full, StringComparison.OrdinalIgnoreCase)))
        {
            LogLine("이미 목록에 있음: " + name);
            return false;
        }
        var job = new Job();
        job.Src = full;
        job.Item = new ListViewItem(new[] { name, "대기", "" });
        job.Item.ToolTipText = full;
        job.Item.Tag = job;
        queue.Add(job);
        list.Items.Add(job.Item);
        return true;
    }

    // ------------------------------------------------------- conversion

    void StartNext()
    {
        if (running)
            return;
        var job = queue.FirstOrDefault(j => !j.Done);
        if (job == null)
        {
            SetStatus("대기 중 - 파일을 끌어다 놓으세요");
            return;
        }
        if (pythonExe == null)
        {
            SetStatus("Python이 없어 변환할 수 없습니다");
            return;
        }
        if (rbCustom.Checked && txtOut.Text.Trim().Length == 0)
        {
            SetStatus("저장 폴더를 지정하면 변환을 시작합니다");
            return;
        }
        try
        {
            job.Dst = ComputeDst(job.Src);
        }
        catch (Exception e)
        {
            Finish(job, "실패 (저장 위치 오류)", Color.Firebrick, e.Message);
            StartNext();
            return;
        }
        job.Item.SubItems[2].Text = job.Dst;
        SetItem(job, "변환 중...", Color.RoyalBlue);
        running = true;
        int done = queue.Count(j => j.Done), total = queue.Count;
        SetStatus(string.Format("변환 중 ({0}/{1}): {2}", done + 1, total,
                                Path.GetFileName(job.Src)));
        var kind = KindArg();
        new Thread(delegate () { RunJob(job, kind); }) { IsBackground = true }.Start();
    }

    string ComputeDst(string src)
    {
        string dir;
        if (rbCustom.Checked)
        {
            dir = Path.GetFullPath(txtOut.Text.Trim());
            Directory.CreateDirectory(dir);
        }
        else
            dir = Path.GetDirectoryName(src);
        var stem = Path.GetFileNameWithoutExtension(src);
        var dst = Path.Combine(dir, stem + ".hlt");
        string owner;
        for (int n = 2; claimed.TryGetValue(dst, out owner) &&
                 !string.Equals(owner, src, StringComparison.OrdinalIgnoreCase); n++)
            dst = Path.Combine(dir, stem + " (" + n + ").hlt");
        claimed[dst] = src;
        return dst;
    }

    string KindArg()
    {
        switch (cmbKind.SelectedIndex)
        {
            case 1: return "spec";
            case 2: return "pct";
            case 3: return "pct_en";
            case 4: return "opinion";
            case 5: return "flat";
            default: return null;
        }
    }

    ProcessStartInfo NewPsi(string exe, string args)
    {
        var psi = new ProcessStartInfo(exe, args);
        psi.UseShellExecute = false;
        // a hidden console: EqRender/TerConv inherit it instead of each
        // popping up a console window of their own
        psi.CreateNoWindow = true;
        psi.RedirectStandardInput = true;
        psi.RedirectStandardOutput = true;
        psi.RedirectStandardError = true;
        psi.StandardOutputEncoding = new UTF8Encoding(false);
        psi.StandardErrorEncoding = new UTF8Encoding(false);
        psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
        psi.EnvironmentVariables["PYTHONUNBUFFERED"] = "1";
        psi.WorkingDirectory = here;
        return psi;
    }

    static string Q(string s) { return "\"" + s + "\""; }

    void RunJob(Job job, string kind)
    {
        var started = DateTime.Now;
        var output = new StringBuilder();
        int code = -1;
        string error = null;
        UI(delegate {
            NewLineIfNeeded();
            LogLine("=== " + Path.GetFileName(job.Src) + "  ->  " + job.Dst);
        });
        try
        {
            // the 2-argument form names the output; no --pause-on-review-error,
            // which would wait for Enter on a console nobody sees
            var args = pythonPrefix + Q(Path.Combine(here, "docx2hlt.py")) + " " +
                       Q(job.Src) + " " + Q(job.Dst) +
                       (kind != null ? " --kind=" + kind : "");
            using (var p = Process.Start(NewPsi(pythonExe, args)))
            {
                proc = p;
                p.StandardInput.Close();
                var tOut = new Thread(delegate () { Pump(p.StandardOutput, output, 0); });
                var tErr = new Thread(delegate () { Pump(p.StandardError, output, 1); });
                tOut.IsBackground = tErr.IsBackground = true;
                tOut.Start();
                tErr.Start();
                p.WaitForExit();
                tOut.Join();
                tErr.Join();
                code = p.ExitCode;
            }
        }
        catch (Exception e)
        {
            error = e.Message;
        }
        proc = null;
        string text;
        lock (output) text = output.ToString();
        UI(delegate { JobEnded(job, started, code, error, text); });
    }

    void JobEnded(Job job, DateTime started, int code, string error, string text)
    {
        running = false;
        var secs = (int)(DateTime.Now - started).TotalSeconds;
        var took = secs >= 60 ? (secs / 60) + "분 " + (secs % 60) + "초" : secs + "초";
        bool made = false;
        try
        {
            made = File.Exists(job.Dst) &&
                   File.GetLastWriteTime(job.Dst) >= started.AddSeconds(-2);
        }
        catch { }
        if (error != null)
            Finish(job, "실패 (실행 오류)", Color.Firebrick, error);
        else if (code == 0 && made)
        {
            var state = "완료";
            var color = Color.ForestGreen;
            if (text.Contains(ReviewFailMark))
            {
                state = "완료 (검토 파일 실패)";
                color = Color.DarkOrange;
            }
            else if (text.Contains(ReviewNoteMark))
            {
                state = "완료 (비교 문서 확인 필요)";
                color = Color.DarkOrange;
            }
            if (text.Contains(StaleViewMark))
            {
                state += " - 열린 NKEditor 창 주의";
                color = Color.DarkOrange;
            }
            Finish(job, state + ", " + took, color, null);
        }
        else
            Finish(job, "실패 (코드 " + code + "), " + took, Color.Firebrick,
                   "로그를 확인하세요");
        StartNext();
    }

    void Finish(Job job, string state, Color color, string note)
    {
        job.Done = true;
        SetItem(job, state, color);
        NewLineIfNeeded();
        LogLine("--- " + Path.GetFileName(job.Src) + ": " + state +
                (note != null ? " - " + note : ""));
    }

    // Reads a child stream in chunks (not lines) so the '\r' progress line
    // shows live, and hands the pieces to the UI thread in order.
    void Pump(StreamReader r, StringBuilder all, int stream)
    {
        var buf = new char[4096];
        int n;
        try
        {
            while ((n = r.Read(buf, 0, buf.Length)) > 0)
            {
                lock (all) all.Append(buf, 0, n);
                var ops = new List<string>();   // "\r", "\n" or a text piece
                var cur = new StringBuilder();
                for (int i = 0; i < n; i++)
                {
                    char c = buf[i];
                    if (c == '\r' || c == '\n')
                    {
                        if (cur.Length > 0) { ops.Add(cur.ToString()); cur.Clear(); }
                        ops.Add(c.ToString());
                    }
                    else
                        cur.Append(c);
                }
                if (cur.Length > 0)
                    ops.Add(cur.ToString());
                UI(delegate { ApplyOps(ops, stream); });
            }
        }
        catch { }
    }

    void ApplyOps(List<string> ops, int stream)
    {
        if (stream != lastStream)
        {
            NewLineIfNeeded();
            lastStream = stream;
        }
        foreach (var op in ops)
        {
            if (op == "\n")
                NewLine();
            else if (op == "\r")
                crPending = true;     // "\r\n" ends the line; "\r" + text redraws it
            else
            {
                if (crPending && log.TextLength > lineStart)
                {
                    log.Select(lineStart, log.TextLength - lineStart);
                    log.SelectedText = "";
                }
                crPending = false;
                log.AppendText(op);
            }
        }
    }

    void NewLine()
    {
        log.AppendText(Environment.NewLine);
        lineStart = log.TextLength;
        crPending = false;
    }

    void NewLineIfNeeded()
    {
        if (log.TextLength > lineStart)
            NewLine();
        crPending = false;
    }

    void LogLine(string s)
    {
        NewLineIfNeeded();
        log.AppendText(s);
        NewLine();
        lastStream = -1;
    }

    // ---------------------------------------------------------- actions

    void CancelPending()
    {
        foreach (var j in queue.Where(j => !j.Done && j.Item.SubItems[1].Text == "대기"))
        {
            j.Done = true;
            SetItem(j, "취소됨", Color.Gray);
        }
        if (!running)
            SetStatus("대기 중 - 파일을 끌어다 놓으세요");
    }

    void ClearFinished()
    {
        foreach (var j in queue.Where(j => j.Done).ToList())
        {
            list.Items.Remove(j.Item);
            queue.Remove(j);
        }
    }

    void OpenResultFolder()
    {
        var job = SelectedJob();
        if (job != null)
        {
            RevealSelected();
            return;
        }
        if (rbCustom.Checked && Directory.Exists(txtOut.Text))
            Process.Start("explorer.exe", Q(txtOut.Text));
        else
            MessageBox.Show(this, "목록에서 파일을 선택하세요.", Text);
    }

    Job SelectedJob()
    {
        return list.SelectedItems.Count > 0 ? list.SelectedItems[0].Tag as Job : null;
    }

    void RevealSelected()
    {
        var job = SelectedJob();
        if (job == null)
            return;
        if (job.Dst != null && File.Exists(job.Dst))
            Process.Start("explorer.exe", "/select," + Q(job.Dst));
        else if (File.Exists(job.Src))
            Process.Start("explorer.exe", "/select," + Q(job.Src));
    }

    bool PickOutFolder()
    {
        using (var d = new FolderBrowserDialog())
        {
            d.Description = "변환된 .hlt 파일을 저장할 폴더";
            if (Directory.Exists(txtOut.Text))
                d.SelectedPath = txtOut.Text;
            if (d.ShowDialog(this) != DialogResult.OK)
                return false;
            txtOut.Text = d.SelectedPath;
            return true;
        }
    }

    void BrowseOut()
    {
        if (!PickOutFolder())
            return;
        rbCustom.Checked = true;
        SaveSettings();
        StartNext();
    }

    bool inModeChange;

    // wired to rbCustom only: it fires on both check and uncheck
    void OutModeChanged()
    {
        if (inModeChange)
            return;
        inModeChange = true;
        try
        {
            if (rbCustom.Checked && txtOut.Text.Length == 0 && Visible &&
                    !PickOutFolder())
                rbSame.Checked = true;
            txtOut.Enabled = btnBrowse.Enabled = rbCustom.Checked;
            SaveSettings();
        }
        finally
        {
            inModeChange = false;
        }
        StartNext();
    }

    void OnClosing(object sender, FormClosingEventArgs e)
    {
        var p = proc;
        if (running && p != null)
        {
            var r = MessageBox.Show(this,
                "변환 중입니다. 중단하고 닫을까요?\n(중단하면 이번 파일은 변환되지 않습니다)",
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (r != DialogResult.Yes)
            {
                e.Cancel = true;
                return;
            }
            try
            {
                // /T also ends the NKEditor/Word helpers this run started
                var k = new ProcessStartInfo("taskkill.exe",
                    "/PID " + p.Id + " /T /F");
                k.CreateNoWindow = true;
                k.UseShellExecute = false;
                Process.Start(k).WaitForExit(10000);
            }
            catch { }
        }
        SaveSettings();
    }

    // --------------------------------------------------------- settings

    void LoadSettings()
    {
        try
        {
            if (!File.Exists(iniPath))
                return;
            string mode = null, dir = null;
            foreach (var line in File.ReadAllLines(iniPath, Encoding.UTF8))
            {
                var i = line.IndexOf('=');
                if (i <= 0) continue;
                var key = line.Substring(0, i).Trim();
                var val = line.Substring(i + 1).Trim();
                if (key == "out_mode") mode = val;
                else if (key == "out_dir") dir = val;
            }
            if (!string.IsNullOrEmpty(dir))
                txtOut.Text = dir;
            if (mode == "custom" && txtOut.Text.Length > 0)
                rbCustom.Checked = true;
        }
        catch { }
        txtOut.Enabled = btnBrowse.Enabled = rbCustom.Checked;
    }

    void SaveSettings()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(iniPath));
            File.WriteAllLines(iniPath, new[] {
                "out_mode=" + (rbCustom.Checked ? "custom" : "same"),
                "out_dir=" + txtOut.Text }, Encoding.UTF8);
        }
        catch { }
    }

    // ----------------------------------------------------------- helpers

    void SetItem(Job job, string state, Color color)
    {
        job.Item.SubItems[1].Text = state;
        job.Item.ForeColor = color;
    }

    void SetStatus(string s) { status.Text = s; }

    void UI(Action a)
    {
        try
        {
            if (IsHandleCreated && !IsDisposed)
                BeginInvoke(a);
        }
        catch (InvalidOperationException) { }
    }
}
