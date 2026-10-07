// TerConv.exe — 32-bit helper that drives hwplt32.dll (Sub Systems TE Edit Control)
// to build an NKEditor .hlt document from a binary job file.
//
// Usage: TerConv.exe <jobfile> <outfile.hlt> [NKEditorDir]
//
// Job file format (little-endian):
//   u32 imgDirLen; u8 imgDir[imgDirLen]   // dir holding HLT%N.HIF staging files (utf8)
//   u32 lineCount
//   per line:
//     u8 kind                            // 0 = text line, 1 = image
//     kind 0: u32 textLen; u8 text[textLen]   // CP949 bytes
//             u8  terminated                  // 1 = paragraph end (0x15), 0 = soft break
//             u32 fmtCount
//             per fmt: u32 colStart, u32 colEnd, u32 styleMask   // byte offsets in line
//                        i32 sizeX2                            // font size in half-points, -1 = unset
//                        u32 colorBGR                          // 0x00BBGGRR, 0xFFFFFFFF = unset
//     both kinds end with:
//             u8  hasPara                       // 1 = para spec follows
//             i32 leftInd, i32 rightInd, i32 firstInd  // twips; first<0 = hanging
//             u32 pflags                        // para flags (alignment bits)
//             i32 spaceBef, i32 spaceAft, i32 spaceBet // twips
//     kind 1: u32 pathLen; u8 path[pathLen]   // utf8 source image path
//             i32 dispW_pt, i32 dispH_pt      // display size in points
//             i32 pxW, i32 pxH                // pixel size
//
// styleMask: 1=bold 2=italic 4=underline 8=strike 0x10=sub 0x20=super
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Linq;
using System.Collections.Generic;

public static class TerConv {
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool LoadTerControl();
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool SaveTerFile(IntPtr h, [MarshalAs(UnmanagedType.LPStr)] string f);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool TerInsertLine(IntPtr h, byte[] text, int fontId, int paraId, int cellId, int reserved, int repaint);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool SelectTerText(IntPtr h, int l1, int c1, int l2, int c2, bool repaint);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool SetTerCharStyle(IntPtr h, int styles, int onOff, bool repaint);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool DeselectTerText(IntPtr h, bool repaint);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int TerInsertPictureFile(IntPtr h, [MarshalAs(UnmanagedType.LPStr)] string f, int embed, int align, int insert);
    // TerSetPictSize(h, objectId, widthPt, heightPt): stores the image's display
    // size in points, converting to internal units itself — the only reliable way
    // to size an image (the raw record fields are easy to get wrong).
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool TerSetPictSize(IntPtr h, int objectId, int w, int ht);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int TerSetImagePath(IntPtr h, [MarshalAs(UnmanagedType.LPStr)] string p);
    // Anchors a created object inline in the text at the cursor: writes a 0x5f
    // marker char and records the object id in the parallel reference array.
    // Without it the object floats and piles up at the end in reverse order.
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int TerInsertObjectId(IntPtr h, int objectId, int repaint);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int GetTerLine(IntPtr h, int line, byte[] text, int[] objids);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int ReadTerFile(IntPtr h, [MarshalAs(UnmanagedType.LPStr)] string f);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int SetTerCursorPos(IntPtr h, int line, int col, int flag);
    // Paragraph format ids: creates/dedups a 0x40-byte para record. Fields:
    //   left/right/first indents in twips (first<0 = hanging), then tabId,
    //   cellId, auxId, shading, frameId, spaceBefore/After/Between (twips),
    //   flags (alignment: 0x800 default-left; mutually exclusive group 1,2,0x400,0x800)
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int TerCreateParaId(IntPtr h, int reuseId, int shared,
        int left, int right, int first, int tabId, int cellId, int auxId,
        int shading, int frameId, int spaceBefore, int spaceAfter, int spaceBetween, int flags);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool SetTerPointSize(IntPtr h, int size, bool repaint);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool SetTerColor(IntPtr h, int color, bool repaint);
    // 문서종류 — integer code stored at doc+0x1a37d, read/written by these.
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int TerGetDocType(IntPtr h);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern bool TerSetDocType(IntPtr h, int docType);
    [DllImport("hwplt32.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int TerImportFile(IntPtr h, [MarshalAs(UnmanagedType.LPStr)] string f, int a, int b);

    // ImageGear (igcore13d.dll) licensing — same sequence NKEditor.exe runs at
    // startup. Without it IG_info_get (called inside TerInsertPictureFile) fails,
    // the image object's pixel handle stays null, and the saved .hlt embeds the
    // raw staged file which NKEditor then renders as a broken (X) box.
    [DllImport("igcore13d.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int IG_lic_solution_name_set([MarshalAs(UnmanagedType.LPStr)] string s);
    [DllImport("igcore13d.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int IG_lic_solution_key_set(uint a, uint b, uint c, uint d);
    [DllImport("igcore13d.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int IG_lic_OEM_license_key_set([MarshalAs(UnmanagedType.LPStr)] string s);
    [DllImport("igcore13d.dll", CallingConvention = CallingConvention.StdCall)]
    static extern int IG_comm_comp_attach([MarshalAs(UnmanagedType.LPStr)] string s);

    [DllImport("ole32.dll")] static extern int OleInitialize(IntPtr r);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern bool SetDllDirectoryW(string p);
    [DllImport("kernel32.dll", CharSet = CharSet.Ansi)] static extern IntPtr LoadLibraryA(string p);
    [DllImport("kernel32.dll", CharSet = CharSet.Ansi)] static extern IntPtr GetModuleHandleA(string m);
    [DllImport("user32.dll", CharSet = CharSet.Ansi)]
    static extern IntPtr CreateWindowExA(uint ex, string cls, string name, uint style,
        int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr inst, IntPtr param);
    [DllImport("user32.dll")] static extern bool GetMessageA(out MSG m, IntPtr h, uint mn, uint mx);
    [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG m);
    [DllImport("user32.dll")] static extern IntPtr DispatchMessageA(ref MSG m);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Ansi)] static extern int GetClassNameA(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Ansi)] static extern int GetWindowTextA(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern IntPtr SendMessageA(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern int GetWindowLongA(IntPtr h, int idx);

    struct MSG { public IntPtr hwnd; public uint message; public IntPtr w; public IntPtr l; public uint time; public int x, y; }
    class ImgRec { public string src; public int pw, ph, dw, dh, line; }
    // One parsed output line from the job stream, before placement.
    class LineRec {
        public int sec;                 // section index (element slot it fills)
        public byte kind;               // 0 text, 1 image
        public byte[] text; public bool term;
        public List<int[]> fmts = new List<int[]>();   // s,e,mask,sizeX2,colorBGR
        public int paraId;
        public string src; public int dw, dh, pw, ph;  // image
        public int land = -1;           // final document line index after insert
    }

    static string nkDir = @"C:\KipoNet\NKEditor";
    static volatile bool done = false;
    static volatile int stage = 0;
    static volatile string error = null;
    static byte[] job; static string outFile; static string dumpPath; static string imgTest;
    static string rtIn, rtOut; static string fontTest; static string dtPath; static int dtSet = -999;
    static string impPath; static int impA = 0, impB = 0;
    static int rtLine = -1; static string rtText = "PROBE-INSERTED";

    static void Log(string s) { Console.WriteLine(s); Console.Out.Flush(); }

    static void ClickDialogs() {
        EnumWindows(delegate (IntPtr h, IntPtr l) {
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (pid != System.Diagnostics.Process.GetCurrentProcess().Id || !IsWindowVisible(h)) return true;
            var c = new StringBuilder(64); GetClassNameA(h, c, 64);
            if (c.ToString() != "#32770") return true;
            EnumChildWindows(h, delegate (IntPtr ch, IntPtr l2) {
                var cc = new StringBuilder(64); GetClassNameA(ch, cc, 64);
                if (cc.ToString() != "Button") return true;
                var tt = new StringBuilder(64); GetWindowTextA(ch, tt, 64);
                string s = tt.ToString();
                if (s.Contains("확인") || s.Contains("OK") || s.Contains("Yes") || s.Contains("예"))
                    SendMessageA(ch, 0x00F5, IntPtr.Zero, IntPtr.Zero);   // BM_CLICK
                return true;
            }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
    }

    static void Watchdog() { while (!done) { ClickDialogs(); Thread.Sleep(200); } }

    static uint ReadU32(byte[] b, ref int p) { uint v = (uint)(b[p] | b[p + 1] << 8 | b[p + 2] << 16 | b[p + 3] << 24); p += 4; return v; }

    static string ReadCStr(IntPtr basePtr, int off) {
        var sb = new StringBuilder();
        for (int i = 0; i < 512; i++) { byte c = Marshal.ReadByte(basePtr, off + i); if (c == 0) break; sb.Append((char)c); }
        return sb.ToString();
    }
    // basename, lowercased — unique within one docx extraction (image1.png, ...)
    static string Base(string s) {
        s = s.Replace('/', '\\').Trim().ToLowerInvariant();
        int i = s.LastIndexOf('\\'); return i >= 0 ? s.Substring(i + 1) : s;
    }

    static void InitImageGear() {
        try {
            int r1 = IG_lic_solution_name_set("Korea Intellectual Property");
            int r2 = IG_lic_solution_key_set(0x7bfcd7bf, 0xcd72ab7d, 0xceff2dc2, 0xa2b242b2);
            int r3 = IG_lic_OEM_license_key_set("1.0.R12N5txzGgak21KuxoCtaoGsIuEFMFGt6uCFxgnY6dxTSzCY2T2dDsGd23M1x1CkXFDdGF2s5NxsGoDuXkSzx");
            int a1 = IG_comm_comp_attach("JPEG2K");
            int a2 = IG_comm_comp_attach("LZW");
            Log("ig lic=" + r1 + "," + r2 + "," + r3 + " attach JPEG2K=" + a1 + " LZW=" + a2);
        } catch (Exception ex) { Log("warn: ig init " + ex.Message); }
    }

    // Reads the per-line para spec from the job stream; creates (or reuses via
    // cache) a paraId with TerCreateParaId and returns it, or -1 for default.
    static int ReadParaSpec(byte[] job, ref int p, IntPtr ter, IntPtr doc, Dictionary<string, int> cache) {
        byte has = job[p++];
        // 0 = default para. Must return 0 (not -1): a new line inherits the
        // previous line's paraId when -1 is passed, so plain paragraphs would
        // pick up the indent/alignment of whatever preceded them.
        if (has == 0) return 0;
        int left = (int)ReadU32(job, ref p), right = (int)ReadU32(job, ref p), first = (int)ReadU32(job, ref p);
        uint pfl = ReadU32(job, ref p);
        int spB = (int)ReadU32(job, ref p), spA = (int)ReadU32(job, ref p), spBt = (int)ReadU32(job, ref p);
        string key = left + "," + right + "," + first + "," + pfl + "," + spB + "," + spA + "," + spBt;
        int id;
        if (cache.TryGetValue(key, out id)) return id;
        // unused id fields are 0 (matching para0's defaults), not -1
        id = TerCreateParaId(ter, -1, 0, left, right, first, 0, 0, 0, 0, 0, spB, spA, spBt, (int)pfl);
        if (id < 0) { Log("warn: paraId create failed"); return -1; }
        // TerCreateParaId stores -1 into several id slots (tab/cell/aux/frame) when
        // given 0, but the default para0 keeps them 0. A -1 in the style/font slot
        // makes the renderer draw nothing for that line, so reset every field that
        // is not one we explicitly set back to 0.
        IntPtr pt0 = Marshal.ReadIntPtr(doc, 0x9324);
        if (pt0 != IntPtr.Zero) {
            IntPtr pr = IntPtr.Add(pt0, id * 0x40);
            int[] keep = { 0x0c, 0x10, 0x14, 0x18, 0x30, 0x34, 0x38 };
            for (int off = 0; off < 0x40; off += 4) {
                if (Array.IndexOf(keep, off) < 0) Marshal.WriteInt32(pr, off, 0);
            }
        }
        cache[key] = id;
        return id;
    }

    // Read the loaded document's per-line text (cp949, trimmed) so we can find
    // the skeleton's 【element】 markers and where each content region starts.
    static List<string> LineTexts(IntPtr doc) {
        var res = new List<string>();
        try {
            IntPtr la = Marshal.ReadIntPtr(doc, 0x170); int nl = Marshal.ReadInt32(doc, 0x52d);
            for (int i = 0; i < nl; i++) {
                IntPtr ls = Marshal.ReadIntPtr(la, i * 4);
                if (ls == IntPtr.Zero) { res.Add(""); continue; }
                IntPtr tp = Marshal.ReadIntPtr(ls, 0); int cl = Marshal.ReadInt32(ls, 0xa);
                if (tp == IntPtr.Zero || cl <= 0) { res.Add(""); continue; }
                byte[] bb = new byte[Math.Min(cl, 200)]; Marshal.Copy(tp, bb, 0, bb.Length);
                res.Add(Encoding.GetEncoding(949).GetString(bb).Replace("\x15", "").Replace("\r", "").Trim());
            }
        } catch { }
        return res;
    }

    static void DumpDoc(IntPtr doc) {
        try {
            IntPtr la = Marshal.ReadIntPtr(doc, 0x170); int nl = Marshal.ReadInt32(doc, 0x52d);
            IntPtr pt = Marshal.ReadIntPtr(doc, 0x9324); int np = Marshal.ReadInt32(doc, 0x9328);
            Log("dump lines=" + nl + " paras=" + np);
            for (int li = 0; li < nl; li++) {
                IntPtr ls = Marshal.ReadIntPtr(la, li * 4); if (ls == IntPtr.Zero) continue;
                IntPtr tp = Marshal.ReadIntPtr(ls, 0); int cl = Marshal.ReadInt32(ls, 0xa);
                int pid = Marshal.ReadInt32(ls, 0x12);
                if (tp == IntPtr.Zero || cl <= 0 || cl > 8000) continue;
                byte[] bb = new byte[cl]; Marshal.Copy(tp, bb, 0, cl);
                int nm = 0; foreach (byte c in bb) if (c == 0x5f) nm++;
                string txt = Encoding.GetEncoding(949).GetString(bb).Replace("\x5f", "[OBJ]").Replace("\r", " ").Trim();
                if (txt.Length > 40) txt = txt.Substring(0, 40);
                var pf = new StringBuilder();
                if (pt != IntPtr.Zero && pid >= 0 && pid < np) {
                    IntPtr pr = IntPtr.Add(pt, pid * 0x40);
                    for (int off = 0; off < 0x40; off += 4) {
                        int v = Marshal.ReadInt32(pr, off);
                        if (v != 0) pf.Append(" +").Append(off.ToString("x2")).Append("=").Append(v);
                    }
                }
                if (nm > 0 || txt.Length > 0 || pf.Length > 0)
                    Log("L" + li + (nm > 0 ? " OBJx" + nm : "") + " pid=" + pid + pf + " | " + txt);
            }
            IntPtr fa = Marshal.ReadIntPtr(doc, 0x93bd); int fc = Marshal.ReadInt32(doc, 0x9481);
            Log("fonts=" + fc);
            for (int i = 0; i < fc && i < 20; i++) {
                IntPtr fp = IntPtr.Add(fa, i * 0xc0);
                int sz = Marshal.ReadInt32(fp, 0x68), st = Marshal.ReadInt32(fp, 0x6c);
                string nm = ReadCStr(fp, 0);
                Log("  font" + i + " sz=" + sz + " style=0x" + st.ToString("x") + " " + nm);
            }
            IntPtr oa = Marshal.ReadIntPtr(doc, 0x9366); int oc = Marshal.ReadInt32(doc, 0x936e);
            Log("objects=" + oc);
            for (int i = 0; i < oc && i < 60; i++) {
                IntPtr ob = new IntPtr(oa.ToInt64() + (long)i * 0x1ce);
                int type = Marshal.ReadInt32(ob, 0x56); int h5e = Marshal.ReadInt32(ob, 0x5e);
                int f1b6 = Marshal.ReadInt32(ob, 0x1b6); string p = ReadCStr(ob, 0x96);
                Log("  obj" + i + " type=" + type + " 5e=" + h5e + " act=" + f1b6 + " " + Base(p));
                var fs = new StringBuilder();
                for (int off = 0; off < 0x1ce; off += 4) {
                    int v = Marshal.ReadInt32(ob, off);
                    if (v != 0) fs.Append(" ").Append(off.ToString("x2")).Append("=").Append(v.ToString("x8"));
                }
                Log("     f:" + fs);
            }
        } catch (Exception ex) { Log("dump err " + ex.Message); }
    }

    static void Work() {
        try {
            OleInitialize(IntPtr.Zero);
            SetDllDirectoryW(nkDir);
            if (LoadLibraryA(nkDir + @"\hwplt32.dll") == IntPtr.Zero) { error = "LoadLibrary hwplt32.dll failed"; return; }
            InitImageGear();
            if (!LoadTerControl()) { error = "LoadTerControl failed"; return; }
            IntPtr host = CreateWindowExA(0, "Static", "h", 0x80000000u, 0, 0, 100, 100,
                IntPtr.Zero, IntPtr.Zero, GetModuleHandleA(null), IntPtr.Zero);
            IntPtr ter = CreateWindowExA(0, "Ter32Class", "", 0x56000000u, 0, 0, 900, 700,
                host, IntPtr.Zero, GetModuleHandleA(null), IntPtr.Zero);
            if (ter == IntPtr.Zero) { error = "CreateWindowEx Ter32Class failed"; return; }
            Log("ter=" + ter);
            IntPtr doc = new IntPtr(GetWindowLongA(ter, 0));
            stage = 1;
            if (fontTest != null) {
                // Probe SetTerPointSize units: insert three lines, size each with a
                // different value, dump the resulting font records +0x68.
                for (int i = 0; i < 3; i++) {
                    byte[] tb = Encoding.GetEncoding(949).GetBytes("LINE" + i + " size" + (i == 0 ? 10 : i == 1 ? 20 : 24));
                    TerInsertLine(ter, tb, -1, -1, -1, -1, 0);
                    SelectTerText(ter, i, 0, i, tb.Length, false);
                    int sv = i == 0 ? 10 : i == 1 ? 20 : 24;
                    SetTerPointSize(ter, sv, false);
                    DeselectTerText(ter, false);
                }
                doc = new IntPtr(GetWindowLongA(ter, 0));
                DumpDoc(doc); done = true; stage = 4; return;
            }
            if (dtPath != null) {
                // Read (and optionally set+save) the 문서종류 integer for a doc.
                int l = ReadTerFile(ter, dtPath);
                Log("read=" + l + " doctype=" + TerGetDocType(ter));
                if (dtSet != -999) {
                    Log("set-> " + TerSetDocType(ter, dtSet) + " now=" + TerGetDocType(ter));
                    if (outFile != null) Log("save=" + SaveTerFile(ter, outFile));
                }
                DumpDoc(doc); done = true; stage = 4; return;
            }
            if (rtIn != null) {
                // Round-trip test: load a structured doc, insert a probe line, save,
                // to confirm ReadTerFile+SaveTerFile preserves the element structure.
                int l = ReadTerFile(ter, rtIn); Log("read=" + l);
                int targetLine = rtLine >= 0 ? rtLine : 1;
                SetTerCursorPos(ter, targetLine, 0, 0);
                TerInsertLine(ter, Encoding.GetEncoding(949).GetBytes(rtText), -1, -1, -1, -1, 0);
                Log("save=" + SaveTerFile(ter, rtOut)); DumpDoc(doc); done = true; stage = 4; return;
            }
            if (impPath != null) {
                int ir = TerImportFile(ter, impPath, impA, impB);
                Log("TerImportFile=" + ir + " doctype=" + TerGetDocType(ter));
                doc = new IntPtr(GetWindowLongA(ter, 0));
                if (outFile != null) Log("save=" + SaveTerFile(ter, outFile));
                DumpDoc(doc); done = true; stage = 4; return;
            }
            if (dumpPath != null) {
                if (imgTest != null) {
                    // Insert an image and dump the object record TerInsertPictureFile
                    // populates — reveals the natural width/height field layout.
                    int oid = TerInsertPictureFile(ter, imgTest, 0, 0, 0);
                    Log("insert oid=" + oid);
                    doc = new IntPtr(GetWindowLongA(ter, 0));
                    DumpDoc(doc); done = true; stage = 4; return;
                }
                int lr = ReadTerFile(ter, dumpPath);
                Log("ReadTerFile=" + lr);
                doc = new IntPtr(GetWindowLongA(ter, 0));
                DumpDoc(doc); done = true; stage = 4; return;
            }

            int p = 0;
            uint dlen = ReadU32(job, ref p);
            string imgdir = Encoding.UTF8.GetString(job, p, (int)dlen); p += (int)dlen;
            // TerSetImagePath strips the last path component (it expects a file path
            // and keeps only the directory part). Trailing "\" keeps imgdir intact so
            // the save reads HLT%N.HIF from our staging dir, not its parent.
            if (imgdir.Length > 0) TerSetImagePath(ter, imgdir.TrimEnd('\\') + "\\");

            // Structured mode: a non-empty skeleton path means build on that
            // .frm form so the output carries the real KIPO element tree.
            uint slen = ReadU32(job, ref p);
            string skel = slen > 0 ? Encoding.UTF8.GetString(job, p, (int)slen) : ""; p += (int)slen;
            uint nsec = ReadU32(job, ref p);
            var secNames = new List<string>();
            for (uint i = 0; i < nsec; i++) {
                uint tl = ReadU32(job, ref p);
                secNames.Add(Encoding.GetEncoding(949).GetString(job, p, (int)tl)); p += (int)tl;
            }
            uint nlines = ReadU32(job, ref p);
            bool structured = skel.Length > 0;
            Log("lines=" + nlines + " secs=" + nsec + " structured=" + structured);

            if (structured) {
                int lr = ReadTerFile(ter, skel);
                Log("skel read=" + lr);
                doc = new IntPtr(GetWindowLongA(ter, 0));
            }

            var fmts = new List<int[]>();
            var imgInfo = new Dictionary<string, ImgRec>();  // basename -> src+dims
            var imgJobs = new List<ImgRec>();                // image order -> src+dims+reserved line
            var paraIds = new Dictionary<string, int>();     // para spec -> paraId
            int nimg = 0;

            // Pass 1a: parse every job line into a LineRec (paraIds are created
            // now, against whatever doc is loaded). Placement happens in pass 1b.
            var recs = new List<LineRec>();
            for (int li = 0; li < (int)nlines; li++) {
                var r = new LineRec { sec = (int)ReadU32(job, ref p) };
                r.kind = job[p++];
                if (r.kind == 1) {
                    uint plen = ReadU32(job, ref p);
                    r.src = Encoding.UTF8.GetString(job, p, (int)plen); p += (int)plen;
                    r.dw = (int)ReadU32(job, ref p); r.dh = (int)ReadU32(job, ref p);
                    r.pw = (int)ReadU32(job, ref p); r.ph = (int)ReadU32(job, ref p);
                } else {
                    uint tlen = ReadU32(job, ref p);
                    r.text = new byte[tlen];
                    Array.Copy(job, p, r.text, 0, (int)tlen); p += (int)tlen;
                    r.term = job[p++] != 0;
                    uint nfmt = ReadU32(job, ref p);
                    for (uint k = 0; k < nfmt; k++) {
                        int s = (int)ReadU32(job, ref p), e = (int)ReadU32(job, ref p), mk = (int)ReadU32(job, ref p);
                        int sz = (int)ReadU32(job, ref p); uint col = ReadU32(job, ref p);
                        r.fmts.Add(new int[] { s, e, mk, sz, (int)col });
                    }
                }
                r.paraId = ReadParaSpec(job, ref p, ter, doc, paraIds);
                recs.Add(r);
            }

            if (!structured) {
                // Flat doc: append every line at the moving cursor.
                int cur = 0;
                foreach (var r in recs) {
                    if (r.kind == 1) {
                        int reserved = Marshal.ReadInt32(doc, 0x525);
                        if (!TerInsertLine(ter, new byte[] { 0x15 }, -1, r.paraId, -1, -1, 0)) { error = "img reserve failed"; return; }
                        r.land = reserved;
                        imgJobs.Add(new ImgRec { src = r.src, pw = r.pw, ph = r.ph, dw = r.dw, dh = r.dh, line = reserved });
                        imgInfo[Base(r.src)] = new ImgRec { src = r.src, pw = r.pw, ph = r.ph, dw = r.dw, dh = r.dh, line = reserved };
                        cur++;
                        continue;
                    }
                    byte[] buf = new byte[r.text.Length + (r.term ? 1 : 0)];
                    Array.Copy(r.text, buf, r.text.Length);
                    if (r.term) buf[r.text.Length] = 0x15;
                    if (!TerInsertLine(ter, buf, -1, r.paraId, -1, -1, 0)) { error = "TerInsertLine failed"; return; }
                    r.land = cur;
                    foreach (var f in r.fmts) fmts.Add(new int[] { cur, f[0], f[1], f[2], f[3], f[4] });
                    cur++;
                }
            } else {
                // Structured: route each section's lines into its skeleton
                // element. Insert positions are the line right after the
                // element's 【marker】; walk elements top-down so an earlier
                // insert only shifts lines below it (tracked via `offset`).
                var texts = LineTexts(doc);
                var elemMarker = new Dictionary<string, int>();
                var markerRe = new System.Text.RegularExpressions.Regex(@"^【(.+)】");
                for (int i = 0; i < texts.Count; i++) {
                    var mm = markerRe.Match(texts[i]);
                    if (mm.Success && !elemMarker.ContainsKey(mm.Groups[1].Value))
                        elemMarker[mm.Groups[1].Value] = i;
                }
                // group recs by section index (preserve order)
                var bySec = new Dictionary<int, List<LineRec>>();
                foreach (var r in recs) {
                    if (!bySec.ContainsKey(r.sec)) bySec[r.sec] = new List<LineRec>();
                    bySec[r.sec].Add(r);
                }
                // markerLine -> content recs (merge same-element sections; orphans
                // flow into the element that follows them in docx order)
                var atPos = new SortedDictionary<int, List<LineRec>>();
                var pending = new List<LineRec>();
                int lastPos = -1;
                for (int si = 0; si < (int)nsec; si++) {
                    List<LineRec> lines; if (!bySec.TryGetValue(si, out lines)) lines = new List<LineRec>();
                    int ml = (si < secNames.Count && elemMarker.ContainsKey(secNames[si])) ? elemMarker[secNames[si]] : -1;
                    if (ml < 0) { pending.AddRange(lines); continue; }
                    if (!atPos.ContainsKey(ml)) atPos[ml] = new List<LineRec>();
                    atPos[ml].AddRange(pending); pending.Clear();
                    atPos[ml].AddRange(lines); lastPos = ml;
                }
                if (pending.Count > 0) {
                    if (lastPos >= 0) atPos[lastPos].AddRange(pending);
                    else if (atPos.Count > 0) atPos[atPos.Keys.Min()].InsertRange(0, pending);
                }
                int offset = 0;
                foreach (var kv in atPos) {
                    int pos = kv.Key + 1 + offset;
                    foreach (var r in kv.Value) {
                        r.land = pos;
                        SetTerCursorPos(ter, pos, 0, 0);
                        if (r.kind == 1) {
                            if (!TerInsertLine(ter, new byte[] { 0x15 }, -1, r.paraId, -1, -1, 0)) { Log("warn: img reserve at " + pos); continue; }
                            imgJobs.Add(new ImgRec { src = r.src, pw = r.pw, ph = r.ph, dw = r.dw, dh = r.dh, line = pos });
                            imgInfo[Base(r.src)] = new ImgRec { src = r.src, pw = r.pw, ph = r.ph, dw = r.dw, dh = r.dh, line = pos };
                        } else {
                            byte[] buf = new byte[r.text.Length + (r.term ? 1 : 0)];
                            Array.Copy(r.text, buf, r.text.Length);
                            if (r.term) buf[r.text.Length] = 0x15;
                            if (!TerInsertLine(ter, buf, -1, r.paraId, -1, -1, 0)) { Log("warn: insert at " + pos); continue; }
                            foreach (var f in r.fmts) fmts.Add(new int[] { pos, f[0], f[1], f[2], f[3], f[4] });
                        }
                        pos++;
                    }
                    offset += kv.Value.Count;
                }
            }
            nimg = imgJobs.Count;
            Log("images=" + nimg);
            // Pass 2: anchor each image at its reserved line. No text is appended
            // between markers, so the marker stays put instead of riding to the end.
            foreach (ImgRec ij in imgJobs) {
                SetTerCursorPos(ter, ij.line, 0, 0);
                int oid = TerInsertPictureFile(ter, ij.src, 0, 0, 0);
                if (oid < 0) { Log("warn: pic insert failed " + ij.src); continue; }
                TerInsertObjectId(ter, oid, 0);
            }
            stage = 2;

            // Post-pass: the object array is now stable. Match each embedded-image
            // object (flag 0x80, no data index yet) to its source via the path the
            // control stored at +0x96, then assign a data index, stage the HIF file
            // (raw source bytes), and fill the dimension fields the failed
            // IG_info_get left unset.
            if (nimg > 0) {
                IntPtr arr = Marshal.ReadIntPtr(doc, 0x9366);
                int ocnt = Marshal.ReadInt32(doc, 0x936e);
                Log("workdir=" + ReadCStr(doc, 0x1a65e));
                Log("imagepath=" + ReadCStr(doc, 0x1a75e));
                int hifIdx = 0, assigned = 0;
                for (int oi = 0; oi < ocnt; oi++) {
                    IntPtr ob = IntPtr.Add(arr, oi * 0x1ce);
                    if ((Marshal.ReadInt32(ob, 0x1b2) & 0x80) == 0 || Marshal.ReadInt32(ob, 0x4) >= 0) continue;
                    ImgRec r;
                    if (!imgInfo.TryGetValue(Base(ReadCStr(ob, 0x96)), out r)) { Log("warn: no src for obj " + oi); continue; }
                    hifIdx++;
                    try {
                        // The save materializes obj+0x4=N by reading
                        // GetTempPath()\HLT%N.HIF in the format
                        //   [u32 little-endian byte length][image data]
                        // then feeding `image data` to IG_load_mem. Writing the
                        // raw source bytes without the length prefix makes the
                        // reader consume the first 4 bytes of the image as the
                        // length and pass a truncated/corrupt buffer to
                        // IG_load_mem -> decode fails -> broken (X) image box.
                        byte[] src = File.ReadAllBytes(r.src);
                        byte[] hif = new byte[src.Length + 4];
                        hif[0] = (byte)src.Length; hif[1] = (byte)(src.Length >> 8);
                        hif[2] = (byte)(src.Length >> 16); hif[3] = (byte)(src.Length >> 24);
                        Array.Copy(src, 0, hif, 4, src.Length);
                        File.WriteAllBytes(Path.Combine(imgdir, "HLT" + hifIdx.ToString("x") + ".HIF"), hif);
                    } catch (Exception ex) { Log("warn: hif write " + ex.Message); }
                    Marshal.WriteInt32(ob, 0x4, hifIdx);          // data index -> HLT%N.HIF
                    // Reproduce the docx display extent via the proper API — it
                    // converts points to internal units and writes the right record
                    // fields (writing raw offsets gets width/height transposed).
                    if (r.dw > 0 && r.dh > 0) TerSetPictSize(ter, oi, r.dw, r.dh);
                    assigned++;
                }
                Log("img objects assigned=" + assigned + "/" + nimg);
            }

            int applied = 0;
            foreach (var f in fmts) {
                if (!SelectTerText(ter, f[0], f[1], f[0], f[2], false)) { Log("warn: select failed line=" + f[0]); continue; }
                bool ok = true;
                if (f[3] != 0) ok = SetTerCharStyle(ter, f[3], 1, false);
                // SetTerPointSize takes whole points (font rec +0x28 renders pt,
                // not the half-points that w:sz uses) — so halve the job value.
                if (f[4] >= 0) SetTerPointSize(ter, (f[4] + 1) / 2, false);
                if (f[5] != -1) SetTerColor(ter, f[5], false);         // 0x00BBGGRR
                if (ok) applied++;
                DeselectTerText(ter, false);
            }
            Log("fmts=" + applied + "/" + fmts.Count);
            stage = 3;

            if (structured) {
                // 문서종류 = 특허명세서 (0x8000): marks the doc as a structured
                // patent specification so 파일정보 and 제출파일 recognise it.
                Log("doctype set-> " + TerSetDocType(ter, 0x8000) + " now=" + TerGetDocType(ter));
            }
            if (!SaveTerFile(ter, outFile)) { error = "SaveTerFile failed"; return; }
            Log("saved=" + outFile);
            stage = 4;
        } catch (Exception ex) { error = ex.ToString(); }
    }

    public static int Main(string[] args) {
        if (args.Length < 2) { Console.Error.WriteLine("usage: TerConv.exe <jobfile> <outfile.hlt> [NKEditorDir]"); return 2; }
        if (args[0] == "dump") { dumpPath = args[1]; }
        else if (args[0] == "rt") { rtIn = args[1]; rtOut = args[2];
            if (args.Length > 3) rtLine = int.Parse(args[3]);
            if (args.Length > 4) rtText = args[4]; }
        else if (args[0] == "fonttest") { fontTest = args[1]; }
        else if (args[0] == "imgtest") { imgTest = args[1]; dumpPath = "x"; }
        else if (args[0] == "doctype") { dtPath = args[1]; if (args.Length > 2) dtSet = int.Parse(args[2]); if (args.Length > 3) outFile = args[3]; }
        else if (args[0] == "import") { impPath = args[1]; if (args.Length > 2) outFile = args[2]; if (args.Length > 3) impA = int.Parse(args[3]); if (args.Length > 4) impB = int.Parse(args[4]); }
        else {
            if (args.Length > 2) nkDir = args[2];
            try { job = File.ReadAllBytes(args[0]); } catch (Exception ex) { Console.Error.WriteLine("jobfile: " + ex.Message); return 2; }
            outFile = args[1];
        }

        var wd = new Thread(Watchdog); wd.IsBackground = true; wd.Start();
        var t = new Thread(Work); t.SetApartmentState(ApartmentState.STA); t.Start();
        for (int i = 0; i < 1200 && stage < 4 && error == null; i++) Thread.Sleep(100);
        done = true;
        if (error != null) { Console.Error.WriteLine("ERROR: " + error); return 1; }
        if (stage < 4) { Console.Error.WriteLine("ERROR: timeout at stage " + stage); return 1; }
        return 0;
    }
}
