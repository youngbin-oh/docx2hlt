// DocCmp.exe — compare a source .docx against a converted .rtf and write a
// Word tracked-changes comparison document.
//
//   DocCmp.exe original.docx converted.rtf out.compare.docx
//
// Uses Word COM (late binding — no PIA dependency), same pattern as
// EqRender.cs. CompareDocuments produces a new document where every
// difference between the source and the round-tripped RTF appears as a
// tracked insertion/deletion — exactly what a reviewer needs to see what
// the hlt conversion changed.
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

public class DocCmp
{
    static object M = Type.Missing;

    static object Call(object obj, string name, params object[] args)
    {
        return obj.GetType().InvokeMember(name,
            BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding,
            null, obj, args);
    }

    static object Prop(object obj, string name, params object[] args)
    {
        return obj.GetType().InvokeMember(name, BindingFlags.GetProperty,
            null, obj, args.Length > 0 ? args : null);
    }

    static void SetProp(object obj, string name, object val)
    {
        obj.GetType().InvokeMember(name, BindingFlags.SetProperty,
            null, obj, new object[] { val });
    }

    static string ParaText(object rng)
    {
        return ((string)Prop(rng, "Text")).Trim('\r', '\a', '\n', ' ', '\t',
            '\u000b');
    }

    // With CompareTables off, a source table the conversion turned into an
    // image (hlt tables are TIFF) vanishes from the compare doc without a
    // trace: only the image shows up as inserted. With CompareTables on, Word
    // puts that image inside the first cell of the deleted table, untracked,
    // so Accept All would drop the image too. Instead copy each source table
    // that has no table at its spot in the compare doc back in, right after
    // the nearest non-empty paragraph before it (e.g. 【표 1】), and delete it
    // with tracking on — it then reads as a tracked table deletion next to
    // the inserted image.
    // Word's compare leaves the table's replacement image as an untracked run
    // (only its paragraph mark is a revision), so Reject All would keep it.
    // Re-insert each image of the picture-only paragraphs right after `pos`
    // with tracking on and drop the untracked original.
    static void MarkImagesInserted(object dCmp, int pos)
    {
        int docEnd = (int)Prop(Prop(dCmp, "Content"), "End");
        object paras = Prop(Call(dCmp, "Range", pos, Math.Min(pos + 1, docEnd)),
                            "Paragraphs");
        if ((int)Prop(paras, "Count") == 0) return;
        object para = Call(paras, "Item", 1);
        for (int k = 0; k < 3 && para != null; k++)
        {
            object pr = Prop(para, "Range");
            object shapes = Prop(pr, "InlineShapes");
            int ns = (int)Prop(shapes, "Count");
            string txt = ParaText(pr).Replace("/", "").Replace("\u0001", "");
            if (ns > 0 && txt.Length > 0) break;      // body text: stop
            for (int s = ns; s >= 1; s--)
            {
                object sr = Prop(Call(shapes, "Item", s), "Range");
                if ((int)Prop(Prop(sr, "Revisions"), "Count") > 0) continue;
                int a = (int)Prop(sr, "Start"), b = (int)Prop(sr, "End");
                SetProp(dCmp, "TrackRevisions", true);
                SetProp(Call(dCmp, "Range", a, a), "FormattedText",
                        Prop(sr, "FormattedText"));
                SetProp(dCmp, "TrackRevisions", false);
                int len = b - a;
                Call(Call(dCmp, "Range", a + len, b + len), "Delete");
            }
            if (ns > 0) break;
            try { para = Call(para, "Next"); } catch { para = null; }
        }
        SetProp(dCmp, "TrackRevisions", false);
    }

    static int RestoreDeletedTables(object app, object dOrig, object dCmp)
    {
        object tables = Prop(dOrig, "Tables");
        int n = (int)Prop(tables, "Count");
        if (n == 0) return 0;
        string author = "비교";      // Word's Korean compare author
        object revs = Prop(dCmp, "Revisions");
        if ((int)Prop(revs, "Count") > 0)
            author = (string)Prop(Call(revs, "Item", 1), "Author");
        string oldUser = (string)Prop(app, "UserName");
        bool oldTrack = (bool)Prop(dCmp, "TrackRevisions");
        int restored = 0, searchFrom = 0;
        try
        {
            SetProp(app, "UserName", author);
            for (int i = 1; i <= n; i++)
            {
                object tRange = Prop(Call(tables, "Item", i), "Range");
                int tStart = (int)Prop(tRange, "Start");
                // anchor: last non-empty paragraph before the table
                string anchor = null;
                object before = Prop(Call(dOrig, "Range", 0, tStart),
                                     "Paragraphs");
                int pc = (int)Prop(before, "Count");
                for (int k = pc; k >= 1 && k > pc - 6 && anchor == null; k--)
                {
                    string s = ParaText(Prop(Call(before, "Item", k), "Range"));
                    if (s.Length > 0) anchor = s;
                }
                if (anchor == null)
                {
                    Console.WriteLine("WARN table " + i +
                        ": no text before it to place a tracked deletion");
                    continue;
                }
                if (anchor.Length > 200) anchor = anchor.Substring(anchor.Length - 200);
                int docEnd = (int)Prop(Prop(dCmp, "Content"), "End");
                object rng = Call(dCmp, "Range", searchFrom, docEnd);
                object find = Prop(rng, "Find");
                Call(find, "ClearFormatting");
                SetProp(find, "Text", anchor.Replace("^", "^^"));
                SetProp(find, "Forward", true);
                SetProp(find, "Wrap", 0);            // wdFindStop
                SetProp(find, "MatchWildcards", false);
                SetProp(find, "MatchCase", true);
                if (!(bool)Call(find, "Execute"))
                {
                    Console.WriteLine("WARN table " + i + ": anchor '" +
                        (anchor.Length > 30 ? anchor.Substring(0, 30) : anchor) +
                        "' not found in the compare doc");
                    continue;
                }
                object para = Call(Prop(rng, "Paragraphs"), "Item", 1);
                int p = (int)Prop(Prop(para, "Range"), "End");
                searchFrom = p;
                if (p >= docEnd) continue;
                // the table survived the conversion: Word compared it already
                if ((bool)Prop(Call(dCmp, "Range", p, p), "Information", 12))
                    continue;                        // wdWithInTable
                SetProp(dCmp, "TrackRevisions", false);
                SetProp(Call(dCmp, "Range", p, p), "FormattedText",
                        Prop(tRange, "FormattedText"));
                object ins = Prop(Call(dCmp, "Range", p, p + 1), "Tables");
                if ((int)Prop(ins, "Count") == 0)
                {
                    Console.WriteLine("WARN table " + i + ": copy failed");
                    continue;
                }
                object tbl = Call(ins, "Item", 1);
                int tEnd = (int)Prop(Prop(tbl, "Range"), "End");
                MarkImagesInserted(dCmp, tEnd);
                SetProp(dCmp, "TrackRevisions", true);
                Call(tbl, "Delete");
                searchFrom = tEnd;
                restored++;
            }
        }
        finally
        {
            try { SetProp(dCmp, "TrackRevisions", oldTrack); } catch { }
            try { SetProp(app, "UserName", oldUser); } catch { }
        }
        return restored;
    }

    public static int Main(string[] args)
    {
        if (args.Length < 3)
        {
            Console.Error.WriteLine("usage: DocCmp.exe orig.docx conv.rtf out.docx");
            return 2;
        }
        string orig = Path.GetFullPath(args[0]);
        string rev = Path.GetFullPath(args[1]);
        string outp = Path.GetFullPath(args[2]);

        object app = null, docs = null, dOrig = null, dRev = null, dCmp = null;
        try
        {
            Type t = Type.GetTypeFromProgID("Word.Application");
            app = Activator.CreateInstance(t);
            try { Prop(app, "Visible"); } catch { }
            app.GetType().InvokeMember("Visible", BindingFlags.SetProperty,
                null, app, new object[] { false });
            app.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty,
                null, app, new object[] { 0 });   // wdAlertsNone

            docs = Prop(app, "Documents");
            // Open(FileName, ...15 optional)
            object[] oargs = new object[16];
            oargs[0] = orig;
            for (int i = 1; i < oargs.Length; i++) oargs[i] = M;
            dOrig = Call(docs, "Open", oargs);

            object[] rargs = new object[16];
            rargs[0] = rev;
            for (int i = 1; i < rargs.Length; i++) rargs[i] = M;
            dRev = Call(docs, "Open", rargs);

            // CompareDocuments(Original, Revised, Destination, Granularity,
            //   CompareFormatting, CompareCaseChanges, CompareWhitespace,
            //   CompareTables, CompareHeaders, CompareFooters, CompareFootnotes,
            //   CompareTextboxes, CompareFields, CompareComments, CompareMoves,
            //   RevisedAuthor, IgnoreAllComparisonWarnings)
            // Formatting compare is OFF by default: the rtf round-trip loses
            // fonts/spacing anyway, which buries the character-level changes
            // the reviewer actually cares about. '/fmt' turns it back on.
            bool fmt = Array.IndexOf(args, "/fmt") >= 0;
            try
            {
                // named-parameter binding: correct regardless of the Word
                // version's parameter ordering. Formatting compare OFF by
                // default — the rtf round-trip loses fonts/spacing anyway,
                // which buries the character changes reviewers care about.
                object[] cargs = {
                    dOrig, dRev,
                    2,                 // Destination: new doc
                    fmt,               // CompareFormatting
                    fmt,               // CompareTables
                    false, false, false, false, false, false,
                    true               // IgnoreAllComparisonWarnings
                };
                string[] cnames = {
                    "OriginalDocument", "RevisedDocument",
                    "Destination", "CompareFormatting", "CompareTables",
                    "CompareHeaders", "CompareFootnotes",
                    "CompareTextboxes", "CompareFields", "CompareComments",
                    "CompareMoves", "IgnoreAllComparisonWarnings"
                };
                dCmp = app.GetType().InvokeMember("CompareDocuments",
                    BindingFlags.InvokeMethod, null, app, cargs,
                    null, null, cnames);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("named call failed: " +
                    (ex.InnerException ?? ex).Message);
                // fallback: minimal positional call (defaults everything on)
                object[] c2 = new object[14];
                c2[0] = dOrig; c2[1] = dRev;
                for (int i = 2; i < c2.Length; i++) c2[i] = M;
                dCmp = Call(app, "CompareDocuments", c2);
            }
            if (dCmp == null)
            {
                Console.Error.WriteLine("CompareDocuments returned null");
                return 3;
            }

            if (!fmt)
            {
                try
                {
                    int nt = RestoreDeletedTables(app, dOrig, dCmp);
                    if (nt > 0)
                        Console.WriteLine("tables shown as tracked deletions: " + nt);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("WARN table restore failed: " +
                        (ex.InnerException ?? ex).Message);
                }
            }

            // SaveAs2(FileName, FileFormat=16 wdFormatDocumentDefault)
            object[] sargs = new object[17];
            sargs[0] = outp; sargs[1] = 16;
            for (int i = 2; i < sargs.Length; i++) sargs[i] = M;
            Call(dCmp, "SaveAs2", sargs);
            Console.WriteLine("OK -> " + outp);
            return 0;
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("DocCmp failed: " + e);
            return 1;
        }
        finally
        {
            // close docs without saving; compare doc already saved
            foreach (object d in new object[] { dCmp, dRev, dOrig })
            {
                if (d == null) continue;
                try { Call(d, "Close", new object[] { 0 }); } catch { }
            }
            if (app != null)
            {
                try { Call(app, "Quit", new object[] { 0, M, M }); } catch { }
                try { Marshal.FinalReleaseComObject(app); } catch { }
            }
        }
    }
}
