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

    static object Prop(object obj, string name)
    {
        return obj.GetType().InvokeMember(name, BindingFlags.GetProperty,
            null, obj, null);
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
