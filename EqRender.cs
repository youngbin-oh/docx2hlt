// EqRender — render every Word OMML equation (OMath) in a .docx to a
// tightly-trimmed 300dpi CCITT-G4 TIFF, without modifying the document.
//
//   EqRender.exe <in.docx> <outdir>
//
// For each equation writes <outdir>\eq<NNNN>.tif, where NNNN is the
// equation's 1-based index in Word's doc.OMaths (document order, main story,
// tracked-deleted equations included). Keying by that index instead of a
// running counter lets docx2hlt align images with the m:oMath elements it
// walks even when some equations are skipped (deleted) or fail to render.
// It also
// appends "index<TAB>w_mm<TAB>h_mm" to <outdir>\eqmap.tsv where w/h are the
// intended display size (physical size of the rendered range scaled by the
// trim ratio). docx2hlt splices these images back into the paragraph text.
//
// Approach: Range.CopyAsPicture() puts an EnhancedMetafile of the equation
// — exactly as Word renders it — on the clipboard. We pull the EMF bits off
// the clipboard, rasterize at 300dpi, trim whitespace, threshold to bilevel
// and save as Group-4 TIFF (same recipe as the VSTO add-in's
// BuildTrimmedTiffBytes, minus the document mutation).
// Equations take the EMF from Range.EnhMetaFileBits instead when Word gives
// it: the clipboard round trip costs 5-10 s per object over a remote-desktop
// session (148 equations, 21 min; EnhMetaFileBits: 25 s, same
// drawing and size). Tables stay on the clipboard: EnhMetaFileBits lays a
// table wider than the text column out narrower than Word shows it.
//
// Uses late-bound Word COM (no interop assembly needed) and plain .NET 4.x.
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

internal static class EqRender
{
    private const float Dpi = 300f;
    private const int CropPaddingPx = 6;
    private const float MaxWidthMm = 165f;    // KIPO page box (E-218)
    private const float MaxHeightMm = 222f;

    [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr hWndNewOwner);
    [DllImport("user32.dll")] private static extern bool CloseClipboard();
    [DllImport("user32.dll")] private static extern IntPtr GetClipboardData(uint uFormat);
    [DllImport("user32.dll")] private static extern bool EmptyClipboard();
    [DllImport("gdi32.dll")] private static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, string file);
    [DllImport("gdi32.dll")] private static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, byte[] lpbBuffer);
    [DllImport("gdi32.dll")] private static extern bool DeleteEnhMetaFile(IntPtr hemf);
    [DllImport("gdi32.dll")] private static extern uint GetEnhMetaFileHeader(IntPtr hemf, uint cbBuffer, byte[] lpemh);

    private const uint CF_ENHMETAFILE = 14;

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("usage: EqRender <in.docx> <outdir>");
            return 2;
        }
        try
        {
            return Run(args[0], args[1]);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("EQRENDER-FAIL " + ex);
            return 1;
        }
    }

    private static int Run(string srcPath, string outdir)
    {
        Directory.CreateDirectory(outdir);
        // Work on a temp copy: the user may already have the doc open in Word.
        string copy = Path.Combine(outdir, "_eqsrc" + Guid.NewGuid().ToString("N") + ".docx");
        File.Copy(srcPath, copy);

        Type wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Console.Error.WriteLine("EQRENDER-FAIL Word is not installed");
            return 1;
        }
        dynamic app = Activator.CreateInstance(wordType);
        dynamic doc = null;
        int rendered = 0, eqSkipped = 0;
        var map = new StringWriter();
        var tmap = new StringWriter();
        int topLvl = 0, tblSkipped = 0, tblRendered = 0;
        try
        {
            app.Visible = false;
            app.DisplayAlerts = 0;              // wdAlertsNone
            app.Options.CheckGrammarAsYouType = false;
            app.Options.CheckSpellingAsYouType = false;
            app.AutomationSecurity = 3;         // msoAutomationSecurityForceDisable
            doc = app.Documents.Open(
                copy, ConfirmConversions: false, ReadOnly: true,
                AddToRecentFiles: false, Visible: false,
                OpenAndRepair: true, NoEncodingDialog: true);

            dynamic omaths = doc.OMaths;
            int count = 0;
            try { count = omaths.Count; } catch { }
            Console.WriteLine("OMATHS " + count);

            for (int i = 1; i <= count; i++)
            {
                dynamic om = null;
                dynamic range = null;
                try
                {
                    om = omaths[i];
                    if (ContainsDeletedRevision(om.Range))
                    {
                        eqSkipped++;
                        continue;
                    }
                    range = om.Range.Duplicate;
                    byte[] emf = RangeEmfBits(range) ?? CopyRangeAsEmf(range);
                    if (emf == null || emf.Length == 0)
                    {
                        Console.Error.WriteLine("eq " + i + ": no EMF on clipboard");
                        continue;
                    }
                    string tif = Path.Combine(outdir, string.Format("eq{0:D4}.tif", i));
                    float wmm, hmm;
                    SaveTrimmedTiff(UnclipEquationEmf(emf), tif, out wmm, out hmm);
                    map.WriteLine(i + "\t" + Mm(wmm) + "\t" + Mm(hmm));
                    rendered++;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("eq " + i + ": " + ex.Message);
                }
                finally
                {
                    Release((object)range); Release((object)om);
                }
            }
            Release((object)omaths);
            Console.WriteLine(string.Format(
                "EQSUMMARY total={0} skipped={1} rendered={2}",
                count, eqSkipped, rendered));

            // Tables: same visual pipeline — every top-level w:tbl
            // (NestingLevel 1) becomes tbl<ordinal>.tif where ordinal is
            // its document-order index among top-level tables, so docx2hlt
            // can align them with the w:tbl nodes it walks. A table that
            // fails to render simply has no tblmap.tsv entry and the
            // caller falls back to flattening just that table's text.
            dynamic tables = doc.Tables;
            int tcount = 0;
            try { tcount = tables.Count; } catch { }
            Console.WriteLine("TABLES " + tcount);

            for (int i = 1; i <= tcount; i++)
            {
                dynamic tbl = null;
                dynamic range = null;
                try
                {
                    tbl = tables[i];
                    int nest = 1;
                    try { nest = tbl.NestingLevel; } catch { }
                    if (nest != 1)
                    {
                        continue;
                    }
                    topLvl++;
                    if (ContainsDeletedRevision(tbl.Range))
                    {
                        tblSkipped++;
                        continue;
                    }
                    range = tbl.Range.Duplicate;
                    byte[] emf = CopyRangeAsEmf(range);
                    if (emf == null || emf.Length == 0)
                    {
                        Console.Error.WriteLine("tbl " + i + ": no EMF on clipboard");
                        continue;
                    }
                    string tif = Path.Combine(outdir, string.Format("tbl{0:D4}.tif", topLvl));
                    float wmm, hmm;
                    SaveTrimmedTiff(emf, tif, out wmm, out hmm);
                    tmap.WriteLine(topLvl + "\t" + Mm(wmm) + "\t" + Mm(hmm));
                    tblRendered++;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("tbl " + i + ": " + ex.Message);
                }
                finally
                {
                    Release((object)range); Release((object)tbl);
                }
            }
            Release((object)tables);
            Console.WriteLine(string.Format(
                "TBLSUMMARY toplevel={0} skipped={1} rendered={2}",
                topLvl, tblSkipped, tblRendered));
        }
        finally
        {
            try { if (doc != null) doc.Close(false); } catch { }
            try { app.Quit(false); } catch { }
            Release((object)doc); Release((object)app);
            try { File.Delete(copy); } catch { }
        }
        File.WriteAllText(Path.Combine(outdir, "eqmap.tsv"), map.ToString());
        File.WriteAllText(Path.Combine(outdir, "tblmap.tsv"), tmap.ToString());
        Console.WriteLine("OK rendered=" + rendered + " tables=" + tblRendered);
        return 0;
    }

    // --- clipboard ---------------------------------------------------------

    private static byte[] RangeEmfBits(dynamic range)
    {
        try
        {
            byte[] b = range.EnhMetaFileBits as byte[];
            return b != null && b.Length > 0 ? b : null;
        }
        catch
        {
            return null;
        }
    }

    private static byte[] CopyRangeAsEmf(dynamic range)
    {
        for (int attempt = 0; attempt < 12; attempt++)
        {
            try
            {
                OpenClipboard(IntPtr.Zero);
                EmptyClipboard();
                CloseClipboard();
            }
            catch { }

            try { range.CopyAsPicture(); }
            catch (Exception ex)
            {
                Console.Error.WriteLine("CopyAsPicture: " + ex.Message);
            }

            byte[] emf = ReadClipboardEmf();
            if (emf != null && emf.Length > 0)
            {
                return emf;
            }
            Application.DoEvents();
            Thread.Sleep(200);
        }
        return null;
    }

    private static byte[] ReadClipboardEmf()
    {
        try
        {
            if (!OpenClipboard(IntPtr.Zero))
            {
                return null;
            }
            IntPtr h = GetClipboardData(CF_ENHMETAFILE);
            if (h == IntPtr.Zero)
            {
                return null;
            }
            IntPtr copy = CopyEnhMetaFile(h, null);
            if (copy == IntPtr.Zero)
            {
                return null;
            }
            try
            {
                uint n = GetEnhMetaFileBits(copy, 0, null);
                if (n == 0)
                {
                    return null;
                }
                byte[] buf = new byte[n];
                GetEnhMetaFileBits(copy, n, buf);
                return buf;
            }
            finally
            {
                DeleteEnhMetaFile(copy);
            }
        }
        catch
        {
            return null;
        }
        finally
        {
            CloseClipboard();
        }
    }

    // --- equation EMF unclip ----------------------------------------------

    private const uint EMR_EXTTEXTOUTA = 83;
    private const uint EMR_EXTTEXTOUTW = 84;
    private const uint EMR_EOF = 14;
    private const uint ETO_CLIPPED = 0x4;

    // Word's CopyAsPicture of an OMath gives a picture one text line tall
    // (18pt) and draws every glyph with ETO_CLIPPED to that line box, so a
    // fraction's numerator top and its whole denominator are cut off even
    // though their text records are in the EMF. Clear ETO_CLIPPED on the text
    // records and grow the header's bounds/frame vertically to cover every
    // non-blank text baseline (plus one line of margin each side); the
    // render is trimmed afterwards anyway. Tables are not touched — their
    // per-cell text clipping is real.
    private static byte[] UnclipEquationEmf(byte[] src)
    {
        if (src == null || src.Length < 88 || BitConverter.ToUInt32(src, 0) != 1)
        {
            return src;
        }
        byte[] emf = (byte[])src.Clone();
        int bTop = BitConverter.ToInt32(emf, 12), bBottom = BitConverter.ToInt32(emf, 20);
        int fTop = BitConverter.ToInt32(emf, 28), fBottom = BitConverter.ToInt32(emf, 36);
        int devY = BitConverter.ToInt32(emf, 76), mmY = BitConverter.ToInt32(emf, 84);
        if (devY <= 0 || mmY <= 0)
        {
            return src;
        }
        double hmmPerPx = mmY * 100.0 / devY;            // frame unit: 0.01 mm
        int lineH = Math.Max(1, (int)Math.Round((fBottom - fTop) / hmmPerPx));
        int minY = int.MaxValue, maxY = int.MinValue;

        int off = 0;
        while (off + 8 <= emf.Length)
        {
            uint type = BitConverter.ToUInt32(emf, off);
            int size = BitConverter.ToInt32(emf, off + 4);
            if (size < 8 || off + size > emf.Length || type == EMR_EOF)
            {
                break;
            }
            if ((type == EMR_EXTTEXTOUTW || type == EMR_EXTTEXTOUTA) && size >= 76)
            {
                // EMRTEXT at +36: ptlReference(x,y) nChars offString options rcl offDx
                int y = BitConverter.ToInt32(emf, off + 40);
                int nChars = BitConverter.ToInt32(emf, off + 44);
                int offString = BitConverter.ToInt32(emf, off + 48);
                uint options = BitConverter.ToUInt32(emf, off + 52);
                BitConverter.GetBytes(options & ~ETO_CLIPPED).CopyTo(emf, off + 52);
                if (!IsBlankText(emf, off + offString, nChars, type == EMR_EXTTEXTOUTW))
                {
                    minY = Math.Min(minY, y);
                    maxY = Math.Max(maxY, y);
                }
            }
            off += size;
        }
        if (minY == int.MaxValue)
        {
            return emf;
        }
        // baselines -> glyph extent: one line height above the top baseline
        // and below the bottom one is ample for ascenders/descenders. (Word's
        // equation EMFs are MM_TEXT with origin 0, so the logical baselines
        // are already in the header's device units.)
        int newTop = Math.Min(bTop, minY - lineH);
        int newBottom = Math.Max(bBottom, maxY + lineH);
        BitConverter.GetBytes(newTop).CopyTo(emf, 12);
        BitConverter.GetBytes(newBottom).CopyTo(emf, 20);
        BitConverter.GetBytes(fTop - (int)Math.Round((bTop - newTop) * hmmPerPx)).CopyTo(emf, 28);
        BitConverter.GetBytes(fBottom + (int)Math.Round((newBottom - bBottom) * hmmPerPx)).CopyTo(emf, 36);
        return emf;
    }

    private static bool IsBlankText(byte[] emf, int at, int nChars, bool wide)
    {
        int width = wide ? 2 : 1;
        if (nChars <= 0 || at < 0 || at + nChars * width > emf.Length)
        {
            return true;
        }
        for (int i = 0; i < nChars; i++)
        {
            int c = wide ? BitConverter.ToUInt16(emf, at + 2 * i) : emf[at + i];
            if (!char.IsWhiteSpace((char)c))
            {
                return false;
            }
        }
        return true;
    }

    // --- render + trim + save ----------------------------------------------

    private static void SaveTrimmedTiff(byte[] emfBytes, string outTif,
                                        out float wMm, out float hMm)
    {
        string tmp = Path.Combine(Path.GetTempPath(), "eqr_" + Guid.NewGuid().ToString("N") + ".emf");
        wMm = hMm = 0;
        try
        {
            File.WriteAllBytes(tmp, emfBytes);
            using (var mf = new Metafile(tmp))
            {
                // physical size of the rendered range (mm)
                float hres = mf.HorizontalResolution;
                float vres = mf.VerticalResolution;
                if (hres <= 0) hres = 96f;
                if (vres <= 0) vres = 96f;
                float origWmm = mf.Width / hres * 25.4f;
                float origHmm = mf.Height / vres * 25.4f;

                int w = Math.Max(1, (int)Math.Ceiling(origWmm / 25.4f * Dpi));
                int h = Math.Max(1, (int)Math.Ceiling(origHmm / 25.4f * Dpi));
                FitWithin(ref w, ref h,
                          MmToPx(MaxWidthMm), MmToPx(MaxHeightMm));

                using (var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                {
                    bmp.SetResolution(Dpi, Dpi);
                    using (var g = Graphics.FromImage(bmp))
                    {
                        g.Clear(Color.White);
                        g.SmoothingMode = SmoothingMode.HighQuality;
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.CompositingQuality = CompositingQuality.HighQuality;
                        g.DrawImage(mf, new Rectangle(0, 0, w, h));
                    }
                    Rectangle bounds = FindContentBounds(bmp);
                    if (bounds.IsEmpty)
                    {
                        throw new InvalidOperationException("empty render");
                    }
                    using (Bitmap cropped = bmp.Clone(bounds, bmp.PixelFormat))
                    using (Bitmap mono = ToMonochrome(cropped))
                    {
                        mono.SetResolution(Dpi, Dpi);
                        SaveGroup4(mono, outTif);
                        // display size = physical size scaled by trim ratio
                        wMm = origWmm * (float)bounds.Width / w;
                        hMm = origHmm * (float)bounds.Height / h;
                    }
                }
            }
        }
        finally
        {
            try { File.Delete(tmp); } catch { }
        }
    }

    private static Rectangle FindContentBounds(Bitmap bmp)
    {
        BitmapData data = bmp.LockBits(
            new Rectangle(0, 0, bmp.Width, bmp.Height),
            ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        int minX = bmp.Width, minY = bmp.Height, maxX = -1, maxY = -1;
        try
        {
            int stride = Math.Abs(data.Stride);
            byte[] buf = new byte[stride * bmp.Height];
            Marshal.Copy(data.Scan0, buf, 0, buf.Length);
            for (int y = 0; y < bmp.Height; y++)
            {
                int row = y * stride;
                for (int x = 0; x < bmp.Width; x++)
                {
                    int p = row + x * 4;
                    byte b = buf[p], g = buf[p + 1], r = buf[p + 2], a = buf[p + 3];
                    if (a > 5 && !(r > 250 && g > 250 && b > 250))
                    {
                        if (x < minX) minX = x;
                        if (y < minY) minY = y;
                        if (x > maxX) maxX = x;
                        if (y > maxY) maxY = y;
                    }
                }
            }
        }
        finally
        {
            bmp.UnlockBits(data);
        }
        if (maxX < minX || maxY < minY)
        {
            return Rectangle.Empty;
        }
        minX = Math.Max(0, minX - CropPaddingPx);
        minY = Math.Max(0, minY - CropPaddingPx);
        maxX = Math.Min(bmp.Width - 1, maxX + CropPaddingPx);
        maxY = Math.Min(bmp.Height - 1, maxY + CropPaddingPx);
        return Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
    }

    private static Bitmap ToMonochrome(Bitmap src)
    {
        var mono = new Bitmap(src.Width, src.Height, PixelFormat.Format1bppIndexed);
        ColorPalette pal = mono.Palette;
        pal.Entries[0] = Color.Black;
        pal.Entries[1] = Color.White;
        mono.Palette = pal;

        BitmapData sd = src.LockBits(new Rectangle(0, 0, src.Width, src.Height),
                                     ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        BitmapData md = mono.LockBits(new Rectangle(0, 0, mono.Width, mono.Height),
                                      ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed);
        try
        {
            int ss = Math.Abs(sd.Stride), ms = Math.Abs(md.Stride);
            byte[] sbuf = new byte[ss * src.Height];
            byte[] mbuf = new byte[ms * mono.Height];
            Marshal.Copy(sd.Scan0, sbuf, 0, sbuf.Length);
            for (int i = 0; i < mbuf.Length; i++) mbuf[i] = 0xFF;
            for (int y = 0; y < src.Height; y++)
            {
                int sr = y * ss, mr = y * ms;
                for (int x = 0; x < src.Width; x++)
                {
                    int p = sr + x * 4;
                    int lum = (sbuf[p + 2] * 299 + sbuf[p + 1] * 587 + sbuf[p] * 114) / 1000;
                    if (lum < 200)
                    {
                        mbuf[mr + (x >> 3)] &= (byte)~(0x80 >> (x & 7));
                    }
                }
            }
            Marshal.Copy(mbuf, 0, md.Scan0, mbuf.Length);
        }
        finally
        {
            src.UnlockBits(sd);
            mono.UnlockBits(md);
        }
        return mono;
    }

    private static void SaveGroup4(Image img, string path)
    {
        ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders()
            .FirstOrDefault(c => c.FormatID == ImageFormat.Tiff.Guid);
        using (var ep = new EncoderParameters(1))
        {
            ep.Param[0] = new EncoderParameter(
                System.Drawing.Imaging.Encoder.Compression,
                Convert.ToInt64(EncoderValue.CompressionCCITT4));
            img.Save(path, codec, ep);
        }
    }

    private static void FitWithin(ref int w, ref int h, int maxW, int maxH)
    {
        if (w <= maxW && h <= maxH) return;
        double s = Math.Min(maxW / (double)w, maxH / (double)h);
        w = Math.Max(1, (int)Math.Floor(w * s));
        h = Math.Max(1, (int)Math.Floor(h * s));
    }

    // Invariant culture: a comma-decimal locale would write "12,5" and break
    // docx2hlt's float() parse of the map files.
    private static string Mm(float mm)
    {
        return mm.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static int MmToPx(float mm)
    {
        return Math.Max(1, (int)Math.Ceiling(mm / 25.4f * Dpi));
    }

    private static bool ContainsDeletedRevision(dynamic range)
    {
        try
        {
            dynamic revs = range.Revisions;
            int n = revs.Count;
            for (int i = 1; i <= n; i++)
            {
                dynamic rev = revs[i];
                try
                {
                    int t = rev.Type;   // 2=delete,14=movedFrom,17=cellDeletion
                    if (t == 2 || t == 14 || t == 17) return true;
                }
                finally { Release((object)rev); }
            }
            Release((object)revs);
        }
        catch { }
        return false;
    }

    private static void Release(object o)
    {
        if (o != null && Marshal.IsComObject(o))
        {
            Marshal.ReleaseComObject(o);
        }
    }
}
