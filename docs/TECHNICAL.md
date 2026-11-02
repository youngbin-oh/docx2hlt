# docx2hlt — DOCX → NKEditor .hlt converter

Converts Microsoft Word `.docx` files into `.hlt` files that open in
NKEditor (지식재산처 전자문서 작성 S/W, KIPO e-filing editor).

The document type is detected first and printed (`document type: ...`);
each kind has its own output path:

- **국내 특허 명세서 (spec)** — a `.docx` that contains patent-section
  headings (【발명의 명칭】, 【기술분야】, 【배경기술】, 【발명의 내용】,
  【청구항 …】, 【요약】, 【도면 …】 …) is converted into a **real
  structured 특허명세서**: the structure pane (구조창) populates with
  발명의 설명 → 청구범위 → 청구항 1…N, and **제출파일(생성/변환)** runs
  on it without the "문서 종류가 잘못되었습니다" error.
  (KIPO XML root `<KIPO><PatentCAFDOC>`.)
- **국문 PCT 명세서 (pct)** — PCT명세서(국제출원) 한글서식. The same
  parse, emitted as WIPO ST.36 `<application-body lang="ko">`
  (application-body-v1-6.dtd): 【청구의 범위】, 【기술적 해결방법】,
  【발명의 실시를 위한 형태】, 【요약서】 without 【요약】/【대표도】.
  NKEditor builds that import as document kind **03** (header byte 0x20;
  PctPat.frm, PCT번역문 제출서 = 국내진입), and 제출파일 변환 then fails
  with E-259 "이 문서는 특허법 시행규칙에 따라 작성되지 않았습니다". A new
  국제출원 한글서식 document (PctPatHD.frm) saves as kind **0x48** with an
  otherwise identical header, so after the save docx2hlt sets that byte
  03 -> 48 (`set_pct_intl_kind`); verified: a patched Korean PCT file
  passes 제출파일 변환. Any other value is left alone with a warning.
  Claim lines are separate sibling `<claim-text>` elements (this import
  drops `<br/>` inside claim-text). `<disclosure>` and
  `<mode-for-invention>` replace the domestic `<summary-of-invention>` /
  `<description-of-embodiments>`; each `<figure>` holds one image (extra
  images are dropped with a warning); a docx 【대표도】 is ignored.
  (`<KIPO><PCTPatentCAFDOC>` does not work - NKEditor builds a domestic
  spec from it; `PCTPatentDOC` comes out as raw text.)
- **영문 PCT 명세서 (pct_en)** — PCT명세서(국제출원) 영문서식. Same
  ST.36 build with `<application-body lang="en">`, from English
  headings (【DESCRIPTION】, 【Invention Title】, 【Technical Field】,
  【Background Art】, 【Disclosure】, 【Technical Problem】, 【Technical
  Solution】, 【Advantageous Effects】, 【Description of Drawings】,
  【Best Mode】, 【Mode for Invention】, 【Industrial Applicability】,
  【CLAIMS】, 【Claim N】, 【ABSTRACT】, 【DRAWINGS】, 【Figure N】/【FIGURE
  1A】; case and spacing are ignored, see `HEAD2XML_EN`). NKEditor saves
  this import as kind **0x49** directly - the kind of filed
  English PCT .hlt files - so no header patch is needed (a 03 would be
  set to 49). Text right under 【ABSTRACT】 is the abstract. A claim set
  without 【Claim N】 headings ("1. A method ..." paragraphs under
  【CLAIMS】 in a filed spec) is split on the numbers, which must run 1, 2, 3...
  (a "2." inside a claim body does not start claim 2). Figure numbers keep
  their letter in upper case (1A).
- **의견서 (opinion)** — emitted as `<KIPO><OpinionDOC><Opinion>`, the
  vocabulary of an NKEditor-exported 의견서: `<P indent align>`, `<B>`,
  `<U>`, `<I>`, `<SB>`/`<SP>`, `<BR/>`, images as `<EMI>` (JPEG,
  `opnNNNNN.jpg`, mm), tables as **native OASIS tables** (column widths,
  horizontal/vertical merges) instead of images. The docx 【의견내용】
  line is dropped (the Opinion element is that item). Indent follows the
  docx2hml rule: 【…】 lines, bare [청구항 N], 인용발명N captions 0,
  other paragraphs 14. Header kind byte 07.
- **보정서 (amend)** — only detected: 보정서 .hlt files are made from the
  .fin, so an amendment docx goes flat (its 【청구항 N】 lines would
  otherwise pass it off as a spec).
- **Flat** — any other `.docx` is converted as a plain NKEditor
  document (no element tree), preserving text/formatting/images.

Kind detection (`detect_kind`), first match wins:

1. a 【의견내용】 marker -> opinion; a line starting with 【보정대상 항목】 /
   【보정방법】 / 【보정내용】 -> amend;
2. filename contains 의견서 -> opinion, 보정서 -> amend (not in a foreign
   case path such as `SP00004-IN0` / `_CN0_`: those are associate
   letters);
3. with patent-section headings: a `WO` country token anywhere in the path
   (`SP00001-WO0`, `...-WO_SP00002`) -> pct; a `KR0`-style token in a
   folder name (`P0000001-KR0`) -> spec (KR0 working drafts sometimes carry
   PCT headings but were filed as domestic specs); PCT-only headings
   (【청구의 범위】, 【기술적 해결방법】, 【발명의 실시를 위한 형태】,
   【발명의 실시를 위한 최선의 형태】, 【산업상 이용가능성】) -> pct; else spec;
4. otherwise flat.

English ST.36 headings (at least 3, outnumbering the Korean ones) mean
pct_en before step 3, whatever the case folder says.

The bare word "PCT" is not a signal (national-stage refs carry it too).
Override with `--kind=spec|pct|pct_en|opinion|flat` (`--flat` =
`--kind=flat`).

Headings the form has no element for (e.g. 【배경 및 요약】, 【Summary of
the Drawings】) are not items: the line stays as plain `[...]` text inside
the section before it, and the run reports each one as a warning. Body
text right under a container heading (【청구범위】, 【도면】, 【요약서】,
【CLAIMS】, ...) that is not a numbered claim has no XML slot and is also
reported. 【수학식 N】 / 【표 N】 / 【Equation N】 / 【Table N】 lines are
plain items and are not reported.

## Usage

```
docx2hlt.bat input.docx output.hlt
```

or directly:

```
python docx2hlt.py input.docx output.hlt
```

Batch (sequential — the structured path drives the NKEditor GUI):

```
python docx2hlt.py a.docx b.docx c.docx    # each -> <same name>.hlt beside it
```

Right-click also accepts multiple selected `.docx` files (one window, files
converted one after another) once `register_ctx.ps1` has been re-run.

## Right-click (Explorer) usage

Run `register_ctx.ps1` once (PowerShell) to add a **"HLT 파일로 변환"**
verb to the `.docx` right-click menu. After that, right-click any
`.docx` in Explorer → **HLT 파일로 변환** — it writes `<same name>.hlt`
in the same folder. Remove it with `unregister_ctx.ps1`.

## Drag-and-drop window (docx2hlt_gui.exe)

Run `docx2hlt_gui.exe` (or a shortcut to it) and drop `.docx` files or
folders (top-level `.docx` only) on the window, or on the exe's icon.
They are converted one after another, with the same live log the console
shows. No right-click registration is needed, so sharing this whole
folder is enough to install it on another PC.

- Output: next to each `.docx` by default; **저장 위치 → 지정한 폴더**
  sends the `.hlt` and its review files to one folder instead. Two
  same-named `.docx` from different folders become `name.hlt` and
  `name (2).hlt` there. The choice is kept per user in
  `%APPDATA%\docx2hlt\gui.ini`.
- 문서 종류: 자동 판별 by default; the other entries pass `--kind=...`.
- Python: `python\python.exe` in this folder when bundled, otherwise
  `python.exe` on PATH (the WindowsApps store stub is skipped), then the
  `py` launcher. The window logs which one it uses and whether Pillow is
  there. `docx2hlt.bat` and the right-click `docx2hlt_ctx.cmd` also use
  the bundled one first.
- On start it sets NKEditor's per-user `RUNASINVOKER` compatibility flag
  (HKCU) when missing, and warns when NKEditor or Word is absent.
- Run it with normal rights: an elevated window cannot receive drops
  from Explorer.
- Rebuild with:
  `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /nologo
   /target:winexe /codepage:65001 /out:docx2hlt_gui.exe
   /r:System.Windows.Forms.dll /r:System.Drawing.dll docx2hlt_gui.cs`

## Requirements

- Windows with **NKEditor** installed (default `C:\KipoNet\NKEditor`).
  - Patent documents need the full NKEditor app — it builds the
    element tree and writes the `.hlt` itself via Save-As automation.
  - Non-patent (flat) conversion only needs `hwplt32.dll` from that
    directory. If installed elsewhere, pass the dir as a 3rd arg to
    `TerConv.exe` or edit `NK_DIR` in `docx2hlt.py`.
- Python 3 (stdlib) **plus Pillow** (`pip install Pillow`) — used to
  normalize embedded images. Or no install at all: the `python\` folder
  holds the official embeddable Python 3.12.10 (python.org
  `python-3.12.10-embed-amd64.zip`, PSF-signed) with Pillow 12.3.0 in
  `python\Lib\site-packages`. An embeddable Python does not put the
  script's folder on `sys.path`, so `python\python312._pth` lists `..`
  (this folder) for `import nkwin`; keep that line when upgrading it.
- `TerConv.exe` — prebuilt 32-bit helper in this folder (flat path).
- `nkwin.py` — ctypes/Win32 driver that imports the generated KIPO XML
  into NKEditor and triggers Save-As (structured path).
- **Microsoft Word** — needed when the docx contains OMML equations
  (`m:oMath`) or tables (`w:tbl`): `EqRender.exe` renders each
  equation **and each table** to a trimmed 300dpi G4 TIFF via
  `Range.CopyAsPicture` → clipboard EMF → rasterize → crop. Tables keep
  their full visual layout (grid, borders, merged cells) as an inline
  image at their original document position — the alternative
  (flattening cells into paragraphs) loses the table structure.
  Because rendered tables/equations are pixels, their text is NOT
  subject to KIPO charset sanitization; unsupported characters inside
  them are harmless. Rendering is **retried a few times** before any
  linear-text fallback (Word COM fails transiently when a running
  WINWORD sits on a modal prompt — close Word's dialogs and rerun if
  you see the WARNING in the charset report). Without Word/rendering,
  equations degrade to linear text and tables fall back to cell-flattened
  paragraphs.
  Word's `CopyAsPicture` of an equation is only one text line (18pt)
  tall and clips every glyph to that line box (`ETO_CLIPPED`), which
  cut off fraction denominators and numerator tops. EqRender clears
  that flag on the equation EMF's text records and grows the picture
  frame to cover all baselines before rasterizing (tables untouched).
  Equation images are keyed by their index in Word's `doc.OMaths`
  (`eqNNNN.tif`), and `docx2hlt.py` numbers the `m:oMath` elements the
  same way (`main_story_omaths`: document order, main story, inside
  tables too, tracked deletions included, text boxes excluded). So an
  equation inside a rendered table, a tracked-deleted equation, or one
  that fails to render can't shift later equations onto the wrong
  image. If Word's equation count differs from the XML's, no equation
  image is used: every equation becomes linear text and a warning is
  written, because a wrong equation is worse than linear text.
  Equations take their EMF from `Range.EnhMetaFileBits` (no clipboard)
  when Word provides it: over a remote-desktop session each clipboard
  copy cost 5-10 s (148 equations: 21 min; 370 equations: would
  be ~50 min), EnhMetaFileBits ~0.2 s with the same drawing and mm size.
  Tables stay on the clipboard, because EnhMetaFileBits lays out a table
  wider than the text column narrower than Word shows it (one
  table: 161 mm instead of its 180 mm grid). The EqRender timeout scales with
  the object count (15 s each, at least 600 s), and a timeout is not
  retried: retries only timed out again and each left a hidden WINWORD.
  Rebuild with:
  `csc -target:exe -out:EqRender.exe -r:System.Drawing.dll
   -r:System.Windows.Forms.dll -r:Microsoft.CSharp.dll EqRender.cs`
- `DocCmp.exe` — Word COM `CompareDocuments` driver for the
  `.compare.docx` review artifact. Rebuild with:
  `csc -target:exe -out:DocCmp.exe DocCmp.cs`

## How it works — structured (patent) path

The `.hlt` patent structure is built by NKEditor's own import engine,
not synthesized by hand:

1. `docx2hlt.py` parses the docx and detects patent sections
   (`_is_patent_doc`). For a patent doc it calls `build_patent_xml`,
   which maps each section to the KIPO **PatentCAFDOC** schema and
   writes a `PatentCAFDOC` XML document plus `patNNNNN.jpg` images:
   `description` → invention-title / technical-field / background-art /
   summary-of-invention{tech-problem, tech-solution, advantageous-effects}
   / description-of-drawings / description-of-embodiments, then `claims`,
   `abstract` (+abstract-figure figref), `drawings` (`<figure num>` +
   `<img he wi file img-format>`).
   - Paragraph `<p>` elements get sequential `num` attributes; claims
     and figures are numbered independently.
   - Sections that have no content are omitted — `tech-problem` /
     `advantageous-effects` etc. are `(heading*, p+)+` and an *empty*
     element makes NKEditor's submission export fail with
     "XML 변환시 오류가 발생했습니다".
   - `abstract-figure`'s `<figref num>` is resolved against the figure
     numbers actually emitted (docx "도 3" that split into 3a/3b
     resolves to a real figure; a dangling reference also breaks export).
   - `【수학식 N】`/`【표 N】`/`【화학식 N】` etc. are NOT emitted as KIPO
     식별항목 — their content (linear text) can't satisfy the PIC/TABLE
     type check (S-009) and styled markers trip E-260. They are written
     as plain `[수학식 N]` text and the equation/table itself is an
     inline `<img>` inside the body `<p>`. Mid-paragraph `【…】` are
     likewise rewritten to `[…]` (E-204/205/206 forbid the brackets
     outside real identification items).
   - Images keep their position inside paragraphs (`<img>` spliced at
     the original text offset) — only images under a `【도 N】` heading
     become `<drawings><figure>` elements. Images are emitted as-is for
     TIFF/JPEG; near-bilevel art (equations, diagrams) becomes Group-4
     TIFF, other content JPEG q92; `he`/`wi` are clamped to the
     165×222mm KIPO page box (E-218).
2. `nkwin.py` launches `NKEditor.exe <that xml>`. NKEditor imports the
   PatentCAFDOC XML and builds the native structured document tree.
   The driver waits for the structure pane's `SysTreeView32` to
   populate (PID-scoped, dismisses recovery/file-not-found prompts).
3. It then triggers Save-As deterministically by posting
   `WM_COMMAND 0xE104` (ID_FILE_SAVE_AS) to the frame — no screen
   coordinates, so it survives remote-desktop occlusion and never needs
   input focus — and clicks the dialog's 저장 button (`BM_CLICK`).
   The default save type is already `*.HLT`, so NKEditor writes a real
   binary `.hlt` into the dialog's last-used folder (which can be
   another case's folder), and it is moved from there to the destination.
   If the structured path fails for any reason it falls back to the
   flat path below.

NKEditor 6.3 is not single-instance: every launch is a separate process
with its own window, even when the user already has NKEditor open. The
driver therefore starts its own editors minimized without activation
(`SW_SHOWMINNOACTIVE`), so they do not pop up over the user's window,
and it never sends anything to an editor it did not start. If an editor
the user already had open shows the output `.hlt`, that window still
holds the pre-conversion document; the run says so at the end (and the
right-click wrapper waits for Enter), because saving there would
overwrite the new `.hlt`.

## How it works — flat (non-patent) path

The `.hlt` binary format is proprietary and interlinked, so instead of
synthesizing it, the converter drives the same engine NKEditor uses:

1. `docx2hlt.py` parses `word/document.xml` — paragraphs, runs,
   bold/italic/underline, and `<w:drawing>`/`<w:pict>` image anchors
   (resolved via `document.xml.rels` to `word/media/*`) — and emits a
   compact binary *job file*. Each embedded image is normalized to an
   uncompressed TIFF (Pillow) because the editor's save path embeds
   TIFF reliably but can't decode every PNG variant.
   Text is encoded as CP949; paragraphs longer than the editor's
   299-byte insert limit are split at character boundaries into
   soft-broken continuation lines.
2. `TerConv.exe` (x86, .NET Framework — no external deps) loads
   `hwplt32.dll`, creates a `Ter32Class` edit control on a pumping
   thread, then:
   - Pass 1 — `TerInsertLine` per text line (`\x15` = paragraph
     terminator). For each image paragraph it inserts an *empty*
     reserved line and records its index.
   - Pass 2 — `SetTerCursorPos` to each reserved line, then
     `TerInsertPictureFile` + `TerInsertObjectId` to drop the image's
     `0x5f` object marker there. Anchoring must happen after all text
     is in place: the marker binds to the moving append cursor, so if
     it were anchored mid-pass every later `TerInsertLine` would push
     it toward the end — the classic "all images pile up on the last
     page in reverse order" symptom.
   - post-pass: each image object gets a data index (`+0x4`), a staged
     `HLT%N.HIF` file under `%TEMP%` in `[u32 len][image bytes]` form
     (the save's `GetTempPath()+HLT%N.HIF` reader feeds it to
     `IG_load_mem`), and the pixel/display dimension fields that
     `IG_info_get` fails to fill — the save then embeds the image data
     into the `.hlt`. (`GetTempPath` is redirected to the per-job
     staging dir via the process `TMP`/`TEMP` env vars.)
   - Each line's paragraph spec (indent/alignment/spacing) is turned
     into a paragraph id via `TerCreateParaId` and passed to
     `TerInsertLine`. (`TerCreateParaId` writes `-1` into unused id
     slots; those are reset to 0 to match the default paragraph record —
     a `-1` in the style slot makes the line's text not paint.)
   - `SelectTerText` + `SetTerCharStyle` (bold/italic/underline/
     strike/sub/super), `SetTerPointSize`, `SetTerColor` per run
   - `SaveTerFile` → `.hlt`

## Job file format

```
u32 imgDirLen ; u8 imgDir[imgDirLen]   // dir for HLT%N.HIF staging
u32 lineCount
per line:
  u8 kind                            // 0=text, 1=image
  kind 0: u32 textLen ; u8 text[textLen]     // CP949
          u8  terminated                     // 1=paragraph end, 0=soft break
          u32 fmtCount ; per fmt: u32 colStart, u32 colEnd, u32 styleMask,
                                   i32 sizeHalfPt, u32 colorBGR
  kind 1: u32 pathLen ; u8 path[pathLen]     // utf8 source image path
          i32 dispW_pt ; i32 dispH_pt        // display size (points)
          i32 pxW ; i32 pxH                  // pixel size
  both:   u8 hasPara                        // 1 = para spec follows
          [ i32 leftIndent, i32 rightIndent, i32 firstIndent,   // twips
            u32 paraFlags,                  // alignment bits
            i32 spaceBefore, i32 spaceAfter, i32 spaceBetween ] // twips
```

styleMask: 1=bold, 2=italic, 4=underline, 8=strikethrough,
0x10=superscript, 0x20=subscript. Columns are CP949 byte offsets.
paraFlags alignment group (mutually exclusive): 1=center, 2=right,
0x400=justify, 0x800=left/default.

## Preserved

- Paragraphs, Korean/English text, tabs→spaces
- Bold, italic, underline, strikethrough (run-level + style level)
- **Subscript / superscript** (`w:vertAlign`) — TE char-style bits
- **Indentation** — left, first-line, and hanging (`w:ind`, twips)
- **Alignment** — center, right, justify (`w:jc`)
- **Paragraph spacing** — space before/after (`w:spacing`)
- Font size (`w:sz`, half-points → points) and text color (`w:color`);
  sizes resolve through `w:basedOn` style chains and `docDefaults`
- **Embedded images/drawings** (PNG, JPEG, TIFF, GIF, BMP — normalized
  to TIFF, anchored at their document position, embedded into the .hlt,
  rendered at the DOCX-specified display size via `TerSetPictSize`)
- Tables are flattened to plain paragraphs (cell images are kept)
- Long paragraphs (auto-split, render continuously)
- **Tracked changes** — the accepted (final) view is converted: insertions
  are kept; deleted text, images and equations (`w:del`, `w:moveFrom`)
  are dropped
- Hyperlink, field and content-control text is read once. Text-box text
  is read once, inline. `mc:AlternateContent` follows the OOXML rule:
  use a DrawingML Choice, otherwise the Fallback (e.g. the `w:t` behind a
  `w16se:symEx` emoji symbol such as →)

## Character set handling (미지원 문자)

KIPO's editor can only represent the glyph set published on
`patent.go.kr` (KS X 1001 Hangul + 특수문자 + ASCII — mirrored in
`kipo_chars.txt`, ~3,400 chars). Anything outside that set cannot
survive a CP949-encoded `.hlt`.

- Every text run is passed through `sanitize_text()` before encoding:
  unsupported chars are replaced by the closest supported substitute
  (`CHAR_SUBS` in `docx2hlt.py`: `‖`→`∥`, `≈`→`≒`, `—`→`―`, `•`→`·`,
  `°`→`˚`, `µ`→`μ`, circled numerals →`(N)`, zero-width chars dropped,
  …). Unmappable chars are dropped rather than silently becoming `?`.
- Whenever a substitution/drop happens, a review file
  `<name>.charset.txt` **and `<name>.charset.docx`** (a Word table:
  코드포인트 | 원본 | 치환 | 문맥) is written next to the output.
  Conversion warnings (e.g. equation-render failure) appear at the top.
  Stale reports from earlier runs are removed when a conversion is
  clean.
- Equation images are unaffected (they're pixels, not text).

## Review artifacts (automatic)

Every conversion additionally produces, next to the `.hlt`:

- `<name>.rtf` — the generated `.hlt` re-opened in NKEditor and saved
  as RTF. This is *what the editor actually ended up holding*.
- `<name>.compare.docx` — `Word.CompareDocuments(원본.docx, rtf)`:
  a tracked-changes document. Substituted/dropped characters and any
  formatting loss appear as revisions — replaces the manual
  "save as rtf → 문서 비교" step. Requires Word for the compare step
  only; the `.rtf` is always produced when the structured path ran.
  Tables the conversion turned into images (spec/PCT tables become TIFF)
  would otherwise vanish from the compare without a trace: Word's
  compare (table compare off) drops them, and with table compare on it
  puts the image inside the deleted table's first cell, untracked. So
  DocCmp copies each such source table back in after the nearest
  non-empty paragraph before it (e.g. 【표 1】) and deletes it with
  tracking on, and re-inserts the image paragraph(s) right after it as a
  tracked insertion. Tables that are still tables in the rtf (opinion
  tables) are left to Word's own compare. An anchor DocCmp cannot find
  is reported as a review issue.

The export opens a byte-identical copy of the `.hlt` under a short temp
path: NKEditor 6.3 (2026.10.22.01) crashes with heap corruption
(0xc0000374) while saving RTF when the opened `.hlt` path is long
(reproduced at 146+ CP949 bytes, e.g. a case folder's `Sent` path; never
at 44-136). If NKEditor still dies during the export, it is retried once. A
review-file failure never fails the conversion: the reasons are printed
at the end of the run, and with `--pause-on-review-error` (the
right-click wrapper passes it) the window waits for Enter instead of
closing on its own. (An exception in the review step - an `.rtf` still
locked by the editor - used to be caught by the structured path's
handler and sent the finished `.hlt` down the flat path, which
overwrote it; the review now runs outside that handler.) Both NKEditor
saves wait until the written file stops growing before closing the
editor, and the close loop answers the 'XML 변환' prompt for up to 30 s:
a large RTF export (an English PCT spec) was still being written when the editor
was closed, and the prompt came after the old 8 s loop had given up,
leaving the editor open with the file locked.

Imported text also carries:

- Word list labels (bullets and numbers from `numbering.xml`, e.g.
  "- ", "○ ", "2. ", "a) ") as text at the start of the paragraph, the
  way the filed .hlt files have them; Symbol/Wingdings bullets map to
  •/■/○ etc. (`BULLET_MAP`). Paragraph-number lists whose level text
  has brackets ("[%1]", "[00%1]" -> [0001]) are skipped - NKEditor
  numbers paragraphs itself.
- Embedded OLE objects (`w:object`, e.g. MathType equations) as their
  preview image, sized from the VML shape style; WMF/EMF previews are
  rasterized at 600 dpi (Pillow's default 72 dpi gives a 26x18 px blot).
  They used to be dropped without a trace.

A sub/superscript run that is also bold, italic or underlined loses the
bold/italic/underline (the script itself stays): 제출파일 변환 rejects
the combination with E-001/E-002 "문자 스타일 중 BOLD, Italic,
UnderLine과 아래첨자를 같이 사용 할 수 없습니다" (e.g. italic
N<sub>CSI-PORT</sub>). Each such run is listed in one warning.

While it runs, the console shows each step with a ticking elapsed time,
and the Word picture step counts the equations/tables done so far
("그림 변환 중: 수식 354/788, 표 0/3 (5분 12초 경과)"); a spec with
hundreds of equations takes minutes there and used to look hung.

Review workflow: open `.compare.docx` → accept what's fine, fix the
docx where a substitution is wrong → reconvert. The `.charset.docx`
table (when present) is the checklist of risky substitutions.

## Limitations

- Embedded images are re-rendered by the editor at a fixed internal
  resolution, so each takes ~2 MB of uncompressed data — output files
  are larger than hand-authored NKEditor files.
- Documents with very many distinct embedded images (roughly >120,
  e.g. specs that inline every equation as an image) may drop some —
  a render-resource ceiling inside `hwplt32.dll`. Ordinary specs with
  tens of drawing figures are unaffected.
- NKEditor may show a "possibly incorrectly-saved file" prompt on open;
  choosing 예/Yes opens the document normally.
- Table grid structure is not reproduced (cells become paragraphs).
- Typefaces/fonts, headers/footers, footnotes, page setup, margins,
  lists/numbering, and tab stops are not carried over.
- The structure pane is populated only for patent-sectioned docs via
  the structured path above; plain documents are still flat.
- The structured path drives the real NKEditor GUI (it must be
  installed and licensed); a headless/locked session may not let it
  open windows. Flat conversion needs no GUI.

## Rebuilding TerConv.exe

```
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /platform:x86 TerConv.cs
```

Must be 32-bit — `hwplt32.dll` is x86. The `Framework` (not
`Framework64`) compiler with `/platform:x86` produces the right PE.
