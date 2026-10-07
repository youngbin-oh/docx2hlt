# docx2hlt — DOCX → NKEditor .hlt converter

Converts Microsoft Word `.docx` files into `.hlt` files that open in
NKEditor (지식재산처 전자문서 작성 S/W, KIPO e-filing editor).

Two output paths, chosen automatically:

- **Patent / structured** — a `.docx` that contains patent-section
  headings (【발명의 명칭】, 【기술분야】, 【배경기술】, 【발명의 내용】,
  【청구항 …】, 【요약】, 【도면 …】 …) is converted into a **real
  structured 특허명세서**: the structure pane (구조창) populates with
  발명의 설명 → 청구범위 → 청구항 1…N, and **제출파일(생성/변환)** runs
  on it without the "문서 종류가 잘못되었습니다" error.
- **Flat** — any other `.docx` is converted as a plain NKEditor
  document (no element tree), preserving text/formatting/images.

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

## Requirements

- Windows with **NKEditor** installed (default `C:\KipoNet\NKEditor`).
  - Patent documents need the full NKEditor app — it builds the
    element tree and writes the `.hlt` itself via Save-As automation.
  - Non-patent (flat) conversion only needs `hwplt32.dll` from that
    directory. If installed elsewhere, pass the dir as a 3rd arg to
    `TerConv.exe` or edit `NK_DIR` in `docx2hlt.py`.
- Python 3 (stdlib) **plus Pillow** (`pip install Pillow`) — used to
  normalize embedded images.
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
  Equation images are keyed by their index in Word's `doc.OMaths`
  (`eqNNNN.tif`), and `docx2hlt.py` numbers the `m:oMath` elements the
  same way (`main_story_omaths`: document order, main story, inside
  tables too, tracked deletions included, text boxes excluded). So an
  equation inside a rendered table, a tracked-deleted equation, or one
  that fails to render can't shift later equations onto the wrong
  image. If Word's equation count differs from the XML's, no equation
  image is used: every equation becomes linear text and a warning is
  written, because a wrong equation is worse than linear text.
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
   binary `.hlt` under `Data\Hlt`, which is copied to the destination.
   If the structured path fails for any reason it falls back to the
   flat path below.

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
