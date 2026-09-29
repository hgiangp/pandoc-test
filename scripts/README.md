# docx → markdown pipeline (pandoc)

pandoc silently drops drawings made with Word Shapes (canvas, group, autoshape, SmartArt,
chart). Some data sets also have two traits of their own (profile `review-markup`):
- **Red boxes** that reviewers drew to highlight content. They carry no information, but
  they split figures into pieces or produce extra images of empty red boxes.
- Figures and tables are placed together with their caption in a **one-cell layout
  table**. pandoc therefore outputs the whole block as a grid table, and the image or table
  appears boxed inside a cell.

The pipeline below handles these problems before and during the pandoc step.

```
input.docx
 ├─[1] preflight   Test-DocxDrawings.ps1        count the drawings pandoc would drop
 ├─[2] cleanup     cleanup\ (Python)            data cleanup by profile, can be switched off → input.clean.docx
 │                   ReviewerBoxes → UnwrapLayoutTables   (profile review-markup)
 ├─[3] convert     Convert-ShapesToPictures.ps1 render drawings to PNG with Word       → input.shapes.docx
 ├─[4] prep        cleanup\ + profiles\pandoc.json  work around pandoc limits, always on → input.pandoc.docx
 │                   ExpandSimpleFields
 ├─[5] pandoc      -t gfm --wrap=none --lua-filter=pandoc\figures.lua             → input.md + images\
 ├─[6] media       Convert-MediaToPng.ps1       convert leftover EMF/WMF to PNG, fix links
 ├─[7] publish     Publish-Output.ps1           md + only the images it links → input.out\
 └─[8] gate        Test-DocxDrawings.ps1        check that no drawing was lost
```

- The **cleanup** stage handles problems specific to the **data** and is switched per
  profile. The **convert** and **prep** stages work around **pandoc** limitations and apply
  to every document. With cleanup off, the pipeline runs normally with the other stages.
- Cleanup runs before convert, so that red boxes are not merged with real figures, and so
  that the PNG is inserted outside the table, right before the caption.
- Every stage writes a new docx. All intermediate files are kept in `input.work\` so they
  can be opened in Word for inspection. The input file is never modified.
- The deliverable is in its own folder `input.out\`: only the md file and the images it links.

### Principle: which stage handles what?

When a new case comes up, pick the stage by the **nature** of the problem:

| Nature of the problem | Handled in | Examples |
|---|---|---|
| **Structural noise in the source data**: things that are not content, caused by how the document was authored or reviewed | **Cleanup** (Python rules that edit the docx XML, switched per profile) | Reviewer red boxes, one-cell layout tables |
| **pandoc limitation**: the object must be **rendered** to be readable | **Convert** (uses Word) | Shape, canvas, SmartArt, chart, OLE, EMF |
| **pandoc limitation**: valid XML that pandoc reads wrongly | **Prep** (Python rules in `profiles\pandoc.json`, always on, runs after the last Word save) | `w:fldSimple`: pandoc drops the field text |
| **Output presentation**, depends on who reads the markdown | **Lua filter in pandoc** (works on the AST) | Image alt text, title and id, table format, table captions |
| Pure text edits on the `.md` file | A text post-processing step, **only when really needed** | Whitespace normalization |

Why:
- **Fixing the structure at the source lets pandoc understand the structure.** For
  example, once the layout table is unwrapped in the docx, pandoc pairs the image with its
  caption into a figure, attaches captions to tables and keeps bookmarks. Unwrapping later
  would mean rebuilding all of that by hand.
- At the docx level, Word information is still available (Caption style, SEQ fields, shape
  colors and geometry); in markdown it is gone.
- **Do not fix structure with regexes on the `.md` file.** Grid tables, escape characters
  and line breaks make that very fragile. Whatever must happen after pandoc is done on the
  AST with a Lua filter.
- Logic specific to one data set belongs in a cleanup profile, never in the shared stages
  (convert, pandoc).

## Installation (once per machine)

Full requirements (Windows, activated Word, PowerShell 5.1, which stage needs what) and the
end-user guide: see [README.md](../README.md) in the repository root.

| Component | Used by | Install |
|---|---|---|
| Windows + Word desktop 2013 or newer | convert | – |
| Windows PowerShell 5.1 | everything | Built into Windows |
| pandoc | pandoc | `winget install --id JohnMacFarlane.Pandoc` |
| [uv](https://docs.astral.sh/uv/) | cleanup and prep (prep always runs unless `-NoPandocPrep`) | `winget install --id astral-sh.uv -e`. **No separate Python install needed**: on the first run uv downloads Python 3.12 (pinned in `cleanup\.python-version`) and the exact library versions from `cleanup\uv.lock` (needs internet the first time) |

## Quick start: the .bat files

Double-click, or drag and drop a docx file onto a `.bat` file:

| File | What it does |
|---|---|
| `run-review-markup.bat input.docx` | **Full pipeline with profile `review-markup`** (removes red boxes) |
| `run-all.bat input.docx [options]` | Full pipeline. **No** cleanup by default |
| `cleanup.bat input.docx --profile review-markup [--mode Report]` | Cleanup stage only (`uv run ... docx-cleanup`) |
| `convert.bat input.docx [-DryRun]` | Convert stage only |

```bat
rem First run on new data: report only, remove nothing, to check what is detected
run-all.bat input.docx -Profile review-markup -CleanupMode Report

rem Real run with profile review-markup
run-review-markup.bat input.docx

rem Compare with no cleanup / no pandoc filter
run-all.bat input.docx -NoCleanup
run-all.bat input.docx -Profile review-markup -NoFigureFilter
```

The results are in two folders next to `input.docx`:

```
input.out\          deliverable, rebuilt on every run
  input.md
  images\           only the images the md links (PNG), links like images/image19.png
input.work\         intermediate files for inspection
```

| File in `input.work\` | Content |
|---|---|
| `input.clean.docx` | After cleanup (only when the profile enables a rule) |
| `input.cleanup-manifest.csv` | What the cleanup rules found and did |
| `input.shapes.docx`, `input.shapes\` | After convert: PNG, EMF and `manifest.csv` |
| `input.pandoc.docx`, `input.pandoc-manifest.csv` | After prep: the file pandoc actually reads, and the list of changes |
| `input.md`, `images\media\` | Raw pandoc output (after the media stage) |
| `input.pipeline.log` | Log of the whole run |

Change the locations with `-OutputDir` and `-WorkDir`. Both folders are named after the
input file, so several docx files in one folder do not overwrite each other's images.

The PNGs in `input.shapes\` need not be copied to the output. The convert stage **embeds**
them in `input.shapes.docx`, so pandoc extracts them into `images\media\` like any other
picture, named `imageN.png`. To find which drawing an image in the md came from, look at its
`data-shape="S001"` attribute and search `input.shapes\manifest.csv`.

Exit code: `0` = PASS, `3` = finished with issues, `1` = error. Set `set NOPAUSE=1` to run
without pausing, e.g. in CI. The logic lives in `Invoke-Pipeline.ps1`; the `.bat` files are
only thin launchers.

## Data cleanup

### Profiles

Each profile is a JSON file in `profiles\`. Every rule in a profile has a `mode`:

| mode | Behavior |
|---|---|
| `Off` | Does not run |
| `Report` | Detects and writes to the manifest, does not modify the docx |
| `Apply` | Detects and fixes (removes) |

- `default.json`: every rule `Off`. The cleanup stage is skipped.
- `pandoc.json`: **not a cleanup profile**. It is the fixed profile of the prep stage,
  always applied right before pandoc. Do not pass it to `-Profile`.
- `review-markup.json`: `ReviewerBoxes` then `UnwrapLayoutTables`, both `Apply`. **Rules run
  in the order they appear in the profile.** ReviewerBoxes must run first, so that red boxes
  inside a layout table are removed before the table content is moved out.
- `-CleanupMode Report` (in `run-all.bat`) or `--mode Report` (in `cleanup.bat`) forces
  every **enabled** rule to run in Report mode. It never enables a rule that is `Off`.
- A misspelled parameter name in a profile stops the run with a clear error, instead of
  silently falling back to the default value.

### Rule `ReviewerBoxes`: reviewer red boxes

A shape is treated as a highlight box when **all** of these hold:

| Condition | Parameter | Default |
|---|---|---|
| Rectangle or rounded rectangle | `geometries` | `["rect", "roundRect"]` |
| No fill, or a nearly transparent fill | `maxFillOpacity` | `0.1` |
| Red outline: close to a listed color, or a clear red hue | `colors`, `colorTolerance`, `matchRedHue` | `FF0000, C00000`; `60`; `true` |
| No text inside | – | – |
| No connector (arrow) attached. Boxes in a diagram usually have arrows attached; highlight boxes do not | – | – |

| Confidence | When | In Apply mode |
|---|---|---|
| `high` | All conditions hold, stand-alone shape | Removed |
| `medium` | All conditions hold, but inside a group or canvas | Removed only with `includeInGroups: true` |
| `low` | Red or reddish outline, and exactly one condition fails | Reported only, for tuning |

Colors are resolved when set directly (`srgbClr`), from the theme (`schemeClr`, including
lighter/darker variants), from the shape style (`lnRef`), and for VML shapes. The
`mc:Fallback` part (a VML copy of the same shape) is not counted twice.

On removal:

- **Stand-alone shape**: the whole `mc:AlternateContent` block is removed, including the
  Fallback. The text under the box is kept.
- **Shape in a group/canvas**: only the child shape is removed. The group's Fallback is
  dropped because it no longer matches the new content; Word regenerates it in the convert
  stage. If the group has no shapes left, the group is removed too.

### Rule `UnwrapLayoutTables`: one-cell tables wrapping a figure/table and its caption

A table is treated as a **layout table** when all of these hold:

| Condition | Parameter | Default |
|---|---|---|
| One column only (every row has exactly one cell) | – | – |
| At most `maxRows` rows, e.g. figure in row 1 and caption in row 2 | `maxRows` | `4` |
| Holds a drawing (drawing, picture, OLE) or a nested table | – | – |
| Has a caption: Caption style (or a style based on it), or a paragraph with a `SEQ` field | – | – |

The Caption style is matched by **name** (`caption`), not by id. Localized Word
(Japanese, Vietnamese…) stores ids like `a3`, but the name of a built-in style is always
English.

| Confidence | When | In Apply mode |
|---|---|---|
| `high` | All conditions hold | Unwrapped: the cell content (paragraphs, images, nested tables, bookmarks) moves out in order |
| `medium` | Has a figure/nested table but no caption | Unwrapped only with `unwrapWithoutCaption: true` |
| `low` | More than `maxRows` rows, or text only (e.g. a Note box) | Reported only |

Tables with more than one column, e.g. two figures side by side, are **never** unwrapped.

`_Ref…` bookmarks (targets of links like "Fig. 3‑2", "Table 1‑2") move with the caption,
so cross references keep working after unwrapping.

### Reading `input.cleanup-manifest.csv`

| Column | Meaning |
|---|---|
| `Action` | `removed`, `unwrapped`, `reported`, or `reported (set … to …)` when a parameter must be enabled for the rule to act |
| `Confidence` / `Reason` | Confidence level; for `low`, the failed condition (`fails: fill`, `fails: text`, …) |
| `Part`, `Paragraph` | Document part (document, header, …) and index of the paragraph that holds the shape |
| `Location` | Text of that paragraph, or of the next non-empty one (usually the caption), to find the shape in Word |
| `ShapeId`, `ShapeName`, `Container` | Id and name in Word; `top`, `group` or `canvas` |
| `Features` | Measured features (JSON): geometry, line color, line width, dash style, fill opacity, … |

**Tuning**:

- A `low` row with `fails: color` that really is a highlight box: add its `Features.lineColor` to `colors`.
- A `medium` row that really is a highlight box: set `includeInGroups: true`.
- A shape removed by mistake: switch the rule to `Report` and send the matching manifest row.

### Adding a rule

1. Create `cleanup\docx_cleanup\rules\<rule_name>.py`, subclass `Rule` (`model.py`) and
   implement `detect()` (must not modify the docx) and `apply()`.
2. Register the class in `registry.py`.
3. Add its configuration to the right profile, following "Principle: which stage handles
   what?": `profiles\<name>.json` for cleanup of one data set, `profiles\pandoc.json` for a
   pandoc workaround that applies to every document.
4. Write tests in `cleanup\tests\`. `conftest.py` has helpers that build sample docx files
   (shape, group, canvas, VML).

The safe removal operations (AlternateContent/Fallback handling, pruning empty runs,
validating the file after writing) live in `DocxPackage` (`package.py`). New rules should
call them instead of editing the XML directly.

### Managing the package with uv

`cleanup\` is a uv project:

| File | Role |
|---|---|
| `pyproject.toml` | Dependencies (`lxml`), dev group (`pytest`) and the `docx-cleanup` command |
| `uv.lock` | Pins the exact version of every library, **must be committed** |
| `.python-version` | Python version uv uses (3.12). The code is compatible with Python ≥ 3.9 |

The pipeline calls `uv run --project cleanup --locked --no-dev docx-cleanup ...`:
- `--locked`: fail if `uv.lock` does not match `pyproject.toml`, instead of silently installing other versions.
- `--no-dev`: do not install pytest on the processing machine.

```bash
cd scripts/cleanup
uv sync                                  # create .venv (including the dev group)
uv run pytest                            # run the tests, no Windows needed
PANDOC=/path/to/pandoc uv run pytest     # add end-to-end tests with real pandoc (skipped without pandoc)
uv run --isolated --python 3.9 pytest    # check Python 3.9 compatibility
uv add <package>                         # add a dependency (updates uv.lock)
uv lock --upgrade                        # upgrade the libraries
```

## Prep stage: pandoc workarounds (`profiles\pandoc.json`)

This stage runs **after the last Word save** (the convert stage) and right before pandoc,
for every document. Switch it off for comparison with `-NoPandocPrep`. Changes are written
to `input.pandoc-manifest.csv`.

### Rule `ExpandSimpleFields`

Word stores a field (caption number, cross reference…) in one of two forms: a **complex
field** (`w:fldChar` begin/separate/end) or the **compact** `w:fldSimple`. pandoc (checked
3.1–3.11) **drops the text** of every `w:fldSimple`:

| In Word | Field | pandoc without prep | pandoc with prep |
|---|---|---|---|
| `Table 1‑3 Example` | `SEQ`, `STYLEREF` as `fldSimple` | `Example`: label and number lost | `Table 1‑3 Example` |
| `Fig. 4‑1 Example of` | `STYLEREF` as `fldSimple` | `Fig. ‑1 Example of` | `Fig. 4‑1 Example of` |
| `shown in Table 1‑3` | `REF` as `fldSimple` | `shown in`: reference lost | `shown in [Table 1‑3](#_Ref…)` (pandoc 3.11; older versions give plain text) |

The rule rewrites each `w:fldSimple` as a complex field with the same instruction, result
and formatting. The document looks exactly the same in Word.

In the sample documents checked, the `fldSimple` fields (STYLEREF, SEQ) were already in the
original file, all of them caption numbers. They are not created by Word when saving in the
convert stage.

## Media stage: `Convert-MediaToPng.ps1`

Markdown viewers and browsers cannot display EMF/WMF. Most such images are already turned
into PNG by the convert stage (inline pictures, OLE objects, pictures inside a
group/canvas), but some still get through: **floating EMF/WMF pictures**, or runs with
`-KeepMetafiles`.

This stage is the last safety net: it scans `images\`, renders every `.emf`/`.wmf` file to
PNG with GDI+ (the same renderer as the convert stage, `lib\ShapeRaster.cs`), and rewrites
**only the file name** in the markdown, touching no other text. The original file is kept
unless `-RemoveOriginals` is given. Switch the stage off with `-NoMediaConvert`.

These images are not trimmed, unlike drawings, because the margins of a picture can be part
of its content.

## pandoc stage: output format and `pandoc\figures.lua`

### Output format: GFM

The pipeline outputs **GFM** (GitHub Flavored Markdown, `-t gfm`), not Pandoc Markdown
(`-t markdown`), to balance human readers and AI:

| Criterion | Pandoc Markdown | **GFM (default)** |
|---|---|---|
| Human readers (GitHub, VS Code, GitLab) | ❌ Grid tables `+---+`, `{…}` and `: caption` show as raw text | ✅ Rendered correctly |
| AI, simple tables | Grid table: lots of padding spaces, costs tokens | Pipe table: most compact |
| AI, complex tables (multi-paragraph cells, caption, no header row) | Grid table: a cell's content is split over several lines | HTML `<table>` + `<caption>`: clear cell boundaries |
| AI, figures | Image + `{alt=…}` attribute | `<figure>` with the image (with alt text) and `<figcaption>` |

GFM pipe tables are limited: one line per cell, a header row is required, no caption. pandoc
outputs any table that breaks these rules as HTML. For a simple table to become a pipe table,
its first row must be marked "Repeat as header row" in Word.

**Requirement for downstream systems (RAG/LLM):** keep the HTML tags in the markdown, and
never split a chunk inside a `<table>` or `<figure>`.

To output Pandoc Markdown for comparison: `run-all.bat input.docx -OutputFormat markdown`.

Unwrapping layout tables (cleanup) is still **needed** with GFM. Without it, figures and data
tables become HTML tables nested in another HTML table: readers still see a box around them,
and an AI cannot tell which figure or table a caption belongs to.

### `figures.lua`

A Lua filter that runs inside pandoc and tidies the images created by the convert stage:

| Task | Before | After |
|---|---|---|
| Title `shape2png:S001` (link to the manifest) becomes an attribute | `"shape2png:S001"` | `data-shape="S001"` |
| Alt text with the text of the drawing: kept for LLMs, without the `\|` that collides with table syntax | `Drawing converted to image. Text: A \| B` | `Text in figure: A; B` |
| Bookmark in a figure caption moves to the figure | `<span id="_Ref1" class="anchor">` inside the caption | `<figure id="_Ref1">` |
| **Table caption above the table** (`gfm` only). Pipe tables have no caption syntax, so pandoc puts the caption below; HTML tables use `<caption>`. The filter produces one consistent form, in the same order as in Word | Caption below a pipe table, or in `<caption>` | Caption paragraph (with its bookmark) before the table |

When an image is directly followed by a caption paragraph, pandoc combines them into a
figure. With GFM:

```html
<figure id="_Ref100000001">
<img src="images/media/image19.png" style="width:2.5in;height:2.38in" data-shape="S001" alt="Text in figure: Idle; Running; Stopped" />
<figcaption><p>Fig. 3‑2 State transitions</p></figcaption>
</figure>
```

A table caption is a paragraph right before the table, starting with the bookmark
`<span id="_Ref…" class="anchor"></span>`. Cross-reference links
(`[Fig. 3‑2](#_Ref100000001)`, `[Table 1‑2](#_Ref100000002)`) point to the figure `id` or to
the table caption bookmark.

`--wrap=none`: never break image or link syntax over several lines.

Tested with pandoc 3.1.11 and 3.11, both `gfm` and `markdown`. The filter needs pandoc ≥ 3.0.

## Running the convert stage by hand

```powershell
cd <scripts folder>

# 0. Preflight: count the objects pandoc would drop (no Word needed)
powershell -ExecutionPolicy Bypass -File .\Test-DocxDrawings.ps1 -Docx .\input.docx

# 1. Dry run: list only, change nothing -> input.shapes\manifest.csv
powershell -ExecutionPolicy Bypass -File .\Convert-ShapesToPictures.ps1 -InputPath .\input.docx -DryRun

# 2. Convert -> input.shapes.docx + input.shapes\S001.png, S001.emf, ..., manifest.csv
powershell -ExecutionPolicy Bypass -File .\Convert-ShapesToPictures.ps1 -InputPath .\input.docx

# 3. Run pandoc on the converted file
pandoc -f docx -t gfm --wrap=none --extract-media=./images --lua-filter=.\pandoc\figures.lua .\input.shapes.docx -o output.md

# 4. Gate: fail if objects pandoc would drop remain; warn (no fail) if captions > images
powershell -ExecutionPolicy Bypass -File .\Test-DocxDrawings.ps1 -Docx .\input.shapes.docx -Markdown .\output.md
```

**Do not use the clipboard** while the convert stage runs: the script copies and pastes to
capture floating shapes.

## Convert stage options

| Option | Default | Meaning |
|---|---|---|
| `-OutputPath` | `<input>.shapes.docx` | Output docx |
| `-ImageDir` | `<input>.shapes\` | Folder for PNG, EMF (for debugging) and `manifest.csv` |
| `-Dpi` | 200 | PNG resolution |
| `-IncludeTextBoxes` | off | Also convert stand-alone text boxes. Kept by default so no text is lost |
| `-KeepMetafiles` | off | Keep OLE objects (Visio...) and EMF/WMF pictures as they are. By default they become PNG because markdown viewers cannot show EMF |
| `-NoCluster` | off | Do not merge loose shapes of the same figure |
| `-NoTrim` | off | Do not crop white margins |
| `-DryRun` | off | List only |
| `-Visible` | off | Show the Word window for debugging |
| `-UnlinkShapeFields` | off | Turn fields inside shapes into plain text before rendering. Use only if images still show `Error! Reference source not found.` |
| `-NoHeadingNumbers` | off | Do not write Word heading numbers ("7.5") into the heading text |
| `-NoPageInfo` | off | Leave the manifest `Page` column empty. Reading a page number makes Word repaginate, so this speeds up long documents |

Exit code: `0` = OK, `2` = some objects failed to convert (see the `Error` column of the manifest).

**Speed.** The convert stage prints the time of each part (floating shapes, inline shapes,
heading numbers, save) to find the slow spot. The script turns off background pagination,
spelling and grammar checks and screen updating (unless `-Visible`), and restores these
options before quitting Word. On long documents, `-NoPageInfo` usually helps the most.

## How the convert stage works

1. **Floating shapes** (`Document.Shapes`):
   - Loose shapes of the same figure are **grouped** before rendering. Two shapes belong to
     the same figure when they share a `Fig./Figure/Hình` caption and page, or the same
     anchor paragraph. This keeps a figure from being cut into several small images.
   - The drawing is captured by trying, in order:
     1. `ConvertToInlineShape()`, which only works for pictures/OLE.
     2. Copy, then read the EMF from the clipboard through Win32.
     3. Paste Special as EMF into a temporary document.
   - The PNG is inserted as its own paragraph (Normal style) right before the anchor
     paragraph, then the original shape is deleted.
2. **Inline shapes** (`Document.InlineShapes`): classified by their XML
   (`Range.WordOpenXML`), not by COM type, because Word does not type inline DrawingML
   shapes/groups/canvases reliably. The drawing is captured with `Range.EnhMetaFileBits`
   and replaced in place.
3. **Fields are locked before converting.** When Word converts a shape or renders its EMF,
   it updates the fields inside it, e.g. a cross reference "refer to 7.5". At that moment
   the shape is detached from the text, the bookmark is out of reach, and the field turns
   into `Error! Reference source not found.`, which is then drawn into the image. Right after
   opening the document, the script sets `Locked = True` on **every field of the document**
   (every story, including shapes) and turns off `Options.UpdateFieldsAtPrint`. Word never
   updates a locked field, so its current value is kept. If the error still shows, use
   `-UnlinkShapeFields` to turn the fields inside shapes (`wdTextFrameStory`) into plain text.
4. **Heading numbers are written into the headings.** pandoc ignores Word's automatic
   numbering, so `7.5 System settings` becomes `System settings`, and every "refer to 7.5"
   loses its target. Word knows each heading's number (`ListFormat.ListString`), so the
   script writes it as plain text at the start of the heading, then turns off the automatic
   numbering of that paragraph so the intermediate file does not show the number twice.
   Only paragraphs with outline level 1–9 are affected; numbered lists in the body are left
   alone, since pandoc already turns them into markdown lists. Headings that already start
   with their number are skipped. Switch off with `-NoHeadingNumbers`.
   Locking all fields (item 3) is also needed here: `STYLEREF` fields in captions take the
   chapter number from this numbering, so Word must not update them.
5. The EMF is rendered to PNG with GDI+. The PNG carries the DPI so Word keeps the printed size.
6. **Text inside the drawing** (state boxes, arrow labels...) is written into the image alt
   text as `Drawing converted to image. Text: Idle | Running ...`; `figures.lua` turns it
   into `Text in figure: Idle; Running ...`. The information stays searchable and usable
   for RAG/LLMs.

## Known limitations

Improvements and known issues that are not handled yet are tracked in [BACKLOG.md](BACKLOG.md).

- If a caption sits in a floating text box next to the figure, that text box can be merged
  into the cluster and rendered into the image. The caption is then only in the alt text.
  Check clusters in the manifest.
- Macros in the input file never run (`AutomationSecurity = ForceDisable`).
- Machines locked down with AppLocker/WDAC (PowerShell in Constrained Language Mode) cannot
  run `Add-Type`/COM. The script reports this clearly at the start.
- Documents with password-protected Restrict Editing: remove the protection in Word first.
- Only the main document body is processed. Shapes in headers, footers, footnotes and
  comments are not handled yet.
- Grouping loose shapes relies on captions. Shapes of a figure without a caption that are
  anchored in several paragraphs can come out as several images. Check the manifest; if
  needed, group them by hand in Word (Select, then Group).
- Word cannot group a drawing canvas with other shapes. The cluster is then split and each
  shape converted separately (manifest shows `Action=split`).
- **Floating** EMF/WMF pictures are not converted in the convert stage; the media stage
  (`Convert-MediaToPng.ps1`) turns them into PNG after pandoc.
- Fonts rendered by GDI+ can look slightly different from Word. For a pixel-exact result,
  export a PDF from Word and crop.

## Publish stage: `Publish-Output.ps1`

Builds `input.out\` from the md in `input.work\`:

- Copies only the files in `images\media\` that the md links. Original EMF/WMF files (that
  already have a PNG) and images left from earlier runs are skipped. The log lists unused files.
- Flattens the folder: `images/media/image19.png` becomes `images/image19.png`. Only those
  exact link paths are rewritten; no other text is touched.
- Warns (pipeline exit code `3`) when the md links a missing file, or still links EMF/WMF.
- The folder is built as `input.out.tmp\` and then swapped in for `input.out\`. A failed run
  leaves the previous output as it was.
- The output folder is deleted only when it is empty or holds the `.pipeline-output` marker
  file created by the script. Pointing `-OutputDir` at a folder with other data is refused.

## After a test run, send back

- `input.work\input.pipeline.log`
- `input.work\input.cleanup-manifest.csv`, especially red boxes removed by mistake or missed
- `input.work\input.shapes\manifest.csv`
- The console log, especially `FAILED` lines
- The output of `Test-DocxDrawings.ps1` before and after convert
- 1–2 PNG images, e.g. of a diagram with a lot of text, to judge the render quality

## Troubleshooting

| Symptom | Fix |
|---|---|
| `uv not found in PATH` | `winget install --id astral-sh.uv -e`, then reopen the terminal; or run with `-NoCleanup -NoPandocPrep` |
| `The lockfile ... needs to be updated` | `uv.lock` does not match `pyproject.toml`: run `uv lock` in `cleanup\` and commit |
| First cleanup run fails to download | uv needs internet to download Python and libraries the first time. Behind a proxy, set `HTTPS_PROXY` |
| `running scripts is disabled` | Run through `powershell -ExecutionPolicy Bypass -File ...` |
| A file downloaded from the internet does not open | `Unblock-File .\input.docx` |
| Hang or COM error | Close all Word instances (`Get-Process WINWORD \| Stop-Process`), then run again with `-Visible` to see what Word shows |
| Many `FAILED` lines with `clipboard`/`Paste Special` | Do not use the clipboard during the run; turn off clipboard manager apps |
| Images show `Error! Reference source not found.` | Word updated a field in the shape while rendering. The script locks fields; if it still happens, run again with `-UnlinkShapeFields`. The manifest `Text` column shows the correct text |
| PNG cropped too much or with extra margins | Run with `-NoTrim` and compare with the matching `.emf` file |
