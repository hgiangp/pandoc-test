# Backlog

Known issues and improvements that are **not handled yet**, to be agreed on and done later.
Priority: **P1** = directly affects the output content; **P2** = quality or usability;
**P3** = technical or operational.

## Known issues

| ID | Priority | Issue | Facts so far | Next step |
|---|---|---|---|---|
| – | – | (no open known issues) | | |

## Open decisions

| ID | Priority | Question | Options |
|---|---|---|---|
| D-02 | P2 | **Text inside figures** for downstream steps (LLM…) | Currently in the `alt` attribute. Alternatives: a text block right below the image; an LLM that reads the image and writes a description; converting state diagrams to Mermaid |

## Improvements by stage

### Cleanup

| ID | Priority | Item |
|---|---|---|
| C-01 | P1 | **Tune ReviewerBoxes on real data** (colors, opacity, `includeInGroups`) based on the `low`/`medium` rows in `cleanup-manifest.csv` |
| C-02 | P2 | UnwrapLayoutTables: **multi-column** layout tables (two figures side by side) are not handled. They could be split into consecutive figures |
| C-03 | P3 | UnwrapLayoutTables: one-cell tables with **text only** (Note boxes) are only reported. They could become blockquotes |
| C-04 | P3 | ReviewerBoxes: a `Mark` mode that highlights the boxed content (`<mark>`) instead of only removing the box. Only if the boxed parts carry meaning |
| C-05 | P3 | Red boxes **already drawn into screenshots** (raster): out of scope. Images with many red pixels could be flagged for manual review |

### Convert (Word)

| ID | Priority | Item |
|---|---|---|
| V-01 | P2 | Shapes in headers, footers, footnotes and comments are not handled |
| V-03 | P2 | Loose shapes are grouped into one figure by caption: figures without a caption anchored in several paragraphs can be split. A caption in a floating text box can be rendered into the image |
| V-04 | P3 | The ` \| ` separator is reused in the alt text (`Build-AltText`), and `figures.lua` relies on it to produce `; `. Without the filter, `\|` still appears |
| V-05 | P3 | Performance: open and close Word once for a batch of documents instead of once per file |
| V-06 | P2 | Convert stage performance: timings per part are in place. Next: drop the per-inline-shape `Range.WordOpenXML` call (classify once from the XML), and replace the fixed 300 ms `Start-Sleep` after each copy with an adaptive wait |
| V-07 | P3 | Parallelism: several documents at once (one Word process each; the clipboard needs a lock since it is shared by the whole machine), and parallel PNG rendering in the convert stage |

### pandoc / output

| ID | Priority | Item |
|---|---|---|
| P-01 | P3 | pandoc renames images to `imageN.png`, so only the `data-shape` attribute links an image to the manifest. A filter could name files by Id (`S001.png`) |
| P-02 | P2 | **Fewer tokens**: GFM keeps presentation-only attributes (`style="width:…"`, `<colgroup>`, table `style`). A filter could drop them |
| P-03 | P2 | **RAG requirement**: keep HTML tags, never split a chunk inside `<table>` or `<figure>`. Needs checking during integration |
| P-04 | P3 | Tables without a header row in Word become HTML tables even when they are very simple. The first row could be treated as a header when it is bold or shaded |
| P-06 | P3 | With pandoc older than 3.11, `REF` cross references become plain text instead of links. Standardize on pandoc ≥ 3.11 on the processing machine |
| P-05 | P3 | GitHub prefixes `id`s with `user-content-`, so cross-reference links (`#_Ref…`) may not jump correctly on GitHub. VS Code works |

### Checks & operations

| ID | Priority | Item |
|---|---|---|
| Q-01 | P2 | The caption vs. image count in the gate is only a heuristic. It could cover tables too and count on the pandoc AST instead of with regexes |
| Q-02 | P3 | Move `Test-DocxDrawings.ps1` into the Python package to share the XML classifier with the cleanup stage |
| Q-03 | P3 | CI: `uv run pytest`, PSScriptAnalyzer for the `.ps1` scripts, and a Windows integration test with an anonymized sample file |
| Q-04 | P3 | Batch runs over many docx files, with a summary report |

## Done

| Issue | Solution |
|---|---|
| pandoc silently drops drawings made in Word | `Convert-ShapesToPictures.ps1` |
| Reviewer red boxes split figures and produce images of red boxes | Cleanup rule `ReviewerBoxes` |
| Figures and tables boxed in a one-cell table (grid table) | Cleanup rule `UnwrapLayoutTables` |
| Alt text and title break the image syntax (`\|`, line breaks, `shape2png`) | `pandoc\figures.lua` + `--wrap=none` |
| Word automatic heading numbers lost (`7.5 System settings` → `System settings`), breaking every "refer to 7.5". pandoc does not read `numbering.xml` for headings, and `--number-sections` does not work with GFM | The convert stage writes `ListFormat.ListString` into the heading, after locking all fields |
| EMF/WMF images reached the markdown and could not be displayed, e.g. floating pictures (was V-02) | Media stage `Convert-MediaToPng.ps1`: render to PNG and fix the links |
| Images rendered `Error! Reference source not found.` instead of the cross reference inside the drawing (e.g. "refer to 7.5"). Cause: Word updates fields while converting/rendering a shape, when the bookmark is out of reach | The convert stage locks the fields before converting and turns off `UpdateFieldsAtPrint`. Fallback: `-UnlinkShapeFields` |
| Caption numbers, the "Table" label and cross references lost (was B-01). Cause: `w:fldSimple` fields (STYLEREF, SEQ) already in the original document; pandoc drops the text of `fldSimple` | Prep stage, rule `ExpandSimpleFields` (`profiles\pandoc.json`): rewrite as complex fields |
| Table captions moved below the table (GFM pipe tables have no caption) | `figures.lua`: caption becomes a paragraph above the table |
| Correct markdown that does not render (grid tables, `{…}`, `: caption`). Was D-01 | Output `gfm`: simple tables become pipe tables, complex tables and figures become HTML. Table caption bookmarks (was B-02) become valid HTML `id`s |
| Wrong path in `.bat` files after `shift` | Save the script folder before `shift` |
