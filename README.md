# docx → markdown (pandoc)

Converts Word documents (`.docx`) to markdown (GFM) with images, without losing drawings
made with Word Shapes (canvas, group, autoshape, SmartArt, chart, Visio/OLE, EMF/WMF). Plain
pandoc drops these drawings without any warning.

This file is the **user guide**. Technical details of each stage, the cleanup rules and how
to extend them are in [scripts/README.md](scripts/README.md). Known issues are tracked in
[scripts/BACKLOG.md](scripts/BACKLOG.md).

## System requirements

### Processing machine (required)

| Component | Version | Notes |
|---|---|---|
| **Windows** | 10 or 11 | The convert stage drives Word through COM and uses the Windows clipboard, so it cannot run on macOS/Linux |
| **Microsoft Word desktop** (Office) | 2013 or newer, 32 or 64 bit | Must be installed on the machine (Microsoft 365 Apps, Office 2016/2019/2021/2024…) and **activated**. Word for the web and Word for Mac do not work. Unactivated Word opens documents read-only and cannot save |
| **Windows PowerShell** | 5.1 | Built into Windows 10/11. The `.bat` files call `powershell.exe` (5.1), not PowerShell 7 (`pwsh`) |
| **pandoc** | ≥ 3.0, 3.11+ recommended | `winget install --id JohnMacFarlane.Pandoc`. Before 3.11, cross references come out as plain text instead of links |
| **uv** | recent | `winget install --id astral-sh.uv -e`. **No Python install needed**: on the first run uv downloads Python 3.12 and the pinned libraries (needs internet the first time; behind a proxy, set `HTTPS_PROXY`) |

After installing pandoc and uv, **open a new terminal** so that PATH is updated. Check:

```bat
pandoc --version
uv --version
```

Other conditions on the processing machine:

- PowerShell must run in **FullLanguage** mode. Machines locked down with AppLocker/WDAC
  (Constrained Language Mode) cannot run the convert stage; the script reports this at the start.
- Word must not be blocked by first-run dialogs (sign-in, default file format, activation).
  Open Word by hand once and close every dialog before running.
- **Do not use the clipboard** (copy/paste) while the convert stage runs, and turn off
  clipboard manager apps. The script captures floating shapes through the clipboard, so its
  current content is cleared.

### Which stage needs what

| Stage | Windows + Word | pandoc | uv |
|---|:---:|:---:|:---:|
| [1] preflight, [8] gate (`Test-DocxDrawings.ps1`) | Windows (PowerShell), no Word | | |
| [2] cleanup, [4] prep (Python) | | | ✔ (also runs on macOS/Linux) |
| [3] convert (`Convert-ShapesToPictures.ps1`) | ✔ | | |
| [5] pandoc | | ✔ | |
| [6] media (`Convert-MediaToPng.ps1`) | Windows (GDI+), no Word | | |
| [7] publish (`Publish-Output.ps1`) | Windows (PowerShell) | | |

On macOS/Linux, only the Python part can be developed and tested (`cd scripts/cleanup && uv run pytest`).

## Usage

### 1. Run

Drag and drop a `.docx` file onto one of the `.bat` files in `scripts\`, or run from `cmd`:

```bat
cd scripts

rem Documents with reviewer red boxes and one-cell layout tables: profile review-markup
run-review-markup.bat D:\data\input.docx

rem Regular documents: no cleanup
run-all.bat D:\data\input.docx
```

For a new data set, first run in **report-only** mode, check `input.cleanup-manifest.csv`,
then do the real run:

```bat
run-all.bat D:\data\input.docx -Profile review-markup -CleanupMode Report
```

A long document can take a few minutes in the convert stage. Word runs hidden; there is no
need to open or touch it.

### 2. Get the result

The result is next to the input file:

```
D:\data\
  input.docx          original file, never modified
  input.out\          DELIVERABLE (rebuilt on every run)
    input.md
    images\           only the PNG images that input.md uses
  input.work\         intermediate files, log and manifests for inspection
```

Only `input.out\` is needed. Do not save your own files there: the folder is deleted and
rebuilt on every run.

### 3. Read the run result

The last console line tells the result:

| Exit code | Last line | Meaning |
|---|---|---|
| `0` | `[PASS] ...` | Every drawing was converted to an image |
| `3` | `[CHECK] Finished with issues ...` | Output was produced, but needs checking: a drawing failed to convert (`FAILED`), a broken image link, or EMF/WMF left |
| `1` | `[ERROR] ...` | Error, no new output. Read the error message |

With exit code `3`, look at the files in `input.work\`:

| File | What to look for |
|---|---|
| `input.pipeline.log` | The full log of the run |
| `input.shapes\manifest.csv` | Every drawing: converted or not, and the error (`Action`, `Error` columns) |
| `input.cleanup-manifest.csv` | Red boxes removed, layout tables unwrapped, and cases only reported |

### Common options

Add them after the file name, e.g. `run-all.bat input.docx -Dpi 300 -NoPageInfo`:

| Option | Effect |
|---|---|
| `-Profile review-markup` | Remove reviewer red boxes and unwrap layout tables (`run-review-markup.bat` does this) |
| `-CleanupMode Report` | Cleanup only reports, changes nothing |
| `-Dpi 300` | Sharper images (default 200) |
| `-NoPageInfo` | Faster on long documents (no page numbers in the manifest) |
| `-IncludeTextBoxes` | Also turn stand-alone text boxes into images (kept as text by default) |
| `-UnlinkShapeFields` | Use when images show `Error! Reference source not found.` |
| `-NoHeadingNumbers` | Do not write heading numbers ("7.5") into the heading text |
| `-OutputDir <folder>` | Where to write the result (default `input.out` next to the input) |
| `-OutputFormat markdown` | Pandoc Markdown instead of GFM |

Full list: run `run-all.bat` without arguments, or see [scripts/README.md](scripts/README.md).

Unattended runs (from another script or CI): `set NOPAUSE=1` before calling the `.bat` file.

## Quick troubleshooting

| Symptom | Fix |
|---|---|
| `pandoc not found` / `uv not found` | Install as listed above, then open a new terminal |
| `running scripts is disabled` | Run through the `.bat` files (they pass `-ExecutionPolicy Bypass`) |
| `needs FullLanguage` | The machine is locked down by AppLocker/WDAC. Ask IT to allow it, or use another machine |
| `Document is protected` | Remove Restrict Editing in Word and run again |
| Hang, or COM error | Close all Word instances (`taskkill /IM WINWORD.EXE /F`), run again with `-Visible` to see which dialog Word shows |
| Many `FAILED` lines about the clipboard | Do not copy/paste during the run; turn off clipboard manager apps |
| A file downloaded from the internet does not open | Right-click → Properties → Unblock, or `Unblock-File .\input.docx` |

Other cases: see "Troubleshooting" in [scripts/README.md](scripts/README.md).
