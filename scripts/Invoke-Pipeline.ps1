<#
.SYNOPSIS
  docx -> markdown pipeline: preflight -> data cleanup -> shapes to PNG -> pandoc prep -> pandoc -> publish -> gate.

.DESCRIPTION
  The deliverable goes to its own folder next to the input (-OutputDir), rebuilt on every run:

    input.out\input.md               final markdown
    input.out\images\                only the images it links (PNG), links "images/<file>"

  Each stage reads a .docx and writes a new one; every intermediate file is kept in the work
  folder (-WorkDir) so it can be opened in Word:

    input.work\input.clean.docx                after data cleanup (only when the profile enables a rule)
    input.work\input.cleanup-manifest.csv      what the cleanup rules found / removed
    input.work\input.shapes.docx               after rasterizing drawings
    input.work\input.shapes\                   rendered PNG/EMF + manifest.csv
    input.work\input.pandoc.docx               after pandoc compatibility fixes (profiles\pandoc.json)
    input.work\input.pandoc-manifest.csv       what those fixes changed
    input.work\input.md, images\media\         raw pandoc output (EMF/WMF converted to PNG)
    input.work\input.pipeline.log              transcript of this run

  The rendered drawings in input.shapes\ are embedded in input.shapes.docx, so pandoc
  extracts them into images\media\ like any other picture; the output needs no copy of them.
  File names inside the work folder keep the input name: Word cannot open two documents
  with the same name at once.

  Data cleanup is controlled by a profile in profiles\<name>.json. The default profile
  enables no rule, so the stage is skipped. The pandoc prep stage always runs (unless
  -NoPandocPrep): it applies profiles\pandoc.json, fixes for pandoc limitations that
  every document needs (e.g. w:fldSimple fields whose text pandoc drops). Both stages use
  the cleanup package, a uv project (cleanup\pyproject.toml + uv.lock); uv installs the
  pinned Python and dependencies on first use.

  Exit code: 0 = pass, 3 = finished with issues (see output), 1 = error.

.EXAMPLE
  .\Invoke-Pipeline.ps1 -InputPath .\input.docx                         # no cleanup
  .\Invoke-Pipeline.ps1 -InputPath .\input.docx -Profile ns             # NS cleanup rules
  .\Invoke-Pipeline.ps1 -InputPath .\input.docx -Profile ns -CleanupMode Report   # only report
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputPath,
    # Cleanup profile name (profiles\<name>.json). Not named $Profile: that is a PowerShell automatic variable
    [Alias('Profile')][string]$CleanupProfile = 'default',
    # Skip the data cleanup stage regardless of the profile
    [switch]$NoCleanup,
    # Override the mode of every enabled rule: Report = only list findings, Apply = modify
    [ValidateSet('', 'Report', 'Apply')][string]$CleanupMode = '',
    # Skip the pandoc compatibility fixes (for comparison); then uv is not needed without cleanup
    [switch]$NoPandocPrep,
    # Skip converting EMF/WMF files that pandoc extracted into PNG
    [switch]$NoMediaConvert,
    # Run pandoc without pandoc\figures.lua (for comparison)
    [switch]$NoFigureFilter,
    # gfm: renders on GitHub/VS Code, complex tables and figures as HTML (default)
    # markdown: Pandoc Markdown (grid tables, {attributes}), only pandoc-aware tools render it
    [ValidateSet('gfm', 'markdown')][string]$OutputFormat = 'gfm',
    # Final markdown + images; default: <input folder>\<name>.out (replaced on every run)
    [string]$OutputDir,
    # Intermediate files; default: <input folder>\<name>.work
    [string]$WorkDir,
    # Passed through to Convert-ShapesToPictures.ps1
    [int]$Dpi = 200,
    [switch]$IncludeTextBoxes,
    [switch]$KeepMetafiles,
    [switch]$NoCluster,
    [switch]$NoTrim,
    [switch]$Visible,
    [switch]$UnlinkShapeFields,
    [switch]$NoHeadingNumbers,
    [switch]$NoPageInfo
)

$InputPath = (Resolve-Path -LiteralPath $InputPath).Path
$inputDir = Split-Path $InputPath -Parent
$name = [IO.Path]::GetFileNameWithoutExtension($InputPath)
function Resolve-FullPath([string]$p) {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p).TrimEnd('\', '/')
}
$workDir = if ($WorkDir) { Resolve-FullPath $WorkDir } else { Join-Path $inputDir "$name.work" }
$outDir = if ($OutputDir) { Resolve-FullPath $OutputDir } else { Join-Path $inputDir "$name.out" }
# the pandoc stage deletes images\ in the work folder, so it must not be a folder of the user's
if ($workDir -ieq $inputDir.TrimEnd('\', '/')) {
    Write-Host "[ERROR] WorkDir must differ from the input folder: $workDir" -ForegroundColor Red
    exit 1
}
New-Item -ItemType Directory -Force -Path $workDir | Out-Null
$cleanDocx = Join-Path $workDir "$name.clean.docx"
$cleanManifest = Join-Path $workDir "$name.cleanup-manifest.csv"
$shapesDocx = Join-Path $workDir "$name.shapes.docx"
$shapesDir = Join-Path $workDir "$name.shapes"
$prepDocx = Join-Path $workDir "$name.pandoc.docx"
$prepManifest = Join-Path $workDir "$name.pandoc-manifest.csv"
$mdName = "$name.md"
$mediaDir = Join-Path $workDir 'images'
$logPath = Join-Path $workDir "$name.pipeline.log"

function Write-Stage([string]$text) {
    Write-Host ''
    Write-Host "===== $text =====" -ForegroundColor Cyan
}

# Runs a PowerShell script in-process and returns its exit code (0 if it did not call exit)
function Invoke-Script([string]$path, [hashtable]$arguments) {
    $global:LASTEXITCODE = 0
    try {
        & $path @arguments | Out-Host
    } catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        return 1
    }
    return [int]$global:LASTEXITCODE
}

# Runs the cleanup package (docx in -> docx out) with a profile; throws on failure
function Invoke-DocxCleanup([string]$in, [string]$out, [string]$profileName, [string]$manifest, [string]$mode) {
    # --locked: fail instead of silently re-resolving if uv.lock is out of date
    # --no-dev: do not install test tools on the processing machine
    $uvArgs = @('run', '--project', (Join-Path $PSScriptRoot 'cleanup'), '--locked', '--no-dev',
                'docx-cleanup', $in, '-o', $out, '--profile', $profileName, '--manifest', $manifest)
    if ($mode) { $uvArgs += @('--mode', $mode) }
    & uv @uvArgs | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "docx-cleanup ($profileName) failed with exit code $LASTEXITCODE." }
}

function Test-CleanupEnabled([string]$profileName) {
    $path = Join-Path $PSScriptRoot "profiles\$profileName.json"
    if (-not (Test-Path -LiteralPath $path)) { throw "Profile not found: $path" }
    $cfg = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $cfg.rules) { return $false }
    foreach ($p in $cfg.rules.PSObject.Properties) {
        if ($p.Value.mode -and $p.Value.mode -ne 'Off') { return $true }
    }
    return $false
}

$exitCode = 1
try { Start-Transcript -Path $logPath -Force | Out-Null } catch {}
try {
    Write-Host "Input:   $InputPath"
    Write-Host "Work:    $workDir"
    Write-Host "Output:  $outDir"
    Write-Host "Profile: $(if ($NoCleanup) { '(cleanup disabled)' } else { $CleanupProfile })"

    if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
        throw 'pandoc not found in PATH. Install: winget install --id JohnMacFarlane.Pandoc'
    }

    # ---- decide which uv stages run before doing any work, so a missing uv fails fast
    $runCleanup = (-not $NoCleanup) -and (Test-CleanupEnabled $CleanupProfile)
    $runPrep = -not $NoPandocPrep
    if (($runCleanup -or $runPrep) -and -not (Get-Command uv -ErrorAction SilentlyContinue)) {
        throw 'uv not found in PATH. Install: winget install --id astral-sh.uv -e  (or run with -NoCleanup -NoPandocPrep)'
    }

    # ---- 1. preflight
    Write-Stage '[1/8] Preflight: drawings pandoc would drop'
    Invoke-Script "$PSScriptRoot\Test-DocxDrawings.ps1" @{ Docx = $InputPath } | Out-Null

    # ---- 2. data cleanup
    Write-Stage '[2/8] Data cleanup'
    $convertInput = $InputPath
    if (-not $runCleanup) {
        Write-Host "Skipped ($(if ($NoCleanup) { '-NoCleanup' } else { "no rule enabled in profile '$CleanupProfile'" }))."
    } else {
        Invoke-DocxCleanup $InputPath $cleanDocx $CleanupProfile $cleanManifest $CleanupMode
        $convertInput = $cleanDocx
    }

    # ---- 3. drawings -> PNG
    Write-Stage '[3/8] Convert drawings to PNG with Word'
    $convArgs = @{
        InputPath = $convertInput; OutputPath = $shapesDocx; ImageDir = $shapesDir; Dpi = $Dpi
        IncludeTextBoxes = $IncludeTextBoxes; KeepMetafiles = $KeepMetafiles
        NoCluster = $NoCluster; NoTrim = $NoTrim; Visible = $Visible
        UnlinkShapeFields = $UnlinkShapeFields; NoHeadingNumbers = $NoHeadingNumbers; NoPageInfo = $NoPageInfo
    }
    $convertRc = Invoke-Script "$PSScriptRoot\Convert-ShapesToPictures.ps1" $convArgs
    if ($convertRc -eq 2) { Write-Warning "Some drawings failed to convert - see $shapesDir\manifest.csv" }
    elseif ($convertRc -ne 0) { throw "Conversion failed with exit code $convertRc." }
    if (-not (Test-Path -LiteralPath $shapesDocx)) { throw "Converted file was not created: $shapesDocx" }

    # ---- 4. pandoc compatibility fixes (after the last Word save, right before pandoc)
    Write-Stage '[4/8] Prepare for pandoc'
    $pandocInput = $shapesDocx
    if ($runPrep) {
        Invoke-DocxCleanup $shapesDocx $prepDocx 'pandoc' $prepManifest ''
        $pandocInput = $prepDocx
    } else {
        Write-Host 'Skipped (-NoPandocPrep).'
    }

    # ---- 5. pandoc (inside the work folder so image links stay relative: images/media/...)
    #      --wrap=none: never break an image/link over several lines
    #      figures.lua: clean title/alt of converted drawings, move caption anchors to figures
    Write-Stage '[5/8] pandoc'
    $pandocArgs = @('-f', 'docx', '-t', $OutputFormat, '--wrap=none', '--extract-media=./images')
    if (-not $NoFigureFilter) { $pandocArgs += "--lua-filter=$(Join-Path $PSScriptRoot 'pandoc\figures.lua')" }
    Write-Host ("pandoc " + ($pandocArgs -join ' '))
    # media from an earlier run would otherwise stay and be converted again in step 6
    if (Test-Path -LiteralPath $mediaDir) { Remove-Item -LiteralPath $mediaDir -Recurse -Force }
    Push-Location $workDir
    try {
        & pandoc @pandocArgs $pandocInput -o $mdName | Out-Host
        $pandocRc = $LASTEXITCODE
    } finally { Pop-Location }
    if ($pandocRc -ne 0) { throw "pandoc failed with exit code $pandocRc." }
    Write-Host "Written: $(Join-Path $workDir $mdName)"

    # ---- 6. EMF/WMF that reached the output (floating metafile pictures, -KeepMetafiles, ...)
    Write-Stage '[6/8] Convert extracted EMF/WMF to PNG'
    if ($NoMediaConvert) {
        Write-Host 'Skipped (-NoMediaConvert).'
    } else {
        $mediaRc = Invoke-Script "$PSScriptRoot\Convert-MediaToPng.ps1" @{
            Markdown = (Join-Path $workDir $mdName); MediaDir = $mediaDir; Dpi = $Dpi
        }
        if ($mediaRc -ne 0) { Write-Warning 'Some EMF/WMF files could not be converted.' }
    }

    # ---- 7. output folder: the markdown + only the images it links
    Write-Stage '[7/8] Publish output'
    $publishRc = Invoke-Script "$PSScriptRoot\Publish-Output.ps1" @{
        Markdown = (Join-Path $workDir $mdName); OutputDir = $outDir; MediaDir = $mediaDir
    }
    if ($publishRc -eq 1) { throw "Publishing to $outDir failed." }
    $outMd = Join-Path $outDir $mdName

    # ---- 8. gate (on the published markdown: what the reader gets)
    Write-Stage '[8/8] Gate: final docx + markdown'
    $gateRc = Invoke-Script "$PSScriptRoot\Test-DocxDrawings.ps1" @{ Docx = $pandocInput; Markdown = $outMd }

    Write-Host ''
    Write-Host "Result: $outMd"
    if ($gateRc -eq 0 -and $convertRc -eq 0 -and $publishRc -eq 0) {
        Write-Host '[PASS] All drawings converted, no drawing objects left for pandoc to drop.' -ForegroundColor Green
        $exitCode = 0
    } else {
        Write-Host "[CHECK] Finished with issues: convert=$convertRc, publish=$publishRc, gate=$gateRc. Review the output above and the manifests." -ForegroundColor Yellow
        $exitCode = 3
    }
} catch {
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
} finally {
    try { Stop-Transcript | Out-Null } catch {}
}
exit $exitCode
