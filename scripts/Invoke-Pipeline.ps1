<#
.SYNOPSIS
  docx -> markdown pipeline: preflight -> data cleanup -> shapes to PNG -> pandoc prep -> pandoc -> gate.

.DESCRIPTION
  Each stage reads a .docx and writes a new one; every intermediate file is kept next to
  the input so it can be opened in Word:

    input.clean.docx                 after data cleanup (only when the profile enables a rule)
    input.cleanup-manifest.csv       what the cleanup rules found / removed
    input.shapes.docx                after rasterizing drawings
    input.shapes\                    rendered PNG/EMF + manifest.csv
    input.pandoc.docx                after pandoc compatibility fixes (profiles\pandoc.json)
    input.pandoc-manifest.csv        what those fixes changed
    input.md, images\media\          pandoc output (EMF/WMF converted to PNG)
    input.pipeline.log               transcript of this run

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
$workDir = Split-Path $InputPath -Parent
$name = [IO.Path]::GetFileNameWithoutExtension($InputPath)
$cleanDocx = Join-Path $workDir "$name.clean.docx"
$cleanManifest = Join-Path $workDir "$name.cleanup-manifest.csv"
$shapesDocx = Join-Path $workDir "$name.shapes.docx"
$shapesDir = Join-Path $workDir "$name.shapes"
$prepDocx = Join-Path $workDir "$name.pandoc.docx"
$prepManifest = Join-Path $workDir "$name.pandoc-manifest.csv"
$mdName = "$name.md"
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
    Write-Stage '[1/7] Preflight: drawings pandoc would drop'
    Invoke-Script "$PSScriptRoot\Test-DocxDrawings.ps1" @{ Docx = $InputPath } | Out-Null

    # ---- 2. data cleanup
    Write-Stage '[2/7] Data cleanup'
    $convertInput = $InputPath
    if (-not $runCleanup) {
        Write-Host "Skipped ($(if ($NoCleanup) { '-NoCleanup' } else { "no rule enabled in profile '$CleanupProfile'" }))."
    } else {
        Invoke-DocxCleanup $InputPath $cleanDocx $CleanupProfile $cleanManifest $CleanupMode
        $convertInput = $cleanDocx
    }

    # ---- 3. drawings -> PNG
    Write-Stage '[3/7] Convert drawings to PNG with Word'
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
    Write-Stage '[4/7] Prepare for pandoc'
    $pandocInput = $shapesDocx
    if ($runPrep) {
        Invoke-DocxCleanup $shapesDocx $prepDocx 'pandoc' $prepManifest ''
        $pandocInput = $prepDocx
    } else {
        Write-Host 'Skipped (-NoPandocPrep).'
    }

    # ---- 5. pandoc (inside the input folder so image links stay relative: images/media/...)
    #      --wrap=none: never break an image/link over several lines
    #      figures.lua: clean title/alt of converted drawings, move caption anchors to figures
    Write-Stage '[5/7] pandoc'
    $pandocArgs = @('-f', 'docx', '-t', $OutputFormat, '--wrap=none', '--extract-media=./images')
    if (-not $NoFigureFilter) { $pandocArgs += "--lua-filter=$(Join-Path $PSScriptRoot 'pandoc\figures.lua')" }
    Write-Host ("pandoc " + ($pandocArgs -join ' '))
    Push-Location $workDir
    try {
        & pandoc @pandocArgs $pandocInput -o $mdName | Out-Host
        $pandocRc = $LASTEXITCODE
    } finally { Pop-Location }
    if ($pandocRc -ne 0) { throw "pandoc failed with exit code $pandocRc." }
    Write-Host "Written: $(Join-Path $workDir $mdName)"

    # ---- 6. EMF/WMF that reached the output (floating metafile pictures, -KeepMetafiles, ...)
    Write-Stage '[6/7] Convert extracted EMF/WMF to PNG'
    if ($NoMediaConvert) {
        Write-Host 'Skipped (-NoMediaConvert).'
    } else {
        $mediaRc = Invoke-Script "$PSScriptRoot\Convert-MediaToPng.ps1" @{
            Markdown = (Join-Path $workDir $mdName); MediaDir = (Join-Path $workDir 'images'); Dpi = $Dpi
        }
        if ($mediaRc -ne 0) { Write-Warning 'Some EMF/WMF files could not be converted.' }
    }

    # ---- 7. gate
    Write-Stage '[7/7] Gate: final docx + markdown'
    $gateRc = Invoke-Script "$PSScriptRoot\Test-DocxDrawings.ps1" @{ Docx = $pandocInput; Markdown = (Join-Path $workDir $mdName) }

    Write-Host ''
    if ($gateRc -eq 0 -and $convertRc -eq 0) {
        Write-Host '[PASS] All drawings converted, no drawing objects left for pandoc to drop.' -ForegroundColor Green
        $exitCode = 0
    } else {
        Write-Host "[CHECK] Finished with issues: convert=$convertRc, gate=$gateRc. Review the output above and the manifests." -ForegroundColor Yellow
        $exitCode = 3
    }
} catch {
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
} finally {
    try { Stop-Transcript | Out-Null } catch {}
}
exit $exitCode
