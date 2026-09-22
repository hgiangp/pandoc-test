<#
.SYNOPSIS
  docx -> markdown pipeline: preflight -> data cleanup -> shapes to PNG -> pandoc -> gate.

.DESCRIPTION
  Each stage reads a .docx and writes a new one; every intermediate file is kept next to
  the input so it can be opened in Word:

    input.clean.docx                 after data cleanup (only when the profile enables a rule)
    input.cleanup-manifest.csv       what the cleanup rules found / removed
    input.shapes.docx                after rasterizing drawings
    input.shapes\                    rendered PNG/EMF + manifest.csv
    input.md, images\media\          pandoc output
    input.pipeline.log               transcript of this run

  Data cleanup is controlled by a profile in profiles\<name>.json. The default profile
  enables no rule, so the stage is skipped (and uv/Python are not needed). The cleanup
  package is a uv project (cleanup\pyproject.toml + uv.lock); uv installs the pinned
  Python and dependencies on first use.

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
    [switch]$Visible
)

$InputPath = (Resolve-Path -LiteralPath $InputPath).Path
$workDir = Split-Path $InputPath -Parent
$name = [IO.Path]::GetFileNameWithoutExtension($InputPath)
$cleanDocx = Join-Path $workDir "$name.clean.docx"
$cleanManifest = Join-Path $workDir "$name.cleanup-manifest.csv"
$shapesDocx = Join-Path $workDir "$name.shapes.docx"
$shapesDir = Join-Path $workDir "$name.shapes"
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

    # ---- decide cleanup before doing any work, so a missing uv fails fast
    $runCleanup = (-not $NoCleanup) -and (Test-CleanupEnabled $CleanupProfile)
    if ($runCleanup -and -not (Get-Command uv -ErrorAction SilentlyContinue)) {
        throw 'uv not found in PATH. Install: winget install --id astral-sh.uv -e  (or run with -NoCleanup)'
    }

    # ---- 1. preflight
    Write-Stage '[1/5] Preflight: drawings pandoc would drop'
    Invoke-Script "$PSScriptRoot\Test-DocxDrawings.ps1" @{ Docx = $InputPath } | Out-Null

    # ---- 2. data cleanup
    Write-Stage '[2/5] Data cleanup'
    $convertInput = $InputPath
    if (-not $runCleanup) {
        Write-Host "Skipped ($(if ($NoCleanup) { '-NoCleanup' } else { "no rule enabled in profile '$CleanupProfile'" }))."
    } else {
        # --locked: fail instead of silently re-resolving if uv.lock is out of date
        # --no-dev: do not install test tools on the processing machine
        $uvArgs = @('run', '--project', (Join-Path $PSScriptRoot 'cleanup'), '--locked', '--no-dev',
                    'docx-cleanup', $InputPath, '-o', $cleanDocx, '--profile', $CleanupProfile,
                    '--manifest', $cleanManifest)
        if ($CleanupMode) { $uvArgs += @('--mode', $CleanupMode) }
        & uv @uvArgs | Out-Host
        $rc = $LASTEXITCODE
        if ($rc -ne 0) { throw "Data cleanup failed with exit code $rc." }
        $convertInput = $cleanDocx
    }

    # ---- 3. drawings -> PNG
    Write-Stage '[3/5] Convert drawings to PNG with Word'
    $convArgs = @{
        InputPath = $convertInput; OutputPath = $shapesDocx; ImageDir = $shapesDir; Dpi = $Dpi
        IncludeTextBoxes = $IncludeTextBoxes; KeepMetafiles = $KeepMetafiles
        NoCluster = $NoCluster; NoTrim = $NoTrim; Visible = $Visible
    }
    $convertRc = Invoke-Script "$PSScriptRoot\Convert-ShapesToPictures.ps1" $convArgs
    if ($convertRc -eq 2) { Write-Warning "Some drawings failed to convert - see $shapesDir\manifest.csv" }
    elseif ($convertRc -ne 0) { throw "Conversion failed with exit code $convertRc." }
    if (-not (Test-Path -LiteralPath $shapesDocx)) { throw "Converted file was not created: $shapesDocx" }

    # ---- 4. pandoc (inside the input folder so image links stay relative: images/media/...)
    #      --wrap=none: never break an image/link over several lines
    #      figures.lua: clean title/alt of converted drawings, move caption anchors to figures
    Write-Stage '[4/5] pandoc'
    $pandocArgs = @('-f', 'docx', '-t', $OutputFormat, '--wrap=none', '--extract-media=./images')
    if (-not $NoFigureFilter) { $pandocArgs += "--lua-filter=$(Join-Path $PSScriptRoot 'pandoc\figures.lua')" }
    Write-Host ("pandoc " + ($pandocArgs -join ' '))
    Push-Location $workDir
    try {
        & pandoc @pandocArgs $shapesDocx -o $mdName | Out-Host
        $pandocRc = $LASTEXITCODE
    } finally { Pop-Location }
    if ($pandocRc -ne 0) { throw "pandoc failed with exit code $pandocRc." }
    Write-Host "Written: $(Join-Path $workDir $mdName)"

    # ---- 5. gate
    Write-Stage '[5/5] Gate: converted docx + markdown'
    $gateRc = Invoke-Script "$PSScriptRoot\Test-DocxDrawings.ps1" @{ Docx = $shapesDocx; Markdown = (Join-Path $workDir $mdName) }

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
