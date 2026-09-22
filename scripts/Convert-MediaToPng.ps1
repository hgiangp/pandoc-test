<#
.SYNOPSIS
  Convert EMF/WMF files extracted by pandoc into PNG and update the links in the markdown.

.DESCRIPTION
  Markdown viewers and browsers cannot display EMF/WMF. Most metafiles are already turned
  into PNG by Convert-ShapesToPictures.ps1, but some reach the output anyway: floating
  EMF/WMF pictures, runs with -KeepMetafiles, or sources not covered yet. This is the
  safety net at the end of the pipeline: whatever metafile is still referenced by the
  markdown gets rendered to PNG next to it, and the link is repointed.

  Only the file name changes in the markdown ("image3.emf" -> "image3.png"); no other text
  is touched. The original metafile is kept unless -RemoveOriginals is given.

  Needs Windows (GDI+ renders the metafile). Exit code: 0 = ok, 1 = a file failed.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\Convert-MediaToPng.ps1 -Markdown .\input.md
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Markdown,
    # Where pandoc put the extracted media; default: <markdown folder>\images
    [string]$MediaDir,
    [ValidateRange(72, 600)][int]$Dpi = 200,
    [switch]$RemoveOriginals
)

$ErrorActionPreference = 'Stop'

if (-not ('ShapeRaster' -as [type])) {
    Add-Type -ReferencedAssemblies System.Drawing `
        -TypeDefinition (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'lib\ShapeRaster.cs') -Raw)
}

$Markdown = (Resolve-Path -LiteralPath $Markdown).Path
if (-not $MediaDir) { $MediaDir = Join-Path (Split-Path $Markdown -Parent) 'images' }
if (-not (Test-Path -LiteralPath $MediaDir)) {
    Write-Host "No media folder ($MediaDir) - nothing to convert."
    exit 0
}

$files = @(Get-ChildItem -LiteralPath $MediaDir -Recurse -File |
    Where-Object { $_.Extension -match '^\.(emf|wmf)$' })
if ($files.Count -eq 0) {
    Write-Host 'No EMF/WMF files in the media folder.'
    exit 0
}

$md = Get-Content -LiteralPath $Markdown -Raw -Encoding UTF8
$converted = 0
$failed = 0
foreach ($file in $files) {
    $png = [IO.Path]::ChangeExtension($file.FullName, '.png')
    try {
        # no trimming: unlike a drawing, a picture's margins can be part of the image
        $size = [ShapeRaster]::RenderEmfToPng([IO.File]::ReadAllBytes($file.FullName), $png, $Dpi, $false, 0, 10000)
        $converted++
        Write-Host ("  {0} -> {1} ({2}x{3})" -f $file.Name, [IO.Path]::GetFileName($png), $size[0], $size[1])
    } catch {
        $failed++
        Write-Warning ("  {0}: {1}" -f $file.Name, $_.Exception.Message)
        continue
    }
    # only the file name is rewritten, wherever it appears (markdown link or <img src=...>)
    $md = $md.Replace($file.Name, [IO.Path]::GetFileName($png))
    if ($RemoveOriginals) { Remove-Item -LiteralPath $file.FullName -Force }
}

# UTF-8 without BOM (Set-Content -Encoding UTF8 would add one in PowerShell 5.1)
[IO.File]::WriteAllText($Markdown, $md, (New-Object Text.UTF8Encoding $false))
Write-Host ("Converted {0} metafile(s), {1} failed. Updated: {2}" -f $converted, $failed, $Markdown)
if ($failed -gt 0) { exit 1 }
