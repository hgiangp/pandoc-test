<#
.SYNOPSIS
  Copy the final markdown and only the images it links into a clean output folder.

.DESCRIPTION
  pandoc writes the markdown in the work folder next to images\media\, which also holds
  files the reader does not need: EMF/WMF originals kept by Convert-MediaToPng.ps1, or
  media left from an earlier run. This script builds the deliverable:

    <OutputDir>\<name>.md
    <OutputDir>\images\<file>     every media file the markdown links, nothing else

  Links "images/media/x.png" become "images/x.png" (pandoc always adds the media
  subfolder). Only those exact link paths are rewritten; no other text is touched.

  The folder is built as <OutputDir>.tmp and swapped in at the end, so a failed run leaves
  the previous output as it was. An existing folder is replaced only when it is empty or
  was created by this script (marker file .pipeline-output); any other folder is refused.

  Exit code: 0 = ok, 2 = written with issues (broken link, EMF/WMF in the output), 1 = error.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\Publish-Output.ps1 -Markdown .\input.work\input.md -OutputDir .\input.out
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Markdown,
    [Parameter(Mandatory = $true)][string]$OutputDir,
    # Where pandoc put the extracted media; default: <markdown folder>\images
    [string]$MediaDir
)

$ErrorActionPreference = 'Stop'
$markerName = '.pipeline-output'
$sep = [IO.Path]::DirectorySeparatorChar

function Resolve-FullPath([string]$p) {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p).TrimEnd('\', '/')
}

# $inner is $outer itself or somewhere below it
function Test-Within([string]$inner, [string]$outer) {
    $inner.Equals($outer, [StringComparison]::OrdinalIgnoreCase) -or
        $inner.StartsWith($outer + $sep, [StringComparison]::OrdinalIgnoreCase)
}

# Refuse to delete a folder this script did not create
function Assert-Replaceable([string]$dir) {
    if (-not (Test-Path -LiteralPath $dir)) { return }
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) { throw "Not a folder: $dir" }
    if (Test-Path -LiteralPath (Join-Path $dir $markerName)) { return }
    if (@(Get-ChildItem -LiteralPath $dir -Force).Count -eq 0) { return }
    throw "Refusing to replace '$dir': it is not empty and was not created by this pipeline (no $markerName file)."
}

try {
    $Markdown = (Resolve-Path -LiteralPath $Markdown).Path
    $mdDir = Split-Path $Markdown -Parent
    $MediaDir = if ($MediaDir) { Resolve-FullPath $MediaDir } else { Join-Path $mdDir 'images' }
    $OutputDir = Resolve-FullPath $OutputDir
    $tmpDir = $OutputDir + '.tmp'

    if ([IO.Path]::GetPathRoot($OutputDir).TrimEnd('\', '/') -eq $OutputDir) { throw "OutputDir must not be a drive root: $OutputDir" }
    if ((Test-Within $mdDir $OutputDir) -or (Test-Within $OutputDir $mdDir)) {
        throw "OutputDir ($OutputDir) must be separate from the markdown folder ($mdDir)."
    }
    if ((Test-Path -LiteralPath $MediaDir) -and -not (Test-Within $MediaDir $mdDir)) {
        throw "MediaDir ($MediaDir) must be inside the markdown folder: pandoc links media relative to it."
    }
    Assert-Replaceable $OutputDir
    Assert-Replaceable $tmpDir

    if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force }
    $tmpImages = Join-Path $tmpDir 'images'
    New-Item -ItemType Directory -Force -Path $tmpImages | Out-Null
    [IO.File]::WriteAllText((Join-Path $tmpDir $markerName),
        "Created by Publish-Output.ps1. The whole folder is replaced on the next pipeline run.`r`n")

    $md = Get-Content -LiteralPath $Markdown -Raw -Encoding UTF8
    $original = $md
    $known = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $issues = 0
    $copied = 0
    $unused = New-Object System.Collections.Generic.List[string]
    $mediaFiles = @()
    if (Test-Path -LiteralPath $MediaDir) { $mediaFiles = @(Get-ChildItem -LiteralPath $MediaDir -Recurse -File) }

    foreach ($file in $mediaFiles) {
        # the link as pandoc wrote it: relative to the markdown folder, "/" separators
        $link = $file.FullName.Substring($mdDir.Length + 1).Replace('\', '/')
        [void]$known.Add($link)
        $newLink = 'images/' + $file.Name
        # whole link only: "image1.png" must not match inside "image11.png" or "image1.png.bak"
        $pattern = '(?<![\w-])' + [regex]::Escape($link) + '(?![\w.-])'
        if (-not [regex]::IsMatch($md, $pattern)) {
            $unused.Add($file.Name)
            continue
        }
        $target = Join-Path $tmpImages $file.Name
        if (Test-Path -LiteralPath $target) { throw "Two linked media files share the name $($file.Name)." }
        Copy-Item -LiteralPath $file.FullName -Destination $target
        $md = [regex]::Replace($md, $pattern, $newLink.Replace('$', '$$'))
        $copied++
        if ($file.Extension -match '^\.(emf|wmf)$') {
            $issues++
            Write-Warning "  $($file.Name): EMF/WMF in the output - markdown viewers cannot display it."
        }
    }

    # links into the media folder that match no file (checked before the rewrite: the new
    # links "images/x.png" would also look like links into the default media folder)
    $broken = @()
    if ((Test-Within $MediaDir $mdDir) -and $MediaDir.Length -gt $mdDir.Length) {
        $mediaLink = [regex]::Escape($MediaDir.Substring($mdDir.Length + 1).Replace('\', '/') + '/')
        $broken = @([regex]::Matches($original, '(?<![\w-])(?:\./)?(' + $mediaLink + '[^\s)"''<>]+)') |
            ForEach-Object { $_.Groups[1].Value } | Where-Object { -not $known.Contains($_) } | Select-Object -Unique)
    }
    foreach ($b in $broken) {
        $issues++
        Write-Warning "  Broken link, file not found: $b"
    }

    $outMd = Join-Path $tmpDir ([IO.Path]::GetFileName($Markdown))
    # UTF-8 without BOM (Set-Content -Encoding UTF8 would add one in PowerShell 5.1)
    [IO.File]::WriteAllText($outMd, $md, (New-Object Text.UTF8Encoding $false))

    if (Test-Path -LiteralPath $OutputDir) { Remove-Item -LiteralPath $OutputDir -Recurse -Force }
    Rename-Item -LiteralPath $tmpDir -NewName (Split-Path $OutputDir -Leaf)

    Write-Host ("Published {0} image(s), {1} unused media file(s) left out{2}" -f $copied, $unused.Count,
        $(if ($unused.Count) { ': ' + ($unused -join ', ') } else { '' }))
    Write-Host "Output: $OutputDir"
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
if ($issues -gt 0) { exit 2 }
exit 0
