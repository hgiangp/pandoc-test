<#
.SYNOPSIS
  Classify every drawing object in a .docx by what pandoc does with it, and optionally
  compare the number of figure captions with the number of images in pandoc's markdown.
  Does not need Word.

.DESCRIPTION
  Each <w:drawing> (DrawingML), <w:pict> (VML) and <w:object> (OLE) is classified once.
  mc:Fallback content is ignored because it duplicates the mc:Choice content.

  Exit code: 0 = nothing lost, 1 = objects that pandoc drops are present.
  A caption/image mismatch in the markdown is only a warning (heuristic).

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\Test-DocxDrawings.ps1 -Docx .\input.docx
  powershell -ExecutionPolicy Bypass -File .\Test-DocxDrawings.ps1 -Docx .\input.shapes.docx -Markdown .\output.md
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Markdown
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Docx = (Resolve-Path -LiteralPath $Docx).Path
$zip = [IO.Compression.ZipFile]::OpenRead($Docx)
$parts = @{}
$rels = @{}
try {
    foreach ($e in $zip.Entries) {
        if ($e.FullName -match '^word/(_rels/)?(document|header\d*|footer\d*|footnotes|endnotes)\.xml(\.rels)?$') {
            $sr = New-Object IO.StreamReader($e.Open())
            try { $text = $sr.ReadToEnd() } finally { $sr.Dispose() }
            if ($e.FullName -like '*.rels') { $rels[$e.FullName] = $text } else { $parts[$e.FullName] = $text }
        }
    }
    $media = @($zip.Entries | Where-Object { $_.FullName -like 'word/media/*' } | ForEach-Object { $_.FullName })
} finally { $zip.Dispose() }

function Has([string]$xml, [string]$tag) { [regex]::IsMatch($xml, '<' + [regex]::Escape($tag) + '[\s>/]') }

# relationship id -> target, to tell EMF/WMF pictures apart
function Get-RelTargets([string]$part) {
    $map = @{}
    $relName = ($part -replace '^word/', 'word/_rels/') + '.rels'
    if ($rels.ContainsKey($relName)) {
        foreach ($m in [regex]::Matches($rels[$relName], '<Relationship [^>]*>')) {
            $id = [regex]::Match($m.Value, 'Id="([^"]+)"').Groups[1].Value
            $tg = [regex]::Match($m.Value, 'Target="([^"]+)"').Groups[1].Value
            if ($id) { $map[$id] = $tg }
        }
    }
    return $map
}

function Test-Metafile([string]$block, $relMap) {
    foreach ($m in [regex]::Matches($block, 'r:(?:embed|id)="([^"]+)"')) {
        $tg = $relMap[$m.Groups[1].Value]
        if ($tg -and $tg -match '\.(emf|wmf)$') { return $true }
    }
    return $false
}

# Category -> meaning. LOST = pandoc drops it.
$meaning = [ordered]@{
    'Canvas'      = 'LOST   drawing canvas'
    'Group'       = 'LOST   group shape'
    'Shape'       = 'LOST   autoshape / line / arrow'
    'SmartArt'    = 'LOST   SmartArt'
    'Chart'       = 'LOST   chart'
    'TextBox'     = 'INFO   standalone text box (kept by converter; verify its text in the markdown)'
    'OLE'         = 'EMF    OLE object (Visio...) - only its EMF/WMF preview is extracted'
    'Picture-EMF' = 'EMF    picture in EMF/WMF - extracted but not viewable in markdown'
    'Picture'     = 'OK     raster picture'
}

$rows = New-Object System.Collections.Generic.List[object]
foreach ($name in ($parts.Keys | Sort-Object)) {
    $xml = [regex]::Replace($parts[$name], '(?s)<mc:Fallback>.*?</mc:Fallback>', '')
    $relMap = Get-RelTargets $name
    $counts = @{}

    # OLE objects first, then remove them so their VML preview is not counted again
    $counts['OLE'] = [regex]::Matches($xml, '(?s)<w:object[\s>].*?</w:object>').Count
    $xml = [regex]::Replace($xml, '(?s)<w:object[\s>].*?</w:object>', '')

    foreach ($m in [regex]::Matches($xml, '(?s)<w:drawing>.*?</w:drawing>')) {
        $b = $m.Value
        $cat =
            if (Has $b 'wpc:wpc') { 'Canvas' }
            elseif (Has $b 'wpg:wgp') { 'Group' }
            elseif (Has $b 'dgm:relIds') { 'SmartArt' }
            elseif (Has $b 'c:chart') { 'Chart' }
            elseif (Has $b 'wps:wsp') { if ($b -match 'txBox="(1|true)"') { 'TextBox' } else { 'Shape' } }
            elseif (Has $b 'pic:pic') { if (Test-Metafile $b $relMap) { 'Picture-EMF' } else { 'Picture' } }
            else { 'Shape' }
        $counts[$cat]++
    }

    foreach ($m in [regex]::Matches($xml, '(?s)<w:pict[\s>].*?</w:pict>')) {
        $b = $m.Value
        $cat =
            if (Has $b 'v:group') { 'Group' }
            elseif ((Has $b 'v:imagedata') -and -not (Has $b 'v:textbox')) {
                if (Test-Metafile $b $relMap) { 'Picture-EMF' } else { 'Picture' }
            }
            elseif ((Has $b 'v:textbox') -and ($b -match '_x0000_t202|o:spt="202"')) { 'TextBox' }
            else { 'Shape' }
        $counts[$cat]++
    }

    foreach ($cat in $meaning.Keys) {
        if ($counts[$cat]) {
            $rows.Add([pscustomobject]@{ Part = $name; Object = $cat; Count = $counts[$cat]; Meaning = $meaning[$cat] })
        }
    }
}

Write-Host "`n== $Docx"
if ($rows.Count) { $rows | Format-Table -AutoSize | Out-String -Width 200 | Write-Host } else { Write-Host '(no drawing objects)' }

$metafiles = @($media | Where-Object { $_ -match '\.(emf|wmf)$' })
Write-Host ("Media files: {0} total, {1} EMF/WMF" -f $media.Count, $metafiles.Count)

$lost = 0
foreach ($r in $rows) { if ($r.Meaning -like 'LOST*') { $lost += $r.Count } }
Write-Host ("Objects pandoc will drop: {0}" -f $lost)

if ($Markdown) {
    $md = Get-Content -LiteralPath $Markdown -Raw -Encoding UTF8
    # alt text may contain escaped brackets (\]), one level of nested brackets ([]{#id .anchor})
    # and line breaks
    $images = [regex]::Matches($md, '!\[(?:\\.|\[[^\]]*\]|[^\]\\])*\]\(|<img\s').Count
    $mdEmf = [regex]::Matches($md, '\.(emf|wmf)[)"\s]').Count
    # heuristic: "Fig. N-N" at the start of a line / table cell / figure caption, in both
    # output formats: "![Fig. N-N" (markdown) or "<figcaption><p>Fig. N-N" / "<p>Fig. N-N" (gfm),
    # optionally after an anchor span ("[]{#id .anchor}" or "<span id=.. class=anchor></span>")
    $capMatches = [regex]::Matches($md, '(?m)(?:^|\||<figcaption>|<p>)\s*(?:!\[)?(?:\[\]\{[^}]*\}\s*|<span[^>]*>\s*</span>\s*)?(?:Fig\.?|Figure|H[i\u00ec]nh)\s*(\d+(?:[.\-\u2011\u2013]\d+)*)')
    $captions = @($capMatches | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    Write-Host "`n== $Markdown"
    Write-Host ("Images: {0} (EMF/WMF references: {1}); distinct figure captions: {2}" -f $images, $mdEmf, $captions.Count)
    if ($captions.Count -gt $images) {
        Write-Warning ("{0} captions but only {1} images - some figures are probably missing (heuristic)." -f $captions.Count, $images)
    }
}

if ($lost -gt 0) { exit 1 }
