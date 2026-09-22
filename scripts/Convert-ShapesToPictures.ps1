<#
.SYNOPSIS
  Rasterize Word drawing objects (drawing canvas, group, autoshape, SmartArt, chart,
  OLE/Visio, EMF/WMF pictures) into PNG pictures so that pandoc can extract them.

.DESCRIPTION
  pandoc's docx reader only extracts raster pictures (pic:pic / v:imagedata). Vector
  drawings made with Insert > Shapes are dropped silently. This script opens the document
  with Microsoft Word (COM automation), renders every such drawing to PNG through Word's
  own EMF output, and replaces the drawing with an inline PNG picture. The text found
  inside the drawing is written into the picture's alt text (pandoc emits it as ![alt]).

  The input file is never modified; the result is saved as a new .docx.

  Requirements: Windows + Microsoft Word desktop (2013 or newer), Windows PowerShell 5.1.

.EXAMPLE
  # 1. Inventory only - nothing is changed, writes manifest.csv
  powershell -ExecutionPolicy Bypass -File .\Convert-ShapesToPictures.ps1 -InputPath .\input.docx -DryRun

.EXAMPLE
  # 2. Convert, then run pandoc on the result
  powershell -ExecutionPolicy Bypass -File .\Convert-ShapesToPictures.ps1 -InputPath .\input.docx
  pandoc -f docx -t markdown --extract-media=./images .\input.shapes.docx -o output.md
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    # Default: <input>.shapes.docx next to the input
    [string]$OutputPath,
    # Default: <input>.shapes\ next to the output (PNG, EMF, manifest.csv)
    [string]$ImageDir,
    [ValidateRange(72, 600)][int]$Dpi = 200,
    # Also rasterize standalone floating text boxes (they usually hold prose, so off by default)
    [switch]$IncludeTextBoxes,
    # Keep OLE objects and EMF/WMF pictures as they are (pandoc then outputs .emf/.wmf files)
    [switch]$KeepMetafiles,
    # Do not merge loose floating shapes that belong to the same figure into one picture
    [switch]$NoCluster,
    # Do not crop white margins around the rendered drawing
    [switch]$NoTrim,
    # Only list what would be converted
    [switch]$DryRun,
    # Show the Word window (debugging)
    [switch]$Visible,
    # Turn fields inside shapes into plain text instead of only locking them. Use this if
    # rendered images still show "Error! Reference source not found."
    [switch]$UnlinkShapeFields
)

$ErrorActionPreference = 'Stop'

# Add-Type and COM are blocked in Constrained Language Mode (AppLocker/WDAC-managed machines)
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    $mode = $ExecutionContext.SessionState.LanguageMode
    throw "PowerShell is running in $mode mode; this script needs FullLanguage (Add-Type + COM). Ask IT to allow it, or run on a machine without AppLocker/WDAC script restrictions."
}

Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.IO;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Threading;

public static class ShapeRaster
{
    [DllImport("user32.dll", SetLastError = true)] static extern bool OpenClipboard(IntPtr hWndNewOwner);
    [DllImport("user32.dll", SetLastError = true)] static extern bool CloseClipboard();
    [DllImport("user32.dll", SetLastError = true)] static extern bool EmptyClipboard();
    [DllImport("user32.dll")] static extern bool IsClipboardFormatAvailable(uint format);
    [DllImport("user32.dll")] static extern IntPtr GetClipboardData(uint uFormat);
    [DllImport("gdi32.dll")] static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, byte[] lpbBuffer);
    const uint CF_ENHMETAFILE = 14;

    static bool Open()
    {
        for (int i = 0; i < 30; i++)
        {
            if (OpenClipboard(IntPtr.Zero)) return true;
            Thread.Sleep(100);
        }
        return false;
    }

    public static void ClearClipboard()
    {
        if (!Open()) return;
        try { EmptyClipboard(); } finally { CloseClipboard(); }
    }

    // Returns the Enhanced Metafile currently on the clipboard, or null.
    public static byte[] GetClipboardEmf()
    {
        if (!Open()) return null;
        try
        {
            if (!IsClipboardFormatAvailable(CF_ENHMETAFILE)) return null;
            IntPtr h = GetClipboardData(CF_ENHMETAFILE);
            if (h == IntPtr.Zero) return null;
            uint size = GetEnhMetaFileBits(h, 0, null);
            if (size == 0) return null;
            byte[] buf = new byte[size];
            GetEnhMetaFileBits(h, size, buf);
            return buf;
        }
        finally { CloseClipboard(); }
    }

    // Renders an EMF to PNG at the given DPI (physical size is preserved through the PNG DPI).
    // Returns { widthPx, heightPx }.
    public static int[] RenderEmfToPng(byte[] emf, string pngPath, int dpi, bool trim, int pad, int maxSide)
    {
        using (var ms = new MemoryStream(emf))
        using (var mf = new Metafile(ms))
        {
            MetafileHeader h = mf.GetMetafileHeader();
            double wIn = h.Bounds.Width / (double)h.DpiX;
            double hIn = h.Bounds.Height / (double)h.DpiY;
            float effDpi = dpi;
            int w = (int)Math.Ceiling(wIn * dpi);
            int ht = (int)Math.Ceiling(hIn * dpi);
            if (w < 1 || ht < 1) throw new InvalidOperationException("Empty metafile");
            if (Math.Max(w, ht) > maxSide)
            {
                double f = maxSide / (double)Math.Max(w, ht);
                w = Math.Max(1, (int)(w * f));
                ht = Math.Max(1, (int)(ht * f));
                effDpi = (float)(dpi * f);
            }

            using (var bmp = new Bitmap(w, ht, PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                    g.DrawImage(mf, new Rectangle(0, 0, w, ht));
                }
                Rectangle crop = trim ? ContentBounds(bmp, pad) : new Rectangle(0, 0, w, ht);
                using (var outBmp = bmp.Clone(crop, PixelFormat.Format24bppRgb))
                {
                    outBmp.SetResolution(effDpi, effDpi);
                    outBmp.Save(pngPath, ImageFormat.Png);
                    return new int[] { outBmp.Width, outBmp.Height };
                }
            }
        }
    }

    static Rectangle ContentBounds(Bitmap bmp, int pad)
    {
        var full = new Rectangle(0, 0, bmp.Width, bmp.Height);
        BitmapData data = bmp.LockBits(full, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        int stride = data.Stride;
        byte[] px = new byte[stride * bmp.Height];
        Marshal.Copy(data.Scan0, px, 0, px.Length);
        bmp.UnlockBits(data);

        int minX = bmp.Width, minY = bmp.Height, maxX = -1, maxY = -1;
        for (int y = 0; y < bmp.Height; y++)
        {
            int row = y * stride;
            for (int x = 0; x < bmp.Width; x++)
            {
                int i = row + x * 4;
                if (px[i] < 245 || px[i + 1] < 245 || px[i + 2] < 245)
                {
                    if (x < minX) minX = x;
                    if (x > maxX) maxX = x;
                    if (y < minY) minY = y;
                    if (y > maxY) maxY = y;
                }
            }
        }
        if (maxX < 0) return full;
        minX = Math.Max(0, minX - pad);
        minY = Math.Max(0, minY - pad);
        maxX = Math.Min(bmp.Width - 1, maxX + pad);
        maxY = Math.Min(bmp.Height - 1, maxY + pad);
        return new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
    }
}
'@

# ---------------------------------------------------------------- constants

$MsoTypeName = @{
    1 = 'AutoShape'; 2 = 'Callout'; 3 = 'Chart'; 4 = 'Comment'; 5 = 'Freeform'; 6 = 'Group'
    7 = 'EmbeddedOLEObject'; 8 = 'FormControl'; 9 = 'Line'; 10 = 'LinkedOLEObject'
    11 = 'LinkedPicture'; 12 = 'OLEControlObject'; 13 = 'Picture'; 14 = 'Placeholder'
    15 = 'TextEffect'; 16 = 'Media'; 17 = 'TextBox'; 18 = 'ScriptAnchor'; 19 = 'Table'
    20 = 'Canvas'; 21 = 'Diagram'; 22 = 'Ink'; 23 = 'InkComment'; 24 = 'SmartArt'; 28 = 'Graphic'
}
$InlineTypeName = @{
    1 = 'EmbeddedOLEObject'; 2 = 'LinkedOLEObject'; 3 = 'Picture'; 4 = 'LinkedPicture'
    5 = 'OLEControlObject'; 6 = 'HorizontalLine'; 7 = 'PictureBullet'; 8 = 'PictureHorizontalLine'
    9 = 'LinkedPictureHorizontalLine'; 10 = 'ScriptAnchor'; 11 = 'OWSAnchor'; 12 = 'Chart'
    13 = 'Diagram'; 14 = 'LockedCanvas'; 15 = 'SmartArt'; 16 = 'WebVideo'
}

# Floating shape types that pandoc drops and that we rasterize
$FloatingConvert = @(1, 2, 3, 5, 6, 9, 15, 20, 21, 24)
if ($IncludeTextBoxes) { $FloatingConvert += 17 }
if (-not $KeepMetafiles) { $FloatingConvert += 7, 10 }
# Floating shape types that are "drawing strokes" - their presence makes a cluster a figure
$DrawingStroke = @(1, 2, 5, 9, 20, 21)

# XML markers (inside a range's WordOpenXML) of content pandoc cannot extract
$DrawingMarkers = @('wpc:wpc', 'wpg:wgp', 'wps:wsp', 'dgm:relIds', 'c:chart', 'v:group',
                    'v:rect', 'v:roundrect', 'v:oval', 'v:line', 'v:polyline', 'v:arc', 'v:curve')

$CaptionRegex = '^\s*(Fig\.?|Figure|H[i\u00ec]nh)\s*\d'

$wdCollapseStart = 1
$wdActiveEndPageNumber = 3
$wdFormatDocumentDefault = 16
$wdDoNotSaveChanges = 0
$wdPasteEnhancedMetafile = 9
$wdInLine = 0
$wdStyleNormal = -1
$wdNoProtection = -1
$msoAutomationSecurityForceDisable = 3
$wdTextFrameStory = 5

# ---------------------------------------------------------------- paths

function Resolve-FullPath([string]$p) {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p)
}

$InputPath = (Resolve-Path -LiteralPath $InputPath).Path
$baseName = [IO.Path]::GetFileNameWithoutExtension($InputPath)
if (-not $OutputPath) { $OutputPath = Join-Path (Split-Path $InputPath -Parent) "$baseName.shapes.docx" }
$OutputPath = Resolve-FullPath $OutputPath
if (-not $ImageDir) { $ImageDir = Join-Path (Split-Path $OutputPath -Parent) "$baseName.shapes" }
$ImageDir = Resolve-FullPath $ImageDir
if ($OutputPath -eq $InputPath) { throw 'OutputPath must differ from InputPath.' }
New-Item -ItemType Directory -Force -Path $ImageDir | Out-Null

# ---------------------------------------------------------------- helpers

$script:nextId = 0
$script:manifest = New-Object System.Collections.Generic.List[object]

function New-Entry($kind, $typeCode, $typeName, $name) {
    $script:nextId++
    $e = [ordered]@{
        Id = ('S{0:D3}' -f $script:nextId); Kind = $kind; TypeCode = $typeCode; TypeName = $typeName
        Name = $name; Page = ''; Caption = ''; Action = ''; Method = ''; Png = ''; PngPx = ''
        Text = ''; Error = ''
    }
    $script:manifest.Add($e)
    return $e
}

function Normalize-Text([string]$t) {
    if (-not $t) { return '' }
    # Word returns '/' for an inline picture and \a for a table cell end
    (($t -replace '[\r\n\v\a\f]+', ' / ') -replace '\s+', ' ').Trim([char[]]' /')
}

# Word updates REF/PAGEREF fields while a shape is converted or rendered. The bookmark then
# is not in scope and the field renders as "Error! Reference source not found." - which ends
# up in the PNG. Locked fields are never updated, so the current result is kept.
function Protect-StoryFields($story, [bool]$unlink) {
    $count = 0
    # backwards: Unlink() removes the field from the collection
    for ($i = $story.Fields.Count; $i -ge 1; $i--) {
        try {
            $field = $story.Fields.Item($i)
            if ($unlink) { $field.Unlink() } else { $field.Locked = $true }
            $count++
        } catch {}
    }
    return $count
}

function Protect-ShapeFields($doc, [bool]$unlink) {
    $count = 0
    try {
        $story = $doc.StoryRanges.Item($wdTextFrameStory)
    } catch {
        return 0          # the document has no shape with text
    }
    while ($story) {
        $count += Protect-StoryFields $story $unlink
        $story = $story.NextStoryRange
    }
    return $count
}

# Text inside a floating shape (recurses into groups and canvases)
function Get-ShapeText($s) {
    $out = New-Object System.Collections.Generic.List[string]
    $type = 0
    try { $type = [int]$s.Type } catch {}
    if ($type -eq 6) {
        try { foreach ($i in $s.GroupItems) { foreach ($t in (Get-ShapeText $i)) { $out.Add($t) } } } catch {}
    }
    if ($type -eq 20) {
        try { foreach ($i in $s.CanvasItems) { foreach ($t in (Get-ShapeText $i)) { $out.Add($t) } } } catch {}
    }
    try {
        if ($s.TextFrame.HasText) {
            $t = Normalize-Text $s.TextFrame.TextRange.Text
            if ($t) { $out.Add($t) }
        }
    } catch {}
    return $out
}

function Strip-Fallback([string]$xml) {
    [regex]::Replace($xml, '(?s)<mc:Fallback>.*?</mc:Fallback>', '')
}

# Text inside drawings of an OOXML fragment (text boxes, SmartArt, charts)
function Get-XmlText([string]$xml) {
    $seen = New-Object System.Collections.Generic.List[string]
    foreach ($tb in [regex]::Matches($xml, '(?s)<w:txbxContent>(.*?)</w:txbxContent>')) {
        foreach ($p in [regex]::Matches($tb.Groups[1].Value, '(?s)<w:p[ >].*?</w:p>')) {
            $t = (([regex]::Matches($p.Value, '<w:t(?: [^>]*)?>([^<]*)</w:t>') | ForEach-Object { $_.Groups[1].Value }) -join '')
            $t = Normalize-Text ([Net.WebUtility]::HtmlDecode($t))
            if ($t -and -not $seen.Contains($t)) { $seen.Add($t) }
        }
    }
    foreach ($m in [regex]::Matches($xml, '<a:t>([^<]*)</a:t>')) {
        $t = Normalize-Text ([Net.WebUtility]::HtmlDecode($m.Groups[1].Value))
        if ($t -and -not $seen.Contains($t)) { $seen.Add($t) }
    }
    return $seen
}

function Has-Tag([string]$xml, [string]$tag) {
    [regex]::IsMatch($xml, '<' + [regex]::Escape($tag) + '[\s>/]')
}

# Decide whether an inline shape needs rasterizing, based on its OOXML (not its COM type,
# because inline DrawingML shapes/groups/canvases are not reliably typed by Word).
function Get-InlineDecision($is) {
    $xml = Strip-Fallback $is.Range.WordOpenXML
    foreach ($m in $DrawingMarkers) {
        if (Has-Tag $xml $m) { return @{ Convert = $true; Reason = $m; Xml = $xml } }
    }
    if ((Has-Tag $xml 'v:shape') -and -not (Has-Tag $xml 'v:imagedata')) {
        return @{ Convert = $true; Reason = 'v:shape'; Xml = $xml }
    }
    if (Has-Tag $xml 'o:OLEObject') {
        return @{ Convert = (-not $KeepMetafiles); Reason = 'o:OLEObject'; Xml = $xml }
    }
    if ([regex]::IsMatch($xml, 'pkg:name="/word/media/[^"]+\.(emf|wmf)"', 'IgnoreCase')) {
        return @{ Convert = (-not $KeepMetafiles); Reason = 'emf/wmf picture'; Xml = $xml }
    }
    if ((Has-Tag $xml 'pic:pic') -or (Has-Tag $xml 'v:imagedata')) {
        return @{ Convert = $false; Reason = 'raster picture (pandoc ok)'; Xml = $xml }
    }
    return @{ Convert = $false; Reason = 'unknown content'; Xml = $xml }
}

function Get-Page($range) {
    try { return [int]$range.Information($wdActiveEndPageNumber) } catch { return '' }
}

# Caption of the figure a range belongs to: the paragraph itself, or the first non-empty
# paragraph after it (skipping the empty paragraphs that often hold a drawing's space).
function Get-NearCaption($range) {
    try {
        $p = $range.Paragraphs.Item(1)
        for ($i = 0; $i -lt 30 -and $p; $i++) {
            $t = Normalize-Text $p.Range.Text
            if ($t) {
                if ($t -match $CaptionRegex) { return $t }
                return ''
            }
            $p = $p.Next()
        }
    } catch {}
    return ''
}

function Build-AltText($texts) {
    $alt = 'Drawing converted to image'
    $joined = (@($texts) | Where-Object { $_ }) -join ' | '
    if ($joined) { $alt += '. Text: ' + $joined }
    if ($alt.Length -gt 1500) { $alt = $alt.Substring(0, 1500) + '...' }
    return $alt
}

function Save-Render($entry, [byte[]]$emf) {
    $emfPath = Join-Path $ImageDir ($entry.Id + '.emf')
    $pngPath = Join-Path $ImageDir ($entry.Id + '.png')
    [IO.File]::WriteAllBytes($emfPath, $emf)
    $px = [ShapeRaster]::RenderEmfToPng($emf, $pngPath, $Dpi, (-not $NoTrim), 8, 10000)
    $entry.Png = [IO.Path]::GetFileName($pngPath)
    $entry.PngPx = '{0}x{1}' -f $px[0], $px[1]
    return $pngPath
}

function Insert-Picture($doc, $range, $pngPath, $entry, $texts) {
    $pic = $doc.InlineShapes.AddPicture($pngPath, $false, $true, $range)
    $pic.AlternativeText = Build-AltText $texts
    $pic.Title = 'shape2png:' + $entry.Id
    return $pic
}

# Rasterize an inline shape in place
function Convert-Inline($doc, $is, $entry, $texts) {
    $r = $is.Range
    $bits = $r.EnhMetaFileBits
    if (-not $bits) { throw 'Range.EnhMetaFileBits returned nothing' }
    $png = Save-Render $entry ([byte[]]$bits)
    $pos = $r.Duplicate
    $pos.Collapse($wdCollapseStart)
    $is.Delete()
    Insert-Picture $doc $pos $png $entry $texts | Out-Null
}

# Get an EMF of a floating shape via copy -> clipboard (Win32), or copy -> Paste Special EMF
function Get-FloatingEmf($word, $shape) {
    [ShapeRaster]::ClearClipboard()
    try { $shape.Select() } catch { throw "Cannot select shape (try -Visible): $($_.Exception.Message)" }
    $word.Selection.Copy()
    Start-Sleep -Milliseconds 300
    $emf = [ShapeRaster]::GetClipboardEmf()
    if ($emf) { return @{ Emf = $emf; Method = 'clipboard-emf' } }

    if (-not $script:tmpDoc) { $script:tmpDoc = $word.Documents.Add() }
    $script:tmpDoc.Content.Delete() | Out-Null
    $m = [Type]::Missing
    $script:tmpDoc.Content.PasteSpecial($m, $false, $wdInLine, $false, $wdPasteEnhancedMetafile)
    if ($script:tmpDoc.InlineShapes.Count -lt 1) { throw 'Paste Special (EMF) produced no picture' }
    $bits = $script:tmpDoc.InlineShapes.Item(1).Range.EnhMetaFileBits
    if (-not $bits) { throw 'EnhMetaFileBits of pasted EMF returned nothing' }
    return @{ Emf = [byte[]]$bits; Method = 'pastespecial-emf' }
}

# Rasterize a floating shape; the picture goes into its own paragraph before the anchor
function Convert-Floating($word, $doc, $shape, $entry) {
    $texts = Get-ShapeText $shape
    $entry.Text = ($texts -join ' | ')

    # Word only supports this for pictures/OLE objects; other shapes fall through to the clipboard
    $is = $null
    try { $is = $shape.ConvertToInlineShape() } catch {}
    if ($is) {
        $entry.Method = 'convert-to-inline'
        Convert-Inline $doc $is $entry $texts
        return
    }

    $res = Get-FloatingEmf $word $shape
    $entry.Method = $res.Method
    $png = Save-Render $entry $res.Emf

    $para = $shape.Anchor.Paragraphs.Item(1).Range
    $pos = $para.Duplicate
    $pos.Collapse($wdCollapseStart)
    if (Normalize-Text $para.Text) {
        $para.InsertParagraphBefore()
        $pos = $doc.Range($para.Start, $para.Start)
        $pos.Style = $wdStyleNormal
    }
    Insert-Picture $doc $pos $png $entry $texts | Out-Null
    $shape.Delete()
}

# ---------------------------------------------------------------- main

$word = $null
$doc = $null
$script:tmpDoc = $null
try {
    Write-Host "Starting Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Visible
    $word.DisplayAlerts = 0
    # Documents opened through COM run macros by default; never run macros of the input file
    $word.AutomationSecurity = $msoAutomationSecurityForceDisable

    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
    $doc = $word.Documents.Open($InputPath, $false, $true, $false)
    Write-Host ("Opened {0} (compatibility mode {1})" -f $InputPath, $doc.CompatibilityMode)
    if (-not $DryRun) {
        $doc.SaveAs2($OutputPath, $wdFormatDocumentDefault)
        $doc.TrackRevisions = $false
        if ($doc.ProtectionType -ne $wdNoProtection) {
            try { $doc.Unprotect() } catch {
                throw 'Document is protected (Restrict Editing) with a password; remove the protection in Word first.'
            }
        }
    }
    $doc.Activate()

    # ---- keep the fields inside shapes as they are now (before any conversion)
    if (-not $DryRun) {
        try { $word.Options.UpdateFieldsAtPrint = $false } catch {}
        $protected = Protect-ShapeFields $doc ([bool]$UnlinkShapeFields)
        Write-Host ("Fields inside shapes {0}: {1}" -f $(if ($UnlinkShapeFields) { 'unlinked' } else { 'locked' }), $protected)
    }

    # ---- pass 1: floating shapes (main story)
    $floating = @()
    for ($i = 1; $i -le $doc.Shapes.Count; $i++) { $floating += $doc.Shapes.Item($i) }
    Write-Host "Floating shapes: $($floating.Count)"

    $items = @()
    foreach ($s in $floating) {
        $t = [int]$s.Type
        $anchor = $s.Anchor
        $items += [pscustomobject]@{
            Shape = $s; Type = $t; Name = [string]$s.Name
            AnchorStart = $anchor.Paragraphs.Item(1).Range.Start
            Page = Get-Page $anchor; Caption = Get-NearCaption $anchor
        }
    }

    # Group loose shapes of one figure (same caption on same page, or same anchor paragraph)
    $units = @()
    if ($NoCluster) {
        foreach ($it in $items) { $units += , @($it) }
    } else {
        $groups = $items | Group-Object {
            if ($_.Caption) { 'cap|' + $_.Page + '|' + $_.Caption } else { 'para|' + $_.AnchorStart }
        }
        foreach ($g in $groups) {
            $members = @($g.Group)
            $hasStroke = @($members | Where-Object { $DrawingStroke -contains $_.Type }).Count -gt 0
            if ($members.Count -gt 1 -and $hasStroke) { $units += , $members }
            else { foreach ($m in $members) { $units += , @($m) } }
        }
    }

    foreach ($unit in $units) {
        $first = $unit[0]
        if ($unit.Count -gt 1) {
            $names = ($unit | ForEach-Object { $_.Name }) -join ', '
            $entry = New-Entry 'Floating' 'cluster' ("Cluster of $($unit.Count)") $names
        } else {
            $typeName = $MsoTypeName[$first.Type]; if (-not $typeName) { $typeName = 'Unknown' }
            $entry = New-Entry 'Floating' $first.Type $typeName $first.Name
        }
        $entry.Page = $first.Page
        $entry.Caption = $first.Caption

        $convert = ($unit.Count -gt 1) -or ($FloatingConvert -contains $first.Type)
        if (-not $convert) {
            $entry.Action = 'skip'
            if ($first.Type -eq 17) { $entry.Error = 'text box kept (use -IncludeTextBoxes)' }
            continue
        }
        if ($DryRun) {
            $entry.Action = 'would-convert'
            $entry.Text = ((@($unit | ForEach-Object { Get-ShapeText $_.Shape })) -join ' | ')
            continue
        }

        # A cluster is grouped into one shape first; if Word refuses (e.g. a canvas cannot be
        # grouped), its members are converted one by one.
        $targets = @()
        if ($unit.Count -gt 1) {
            try {
                # unique temporary names, since Word shape names can repeat
                $k = 0
                foreach ($u in $unit) { $k++; $u.Shape.Name = ('s2p_{0}_{1}' -f $entry.Id, $k) }
                $tmpNames = [object[]]($unit | ForEach-Object { $_.Shape.Name })
                $targets += , @($doc.Shapes.Range($tmpNames).Group(), $entry, 'grouped+')
            } catch {
                $entry.Action = 'split'
                $entry.Error = "Could not group cluster: $($_.Exception.Message)"
                foreach ($u in $unit) {
                    if ($FloatingConvert -notcontains $u.Type) { continue }
                    $sub = New-Entry 'Floating' $u.Type $MsoTypeName[$u.Type] $u.Name
                    $sub.Page = $u.Page
                    $sub.Caption = $u.Caption
                    $targets += , @($u.Shape, $sub, '')
                }
            }
        } else {
            $targets += , @($first.Shape, $entry, '')
        }

        foreach ($tg in $targets) {
            $shape = $tg[0]; $e = $tg[1]; $prefix = $tg[2]
            try {
                Convert-Floating $word $doc $shape $e
                $e.Method = $prefix + $e.Method
                $e.Action = 'converted'
                Write-Host ("  [{0}] converted {1} p.{2} {3}" -f $e.Id, $e.TypeName, $e.Page, $e.Caption)
            } catch {
                $e.Action = 'FAILED'
                $e.Error = $_.Exception.Message
                Write-Warning ("  [{0}] FAILED {1} p.{2}: {3}" -f $e.Id, $e.TypeName, $e.Page, $e.Error)
            }
        }
    }

    # ---- pass 2: inline shapes (main story)
    $inline = @()
    for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) { $inline += $doc.InlineShapes.Item($i) }
    Write-Host "Inline shapes: $($inline.Count)"

    foreach ($is in $inline) {
        $title = ''
        try { $title = [string]$is.Title } catch {}
        if ($title -like 'shape2png:*') { continue }

        $t = [int]$is.Type
        $typeName = $InlineTypeName[$t]; if (-not $typeName) { $typeName = 'Unknown' }
        $entry = New-Entry 'Inline' $t $typeName ''
        $entry.Page = Get-Page $is.Range
        $entry.Caption = Get-NearCaption $is.Range

        try {
            $d = Get-InlineDecision $is
            $texts = Get-XmlText $d.Xml
            $entry.Text = ($texts -join ' | ')
            $entry.Method = $d.Reason
            if (-not $d.Convert) { $entry.Action = 'skip'; continue }
            if ($DryRun) { $entry.Action = 'would-convert'; continue }

            Convert-Inline $doc $is $entry $texts
            $entry.Action = 'converted'
            Write-Host ("  [{0}] converted {1} ({2}) p.{3} {4}" -f $entry.Id, $typeName, $d.Reason, $entry.Page, $entry.Caption)
        } catch {
            $entry.Action = 'FAILED'
            $entry.Error = $_.Exception.Message
            Write-Warning ("  [{0}] FAILED {1} p.{2}: {3}" -f $entry.Id, $typeName, $entry.Page, $entry.Error)
        }
    }

    if (-not $DryRun) {
        $doc.Save()
        Write-Host "Saved: $OutputPath"
    }
}
finally {
    if ($script:tmpDoc) { try { $script:tmpDoc.Close($wdDoNotSaveChanges) } catch {} }
    if ($doc) { try { $doc.Close($wdDoNotSaveChanges) } catch {} }
    if ($word) {
        try { $word.Quit($wdDoNotSaveChanges) } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $csv = Join-Path $ImageDir 'manifest.csv'
    $script:manifest | ForEach-Object { [pscustomobject]$_ } |
        Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Host "Manifest: $csv"
}

$summary = $script:manifest | Group-Object { $_.Action } | ForEach-Object { '{0}={1}' -f $_.Name, $_.Count }
Write-Host ("Summary: " + ($summary -join ', '))
$failed = @($script:manifest | Where-Object { $_.Action -eq 'FAILED' }).Count
if ($failed -gt 0) { exit 2 }
