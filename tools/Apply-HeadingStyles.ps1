param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [string]$CopySuffix = ".styled",
    [string]$StartAfter,
    [switch]$IncludeFrontMatter,
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Initialize-Unicode.ps1")

$WdStyleHeading1 = -2
$WdStyleHeading2 = -3
$WdStyleHeading3 = -4
$WdOutlineLevel1 = 1
$WdOutlineLevel2 = 2
$WdOutlineLevel3 = 3

function Join-Chars {
    param([int[]]$CodePoints)
    return -join ($CodePoints | ForEach-Object { [char]$_ })
}

$TocTitle = Join-Chars @(0x76EE, 0x5F55)
$FigureListTitle = Join-Chars @(0x63D2, 0x56FE, 0x6E05, 0x5355)
$TableListTitle = Join-Chars @(0x9644, 0x8868, 0x6E05, 0x5355)

function Get-ParagraphText {
    param($Paragraph)
    $text = [string]$Paragraph.Range.Text
    return ($text -replace "[`r`a]", "").Trim()
}

function Copy-DocumentForUpdate {
    param(
        [System.IO.FileInfo]$Document,
        [string]$Suffix
    )

    if ([string]::IsNullOrWhiteSpace($Suffix)) {
        return $Document
    }

    $copyName = "{0}{1}{2}" -f $Document.BaseName, $Suffix, $Document.Extension
    $copyPath = Join-Path $Document.DirectoryName $copyName
    Copy-Item -LiteralPath $Document.FullName -Destination $copyPath -Force
    return Get-Item -LiteralPath $copyPath
}

function Get-HeadingLevel {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return 0
    }

    if ($Text -match '^\d+\.\d+\.\d+\s*\S+') {
        return 3
    }
    if ($Text -match '^\d+\.\d+\s*\S+') {
        return 2
    }
    if ($Text -match '^\d+\s+\S+') {
        return 1
    }

    return 0
}

function Find-StartPosition {
    param($Document)

    if ($IncludeFrontMatter) {
        return 0
    }

    if (-not [string]::IsNullOrWhiteSpace($StartAfter)) {
        for ($i = 1; $i -le [int]$Document.Paragraphs.Count; $i++) {
            $paragraph = $Document.Paragraphs.Item($i)
            if ((Get-ParagraphText $paragraph) -eq $StartAfter) {
                return [int]$paragraph.Range.End
            }
        }
        throw "Could not find StartAfter paragraph: $StartAfter"
    }

    $frontMatterTitles = @($TocTitle, $FigureListTitle, $TableListTitle)
    $start = 0
    $scanLimit = [Math]::Min(250, [int]$Document.Paragraphs.Count)
    for ($i = 1; $i -le $scanLimit; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        $text = Get-ParagraphText $paragraph
        if ($frontMatterTitles -contains $text) {
            $start = [int]$paragraph.Range.End
        }
    }

    return $start
}

$source = Get-Item -LiteralPath $Path
if ($source.PSIsContainer) {
    throw "Path must be a Word document, not a directory."
}

$output = Copy-DocumentForUpdate -Document $source -Suffix $CopySuffix
Write-Host ("Source: {0}" -f $source.FullName)
Write-Host ("Output: {0}" -f $output.FullName)

$wordType = [type]::GetTypeFromProgID("Word.Application")
if ($null -eq $wordType) {
    throw "Microsoft Word COM automation is not available on this machine."
}

$word = $null
$doc = $null
try {
    $word = [Activator]::CreateInstance($wordType)
    $word.Visible = [bool]$Visible
    $word.DisplayAlerts = 0
    try { $word.ScreenUpdating = [bool]$Visible } catch { }

    $doc = $word.Documents.Open($output.FullName, $false, $false, $false)
    $startPosition = Find-StartPosition -Document $doc

    $counts = @{ 1 = 0; 2 = 0; 3 = 0 }
    for ($i = 1; $i -le [int]$doc.Paragraphs.Count; $i++) {
        $paragraph = $doc.Paragraphs.Item($i)
        if ([int]$paragraph.Range.Start -lt $startPosition) {
            continue
        }

        $text = Get-ParagraphText $paragraph
        $level = Get-HeadingLevel -Text $text
        if ($level -eq 0) {
            continue
        }

        if ($level -eq 1) {
            $paragraph.Range.Style = $doc.Styles.Item($WdStyleHeading1)
            $paragraph.Range.ParagraphFormat.OutlineLevel = $WdOutlineLevel1
        }
        elseif ($level -eq 2) {
            $paragraph.Range.Style = $doc.Styles.Item($WdStyleHeading2)
            $paragraph.Range.ParagraphFormat.OutlineLevel = $WdOutlineLevel2
        }
        elseif ($level -eq 3) {
            $paragraph.Range.Style = $doc.Styles.Item($WdStyleHeading3)
            $paragraph.Range.ParagraphFormat.OutlineLevel = $WdOutlineLevel3
        }

        $counts[$level]++
    }

    $doc.Save()
    Write-Host ("Heading 1 styled: {0}" -f $counts[1])
    Write-Host ("Heading 2 styled: {0}" -f $counts[2])
    Write-Host ("Heading 3 styled: {0}" -f $counts[3])
    Write-Host "Done"
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close(-1) } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
