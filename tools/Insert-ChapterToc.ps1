param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [string]$CopySuffix = ".with-toc",
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Initialize-Unicode.ps1")

$WdFieldEmpty = -1
$WdStyleHeading1 = -2

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

function Find-ParagraphByExactText {
    param(
        $Document,
        [string]$Text,
        [int]$StartIndex = 1,
        [int]$EndIndex = 0
    )

    if ($EndIndex -le 0 -or $EndIndex -gt [int]$Document.Paragraphs.Count) {
        $EndIndex = [int]$Document.Paragraphs.Count
    }

    for ($i = $StartIndex; $i -le $EndIndex; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        if ((Get-ParagraphText $paragraph) -eq $Text) {
            return $paragraph
        }
    }

    return $null
}

function Find-ParagraphIndexByExactText {
    param(
        $Document,
        [string]$Text,
        [int]$StartIndex = 1,
        [int]$EndIndex = 0
    )

    if ($EndIndex -le 0 -or $EndIndex -gt [int]$Document.Paragraphs.Count) {
        $EndIndex = [int]$Document.Paragraphs.Count
    }

    for ($i = $StartIndex; $i -le $EndIndex; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        if ((Get-ParagraphText $paragraph) -eq $Text) {
            return $i
        }
    }

    return 0
}

function Delete-ParagraphsBetweenIndexes {
    param(
        $Document,
        [int]$StartIndex,
        [int]$EndIndex
    )

    for ($i = $EndIndex - 1; $i -gt $StartIndex; $i--) {
        try {
            [void]$Document.Paragraphs.Item($i).Range.Delete()
        }
        catch {
        }
    }
}

function Insert-TocAfterParagraph {
    param(
        $Document,
        $Paragraph
    )

    $insertAt = [int]$Paragraph.Range.End
    $insertRange = $Document.Range($insertAt, $insertAt)
    $insertRange.InsertAfter("`r")
    $fieldRange = $Document.Range($insertAt, $insertAt)
    $field = $Document.Fields.Add($fieldRange, $WdFieldEmpty, 'TOC \o "1-3" \h \z \u', $false)
    try {
        [void]$field.Update()
    }
    catch {
    }
}

function Ensure-FrontMatterHeading {
    param(
        $Document,
        $Paragraph
    )

    if ($null -eq $Paragraph) {
        return
    }

    try {
        $Paragraph.Range.Style = $Document.Styles.Item($WdStyleHeading1)
    }
    catch {
    }
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
    $tocIndex = Find-ParagraphIndexByExactText -Document $doc -Text $TocTitle -StartIndex 1 -EndIndex 50
    $figureIndex = Find-ParagraphIndexByExactText -Document $doc -Text $FigureListTitle -StartIndex 1 -EndIndex 180

    if ($tocIndex -eq 0 -or $figureIndex -eq 0) {
        throw "Could not find the TOC title and the next front-matter title."
    }

    Delete-ParagraphsBetweenIndexes -Document $doc -StartIndex $tocIndex -EndIndex $figureIndex
    $tocTitle = Find-ParagraphByExactText -Document $doc -Text $TocTitle -StartIndex 1 -EndIndex 50
    if ($null -eq $tocTitle) {
        throw "Could not find the TOC title after replacing the typed directory entries."
    }
    $figureTitle = Find-ParagraphByExactText -Document $doc -Text $FigureListTitle -StartIndex 1 -EndIndex 180
    $tableTitle = Find-ParagraphByExactText -Document $doc -Text $TableListTitle -StartIndex 1 -EndIndex 220
    Ensure-FrontMatterHeading -Document $doc -Paragraph $tocTitle
    Ensure-FrontMatterHeading -Document $doc -Paragraph $figureTitle
    Ensure-FrontMatterHeading -Document $doc -Paragraph $tableTitle
    Insert-TocAfterParagraph -Document $doc -Paragraph $tocTitle

    try {
        [void]$doc.Repaginate()
        for ($i = 1; $i -le [int]$doc.TablesOfContents.Count; $i++) {
            $toc = $doc.TablesOfContents.Item($i)
            [void]$toc.Update()
            [void]$toc.UpdatePageNumbers()
        }
    }
    catch {
    }

    $doc.Save()
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
