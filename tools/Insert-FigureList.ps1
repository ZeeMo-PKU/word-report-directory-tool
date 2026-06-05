param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [string]$CopySuffix = ".with-figures",
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Initialize-Unicode.ps1")

$WdFieldEmpty = -1
$WdStyleCaption = -35
$WdAlignParagraphCenter = 1

function Join-Chars {
    param([int[]]$CodePoints)
    return -join ($CodePoints | ForEach-Object { [char]$_ })
}

$FigureListTitle = Join-Chars @(0x63D2, 0x56FE, 0x6E05, 0x5355)
$TableListTitle = Join-Chars @(0x9644, 0x8868, 0x6E05, 0x5355)
$FigureLabel = Join-Chars @(0x56FE)
$FirstBodyHeading = "1 " + (Join-Chars @(0x5F15, 0x8A00))

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

function Convert-FigureCaption {
    param(
        $Document,
        $Paragraph,
        [string]$CaptionText
    )

    $hasSeq = $false
    try {
        for ($i = 1; $i -le [int]$Paragraph.Range.Fields.Count; $i++) {
            $code = [string]$Paragraph.Range.Fields.Item($i).Code.Text
            if ($code -match ('SEQ\s+' + [regex]::Escape($FigureLabel))) {
                $hasSeq = $true
                break
            }
        }
    }
    catch {
    }

    if ($hasSeq) {
        return
    }

    $start = [int]$Paragraph.Range.Start
    $end = [int]$Paragraph.Range.End
    if ($end -le $start) {
        return
    }

    $textRange = $Document.Range($start, $end - 1)
    $textRange.Text = ("{0}  {1}" -f $FigureLabel, $CaptionText)

    $fieldAt = $start + $FigureLabel.Length + 1
    $fieldRange = $Document.Range($fieldAt, $fieldAt)
    $field = $Document.Fields.Add($fieldRange, $WdFieldEmpty, ('SEQ {0} \* ARABIC' -f $FigureLabel), $false)
    try { [void]$field.Update() } catch { }

    try {
        $Paragraph.Range.Style = $Document.Styles.Item($WdStyleCaption)
        $Paragraph.Range.ParagraphFormat.Alignment = $WdAlignParagraphCenter
    }
    catch {
    }
}

function Convert-FigureCaptionsAfter {
    param(
        $Document,
        [int]$StartIndex
    )

    $items = New-Object System.Collections.Generic.List[object]
    $pattern = '^' + [regex]::Escape($FigureLabel) + '\s+\d+(?:\.\d+)*\s+(.+)$'

    for ($i = $StartIndex; $i -le [int]$Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        $text = Get-ParagraphText $paragraph
        if ($text -match $pattern) {
            $captionText = $Matches[1].Trim()
            Convert-FigureCaption -Document $Document -Paragraph $paragraph -CaptionText $captionText
            $items.Add([pscustomobject]@{
                Paragraph = $i
                Page = $paragraph.Range.Information(3)
                Caption = ($FigureLabel + " " + ($items.Count + 1).ToString() + " " + $captionText)
            }) | Out-Null
        }
    }

    return $items
}

function Insert-FigureListAfterTitle {
    param(
        $Document,
        [int]$FigureTitleIndex
    )

    $paragraph = $Document.Paragraphs.Item($FigureTitleIndex)
    $insertAt = [int]$paragraph.Range.End
    $insertRange = $Document.Range($insertAt, $insertAt)
    $insertRange.InsertAfter("`r")
    $fieldRange = $Document.Range($insertAt, $insertAt)
    $fieldCode = 'TOC \h \z \c "' + $FigureLabel + '"'
    $field = $Document.Fields.Add($fieldRange, $WdFieldEmpty, $fieldCode, $false)
    try { [void]$field.Update() } catch { }
}

function Update-WordFields {
    param($Document)

    try { [void]$Document.Repaginate() } catch { }
    try { [void]$Document.Fields.Update() } catch { }

    try {
        for ($i = 1; $i -le [int]$Document.TablesOfContents.Count; $i++) {
            $toc = $Document.TablesOfContents.Item($i)
            try { [void]$toc.Update() } catch { }
            try { [void]$toc.UpdatePageNumbers() } catch { }
        }
    }
    catch {
    }

    try {
        for ($i = 1; $i -le [int]$Document.TablesOfFigures.Count; $i++) {
            try { [void]$Document.TablesOfFigures.Item($i).Update() } catch { }
        }
    }
    catch {
    }

    try { [void]$Document.Repaginate() } catch { }
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
    $figureIndex = Find-ParagraphIndexByExactText -Document $doc -Text $FigureListTitle -StartIndex 1 -EndIndex 220
    $tableIndex = Find-ParagraphIndexByExactText -Document $doc -Text $TableListTitle -StartIndex 1 -EndIndex 260
    $bodyStartIndex = Find-ParagraphIndexByExactText -Document $doc -Text $FirstBodyHeading -StartIndex 1

    if ($figureIndex -eq 0 -or $tableIndex -eq 0) {
        throw "Could not find the figure list title and table list title."
    }
    if ($bodyStartIndex -eq 0) {
        throw "Could not find the first body heading."
    }

    Delete-ParagraphsBetweenIndexes -Document $doc -StartIndex $figureIndex -EndIndex $tableIndex
    $figureIndex = Find-ParagraphIndexByExactText -Document $doc -Text $FigureListTitle -StartIndex 1 -EndIndex 220
    $bodyStartIndex = Find-ParagraphIndexByExactText -Document $doc -Text $FirstBodyHeading -StartIndex 1

    $items = Convert-FigureCaptionsAfter -Document $doc -StartIndex $bodyStartIndex
    Insert-FigureListAfterTitle -Document $doc -FigureTitleIndex $figureIndex
    Update-WordFields -Document $doc

    $doc.Save()
    Write-Host ("Converted figure captions: {0}" -f $items.Count)
    foreach ($item in $items) {
        Write-Host ("  p{0}, page {1}: {2}" -f $item.Paragraph, $item.Page, $item.Caption)
    }
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
