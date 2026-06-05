param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [string]$TemplatePath,
    [string]$CopySuffix = ".toc",
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Initialize-Unicode.ps1")

$WdCollapseEnd = 0
$WdFieldEmpty = -1
$WdStyleNormal = -1
$WdStyleCaption = -35
$WdAlignParagraphLeft = 0
$WdAlignParagraphCenter = 1
$WdOutlineLevel1 = 1
$WdOrganizerObjectStyles = 0

function Join-Chars {
    param([int[]]$CodePoints)
    return -join ($CodePoints | ForEach-Object { [char]$_ })
}

$TocTitle = Join-Chars @(0x76EE, 0x5F55)
$FigureListTitle = Join-Chars @(0x63D2, 0x56FE, 0x6E05, 0x5355)
$TableListTitle = Join-Chars @(0x9644, 0x8868, 0x6E05, 0x5355)
$FigureLabel = Join-Chars @(0x56FE)
$TableLabel = Join-Chars @(0x8868)

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

function Find-FirstLevelOneHeadingAfter {
    param(
        $Document,
        [int]$AfterStart
    )

    for ($i = 1; $i -le [int]$Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        if ([int]$paragraph.Range.Start -le $AfterStart) {
            continue
        }

        $text = Get-ParagraphText $paragraph
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        try {
            if ([int]$paragraph.Range.ParagraphFormat.OutlineLevel -eq $WdOutlineLevel1) {
                return $paragraph
            }
        }
        catch {
        }

        if ($text -match '^\d+(\.\d+)*\s+') {
            return $paragraph
        }
    }

    return $null
}

function Delete-BetweenParagraphs {
    param(
        $Document,
        $StartParagraph,
        $EndParagraph
    )

    $start = [int]$StartParagraph.Range.End
    $end = [int]$EndParagraph.Range.Start
    if ($end -gt $start) {
        $range = $Document.Range($start, $end)
        [void]$range.Delete()
    }
}

function Format-FrontMatterTitle {
    param($Document, $Paragraph)

    try {
        $Paragraph.Range.Style = $Document.Styles.Item($WdStyleNormal)
    }
    catch {
    }

    try {
        $Paragraph.Range.ParagraphFormat.Alignment = $WdAlignParagraphCenter
        $Paragraph.Range.Font.Bold = $true
    }
    catch {
    }
}

function Insert-FieldAfterParagraph {
    param(
        $Document,
        $Paragraph,
        [string]$FieldCode
    )

    $insertAt = [int]$Paragraph.Range.End
    $insertRange = $Document.Range($insertAt, $insertAt)
    $insertRange.InsertAfter("`r")
    $fieldRange = $Document.Range($insertAt, $insertAt)
    $field = $Document.Fields.Add($fieldRange, $WdFieldEmpty, $FieldCode, $false)
    try {
        [void]$field.Update()
    }
    catch {
    }
    return $field
}

function Convert-CaptionParagraph {
    param(
        $Document,
        $Paragraph,
        [string]$Label,
        [string]$CaptionText
    )

    if ($Paragraph.Range.Fields.Count -gt 0) {
        for ($i = 1; $i -le [int]$Paragraph.Range.Fields.Count; $i++) {
            $code = [string]$Paragraph.Range.Fields.Item($i).Code.Text
            if ($code -match ('SEQ\s+' + [regex]::Escape($Label))) {
                return
            }
        }
    }

    $paragraphStart = [int]$Paragraph.Range.Start
    $paragraphEnd = [int]$Paragraph.Range.End
    if ($paragraphEnd -le $paragraphStart) {
        return
    }

    $textRange = $Document.Range($paragraphStart, $paragraphEnd - 1)
    $textRange.Text = ("{0}  {1}" -f $Label, $CaptionText)

    $fieldAt = $paragraphStart + $Label.Length + 1
    $fieldRange = $Document.Range($fieldAt, $fieldAt)
    $fieldCode = ('SEQ {0} \* ARABIC' -f $Label)
    $field = $Document.Fields.Add($fieldRange, $WdFieldEmpty, $fieldCode, $false)

    try {
        [void]$field.Update()
    }
    catch {
    }

    try {
        $Paragraph.Range.Style = $Document.Styles.Item($WdStyleCaption)
    }
    catch {
    }

    try {
        $Paragraph.Range.ParagraphFormat.Alignment = $WdAlignParagraphCenter
    }
    catch {
    }
}

function Convert-CaptionsAfter {
    param(
        $Document,
        [int]$StartPosition
    )

    $convertedFigures = 0
    $convertedTables = 0

    for ($i = 1; $i -le [int]$Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        if ([int]$paragraph.Range.Start -lt $StartPosition) {
            continue
        }

        $text = Get-ParagraphText $paragraph
        if ($text -match ('^' + [regex]::Escape($FigureLabel) + '\s+\d+(?:\.\d+)*\s+(.+)$')) {
            Convert-CaptionParagraph -Document $Document -Paragraph $paragraph -Label $FigureLabel -CaptionText $Matches[1].Trim()
            $convertedFigures++
        }
        elseif ($text -match ('^' + [regex]::Escape($TableLabel) + '\s+\d+(?:\.\d+)*\s+(.+)$')) {
            Convert-CaptionParagraph -Document $Document -Paragraph $paragraph -Label $TableLabel -CaptionText $Matches[1].Trim()
            $convertedTables++
        }
    }

    return [pscustomobject]@{
        Figures = $convertedFigures
        Tables = $convertedTables
    }
}

function Copy-TemplateStyles {
    param(
        $Word,
        [string]$Template,
        [string]$Destination
    )

    if ([string]::IsNullOrWhiteSpace($Template)) {
        return
    }
    if (-not (Test-Path -LiteralPath $Template)) {
        throw "Template not found: $Template"
    }

    foreach ($styleName in @("TOC 1", "TOC 2", "TOC 3")) {
        try {
            $Word.OrganizerCopy($Template, $Destination, $styleName, $WdOrganizerObjectStyles)
        }
        catch {
        }
    }
}

function Update-FieldsInDocument {
    param($Document)

    try {
        [void]$Document.Repaginate()
    }
    catch {
    }

    try {
        [void]$Document.Fields.Update()
    }
    catch {
    }

    try {
        for ($i = 1; $i -le [int]$Document.TablesOfContents.Count; $i++) {
            $toc = $Document.TablesOfContents.Item($i)
            [void]$toc.Update()
            [void]$toc.UpdatePageNumbers()
        }
    }
    catch {
    }

    try {
        for ($i = 1; $i -le [int]$Document.TablesOfFigures.Count; $i++) {
            [void]$Document.TablesOfFigures.Item($i).Update()
        }
    }
    catch {
    }

    try {
        [void]$Document.Repaginate()
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

    Copy-TemplateStyles -Word $word -Template $TemplatePath -Destination $output.FullName

    $doc = $word.Documents.Open($output.FullName, $false, $false, $false)

    $tocTitle = Find-ParagraphByExactText -Document $doc -Text $TocTitle -StartIndex 1 -EndIndex 50
    $figureTitle = Find-ParagraphByExactText -Document $doc -Text $FigureListTitle -StartIndex 1 -EndIndex 180
    $tableTitle = Find-ParagraphByExactText -Document $doc -Text $TableListTitle -StartIndex 1 -EndIndex 220

    if ($null -eq $tocTitle -or $null -eq $figureTitle -or $null -eq $tableTitle) {
        throw "Could not find the front-matter titles for TOC, figure list, and table list."
    }

    $bodyFirstHeading = Find-FirstLevelOneHeadingAfter -Document $doc -AfterStart ([int]$tableTitle.Range.Start)
    if ($null -eq $bodyFirstHeading) {
        throw "Could not find the first body heading after the table list."
    }

    Delete-BetweenParagraphs -Document $doc -StartParagraph $tableTitle -EndParagraph $bodyFirstHeading
    $figureTitle = Find-ParagraphByExactText -Document $doc -Text $FigureListTitle -StartIndex 1 -EndIndex 180
    $tableTitle = Find-ParagraphByExactText -Document $doc -Text $TableListTitle -StartIndex 1 -EndIndex 220
    Delete-BetweenParagraphs -Document $doc -StartParagraph $figureTitle -EndParagraph $tableTitle
    $tocTitle = Find-ParagraphByExactText -Document $doc -Text $TocTitle -StartIndex 1 -EndIndex 50
    $figureTitle = Find-ParagraphByExactText -Document $doc -Text $FigureListTitle -StartIndex 1 -EndIndex 180
    Delete-BetweenParagraphs -Document $doc -StartParagraph $tocTitle -EndParagraph $figureTitle

    $tocTitle = Find-ParagraphByExactText -Document $doc -Text $TocTitle -StartIndex 1 -EndIndex 50
    $figureTitle = Find-ParagraphByExactText -Document $doc -Text $FigureListTitle -StartIndex 1 -EndIndex 180
    $tableTitle = Find-ParagraphByExactText -Document $doc -Text $TableListTitle -StartIndex 1 -EndIndex 220
    $bodyFirstHeading = Find-FirstLevelOneHeadingAfter -Document $doc -AfterStart ([int]$tableTitle.Range.Start)

    foreach ($title in @($tocTitle, $figureTitle, $tableTitle)) {
        Format-FrontMatterTitle -Document $doc -Paragraph $title
    }

    $captionCounts = Convert-CaptionsAfter -Document $doc -StartPosition ([int]$bodyFirstHeading.Range.Start)

    $tableCode = 'TOC \h \z \c "' + $TableLabel + '"'
    $figureCode = 'TOC \h \z \c "' + $FigureLabel + '"'
    $chapterCode = 'TOC \o "1-3" \h \z \u'

    Insert-FieldAfterParagraph -Document $doc -Paragraph $tableTitle -FieldCode $tableCode | Out-Null
    $figureTitle = Find-ParagraphByExactText -Document $doc -Text $FigureListTitle -StartIndex 1 -EndIndex 180
    Insert-FieldAfterParagraph -Document $doc -Paragraph $figureTitle -FieldCode $figureCode | Out-Null
    $tocTitle = Find-ParagraphByExactText -Document $doc -Text $TocTitle -StartIndex 1 -EndIndex 50
    Insert-FieldAfterParagraph -Document $doc -Paragraph $tocTitle -FieldCode $chapterCode | Out-Null

    Update-FieldsInDocument -Document $doc
    $doc.Save()

    Write-Host ("Converted figure captions: {0}" -f $captionCounts.Figures)
    Write-Host ("Converted table captions: {0}" -f $captionCounts.Tables)
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
