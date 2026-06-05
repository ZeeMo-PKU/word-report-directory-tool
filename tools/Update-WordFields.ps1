param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias("FullName")]
    [string[]]$Path,

    [switch]$NoBackup,
    [switch]$Visible,
    [switch]$NoOpenUpdateFlag,
    [string]$CopySuffix
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Initialize-Unicode.ps1")

function Resolve-InputDocuments {
    param([string[]]$InputPaths)

    foreach ($inputPath in $InputPaths) {
        if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($inputPath)) {
            Get-ChildItem -Path $inputPath -File |
                Where-Object { $_.Extension -match '^\.docx?$|^\.docm$|^\.dotx$|^\.dotm$' -and -not $_.Name.StartsWith('~$') }
        }
        else {
            $item = Get-Item -LiteralPath $inputPath
            if ($item.PSIsContainer) {
                Get-ChildItem -LiteralPath $item.FullName -File -Filter "*.docx" |
                    Where-Object { -not $_.Name.StartsWith('~$') }
            }
            else {
                if ($item.Name.StartsWith('~$')) {
                    continue
                }
                $item
            }
        }
    }
}

function Backup-Document {
    param([System.IO.FileInfo]$Document)

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupName = "{0}.before-field-update.{1}{2}" -f $Document.BaseName, $stamp, $Document.Extension
    $backupPath = Join-Path $Document.DirectoryName $backupName
    Copy-Item -LiteralPath $Document.FullName -Destination $backupPath -Force
    return $backupPath
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

function Update-FieldCollection {
    param($Fields)

    if ($null -eq $Fields) {
        return
    }

    try {
        $count = [int]$Fields.Count
    }
    catch {
        return
    }

    for ($i = 1; $i -le $count; $i++) {
        try {
            $field = $Fields.Item($i)
            $field.Locked = $false
            [void]$field.Update()
        }
        catch {
            # Some fields can fail while their surrounding collection still updates.
        }
    }

    try {
        [void]$Fields.Update()
    }
    catch {
    }
}

function Update-FieldsInStoryRanges {
    param($Document)

    try {
        $storyRanges = $Document.StoryRanges
        for ($i = 1; $i -le [int]$storyRanges.Count; $i++) {
            $range = $storyRanges.Item($i)
            while ($null -ne $range) {
                Update-FieldCollection -Fields $range.Fields
                try {
                    $range = $range.NextStoryRange
                }
                catch {
                    $range = $null
                }
            }
        }
    }
    catch {
    }
}

function Update-ShapeTextFields {
    param($Shapes)

    if ($null -eq $Shapes) {
        return
    }

    try {
        $count = [int]$Shapes.Count
    }
    catch {
        return
    }

    for ($i = 1; $i -le $count; $i++) {
        try {
            $shape = $Shapes.Item($i)
            if ($shape.TextFrame.HasText -ne 0) {
                Update-FieldCollection -Fields $shape.TextFrame.TextRange.Fields
            }
        }
        catch {
        }

        try {
            if ([int]$shape.GroupItems.Count -gt 0) {
                Update-ShapeTextFields -Shapes $shape.GroupItems
            }
        }
        catch {
        }
    }
}

function Update-SpecialTables {
    param($Document)

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
            [void]$Document.TablesOfFigures.Item($i).Update()
        }
    }
    catch {
    }

    try {
        for ($i = 1; $i -le [int]$Document.TablesOfAuthorities.Count; $i++) {
            [void]$Document.TablesOfAuthorities.Item($i).Update()
        }
    }
    catch {
    }

    try {
        for ($i = 1; $i -le [int]$Document.Indexes.Count; $i++) {
            [void]$Document.Indexes.Item($i).Update()
        }
    }
    catch {
    }
}

function Update-AllDocumentFields {
    param($Document)

    try { [void]$Document.Repaginate() } catch { }
    Update-FieldsInStoryRanges -Document $Document
    Update-ShapeTextFields -Shapes $Document.Shapes

    try {
        for ($i = 1; $i -le [int]$Document.Sections.Count; $i++) {
            $section = $Document.Sections.Item($i)
            for ($j = 1; $j -le [int]$section.Headers.Count; $j++) {
                Update-ShapeTextFields -Shapes $section.Headers.Item($j).Shapes
            }
            for ($j = 1; $j -le [int]$section.Footers.Count; $j++) {
                Update-ShapeTextFields -Shapes $section.Footers.Item($j).Shapes
            }
        }
    }
    catch {
    }

    Update-SpecialTables -Document $Document
    try { [void]$Document.Repaginate() } catch { }
    Update-SpecialTables -Document $Document
}

function Enable-UpdateFieldsOnOpen {
    param([string]$DocxPath)

    $extension = [System.IO.Path]::GetExtension($DocxPath).ToLowerInvariant()
    if ($extension -notin @(".docx", ".docm", ".dotx", ".dotm")) {
        return
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("word-fields-" + [guid]::NewGuid().ToString("N") + $extension)
    Copy-Item -LiteralPath $DocxPath -Destination $tempPath -Force

    $zip = [System.IO.Compression.ZipFile]::Open($tempPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry("word/settings.xml")
        if ($null -eq $entry) {
            return
        }

        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $xmlText = $reader.ReadToEnd()
        $reader.Dispose()

        [xml]$xml = $xmlText
        $nsUri = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("w", $nsUri)

        $settings = $xml.SelectSingleNode("/w:settings", $ns)
        if ($null -eq $settings) {
            return
        }

        $updateFields = $xml.SelectSingleNode("/w:settings/w:updateFields", $ns)
        if ($null -eq $updateFields) {
            $updateFields = $xml.CreateElement("w", "updateFields", $nsUri)
            [void]$settings.AppendChild($updateFields)
        }

        $valAttr = $xml.CreateAttribute("w", "val", $nsUri)
        $valAttr.Value = "true"
        [void]$updateFields.Attributes.SetNamedItem($valAttr)

        $entry.Delete()
        $newEntry = $zip.CreateEntry("word/settings.xml")
        $writeStream = $newEntry.Open()
        $writer = New-Object System.IO.StreamWriter($writeStream, [System.Text.Encoding]::UTF8)
        $writer.Write($xml.OuterXml)
        $writer.Dispose()
    }
    finally {
        $zip.Dispose()
    }

    Move-Item -LiteralPath $tempPath -Destination $DocxPath -Force
}

$documents = @(Resolve-InputDocuments -InputPaths $Path)
if ($documents.Count -eq 0) {
    throw "No Word documents matched the input path."
}

$wordType = [type]::GetTypeFromProgID("Word.Application")
if ($null -eq $wordType) {
    throw "Microsoft Word COM automation is not available on this machine."
}

$word = $null
try {
    $word = [Activator]::CreateInstance($wordType)
    $word.Visible = [bool]$Visible
    $word.DisplayAlerts = 0
    try { $word.ScreenUpdating = [bool]$Visible } catch { }

    foreach ($sourceDocument in $documents) {
        $document = Copy-DocumentForUpdate -Document $sourceDocument -Suffix $CopySuffix
        if ($document.FullName -ne $sourceDocument.FullName) {
            Write-Host ("Copied source: {0}" -f $sourceDocument.FullName)
            Write-Host ("Updating copy: {0}" -f $document.FullName)
        }
        else {
            Write-Host ("Updating fields: {0}" -f $document.FullName)
        }

        if ((-not $NoBackup) -and [string]::IsNullOrWhiteSpace($CopySuffix)) {
            $backupPath = Backup-Document -Document $document
            Write-Host ("  Backup: {0}" -f $backupPath)
        }

        $doc = $null
        try {
            $doc = $word.Documents.Open($document.FullName, $false, $false, $false)
            Update-AllDocumentFields -Document $doc
            $doc.Save()
            $doc.Close(-1)
            $doc = $null

            if (-not $NoOpenUpdateFlag) {
                Enable-UpdateFieldsOnOpen -DocxPath $document.FullName
            }
            Write-Host "  Done"
        }
        catch {
            if ($null -ne $doc) {
                try { $doc.Close(0) } catch { }
            }
            throw
        }
    }
}
finally {
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
