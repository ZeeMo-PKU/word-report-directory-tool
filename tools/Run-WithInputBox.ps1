Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Initialize-Unicode.ps1")

Add-Type -AssemblyName System.Windows.Forms

$logPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "WordReportTool-run-log.txt"

function Show-Message {
    param(
        [string]$Text,
        [string]$Title = "Word Report Tool"
    )
    [System.Windows.Forms.MessageBox]::Show($Text, $Title) | Out-Null
}

function Get-LastLogLines {
    param(
        [string]$Path,
        [int]$Count = 12
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return ""
    }

    return ((Get-Content -LiteralPath $Path -Encoding UTF8 -Tail $Count) -join "`n")
}

try {
    "==== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====" | Set-Content -LiteralPath $logPath -Encoding UTF8

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select a Word .docx file"
    $dialog.Filter = "Word document (*.docx)|*.docx"
    $dialog.Multiselect = $false

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Show-Message "No file selected. Canceled."
        exit 0
    }

    $inputPath = $dialog.FileName
    "input=$inputPath" | Add-Content -LiteralPath $logPath -Encoding UTF8

    if (-not (Test-Path -LiteralPath $inputPath)) {
        throw "File does not exist: $inputPath"
    }
    if ([System.IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne ".docx") {
        throw "Please select a .docx file."
    }

    $scriptPath = Join-Path $PSScriptRoot "Update-ReportDirectories.ps1"
    $output = & $scriptPath -Path $inputPath -NoBackup *>&1
    $outputLines = @($output | ForEach-Object { [string]$_ })
    $outputLines | Add-Content -LiteralPath $logPath -Encoding UTF8

    $updatedLine = $outputLines | Where-Object { $_ -like "updated=*" } | Select-Object -Last 1
    $updatedPath = if ($updatedLine) { [string]$updatedLine -replace "^updated=", "" } else { "" }

    if ($updatedPath -and (Test-Path -LiteralPath $updatedPath)) {
        Show-Message "Done.`n`nOutput file:`n$updatedPath"
        Start-Process -FilePath explorer.exe -ArgumentList ("/select,`"$updatedPath`"")
    }
    else {
        Show-Message "Finished, but output file was not found.`n`nSee log:`n$logPath"
    }
}
catch {
    "ERROR=$($_.Exception.Message)" | Add-Content -LiteralPath $logPath -Encoding UTF8
    $details = Get-LastLogLines -Path $logPath
    Show-Message "Failed:`n$($_.Exception.Message)`n`nDetails:`n$details`n`nLog:`n$logPath"
    exit 1
}
