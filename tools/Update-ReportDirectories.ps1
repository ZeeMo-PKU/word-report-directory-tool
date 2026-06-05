param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [string]$CopySuffix = ".with-directories",
    [string]$PythonPath = "",
    [switch]$NoBackup,
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Initialize-Unicode.ps1")

function Get-Python {
    param([string]$ExplicitPath)

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        return $ExplicitPath
    }

    $toolRoot = Split-Path -Parent $PSScriptRoot
    $localVenv = Join-Path $toolRoot ".venv\Scripts\python.exe"
    if (Test-Path -LiteralPath $localVenv) {
        return $localVenv
    }

    $py = Get-Command python -ErrorAction SilentlyContinue
    if ($null -ne $py) {
        return $py.Source
    }

    throw "Python was not found. Run Install-Dependencies.bat first, or pass -PythonPath C:\path\to\python.exe."
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
    if (Test-Path -LiteralPath $copyPath) {
        $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $copyName = "{0}{1}.{2}{3}" -f $Document.BaseName, $Suffix, $stamp, $Document.Extension
        $copyPath = Join-Path $Document.DirectoryName $copyName
    }
    Copy-Item -LiteralPath $Document.FullName -Destination $copyPath -Force
    return Get-Item -LiteralPath $copyPath
}

function Update-MainTocWithWord {
    param(
        [string]$DocumentPath,
        [switch]$ShowWord
    )

    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    $word.DisplayAlerts = 0
    try {
        $resolvedPath = (Get-Item -LiteralPath $DocumentPath).FullName
        $doc = $word.Documents.Open($resolvedPath, $false, $false)
        try {
            try {
                $count = [int]$doc.TablesOfContents.Count
            }
            catch {
                $count = 0
            }
            if ($count -gt 0) {
                [void]$doc.TablesOfContents.Item(1).Update()
            }
            $doc.Repaginate()
            if ($count -gt 0) {
                [void]$doc.TablesOfContents.Item(1).UpdatePageNumbers()
            }
            $doc.Save()
            Write-Output "main_toc_updated=True"
            Write-Output ("word_toc_count={0}" -f $count)
        }
        finally {
            $doc.Close($true)
        }
    }
    finally {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}

function Invoke-PythonReportScript {
    param(
        [string]$Python,
        [string[]]$Arguments
    )

    function Quote-ProcessArgument {
        param([string]$Argument)

        if ($null -eq $Argument) {
            return '""'
        }

        return '"' + ($Argument -replace '(\\*)"', '$1$1\"' -replace '(\\+)$', '$1$1') + '"'
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Python
    $psi.Arguments = (($Arguments | ForEach-Object { Quote-ProcessArgument $_ }) -join " ")
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8

    $process = [System.Diagnostics.Process]::Start($psi)
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        Write-Output (($stdout -split "`r?`n") | Where-Object { $_ -ne "" })
    }

    if ($process.ExitCode -ne 0) {
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            Write-Output (($stderr -split "`r?`n") | Where-Object { $_ -ne "" })
        }
        throw "Python step failed with exit code $($process.ExitCode)."
    }

    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        Write-Output (($stderr -split "`r?`n") | Where-Object { $_ -ne "" })
    }
}

$document = Get-Item -LiteralPath $Path
if ($document.PSIsContainer) {
    throw "Path must be a .docx file, not a directory."
}
if ($document.Name.StartsWith("~$")) {
    throw "Refusing to process a temporary Word lock file: $($document.FullName)"
}

$target = Copy-DocumentForUpdate -Document $document -Suffix $CopySuffix
$python = Get-Python -ExplicitPath $PythonPath
$script = Join-Path $PSScriptRoot "report_directories.py"

$prepareArgs = @($script, $target.FullName, "--phase", "prepare")
if ($NoBackup) {
    $prepareArgs += "--no-backup"
}

Invoke-PythonReportScript -Python $python -Arguments $prepareArgs
Update-MainTocWithWord -DocumentPath $target.FullName -ShowWord:$Visible
Invoke-PythonReportScript -Python $python -Arguments @($script, $target.FullName, "--phase", "finalize")

Write-Output ("updated={0}" -f $target.FullName)
