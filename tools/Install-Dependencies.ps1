Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Initialize-Unicode.ps1")

Add-Type -AssemblyName System.Windows.Forms

function Show-Message {
    param(
        [string]$Text,
        [string]$Title = "Word Report Tool Setup"
    )
    [System.Windows.Forms.MessageBox]::Show($Text, $Title) | Out-Null
}

function Find-SystemPython {
    $candidates = @()

    $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
    if ($null -ne $pyLauncher) {
        $candidates += @{ Command = $pyLauncher.Source; Args = @("-3") }
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($null -ne $python) {
        $candidates += @{ Command = $python.Source; Args = @() }
    }

    foreach ($candidate in $candidates) {
        try {
            $version = & $candidate.Command @($candidate.Args + @("-c", "import sys; print(sys.version_info[:2])")) 2>$null
            if ($LASTEXITCODE -eq 0 -and $version) {
                return $candidate
            }
        }
        catch {
        }
    }

    throw "Python 3 was not found. Please install Python 3 from https://www.python.org/downloads/ and tick 'Add python.exe to PATH'."
}

$toolRoot = Split-Path -Parent $PSScriptRoot
$venvDir = Join-Path $toolRoot ".venv"
$venvPython = Join-Path $venvDir "Scripts\python.exe"

try {
    if (-not (Test-Path -LiteralPath $venvPython)) {
        $systemPython = Find-SystemPython
        & $systemPython.Command @($systemPython.Args + @("-m", "venv", $venvDir))
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to create the local Python environment."
        }
    }

    & $venvPython -m pip install --upgrade pip python-docx lxml
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to install Python dependencies."
    }

    Show-Message "Setup complete.`n`nYou can now double-click Run-WordReportTool.bat."
}
catch {
    Show-Message "Setup failed:`n$($_.Exception.Message)"
    exit 1
}
