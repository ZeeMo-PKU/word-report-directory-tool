$WordReportToolUtf8NoBom = New-Object System.Text.UTF8Encoding($false)

try {
    [Console]::InputEncoding = $WordReportToolUtf8NoBom
}
catch {
}

try {
    [Console]::OutputEncoding = $WordReportToolUtf8NoBom
}
catch {
}

$OutputEncoding = $WordReportToolUtf8NoBom
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
