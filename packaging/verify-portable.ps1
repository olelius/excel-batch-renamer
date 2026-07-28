$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$distPath = Join-Path $projectRoot ".artifacts\portable"
$testDirectory = Join-Path $distPath "ExcelBatchRenamer-Test"
$releaseDirectory = Join-Path $distPath "ExcelBatchRenamer"
$testExecutable = Join-Path $testDirectory "ExcelBatchRenamer-Test.exe"
$releaseExecutable = Join-Path $releaseDirectory "ExcelBatchRenamer.exe"
$ucrtRedistRoot = Join-Path $projectRoot (
    ".tools\windows-sdk\ucrt-extracted\Windows Kits\10\Redist"
)
$ucrtVersionDirectory = Get-ChildItem `
    -LiteralPath $ucrtRedistRoot `
    -Directory `
    -ErrorAction SilentlyContinue |
    Sort-Object Name -Descending |
    Select-Object -First 1
if ($null -eq $ucrtVersionDirectory) {
    throw "Project-local Windows SDK UCRT redistributables were not found."
}
$ucrtDirectory = Join-Path $ucrtVersionDirectory.FullName "ucrt\DLLs\x64"
$ucrtNames = @(
    Get-ChildItem -LiteralPath $ucrtDirectory -File -Filter "*.dll" |
        Select-Object -ExpandProperty Name
)

foreach ($requiredPath in @(
    $testExecutable,
    $releaseExecutable,
    (Join-Path $testDirectory "README.txt"),
    (Join-Path $releaseDirectory "README.txt"),
    (Join-Path $testDirectory "python38.dll"),
    (Join-Path $releaseDirectory "python38.dll"),
    (Join-Path $testDirectory "_tkinter.pyd"),
    (Join-Path $releaseDirectory "_tkinter.pyd"),
    (Join-Path $testDirectory "VCRUNTIME140.dll"),
    (Join-Path $releaseDirectory "VCRUNTIME140.dll")
)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Portable directory is missing a required file: $requiredPath"
    }
}

foreach ($portableDirectory in @($testDirectory, $releaseDirectory)) {
    foreach ($ucrtName in $ucrtNames) {
        $ucrtPath = Join-Path $portableDirectory $ucrtName
        if (-not (Test-Path -LiteralPath $ucrtPath -PathType Leaf)) {
            throw "Portable directory is missing UCRT file: $ucrtPath"
        }
    }
    foreach ($tkDirectoryName in @("tcl", "tcl8", "tk")) {
        $tkDirectory = Join-Path $portableDirectory $tkDirectoryName
        if (-not (Test-Path -LiteralPath $tkDirectory -PathType Container)) {
            throw "Portable directory is missing Tcl/Tk data: $tkDirectory"
        }
    }
}

& $testExecutable --smoke-test
if ($LASTEXITCODE -ne 0) {
    throw "Console build smoke test failed with exit code $LASTEXITCODE"
}

$releaseProcess = Start-Process `
    -FilePath $releaseExecutable `
    -ArgumentList "--smoke-test" `
    -Wait `
    -PassThru `
    -WindowStyle Hidden
if ($releaseProcess.ExitCode -ne 0) {
    throw "Windowed build smoke test failed with exit code $($releaseProcess.ExitCode)"
}

$testFileCount = (Get-ChildItem -LiteralPath $testDirectory -Recurse -File).Count
$releaseFileCount = (Get-ChildItem -LiteralPath $releaseDirectory -Recurse -File).Count
$testBytes = (
    Get-ChildItem -LiteralPath $testDirectory -Recurse -File |
        Measure-Object -Property Length -Sum
).Sum
$releaseBytes = (
    Get-ChildItem -LiteralPath $releaseDirectory -Recurse -File |
        Measure-Object -Property Length -Sum
).Sum

Write-Output "Console build: $testFileCount files, $testBytes bytes"
Write-Output "Windowed build: $releaseFileCount files, $releaseBytes bytes"
Write-Output "Both portable directories passed the current-machine smoke test."
