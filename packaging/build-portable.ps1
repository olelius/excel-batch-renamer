$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonPath = Join-Path $projectRoot ".venv\Scripts\python.exe"
$entryPath = Join-Path $projectRoot "src\excel_batch_renamer\app.py"
$sourcePath = Join-Path $projectRoot "src"
$readmePath = Join-Path $PSScriptRoot "README.txt"
$artifactRoot = Join-Path $projectRoot ".artifacts"
$distPath = Join-Path $artifactRoot "portable"
$specPath = Join-Path $artifactRoot "specs"
$ucrtRedistRoot = Join-Path $projectRoot (
    ".tools\windows-sdk\ucrt-extracted\Windows Kits\10\Redist"
)

if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    throw "Project virtual environment not found: $pythonPath"
}

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
$ucrtFiles = @(
    Get-ChildItem -LiteralPath $ucrtDirectory -File -Filter "*.dll"
)
if ($ucrtFiles.Count -lt 2) {
    throw "Project-local x64 UCRT redistributables are incomplete."
}

New-Item -ItemType Directory -Force -Path $distPath | Out-Null
New-Item -ItemType Directory -Force -Path $specPath | Out-Null

function Build-PortableApplication {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateSet("console", "windowed")]
        [string]$Mode
    )

    $workPath = Join-Path $artifactRoot "pyinstaller\$Name"
    $modeArgument = if ($Mode -eq "console") { "--console" } else { "--windowed" }

    $arguments = @(
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        $modeArgument,
        "--noupx",
        "--name",
        $Name,
        "--paths",
        $sourcePath,
        "--add-data",
        "$readmePath;.",
        "--distpath",
        $distPath,
        "--workpath",
        $workPath,
        "--specpath",
        $specPath
    )
    foreach ($ucrtFile in $ucrtFiles) {
        $arguments += @("--add-binary", "$($ucrtFile.FullName);.")
    }
    $arguments += $entryPath

    & $pythonPath @arguments

    if ($LASTEXITCODE -ne 0) {
        throw "$Name build failed with exit code $LASTEXITCODE"
    }
}

Push-Location $projectRoot
try {
    Build-PortableApplication -Name "ExcelBatchRenamer-Test" -Mode "console"
    Build-PortableApplication -Name "ExcelBatchRenamer" -Mode "windowed"
}
finally {
    Pop-Location
}

Write-Output "Portable directories generated: $distPath"
