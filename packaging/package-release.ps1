$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$portableRoot = Join-Path $projectRoot ".artifacts\portable"
$releaseRoot = Join-Path $projectRoot ".artifacts\release"
$testDirectory = Join-Path $portableRoot "ExcelBatchRenamer-Test"
$releaseDirectory = Join-Path $portableRoot "ExcelBatchRenamer"
$testArchive = Join-Path $releaseRoot "ExcelBatchRenamer-Test-win7-x64.zip"
$releaseArchive = Join-Path $releaseRoot "ExcelBatchRenamer-win7-x64.zip"
$checksumFile = Join-Path $releaseRoot "SHA256SUMS.txt"

foreach ($requiredDirectory in @($testDirectory, $releaseDirectory)) {
    if (-not (Test-Path -LiteralPath $requiredDirectory -PathType Container)) {
        throw "Portable directory not found: $requiredDirectory"
    }
}

New-Item -ItemType Directory -Force -Path $releaseRoot | Out-Null

Compress-Archive `
    -LiteralPath $testDirectory `
    -DestinationPath $testArchive `
    -CompressionLevel Optimal `
    -Force
Compress-Archive `
    -LiteralPath $releaseDirectory `
    -DestinationPath $releaseArchive `
    -CompressionLevel Optimal `
    -Force

$checksumLines = foreach ($archive in @($testArchive, $releaseArchive)) {
    $hash = Get-FileHash -LiteralPath $archive -Algorithm SHA256
    "$($hash.Hash.ToLowerInvariant())  $([IO.Path]::GetFileName($archive))"
}
$checksumLines | Set-Content -LiteralPath $checksumFile -Encoding ASCII

Get-Item -LiteralPath $testArchive, $releaseArchive, $checksumFile |
    Select-Object Name, Length, LastWriteTime
