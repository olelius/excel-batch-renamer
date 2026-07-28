$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$artifactRoot = (Resolve-Path (Join-Path $projectRoot ".artifacts")).Path
$releaseRoot = Join-Path $artifactRoot "release"
$extractRoot = Join-Path $artifactRoot (
    "zip-smoke-" + [Guid]::NewGuid().ToString("N")
)
$testExtract = Join-Path $extractRoot "test"
$releaseExtract = Join-Path $extractRoot "release"

New-Item -ItemType Directory -Force -Path $testExtract, $releaseExtract |
    Out-Null

try {
    Expand-Archive `
        -LiteralPath (
            Join-Path $releaseRoot "ExcelBatchRenamer-Test-win7-x64.zip"
        ) `
        -DestinationPath $testExtract
    Expand-Archive `
        -LiteralPath (
            Join-Path $releaseRoot "ExcelBatchRenamer-win7-x64.zip"
        ) `
        -DestinationPath $releaseExtract

    $testExecutable = Join-Path $testExtract (
        "ExcelBatchRenamer-Test\ExcelBatchRenamer-Test.exe"
    )
    $releaseExecutable = Join-Path $releaseExtract (
        "ExcelBatchRenamer\ExcelBatchRenamer.exe"
    )

    & $testExecutable --smoke-test
    if ($LASTEXITCODE -ne 0) {
        throw "Extracted console build failed with exit code $LASTEXITCODE"
    }

    $releaseProcess = Start-Process `
        -FilePath $releaseExecutable `
        -ArgumentList "--smoke-test" `
        -Wait `
        -PassThru `
        -WindowStyle Hidden
    if ($releaseProcess.ExitCode -ne 0) {
        throw (
            "Extracted windowed build failed with exit code " +
            $releaseProcess.ExitCode
        )
    }

    Write-Output "Both extracted release archives passed smoke tests."
}
finally {
    if (Test-Path -LiteralPath $extractRoot -PathType Container) {
        $resolvedExtractRoot = (Resolve-Path -LiteralPath $extractRoot).Path
        $expectedPrefix = $artifactRoot.TrimEnd("\") + "\zip-smoke-"
        if (-not $resolvedExtractRoot.StartsWith(
            $expectedPrefix,
            [StringComparison]::OrdinalIgnoreCase
        )) {
            throw "Refusing to remove a path outside the expected artifact scope."
        }
        Remove-Item -LiteralPath $resolvedExtractRoot -Recurse -Force
    }
}
