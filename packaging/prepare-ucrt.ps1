$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$sdkRoot = Join-Path $projectRoot ".tools\windows-sdk"
$installerPath = Join-Path $sdkRoot "winsdksetup.exe"
$layoutPath = Join-Path $sdkRoot "layout"
$extractPath = Join-Path $sdkRoot "ucrt-extracted"
$installerUrl = (
    "https://download.microsoft.com/download/" +
    "e119c04b-71aa-4067-ac3c-360c2e13d209/" +
    "windowssdk/winsdksetup.exe"
)

New-Item -ItemType Directory -Force -Path $sdkRoot | Out-Null

if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
}

$signature = Get-AuthenticodeSignature -LiteralPath $installerPath
if ($signature.Status -ne "Valid") {
    throw "Windows SDK installer signature is not valid: $($signature.Status)"
}

New-Item -ItemType Directory -Force -Path $layoutPath | Out-Null
$layoutProcess = Start-Process `
    -FilePath $installerPath `
    -ArgumentList @(
        "/layout",
        $layoutPath,
        "/features",
        "OptionId.DesktopCPPx64",
        "/quiet",
        "/norestart",
        "/ceip",
        "off"
    ) `
    -Wait `
    -PassThru `
    -WindowStyle Hidden
if ($layoutProcess.ExitCode -ne 0) {
    throw "Windows SDK layout failed with exit code $($layoutProcess.ExitCode)"
}

$ucrtMsi = Join-Path (
    Join-Path $layoutPath "Installers"
) "Universal CRT Redistributable-x86_en-us.msi"
if (-not (Test-Path -LiteralPath $ucrtMsi -PathType Leaf)) {
    throw "Universal CRT Redistributable MSI was not downloaded."
}

New-Item -ItemType Directory -Force -Path $extractPath | Out-Null
$msiProcess = Start-Process `
    -FilePath "msiexec.exe" `
    -ArgumentList @(
        "/a",
        ('"' + $ucrtMsi + '"'),
        ('TARGETDIR="' + $extractPath + '"'),
        "/qn",
        "/norestart"
    ) `
    -Wait `
    -PassThru `
    -WindowStyle Hidden
if ($msiProcess.ExitCode -ne 0) {
    throw "UCRT extraction failed with exit code $($msiProcess.ExitCode)"
}

$ucrtBase = Get-ChildItem `
    -LiteralPath $extractPath `
    -Recurse `
    -File `
    -Filter "ucrtbase.dll" |
    Where-Object { $_.FullName -match "\\DLLs\\x64\\ucrtbase\.dll$" } |
    Select-Object -First 1
if ($null -eq $ucrtBase) {
    throw "Extracted x64 UCRT files were not found."
}

Write-Output "Project-local x64 UCRT files are ready: $($ucrtBase.DirectoryName)"
