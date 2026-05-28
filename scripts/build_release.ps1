[CmdletBinding()]
param(
    [string]$Version = "1.0.0",
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"

function Invoke-NativeCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Comanda a eșuat cu codul ${LASTEXITCODE}: $FilePath $($Arguments -join ' ')"
    }
}

function Resolve-InnoSetupCompiler {
    $command = Get-Command iscc -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $candidates = @(
        "$Env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "${Env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$Env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    return $null
}

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$IconSource = Join-Path $ProjectRoot "letter-s2.png"
$IconOutput = Join-Path $ProjectRoot "build\transareapp.ico"
$SpecFile = Join-Path $ProjectRoot "TransareApp.spec"
$InstallerScript = Join-Path $ProjectRoot "installer\TransareApp.iss"

if (-not (Test-Path $PythonExe)) {
    throw "Nu găsesc interpretul Python din .venv: $PythonExe"
}

if (-not (Test-Path $IconSource)) {
    throw "Nu găsesc imaginea pentru iconiță: $IconSource"
}

Write-Host "Instalez dependențele de runtime și build..."
Invoke-NativeCommand $PythonExe -m pip install --disable-pip-version-check -r (Join-Path $ProjectRoot "requirements.txt")
Invoke-NativeCommand $PythonExe -m pip install --disable-pip-version-check -r (Join-Path $ProjectRoot "requirements-build.txt")

Write-Host "Generez iconița ICO..."
Invoke-NativeCommand $PythonExe (Join-Path $ProjectRoot "scripts\png_to_ico.py") $IconSource $IconOutput

Write-Host "Construiesc distribuția PyInstaller..."
Invoke-NativeCommand $PythonExe -m PyInstaller --noconfirm --clean $SpecFile

if ($SkipInstaller) {
    Write-Host "Installer-ul a fost sărit la cerere. Executabilul este în dist\TransareApp\."
    exit 0
}

$IsccExe = Resolve-InnoSetupCompiler
if (-not $IsccExe) {
    Write-Warning "Inno Setup nu este instalat. Executabilul este în dist\TransareApp\."
    Write-Warning "Instalează Inno Setup 6 și rulează din nou scriptul pentru a genera setup.exe."
    exit 0
}

Write-Host "Construiesc installer-ul Inno Setup..."
Invoke-NativeCommand $IsccExe "/DMyAppVersion=$Version" $InstallerScript

Write-Host "Build finalizat. Installer-ul este în dist\installer\."
