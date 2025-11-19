param(
    [ValidateSet("Register", "Unregister")]
    [string]$Action = "Register",

    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Administrator rights are required. Re-run this script from an elevated PowerShell session."
}

$projectPath = Join-Path $PSScriptRoot "CropperShellExtension" "CropperShellExtension.csproj"
$publishDir = Join-Path $PSScriptRoot "CropperShellExtension" "bin" $Configuration "net8.0-windows" "win-x64" "publish"

Write-Host "Building CropperShellExtension ($Configuration)..." -ForegroundColor Cyan
dotnet publish $projectPath -c $Configuration -r win-x64 --no-self-contained

$comHost = Join-Path $publishDir "CropperShellExtension.comhost.dll"
if (-not (Test-Path $comHost)) {
    throw "Unable to locate comhost at $comHost."
}

$args = @()
if ($Action -eq "Unregister") {
    $args += "/u"
}

$args += "/s"
$args += "`"$comHost`""

Write-Host "$Action shell extension via regsvr32..." -ForegroundColor Cyan
$process = Start-Process -FilePath "regsvr32.exe" -ArgumentList $args -PassThru -Wait

if ($process.ExitCode -ne 0) {
    throw "regsvr32 exited with code $($process.ExitCode)."
}

Write-Host "regsvr32 completed successfully." -ForegroundColor Green
