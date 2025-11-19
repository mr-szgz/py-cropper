param(
    [ValidateSet("Install", "Register", "Unregister")]
    [string]$Action = "Install",

    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$Force
)

$ErrorActionPreference = "Stop"

$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Administrator rights are required. Re-run this script from an elevated PowerShell session."
}

function Invoke-Regsvr32 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComHostPath,
        [switch]$Unregister,
        [switch]$IgnoreMissing
    )

    if (-not (Test-Path $ComHostPath)) {
        if ($IgnoreMissing) {
            Write-Host "Skipping regsvr32 $($Unregister ? "unregister" : "register") because $ComHostPath does not exist." -ForegroundColor Yellow
            return
        }

        throw "Unable to locate comhost at $ComHostPath."
    }

    $args = @("/s")
    if ($Unregister) {
        $args += "/u"
    }

    $args += "`"$ComHostPath`""

    $verb = $Unregister ? "Unregistering" : "Registering"
    Write-Host "$verb shell extension via regsvr32..." -ForegroundColor Cyan

    $process = Start-Process -FilePath "regsvr32.exe" -ArgumentList $args -PassThru -Wait

    if ($process.ExitCode -ne 0) {
        throw "regsvr32 exited with code $($process.ExitCode)."
    }

    $resultVerb = $Unregister ? "Unregister" : "Register"
    Write-Host "regsvr32 $resultVerb completed successfully." -ForegroundColor Green
}

$projectPath = Join-Path $PSScriptRoot "CropperShellExtension" "CropperShellExtension.csproj"
$publishDir = Join-Path $PSScriptRoot "CropperShellExtension" "bin" $Configuration "net8.0-windows" "win-x64" "publish"
$comHost = Join-Path $publishDir "CropperShellExtension.comhost.dll"

$shouldUnregister = $Action -eq "Unregister" -or $Action -eq "Install" -or ($Force -and $Action -eq "Register")
$shouldRegister = $Action -eq "Install" -or $Action -eq "Register"

if ($shouldUnregister) {
    $ignoreMissing = $Action -ne "Unregister" -or $Force
    Invoke-Regsvr32 -ComHostPath $comHost -Unregister -IgnoreMissing:$ignoreMissing
}

if ($shouldRegister) {
    Write-Host "Building CropperShellExtension ($Configuration)..." -ForegroundColor Cyan
    dotnet publish $projectPath -c $Configuration -r win-x64 --no-self-contained

    if (-not (Test-Path $comHost)) {
        throw "Unable to locate comhost at $comHost."
    }

    Invoke-Regsvr32 -ComHostPath $comHost
}

Write-Host "Completed action '$Action'." -ForegroundColor Green
