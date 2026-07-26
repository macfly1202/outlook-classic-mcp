[CmdletBinding()]
param(
    [ValidateSet("user", "project", "local")]
    [string]$Scope = "user",

    [switch]$SkipUvInstall
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$PluginId = "outlook@outlook-classic-mcp"
$MarketplaceName = "outlook-classic-mcp"
$MarketplaceSource = "macfly1202/outlook-classic-mcp"

function Invoke-Claude {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [switch]$IgnoreFailure
    )

    & claude @Arguments
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0 -and -not $IgnoreFailure) {
        throw "claude $($Arguments -join ' ') failed with exit code $exitCode."
    }
    return $exitCode
}

function Resolve-Uv {
    $uvCommand = Get-Command uv -ErrorAction SilentlyContinue
    if ($uvCommand) {
        return $uvCommand.Source
    }

    $candidate = Join-Path $env:USERPROFILE ".local\bin\uv.exe"
    if (Test-Path $candidate) {
        return $candidate
    }
    return $null
}

Write-Host ""
Write-Host "Outlook Classic MCP fork installer" -ForegroundColor Cyan
Write-Host "Replacing the installed plugin with $MarketplaceSource" -ForegroundColor Cyan
Write-Host ""

if (-not (Get-Command claude -ErrorAction SilentlyContinue)) {
    throw "Claude Code is not available in PATH. Install or update Claude Code, then re-run this script."
}

$uvPath = Resolve-Uv
if (-not $uvPath -and -not $SkipUvInstall) {
    Write-Host "Installing uv..." -ForegroundColor Yellow
    Invoke-RestMethod "https://astral.sh/uv/install.ps1" | Invoke-Expression
    $uvPath = Resolve-Uv
}
if (-not $uvPath) {
    throw "uv is not available. Install it from https://docs.astral.sh/uv/ and re-run this script."
}

Write-Host "Removing previous Outlook plugin installations..." -ForegroundColor Yellow
foreach ($installedScope in @("user", "project", "local")) {
    Invoke-Claude -Arguments @(
        "plugin", "uninstall", $PluginId,
        "--scope", $installedScope,
        "--yes"
    ) -IgnoreFailure | Out-Null
}

Write-Host "Removing the previous marketplace..." -ForegroundColor Yellow
Invoke-Claude -Arguments @(
    "plugin", "marketplace", "remove", $MarketplaceName
) -IgnoreFailure | Out-Null

Write-Host "Clearing the cached PyPI build..." -ForegroundColor Yellow
& $uvPath cache clean outlook-classic-mcp 2>$null

Write-Host "Adding the fork marketplace..." -ForegroundColor Yellow
Invoke-Claude -Arguments @(
    "plugin", "marketplace", "add", $MarketplaceSource
) | Out-Null

Write-Host "Installing $PluginId at $Scope scope..." -ForegroundColor Yellow
Invoke-Claude -Arguments @(
    "plugin", "install", $PluginId,
    "--scope", $Scope
) | Out-Null

Write-Host ""
Write-Host "Installation complete." -ForegroundColor Green
Write-Host "Restart Claude Code, or run /reload-plugins in an active session."
Write-Host "Then call outlook_whoami to verify the Outlook connection."
