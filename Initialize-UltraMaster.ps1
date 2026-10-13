<#
.SYNOPSIS
    One-time (idempotent) bootstrap for the UltraMaster environment on Windows.

.DESCRIPTION
    Sets machine-level environment variables, stores secrets in Windows
    Credential Manager, installs Python/Node.js, and installs pip/npm
    dependencies. Safe to re-run: each step is skipped when already done.

    Run once from an elevated PowerShell prompt:
        .\Initialize-UltraMaster.ps1

    From UltraMaster-Orchestrator.ps1, call it only when setup is missing:
        if (-not $env:PROJECT_ROOT) { & "$PSScriptRoot\Initialize-UltraMaster.ps1" }
    and read secrets with:
        . "$PSScriptRoot\Initialize-UltraMaster.ps1" -FunctionsOnly
        $clientSecret = Get-UltraMasterSecret -Name 'GRAPH_CLIENT_SECRET'

.PARAMETER FunctionsOnly
    Only define the helper functions (for dot-sourcing); perform no setup.

.PARAMETER ResetSecrets
    Prompt for every secret again, even if it is already stored.
#>
[CmdletBinding()]
param(
    [switch]$FunctionsOnly,
    [switch]$ResetSecrets
)

# --- Credential Manager helpers ----------------------------------------------
# Calls the Win32 CredWrite/CredRead APIs directly, so secret values never
# appear on a command line (cmdkey /pass:... is visible in the process list
# and in PowerShell history).
if (-not ('UltraMaster.CredMan' -as [type])) {
    Add-Type -Namespace UltraMaster -Name CredMan -MemberDefinition @'
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct CREDENTIAL {
    public int Flags;
    public int Type;
    public string TargetName;
    public string Comment;
    public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
    public int CredentialBlobSize;
    public IntPtr CredentialBlob;
    public int Persist;
    public int AttributeCount;
    public IntPtr Attributes;
    public string TargetAlias;
    public string UserName;
}
[DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool CredWrite(ref CREDENTIAL credential, int flags);
[DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool CredRead(string target, int type, int flags, out IntPtr credential);
[DllImport("advapi32.dll")]
public static extern void CredFree(IntPtr buffer);
'@
}

$script:CredTypeGeneric = 1
$script:CredPersistLocalMachine = 2
$script:SecretPrefix = 'UltraMaster/'

function Set-UltraMasterSecret {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][securestring]$Value
    )
    $plain = [Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($Value)
    try {
        $cred = New-Object UltraMaster.CredMan+CREDENTIAL
        $cred.Type = $script:CredTypeGeneric
        $cred.TargetName = $script:SecretPrefix + $Name
        $cred.UserName = 'UltraMaster'
        $cred.Persist = $script:CredPersistLocalMachine
        $cred.CredentialBlob = $plain
        $cred.CredentialBlobSize = $Value.Length * 2
        if (-not [UltraMaster.CredMan]::CredWrite([ref]$cred, 0)) {
            throw "CredWrite failed for $Name (Win32 error $([Runtime.InteropServices.Marshal]::GetLastWin32Error()))."
        }
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($plain)
    }
}

function Get-UltraMasterSecret {
    param([Parameter(Mandatory)][string]$Name)
    $ptr = [IntPtr]::Zero
    if (-not [UltraMaster.CredMan]::CredRead($script:SecretPrefix + $Name, $script:CredTypeGeneric, 0, [ref]$ptr)) {
        return $null
    }
    try {
        $cred = [Runtime.InteropServices.Marshal]::PtrToStructure($ptr, [type][UltraMaster.CredMan+CREDENTIAL])
        if ($cred.CredentialBlobSize -eq 0) { return '' }
        return [Runtime.InteropServices.Marshal]::PtrToStringUni($cred.CredentialBlob, $cred.CredentialBlobSize / 2)
    }
    finally {
        [UltraMaster.CredMan]::CredFree($ptr)
    }
}

if ($FunctionsOnly) { return }

$ErrorActionPreference = 'Stop'

# --- Preconditions ------------------------------------------------------------
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this script from an elevated PowerShell prompt (machine-level environment variables and winget installs need admin rights).'
}

Write-Host 'Initializing UltraMaster environment...'

# --- Environment variables ------------------------------------------------------
$machineVars = [ordered]@{
    NODE_ENV         = 'production'
    LOG_LEVEL        = 'info'
    PROJECT_ROOT     = 'C:\Scripts\UltraMaster'
    GRAPH_START_DATE = '2023-07-05'
}
foreach ($entry in $machineVars.GetEnumerator()) {
    [Environment]::SetEnvironmentVariable($entry.Key, $entry.Value, 'Machine')
    # SetEnvironmentVariable(..., 'Machine') does not affect the running session.
    Set-Item -Path "Env:$($entry.Key)" -Value $entry.Value
}
# Git identity is per user, not per machine.
[Environment]::SetEnvironmentVariable('GIT_AUTHOR_NAME', 'Stamos', 'User')
$env:GIT_AUTHOR_NAME = 'Stamos'

if (-not (Test-Path $env:PROJECT_ROOT)) {
    New-Item -ItemType Directory -Path $env:PROJECT_ROOT | Out-Null
}

# --- Secrets (Windows Credential Manager) --------------------------------------
# Values are typed in at a hidden prompt, never stored in this file.
$secretNames = 'GRAPH_CLIENT_ID', 'GRAPH_CLIENT_SECRET', 'GRAPH_TENANT_ID', 'HOSPITAL_API_KEY'
foreach ($name in $secretNames) {
    if (-not $ResetSecrets -and $null -ne (Get-UltraMasterSecret -Name $name)) {
        Write-Host "Secret $name already stored; skipping."
        continue
    }
    $value = Read-Host -Prompt "Enter value for $name" -AsSecureString
    if ($value.Length -eq 0) {
        Write-Warning "No value entered for $name; skipping."
        continue
    }
    Set-UltraMasterSecret -Name $name -Value $value
    Write-Host "Secret $name stored in Credential Manager."
}

# --- Dependencies ---------------------------------------------------------------
function Update-SessionPath {
    # winget updates PATH in the registry only; reload it so python/pip/npm
    # are found in this same session.
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
}

function Install-WingetPackage {
    param([Parameter(Mandatory)][string]$Id)
    $null = winget list --id $Id -e --accept-source-agreements
    if ($LASTEXITCODE -eq 0) {
        Write-Host "$Id already installed; skipping."
        return
    }
    Write-Host "Installing $Id..."
    winget install --id $Id -e --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) { throw "winget install $Id failed (exit code $LASTEXITCODE)." }
}

if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    throw 'winget is not available. Install "App Installer" from the Microsoft Store, then re-run.'
}

Install-WingetPackage -Id 'Python.Python.3.12'
Install-WingetPackage -Id 'OpenJS.NodeJS.LTS'
Update-SessionPath

Write-Host 'Installing Python packages...'
python -m pip install --upgrade requests openpyxl pandas
if ($LASTEXITCODE -ne 0) { throw "pip install failed (exit code $LASTEXITCODE)." }

$packageJson = Join-Path $env:PROJECT_ROOT 'package.json'
if (Test-Path $packageJson) {
    Write-Host 'Installing npm packages...'
    Push-Location $env:PROJECT_ROOT
    try {
        npm install
        if ($LASTEXITCODE -ne 0) { throw "npm install failed (exit code $LASTEXITCODE)." }
    }
    finally {
        Pop-Location
    }
}
else {
    Write-Host "No package.json in $env:PROJECT_ROOT; skipping npm install."
}

Write-Host 'UltraMaster environment initialized.'
