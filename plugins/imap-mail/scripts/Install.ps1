# Installs the ImapMail PowerShell module for the current user and runs setup.
# Run with:  powershell -ExecutionPolicy Bypass -File Install.ps1
$ErrorActionPreference = 'Stop'
$src = $PSScriptRoot
$dest = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules\ImapMail'
New-Item -ItemType Directory -Force $dest | Out-Null
foreach ($f in 'ImapMail.psm1', 'ImapMail.cs', 'ImapMail.Calendar.ps1') {
    Copy-Item (Join-Path $src $f) $dest -Force
    Unblock-File (Join-Path $dest $f)
}
Write-Host "Installed ImapMail module to $dest" -ForegroundColor Green
Set-ExecutionPolicy -Scope Process Bypass -Force
Import-Module ImapMail -Force
Initialize-ImapMail
Read-Host 'Press Enter to close'
