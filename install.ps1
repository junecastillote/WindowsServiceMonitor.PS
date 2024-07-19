[CmdletBinding()]
param ()

if ($PSEdition -eq 'Desktop') {
    Write-Output "This module requires PowerShell Core."
    return $null
}

[bool]$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (!$isAdmin) {
    Write-Output "Installing the module requires elevation. Run PowerShell as admin and try again."
    return $null
}

$moduleManifest = Get-ChildItem -Path $PSScriptRoot\WindowsServiceMonitor.PS.psd1
$Moduleinfo = Test-ModuleManifest -Path ($moduleManifest.FullName)

Remove-Module ($Moduleinfo.Name) -ErrorAction SilentlyContinue

# PowerShell Core
# if ($PSEdition -eq 'Core') {
$ModulePath = "$env:ProgramFiles\PowerShell\Modules"
# }

# Register event source
# if (-not [System.Diagnostics.EventLog]::SourceExists('JFC.GroupMembersSync.PS')) {
#     [System.Diagnostics.EventLog]::CreateEventSource('JFC.GroupMembersSync.PS', 'Application')
# }

$ModulePath = ($ModulePath + "\$($Moduleinfo.Name.ToString())\$($Moduleinfo.Version.ToString())")

Write-Output "Module installation path: [$($ModulePath)]"
Write-Output "Module version: [$($Moduleinfo.Version.ToString())]"

if (!(Test-Path $ModulePath)) {
    New-Item -Path $ModulePath -ItemType Directory -Force | Out-Null
}

try {
    Copy-Item -Path $PSScriptRoot\* -Include *.psd1, *.psm1 -Destination $ModulePath -Force -Confirm:$false -ErrorAction Stop
    Copy-Item -Path $PSScriptRoot\source -Recurse -Destination $ModulePath -Force -Confirm:$false -ErrorAction Stop
    Write-Output ""
    Write-Output "Success. Installed to [$ModulePath]"
    Write-Output ""
    Get-ChildItem -Recurse $ModulePath | Unblock-File -Confirm:$false
    Import-Module $($Moduleinfo.Name.ToString()) -Force
}
catch {
    Write-Output ""
    Write-Output "Failed"
    Write-Output $_.Exception.Message
    Write-Output ""
}