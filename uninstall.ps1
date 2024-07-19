[CmdletBinding()]
param (
    [Parameter()]
    [switch]
    $All
)
$moduleManifest = Test-ModuleManifest -Path $((Get-ChildItem -Path $PSScriptRoot\WindowsServiceMonitor.PS.psd1).FullName)

if ($All) {
    $module = Get-Module $($moduleManifest.Name) -ListAvailable
}
else {
    $module = Get-Module $($moduleManifest.Name) -ListAvailable | Where-Object { $_.Version -eq $moduleManifest.Version }
}

[bool]$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (!$isAdmin) {
    Write-Output "Uninstallation requires elevation. Run PowerShell as admin and try again."
    return $null
}

if ($module) {
    foreach ($Moduleinfo in $module) {

        Remove-Module ($Moduleinfo.Name) -ErrorAction SilentlyContinue

        $ModulePath = $Moduleinfo.ModuleBase
        try {
            Write-Output "Uninstalling $($Moduleinfo.Name) [$($Moduleinfo.Version)]..."
            $null = Remove-Item $ModulePath -Recurse -Confirm:$false -Force -ErrorAction Stop
            Write-Output "Done uninstalling $($Moduleinfo.Name) [$($Moduleinfo.Version)] from $($Moduleinfo.ModuleBase)"
        }
        catch {
            Write-Output ""
            Write-Output "Failed"
            Write-Output $_.Exception.Message
            Write-Output ""
            return $null
        }
    }

    if ((@(Get-ChildItem -Recurse -LiteralPath $(Split-Path $ModulePath) -File).Count -lt 1)) {
        $null = Remove-Item $(Split-Path $ModulePath) -Recurse -Confirm:$false -Force -ErrorAction Stop
    }

    # Register event source
    # if ([System.Diagnostics.EventLog]::SourceExists('JFC.GroupMembersSync.PS')) {
    #     [System.Diagnostics.EventLog]::DeleteEventSource('JFC.GroupMembersSync.PS')
    # }
}
else {
    "[$($moduleManifest.Name)] module not found. Nothing to uninstall." | Out-Default
}

