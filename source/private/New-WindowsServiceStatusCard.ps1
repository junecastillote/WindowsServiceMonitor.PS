# This function creates a consolidated Teams report
# using adaptive cards 1.4.
Function New-WindowsServiceStatusCard {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [PSTypeNameAttribute('WindowsServiceMonitorResult')]
        $InputObject
    )

    $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

    Function New-FactItem {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            $InputObject
        )

        $factHeader = [pscustomobject][ordered]@{
            type  = "Container"
            style = "emphasis"
            bleed = $true
            items = @(
                $([pscustomobject][ordered]@{
                        type      = 'TextBlock'
                        wrap      = $true
                        separator = $true
                        weight    = 'Bolder'
                        text      = "Windows Service Monitor detected [$($InputObject.ServiceNotRunningCount)] services are not running!"
                    } )
            )
        }

        $factSet = [pscustomobject][ordered]@{
            type      = 'FactSet'
            separator = $true
            facts     = @(
                $([pscustomobject][ordered]@{Title = 'Service'; Value = "$($InputObject.DisplayName)" } ),
                $([pscustomobject][ordered]@{Title = 'Status'; Value = $($InputObject.Status.ToString()) } )
            )
        }
        return @($factSet)
    }

    $jsonPayload = $(Get-Content (
        ($moduleInfo.ModuleBase.ToString()) + '\source\private\TeamsConsolidated.json'
        ) -Raw
    ).Replace(
        'vOrganizationName', $InputObject.OrganizationName
    ).Replace(
        'vComputerName', $InputObject.ComputerName
    )
    # $teamsAdaptiveCard = (Get-Content (($moduleInfo.ModuleBase.ToString()) + '\source\private\TeamsConsolidated.json') -Raw | ConvertFrom-Json)
    $teamsAdaptiveCard = ($jsonPayload | ConvertFrom-Json)
    foreach ($item in $InputObject.ServiceNotRunning) {
        $teamsAdaptiveCard.attachments[0].content.body += (New-FactItem -InputObject $item)
    }
    # $teamsAdaptiveCard = (($teamsAdaptiveCard | ConvertTo-Json -Depth 50))
    return $teamsAdaptiveCard
}