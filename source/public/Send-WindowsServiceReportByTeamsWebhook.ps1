Function Send-WindowsServiceReportByTeamsWebhook {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [PSTypeNameAttribute('WindowsServiceMonitorResult')]
        $InputObject,

        [Parameter(Mandatory)]
        [string[]]
        $TeamsWebHookURL
    )

    $teamsAdaptiveCard = New-WindowsServiceStatusCard -InputObject $InputObject


    foreach ($url in $TeamsWebHookURL) {
        SayInfo "Posting alert to Teams with URL [$($url)]"
        $Params = @{
            "URI"         = $url
            "Method"      = 'POST'
            "Body"        = $teamsAdaptiveCard | ConvertTo-Json -Depth 10
            "ContentType" = 'application/json'
        }
        try {
            Invoke-RestMethod @Params -ErrorAction Stop
        }
        catch {
            SayError "Failed to post to channel. $($_.Exception.Message)"
        }
    }
}