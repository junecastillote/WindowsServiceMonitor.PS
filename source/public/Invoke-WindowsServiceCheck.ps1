Function Invoke-WindowsServiceCheck {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ServiceName,

        [Parameter()]
        $IntervalInSeconds
    )

    if ($PSVersionTable.PSEdition -eq 'Core') {
        # $PSStyle.Progress.View = 'Classic'
    }

    do {
        "=============================================================================" | Out-Default
        if ($service = (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue | Select-Object Name, DisplayName, Status)) {
            $service | ForEach-Object {
                "[$($_.DisplayName)],[$($_.Name)] = $($_.Status)" | Out-Default
            }
        }

        $stopped_service = @($service | Where-Object { $_.Status -ne 'Running' })
        "Service(s) not running = $($stopped_service.count)" | Out-Default

        if ($stopped_service.Count -gt 0) {
            TODO: email notification
        }

        if ($IntervalInSeconds) {
            for ($i = $IntervalInSeconds; $i -ge 0; $i--) {
                $percentComplete = (($IntervalInSeconds - $i) / $IntervalInSeconds) * 100
                Write-Progress -Activity "Next check in: " -Status "$i seconds" -PercentComplete $percentComplete
                Start-Sleep -Seconds 1
            }
        }
        else {
            break
        }
    }
    while ($true)
}