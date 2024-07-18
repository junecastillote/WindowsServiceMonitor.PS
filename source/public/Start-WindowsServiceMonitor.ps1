Function Start-WindowsServiceMonitor {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ConfigurationFile
    )
    "=============================================================================" | Say
    # Import the configuration data. Exit on error.
    try {
        $full_path = (Resolve-Path -LiteralPath $ConfigurationFile -ErrorAction Stop).Path
        $config = Import-PowerShellDataFile -LiteralPath $full_path -ErrorAction Stop
    }
    catch {
        $_.Exception.Message | SayError
        return $null
    }

    # Validate configuration
    $isValid = $true

    if ($config.UseCustomServiceListFile) {
        if (!$config.CustomServiceListFile) {
            "The custom service list file is not specified in the CustomServiceListFile configuration." | SayError
            $isValid = $false
        }

        if ($config.CustomServiceListFile) {
            try {
                $custom_service_list = Get-Content $config.CustomServiceListFile -ErrorAction Stop
                $config.Service = $custom_service_list
            }
            catch {
                "Failed to import the custom service list file." | SayError
                $_.Exception.Message | SayError
                $isValid = $false
            }
        }
    }

    if (!($config.UseCustomServiceListFile)) {
        if (!$config.Service) {
            "The list of services to monitor is missing." | SayError
            $isValid = $false
        }
    }

    if (!$config.OrganizationName) {
        "The OrganizationName is missing." | SayError
        $isValid = $false
    }

    if ($config.MaxNumberOfReportsToKeep -lt 1) {
        "MaxNumberOfReportsToKeep must be greater than 0." | SayError
        $isValid = $false
    }

    ## Validate SMTP-related parameters
    if ($config.SendEmailNotificationMethod -eq 'Smtp') {
        if (!$config.SmtpConfiguration.MailFrom) {
            "[MailFrom] value is required." | SayError
            $isValid = $false
        }
        if (!$config.SmtpConfiguration.SmtpServer) {
            "[SmtpServer] value is required." | SayError
            $isValid = $false
        }
        if (!$config.SmtpConfiguration.SmtpPort) {
            "[SmtpPort] value is required." | SayError
            $isValid = $false
        }
        if (!$config.SmtpConfiguration.MailTo) {
            "[MailTo] value is required." | SayError
            $isValid = $false
        }

        if ($config.SmtpConfiguration.CredentialFile) {
            try {
                $smtp_credential = Import-Clixml -LiteralPath $config.SmtpConfiguration.CredentialFile -ErrorAction Stop
                $config.SmtpConfiguration.Add('SmtpCredential', $smtp_credential)
                $config.SmtpConfiguration.Remove('CredentialFile')
            }
            catch {
                $isValid = $false
                "Failed to import the encrypted SMTP credential." | SayError
                $_.Exception.Message | SayError
            }
        }

        $config.SmtpConfiguration.Add('MailPriority', 'Normal')
        $config.SmtpConfiguration.Add('MailSubject', '')
    }

    if (!$config.OutputDirectory) {
        $config.OutputDirectory = $env:TEMP
    }


    if (!$isValid) {
        "=============================================================================" | Say
        return $null
    }

    "Configuration" | Say
    "Use custom service list file input? = [$($config.UseCustomServiceListFile)]" | SayInfo
    if ($config.UseCustomServiceListFile -and $config.CustomServiceListFile) {
        "Custom service list file = [$($config.CustomServiceListFile)]" | SayInfo
    }
    "Services monitored = [$($config.Service -join ',')]" | SayInfo
    "Organization name = [$($config.OrganizationName)]" | SayInfo
    "Monitoring interval = [$($config.IntervalInSeconds) seconds]" | SayInfo
    "Notify non-running only? = [$($config.NotifyIfNotRunningOnly)]" | SayInfo
    "Email notification method = [$($config.SendEmailNotificationMethod)]" | SayInfo
    "Max # of reports to keep = [$($config.MaxNumberOfReportsToKeep)]" | SayInfo


    do {
        "=============================================================================" | Say
        "Service Status Check" | Say
        $result = New-WindowServiceStatusReport -Service $config.Service -OutputDirectory $config.OutputDirectory -OrganizationName $config.OrganizationName -MaxNumberOfReportsToKeep $config.MaxNumberOfReportsToKeep


        # if ($config.SendEmailNotificationMethod -eq 'Smtp') {
        $smtp_splat = $config.SmtpConfiguration

        if ($result.ServiceNotRunningCount) {
            $smtp_splat.MailPriority = 'High'
            $smtp_splat.MailSubject = "[CRITICAL] | Windows Service Monitor detected [$($result.ServiceNotRunningCount)] services are not running!"
        }

        if (!$result.ServiceNotRunningCount) {
            $smtp_splat.MailPriority = 'Low'
            $smtp_splat.MailSubject = "[NORMAL] | Windows Service Monitor"
        }

        # Display service status on the screen
        $result.ServiceNotRunning | ForEach-Object {
            "[$($_.DisplayName)] : [$($_.Status)]" | SayInfo -Color Red
        }

        $result.ServiceRunning | ForEach-Object {
            "[$($_.DisplayName)] : [$($_.Status)]" | SayInfo
        }

        "Monitored services not running = [$($result.ServiceNotRunningCount)]" | SayInfo
        "Report file @ [$($result.ServiceReportFile)]" | SayInfo

        if ($config.SendEmailNotificationMethod -eq 'Smtp') {
            if ($config.NotifyIfNotRunningOnly -and $result.ServiceNotRunningCount -gt 0) {
                "Sending alert notification..." | SayInfo
                try {
                    $result | Send-WindowsServiceReportBySmtp @smtp_splat -ErrorAction Stop
                    "Done." | SayInfo
                }
                catch {
                    "Failed to send email notification." | SayError
                    $_.Exception.Message | SayError
                }
            }
        }

        ## Housekeeping
        if ($history_to_delete = Get-ChildItem "$($config.OutputDirectory)\WindowsServiceMonitor.PS*report.html" -File | Sort-Object CreationTime -Descending | Select-Object -Skip $($config.MaxNumberOfReportsToKeep)) {
            $history_to_delete | Remove-Item -Confirm:$false -Force -ErrorAction Continue
        }

        ## Terminate loop if -IntervalInSeconds is not specified (no interval)
        if (!$config.IntervalInSeconds) {
            break
        }

        # Suspend for the duration of IntervalInSeconds
        for ($i = ($config.IntervalInSeconds); $i -gt 0; $i--) {
            $percentComplete = (($config.IntervalInSeconds - $i) / $config.IntervalInSeconds) * 100
            Write-Progress -Activity "Next check in: " -Status "$i seconds" -PercentComplete $percentComplete
            Start-Sleep -Seconds 1
        }
    }
    while ($true)
}