Function Send-WindowsServiceReportBySmtp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline = $true)]
        [PSTypeNameAttribute('WindowsServiceMonitorResult')]
        $WindowsServiceMonitorResult,

        [Parameter(Mandatory)]
        [string]
        $MailFrom,

        [Parameter()]
        [string[]]
        $MailTo,

        [Parameter()]
        [string[]]
        $MailCc,

        [Parameter()]
        [string[]]
        $MailBcc,

        [Parameter(Mandatory)]
        [string]
        $SmtpServer,

        [Parameter()]
        [int]
        $SmtpPort = 25,

        [Parameter()]
        [pscredential]
        $SmtpCredential,

        [Parameter(Mandatory)]
        [string]
        $MailSubject,

        [Parameter()]
        [switch]
        $SmtpSSLEnabled,

        # Parameter help description
        [Parameter()]
        [ValidateSet('Low', 'Normal', 'High')]
        [string]
        $MailPriority = 'Normal'
    )

    $mailSplat = @{
        SmtpServer = $SmtpServer
        Port       = $SmtpPort
        UseSsl     = $SmtpSSLEnabled
        BodyAsHtml = $true
        # Body       = $MailBody
        Body       = $WindowsServiceMonitorResult.ServiceReportHtml
        Subject    = $MailSubject
        From       = $MailFrom
        Priority   = $MailPriority
    }

    if ($SmtpCredential) {
        $mailSplat.Add('Credential', $SmtpCredential)
    }

    if ($MailTo) {
        $mailSplat.Add('To', $MailTo)
    }

    if ($MailCc) {
        $mailSplat.Add('Cc', $MailCc)
    }

    if ($MailBcc) {
        $mailSplat.Add('Bcc', $MailBcc)
    }

    Send-MailMessage @mailSplat -WarningAction SilentlyContinue

}