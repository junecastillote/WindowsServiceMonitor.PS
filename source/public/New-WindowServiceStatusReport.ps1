Function New-WindowServiceStatusReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string[]]
        $Service,

        [Parameter()]
        [string]
        $OutputDirectory = $($env:temp),

        [Parameter()]
        [int]
        $MaxNumberOfReportsToKeep = 30,

        [Parameter(Mandatory)]
        [string]
        $OrganizationName = ""
    )

    $module_Info = (ThisModule)
    $notification_template = "$($module_Info.ModuleBase)\source\private\notification_template.html"
    $report_html_file = "$($OutputDirectory)\WindowsServiceMonitor.PS_$(([datetime]::Now).ToString('yyyy-MM-dd_HH-mm-ss'))_report.html"

    ## Get service(s) status
    $service_collection = @(Get-Service $Service -ErrorAction SilentlyContinue | Sort-Object Status, DisplayName)

    ## Terminate loop if $service_collection is nothing
    if (!$service_collection) {
        break
    }

    $service_collection = $service_collection | Sort-Object Status, DisplayName
    $not_running_service = @($service_collection | Where-Object { $_.Status -ne 'Running' })
    $running_service = @($service_collection | Where-Object { $_.Status -eq 'Running' })
    $service_report_items_html = [System.Collections.Generic.List[string]]@()

    ## Display each service and status
    ## Not running
    $not_running_service | ForEach-Object {
        $current_service = $_

        $service_report_items_html.Add(
            "`n" +
            '        <tr>' + "`n" +
            '            <td style="border: 1px solid #dddddd; padding: 5px; text-align: left;"">' + $current_service.DisplayName + '</td>' + "`n" +
            '            <td style="border: 1px solid #dddddd; padding: 5px; text-align: left;">' + $current_service.Name + '</td>' + "`n" +
            '            <td style="border: 1px solid #dddddd; padding: 5px; text-align: left; color: red; font-weight: bold;">' + $current_service.Status + '</td>' + "`n" +
            '        </tr>'
        )
    }
    ## Running
    $running_service | ForEach-Object {
        $current_service = $_
        $service_report_items_html.Add(
            "`n" +
            '        <tr>' + "`n" +
            '            <td style="border: 1px solid #dddddd; padding: 5px; text-align: left;">' + $current_service.DisplayName + '</td>' + "`n" +
            '            <td style="border: 1px solid #dddddd; padding: 5px; text-align: left;">' + $current_service.Name + '</td>' + "`n" +
            '            <td style="border: 1px solid #dddddd; padding: 5px; text-align: left; color: green; font-weight: bold;">' + $current_service.Status + '</td>' + "`n" +
            '        </tr>'
        )
    }

    ## Write HTML Report
    $message_body = (Get-Content -LiteralPath $notification_template -Raw)
    $message_body = $message_body.Replace(
        '        <!-- vServiceList -->', ($service_report_items_html)
    ).Replace(
        'vOrganizationName', $OrganizationName
    ).Replace(
        'vComputerName', $env:COMPUTERNAME
    )
    $message_body | Out-File -LiteralPath $report_html_file -Force -Confirm:$false

    ## Housekeeping
    if ($MaxNumberOfReportsToKeep -gt 0) {
        if ($history_to_delete = Get-ChildItem "$($OutputDirectory)\WindowsServiceMonitor.PS*report.html" -File | Sort-Object CreationTime -Descending | Select-Object -Skip $MaxNumberOfReportsToKeep) {
            $history_to_delete | Remove-Item -Confirm:$false -Force -ErrorAction Continue
        }
    }

    ## Compose the output object
    $result = [PSCustomObject]$(
        [ordered]@{
            ServiceNotRunning      = @($not_running_service)
            ServiceRunning         = @($running_service)
            ServiceNotRunningCount = @($not_running_service).Count
            ServiceRunningCount    = @($running_service).Count
            ServiceReportFile      = $report_html_file
            ServiceReportHtml      = $($message_body).ToString()
            PSTypeName             = 'WindowsServiceMonitorResult'
        }
    )
    $visible_properties = [string[]]@('ServiceNotRunning', 'ServiceRunning', 'ServiceReportFile')
    [Management.Automation.PSMemberInfo[]]$default_properties = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet', $visible_properties)
    $result | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $default_properties
    return $result
}