Function Invoke-WindowsServiceCheck {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string[]]
        $ServiceName,

        [Parameter()]
        [int]
        $Interval,

        [Parameter()]
        [ValidateSet(
            'Days',
            'Hours',
            'Minutes',
            'Seconds'
        )]
        [string]
        $Unit
    )


}