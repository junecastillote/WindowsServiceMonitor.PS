# [enum]::GetValues([System.ConsoleColor])
Function Say {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Text,
        [Parameter()]
        $Color = 'Cyan'
    )

    if ($Color) {
        $Host.UI.RawUI.ForegroundColor = $Color
    }
    $Text | Out-Host
    [Console]::ResetColor()
}

Function SayError {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Text,
        [Parameter()]
        $Color = 'Red'
    )
    $Host.UI.RawUI.ForegroundColor = $Color
    "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [ERROR] - $Text" | Out-Host
    [Console]::ResetColor()
}

Function SayInfo {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Text,
        [Parameter()]
        $Color = 'Green'
    )
    $Host.UI.RawUI.ForegroundColor = $Color
    "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [INFO] - $Text" | Out-Host
    [Console]::ResetColor()
}

Function SayWarning {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Text,
        [Parameter()]
        $Color = 'DarkYellow'
    )
    $Host.UI.RawUI.ForegroundColor = $Color
    "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [WARNING] - $Text" | Out-Host
    [Console]::ResetColor()
}

