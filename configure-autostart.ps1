param(
    [string]$TaskPath = '\OpenClaw\',
    [string]$TaskName = 'OutlookMCPServer'
)

$ErrorActionPreference = 'Stop'

$xml = Export-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName
$xml = $xml.Replace(
    '<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>',
    '<MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>'
)
if ($xml -match '<RunLevel>.*?</RunLevel>') {
    $xml = $xml -replace `
        '<RunLevel>.*?</RunLevel>', `
        '<RunLevel>HighestAvailable</RunLevel>'
} else {
    $xml = $xml.Replace(
        '<LogonType>InteractiveToken</LogonType>',
        '<LogonType>InteractiveToken</LogonType>' +
        '<RunLevel>HighestAvailable</RunLevel>'
    )
}

Register-ScheduledTask `
    -TaskPath $TaskPath `
    -TaskName $TaskName `
    -Xml $xml `
    -Force | Out-Null

$task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName
$configuredXml = Export-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName
$multipleInstances = [regex]::Match(
    $configuredXml,
    '<MultipleInstancesPolicy>(.*?)</MultipleInstancesPolicy>'
).Groups[1].Value
[pscustomobject]@{
    RunLevel = $task.Principal.RunLevel
    LogonType = $task.Principal.LogonType
    MultipleInstances = $multipleInstances
    State = $task.State
}
