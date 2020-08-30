##Query registry for Microsoft 365 Apps information##

$M365Platofrm=(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).platform
$M365Version=(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').ClientVersionToReport
$M365Language=(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').ClientCulture
$M365UpdateChannel=(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').UpdateChannel
$M365ProductsInstalled=(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').ProductReleaseIds
$M365UpdatesEnabled=(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').UpdatesEnabled
$M365DeviceName = $env:computername

## Identidy the Update channel ##
if($M365Version -gt $null){
    switch ($M365UpdateChannel)
{
    "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6" {$M365UpdateChannel="Monthly Enterprise Channel"}
    "http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60" {$M365UpdateChannel="Current Channel"}
    "http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be" {$M365UpdateChannel="Current Channel (Preview)"}
    "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" {$M365UpdateChannel="Semi-Annual Enterprise Channel"}
    "http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf" {$M365UpdateChannel="Semi-Annual Enterprise Channel (Preview)"}
    "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f" {$M365UpdateChannel="Beta Channel"}
}

## Build a JSON object with all information collected ##
$OPPInformation=@{DeviceName=$M365DeviceName;Platform=$M365Platofrm;Version=$M365Version;Language=$M365Language;Channel=$M365UpdateChannel;ProductsInstalled=$M365ProductsInstalled;UpdatesEnabled=$M365UpdatesEnabled}
$output=$OPPInformation | ConvertTo-Json -Compress

## Display this information to the package output

Write-Output $output
Exit 0
}else{
    Write-Output "M365 Apps are not installed"
    Exit 1
}
