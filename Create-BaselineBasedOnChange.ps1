## Variables to Edit  ##
$tenant = ""
## End Editing ##
$authority = "https://login.windows.net/$tenant"
$AppId = Get-AutomationVariable -Name 'AppId'
$AppSecret = Get-AutomationVariable -Name 'AppSecret'
$Resource = "deviceManagement/auditEvents"
$graphApiVersion = "Beta"
$uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)"
$previousday = (Get-Date).AddDays(-1)
$currentdate= (get-date).ToShortDateString()

Update-MSGraphEnvironment -AppId $AppId -Quiet
Update-MSGraphEnvironment -AuthUrl $authority -Quiet
Connect-MSGraph -ClientSecret $AppSecret -Quiet
 
$body=Invoke-MSGraphRequest -HttpMethod GET -Url $uri
$body.value | Where-Object {$_.componentname -like "DeviceConfiguration" -and $_.resources.type -like "Windows10*" } | ForEach-Object{
    if(($_ -ne $null) -and ($_.activitydatetime -ge $previousday)){
        $Resource = "deviceManagement/userExperienceAnalyticsBaselines"
        $graphApiVersion = "Beta"
        $newBaselineURL= "https://graph.microsoft.com/$graphApiVersion/$($resource)"
$newBaseline=@"
{
    "displayName":"$($currentdate)"
}
"@

        Invoke-MSGraphRequest -HttpMethod POST -Url $newBaselineURL -Content $newBaseline | out-null
    }
}

