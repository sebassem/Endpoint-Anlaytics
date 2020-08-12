$MonthstoKeepBaselines=5
$tenant = ""
$authority = "https://login.windows.net/$tenant"
$AppId = Get-AutomationVariable -Name 'AppId'
$AppSecret = Get-AutomationVariable -Name 'AppSecret'
$Resource = "deviceManagement/userExperienceAnalyticsBaselines"
$graphApiVersion = "Beta"
$uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)?"+'$orderby'+"=createdDateTime%20desc"
$currentMonth= get-date -Format Y


Update-MSGraphEnvironment -AppId $AppId -Quiet
Update-MSGraphEnvironment -AuthUrl $authority -Quiet
Connect-MSGraph -ClientSecret $AppSecret -Quiet
 
$baselines=Invoke-MSGraphRequest -HttpMethod GET -Url $uri
$numberOfBaselines=$baselines.value.Count

#Baselines cleanup
$monthstodelete=((get-date).AddMonths(-$MonthstoKeepBaselines)).ToString("yyyy-MM-dd")
$uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)?"+'$filter'+"=createdDateTime%20lt%20$monthstodelete"
$baselinesToCleanUp=Invoke-MSGraphRequest -HttpMethod GET -Url $uri
$baselinesToCleanUp.value | foreach-object{
    $baselineID = $_.id
    $deleteUri = "https://graph.microsoft.com/$graphApiVersion/$($resource)/$baselineID"
    Invoke-MSGraphRequest -HttpMethod DELETE -Url $deleteUri | Out-Null
}

#Check if the 100 limit is reached then delete oldest year baselines
if($numberOfBaselines -ge 88){
        $baselines.value | select-object -last 12 | foreach-object{
        $baselineID = $_.id
        $deleteUri = "https://graph.microsoft.com/$graphApiVersion/$($resource)/$baselineID"
        Invoke-MSGraphRequest -HttpMethod DELETE -Url $deleteUri | Out-Null
    }
}
else {
$newBaseline=@"
{
    "displayName":"$($currentMonth)"
}
"@
#Create new Baseline
$newBaselineURL= "https://graph.microsoft.com/$graphApiVersion/$($resource)"
Invoke-MSGraphRequest -HttpMethod POST -Url $newBaselineURL -Content $newBaseline | out-null
}