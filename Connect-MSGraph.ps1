Connect-MSGraph
Update-MSGraphEnvironment -SchemaVersion beta
Invoke-MSGraphRequest -HttpMethod GET -Url ‘deviceManagement/deviceHealthScripts’ | Select-Object -ExpandProperty value  `
| where {$_.displayname -like "M365 Apps Inventory"} | Select-Object DisplayName, ID | FT
$script="f331323b-2bab-4727-9853-b8324e8cf1ec"
$Details = Invoke-MSGraphRequest -HttpMethod GET -Url "deviceManagement/deviceHealthScripts/$script/deviceRunStates?"
$output=$Details.value.preRemediationDetectionScriptOutput
$output | ConvertFrom-Json  | export-csv C:\Users\sebassem\Desktop\M365Inventory.csv -NoTypeInformation

