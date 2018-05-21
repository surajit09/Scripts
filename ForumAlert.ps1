
$webUrl= "https://thethread.carpetright.co.uk/policies"

$SPweb = Get-SPWeb $webUrl
$SPlist = $SPweb.Lists["Forum"]

$users =Get-SPUser -Web $webUrl | where {$_.LoginName -like "*store*"}
foreach ($user in $users){

  $alert = $user.Alerts.Add()
     $alert.Title = "Forum Alert"
     $alert.AlertType = [Microsoft.SharePoint.SPAlertType]::List
     $alert.List = $SPlist
     $alert.DeliveryChannels = [Microsoft.SharePoint.SPAlertDeliveryChannels]::Email
     $alert.EventType = [Microsoft.SharePoint.SPEventType]::Add
     $alert.AlertFrequency = [Microsoft.SharePoint.SPAlertFrequency]::Immediate
     $alert.Update()
}


$SPweb.Dispose()