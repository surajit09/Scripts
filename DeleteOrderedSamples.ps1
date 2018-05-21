
#load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}

$web = Get-SPWeb "https://thethread.carpetright.co.uk/buying"
$list = $web.Lists["Sample Orders"]
$date=""
$toDay= Get-date 
write-host $list.Items.count

foreach ($item in $list.Items)
{

	
  if ([String]::IsNullOrEmpty($item["Date of order"]) -eq $false)
{
	$Ordered= $item["Date of order"] 
	if($toDay -gt $Ordered.AddDays(+90))
	{
		$List.getitembyid($Item.id).Delete()
		#write-host $Ordered
	
	}
   
}
}

$web.Dispose()


