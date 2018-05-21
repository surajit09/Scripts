
#load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}

$web = Get-SPWeb "https://thethread.carpetright.co.uk/buying"
$list = $web.Lists["Sample Orders"]
$date=""

write-host $list.Items.count

foreach ($item in $list.Items)
{

	
  if ([String]::IsNullOrEmpty($item["Date of order"]) -eq $true)
{
  $item["Date of order"] = [System.DateTime]::Now;
  $item.Update();
}
}

$web.Dispose()


