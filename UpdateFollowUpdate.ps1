
#load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}






$web = Get-SPWeb "https://thethread.carpetright.co.uk/Facilities"



$listSiteInspectionDashboard = $web.Lists["RM Site Inspection Dashboard"] 

	
foreach ($item in $listSiteInspectionDashboard.Items)
{

if ([String]::IsNullOrEmpty($item["Last Inspection Date"]) -eq $false )
 {
	
	 $LastInspectionDate=[System.DateTime]$item["Last Inspection Date"]

	$item["Follow Up Date"]=$LastInspectionDate.AddDays(1)
	$item.update()
	


}
}


$web.Dispose()

