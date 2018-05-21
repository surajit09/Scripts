#load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}



$web = Get-SPWeb "https://thethread.carpetright.co.uk/Facilities"

$listSiteInspectionForms = $web.Lists["RM Site Inspection Form"] 

$listSiteInspectionDashboard = $web.Lists["RM Site Inspection Dashboard"] 

$StoreExists=$false




	foreach($LstItem in $listSiteInspectionDashboard.Items)
		{
			$StoreExists=$false
		foreach ($item in $listSiteInspectionForms.Items)
		
		{
			if($item["Store Number"] -eq $LstItem["Store Number"])
			{	
			
				
				$StoreExists=$true
				break;
			}
		
		}

		if($StoreExists -eq $false)
		{
			$LstItem["Followed Up"]=$false
			$LstItem["RAS Comments"]=""
			$LstItem["Last Inspection Date"]=$null
			$LstItem.update()
		
		}
		
		
}


$web.Dispose()

