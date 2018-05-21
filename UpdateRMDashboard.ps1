
#load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}






$web = Get-SPWeb "https://thethread.carpetright.co.uk/Facilities"

$listSiteInspectionForms = $web.Lists["RM Site Inspection Form"] 

$listSiteInspectionDashboard = $web.Lists["RM Site Inspection Dashboard"] 

$stores = @()

$toDay= Get-date 




foreach ($item in $listSiteInspectionForms.Items)
{

	$Created= $item["Created"] 
	
	
	
	if( $Created.ToShortDateString() -eq $toDay.ToShortDateString())
	{
	
		
		 
		 foreach($LstItem in $listSiteInspectionDashboard.Items)
		{
		
		if($item["Store Number"] -eq $LstItem["Store Number"])
		{	
		
			
			$InspectionDate=[System.DateTime]$item["Inspection Date"]
		
			
			$LstItem["Next Inspection date"]= $InspectionDate
			$LstItem["Last Inspection Date"]= $InspectionDate
			
			$LstItem["Completed"]=$true
			$LstItem["Followed Up"]=$false
			$LstItem["RAS Comments"]=""
			$LstItem.update()
			break;
		}
	
		}
		 
		
	}

}

 

foreach ($item in $listSiteInspectionDashboard.Items)
{

if (([String]::IsNullOrEmpty($item["Last Inspection Date"]) -eq $false) -and ([String]::IsNullOrEmpty($item["Next Inspection date"])-eq $false))
 {
	 $InspectionDate=[System.DateTime]$item["Next Inspection date"]
	 $LastInspectionDate=[System.DateTime]$item["Last Inspection Date"]

if( ($InspectionDate.ToShortDateString() -eq $LastInspectionDate.ToShortDateString()) -and ($item["Completed"] -eq $true) -and ($item["Followed Up"]=$true) )
{
	$NextInspectionDate=$InspectionDate.AddMonths(+6)
	$item["Next Inspection date"]=$NextInspectionDate
	$item.update()
	

}
}
}


$web.Dispose()

