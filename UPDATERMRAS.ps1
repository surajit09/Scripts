$web = Get-SPWeb "https://thethread.carpetright.co.uk/Facilities"

#$RM = $web.EnsureUser("carpetright_plc\ssna")

#$RAS=$web.EnsureUser("carpetright_plc\ssvs")

$RAS=$web.EnsureUser("Laura Owen")

$listRMDetails = $web.Lists["RM Site Inspection Dashboard"] 



foreach ($item in $listRMDetails.Items)
{
	
	$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["RAS"].ToString());            
	  
	if($userfield.User.DisplayName -eq 'Sandra Cook')
	{
		
		$item["RAS"]=$RAS


		$item.update()
		
	
	}


}