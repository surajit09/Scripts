

$web = Get-SPWeb "https://thethread.carpetright.co.uk/Facilities"



$DM=$web.EnsureUser("carpetright_plc\sspk")


$listRMDetails = $web.Lists["RM Site Inspection Dashboard"] 



foreach ($item in $listRMDetails.Items)
{
	
	$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["RM"].ToString());            
	  
	if($userfield.User.DisplayName -eq 'Craig Watson')
	{
		
		$item["DM"]=$DM


		$item.update()
		
	
	}


}