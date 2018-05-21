

# load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}


$MyWeb =get-SPWeb "https://thethread.carpetright.co.uk"
$MyWeb2 =get-SPWeb "https://thethread.carpetright.co.uk/Facilities"

$listName1="Store Contact Information"
$listName2="HS Forms Status"
$listName3="HS Forms Frequency"

$LateUpdateDate=""
$StoreAccount=""
$StoreNumber=""
$userNumber=""

$Stores=@("Manager's Weekly Check", "Fire Alarm Check","Emergency Evacuation Log","Monthly Light Test","Training Matrix","Yearly Planner","Airport Steps","Roller Shutter Door","Beds On Display","PEDESTRIAN BOOM TRUCK CHECKS","Carpet Barrow Check","Roll Stock Stands","Step Ladder Check","Easy Lift Check","Warehouse Racking Check","Vacuum Cleaner Check")



$list2 = $MyWeb2.Lists | where{$_.Title -eq $listName2}

$list3 = $MyWeb2.Lists | where{$_.Title -eq $listName3}


$StoreExists=$false;

# Break out if the list has no content. Stops the creation of empty files.

	





foreach ($element in $Stores) 
{
	
	$list = $MyWeb2.Lists | where{$_.Title -eq $element}
	
if($list)
{


 
foreach($item in $List2.items)
{
	
	$LateUpdateDate=""
	$StoreNumber=$item["StoreNumber"]
	
	#$StoreName=$item["StoreName"]
	$StoreExists=$false;
	$DateArr=@()
	<#
	foreach($item1 in $list.Items)
	{	
	
		$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($MyWeb2,$item1["Modified By"].ToString());            
		$userAccount=$userfield.User.DisplayName
		$index=$userAccount.length-4
		$userNumber=$userAccount.substring($index)
		
		#$userName=$userAccount.substring(0,$index)
		#write-host $userName $userNumber
		if($userNumber.tostring() -eq $StoreNumber.tostring())
		{
			$DateArr+=$item1["Modified"]
			
		}
		
	}	
	#>
	foreach($item1 in $list.Items)
	{	
		if($item1.File.ModifiedBy -Match $StoreNumber)
		{
			write-host "Found Match " $StoreNumber   
			$DateArr+=$item1["Modified"]
		
		}
	
	}
	
	if($DateArr.length -ge 1) 
	{
		
		#$DateArr=$DateArr|sort -Descending
		$StoreExists=$true
		
		$LateUpdateDate=$DateArr|sort  | Select -Last 1
		
		$LateUpdateDate=$LateUpdateDate.ToShortDateString();
		
		write-host $LateUpdateDate $StoreNumber
		#$LateUpdateDate=$DateArr[0].ToShortDateString();
		
		
		if([String]::IsNullOrEmpty($LateUpdateDate))
		{
			continue;
		}
		
		if($StoreExists -eq $true)
		{
			if([String]::IsNullOrEmpty($item["Last Update"]))
		{
			
			#$item["Last Update"]="{"+$element+":"+$LateUpdateDate+"}"
			#$item.update()
			
		
		}
		else{
			
			#$item["Last Update"]+="{"+$element+":"+$LateUpdateDate +"}"
			#$item.update()
		}
		}
		 
	
	else
	{
	
	if([String]::IsNullOrEmpty($item["HS Form Name"]))
		{
			
			#$item["HS Form Name"]=$element
			#$item.update()
		
		}
		else{
			
			#$item["HS Form Name"]+=";" +$element
			#$item.update()
		}
	
	
	
	}
		
}
	foreach($ManagerItem in $items)
	{	
	
		
		$StoreAccount=$ManagerItem["Modified By"] 
		
		$userObj = New-Object Microsoft.SharePoint.SPFieldUserValue($MyWeb, $StoreAccount)
		$accountName = $userObj.User.UserLogin
		$accountName=$accountName.substring($accountName.IndexOf("\")+1)
		$accountNumber=$accountName.remove(0,5)
		
		if($accountNumber -eq $StoreNumber)
		{
			
			$StoreExists=$true
			$LateUpdateDate=$ManagerItem["Modified"]
			
			
			
			break;
		}
	
	
	
	}

	#>
	

}
}
	
}
$MyWeb.Dispose();
$MyWeb2.Dispose();
