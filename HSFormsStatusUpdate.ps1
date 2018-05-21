

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

$LateUpdateDate=""
$StoreAccount=""

$Stores=@("Manager's Weekly Check", "Fire Alarm Check","Emergency Evacuation Log","Monthly Light Test","Training Matrix","Yearly Planner","Airport Steps","Roller Shutter Door","Pallet Truck Check","Beds On Display","PEDESTRIAN BOOM TRUCK CHECKS","Carpet Barrow Check","Roll Stock Stands","Step Ladder Check","Warehouse Racking Check","Vacuum Cleaner Check","Carpet Manipulator","Easy Lift Check")

$list1 = $MyWeb.Lists | where{$_.Title -eq $listName1}

$list2 = $MyWeb2.Lists | where{$_.Title -eq $listName2}

$textlog="";
$StoreExists=$false;

# Break out if the list has no content. Stops the creation of empty files.
if ($list1) {
	

if($List2)
{


if($List2.items.count -gt 1)
{
foreach ($item in $List2.items)
{
    
    $List2.getitembyid($Item.id).Delete()
}
}

foreach ($Myitem in $List1.items)
{


 if(($Myitem["Name"] -like "Mobile Terminal*" ) -or ($Myitem["Reg Code"] -like "POLAND*" ) -or ($Myitem["Reg Code"] -eq "9900"))
 {
	continue;
	}
 
	$newItem = $List2.Items.Add()


	$newItem["Division"]=$Myitem["Division Code"]
	$newItem["Region"]=$Myitem["Reg Code"]
	$newItem["StoreName"]=$Myitem["Name"]
	$newItem["StoreNumber"]=$Myitem["Site"]


	$newItem.Update()
}

}



foreach ($element in $Stores) 
{
	
	$list = $MyWeb2.Lists | where{$_.Title -eq $element}
	
if($list)
{

try
{

foreach($item in $List2.items)
{
	
	$LateUpdateDate=""
	$StoreNumber=$item["StoreNumber"]
	#$StoreName=$item["StoreName"]
	$StoreExists=$false;
	$DateArr=@()
	
	foreach($item1 in $list.Items)
	{	
		
		$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($MyWeb2,$item1["Created By"].ToString());            
		$userAccount=$userfield.User.DisplayName
		$index=$userAccount.length-4
		$userNumber=$userAccount.substring($index)
		#$userName=$userAccount.substring(0,$index)
		#write-host $userName $userNumber
		if($userNumber -eq $StoreNumber)
		{
			$DateArr+=$item1["Modified"]
			
		}
			
	}	
	if($DateArr.length -gt 0) 
	{
	
	
		$DateArr=$DateArr|sort -Descending 
		$LateUpdateDate= $DateArr[0].ToShortDateString();
		$StoreExists=$true

		<#if([String]::IsNullOrEmpty($LateUpdateDate))
		{
			continue;
		}#>
		
		if($StoreExists -eq $true)
		{
			if([String]::IsNullOrEmpty($item["Last Update"]))
		{
			
			$item["Last Update"]="{"+$element+":"+$LateUpdateDate+"}"
			$item.update()
			
		
		}
		else{
			
			$item["Last Update"]+="{"+$element+":"+$LateUpdateDate +"}"
			$item.update()
		}
		}
		 
	}
	else
	{
	
	if([String]::IsNullOrEmpty($item["HS Form Name"]))
		{
			
			$item["HS Form Name"]=$element
			$item.update()
		
		}
		else{
			
			$item["HS Form Name"]+=";" +$element
			$item.update()
		}
	
	
	
	}
		

}		

}
catch
{

	
	write-host "Caught an exception:" -ForegroundColor Red
   write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
   write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
	write-host  $element $LateUpdateDate $StoreNumber $userNumber
	
	$ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Send-MailMessage -From thethread@carpetright.co.uk -To surajit.mukherjee@carpetright.co.uk -Subject "Error occurred updating the HS Forms Status!" -SmtpServer relay.carpetright.co.uk  -Body "Error occurred at $FailedItem. The error message was $ErrorMessage Other information:$element $LateUpdateDate $StoreNumber $userNumber"
    


}

}
}
	

}





$MyWeb.Dispose();
$MyWeb2.Dispose();
