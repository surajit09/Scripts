#load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}



$web = Get-SPWeb "https://thethread.carpetright.co.uk/Facilities"


$lstSiteInspectionDashboard = $web.Lists["RM Site Inspection Dashboard"] 
$Arr=@()


$toDay= Get-date
$Body=""
$Forms=@("Training Matrix", "Airport Steps","Roller Shutter Door","Beds On Display","Carpet Barrow Check","Step Ladder Check","Easy Lift Check","Roll Stock Stands","Pallet Truck Check","Vacuum Cleaner Check","Warehouse Racking Check")  
$From = "thethread@carpetright.co.uk"
$To = ""
$Cc = ""
$Bcc="surajit.mukherjee@carpetright.co.uk"
$Subject = "Quarterly report for health and safety forms"
$isFilledAtAll=$false

  foreach($item in $lstSiteInspectionDashboard.Items)
  {
		
		$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["RM"].tostring());
			if(([String]::IsNullOrEmpty($userfield.tostring()) -eq $false))
			{
				$userAccount=$userfield.User.DisplayName
				
				$Arr+=$userAccount
				
			}
			
			
			
  }
$Arr=$Arr | sort  | Get-Unique

#write-host $Arr

foreach ($rm in $Arr)
{

	$Body += " Dear "+$rm+" ,"	+"<br>"
	$Stores=@()
	foreach($item in $lstSiteInspectionDashboard.Items)
	{
		
			$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["RM"].tostring());
			$userAccount=$userfield.User.DisplayName
			$userEmail=$userfield.User.Email
				if($userAccount -eq $rm)
				{
					
					$Stores+=$item["Store Number"]
					
					$rmEmail=$userEmail
					
					$RASfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["RAS"].tostring());
					if(([String]::IsNullOrEmpty($RASfield.tostring()) -eq $false))
					{
						$RASAccount=$RASfield.User.DisplayName
						$RASEmail=$RASfield.User.Email
						
						#write-host $rm $rmEmail $RASAccount $RASEmail
						
					}
				}
			
	
	}
	
	
	
	$StoresToEmail=@()
	
	foreach ($element in $Forms) 
{
	
	$list = $web.Lists | where{$_.Title -eq $element}
	
	
	
if($list)
{
	
	$Body +="List of stores that have not filled the "+$list.Title+" form last quarter: "	+"<br>"
	
		
			foreach($store in $Stores)
			{
				$isFilledAtAll=$false
				#write-host  $store
				$DateArr=@()
			foreach($item in $list.Items)
			{
				if($item.File.Author -Match $store)
				{
					$DateArr+=$item["Modified"]
				}
			
			
			}
		
		if($DateArr.length -gt 1) 
			{
				$isFilledAtAll=$true
				$DateArr=$DateArr|sort -Descending
				$LateUpdateDate=$DateArr[0].ToShortDateString();
				
				#write-host (get-date $LateUpdateDate) $toDay 
				$LastDate=get-date $LateUpdateDate
				$days=( $toDay-$LastDate).Days
	
				#if( (get-date $LateUpdateDate) -gt ($toDay.AddDays(-90)))
				if( $days -gt  92)
				{
					#write-host "Sending email for this"
					foreach($item in $lstSiteInspectionDashboard.Items)
					{
						if($item["Store Number"] -eq $store)
						{
						
							$storeName=$item["Title"]
						}
					}
			
					if($StoresToEmail.length -ge 1)
					{
						$StoresToEmail+="; "+$storeName+" "+$store  
					}
					else{
						$StoresToEmail=$storeName+" "+$store
					}
					
				}
			}
	
	
	
		<#if($isFilledAtAll -eq $false )
		{
			#write-host "Sending email for this"
			foreach($item in $lstSiteInspectionDashboard.Items)
					{
						if($item["Store Number"] -eq $store)
						{
						
							$storeName=$item["Title"]
						}
					}
					if($StoresToEmail.length -ge 1)
					{
						$StoresToEmail+="; "+$storeName+" "+$store
					}
					else{
						$StoresToEmail=$storeName+" "+$store
					}
		}#>
		
	}
	
	
	

$Body +=$StoresToEmail +"<br>"
$Body +="<br>"
$Body +="<br>"

$StoresToEmail=""
}

}


			$Body +="Thanks <br>"
			$Body +="Health and Safety Team"
			
		#if($rm.contains("Bradley"))
		#{
			$To=$rmEmail 
			$Cc = $RASEmail
			
			#$To="surajit.mukherjee@carpetright.co.uk"
			
			Send-MailMessage -From $From -to $To -Cc $Cc  -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk
			
			#Send-MailMessage -From $From -to $To -Subject $Subject -bodyashtml  -Body $Body -SmtpServer csombxex01.uk.cruk.net
			#break;
		#}
		
		$Body=""
		
}

$web.Dispose()