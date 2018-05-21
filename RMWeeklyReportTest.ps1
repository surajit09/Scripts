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
$Forms=@("Manager's Weekly Check", "Fire Alarm Check","SPRINKLER TEST LOG")  
$From = "thethread@carpetright.co.uk"
$To = ""
$Cc = ""
$Bcc="surajit.mukherjee@carpetright.co.uk"
$Subject = "Weekly report for health and safety forms"
$isFilledAtAll=$false

$SprinklerStores=@("0120","0504","0244","0156","5002","0746","1148","0745","0532","1184")



#Get the RMs
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
#Get the stores for each RM


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
	
	#send  email for each form if the form is not filled for more than 7days
	
	foreach ($element in $Forms) 
{
	
	$list = $web.Lists | where{$_.Title -eq $element}
	
	
	
if($list)
{
	write-host $list
	$Body +="List of stores that have not filled the "+$list.Title+" form last week: "	+"<br>"
	
		
			foreach($store in $Stores)
			{
				if($element -eq "SPRINKLER TEST LOG") 
					{
						$found=$false
						foreach($SprinklerStore in	$SprinklerStores)
						{
							if($SprinklerStore -eq $store)
							{
								$found=$true		
							}
								
						}
						
						if($found -eq $false)
						{
							continue
						}
								
					}
				$isFilledAtAll=$false
				#write-host  $store
				$DateArr=@()
					foreach($item in $list.Items)
					{
						#write-host $item.File.Author 
						if($item.File.Author -Match $store)
						{
							write-host $item.File.Author  $store
							$DateArr+=$item["Created"]
						}
					
					
					}
		
					if($DateArr.length -gt 1) 
					{
						$isFilledAtAll=$true
						$DateArr=$DateArr|sort -Descending
						$LateUpdateDate=$DateArr[0].ToShortDateString();
						
						write-host (get-date $LateUpdateDate) $toDay 
						$LastDate=get-date $LateUpdateDate
						$days=( $toDay-$LastDate).Days
						
						#if( (get-date $LateUpdateDate) -lt ($toDay.AddDays(-9)))
						if( $days -gt  9)
						{
							write-host "Sending email for this"
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
			#Send-MailMessage -From $From -to $To -Cc $Cc  -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk
			#break;
		#}
		
		$Body=""
}

$web.Dispose()