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
$Forms=@("Monthly Light Test", "Yearly Planner")  
$From = "thethread@carpetright.co.uk"
$To = ""
$Cc = ""
$Bcc="surajit.mukherjee@carpetright.co.uk"
$Subject = "Monthly report for health and safety forms"
$isFilledAtAll=$false
$StoresToEmail=""

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
	
	
	
$StoresToEmail=""
	
	foreach ($element in $Forms) 
{
	
	$list = $web.Lists | where{$_.Title -eq $element}
	
	
	
if($list)
{
	
	$Body +="List of stores that have not filled the "+$list.Title+" form last month: "	+"<br>"
	
		
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
				#write-host $LateUpdateDate
				$LastDate=get-date $LateUpdateDate
				#write-host $LastDate
				$days=( $toDay-$LastDate).Days
				#write-host $days
				#write-host (get-date $LateUpdateDate) $toDay 
				#write-host $StoresToEmail
				#if( (get-date $LateUpdateDate) - $toDay.AddDays(-30)))
				if( $days -gt  31)
				{
					#write-host "Sending email for this" $store
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