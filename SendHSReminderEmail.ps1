# This script sends a reminder email to all stores to fill health and safety forms on weekly , monthly or half-yearly basis

# load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}


$web = Get-SPWeb "https://thethread.carpetright.co.uk/Facilities"
$list = $web.Lists["Health And Safety Forms Reminder"]
$From = "Health.Safety@carpetright.co.uk"
$To = "All-Stores@carpetright.co.uk"
$Cc = "Charlotte.Nyandoro@carpetright.co.uk", "AllStoreManagers@carpetright.co.uk","Lorraine.Watler@carpetright.co.uk"
$bcc="surajit.mukherjee@carpetright.co.uk","AberdeenWarehouse@carpetright.co.uk"

$login = "carpetright_plc\healthandsafety"
$password = "Alwaysabovespurs16" | Convertto-SecureString -AsPlainText -Force
$credentials = New-Object System.Management.Automation.Pscredential -Argumentlist $login,$password
$Subject = "Health and Safety check list Reminder"
$Body =""
$isRun="No"
$ToSprinklerEmail="Dundee0120@carpetright.co.uk","Edinburgh0504@carpetright.co.uk","Livingston0244@carpetright.co.uk",
"Stirling0156@carpetright.co.uk","Dumfries5002@storeycarpets.co.uk",
"Glasgow0746@carpetright.co.uk","Hamilton1148@carpetright.co.uk","Pollokshaws0745@carpetright.co.uk",
"Bournemouth0532@carpetright.co.uk","Fareham1184@carpetright.co.uk"


foreach ($item in $list.Items)
{


	$toDay= Get-date
	$ReminderDate=$item["Reminder Date"]
	$CourseName=$item["Course Name"]		
	
			
			
			
			if($ReminderDate.ToShortDateString() -eq $toDay.ToShortDateString())
			{
				$body = "<HTML><HEAD><META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /><TITLE></TITLE></HEAD>"
				$body += "<BODY bgcolor=""#FFFFFF"" style=""font-size: Small; font-family: TAHOMA; color: #000000""><P>"
				$Url=new-object Microsoft.SharePoint.SPFieldUrlValue($item["Form Link"])
				$Link=$Url.URL
				
			$Body += " Dear Colleague,"	+"<br>"
			$Body += "Please complete the " + $item["Course Name"] + ".  <br>"
			 $body += "Click <a href=$Link target=""_blank"">here</a> to open the form <br>"
			$Body += "Please ignore this email if you have already completed the check list.<br>"
			$Body +="Thanks <br>"
			$Body +="Health and Safety Team"


			Send-MailMessage -From $From -to $To -Cc $Cc -Bcc $bcc -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk -Credential $credentials
			$isRun="Yes"
}
}

if ($isRun -eq "Yes")
{
foreach ($item in $list.Items)
{
	$CourseName=$item["Course Name"]

if(($CourseName -eq "weekly Sprinkler Check") -or ($CourseName -eq "weekly Fire Alarm Check") -or ($CourseName -eq "Manager's checks weekly"))
			{
				$body = "<HTML><HEAD><META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /><TITLE></TITLE></HEAD>"
				$body += "<BODY bgcolor=""#FFFFFF"" style=""font-size: Small; font-family: TAHOMA; color: #000000""><P>"
				$Url=new-object Microsoft.SharePoint.SPFieldUrlValue($item["Form Link"])
				$Link=$Url.URL
				
			$Body += " Dear Colleague,"	+"<br>"
			$Body += "Please complete the " + $item["Course Name"] + ".  <br>"
			 $body += "Click <a href=$Link target=""_blank"">here</a> to open the form <br>"
			 if($CourseName -eq "Manager's checks weekly")
			 {
				$Body +=  "Please note you can take a print out of the form if you find that is convenient.<br>"
			 }
			 $Body += "Please ignore this email if you have already completed the check list.<br>"
			$Body +="Thanks <br>"
			$Body +="Health and Safety Team"
			
			
			if($CourseName -eq "weekly Sprinkler Check")
			{
				$To=$ToSprinklerEmail
				$Cc = "Charlotte.Nyandoro@carpetright.co.uk"
			
			}
			else{
				$To = "All-Stores@carpetright.co.uk"
				$Cc = "Charlotte.Nyandoro@carpetright.co.uk", "AllStoreManagers@carpetright.co.uk"
			
			}


			Send-MailMessage -From $From -to $To -Cc $Cc -Bcc $bcc -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk -Credential $credentials
	
}
}
}



$web.Dispose()




 