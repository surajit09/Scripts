# This script sends a reminder email to Network Team when a certificate , software or hardware is due to expire. The email will be sent 30 days before

# load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}


$web = Get-SPWeb "https://thethread.carpetright.co.uk/Network"

$Lists=@("Software Licensing & Support Database", "Hardware Support & Contract Database","Certificate Database")

$From = "thethread@carpetright.co.uk"
$To = "Networks@carpetright.co.uk"
#$Cc = "surajit.mukherjee@carpetright.co.uk"
#$bcc="surajit.mukherjee@carpetright.co.uk"
#$login = "carpetright_plc\healthandsafety"
#$password = "Alwaysabovespurs16" | Convertto-SecureString -AsPlainText -Force
#$credentials = New-Object System.Management.Automation.Pscredential -Argumentlist $login,$password
$Subject = "Certificate, Software or Hardware expiration reminder"
$Body =""


foreach ($element in $Lists) 
{
	
	$list = $web.Lists | where{$_.Title -eq $element}
	
if($list)
{

	
foreach ($item in $list.Items)
{
	
	if($list.Title -eq "Certificate Database")
	{
		if($item["Expires On"] -ne $null)
		{
			$ReminderDate=get-date $item["Expires On"] 
		}
		
	}
	else
	{
			if($item["Renewal Date"] -ne $null)
			{
				$ReminderDate=get-date $item["Renewal Date"] 
			}
		}
		
		if($ReminderDate)
		{
			$toDay= Get-date
			
			
			
			if($ReminderDate.ToShortDateString() -eq $toDay.AddDays(+30).ToShortDateString())
			{
				$body = "<HTML><HEAD><META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /><TITLE></TITLE></HEAD>"
				$body += "<BODY bgcolor=""#FFFFFF"" style=""font-size: Small; font-family: TAHOMA; color: #000000""><P>"
				
				if($list.Title -eq "Hardware Support & Contract Database")
				{
					$listItem=$item["Contracts"]
					$type="Hardware"
					
				}
				elseif($list.Title -eq "Certificate Database")
				{
					$listItem=$item["Certificate Name"]
					$type="Certificate"
				
				}
				else{
				
					$listItem=$item["Product"]
					$type="Software"
				
				}
				
				
			$Body += " Dear Colleague,"	+"<br>"
			$Body += "Please note the " +$type+"- " +$listItem + " will expire on " +$ReminderDate.ToShortDateString() + ".  <br>"
			 #$body += "Click <a href=$Link target=""_blank"">here</a> to open the form <br>"
			 $Body += "Please take appropriate action.<br>"
			$Body += "Please ignore this email if you have already taken action.<br>"
			$Body +="Thanks <br>"
			$Body +="Networks Team"


			Send-MailMessage -From $From -to $To  -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk 
			
}
		}
}

}

}


$web.Dispose()




 