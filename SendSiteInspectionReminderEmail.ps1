#This script sends reminder email to RM before site inspection is due. 
#This also sends the reminder email if the RM does not fill the site inspection form after the site inspection. 
#The script also send reminder email if RAS does not follow up after the site inspection


# load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}

$web = Get-SPWeb "https://thethread.carpetright.co.uk/Facilities"
$list = $web.Lists["RM Site Inspection Dashboard"]
$From = "thethread@carpetright.co.uk"
$To = ""
$Cc = ""
$Bcc="surajit.mukherjee@carpetright.co.uk"
$Link="https://thethread.carpetright.co.uk/Facilities/Lists/RM%20Site%20Inspection%20Form/AllItems.aspx"
$Dashboard="https://thethread.carpetright.co.uk/Facilities/Lists/RM%20Site%20Inspection%20Dashboard/AllItems.aspx"
$Subject = ""
$Body =""


foreach ($item in $list.Items)
{

if(([String]::IsNullOrEmpty($item["Next Inspection date"]) -eq $false) -and ([String]::IsNullOrEmpty($item["Last Inspection Date"]) -eq $false))
{
	$toDay= Get-date
	
	$InspectionDate=get-date $item["Next Inspection date"]
	$LastInspectionDate= get-date $item["Last Inspection Date"]
	
	
			if(($InspectionDate -lt $toDay.AddDays(+30)) -and ($InspectionDate -gt $toDay))
			{
				            
				$Subject = "Site inspection Reminder for " + $item["Title"] + " "+$item["Store Number"] 
				
				$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["RM"].ToString());
				$To=$userfield.User.Email
				$body = "<HTML><HEAD><META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /><TITLE></TITLE></HEAD>"
				$body += "<BODY bgcolor=""#FFFFFF"" style=""font-size: Small; font-family: TAHOMA; color: #000000""><P>"

			$Body += " Dear "+$userfield.User.DisplayName+" ,"	+"<br>"
			$Body += "You have a site inspection due for " + $item["Title"] + " "+$item["Store Number"] + " on " +$item["Next Inspection date"].ToShortDateString()+ ".  <br>" 
			
			$Body +="Thanks <br>"
			$Body +="Health and Safety Team"
			
			
			Send-MailMessage -From $From -to $To -Bcc $bcc -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk  
			
			$item["Completed"]=$false
			$item["Followed Up"]=$false
			$item["RAS Comments"]=""
			$item.update()
			
			}
			elseif (($InspectionDate -lt $toDay) -and ($item["Completed"] -eq $false) )
					{	
						
						
						
							
							$Subject = "Site inspection form Reminder for " + $item["Title"] + " "+$item["Store Number"] 
				
							$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["RM"].ToString());
							$To=$userfield.User.email;
							

							
						$body = "<HTML><HEAD><META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /><TITLE></TITLE></HEAD>"
						$body += "<BODY bgcolor=""#FFFFFF"" style=""font-size: Small; font-family: TAHOMA; color: #000000""><P>"

						$Body += " Dear "+$userfield.User.DisplayName+" ,"	+"<br>"
						$Body += "Please fill the site inspection form for " + $item["Title"] + " "+$item["Store Number"] + " store visit on " +$item["Next Inspection date"].ToShortDateString()+ ".  <br>" 
						$body += "Click <a href=$Link target=""_blank"">here</a> to go to Site inspection form library <br>"
						$Body +="Thanks <br>"
						$Body +="Health and Safety Team"
						$Cc="";
						
						if($InspectionDate -lt $toDay.AddDays(-30))
						{
							$DMfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["DM"].ToString());
							$Cc=$DMfield.User.email;
							$bcc="Charlotte.Nyandoro@carpetright.co.uk";
							Send-MailMessage -From $From -to $To -Cc $Cc -Bcc $bcc -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk  
						}
						else
						{
							
							Send-MailMessage -From $From -to $To -Bcc $bcc -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk
						
						}
			
					
			
			}
			elseif(($LastInspectionDate -lt $toDay) -and ($item["Completed"] -eq $true)){
			
			if($item["Followed Up"] -eq $false )
			{
				
					
					$Subject = "Site inspection form Reminder for " + $item["Title"] + " "+$item["Store Number"] 
				
				$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["RAS"].ToString());
				$To=$userfield.User.email
				
				$body = "<HTML><HEAD><META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /><TITLE></TITLE></HEAD>"
				$body += "<BODY bgcolor=""#FFFFFF"" style=""font-size: Small; font-family: TAHOMA; color: #000000""><P>"

			$Body += " Dear "+$userfield.User.DisplayName+" ,"	+"<br>"
			$Body += "Please update the follow up check box in the dashboard for " + $item["Title"] + " "+$item["Store Number"] + " store visit on " +$item["Last Inspection Date"].ToShortDateString()+ ".  <br>" 
			$body += "Click <a href=$Dashboard target=""_blank"">here</a> to go to the dashboard <br>"

			$Body +="Thanks <br>"
			$Body +="Health and Safety Team"
			$Cc="";
			
						if($LastInspectionDate -lt $toDay.AddDays(-60))
						{
							$DMfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web,$item["DM"].ToString());
							$Cc=$DMfield.User.email;
							$bcc="Charlotte.Nyandoro@carpetright.co.uk";
							Send-MailMessage -From $From -to $To -Cc $Cc -Bcc $bcc -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk  
						}
						else
						{
							Send-MailMessage -From $From -to $To -Bcc $bcc -Subject $Subject -bodyashtml  -Body $Body -SmtpServer relay.carpetright.co.uk 
						
						}
			  
				
				
			}
			
			}
			
	}		
}

$web.Dispose()






 