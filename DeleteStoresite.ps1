#This  script is used to delete the store sites

# load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}

write-host "This will delete store site, type store number to continue"
$retval = read-host

$webUrl="https://thethread.carpetright.co.uk/" + $retval 

$web = get-spweb https://thethread.carpetright.co.uk/
$list = $web.lists | where {$_.title -eq "Stores"}
$items = $list.items 

$site = Get-SPWeb $webUrl -ErrorVariable err -ErrorAction SilentlyContinue -AssignmentCollection $assignmentCollection

if ($err)
{
   
   Write-Host "Site does not exists or the store number is wrongly typed!! Please note you need to remove the leading zeros in the store number"
}
else
{
   
   
   write-host "continuing to delete $webUrl "

	Remove-SPWeb  -Identity $webUrl  -Confirm:$true

	Write-host "  Say Goodbye to $webUrl" -foregroundcolor red
	
	$item = $list.items | where {$_."Store ID" -eq $retval}
	if($item)
	{
		Write-host "  Say Goodbye to $($item.id) item from the Stores list " -foregroundcolor red
		$list.getitembyid($Item.id).Delete()
	}
	
}

