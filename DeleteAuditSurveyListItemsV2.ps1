
# load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}

write-host "This will delete data, type YES to continue"
$retval = read-host 
if ($retval -ne "YES") 
{
    write-host "exiting - you did not type yes" -foregroundcolor green
    exit
}
write-host "continuing"
 
$web = get-spweb https://thethread.carpetright.co.uk/Audit 
$list = $web.lists | where {$_.title -eq "Audit Survey"}
Write-host "List $($list.title) has $($list.items.count) entries"
$items = $list.items
foreach ($item in $items)
{
    Write-host "  Say Goodbye to $($item.id)" -foregroundcolor red
    $list.getitembyid($Item.id).Delete()
}

