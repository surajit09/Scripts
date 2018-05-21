$webUrl= "https://thethread.carpetright.co.uk/policies" 

$SPweb = Get-SPWeb $webUrl
$SPlist = $SPweb.Lists["Forum"]

$IDS = ""
    foreach($alert in $spweb.alerts)
    {
        if($alert.ListID -eq $SPlist.ID)
        {
			write-host $alert.Title
			
			
        
        }
        
    }
write-host "deleting..."
    foreach($s in $IDS.Split("|"))
    {
		write-host -nonewline "*"
		
    }




