$webUrl= "https://thethread.carpetright.co.uk/Facilities" 

$SPweb = Get-SPWeb $webUrl
$SPlist = $SPweb.Lists["RM Site Inspection Form"]

 $myalerts = @()

$alerts=$spweb.alerts

    foreach($alert in $spweb.alerts)
    {
	
		write-host $alert.Title
        
			if($alert.Title -eq "RM Site Inspection Form")
			{
				
				 $myalerts += $alert
				
				
				
			
			}
			
			
        
        
        
    }
  ### now we have alerts for this site, we can delete them

                foreach ($alertdel in $myalerts)

                {

                    $alerts.Delete($alertdel.ID)

            		write-host $alertdel.ID

                }





