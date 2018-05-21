$webUrl= "https://thethread.carpetright.co.uk/policies" 

$SPweb = Get-SPWeb $webUrl
$SPlist = $SPweb.Lists["Forum"]

 $myalerts = @()

$alerts=$spweb.alerts

    foreach($alert in $spweb.alerts)
    {
	
		write-host $alert.Title
        
			if($alert.Title -eq "Forum Alert")
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





