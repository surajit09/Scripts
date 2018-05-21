#load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}



$web = Get-SPWeb "https://thethread.carpetright.co.uk/Facilities"



$Arr=@()

$Forms=@("Manager's Weekly Check", "Fire Alarm Check","SPRINKLER TEST LOG")  

$isFilledAtAll=$false

$SprinklerStores=@("0120","0504","0244","0156","5002","0746","1148","0745","0532","1184")


	
$list = $web.Lists | where{$_.Title -eq "SPRINKLER TEST LOG"}
	
	
	
if($list)
{
	
			foreach($item in $list.Items)
			{
				
						$found=$false
						foreach($SprinklerStore in	$SprinklerStores)
						{
							if($item["Created By"] -Match $SprinklerStore)
							{
								$found=$true
								
								
								
							}
								
						}
						
						if($found -eq $false)
						{
							#write-host $item["Created By"]
							$list.getitembyid($Item.id).Delete()
						}
								
					
				}


			
}

$web.Dispose()