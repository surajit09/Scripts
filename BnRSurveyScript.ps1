

 
$web = get-spweb https://thethread.carpetright.co.uk/buying 
$list = $web.lists | where {$_.title -eq "Beds And Rugs Survey"}
Write-host "List $($list.title) has $($list.items.count) entries"

$arrSurvey = @()



$items = $list.items

foreach ($item in $items)
{
	
	
	$obj = New-Object System.Object
		
	
	$userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($web ,$item["Created By"].ToString()); 

		$obj | Add-Member -type NoteProperty -name "Store number" -value $userfield.User.DisplayName
  
	$arrSurvey += $obj
   
}


$arrSurvey | Format-Table –AutoSize >> D:/BnRSurvey.txt












 