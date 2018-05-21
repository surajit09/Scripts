

 
$web = get-spweb https://thethread.carpetright.co.uk/buying 
$list = $web.lists | where {$_.title -eq "Beds And Rugs Survey"}
Write-host "List $($list.title) has $($list.items.count) entries"

$arrSurvey = @()
$arrStoreContact = @()


$items = $list.items
$Temp=""
foreach ($item in $items)
{
	$Temp=$item["Created By"]
    	$arrSurvey +=$temp.substring($temp.Length-4)

   
}

$table = @()


$webs = get-spweb https://thethread.carpetright.co.uk 
$listA = $webs.lists | where {$_.title -eq "Store Contact Information"}
Write-host "List $($listA.title) has $($listA.items.count) entries"
$items = $listA.items

foreach ($item in $items)
{
if(($item["Site"] -notmatch "50*") -And ($item["Site"] -notmatch "40*") -And ($item["Site"] -notmatch "90*"))
{
	if($item["Site"] -NotIn $arrSurvey)
	{
		$arrStoreContact +=  $item["Site"] 
		$obj = New-Object System.Object
		$obj | Add-Member -type NoteProperty -name "Store number" -value $item["Site"]
		$obj | Add-Member -type NoteProperty -Name "Store Name" -Value $item["Name"]
		$obj | Add-Member -type NoteProperty -Name "Division Code" -Value $item["Division Code"]
		$obj | Add-Member -type NoteProperty -Name "Reg Code" -Value $item["Reg Code"]

		$table += $obj
    	
   }
}
}

Write-host " $($table.count) stores need to fill surveys"

$table | Format-Table –AutoSize >> D:/BnRSurvey.txt











 