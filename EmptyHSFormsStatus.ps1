


$MyWeb2 =get-SPWeb "https://thethread.carpetright.co.uk/Facilities"
$listName2="HS Forms Status"
$list2 = $MyWeb2.Lists | where{$_.Title -eq $listName2}

if($List2)
{
if($List2.items.count -gt 1)
{
foreach ($item in $List2.items)
{
    $item["HS Form Name"]=""
	 $item["Last Update"]=""
	$item.update()
    
}
}
}

$MyWeb2.dispose()