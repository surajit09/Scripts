


$wa = Get-SPWebApplication https://thethread.carpetright.co.uk
$wa.CrossDomainPhotosEnabled = $true
$wa.Update()
