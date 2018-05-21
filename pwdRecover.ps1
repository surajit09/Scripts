#Begin Setting Script Variables

#$AccountToRetrieve = "Domain\User"


$AccountToRetrieve = "carpetright_plc\spprd_cfg"


#Create Functions

Function VerifyTimerJob ($Filter)
{
$Timer = Get-SPTimerJob | ? {$_.displayname -like $Filter}
If ($Timer)
{
$timer.Delete()
}
}

#Begin Script

$Farm = get-spfarm | select name

$Configdb = Get-SPDatabase | ? {$_.name -eq $Farm.Name.Tostring()}

$ManagedAccount = get-SPManagedAccount $AccountToRetrieve

$WebApplication = new-SPWebApplication -Name "Temp Web Application" -url "https://contoso.carpetright.co.uk" -port 80 -AuthenticationProvider (New-SPAuthenticationProvider) -DatabaseServer $Configdb.server.displayname -DatabaseName TempWebApp_DB -ApplicationPool "Password Retrieval" -ApplicationPoolAccount $ManagedAccount -hostheader "https://contoso.carpetright.co.uk"

$Password = cmd.exe /c $env:windir\system32\inetsrv\appcmd.exe list apppool "Password Retrieval" /text:ProcessModel.Password

Write-Host "Password for Account "  $AccountToRetrieve  " is " $Password

$Filter = "Unprovisioning *" + $Webapplication.Displayname + "*"

VerifyTimerJob($Filter)
Remove-SPWebApplication $WebApplication -DeleteIISSite -RemoveContentDatabases -Confirm:$False
VerifyTimerJob($Filter)

$ProvisionJobs = Get-SPTimerJob | ? {$_.displayname -like "provisioning web application*"}
if ($ProvisionJobs)
{
    foreach ($ProvisionJob in $ProvisionJobs)
    {
        $ProvisionJob.Delete()
    }
}