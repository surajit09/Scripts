# This script checks the contents of document libraries every hour and adds Urls for new documents to announcement list


# load SharePoint Powershell snap-in if not already loaded
$snapin = get-pssnapin | where { $_.Name -eq "Microsoft.SharePoint.PowerShell" }
if($snapin -eq $null)  
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell
}

# Array of the document libraries that store content for store announcements

$DocLibArray = "Updates","Buying Carpets","Buying Beds","Buying Hardflooring","Customer Service", "Marketing Bulletin","Concept Stores","POS Packs","Operations Bulletin", "Peak Trade Communications","News Letters","Operations Procedures","HR Policies and Procedures","Health and Safety policies","Health and Safety Bulletin", "Executive Memos", "Facilities Policies", "H&S Manual (Check Lists)","Risk Assessments", "Training",  "IT Policies" ,"Customer Service","Buying Policies","Finance Policies", "Company Announcements", "eCommerce Policies","Marketing Policies","Warehouse Policies","Templates", "Regional Documents"



$siteURL = "https://thethread.carpetright.co.uk"
$web = get-SPWeb $siteURL 
$AnnouncementList=$web.lists["Announcements"]
$UKAnnouncementList=$web.lists["UK Announcements"]
$IEAnnouncementList=$web.lists["IE Announcements"]
$StoreysAnnouncementList=$web.lists["Storeys Announcements"]

$AnnouncementItems = $AnnouncementList.GetItems()


$site = Get-SPSite($siteURL)



foreach($web in $site.AllWebs) 
{
    foreach($list in $web.Lists)
    {
        if($list.BaseType -eq "DocumentLibrary")
	

        {
        $listName = $list.Title
		
	
	if ($DocLibArray -contains $listName) 
	{
		foreach ($listItem in $list.Items)
		{
   		 
			$count=0
    			
			$creationDate= $listItem["Created"] 
			
			$toDay= Get-date 
			

			if($creationDate.ToShortDateString() -eq $toDay.ToShortDateString())
			{
				
				
				foreach($AnnouncementItem in $AnnouncementItems)
				  
				{


					if($AnnouncementItem.Title.contains($listItem.Name))
					{
						$count=1

						write-host $listItem.Name
						
						break
						
						
					}
					
				}

				if($count -eq 0)
				{

					$DocUrl=$listItem.url
				


					$newItem = $AnnouncementList.Items.Add()
				
				

					$DocumentUrl="<a href='"+$web.url+"/"+$DocUrl+"'  target='"+"_blank"+"'>" + " Click Here" +"</a>"
				
					$newItem["Title"] = $web.Title + " > " +$list.Title+ " > " + $listItem.Name






				$newItem["Body"] =   $DocumentUrl+ " to view the document "
				
				$newItem.Update()
				
				if($DocUrl.contains("UK"))
				{
					
					$UKItem = $UKAnnouncementList.Items.Add()
					$UKItem["Title"]=$web.Title + " > " +$list.Title+ " > " + $listItem.Name
					$UKItem["Body"] =   $DocumentUrl+ " to view the document "
					$UKItem.Update()
				}
				
				elseif($DocUrl.contains("IE"))
				{
					
					$IEItem = $IEAnnouncementList.Items.Add()
					$IEItem["Title"]=$web.Title + " > " +$list.Title+ " > " + $listItem.Name
					$IEItem["Body"] =   $DocumentUrl+ " to view the document "
					$IEItem.Update()
				}
				elseif($DocUrl.contains("Storeys"))
				{
					
					$STItem = $StoreysAnnouncementList.Items.Add()
					$STItem["Title"]=$web.Title + " > " +$list.Title+ " > " + $listItem.Name
					$STItem["Body"] =   $DocumentUrl+ " to view the document "
					$STItem.Update()
				}
				else
				{
					
					$UKItem = $UKAnnouncementList.Items.Add()
					$UKItem["Title"]=$web.Title + " > " +$list.Title+ " > " + $listItem.Name
					$UKItem["Body"] =   $DocumentUrl+ " to view the document "
					$UKItem.Update()
					
					$IEItem = $IEAnnouncementList.Items.Add()
					$IEItem["Title"]=$web.Title + " > " +$list.Title+ " > " + $listItem.Name
					$IEItem["Body"] =   $DocumentUrl+ " to view the document "
					$IEItem.Update()
					
					$STItem = $StoreysAnnouncementList.Items.Add()
					$STItem["Title"]=$web.Title + " > " +$list.Title+ " > " + $listItem.Name
					$STItem["Body"] =   $DocumentUrl+ " to view the document "
					$STItem.Update()
					
				
				
				}
				
				
				
				
				
				}

				


			}

		}

		
       

	}
	}
    }
}

$web.Dispose();
$site.Dispose();