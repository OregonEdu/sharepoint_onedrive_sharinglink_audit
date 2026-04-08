# SharePoint/OneDrive SharingLink Audit/Removal Script
Benjamin Barshaw <benjamin.barshaw@ode.oregon.gov> - IT Operations & Support Network Team Lead - Oregon Department of Education

Requirements: PnP PowerShell Module (https://pnp.github.io/powershell/index.html)
              ExchangeOnlineManagement Module (https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-module) -- NOTE: I needed to use "Connect-ExchangeOnline -DisableWAM"
              "Audit Logs" or "View-Only Audit Logs" role in Exchange Admin Center to be able to use Search-UnifiedAuditLog
              One of the following:
                - SharePoint (or Global) Administrator if you aren't a SiteCollectionAdmin for the OneDrive/SharePoint sites to add yourself
                - A registered PnP Entra application scoped with Sites.ReadWrite.All Graph permissions
                - Already be a SiteCollectionAdmin for the SharePoint/OneDrive sites you wish to audit/remove SharingLinks
 
This script will search the Unified Audit Log for any SharingLinks that have been accessed within as many days back as defined in the $daysToGoBack (default is 30) variable in the user defined variables section directly below and add them to an exclusion list. You can increase 
the days to go back, but by default the Unified Audit Log only goes back 180 days (https://learn.microsoft.com/en-us/purview/audit-log-retention-policies) which is also the maximum time you can search on an unlicensed user. This script will also find any SharingLinks
for SharePoint/OneDrive site(s) that have been created within the amount of days defined in the $exceptionDaysToGoBack (default is 7) variable and also add them to the exclusion list. The logic for this is that if a SharingLink has just been created it may not have been
accessed by the enduser yet giving them a grace period to access it. Any other SharingLinks found on the site(s) have the option to be deleted, cleaning up any SharingLinks that have been created and most likely forgotten. The option to delete is a toggleable flag
in the menu (off by default) -- if you do not set the flag it will only report on what would be deleted by the script and not actually deleted. The script outputs 2 files:

SharingLinks_To_Keep-\<date\>.csv   -- This is a CSV containing FullPath, Object, ObjectType, UserAccessing, DateAccessed, SharingLinkURL, and SharingLinkId headers of all the SharingLinks that will be preserved either because someone used them within 30 days, or because 
                                     they were created within the $exceptionDaysToGoBack window.
                                     
SharingLinks_To_Delete-\<date\>.csv -- This is a CSV containing FullPath, Object, ObjectType, UserAccessing, DateAccessed, SharingLinkURL, and SharingLinkId headers of all the SharingLinks that would be deleted should you run the script and toggle the flag to delete
                                     SharingLinks (menu option #11 -- off by default)
