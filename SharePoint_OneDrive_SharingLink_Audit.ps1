## SharePoint/OneDrive SharingLink Audit/Removal Script
## Benjamin Barshaw <benjamin.barshaw@ode.oregon.gov> - IT Operations & Support Network Team Lead - Oregon Department of Education
#
#  Requirements: PnP PowerShell Module (https://pnp.github.io/powershell/index.html)
#                ExchangeOnlineManagement Module (https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-module) -- NOTE: I needed to use "Connect-ExchangeOnline -DisableWAM"
#                "Audit Logs" or "View-Only Audit Logs" role in Exchange Admin Center to be able to use Search-UnifiedAuditLog
#                One of the following:
#                  - SharePoint (or Global) Administrator if you aren't a SiteCollectionAdmin for the OneDrive/SharePoint sites to add yourself
#                  - A registered PnP Entra application scoped with Sites.ReadWrite.All Graph permissions
#                  - Already be a SiteCollectionAdmin for the SharePoint/OneDrive sites you wish to audit/remove SharingLinks
# 
#  This script will search the Unified Audit Log for any SharingLinks that have been accessed within as many days back as defined in the $daysToGoBack (default is 30) variable in the user defined variables section directly below and add them to an exclusion list. You can increase 
#  the days to go back, but by default the Unified Audit Log only goes back 180 days (https://learn.microsoft.com/en-us/purview/audit-log-retention-policies) which is also the maximum time you can search on an unlicensed user. This script will also find any SharingLinks
#  for SharePoint/OneDrive site(s) that have been created within the amount of days defined in the $exceptionDaysToGoBack (default is 7) variable and also add them to the exclusion list. The logic for this is that if a SharingLink has just been created it may not have been
#  accessed by the enduser yet giving them a grace period to access it. Any other SharingLinks found on the site(s) have the option to be deleted, cleaning up any SharingLinks that have been created and most likely forgotten. The option to delete is a toggleable flag
#  in the menu (off by default) -- if you do not set the flag it will only report on what would be deleted by the script and not actually deleted. The script outputs 2 files:
#
#  SharingLinks_To_Keep-<date>.csv   -- This is a CSV containing FullPath, Object, ObjectType, UserAccessing, DateAccessed, SharingLinkURL, and SharingLinkId headers of all the SharingLinks that will be preserved either because someone used them within 30 days, or because 
#                                       they were created within the $exceptionDaysToGoBack window.
#  SharingLinks_To_Delete-<date>.csv -- This is a CSV containing FullPath, Object, ObjectType, UserAccessing, DateAccessed, SharingLinkURL, and SharingLinkId headers of all the SharingLinks that would be deleted should you run the script and toggle the flag to delete
#                                       SharingLinks (menu option #11 -- off by default) 
#  
### USER DEFINED VARIABLES SECTION ###
#
# PA account to add to user's OneDrive as SiteCollectionAdmin
$paAccount = "<CHANGE_TO_PA_ACCOUNT>"
# Change this to the admin portal of your SharePoint
$spoAdmin = "<CHANGE_TO_SHAREPOINT_ADMIN_PORTAL>"
# Change this to the ClientID of the PnP Entra application
$pnpAppClientId = "<CHANGE_TO_CLIENTID_OF_PNP_POWERSHELL_ENTRA_APP>"
# Change this to your TenantID
$tenantId = "<CHANGE_TO_TENANT_ID>"
# Days to go back in Unified Audit Log search
$daysToGoBack = 30
# Days to go back Unified Audit Log search for grace period
$exceptionDaysToGoBack = 7
# Operations for searching Unified Audit Log
$auditLogOperations = "AnonymousLinkUsed,CompanyLinkUsed,SecureLinkUsed"
# Operations for search Unified Audit Log for exceptions
$exceptionLogOperations = "AnonymousLinkCreated,CompanyLinkCreated,SecureLinkCreated"
# Libraries to scan for SharePoint sites
$sharePointLibraries = @('Documents', 'Site Pages')
### END USER DEFINED VARIABLES SECTION ###
### GLOBAL VARIABLES SECTION - DON'T TOUCH PLEASE ###
#
# Version of the script 
$version = "v1.0"
# Get today's date formatted for a filename
$getDate = Get-Date -Format MMddyyyy
# Toggle for SiteCollectionAdmin check
$siteCollectionAdminFlag = 0
# Toggle for deletion of SharingLinks not used in days defined in the variable $daysToGoBack above -- default is off
$sharingLinksRemovalFlag = 0

# Create our class to hold all the pertinent information of a SharingLink
class PnPSharingLink
{
    # Full URL of the file/folder the SharingLink was made for
    [string]$FullPath
    # The file/folder itself without the URL
    [string]$Object
    # Type of the object -- whether it's a "File" or "Folder"
    [string]$ObjectType
    # Who accessed the SharingLink for exclusions or comma-joined string of all people who have access to and will be affected by the SharingLink deletion
    [string]$UserAccessing
    # Date the SharingLink was accessed for exclusions
    [string]$DateAccessed
    # The URL of the actual SharingLink
    [string]$SharingLinkURL
    # The ID of the SharingLink
    [string]$SharingLinkId
}

# NOTE: All the SiteCollectionAdmin functions (including comments) were stolen from my SharePoint/OneDrive SharingLinks Enumerator script (https://github.com/OregonEdu/sharepoint_onedrive_sharinglinks_enumerator) which was made before the PnP module had the
# handy-dandy SharingLink cmdlets

# This functions checks to see if your PA is already a SiteCollectionAdmin and if so will not add it to the CSV we create to track the sites you added your PA to. The logic here being that there could be site collections that your PA
# SHOULD be SiteCollectionAdmin for and is already. By handling it this way, we won't remove your PA from these sites when/if you remove yourself later from the CSV.
function checkSiteCollectionAdmin($getSPOSites)
{    
    ForEach ($spoSite in $getSPOSites)
    {
        Connect-PnPOnline -Url $spoSite -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                
        If ((Get-PnPSiteCollectionAdmin -ErrorAction Ignore).Email -contains $paAccount) 
        { 
            Write-Host -ForegroundColor DarkYellow "$($paAccount) was already a SiteCollectionAdmin for $($spoSite)! Not exporting to the CSV."
        }
        Else
        {                                                                
            Write-Host -ForegroundColor Cyan "Adding $($spoSite) to SiteCollectionAdmin-$($getDate).csv"
            [PSCustomObject]@{Url = $spoSite} | Export-Csv -NoTypeInformation -Append -Path ".\SiteCollectionAdmin-$getDate.csv"
        }
    }
}

# Add our PA as SiteCollectionAdmin for the sites specified -- this is needed to do ANY kind of work on them
function addSiteCollectionAdmin($getSPOSites)
{
    try
    {
        Connect-PnPOnline -Url $spoAdmin -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
        Write-Host -ForegroundColor Green "Successfully connected to PnP Admin!"
    }
    catch
    {
        Write-Host -ForegroundColor Red "Could not connect to PnP Admin!"
        exit
    }   

    ForEach ($spoSite in $getSPOSites)
    {
        try
        {            
            Set-PnPTenantSite -Identity $spoSite -Owners $paAccount
            Write-Host -ForegroundColor Cyan "Successfully added $($paAccount) as SiteCollectionAdmin to $($spoSite)!"
        }
        catch
        {
            Write-Host -ForegroundColor Red "Could not add PA account to $($spoSite)!"
        }
    }
}

# Remove our PA as SiteCollectionAdmin for the sites specified
function removeSiteCollectionAdmin($getSPOSites)
{
    ForEach ($spoSite in $getSPOSites)
    {
        Connect-PnPOnline -Url $spoSite -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
        Remove-PnPSiteCollectionAdmin -Owners $paAccount
        Write-Host -ForegroundColor Cyan "Successfully removed $($paAccount) as SiteCollectionAdmin for $($spoSite)!"
    }
}

# Audit and remove (if flag set) the SharingLinks that have not been accessed for $daysToGoBack (default 30) or are outside of the $exceptionDaysToGoBack (default 7) grace period
function removePnPSharingLinks($pnpListItems, $exclusionListHashtable)
{
    # Create hashtable to house our file-based SharingLink deletions
    $removePnPFileSharingLinkHash = @{}
    # Create hashtable to house our folder-based SharingLink deletions
    $removePnPFolderSharingLinkHash = @{}

    # Cycle through the entire contents of the Site Library we're working on
    ForEach ($listItem in $pnpListItems)
    {
        # Determine type so we can use the correct PnP cmdlet
        If ($listItem.FileSystemObjectType -eq "File")
        {
            # Check to see if a SharingLink exists
            If (($getPnPFileSharingLink = Get-PnPFileSharingLink -Identity $listItem.FieldValues.FileRef -ErrorAction SilentlyContinue) -ne $null)
            {
                # Populate our PnPSharingLink object for the CSV output
                $pnpSharingLinkObject = [PnPSharingLink]::new()
                $pnpSharingLinkObject.FullPath = (Get-PnPContext).Url + $listItem.FieldValues.FileRef
                $pnpSharingLinkObject.Object = $listItem.FieldValues.FileLeafRef
                $pnpSharingLinkObject.ObjectType = $listItem.FileSystemObjectType
                $tempArray = @()
                # SiteUser is specific to the site itself - User had a broader scope in Entra -- we don't care which they are, we want to know all of them
                $tempArray += $getPnPFileSharingLink.GrantedToIdentitiesV2.SiteUser.Email
                $tempArray += $getPnPFileSharingLink.GrantedToIdentitiesV2.User.Email
                If ($tempArray -eq $null)
                {
                    $pnpSharingLinkObject.UserAccessing = "N/A"                    
                }
                Else
                {
                    # An identity can be both a SiteUser AND a User so we sort the object and remove duplicates
                    $pnpSharingLinkObject.UserAccessing = ($tempArray | Sort-Object | Get-Unique) -join ", "
                }
                # There is no DateAccessed without a Unified Audit Log entry of accessing (as opposed to SharingLinks we are keeping) so we set it as N/A
                $pnpSharingLinkObject.DateAccessed = "N/A"
                $pnpSharingLinkObject.SharingLinkURL = $getPnPFileharingLink.Link.WebUrl
                $pnpSharingLinkObject.SharingLinkId = $getPnPFileSharingLink.Id

                Write-Host -ForegroundColor Magenta "Found SharingLink! " -NoNewline
                Write-Host -ForegroundColor Cyan $($pnpSharingLinkObject.Object) -NoNewline
                Write-Host -ForegroundColor DarkCyan " which is a " -NoNewline
                Write-Host -ForegroundColor Cyan $($pnpSharingLinkObject.ObjectType)
                
                # Check our exclusion list hashtable to see if we are keeping the SharingLink
                If ($exclusionListHashtable.Keys -contains $getPnPFileSharingLink.Id)
                {                    
                    # Determine if the SharingLink is being propagated from inheritance of a folder being shared or if it is unique
                    If ($listItem.FieldValues.FileLeafRef -eq $exclusionListHashtable[$getPnPFileSharingLink.Id])
                    {
                        Write-Host -ForegroundColor DarkMagenta "$($pnpSharingLinkObject.Object) is on the master exclusion list! Skipping..."
                    }
                    Else
                    {
                        Write-Host -ForegroundColor Yellow "$($pnpSharingLinkObject.Object) is on exclusion list from $($exclusionListHashtable[$getPnPFileSharingLink.Id])! Skipping..."
                    }
                }
                Else
                {
                    # If determined we are deleting the SharingLink, add it to a hashtable we'll cycle through at the end and output it to our deletions CSV
                    Write-Host -ForegroundColor DarkYellow "$($pnpSharingLinkObject.Object) not on exclusion list! Adding to removal list..."
                    $removePnPFileSharingLinkHash[$getPnPFileSharingLink.Id] = $listItem.FieldValues.FileRef
                    $pnpSharingLinkObject | Export-Csv -NoTypeInformation -Append -Path "SharingLinks_To_Delete-$getDate.csv"
                }
            }
        }
        ElseIf ($listItem.FileSystemObjectType -eq "Folder")
        {
            # Same logic as for files so no need to comment
            If (($getPnPFolderSharingLink = Get-PnPFolderSharingLink -Folder $listItem.FieldValues.FileRef -ErrorAction SilentlyContinue) -ne $null)
            {
                $pnpSharingLinkObject = [PnPSharingLink]::new()
                $pnpSharingLinkObject.FullPath = (Get-PnPContext).Url + $listItem.FieldValues.FileRef
                $pnpSharingLinkObject.Object = $listItem.FieldValues.FileLeafRef
                $pnpSharingLinkObject.ObjectType = $listItem.FileSystemObjectType
                $tempArray = @()
                $tempArray += $getPnPFileSharingLink.GrantedToIdentitiesV2.SiteUser.Email
                $tempArray += $getPnPFileSharingLink.GrantedToIdentitiesV2.User.Email
                If ($tempArray -eq $null)
                {
                    $pnpSharingLinkObject.UserAccessing = "N/A"                    
                }
                Else
                {
                    $pnpSharingLinkObject.UserAccessing = ($tempArray | Sort-Object | Get-Unique) -join ", "
                }                
                $pnpSharingLinkObject.DateAccessed = "N/A"
                $pnpSharingLinkObject.SharingLinkURL = $getPnPFolderSharingLink.Link.WebUrl
                $pnpSharingLinkObject.SharingLinkId = $getPnPFolderSharingLink.Id

                Write-Host -ForegroundColor Magenta "Found SharingLink! " -NoNewline
                Write-Host -ForegroundColor Cyan $($pnpSharingLinkObject.Object) -NoNewline
                Write-Host -ForegroundColor DarkCyan " which is a " -NoNewline
                Write-Host -ForegroundColor Cyan $($pnpSharingLinkObject.ObjectType)

                If ($exclusionListHashtable.Keys -contains $getPnPFolderSharingLink.Id)
                {
                    If ($listItem.FieldValues.FileLeafRef -eq $exclusionListHashtable[$getPnPFolderSharingLink.Id])
                    {
                        Write-Host -ForegroundColor DarkMagenta "$($pnpSharingLinkObject.Object) is on the master exclusion list! Skipping..."
                    }
                    Else
                    {
                        Write-Host -ForegroundColor Yellow "$($pnpSharingLinkObject.Object) is on exclusion list from $($exclusionListHashtable[$getPnPFolderSharingLink.Id])! Skipping..."
                    }
                }
                Else
                {
                    Write-Host -ForegroundColor DarkYellow "$($pnpSharingLinkObject.Object) not on exclusion list! Adding to removal list..."
                    $removePnPFolderSharingLinkHash[$getPnPFolderSharingLink.Id] = $listItem.FieldValues.FileRef
                    $pnpSharingLinkObject | Export-Csv -NoTypeInformation -Append -Path "SharingLinks_To_Delete-$getDate.csv"
                }
            }
        }
    }

    # Check if we actually are going to remove SharingLinks or if we just wanted to audit what would be removed
    If ($sharingLinksRemovalFlag)
    {
        Write-Host -ForegroundColor Cyan "SharingLinks removal flag set!"

        # Cycle through our hashtables and do the actual SharingLink deletions
        Write-Host -ForegroundColor Yellow "Found $($removePnPFileSharingLinkHash.Count) file SharingLinks to remove! Removing..."
        ForEach ($key in $removePnPFileSharingLinkHash.Keys)
        {
            Write-Host -ForegroundColor Cyan "Removing SharingLink for File $($removePnPFileSharingLinkHash[$key]) ..."
            Remove-PnPFileSharingLink -FileUrl $removePnPFileSharingLinkHash[$key] -Identity $key
        }
        Write-Host -ForegroundColor Yellow "Found $($removePnPFolderSharingLinkHash.Count) folder SharingLinks to remove! Removing..."
        ForEach ($key in $removePnPFolderSharingLinkHash.Keys)
        {
            Write-Host -ForegroundColor Cyan "Removing SharingLink for Folder $($removePnPFolderSharingLinkHash[$key]) ..."
            Remove-PnPFolderSharingLink -Folder $removePnPFolderSharingLinkHash[$key] -Identity $key
        }
    }
    Else
    {
        Write-Host -ForegroundColor Cyan "SharingLinks removal flag not set!"
    }
}

# Audit SharingLinks to determine which ones we want to keep
function auditPnPSharingLinks($getSPOsites)
{
    # Scan through all sites specified from the menu
    ForEach($spoSite in $getSPOsites)
    {
        # Check to see if it's a true SharePoint site or a OneDrive -- OneDrive only has a "Documents" Document Library but for SharePoint we want to check if Site Pages have been shared as well
        If ($spoSite -like "*-my.sharepoint.com/personal/*")
        {
            $siteLibraries = "Documents"
        }
        Else
        {
            $siteLibraries = $sharePointLibraries
        }

        # Check if we're already connected to PnP -- no use authenticating if we already have!
        If ((Get-PnPConnection -ErrorAction SilentlyContinue).Url -eq $spoSite)
        {
            Write-Host -ForegroundColor DarkYellow "Already connected to $($spoSite)! Skipping connection..."
        }
        Else
        {
            try
            {
                # Connect to site via PnP
                Connect-PnPOnline -Url $spoSite -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                Write-Host -ForegroundColor Green "Successfully connected to $($spoSite)!"
            }
            catch
            {
                Write-Host -ForegroundColor Red "Could not connect to PnP site $($spoSite)!"
                #exit
            }   
        }

        # The Search-UnifiedAuditLog uses the SiteId of the SharePoint/OneDrive site instead of a URL
        Write-Host -ForegroundColor Cyan "Retrieving SiteId of $($spoSite)..."

        try
        {
            $getSiteId = (Get-PnPSite -Includes Id).Id | Select-Object -ExpandProperty Guid
        }
        catch
        {
            Write-Host -ForegroundColor Red "Could not retrieve SiteID for $($spoSite)!"
            #exit
        }

        # Check if we're already connected to ExchangeOnline
        If ((Get-ConnectionInformation) -ne $null)
        {
            Write-Host -ForegroundColor DarkYellow "Already connected to Exchange Online! Skipping connection..."
        }
        Else
        {
            try
            {
                # Connect to ExchangeOnline if we're not -- I had to use -DisableWAM to get it work with the latest version of the ExchangeOnlineManagement module (v3.9.2)
                Connect-ExchangeOnline -DisableWAM
                Write-Host -ForegroundColor Green "Successfully connected to ExchangeOnline!"
            }
            catch
            {
                Write-Host -ForegroundColor Red "Could not connect to ExchangeOnline!"
            }
        }

        Write-Host -ForegroundColor Magenta "Retrieving " -NoNewline
        Write-Host -ForegroundColor Yellow $($daysToGoBack) -NoNewline
        Write-Host -ForegroundColor Magenta " days worth of SharePointSharingOperation logs for " -NoNewline
        Write-Host -ForegroundColor Yellow $($spoSite) -NoNewline
        Write-Host -ForegroundColor Magenta " ..."        

        # Combine both the $exceptionDaysToGoBack logs for sites created within the timeframe along with the $daysToGoBack logs for sites accessed within the timeframe
        $allLogs = @()

        # Pull all *created* SharingLinks logs
        $getExceptionLogs = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-$exceptionDaysToGoBack) -EndDate (Get-Date) -SiteIds $getSiteId -RecordType "SharePointSharingOperation" -Operations $exceptionLogOperations -ResultSize 5000 -SessionCommand ReturnLargeSet) | Sort-Object -Property Identity -Unique
        If ($getExceptionLogs.AuditData)
        {
            # Convert logs from JSON to workable PowerShell objects
            $convertExceptionLogsFromJson = ($getExceptionLogs.AuditData | ConvertFrom-Json)
            Write-Host -ForegroundColor Cyan "SharingLinks created within the $($exceptionDaysToGoBack) day grace period found in Unified Audit Logs!"
            # Add all the objects to the array for searching
            $allLogs += $convertExceptionLogsFromJson
        }
        Else
        {
            # If no logs are found we set it to $null which when used in the ForEach() loop will just not do anything since there's nothing there
            $convertExceptionLogsFromJson = $null
            Write-Host -ForegroundColor DarkMagenta "No SharingLinks found created within the $($exceptionDaysToGoBack) day grace period!"
        }        

        # Pull all *accessed* SharingLink logs
        $getLogs = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-$daysToGoBack) -EndDate (Get-Date) -SiteIds $getSiteId -RecordType "SharePointSharingOperation" -Operations $auditLogOperations -ResultSize 5000 -SessionCommand ReturnLargeSet) | Sort-Object -Property Identity -Unique
        If ($getLogs.AuditData)
        {
            $convertLogsFromJson = ($getLogs.AuditData | ConvertFrom-Json)
            Write-Host -ForegroundColor DarkMagenta "SharingLinks usage found within $($daysToGoBack) days in Unified Audit Logs! Adding to exclusion list..."
            $allLogs += $convertLogsFromJson
        }
        Else
        {
            $convertLogsFromJson = $null
            Write-Host -ForegroundColor DarkMagenta "No SharingLinks usage found within $($daysToGoBack) days in Unified Audit Logs! Nothing to add to exclusion list!"
        }

        # Cycle through each library which for OneDrive will just be "Documents" but SharePoint will be "Documents" and "Site Pages"
        ForEach ($library in $siteLibraries)
        {            
            Write-Host -ForegroundColor Cyan "Retrieving all files from $($library) library of $($spoSite)..."
            # Get all files from the Document Library of the site we're connected to
            $getPnPListItems = Get-PnPListItem -List $library -PageSize 5000            
            # Build our hashtable to hold any SharingLink exclusions that we want to keep and not have removed
            $exclusionListHashtable = @{}
            
            # Cycle through the combined audit logs
            ForEach ($objectFound in $allLogs)
            {    
                # Old regexp method of getting absolute path left in for posterity
                #$stripDomain = $objectFound.ObjectId -replace '^https?:\/\/[^\/]+', ''
                # Use the System.Uri class to get the pieces of the URL we need -- and we need both
                $getUri = [System.Uri]$objectFound.ObjectId
                $getBaseUrl = $getUri.GetLeftPart([System.UriPartial]::Authority)
                $getAbsolutePath = $getUri.AbsolutePath

                # Match the format of the Unified Audit Log file/folder to the Get-PnPListItem format
                If (($pnpObject = $getPnPListItems | Where-Object { $_.FieldValues.FileRef -eq $getAbsolutePath }) -ne $null)
                {
                    # Build out our object for exporting to CSV
                    $pnpSharingLinkObject = [PnPSharingLink]::new()
                    $pnpSharingLinkObject.FullPath = $objectFound.ObjectId
                    $pnpSharingLinkObject.Object = $pnpObject.FieldValues.FileLeafRef
                    $pnpSharingLinkObject.ObjectType = $pnpObject.FileSystemObjectType
                    $pnpSharingLinkObject.UserAccessing = $objectFound.UserId
                    $pnpSharingLinkObject.DateAccessed = $objectFound.CreationTime    

                    Write-Host -ForegroundColor Magenta "Found Unified Audit Log entry! " -NoNewline
                    Write-Host -ForegroundColor Cyan $($pnpSharingLinkObject.Object) -NoNewline
                    Write-Host -ForegroundColor DarkCyan " which is a " -NoNewline
                    Write-Host -ForegroundColor Cyan $($pnpSharingLinkObject.ObjectType) -NoNewline
                    Write-Host -ForegroundColor DarkCyan " was accessed by " -NoNewline 
                    Write-Host -ForegroundColor Cyan $($pnpSharingLinkObject.UserAccessing) -NoNewline  
                    Write-Host -ForegroundColor DarkCyan " on " -NoNewline 
                    Write-Host -ForegroundColor Cyan $($pnpSharingLinkObject.DateAccessed)

                    # If there is an error, we can skip doing work on the object -- set a flag to let us know
                    $sharingLinkErrorFlag = 0
                    
                    # Get the SharingLink for folders or set flag if error
                    If ($pnpSharingLinkObject.ObjectType -eq "Folder")
                    {
                        try
                        {
                            $getSharingLinks = Get-PnPFolderSharingLink -Folder $getAbsolutePath
                        }
                        catch
                        {
                            Write-Host -ForegroundColor Red "Problem getting SharingLink Id for $($pnpSharingLinkObject.FullPath)!"
                            $sharingLinkErrorFlag = 1
                        }
                    }
                    # Get the SharingLink for files or set flag if error
                    ElseIf ($pnpSharingLinkObject.ObjectType -eq "File")
                    {
                        try
                        {
                            $getSharingLinks = Get-PnPFileSharingLink -Identity $getAbsolutePath
                        }
                        catch
                        {
                            Write-Host -ForegroundColor Red "Problem getting SharingLink Id for $($pnpSharingLinkObject.FullPath)!"
                            $sharingLinkErrorFlag = 1
                        }
                    }
                    # Catchall if it's something we don't care about
                    Else
                    {
                        Write-Host -ForegroundColor Red "$($pnpSharingLinkObject.ObjectType) is a type we don't know and/or care about! Ignoring..."
                    }

                    # If there's not an error do some work
                    If (! $sharingLinkErrorFlag)
                    {
                        # Setup a flag to tell us if we're going to export the entry or not -- we do this to prevent a lot of duplicates which are unnecessary for files/folders that have multiple SharingLinks but are already excluded
                        $skipExportFlag = 0

                        # Cycle through all the SharingLinks for the object
                        ForEach ($sharingLink in $getSharingLinks)
                        {
                            # Add the SharingLink URL and ID to our object for export
                            $pnpSharingLinkObject.SharingLinkURL = $sharingLink.Link.WebUrl
                            $pnpSharingLinkObject.SharingLinkId = $sharingLink.Id

                            # Since we only searched the time range of $exceptionDaysToGoBack we can safely assume that everything found to be Created is an exception
                            If ($objectFound.Operation -like "*Created")
                            {
                                Write-Host -ForegroundColor Yellow "Exception found! $($pnpSharingLinkObject.FullPath) had a SharingLink created within the $($exceptionDaysToGoBack) day grace period!"
                                # Add exception to our exclusion hashtable
                                $exclusionListHashtable[$sharingLink.Id] = $pnpSharingLinkObject.Object
                            }
                            Else
                            {                                
                                # Make sure an accessed SharingLink doesn't have a duplicate -- otherwise skip
                                If ($exclusionListHashtable.Keys -contains $sharingLink.Id)
                                {
                                    Write-Host -ForegroundColor Yellow "$($sharingLink.Id) already being excluded from $($exclusionListHashtable[$sharingLink.Id])! Skipping..."
                                    $skipExportFlag = 1
                                }                            
                                Else
                                {                                    
                                    # Try to lookup the user the SharingLink was shared with -- if they don't exist they were most likely deleted from Entra
                                    If (($getUser = Get-PnPEntraIDUser -Identity $($objectFound.UserId)) -eq $null)
                                    {
                                        Write-Host -ForegroundColor DarkRed "$($objectFound.UserId) not found in Entra! Not adding to exclusion list..."
                                    }
                                    Else
                                    {      
                                        # Determine which SharingLink was used by the person accessing by looking at the permissions granted for each SharingLink
                                        If (($sharingLink.GrantedToIdentitiesV2.User.Email -contains $getUser.Mail) -or ($sharingLink.GrantedToIdentitiesV2.SiteUser.Email -contains $getUser.Mail))
                                        {
                                            Write-Host -ForegroundColor Yellow "Found $($sharingLink.Id) for $($pnpSharingLinkObject.FullPath) which is most likely the SharingLink accessed by $($objectFound.UserId)! Adding to exclusion list..."                                        
                                            # After finding a match add it to the exclusions list
                                            $exclusionListHashtable[$sharingLink.Id] = $pnpSharingLinkObject.Object
                                        }
                                        Else
                                        {
                                            # If user not found, no exception is added for that specific SharingLink
                                            Write-Host -ForegroundColor DarkYellow "$($sharingLink.Id) was not the SharingLink used by $($objectFound.UserId)! Not adding to exclusion list..."
                                        }                                   
                                    }
                                }
                            }
                        }
                    }                    

                    # If the flag for skipping the export isn't set then export to our CSV
                    If (! $skipExportFlag)
                    {
                        $pnpSharingLinkObject | Export-Csv -NoTypeInformation -Append -Path ".\SharingLinks_To_Keep-$getDate.csv"
                    }
                }                    
            }

            # Begin process of auditing links that would be deleted 
            removePnPSharingLinks $getPnPListItems $exclusionListHashtable
        }
    }
}

# Display our menu 
function showMenu
{
    Write-Host -ForegroundColor Cyan "SharePoint/OneDrive SharingLinks Auditing/Deletion Script $($version) - Benjamin Barshaw <benjamin.barshaw@ode.oregon.gov>"    
    Write-Host -ForegroundColor Yellow "Please make your selection:"    
    Write-Host -ForegroundColor DarkYellow " 1) Load SharePoint sites for tenant"
    Write-Host -ForegroundColor DarkYellow " 2) Load OneDrive sites for tenant"
    Write-Host -ForegroundColor DarkYellow " 3) Load SharePoint & OneDrive sites for tenant"
    Write-Host -ForegroundColor DarkYellow " 4) Load SharePoint or OneDrive for a single site/user"
    Write-Host -ForegroundColor DarkYellow " 5) Load SharePoint or OneDrive sites from a CSV (NOTE: Must have Url header!)"
    Write-Host -ForegroundColor DarkYellow " 6) Toggle check to see if PA account is SiteCollectionAdmin for loaded sites (Will export to SiteCollectionAdmin-<date>.csv if not already admin)"
    Write-Host -ForegroundColor DarkYellow " 7) Add PA account as SiteCollectionAdmin from SiteCollectionAdmin CSV"
    Write-Host -ForegroundColor DarkYellow " 8) Add PA account as SiteCollectionAdmin for ALL loaded sites"
    Write-Host -ForegroundColor DarkYellow " 9) Remove PA account as SiteCollectionAdmin from SiteCollectionAdmin CSV"
    Write-Host -ForegroundColor DarkYellow "10) Remove PA account as SiteCollectionAdmin for ALL loaded sites"
    Write-Host -ForegroundColor DarkYellow "11) Toggle removal of SharingLinks that have not been used in $($daysToGoBack) days"
    Write-Host -ForegroundColor DarkYellow "12) Audit SharingLinks for all sites loaded " -NoNewline
    Write-Host -ForegroundColor Yellow "(WARNING: If SharingLinksRemovalFlag is true it will delete SharingLinks not used in $($daysToGoBack) days and cannot be undone!)"
    Write-Host -ForegroundColor DarkYellow "13) Show menu"    
    Write-Host -ForegroundColor DarkYellow "14) Exit"
    # Display our toggle flags
    Write-Host -ForegroundColor Cyan -NoNewline "SiteCollectionAdmin Flag: "
    Write-Host -ForegroundColor Magenta "$([bool]$siteCollectionAdminFlag)"
    Write-Host -ForegroundColor Cyan -NoNewline "SharingLinksRemovalFlag: "
    Write-Host -ForegroundColor Magenta "$([bool]$sharingLinksRemovalFlag)"
    # Display count of how many sites are loaded
    Write-Host -ForegroundColor Cyan -NoNewline "Loaded sites: "
    If (! $getSPOSites)
    {
        Write-Host -ForegroundColor Magenta "0"
    }
    Else
    {
        Write-Host -ForegroundColor Magenta "$($getSPOSites.Count)"
    }
}

# Display our menu
showMenu

# Start a do/while loop that does a switch() on the integer value menu input
do
{
    [int]$menuResponse = $(Read-Host "Choice (13 to re-display menu; 14 to exit)")

    switch($menuResponse)
    {
        # Grab all SharePoint sites -- we omit the -IncludeOneDriveSites:$true from Get-PnPTenantSite to only get true SharePoint sites
        1
        {
            try
            {
                Connect-PnPOnline -Url $spoAdmin -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                Write-Host -ForegroundColor Green "Successfully connected to PnP Admin!"
            }
            catch
            {
                Write-Host -ForegroundColor Red "Could not connect to PnP Admin!"
                exit
            }

            Write-Host -ForegroundColor Magenta "Collecting all SharePoint sites in tenant..."
            $getSPOSites = Get-PnPTenantSite | Select-Object -ExpandProperty Url
            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) SharePoint sites!"

            # Check if the toggle to check SiteCollectionAdmin is set
            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }
        # This time we include -IncludeOneDriveSites:$true but also -Filter on the return to ONLY get OneDrive sites
        2
        {            
            try
            {
                Connect-PnPOnline -Url $spoAdmin -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                Write-Host -ForegroundColor Green "Successfully connected to PnP Admin!"
            }
            catch
            {
                Write-Host -ForegroundColor Red "Could not connect to PnP Admin!"
                exit
            }

            Write-Host -ForegroundColor Magenta "Collecting all OneDrive sites in tenant..."
            $getSPOSites = Get-PnPTenantSite -IncludeOneDriveSites:$true -Filter "Url -like '-my.sharepoint.com/personal/'" | Select-Object -ExpandProperty Url
            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) OneDrive sites!"

            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }
        # We want both SharePoint/OneDrive so we let Get-PnPTenantSite rip!
        3
        {
            try
            {
                Connect-PnPOnline -Url $spoAdmin -Interactive -ClientId $pnpAppClientId -Tenant $tenantId -WarningAction SilentlyContinue
                Write-Host -ForegroundColor Green "Successfully connected to PnP Admin!"
            }
            catch
            {
                Write-Host -ForegroundColor Red "Could not connect to PnP Admin!"
                exit
            }

            Write-Host -ForegroundColor Magenta "Collecting all SharePoint & OneDrive sites in tenant..."
            $getSPOSites = Get-PnPTenantSite -IncludeOneDriveSites:$true | Select-Object -ExpandProperty Url
            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) SharePoint & OneDrive sites!"

            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }
        # Single site
        4
        {
            $getSPOSites = $(Read-Host -Prompt "URL").TrimEnd('/')
            
            Write-Host -ForegroundColor Cyan "Loading $($getSPOSites.Count) site!"

            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }
        # Read from a pre-populated CSV file -- file must have Url header
        5
        {
            $csvFile = $(Read-Host -Prompt "CSV File (Must have Url header)")
            $importCsv = Import-Csv $csvFile

            $getSPOSites = $importCsv.Url

            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) sites in CSV!"

            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Yellow "Flag for SiteCollectionAdmin check set! This will check all loaded sites to see if $($paAccount) is SiteCollectionAdmin. Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    checkSiteCollectionAdmin $getSPOSites
                }
            }
        }            
        # Toggle the check for SiteCollectionAdmin
        6
        {
            If ($siteCollectionAdminFlag)
            {
                Write-Host -ForegroundColor Magenta "Setting siteCollectionAdminFlag to False..."
                $siteCollectionAdminFlag = 0
            }
            Else
            {
                Write-Host -ForegroundColor Magenta "Setting siteCollectionAdminFlag to True..."
                $siteCollectionAdminFlag = 1
            }

            showMenu
        }
        # This will blindly add whatever account as the SiteCollectionAdmin from a CSV without checking -- be careful!
        7
        {
            $csvFile = $(Read-Host -Prompt "CSV File (Must have Url header)")
            $importCsv = Import-Csv $csvFile

            $getSPOSites = $importCsv.Url

            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) sites in CSV! Loading..."
            Write-Host -ForegroundColor Yellow "This will add $($paAccount) as SiteCollectionAdmin to all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                addSiteCollectionAdmin $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            }
        }
        # Same as 7 but from loaded sites
        8
        {
            Write-Host -ForegroundColor Yellow "This will add $($paAccount) as SiteCollectionAdmin to all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                addSiteCollectionAdmin $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            }
        }
        # Remove the SiteCollectionAdmin from a CSV
        9
        {
            $csvFile = $(Read-Host -Prompt "CSV File (Must have Url header)")
            $importCsv = Import-Csv $csvFile

            $getSPOSites = $importCsv.Url

            Write-Host -ForegroundColor Cyan "Found $($getSPOSites.Count) sites in CSV! Loading..."
            Write-Host -ForegroundColor Yellow "This will remove $($paAccount) as SiteCollectionAdmin for all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                removeSiteCollectionAdmin $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            }
        }
        # Same as 9 but for loaded sites
       10
        {
            Write-Host -ForegroundColor Yellow "This will remove $($paAccount) as SiteCollectionAdmin for all loaded sites. Are you sure?"
            $yesNo = $null
            while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
            {
                $yesNo = Read-Host -Prompt "[Y/N]"
            }
        
            If ($yesNo -eq "y")
            {
                removeSiteCollectionAdmin $getSPOSites
            }
            Else
            {
                Write-Host -ForegroundColor Red "Aborting!"
            }
        }
        # This toggle is to remove all SharingLinks for a site that have not been used in the $daysToGoBack variable in the user-defined variables
       11
        {
            If ($sharingLinksRemovalFlag)
            {
                Write-Host -ForegroundColor Magenta "Setting SharingLinksRemovalFlag to False..."
                $sharingLinksRemovalFlag = 0
            }
            Else
            {
                Write-Host -ForegroundColor Magenta "Setting SharingLinksRemovalFlag to True..."
                $sharingLinksRemovalFlag = 1
            }

            showMenu
        }
        # This does the audit of SharingLinks and subsequent removal if $sharingLinksRemovalFlag is true
       12
        {
            If ($sharingLinksRemovalFlag)
            {
                Write-Host -ForegroundColor Yellow "This will delete ALL SharingLinks for the sites loaded that haven't been used in $($daysToGoBack) days! Are you sure?"
                $yesNo = $null
                while (($yesNo -ne 'y') -and ($yesNo -ne 'n')) 
                {
                    $yesNo = Read-Host -Prompt "[Y/N]"
                }
        
                If ($yesNo -eq "y")
                {
                    auditPnPSharingLinks $getSPOSites                                        
                }
                Else
                {
                    Write-Host -ForegroundColor Red "Aborting!"
                }
            }
            Else
            {
                auditPnPSharingLinks $getSPOSites
            }            
       }            
        # Re-display the menu
       13
        {
            showMenu
        }
    }     
}
while ($menuResponse -ne 14)  # Quit on 14!
# FIN