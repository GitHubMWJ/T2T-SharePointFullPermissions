# ==============================================================================
# Script:      Report-EntraGroupPermissions.ps1
# Description: For tenant migration, report Entra (M365, Security, etc.) groups
#              applied to SharePoint site collections, subsites, and lists.
# ==============================================================================

# --- Import the PnP.PowerShell module only if not already loaded ---
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    Write-Error "PnP.PowerShell is not installed. Please install it (e.g., Install-Module PnP.PowerShell -Force)"
    exit
}
if (-not (Get-Module -Name "PnP.PowerShell")) {
    Import-Module PnP.PowerShell -DisableNameChecking -ErrorAction Stop
}

# --- Check for Microsoft Graph PowerShell modules (informational only) ---
$graphModulesRequired = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Sites")
$missingModules = @()

foreach ($module in $graphModulesRequired) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        $missingModules += $module
    }
}

$graphAvailable = $missingModules.Count -eq 0
if (-not $graphAvailable) {
    Write-Host "Note: Microsoft Graph PowerShell modules are not installed." -ForegroundColor Yellow
    Write-Host "      These modules may be useful for other SharePoint administration tasks." -ForegroundColor Yellow
}

# --- Global variable to hold the report data ---
$global:Report = @()

# --- Define SharePoint system principals to exclude from summary ---
$systemPrincipals = @(
    # SharePoint system accounts
    "Everyone", 
    "Everyone except external users",
    "NT AUTHORITY\authenticated users",
    "NT AUTHORITY\LOCAL SERVICE",
    "Authenticated Users",
    "SharePoint App",
    
    # Microsoft 365 admin roles that are automatically provisioned
    "Global Administrator",
    "SharePoint Administrator",
    "Exchange Administrator", 
    "Teams Administrator",
    "Security Administrator",
    "Compliance Administrator",
    "User Administrator",
    "Billing Administrator",
    "Power Platform Administrator",
    "Dynamics 365 Administrator",
    "Application Administrator",
    "Global Reader"
)

# --- Function: Test Graph API Permissions ---
function Test-GraphPermissionWithPnP {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantAdminUrl,
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )
    
    try {
        # Make sure we're connected to SharePoint admin site
        Connect-PnPOnline -Url $TenantAdminUrl -Interactive -ClientId $ClientId -Tenant $TenantId -PersistLogin -ErrorAction Stop
        
        # Try to get a token from the current PnP connection (no resource parameter needed)
        Write-Host "Testing Graph API access..." -ForegroundColor Cyan
        $accessToken = Get-PnPAccessToken -ErrorAction Stop
        
        # Create headers for a test request
        $headers = @{
            "Authorization" = "Bearer $accessToken"
            "Content-Type" = "application/json"
        }
        
        # Make a minimal test request - just get a single site
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?`$top=1" -Headers $headers -Method Get -ErrorAction Stop
        
        # If we get here, permissions are working
        if ($response.value -and $response.value.Count -gt 0) {
            Write-Host "✓ Graph API access confirmed: Successfully retrieved site data" -ForegroundColor Green
            return $true
        } else {
            Write-Host "? Graph API access seems okay but returned no sites" -ForegroundColor Yellow
            return $true  # Still return true as the API worked, just no data
        }
    }
    catch {
        # Check specific error types
        if ($_ -match "Forbidden" -or $_ -match "Access denied" -or $_ -match "Authorization_RequestDenied") {
            Write-Host "✗ Graph API access denied: The application '$ClientId' doesn't have 'Sites.Read.All' permission" -ForegroundColor Red
            Write-Host "  This permission must be added in Entra ID and admin consent granted" -ForegroundColor Yellow
            return $false
        }
        elseif ($_ -match "invalid_resource" -or $_ -match "AADSTS500011") {
            Write-Host "✗ Graph API access error: Resource not registered for this application" -ForegroundColor Red
            return $false
        }
        else {
            Write-Host "✗ Graph API access error: $_" -ForegroundColor Red
            return $false
        }
    }
}

# --- Function: Process-SharePointGroup ---
# Processes a SharePoint group to find Entra ID groups inside it
function Process-SharePointGroup {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WebUrl,
        [Parameter(Mandatory = $true)]
        [string]$WebTitle,
        [Parameter(Mandatory = $true)]
        [string]$ObjectType,
        [Parameter(Mandatory = $true)]
        [string]$SPGroupName,
        [Parameter(Mandatory = $true)]
        [string]$Permissions
    )
    
    try {
        Write-Host "    Examining SharePoint group members: $SPGroupName" -ForegroundColor DarkGray
        
        # Get members of the SharePoint group
        $groupMembers = Get-PnPGroupMember -Identity $SPGroupName -ErrorAction Stop
        
        if ($null -eq $groupMembers -or $groupMembers.Count -eq 0) {
            Write-Host "      No members found in SharePoint group" -ForegroundColor DarkGray
            return
        }
        
        Write-Host "      Found $($groupMembers.Count) members in SharePoint group" -ForegroundColor DarkGray
        
        foreach ($member in $groupMembers) {
            # Check if this member is an Entra ID/Security group
            if ($member.PrincipalType -eq "SecurityGroup") {
                $groupType = "Unknown"
                                
                # Check for Microsoft 365 Group vs Security Group
                if ($member.LoginName -match "c:0t\.c\|tenant\|") {
                    $groupType = "M365/EntraIDGroup"
                } else {
                    $groupType = "SecurityGroup"
                }
                
                # Skip system principals
                $isSystemPrincipal = $systemPrincipals -contains $member.Title
                if ($isSystemPrincipal) {
                    Write-Host "      Skipping system principal: $($member.Title)" -ForegroundColor DarkGray
                    continue
                }
                
                Write-Host "      Found Entra/Security group in SP group: $($member.Title)" -ForegroundColor White
                
                # Add to report with indicator it's in a SharePoint group
                $global:Report += [PSCustomObject]@{
                    WebUrl           = $WebUrl
                    WebTitle         = $WebTitle
                    ObjectType       = $ObjectType
                    GroupName        = $member.Title
                    GroupType        = $groupType
                    LoginName        = $member.LoginName
                    Permissions      = $Permissions
                    IsSystemPrincipal = $false
                    SharePointGroupContainer = $SPGroupName
                }
            }
        }
    }
    catch {
        Write-Host "    Error processing SharePoint group members: $_" -ForegroundColor Red
    }
}

# --- Function: Process-Web ---
# Recursively processes a SharePoint web (site or subsite) and its lists.
function Process-Web {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WebUrl,
        [Parameter(Mandatory = $true)]
        [string]$ObjectType
    )
    
    # Set a color based on the ObjectType:
    switch ($ObjectType) {
        "Site"    { $color = "Green" }
        "Subsite" { $color = "Cyan" }
        default   { $color = "White" }
    }
    
    Write-Host ("Processing {0}: {1}" -f $ObjectType, $WebUrl) -ForegroundColor $color

    try {
        # Connect to the site - this is crucial as we need a fresh connection for each site
        # Updated to include tenant ID parameter
        Connect-PnPOnline -Url $WebUrl -Interactive -ClientId $clientId -Tenant $tenantId -PersistLogin -ErrorAction Stop
        
        # Get the web object with all properties we need
        $web = Get-PnPWeb -Includes RoleAssignments, Title, HasUniqueRoleAssignments
        
        if ($null -eq $web) {
            Write-Host "  Error: Could not retrieve web object" -ForegroundColor Red
            return
        }
        
        Write-Host "  Retrieved web object successfully" -ForegroundColor DarkGray
        
        # FIX: Handle empty web titles by using URL as fallback
        $webTitle = if ([string]::IsNullOrEmpty($web.Title)) {
            # Use URL as fallback if title is empty (for system sites)
            "System Site: " + $WebUrl.TrimEnd('/').Split('/')[-1]
        } else {
            $web.Title
        }
        
        # Process role assignments directly
        if ($web.RoleAssignments -ne $null) {
            Write-Host "  Found $($web.RoleAssignments.Count) role assignments" -ForegroundColor DarkGray
            
            foreach ($roleAssignment in $web.RoleAssignments) {
                try {
                    # Load required properties of the role assignment
                    $member = Get-PnPProperty -ClientObject $roleAssignment -Property Member
                    
                    if ($null -ne $member) {
                        # Load member properties
                        $memberProps = Get-PnPProperty -ClientObject $member -Property Title, LoginName, PrincipalType, Id
                        
                        # Determine display name
                        $displayName = if (-not [string]::IsNullOrWhiteSpace($member.Title)) {
                            $member.Title
                        } elseif (-not [string]::IsNullOrWhiteSpace($member.LoginName)) {
                            $member.LoginName
                        } else {
                            "Unknown Principal (ID: $($member.Id))"
                        }
                        
                        Write-Host "    Found member: $displayName of type: $($member.PrincipalType)" -ForegroundColor DarkGray
                        
                        # Check if this is a group
                        $isGroup = $member.PrincipalType -eq "SecurityGroup" -or 
                                  $member.PrincipalType -eq "SharePointGroup" -or
                                  $member.PrincipalType -eq "UnifiedGroup"
                        
                        if ($isGroup) {
                            # Determine group type
                            $groupType = "Unknown"
                            
                            if ($member.PrincipalType -eq "SecurityGroup") {
                                # Check for Microsoft 365 Group vs Security Group
                                if ($member.LoginName -match "c:0t\.c\|tenant\|") {
                                    $groupType = "M365/EntraIDGroup"
                                } else {
                                    $groupType = "SecurityGroup"
                                }
                            } elseif ($member.PrincipalType -eq "SharePointGroup") {
                                $groupType = "SharePointGroup"
                            } elseif ($member.PrincipalType -eq "UnifiedGroup") {
                                $groupType = "Microsoft365Group"
                            }
                            
                            # Get permissions
                            $roleBindings = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings
                            $permissions = ($roleBindings | ForEach-Object { $_.Name }) -join "; "
                            
                            # Add a flag to identify SharePoint system principals
                            $isSystemPrincipal = $systemPrincipals -contains $displayName
                            
                            $global:Report += [PSCustomObject]@{
                                WebUrl                 = $WebUrl
                                WebTitle               = $webTitle   # Use the fixed webTitle here
                                ObjectType             = $ObjectType
                                GroupName              = $displayName
                                GroupType              = $groupType
                                LoginName              = $member.LoginName
                                Permissions            = $permissions
                                IsSystemPrincipal      = $isSystemPrincipal
                                SharePointGroupContainer = $null
                            }
                            
                            # If this is a SharePoint group, examine its members for Entra ID groups
                            if ($groupType -eq "SharePointGroup") {
                                Process-SharePointGroup -WebUrl $WebUrl -WebTitle $webTitle -ObjectType $ObjectType `
                                    -SPGroupName $displayName -Permissions $permissions
                            }
                        }
                    } else {
                        Write-Host "    Warning: Member object is null" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "    Error processing role assignment: $_" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "  No role assignments found or unable to retrieve them" -ForegroundColor Yellow
        }

        # Process lists with unique permissions
        $lists = Get-PnPList -Includes HasUniqueRoleAssignments, RoleAssignments
        foreach ($list in $lists) {
            if ($list.HasUniqueRoleAssignments) {
                Write-Host ("  Processing List: {0} (Unique Permissions)" -f $list.Title) -ForegroundColor Yellow
                
                foreach ($roleAssignment in $list.RoleAssignments) {
                    try {
                        $member = Get-PnPProperty -ClientObject $roleAssignment -Property Member
                        
                        if ($null -ne $member) {
                            # Load member properties
                            $memberProps = Get-PnPProperty -ClientObject $member -Property Title, LoginName, PrincipalType, Id
                            
                            # Determine display name
                            $displayName = if (-not [string]::IsNullOrWhiteSpace($member.Title)) {
                                $member.Title
                            } elseif (-not [string]::IsNullOrWhiteSpace($member.LoginName)) {
                                $member.LoginName
                            } else {
                                "Unknown Principal (ID: $($member.Id))"
                            }
                            
                            Write-Host "    Found member: $displayName of type: $($member.PrincipalType)" -ForegroundColor DarkGray
                            
                            # Check if this is a group
                            $isGroup = $member.PrincipalType -eq "SecurityGroup" -or 
                                      $member.PrincipalType -eq "SharePointGroup" -or
                                      $member.PrincipalType -eq "UnifiedGroup"
                            
                            if ($isGroup) {
                                # Determine group type
                                $groupType = "Unknown"
                                
                                if ($member.PrincipalType -eq "SecurityGroup") {
                                    # Check for Microsoft 365 Group vs Security Group
                                    if ($member.LoginName -match "c:0t\.c\|tenant\|") {
                                        $groupType = "M365/EntraIDGroup"
                                    } else {
                                        $groupType = "SecurityGroup"
                                    }
                                } elseif ($member.PrincipalType -eq "SharePointGroup") {
                                    $groupType = "SharePointGroup"
                                } elseif ($member.PrincipalType -eq "UnifiedGroup") {
                                    $groupType = "Microsoft365Group"
                                }
                                
                                # Get permissions
                                $roleBindings = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings
                                $permissions = ($roleBindings | ForEach-Object { $_.Name }) -join "; "
                                
                                # Add a flag to identify SharePoint system principals
                                $isSystemPrincipal = $systemPrincipals -contains $displayName
                                
                                $global:Report += [PSCustomObject]@{
                                    WebUrl                 = $WebUrl
                                    WebTitle               = $webTitle   # Use the fixed webTitle here
                                    ObjectType             = "List: $($list.Title)"
                                    GroupName              = $displayName
                                    GroupType              = $groupType
                                    LoginName              = $member.LoginName
                                    Permissions            = $permissions
                                    IsSystemPrincipal      = $isSystemPrincipal
                                    SharePointGroupContainer = $null
                                }
                                
                                # If this is a SharePoint group, examine its members for Entra ID groups
                                if ($groupType -eq "SharePointGroup") {
                                    Process-SharePointGroup -WebUrl $WebUrl -WebTitle $webTitle -ObjectType "List: $($list.Title)" `
                                        -SPGroupName $displayName -Permissions $permissions
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "    Error processing list role assignment: $_" -ForegroundColor Red
                    }
                }
            }
        }

        # Process subsites
        $subsites = Get-PnPSubWeb
        foreach ($subsite in $subsites) {
            Process-Web -WebUrl $subsite.Url -ObjectType "Subsite"
        }
    }
    catch {
        Write-Host "  Error processing web: $_" -ForegroundColor Red
    }
}

# ==============================================================================
# MAIN SCRIPT
# ==============================================================================

# --- Ask user for required variables and explain them ---
$clientId = Read-Host -Prompt `
    "Enter the PnP PowerShell Application ID (ClientID of your registered app in Entra ID)"
$tenantId = Read-Host -Prompt `
    "Enter your Tenant ID (e.g., 12345678-1234-1234-1234-123456789012)"
$tenantAdminUrl = Read-Host -Prompt `
    "Enter your Tenant Admin URL (e.g., https://yourtenant-admin.sharepoint.com)"

# --- Connect to the Tenant Admin site using PnP Interactive Login ---
Write-Host "Connecting to Tenant Admin site..." -ForegroundColor Cyan
Connect-PnPOnline -Url $tenantAdminUrl -Interactive -ClientId $clientId -Tenant $tenantId -PersistLogin

# --- Site Collection Input: Allow multiple input methods based on available permissions ---
Write-Host "`nGet Site Collections to scan..." -ForegroundColor Magenta

# Build the menu options
$menuOptions = @"
How would you like to provide site URLs? (Enter number)
1. Enter URLs manually (comma-separated)
2. Import from a CSV file
3. Try tenant-level scan via PnP (requires SharePoint Admin permissions)
"@

$inputMethod = Read-Host -Prompt $menuOptions

$allSites = @()

switch ($inputMethod) {
    "1" {
        $siteUrls = Read-Host -Prompt "Enter comma-separated site collection URLs (e.g., https://tenant.sharepoint.com/sites/site1,https://tenant.sharepoint.com/sites/site2)"
        $urlList = $siteUrls -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        
        foreach ($url in $urlList) {
            $allSites += [PSCustomObject]@{
                Url = $url
            }
        }
    }
    "2" {
        $csvPath = Read-Host -Prompt "Enter path to CSV file containing site URLs (should have a column named 'Url' or 'URL')"
        if (Test-Path $csvPath) {
            $importedSites = Import-Csv -Path $csvPath
            
            # Check for the URL column name (either Url or URL)
            $urlColumnName = if ($importedSites[0].PSObject.Properties.Name -contains "Url") { "Url" } 
                            elseif ($importedSites[0].PSObject.Properties.Name -contains "URL") { "URL" }
                            else { $null }
            
            if ($urlColumnName) {
                foreach ($site in $importedSites) {
                    $allSites += [PSCustomObject]@{
                        Url = $site.$urlColumnName
                    }
                }
            } else {
                Write-Host "Error: CSV doesn't have a 'Url' or 'URL' column. Please check the format." -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "Error: CSV file not found at specified path." -ForegroundColor Red
            exit
        }
    }
    "3" {
        Write-Host "Attempting tenant-level scan via PnP. This may fail if you don't have sufficient permissions..." -ForegroundColor Yellow
        try {
            $allSites = Get-PnPTenantSite -Detailed -IncludeOneDriveSites:$false -ErrorAction Stop
        }
        catch {
            Write-Host "Error in tenant scan: $_" -ForegroundColor Red
            Write-Host "Falling back to manual input..." -ForegroundColor Yellow
            
            $siteUrls = Read-Host -Prompt "Enter comma-separated site collection URLs (e.g., https://tenant.sharepoint.com/sites/site1,https://tenant.sharepoint.com/sites/site2)"
            $urlList = $siteUrls -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
            
            foreach ($url in $urlList) {
                $allSites += [PSCustomObject]@{
                    Url = $url
                }
            }
        }
    }
    "4" {
        Write-Host "Invalid choice. Defaulting to manual input." -ForegroundColor Yellow
        $siteUrls = Read-Host -Prompt "Enter comma-separated site collection URLs (e.g., https://tenant.sharepoint.com/sites/site1,https://tenant.sharepoint.com/sites/site2)"
        $urlList = $siteUrls -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        
        foreach ($url in $urlList) {
            $allSites += [PSCustomObject]@{
                Url = $url
            }
        }
    }
    default {
        Write-Host "Invalid choice. Defaulting to manual input." -ForegroundColor Yellow
        $siteUrls = Read-Host -Prompt "Enter comma-separated site collection URLs (e.g., https://tenant.sharepoint.com/sites/site1,https://tenant.sharepoint.com/sites/site2)"
        $urlList = $siteUrls -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        
        foreach ($url in $urlList) {
            $allSites += [PSCustomObject]@{
                Url = $url
            }
        }
    }
}

$totalSites = $allSites.Count
if ($totalSites -eq 0) {
    Write-Host "No sites to scan. Exiting..." -ForegroundColor Red
    exit
}

# --- Recon Scan: Count subsites and lists for the provided site collections ---
Write-Host "`nStarting Recon Scan to count subsites and lists for $totalSites site collections..." `
    -ForegroundColor Magenta

$totalSubsites = 0
$totalLists = 0

$siteIndex = 0
foreach ($site in $allSites) {
    $siteIndex++
    Write-Progress -Activity "Recon Scan" `
        -Status ("Processing site {0} of {1}: {2}" -f $siteIndex, $totalSites, $site.Url) `
        -PercentComplete (($siteIndex / $totalSites) * 100)
    try {
        Connect-PnPOnline -Url $site.Url -Interactive -ClientId $clientId -Tenant $tenantId -PersistLogin -ErrorAction Stop | Out-Null
        $web = Get-PnPWeb
        $subWebs = Get-PnPSubWeb -Recurse
        $totalSubsites += $subWebs.Count
        $lists = Get-PnPList
        $totalLists += $lists.Count
    }
    catch {
        Write-Host ("Error in recon scan for site: {0} - {1}" -f $site.Url, $_) -ForegroundColor Red
    }
}
Write-Host "`nRecon Scan Complete:" -ForegroundColor Magenta
Write-Host ("  Total Site Collections: {0}" -f $totalSites)
Write-Host ("  Total Subsites:         {0}" -f $totalSubsites)
Write-Host ("  Total Lists:            {0}" -f $totalLists)
Write-Host ""

# --- Ask whether to proceed to the full scan ---
$proceed = Read-Host -Prompt `
    "Proceed with full scan (this will retrieve permission details for groups)? (Y/N)"
if ($proceed -notmatch "^[Yy]") {
    Write-Host "Full scan cancelled by user. Exiting..." -ForegroundColor Red
    Disconnect-PnPOnline -ClearPersistedLogin
    exit
}

# --- Option to scan just one site for testing ---
$testMode = Read-Host -Prompt "Run in test mode on a single site? (Y/N)"
if ($testMode -match "^[Yy]") {
    $testSiteUrl = Read-Host -Prompt "Enter the URL of a single site to scan"
    $allSites = $allSites | Where-Object { $_.Url -eq $testSiteUrl }
    
    if ($allSites.Count -eq 0) {
        Write-Host "Site not found. Exiting..." -ForegroundColor Red
        Disconnect-PnPOnline -ClearPersistedLogin
        exit
    }
    
    Write-Host "Test mode enabled. Will only scan: $testSiteUrl" -ForegroundColor Yellow
    $totalSites = 1
}

# --- Full Scan: Process site collections to discover role assignments ---
Write-Host "`nStarting Full Scan to discover Entra/SharePoint group permissions..." `
    -ForegroundColor Magenta
$siteCounter = 0
foreach ($site in $allSites) {
    $siteCounter++
    Write-Progress -Activity "Full Scan" `
        -Status ("Processing Site Collection {0} of {1}: {2}" -f $siteCounter, $totalSites, $site.Url) `
        -PercentComplete (($siteCounter / $totalSites) * 100)
    
    # Process each site directly by URL
    Process-Web -WebUrl $site.Url -ObjectType "Site"
}

# --- Summary Report with Group-Centric Format ---
Write-Host "`nFull Scan completed." -ForegroundColor Magenta

if ($global:Report.Count -eq 0) {
    Write-Host "No group permissions found in the scanned sites." -ForegroundColor Yellow
} else {
    # Filter out SharePoint groups and system principals, but keep Entra ID groups found inside SharePoint groups
    $filteredReport = $global:Report | Where-Object { 
        ($_.GroupType -ne "SharePointGroup" -and -not $_.IsSystemPrincipal) -or $_.SharePointGroupContainer -ne $null
    }
    
    # Group by GroupName and GroupType instead of by site
    $groupSummary = $filteredReport | Group-Object -Property GroupName, GroupType

    # Create clean header with border
    $headerWidth = 80
    $border = "-" * $headerWidth
    
    Write-Host "`nENTRA GROUP ASSIGNMENTS SUMMARY" -NoNewline -ForegroundColor White
    Write-Host "".PadRight(26) -NoNewline
    Write-Host "[Filtered: SP Groups excluded]" -ForegroundColor Yellow
    Write-Host $border -ForegroundColor DarkGray

    if ($groupSummary.Count -eq 0) {
        Write-Host "`nNo Entra ID or Security groups found - only SharePoint groups or system principals were detected." -ForegroundColor Cyan
    } else {
        # Add legend at the top
        Write-Host "LEGEND: [M]=M365/EntraID Group, [S]=Security Group, [SP]=SharePoint Group" -ForegroundColor White
        Write-Host ""
        
        foreach ($groupEntry in $groupSummary) {
            # Extract GroupName and GroupType from the Name property
            $groupParts = $groupEntry.Name -split ', '
            $groupName = $groupParts[0].Trim()
            $groupType = $groupParts[1].Trim()
            
            # Set type code based on group type
            $typeCode = switch ($groupType) {
                "M365/EntraIDGroup" { "[M]" }
                "SecurityGroup" { "[S]" }
                "SharePointGroup" { "[SP]" }
                default { "[?]" }
            }
            
            # Display group information as header
            Write-Host "GROUP: '$groupName' $typeCode - $($groupEntry.Group.Count) total instances" -ForegroundColor Yellow
            Write-Host $border -ForegroundColor DarkGray
            
            # Display LOCATIONS header
            Write-Host "LOCATIONS:" -ForegroundColor White
            
            # Display all locations where this group is found
            foreach ($location in $groupEntry.Group) {
                $siteInfo = $location.WebTitle
                $locationInfo = $location.ObjectType
                
                # Format location info with site context
                $locationDisplay = "• $siteInfo | $locationInfo"
                
                # If found inside a SharePoint group, add that information
                if ($location.SharePointGroupContainer -ne $null) {
                    $locationDisplay += " (via SP Group: $($location.SharePointGroupContainer))"
                }
                
                Write-Host $locationDisplay -ForegroundColor White
            }
            
            # Add border after each group's locations
            Write-Host $border -ForegroundColor DarkGray
            Write-Host ""
        }
    }

    # Group type summary statistics
    $filteredForSummary = $global:Report | Where-Object { -not $_.IsSystemPrincipal -and $_.GroupType -ne "SharePointGroup" }
    $groupTypeSummary = $filteredForSummary | Group-Object -Property GroupType | Select-Object Name, Count
    
    Write-Host "SUMMARY STATISTICS:" -ForegroundColor White
    $groupTypeSummary | ForEach-Object {
        # Set type code based on group type
        $typeCode = switch ($_.Name) {
            "M365/EntraIDGroup" { "[M]" }
            "SecurityGroup" { "[S]" }
            "SharePointGroup" { "[SP]" }
            default { "[?]" }
        }
        Write-Host ("• {0} {1}: {2} instances" -f $typeCode, $_.Name, $_.Count) -ForegroundColor White
    }
    
    # Count of groups found inside SharePoint groups
    $nestedGroupCount = ($global:Report | Where-Object { $_.SharePointGroupContainer -ne $null }).Count
    Write-Host "• Entra ID/Security groups inside SharePoint groups: $nestedGroupCount" -ForegroundColor White
    
    # Report on excluded system principals
    $systemPrincipalCount = ($global:Report | Where-Object { $_.IsSystemPrincipal }).Count
    Write-Host "• Excluded system principals: $systemPrincipalCount" -ForegroundColor White
    
    # Note about permissions in CSV
    Write-Host "`nNOTE: Permissions details are included in CSV export but omitted from display" -ForegroundColor Yellow
}

# --- Optionally save the detailed report to CSV ---
$saveCSV = Read-Host -Prompt "Save detailed report to CSV? (Y/N)"
if ($saveCSV -match "^[Yy]") {
    $csvPath = Read-Host -Prompt `
        "Enter full path to save CSV file (e.g., C:\temp\EntraGroupsReport.csv)"
    try {
        $global:Report | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host ("Report successfully saved to: {0}" -f $csvPath) `
            -ForegroundColor Green
    }
    catch {
        Write-Host "Error saving CSV file: $_" -ForegroundColor Red
    }
}

# --- Clear persisted login before finishing ---
Disconnect-PnPOnline -ClearPersistedLogin
Write-Host "`nScript completed. Persisted sign-in cleared." -ForegroundColor Cyan