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

# --- Global variable to hold the report data ---
$global:Report = @()

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
        Connect-PnPOnline -Url $WebUrl -Interactive -ClientId $clientId -PersistLogin -ErrorAction Stop
        
        # Get the web object with all properties we need
        $web = Get-PnPWeb -Includes RoleAssignments, Title, HasUniqueRoleAssignments
        
        if ($null -eq $web) {
            Write-Host "  Error: Could not retrieve web object" -ForegroundColor Red
            return
        }
        
        Write-Host "  Retrieved web object successfully" -ForegroundColor DarkGray
        
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
                            
                            $global:Report += [PSCustomObject]@{
                                WebUrl      = $WebUrl
                                WebTitle    = $web.Title
                                ObjectType  = $ObjectType
                                GroupName   = $displayName
                                GroupType   = $groupType
                                LoginName   = $member.LoginName
                                Permissions = $permissions
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
                                
                                $global:Report += [PSCustomObject]@{
                                    WebUrl      = $WebUrl
                                    WebTitle    = $web.Title
                                    ObjectType  = "List: $($list.Title)"
                                    GroupName   = $displayName
                                    GroupType   = $groupType
                                    LoginName   = $member.LoginName
                                    Permissions = $permissions
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
    "Enter your ClientID (Application ID) for authentication (required for PnP Interactive Login)"
$tenantAdminUrl = Read-Host -Prompt `
    "Enter your Tenant Admin URL (e.g., https://yourtenant-admin.sharepoint.com)"

# --- Connect to the Tenant Admin site using PnP Interactive Login ---
Write-Host "Connecting to Tenant Admin site..." -ForegroundColor Cyan
Connect-PnPOnline -Url $tenantAdminUrl -Interactive -ClientId $clientId -PersistLogin

# --- Recon Scan: Count all site collections, subsites, and lists ---
Write-Host "`nStarting Recon Scan to count sites, subsites, and lists..." `
    -ForegroundColor Magenta
$allSites = Get-PnPTenantSite -Detailed -IncludeOneDriveSites:$false

$totalSites = $allSites.Count
$totalSubsites = 0
$totalLists = 0

$siteIndex = 0
foreach ($site in $allSites) {
    $siteIndex++
    Write-Progress -Activity "Recon Scan" `
        -Status ("Processing site {0} of {1}: {2}" -f $siteIndex, $totalSites, $site.Url) `
        -PercentComplete (($siteIndex / $totalSites) * 100)
    try {
        Connect-PnPOnline -Url $site.Url -Interactive -ClientId $clientId -PersistLogin | Out-Null
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

# --- Summary Report ---
Write-Host "`nFull Scan completed." -ForegroundColor Magenta

if ($global:Report.Count -eq 0) {
    Write-Host "No group permissions found in the scanned sites." -ForegroundColor Yellow
} else {
    $uniqueGroupSummary = $global:Report | Group-Object -Property GroupName | `
        Select-Object Name, Count
    
    Write-Host "`nSummary of discovered group assignments:" -ForegroundColor Green
    $uniqueGroupSummary | ForEach-Object {
        # Get the group type from the first occurrence
        $groupType = ($global:Report | Where-Object { $_.GroupName -eq $_.Name } | Select-Object -First 1).GroupType
        Write-Host ("Group: {0} | Type: {1} | Occurrences: {2}" -f $_.Name, $groupType, $_.Count) `
            -ForegroundColor White
    }
    
    # Group type summary
    $groupTypeSummary = $global:Report | Group-Object -Property GroupType | `
        Select-Object Name, Count
    
    Write-Host "`nSummary by group type:" -ForegroundColor Green
    $groupTypeSummary | ForEach-Object {
        Write-Host ("Type: {0} | Count: {1}" -f $_.Name, $_.Count) `
            -ForegroundColor White
    }
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
