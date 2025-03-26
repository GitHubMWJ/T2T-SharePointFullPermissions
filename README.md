# 🔍 SharePoint Entra Group Permissions Reporter

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![PnP.PowerShell](https://img.shields.io/badge/PnP.PowerShell-Required-green.svg)](https://pnp.github.io/powershell/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A comprehensive PowerShell tool for reporting Microsoft 365 (Entra ID) groups and security groups assigned to SharePoint sites, subsites, and lists.

<div align="center">
  <img src="https://img.shields.io/badge/SharePoint-Online-2C7CD3" alt="SharePoint Online Compatible" />
  <img src="https://img.shields.io/badge/Microsoft%20365-Groups-7B83EB" alt="Microsoft 365 Groups" />
  <img src="https://img.shields.io/badge/Entra%20ID-Integration-0078D4" alt="Entra ID Integration" />
</div>

## 📋 Table of Contents

- [Purpose](#-purpose)
- [Features](#-features)
- [Prerequisites](#-prerequisites)
- [Installation](#-installation)
- [Usage](#-usage)
- [How It Works](#-how-it-works)
- [Output Explanation](#-output-explanation)
- [Excluded System Principals](#-excluded-system-principals)
- [Troubleshooting](#-troubleshooting)
- [Limitations](#-limitations)
- [License](#-license)

## 🎯 Purpose

During tenant migrations, understanding how groups are applied across your SharePoint environment is critical. This script helps migration teams and administrators to:

- **Discover** all Microsoft 365 Groups, Security Groups, and SharePoint Groups with site permissions
- **Identify** where these groups have access in your SharePoint environment
- **Document** permission levels assigned to each group
- **Generate** reports that can inform your migration strategy and security governance

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🔍 **Comprehensive Discovery** | Scans site collections, subsites, and lists with unique permissions |
| 🏷️ **Group Classification** | Correctly identifies different group types (M365 Groups, Security Groups, SharePoint Groups) |
| 🔐 **Permission Mapping** | Reports which permission levels are assigned to each group |
| 📑 **Multiple Site Input Methods** | Scan sites via manual URL entry, CSV import, or tenant-level scan |
| ⚡ **Recon Mode** | Provides a quick overview of the environment size before performing the full scan |
| 🧪 **Test Mode** | Allows testing against a single site before scanning the entire tenant |
| 📊 **Visual Progress** | Uses color-coded output and progress bars to track scanning progress |
| 📂 **CSV Export** | Exports detailed findings for documentation and further analysis |
| 🔄 **Nested Group Detection** | Identifies Entra ID groups nested inside SharePoint groups |
| 🛡️ **System Principal Filtering** | Excludes system principals and admin roles from reports for cleaner migration planning |

## 📋 Prerequisites

- PowerShell 5.1 or higher
- [PnP.PowerShell](https://pnp.github.io/powershell/) module installed
- An app registration in Entra ID with:
  - **SharePoint**: `Sites.FullControl.All`
- Your Entra ID tenant ID (GUID format)
- An account with read access to the SharePoint sites you want to scan

## 💻 Installation

1. Install the PnP PowerShell module if you haven't already:

```powershell
Install-Module -Name "PnP.PowerShell" -Force
```

2. Download the `SharePoint-EntraGroupsScanner-Full.ps1` script to your local machine.

3. Ensure you have your Entra ID App Registration ClientID and Tenant ID ready.

## 🚀 Usage

1. Run the script in PowerShell:

```powershell
.\SharePoint-EntraGroupsScanner-Full.ps1
```

2. When prompted, enter:
   - The PnP PowerShell Application ID (ClientID of your registered app in Entra ID)
   - Your Tenant ID (e.g., `12345678-1234-1234-1234-123456789012`)
   - Your tenant admin URL (e.g., `https://yourtenant-admin.sharepoint.com`)

3. Select how you want to provide site URLs:
   - Enter URLs manually (comma-separated)
   - Import from a CSV file (with a column named 'Url' or 'URL')
   - Try tenant-level scan via PnP (requires SharePoint Admin permissions)

4. The script will first perform a recon scan to count site collections, subsites, and lists.

5. After reviewing the recon scan results, choose whether to proceed with the full scan.

6. Optionally enable test mode to scan just a single site.

7. Review the results in the console, including:
   - A summary of discovered group assignments
   - A breakdown of group types found

8. Optionally export the detailed findings to a CSV file.

## ⚙️ How It Works

The script operates in several phases:

```mermaid
graph TD
    A[Connection with Client & Tenant IDs] --> B[Site Collection Input Methods]
    B -->|Manual URLs| C[Site List]
    B -->|CSV Import| C
    B -->|PnP Tenant Scan| C
    C --> D[Recon Scan]
    D --> E{Proceed with full scan?}
    E -->|Yes| F[Full Scan]
    E -->|No| K[Exit]
    F --> G[Process Sites]
    G --> H[Process Subsites]
    H --> I[Process Lists]
    G --> J[Identify Groups & Permissions]
    H --> J
    I --> J
    J --> L[Generate Report]
    L --> M[Export CSV]
```

1. **Connection**: Uses PnP PowerShell to connect to your tenant admin site with your provided ClientID and Tenant ID.

2. **Site Collection Input**: Provides multiple ways to specify which sites to scan:
   - Manual entry of comma-separated URLs
   - Import from a CSV file
   - PnP tenant-level scan (requires SharePoint Admin permissions)

3. **Recon Scan**: Quickly counts all site collections, subsites, and lists in your environment to give you an overview of the scope.

4. **Full Scan (if approved)**: 
   - Processes each site collection in the provided list
   - For each site, retrieves all role assignments and identifies group principals
   - Examines lists with unique permissions to find group assignments
   - Recursively processes all subsites using the same approach
   - Collects information about the groups, their types, and permission levels
   - Inspects SharePoint groups to find nested Entra ID groups

5. **Reporting**: Provides a color-coded summary of findings in the console and optionally exports detailed data to CSV.

## 📝 Output Explanation

The script generates a report with the following information:

| Field | Description |
|-------|-------------|
| **WebUrl** | The URL of the site or subsite where the group has permissions |
| **WebTitle** | The title of the site or subsite |
| **ObjectType** | Whether this is a Site, Subsite, or List (with list name) |
| **GroupName** | The display name of the group |
| **GroupType** | The type of group (M365/EntraIDGroup, SecurityGroup, SharePointGroup, etc.) |
| **LoginName** | The internal login name of the group, which helps identify the group type |
| **Permissions** | The permission levels assigned to the group (e.g., "Full Control; Design; Edit") |
| **IsSystemPrincipal** | Indicates if this is a system principal (excluded from summary display) |
| **SharePointGroupContainer** | If the group is nested within a SharePoint group, this shows the container group name |

## 🔒 Excluded System Principals

The script automatically excludes certain system principals and admin roles from the report to provide cleaner output that focuses on groups that actually need to be migrated between tenants. These excluded items include:

### SharePoint System Principals
- Everyone
- Everyone except external users
- NT AUTHORITY\authenticated users
- NT AUTHORITY\LOCAL SERVICE
- Authenticated Users
- SharePoint App

### Microsoft 365 Admin Roles
- Global Administrator
- SharePoint Administrator
- Exchange Administrator
- Teams Administrator
- Security Administrator
- Compliance Administrator
- User Administrator
- Billing Administrator
- Power Platform Administrator
- Dynamics 365 Administrator
- Application Administrator
- Global Reader

**Why are these excluded?** 
- System principals are automatically created in every SharePoint environment
- Admin roles represent claims-based identities that are auto-provisioned in the destination tenant
- These principals appear in SharePoint with special identifiers (e.g., `c:0t.c|tenant|[GUID]`)
- They don't need to be migrated as part of tenant-to-tenant migrations
- Excluding them provides cleaner, more actionable reports focused on actual groups that need migration planning

The script counts these excluded principals separately in the summary statistics.

## 🔧 Troubleshooting

### Common Issues:

<details>
<summary><b>Authentication Failed</b></summary>
<ul>
<li>Ensure your App Registration has the correct permissions</li>
<li>Check that the ClientID and Tenant ID are entered correctly</li>
<li>Make sure you have access to the tenant admin site</li>
</ul>
</details>

<details>
<summary><b>Script Runs Slowly</b></summary>
<ul>
<li>The script processes every site, subsite, and list with unique permissions</li>
<li>Large tenants may take significant time to process</li>
<li>Consider using test mode on a subset of sites first</li>
</ul>
</details>

<details>
<summary><b>Missing Groups</b></summary>
<ul>
<li>If groups aren't appearing in the report, check if they're actually assigned directly to sites</li>
<li>Some groups might be nested within SharePoint groups</li>
<li>The script will detect Entra ID groups inside SharePoint groups and include them in the report</li>
</ul>
</details>

<details>
<summary><b>"ClientObject is null" Errors</b></summary>
<ul>
<li>This can happen if there are connection issues</li>
<li>The script handles this by creating a fresh connection for each site</li>
<li>Check network stability and permissions</li>
</ul>
</details>

<details>
<summary><b>CSV Import Issues</b></summary>
<ul>
<li>Ensure your CSV file has a column named 'Url' or 'URL'</li>
<li>Check that all URLs in the CSV are valid SharePoint site URLs</li>
<li>Verify the file path to your CSV is correct</li>
</ul>
</details>

## ⚠️ Limitations

- The script doesn't report on individual user permissions, only groups
- Nested group memberships beyond the first level are not expanded (e.g., if a Security Group is a member of another Security Group)
- Performance may be affected in very large tenants with thousands of sites
- The script doesn't report item-level permissions, only site, subsite and list-level permissions
- Using delegated permissions with Graph API is not supported for comprehensive site collection discovery

## 📄 License

This script is provided "as is" without warranty of any kind, either expressed or implied.

---

**Note**: Always test scripts in a non-production environment before running them against your production tenant.
