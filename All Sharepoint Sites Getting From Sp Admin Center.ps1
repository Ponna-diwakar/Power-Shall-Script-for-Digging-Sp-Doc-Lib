# Install the ImportExcel module if not already installed
# Install-Module -Name ImportExcel -Scope CurrentUser


# Variables
$ClientID = "9944911f-691c-4c83-89ca-dc2021938042"
$SiteURL = "https://qtpower.sharepoint.com/sites/Master123"
$ExportPath = "C:\Users\ponna\OneDrive\Documents"

# Connect to SharePoint Online Admin Center (Interactive Login)
Connect-PnPOnline -Url $SiteURL -ClientId $ClientID -Interactive

# Get all sites
$Sites = Get-PnPList 

# Export to real Excel file
$Sites | Export-Excel -Path $ExportPath -AutoSize -BoldTopRow -FreezeTopRow

Write-Host "All site information exported to $ExportPath"

