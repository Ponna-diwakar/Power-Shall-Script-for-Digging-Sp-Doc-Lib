# Install the ImportExcel module if not already installed
# Install-Module -Name ImportExcel -Scope CurrentUser
# git init
# git status
# git add "All Sharepoint Sites Getting From Sp Admin Center.ps1"
# (here i you want to all folder to push use ""git add ."")
# git commit -m "This Script Loop All Sp Site From Admin Center And give you Excel Report "
# git remote add origin "https://github.com/Ponna-diwakar/Power-Shall-Script-for-Digging-Sp-Doc-Lib.git"
# git remote -v
# git branch -M main
# git push -u origin main


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

