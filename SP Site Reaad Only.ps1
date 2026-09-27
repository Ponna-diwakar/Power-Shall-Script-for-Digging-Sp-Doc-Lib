# Connect to SharePoint Admin Center
Connect-PnPOnline `
    -Url "https://qtpower-admin.sharepoint.com" `
    -ClientId "a6b628b8-536a-4fa3-a471-97bad9b61e66" `
    -Interactive

# Lock destination site as Read-Only
Set-PnPTenantSite `
    -Url "https://qtpower.sharepoint.com/sites/Master123-old" `
    -LockState Unlock

# Verify
Get-PnPTenantSite `
    -Url "https://qtpower.sharepoint.com/sites/Master123-old" |
    Select-Object Url, LockState