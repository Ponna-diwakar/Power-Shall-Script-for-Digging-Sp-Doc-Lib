# ===================================================================
# Script: Generate-SGUM.ps1
# Purpose: Export all users from SharePoint site into a clean SGUM file
# Author: Manish (QT Power Soft Solutions)
# Compatible with: PnP.PowerShell v3.x
# ===================================================================

Import-Module PnP.PowerShell

# ---- 🔧 CONFIGURATION ----
$siteUrl = "https://qtpower.sharepoint.com/sites/Master123-old"
$clientId = "a6b628b8-536a-4fa3-a471-97bad9b61e66"
$tenant = "qtpower.onmicrosoft.com"
$ExportPath = "C:\Users\ponna\OneDrive\Desktop\Error&Output\SGUM_Users.csv"   # Output file path

# ---- 🧩 CONNECT TO SHAREPOINT ----
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenant -Interactive

# ---- 📥 FETCH ALL USERS ----
Write-Host "Fetching all users from $siteUrl ..." -ForegroundColor Cyan

# In v3.x, Get-PnPUser doesn’t support -PageSize, so we simply call it directly
$users = Get-PnPUser | Where-Object { $_.PrincipalType -eq "User" }

if (-not $users -or $users.Count -eq 0) {
    Write-Host "⚠️ No users found or insufficient permissions." -ForegroundColor Red
    exit
}

Write-Host "✅ Found $($users.Count) users." -ForegroundColor Green

# ---- 🧹 CLEAN AND PREPARE SGUM DATA ----
$results = @()
foreach ($user in $users) {
    $cleanLogin = $user.LoginName
    $cleanLogin = $cleanLogin -replace "i:0#.f\|membership\|", ""
    $cleanLogin = $cleanLogin -replace "c:0o.c\|federateddirectoryclaimprovider\|", ""
    $cleanLogin = $cleanLogin -replace "c:0-.f\|rolemanager\|", ""
    $cleanLogin = $cleanLogin -replace "c:0!.s\|windows", ""
    $cleanLogin = $cleanLogin.Trim()

    if ($cleanLogin -ne "") {
        $results += [PSCustomObject]@{
            SourceLogin      = $cleanLogin
            DestinationLogin = ""    # Fill later for mapping
        }
    }
}

# ---- 💾 EXPORT SGUM FILE ----
$results | Sort-Object SourceLogin | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "`n🎯 Export Complete!" -ForegroundColor Green
Write-Host "✅ SGUM file created successfully: $ExportPath"
Write-Host "👉 Open it in Excel, fill in the 'DestinationLogin' column, save, and import into ShareGate."
