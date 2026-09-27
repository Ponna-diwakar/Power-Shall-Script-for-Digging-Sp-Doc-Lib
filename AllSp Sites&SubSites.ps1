# =========================
# Configuration
# =========================
$adminUrl = "https://qtpower-admin.sharepoint.com"
$clientId = "a6b628b8-536a-4fa3-a471-97bad9b61e66"
$exportPath = "C:\Users\ponna\OneDrive\Desktop\SPO_AllSites_Report.csv"

# =========================
# Connect to Admin Center
# =========================
Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $clientId

# Get all site collections
$siteCollections = Get-PnPTenantSite

$results = @()

foreach($siteCollection in $siteCollections)
{
    Write-Host "Processing Site Collection: $($siteCollection.Url)" -ForegroundColor Cyan

    try
    {
        Connect-PnPOnline -Url $siteCollection.Url -Interactive -ClientId $clientId

        # Root Web
        $rootWeb = Get-PnPWeb

        $results += [PSCustomObject]@{
            SiteName     = $rootWeb.Title
            SiteURL      = $rootWeb.Url
            TemplateCode = "$($rootWeb.WebTemplate)#$($rootWeb.Configuration)"
            SiteLevel    = "Root Site"
        }

        # All subsites recursively
        $subsites = Get-PnPSubWeb -Recurse

        foreach($subsite in $subsites)
        {
            # Determine level
            $rootSegments = ($rootWeb.Url.TrimEnd('/') -split '/').Count
            $subSegments  = ($subsite.Url.TrimEnd('/') -split '/').Count

            $level = if(($subSegments - $rootSegments) -eq 1)
            {
                "Subsite"
            }
            else
            {
                "Nested Subsite"
            }

            $results += [PSCustomObject]@{
                SiteName     = $subsite.Title
                SiteURL      = $subsite.Url
                TemplateCode = "$($subsite.WebTemplate)#$($subsite.Configuration)"
                SiteLevel    = $level
            }
        }
    }
    catch
    {
        Write-Warning "Failed to process $($siteCollection.Url)"
    }
}

# Export Report
$results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

Write-Host "Report exported to $exportPath" -ForegroundColor Green