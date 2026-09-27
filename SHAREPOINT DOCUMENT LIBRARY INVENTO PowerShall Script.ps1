# ============================================================
# SHAREPOINT DOCUMENT LIBRARY INVENTORY
# READ-ONLY - NO SHAREPOINT CONTENT IS MODIFIED
# ============================================================

# -----------------------------
# CONFIGURATION
# -----------------------------

$SiteUrl = "https://qtpower.sharepoint.com/sites/Master123-old"
$Library = "Shared Documents"
$ClientId = "a6b628b8-536a-4fa3-a471-97bad9b61e66"
$OutputFile = "C:\Users\ponna\OneDrive\Desktop\Error&Output\SharePoint_Source_Inventory2.csv"
$ErrorFile = "C:\Users\ponna\OneDrive\Desktop\Error&Output\SharePoint_Source_Inventory_Errors.csv"

# Number of items retrieved per SharePoint request
$PageSize = 1000


# -----------------------------
# PREPARE OUTPUT FOLDER
# -----------------------------

$OutputFolder = Split-Path $OutputFile -Parent

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}


# -----------------------------
# CONNECT TO SHAREPOINT
# -----------------------------

Write-Host ""
Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
Write-Host "Site: $SiteUrl"
Write-Host "Library: $Library"
Write-Host ""

Connect-PnPOnline `
    -Url $SiteUrl `
    -ClientId $ClientId `
    -Interactive


# -----------------------------
# VERIFY LIBRARY
# -----------------------------

$List = Get-PnPList -Identity $Library -ErrorAction Stop

Write-Host "Library found: $($List.Title)" -ForegroundColor Green
Write-Host "Starting inventory..." -ForegroundColor Cyan
Write-Host ""


# -----------------------------
# GET LIBRARY ITEMS
# READ ONLY
# -----------------------------

$StartTime = Get-Date

$Items = Get-PnPListItem `
    -List $Library `
    -PageSize $PageSize `
    -Fields `
        "ID",
        "UniqueId",
        "FileRef",
        "FileLeafRef",
        "FSObjType",
        "File_x0020_Size",
        "Created",
        "Modified",
        "Author",
        "Editor",
        "ContentTypeId",
        "ContentType",
        "FileDirRef" `
    -ErrorAction Stop


# -----------------------------
# PROCESS ITEMS
# -----------------------------

$Results = foreach ($Item in $Items) {

    $Values = $Item.FieldValues

    $ItemType = if ($Values["FSObjType"] -eq 1) {
        "Folder"
    }
    else {
        "File"
    }

    $FileSizeBytes = 0

    if ($ItemType -eq "File" -and $null -ne $Values["File_x0020_Size"]) {
        $FileSizeBytes = [int64]$Values["File_x0020_Size"]
    }

    $FileSizeKB = if ($FileSizeBytes -gt 0) {
        [math]::Round($FileSizeBytes / 1KB, 2)
    }
    else {
        0
    }

    $FileSizeMB = if ($FileSizeBytes -gt 0) {
        [math]::Round($FileSizeBytes / 1MB, 2)
    }
    else {
        0
    }

    $FileExtension = ""

    if ($ItemType -eq "File") {
        $FileExtension = [System.IO.Path]::GetExtension(
            [string]$Values["FileLeafRef"]
        )
    }

    [PSCustomObject]@{
        ID = $Values["ID"]

        UniqueId = $Values["UniqueId"]

        ItemType = $ItemType

        Name = $Values["FileLeafRef"]

        FullPath = $Values["FileRef"]

        ParentFolder = $Values["FileDirRef"]

        Extension = $FileExtension

        SizeBytes = $FileSizeBytes

        SizeKB = $FileSizeKB

        SizeMB = $FileSizeMB

        Created = $Values["Created"]

        Modified = $Values["Modified"]

        CreatedBy = if ($null -ne $Values["Author"]) {
            $Values["Author"].LookupValue
        }
        else {
            ""
        }

        ModifiedBy = if ($null -ne $Values["Editor"]) {
            $Values["Editor"].LookupValue
        }
        else {
            ""
        }

        ContentType = if ($null -ne $Values["ContentType"]) {
            $Values["ContentType"].Name
        }
        else {
            ""
        }

        ContentTypeId = $Values["ContentTypeId"]
    }
}


# -----------------------------
# EXPORT CSV
# -----------------------------

$Results |
    Export-Csv `
        -Path $OutputFile `
        -NoTypeInformation `
        -Encoding UTF8


# -----------------------------
# SUMMARY
# -----------------------------

$EndTime = Get-Date

$Elapsed = $EndTime - $StartTime

$TotalItems = $Results.Count

$TotalFiles = @(
    $Results |
        Where-Object { $_.ItemType -eq "File" }
).Count

$TotalFolders = @(
    $Results |
        Where-Object { $_.ItemType -eq "Folder" }
).Count

$TotalSizeBytes = (
    $Results |
        Where-Object { $_.ItemType -eq "File" } |
        Measure-Object -Property SizeBytes -Sum
).Sum

if ($null -eq $TotalSizeBytes) {
    $TotalSizeBytes = 0
}

$TotalSizeGB = [math]::Round(
    $TotalSizeBytes / 1GB,
    2
)


# -----------------------------
# DISPLAY RESULTS
# -----------------------------

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host " INVENTORY COMPLETED" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green

Write-Host "Total Objects : $TotalItems"
Write-Host "Total Files   : $TotalFiles"
Write-Host "Total Folders : $TotalFolders"
Write-Host "Total Size    : $TotalSizeGB GB"
Write-Host "Elapsed Time  : $($Elapsed.ToString())"

Write-Host ""

Write-Host "CSV created at:" -ForegroundColor Cyan
Write-Host $OutputFile -ForegroundColor Yellow

Write-Host ""

Write-Host "NO SharePoint content was modified." -ForegroundColor Green
Write-Host "NO files were downloaded." -ForegroundColor Green
Write-Host "NO folders were modified." -ForegroundColor Green
Write-Host "NO files were uploaded/deleted/moved." -ForegroundColor Green

Write-Host ""