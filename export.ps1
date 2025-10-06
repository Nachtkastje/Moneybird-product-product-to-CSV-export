<#
.SYNOPSIS
    Exports all products from a Moneybird administration to a CSV file using the Moneybird API.

.DESCRIPTION
    This script fetches all products from the specified Moneybird administration (ID: 1) via the API,
    handling pagination to retrieve all pages, and exports only the Title and Price to a CSV file.
    Requires PowerShell 5.1 or later.

.PARAMETER OutputPath
    The path to the output CSV file (default: 'MoneybirdProducts.csv' in the current directory).

.EXAMPLE
    .\Export-MoneybirdProducts.ps1 -OutputPath 'C:\Exports\products.csv'
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "MoneybirdProducts.csv"
)

# Configuration
$AdministrationId = "123456_YOUR_ID"
$AccessToken = "YOUR_API_KEY"
$baseUrl = "https://moneybird.com/api/v2"

# Headers for API requests
$headers = @{
    "Authorization" = "Bearer $AccessToken"
    "Accept" = "application/json"
}

# Function to fetch a page of products
function Get-ProductsPage {
    param(
        [string]$Page = 1
    )
    $uri = "$baseUrl/$AdministrationId/products.json?per_page=100&page=$Page"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -ErrorAction Stop
        return $response
    } catch {
        Write-Error "Failed to fetch products from page $Page : $($_.Exception.Message)"
        return $null
    }
}

# Main logic: Fetch all pages and collect products
$allProducts = @()
$page = 1
do {
    Write-Host "Fetching page $page..."
    $pageData = Get-ProductsPage -Page $page
    if ($pageData -and $pageData.Count -gt 0) {
        $allProducts += $pageData
        $page++
    } else {
        break
    }
} while ($true)

if ($allProducts.Count -eq 0) {
    Write-Warning "No products found or failed to fetch data. Please verify your Administration ID and Access Token."
    return
}

# Select only Title and Price for CSV
$productsForExport = $allProducts | ForEach-Object {
    [PSCustomObject]@{
        Title = $_.title
        Price = $_.price
    }
}

# Export to CSV with ? delimiter
try {
    $productsForExport | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force -Delimiter "?"
    Write-Host "Successfully exported $($allProducts.Count) products to $OutputPath"
} catch {
    Write-Error "Failed to export CSV: $($_.Exception.Message)"
}
