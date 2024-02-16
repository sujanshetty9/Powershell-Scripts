Write-Host "google.com" -BackgroundColor Red -ForegroundColor Green
# Import the ImportExcel module
Import-Module ImportExcel

# Path to your Excel file
$excelFilePath = "C:\path\to\your\file.xlsx"

# Load data from Excel
$data = Import-Excel -Path $excelFilePath

# Initialize a new array to store the transformed data
$transformedData = @()

# Iterate through each row of the data
foreach ($row in $data) {
    # Create a new object with the desired columns
    $newRow = [PSCustomObject]@{
        'User ID' = $row.'User Name'
        'User Name' = $row.'User Name'
        'Accessed VDI' = $row.'Accessed VDI or App'
        'App' = if ($row.'Accessed VDI or App' -like '*VDI*') { $null } else { $row.'Accessed VDI or App' }
    }
    # Add the new object to the transformed data array
    $transformedData += $newRow
}

# Export the transformed data to a new Excel file
$transformedData | Export-Excel -Path "C:\path\to\transformed_file.xlsx" -AutoSize -ClearSheet -Force
