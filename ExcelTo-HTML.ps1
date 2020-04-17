<#
.SYNOPSIS
    ExcelTo-HTML.ps1 - convert Excel to HTML

.DESCRIPTION
    Script for converting an Excel workbook to HTML format.

.NOTES
    File Name:  ExcelTo-HTML.ps1
    Author:     Marcus Libäck <marcus.liback@gmail.com>
    Requires:   PowerShell v4

.EXAMPLE
    ExcelTo-HTML.ps1 -InFile Foo.xlsx -ExportFile Foo.html -Refresh
#>

# Command line parameters
Param (
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Configuration file containing the connection information to the database (required)"
    )]
    [string] $InFile
    ,
    [Parameter(
        Mandatory=$false,
        HelpMessage = "Filename of the resulting HTML document (optional)"
    )]
    [string] $ExportFile
    ,
    [Parameter(
        Mandatory=$false,
        HelpMessage = "Refresh all elements in workbook before publishing to HTML file (optional)"
    )]
    [switch] $Refresh
)

# -----------------------------------------------------------------------------
# Support functions
# -----------------------------------------------------------------------------
function PathToAbsolute([string] $path)
{
    if ($path -eq "") {
        $path = Convert-Path .
    }
    if (-not $(Split-Path -IsAbsolute $path)) {
        $path = Convert-Path (Join-Path $(Convert-Path .) $path)
    }

    return $path
}

function CleanupExcelInstance
{
    #$workbook.Close()
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($workbook)
    [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel)
    [System.GC]::Collect()
}

# -----------------------------------------------------------------------------
# Set parameters & settings
# -----------------------------------------------------------------------------
# Set up path and filename for the exported PDF file
$_InPath    = Split-Path $InFile -Parent
$_InFile    = Split-Path $InFile -Leaf

# Set up path and filename for the exported PDF file
if ($ExportFile) {
    $_ExportPath    = Split-Path $ExportFile -Parent
    $_ExportFile    = Split-Path $ExportFile -Leaf
} else {
    $_ExportPath    = Convert-Path .
    $_ExportFile    = "$((Get-Item $MyInvocation.MyCommand.Definition).BaseName).html"
}

# Convert relative paths to absolute
$_InPath        = PathToAbsolute($_InPath)
$_ExportPath    = PathToAbsolute($_ExportPath)

# -----------------------------------------------------------------------------
# Main program
# -----------------------------------------------------------------------------
# Create new Excel COM-object
$excel = New-Object -ComObject Excel.Application
$excel.Visible          = $false
$excel.DisplayAlerts    = $false
$xlSourceType           = "Microsoft.Office.Interop.Excel.xlSourceType" -as [type]

# Open Excel file
try {
    $workbook    = $excel.workbooks.open("$_InPath\$_InFile")
    $worksheet    = $workbook.ActiveSheet
} catch {
    Write-Output "" "Error! Could open Excel file, exiting script!"
    Write-Output "" "Message:" "--------" "$($_.Exception.Message)" "$($_.Exception.ItemName)"
    CleanupExcelInstance
    exit 1
}

# Refresh all sheets before continuing (ensures all data is current)
if ($Refresh) {
    try {
        $workbook.RefreshAll()
        $excel.Application.CalculateUntilAsyncQueriesDone()
    } catch {
        Write-Output "" "Error! Could not refresh Excel file, exiting script!"
        Write-Output "" "Message:" "--------" "$($_.Exception.Message)" "$($_.Exception.ItemName)"
        CleanupExcelInstance
        exit 1
    }
}

# Export to HTML
try {
    $workbook.PublishObjects.Add($xlSourceType::xlSourceSheet, "$_ExportPath\$_ExportFile", $worksheet.Name, $false, $false, $worksheet.Name, $false).Publish($true).Delete
} catch {
    Write-Output "" "Error! Could not export HTML file, exiting script!"
    Write-Output "" "Message:" "--------" "$($_.Exception.Message)" "$($_.Exception.ItemName)"
    CleanupExcelInstance
    exit 1
}

# Save original Excel file again, quit excel and clean up
$workbook.SaveAs("$_InPath\$_InFile")
Write-Output "Excel to HTML:`t $_ExportPath\$_ExportFile successfully created"
CleanupExcelInstance
