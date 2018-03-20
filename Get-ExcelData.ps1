<# Written by @Tr4pSec #>

param (
    [Parameter(Mandatory=$True)]
    [string]$Names = ''
)

$Path = "$home\documents\somebook1.xlsx" #Enter EXCEL file path
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $False
$Workbook = $Excel.Workbooks.open($Path)
$Worksheet = $Workbook.Worksheets.item("Sheet1")
$Worksheet.activate()
$SearchString = "*$Names*"
$Range = $Worksheet.Range("E1").EntireColumn
$Search = $Range.find($SearchString)

Write-Verbose "Found name $($Search.text) in row $($search.row)" -Verbose

$Row = $Search.Row
$InvoiceCell = "B$Row"
$Invoice = $Worksheet.Range("$InvoiceCell")
$Invoice = $Invoice.Text
Write-Verbose "Invoice number found! ..." -Verbose

$DueCell = "H$Row"
$Due = $Worksheet.Range("$DueCell")
$Due = $Due.Text
Write-Verbose "Due date found ..." -Verbose

$AmountCell = "L$Row"
$Amount = $Worksheet.Range("$AmountCell")
$Amount = $Amount.Text -replace ' '

Write-Verbose "Amount found ..." -Verbose

$thething = "Good day,
Our records indicate that invoice# $Invoice with the amount of $Amount to the whatever $($Search.text) is overdue as of $Due.  
"
$thething | clip

$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Search) | Out-Null
Remove-Variable Excel
Remove-Variable Workbook
Remove-Variable Worksheet
Remove-Variable Range
Remove-Variable Search
[System.GC]::Collect()

Write-Verbose "Finished! Copied to clipboard" -Verbose