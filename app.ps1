# param (
#     [Parameter(Mandatory=$true)][string]$excelFile = "appFieldDef.xlsm",
#     [Parameter(Mandatory=$true)][string]$macro = "FieldDefMain.Batch",
#     [Parameter(Mandatory=$true)][string]$formName = "FIELD_DEF",
#     [Parameter(Mandatory=$true)][boolean]$outTemplOn = $true,
#     [Parameter(Mandatory=$true)][boolean]$moveOn = $false
# )

$excelFile = "appFieldDef.xlsm"
$macro = "FieldDefMain.Batch"
$formName = "FIELD_DEF"
$outTemplOn = $true
$moveOn = $false
##
$curFolder = pwd 
$fullpath = Join-Path $curFolder.Path $excelFile
$excel = new-object -comobject excel.application
$excel.Visible = $false
$workbook = $excel.workbooks.open($fullpath)
$excel.Run($macro, $formName, $outTemplOn, $moveOn)
$workbook.close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel
