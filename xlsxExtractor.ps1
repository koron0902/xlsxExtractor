# receive xlsx file path
Param($xlsx)

#
if($xlsx -eq $null){
    Write-Host "ERROR : option -xlsx is not assigned"
    Write-Host "How :"
    Write-Host "    xlsxExtractor /path/to/file.xlsx"
    Write-Host "    xlsxExtractor -xlsx /path/to/file.xlsx"
    exit(-1)
}

# check file existence
if(-Not(Test-Path $xlsx)) {
    Write-Host "ERROR : File", $xlsx, "does NOT exist"
    exit(-1)
}

# convert to absolute path for excel object
if(-Not(Split-Path $xlsx -IsAbsolute)){
    $xlsx = Convert-Path $xlsx
}

# start excel server
$excel = New-Object -ComObject Excel.Application
$excel.Workbooks.Open($xlsx)
$excel.DisplayAlerts = $false

# split xlsx path to get save directory 
$parent = Split-Path -Parent $xlsx
$filename = [System.IO.Path]::GetFileNameWithoutExtension($xlsx) + "_"
$prefix = Join-Path $parent $filename


# convert all sheet to csv 
for($i = 0;$i -lt $excel.Worksheets.Count; $i++){
    $sheet = $excel.Worksheets.Item($($i+1))
    $sheet.SaveAs($($prefix + $sheet.Name + ".csv"), [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)
	[void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($sheet)
}

# quit excel server
$excel.Workbooks.Close()
$excel.Quit()


# release object to safe
[void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel)
[void][System.GC]::Collect()
