<#
close excel connection
#>
function Release-Ref ($ref) { 
([System.Runtime.InteropServices.Marshal]::ReleaseComObject( 
[System.__ComObject]$ref) -gt 0) 
[System.GC]::Collect() 
[System.GC]::WaitForPendingFinalizers() 
} 

<# Worle file names #>
$arrExcelValuesWPFilenames = @() 
<# Customer file names #>
$arrExcelValuesCTFilenames = @()
<# missed file names on working path #>
$arrExcelValmissFilenames = @()
 
$objExcel = new-object -comobject excel.application  
$objExcel.Visible = $True  
$objWorkbook = $objExcel.Workbooks.Open("C:\filenamechange\pingzuo1.xlsx") 
$objWorksheet = $objWorkbook.Worksheets.Item(1) 

$i = 1 
 
Do { 
    $arrExcelValuesWPFilenames += $objWorksheet.Cells.Item($i, 1).Value()
    $i++ 
} 
While ($objWorksheet.Cells.Item($i,1).Value() -ne $null)
<# change working path #>
Set-Location -Path "C:\filenamechange"
<#check any missed file, if there is any missed files, force the program terminate and ask user try again#>
$missfilecount = 0
foreach ($arrExcelValuesWPFilename in $arrExcelValuesWPFilenames)
    {
        if ((Get-ChildItem).name -match [System.Text.RegularExpressions.Regex]::Escape($arrExcelValuesWPFilename))
        {
            Write-Host "find" $arrExcelValuesWPFilename
        }
        else
        {
            Write-Host "can not find file start wtih " $arrExcelValuesWPFilename -ForegroundColor red -BackgroundColor White
            $arrExcelValmissFilenames = $arrExcelValmissFilenames + $arrExcelValuesWPFilename
            $missfilecount++
        }    
    }
If ($missfilecount -ne 0)
    {
        <# note user there are file missed #>
        Write-Host "There are $missfilecount files missed, please check your files, ten run this program again"
        <#Ask user want to change file names or not#>
        $ChangeNames = Read-Host -Prompt 'Do you still want to change file names (y/n)'
        If ($ChangeNames -eq 'n')
            {
                Write-Host "Please check the missed file names, run the script again, see you"
                <# terminate powershell and exit#>
                $objExcel.workbooks.Close()
                $a = $objExcel.Quit()
                $a = Release-Ref($objWorksheet) 
                $a = Release-Ref($objWorkbook) 
                $a = Release-Ref($objExcel)
                exit 1
            }
        else
            {
                Write-Host "We start change the file name for you, the result will come out soon ..."
            }
    }


$j = 1

Do {
    $arrExcelValuesCTFilenames += $objWorksheet.Cells.Item($j, 2).text
    $j++
}
While ($objWorksheet.Cells.Item($j,2).Value() -ne $null)

$i--
$j--

If (($i -eq $j) -or ($ChangeNames -ne 'n')){
        Write-Host "We rename files now..."
        for ($n=0;$n -lt $i; $n++){
                write-host $arrExcelValuesWPFilenames[$n] -NoNewline
                $sourcefilename = $arrExcelValuesWPFilenames[$n]
                Write-Host " --> " -NoNewline
                write-host $arrExcelValuesCTFilenames[$n]
                $desfilename = $arrExcelValuesCTFilenames[$n]
                $filterstring = "$sourcefilename"+'*'
                Get-ChildItem -Filter "$filterstring"|Rename-Item -NewName{$_.name -replace "$sourcefilename","$desfilename"}
        }
    }
Else {
        Write-Output "Please check your excel file"
    }

$objExcel.workbooks.Close()
$a = $objExcel.Quit()
$a = Release-Ref($objWorksheet) 
$a = Release-Ref($objWorkbook) 
$a = Release-Ref($objExcel)