$month=$args[0]
$year=$args[1]
[int]$from=$args[2]
[int]$to=$args[3]
#Write-Host $from.GetType()
$name="Dhanalkshmi_F&P_Pesticides_Sale_Report_"+$month+"_"+$year
$fname=$name+".xlsx"
$excel = New-Object -ComObject excel.application 
$path=$PSScriptRoot+"\Pesticides\"+$fname
$wb = $excel.Workbooks.Open($path)
$ExcelWorkSheet = $wb.Sheets.Item("Form")
$i=4
$j=$to - $from
$j+=$i
[int]$credit_total=0
[int]$cash_total=0
[int]$return_total=0
for($i=4;$i -le $j;$i++){
$credit_total+=$ExcelWorkSheet.cells.Item($i, 2).value2
$cash_total+=$ExcelWorkSheet.cells.Item($i, 3).value2
$return_total+=$ExcelWorkSheet.cells.Item($i, 4).value2
#$ExcelWorkSheet.Cells.Item($i,1)= $from
$from+=1
}
$i+=2
$ExcelWorkSheet.Cells.Item($i,1)= 'Total'
$ExcelWorkSheet.Cells.Item($i,1).Font.Bold=$True
$ExcelWorkSheet.Cells.Item($i,2)= $credit_total
$ExcelWorkSheet.Cells.Item($i,3)= $cash_total
$ExcelWorkSheet.Cells.Item($i,4)= $return_total
$i+=2
$ExcelWorkSheet.Cells.Item($i,1)= 'Credit ='
$ExcelWorkSheet.Cells.Item($i,1).Font.Bold=$True
$ExcelWorkSheet.Cells.Item($i,2)= $credit_total
$i+=1
$ExcelWorkSheet.Cells.Item($i,1)= 'Cash ='
$ExcelWorkSheet.Cells.Item($i,1).Font.Bold=$True
$ExcelWorkSheet.Cells.Item($i,2)= $cash_total
$i+=1
$ExcelWorkSheet.Cells.Item($i,1)= 'Total Sale ='
$ExcelWorkSheet.Cells.Item($i,1).Font.Bold=$True
$ExcelWorkSheet.Cells.Item($i,2)= $cash_total+$credit_total
$i+=2
$ExcelWorkSheet.Cells.Item($i,1)= '18% ='
$ExcelWorkSheet.Cells.Item($i,1).Font.Bold=$True
$ExcelWorkSheet.Cells.Item($i,2)= $cash_total+$credit_total
$i+=1
$ExcelWorkSheet.Cells.Item($i,1)= 'Return ='
$ExcelWorkSheet.Cells.Item($i,1).Font.Bold=$True
$ExcelWorkSheet.Cells.Item($i,2)= $return_total
$wb.Save()
$wb.close($true)