$month=$args[0]
$year=$args[1]
[int]$from=$args[2]
[int]$to=$args[3]
#Write-Host $from.GetType()
$name="Dhanalkshmi_F&P_Fertilizers_Sale_Report_"+$month+"_"+$year
$excel = New-Object -ComObject excel.application 
$excel.visible = $True
$workbook = $excel.Workbooks.Add()
$uregwksht= $workbook.Worksheets.Item(1) 
$outputpath=$PSScriptRoot+"\Fertilizers\$name"
$uregwksht.Name = 'Form'
$row = 1 
$Column = 1
$uregwksht.Cells.Item($row,$column)= $name
$MergeCells = $uregwksht.Range('A1:F2')
$MergeCells.Select() 
$MergeCells.MergeCells = $true 
$uregwksht.Cells(1, 1).HorizontalAlignment = -4108
$uregwksht.Cells.Item(1,1).Font.Size = 16 
$uregwksht.Cells.Item(1,1).Font.Bold=$True 
#$uregwksht.Cells.Item(1,1).Font.Name = 'Times New Roman' 
#$uregwksht.Cells.Item(1,1).Font.ThemeFont = 1 
#$uregwksht.Cells.Item(1,1).Font.ThemeColor = 4 
#$uregwksht.Cells.Item(1,1).Font.ColorIndex = 55 
#$uregwksht.Cells.Item(1,1).Font.Color = 8210719
$uregwksht.Cells.Item(3,1)  = 'BILL NUMBER'
$uregwksht.Cells.Item(3,1).Font.Bold=$True 
$uregwksht.Cells.Item(3,2) = 'CREDIT' 
$uregwksht.Cells.Item(3,2).Font.Bold=$True 
$uregwksht.Cells.Item(3,3)  = 'CASH' 
$uregwksht.Cells.Item(3,3).Font.Bold=$True 
$uregwksht.Cells.Item(3,4) = '12%'
$uregwksht.Cells.Item(3,4).Font.Bold=$True 
$uregwksht.Cells.Item(3,5) = '18%'
$uregwksht.Cells.Item(3,5).Font.Bold=$True 
$uregwksht.Cells.Item(3,6)= 'RETURN BILLS'
$uregwksht.Cells.Item(3,6).Font.Bold=$True 
$usedRange = $uregwksht.UsedRange
$uregwksht.Cells.Item(1,1).Font.Size = 12
$usedRange.EntireColumn.ColumnWidth=12
$i=4

$j=$to - $from
$j+=$i
for($i=4;$i -le $j;$i++){

$uregwksht.Cells.Item($i,1)= $from
$from+=1
}
$workbook.SaveAs($outputpath) 
$excel.Quit()