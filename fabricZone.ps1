$file = "c:\temp\fabric.xlsx"
$sheetNae = "fabricb"
$fab = "B"
##Create an  instance of Excel.Application and Open Excel file 

$objExcel = New-Object -ComObject Excel.Application 
$workbook = $objExcel.Workbooks.Open($file) 
$sheet = $workbook.Worksheets.Item($sheetName) 
$objExcel.Visible=$false 
##Count max row 

$rowMax = ($sheet.UsedRange.Rows).count 
##Declare the starting positions 

$row,$colCluster = 1,1 
$row,$colHost = 1,2 
$row,$colArray = 1,3
$row,$colPort = 1,4
$row,$colFiber = 1,5
 
##loop to get values and store it 

for ($i=1; $i -le $rowMax-1; $i++) 
{ 
$vmhost = $sheet.Cells.Item($row+$i,$colHost).text 
$array = $sheet.Cells.Item($row+$i,$colArray).text 
$port = $sheet.Cells.Item($row+$i,$colPort).text 
if ($fab -eq "A") 
    {
    $vsan = "1000"
    $hba = "vhba0"
    $zoneset = "Fabric_A_Corp"
    }
else
    {
    $vsan = "2000"
    $hba = "vhba1"
    $zoneset = "Fabric_B_Corp"
    }

Write-Host ("! New Line :" +$vmhost, $fab)
Write-Host ("    zone name HDS_"+ $array +"_" + $port +"_" + $vmhost + "_" + $hba + " vsan " + $vsan)
Write-Host ("        member device-alias HDS_" + $array + "_" + $port)
Write-Host ("        member device-alias " + $vmhost + "_" + $hba)
Write-Host ("   exit")
}
$row,$colCluster = 1,1 
$row,$colHost = 1,2 
$row,$colArray = 1,3
$row,$colPort = 1,4
$row,$colFiber = 1,5
 
##loop to get values and store it 

Write-Host ("     zoneset  name " + $zoneset + " vsan " + $vsan)
for ($i=1; $i -le $rowMax-1; $i++) 
{ 
$vmhost = $sheet.Cells.Item($row+$i,$colHost).text 
$array = $sheet.Cells.Item($row+$i,$colArray).text 
$port = $sheet.Cells.Item($row+$i,$colPort).text 

Write-Host ("        member HDS_"+ $array +"_" + $port +"_" + $vmhost + "_" + $hba)
} 
Write-Host ("   exit")
write-host("#### Coments only #####  see below")
Write-Host ("!zoneset activate name <zoneset name> vsan xxxx")
write-host ("!show zone pending-diff")
write-host("!zone commit vsan xxxx")
write-host("!show zone active")
write-host("!end")
write-host("!copy run start")


##close excel file
$objExcel.Workbooks.Close()
$objExcel.Quit()
 
