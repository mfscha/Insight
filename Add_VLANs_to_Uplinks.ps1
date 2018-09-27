Connect-Ucs 10.207.27.105 -Credential (Get-Credential)
 
$csv = Import-Csv C:\Users\kwallace\Dropbox\UCS\Scripts\vlan.csv
 
foreach ($row in $csv)
{
write-host "Adding VLAN $($row.vlanid) with name $($row.name)"
Add-Ucsvlan -Id $row.vlanid -Name $row.name -Sharing none -LanCloud (Get-UcsLanCloud)
}
 
Disconnect-Ucs