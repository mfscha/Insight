Connect-Ucs 10.207.27.105 -Credential (Get-Credential)

$vnic_template = "ESXi_VM_B"
 
$csv = Import-Csv C:\Users\kwallace\Dropbox\UCS\Scripts\vlan.csv
 
foreach ($row in $csv)
    {
    Invoke-UcsXml -XmlQuery "<configConfMo cookie='$curConUCS.cookie'><inConfig><vnicLanConnTempl dn='org-root/lan-conn-templ-$vnic_template'
    status='created,modified'><vnicEtherIf defaultNet='no' rn='if-$($row.name)'></vnicEtherIf></vnicLanConnTempl></inConfig></configConfMo>"
    }
 
Disconnect-Ucs