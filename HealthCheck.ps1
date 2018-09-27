###############################################################################################
# Cisco UCS Health Check Report Version 3 by Jeremy Waldrop of Varrow - jwaldrop@varrow.com   #                			     
###############################################################################################	

######################################################################################################################################
# Setup instructions 											   			     #
#   - Microsoft PowerShell 2 is required. PowerShell 2 is part of Windows 7/2008 R2			   			     #
#   - Cisco UCS PowerTool v 0.98 or higher for UCSM - http://developer.cisco.com/web/unifiedcomputing/pshell-download                #
#   - A local UCSM read-only user or domain user						     		                     #
#   - Put the user name in the $user variable below										     #
#   - Put the password in the $password variable below										     #
#   - Put the UCSM VIP IP address in the 0.0.0.0 feilds									             #
#   - For email key in your SMTP server and email address and change the $enablemail variable to yes   			             #
#   - The default path for the HTML file is C:\ if you want it in a different path change the $ReportFile variable                   #
#   - Rename the UCS-HealthCheck-v3.ps1.txt to just .ps1						                             #
#																     #
######################################################################################################################################

#########################################
# Import the Cisco UCS PowerTool module #
#########################################
Import-Module CiscoUcsPS

####################################################
# Authenticate to UCSM with a local read-only user #
####################################################
$user = "psuser"
$password = "yourpasswordhere" | ConvertTo-SecureString -AsPlainText -Force
$cred = New-Object system.Management.Automation.PSCredential($user, $password)
$handle1 = Connect-Ucs 0.0.0.0 -NotDefault -Credential $cred
$handle2 = Connect-Ucs 0.0.0.0 -NotDefault -Credential $cred
$handleArray = $handle1, $handle2

##################
# Mail variables #
##################
$enablemail="no"
$smtpServer = "smtpserver" 
$mailfrom = "Cisco UCS Healtcheck <ucshcheck@yourdomain.com>"
$mailto = "user@yourdomain.com"

########################
# Create the HTML file #
########################
$ReportFile = "C:\UCS-HealthCheck-v3.html"
New-Item -ItemType file $ReportFile -Force

###################
# Define the HTML #
###################
Add-Content $ReportFile "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN'  'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>"
Add-Content $ReportFile "<html xmlns='http://www.w3.org/1999/xhtml'>"
Add-Content $ReportFile "<head>"
Add-Content $ReportFile "<style type='text/css'>"
Add-Content $ReportFile "div.content {"
Add-Content $ReportFile "    border: #48f solid 3px;"
Add-Content $ReportFile "    clear: left;"
Add-Content $ReportFile "    padding: 1em;"
Add-Content $ReportFile "    font-family: Tahoma;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "div.content.inactive {"
Add-Content $ReportFile "	display: none;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc {"
Add-Content $ReportFile "    height: 2em;"
Add-Content $ReportFile "    list-style: none;"
Add-Content $ReportFile "    margin: 0;"
Add-Content $ReportFile "    padding: 0;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc a {"
Add-Content $ReportFile "    background: #bdf url(tabs.gif);"
Add-Content $ReportFile "    color: #008;"
Add-Content $ReportFile "    display: block;"
Add-Content $ReportFile "    float: left;"
Add-Content $ReportFile "    height: 2em;"
Add-Content $ReportFile "    padding-left: 10px;"
Add-Content $ReportFile "    text-decoration: none;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc a:hover {"
Add-Content $ReportFile "    background-color: #3af;"
Add-Content $ReportFile "    background-position: 0 -120px;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc a:hover span {"
Add-Content $ReportFile "    background-position: 100% -120px;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc li {"
Add-Content $ReportFile "    float: left;"
Add-Content $ReportFile "    margin: 0 1px 0 0;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc li a.active {"
Add-Content $ReportFile "    background-color: #48f;"
Add-Content $ReportFile "    background-position: 0 -60px;"
Add-Content $ReportFile "    color: #fff;"
Add-Content $ReportFile "    font-weight: bold;"
Add-Content $ReportFile "    font-family: Tahoma;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc li a.active span {"
Add-Content $ReportFile "    background-position: 100% -60px;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc span {"
Add-Content $ReportFile "    background: url(tabs.gif) 100% 0;"
Add-Content $ReportFile "    display: block;"
Add-Content $ReportFile "    line-height: 2em;"
Add-Content $ReportFile "    padding-right: 10px;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "</style>"

####################
# Define HTML Tabs #
####################
Add-Content $ReportFile "<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />"
Add-Content $ReportFile "<title>UCS Health Check Report</title>"
Add-Content $ReportFile "</head>"
Add-Content $ReportFile "<body>"
Add-Content $ReportFile "<body style=font-family:Tahoma>"
Add-Content $ReportFile "<h1>UCS Health Check Report</h1>"
Add-Content $ReportFile "<ol id='toc'>"    
Add-Content $ReportFile "<li><a href='#page-1'><span>Hardware Inventory</span></a></li>" 
Add-Content $ReportFile "<li><a href='#page-2'><span>Cluster State</span></a></li>"
Add-Content $ReportFile "<li><a href='#page-3'><span>Configuration Summary</span></a></li>"   
Add-Content $ReportFile "<li><a href='#page-4'><span>Environmental Statistics</span></a></li>"
Add-Content $ReportFile "<li><a href='#page-5'><span>Ethernet Statistics</span></a></li>"
Add-Content $ReportFile "<li><a href='#page-6'><span>SAN FC Statistics</span></a></li>"
Add-Content $ReportFile "<li><a href='#page-7'><span>Fault Report</span></a></li>"
Add-Content $ReportFile "</ol>"

###########################
# Define HTML Table Style #
###########################
Add-Content $ReportFile "<style>BODY{background-color:LightGoldenRodYellow;}TABLE{font-family: Tahoma;border-width: 2px;border-style: solid;border-color: black;border-collapse: collapse;}TH{border-width: 2px;padding: 2px;border-style: solid;border-color: black;background-color:LightGray}TD{border-width: 2px;padding: 2px;border-style: solid;border-color: black;background-color:white}</style>"

#################################
# Tab 1, UCS Hardware Inventory #
#################################
Add-Content $ReportFile "<div class='content' id='page-1'>"

################################
# Get UCS Fabric Interconnects #
################################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnects</H2>"
Get-UcsNetworkElement -Ucs $handleArray | Select-Object Ucs,Rn,OobIfIp,OobIfMask,OobIfGw,Operability,Model,Serial | ConvertTo-Html -Fragment | Out-File ucsfis.html
Get-Content ucsfis.html | Add-Content $ReportFile

#########################################
# Get UCS Fabric Interconnect inventory #
#########################################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect Inventory</H2>"
Get-UcsFiModule -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,Model,Descr,OperState,State,Serial | ConvertTo-Html -Fragment | Out-File ucsfi-inv.html
Get-Content ucsfi-inv.html | Add-Content $ReportFile

#########################
# Get UCS License usage #
#########################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect License Usage</H2>"
Get-UcsLicense -Ucs $handleArray | Sort-Object -Property Ucs | Select-Object Ucs,Scope,AbsQuant,UsedQuant,PeerStatus,OperState | ConvertTo-Html -Fragment | Out-File ucsfilic-inv.html
Get-Content ucsfilic-inv.html | Add-Content $ReportFile

########################
# Get UCS Chassis info #
########################
Add-Content $ReportFile "<H2>Cisco UCS Chassis Inventory</H2>"
Get-UcsChassis -Ucs $handleArray | Sort-Object -Property Ucs,Rn | Select-Object Ucs,Rn,Model,OperState,Thermal,Serial | ConvertTo-Html -Fragment | Out-File ucschassis-inv.html
Get-Content ucschassis-inv.html | Add-Content $ReportFile

##############################
# Get chassis IOM (FEX) info #
##############################
Add-Content $ReportFile "<H2>Cisco UCS IOM (FEX) Inventory</H2>"
Get-UcsIom -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,ChassisId,Rn,Model,OperState,Side,Thermal,Serial | ConvertTo-Html -Fragment | Out-File ucsiom-inv.html
Get-Content ucsiom-inv.html | Add-Content $ReportFile

#######################################
# Get all UCS servers and server info #
#######################################
Add-Content $ReportFile "<H2>Cisco UCS Server Inventory</H2>"
Get-UcsBlade -Ucs $handleArray | Sort-Object -Property Ucs,ChassisID,SlotID | Select-Object Ucs,ChassisId,SlotId,Model,AvailableMemory,AssignedToDn,OperState,Operability,OperPower,Serial | ConvertTo-Html -Fragment | Out-File ucsserver-inv.html
Get-Content ucsserver-inv.html | Add-Content $ReportFile

################################################
# Get UCS server adaptor (mezzanine card) info #
################################################
Add-Content $ReportFile "<H2>Cisco UCS Server Adaptor Inventory</H2>"
Get-UcsAdaptorUnit -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,ChassisId,BladeId,Model | ConvertTo-Html -Fragment | Out-File ucsadaptor-inv.html
Get-Content ucsadaptor-inv.html| Add-Content $ReportFile

#################################
# Get UCS server processor info #
#################################
Add-Content $ReportFile "<H2>Cisco UCS Server CPU Inventory</H2>"
Get-UcsProcessorUnit -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,SocketDesignation,Cores,CoresEnabled,Threads,Speed,OperState,Thermal,Model | Where-Object {$_.OperState -ne "removed"} | ConvertTo-Html -Fragment | Out-File ucsservercpu-inv.html
Get-Content ucsservercpu-inv.html| Add-Content $ReportFile

##############################
# Get UCS server memory info #
##############################
Add-Content $ReportFile "<H2>Cisco UCS Server Memory Inventory</H2>"
Get-UcsMemoryUnit -Ucs $handleArray | Sort-Object -Property Ucs,Dn,Location | where {$_.Capacity -ne "unspecified"} | Select-Object -Property Ucs,Dn,Location,Capacity,Clock,OperState,Model | ConvertTo-Html -Fragment | Out-File ucsservermemory-inv.html
Get-Content ucsservermemory-inv.html| Add-Content $ReportFile

################
# End of Tab 1 #
################
Add-Content $ReportFile "</div>"

############################
# Tab 2, UCS Cluster State #
############################
Add-Content $ReportFile "<div class='content' id='page-2'>"

#########################
# Get UCS Cluster State #
#########################
Add-Content $ReportFile "<H2>Cisco UCS Cluster State</H2>"
Get-UcsStatus -Ucs $handleArray | Select-Object Name,VirtualIpv4Address,HaConfiguration,HaReadiness,HaReady,FiALeadership,FiAOobIpv4Address,FiAManagementServicesState,FiBLeadership,FiBOobIpv4Address,FiBManagementServicesState | ConvertTo-Html -Fragment | Out-File ucscluster.html
Get-Content ucscluster.html | Add-Content $ReportFile

################
# End of Tab 2 #
################
Add-Content $ReportFile "</div>"

############################
# Tab 3, UCS Configuration #
############################
Add-Content $ReportFile "<div class='content' id='page-3'>"

###########################################
# Get UCS Cluster Configuration and State #
###########################################
Add-Content $ReportFile "<H2>Cisco UCS Cluster Configuration</H2>"
Get-UcsStatus -Ucs $handleArray | Select-Object Name,VirtualIpv4Address,HaConfiguration,HaReadiness,HaReady,FiALeadership,FiAOobIpv4Address,FiAOobIpv4DefaultGateway,FiAManagementServicesState,FiBLeadership,FiBOobIpv4Address,FiBOobIpv4DefaultGateway,FiBManagementServicesState | ConvertTo-Html -Fragment | Out-File ucsclusterconfig-inv.html
Get-Content ucsclusterconfig-inv.html | Add-Content $ReportFile

###########################################
# Get UCS Global chassis discovery policy #
###########################################
Add-Content $ReportFile "<H2>Cisco UCS Chassis Discovery Policy</H2>"
Get-UcsChassisDiscoveryPolicy -Ucs $handleArray | Select-Object Ucs,Rn,Action | ConvertTo-Html -Fragment | Out-File ucschassisdisc-pol.html
Get-Content ucschassisdisc-pol.html | Add-Content $ReportFile

##################################################
# Get UCS Global chassis power redundancy policy #
##################################################
Add-Content $ReportFile "<H2>Cisco UCS Chassis Power Redundancy Policy</H2>"
Get-UcsPowerControlPolicy -Ucs $handleArray | Select-Object Ucs,Rn,Redundancy | ConvertTo-Html -Fragment | Out-File ucschassispower-pol.html
Get-Content ucschassispower-pol.html | Add-Content $ReportFile

#########################
# Get UCS Organizations #
#########################
Add-Content $ReportFile "<H2>Cisco UCS Organizations</H2>"
Get-UcsOrg -Ucs $handleArray | Select-Object Ucs,Name,Dn | ConvertTo-Html -Fragment | Out-File ucsorgs.html
Get-Content ucsorgs.html | Add-Content $ReportFile

####################################################
# Get running firmware version from all components #
####################################################
Add-Content $ReportFile "<H2>Cisco UCS Firmware Versions</H2>"
Get-UcsFirmwareRunning -Ucs $handleArray | Select-Object Ucs,Dn,Type,Version |Sort-Object -Property Ucs,Dn | Where-Object -FilterScript {$_.Type -notlike "unspecified"} | ConvertTo-Html -Fragment | Out-File ucsfirmwarever.html
Get-Content ucsfirmwarever.html | Add-Content $ReportFile

##############################
# Get UCS LAN Switching Mode #
##############################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect Ethernet Switching Mode</H2>"
Get-UcsLanCloud -Ucs $handleArray | Select-Object Ucs,Rn,Mode | ConvertTo-Html -Fragment | Out-File ucsfiethmode.html
Get-Content ucsfiethmode.html | Add-Content $ReportFile

##############################
# Get UCS SAN Switching Mode #
##############################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect Fibre Channel Switching Mode</H2>"
Get-UcsSanCloud -Ucs $handleArray | Select-Object Ucs,Rn,Mode | ConvertTo-Html -Fragment | Out-File ucsfifcmode.html
Get-Content ucsfifcmode.html | Add-Content $ReportFile

#################################################################
# Get UCS Fabric Interconnect Ethernet port usage and role info #
#################################################################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect Ethernet Port Configuration</H2>"
Get-UcsFabricPort -Ucs $handleArray | Select-Object Ucs,Dn,IfRole,LicState,Mode,OperState,OperSpeed,XcvrType | Where-Object {$_.OperState -eq "up"} | ConvertTo-Html -Fragment | Out-File ucsfiethportroles.html
Get-Content ucsfiethportroles.html | Add-Content $ReportFile

#######################################################
# Get UCS Fabric Interconnect to Chassis port mapping #
#######################################################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect Chassis IOM Mappings</H2>"
Get-UcsEtherSwitchIntFIo -Ucs $handleArray | Select-Object Ucs,ChassisId,Discovery,Model,OperState,SwitchId,PeerSlotId,PeerPortId,SlotId,PortId,XcvrType | ConvertTo-Html -Fragment | Out-File ucsfichassisiommap.html
Get-Content ucsfichassisiommap.html | Add-Content $ReportFile

###############################################
# Get UCS Fabric Interconnect FC Uplink Ports #
###############################################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect FC Uplink Ports</H2>"
Get-UcsFiFcPort -Ucs $handleArray | Select-Object Ucs,EpDn,SwitchId,SlotId,PortId,LicState,Mode,OperSpeed,OperState | sort-object -descending  | where-object {$_.OperState -ne "sfp-not-present"} | ConvertTo-Html -Fragment | Out-File ucsfifcuplinkports.html
Get-Content ucsfifcuplinkports.html | Add-Content $ReportFile

##################################################
# Get SAN Fiber Channel Uplink Port Channel info #
##################################################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect FC Uplink Port Channels</H2>"
Get-UcsFcUplinkPortChannel -Ucs $handleArray | Select-Object Ucs,Dn,Name,OperSpeed,OperState,Transport | ConvertTo-Html -Fragment | Out-File ucsfifcpc.html
Get-Content ucsfifcpc.html | Add-Content $ReportFile

##############################################
# Get Ethernet LAN Uplink Port Channel info #
##############################################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect Ethernet Uplink Port Channels</H2>"
Get-UcsUplinkPortChannel -Ucs $handleArray | Sort-Object -Property Ucs,Name | Select-Object Ucs,Dn,Name,OperSpeed,OperState,Transport | ConvertTo-Html -Fragment | Out-File ucsfiethpc.html
Get-Content ucsfiethpc.html | Add-Content $ReportFile

#############################################################
# Get Ethernet LAN Uplink Port Channel port membership info #
#############################################################
Add-Content $ReportFile "<H2>Cisco UCS Fabric Interconnect Ethernet Uplink Port Channel Members</H2>"
Get-UcsUplinkPortChannelMember -Ucs $handleArray | Sort-Object -Property Ucs,Dn |Select-Object Ucs,Dn,Membership | ConvertTo-Html -Fragment | Out-File ucsfiethpcmembers.html
Get-Content ucsfiethpcmembers.html | Add-Content $ReportFile

########################################
# Get UCS Native Authentication Source #
########################################
Add-Content $ReportFile "<H2>Cisco UCS Native Authentication</H2>"
Get-UcsNativeAuth -Ucs $handleArray | Select-Object Ucs,Rn,DefLogin,ConLogin | ConvertTo-Html -Fragment | Out-File ucsnativeauth.html
Get-Content ucsnativeauth.html | Add-Content $ReportFile

############################
# Get UCS LDAP server info #
############################
Add-Content $ReportFile "<H2>Cisco UCS LDAP Providers</H2>"
Get-UcsLdapProvider -Ucs $handleArray | Select-Object Ucs,Name,Rootdn,Basedn,Attribute | ConvertTo-Html -Fragment | Out-File ucsldapproviders.html
Get-Content ucsldapproviders.html | Add-Content $ReportFile

###############################
# Get UCS LDAP group mappings #
###############################
Add-Content $ReportFile "<H2>Cisco UCS LDAP Group Mappings</H2>"
Get-UcsLdapGroupMap -Ucs $handleArray | Select-Object Ucs,Name | ConvertTo-Html -Fragment | Out-File ucsldapgroupmaps.html
Get-Content ucsldapgroupmaps.html | Add-Content $ReportFile

#####################################
# Get UCS user and LDAP group roles #
#####################################
Add-Content $ReportFile "<H2>Cisco UCS User Roles</H2>"
Get-UcsUserRole -Ucs $handleArray | Select-Object Ucs,Name,Dn | Where-Object {$_.Dn -like "sys/ldap-ext*"} | ConvertTo-Html -Fragment | Out-File ucsuserroles.html
Get-Content ucsuserroles.html | Add-Content $ReportFile

############################
# Get UCS Call Home config #
############################
Add-Content $ReportFile "<H2>Cisco UCS Call Home Configuration</H2>"
Get-UcsCallhome -Ucs $handleArray | Sort-Object -Property Ucs | Select-Object Ucs,AdminState | ConvertTo-Html -Fragment | Out-File ucscallhomecnfg.html
Get-Content ucscallhomecnfg.html | Add-Content $ReportFile

#################################
# Get UCS Call Home SMTP Server #
#################################
Add-Content $ReportFile "<H2>Cisco UCS Call Home SMTP Server</H2>"
Get-UcsCallhomeSmtp -Ucs $handleArray | Sort-Object -Property Ucs | Select-Object Ucs,Host | ConvertTo-Html -Fragment | Out-File ucscallhomesmtpsrv.html
Get-Content ucscallhomesmtpsrv.html | Add-Content $ReportFile

################################
# Get UCS Call Home Recipients #
################################
Add-Content $ReportFile "<H2>Cisco UCS Call Home Recipients</H2>"
Get-UcsCallhomeRecipient -Ucs $handleArray | Sort-Object -Property Ucs | Select-Object Ucs,Dn,Email | ConvertTo-Html -Fragment | Out-File ucscallhomerecipients.html
Get-Content ucscallhomerecipients.html | Add-Content $ReportFile

##############################
# Get UCS SNMP Configuration #
##############################
Add-Content $ReportFile "<H2>Cisco UCS SNMP Configuration</H2>"
Get-UcsSnmp -Ucs $handleArray | Sort-Object -Property Ucs | Select-Object Ucs,AdminState,Community,SysContact,SysLocation | ConvertTo-Html -Fragment | Out-File ucscallsnmpcfg.html
Get-Content ucscallsnmpcfg.html | Add-Content $ReportFile


#######################
# Get UCS DNS Servers #
#######################
Add-Content $ReportFile "<H2>Cisco UCS DNS Servers</H2>"
Get-UcsDnsServer -Ucs $handleArray | Select-Object Ucs,Name | ConvertTo-Html -Fragment | Out-File ucsdnsservers.html
Get-Content ucsdnsservers.html | Add-Content $ReportFile

#######################
# Get UCS NTP Servers #
#######################
Add-Content $ReportFile "<H2>Cisco UCS NTP Servers</H2>"
Get-UcsNtpServer -Ucs $handleArray | Select-Object Ucs,Name | ConvertTo-Html -Fragment | Out-File ucsntpservers.html
Get-Content ucsntpservers.html | Add-Content $ReportFile

####################
# Get UCS Timezone #
####################
Add-Content $ReportFile "<H2>Cisco UCS Timezone</H2>"
Get-UcsTimezone -Ucs $handleArray | Select-Object Ucs,Timezone | ConvertTo-Html -Fragment | Out-File ucstimezone.html
Get-Content ucstimezone.html | Add-Content $ReportFile

#############################
# Get UCS IP CIMC MGMT Pool #
#############################
Add-Content $ReportFile "<H2>Cisco UCS CIMC IP Pool</H2>"
Get-UcsIpPoolBlock -Ucs $handleArray | Select-Object Ucs,Dn,From,To,Subnet,DefGw | Where-Object {$_.DefGw -ne "0.0.0.0"} | ConvertTo-Html -Fragment | Out-File ucscimcippool.html
Get-Content ucscimcippool.html | Add-Content $ReportFile

#########################################
# Get UCS IP CIMC MGMT Pool Assignments #
#########################################
Add-Content $ReportFile "<H2>Cisco UCS CIMC IP Pool Assignments</H2>"
Get-UcsIpPoolAddr -Ucs $handleArray | Sort-Object -Property Ucs,AssignedToDn | where {$_.Assigned -eq "yes"} | Select-Object Ucs,AssignedToDn,Id | ConvertTo-Html -Fragment | Out-File ucscimcpooladdr.html
Get-Content ucscimcpooladdr.html | Add-Content $ReportFile

#############################
# Get UCS UUID Suffix Pools #
#############################
Add-Content $ReportFile "<H2>Cisco UCS UUID Pools</H2>"
Get-UcsUuidSuffixPool -Ucs $handleArray | Select-Object Ucs,Name,Prefix,Size,Assigned | ConvertTo-Html -Fragment | Out-File ucsuuidpool.html
Get-Content ucsuuidpool.html | Add-Content $ReportFile

###################################
# Get UCS UUID Suffix Pool Blocks # 
###################################
Add-Content $ReportFile "<H2>Cisco UCS UUID Pool Blocks</H2>"
Get-UcsUuidSuffixBlock -Ucs $handleArray | Select-Object Ucs,Dn,From,To | ConvertTo-Html -Fragment | Out-File ucsuuidpoolblocks.html
Get-Content ucsuuidpoolblocks.html | Add-Content $ReportFile

######################################
# Get UCS UUID UUID Pool Assignments #
######################################
Add-Content $ReportFile "<H2>Cisco UCS UUID Pool Assignments</H2>"
Get-UcsUuidpoolAddr -Ucs $handleArray | Where-Object {$_.Assigned -ne "no"} | select-object Ucs,AssignedToDn,Id | sort-object -property Ucs,AssignedToDn | ConvertTo-Html -Fragment | Out-File ucsuuidpooladdr.html
Get-Content ucsuuidpooladdr.html | Add-Content $ReportFile

##############################
# Get UCS MAC Address Pools #
#############################
Add-Content $ReportFile "<H2>Cisco UCS MAC Address Pools</H2>"
Get-UcsMacPool -Ucs $handleArray | Select-Object Ucs, Name,Size,Assigned | ConvertTo-Html -Fragment | Out-File ucsmacpools.html
Get-Content ucsmacpools.html | Add-Content $ReportFile

##################################
# Get UCS MAC Address Pool Blocks#
##################################
Add-Content $ReportFile "<H2>Cisco UCS MAC Address Pool Blocks</H2>"
Get-UcsMacMemberBlock -Ucs $handleArray | Select-Object Ucs,Dn,From,To | ConvertTo-Html -Fragment | Out-File ucsmacpoolblocks.html
Get-Content ucsmacpoolblocks.html | Add-Content $ReportFile

################################
# Get UCS MAC Pool Assignments #
################################
Add-Content $ReportFile "<H2>Cisco UCS MAC Address Pool Assignments</H2>"
Get-UcsVnic -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,IdentPoolName,Addr | where {$_.Addr -ne "derived"} | ConvertTo-Html -Fragment | Out-File ucsmacpooladdr.html
Get-Content ucsmacpooladdr.html | Add-Content $ReportFile

######################
# Get UCS WWNN Pools #
######################
Add-Content $ReportFile "<H2>Cisco UCS WWN Pools</H2>"
Get-UcsWwnPool -Ucs $handleArray | Select-Object Ucs,Name,Purpose,Size,Assigned | ConvertTo-Html -Fragment | Out-File ucswwnpools.html
Get-Content ucswwnpools.html | Add-Content $ReportFile

######################################
# Get UCS WWNN/WWPN Pool Assignments #
######################################
Add-Content $ReportFile "<H2>Cisco UCS WWN Pool Assignments</H2>"
Get-UcsVhba -Ucs $handleArray | Sort-Object -Property Ucs,Addr | Select-Object Ucs,Dn,IdentPoolName,NodeAddr,Addr | where {$_.NodeAddr -ne "vnic-derived"} | ConvertTo-Html -Fragment | Out-File ucswwnpooladdr.html
Get-Content ucswwnpooladdr.html | Add-Content $ReportFile

##################################################
# Get UCS WWNN/WWPN vHBA and adaptor Assignments #
##################################################
Add-Content $ReportFile "<H2>Cisco UCS vHBA Details</H2>"
Get-UcsAdaptorHostFcIf -Ucs $handleArray | sort-object -Property Ucs,VnicDn -Descending | Select-Object Ucs,VnicDn,Vendor,Model,LinkState,SwitchId,NodeWwn,Wwn | Where-Object {$_.NodeWwn -ne "00:00:00:00:00:00:00:00"} | ConvertTo-Html -Fragment | Out-File ucsvhbadetails.html
Get-Content ucsvhbadetails.html | Add-Content $ReportFile

########################
# Get UCS Server Pools #
########################
Add-Content $ReportFile "<H2>Cisco UCS Server Pools</H2>"
Get-UcsServerPool -Ucs $handleArray | Select-Object Ucs,Name,Assigned | ConvertTo-Html -Fragment | Out-File ucsserverpools.html
Get-Content ucsserverpools.html | Add-Content $ReportFile

###################################
# Get UCS Server Pool Assignments #
###################################
Add-Content $ReportFile "<H2>Cisco UCS Server Pool Assignments</H2>"
Get-UcsServerPoolAssignment -Ucs $handleArray | Select-Object Ucs,Name,AssignedToDn | ConvertTo-Html -Fragment | Out-File ucsserverpoolassigned.html
Get-Content ucsserverpoolassigned.html | Add-Content $ReportFile

###################################
# Get UCS QoS Class Configuration #
###################################
Add-Content $ReportFile "<H2>Cisco UCS QoS System Class Configuration</H2>"
Get-UcsQosClass -Ucs $handleArray | Select-Object Ucs,Priority,AdminState,Cos,Weight,Drop,Mtu | ConvertTo-Html -Fragment | Out-File ucsqossys.html
Get-Content ucsqossys.html | Add-Content $ReportFile

########################
# Get UCS QoS Policies #
########################
Add-Content $ReportFile "<H2>Cisco UCS QoS Policies</H2>"
Get-UcsQosPolicy -Ucs $handleArray | Select-Object Ucs,Name | ConvertTo-Html -Fragment | Out-File ucsqospol.html
Get-Content ucsqospol.html | Add-Content $ReportFile

####################################
# Get UCS Network Control Policies #
####################################
Add-Content $ReportFile "<H2>Cisco UCS Network Control Policies</H2>"
Get-UcsNetworkControlPolicy -Ucs $handleArray | Select-Object Ucs,Name,Cdp,UplinkFailAction | ConvertTo-Html -Fragment | Out-File ucsnetctrlpol.html
Get-Content ucsnetctrlpol.html | Add-Content $ReportFile

##########################
# Get UCS vNIC Templates #
##########################
Add-Content $ReportFile "<H2>Cisco UCS vNIC Templates</H2>"
Get-UcsVnicTemplate -Ucs $handleArray | Select-Object Ucs,Name,Descr,SwitchId,TemplType,IdentPoolName,Mtu,NwCtrlPolicyName,QosPolicyName | ConvertTo-Html -Fragment | Out-File ucsvnictemplts.html
Get-Content ucsvnictemplts.html | Add-Content $ReportFile

##########################
# Get UCS vHBA Templates #
##########################
Add-Content $ReportFile "<H2>Cisco UCS vHBA Templates</H2>"
Get-UcsVhbaTemplate -Ucs $handleArray | Select-Object Ucs,Name,Descr,SwitchId,TemplType,QosPolicyName | ConvertTo-Html -Fragment | Out-File ucsvhbatemplts.html
Get-Content ucsvhbatemplts.html | Add-Content $ReportFile

#################################
# Get UCS vHBA to VSAN Mappings #
#################################
Add-Content $ReportFile "<H2>Cisco UCS vHBA to VSAN Mappings</H2>"
Get-UcsVhbaInterface -Ucs $handleArray | Select-Object Ucs,Dn,OperVnetName,Initiator | Where-Object {$_.Initiator -ne "00:00:00:00:00:00:00:00"} | ConvertTo-Html -Fragment | Out-File ucsvhbasanmap.html
Get-Content ucsvhbasanmap.html | Add-Content $ReportFile

#####################################
# Get UCS Service Profile Templates #
#####################################
Add-Content $ReportFile "<H2>Cisco UCS Service Profile Templates</H2>"
Get-UcsServiceProfile -Ucs $handleArray | Where-object {$_.UuidSuffix -eq "0000-000000000000"}  | Sort-object -Property Ucs,Name | Select-Object Ucs,Dn,Name,BiosProfileName,BootPolicyName,HostFwPolicyName,LocalDiskPolicyName,MaintPolicyName,MgmtFwPolicyName,VconProfileName | ConvertTo-Html -Fragment | Out-File ucssptemplts.html
Get-Content ucssptemplts.html | Add-Content $ReportFile

############################
# Get UCS Service Profiles #
############################
Add-Content $ReportFile "<H2>Cisco UCS Service Profiles</H2>"
Get-UcsServiceProfile -Ucs $handleArray | Where-object {$_.UuidSuffix -ne "0000-000000000000"}  | Sort-object -Property Ucs,Name | Select-Object Ucs,Dn,Name,AssocState,PnDn,BiosProfileName,IdentPoolName,Uuid,BootPolicyName,HostFwPolicyName,LocalDiskPolicyName,MaintPolicyName,MgmtFwPolicyName,VconProfileName,OperState | ConvertTo-Html -Fragment | Out-File ucssps.html
Get-Content ucssps.html | Add-Content $ReportFile

##########################
# Get UCS Ethernet VLANs #
##########################
Add-Content $ReportFile "<H2>Cisco UCS Ethernet VLANs</H2>"
Get-UcsVlan -Ucs $handleArray | where {$_.IfRole -eq "network"} | Sort-Object -Property Ucs,Id | Select-Object Ucs,Id,Name,SwitchId | ConvertTo-Html -Fragment | Out-File ucsvlans.html
Get-Content ucsvlans.html | Add-Content $ReportFile

##########################################
# Get UCS Ethernet VLAN to vNIC Mappings #
##########################################
Add-Content $ReportFile "<H2>Cisco UCS Ethernet VLAN to vNIC Mappings</H2>"
Get-UcsAdaptorVlan -Ucs $handleArray | sort-object Ucs,Dn |Select-Object Ucs,Dn,Name,Id,SwitchId | ConvertTo-Html -Fragment | Out-File ucsvnic2vlans.html
Get-Content ucsvnic2vlans.html | Add-Content $ReportFile

#########################
# Get UCS FC VSAN info #
#########################
Add-Content $ReportFile "<H2>Cisco UCS FC VSANs</H2>"
Get-UcsVsan -Ucs $handleArray | Select-Object Ucs,Dn,Id,FcoeVlan,DefaultZoning | ConvertTo-Html -Fragment | Out-File ucsvsans.html
Get-Content ucsvsans.html | Add-Content $ReportFile

########################################
# Get UCS FC Port Channel VSAN Mapping #
########################################
Add-Content $ReportFile "<H2>Cisco UCS FC VSAN to FC Port Mappings</H2>"
Get-UcsVsanMemberFcPortChannel -Ucs $handleArray | Select-Object Ucs,EpDn,IfType | ConvertTo-Html -Fragment | Out-File ucsvsanmap.html
Get-Content ucsvsanmap.html | Add-Content $ReportFile


################
# End of Tab 3 #
################
Add-Content $ReportFile "</div>"

#######################################
# Tab 4, UCS Environmental Statistics #
#######################################
Add-Content $ReportFile "<div class='content' id='page-4'>"

###################################
# Get UCS chassis power usage stats #
###################################
Add-Content $ReportFile "<H2>Cisco UCS Chassis Power Stats</H2>"
Get-UcsChassisStats -Ucs $handleArray | Select-Object Ucs,Dn,InputPower,InputPowerAvg,InputPowerMax,InputPowerMin,OutputPower,OutputPowerAvg,OutputPowerMax,OutputPowerMin,Suspect | ConvertTo-Html -Fragment | Out-File ucschassispwrstats.html
Get-Content ucschassispwrstats.html | Add-Content $ReportFile

####################################
# Get UCS chassis and FI PSU status #
####################################
Add-Content $ReportFile "<H2>Cisco UCS Chassis and Fabric Interconnect Power Supply Status</H2>"
Get-UcsPsu -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,OperState,Perf,Power,Thermal,Voltage | ConvertTo-Html -Fragment | Out-File ucspsus.html
Get-Content ucspsus.html | Add-Content $ReportFile

#############################
# Get UCS chassis PSU stats #
#############################
Add-Content $ReportFile "<H2>Cisco UCS Chassis Power Supply Stats</H2>"
Get-UcsPsuStats -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,AmbientTemp,AmbientTempAvg,Input210v,Input210vAvg,Output12v,Output12vAvg,OutputCurrentAvg,OutputPowerAvg,Suspect | ConvertTo-Html -Fragment | Out-File ucspsustats.html
Get-Content ucspsustats.html | Add-Content $ReportFile

####################################
# Get UCS chassis and FI fan stats #
####################################
Add-Content $ReportFile "<H2>Cisco UCS Chassis and Fabric Interconnect Fan Stats</H2>"
get-ucsfan -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,Module,Id,Perf,Power,OperState,Thermal | ConvertTo-Html -Fragment | Out-File ucsfanstats.html
Get-Content ucsfanstats.html | Add-Content $ReportFile

##############################################
# Get UCS chassis IO Module (fex) temp stats #
##############################################
Add-Content $ReportFile "<H2>Cisco UCS Chassis IO Module (FEX) Temperature Stats</H2>"
Get-UcsEquipmentIOCardStats -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,AmbientTemp,AmbientTempAvg,Temp,TempAvg,Suspect | ConvertTo-Html -Fragment | Out-File ucsiomtempstats.html
Get-Content ucsiomtempstats.html | Add-Content $ReportFile

###################################
# Get UCS blade power usage stats #
###################################
Add-Content $ReportFile "<H2>Cisco UCS Server Power Stats</H2>"
Get-UcsComputeMbPowerStats -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,ConsumedPower,ConsumedPowerAvg,ConsumedPowerMax,InputCurrent,InputCurrentAvg,InputVoltage,InputVoltageAvg,Suspect | ConvertTo-Html -Fragment | Out-File ucschassispwr.html
Get-Content ucschassispwr.html | Add-Content $ReportFile

##################################
# Get UCS blade temprature stats #
##################################
Add-Content $ReportFile "<H2>Cisco UCS Server Temperature Stats (in Celcius)</H2>"
Get-UcsComputeMbTempStats -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,FmTempSenIo,FmTempSenIoAvg,FmTempSenIoMax,FmTempSenRear,FmTempSenRearAvg,FmTempSenRearMax,Suspect | ConvertTo-Html -Fragment | Out-File ucsbladetemps.html
Get-Content ucsbladetemps.html | Add-Content $ReportFile

###################################
# Get UCS Memory temprature stats #
###################################
Add-Content $ReportFile "<H2>Cisco Memory Temperature Stats (in Celcius)</H2>"
Get-UcsMemoryUnitEnvStats -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,Temperature,TemperatureAvg,TemperatureMax,Suspect | ConvertTo-Html -Fragment | Out-File ucsmemtempstats.html
Get-Content ucsmemtempstats.html | Add-Content $ReportFile

##########################################
# Get UCS CPU Power and temprature stats #
##########################################
Add-Content $ReportFile "<H2>Cisco CPU Power and Temperature Stats (in Celcius)</H2>"
Get-UcsProcessorEnvStats -Ucs $handleArray | Sort-Object -Property Ucs,Dn |Select-Object Ucs,Dn,InputCurrent,InputCurrentAvg,InputCurrentMax,Temperature,TemperatureAvg,TemperatureMax,Suspect | ConvertTo-Html -Fragment | Out-File ucscputempstats.html
Get-Content ucscputempstats.html | Add-Content $ReportFile


################
# End of Tab 4 #
################
Add-Content $ReportFile "</div>"

#######################################
# Tab 5, UCS Ethernet Statistics #
#######################################
Add-Content $ReportFile "<div class='content' id='page-5'>"

##############################################
# Get UCS LAN Uplink Port Channel Loss Stats #
##############################################
Add-Content $ReportFile "<H2>Cisco LAN Uplink Port Channel Loss Stats</H2>"
Get-UcsUplinkPortChannel -Ucs $handleArray | Get-UcsEtherLossStats | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,ExcessCollision,ExcessCollisionDeltaAvg,LateCollision,LateCollisionDeltaAvg,MultiCollision,MultiCollisionDeltaAvg,SingleCollision,SingleCollisionDeltaAvg | ConvertTo-Html -Fragment | Out-File ucsuplanlossstats.html
Get-Content ucsuplanlossstats.html | Add-Content $ReportFile

#################################################
# Get UCS LAN Uplink Port Channel Receive Stats #
#################################################
Add-Content $ReportFile "<H2>Cisco LAN Uplink Port Channel Receive Stats</H2>"
Get-UcsUplinkPortChannel -Ucs $handleArray | Get-UcsEtherRxStats | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,BroadcastPackets,BroadcastPacketsDeltaAvg,JumboPackets,JumboPacketsDeltaAvg,MulticastPackets,MulticastPacketsDeltaAvg,TotalBytes,TotalBytesDeltaAvg,TotalPackets,TotalPacketsDeltaAvg,Suspect | ConvertTo-Html -Fragment | Out-File ucsuplanrxtats.html
Get-Content ucsuplanrxtats.html | Add-Content $ReportFile

##############################################
# Get UCS LAN Uplink Port Channel Transmit Stats #
##############################################
Add-Content $ReportFile "<H2>Cisco LAN Uplink Port Channel Trasmit Stats</H2>"
Get-UcsUplinkPortChannel -Ucs $handleArray | Get-UcsEtherTxStats | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,BroadcastPackets,BroadcastPacketsDeltaAvg,JumboPackets,JumboPacketsDeltaAvg,MulticastPackets,MulticastPacketsDeltaAvg,TotalBytes,TotalBytesDeltaAvg,TotalPackets,TotalPacketsDeltaAvg,Suspect | ConvertTo-Html -Fragment | Out-File ucsuplantxtats.html
Get-Content ucsuplantxtats.html | Add-Content $ReportFile

######################
# Get UCS vNIC Stats #
######################
Add-Content $ReportFile "<H2>Cisco UCS vNIC Stats</H2>"
Get-UcsAdaptorVnicStats -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,BytesRx,BytesRxDeltaAvg,BytesTx,BytesTxDeltaAvg,PacketsRx,PacketsRxDeltaAvg,PacketsTx,PacketsTxDeltaAvg,DroppedRx,DroppedRxDeltaAvg,DroppedTx,DroppedTxDeltaAvg,ErrorsTx,ErrorsTxDeltaAvg,Suspect | ConvertTo-Html -Fragment | Out-File ucsvnictats.html
Get-Content ucsvnictats.html | Add-Content $ReportFile

################
# End of Tab 5 #
################
Add-Content $ReportFile "</div>"

############################
# Tab 6, UCS FC Statistics #
############################
Add-Content $ReportFile "<div class='content' id='page-6'>"

##############################################
# Get UCS LAN Uplink Port Channel Loss Stats #
##############################################
Add-Content $ReportFile "<H2>Cisco UCS FC Uplink Port Stats</H2>"
Get-UcsFcErrStats -Ucs $handleArray | Sort-Object -Property Ucs,Dn | Select-Object Ucs,Dn,CrcRx,CrcRxDeltaAvg,DiscardRx,DiscardRxDeltaAvg,DiscardTx,DiscardTxDeltaAvg,LinkFailures,SignalLosses,Suspect | ConvertTo-Html -Fragment | Out-File ucsfcuplinkstats.html
Get-Content ucsfcuplinkstats.html | Add-Content $ReportFile


################
# End of Tab 6 #
################
Add-Content $ReportFile "</div>"

###########################
# Tab 7, UCS Fault Report #
###########################
Add-Content $ReportFile "<div class='content' id='page-7'>"

#########################################
# Get all UCS Faults sorted by severity #
#########################################
Add-Content $ReportFile "<H2>Cisco UCS Faults</H2>"
Get-UcsFault -Ucs $handleArray | Sort-Object -Property Ucs,Severity -Descending | Select-Object Ucs,Severity,Created,Descr,dn | ConvertTo-Html -Fragment | Out-File ucsfaults.html
Get-Content ucsfaults.html | Add-Content $ReportFile

################
# End of Tab 7 #
################
Add-Content $ReportFile "</div>"

Add-Content $ReportFile "</body>"


###################################
# Javascript to Activate the Tabs #
###################################

Add-Content $ReportFile "<script type='text/javascript'>"
Add-Content $ReportFile "// Wrapped in a function so as to not pollute the global scope."
Add-Content $ReportFile "var activatables = (function () {"
Add-Content $ReportFile "// The CSS classes to use for active/inactive elements."
Add-Content $ReportFile "var activeClass = 'active';"
Add-Content $ReportFile "var inactiveClass = 'inactive';"
Add-Content $ReportFile "  "
Add-Content $ReportFile "var anchors = {}, activates = {};"
Add-Content $ReportFile "var regex = /#([A-Za-z][A-Za-z0-9:._-]*)$/;"
Add-Content $ReportFile "  "
Add-Content $ReportFile "// Find all anchors (<a href='#something'>.)"
Add-Content $ReportFile "var temp = document.getElementsByTagName('a');"
Add-Content $ReportFile "for (var i = 0; i < temp.length; i++) {"
Add-Content $ReportFile "     var a = temp[i];"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	// Make sure the anchor isn't linking to another page."
Add-Content $ReportFile "	if ((a.pathname != location.pathname &&"
Add-Content $ReportFile "		'/' + a.pathname != location.pathname) ||"
Add-Content $ReportFile "		a.search != location.search) continue;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	// Make sure the anchor has a hash part."
Add-Content $ReportFile "	var match = regex.exec(a.href);"
Add-Content $ReportFile "	if (!match) continue;"
Add-Content $ReportFile "	var id = match[1];"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	// Add the anchor to a lookup table."
Add-Content $ReportFile "	if (id in anchors)"
Add-Content $ReportFile "		anchors[id].push(a);"
Add-Content $ReportFile "	else"
Add-Content $ReportFile "		anchors[id] = [a];"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// Adds/removes the active/inactive CSS classes depending on whether the"
Add-Content $ReportFile "// element is active or not."
Add-Content $ReportFile "function setClass(elem, active) {"
Add-Content $ReportFile "	var classes = elem.className.split(/\s+/);"
Add-Content $ReportFile "	var cls = active ? activeClass : inactiveClass, found = false;"
Add-Content $ReportFile "	for (var i = 0; i < classes.length; i++) {"
Add-Content $ReportFile "		if (classes[i] == activeClass || classes[i] == inactiveClass) {"
Add-Content $ReportFile "			if (!found) {"
Add-Content $ReportFile "				classes[i] = cls;"
Add-Content $ReportFile "				found = true;"
Add-Content $ReportFile "			} else {"
Add-Content $ReportFile "				delete classes[i--];"
Add-Content $ReportFile "			}"
Add-Content $ReportFile "		}"
Add-Content $ReportFile "	}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	if (!found) classes.push(cls);"
Add-Content $ReportFile "	elem.className = classes.join(' ');"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// Functions for managing the hash."
Add-Content $ReportFile "function getParams() {"
Add-Content $ReportFile "	var hash = location.hash || '#';"
Add-Content $ReportFile "	var parts = hash.substring(1).split('&');"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	var params = {};"
Add-Content $ReportFile "	for (var i = 0; i < parts.length; i++) {"
Add-Content $ReportFile "		var nv = parts[i].split('=');"
Add-Content $ReportFile "		if (!nv[0]) continue;"
Add-Content $ReportFile "		params[nv[0]] = nv[1] || null;"
Add-Content $ReportFile "	}"
Add-Content $ReportFile "   "	
Add-Content $ReportFile "	return params;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "function setParams(params) {"
Add-Content $ReportFile "	var parts = [];"
Add-Content $ReportFile "	for (var name in params) {"
Add-Content $ReportFile "		// One of the following two lines of code must be commented out. Use the"
Add-Content $ReportFile "		// first to keep empty values in the hash query string; use the second"
Add-Content $ReportFile "		// to remove them."
Add-Content $ReportFile "		//parts.push(params[name] ? name + '=' + params[name] : name);"
Add-Content $ReportFile "		if (params[name]) parts.push(name + '=' + params[name]);"
Add-Content $ReportFile "	}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	location.hash = knownHash = '#' + parts.join('&');"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// Looks for changes to the hash."
Add-Content $ReportFile "var knownHash = location.hash;"
Add-Content $ReportFile "function pollHash() {"
Add-Content $ReportFile "	var hash = location.hash;"
Add-Content $ReportFile "	if (hash != knownHash) {"
Add-Content $ReportFile "		var params = getParams();"
Add-Content $ReportFile "		for (var name in params) {"
Add-Content $ReportFile "			if (!(name in activates)) continue;"
Add-Content $ReportFile "			activates[name](params[name]);"
Add-Content $ReportFile "		}"
Add-Content $ReportFile "		knownHash = hash;"
Add-Content $ReportFile "	}"
Add-Content $ReportFile "}"
Add-Content $ReportFile "setInterval(pollHash, 250);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "function getParam(name) {"
Add-Content $ReportFile "	var params = getParams();"
Add-Content $ReportFile "	return params[name];"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "function setParam(name, value) {"
Add-Content $ReportFile "	var params = getParams();"
Add-Content $ReportFile "	params[name] = value;"
Add-Content $ReportFile "	setParams(params);"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// If the hash is currently set to something that looks like a single id,"
Add-Content $ReportFile "// automatically activate any elements with that id."
Add-Content $ReportFile "var initialId = null;"
Add-Content $ReportFile "var match = regex.exec(knownHash);"
Add-Content $ReportFile "if (match) {"
Add-Content $ReportFile "	initialId = match[1];"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// Takes an array of either element IDs or a hash with the element ID as the key"
Add-Content $ReportFile "// and an array of sub-element IDs as the value."
Add-Content $ReportFile "// When activating these sub-elements, all parent elements will also be"
Add-Content $ReportFile "// activated in the process."
Add-Content $ReportFile "function makeActivatable(paramName, activatables) {"
Add-Content $ReportFile "	var all = {}, first = initialId;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	// Activates all elements for a specific id (and inactivates the others.)"
Add-Content $ReportFile "	function activate(id) {"
Add-Content $ReportFile "		if (!(id in all)) return false;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "		for (var cur in all) {"
Add-Content $ReportFile "			if (cur == id) continue;"
Add-Content $ReportFile "			for (var i = 0; i < all[cur].length; i++) {"
Add-Content $ReportFile "				setClass(all[cur][i], false);"
Add-Content $ReportFile "			}"
Add-Content $ReportFile "		}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "		for (var i = 0; i < all[id].length; i++) {"
Add-Content $ReportFile "			setClass(all[id][i], true);"
Add-Content $ReportFile "		}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "		setParam(paramName, id);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "		return true;"
Add-Content $ReportFile "	}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	activates[paramName] = activate;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	function attach(item, basePath) {"
Add-Content $ReportFile "		if (item instanceof Array) {"
Add-Content $ReportFile "			for (var i = 0; i < item.length; i++) {"
Add-Content $ReportFile "				attach(item[i], basePath);"
Add-Content $ReportFile "			}"
Add-Content $ReportFile "		} else if (typeof item == 'object') {"
Add-Content $ReportFile "			for (var p in item) {"
Add-Content $ReportFile "				var path = attach(p, basePath);"
Add-Content $ReportFile "				attach(item[p], path);"
Add-Content $ReportFile "			}"
Add-Content $ReportFile "		} else if (typeof item == 'string') {"
Add-Content $ReportFile "			var path = basePath ? basePath.slice(0) : [];"
Add-Content $ReportFile "			var e = document.getElementById(item);"
Add-Content $ReportFile "			if (e)"
Add-Content $ReportFile "				path.push(e);"
Add-Content $ReportFile "			else "
Add-Content $ReportFile "				return;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "			if (!first) first = item;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "			// Store the elements in a lookup table."
Add-Content $ReportFile "			all[item] = path;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "			// Attach a function that will activate the appropriate element"
Add-Content $ReportFile "			// to all anchors."
Add-Content $ReportFile "			if (item in anchors) {"
Add-Content $ReportFile "				// Create a function that will call the 'activate' function with"
Add-Content $ReportFile "				// the proper parameters. It will be used as the event callback."
Add-Content $ReportFile "				var func = (function (id) {"
Add-Content $ReportFile "					return function (e) {"
Add-Content $ReportFile "						activate(id);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "						if (!e) e = window.event;"
Add-Content $ReportFile "						if (e.preventDefault) e.preventDefault();"
Add-Content $ReportFile "						e.returnValue = false;"
Add-Content $ReportFile "						return false;"
Add-Content $ReportFile "					};"
Add-Content $ReportFile "				})(item);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "				for (var i = 0; i < anchors[item].length; i++) {"
Add-Content $ReportFile "					var a = anchors[item][i];"
Add-Content $ReportFile "   "
Add-Content $ReportFile "					if (a.addEventListener) {"
Add-Content $ReportFile "						a.addEventListener('click', func, false);"
Add-Content $ReportFile "					} else if (a.attachEvent) {"
Add-Content $ReportFile "						a.attachEvent('onclick', func);"
Add-Content $ReportFile "					} else {"
Add-Content $ReportFile "						throw 'Unsupported event model.';"
Add-Content $ReportFile "					}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "					all[item].push(a);"
Add-Content $ReportFile "				}"
Add-Content $ReportFile "			}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "			return path;"
Add-Content $ReportFile "		} else {"
Add-Content $ReportFile "			throw 'Unexpected type.';"
Add-Content $ReportFile "		}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "		return basePath;"
Add-Content $ReportFile "	}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	attach(activatables);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "	// Activate an element."
Add-Content $ReportFile "	if (first) activate(getParam(paramName)) || activate(first);"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "return makeActivatable;"
Add-Content $ReportFile "})();"
Add-Content $ReportFile "   "
Add-Content $ReportFile "activatables('page', ['page-1', 'page-2', 'page-3','page-4','page-5','page-6','page-7']);"
Add-Content $ReportFile "</script>"
Add-Content $ReportFile "</html>"

#####################
# Remove temp files #
#####################
Remove-Item ucsfis.html
Remove-Item ucsfi-inv.html
Remove-Item ucschassis-inv.html
Remove-Item ucsiom-inv.html
Remove-Item ucsserver-inv.html
Remove-Item ucsadaptor-inv.html
Remove-Item ucsservercpu-inv.html
Remove-Item ucsservermemory-inv.html
Remove-Item ucsclusterconfig-inv.html
Remove-Item ucschassisdisc-pol.html
Remove-Item ucschassispower-pol.html
Remove-Item ucscluster.html
Remove-Item ucsfirmwarever.html
Remove-Item ucsfiethmode.html
Remove-Item ucsfifcmode.html
Remove-Item ucsfiethportroles.html
Remove-Item ucsfichassisiommap.html
Remove-Item ucsfifcuplinkports.html
Remove-Item ucsfifcpc.html 
Remove-Item ucsfiethpc.html
Remove-Item ucsfiethpcmembers.html
Remove-Item ucsorgs.html
Remove-Item ucsnativeauth.html
Remove-Item ucsldapproviders.html
Remove-Item ucsldapgroupmaps.html
Remove-Item ucsuserroles.html
Remove-Item ucsdnsservers.html
Remove-Item ucsntpservers.html
Remove-Item ucstimezone.html
Remove-Item ucscimcippool.html
Remove-Item ucscimcpooladdr.html
Remove-Item ucsuuidpool.html
Remove-Item ucsuuidpoolblocks.html
Remove-Item ucsuuidpooladdr.html
Remove-Item ucsmacpools.html
Remove-Item ucsmacpoolblocks.html
Remove-Item ucsmacpooladdr.html
Remove-Item ucswwnpools.html
Remove-Item ucswwnpooladdr.html
Remove-Item ucsserverpools.html
Remove-Item ucsserverpoolassigned.html
Remove-Item ucsqossys.html
Remove-Item ucsqospol.html
Remove-Item ucsnetctrlpol.html
Remove-Item ucsvnictemplts.html
Remove-Item ucsvhbatemplts.html
Remove-Item ucsvhbasanmap.html
Remove-Item ucschassispwr.html
Remove-Item ucschassispwrstats.html
Remove-Item ucsbladetemps.html
Remove-Item ucsmemtempstats.html
Remove-Item ucscputempstats.html
Remove-Item ucsfanstats.html
Remove-Item ucspsus.html
Remove-Item ucspsustats.html
Remove-Item ucsuplanlossstats.html
Remove-Item ucsuplanrxtats.html
Remove-Item ucsuplantxtats.html
Remove-Item ucsvnictats.html
Remove-Item ucsfcuplinkstats.html
Remove-Item ucsfaults.html
Remove-Item ucssptemplts.html
Remove-Item ucssps.html
Remove-Item ucsvlans.html
Remove-Item ucsvsans.html
Remove-Item ucsvsanmap.html
Remove-Item ucsfilic-inv.html
Remove-Item ucscallhomecnfg.html
Remove-Item ucscallhomesmtpsrv.html
Remove-Item ucscallhomerecipients.html
Remove-Item ucscallsnmpcfg.html
Remove-Item ucsvhbadetails.html
Remove-Item ucsvnic2vlans.html
Remove-Item ucsiomtempstats.html

######################
# E-mail HTML output #
######################
if ($enablemail -match "yes") 
{ 
$msg = new-object Net.Mail.MailMessage
$att = new-object Net.Mail.Attachment($ReportFile)
$smtp = new-object Net.Mail.SmtpClient($smtpServer) 
$msg.From = $mailfrom
$msg.To.Add($mailto) 
$msg.Subject = �Cisco UCS Health Check�
$msg.Body = �Cisco UCS Health Check, open the attached HTML file to view the report.�
$msg.Attachments.Add($att) 
$smtp.Send($msg)
}