# NTAPMorningHealthCheckReport
Get-MorningHealthCheckReport.ps1 performs Health/Performance checks on NetApp cDOT/ONTAP 9 storage controllers

This script queries OCUM and OPM (version 7.1) application Databases to perform Health Checks

####This is a csv file with contents as below:
**cluster**|**location**|**ocumServer**|**opmServer**
:-----:|:-----:|:-----:|:-----:
snowy001|Sydney|192.168.100.135|192.168.100.137
thunder001|Rockdale|192.168.100.135|192.168.100.137