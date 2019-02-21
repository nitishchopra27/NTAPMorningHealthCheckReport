#requires -version 4
<#
.SYNOPSIS
  Get-MorningHealthCheckReport.ps1 performs Helath/Performance checks on NetApp cDOT/ONTAP 9 storage controllers

.DESCRIPTION
  This script queries OCUM and OPM (version 7.1) application Databases to perform Health Checks

.PARAMETER settingsFilePath
  This is a csv file with contents as below
  cluster,location,ocumServer,opmServer
  snowy001,Sydney,192.168.100.135,192.168.100.137
  thunder001,Rockdale,192.168.100.135,192.168.100.137

.INPUTS
  None

.OUTPUTS
  None

.NOTES
  Version:        1.0
  Author:         Nitish Chopra (nitish@netapp.com)
  Creation Date:  13/07/2018
  Purpose/Change: Automate Daily Health Checks

.EXAMPLE
  Run the script and provide an input csv file
  
  .\Get-MorningHealthCheckReport.ps1 -settingsFilePath .\inputs.csv
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------
[CmdletBinding()]
Param (
  [Parameter(Mandatory=$True,ValueFromPipeLine=$True,ValueFromPipeLineByPropertyName=$True,HelpMessage="Location of csv file")]
  [string[]]$settingsFilePath = (Read-Host "Location of Config File"),
  [Parameter(Mandatory=$False,ValueFromPipeLine=$True,ValueFromPipeLineByPropertyName=$True,HelpMessage="Location of csv file with Management Servers Detail")]
  [string[]]$mgmtServersFilePath = (Read-Host "Location of mgmt Servers File")
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Any Global Declarations go here
[String]$scriptPath     = $PSScriptRoot
[String]$logName        = "Morning_Health_Check_Report_Log.csv"
[String]$scriptLogP     = $scriptPath + "\Logs"
[String]$scriptLogPath  = $scriptPath + "\Logs\" + (Get-Date -uformat "%Y-%m-%d-%H-%M") + "-" + $logName
[String]$htmlName       = "Morning_Health_Check_Report.html"
[String]$scriptHTMLLogPath  = $scriptPath + "\Logs\" + (Get-Date -uformat "%Y-%m-%d-%H-%M") + "-" + $htmlName
[array]$report = @()
[string]$startDate = ((Get-Date).AddHours(-24)).ToString("yyyy:MM:dd:HH")
[string]$strtDate = "'"+$startDate+"'"
[string]$endDate = (Get-Date).ToString("yyyy:MM:dd:HH")
[string]$eDate = "'"+$endDate+"'"
[string]$reportime = (Get-Date).ToString("dd/MM/yyyy HH:mm")
[string]$reportTimeSubject = (Get-Date).ToString("yyyyMMdd")
[string]$ntapImage = "$scriptPath" + "\netapp.jpg"
[string]$rcc_toolcheck_loc = "\\snowySVM050.lab.local\Reports\HealthChecks"
#..................................................................................
# Array with cluster names. These names will be used by HTM Table DR Config Validation
[array]$drClusterArray = @("au1111ntap001",
                           "au2222ntap001",
                           "au3333ntap001",
                           "au4444ntap001",
                           "au5555ntap001",
                           "au6666ntap001",
                           "au7777ntap001",
                           "au8888ntap001",
                           "au9999ntap001",
                           "au1010ntap001",
                           "au1212ntap001",
                           "au1414ntap001")
#..................................................................................
# Email Settings are at the bottom of this script, Modify before using this script
#..................................................................................

#-----------------------------------------------------------[Functions]------------------------------------------------------------
Function MySQL {
  Param(
    [Parameter(
    Mandatory = $true,
    ParameterSetName = '',
    ValueFromPipeline = $true)]
    [string]$Query,
    [Parameter(
    Mandatory = $true,
    ParameterSetName = '',
    ValueFromPipeline = $true)]
    [string]$dbServer,
    [Parameter(
    Mandatory = $true,
    ParameterSetName = '',
    ValueFromPipeline = $true)]
    [string]$switchString
  )

  $MySQLAdminUserName = 'reportuser'
  $MySQLAdminPassword = 'Netapp123'


  $MySQLDatabase = switch ( $switchString ) {
    "ocum" {'ocum_report'}
    "model" {'netapp_model_view'}
    "performance" {'netapp_performance'}
  }
  $MySQLHost = $dbServer
  $ConnectionString = "server=" + $MySQLHost + ";port=3306;Integrated Security=False;uid=" + $MySQLAdminUserName + ";pwd=" + $MySQLAdminPassword + ";database="+$MySQLDatabase

  Try {
    #[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
    [void][System.Reflection.Assembly]::LoadFrom("E:\ssh\L080898\MySql.Data.dll")
    $Connection = New-Object MySql.Data.MySqlClient.MySqlConnection
    $Connection.ConnectionString = $ConnectionString
    $Connection.Open()

    $Command = New-Object MySql.Data.MySqlClient.MySqlCommand($Query, $Connection)
    $DataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($Command)
    $DataSet = New-Object System.Data.DataSet
    $RecordCount = $dataAdapter.Fill($dataSet, "data")
    $DataSet.Tables[0]
  }

  Catch {
    Write-Log -Message "ERROR : Unable to run query : $query `n$Error[0]" -Severity Error
  }

  Finally {
    $Connection.Close()
  }
}
function Write-ErrMsg {
  [CmdletBinding()]
  Param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$msg
  )
  Process {
    $fg_color = "White"
    $bg_color = "Red"
    Write-host $msg -ForegroundColor $fg_color -BackgroundColor $bg_color
  }
}
function Write-Msg {
  [CmdletBinding()]
  Param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$msg
  )
  Process {
    $color = "yellow"
    Write-host ""
    Write-host $msg -foregroundcolor $color
    Write-host ""
  }
}
function Write-Log {
  [CmdletBinding()]
  Param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Message,
 
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('Information','Success','Error')]
    [string]$Severity = 'Information'
  )
  Process { 
    [pscustomobject]@{
    #"Time" = (Get-Date -f g);
    "Severity" = $Severity;
    "Message" = $Message;
    } | Export-Csv -Path $scriptLogPath -Append -NoTypeInformation
  }
}
Function New-dbHealthHTMLTableCell {
  [CmdletBinding()]
  Param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [Object[]]$arrayline,
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$lineitem
  )
  Begin {
    $htmltablecell = $null
  } 
  Process {
    switch ($($arrayline."$lineitem")){
      $success {$htmltablecell = "<td class=""pass"">$($arrayline."$lineitem")</td>"}
      "Pass" {$htmltablecell = "<td class=""pass"">$($arrayline."$lineitem")</td>"}
      "Warn" {$htmltablecell = "<td class=""warn"">$($arrayline."$lineitem")</td>"}
      "Fail" {$htmltablecell = "<td class=""fail"">$($arrayline."$lineitem")</td>"}
      default {$htmltablecell = "<td>$($arrayline."$lineitem")</td>"}
    }
    return $htmltablecell
   }
}
function Average($array) {
  $RunningTotal = 0
  foreach ($i in $array) {
    $RunningTotal += $i
  }
  return ([decimal]($RunningTotal) / [decimal]($array.Length))
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Create Log Directory
if ( -not (Test-Path $scriptLogP) ) { 
       Try{
          New-Item -Type directory -Path $scriptLogP -ErrorAction Stop | Out-Null
       }
       Catch{
          Exit -1;
       }
}
Write-Log -Message "=====================================" -Severity Information
Write-Log -Message " NetApp Environment Report " -Severity Information

# Test for input file 
if (-not (Test-Path $settingsFilePath)) {
    Write-Log -Message "Script Configuration file with Parameters not found." -Severity Error
    Write-Log -Message "Exiting Script." -Severity Error
    exit
}
# Import contents of input csv file
try {
  $contents = Import-Csv -Path $settingsFilePath
  Write-Log -Message "Contents of $settingsFilePath are imported successfully" -Severity Success
}
catch {
  Write-Log -Message “Cannot Get Cluster HA Info: $_.” -Severity Error
  exit 
}

# Create HTML Report

#Common HTML head and styles
$htmlhead="<html>
            <style>
             BODY{font-family: Arial; font-size: 8pt;}
             H1{font-size: 16px;}
             H2{font-size: 14px;}
             H3{font-size: 12px;}
             TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
             TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
             TD{border: 1px solid black; padding: 5px; }
             td.pass{background: #7FFF00;}
             td.warn{background: #FFE600;}
             td.fail{background: #FF0000; color: #ffffff;}
             td.info{background: #85D4FF;}
             </style>
             <body>
             <h1 align=""center"">NetApp Environment Report</h1>
             <h3 align=""center"">Generated: $reportime</h3>"
 
$ocumServersummaryhtml = "<h3>NetApp Health Check Summary</h3>
                         <p>This report contains Health Checks performed on NetApp ONTAP 9 storage systems.</p>"

# Parse through the contents of input csv file, execute Health Checks and create HTML Tables for each cluster
Write-Log -Message '***************************************************************************************************' -Severity Information
$contents | % {
  $cluster    = $_.cluster
  $location   = $_.location
  $ocumServer = $_.ocumserver
  $opmServer  = $_.opmServer

  Write-Log -Message "Executing Health Check for $cluster" -Severity Information
  # create an empty array where cluster heatlh object properties will be added
  [array]$report = $null 
  # get cluster info
  try {
    $clusterInfo = MySQL -Query "SELECT id,name FROM ocum_report.cluster WHERE ocum_report.cluster.name=`"$cluster`"" -dbServer $ocumServer -switchString "ocum"
    Write-Log -Message "Successfully Executed SQL Query: SELECT id,name FROM ocum_report.cluster WHERE ocum_report.cluster.name=$cluster" -Severity Success
  }
  catch {
    Write-Log -Message “Failed Executing SQL Query: SELECT id,name FROM ocum_report.cluster WHERE ocum_report.cluster.name=$cluster" -Severity Error
    Write-Log -Message " $_. " -Severity Error
  }
  # Get the nodes in the cluster
  $clusterID = $($clusterInfo.id)
  try {
    $nodes = MySQL -Query "SELECT id,name FROM ocum_report.clusternode WHERE ocum_report.clusternode.clusterId=`"$clusterID`" ORDER BY name" -dbServer $ocumServer -switchString "ocum"
    Write-Log -Message "Successfully Executed SQL Query: SELECT id,name FROM ocum_report.clusternode WHERE ocum_report.clusternode.clusterId=$clusterID" -Severity Success
  }
  catch {
    Write-Log -Message “Failed Executing SQL Query: SELECT id,name FROM ocum_report.clusternode WHERE ocum_report.clusternode.clusterId=$clusterID" -Severity Error
    Write-Log -Message " $_. " -Severity Error
  }
  
  # Iterate through each node, perform Health Checks and append HTML Table rows
  $($nodes) | % {
    $node = $_.name
    Write-Log -Message "Executing Health Check for $node" -Severity Information
    $nodeid = $_.id
    
    # A null string where Notes get appended. These notes have information on spares, aggregates etc.
    [string]$htmlTableNote = $null
    
    # Create a PSObject where node Health Properties will be added and converted to HTML table rows
    $ocumServerObj = New-Object PSObject
    $ocumServerObj | Add-Member NoteProperty -Name "Location" -Value $location
    
    # Get Cluster nodes Health Status
    try {
      $nodesHealth = MySQL -Query "SELECT name,uptime,isNodeHealthy,currentMode,giveBackState FROM ocum_report.clusternode WHERE ocum_report.clusternode.name=`"$node`"" -dbServer $ocumServer -switchString "ocum"
      Write-Log -Message "Successfully Executed SQL Query: SELECT name,uptime,isNodeHealthy,currentMode,giveBackState FROM ocum_report.clusternode WHERE ocum_report.clusternode.name=$node" -Severity Success
    }
    catch {
      Write-Log -Message “Failed Executing SQL Query: SELECT name,uptime,isNodeHealthy,currentMode,giveBackState FROM ocum_report.clusternode WHERE ocum_report.clusternode.name=$node" -Severity Error
      Write-Log -Message " $_. " -Severity Error
    }
    <#
    # OCUM version 7.1 does not report Hot Spares for Partitioned Drives.
    # Omitting reporting on Hot Spares till we upgrade to OCUM version 7.2 or 9
    try {
      $nodeSpares = MySQL -Query "SELECT name AS 'hotSpares' FROM ocum_report.disk WHERE ocum_report.disk.homeNodeId=`"$nodeid`" AND ocum_report.disk.containerType LIKE 'spare'" -dbServer $ocumServer -switchString "ocum"
      Write-Log -Message "Successfully Executed SQL Query: SELECT name,uptime,isNodeHealthy,currentMode,giveBackState FROM ocum_report.clusternode WHERE ocum_report.clusternode.name=$node" -Severity Success
    }
    catch {
      Write-Log -Message “Failed Executing SQL Query: SELECT name AS 'hotSpares' FROM ocum_report.disk WHERE ocum_report.disk.homeNodeId=$nodeid AND ocum_report.disk.containerType LIKE 'spare'" -Severity Error
      Write-Log -Message " $_. " -Severity Error
    }
    #>
    try {
      $offlineAggrs = MySQL -Query "SELECT name FROM ocum_report.aggregate WHERE ocum_report.aggregate.nodeId=`"$nodeid`" AND ocum_report.aggregate.state NOT LIKE '%online%'" -dbServer $ocumServer -switchString "ocum"
      Write-Log -Message "Successfully Executed SQL Query: SELECT name FROM ocum_report.aggregate WHERE ocum_report.aggregate.nodeId=$nodeid AND ocum_report.aggregate.state NOT LIKE '%online%'" -Severity Success
    }
    catch {
      Write-Log -Message “Failed Executing SQL Query: SELECT name FROM ocum_report.aggregate WHERE ocum_report.aggregate.nodeId=$nodeid AND ocum_report.aggregate.state NOT LIKE '%online%'" -Severity Error
      Write-Log -Message " $_. " -Severity Error
    }
    try {
      $aggrs = MySQL -Query "SELECT name,sizeUsedPercent FROM ocum_report.aggregate WHERE ocum_report.aggregate.nodeId=`"$nodeid`" AND ocum_report.aggregate.hasLocalRoot=0" -dbServer $ocumServer -switchString "ocum"
      Write-Log -Message "Successfully Executed SQL Query: SELECT name,sizeUsedPercent FROM ocum_report.aggregate WHERE ocum_report.aggregate.nodeId=$nodeid AND ocum_report.aggregate.hasLocalRoot=0" -Severity Success
    }
    catch {
      Write-Log -Message “Failed Executing SQL Query: SELECT name,sizeUsedPercent FROM ocum_report.aggregate WHERE ocum_report.aggregate.nodeId=$nodeid AND ocum_report.aggregate.hasLocalRoot=0" -Severity Error
      Write-Log -Message " $_. " -Severity Error
    }
    # Create two empty arrays
    # $wararray contains aggregates that are 75%<= aggrUsed <85%
    # $redarray contains aggregates that are aggrUsed >85%
    Write-Log -Message "Checking if any aggregates are above 75% used" -Severity Information
    $wararray = @()
    $redarray = @()
    $($aggrs) | % {if(($_.sizeUsedPercent -ge 75) -and ($($_.sizeUsedPercent) -lt 85)) {$wararray += '{0} is {1}%' -f $($_.name), $($_.sizeUsedPercent)}}
    $($aggrs) | % {if($($_.sizeUsedPercent) -ge 85) {$redarray += '{0} is {1}%' -f $($_.name), $($_.sizeUsedPercent)}}
    $wararray -join " ; " | Out-Null
    $redarray -join " ; " | Out-Null
    
    $ocumServerObj | Add-Member NoteProperty -Name "NodeName" -Value $node
    
    Write-Log -Message "Starting Health Checks" -Severity Information
    #***************************************************************************************************
    # Node Health Checks
    #***************************************************************************************************
    # node health
    Write-Log -Message "Checking if $node is Healthy" -Severity Information
    if ($($nodesHealth.isNodeHealthy) -eq $true) {
      $ocumServerObj | Add-Member NoteProperty -Name "Health" -Value "Pass" -Force
      Write-Log -Message "$node is Healthy" -Severity Success
      #$ocumServerObj | Add-Member NoteProperty -Name "Health" -Value "$($nodesHealth.isNodeHealthy)"
    }
    else {
      $ocumServerObj | Add-Member NoteProperty -Name "Health" -Value "Fail"
      Write-Log -Message "$node is not Healthy" -Severity Error
    }
    # uptime
    Write-Log -Message "Checking Uptime of $node" -Severity Information
    if ($($nodesHealth.uptime) -gt 86400) {
      $ocumServerObj | Add-Member NoteProperty -Name "Uptime" -Value "Pass" -Force
      Write-Log -Message "$node has an uptime of more than 24 hours" -Severity Success
      #$ocumServerObj | Add-Member NoteProperty -Name "Uptime" -Value "$($nodesHealth.uptime)"
    }
    else {
      $ocumServerObj | Add-Member NoteProperty -Name "Uptime" -Value "Fail"
      Write-Log -Message "$node has an uptime of less than 24 hours" -Severity Error
    }
    # ha status
    Write-Log -Message "Checking HA Status of $node" -Severity Information
    if ($($nodesHealth.currentMode) -eq 'ha') {
      $ocumServerObj | Add-Member NoteProperty -Name "HAStatus" -Value "Pass" -Force
      Write-Log -Message "$node is HA" -Severity Success
      #$ocumServerObj | Add-Member NoteProperty -Name "HAStatus" -Value "$($nodesHealth.currentMode)"
    }
    else {
      $ocumServerObj | Add-Member NoteProperty -Name "HAStatus" -Value "Fail"
      Write-Log -Message "$node is not HA" -Severity Error
    }
    <#
    # hot spares
    # OCUM version 7.1 does not report Hot Spares for Partitioned Drives.
    # Omitting reporting on Hot Spares till we upgrade to OCUM version 7.2 or 9
    Write-Log -Message "Checking available Hot Spares on $node" -Severity Information
    if ($(($nodeSpares.hotSpares).count) -ge 2) {
      $ocumServerObj | Add-Member NoteProperty -Name "HotSpares" -Value "Pass" -Force
      Write-Log -Message "$node has more than 2 Hot Spares" -Severity Success
      #$ocumServerObj | Add-Member NoteProperty -Name "HotSpares" -Value "$(($nodeSpares.hotSpares).count)"
      $htmlTableNote += "[ $(($nodeSpares.hotSpares).count) ] Hot Spares " | Out-String
    }
    else {
      #$ocumServerObj | Add-Member NoteProperty -Name "Hot Spares" -Value "Fail" -Force
      $ocumServerObj | Add-Member NoteProperty -Name "HotSpares" -Value "Warn"
      Write-Log -Message "$node has less than 2 Hot Spares" -Severity Error
      $htmlTableNote += "[ $(($nodeSpares.hotSpares).count) ] Hot Spares " | Out-String
    }
    #>
    # capacity
    Write-Log -Message "Updating Capacity Property of $node" -Severity Information
    if (($wararray -gt 0) -or ($redarray -gt 0)) {
      $aggrstring = ($wararray + $redarray) -join " ; " | out-string
      if ($redarray -gt 0) {
        $ocumServerObj | Add-Member NoteProperty -Name "Capacity" -Value "Fail"
      }
      else {
        $ocumServerObj | Add-Member NoteProperty -Name "Capacity" -Value "Warn"
      }
      $htmlTableNote = ($htmlTableNote + $aggrstring) | Out-String
      #$htmlTableNote = ($htmlTableNote + " ; " + $aggrstring) | Out-String
      Write-Log -Message "$node has aggregates where used capcity is > 75%" -Severity Error
    }
    else {
      $ocumServerObj | Add-Member NoteProperty -Name "Capacity" -Value "Pass"
      Write-Log -Message "All aggregates on $node are below 75% used capacity" -Severity Success
    }
    $ocumServerObj | Add-Member NoteProperty -Name "Notes" -Value "$htmlTableNote"
    
    Write-Log -Message "Completed Health Checks" -Severity Information
    
    #***************************************************************************************************
    # Check Node Performance
    #***************************************************************************************************
    Write-Log -Message "Starting Performance Checks on $node" -Severity Information
    try {
      $perfnodeid = MySQL -Query "SELECT objid,name FROM node WHERE name LIKE `"%$node`"" -dbServer $opmServer -switchString model
      Write-Log -Message "Successfully Executed SQL Query: SELECT objid,name FROM node WHERE name LIKE $node" -Severity Success
    }
    catch {
      Write-Log -Message “Failed Executing SQL Query: SELECT objid,name FROM node WHERE name LIKE $node" -Severity Error
      Write-Log -Message " $_. " -Severity Error
    }
    try {
      $perfavgLatency = MySQL -Query "SELECT round((avgLatency/1000),2) AS avgLatency FROM sample_node WHERE objid=`"$($perfnodeid.objid)`" AND (Date_Format(FROM_UNIXTIME(time/1000),'%Y:%m:%d:%H') between $strtDate and $eDate)" -dbServer $opmServer -switchString performance
      Write-Log -Message "Successfully Executed SQL Query: SELECT round((avgLatency/1000),2) AS avgLatency FROM sample_node WHERE objid=$($perfnodeid.objid) AND (Date_Format(FROM_UNIXTIME(time/1000),'%Y:%m:%d:%H') between $strtDate and $eDate)" -Severity Success
    }
    catch {
      Write-Log -Message “Failed Executing SQL Query: Successfully Executed SQL Query: SELECT round((avgLatency/1000),2) AS avgLatency FROM sample_node WHERE objid=$($perfnodeid.objid) AND (Date_Format(FROM_UNIXTIME(time/1000),'%Y:%m:%d:%H') between $strtDate and $eDate)" -Severity Error
      Write-Log -Message " $_. " -Severity Error
    }
    Write-Log -Message "There are $(($perfavgLatency.avgLatency).count) entries in perfavgLatency Array" -Severity Information

    # calculate average latency for last 24 hours and if the value is > 15 ms, report as RED
    $perfobj = Average($($perfavgLatency.avgLatency))
    $perfobj = [math]::Round($perfobj,2)
    Write-Log -Message "$node has an average latency of : $perfobj : for last 24 hours"
    
    if ($perfobj -gt 15) {
      $ocumServerObj | Add-Member NoteProperty -Name "NodeLatency" -Value "Fail"
    }
    else {$ocumServerObj | Add-Member NoteProperty -Name "NodeLatency" -Value "Pass"}
    Write-Log -Message "Completed Performance Checks on $node" -Severity Information
    <#
    # Below code will check if SLO on latency is breached
    # SLO is breached with a node has a latnecy of >15ms for 15 minute period
    # In below code, i create batches with three latency values, take average of each batch. If avg of any batch is >15, we breach SLO
    # $perfobj references to the Latency values in array perfavgLatency   
    $perfobj = $($perfavgLatency.avgLatency)
    $myPerfTmpObj = @()
    $latencyBreached = $false
    $k=1
    for($i = 0; $i -lt $perfobj.Length; $i += 3) {
      # end index
      $j = $i + 2
      if ($j -ge $perfobj.Length) {
        $j = $perfobj.Length - 1
      }
      # create tmpObj which contains items 1-3, then items 4-6, then 7-9 etc
      $myPerfTmpObj += [math]::Round(($perfobj[$i..$j] | Measure-Object -Average).Average)
    }
    Write-Log -Message "There are $(($myPerfTmpObj).count) batches for calculating Avg Latency" -Severity Information
        
    foreach ($item in $myPerfTmpObj ) {
      if ($item -ge 15) {
        $latencyBreached = $true
        Write-Log -Message "Latency value of $item was seen on $node" -Severity Information
      }
    }
    if ($latencyBreached -eq $true) {
      $ocumServerObj | Add-Member NoteProperty -Name "NodeLatency" -Value "Fail"
    }
    else {$ocumServerObj | Add-Member NoteProperty -Name "NodeLatency" -Value "Pass"}
    Write-Log -Message "Completed Performance Checks on $node" -Severity Information
    #>
    # Update HTML report with all the object properties collected above
    $report = $report + $ocumServerObj
  }
  
  # Create an HTML table with rows
  $htmltableheader = "<h3>$location</h3>
                    <p>
                    <table>
                    <tr>
                    <th>Hostname</th>
                    <th>Health</th>
                    <th>Availability</th>
                    <th>HA Status</th>
                    <th>Capacity</th>
                    <th>Performance</th>
                    <th>Notes</th>
                    </tr>"

  $ocumServerhealthhtmltable = $ocumServerhealthhtmltable + $htmltableheader


  foreach ($reportline in $report) {
    $htmltablerow = "<tr>"
    $htmltablerow += "<td>$($reportline.NodeName)</td>"
    $htmltablerow += (New-dbHealthHTMLTableCell $reportline "Health")
    $htmltablerow += (New-dbHealthHTMLTableCell $reportline "Uptime")
    $htmltablerow += (New-dbHealthHTMLTableCell $reportline "HAStatus")
    $htmltablerow += (New-dbHealthHTMLTableCell $reportline "Capacity")
    $htmltablerow += (New-dbHealthHTMLTableCell $reportline "NodeLatency")
    $htmltablerow += "<td>$($reportline.Notes)</td>"

    $ocumServerhealthhtmltable = $ocumServerhealthhtmltable + $htmltablerow
  }   
  Write-Log -Message "Successfully updated HTML table for $cluster" -Severity Success

  $ocumServerhealthhtmltable = $ocumServerhealthhtmltable + "</table></p>"
  Write-Log -Message '***************************************************************************************************' -Severity Information
}

# Mgmt Servers Connectivity Information
# Test for input file 
if (Test-Path $mgmtServersFilePath) {

  # create an empty array where cluster heatlh object properties will be added
  [array]$mgmtreport = $null 

  Write-Log -Message "Mgmt Server information file found" -Severity Success
  # Import contents of input csv file
  try {
    $mgmtcontents = Import-Csv -Path $mgmtServersFilePath
    Write-Log -Message "Contents of $mgmtServersFilePath are imported successfully" -Severity Success
  }
  catch {
    Write-Log -Message “Cannot Get contents of $mgmtServersFilePath : $_.” -Severity Error
  }

  $mgmtcontents | % {
    $url = $_.url
    $app = $_.app

    # Create a PSObject where node Health Properties will be added and converted to HTML table rows
    $mgmtServerObj = New-Object PSObject
    $mgmtServerObj | Add-Member NoteProperty -Name "Hostname" -Value "$app"
    # A null string where Notes get appended. These notes have information on spares, aggregates etc.
    [string]$mgmthtmlTableNote = $null

    Try {
      [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$True}
      $request  = [System.Net.WebRequest]::Create($url)
      $response = $request.GetResponse()
      $status   = [int]$response.StatusCode
    }
    Catch {
      Write-Log -Message $("Failed enumerating status for ""$app"". Error " + $_.Exception.Message) -Severity Error
      $status = $Null
      $response = $Null
    }
    If($status -ne $Null) { 
      Write-Log -Message $("$app +  is OK!") -Severity Success
      $mgmtServerObj | Add-Member NoteProperty -Name "Availability" -Value "Pass"
    }
    Else {
      Write-Log -Message $("$app +  may be down please check!") -Severity Error
      $message = $app + " may be down, please check!"
      $mgmtServerObj | Add-Member NoteProperty -Name "Availability" -Value "Fail"
      $mgmthtmlTableNote += $message | Out-String
      $mgmtServerObj | Add-Member NoteProperty -Name "Notes" -Value "$mgmthtmlTableNote"
    }
    # Update HTML report with all the object properties collected above
    $mgmtreport = $mgmtreport + $mgmtServerObj
  }

  # Create an HTML table with rows
  $mgmthtmltableheader = "<h3>Hosting Platform</h3>
                    <p>
                    <table>
                    <tr>
                    <th>Hostname</th>
                    <th>Availability</th>
                    <th>Notes</th>
                    </tr>"

  $mgmthealthhtmltable = $mgmthealthhtmltable + $mgmthtmltableheader

  foreach ($mgmtreportline in $mgmtreport) {
    $mgmthtmltablerow = "<tr>"
    $mgmthtmltablerow += "<td>$($mgmtreportline.Hostname)</td>"
    $mgmthtmltablerow += (New-dbHealthHTMLTableCell $mgmtreportline "Availability")
    $mgmthtmltablerow += "<td>$($mgmtreportline.Notes)</td>"

    $mgmthealthhtmltable = $mgmthealthhtmltable + $mgmthtmltablerow
  }   
  Write-Log -Message "Successfully updated HTML table for Mgmt Applicatons" -Severity Success

  $mgmthealthhtmltable = $mgmthealthhtmltable + "</table></p>"
  Write-Log -Message '***************************************************************************************************' -Severity Information
  # Complete HTML Code
  $mgmthtmltail = "</body>
               </html>" 
  $mgmthtmlreport = $mgmthtmlhead + $mgmthealthhtmltable + $mgmthtmltail
  $mgmttableyes = $true
}
#######################################
# RCC Mgmt Servers/Applications Connectivity Status
# Test for input file 
$rcc_toolcheckfile = (Get-ChildItem -Path $rcc_toolcheck_loc | Sort LastWriteTime | select -Last 1).Name
$rcc_toolcheck_file = $rcc_toolcheck_loc+"\"+$rcc_toolcheckfile
if (Test-Path $rcc_toolcheck_file) {

  # create an empty array where cluster heatlh object properties will be added
  [array]$rccmgmtreport = $null 

  Write-Log -Message "RCC Mgmt Server information file found" -Severity Success
  # Import contents of input csv file
  try {
    $rccmgmtcontents = Import-Csv -Path $rcc_toolcheck_file
    Write-Log -Message "Contents of $rcc_toolcheck_file are imported successfully" -Severity Success
  }
  catch {
    Write-Log -Message “Cannot Get contents of $rcc_toolcheck_file : $_.” -Severity Error
  }

  $rccmgmtcontents | % {
    $rccapp    = $_.app
    $rccstatus = $_.status

    # Create a PSObject where node Health Properties will be added and converted to HTML table rows
    $rccmgmtServerObj = New-Object PSObject
    $rccmgmtServerObj | Add-Member NoteProperty -Name "Hostname" -Value "$rccapp"
    # A null string where Notes get appended. These notes have information on spares, aggregates etc.
    [string]$rccmgmthtmlTableNote = $null

    $rccmgmtServerObj | Add-Member NoteProperty -Name "Availability" -Value "$rccstatus"
    
    # Update HTML report with all the object properties collected above
    $rccmgmtreport = $rccmgmtreport + $rccmgmtServerObj
  }

  # Create an HTML table with rows
  $rccmgmthtmltableheader = "<h3>Hosting Platform</h3>
                    <p>
                    <table>
                    <tr>
                    <th>Hostname</th>
                    <th>Availability</th>
                    <th>Notes</th>
                    </tr>"

  $rccmgmthealthhtmltable = $rccmgmthealthhtmltable + $rccmgmthtmltableheader

  foreach ($rccmgmtreportline in $rccmgmtreport) {
    $rccmgmthtmltablerow = "<tr>"
    $rccmgmthtmltablerow += "<td>$($rccmgmtreportline.Hostname)</td>"
    $rccmgmthtmltablerow += (New-dbHealthHTMLTableCell $rccmgmtreportline "Availability")
    $rccmgmthtmltablerow += "<td>$($rccmgmtreportline.Notes)</td>"

    $rccmgmthealthhtmltable = $rccmgmthealthhtmltable + $rccmgmthtmltablerow
  }   
  Write-Log -Message "Successfully updated HTML table for RCC Mgmt Applicatons" -Severity Success

  $rccmgmthealthhtmltable = $rccmgmthealthhtmltable + "</table></p>"
  Write-Log -Message '***************************************************************************************************' -Severity Information
  # Complete HTML Code
  $rccmgmthtmltail = "</body>
               </html>" 
  $rccmgmthtmlreport = $rccmgmthtmlhead + $rccmgmthealthhtmltable + $rccmgmthtmltail
  $rccmgmttableyes = $true
}
#######################################
# Create HTML Table for Open Changes
# create an empty array where cluster heatlh object properties will be added
[array]$changereport = $null 

Write-Log -Message "Creating Table for Open Changes" -Severity Information
# Create a PSObject for Open changes HTML Table
$changeObj = New-Object PSObject
$changeObj | Add-Member NoteProperty -Name "ChangeNumber" -Value "N/A"
$changeObj | Add-Member NoteProperty -Name "DaysOpen" -Value ""
$changeObj | Add-Member NoteProperty -Name "Notes" -Value ""
# Update HTML report with all the object properties collected above
$changereport = $changereport + $changeObj

# Create an HTML table with rows
$changehtmltableheader = "<h3>Open Changes</h3>
                    <p>
                    <table>
                    <tr>
                    <th>Change Number</th>
                    <th>Days Open</th>
                    <th>Notes</th>
                    </tr>"
$changehealthhtmltable = $changehealthhtmltable + $changehtmltableheader

foreach ($changereportline in $changereport) {
    $changehtmltablerow = "<tr>"
    $changehtmltablerow += "<td>$($changereportline.ChangeNumber)</td>"
    $changehtmltablerow += "<td>$($changereportline.DaysOpen)</td>"
    $changehtmltablerow += "<td>$($changereportline.Notes)</td>"

    $changehealthhtmltable = $changehealthhtmltable + $changehtmltablerow
}   
Write-Log -Message "Successfully updated HTML table for Open Changes" -Severity Success

$changehealthhtmltable = $changehealthhtmltable + "</table></p>"
Write-Log -Message '***************************************************************************************************' -Severity Information
# Complete HTML Code
$changehtmltail = "</body>
                   </html>" 
$changehtmlreport = $changehtmlhead + $changehealthhtmltable + $changehtmltail
#######################################
# Create HTML Table for DR Config
# create an empty array where DR Config object properties will be added
[array]$drreport = $null 

Write-Log -Message "Creating Table for DR Config" -Severity Information
$drClusterArray | % {
  # Create a PSObject for Open drs HTML Table
  $drObj = New-Object PSObject
  $drObj | Add-Member NoteProperty -Name "Cluster" -Value "$_"
  $drObj | Add-Member NoteProperty -Name "Checked" -Value "Pass"
  $drObj | Add-Member NoteProperty -Name "Notes" -Value ""
  # Update HTML report with all the object properties collected above
  $drreport = $drreport + $drObj
}
# Create an HTML table with rows
$drhtmltableheader = "<h3>DR Config Validation</h3>
                    <p>
                    <table>
                    <tr>
                    <th>Cluster</th>
                    <th>Checked</th>
                    <th>Notes</th>
                    </tr>"
$drhealthhtmltable = $drhealthhtmltable + $drhtmltableheader

foreach ($drreportline in $drreport) {
    $drhtmltablerow = "<tr>"
    $drhtmltablerow += "<td>$($drreportline.Cluster)</td>"
    $drhtmltablerow += "<td class=""pass"">$($drreportline.Checked)</td>"
    $drhtmltablerow += "<td>$($drreportline.Notes)</td>"

    $drhealthhtmltable = $drhealthhtmltable + $drhtmltablerow
}   
Write-Log -Message "Successfully updated HTML table for DR Configs" -Severity Success

$drhealthhtmltable = $drhealthhtmltable + "</table></p>"
Write-Log -Message '***************************************************************************************************' -Severity Information
# Complete HTML Code
$drhtmltail = "</body>
               </html>" 
$drhtmlreport = $drhtmlhead + $drhealthhtmltable + $drhtmltail
#######################################

# Complete HTML Code for OCUM Health
$htmltail = "</body>
             </html>" 
$htmlreport = $htmlhead + $ocumServersummaryhtml + $ocumServerhealthhtmltable + $htmltail

if (($mgmtTableyes -eq $True) -and ($rccmgmttableyes = $True)){
  $htmlreport = $htmlreport + $mgmthtmlreport + $rccmgmthtmlreport + $changehtmlreport + $drhtmlreport
}
elseif ($mgmtTableyes -eq $True) {
  $htmlreport = $htmlreport + $mgmthtmlreport + $changehtmlreport + $drhtmlreport
}
elseif ($rccmgmtTableyes -eq $True) {
  $htmlreport = $htmlreport + $rccmgmthtmlreport + $changehtmlreport + $drhtmlreport
}
else {
  $htmlreport = $htmlreport + $changehtmlreport + $drhtmlreport
}

Start-Sleep -Seconds 5

# SEND EMAIL MESSAGE
[string[]]$recipients  = "nitish.chopra@netapp.com","nitish.chopra@lab.local" 
$splat = @{
  'to' = $recipients;
  'subject' = "NetApp Environment Report - " + $reportTimeSubject;
  'SmtpServer' = "appsmtp.lab.local";
  'from' = "NSO_Automation@lab.local";
  'body' = $htmlreport;
  'BodyAsHtml' = $true;
}
Send-MailMessage @splat -Encoding ([System.Text.Encoding]::UTF8)