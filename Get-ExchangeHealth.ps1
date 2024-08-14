
<#PSScriptInfo

.VERSION 6.0

.GUID d2f58251-eef9-4c92-b19c-10e98387b5c3

.AUTHOR June Castillote

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI https://github.com/junecastillote/Get-ExchangeHealth

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

.PRIVATEDATA
{
  "ReleaseDate": "2024-03-13"
}

#>

<#

.DESCRIPTION
 Use Get-ExchangeHealth.ps1 for gathering and reporting the overall Exchange Server health.

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$configFile,
    [Parameter(Mandatory = $false)]
    [switch]$enableDebug
)
$script_info = Test-ScriptFileInfo $MyInvocation.MyCommand.Definition
# $script_data = $script_info.PrivateData | ConvertFrom-Json
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
if ($enableDebug) { Start-Transcript -Path ($script_root + "\debugLog.txt") }

Function BuildNumberToName {
    param (
        [string]$BuildNumber
    )
    $definitions = Import-Csv "$script_root\ExchangeBuildNumbers.csv"
    $definitions | Where-Object { $_.'Build Number' -eq $BuildNumber }
}


Function TrimExchangeVersion {
    # This function formats the AdminDisplayVersion to Build Number
    # ie. "Version 15.2 (Build 1258.12)" to "15.2.1258.12"
    param (
        [string]$AdminDisplayVersion
    )
    $AdminDisplayVersion.ToString().Replace('Version ', '').Replace(' (Build ', '.').replace(')', '')
}

Function AdminDisplayVersionToName {
    param(
        [string]$AdminDisplayVersion
    )
    BuildNumberToName (TrimExchangeVersion $AdminDisplayVersion)
}

#Import Configuration File
if ((Test-Path $configFile) -eq $false) {
    Write-Host "ERROR: File $($configFile) does not exist. Script cannot continue" -ForegroundColor Yellow
    "ERROR: File $($configFile) does not exist. Script cannot continue" | Out-File error.txt
    if ($enableDebug) { Stop-Transcript }
    return $null
}

$config = Import-PowerShellDataFile $configFile

# Start Script
# $hr = "=" * ($script_info.ProjectUri.OriginalString.Length + 15)
$hr = "=" * ($script_info.ProjectUri.OriginalString.Length)
Write-Host $hr -ForegroundColor Yellow
# Write-Host "Name         : $($script_info.Name) $($script_info.Version)" -ForegroundColor Yellow
Write-Host "$($script_info.Name) $($script_info.Version)" -ForegroundColor Yellow
# Write-Host "Version      : $($script_info.Version)" -ForegroundColor Yellow
# Write-Host "Release Date : $($script_data.ReleaseDate)" -ForegroundColor Yellow
Write-Host "$($script_info.ProjectUri.OriginalString)" -ForegroundColor Yellow
# Write-Host "Repository   : $($script_info.ProjectUri.OriginalString)" -ForegroundColor Yellow
Write-Host $hr -ForegroundColor Yellow
Write-Host ''
Write-Host (Get-Date) ': Begin' -ForegroundColor Green
Write-Host (Get-Date) ': Setting Paths and Variables' -ForegroundColor Yellow

#Define Variables
$testCount = 0
$testFailed = 0
$testPassed = 0
$percentPassed = 0
$overAllResult = "PASSED"
$errSummary = ""
$today = '{0:dd-MMM-yyyy hh:mm tt}' -f (Get-Date)

$css_string = @'
<style type="text/css">
#HeadingInfo
	{
		font-family:"Segoe UI";
		width:100%;
		border-collapse:collapse;
	}
#HeadingInfo td, #HeadingInfo th
	{
		font-size:0.8em;
		padding:3px 7px 2px 7px;
	}
#HeadingInfo th
	{
		font-size:2.0em;
		font-weight:normal;
		text-align:left;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#604767;
		color:#fff;
	}
#SectionLabels
	{
		font-family:"Segoe UI";
		width:100%;
		border-collapse:collapse;
	}
#SectionLabels th.data
	{
		font-size:2.0em;
		text-align:left;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#fff;
		color:#000;
	}
#data
	{
		font-family:"Segoe UI";
		width:100%;
		border-collapse:collapse;
	}
#data td, #data th
	{
		font-size:0.8em;
		border:1px solid #DDD;
		padding:3px 7px 2px 7px;
	}
#data th
	{
		font-size:0.8em;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#00B388;
		color:#fff; text-align:left;
	}
#data td
	{ 	font-size:0.8em;
		padding-top:5px;
		padding-bottom:4px;
		text-align:left;
	}
#data td.bad
	{ 	font-size:0.8em;
		font-weight: bold;
		padding-top:5px;
		padding-bottom:4px;
		color:#f04953;
	}
#data td.good
	{ 	font-size:0.8em;
		font-weight: bold;
		padding-top:5px;
		padding-bottom:4px;
		color:#01a982;
	}

.status {
	width: 10px;
	height: 10px;
	margin-right: 7px;
	margin-bottom: 0px;
	background-color: #CCC;
	background-position: center;
	opacity: 0.8;
	display: inline-block;
}
.green {
	background: #01a982;
}
.purple {
	background: #604767;
}
.orange {
	background: #ffd144;
}
.red {
	background: #f04953;
}
</style>
</head>
<body>
'@

#Thresholds from config
[int]$t_lastfullbackup = $config.thresholds.LastFullBackup
[int]$t_lastincrementalbackup = $config.thresholds.LastIncrementalBackup
[double]$t_DiskBadPercent = $config.thresholds.DiskSpaceFree
[int]$t_mQueue = $config.thresholds.MailQueueCount
[int]$t_copyQueue = $config.thresholds.CopyQueueLenght
[int]$t_replayQueue = $config.thresholds.ReplayQueueLenght
[double]$t_cpuUsage = $config.thresholds.CpuUsage
[double]$t_ramUsage = $config.thresholds.RamUsage

#Options from config
$RunCPUandMemoryReport = $config.reportOptions.RunCPUandMemoryReport
$RunServerHealthReport = $config.reportOptions.RunServerHealthReport
$RunMdbReport = $config.reportOptions.RunMdbReport
$RunComponentReport = $config.reportOptions.RunComponentReport
$RunPdbReport = $config.reportOptions.RunPdbReport
$RunDBCopyReport = $config.reportOptions.RunDBCopyReport
$RunDAGReplicationReport = $config.reportOptions.RunDAGReplicationReport
$RunQueueReport = $config.reportOptions.RunQueueReport
$RunDiskReport = $config.reportOptions.RunDiskReport
$SendReportViaEmail = $config.reportOptions.SendReportViaEmail
# $reportfile = $config.reportOptions.ReportFile
$reportfile = (New-Item -ItemType File -Path $config.reportOptions.ReportFile -Force).FullName


#Mail settings
$CompanyName = $config.mailAndReportParameters.CompanyName
$MailSubject = $config.mailAndReportParameters.MailSubject
$MailServer = $config.mailAndReportParameters.MailServer
$MailSender = $config.mailAndReportParameters.MailSender
$MailTo = @($config.mailAndReportParameters.MailTo)
$MailCc = @($config.mailAndReportParameters.MailCc)
$MailBcc = @($config.mailAndReportParameters.MailBcc)

#Exclusions
$IgnoreServer = @($config.exclusions.IgnoreServer)
$IgnoreDatabase = @($config.exclusions.IgnoreDatabase)
$IgnorePFDatabase = @($config.exclusions.IgnorePFDatabase)
$IgnoreComponent = @($config.exclusions.IgnoreComponent)

Function Get-CPUAndMemoryLoad ($exchangeServers) {
    $stats_collection = @()
    $TopProcessCPU = ""
    $tCounter = 0
    foreach ($exchangeServer in $exchangeServers) {
        #Get CPU Usage
        $x = Get-Counter -Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -computer $exchangeServer.Name | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue

        Write-Host (Get-Date) ": Getting CPU Load for" $exchangeServer.Name -ForegroundColor Yellow
        $cpuMemObject = "" | Select-Object Server, CPU_Usage, Top_CPU_Consumers, Total_Memory_KB, Memory_Free_KB, Memory_Used_KB, Memory_Free_Percent, Memory_Used_Percent, Top_Memory_Consumers
        $cpuMemObject.Server = $exchangeServer.Name
        $cpuMemObject.CPU_Usage = "{0:N0}" -f ($x.cookedvalue)

        #Get Top 3
        $TopProcessCPU = ""
        $y = Get-Counter '\Process(*)\% Processor Time' -computer $exchangeServer.Name | Select-Object -ExpandProperty countersamples | Where-Object { $_.instancename -ne 'idle' -and $_.instancename -ne '_total' } | Select-Object -Property instancename, cookedvalue | Sort-Object -Property cookedvalue -Descending | Select-Object -First 5
        foreach ($tproc in $y) {
            $z = "$($tproc.instancename) `n"
            #$TopProcessCPU += "$z"
            if ($tCounter -ne ($y.count - 1)) {
                $TopProcessCPU += "$z"
                $tCounter = $tCounter + 1
            }
            else {
                $TopProcessCPU += "$z"
            }
        }
        $cpuMemObject.Top_CPU_Consumers = $TopProcessCPU

        Write-Host (Get-Date) ": Getting Memory Load for" $exchangeServer.Name -ForegroundColor Yellow
        $memObj = Get-WmiObject -ComputerName $exchangeServer.Name -Class Win32_operatingsystem -Property CSName, TotalVisibleMemorySize, FreePhysicalMemory
        $cpuMemObject.Total_Memory_KB = $memObj.TotalVisibleMemorySize
        $cpuMemObject.Memory_Free_KB = $memObj.FreePhysicalMemory
        $cpuMemObject.Memory_Used_KB = ($cpuMemObject.Total_Memory_KB - $cpuMemObject.Memory_Free_KB)
        $cpuMemObject.Memory_Used_Percent = "{0:N0}" -f (($cpuMemObject.Memory_Used_KB / $cpuMemObject.Total_Memory_KB) * 100)
        $cpuMemObject.Memory_Free_Percent = "{0:N0}" -f (($cpuMemObject.Memory_Free_KB / $cpuMemObject.Total_Memory_KB) * 100)

        #Get the Top Memory Consumers
        $processes = Get-Process -ComputerName $exchangeServer.Name | Group-Object -Property ProcessName
        $proc_collection = @()
        foreach ($process in $processes) {
            $tempproc = "" | Select-Object Server, ProcessName, MemoryUsed
            $tempproc.ProcessName = $process.Name
            $tempproc.MemoryUsed = (($process.Group | Measure-Object WorkingSet -Sum).sum / 1kb)
            $proc_collection += $tempproc
        }

        $proclist = $proc_collection | Sort-Object MemoryUsed -Descending | Select-Object -First 5

        $TopProcessMemory = ""
        foreach ($proc in $proclist) {
            $topProc = "$($proc.ProcessName) | $($proc.MemoryUsed) KB `n"
            $TopProcessMemory += $topProc
        }

        $cpuMemObject.Top_Memory_Consumers = $TopProcessMemory


        $stats_collection += $cpuMemObject

    }
    Return $stats_collection
}

#Ping function
Function Ping-Server ($server) {
    $ping = Test-Connection $server -Quiet -Count 1
    return $ping
}

Function Get-MdbStatistics ($mailboxdblist) {
    Write-Host (Get-Date) ': Mailbox Database Check... ' -ForegroundColor Yellow # -NoNewline
    $stats_collection = @()
    foreach ($mailboxdb in $mailboxdblist) {
        if (Ping-Server($mailboxdb.Server.Name) -eq $true) {
            $mdbobj = "" | Select-Object Name, Mounted, MountedOnServer, ActivationPreference, DatabaseSize, AvailableNewMailboxSpace, ActiveMailboxCount, DisconnectedMailboxCount, TotalItemSize, TotalDeletedItemSize, EdbFilePath, LogFolderPath, LogFilePrefix, LastFullBackup, LastIncrementalBackup, BackupInProgress, MapiConnectivity, EDBFreeSpace, LogFreeSpace
            if ($mailboxdb.Mounted -eq $true) {
                $mdbStat = Get-MailboxStatistics -Database $mailboxdb
                $mbxItemSize = $mdbStat | ForEach-Object { $_.TotalItemSize.Value } | Measure-Object -Sum
                $mbxDelSize = $mdbStat | ForEach-Object { $_.TotalDeletedItemSize.Value } | Measure-Object -Sum
                $mdbobj.ActiveMailboxCount = ($mdbStat | Where-Object { !$_.DisconnectDate }).count
                $mdbobj.DisconnectedMailboxCount = ($mdbStat | Where-Object { $_.DisconnectDate }).count
                $mdbobj.TotalItemSize = "{0:N2}" -f ($mbxItemSize.sum / 1GB)
                $mdbobj.TotalDeletedItemSize = "{0:N2}" -f ($mbxDelSize.sum / 1GB)
                $mdbobj.MountedOnServer = $mailboxdb.Server.Name
                $mdbobj.ActivationPreference = $mailboxdb.ActivationPreference | Where-Object { $_.Key -eq $mailboxdb.Server.Name }
                $mdbobj.LastFullBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $mailboxdb.LastFullBackup
                $mdbobj.LastIncrementalBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $mailboxdb.LastIncrementalBackup
                $mdbobj.BackupInProgress = $mailboxdb.BackupInProgress
                $mdbobj.DatabaseSize = "{0:N2}" -f ($mailboxdb.DatabaseSize.tobytes() / 1GB)
                $mdbobj.AvailableNewMailboxSpace = "{0:N2}" -f ($mailboxdb.AvailableNewMailboxSpace.tobytes() / 1GB)
                $mdbobj.MapiConnectivity = Test-MapiConnectivity -Database $mailboxdb.Identity -PerConnectionTimeout 10
                #Get Disk Details
                $dbDrive = (Get-WmiObject Win32_LogicalDisk -Computer $mailboxdb.Server.Name | Where-Object { $_.DeviceID -eq $mailboxdb.EdbFilePath.DriveName })
                $logDrive = (Get-WmiObject Win32_LogicalDisk -Computer $mailboxdb.Server.Name | Where-Object { $_.DeviceID -eq $mailboxdb.LogFolderPath.DriveName })
                $mdbobj.EDBFreeSpace = "{0:N2}" -f ($dbDrive.Size / 1GB) + '[' + "{0:N2}" -f ($dbDrive.FreeSpace / 1GB) + ']'
                $mdbobj.LogFreeSpace = "{0:N2}" -f ($logDrive.Size / 1GB) + '[' + "{0:N2}" -f ($logDrive.FreeSpace / 1GB) + ']'
            }
            else {
                $mdbobj.ActiveMailboxCount = "DISMOUNTED"
                $mdbobj.DisconnectedMailboxCount = "DISMOUNTED"
                $mdbobj.TotalItemSize = "DISMOUNTED"
                $mdbobj.TotalDeletedItemSize = "DISMOUNTED"
                $mdbobj.MountedOnServer = "DISMOUNTED"
                $mdbobj.ActivationPreference = "DISMOUNTED"
                $mdbobj.LastFullBackup = "DISMOUNTED"
                $mdbobj.LastIncrementalBackup = "DISMOUNTED"
                $mdbobj.BackupInProgress = "DISMOUNTED"
                $mdbobj.DatabaseSize = "DISMOUNTED"
                $mdbobj.AvailableNewMailboxSpace = "DISMOUNTED"
                $mdbobj.MapiConnectivity = "Failed"
                #Get Disk Details
                $dbDrive = "DISMOUNTED"
                $logDrive = "DISMOUNTED"
                $mdbobj.EDBFreeSpace = "DISMOUNTED"
                $mdbobj.LogFreeSpace = "DISMOUNTED"
            }
            $mdbobj.Name = $mailboxdb.name
            $mdbobj.EdbFilePath = $mailboxdb.EdbFilePath
            $mdbobj.LogFolderPath = $mailboxdb.LogFolderPath
            $mdbobj.Mounted = $mailboxdb.Mounted
        }
        else {
            $mdbobj = "" | Select-Object Name, Mounted, MountedOnServer, ActivationPreference, DatabaseSize, AvailableNewMailboxSpace, ActiveMailboxCount, DisconnectedMailboxCount, TotalItemSize, TotalDeletedItemSize, EdbFilePath, LogFolderPath, LogFilePrefix, LastFullBackup, LastIncrementalBackup, BackupInProgress, MapiConnectivity, EDBFreeSpace, LogFreeSpace
            $mdbobj.Name = $mailboxdb.name
            $mdbobj.EdbFilePath = $mailboxdb.EdbFilePath
            $mdbobj.LogFolderPath = $mailboxdb.LogFolderPath
            $mdbobj.Mounted = "$($mailboxdb.Server.Name): Connection/Ping Failed"
            $mdbobj.MountedOnServer = "-"
            $mdbobj.ActivationPreference = "-"
            $mdbobj.LastFullBackup = "-"
            $mdbobj.LastIncrementalBackup = "-"
            $mdbobj.BackupInProgress = "-"
            $mdbobj.DatabaseSize = "-"
            $mbxItemSize = "-"
            $mbxDelSize = "-"
            $mdbobj.TotalItemSize = "-"
            $mdbobj.TotalDeletedItemSize = "-"
            $mdbobj.ActiveMailboxCount = "-"
            $mdbobj.DisconnectedMailboxCount = "-"
            $mdbobj.AvailableNewMailboxSpace = "-"
            $mdbobj.MapiConnectivity = "-"
            $mdbobj.EDBFreeSpace = "-"
            $mdbobj.LogFreeSpace = "-"
        }
        $stats_collection += $mdbobj
    }
    # Write-Host 'Done' -ForegroundColor Green
    return $stats_collection
}

Function Get-PdbStatistics ($pfdblist) {
    Write-Host (Get-Date) ': Public Folder Database Check... ' -ForegroundColor Yellow # -NoNewline
    $stats_collection = @()
    foreach ($pfdb in $pfdblist) {
        $pfdbobj = "" | Select-Object Name, Mounted, MountedOnServer, DatabaseSize, AvailableNewMailboxSpace, FolderCount, TotalItemSize, LastFullBackup, LastIncrementalBackup, BackupInProgress, MapiConnectivity
        $pfdbobj.Name = $pfdb.Name
        $pfdbobj.Mounted = $pfdb.Mounted
        if ($pfdb.Mounted -eq $true) {
            $pfdbobj.MountedOnServer = $pfdb.Server.Name
            $pfdbobj.DatabaseSize = "{0:N2}" -f ($pfdb.DatabaseSize.tobytes() / 1GB)
            $pfdbobj.AvailableNewMailboxSpace = "{0:N2}" -f ($pfdb.AvailableNewMailboxSpace.tobytes() / 1GB)
            $pfdbobj.LastFullBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $pfdb.LastFullBackup
            $pfdbobj.LastIncrementalBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $pfdb.LastIncrementalBackup
            $pfdbobj.BackupInProgress = $pfdb.BackupInProgress
            $pfdbobj.MapiConnectivity = Test-MapiConnectivity -Database $pfdb.Identity -PerConnectionTimeout 10
        }
        else {
            $pfdbobj.MountedOnServer = "DISMOUNTED"
            $pfdbobj.DatabaseSize = "DISMOUNTED"
            $pfdbobj.AvailableNewMailboxSpace = "DISMOUNTED"
            $pfdbobj.LastFullBackup = "DISMOUNTED"
            $pfdbobj.LastIncrementalBackup = "DISMOUNTED"
            $pfdbobj.BackupInProgress = "DISMOUNTED"
            $pfdbobj.MapiConnectivity = "DISMOUNTED"
        }

        $stats_collection += $pfdbobj
    }
    # Write-Host "Done" -ForegroundColor Green
    return $stats_collection
}

Function Get-DiskSpaceStatistics ($serverlist) {
    Write-Host (Get-Date) ': Disk Space Check... ' -ForegroundColor Yellow # -NoNewline
    $stats_collection = @()
    foreach ($server in $serverlist) {
        try {
            $diskObj = Get-WmiObject Win32_LogicalDisk -Filter 'DriveType=3' -computer $server | Select-Object SystemName, DeviceID, VolumeName, Size, FreeSpace
            foreach ($disk in $diskObj) {
                $serverobj = "" | Select-Object SystemName, DeviceID, VolumeName, Size, FreeSpace, PercentFree
                $serverobj.SystemName = $disk.SystemName
                $serverobj.DeviceID = $disk.DeviceID
                $serverobj.VolumeName = $disk.VolumeName
                $serverobj.Size = "{0:N2}" -f ($disk.Size / 1GB)
                $serverobj.FreeSpace = "{0:N2}" -f ($disk.FreeSpace / 1GB)
                [int]$serverobj.PercentFree = "{0:N0}" -f (($disk.freespace / $disk.size) * 100)
                $stats_collection += $serverobj
            }
        }
        catch {
            $serverobj = "" | Select-Object SystemName, DeviceID, VolumeName, Size, FreeSpace, PercentFree
            $serverobj.SystemName = $server
            $serverobj.DeviceID = $disk.DeviceID
            $serverobj.VolumeName = $disk.VolumeName
            $serverobj.Size = 0
            $serverobj.FreeSpace = 0
            [int]$serverobj.PercentFree = 20000
            $stats_collection += $serverobj
        }
    }
    # Write-Host "Done" -ForegroundColor Green
    return $stats_collection
}

Function Get-ReplicationHealth {
    Write-Host (Get-Date) ': Replication Health Check... ' -ForegroundColor Yellow # -NoNewline
    $stats_collection = Get-MailboxServer | Where-Object { $_.DatabaseAvailabilityGroup } | Sort-Object Name | ForEach-Object { Test-ReplicationHealth -Identity $_ }
    # Write-Host "Done" -ForegroundColor Green
    return $stats_collection
}

Function Get-MailQueues ($transportServerList) {
    Write-Host (Get-Date) ': Mail Queue Check... ' -ForegroundColor Yellow # -NoNewline
    #$stats_collection = get-TransportServer | Where-Object {$_.ServerRole -notmatch 'Edge'} | Sort-Object Name | ForEach-Object {Get-Queue -Server $_}
    $stats_collection = $transportServerList | Sort-Object Name | ForEach-Object { Get-Queue -Server $_ | Where-Object { $_.Identity -notmatch 'Shadow' } }
    # Write-Host "Done" -ForegroundColor Green
    return $stats_collection
}

Function Get-ServerHealth ($serverlist) {
    Write-Host (Get-Date) ': Server Status Check... ' -ForegroundColor Yellow # -NoNewline
    $stats_collection = @()
    foreach ($server in $serverlist) {
        if (Ping-Server($server.name) -eq $true) {
            $exchange_product = (AdminDisplayVersionToName -AdminDisplayVersion $server.AdminDisplayVersion)

            $serverOS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server

            $serverobj = "" | Select-Object Server, ProductName, BuildNumber, KB, Version, Edition, Connectivity, ADSite, UpTime, HubTransportRole, ClientAccessRole, MailboxRole, MailFlow, MessageLatency
            $timespan = $serverOS.ConvertToDateTime($serverOS.LocalDateTime) - $serverOS.ConvertToDateTime($serverOS.LastBootUpTime)
            [int]$uptime = "{0:00}" -f $timespan.TotalHours

            $serverobj.Server = $server.Name
            $serverobj.ProductName = $exchange_product.'Product Name'
            $serverobj.BuildNumber = $exchange_product.'Build Number'
            $serverobj.KB = $exchange_product.KB
            $serverobj.Edition = $server.Edition
            $serverobj.UpTime = $uptime
            $serverobj.Connectivity = "Passed"
            $serviceStatus = Test-ServiceHealth -Server $server
            $serverobj.HubTransportRole = ""
            $serverobj.ClientAccessRole = ""
            $serverobj.MailboxRole = ""
            $site = ($server.site.ToString()).Split("/")
            $serverObj.ADSite = $site[-1]
            foreach ($service in $serviceStatus) {
                if ($service.Role -eq 'Hub Transport Server Role') {
                    if ($service.RequiredServicesRunning -eq $true) {
                        $serverobj.HubTransportRole = "Passed"
                    }
                    else {
                        $serverobj.HubTransportRole = "Failed"
                    }
                }

                if ($service.Role -eq 'Client Access Server Role') {
                    if ($service.RequiredServicesRunning -eq $true) {
                        $serverobj.ClientAccessRole = "Passed"
                    }
                    else {
                        $serverobj.ClientAccessRole = "Failed"
                    }
                }

                if ($service.Role -eq 'Mailbox Server Role') {
                    if ($service.RequiredServicesRunning -eq $true) {
                        $serverobj.MailboxRole = "Passed"
                    }
                    else {
                        $serverobj.MailboxRole = "Failed"
                    }
                }
            }
            #Mail Flow
            if ($server.serverrole -match 'Mailbox' -AND $activeServers -contains $server.name) {
                $mailflowresult = $null
                $result = Test-MailFlow -TargetMailboxServer $server.Name
                $mailflowresult = $result.TestMailflowResult
                $serverObj.MailFlow = $mailflowresult
            }
        }
        else {
            $serverobj = "" | Select-Object Server, Connectivity, ADSite, UpTime, HubTransportRole, ClientAccessRole, MailboxRole

            $site = ($server.site.ToString()).Split("/")
            $serverObj.ADSite = $site[-1]
            $serverobj.Server = $server.Name
            $serverobj.Connectivity = "Failed"
            $serverobj.UpTime = "Cannot retrieve up time"
            $serverobj.HubTransportRole = "Failed"
            $serverobj.ClientAccessRole = "Failed"
            $serverobj.MailboxRole = "Failed"
            $serverObj.MailFlow = "Failed"
            $serverObj.MessageLatency = "Failed"
        }
        $stats_collection += $serverobj
    }
    # Write-Host "Done" -ForegroundColor Green
    #Write-Host $stats_collection
    return $stats_collection
}

Function Get-ServerHealthReport ($serverhealthinfo) {
    Write-Host (Get-Date) ': Server Health Report... ' -ForegroundColor Yellow # -NoNewline
    $mResult = "<tr><td>Server Health Status</td><td class = ""good"">Passed</td></tr>"
    $testFailed = 0
    $mbody = @()
    $errString = @()
    #$currentServer = ""
    $mbody += '<table id="SectionLabels"><tr><th class="data">Server Health Status</th></tr></table>'
    $mbody += '<table id="data">'
    # $mbody += '<tr><th>Server</th><th>Version / Edition</th><th>Site</th><th>Connectivity</th><th>Up Time (Hours)</th><th>Hub Transport Role</th><th>Client Access Role</th><th>Mailbox Role</th><th>Mail Flow</th></tr>'
    $mbody += '<tr><th>Server Name</th><th>Exchange Server Info</th><th>Site</th><th>Connectivity</th><th>Up Time (Hours)</th><th>Hub Transport Role</th><th>Client Access Role</th><th>Mailbox Role</th><th>Mail Flow</th></tr>'
    foreach ($server in $serverhealthinfo) {
        $mbody += "<tr><td>$($server.server)</td><td>Name: $($server.ProductName)<br/>Build: $($server.BuildNumber) [$($server.KB)]<br/>Edition: $($server.Edition)</td><td>$($server.ADSite)</td>"
        #Uptime
        if ($server.UpTime -lt 24) {
            #$errString += "<tr><td>Server Up Time</td></td><td>$($server.server) - up time [$($server.Uptime)] is less than 24 hours</td></tr>"
            $mbody += "<td class = ""good"">$($server.Connectivity)</td><td class = ""bad"">$($server.UpTime)</td>"
        }
        elseif ($server.Uptime -eq 'Cannot retrieve up time') {
            $errString += "<tr><td>Server Connectivity</td></td><td>$($server.server) - connection test failed. SERVER MIGHT BE DOWN!!!</td></tr>"
            $mbody += "<td class = ""bad"">$($server.Connectivity)</td><td class = ""bad"">$($server.UpTime)</td>"
        }
        else {
            $mbody += "<td class = ""good"">$($server.Connectivity)</td><td class = ""good"">$($server.UpTime)</td>"
        }
        #Transport Role
        if ($server.HubTransportRole -eq 'Passed') {
            $mbody += '<td class = "good">Passed</td>'
        }
        elseif ($server.HubTransportRole -eq 'Failed') {
            $errString += "<tr><td>Role Services</td></td><td>$($server.server) - not all required Hub Transport Role services are running</td></tr>"
            $mbody += '<td class = "bad">Failed</td>'
        }
        else {
            $mbody += '<td class = "good"></td>'
        }
        #CAS Role
        if ($server.ClientAccessRole -eq 'Passed') {
            $mbody += '<td class = "good">Passed</td>'
        }
        elseif ($server.ClientAccessRole -eq 'Failed') {
            $errString += "<tr><td>Role Services</td></td><td>$($server.server) - not all required Client Access Role services are running</td></tr>"
            $mbody += '<td class = "bad">Failed</td>'
        }
        else {
            $mbody += '<td class = "good"></td>'
        }
        #Mailbox Role
        if ($server.MailboxRole -eq 'Passed') {
            $mbody += '<td class = "good">Passed</td>'
        }
        elseif ($server.MailboxRole -eq 'Failed') {
            $errString += "<tr><td>Role Services</td></td><td>$($server.server) - not all required Mailbox Role services are running</td></tr>"
            $mbody += '<td class = "bad">Failed</td>'
        }
        else {
            $mbody += '<td class = "good"></td>'
        }

        #Mail Flow
        #write-host $server.server $server.MailFlow
        if ($server.MailFlow -eq "Failed") {
            $errString += "<tr><td>Mail Flow</td></td><td>$($db.Name) - Mail Flow Result FAILED</td></tr>"
            $mbody += '<td class = "bad">Failed</td>'
        }
        elseif ($server.MailFlow -eq 'Success') {
            $mbody += '<td class = "good">Success</td>'
        }
        else {
            $mbody += '<td class = "good"></td>'
        }
        $mbody += '</tr>'
    }
    if ($errString) { $mResult = "<tr><td>Server Health Status</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }
    # Write-Host "Done" -ForegroundColor Green
    return $mbody, $errString, $mResult, $testFailed
}

Function Get-DatabaseCopyStatus ($mailboxdblist) {
    Write-Host (Get-Date) ': Mailbox Database Copy Status Check... ' -ForegroundColor Yellow # -NoNewline
    $stats_collection = @()

    foreach ($db in $mailboxdblist) {
        if ($db.DatabaseCopies.Count -lt 2) {
            continue
        }
        #if ($db.MasterType -eq 'DatabaseAvailabilityGroup')
        #{
        foreach ($dbCopy in $db.DatabaseCopies) {
            $temp = "" | Select-Object Name, Status, CopyQueueLength, LogCopyQueueIncreasing, ReplayQueueLength, LogReplayQueueIncreasing, ContentIndexState, ContentIndexErrorMessage
            $dbStatus = Get-MailboxDatabaseCopyStatus -Identity $dbCopy
            $temp.Name = $dbStatus.Name
            $temp.Status = $dbStatus.Status
            $temp.CopyQueueLength = $dbStatus.CopyQueueLength
            $temp.LogCopyQueueIncreasing = $dbStatus.LogCopyQueueIncreasing
            $temp.ReplayQueueLength = $dbStatus.ReplayQueueLength
            $temp.LogReplayQueueIncreasing = $dbStatus.LogReplayQueueIncreasing
            if ($db.IndexEnabled -eq $false) {
                $temp.ContentIndexState = "Disabled"
                $temp.ContentIndexErrorMessage = $dbStatus.ContentIndexErrorMessage
            }
            else {
                $temp.ContentIndexState = $dbStatus.ContentIndexState
                $temp.ContentIndexErrorMessage = $dbStatus.ContentIndexErrorMessage
            }
            $stats_collection += $temp
        }
        #}
    }
    # Write-Host "Done" -ForegroundColor Green
    return $stats_collection | Sort-Object Name
}

Function Get-DAGCopyStatusReport ($mdbCopyStatus) {
    Write-Host (Get-Date) ': Mailbox Database Copy Status... ' -ForegroundColor Yellow # -NoNewline
    $mResult = "<tr><td>Mailbox Database Copy Status</td><td class = ""good"">Passed</td></tr>"
    $testFailed = 0
    $mbody = @()
    $errString = @()
    $mbody += '<table id="SectionLabels"><tr><th class="data">Mailbox Database Copy Status</th></tr></table>'
    $mbody += '<table id="data">'
    $mbody += '<tr><th>Name</th><th>Status</th><th>CopyQueueLength</th><th>LogCopyQueueIncreasing</th><th>ReplayQueueLength</th><th>LogReplayQueueIncreasing</th><th>ContentIndexState</th><th>ContentIndexErrorMessage</th></tr>'

    foreach ($mdbCopy in $mdbCopyStatus) {

        $mbody += "<tr><td>$($mdbCopy.Name)</td>"

        #Status
        if ($mdbCopy.Status -eq 'Mounted' -or $mdbCopy.Status -eq 'Healthy') {
            $mbody += "<td class = ""good"">$($mdbCopy.Status)</td>"
        }
        else {
            $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - Status is [$($mdbCopy.Status)]</td></tr>"
            $mbody += "<td class = ""bad"">$($mdbCopy.Status)</td>"
        }
        #CopyQueueLength
        if ($mdbCopy.CopyQueueLength -ge $t_copyQueue) {
            $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - CopyQueueLength [$($mdbCopy.CopyQueueLength)] is >= $($t_copyQueue)</td></tr>"
            $mbody += "<td class = ""bad"">$($mdbCopy.CopyQueueLength)</td>"
        }
        else {
            $mbody += "<td class = ""good"">$($mdbCopy.CopyQueueLength)</td>"
        }
        #LogCopyQueueIncreasing
        if ($mdbCopy.LogCopyQueueIncreasing -eq $true) {
            $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - LogCopyQueueIncreasing</tr>"
            $mbody += "<td class = ""bad"">$($mdbCopy.LogCopyQueueIncreasing)</td>"
        }
        else {
            $mbody += "<td class = ""good"">$($mdbCopy.LogCopyQueueIncreasing)</td>"
        }
        #ReplayQueueLength
        if ($mdbCopy.ReplayQueueLength -ge $t_replayQueue) {
            $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - ReplayQueueLength [$($mdbCopy.CopyQueueLength)] is >= $($t_replayQueue)</td></tr>"
            $mbody += "<td class = ""bad"">$($mdbCopy.ReplayQueueLength)</td>"
        }
        else {
            $mbody += "<td class = ""good"">$($mdbCopy.ReplayQueueLength)</td>"
        }
        #LogReplayQueueIncreasing
        if ($mdbCopy.LogReplayQueueIncreasing -eq $true) {
            $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - LogReplayQueueIncreasing</tr>"
            $mbody += "<td class = ""bad"">$($mdbCopy.LogReplayQueueIncreasing)</td>"
        }
        else {
            $mbody += "<td class = ""good"">$($mdbCopy.LogReplayQueueIncreasing)</td>"
        }
        #ContentIndexState
        if ($mdbCopy.ContentIndexState -eq "Healthy") {
            $mbody += "<td class = ""good"">$($mdbCopy.ContentIndexState)</td>"
        }
        elseif ($mdbCopy.ContentIndexState -eq "Disabled") {
            $mbody += "<td class = ""good"">$($mdbCopy.ContentIndexState)</td>"
        }
        elseif ($mdbCopy.ContentIndexState -eq "NotApplicable") {
            $mbody += "<td class = ""good"">$($mdbCopy.ContentIndexState)</td>"
        }
        else {
            $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - ContentIndexState is $($mdbCopy.ContentIndexState)</tr>"
            $mbody += "<td class = ""bad"">$($mdbCopy.ContentIndexState)</td>"
        }
        #ContentIndexErrorMessage
        $mbody += "<td class = ""bad"">$($mdbCopy.ContentIndexErrorMessage)</td>"
    }
    $mbody += '</tr>'
    if ($errString) { $mResult = "<tr><td>Mailbox Database Copy Status</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }
    # Write-Host "Done" -ForegroundColor Green
    return $mbody, $errString, $mResult, $testFailed
}

Function Get-ExServerComponents ($exServerList) {
    Write-Host (Get-Date) ': Server Component State... ' -ForegroundColor Yellow # -NoNewline
    foreach ($exServer in $exServerList) {
        #$stats_collection += (Get-ServerComponentState $exServer | Where-Object {$_.State -ne 'Active'} | Select-Object Identity,Component,State)
        $stats_collection += (Get-ServerComponentState $exServer | Where-Object { $_.Component -notin $IgnoreComponent } | Select-Object Identity, Component, State)
    }
    # Write-Host "Done" -ForegroundColor Green
    return $stats_collection
}

Function Get-QueueReport ($queueInfo) {
    Write-Host (Get-Date) ': Mail Queue Report... ' -ForegroundColor Yellow # -NoNewline
    $mResult = "<tr><td>Mail Queue</td><td class = ""good"">Passed</td></tr>"
    $testFailed = 0
    $mbody = @()
    $errString = @()
    $currentServer = ""
    $mbody += '<table id="SectionLabels"><tr><th class="data">Mail Queue</th></tr></table>'
    $mbody += '<table id="data">'

    foreach ($queue in $queueInfo) {
        $xq = $queue.Identity.ToString()
        $transportServer = $xq.split("\")
        if ($currentServer -ne $transportServer[0]) {
            $currentServer = $transportServer[0]
            $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Delivery Type</th><th>Status</th><th>Message Count</th><th>Next Hop Domain</th><th>Last Error</th></tr>'
        }

        if ($queue.MessageCount -ge $t_mQueue) {
            $errString += "<tr><td>Mail Queue</td></td><td>$($transportServer[0]) - $($queue.Identity) - Message Count is >= $($t_mQueue)</td></tr>"
            $mbody += "<tr><td>$($queue.Identity)</td><td>$($queue.DeliveryType)</td><td>$($queue.Status)</td><td class = ""bad"">$($queue.MessageCount)</td><td>$($queue.NextHopDomain)</td><td>$($queue.LastError)</td></tr>"
        }
        else {
            $mbody += "<tr><td>$($queue.Identity)</td><td>$($queue.DeliveryType)</td><td>$($queue.Status)</td><td>$($queue.MessageCount)</td><td>$($queue.NextHopDomain)</td><td>$($queue.LastError)</td></tr>"
        }

    }
    if ($errString) { $mResult = "<tr><td>Mail Queue</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }
    # Write-Host "Done" -ForegroundColor Green
    return $mbody, $errString, $mResult, $testFailed
}

Function Get-ReplicationReport ($replInfo) {
    Write-Host (Get-Date) ': Replication Health Report... ' -ForegroundColor Yellow # -NoNewline
    $mResult = "<tr><td>DAG Members Replication</td><td class = ""good"">Passed</td></tr>"
    $testFailed = 0
    $mbody = @()
    $errString = @()
    $currentServer = ""
    $mbody += '<table id="SectionLabels"><tr><th class="data">DAG Members Replication</th></tr></table>'
    $mbody += '<table id="data">'

    foreach ($repl in $replInfo) {
        if ($currentServer -ne $repl.Server) {
            $currentServer = $repl.Server
            $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Result</th><th>Error</th></tr>'
        }

        if ($repl.Result -match "Pass") {
            $mbody += "<tr><td>$($repl.Check)</td><td>$($repl.Result)</td><td>$($repl.Error)</td></tr>"
        }
        else {
            $errString += "<tr><td>Replication</td></td><td>$($currentServer) - $($repl.Check) is $($repl.Result) - $($repl.Error)</td></tr>"
            $mbody += "<tr><td>$($repl.Check)</td><td class = ""bad"">$($repl.Result)</td><td>$($repl.Error)</td></tr>"
        }
    }
    $mbody += ""



    if ($errString) { $mResult = "<tr><td>DAG Members Replication</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }
    # Write-Host "Done" -ForegroundColor Green
    return $mbody, $errString, $mResult, $testFailed
}

Function Get-ServerComponentStateReport ($serverComponentStateInfo) {

    Write-Host (Get-Date) ': Server Component State... ' -ForegroundColor Yellow # -NoNewline
    $mResult = "<tr><td>Server Component State</td><td class = ""good"">Passed</td></tr>"
    $testFailed = 0
    $mbody = @()
    $errString = @()
    $currentServer = ""
    $mbody += '<table id="SectionLabels"><tr><th class="data">Server Component State</th></tr></table>'
    $mbody += '<table id="data">'

    foreach ($componentInfo in $serverComponentStateInfo) {
        if ($currentServer -ne $componentInfo.Identity) {
            $currentServer = $componentInfo.Identity
            $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Component State</th></tr>'
        }

        [string]$componentName = $componentInfo.Component
        [string]$componentState = $componentInfo.State

        if ($componentState -ne 'Active') {
            $errString += "<tr><td>Component State</td></td><td>$($currentServer) - $($componentName) [$($componentState)]</td></tr>"
            $mbody += "<tr><td>$($componentName)</td><td class = ""bad"">$($componentState)</td></tr>"
        }
        else {
            $mbody += "<tr><td>$($componentName)</td><td class = ""good"">$($componentState)</td></tr>"
        }
    }
    if ($errString) { $mResult = "<tr><td>Server Component State</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }
    # Write-Host "Done" -ForegroundColor Green
    return $mbody, $errString, $mResult, $testFailed
}

Function Get-DiskReport ($diskinfo) {
    Write-Host (Get-Date) ': Disk Space Report... ' -ForegroundColor Yellow # -NoNewline
    $mResult = "<tr><td>Disk Space</td><td class = ""good"">Passed</td></tr>"
    $testFailed = 0
    $mbody = @()
    $errString = @()
    $currentServer = ""
    $mbody += '<table id="SectionLabels"><tr><th class="data">Disk Space</th></tr></table>'
    $mbody += '<table id="data">'
    foreach ($diskdata in $diskinfo) {
        if ($currentServer -ne $diskdata.SystemName) {
            $currentServer = $diskdata.SystemName
            $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Size (GB)</th><th>Free (GB)</th><th>Free (%)</th></tr>'
        }

        if ($diskdata.PercentFree -eq 20000) {
            $errString += "<tr><td>Disk</td></td><td>$($currentServer) - Error Fetching Data</td></tr>"
            $mbody += "<tr><td>$($diskdata.DeviceID) [$($diskdata.VolumeName)] </td><td>$($diskdata.Size)</td><td>$($diskdata.FreeSpace)</td><td class = ""bad"">Error Fetching Data</td></tr>"
        }
        elseif ($diskdata.PercentFree -ge $t_DiskBadPercent) {
            $mbody += "<tr><td>$($diskdata.DeviceID) [$($diskdata.VolumeName)] </td><td>$($diskdata.Size)</td><td>$($diskdata.FreeSpace)</td><td class = ""good"">$($diskdata.PercentFree)</td></tr>"
        }
        else {
            $errString += "<tr><td>Disk</td></td><td>$($currentServer) - $($diskdata.DeviceID) [$($diskdata.VolumeName)] [$($diskdata.FreeSpace) GB / $($diskdata.PercentFree)%] is <= $($t_DiskBadPercent)% Free</td></tr>"
            $mbody += "<tr><td>$($diskdata.DeviceID) [$($diskdata.VolumeName)] </td><td>$($diskdata.Size)</td><td>$($diskdata.FreeSpace)</td><td class = ""bad"">$($diskdata.PercentFree)</td></tr>"
        }

    }
    if ($errString) { $mResult = "<tr><td>Disk Space </td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }
    # Write-Host "Done" -ForegroundColor Green
    return $mbody, $errString, $mResult, $testFailed
}

Function Get-MdbReport ($dblist) {
    Write-Host (Get-Date) ': Mailbox Database Report... ' -ForegroundColor Yellow # -NoNewline
    $mResult = "<tr><td>Mailbox Database Status</td><td class = ""good"">Passed</td></tr>"
    $testFailed = 0
    $mbody = @()
    $errString = @()
    $mbody += '<table id="SectionLabels"><tr><th class="data">Mailbox Database Status</th></tr></table>'
    $mbody += '<table id="data"><tr><th>[Name][EDB Path][Log Path]</th><th>Mounted</th><th>On Server [Preference]</th><th>EDB Disk Size [Free] <br /> Log Disk Size [Free]</th><th>Size (GB)</th><th>White Space (GB)</th><th>Active Mailbox</th><th>Disconnected Mailbox</th><th>Item Size (GB)</th><th>Deleted Items Size (GB)</th><th>Full Backup</th><th>Incremental Backup</th><th>Backup In Progress</th><th>Mapi Connectivity</th></tr>'
    ForEach ($db in $dblist) {
        #$dbDetails = Get-MailboxDatabase $db.Name
        if ($db.mounted -eq $true) {
            #Calculate backup age----------------------------------------------------------
            if ($db.LastFullBackup -ne '') {
                $lastfullbackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastFullBackup
                $lastfullbackupelapsed = New-TimeSpan -Start $db.LastFullBackup
            }
            Else {
                $lastfullbackupelapsed = ''
                $lastfullbackup = '[NO DATA]'
            }

            if ($db.LastIncrementalBackup -ne '') {
                $lastincrementalbackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastIncrementalBackup
                $lastincrementalbackupelapsed = New-TimeSpan -Start $db.LastIncrementalBackup
            }
            Else {
                $lastincrementalbackupelapsed = ''
                $lastincrementalbackup = '[NO DATA]'
            }

            if ($t_lastfullbackup -eq 0) {
                [int]$full_backup_age = -1
            }
            else {
                [int]$full_backup_age = $lastfullbackupelapsed.totaldays
            }

            if ($t_lastincrementalbackup -eq 0) {
                [int]$incremental_backup_age = -1
            }
            else {
                [int]$incremental_backup_age = $lastincrementalbackupelapsed.totaldays
            }
            #-------------------------------------------------------------------------------
            $mbody += '<tr>'
            $mbody += '<td>[' + $db.Name + ']<br />[' + $db.EdbFilePath + ']<br />[' + $db.LogFolderPath + ']</td>'
            if ($db.Mounted -eq $true) {
                $mbody += '<td class = "good">' + $db.Mounted + '</td>'
            }
            Else {
                $errString += "<tr><td>Database Mount</td></td><td>$($db.Name) - is NOT MOUNTED</td></tr>"
                $mbody += '<td class = "bad">' + $db.Mounted + '</td>'
            }

            if ($db.ActivationPreference.Value -eq 1) {
                $mbody += '<td class = "good">' + $db.MountedOnServer + ' [' + $db.ActivationPreference.value + ']' + '</td>'
            }
            Else {
                $errString += "<tr><td>Database Activation</td></td><td>$($db.Name) - is mounted on $($db.MountedOnServer) which is NOT the preferred active server</td></tr>"
                $mbody += '<td class = "bad">' + $db.MountedOnServer + ' [' + $db.ActivationPreference.value + ']' + '</td>'
            }

            $mbody += '<td>' + $db.EDBFreeSpace + '<br />' + $db.LogFreeSpace + '</td>'
            $mbody += '<td>' + $db.DatabaseSize + '</td><td>' + $db.AvailableNewMailboxSpace + '</td><td>' + $db.ActiveMailboxCount + '</td><td>' + $db.DisconnectedMailboxCount + '</td><td>' + $db.TotalItemSize + '</td><td>' + $db.TotalDeletedItemSize + '</td>'

            if ($full_backup_age -gt $t_lastfullbackup) {
                $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date [$lastfullbackup] is OLDER than $($t_lastfullbackup) Day(s)</td></tr>"
                $mbody += '<td class = "bad">' + $lastfullbackup + '</td>'
            }
            elseif ($lastfullbackup -eq '[NO DATA]' -and $t_lastfullbackup -ne 0) {
                $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date [$lastfullbackup] is OLDER than $($t_lastfullbackup) Day(s)</td></tr>"
                $mbody += '<td class = "bad">' + $lastfullbackup + '</td>'
            }
            Else {
                $mbody += '<td class = "good">' + $lastfullbackup + '</td>'
            }

            if ($incremental_backup_age -gt $t_lastincrementalbackup) {
                $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last incremental backup date [$lastincrementalbackup] is OLDER than $($t_lastincrementalbackup) Day(s)</td></tr>"
                $mbody += '<td class = "bad">' + $lastincrementalbackup + '</td>'
            }
            Else {
                $mbody += '<td class = "good"> ' + $lastincrementalbackup + '</td>'
            }

            $mbody += '</td><td>' + $db.BackupInProgress + '</td>'

            if ($db.MapiConnectivity.Result.Value -eq 'Success') {
                $mbody += '<td class = "good"> ' + $db.MapiConnectivity.Result.Value + '</td>'
            }
            else {
                $errString += "<tr><td>MAPI Connectivity</td></td><td>$($db.Name) - MAPI Connectivity Result is $($db.MapiConnectivity.Result.Value)</td></tr>"
                $mbody += '<td class = "bad"> ' + $db.MapiConnectivity.Result.Value + '</td>'
            }
        }
        else {
            $errString += "<tr><td>Mailbox Datababase</td></td><td>$($db.Name) is DISMOUNTED</td></tr>"
            $mbody += "<tr><td>$($db.Name)</td><td class = ""bad"">$($db.Mounted)</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td></tr>"
        }
        $mbody += '</tr>'
    }
    if ($errString) { $mResult = "<tr><td>Mailbox Database Status</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }
    # Write-Host "Done" -ForegroundColor Green
    return $mbody, $errString, $mResult, $testFailed
}

Function Get-PdbReport ($dblist) {
    Write-Host (Get-Date) ': Public Folder Database Report... ' -ForegroundColor Yellow # -NoNewline
    $mResult = "<tr><td>Public Folder Database Status</td><td class = ""good"">Passed</td></tr>"
    $testFailed = 0
    $mbody += '<table id="SectionLabels"><tr><th class="data">Public Folder Database</th></tr></table>'
    $mbody += '<table id="data"><tr><th>Name</th><th>Mounted</th><th>On Server</th><th>Size (GB)</th><th>White Space (GB)</th><th>Full Backup</th><th>Incremental Backup</th><th>Backup In Progress</th><th>MAPI Connectivity</th></tr>'
    ForEach ($db in $dblist) {
        if ($db.Mounted -eq $true) {
            #Calculate backup age----------------------------------------------------------
            if ($db.LastFullBackup -ne '') {
                $lastfullbackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastFullBackup
                $lastfullbackupelapsed = New-TimeSpan -Start $db.LastFullBackup
            }
            Else {
                $lastfullbackupelapsed = ''
                $lastfullbackup = '[NO DATA]'
            }

            if ($db.LastIncrementalBackup -ne '') {
                $lastincrementalbackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastIncrementalBackup
                $lastincrementalbackupelapsed = New-TimeSpan -Start $db.LastIncrementalBackup
            }
            Else {
                $lastincrementalbackupelapsed = ''
                $lastincrementalbackup = '[NO DATA]'
            }
            [int]$full_backup_age = $lastfullbackupelapsed.totaldays
            [int]$incremental_backup_age = $lastincrementalbackupelapsed.totaldays
            #-------------------------------------------------------------------------------
            $mbody += '<tr>'
            $mbody += '<td>' + $db.Name + '</td>'
            if ($db.Mounted -eq $true) {
                $mbody += '<td class = "good">' + $db.Mounted + '</td>'
            }
            Else {
                $errString += "<tr><td>Database Mount</td></td><td>$($db.Name) - is NOT MOUNTED</td></tr>"
                $mbody += '<td class = "bad">' + $db.Mounted + '</td>'
            }

            $mbody += '<td>' + $db.MountedOnServer + '</td><td>' + $db.DatabaseSize + '</td><td>' + $db.AvailableNewMailboxSpace + '</td>'

            if ($full_backup_age -gt $t_lastfullbackup) {
                $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date [$lastfullbackup] is OLDER than $($t_lastfullbackup) days</td></tr>"
                $mbody += '<td class = "bad">' + $lastfullbackup + '</td>'
            }
            Else {
                $mbody += '<td class = "good">' + $lastfullbackup + '</td>'
            }

            if ($incremental_backup_age -gt $t_lastincrementalbackup) {
                $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last incremental backup date [$lastfullbackup] is OLDER than $($t_lastincrementalbackup) days</td></tr>"
                $mbody += '<td class = "bad">' + $lastincrementalbackup + '</td>'
            }
            Else {
                $mbody += '<td class = "good"> ' + $lastincrementalbackup + '</td>'
            }
            $mbody += '</td><td>' + $db.BackupInProgress + '</td>'

            if ($db.MapiConnectivity.Result.Value -eq 'Success') {
                $mbody += '<td class = "good"> ' + $db.MapiConnectivity.Result.Value + '</td>'
            }
            else {
                $mbody += '<td class = "bad"> ' + $db.MapiConnectivity.Result.Value + '</td>'
            }
        }
        else {
            $errString += "<tr><td>Public Folder Datababase</td></td><td>$($db.Name) is DISMOUNTED</td></tr>"
            $mbody += "<tr><td>$($db.Name)</td><td class = ""bad"">$($db.Mounted)</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td></tr>"
        }

        $mbody += '</tr>'
    }
    if ($errString) { $mResult = "<tr><td>Public Folder Database Status</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }
    # Write-Host "Done" -ForegroundColor Green
    return $mbody, $errString, $mResult, $testFailed
}

Function Get-CPUAndMemoryReport ($cpuAndMemDataResult) {
    Write-Host (Get-Date) ': CPU and Memory Report... ' -ForegroundColor Yellow # -NoNewline
    $mResult = "<tr><td>CPU and Memory Usage</td><td class = ""good"">Passed</td></tr>"
    $mbody = @()
    $errString = @()
    $testFailed = 0
    $currentServer = ""
    $mbody += '<table id="SectionLabels"><tr><th class="data">CPU and Memory Load</th></tr></table>'
    $mbody += '<table id="data">'

    foreach ($cpuAndMemData in $cpuAndMemDataResult) {
        $Top_CPU_Consumers = $cpuAndMemData.Top_CPU_Consumers -replace "`n", "<br />"
        $Top_Memory_Consumers = $cpuAndMemData.Top_Memory_Consumers -replace "`n", "<br />"

        if ($currentServer -ne $cpuAndMemData.Server) {
            $currentServer = $cpuAndMemData.Server
            $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>CPU Load</th><th>CPU Top Processes</th><th>Memory Load</th><th>Memory Top Processes</th></tr>'
        }

        if ([int]$cpuAndMemData.CPU_Usage -lt $t_cpuUsage) {
            $mbody += "<tr><td></td><td class = ""good"">$($cpuAndMemData.CPU_Usage)%</td><td>$($Top_CPU_Consumers)</td>"
        }
        elseif ([int]$cpuAndMemData.CPU_Usage -ge $t_cpuUsage) {
            $mbody += "<tr><td></td><td class = ""bad"">$($cpuAndMemData.CPU_Usage)%</td><td>$($Top_CPU_Consumers)</td>"
            $errString += "<tr><td>CPU</td></td><td>$($currentServer) - $($cpuAndMemData.CPU_Usage)% CPU Load IS OVER the $($t_cpuUsage)% threshold </td></tr>"
        }


        if ([int]$cpuAndMemData.Memory_Used_Percent -lt $t_ramUsage) {
            $mbody += "<td class = ""good"">$($cpuAndMemData.Memory_Used_Percent)%</td><td>$($Top_Memory_Consumers)</td></tr>"
        }
        elseif ([int]$cpuAndMemData.Memory_Used_Percent -ge $t_ramUsage) {
            $errString += "<td>Memory</td></td><td>$($currentServer) - $($cpuAndMemData.Memory_Used_Percent)% Memory Load IS OVER the $($t_ramUsage)% threshold </td></tr>"
            $mbody += "<td class = ""bad"">$($cpuAndMemData.Memory_Used_Percent)%</td><td>$($Top_Memory_Consumers)</td></tr>"
        }
    }

    if ($errString) { $mResult = "<tr><td>CPU and Memory Usage</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }
    # Write-Host "Done" -ForegroundColor Green
    return $mbody, $errString, $mResult, $testFailed
}

#SCRIPT BEGIN---------------------------------------------------------------

#Get-List of Exchange Servers and assign to array----------------------------
Write-Host (Get-Date) ': Building List of Servers - excluding Edge' -ForegroundColor Yellow
$temp_ExServerList = Get-ExchangeServer | Where-Object { $_.ServerRole -notmatch 'Edge' } | Sort-Object Name
$dagMemberCount = Get-MailboxServer | Where-Object { $_.DatabaseAvailabilityGroup }
if (!$dagMemberCount) { $dagMemberCount = @() }

#Get rid of excluded Servers
$ExServerList = @()
foreach ($ExServer in $temp_ExServerList) {
    if ($IgnoreServer -notcontains $ExServer.Name) {
        $exServerList += $ExServer
    }
}
$nonEx2010 = $ExServerList | Where-Object { $_.AdminDisplayVersion -notlike "Version 14*" }
$nonEx2010transportServers = @()
$nonEx2010transportServers += $ExServerList | Where-Object { $_.AdminDisplayVersion -notlike "Version 14*" -and $_.ServerRole -match 'Mailbox' }
$Ex2010TransportServers = @()
$Ex2010TransportServers += $ExServerList | Where-Object { $_.AdminDisplayVersion -like "Version 14*" -and $_.ServerRole -match 'HubTransport' }
$transportServers = @()
$transportServers += $nonEx2010transportServers + $Ex2010TransportServers
#----------------------------------------------------------------------------
#Get-List of Mailbox Database and assign to array----------------------------
if ($RunMdbReport -eq $true -OR $RunDBCopyReport -eq $true) {
    Write-Host (Get-Date) ': Building List of Mailbox Database' -ForegroundColor Yellow
    $temp_ExMailboxDBList = Get-MailboxDatabase -Status | Where-Object { $_.Recovery -eq $False }
    #Get rid of excluded Mailbox Database
    $ExMailboxDBList = @()
    $activeServers = @()
    foreach ($ExMailboxDB in $temp_ExMailboxDBList) {
        if ($IgnoreDatabase -notcontains $ExMailboxDB.Name) {
            $ExMailboxDBList += $ExMailboxDB
            $activeServers += ($ExMailboxDB.MountedOnServer).Split(".")[0]
        }
    }
    $activeServers = $activeServers | Select-Object -Unique
}
#----------------------------------------------------------------------------
#Get-List of Public Folder Database and assign to array----------------------
if ($RunPdbReport -eq $true) {
    Write-Host (Get-Date) ': Building List of Public Folder Database' -ForegroundColor Yellow
    $temp_ExPFDBList = Get-PublicFolderDatabase -Status | Where-Object { $_.Recovery -eq $False }
    if (!$temp_ExPFDBList) { $temp_ExPFDBList = @() }
    $ExPFDBList = @()

    #Get rid of excluded PF Database
    foreach ($ExPFDB in $temp_ExPFDBList) {
        if ($IgnorePFDatabase -notcontains $ExPFDB.Name) {
            $ExPFDBList += $ExPFDB
        }
    }

}
#----------------------------------------------------------------------------

#Begin Data Extraction-------------------------------------------------------
Write-Host '==================================================================' -ForegroundColor Green
Write-Host (Get-Date) ': Begin Data Extraction' -ForegroundColor Yellow
if ($RunCPUandMemoryReport -eq $true) { $cpuHealthData = Get-CPUAndMemoryLoad($ExServerList) ; $testCount++ }
if ($RunServerHealthReport -eq $true) { $serverhealthdata = Get-ServerHealth($ExServerList) ; $testCount++ }
if ($RunComponentReport -eq $true -AND $nonEx2010.count -gt 0) { $componentHealthData = Get-ExServerComponents ($nonEx2010) ; $testCount++ }
if ($RunMdbReport -eq $true) { $mdbdata = Get-MdbStatistics ($ExMailboxDBList) | Sort-Object Name ; $testCount++ }
if ($RunPdbReport -eq $true -AND $ExPFDBList.Count -gt 0) { $pdbdata = Get-PdbStatistics ($ExPFDBList) ; $testCount++ }
if ($RunDBCopyReport -eq $true) { $dagCopyData = Get-DatabaseCopyStatus ($ExMailboxDBList) ; $testCount++ }
if ($RunDAGReplicationReport -eq $true -and $dagMemberCount.count -gt 0) { $repldata = Get-ReplicationHealth ; $testCount++ }
if ($RunQueueReport -eq $true) { $queueData = Get-MailQueues ($transportServers) ; $testCount++ }
if ($RunDiskReport -eq $true) { $diskdata = Get-DiskSpaceStatistics($ExServerList) ; $testCount++ }

#----------------------------------------------------------------------------
# Build Report --------------------------------------------------------------
Write-Host '==================================================================' -ForegroundColor Green
Write-Host (Get-Date) ': Create Report' -ForegroundColor Yellow
if ($RunCPUandMemoryReport -eq $true) {
    $cpuAndMemoryCheckResult, $cpuError, $cpuResult, $cpuFailed = Get-CPUAndMemoryReport ($cpuHealthData)
    $errSummary += $cpuError
    $testFailed += $cpuFailed
}
if ($RunServerHealthReport -eq $true) { $serverhealthreport, $sError, $sResult, $sFailed = Get-ServerHealthReport ($serverhealthdata) ; $errSummary += $sError; $testFailed += $sFailed }
if ($RunComponentReport -eq $true -AND $nonEx2010.count -gt 0) { $componentHealthReport, $cError, $cResult, $cFailed = Get-ServerComponentStateReport ($componentHealthData) ; $errSummary += $cError; $testFailed += $cFailed }
if ($RunMdbReport -eq $true) { $mdbreport, $mError, $mdbResult, $mdbFailed = Get-MdbReport ($mdbdata) ; $errSummary += $mError; $testFailed += $mdbFailed }
if ($RunDBCopyReport -eq $true) { $dbcopyreport, $dbCopyError, $dbResult, $dbFailed = Get-DAGCopyStatusReport ($dagCopyData) ; $errSummary += $dbCopyError; $testFailed += $dbFailed }
if ($RunDAGReplicationReport -eq $true -and $dagMemberCount.count -gt 0) { $replicationreport, $rError, $rResult, $rFailed = Get-ReplicationReport ($repldata) ; $errSummary += $rError; $testFailed += $rFailed }
if ($RunPdbReport -eq $true -AND $ExPFDBList.Count -gt 0) { $pdbreport, $pdbError, $pdbResult, $pdbFailed = Get-PdbReport ($pdbdata) ; $errSummary += $pdbError; $testFailed += $pdbFailed }
if ($RunQueueReport -eq $true) { $queuereport, $qError, $qResult, $qFailed = Get-QueueReport($queueData) ; $errSummary += $qError; $testFailed += $qFailed }
if ($RunDiskReport -eq $true) { $diskreport, $dError, $dResult, $dFailed = Get-DiskReport ($diskdata) ; $errSummary += $dError; $testFailed += $dFailed }



$mail_body = "<html><head><title>[$($CompanyName)] $($MailSubject) $($today)</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
Write-Host (Get-Date) ': Formatting Report' -ForegroundColor Yellow
$mail_body += $css_string
$mail_body += '<table id="HeadingInfo">'
$mail_body += '<tr><th>' + $CompanyName + '<br />' + $MailSubject + '<br />' + $today + '</th></tr>'
$mail_body += '</table>'

##Set Individual Test Results
$testPassed = $testCount - $testFailed
$percentPassed = ($testPassed / $testCount) * 100
$percentPassed = [math]::Round($percentPassed)
if ($testPassed -lt $testCount) { $overAllResult = "FAILED" }

$mail_body += '<table id="SectionLabels">'
$mail_body += "<tr><th class=""data"">Overall Health: $($percentPassed)% - $($overAllResult)</th></tr></table>"
$mail_body += '<table id="data"><tr><th>Test</th><th>Result</th></tr>'
if ($RunCPUandMemoryReport -eq $true) { $mail_body += $cpuResult }
if ($RunServerHealthReport -eq $true) { $mail_body += $sResult }
if ($RunComponentReport -eq $true -AND $nonEx2010.count -gt 0) { $mail_body += $cResult }
if ($RunMdbReport -eq $true) { $mail_body += $mdbResult }
if ($RunDBCopyReport -eq $true) { $mail_body += $dbResult }
if ($RunDAGReplicationReport -eq $true -and $dagMemberCount.count -gt 0) { $mail_body += $rResult }
if ($RunPdbReport -eq $true -AND $ExPFDBList.Count -gt 0) { $mail_body += $pdbResult }
if ($RunQueueReport -eq $true) { $mail_body += $qResult }
if ($RunDiskReport -eq $true) { $mail_body += $dResult }
$mail_body += '</table>'
if ($overAllResult -eq 'FAILED') {
    $mail_body += '<table id="SectionLabels">'
    $mail_body += '<tr><th class="data">Issues</th></tr></table>'
    $mail_body += '<table id="data"><tr><th>Check Item</th><th>Details</th></tr>'
    $mail_body += $errSummary
    $mail_body += '</table>'
}

if ($RunCPUandMemoryReport -eq $true) { $mail_body += $cpuAndMemoryCheckResult ; $mail_body += '</table>' }
if ($RunServerHealthReport -eq $true) { $mail_body += $serverhealthreport ; $mail_body += '</table>' }
if ($RunComponentReport -eq $true -AND $nonEx2010.count -gt 0) { $mail_body += $componentHealthReport ; $mail_body += '</table>' }
if ($RunMdbReport -eq $true) { $mail_body += $mdbreport ; $mail_body += '</table>' }
if ($RunDAGReplicationReport -eq $true) { $mail_body += $replicationreport ; $mail_body += '</table>' }
if ($RunDBCopyReport -eq $true) { $mail_body += $dbcopyreport ; $mail_body += '</table>' }
if ($RunPdbReport -eq $true) { $mail_body += $pdbreport ; $mail_body += '</table>' }
if ($RunQueueReport -eq $true) { $mail_body += $queuereport ; $mail_body += '</table>' }
if ($RunDiskReport -eq $true) { $mail_body += $diskreport ; $mail_body += '</table>' }
$mail_body += '<p><table id="SectionLabels">'
$mail_body += '<tr><th>----END of REPORT----</th></tr></table></p>'
$mail_body += '<p><font size="2" face="Tahoma"><u>Report Paremeters</u><br />'
$mail_body += '<b>[THRESHOLD]</b><br />'
$mail_body += 'Last Full Backup: ' + $t_lastfullbackup + ' Day(s)<br />'
$mail_body += 'Last Incremental Backup: ' + $t_lastincrementalbackup + ' Day(s)<br />'
$mail_body += 'Mail Queue: ' + $t_mQueue + '<br />'
$mail_body += 'Copy Queue: ' + $t_copyQueue + '<br />'
$mail_body += 'Replay Queue: ' + $t_replayQueue + '<br />'
$mail_body += 'Disk Space Critical: ' + $t_DiskBadPercent + ' (%) <br />'
$mail_body += 'CPU: ' + $t_cpuUsage + ' (%) <br />'
$mail_body += 'Memory: ' + $t_ramUsage + ' (%) <br /><br />'
$mail_body += '<b>[MAIL]</b><br />'
$mail_body += 'SMTP Server: ' + $MailServer + '<br /><br />'
$mail_body += '<b>[REPORT]</b><br />'
$mail_body += 'Generated from Server: ' + ($env:computername) + '<br />'
$mail_body += 'Script File: ' + $MyInvocation.MyCommand.Definition + '<br />'
$mail_body += 'Config File: ' + $configFile + '<br />'
$mail_body += 'Report File: ' + $reportfile + '<br />'
$mail_body += 'Recipients: ' + ($MailTo -join ';') + '<br /><br />'
$mail_body += '<b>[EXCLUSIONS]</b><br />'
$mail_body += 'Excluded Servers: ' + (@($config.exclusions.IgnoreServer) -join ';') + '<br />'
$mail_body += 'Excluded Components: ' + (@($config.exclusions.IgnoreComponent) -join ';') + '<br />'
$mail_body += 'Excluded Mailbox Database: ' + (@($config.exclusions.IgnoreDatabase) -join ';') + '<br />'
$mail_body += 'Excluded Public Database: ' + (@($config.exclusions.IgnorePFDatabase) -join ';') + '<br /><br />'
$mail_body += '</p><p>'
$mail_body += '<a href="' + $script_info.ProjectUri.OriginalString + '">' + $script_info.Name + ' ' + $script_info.Version.ToString() + '</a></p>'
$mail_body += '</html>'
# $mbody = $mbox -replace "&lt;", "<"
# $mbody = $mbox -replace "&gt;", ">"
$mail_body | Out-File $reportfile
Write-Host (Get-Date) ': HTML Report saved to file -' $reportfile -ForegroundColor Yellow
#----------------------------------------------------------------------------
# Mail Parameters------------------------------------------------------------
# Add CC= and/or BCC= lines if you want to add recipients for CC and BCC
$params = @{
    Body       = $mail_body
    BodyAsHtml = $true
    Subject    = "[$($CompanyName)] $($MailSubject) $($today)"
    From       = $MailSender
    SmtpServer = $MailServer
    UseSsl     = $config.mailAndReportParameters.SSLEnabled
    Port       = $config.mailAndReportParameters.Port
}

if ($MailTo) { $params.Add('To', $MailTo) }
if ($MailCc) { $params.Add('Cc', $MailCc) }
if ($MailBcc) { $params.Add('Bcc', $MailBcc) }

#----------------------------------------------------------------------------
# Send Report----------------------------------------------------------------
if ($SendReportViaEmail -eq $true) { Write-Host (Get-Date) ': Sending Report' -ForegroundColor Yellow ; Send-MailMessage @params }
#----------------------------------------------------------------------------
Write-Host ""
Write-Host "======================================"
Write-Host "Enabled Tests: $($testCount)"
Write-Host "Failed: $($testFailed)"
Write-Host "Passed: $($testPassed) [$($percentPassed)%]"
Write-Host "Overall Result: " -NoNewline
if ($overAllResult -eq "Failed") { Write-Host "$($overAllResult)" -ForegroundColor RED } else { Write-Host "$($overAllResult)" -ForegroundColor GREEN }
Write-Host "======================================"
Write-Host ""
Write-Host (Get-Date) ': End' -ForegroundColor Green
#SCRIPT END------------------------------------------------------------------
#https://www.lazyexchangeadmin.com/2015/03/database-backup-and-disk-space-report.html
if ($enableDebug) { Stop-Transcript }


