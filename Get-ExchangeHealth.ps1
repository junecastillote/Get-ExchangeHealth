#parameter bindings
[CmdletBinding()]
param (
	[Parameter(Mandatory = $true)]
	[string]$configFile,
	[Parameter(Mandatory = $false)]
	[switch]$enableDebug
)

#Import Configuration File
if ((Test-Path $configFile) -eq $false)
{
	Write-Host "ERROR: File $($configFile) does not exist. Script cannot continue" -ForegroundColor Yellow
	"ERROR: File $($configFile) does not exist. Script cannot continue" | Out-File error.txt
	EXIT
}
[xml]$config = gc $configFile

if ($Debug)
{
	$ErrorActionPreference="SilentlyContinue"
	$WarningPreference="SilentlyContinue"
}

#Start Script
$scriptVersion = "5.1"
Write-Host '=================================================' -ForegroundColor Yellow
Write-Host '              Get-ExchangeHealth		         ' -ForegroundColor Yellow
Write-Host '           june.castillote@gmail.com    	     ' -ForegroundColor Yellow
Write-Host '=================================================' -ForegroundColor Yellow
#http://shaking-off-the-cobwebs.blogspot.com/2015/03/database-backup-and-disk-space-report.html
Write-Host ''
Write-Host (Get-Date) ': Begin' -ForegroundColor Green
Write-Host (Get-Date) ': Setting Paths and Variables' -ForegroundColor Yellow

#Define Variables
$errSummary = ""
$today = '{0:dd-MMM-yyyy hh:mm tt}' -f (Get-Date)
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
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
<hr />
'@

#Thresholds from config
[int]$t_lastfullbackup = $config.configuration.thresholds.LastFullBackup
[int]$t_lastincrementalbackup = $config.configuration.thresholds.LastIncrementalBackup
[int]$t_DiskBadPercent = $config.configuration.thresholds.DiskSpaceFree
[int]$t_mQueue = $config.configuration.thresholds.MailQueueCount

#Options from config
$RunServerHealthReport = $config.configuration.reportOptions.RunServerHealthReport
$RunMdbReport = $config.configuration.reportOptions.RunMdbReport
$RunPdbReport = $config.configuration.reportOptions.RunPdbReport
$RunDAGCopyReport = $config.configuration.reportOptions.RunDAGCopyReport
$RunDAGReplicationReport = $config.configuration.reportOptions.RunDAGReplicationReport
$RunQueueReport = $config.configuration.reportOptions.RunQueueReport
$RunDiskReport = $config.configuration.reportOptions.RunDiskReport
$SendReportViaEmail = $config.configuration.reportOptions.SendReportViaEmail
$reportfile = $config.configuration.reportOptions.ReportFile

#Mail settings
$CompanyName = $config.configuration.mailAndReportParameters.CompanyName
$MailSubject = $config.configuration.mailAndReportParameters.MailSubject
$MailServer = $config.configuration.mailAndReportParameters.MailServer
$MailSender = $config.configuration.mailAndReportParameters.MailSender
$MailTo = $config.configuration.mailAndReportParameters.MailTo

#Import Exchange 2010 Shell Snap-In if not already added

	if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
	{
		try
		{
			Write-Host (Get-Date) ': Add Exchange Snap-in' -ForegroundColor Yellow
			Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
		}
		catch
		{
			Write-Warning $_.Exception.Message
			EXIT
		}
	}

#Ping function
Function Ping-Server ($server)
{
	$ping = Test-Connection $server -quiet -count 1
	return $ping
}

Function Get-MdbStatistics ($mailboxdblist){
Write-Host (Get-Date) ': Mailbox Database Check... ' -ForegroundColor Yellow -NoNewLine
	$stats_collection = @()
		foreach ($mailboxdb in $mailboxdblist)
		{
			if (Ping-Server($mailboxdb.Server.Name) -eq $true)
			{
				$mdbStat = Get-MailboxStatistics -Database $mailboxdb
				$mdbobj = "" | Select Name,Mounted,MountedOnServer,ActivationPreference,DatabaseSize,AvailableNewMailboxSpace,ActiveMailboxCount,DisconnectedMailboxCount,TotalItemSize,TotalDeletedItemSize,EdbFilePath,LogFolderPath,LogFilePrefix,LastFullBackup,LastIncrementalBackup,BackupInProgress,MapiConnectivity,EDBFreeSpace,LogFreeSpace
				$mdbobj.Name = $mailboxdb.name
				$mdbobj.EdbFilePath = $mailboxdb.EdbFilePath
				$mdbobj.LogFolderPath = $mailboxdb.LogFolderPath
				$mdbobj.Mounted = $mailboxdb.Mounted
				$mdbobj.MountedOnServer = $mailboxdb.Server.Name
				$mdbobj.ActivationPreference = $mailboxdb.ActivationPreference | ?{$_.Key -eq $mailboxdb.Server.Name}
				$mdbobj.LastFullBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $mailboxdb.LastFullBackup
				$mdbobj.LastIncrementalBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $mailboxdb.LastIncrementalBackup
				$mdbobj.BackupInProgress = $mailboxdb.BackupInProgress
				$mdbobj.DatabaseSize = "{0:N2}" -f ($mailboxdb.DatabaseSize.tobytes() / 1GB)
				#$mbxItemSize = Get-MailboxStatistics -Database $mailboxdb | %{$_.TotalItemSize.Value} | Measure-Object -sum
				$mbxItemSize = $mdbStat | %{$_.TotalItemSize.Value} | Measure-Object -sum
				#$mbxDelSize = Get-MailboxStatistics -Database $mailboxdb | %{$_.TotalDeletedItemSize.Value} | Measure-Object -sum
				$mbxDelSize = $mdbStat | %{$_.TotalDeletedItemSize.Value} | Measure-Object -sum
				$mdbobj.TotalItemSize = "{0:N2}" -f ($mbxItemSize.sum / 1GB)
				$mdbobj.TotalDeletedItemSize = "{0:N2}" -f ($mbxDelSize.sum / 1GB)
				$mdbobj.ActiveMailboxCount = ($mdbStat | where {$_.DisconnectDate -eq $null}).count
				$mdbobj.DisconnectedMailboxCount = ($mdbStat | where {$_.DisconnectDate -ne $null}).count
				$mdbobj.AvailableNewMailboxSpace = "{0:N2}" -f ($mailboxdb.AvailableNewMailboxSpace.tobytes() / 1GB)
				$mdbobj.MapiConnectivity = Test-MapiConnectivity -Database $mailboxdb.Identity -PerConnectionTimeout 10
				
				#Get Disk Details
				$dbDrive = (Get-WmiObject Win32_LogicalDisk -Computer $mailboxdb.Server.Name | ?{$_.DeviceID -eq $mailboxdb.EdbFilePath.DriveName})
				$logDrive = (Get-WmiObject Win32_LogicalDisk -Computer $mailboxdb.Server.Name | ?{$_.DeviceID -eq $mailboxdb.LogFolderPath.DriveName})
				
				$mdbobj.EDBFreeSpace = "{0:N2}" -f ($dbDrive.Size / 1GB) + '[' + "{0:N2}" -f ($dbDrive.FreeSpace / 1GB) + ']'
				$mdbobj.LogFreeSpace = "{0:N2}" -f ($logDrive.Size / 1GB) + '[' + "{0:N2}" -f ($logDrive.FreeSpace / 1GB) + ']'
			}
			else
			{
				$mdbobj = "" | Select Name,Mounted,MountedOnServer,ActivationPreference,DatabaseSize,AvailableNewMailboxSpace,ActiveMailboxCount,DisconnectedMailboxCount,TotalItemSize,TotalDeletedItemSize,EdbFilePath,LogFolderPath,LogFilePrefix,LastFullBackup,LastIncrementalBackup,BackupInProgress,MapiConnectivity,EDBFreeSpace,LogFreeSpace
				$mdbobj.Name = $mailboxdb.name
				$mdbobj.EdbFilePath = $mailboxdb.EdbFilePath
				$mdbobj.LogFolderPath = $mailboxdb.LogFolderPath
				$mdbobj.Mounted = "$($mailboxdb.Server.Name): Connection Failed"
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
		
			
			$stats_collection+=$mdbobj
		}
Write-Host 'Done' -ForegroundColor Green
return $stats_collection
}

Function Get-PdbStatistics ($pfdblist){
Write-Host (Get-Date) ': Public Folder Database Check... ' -ForegroundColor Yellow -NoNewLine
	$stats_collection = @()
		foreach ($pfdb in $pfdblist)
		{
			$pfdbobj = "" | Select Name,Mounted,MountedOnServer,DatabaseSize,AvailableNewMailboxSpace,FolderCount,TotalItemSize,LastFullBackup,LastIncrementalBackup,BackupInProgress,MapiConnectivity
			$pfdbobj.Name = $pfdb.Name
			$pfdbobj.Mounted = $pfdb.Mounted
			$pfdbobj.MountedOnServer = $pfdb.Server.Name
			$pfdbobj.DatabaseSize = "{0:N2}" -f ($pfdb.DatabaseSize.tobytes() / 1GB)
			$pfdbobj.AvailableNewMailboxSpace = "{0:N2}" -f ($pfdb.AvailableNewMailboxSpace.tobytes() / 1GB)
			$pfdbobj.LastFullBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $pfdb.LastFullBackup
			$pfdbobj.LastIncrementalBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $pfdb.LastIncrementalBackup
			$pfdbobj.BackupInProgress = $pfdb.BackupInProgress
			$pfdbobj.MapiConnectivity = Test-MapiConnectivity -Database $pfdb.Identity -PerConnectionTimeout 10
			$stats_collection+=$pfdbobj
		}
Write-Host "Done" -ForegroundColor Green
	return $stats_collection
}

Function Get-DiskSpaceStatistics ($serverlist) {
	Write-Host (Get-Date) ': Disk Space Check... ' -ForegroundColor Yellow -NoNewLine
	$stats_collection = @()
		foreach ($server in $serverlist)
		{
			try
				{
				$diskObj = Get-WmiObject Win32_LogicalDisk -Filter 'DriveType=3' -computer $server | Select-Object SystemName, DeviceID, VolumeName, Size, FreeSpace
				foreach ($disk in $diskObj)
					{
							$serverobj = "" | Select SystemName, DeviceID, VolumeName, Size, FreeSpace, PercentFree
							$serverobj.SystemName = $disk.SystemName
							$serverobj.DeviceID = $disk.DeviceID
							$serverobj.VolumeName = $disk.VolumeName
							$serverobj.Size = "{0:N2}" -f ($disk.Size / 1GB)
							$serverobj.FreeSpace = "{0:N2}" -f ($disk.FreeSpace / 1GB)
							[int]$serverobj.PercentFree = "{0:N0}" -f (($disk.freespace/$disk.size) * 100)
							$stats_collection+=$serverobj						
					}
				}
			catch
				{
							$serverobj = "" | Select SystemName, DeviceID, VolumeName, Size, FreeSpace, PercentFree
							$serverobj.SystemName = $server
							$serverobj.DeviceID = $disk.DeviceID
							$serverobj.VolumeName = $disk.VolumeName
							$serverobj.Size = 0
							$serverobj.FreeSpace = 0
							[int]$serverobj.PercentFree = 20000
							$stats_collection+=$serverobj	
				}
		}
	Write-Host "Done" -ForegroundColor Green
	return $stats_collection
}

Function Get-ReplicationHealth {
	Write-Host (Get-Date) ': Replication Health Check... ' -ForegroundColor Yellow -NoNewLine
	$stats_collection = Get-MailboxServer | ?{$_.DatabaseAvailabilityGroup -ne $null} | Sort-Object Name | %{Test-ReplicationHealth -Identity $_}
	Write-Host "Done" -ForegroundColor Green
	return $stats_collection
}

Function Get-MailQueues{
	Write-Host (Get-Date) ': Mail Queue Check... ' -ForegroundColor Yellow -NoNewLine
	$stats_collection = get-TransportServer | ?{$_.ServerRole -notmatch 'Edge'} | Sort-Object Name | %{Get-Queue -Server $_}
	Write-Host "Done" -ForegroundColor Green
	return $stats_collection
}

Function Get-ServerHealth ($serverlist) 
{
Write-Host (Get-Date) ': Server Status Check... ' -ForegroundColor Yellow -NoNewLine
	$stats_collection = @()
		foreach ($server in $serverlist)
		{
			if (Ping-Server($server.name) -eq $true)
				{
					$serverOS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server
				
					$serverobj = "" | Select Server,Connectivity,ADSite,UpTime,HubTransportRole,ClientAccessRole,MailboxRole,MailFlow
					$timespan = $serverOS.ConvertToDateTime($serverOS.LocalDateTime) - $serverOS.ConvertToDateTime($serverOS.LastBootUpTime)
					[int]$uptime = "{0:00}" -f $timespan.TotalHours			
					
					$serverobj.Server = $server.Name
					$serverobj.UpTime = $uptime
					$serverobj.Connectivity = "Passed"
					$serviceStatus = Test-ServiceHealth -Server $server
					$serverobj.HubTransportRole = ""
					$serverobj.ClientAccessRole = ""
					$serverobj.MailboxRole = ""
					$site = ($server.site.ToString()).Split("/")
					$serverObj.ADSite = $site[-1]
					foreach ($service in $serviceStatus)
						{
							if ($service.Role -eq 'Hub Transport Server Role')
								{
									if ($service.RequiredServicesRunning -eq $true)
										{
											$serverobj.HubTransportRole = "Passed"
										}
									else
										{
											$serverobj.HubTransportRole = "Failed"
										}
								}
								
							if ($service.Role -eq 'Client Access Server Role')
								{
									if ($service.RequiredServicesRunning -eq $true)
										{
											$serverobj.ClientAccessRole = "Passed"
										}
									else
										{
											$serverobj.ClientAccessRole = "Failed"
										}
								}
								
							if ($service.Role -eq 'Mailbox Server Role')
								{
									if ($service.RequiredServicesRunning -eq $true)
										{
											$serverobj.MailboxRole = "Passed"
										}
									else
										{
											$serverobj.MailboxRole = "Failed"
										}
								}
						}
						
					if ($server.serverrole -match 'Mailbox')
					{
						#Exchange 2013
						if ($server.AdminDisplayVersion -like 'Version 15*') 
						{
							$mailflowresult = $null
							$url = (Get-PowerShellVirtualDirectory -Server $server -AdPropertiesOnly | Where {$_.Name -eq "Powershell (Default Web Site)"}).InternalURL.AbsoluteUri
							$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $url -ErrorAction STOP
							
							try
							{
								$result = Invoke-Command -Session $session {Test-Mailflow} -ErrorAction STOP
								$mailflowresult = $result.TestMailflowResult
							}
							catch
							{
								$mailflowresult = "Fail"
							}
							$serverObj.MailFlow = $mailflowresult
						}
						#Exchange 2010
						elseif ($server.AdminDisplayVersion -like 'Version 14*')
						{
							$serverObj.MailFlow = Test-MailFlow $server.Name
						}
					}
				}
				else
				{
				$serverobj = "" | Select Server,Connectivity,ADSite,UpTime,HubTransportRole,ClientAccessRole,MailboxRole
				
				$site = ($server.site.ToString()).Split("/")
				$serverObj.ADSite = $site[-1]
				$serverobj.Server = $server.Name
				$serverobj.Connectivity = "Failed"
				$serverobj.UpTime = "Cannot retrieve up time"
				$serverobj.HubTransportRole = "Failed"
				$serverobj.ClientAccessRole = "Failed"
				$serverobj.MailboxRole = "Failed"
				$serverObj.MailFlow = "Failed"
				}
			$stats_collection += $serverobj
		}
	Write-Host "Done" -ForegroundColor Green
	return $stats_collection
}

Function Get-DAGCopyStatus ($mailboxdblist) {
Write-Host (Get-Date) ': DAG Copy Status Check... ' -ForegroundColor Yellow -NoNewLine
	$stats_collection = @()
	
		foreach ($db in $mailboxdblist)
		{
			if ($db.MasterType -eq 'DatabaseAvailabilityGroup')
			{				
				foreach ($dbCopy in $db.DatabaseCopies)
				{
					$temp = "" | Select Name,Status,CopyQueueLength,LogCopyQueueIncreasing,ReplayQueueLength,LogReplayQueueIncreasing,ContentIndexState,ContentIndexErrorMessage
					$dbStatus = Get-MailboxDatabaseCopyStatus -Identity $dbCopy
					$temp.Name = $dbStatus.Name
					$temp.Status = $dbStatus.Status
					$temp.CopyQueueLength = $dbStatus.CopyQueueLength
					$temp.LogCopyQueueIncreasing = $dbStatus.LogCopyQueueIncreasing
					$temp.ReplayQueueLength = $dbStatus.ReplayQueueLength
					$temp.LogReplayQueueIncreasing = $dbStatus.LogReplayQueueIncreasing
					$temp.ContentIndexState = $dbStatus.ContentIndexState
					$temp.ContentIndexErrorMessage = $dbStatus.ContentIndexErrorMessage
					$stats_collection += $temp
				}				
			}		
		}
	Write-Host "Done" -ForegroundColor Green
	return $stats_collection | sort-object Name
}


Function Create-DAGCopyStatusReport ($mdbCopyStatus) {
	Write-Host (Get-Date) ': DAG Copy Status... ' -ForegroundColor Yellow -NoNewLine
	$mbody = @()
	$errString = @()
	$mbody += '<table id="SectionLabels"><tr><th class="data">Mailbox Database Copy Status</th></tr></table>'
	$mbody += '<table id="data">'
	$mbody += '<tr><th>Name</th><th>Status</th><th>CopyQueueLength</th><th>LogCopyQueueIncreasing</th><th>ReplayQueueLength</th><th>LogReplayQueueIncreasing</th><th>ContentIndexState</th><th>ContentIndexErrorMessage</th></tr>'
	
	foreach ($mdbCopy in $mdbCopyStatus)
	{
		
		$mbody += "<tr><td>$($mdbCopy.Name)</td>"
		
		#Status
		if ($mdbCopy.Status -eq 'Mounted' -or $mdbCopy.Status -eq 'Healthy')
			{
				$mbody += "<td class = ""good"">$($mdbCopy.Status)</td>"
			}
		else
			{
				$errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - Status is [$($mdbCopy.Status)]</td></tr>"
				$mbody += "<td class = ""bad"">$($mdbCopy.Status)</td>"
			}
		#CopyQueueLength
		if ($mdbCopy.CopyQueueLength -ge 5)
			{
				$errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - CopyQueueLength [$($mdbCopy.CopyQueueLength)] is >= 5</td></tr>"
				$mbody += "<td class = ""bad"">$($mdbCopy.CopyQueueLength)</td>"
			}
		else
			{
				$mbody += "<td class = ""good"">$($mdbCopy.CopyQueueLength)</td>"
			}
		#LogCopyQueueIncreasing
		if ($mdbCopy.LogCopyQueueIncreasing -eq $true)
			{
				$errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - LogCopyQueueIncreasing</tr>"
				$mbody += "<td class = ""bad"">$($mdbCopy.LogCopyQueueIncreasing)</td>"
			}
		else
			{
				$mbody += "<td class = ""good"">$($mdbCopy.LogCopyQueueIncreasing)</td>"				
			}
		#ReplayQueueLength
		if ($mdbCopy.ReplayQueueLength -ge 5)
			{
				$errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - ReplayQueueLength [$($mdbCopy.CopyQueueLength)] is >= 5</td></tr>"
				$mbody += "<td class = ""bad"">$($mdbCopy.ReplayQueueLength)</td>"
			}
		else
			{
				$mbody += "<td class = ""good"">$($mdbCopy.ReplayQueueLength)</td>"
			}		
		#LogReplayQueueIncreasing
		if ($mdbCopy.LogReplayQueueIncreasing -eq $true)
			{
				$errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - LogReplayQueueIncreasing</tr>"
				$mbody += "<td class = ""bad"">$($mdbCopy.LogReplayQueueIncreasing)</td>"
			}
		else
			{
				$mbody += "<td class = ""good"">$($mdbCopy.LogReplayQueueIncreasing)</td>"				
			}
		#ContentIndexState
		if ($mdbCopy.ContentIndexState -ne "Healthy")
			{
				$errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - ContentIndexState is $($mdbCopy.ContentIndexState)</tr>"
				$mbody += "<td class = ""bad"">$($mdbCopy.ContentIndexState)</td>"
			}
		else
			{
				$mbody += "<td class = ""good"">$($mdbCopy.ContentIndexState)</td>"				
			}
		#ContentIndexErrorMessage
		$mbody += "<td class = ""bad"">$($mdbCopy.ContentIndexErrorMessage)</td>"
	}
	$mbody += '</tr>'

Write-Host "Done" -ForegroundColor Green
return $mbody,$errString
}

Function Create-ServerHealthReport ($serverhealthinfo) {
	Write-Host (Get-Date) ': Server Health Report... ' -ForegroundColor Yellow -NoNewLine
	$mbody = @()
	$errString = @()
	$currentServer = ""
	$mbody += '<table id="SectionLabels"><tr><th class="data">Server Health Status</th></tr></table>'
	$mbody += '<table id="data">'
	$mbody += '<tr><th>Server</th><th>Site</th><th>Connectivity</th><th>Up Time (Hours)</th><th>Hub Transport Role</th><th>Client Access Role</th><th>Mailbox Role</th><th>Mail Flow</th></tr>'
	foreach ($server in $serverhealthinfo)
	{
		$mbody += "<tr><td>$($server.server)</td><td>$($server.ADSite)</td>"
		#Uptime
		if ($server.UpTime -lt 24)
			{
				$errString += "<tr><td>Server Up Time</td></td><td>$($server.server) - up time [$($server.Uptime)] is less than 24 hours</td></tr>"
				$mbody += "<td class = ""good"">$($server.Connectivity)</td><td class = ""bad"">$($server.UpTime)</td>"
			}
		elseif ($server.Uptime -eq 'Cannot retrieve up time')
			{
				$errString += "<tr><td>Server Connectivity</td></td><td>$($server.server) - connection test failed. SERVER MIGHT BE DOWN!!!</td></tr>"
				$mbody += "<td class = ""bad"">$($server.Connectivity)</td><td class = ""bad"">$($server.UpTime)</td>"
			}
		else
			{
				$mbody += "<td class = ""good"">$($server.Connectivity)</td><td class = ""good"">$($server.UpTime)</td>"
			}
		#Transport Role
		if ($server.HubTransportRole -eq 'Passed')
			{
				$mbody += '<td class = "good">Passed</td>'
			}
		elseif ($server.HubTransportRole -eq 'Failed')
			{
				$errString += "<tr><td>Role Services</td></td><td>$($server.server) - not all required Hub Transport Role services are running</td></tr>"
				$mbody += '<td class = "bad">Failed</td>'
			}
		else
			{
				$mbody += '<td class = "good"></td>'
			}
		#CAS Role
		if ($server.ClientAccessRole -eq 'Passed')
			{
				$mbody += '<td class = "good">Passed</td>'
			}
		elseif ($server.ClientAccessRole -eq 'Failed')
			{
				$errString += "<tr><td>Role Services</td></td><td>$($server.server) - not all required Client Access Role services are running</td></tr>"
				$mbody += '<td class = "bad">Failed</td>'
			}
		else
			{
				$mbody += '<td class = "good"></td>'
			}
		#Mailbox Role
		if ($server.MailboxRole -eq 'Passed')
			{
				$mbody += '<td class = "good">Passed</td>'
			}
		elseif ($server.MailboxRole -eq 'Failed')
			{
				$errString += "<tr><td>Role Services</td></td><td>$($server.server) - not all required Mailbox Role services are running</td></tr>"
				$mbody += '<td class = "bad">Failed</td>'
			}
		else
			{
				$mbody += '<td class = "good"></td>'
			}
			
		#Mail Flow
		if ($server.MailFlow -eq "Failed")
			{
				$errString += "<tr><td>Mail Flow</td></td><td>$($db.Name) - Mail Flow Result FAILED</td></tr>"
				$mbody += '<td class = "bad">Failed</td>'
			}
		elseif ($server.MailFlow = 'Success')
			{
				$mbody += '<td class = "good">Success</td>'
			}
		elseif ($server.MailFlow -eq "" -or $server.MailFlow -eq $null)
			{
				$mbody += '<td class = "good"></td>'
			}
			
		$mbody += '</tr>'
	}
Write-Host "Done" -ForegroundColor Green
return $mbody,$errString
}

Function Create-QueueReport ($queueInfo){
	Write-Host (Get-Date) ': Mail Queue Report... ' -ForegroundColor Yellow -NoNewLine
	$mbody = @()
	$errString = @()
	$currentServer = ""
	$mbody += '<table id="SectionLabels"><tr><th class="data">Mail Queue</th></tr></table>'
	$mbody += '<table id="data">'
	
	foreach ($queue in $queueInfo)
	{
		$xq = $queue.Identity.ToString()
		$transportServer = $xq.split("\")
		if ($currentServer -ne $transportServer[0])
		{
			$currentServer = $transportServer[0]
			$mbody += '<tr><th><b><u>'+ $currentServer + '</b></u></th><th>Delivery Type</th><th>Status</th><th>Message Count</th><th>Next Hop Domain</th><th>Last Error</th></tr>'
		}
			
		if ($queue.MessageCount -ge $t_mQueue)
		{
			$errString += "<tr><td>Mail Queue</td></td><td>$($transportServer[0]) - $($queue.Identity) - Message Count is >= $($t_mQueue)</td></tr>"
			$mbody += "<tr><td>$($queue.Identity)</td><td>$($queue.DeliveryType)</td><td>$($queue.Status)</td><td class = ""bad"">$($queue.MessageCount)</td><td>$($queue.NextHopDomain)</td><td>$($queue.LastError)</td></tr>"
		}
		else
		{
			$mbody += "<tr><td>$($queue.Identity)</td><td>$($queue.DeliveryType)</td><td>$($queue.Status)</td><td>$($queue.MessageCount)</td><td>$($queue.NextHopDomain)</td><td>$($queue.LastError)</td></tr>"
		}
		
	}
Write-Host "Done" -ForegroundColor Green
return $mbody,$errString
}

Function Create-ReplicationReport ($replInfo) {
Write-Host (Get-Date) ': Replication Health Report... ' -ForegroundColor Yellow -NoNewLine
$mbody = @()
$errString = @()
$currentServer = ""
$mbody += '<table id="SectionLabels"><tr><th class="data">DAG Members Replication</th></tr></table>'
$mbody += '<table id="data">'

	foreach ($repl in $replInfo)
	{
		if ($currentServer -ne $repl.Server)
		{
			$currentServer = $repl.Server			
			$mbody += '<tr><th><b><u>'+ $currentServer + '</b></u></th><th>Result</th><th>Error</th></tr>'
		}
	
		if ($repl.Result -match "Pass") 
		{
			$mbody += "<tr><td>$($repl.Check)</td><td>$($repl.Result)</td><td>$($repl.Error)</td></tr>"
		}
		else
		{
			$errString += "<tr><td>Replication</td></td><td>$($currentServer) - $($repl.Check) is $($repl.Result) - $($repl.Error)</td></tr>"
			$mbody += "<tr><td>$($repl.Check)</td><td class = ""bad"">$($repl.Result)</td><td>$($repl.Error)</td></tr>"
		}
	}
	$mbody += "<hr />"
Write-Host "Done" -ForegroundColor Green
return $mbody,$errString
}
	
Function Create-DiskReport ($diskinfo){
Write-Host (Get-Date) ': Disk Space Report... ' -ForegroundColor Yellow -NoNewLine
$mbody = @()
$errString = @()
$currentServer = ""
$mbody += '<table id="SectionLabels"><tr><th class="data">Disk Space</th></tr></table>'
$mbody += '<table id="data">'

	foreach ($diskdata in $diskinfo)
	{
			if ($currentServer -ne $diskdata.SystemName)
			{
				$currentServer = $diskdata.SystemName
				$mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Size (GB)</th><th>Free (GB)</th><th>Free (%)</th></tr>'
			}

			if ($diskdata.PercentFree -eq 20000)
			{
				$errString += "<tr><td>Disk</td></td><td>$($currentServer) - Error Fetching Data</td></tr>"
				$mbody += "<tr><td>$($diskdata.DeviceID) [$($diskdata.VolumeName)] </td><td>$($diskdata.Size)</td><td>$($diskdata.FreeSpace)</td><td class = ""bad"">Error Fetching Data</td></tr>"
			}
			elseif ($diskdata.PercentFree -ge $t_DiskBadPercent)
			{
				$mbody += "<tr><td>$($diskdata.DeviceID) [$($diskdata.VolumeName)] </td><td>$($diskdata.Size)</td><td>$($diskdata.FreeSpace)</td><td class = ""good"">$($diskdata.PercentFree)</td></tr>"
			}
			else
			{
				$errString += "<tr><td>Disk</td></td><td>$($currentServer) - $($diskdata.DeviceID) [$($diskdata.VolumeName)] [$($diskdata.FreeSpace) GB / $($diskdata.PercentFree)%] is <= $($t_DiskBadPercent)% Free</td></tr>"
				$mbody += "<tr><td>$($diskdata.DeviceID) [$($diskdata.VolumeName)] </td><td>$($diskdata.Size)</td><td>$($diskdata.FreeSpace)</td><td class = ""bad"">$($diskdata.PercentFree)</td></tr>"
			}
			
	}
Write-Host "Done" -ForegroundColor Green
return $mbody,$errString
}

Function Create-MdbReport ($dblist){
Write-Host (Get-Date) ': Mailbox Database Report... ' -ForegroundColor Yellow -NoNewLine
$mbody = @()
$errString = @()
$mbody += '<table id="SectionLabels"><tr><th class="data">Mailbox Database</th></tr></table>'
$mbody += '<table id="data"><tr><th>[Name][EDB Path][Log Path]</th><th>Mounted</th><th>On Server [Preference]</th><th>EDB Disk Size [Free] <br /> Log Disk Size [Free]</th><th>Size (GB)</th><th>White Space (GB)</th><th>Active Mailbox</th><th>Disconnected Mailbox</th><th>Item Size (GB)</th><th>Deleted Items Size (GB)</th><th>Full Backup</th><th>Incremental Backup</th><th>Backup In Progress</th><th>Mapi Connectivity</th></tr>'
ForEach ($db in $dblist)
{
	#$dbDetails = Get-MailboxDatabase $db.Name
	if ($db.mounted -eq $true)
	{
		#Calculate backup age----------------------------------------------------------
		if ($db.LastFullBackup -ne '')
			{
				$lastfullbackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastFullBackup
				$lastfullbackupelapsed = New-TimeSpan -Start $db.LastFullBackup
			}
		Else
			{
				$lastfullbackupelapsed = ''
				$lastfullbackup = '[NO DATA]'
			}
			
		if ($db.LastIncrementalBackup -ne '')
			{
				$lastincrementalbackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastIncrementalBackup
				$lastincrementalbackupelapsed = New-TimeSpan -Start $db.LastIncrementalBackup
			}
		Else
			{
				$lastincrementalbackupelapsed = ''
				$lastincrementalbackup = '[NO DATA]'
			}
			
		[int]$full_backup_age = $lastfullbackupelapsed.totaldays
		[int]$incremental_backup_age = $lastincrementalbackupelapsed.totaldays
		#-------------------------------------------------------------------------------
		$mbody += '<tr>'
		$mbody += '<td>[' + $db.Name + ']<br />['+ $db.EdbFilePath + ']<br />['+ $db.LogFolderPath + ']</td>'
		if ($db.Mounted -eq $true) 
		{
			$mbody += '<td class = "good">' + $db.Mounted + '</td>'
		}
		Else
		{
			$errString += "<tr><td>Database Mount</td></td><td>$($db.Name) - is NOT MOUNTED</td></tr>"
			$mbody += '<td class = "bad">' + $db.Mounted + '</td>'
		}
		
		if ($db.ActivationPreference.Value -eq 1)
		{
			$mbody += '<td class = "good">' + $db.MountedOnServer + ' ['+ $db.ActivationPreference.value +']' +'</td>'
		}
		Else
		{
			$errString += "<tr><td>Database Activation</td></td><td>$($db.Name) - is mounted on $($db.MountedOnServer) which is NOT the preferred active server</td></tr>"
			$mbody += '<td class = "bad">' + $db.MountedOnServer + ' ['+ $db.ActivationPreference.value +']' +'</td>'
		}
		
		$mbody += '<td>' + $db.EDBFreeSpace + '<br />' + $db.LogFreeSpace + '</td>'
		$mbody += '<td>' + $db.DatabaseSize + '</td><td>' + $db.AvailableNewMailboxSpace + '</td><td>' + $db.ActiveMailboxCount + '</td><td>' + $db.DisconnectedMailboxCount + '</td><td>' + $db.TotalItemSize + '</td><td>' + $db.TotalDeletedItemSize + '</td>'
		
		if ($full_backup_age -gt $t_lastfullbackup)
		{
			$errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date [$lastfullbackup] is OLDER than $($t_lastfullbackup) Day(s)</td></tr>"
			$mbody += '<td class = "bad">' + $lastfullbackup + '</td>'		
		}
		Else
		{
			$mbody += '<td class = "good">' + $lastfullbackup + '</td>'
		}
		
		if ($incremental_backup_age -gt $t_lastincrementalbackup)
		{
			$errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last incremental backup date [$lastincrementalbackup] is OLDER than $($t_lastincrementalbackup) Day(s)</td></tr>"
			$mbody += '<td class = "bad">' + $lastincrementalbackup + '</td>'
		}
		Else
		{
			$mbody += '<td class = "good"> ' + $lastincrementalbackup + '</td>'
		}
		
		$mbody += '</td><td>' + $db.BackupInProgress + '</td>'
		
		if ($db.MapiConnectivity.Result.Value -eq 'Success')
		{
			$mbody += '<td class = "good"> ' + $db.MapiConnectivity.Result.Value + '</td>'
		}
		else
		{
			$errString += "<tr><td>MAPI Connectivity</td></td><td>$($db.Name) - MAPI Connectivity Result is $($db.MapiConnectivity.Result.Value)</td></tr>"
			$mbody += '<td class = "bad"> ' + $db.MapiConnectivity.Result.Value + '</td>'
		}
	}
	else
	{
		$mbody += "<tr><td>$($db.Name)</td><td class = ""bad"">$($db.Mounted)</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td></tr>"
	}
	
		
		$mbody += '</tr>'
}
Write-Host "Done" -ForegroundColor Green
return $mbody,$errString
}

Function Create-PdbReport ($dblist){
Write-Host (Get-Date) ': Public Folder Database Report... ' -ForegroundColor Yellow -NoNewLine
$mbody += '<table id="SectionLabels"><tr><th class="data">Public Folder Database</th></tr></table>'
$mbody += '<table id="data"><tr><th>Name</th><th>Mounted</th><th>On Server</th><th>Size (GB)</th><th>White Space (GB)</th><th>Full Backup</th><th>Incremental Backup</th><th>Backup In Progress</th><th>MAPI Connectivity</th></tr>'
ForEach ($db in $dblist)
	{
		#Calculate backup age----------------------------------------------------------
		if ($db.LastFullBackup -ne '')
			{
				$lastfullbackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastFullBackup
				$lastfullbackupelapsed = New-TimeSpan -Start $db.LastFullBackup
			}
		Else
			{
				$lastfullbackupelapsed = ''
				$lastfullbackup = '[NO DATA]'
			}
			
		if ($db.LastIncrementalBackup -ne '')
			{
				$lastincrementalbackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastIncrementalBackup
				$lastincrementalbackupelapsed = New-TimeSpan -Start $db.LastIncrementalBackup
			}
		Else
			{
				$lastincrementalbackupelapsed = ''
				$lastincrementalbackup = '[NO DATA]'
			}
		[int]$full_backup_age = $lastfullbackupelapsed.totaldays
		[int]$incremental_backup_age = $lastincrementalbackupelapsed.totaldays
		#-------------------------------------------------------------------------------
		$mbody += '<tr>'
		$mbody += '<td>' + $db.Name + '</td>'
		if ($db.Mounted -eq $true) 
		{
			$mbody += '<td class = "good">' + $db.Mounted + '</td>'
		}
		Else
		{
			$errString += "<tr><td>Database Mount</td></td><td>$($db.Name) - is NOT MOUNTED</td></tr>"
			$mbody += '<td class = "bad">' + $db.Mounted + '</td>'
		}

		$mbody += '<td>' + $db.MountedOnServer + '</td><td>' + $db.DatabaseSize + '</td><td>' + $db.AvailableNewMailboxSpace + '</td>'
		
		if ($full_backup_age -gt $t_lastfullbackup)
		{
			$errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date [$lastfullbackup] is OLDER than $($t_lastfullbackup) days</td></tr>"
			$mbody += '<td class = "bad">' + $lastfullbackup + '</td>'
		}
		Else
		{
			$mbody += '<td class = "good">' + $lastfullbackup + '</td>'
		}
		
		if ($incremental_backup_age -gt $t_lastincrementalbackup)
		{
			$errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last incremental backup date [$lastfullbackup] is OLDER than $($t_lastincrementalbackup) days</td></tr>"
			$mbody += '<td class = "bad">' + $lastincrementalbackup + '</td>'
		}
		Else
		{
			$mbody += '<td class = "good"> ' + $lastincrementalbackup + '</td>'
		}		
		$mbody += '</td><td>' + $db.BackupInProgress + '</td>'
		
		if ($db.MapiConnectivity.Result.Value -eq 'Success')
		{
			$mbody += '<td class = "good"> ' + $db.MapiConnectivity.Result.Value + '</td>'
		}
		else
		{
			$mbody += '<td class = "bad"> ' + $db.MapiConnectivity.Result.Value + '</td>'
		}
		$mbody += '</tr>'
	}
Write-Host "Done" -ForegroundColor Green
return $mbody,$errString
}

#SCRIPT BEGIN---------------------------------------------------------------

#Get-List of Exchange Servers and assign to array----------------------------
Write-Host (Get-Date) ': Building List of Servers - excluding Edge' -ForegroundColor Yellow
$ExServerList = Get-ExchangeServer | ?{$_.ServerRole -notmatch 'Edge'} | Sort-Object Name
#$ExServerList = Get-ExchangeServer | Sort-Object Name
#----------------------------------------------------------------------------
#Get-List of Mailbox Database and assign to array----------------------------
if ($RunMdbReport -eq $true -OR $RunDAGCopyReport -eq $true) {
	Write-Host (Get-Date) ': Building List of Mailbox Database' -ForegroundColor Yellow
	$ExMailboxDBList = Get-MailboxDatabase -Status | where {$_.Recovery -eq $False}
}
#----------------------------------------------------------------------------
#Get-List of Public Folder Database and assign to array----------------------
if ($RunPdbReport -eq $true) {
	Write-Host (Get-Date) ': Building List of Public Folder Database' -ForegroundColor Yellow
	$ExPFDBList = Get-PublicFolderDatabase -Status | where {$_.Recovery -eq $False}
}
#----------------------------------------------------------------------------

#Begin Data Extraction-------------------------------------------------------
Write-Host '==================================================================' -ForegroundColor Green
Write-Host (Get-Date) ': Begin Data Extraction' -ForegroundColor Yellow

if ($RunServerHealthReport -eq $true) {$serverhealthdata = Get-ServerHealth($ExServerList)}
if ($RunMdbReport -eq $true) {$mdbdata = Get-MdbStatistics ($ExMailboxDBList) | Sort-Object Name}
if ($RunPdbReport -eq $true -AND $ExPFDBList.Count -gt 0) {$pdbdata = Get-PdbStatistics ($ExPFDBList)}
if ($RunDAGCopyReport -eq $true) {$dagCopyData = Get-DAGCopyStatus ($ExMailboxDBList)}
if ($RunDAGReplicationReport -eq $true) {$repldata = Get-ReplicationHealth}
if ($RunQueueReport -eq $true) {$queueData = Get-MailQueues}
if ($RunDiskReport -eq $true) {$diskdata = Get-DiskSpaceStatistics($ExServerList)}
#----------------------------------------------------------------------------
# Build Report --------------------------------------------------------------
Write-Host '==================================================================' -ForegroundColor Green
Write-Host (Get-Date) ': Create Report' -ForegroundColor Yellow
if ($RunServerHealthReport -eq $true) {$serverhealthreport,$sError = Create-ServerHealthReport ($serverhealthdata) ; $errSummary += $sError}
if ($RunMdbReport -eq $true) {$mdbreport,$mError = Create-MdbReport ($mdbdata) ; $errSummary += $mError}
if ($RunPdbReport -eq $true -AND $ExPFDBList.Count -gt 0) {$pdbreport,$pError = Create-PdbReport ($pdbdata) ; $errSummary += $pError}
if ($RunDAGCopyReport -eq $true) {$dbcopyreport,$dbCopyError = Create-DAGCopyStatusReport ($dagCopyData) ; $errSummary += $dbCopyError}
if ($RunDAGReplicationReport -eq $true) {$replicationreport,$rError = Create-ReplicationReport ($repldata) ; $errSummary += $rError}
if ($RunQueueReport -eq $true) {$queuereport,$qError = Create-QueueReport($queueData) ; $errSummary += $qError}
if ($RunDiskReport -eq $true) {$diskreport,$dError = Create-DiskReport ($diskdata) ; $errSummary += $dError}

$mail_body = "<html><head><title>[$($CompanyName)] $($MailSubject) $($today)</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
Write-Host (Get-Date) ': Applying CSS to HTML Report' -ForegroundColor Yellow
$mail_body += $css_string
$mail_body += '<table id="HeadingInfo">'
$mail_body += '<tr><th>' + $CompanyName + '<br />' + $MailSubject + '<br />' + $today + '</th></tr>'
$mail_body += '</table><hr />'
$mail_body += '<table id="SectionLabels">'
$mail_body += '<tr><th class="data">Issues</th></tr></table>'
$mail_body += '<table id="data"><tr><th>Check Item</th><th>Details</th></tr>'
$mail_body += $errSummary
$mail_body += '</table><hr />'
if ($RunServerHealthReport -eq $true) {$mail_body += $serverhealthreport ; $mail_body += '</table><hr />'}
if ($RunMdbReport -eq $true) {$mail_body += $mdbreport ; $mail_body += '</table><hr />'}
if ($RunPdbReport -eq $true) {$mail_body += $pdbreport ; $mail_body += '</table><hr />'}
if ($RunDAGReplicationReport -eq $true) {$mail_body += $replicationreport ; $mail_body += '</table><hr />'}
if ($RunDAGCopyReport -eq $true) {$mail_body += $dbcopyreport ; $mail_body += '</table><hr />'}
if ($RunQueueReport -eq $true) {$mail_body += $queuereport ; $mail_body += '</table><hr />'}
if ($RunDiskReport -eq $true) {$mail_body += $diskreport ; $mail_body += '</table><hr />'}
$mail_body += '<p><table id="SectionLabels">'
$mail_body += '<tr><th>----END of REPORT----</th></tr></table><hr /></p>'
$mail_body += '<p><font size="2" face="Tahoma"><u>Report Paremeters</u><br />'
$mail_body += '<b>[THRESHOLD]</b><br />'
$mail_body += 'Last Full Backup: ' +  $t_lastfullbackup + ' Day(s)<br />'
$mail_body += 'Last Incremental Backup: ' + $t_lastincrementalbackup + ' Day(s)<br />'
$mail_body += 'Mail Queue: ' + $t_mQueue+ '<br />'
$mail_body += 'Disk Space Critical: ' + $t_DiskBadPercent+ ' (%) <br /><br />'
$mail_body += '<b>[MAIL]</b><br />'
$mail_body += 'SMTP Server: ' + $MailServer + '<br /><br />'
$mail_body += '<b>[REPORT]</b><br />'
$mail_body += 'Generated from Server: ' + (gc env:computername) + '<br />'
$mail_body += 'Script File: ' + $MyInvocation.MyCommand.Definition + '<br />'
$mail_body += 'Config File: ' + $configFile + '<br />'
$mail_body += 'Report File: ' + $reportfile + '<br />'
$mail_body += 'Recipients: ' + $MailTo.Split(";") + '<br />'
$mail_body += '</p><p>'
$mail_body += '<a href="http://shaking-off-the-cobwebs.blogspot.com/2015/03/database-backup-and-disk-space-report.html">Exchange Server 2010 Health Check v.'+$scriptVersion+'</a></p>'
$mail_body += '</html>'
$mbody = $mbox -replace "&lt;","<"
$mbody = $mbox -replace "&gt;",">"
$mail_body | Out-File $reportfile
Write-Host (Get-Date) ': HTML Report saved to file -' $reportfile -ForegroundColor Yellow
#----------------------------------------------------------------------------
# Mail Parameters------------------------------------------------------------
# Add CC= and/or BCC= lines if you want to add recipients for CC and BCC
$params = @{
    Body = $mail_body
    BodyAsHtml = $true
    Subject = "[$($CompanyName)] $($MailSubject) $($today)"
    From = $MailSender
	To = $MailTo.Split(";")
    SmtpServer = $MailServer
}
#----------------------------------------------------------------------------
# Send Report----------------------------------------------------------------
if ($SendReportViaEmail -eq $true) {Write-Host (Get-Date) ': Sending Report' -ForegroundColor Yellow ; Send-MailMessage @params}
#----------------------------------------------------------------------------
Write-Host (Get-Date) ': End' -ForegroundColor Green
#SCRIPT END------------------------------------------------------------------
