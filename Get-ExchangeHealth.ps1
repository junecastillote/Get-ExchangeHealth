Write-Host '===================================================' -ForegroundColor Yellow
Write-Host '>>          Get-ExchangeHealth v4.2              <<' -ForegroundColor Yellow
Write-Host '>>         june.castillote@gmail.com             <<' -ForegroundColor Yellow
Write-Host '===================================================' -ForegroundColor Yellow
#http://shaking-off-the-cobwebs.blogspot.com/2015/03/database-backup-and-disk-space-report.html
Write-Host ''
Write-Host (Get-Date) ': Begin' -ForegroundColor Green
Write-Host (Get-Date) ': Setting Paths and Variables' -ForegroundColor Yellow
#$ErrorActionPreference="SilentlyContinue"
$WarningPreference="SilentlyContinue";
#>>Define Variables---------------------------------------------------------------
$errSummary = ""
$today = '{0:dd-MMM-yyyy hh:mm tt}' -f (Get-Date)
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$css_string = '<style type="text/css"> #HeadingInfo { font-family:Tahoma, "Trebuchet MS", Arial, Helvetica, sans-serif; width:100%; border-collapse:collapse; } #HeadingInfo td, #HeadingInfo th { font-size:0.9em; padding:3px 7px 2px 7px; } #HeadingInfo th  { font-size:1.0em; font-weight:bold; text-align:center; padding-top:5px; padding-bottom:4px; background-color:#CC3300; color:#fff; } #SectionLabels { font-family:Tahoma, "Trebuchet MS", Arial, Helvetica, sans-serif; width:100%; border-collapse:collapse; } #SectionLabels th.data { font-size:0.8em; text-align:center; padding-top:5px; padding-bottom:4px; background-color:#A7C942; color:#fff; } #data { font-family:Consolas,Tahoma, "Trebuchet MS", Arial, Helvetica, sans-serif; width:100%; border-collapse:collapse; } #data td, #data th  { font-size:0.8em; border:1px solid #98bf21; padding:3px 7px 2px 7px; } #data th  { font-size:0.8em; padding-top:5px; padding-bottom:4px; background-color:#A7C942; color:#fff; text-align:left; } #data td { font-size:0.8em; padding-top:5px; padding-bottom:4px; text-align:left; } #data td.bad { font-size:0.8em; font-weight: bold; padding-top:5px; padding-bottom:4px; background-color:red; } #data td.good { font-size:0.8em; font-weight: bold; padding-top:5px; padding-bottom:4px; color:green; }</style> </head> <body> <hr />'
$reportfile = $script_root + "\DbAndDiskReport_" + ('{0:dd_MMM_yyyy}' -f (Get-Date)) + ".html"
#>>------------------------------------------------------------------------------
#>>Thresholds--------------------------------------------------------------------
[int]$t_lastfullbackup = 7
[int]$t_lastincrementalbackup = 1
[int]$t_DiskBadPercent = 12
[int]$t_mQueue = 20
#>>------------------------------------------------------------------------------
#>>Options, set to $false if you do not want to run a specific report------------
$RunMdbReport = $true
$RunPdbReport = $true
$RunDAGReplicationReport = $true
$RunQueueReport = $true
$RunDiskReport = $true
$SendReportViaEmail = $true
#>>------------------------------------------------------------------------------
#>>Mail
$CompanyName = 'Boral Limited'
$MailSubject = '[BORAL] Exchange Service Health Report '
$MailServer = 'cluster4out.us.messagelabs.com'
$MailSender = 'Boral PostMaster <exchange-Admin@boral.com.au>'
#$MailTo = 'tito.castillote-jr@hpe.com'
$MailTo = 'gcp.messaging.exchangets@hp.com'
$MailCC = ''
$MailBCC = ''
#>>------------------------------------------------------------------------------
#>>Import Exchange 2010 Shell Snap-In if not already added-----------------------
if ($RunMdbReport -eq $true -OR $RunPdbReport -eq $true -OR $RunDAGReplicationReport -eq $true -OR $RunQueueReport -eq $true)
{
	if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
	{
		try
		{
			Write-Host (Get-Date) ': Add Exchange 2010 Snap-in' -ForegroundColor Yellow
			Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
		}
		catch
		{
			Write-Warning $_.Exception.Message
			EXIT
		}
	}
}
#>>------------------------------------------------------------------------------
Function Get-ServerHealth ($serverlist) {
Write-Host (Get-Date) ': Server Status Check... ' -ForegroundColor Yellow -NoNewLine
	$stats_collection = @()
		foreach ($server in $serverlist)
		{
			$serverOS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -ErrorAction STOP
			$timespan = $serverOS.ConvertToDateTime($serverOS.LocalDateTime) – $serverOS.ConvertToDateTime($serverOS.LastBootUpTime)
			[int]$uptime = "{0:00}" -f $timespan.TotalHours
			
			$serverobj = "" | Server,UpTime,RoleServicesUp,RoleServicesDown
			$serverobj.UpTime = $uptime
		}
}
Function Get-MdbStatistics ($mailboxdblist){
Write-Host (Get-Date) ': Mailbox Database Check... ' -ForegroundColor Yellow -NoNewLine
	$stats_collection = @()
		foreach ($mailboxdb in $mailboxdblist)
		{
			$mdbobj = "" | Select Name,Mounted,MountedOnServer,ActivationPreference,DatabaseSize,AvailableNewMailboxSpace,ActiveMailboxCount,DisconnectedMailboxCount,TotalItemSize,TotalDeletedItemSize,EdbFilePath,LogFolderPath,LogFilePrefix,LastFullBackup,LastIncrementalBackup,BackupInProgress,MapiConnectivity,MailFlow
			$mdbobj.Name = $mailboxdb.name
			$mdbobj.Mounted = $mailboxdb.Mounted
			$mdbobj.MountedOnServer = $mailboxdb.Server.Name
			$mdbobj.ActivationPreference = $mailboxdb.ActivationPreference | ?{$_.Key -eq $mailboxdb.Server.Name}
			$mdbobj.LastFullBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $mailboxdb.LastFullBackup
			$mdbobj.LastIncrementalBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $mailboxdb.LastIncrementalBackup
			$mdbobj.BackupInProgress = $mailboxdb.BackupInProgress
			$mdbobj.DatabaseSize = "{0:N2}" -f ($mailboxdb.DatabaseSize.tobytes() / 1GB)
			$mbxItemSize = Get-MailboxStatistics -Database $mailboxdb | %{$_.TotalItemSize.Value} | Measure-Object -sum
			$mbxDelSize = Get-MailboxStatistics -Database $mailboxdb | %{$_.TotalDeletedItemSize.Value} | Measure-Object -sum
			$mdbobj.TotalItemSize = "{0:N2}" -f ($mbxItemSize.sum / 1GB)
			$mdbobj.TotalDeletedItemSize = "{0:N2}" -f ($mbxDelSize.sum / 1GB)
			$mdbobj.ActiveMailboxCount = (Get-MailboxStatistics -Database $mailboxdb | where {$_.DisconnectDate -eq $null}).count
			$mdbobj.DisconnectedMailboxCount = (Get-MailboxStatistics -Database $mailboxdb | where {$_.DisconnectDate -ne $null}).count
			$mdbobj.AvailableNewMailboxSpace = "{0:N2}" -f ($mailboxdb.AvailableNewMailboxSpace.tobytes() / 1GB)
			#MAPI CONNECTIVITY
			$mdbobj.MapiConnectivity = Test-MapiConnectivity -Database $mailboxdb.Identity -PerConnectionTimeout 10
			$mdbobj.MailFlow = Test-Mailflow $mailboxdb.Server.Name
			
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
	$stats_collection = get-TransportServer | Sort-Object Name | %{Get-Queue -Server $_}
	Write-Host "Done" -ForegroundColor Green
	return $stats_collection
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
				$errString += "<tr><td>Disk</td></td><td>$($currentServer) - $($diskdata.DeviceID) [$($diskdata.VolumeName)] is <= $($t_DiskBadPercent)% Free</td></tr>"
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
$mbody += '<table id="data"><tr><th>Name</th><th>Mounted</th><th>On Server [Preference]</th><th>Size (GB)</th><th>White Space (GB)</th><th>Active Mailbox</th><th>Disconnected Mailbox</th><th>Item Size (GB)</th><th>Deleted Items Size (GB)</th><th>Full Backup</th><th>Incremental Backup</th><th>Backup In Progress</th><th>Mapi Connectivity</th><th>Mail Flow</th></tr>'
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
			
		[int]$full_backup_age = $lastfullbackupelapsed.days
		[int]$incremental_backup_age = $lastincrementalbackupelapsed.days
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
		
		if ($db.ActivationPreference.Value -eq 1)
		{
			$mbody += '<td class = "good">' + $db.MountedOnServer + ' ['+ $db.ActivationPreference.value +']' +'</td>'
		}
		Else
		{
			$errString += "<tr><td>Database Activation</td></td><td>$($db.Name) - is mounted on $($db.MountedOnServer) which is NOT the preferred active server</td></tr>"
			$mbody += '<td class = "bad">' + $db.MountedOnServer + ' ['+ $db.ActivationPreference.value +']' +'</td>'
		}
		
		$mbody += '<td>' + $db.DatabaseSize + '</td><td>' + $db.AvailableNewMailboxSpace + '</td><td>' + $db.ActiveMailboxCount + '</td><td>' + $db.DisconnectedMailboxCount + '</td><td>' + $db.TotalItemSize + '</td><td>' + $db.TotalDeletedItemSize + '</td>'
		
		if ($full_backup_age -gt $t_lastfullbackup)
		{
			$errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date is OLDER than $($t_lastfullbackup) days</td></tr>"
			$mbody += '<td class = "bad">' + $lastfullbackup + '</td>'
		}
		Else
		{
			$mbody += '<td class = "good">' + $lastfullbackup + '</td>'
		}
		
		if ($incremental_backup_age -gt $t_lastincrementalbackup)
		{
			$errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last incremental backup date OLDER than $($t_lastincrementalbackup) days</td></tr>"
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
		
		if ($db.MailFlow.TestMailflowResult -eq 'Success')
		{
			$mbody += '<td class = "good"> ' + $db.MailFlow.TestMailflowResult + '</td>'
		}
		else
		{
			$errString += "<tr><td>Mail Flow</td></td><td>$($db.Name) - Mail Flow Result is $($db.MailFlow.TestMailflowResult)</td></tr>"
			$mbody += '<td class = "bad"> ' + $db.MailFlow.TestMailflowResult + '</td>'
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
		[int]$full_backup_age = $lastfullbackupelapsed.days
		[int]$incremental_backup_age = $lastincrementalbackupelapsed.days
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
			$errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date OLDER than $($t_lastfullbackup) days</td></tr>"
			$mbody += '<td class = "bad">' + $lastfullbackup + '</td>'
		}
		Else
		{
			$mbody += '<td class = "good">' + $lastfullbackup + '</td>'
		}
		
		if ($incremental_backup_age -gt $t_lastincrementalbackup)
		{
			$errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last incremental backup date OLDER than $($t_lastincrementalbackup) days</td></tr>"
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

#>>SCRIPT BEGIN---------------------------------------------------------------

#>>Get-List of Exchange Servers and assign to array----------------------------
Write-Host (Get-Date) ': Building List of Servers' -ForegroundColor Yellow
$ExServerList = Get-ExchangeServer | Sort-Object Name
#>>----------------------------------------------------------------------------
#>>Get-List of Mailbox Database and assign to array----------------------------
if ($RunMdbReport -eq $true) {
	Write-Host (Get-Date) ': Building List of Mailbox Database' -ForegroundColor Yellow
	$ExMailboxDBList = Get-MailboxDatabase -Status | where {$_.Recovery -eq $False}
}
#>>----------------------------------------------------------------------------
#>>Get-List of Public Folder Database and assign to array----------------------
if ($RunPdbReport -eq $true) {
	Write-Host (Get-Date) ': Building List of Public Folder Database' -ForegroundColor Yellow
	$ExPFDBList = Get-PublicFolderDatabase -Status | where {$_.Recovery -eq $False}
}
#>>----------------------------------------------------------------------------
#>>Begin Data Extraction-------------------------------------------------------
Write-Host '==================================================================' -ForegroundColor Green
Write-Host (Get-Date) ': Begin Data Extraction' -ForegroundColor Yellow

if ($RunMdbReport -eq $true) {$mdbdata = Get-MdbStatistics ($ExMailboxDBList) | Sort-Object Name}
if ($RunPdbReport -eq $true) {$pdbdata = Get-PdbStatistics ($ExPFDBList)}
if ($RunDAGReplicationReport -eq $true) {$repldata = Get-ReplicationHealth}
if ($RunQueueReport -eq $true) {$queueData = Get-MailQueues}
if ($RunDiskReport -eq $true) {$diskdata = Get-DiskSpaceStatistics($ExServerList)}
#>>----------------------------------------------------------------------------
#>> Build Report --------------------------------------------------------------
Write-Host '==================================================================' -ForegroundColor Green
Write-Host (Get-Date) ': Create Report' -ForegroundColor Yellow
if ($RunMdbReport -eq $true) {$mdbreport,$mError = Create-MdbReport ($mdbdata) ; $errSummary += $mError}
if ($RunPdbReport -eq $true) {$pdbreport,$pError = Create-PdbReport ($pdbdata) ; $errSummary += $pError}
if ($RunDAGReplicationReport -eq $true) {$replicationreport,$rError = Create-ReplicationReport ($repldata) ; $errSummary += $rError}
if ($RunQueueReport -eq $true) {$queuereport,$qError = Create-QueueReport($queueData) ; $errSummary += $qError}
if ($RunDiskReport -eq $true) {$diskreport,$dError = Create-DiskReport ($diskdata) ; $errSummary += $dError}

$mail_body = '<html><head><title>' + ($MailSubject+$today) + ' </title><meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />'
Write-Host (Get-Date) ': Applying CSS to HTML Report' -ForegroundColor Yellow
$mail_body += $css_string
$mail_body += '<table id="HeadingInfo">'
$mail_body += '<tr><th>' + $CompanyName + '<br />' + $MailSubject + '<br />' + $today + '</th></tr>'
$mail_body += '</table><hr />'
$mail_body += '<table id="SectionLabels">'
$mail_body += '<tr><th class="data">----SUMMARY----</th></tr></table>'
$mail_body += '<table id="data"><tr><th>Check Item</th><th>Details</th></tr>'
$mail_body += $errSummary
$mail_body += '</table><hr />'
if ($RunMdbReport -eq $true) {$mail_body += $mdbreport ; $mail_body += '</table><hr />'}
if ($RunPdbReport -eq $true) {$mail_body += $pdbreport ; $mail_body += '</table><hr />'}
if ($RunDAGReplicationReport -eq $true) {$mail_body += $replicationreport ; $mail_body += '</table><hr />'}
if ($RunQueueReport -eq $true) {$mail_body += $queuereport ; $mail_body += '</table><hr />'}
if ($RunDiskReport -eq $true) {$mail_body += $diskreport ; $mail_body += '</table><hr />'}
$mail_body += '<p>'
$mail_body += '<table id="SectionLabels">'
$mail_body += '<tr><th>----END of REPORT----</th></tr></table><hr />'
$mail_body += '<p><font size="2" face="Tahoma"><u>Report Paremeters</u><br />'
$mail_body += '<b>[THRESHOLD]</b><br />'
$mail_body += 'Last Full Backup: ' +  $t_lastfullbackup + ' day(s)<br />'
$mail_body += 'Last Incremental Backup: ' + $t_lastincrementalbackup + ' day(s)<br />'
$mail_body += 'Mail Queue: ' + $t_mQueue+ '<br />'
$mail_body += 'Disk Space Critical: ' + $t_DiskBadPercent+ ' (%) <br /><br />'
$mail_body += '<b>[MAIL]</b><br />'
$mail_body += 'SMTP Server: ' + $MailServer + '<br /><br />'
$mail_body += '<b>[REPORT]</b><br />'
$mail_body += 'Generated from Server: ' + (gc env:computername) + '<br />'
$mail_body += 'Script Path: ' + $script_root
$mail_body += '<p>'
$mail_body += '<a href="http://shaking-off-the-cobwebs.blogspot.com/2015/03/database-backup-and-disk-space-report.html">Exchange Server 2010 Health Check v.4.1</a>'
$mbody = $mbox -replace "&lt;","<"
$mbody = $mbox -replace "&gt;",">"
$mail_body | Out-File $reportfile
Write-Host (Get-Date) ': HTML Report saved to file -' $reportfile -ForegroundColor Yellow
#>>----------------------------------------------------------------------------
#>> Mail Parameters------------------------------------------------------------
#>> Add CC= and/or BCC= lines if you want to add recipients for CC and BCC
$params = @{
    Body = $mail_body
    BodyAsHtml = $true
    Subject = "$MailSubject$today"
    From = $MailSender
	To = $MailTo.Split(",")
    SmtpServer = $MailServer
}
#>>----------------------------------------------------------------------------
#>> Send Report----------------------------------------------------------------
if ($SendReportViaEmail -eq $true) {Write-Host (Get-Date) ': Sending Report' -ForegroundColor Yellow ; Send-MailMessage @params}
#>>----------------------------------------------------------------------------
Write-Host (Get-Date) ': End' -ForegroundColor Green
#>>SCRIPT END------------------------------------------------------------------