<h3>
What the script does?</h3>
The script performs several checks on your Exchange Servers like the ones below:
<ul>
<li>Server Health (Up Time, Server Roles Services, Mail flow,...)</li>
<li>Mailbox Database Status (Mounted, Backup, Size, and Space, Mailbox Count, Paths,...)</li>
<li>Public Folder Database Status (Mount, Backup, Size, and Space,...)</li>
<li>Database Copy Status</li>
<li>Database Replication Status</li>
<li>Mail Queue</li>
<li>Disk Space</li>
<li>Server Components (for Exchange 2013/2016)</li>
</ul>
<div>
Then an HTML report will be generated and can be sent via email if enabled in the configuration file.</div>
<div>

</div>
<div>
I have not tested this for Exchange 2016 but in theory, this should work just as well (let me know if it doesn't). </div>
<div>

</div>
<h3>
Sample Output </h3>
<div>
<a href="http://www.lazyexchangeadmin.com/p/blog-page.html" target="_blank">HTML page sample</a>

</div>
<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://3.bp.blogspot.com/-SFmtGlRE9Yw/WNz6NsZjDkI/AAAAAAAAAxg/_dEalqrbNwc1dhh0aw9ISzHbvSaAxnTVgCLcB/s1600/01%2BOverall.png"><img width="640" height="198" src="https://3.bp.blogspot.com/-SFmtGlRE9Yw/WNz6NsZjDkI/AAAAAAAAAxg/_dEalqrbNwc1dhh0aw9ISzHbvSaAxnTVgCLcB/s640/01%2BOverall.png" border="0"></a></div>
<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://4.bp.blogspot.com/-yf0SIGXlG6A/WNz6NUrhsJI/AAAAAAAAAxc/g2b4YSMo4ekiydMo6KhqZw3m95989JL7QCLcB/s1600/02%2BServer%2BHealth.png"><img width="640" height="77" src="https://4.bp.blogspot.com/-yf0SIGXlG6A/WNz6NUrhsJI/AAAAAAAAAxc/g2b4YSMo4ekiydMo6KhqZw3m95989JL7QCLcB/s640/02%2BServer%2BHealth.png" border="0"></a></div>

<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://1.bp.blogspot.com/--fpTPmzfit0/WNz6NtqOeKI/AAAAAAAAAxk/LNEWDjqJWK01BZFDKSGmIMfC9-oqov0ugCLcB/s1600/03%2BServer%2BComponent.png"><img width="640" height="190" src="https://1.bp.blogspot.com/--fpTPmzfit0/WNz6NtqOeKI/AAAAAAAAAxk/LNEWDjqJWK01BZFDKSGmIMfC9-oqov0ugCLcB/s640/03%2BServer%2BComponent.png" border="0"></a></div>

<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://2.bp.blogspot.com/-LRjHtKlZRNU/WNz6OXJfhpI/AAAAAAAAAxo/giGxAJGM45cbaGF3Kqx4CW6SrBTdWULTQCLcB/s1600/04%2BMailbox%2BDatabase.png"><img width="640" height="192" src="https://2.bp.blogspot.com/-LRjHtKlZRNU/WNz6OXJfhpI/AAAAAAAAAxo/giGxAJGM45cbaGF3Kqx4CW6SrBTdWULTQCLcB/s640/04%2BMailbox%2BDatabase.png" border="0"></a></div>

<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://4.bp.blogspot.com/-DZ7n4VejqrE/WNz6OVREf4I/AAAAAAAAAxs/mYpWhbV8vHwSZrqDWvHPPLX80UgrkAoXgCLcB/s1600/05%2BReplication.png"><img width="640" height="153" src="https://4.bp.blogspot.com/-DZ7n4VejqrE/WNz6OVREf4I/AAAAAAAAAxs/mYpWhbV8vHwSZrqDWvHPPLX80UgrkAoXgCLcB/s640/05%2BReplication.png" border="0"></a></div>

<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://3.bp.blogspot.com/-RIYSPCWbb9g/WNz6OQ6kS6I/AAAAAAAAAxw/kVFnZC05lNcJH1eRopBXQKJT85BX78niQCLcB/s1600/06%2BDatabase%2BCopy.png"><img width="640" height="136" src="https://3.bp.blogspot.com/-RIYSPCWbb9g/WNz6OQ6kS6I/AAAAAAAAAxw/kVFnZC05lNcJH1eRopBXQKJT85BX78niQCLcB/s640/06%2BDatabase%2BCopy.png" border="0"></a></div>

<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://2.bp.blogspot.com/-iGluYFsf-ZA/WNz6O4dq5-I/AAAAAAAAAx4/THBflzhCZUMW0qtgSP2jTdNh-3McViHEgCLcB/s1600/07%2BMail%2BQueue.png"><img width="640" height="76" src="https://2.bp.blogspot.com/-iGluYFsf-ZA/WNz6O4dq5-I/AAAAAAAAAx4/THBflzhCZUMW0qtgSP2jTdNh-3McViHEgCLcB/s640/07%2BMail%2BQueue.png" border="0"></a></div>

<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://4.bp.blogspot.com/-S-n0kwMpPls/WNz6O87rdCI/AAAAAAAAAx0/_XQMHGqYru83WPfCwcGKppUxDia6h1Z1QCLcB/s1600/08%2BDisk%2BSpace.png"><img width="640" height="152" src="https://4.bp.blogspot.com/-S-n0kwMpPls/WNz6O87rdCI/AAAAAAAAAx0/_XQMHGqYru83WPfCwcGKppUxDia6h1Z1QCLcB/s640/08%2BDisk%2BSpace.png" border="0"></a></div>

<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://2.bp.blogspot.com/-3JYLfX7twhY/WNz6PP5wMoI/AAAAAAAAAx8/RnFUF8-p0aE2e_CojZDZP4G7Z12U59woACLcB/s1600/09%2BConfig.png"><img width="320" height="248" src="https://2.bp.blogspot.com/-3JYLfX7twhY/WNz6PP5wMoI/AAAAAAAAAx8/RnFUF8-p0aE2e_CojZDZP4G7Z12U59woACLcB/s320/09%2BConfig.png" border="0"></a></div>

<div>

</div>
<h3>
Parameters</h3>
<div>
<ul>
<li><b>-configFile</b>, to specify the XML file that contain the configuration for the script.</li>
<li><b>-enableDebug</b>, optional switch to start a transcript output to debugLog.txt</li>
</ul>
<h3>
Configuration File</h3>
</div>
<div>
The configuration file is an XML file containing the options, thresholds, mail settings, exclusion that will be used by the script. The snapshot of the configuration file is seen below:</div>
<div>

</div>
<div class="separator" style="text-align: center; clear: both;">
<a style="margin-right: 1em; margin-left: 1em;" href="https://3.bp.blogspot.com/-aFjccM_rIHM/WNz88KYmCRI/AAAAAAAAAyI/Hb2u9KSvzvEKJS5POvwp-OeIY26mK4GpwCLcB/s1600/10%2BXML.png"><img width="400" height="390" src="https://3.bp.blogspot.com/-aFjccM_rIHM/WNz88KYmCRI/AAAAAAAAAyI/Hb2u9KSvzvEKJS5POvwp-OeIY26mK4GpwCLcB/s400/10%2BXML.png" border="0"></a></div>
<div class="separator" style="text-align: left; clear: both;">
</div>
<div>

</div>
<h4>
reportOptions</h4>
This section can be toggled by changing values with "true" or "false"

<ul>
<li><b>RunServerHealthReport </b>- Run test and report the Server Health status</li>
<li><b>RunMdbReport </b>- Mailbox Database test and report</li>
<li><b>RunComponentReport </b>- Server Components check (Exchange 2013/2016)</li>
<li><b>RunPdbReport </b>- For checking the Public Folder database(s)</li>
<li><b>RunDAGReplicationReport </b>- Check and test replication status</li>
<li><b>RunQueueReport </b>- Inspect mail queue count</li>
<li><b>RunDiskReport </b>- Disk space report for each server</li>
<li><b>RunDBCopyReport </b>- Checking the status of the Database Copies</li>
<li><b>SendReportViaEmail </b>- Option to send the HTML report via email</li>
<li><b>ReportFile </b>- File path and name of the HTML Report</li>
</ul>
<h4>
thresholds</h4>

This section defines at which levels the script will report a problem for each check item

<ul>
<li><b>LastFullBackup </b>- age of full backup in days. Setting this to zero (0) will cause the script to ignore this threshold</li>
<li><b>LastIncrementalBackup </b>- age of incremental backup in days. Setting this to zero (0) will cause the script to ignore this threshold.</li>
<li><b>DiskSpaceFree </b>- percent (%) of free disk space left</li>
<li><b>MailQueueCount </b>- Mail transport queue threshold</li>
<li><b>CopyQueueLenght </b>- CopyQueueLenght threshold for the DAG replication</li>
<li><b>ReplayQueueLenght </b>- ReplayQueueLenght threshold</li>
<li><strong>cpuUsage</strong> – CPU usage threshold %</li>
<li><strong>ramUsage</strong> – Memory usage threshold %</li>
</ul>
<h4>

</h4>
<h4>
mailAndReportParameters</h4>
This section specifies the mail parameters
<div>
<ul>
<li><b>CompanyName </b>- the name of the organization or company that you want to appear in the banner of the report</li>
<li><b>MailSubject </b>- Subject of the email report</li>
<li><b>MailServer </b>- The SMTP Relay server</li>
<li><b>MailSender </b>- Mail sender address</li>
<li><b>MailTo </b>- Recipient address. For multiple recipients, separate the addresses with a semi-colon (;)</li>
</ul>
<h4>
exclusions</h4>
</div>
<div>
This section is where the exclusion can be defined.</div>
<div>
<ul>
<li><b>IgnoreServer </b>- List of servers to be ignored by the script. Separate with a comma (,) with no spaces.</li>
<li><b>IgnoreDatabase </b>- List of Mailbox Database to be ignored by the script. Separate with a comma (,) with no spaces.</li>
<li><b>IgnorePFDatabase </b>- List of Public Folder Database to be ignored by the script. Separate with a comma (,) with no spaces.</li>
</ul>
<h3>
How to Use</h3>
</div>
<h4>
Run manually using Exchange Management Shell</h4>
<div class="separator" style="text-align: center; clear: both;">
<a style="clear: left; margin-right: 1em; margin-bottom: 1em; float: left;" href="https://1.bp.blogspot.com/-E4ry80tCFFY/WN0CJMdiu7I/AAAAAAAAAyY/wvo3LKqtkDomdwu1p6y7MyhiNxY7zpIVgCLcB/s1600/11%2BUsage.png"><img src="https://1.bp.blogspot.com/-E4ry80tCFFY/WN0CJMdiu7I/AAAAAAAAAyY/wvo3LKqtkDomdwu1p6y7MyhiNxY7zpIVgCLcB/s1600/11%2BUsage.png" border="0"></a></div>
<div>

</div>
<div>

</div>
<div>

</div>
<div>
<b>[PS] C:\scripts&gt;.\Get-ExchangeHealth.ps1 -configFile .\config.xml -enableDebug</b></div>
<div>

</div>
<div>
<strong><u>Note:</u> <font color="#ff0000">This must be run within the Exchange Management Shell session. Avoid using this inside the normal PowerShell session, with the Exchange 2010 PSSnapin loaded especially for Exchange 2013 servers.</font></strong></div>
<div>

</div>
<h4>
Task Scheduler</h4>
<div>
Create a task in Task Scheduler with this action:</div>
<div>

</div>
<div>
<b>Program/script</b>: powershell.exe </div>
<div>
<b>Add arguments</b>: -Command ".'C:\Program Files\Microsoft\Exchange Server\V15\Bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto; C:\scripts\Get-ExchangeHealth.ps1 -configFile C:\scripts\Get-\config.xml"</div>
<div>

</div>
<h3>
Download</h3>
<div>
You can find the latest script and config.xml file in this GitHub repository:</div>
<div>
<a href="https://github.com/junecastillote/Get-ExchangeHealth">https://github.com/junecastillote/Get-ExchangeHealth</a></div>
<div>

</div>
<h3>
Change Logs</h3>
<h4>
Version 5.4 (Latest)</h4>
<ul>
<li>Added CPU and Memory Utilization Checks</li>
<ul>
<li>New configuration in config.xml (ramUsage, cpuUsage, RunCPUandMemoryReport)</li>
</ul>
<li>Code clean-up</li>
</ul>
<h4>
Version 5.3 (skipped, crappy version)</h4>
<div>
<h4>
Version 5.2 (Latest)</h4>
<div>
<ul>
<li>Several code cleanups</li>
<li>Renamed RunDAGCopyReport to RunDBCopyReport (because it makes more sense to call it that)</li>
<li>Removed Add-PSSnapin code because it didn't play well with Exchange 2013. Hence the need to use Exchange Management Shell.</li>
<li>Added the Version and Edition of Exchange Server column in Server Health report</li>
<li>Revised the Get-MailQueues function</li>
<li>Added AdminDisplayVersion identification in Get-ServerHealth function</li>
<li>Removed PowerShell-Remoting code from the Mail flow test</li>
<li>Renamed Get-DAGCopyStatus function to Get-DatabaseCopyStatus</li>
<li>Added CopyQueueLenght threshold in XMLconfig file</li>
<li>Added ReplayQueueLenght threshold in XMLconfig file</li>
<li>Added Server Components check (for Exchange 2013)</li>
<li>Added RunComponentReport option in XML config file</li>
<li>Added <exclusions> section in XML config file</exclusions></li>
<li>Added IgnoreServer,IgnoreDatabase,IgnorePFDatabase fields in XML config file</li>
<li>Added counter for the number of tests, passed and failed.</li>
<li>Added percentage computation for overall health</li>
<li>The "Issues" table is not visible from the report if there are no actual issues detected</li>
<li>Added individual test results summary</li>
<li>Added logic to not run DAG checks if there are no DAGs</li>
<li>Added logic to not run Mail Flow test against Mailbox Servers with no Active Mailbox Database</li>
</ul>
</div>
<h4>
Version 5.1 </h4>
<div>
<ul>
<li>Added new parameter "configFile" where you will need to specify the configuration XML file which contains the variables that used to be included inside the script in previous versions.</li>
<li>Moved the Variables to an outside XML file (default is config.xml) You can create different XML files with different configurations/variables if desired.</li>
<li>To run: "Get-ExchangeHealth.ps1 -configFile config.XML"</li>
<li>Added DAG Copy Status</li>
<li>Fixed the Math for getting back up age</li>
</ul>
</div>
<h4>
Version 4.4b</h4>
<div>
<ul>
<li>Corrected version information within the script</li>
<li>Added BCC and CC line within the @params variable block, but are commented out.</li>
<li>Added comments to:</li>
<li>[int]$t_lastincrementalbackup</li>
<li>[int]$t_lastfullbackup</li>
<li>Added comments to:</li>
<li>$MailCC</li>
<li>$MailBCC</li>
</ul>
</div>
<h4>
Version 4.4</h4>
<div>
<ul>
<li>Added Test-MailFlow Handle for Exchange 2013</li>
<li>Moved Test-MailFlow Result to Server Health Status Report</li>
<li>Exclude Edge Servers from Testing</li>
<li>Public Folder Database Report will not run if database count is 0</li>
</ul>
</div>
<h4>
Version 4.3</h4>
<div>
<ul>
<li>Renamed script to Get-ExchangeHealth.ps1</li>
<li>Added Test-MapiConnectivity</li>
<li>Added Test-MailFlow</li>
<li>Added Services Status</li>
<li>Added DNS/Ping Test</li>
<li>Added Server Up Time</li>
<li>Changed Backup Threshold from Days to Hours</li>
</ul>
</div>
<h4>
Version 4.0</h4>
<div>
<ul>
<li>Added DAG Members Replication Checks </li>
<li>Added Mail Queue Checks</li>
<li>Added DB Activation Preference Check</li>
<li>Added "Summary" Section</li>
<li>Fixed HTML character recognition issues</li>
</ul>
</div>
</div>
<h4>
Older versions --- :)</h4>
