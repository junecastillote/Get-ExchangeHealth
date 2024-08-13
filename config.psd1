@{
    reportOptions           = @{
        RunCPUandMemoryReport   = $true
        RunServerHealthReport   = $true
        RunMdbReport            = $true
        RunComponentReport      = $true
        RunPdbReport            = $false
        RunDAGReplicationReport = $false
        RunQueueReport          = $true
        RunDiskReport           = $true
        RunDBCopyReport         = $false
        SendReportViaEmail      = $false
        ReportFile              = "C:\Scripts\Get-ExchangeHealth\XYZ_Exchange_Hourly_Report.html"
    }
    thresholds              = @{
        LastFullBackup        = 7
        LastIncrementalBackup = 1
        DiskSpaceFree         = 12
        MailQueueCount        = 20
        CopyQueueLenght       = 10
        ReplayQueueLenght     = 10
        CpuUsage              = 60
        RamUsage              = 80
    }
    mailAndReportParameters = @{
        CompanyName = "XYZ"
        MailSubject = "Exchange Service Health Report"
        MailServer  = "192.168.56.30"
        MailSender  = "XYZ PostMaster <exchange-Admin@XYZ.com.au>"
        MailTo      = @('administrator@xyz.com')
        MailCc      = @()
        MailBcc     = @()
        SSLEnabled  = $false
        Port        = 25
    }
    exclusions              = @{
        IgnoreServer     = @()
        IgnoreDatabase   = @()
        IgnorePFDatabase = @()
    }
}