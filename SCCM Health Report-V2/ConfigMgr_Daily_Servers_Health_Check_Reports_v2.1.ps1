<#
.SYNOPSIS
    ConfigMgr daily servers health check report (v2). Read-only: it collects data and builds an HTML (+CSV) report.

.DESCRIPTION
    - Uses the same ConfigFile.xml as v1. Every new key is optional and has a default.
    - Run interactively and it asks which servers to check (ENTER keeps the value from ConfigFile.xml).
      Run with -NonInteractive (or from a scheduled task with no logged-on user) and it uses ConfigFile.xml only.
    - Collection is separated from rendering: each check returns a "section" (title, columns, rows),
      and one renderer turns all sections into HTML and CSV.
    - A failing check never stops the report: it appears as a Warning row with the error message.
    - Nothing is changed on any server. Only reads: CIM (DCOM), WinRM (inboxes, certificates), SQL SELECT, event logs, HTTP GET.

.PARAMETER OutputPath
    Folder with ConfigFile.xml, the report files, the log and Historical_Reports.

.PARAMETER ConfigPath
    Full path of the config file. Default: <OutputPath>\ConfigFile.xml

.PARAMETER NonInteractive
    Never prompt. Use it in scheduled tasks.

.EXAMPLE
    .\ConfigMgr_Daily_Servers_Health_Check_Reports_v2.ps1
    .\ConfigMgr_Daily_Servers_Health_Check_Reports_v2.ps1 -NonInteractive

.NOTES
    Version : 02.00
    Based on: ConfigMgr_Daily_Servers_Health_Check_Reports v01.00 (Karthikeyan Anbarasu, 28/01/2018)
    Target  : Windows PowerShell 4.0+ (Windows Server 2012 R2 or later)
#>
[CmdletBinding()]
param(
    [string]$OutputPath = 'D:\Scripts\ConfigMgr_Daily_Servers_Health_Check_Reports',
    [string]$ConfigPath,
    [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '02.00'
if (-not $ConfigPath) { $ConfigPath = Join-Path $OutputPath 'ConfigFile.xml' }
$Interactive = (-not $NonInteractive) -and [Environment]::UserInteractive

# =====================================================================================================
# Basics: log, config helpers
# =====================================================================================================
$LogFile = Join-Path $OutputPath 'ConfigMgr_Daily_Servers_Health_Check_Reports.log'
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    try { Add-Content -Path $LogFile -Value $line } catch { }
}

function Get-Setting {
    param($Node, [string]$Default = '')
    if ($null -ne $Node -and -not [string]::IsNullOrWhiteSpace([string]$Node)) { return ([string]$Node).Trim() }
    return $Default
}

$Config = $null
if (Test-Path -LiteralPath $ConfigPath) {
    [xml]$Config = Get-Content -LiteralPath $ConfigPath
} elseif (-not $Interactive) {
    Write-Log "Config file not found: $ConfigPath (and the run is not interactive)." 'ERROR'
    exit 1
} else {
    Write-Log "Config file not found: $ConfigPath. You will be asked for every value." 'WARN'
}
$S = $Config.Settings

# =====================================================================================================
# Questions: which servers / site data to collect from
# =====================================================================================================
$NamePattern = '^[A-Za-z0-9][A-Za-z0-9._-]{0,62}$'

function Read-Answer {
    # Asks (when interactive), validates, and returns the cleaned value(s). ENTER keeps the default.
    param(
        [string]$Label,
        [string]$Default,
        [string]$Pattern = $NamePattern,
        [switch]$Single,         # one value only; the text is NOT split (allows SERVER\INSTANCE,PORT)
        [switch]$KeepCase,       # do not upper-case (database names and SQL instances can be case-sensitive)
        [switch]$AllowEmpty
    )
    while ($true) {
        $answer = $Default
        if ($Interactive) {
            $shown = ''
            if ($Default) { $shown = " [$Default]" }
            $typed = Read-Host "$Label$shown"
            if (-not [string]::IsNullOrWhiteSpace($typed)) { $answer = $typed }
        }
        if ($Single) { $items = @($answer.Trim()) } else { $items = @($answer -split '[,;\s]+' | Where-Object { $_ }) }
        $items = @($items | Where-Object { $_ })
        if ($items.Count -eq 0) {
            if ($AllowEmpty) { return @() }
            $problem = 'a value is required'
        } else {
            $bad = @($items | Where-Object { $_ -notmatch $Pattern })
            if ($bad.Count -eq 0) {
                if ($KeepCase) { return @($items | Select-Object -Unique) }
                return @($items | ForEach-Object { $_.ToUpper() } | Select-Object -Unique)
            }
            $problem = "invalid value(s): $($bad -join ', ')"
        }
        if (-not $Interactive) { Write-Log "$Label - $problem" 'ERROR'; exit 1 }
        Write-Host "  $problem. Try again." -ForegroundColor Yellow
    }
}

if ($Interactive) {
    Write-Host ''
    Write-Host 'ConfigMgr health check: confirm the servers to collect from (ENTER keeps the value in brackets).' -ForegroundColor Magenta
    Write-Host ''
}
# @( )[0]: a single answer would otherwise come back as a plain string and [0] would return its first character
$SqlPattern      = '^[A-Za-z0-9][A-Za-z0-9._\\-]{0,100}(,\d{1,5})?$'
$SiteCode        = @(Read-Answer 'Site code' (Get-Setting $S.SCCMSettings.SiteCode) '^[A-Za-z0-9]{3}$' -Single)[0]
$Servers         = @(Read-Answer 'Servers to check (availability, CPU/RAM, disk, certificates)' (Get-Setting $S.SCCMSettings.strServers))
$MPServers       = @(Read-Answer 'Management Point servers' (Get-Setting $S.SCCMSettings.strMPServers) -AllowEmpty)
$ServicesServers = @(Read-Answer 'Servers to check ConfigMgr services' (Get-Setting $S.SCCMSettings.strServicesServers) -AllowEmpty)
$Provider        = @(Read-Answer 'SMS Provider / site server' (Get-Setting $S.SCCMSettings.SMSProviderServerName) -Single)[0]
$SiteDbServer    = @(Read-Answer 'Site database server (SQL)' (Get-Setting $S.SCCMSettings.SMSDBServerName) -Single)[0]
$CentralDbServer = @(Read-Answer 'Central database server (SQL)' (Get-Setting $S.CentralSettings.SCCMCentralDBServerName) $SqlPattern -Single -KeepCase)[0]
$CentralDbName   = @(Read-Answer 'Central database name' (Get-Setting $S.CentralSettings.SCCMCentralDBName) '^[A-Za-z0-9_]{1,128}$' -Single -KeepCase)[0]

$InboxViaWinRM = (Get-Setting $S.HealthCheckCustomSettings.CheckInboxRpt 'Yes') -eq 'Yes'
if ($Interactive) {
    $a = Read-Host "Read the inboxes folders through WinRM on $Provider? (Y/N) [$(if ($InboxViaWinRM) {'Y'} else {'N'})]"
    if ($a -match '^[Yy]') { $InboxViaWinRM = $true } elseif ($a -match '^[Nn]') { $InboxViaWinRM = $false }
}

# =====================================================================================================
# Settings (ConfigFile.xml, new keys optional)
# =====================================================================================================
$ProjectName   = Get-Setting $S.SCCMSettings.ProjectName 'SCCM Health Check'
$OutputName    = Get-Setting $S.SCCMSettings.OutputFileName 'ConfigMgr_Daily_Servers_Health_Check_Reports'
$CompanyName   = Get-Setting $S.SCCMSettings.CompanyName 'Bauducco'
$D             = $S.DefaultSettings
$InboxWarn     = [int](Get-Setting $D.InboxWarningCount 1000)
$InboxCrit     = [int](Get-Setting $D.InboxCriticalCount 5000)
$DiskWarnPct   = [double](Get-Setting $D.WarningDiskSpacePercentage 15)
$DiskCritPct   = [double](Get-Setting $D.CriticalDiskSpacePercentage 10)
$CheckSiteBkp  = (Get-Setting $D.CheckSiteBackup 'Yes') -eq 'Yes'
$CheckDbBkp    = (Get-Setting $D.CheckDBBackup 'Yes') -eq 'Yes'
$HistoryDays   = [Math]::Abs([int](Get-Setting $D.HistoryRpt -30))
$CpuWarn       = [double](Get-Setting $D.CpuWarningPercentage 80)
$CpuCrit       = [double](Get-Setting $D.CpuCriticalPercentage 90)
$MemWarn       = [double](Get-Setting $D.MemoryWarningPercentage 85)
$MemCrit       = [double](Get-Setting $D.MemoryCriticalPercentage 95)
$ClientWarnPct = [double](Get-Setting $D.ClientWarningPercentage 95)
$ClientCritPct = [double](Get-Setting $D.ClientCriticalPercentage 85)
$CompWarnPct   = [double](Get-Setting $D.ComplianceWarningPercentage 95)
$CompCritPct   = [double](Get-Setting $D.ComplianceCriticalPercentage 85)
$CertWarnDays  = [int](Get-Setting $D.CertWarningDays 60)
$CertCritDays  = [int](Get-Setting $D.CertCriticalDays 30)
$SupWarnHours  = [double](Get-Setting $D.SupSyncWarningHours 30)
$SupCritHours  = [double](Get-Setting $D.SupSyncCriticalHours 72)
$BackupStart   = Get-Setting $D.BackupWindowStart '21:00'
$BackupEnd     = Get-Setting $D.BackupWindowEnd '06:00'
$MpScheme      = Get-Setting $D.MPScheme 'http'
$MaxRows       = [int](Get-Setting $D.MaxRowsPerSection 200)

$HC = $S.HealthCheckCustomSettings
function Test-Flag {
    param($Node, [string]$Default = 'Yes')
    return ((Get-Setting $Node $Default) -eq 'Yes')
}
$Run = @{
    Availability = Test-Flag $HC.CheckServersAvailabilityRpt
    Resources    = Test-Flag $HC.CheckResourcesRpt
    Disk         = Test-Flag $HC.CheckServersDiskSpaceRpt
    MP           = Test-Flag $HC.CheckServersMPRpt
    SiteServices = Test-Flag $HC.CheckSiteServersServicesRpt
    SqlServices  = Test-Flag $HC.CheckSQLServerServicesRpt
    Backups      = Test-Flag $HC.CheckBackupsRpt
    Inbox        = $InboxViaWinRM
    IssueSites   = Test-Flag $HC.CheckIssueSiteServersRpt
    Components   = Test-Flag $HC.CheckCompRpt
    ClientHealth = Test-Flag $HC.CheckClientHealthRpt
    Sup          = Test-Flag $HC.CheckSupRpt
    Compliance   = Test-Flag $HC.CheckComplianceRpt
    DpStorage    = Test-Flag $HC.CheckDPStorageRpt
    Content      = Test-Flag $HC.CheckWaitingContentRpt
    Certificates = Test-Flag $HC.CheckCertificatesRpt
    Csv          = Test-Flag $HC.GenerateCSVRpt
}
$SccmServices = 'IISADMIN', 'W3SVC', 'Winmgmt', 'CcmExec', 'SMS_EXECUTIVE', 'SMS_SITE_COMPONENT_MANAGER', 'SMS_SITE_VSS_WRITER'
$SqlServices  = 'Winmgmt', 'CcmExec', 'SMS_EXECUTIVE', 'SMS_REPORTING_POINT', 'ReportServer', 'MSSQLSERVER'

# Email (same keys as v1; section is optional)
$E = $S.EmailSettings
$TriggerMail = Test-Flag $E.TriggerMail 'No'

# =====================================================================================================
# Section model: every check returns one or more sections
# =====================================================================================================
function New-Col {
    param([string]$Header, [string]$Prop, [string]$Align = 'center')
    return @{ H = $Header; P = $Prop; A = $Align }
}

function New-Section {
    param(
        [string]$Id, [string]$Title, [object[]]$Columns, [object[]]$Rows,
        [string]$Note = '', [string]$EmptyText = ''
    )
    # A check with no output gives $null, and @($null) would be one empty row: filter it out
    $clean = @($Rows | Where-Object { $null -ne $_ })
    return [pscustomobject]@{ Id = $Id; Title = $Title; Columns = $Columns; Rows = $clean; Note = $Note; EmptyText = $EmptyText }
}

function Get-Severity {
    param([string]$Status)
    switch -Regex ($Status) {
        '^(Critical|Offline|DNS_Issue|Failed)$' { return 'critical' }
        '^(Warning|WMI_Issue|WinRM_Issue)$'     { return 'warning' }
        '^Ok$'                                  { return 'ok' }
        default                                 { return 'na' }
    }
}

function ConvertTo-Status {
    # Normalizes status text coming from SQL ('OK' -> 'Ok')
    param([string]$Text)
    if ($Text -match '^\s*ok\s*$') { return 'Ok' }
    if ($Text -match '^\s*warning\s*$') { return 'Warning' }
    if ($Text -match '^\s*critical\s*$') { return 'Critical' }
    return 'NA'
}

function Get-LevelStatus {
    # Higher is worse (CPU, memory, inbox counts)
    param([double]$Value, [double]$Warn, [double]$Crit)
    if ($Value -ge $Crit) { return 'Critical' }
    if ($Value -ge $Warn) { return 'Warning' }
    return 'Ok'
}

function Get-FloorStatus {
    # Lower is worse (free space %, active client %, compliance %)
    param([double]$Value, [double]$Warn, [double]$Crit)
    if ($Value -lt $Crit) { return 'Critical' }
    if ($Value -lt $Warn) { return 'Warning' }
    return 'Ok'
}

function Format-Pct {
    param([double]$Part, [double]$Total)
    if ($Total -le 0) { return 0 }
    return [Math]::Round(($Part / $Total) * 100, 1)
}

function Invoke-Check {
    # A check that throws becomes a Warning row instead of stopping the whole report
    param([string]$Id, [string]$Title, [scriptblock]$Script)
    Write-Log "Checking: $Title"
    try {
        return @(& $Script)
    } catch {
        $msg = $_.Exception.Message
        Write-Log "$Title failed: $msg" 'ERROR'
        $row = [pscustomobject]@{ Detail = "Check could not run: $msg"; Status = 'Warning' }
        return @(New-Section -Id $Id -Title $Title -Columns @(New-Col 'Detail' 'Detail' 'left') -Rows @($row))
    }
}

# =====================================================================================================
# Server state: resolve, ping and open ONE CIM session per server (v1 repeated this 4 times)
# =====================================================================================================
$ServerState = @{}

function Get-ServerState {
    param([string]$Name)
    if ($ServerState.ContainsKey($Name)) { return $ServerState[$Name] }
    $s = [pscustomobject]@{
        Name = $Name; IP = ''; OS = 'NA'; Domain = 'NA'; Uptime = 'NA'; Status = 'Ok'; Reason = '';
        Session = $null; CpuPct = $null; MemTotalGB = $null; MemUsedPct = $null
    }
    $ServerState[$Name] = $s
    try { $ip = [System.Net.Dns]::GetHostAddresses($Name) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1 } catch { $ip = $null }
    if (-not $ip) { $s.Status = 'DNS_Issue'; $s.Reason = 'Name does not resolve in DNS'; return $s }
    $s.IP = $ip.IPAddressToString
    if (-not (Test-Connection -ComputerName $Name -Count 1 -Quiet)) { $s.Status = 'Offline'; $s.Reason = 'No ping reply'; return $s }
    try {
        $opt = New-CimSessionOption -Protocol Dcom
        $s.Session = New-CimSession -ComputerName $Name -SessionOption $opt -OperationTimeoutSec 30
        $os  = Get-CimInstance -CimSession $s.Session -ClassName Win32_OperatingSystem
        $cs  = Get-CimInstance -CimSession $s.Session -ClassName Win32_ComputerSystem
        $cpu = Get-CimInstance -CimSession $s.Session -ClassName Win32_Processor | Measure-Object -Property LoadPercentage -Average
        $s.OS = $os.Caption
        $s.Domain = $cs.Domain
        $span = (Get-Date) - $os.LastBootUpTime
        $s.Uptime = '{0}d {1}h' -f $span.Days, $span.Hours
        $s.CpuPct = [Math]::Round([double]$cpu.Average, 0)
        $s.MemTotalGB = [Math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
        $s.MemUsedPct = [Math]::Round((1 - ($os.FreePhysicalMemory / $os.TotalVisibleMemorySize)) * 100, 0)
    } catch {
        $s.Status = 'WMI_Issue'
        $s.Reason = $_.Exception.Message
    }
    return $s
}

# =====================================================================================================
# SQL helper (SELECT only, integrated security, static queries; names were validated above)
# =====================================================================================================
function Invoke-CmSql {
    param([string]$Query)
    $cs = "Server=$CentralDbServer;Database=$CentralDbName;Integrated Security=SSPI;Application Name=ConfigMgr HealthCheck;Connect Timeout=15"
    $conn = New-Object System.Data.SqlClient.SqlConnection $cs
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = 120
        $table = New-Object System.Data.DataTable
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        [void]$adapter.Fill($table)
        $out = @()
        foreach ($r in $table.Rows) {
            $o = [ordered]@{}
            foreach ($c in $table.Columns) {
                $v = $r[$c.ColumnName]
                if ($v -is [System.DBNull]) { $v = $null }
                $o[$c.ColumnName] = $v
            }
            $out += [pscustomobject]$o
        }
        return $out
    } finally {
        $conn.Close()
    }
}

function Resolve-CmObject {
    # Returns the first candidate view/table that exists in the central database.
    # If none exists, throws a message that lists similar names, so the report itself tells what to use.
    param([string[]]$Candidates, [string]$LikePattern)
    $in = ($Candidates | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ','
    $hit = @(Invoke-CmSql "SELECT name FROM sys.all_objects WHERE name IN ($in) AND type IN ('V','U')" | ForEach-Object { $_.name })
    foreach ($c in $Candidates) { if ($hit -contains $c) { return $c } }
    $like = $LikePattern -replace "'", "''"
    $near = @(Invoke-CmSql "SELECT TOP 15 name FROM sys.all_objects WHERE name LIKE '$like' AND type IN ('V','U') ORDER BY name" | ForEach-Object { $_.name })
    throw "None of these exist in ${CentralDbName}: $($Candidates -join ', '). Similar names found: $(if ($near.Count) { $near -join ', ' } else { 'none' })"
}

function Format-Date {
    param($Value)
    if ($null -eq $Value) { return '' }
    return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm')
}

# =====================================================================================================
# Checks
# =====================================================================================================
function Get-AvailabilitySection {
    $rows = foreach ($n in $Servers) {
        $st = Get-ServerState $n
        [pscustomobject]@{ Server = $n; IP = $st.IP; OS = $st.OS; Domain = $st.Domain; Uptime = $st.Uptime; Status = $st.Status; Detail = $st.Reason }
    }
    $cols = @((New-Col 'ServerName' 'Server' 'left'), (New-Col 'IP Address' 'IP'), (New-Col 'Operating System' 'OS'), (New-Col 'Domain' 'Domain'), (New-Col 'Uptime' 'Uptime'))
    return New-Section -Id 'availability' -Title 'Servers Availability' -Columns $cols -Rows $rows
}

function Get-ResourcesSection {
    $rows = foreach ($n in $Servers) {
        $st = Get-ServerState $n
        if ($st.Status -ne 'Ok') {
            [pscustomobject]@{ Server = $n; Cpu = 'NA'; MemTotal = 'NA'; MemUsed = 'NA'; Status = $st.Status; Detail = $st.Reason }
            continue
        }
        $a = Get-LevelStatus $st.CpuPct $CpuWarn $CpuCrit
        $b = Get-LevelStatus $st.MemUsedPct $MemWarn $MemCrit
        $status = 'Ok'
        if ($a -eq 'Critical' -or $b -eq 'Critical') { $status = 'Critical' } elseif ($a -eq 'Warning' -or $b -eq 'Warning') { $status = 'Warning' }
        [pscustomobject]@{ Server = $n; Cpu = "$($st.CpuPct) %"; MemTotal = $st.MemTotalGB; MemUsed = "$($st.MemUsedPct) %"; Status = $status; Detail = '' }
    }
    $cols = @((New-Col 'ServerName' 'Server' 'left'), (New-Col 'CPU now' 'Cpu'), (New-Col 'RAM Total (GB)' 'MemTotal'), (New-Col 'RAM Used' 'MemUsed'))
    return New-Section -Id 'resources' -Title 'CPU & Memory' -Columns $cols -Rows $rows -Note "CPU is a snapshot taken during the run. Warning from $CpuWarn % CPU or $MemWarn % RAM; Critical from $CpuCrit % or $MemCrit %."
}

function Get-DiskSection {
    $rows = foreach ($n in $Servers) {
        $st = Get-ServerState $n
        if ($st.Status -ne 'Ok') {
            [pscustomobject]@{ Server = $n; Drive = 'NA'; Vol = 'NA'; Total = 'NA'; Used = 'NA'; Free = 'NA'; FreePct = 'NA'; Status = $st.Status; Detail = $st.Reason }
            continue
        }
        foreach ($d in (Get-CimInstance -CimSession $st.Session -ClassName Win32_LogicalDisk -Filter 'DriveType = 3')) {
            $total = [Math]::Round($d.Size / 1GB, 2)
            $free  = [Math]::Round($d.FreeSpace / 1GB, 2)
            $pct   = Format-Pct $d.FreeSpace $d.Size
            [pscustomobject]@{ Server = $n; Drive = $d.DeviceID; Vol = $d.VolumeName; Total = $total; Used = [Math]::Round($total - $free, 2); Free = $free; FreePct = $pct; Status = (Get-FloorStatus $pct $DiskWarnPct $DiskCritPct); Detail = '' }
        }
    }
    $cols = @((New-Col 'ServerName' 'Server' 'left'), (New-Col 'Drive' 'Drive'), (New-Col 'VolName' 'Vol'), (New-Col 'Total (GB)' 'Total'), (New-Col 'Used (GB)' 'Used'), (New-Col 'Free (GB)' 'Free'), (New-Col 'Free %' 'FreePct'))
    return New-Section -Id 'disk' -Title 'Servers Disk Space' -Columns $cols -Rows $rows -Note "Warning below $DiskWarnPct % free; Critical below $DiskCritPct % free."
}

function Test-MpUrl {
    param([string]$Url)
    $r = @{ Text = ''; Code = ''; Ok = $false; Failed = $false }
    try {
        $req = [System.Net.WebRequest]::Create($Url)
        $req.Timeout = 10000
        $resp = $req.GetResponse()
        $r.Code = [int]$resp.StatusCode
        $r.Text = [string]$resp.StatusCode
        $r.Ok = ($r.Code -eq 200)
        $resp.Close()
    } catch [System.Net.WebException] {
        $r.Failed = $true
        if ($_.Exception.Response) { $r.Code = [int]$_.Exception.Response.StatusCode; $r.Text = [string]$_.Exception.Response.StatusCode }
        else { $r.Code = '-'; $r.Text = $_.Exception.Status.ToString() }
    }
    return $r
}

function Get-MpSection {
    $rows = foreach ($n in $MPServers) {
        $cert = Test-MpUrl "${MpScheme}://$n/sms_mp/.sms_aut?mpcert"
        $list = Test-MpUrl "${MpScheme}://$n/sms_mp/.sms_aut?mplist"
        $status = 'Warning'
        if ($cert.Ok -and $list.Ok) { $status = 'Ok' } elseif ($cert.Failed -or $list.Failed) { $status = 'Critical' }
        [pscustomobject]@{ Server = $n; Site = $SiteCode; CertText = $cert.Text; CertCode = $cert.Code; ListText = $list.Text; ListCode = $list.Code; Status = $status; Detail = '' }
    }
    $cols = @((New-Col 'ServerName' 'Server' 'left'), (New-Col 'Site Code' 'Site'), (New-Col 'MPCert Status' 'CertText'), (New-Col 'MPCert Code' 'CertCode'), (New-Col 'MPList Status' 'ListText'), (New-Col 'MPList Code' 'ListCode'))
    return New-Section -Id 'mp' -Title 'Management Points' -Columns $cols -Rows $rows
}

function Get-ServiceRows {
    param([string[]]$ServerNames, [string[]]$ServiceNames)
    foreach ($n in $ServerNames) {
        $st = Get-ServerState $n
        if ($st.Status -ne 'Ok') {
            [pscustomobject]@{ Server = $n; DisplayName = 'NA'; Name = 'NA'; StartMode = 'NA'; State = 'NA'; Status = $st.Status; Detail = $st.Reason }
            continue
        }
        $found = Get-CimInstance -CimSession $st.Session -ClassName Win32_Service | Where-Object { $ServiceNames -contains $_.Name }
        foreach ($svc in $found) {
            # v1 judged only the start mode: an Auto service that was stopped still showed Ok
            if ($svc.StartMode -eq 'Disabled') { $status = 'Critical' }
            elseif ($svc.StartMode -eq 'Auto' -and $svc.State -ne 'Running') { $status = 'Critical' }
            elseif ($svc.StartMode -eq 'Manual') { $status = 'Warning' }
            else { $status = 'Ok' }
            [pscustomobject]@{ Server = $n; DisplayName = $svc.DisplayName; Name = $svc.Name; StartMode = $svc.StartMode; State = $svc.State; Status = $status; Detail = '' }
        }
    }
}

function Get-ServiceColumns {
    return @((New-Col 'ServerName' 'Server' 'left'), (New-Col 'Display Name' 'DisplayName' 'left'), (New-Col 'Name' 'Name' 'left'), (New-Col 'StartMode' 'StartMode'), (New-Col 'State' 'State'))
}

function Test-EventInWindow {
    param([string]$Computer, [string]$Provider, [int]$Id, [datetime]$After, [datetime]$Before)
    try {
        return Get-WinEvent -ComputerName $Computer -FilterHashtable @{ LogName = 'Application'; ProviderName = $Provider; Id = $Id; StartTime = $After; EndTime = $Before } -MaxEvents 1 -ErrorAction Stop
    } catch {
        if ($_.FullyQualifiedErrorId -like '*NoMatchingEventsFound*') { return $null }
        throw
    }
}

function Get-BackupSection {
    $today  = (Get-Date).Date
    $after  = $today.AddDays(-1).Add([timespan]::Parse($BackupStart))
    $before = $today.Add([timespan]::Parse($BackupEnd))
    $window = '{0:yyyy-MM-dd HH:mm} to {1:yyyy-MM-dd HH:mm}' -f $after, $before
    $rows = @()
    $targets = @()
    if ($CheckSiteBkp) { $targets += , @($Provider, 'Site backup', 'SMS Server', 6833) }
    if ($CheckDbBkp)   { $targets += , @($SiteDbServer, 'Database backup', 'MSSQLSERVER', 18264) }
    foreach ($t in $targets) {
        $row = [pscustomobject]@{ Server = $t[0]; Kind = $t[1]; Window = $window; Last = ''; Result = ''; Status = 'Ok'; Detail = '' }
        try {
            $ev = Test-EventInWindow -Computer $t[0] -Provider $t[2] -Id $t[3] -After $after -Before $before
            if ($ev) { $row.Last = Format-Date $ev.TimeCreated; $row.Result = 'Success' }
            else { $row.Result = 'Not found'; $row.Status = 'Critical' }
        } catch {
            $row.Result = 'Could not read event log'; $row.Status = 'Warning'; $row.Detail = $_.Exception.Message
        }
        $rows += $row
    }
    $cols = @((New-Col 'ServerName' 'Server' 'left'), (New-Col 'Backup' 'Kind'), (New-Col 'Checked window' 'Window'), (New-Col 'Last event' 'Last'), (New-Col 'Result' 'Result'))
    return New-Section -Id 'backups' -Title 'Backups' -Columns $cols -Rows $rows
}

function Get-InboxFolderStats {
    # Same read-only logic as the remote block, for a path this machine can reach (UNC fallback)
    param([string]$Root, [string[]]$Relative)
    foreach ($r in $Relative) {
        $path = Join-Path $Root $r
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $files = @(Get-ChildItem -LiteralPath $path -File -ErrorAction SilentlyContinue)
        $sum = ($files | Measure-Object -Property Length -Sum).Sum
        [pscustomobject]@{ Folder = $r; Path = $path; Count = $files.Count; SizeMB = [Math]::Round(([double]$sum) / 1MB, 2) }
    }
}

function Get-InboxSection {
    $folders = @(
        'ccr.box', 'ccrretry.box', 'auth\ddm.box', 'auth\ddm.box\BAD_DDRS', 'auth\statesys.box\incoming', 'auth\sinv.box',
        'auth\dataldr.box', 'auth\dataldr.box\process', 'auth\dataldr.box\retry', 'auth\dataldr.box\BADMIFS',
        'schedule.box', 'schedule.box\tosend', 'schedule.box\outboxes', 'statmgr.box', 'statmgr.box\statmsgs',
        'replmgr.box', 'replmgr.box\incoming', 'replmgr.box\outbound', 'offermgr.box', 'compsumm.box', 'despoolr.box', 'despoolr.box\receive'
    )
    $cols = @((New-Col 'Folder' 'Folder' 'left'), (New-Col 'Path' 'Path' 'left'), (New-Col 'Files' 'Count'), (New-Col 'Size (MB)' 'SizeMB'))
    $method = "WinRM on $Provider"
    $found = $null
    try {
        $found = Invoke-Command -ComputerName $Provider -ArgumentList (, $folders) -ErrorAction Stop -ScriptBlock {
            param($Rel)
            $root = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\SMS\Identification' -ErrorAction Stop).'Installation Directory'
            foreach ($r in $Rel) {
                $path = Join-Path $root ('inboxes\' + $r)
                if (-not (Test-Path -LiteralPath $path)) { continue }
                $files = @(Get-ChildItem -LiteralPath $path -File -ErrorAction SilentlyContinue)
                $sum = ($files | Measure-Object -Property Length -Sum).Sum
                [pscustomobject]@{ Folder = $r; Path = $path; Count = $files.Count; SizeMB = [Math]::Round(([double]$sum) / 1MB, 2) }
            }
        }
    } catch {
        # Show the real reason in the report (it used to be only in a tooltip), then try the v1 way (UNC share)
        $winrmError = ($_.Exception.Message -replace '\s+', ' ').Trim()
        if ($winrmError.Length -gt 240) { $winrmError = $winrmError.Substring(0, 240) + '...' }
        Write-Log "Inboxes: WinRM to ${Provider} failed: ${winrmError}" 'WARN'
        $unc = "\\$Provider\SMS_$SiteCode\inboxes"
        if (Test-Path -LiteralPath $unc) {
            $found = @(Get-InboxFolderStats -Root $unc -Relative $folders)
            $method = "UNC share ${unc} (WinRM failed: ${winrmError})"
        } else {
            $row = [pscustomobject]@{ Folder = 'NA'; Path = "Not collected. WinRM error: ${winrmError} | UNC ${unc} is not accessible."; Count = 'NA'; SizeMB = 'NA'; Status = 'WinRM_Issue'; Detail = $winrmError }
            return New-Section -Id 'inboxes' -Title 'Inboxes Backlog' -Columns $cols -Rows @($row)
        }
    }
    $found = @($found | Where-Object { $null -ne $_ })
    $rows = foreach ($f in $found) {
        $f | Add-Member -NotePropertyName Status -NotePropertyValue (Get-LevelStatus $f.Count $InboxWarn $InboxCrit) -PassThru | Add-Member -NotePropertyName Detail -NotePropertyValue '' -PassThru
    }
    $missing = $folders.Count - $found.Count
    return New-Section -Id 'inboxes' -Title 'Inboxes Backlog' -Columns $cols -Rows $rows -Note "Collected via $method. Files counted per folder (not recursive). Warning above $InboxWarn files; Critical above $InboxCrit. $missing of $($folders.Count) known folders do not exist on this site."
}

function Get-IssueSitesSection {
    $q = @'
Select SiteStatus.SiteCode,
SiteInfo.ServerName,
SiteInfo.SiteName,
SiteStatus.Updated 'TimeStamp',
Case SiteInfo.Status When 1 Then 'Active' When 2 Then 'Pending' When 3 Then 'Failed' When 4 Then 'Deleted' When 5 Then 'Upgrade' Else ' ' END AS 'SiteState',
Case SiteStatus.Status When 0 Then 'OK' When 1 Then 'Warning' When 2 Then 'Critical' Else ' ' End AS 'Status'
From V_SummarizerSiteStatus SiteStatus Join v_Site SiteInfo on SiteStatus.SiteCode = SiteInfo.SiteCode where SiteInfo.Status <> 1 or SiteStatus.Status = 2 Order By SiteCode
'@
    $rows = foreach ($r in (Invoke-CmSql $q)) {
        $status = ConvertTo-Status $r.Status
        if ($status -eq 'Ok' -and $r.SiteState -ne 'Active') { $status = 'Warning' }
        [pscustomobject]@{ SiteCode = $r.SiteCode; Server = $r.ServerName; SiteName = $r.SiteName; Stamp = (Format-Date $r.TimeStamp); SiteState = $r.SiteState; Status = $status; Detail = '' }
    }
    $cols = @((New-Col 'Site Code' 'SiteCode'), (New-Col 'ServerName' 'Server' 'left'), (New-Col 'Site Name' 'SiteName' 'left'), (New-Col 'TimeStamp' 'Stamp'), (New-Col 'Site State' 'SiteState'))
    return New-Section -Id 'sites' -Title 'Site Status (issues only)' -Columns $cols -Rows $rows -EmptyText 'All sites are Active and OK.'
}

function Get-ComponentSection {
    $q = @'
SELECT distinct SiteCode,
MachineName 'ServerName',
ComponentName,
Case v_componentSummarizer.State When 0 Then 'Stopped' When 1 Then 'Started' When 2 Then 'Paused' When 3 Then 'Installing' When 4 Then 'Re-Installing' When 5 Then 'De-Installing' Else ' ' END AS 'Thread State',
Errors,
Warnings,
Infos,
Case v_componentSummarizer.Type When 0 Then 'Autostarting' When 1 Then 'Scheduled' When 2 Then 'Manual' ELSE ' ' END AS 'StartupType',
CASE AvailabilityState When 0 Then 'Online' When 3 Then 'Offline' ELSE ' ' END AS 'State',
Case v_ComponentSummarizer.Status When 0 Then 'OK' When 1 Then 'Warning' When 2 Then 'Critical' Else ' ' End As 'Status'
from v_ComponentSummarizer Where TallyInterval = '0001128000100008'
and (v_ComponentSummarizer.Status = 2 or v_ComponentSummarizer.Status = 1)
Order By ComponentName,SiteCode
'@
    # v1 selected Warning and Critical but printed Critical only; the checklist says "check for Warning/Critical", so both are shown
    $rows = foreach ($r in (Invoke-CmSql $q)) {
        [pscustomobject]@{ SiteCode = $r.SiteCode; Server = $r.ServerName; Component = $r.ComponentName; Thread = $r.'Thread State'; Errors = $r.Errors; Warns = $r.Warnings; Infos = $r.Infos; Startup = $r.StartupType; State = $r.State; Status = (ConvertTo-Status $r.Status); Detail = '' }
    }
    $rows = @($rows | Where-Object { $null -ne $_ } | Sort-Object @{ Expression = { if ($_.Status -eq 'Critical') { 0 } else { 1 } } }, Component)
    $cols = @((New-Col 'Site Code' 'SiteCode'), (New-Col 'ServerName' 'Server' 'left'), (New-Col 'Component' 'Component' 'left'), (New-Col 'Thread' 'Thread'), (New-Col 'Errors' 'Errors'), (New-Col 'Warns' 'Warns'), (New-Col 'Infos' 'Infos'), (New-Col 'Startup' 'Startup'), (New-Col 'State' 'State'))
    return New-Section -Id 'components' -Title 'Component Status (Warning / Critical)' -Columns $cols -Rows $rows -EmptyText 'No component in Warning or Critical.'
}

function Get-ClientHealthSection {
    $rows = @()
    $sum = @(Invoke-CmSql "SELECT ClientsTotal, ClientsActive, ClientsInactive, ClientsHealthy, ClientsUnhealthy, ClientsHealthUnknown FROM v_CH_ClientSummaryCurrent WHERE CollectionID = 'SMS00001'")
    if ($sum.Count -gt 0) {
        $c = $sum[0]; $t = [double]$c.ClientsTotal
        $act = Format-Pct $c.ClientsActive $t
        $hea = Format-Pct $c.ClientsHealthy $t
        $rows += [pscustomobject]@{ Metric = 'Clients in All Systems'; Count = $c.ClientsTotal; Pct = ''; Status = 'Info'; Detail = '' }
        $rows += [pscustomobject]@{ Metric = 'Active clients'; Count = $c.ClientsActive; Pct = "$act %"; Status = (Get-FloorStatus $act $ClientWarnPct $ClientCritPct); Detail = '' }
        $rows += [pscustomobject]@{ Metric = 'Inactive clients'; Count = $c.ClientsInactive; Pct = "$(Format-Pct $c.ClientsInactive $t) %"; Status = 'Info'; Detail = '' }
        $rows += [pscustomobject]@{ Metric = 'Healthy clients (client check / CCMEval)'; Count = $c.ClientsHealthy; Pct = "$hea %"; Status = (Get-FloorStatus $hea $ClientWarnPct $ClientCritPct); Detail = '' }
        $rows += [pscustomobject]@{ Metric = 'Unhealthy clients'; Count = $c.ClientsUnhealthy; Pct = "$(Format-Pct $c.ClientsUnhealthy $t) %"; Status = 'Info'; Detail = '' }
    }
    $inv = @(Invoke-CmSql @'
SELECT COUNT(*) AS Total,
SUM(CASE WHEN ws.LastHWScan >= DATEADD(day, -7, GETDATE()) THEN 1 ELSE 0 END) AS HwRecent,
SUM(CASE WHEN sw.LastScanDate >= DATEADD(day, -7, GETDATE()) THEN 1 ELSE 0 END) AS SwRecent
FROM v_R_System sys
LEFT JOIN v_GS_WORKSTATION_STATUS ws ON ws.ResourceID = sys.ResourceID
LEFT JOIN v_GS_LastSoftwareScan sw ON sw.ResourceID = sys.ResourceID
WHERE sys.Client0 = 1 AND ISNULL(sys.Obsolete0, 0) = 0 AND ISNULL(sys.Decommissioned0, 0) = 0
'@)
    if ($inv.Count -gt 0 -and [double]$inv[0].Total -gt 0) {
        $t = [double]$inv[0].Total
        $hw = Format-Pct $inv[0].HwRecent $t
        $sw = Format-Pct $inv[0].SwRecent $t
        $rows += [pscustomobject]@{ Metric = 'Hardware inventory in the last 7 days'; Count = $inv[0].HwRecent; Pct = "$hw %"; Status = (Get-FloorStatus $hw $ClientWarnPct $ClientCritPct); Detail = "of $($inv[0].Total) clients" }
        $rows += [pscustomobject]@{ Metric = 'Software inventory in the last 7 days'; Count = $inv[0].SwRecent; Pct = "$sw %"; Status = (Get-FloorStatus $sw $ClientWarnPct $ClientCritPct); Detail = "of $($inv[0].Total) clients" }
    }
    $cols = @((New-Col 'Metric' 'Metric' 'left'), (New-Col 'Clients' 'Count'), (New-Col '%' 'Pct'))
    return New-Section -Id 'clients' -Title 'Client Health' -Columns $cols -Rows $rows -Note "Warning below $ClientWarnPct %; Critical below $ClientCritPct %. Source: site database summary views."
}

function Get-SupSection {
    $states = @{ 6702 = 'Sync done'; 6703 = 'Sync failed'; 6704 = 'In progress (WSUS server)'; 6705 = 'In progress (site database)'; 6706 = 'In progress (Internet-facing WSUS)'; 6707 = 'Out of sync with upstream'; 6708 = 'Done, license terms pending' }
    # v_SUPSyncStatus does not exist in every database; the WMI class SMS_SUPSyncStatus is usually exposed as vSMS_SUPSyncStatus
    $view = Resolve-CmObject -Candidates @('vSMS_SUPSyncStatus', 'v_SUPSyncStatus', 'SUPSyncStatus') -LikePattern '%SUP%Sync%'
    $rows = foreach ($r in (Invoke-CmSql "SELECT WSUSServerName, SiteCode, LastSyncState, LastSuccessfulSyncTime, LastSyncErrorCode, LastSyncStateTime FROM $view ORDER BY SiteCode, WSUSServerName")) {
        $state = [int]$r.LastSyncState
        $text = $states[$state]; if (-not $text) { $text = "State $state" }
        $age = $null
        if ($r.LastSuccessfulSyncTime) { $age = [Math]::Round(((Get-Date) - [datetime]$r.LastSuccessfulSyncTime).TotalHours, 1) }
        if ($state -eq 6703 -or $null -eq $age) { $status = 'Critical' } else { $status = Get-LevelStatus $age $SupWarnHours $SupCritHours }
        if ($status -eq 'Ok' -and $state -notin 6702, 6704, 6705, 6706) { $status = 'Warning' }
        [pscustomobject]@{ Server = $r.WSUSServerName; Site = $r.SiteCode; State = $text; LastOk = (Format-Date $r.LastSuccessfulSyncTime); Age = $age; ErrorCode = $r.LastSyncErrorCode; Status = $status; Detail = '' }
    }
    $cols = @((New-Col 'WSUS / SUP' 'Server' 'left'), (New-Col 'Site' 'Site'), (New-Col 'Last sync state' 'State'), (New-Col 'Last successful sync' 'LastOk'), (New-Col 'Age (h)' 'Age'), (New-Col 'Error code' 'ErrorCode'))
    return New-Section -Id 'sup' -Title 'Software Update Point Synchronization' -Columns $cols -Rows $rows -Note "Warning when the last successful sync is older than $SupWarnHours h; Critical older than $SupCritHours h or failed."
}

function Get-ComplianceSection {
    # NumberTargeted is not a column of v_DeploymentSummary in this database (the other columns are), so the total is the sum of the states
    $q = @'
SELECT SoftwareName, CollectionName,
(ISNULL(NumberSuccess,0) + ISNULL(NumberInProgress,0) + ISNULL(NumberErrors,0) + ISNULL(NumberOther,0) + ISNULL(NumberUnknown,0)) AS Targeted,
NumberSuccess, NumberInProgress, NumberErrors, NumberUnknown, NumberOther
FROM v_DeploymentSummary WHERE FeatureType = 5 ORDER BY SoftwareName
'@
    $rows = foreach ($r in (Invoke-CmSql $q)) {
        if ([double]$r.Targeted -le 0) { continue }
        $pct = Format-Pct $r.NumberSuccess $r.Targeted
        [pscustomobject]@{ Name = $r.SoftwareName; Collection = $r.CollectionName; Targeted = $r.Targeted; Success = $r.NumberSuccess; InProgress = $r.NumberInProgress; Errors = $r.NumberErrors; Unknown = $r.NumberUnknown; Pct = "$pct %"; Status = (Get-FloorStatus $pct $CompWarnPct $CompCritPct); Detail = '' }
    }
    $cols = @((New-Col 'Deployment (Software Update Group)' 'Name' 'left'), (New-Col 'Collection' 'Collection' 'left'), (New-Col 'Targeted' 'Targeted'), (New-Col 'Success' 'Success'), (New-Col 'In progress' 'InProgress'), (New-Col 'Errors' 'Errors'), (New-Col 'Unknown' 'Unknown'), (New-Col 'Compliance' 'Pct'))
    return New-Section -Id 'compliance' -Title 'Patch Compliance (Software Update deployments)' -Columns $cols -Rows $rows -Note "Compliance = Success / (Success + In progress + Errors + Other + Unknown). Warning below $CompWarnPct %; Critical below $CompCritPct %."
}

function Get-NalServer {
    param([string]$Nal)
    $m = [regex]::Match($Nal, '\\\\([^\\"\]]+)')
    if ($m.Success) { return $m.Groups[1].Value.ToUpper() }
    return $Nal
}

function Get-DpStorageSection {
    $rows = foreach ($r in (Invoke-CmSql 'SELECT NALPath, Drive, BytesTotal, BytesFree, PercentFree FROM v_DistributionPointDriveInfo ORDER BY NALPath, Drive')) {
        $totalGB = [Math]::Round([double]$r.BytesTotal / 1MB, 1)   # the view stores KB
        $freeGB  = [Math]::Round([double]$r.BytesFree / 1MB, 1)
        [pscustomobject]@{ Server = (Get-NalServer $r.NALPath); Drive = $r.Drive; Total = $totalGB; Free = $freeGB; FreePct = $r.PercentFree; Status = (Get-FloorStatus ([double]$r.PercentFree) $DiskWarnPct $DiskCritPct); Detail = '' }
    }
    $cols = @((New-Col 'Distribution Point' 'Server' 'left'), (New-Col 'Drive' 'Drive'), (New-Col 'Total (GB)' 'Total'), (New-Col 'Free (GB)' 'Free'), (New-Col 'Free %' 'FreePct'))
    return New-Section -Id 'dpstorage' -Title 'Distribution Point Storage' -Columns $cols -Rows $rows -Note "Warning below $DiskWarnPct % free; Critical below $DiskCritPct % free."
}

function Get-ContentSections {
    $q = @'
Select
SubString(dp.ServerNALPath, CHARINDEX('\\', dp.ServerNALPath)+2,(CHARINDEX('"]', dp.ServerNALPath) - CHARINDEX('\\', dp.ServerNALPath))-3) as ServerName,
dp.SiteCode as 'SiteCode',
dp.PackageID as 'PackageID',
p.Name as 'PackageName',
P.SourceVersion as 'SourceVersion',
P.LastRefreshTime as 'LastRefreshTime',
stat.InstallStatus as 'InstallStatus'
from v_DistributionPoint dp left join v_PackageStatusDistPointsSumm stat on
dp.ServerNALPath=stat.ServerNALPath and dp.PackageID=stat.PackageID
left join v_PackageStatus pstat on dp.ServerNALPath=pstat.PkgServer and
dp.PackageID=pstat.PackageID left outer join v_Package p on dp.packageid = p.packageid
where stat.InstallStatus not in ('Package Installation complete')
ORDER BY 1
'@
    $data = @(Invoke-CmSql $q)
    $detail = foreach ($r in $data) {
        [pscustomobject]@{ Server = $r.ServerName; Site = $r.SiteCode; PackageID = $r.PackageID; Package = $r.PackageName; Ver = $r.SourceVersion; Refresh = (Format-Date $r.LastRefreshTime); Install = $r.InstallStatus; Status = 'Warning'; Detail = '' }
    }
    $perDp = foreach ($g in ($data | Group-Object -Property ServerName)) {
        [pscustomobject]@{ Server = $g.Name; Waiting = $g.Count; Status = 'Warning'; Detail = '' }
    }
    $perDp = @($perDp | Sort-Object Waiting -Descending)
    $colsDp = @((New-Col 'Distribution Point' 'Server' 'left'), (New-Col 'Packages not complete' 'Waiting'))
    $colsDetail = @((New-Col 'Distribution Point' 'Server' 'left'), (New-Col 'Site' 'Site'), (New-Col 'PackageID' 'PackageID'), (New-Col 'Package' 'Package' 'left'), (New-Col 'Source Ver' 'Ver'), (New-Col 'Last refresh' 'Refresh'), (New-Col 'Install status' 'Install'))
    New-Section -Id 'contentdp' -Title 'Content Distribution by Distribution Point' -Columns $colsDp -Rows $perDp -EmptyText 'No package is waiting to be distributed.'
    New-Section -Id 'content' -Title 'Content Distribution Detail (not complete)' -Columns $colsDetail -Rows $detail -EmptyText 'No package is waiting to be distributed.'
}

function Get-CertificateSection {
    $names = @($Servers | Where-Object { (Get-ServerState $_).Status -eq 'Ok' })
    $cols = @((New-Col 'ServerName' 'Server' 'left'), (New-Col 'Subject' 'Subject' 'left'), (New-Col 'Expires' 'Expires'), (New-Col 'Days left' 'Days'))
    $rows = @()
    $errs = @()
    if ($names.Count -gt 0) {
        $certs = Invoke-Command -ComputerName $names -ErrorAction SilentlyContinue -ErrorVariable errs -ScriptBlock {
            Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.HasPrivateKey } | ForEach-Object {
                [pscustomobject]@{ Subject = $_.Subject; NotAfter = $_.NotAfter; Thumbprint = $_.Thumbprint }
            }
        }
        foreach ($c in @($certs)) {
            $days = [Math]::Floor(($c.NotAfter - (Get-Date)).TotalDays)
            if ($days -lt -30) { continue }   # long-expired leftovers are ignored; only recently expired ones are shown
            if ($days -le $CertCritDays) { $status = 'Critical' } elseif ($days -le $CertWarnDays) { $status = 'Warning' } else { continue }
            $rows += [pscustomobject]@{ Server = $c.PSComputerName.ToUpper(); Subject = $c.Subject; Expires = (Format-Date $c.NotAfter); Days = $days; Status = $status; Detail = $c.Thumbprint }
        }
    }
    foreach ($e in @($errs)) {
        $target = $e.TargetObject; if (-not $target) { $target = 'unknown' }
        $rows += [pscustomobject]@{ Server = "$target".ToUpper(); Subject = 'WinRM failed, certificates not read'; Expires = 'NA'; Days = 'NA'; Status = 'WinRM_Issue'; Detail = $e.Exception.Message }
    }
    $rows = @($rows | Sort-Object { if ($_.Days -is [int] -or $_.Days -is [double]) { $_.Days } else { 99999 } })
    return New-Section -Id 'certs' -Title 'Certificate Expiration (LocalMachine\My, with private key)' -Columns $cols -Rows $rows -EmptyText "No certificate expires within $CertWarnDays days." -Note "Warning within $CertWarnDays days; Critical within $CertCritDays days or recently expired (last 30 days). Older expired leftovers are ignored."
}

# =====================================================================================================
# Rendering (HTML + CSV). Palette: Accenture colors only, no logo.
# =====================================================================================================
$Css = @'
body { margin:0; background:#F2F2F2; color:#000000; font-family:'Segoe UI', Arial, Helvetica, sans-serif; font-size:13px; line-height:1.4; }
.wrap { width:100%; max-width:1180px; margin:0 auto; background:#FFFFFF; }
.hdr { background:#000000; border-top:6px solid #A100FF; padding:22px 28px 18px 28px; }
.hdr h1 { margin:0; font-size:22px; font-weight:600; color:#FFFFFF; }
.hdr .sub { margin-top:4px; font-size:12px; color:#D2A5FF; }
.banner { padding:11px 28px; font-size:14px; font-weight:600; color:#FFFFFF; }
.banner.critical { background:#C62828; }
.banner.warning { background:#FFB000; color:#000000; }
.banner.ok { background:#2E7D32; }
table.kpis { width:100%; border-collapse:collapse; }
table.kpis td { width:25%; text-align:center; padding:14px 6px; background:#F5EBFF; border-right:2px solid #FFFFFF; }
table.kpis td b { display:block; font-size:28px; color:#460073; }
table.kpis td span { font-size:11px; letter-spacing:.06em; text-transform:uppercase; color:#66646E; }
.pad { padding:6px 28px 18px 28px; }
.sec-title { background:#7500C0; color:#FFFFFF; font-size:14px; font-weight:600; padding:9px 12px; margin-top:22px; }
.note { font-size:11px; color:#66646E; padding:5px 2px 6px 2px; }
table.grid { width:100%; border-collapse:collapse; }
table.grid th { background:#F2F2F2; color:#000000; font-size:12px; text-align:center; padding:7px 8px; border-bottom:2px solid #A100FF; }
table.grid td { padding:6px 8px; text-align:center; border-bottom:1px solid #E3E3E3; }
table.grid td.l { text-align:left; }
table.grid td.n, table.grid th.n { width:34px; color:#66646E; }
table.grid tr.alt td { background:#FAF5FF; }
table.grid td.st { font-weight:600; color:#FFFFFF; width:84px; }
table.grid td.st.ok { background:#2E7D32; }
table.grid td.st.warning { background:#FFB000; color:#000000; }
table.grid td.st.critical { background:#C62828; }
table.grid td.st.na { background:#66646E; }
td.clear { text-align:left; padding:10px 12px; color:#2E7D32; font-weight:600; border-bottom:1px solid #E3E3E3; }
.more { font-size:11px; color:#66646E; padding:6px 2px; }
table.glance td.l a { color:#7500C0; text-decoration:none; font-weight:600; }
.foot { background:#000000; color:#D2A5FF; text-align:center; font-size:11px; padding:14px 28px; border-bottom:6px solid #A100FF; }
.legend span { display:inline-block; padding:3px 10px; margin-right:6px; font-size:11px; font-weight:600; color:#FFFFFF; }
.legend .ok { background:#2E7D32; } .legend .warning { background:#FFB000; color:#000000; } .legend .critical { background:#C62828; } .legend .na { background:#66646E; }
'@

function Enc {
    param($Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-SectionSeverity {
    param($Section)
    if (@($Section.Rows).Count -eq 0) { if ($Section.EmptyText) { return 'ok' } else { return 'na' } }
    $rank = @{ critical = 3; warning = 2; ok = 1; na = 0 }
    $worst = 'na'
    foreach ($r in $Section.Rows) { $sev = Get-Severity $r.Status; if ($rank[$sev] -gt $rank[$worst]) { $worst = $sev } }
    return $worst
}

function ConvertTo-SectionHtml {
    param($Section)
    $sb = New-Object System.Text.StringBuilder
    $rows = @($Section.Rows)
    $span = @($Section.Columns).Count + 2
    [void]$sb.AppendLine("<a id='sec-$($Section.Id)'></a><div class='sec-title'>$(Enc $Section.Title)</div>")
    if ($Section.Note) { [void]$sb.AppendLine("<div class='note'>$(Enc $Section.Note)</div>") }
    [void]$sb.Append("<table class='grid'><thead><tr><th class='n'>#</th>")
    foreach ($c in $Section.Columns) { [void]$sb.Append("<th>$(Enc $c.H)</th>") }
    [void]$sb.AppendLine('<th>Status</th></tr></thead><tbody>')
    if ($rows.Count -eq 0) {
        $text = $Section.EmptyText; if (-not $text) { $text = 'No data returned.' }
        $cls = 'clear'; if (-not $Section.EmptyText) { $cls = 'l' }
        [void]$sb.AppendLine("<tr><td class='$cls' colspan='$span'>$(Enc $text)</td></tr>")
    }
    $i = 0
    foreach ($row in ($rows | Select-Object -First $MaxRows)) {
        $i++
        $alt = ''; if ($i % 2 -eq 0) { $alt = " class='alt'" }
        [void]$sb.Append("<tr$alt><td class='n'>$i</td>")
        foreach ($c in $Section.Columns) {
            $cls = ''; if ($c.A -eq 'left') { $cls = " class='l'" }
            [void]$sb.Append("<td$cls>$(Enc $row.($c.P))</td>")
        }
        $tip = ''; if ($row.Detail) { $tip = " title='$(Enc $row.Detail)'" }
        [void]$sb.AppendLine("<td class='st $(Get-Severity $row.Status)'$tip>$(Enc $row.Status)</td></tr>")
    }
    [void]$sb.AppendLine('</tbody></table>')
    if ($rows.Count -gt $MaxRows) { [void]$sb.AppendLine("<div class='more'>Showing the first $MaxRows of $($rows.Count) rows. The CSV has all of them.</div>") }
    return $sb.ToString()
}

function ConvertTo-ReportHtml {
    param($Sections, [string]$Title, [string]$SubTitle)
    $count = @{ critical = 0; warning = 0; ok = 0; na = 0 }
    foreach ($sec in $Sections) {
        if (@($sec.Rows).Count -eq 0 -and $sec.EmptyText) { $count.ok++ }
        foreach ($r in $sec.Rows) { $count[(Get-Severity $r.Status)]++ }
    }
    if ($count.critical -gt 0) { $overall = 'critical'; $banner = "ATTENTION: $($count.critical) critical and $($count.warning) warning item(s) need review." }
    elseif ($count.warning -gt 0) { $overall = 'warning'; $banner = "$($count.warning) warning item(s) need review. Nothing critical." }
    else { $overall = 'ok'; $banner = 'All checks are OK.' }

    $glance = New-Object System.Text.StringBuilder
    [void]$glance.AppendLine("<div class='sec-title'>At a glance</div><table class='grid glance'><thead><tr><th class='n'>#</th><th>Check</th><th>Critical</th><th>Warning</th><th>Status</th></tr></thead><tbody>")
    $i = 0
    foreach ($sec in $Sections) {
        $i++
        $crit = @($sec.Rows | Where-Object { (Get-Severity $_.Status) -eq 'critical' }).Count
        $warn = @($sec.Rows | Where-Object { (Get-Severity $_.Status) -eq 'warning' }).Count
        $sev = Get-SectionSeverity $sec
        $label = @{ critical = 'Critical'; warning = 'Warning'; ok = 'Ok'; na = 'NA' }[$sev]
        $alt = ''; if ($i % 2 -eq 0) { $alt = " class='alt'" }
        [void]$glance.AppendLine("<tr$alt><td class='n'>$i</td><td class='l'><a href='#sec-$($sec.Id)'>$(Enc $sec.Title)</a></td><td>$crit</td><td>$warn</td><td class='st $sev'>$label</td></tr>")
    }
    [void]$glance.AppendLine('</tbody></table>')

    $body = New-Object System.Text.StringBuilder
    foreach ($sec in $Sections) { [void]$body.Append((ConvertTo-SectionHtml $sec)) }

    $generated = '{0} on {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $env:COMPUTERNAME
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$(Enc $Title)</title>
<style>
$Css
</style>
</head>
<body>
<div class="wrap">
<div class="hdr"><h1>$(Enc $Title)</h1><div class="sub">$(Enc $SubTitle)</div></div>
<div class="banner $overall">$(Enc $banner)</div>
<table class="kpis"><tr>
<td><b>$($count.critical)</b><span>Critical</span></td>
<td><b>$($count.warning)</b><span>Warning</span></td>
<td><b>$($count.ok)</b><span>Ok</span></td>
<td><b>$($count.na)</b><span>Info / NA</span></td>
</tr></table>
<div class="pad">
$($glance.ToString())
$($body.ToString())
<div class="note" style="margin-top:18px"><span class="legend"><span class="ok">Ok</span><span class="warning">Warning</span><span class="critical">Critical</span><span class="na">Info / NA</span></span></div>
</div>
<div class="foot">$(Enc $CompanyName) - internal use only - read-only report generated $(Enc $generated) - script v$ScriptVersion</div>
</div>
</body>
</html>
"@
}

function Export-SectionsCsv {
    param($Sections, [string]$Path)
    $lines = New-Object System.Collections.ArrayList
    foreach ($sec in $Sections) {
        [void]$lines.Add($sec.Title)
        if (@($sec.Rows).Count -eq 0) { [void]$lines.Add('(nothing to report)'); [void]$lines.Add(''); continue }
        $props = @()
        foreach ($c in $sec.Columns) { $p = $c.P; $props += @{ Name = $c.H; Expression = { $_.$p }.GetNewClosure() } }
        $props += @{ Name = 'Status'; Expression = { $_.Status } }
        foreach ($l in ($sec.Rows | Select-Object $props | ConvertTo-Csv -NoTypeInformation)) { [void]$lines.Add($l) }
        [void]$lines.Add('')
    }
    Set-Content -LiteralPath $Path -Value $lines -Encoding UTF8
}

# =====================================================================================================
# Main
# =====================================================================================================
$ReportHtml = Join-Path $OutputPath "$OutputName.html"
$ReportCsv  = Join-Path $OutputPath "$OutputName.csv"
$HistoryDir = Join-Path $OutputPath 'Historical_Reports'
if (-not (Test-Path -LiteralPath $HistoryDir)) { New-Item -Path $HistoryDir -ItemType Directory -Force | Out-Null }

Write-Log "***** Start v$ScriptVersion (interactive=$Interactive) site=$SiteCode central=$CentralDbServer/$CentralDbName *****"
$Sections = New-Object System.Collections.ArrayList
function Add-Sections { param($Items) foreach ($x in @($Items)) { if ($x) { [void]$Sections.Add($x) } } }

try {
    if ($Run.Availability) { Add-Sections (Invoke-Check 'availability' 'Servers Availability' { Get-AvailabilitySection }) }
    if ($Run.Resources)    { Add-Sections (Invoke-Check 'resources' 'CPU & Memory' { Get-ResourcesSection }) }
    if ($Run.Disk)         { Add-Sections (Invoke-Check 'disk' 'Servers Disk Space' { Get-DiskSection }) }
    if ($Run.MP)           { Add-Sections (Invoke-Check 'mp' 'Management Points' { Get-MpSection }) }
    if ($Run.SiteServices) {
        Add-Sections (Invoke-Check 'services' 'ConfigMgr Services' {
            New-Section -Id 'services' -Title 'ConfigMgr Services' -Columns (Get-ServiceColumns) -Rows @(Get-ServiceRows $ServicesServers $SccmServices) -Note 'Critical: disabled, or Auto but not running. Warning: Manual start mode.'
        })
    }
    if ($Run.SqlServices) {
        Add-Sections (Invoke-Check 'sqlservices' 'SQL Server Services' {
            New-Section -Id 'sqlservices' -Title 'SQL Server Services' -Columns (Get-ServiceColumns) -Rows @(Get-ServiceRows @($SiteDbServer) $SqlServices)
        })
    }
    if ($Run.Backups)      { Add-Sections (Invoke-Check 'backups' 'Backups' { Get-BackupSection }) }
    if ($Run.Inbox)        { Add-Sections (Invoke-Check 'inboxes' 'Inboxes Backlog' { Get-InboxSection }) }
    if ($Run.IssueSites)   { Add-Sections (Invoke-Check 'sites' 'Site Status' { Get-IssueSitesSection }) }
    if ($Run.Components)   { Add-Sections (Invoke-Check 'components' 'Component Status' { Get-ComponentSection }) }
    if ($Run.ClientHealth) { Add-Sections (Invoke-Check 'clients' 'Client Health' { Get-ClientHealthSection }) }
    if ($Run.Sup)          { Add-Sections (Invoke-Check 'sup' 'Software Update Point Synchronization' { Get-SupSection }) }
    if ($Run.Compliance)   { Add-Sections (Invoke-Check 'compliance' 'Patch Compliance' { Get-ComplianceSection }) }
    if ($Run.DpStorage)    { Add-Sections (Invoke-Check 'dpstorage' 'Distribution Point Storage' { Get-DpStorageSection }) }
    if ($Run.Content)      { Add-Sections (Invoke-Check 'content' 'Content Distribution' { Get-ContentSections }) }
    if ($Run.Certificates) { Add-Sections (Invoke-Check 'certs' 'Certificate Expiration' { Get-CertificateSection }) }

    $title = "$ProjectName - ConfigMgr Daily Servers Health Check - $(Get-Date -Format 'yyyy-MM-dd')"
    $sub = "Site $SiteCode  |  Provider $Provider  |  Site DB $SiteDbServer  |  Central DB $CentralDbServer\$CentralDbName  |  $(Get-Date -Format F)"
    $html = ConvertTo-ReportHtml -Sections $Sections -Title $title -SubTitle $sub

    # Archive the previous report (only our own two files; v1 deleted every *.html and *.csv in the folder)
    foreach ($ext in 'html', 'csv') {
        $old = Join-Path $OutputPath "$OutputName.$ext"
        if (Test-Path -LiteralPath $old) {
            $stamp = (Get-Item -LiteralPath $old).LastWriteTime.ToString('yyyy-MM-dd_HH-mm')
            Move-Item -LiteralPath $old -Destination (Join-Path $HistoryDir "$OutputName-$stamp.$ext") -Force
        }
    }
    Get-ChildItem -LiteralPath $HistoryDir -File | Where-Object { $_.Name -like "$OutputName-*" -and $_.LastWriteTime -lt (Get-Date).AddDays(-$HistoryDays) } | Remove-Item -Force

    [System.IO.File]::WriteAllText($ReportHtml, $html, (New-Object System.Text.UTF8Encoding($true)))
    if ($Run.Csv) { Export-SectionsCsv -Sections $Sections -Path $ReportCsv }
    Write-Log "Report written: $ReportHtml"

    # Email (same keys as v1). Subject starts with the overall status.
    if ($TriggerMail) {
        $worst = 'OK'
        foreach ($sec in $Sections) {
            $sev = Get-SectionSeverity $sec
            if ($sev -eq 'critical') { $worst = 'CRITICAL'; break }
            if ($sev -eq 'warning') { $worst = 'WARNING' }
        }
        $smtpServer = Get-Setting $E.SMTPServer
        $from = Get-Setting $E.FromAddress 'sccmhealthcheckalert@domainname.com'
        $to  = @((Get-Setting $E.ToAddress)  -split '[,;\s]+' | Where-Object { $_ })
        $cc  = @((Get-Setting $E.CCAddress)  -split '[,;\s]+' | Where-Object { $_ })
        $bcc = @((Get-Setting $E.BCCAddress) -split '[,;\s]+' | Where-Object { $_ })
        if (-not $smtpServer -or $to.Count -eq 0) {
            Write-Log 'Email skipped: SMTPServer or ToAddress is empty in ConfigFile.xml' 'WARN'
        } else {
            $msg = $null; $smtp = $null
            try {
                $msg = New-Object System.Net.Mail.MailMessage
                $msg.From = $from
                foreach ($x in $to)  { $msg.To.Add($x) }
                foreach ($x in $cc)  { $msg.CC.Add($x) }
                foreach ($x in $bcc) { $msg.Bcc.Add($x) }
                $msg.IsBodyHtml = $true
                $msg.Subject = "[$worst] $title"
                $msg.Body = $html
                [void]$msg.Attachments.Add((New-Object System.Net.Mail.Attachment($ReportHtml)))
                if ($Run.Csv) { [void]$msg.Attachments.Add((New-Object System.Net.Mail.Attachment($ReportCsv))) }
                $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer)
                $smtp.Send($msg)
                Write-Log "Email sent to $($to -join ', ') with status $worst"
            } catch {
                Write-Log "Email failed: $($_.Exception.Message)" 'ERROR'
            } finally {
                if ($msg) { $msg.Dispose() }
                if ($smtp) { $smtp.Dispose() }
            }
        }
    } else {
        Write-Log 'Email skipped: TriggerMail is not Yes (v1 had no EmailSettings section in this ConfigFile.xml)'
    }
} finally {
    foreach ($st in $ServerState.Values) { if ($st.Session) { Remove-CimSession -CimSession $st.Session -ErrorAction SilentlyContinue } }
    Write-Log '***** End *****'
}
