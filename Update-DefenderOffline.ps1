<#
.SYNOPSIS
    Update-DefenderOffline.ps1 – Fully Automated Offline Microsoft Defender Antivirus Definitions Update
    Designed for air-gapped, high-security, or segmented networks.

.DESCRIPTION
    This script updates Microsoft Defender antivirus definitions on systems without internet access
    using the official mpam-fe.exe (x64) package. It leverages PowerShell Remoting (WinRM) to copy
    and silently install the update, verifies success, collects logs, and generates beautiful HTML reports.

    Key Features:
     • PowerShell 7+ = automatic parallel processing (up to 10x faster)
     • PowerShell 5.1 = safe serial fallback
     • Real-time progress monitoring in both serial and parallel modes
     • Automatic hosts.conf creation from Active Directory if missing
     • Full dual logging: color console + timestamped file
     • Organized output: .\Logs\ and .\Reports\ (configurable)
     • Optional remote log collection to network share
     • Optional HTML email with report attached
     • Dry-run mode (-WhatIfMode)
     • Fully compatible with gMSA, service accounts, and scheduled tasks
     • Administrative privilege validation

.HOW TO CREATE SMTP CREDENTIALS (for -SendEmail)

    Method 1 - Use the built-in helper function:
    .\Update-DefenderOffline.ps1 -SaveSmtpCredential

    Method 2 - Manual credential creation:
    Run this once interactively to securely save credentials:

    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $cred = Get-Credential -Message "Enter SMTP username and password (e.g. smtpuser@contoso.com)"
    $cred | Export-Clixml -Path (Join-Path $scriptDir "Config\SmtpCredential.xml")

    Then use in your script call:
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $smtpCred = Import-Clixml (Join-Path $scriptDir "Config\SmtpCredential.xml")
    .\Update-DefenderOffline.ps1 ... -SendEmail -SmtpCredential $smtpCred

    The .xml file is encrypted per-user + per-machine and safe for scheduled tasks/gMSA.

.USAGE EXAMPLES

    1. First run (auto-creates hosts.conf from AD):
       .\Update-DefenderOffline.ps1 -SourceSharePath "\\fileserver\DefenderUpdates" -MpamFileName "mpam-feX64.exe"

    2. Weekly run:
       .\Update-DefenderOffline.ps1 -SourceSharePath "\\fs\updates" -MpamFileName "mpam-feX64_1.405.9999.0.exe"

    3. Save SMTP credentials (one-time setup):
       .\Update-DefenderOffline.ps1 -SaveSmtpCredential

    4. With email and secure credentials:
       $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
       $smtpCred = Import-Clixml (Join-Path $scriptDir "Config\SmtpCredential.xml")
       .\Update-DefenderOffline.ps1 `
         -SourceSharePath "\\fs\updates" `
         -MpamFileName "mpam-feX64.exe" `
         -SendEmail `
         -SmtpServer "smtp.contoso.com" -SmtpPort 587 -SmtpUseSsl `
         -From "defender@contoso.com" -To "security@contoso.com" `
         -SmtpCredential $smtpCred

    5. Dry-run (no changes):
       .\Update-DefenderOffline.ps1 ... -WhatIfMode

    6. Max speed on PowerShell 7+:
       .\Update-DefenderOffline.ps1 ... -ParallelThreads 24

.NOTES
    File Name      : Update-DefenderOffline.ps1
    Author         : Kismet Agbasi (GitHub: https://github.com/kismetgerald | Email: KismetG17@gmail.com)
    Ai Contributors: ClaudeAI, and Grok
    Prerequisite   : • PowerShell 5.1 or 7+ (7+ strongly recommended for performance)
                     • WinRM enabled on all target computers (port 5985)
                     • Administrative privileges on target systems
                     • Network share containing the latest mpam-fe.exe (x64)
                     • (Optional) Domain-joined machine or RSAT:ActiveDirectory for auto hosts.conf generation
    
    Version        : 0.0.4
    Created        : 2025-11-27
    Last Updated   : 2026-01-08

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [string[]]$ComputerName,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$SourceSharePath,

    [Parameter(Mandatory = $true)]
    [string]$MpamFileName,

    [string]$TempFolderOnTarget = 'C:\Temp\Update-DefenderOffline',

    [string]$LogSharePath,

    # Email Options
    [switch]$SendEmail,
    [string]$SmtpServer,
    [int]$SmtpPort = 25,
    [string]$From = 'DefenderUpdate@contoso.com',
    [string[]]$To,
    [switch]$SmtpUseSsl,
    [pscredential]$SmtpCredential,

    # Safety & Testing
    [switch]$WhatIfMode,

    # Output Folders
    [string]$LogPath,
    [string]$ReportPath,

    # Performance (PowerShell 7+ only)
    [ValidateRange(1, 32)]
    [int]$ParallelThreads = 16,

    # Credential Management Helper
    [switch]$SaveSmtpCredential
)

# ===================================================================
# Script Initialization
# ===================================================================
$ScriptVersion = "0.0.1"
$ScriptStartTime = Get-Date
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$HostsFile = Join-Path $ScriptDir 'hosts.conf'

# ===================================================================
# Credential Management Helper Mode
# ===================================================================
if ($SaveSmtpCredential) {
    Write-Host "`n=== SMTP Credential Setup ===" -ForegroundColor Cyan
    Write-Host "This will securely save your SMTP credentials for email notifications.`n" -ForegroundColor White

    $configDir = Join-Path $ScriptDir 'Config'
    if (-not (Test-Path $configDir)) {
        New-Item -Path $configDir -ItemType Directory -Force | Out-Null
        Write-Host "Created Config directory: $configDir" -ForegroundColor Green
    }

    $credPath = Join-Path $configDir 'SmtpCredential.xml'

    try {
        $cred = Get-Credential -Message "Enter SMTP username and password (e.g. smtp-user@contoso.com)"

        if ($cred) {
            $cred | Export-Clixml -Path $credPath -Force
            Write-Host "`nCredentials saved successfully!" -ForegroundColor Green
            Write-Host "Location: $credPath" -ForegroundColor Cyan
            Write-Host "`nTo use these credentials, load them in your script:" -ForegroundColor Yellow
            Write-Host "  `$smtpCred = Import-Clixml '$credPath'" -ForegroundColor White
            Write-Host "  .\Update-DefenderOffline.ps1 -SendEmail -SmtpCredential `$smtpCred ...`n" -ForegroundColor White
        }
        else {
            Write-Host "`nCredential creation cancelled." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "`nERROR: Failed to save credentials: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    exit 0
}

# Check for administrative privileges
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "ERROR: This script requires administrative privileges. Please run as Administrator."
    throw "Insufficient privileges. Run PowerShell as Administrator and try again."
}

# Default output folders
if (-not $LogPath)    { $LogPath    = Join-Path $env:SystemDrive 'Logs' }
if (-not $ReportPath) { $ReportPath = Join-Path $ScriptDir 'Reports' }
foreach ($p in $LogPath, $ReportPath) {
    if (-not (Test-Path $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null }
}

$LogFile = Join-Path $LogPath "Update-DefenderOffline_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Thread-safe logging mutex
$Global:LogMutex = [System.Threading.Mutex]::new($false, "DefenderUpdateLogMutex")

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','HEADER')]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp [$Level] $Message"

    try {
        $Global:LogMutex.WaitOne() | Out-Null
        $line | Out-File -FilePath $LogFile -Append -Encoding UTF8
    }
    finally {
        $Global:LogMutex.ReleaseMutex()
    }

    if ([System.Threading.Thread]::CurrentThread.ManagedThreadId -eq 1) {
        switch ($Level) {
            'INFO'    { Write-Host $line -ForegroundColor Cyan }
            'WARN'    { Write-Host $line -ForegroundColor Yellow }
            'ERROR'   { Write-Host $line -ForegroundColor Red }
            'SUCCESS' { Write-Host $line -ForegroundColor Green }
            'HEADER'  { Write-Host $line -ForegroundColor Magenta }
            default   { Write-Host $line }
        }
    }
}

Write-Log "=== Microsoft Defender Offline Update v$ScriptVersion ===" 'HEADER'
Write-Log "Started          : $(Get-Date)"
Write-Log "PowerShell       : $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
Write-Log "Parallel Mode    : $(if ($PSVersionTable.PSVersion.Major -ge 7) { "ENABLED ($ParallelThreads threads)" } else { "DISABLED (PS 5.1)" })"
Write-Log "Log File         : $LogFile"
Write-Log "Report Folder    : $ReportPath"
Write-Log "WhatIf Mode      : $WhatIfMode"
if ($SendEmail) { Write-Log "Email enabled    : Yes$(if ($SmtpCredential) { ' (credentials provided)' } else { '' })" }

# ===================================================================
# Resolve Target Computers
# ===================================================================
function Resolve-TargetComputers {
    if ($ComputerName) {
        Write-Log "Using manually provided list ($($ComputerName.Count) computers)" 'INFO'
        return $ComputerName | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim().ToUpper() }
    }

    if (Test-Path $HostsFile) {
        $list = Get-Content $HostsFile | Where-Object { $_ -notmatch '^\s*#|^\s*$' } | ForEach-Object { $_.Trim().ToUpper() }
        Write-Log "Loaded $($list.Count) computers from hosts.conf" 'SUCCESS'
        return $list
    }

    Write-Log "hosts.conf not found → querying Active Directory..." 'WARN'

    try {
        if (Get-Module -ListAvailable ActiveDirectory -ErrorAction SilentlyContinue) {
            Import-Module ActiveDirectory
            $computers = Get-ADComputer -Filter 'OperatingSystem -like "*Windows*" -and Enabled -eq $true' -Properties Name |
                         Sort-Object Name | Select-Object -ExpandProperty Name
        } else {
            $domain = (Get-CimInstance Win32_ComputerSystem).Domain
            $searcher = [adsisearcher]"(&(objectCategory=computer)(operatingSystem=*Windows*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
            $searcher.SearchRoot = "LDAP://$domain"
            $computers = $searcher.FindAll() | ForEach-Object { $_.Properties.name[0] } | Sort-Object
        }

        $header = @"
# =============================================================================
# hosts.conf – AUTO-GENERATED by Update-DefenderOffline.ps1 v$ScriptVersion
# Generated      : $(Get-Date)
# Domain         : $((Get-CimInstance Win32_ComputerSystem).Domain)
# Total Systems  : $($computers.Count)
# Edit this file to exclude lab/test machines if needed
# =============================================================================

"@
        ($header + ($computers -join "`r`n")) | Out-File -FilePath $HostsFile -Encoding UTF8 -Force
        Write-Log "Created hosts.conf with $($computers.Count) computers" 'SUCCESS'
        return $computers
    }
    catch {
        Write-Log "Failed to query AD: $($_.Exception.Message)" 'ERROR'
        throw "Cannot proceed without target list. Create hosts.conf manually."
    }
}

$TargetComputers = Resolve-TargetComputers
if (-not $TargetComputers) { Write-Log "No target computers found!" 'ERROR'; return }
Write-Log "Will process $($TargetComputers.Count) computers" 'HEADER'

# ===================================================================
# Detect Latest mpam-fe.exe File
# ===================================================================

function Get-LatestMpamFile {
    param([string]$Root)

    $files = Get-ChildItem -Path $Root -Recurse -Filter "mpam-fe*.exe" -ErrorAction SilentlyContinue

    if (-not $files) {
        throw "No mpam-fe*.exe files found under $Root"
    }

    # Pick the newest file by LastWriteTime
    $latest = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    return [pscustomobject]@{
        File    = $latest.FullName
        Version = 'Unknown'  # We will detect version after install
    }
}

# ===================================================================
# Defender Update Function
# ===================================================================
function Invoke-DefenderUpdate {
    param(
        [string]$Computer,
        [string]$SourceFile,
        [string]$TempFolderOnTarget,
        [string]$MpamFileName,
        [switch]$WhatIfMode,
        [string]$LogSharePath
    )

    $result = [pscustomobject]@{
        ComputerName = $Computer
        Status       = 'Unknown'
        OldVersion   = ''
        NewVersion   = ''
        DurationSec  = 0
        Details      = ''
        Attempt      = 1
        Timeout      = $false
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        if ($WhatIfMode) {
            $result.Status = 'WhatIf'
            $result.Details = 'Dry-run mode – no changes made'
            return $result
        }

        if (-not (Test-NetConnection -ComputerName $Computer -Port 5985 -InformationLevel Quiet -WarningAction SilentlyContinue)) {
            throw "WinRM (5985) not reachable"
        }

        $session = New-PSSession -ComputerName $Computer -ErrorAction Stop

        $currentVer = Invoke-Command -Session $session -ScriptBlock {
            try { (Get-MpComputerStatus -ErrorAction Stop).AntivirusSignatureVersion } catch { $null }
        }
        $result.OldVersion = $currentVer

        # No filename-based version parsing.
        # We always install the package and determine the version afterward.

        Invoke-Command -Session $session -ScriptBlock {
            New-Item -Path $using:TempFolderOnTarget -ItemType Directory -Force | Out-Null
        }

        $remoteFile = Invoke-Command -Session $session -ScriptBlock {
            Join-Path $using:TempFolderOnTarget $using:MpamFileName
        }

        Copy-Item -Path $SourceFile -Destination $remoteFile -ToSession $session -Force

        $install = Invoke-Command -Session $session -ScriptBlock {
            $log = Join-Path $using:TempFolderOnTarget "install_$(Get-Date -f 'yyyyMMdd_HHmmss').log"
            $p = Start-Process -FilePath $using:remoteFile -ArgumentList '/q' -Wait -PassThru -NoNewWindow `
                -RedirectStandardOutput $log -RedirectStandardError ($log + '.err')
            [pscustomobject]@{ ExitCode = $p.ExitCode; Log = $log }
        }

        $finalVer = Invoke-Command -Session $session -ScriptBlock {
            try { (Get-MpComputerStatus).AntivirusSignatureVersion } catch { $null }
        }
        $result.NewVersion = $finalVer

        if ($LogSharePath -and (Test-Path $LogSharePath -ErrorAction SilentlyContinue)) {
            try {
                $targetLogDir = Join-Path $LogSharePath $Computer
                if (-not (Test-Path $targetLogDir)) {
                    New-Item -Path $targetLogDir -ItemType Directory -Force | Out-Null
                }
                Copy-Item -Path "$remoteFile.*" -Destination $targetLogDir -FromSession $session -Force -ErrorAction SilentlyContinue
            }
            catch {}
        }

        Invoke-Command -Session $session -ScriptBlock {
            Remove-Item (Split-Path $using:remoteFile -Parent) -Recurse -Force -ErrorAction SilentlyContinue
        }
        Remove-PSSession $session

        if ($install.ExitCode -eq 0 -and $finalVer) {
            $result.Status = 'Success'
            $result.Details = "$currentVer → $finalVer"
        } else {
            $result.Status = 'Failed'
            $result.Details = "Installer exit code: $($install.ExitCode)"
        }
    }
    catch {
        $result.Status = 'Failed'
        $result.Details = $_.Exception.Message -replace "`r`n", " "
    }
    finally {
        $sw.Stop()
        $result.DurationSec = [math]::Round($sw.Elapsed.TotalSeconds, 2)
    }
    return $result
}

# ===================================================================
# Core Update Logic (used in both serial and parallel)
# ===================================================================
$latest = Get-LatestMpamFile -Root $SourceSharePath
$SourceFile = $latest.File
$MpamFileName = Split-Path $SourceFile -Leaf

Write-Log "Latest mpam-fe package detected: $MpamFileName (version $($latest.Version))" "INFO"

if (-not (Test-Path $SourceFile)) {
    Write-Log "CRITICAL: Source file not found: $SourceFile" 'ERROR'
    throw "mpam-fe.exe file missing!"
}

$UpdateScriptBlock = {
    param($Computer, $SourceFile, $TempFolderOnTarget, $MpamFileName, $WhatIfMode, $LogSharePath)

    $result = [pscustomobject]@{
        ComputerName = $Computer
        Status       = 'Unknown'
        OldVersion   = ''
        NewVersion   = ''
        DurationSec  = 0
        Details      = ''
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        if ($WhatIfMode) {
            $result.Status = 'WhatIf'
            $result.Details = 'Dry-run mode – no changes made'
            return $result
        }

        if (-not (Test-NetConnection -ComputerName $Computer -Port 5985 -InformationLevel Quiet -WarningAction SilentlyContinue)) {
            throw "WinRM (5985) not reachable"
        }

        $session = New-PSSession -ComputerName $Computer -ErrorAction Stop

        $currentVer = Invoke-Command -Session $session -ScriptBlock {
            try { (Get-MpComputerStatus -ErrorAction Stop).AntivirusSignatureVersion } catch { $null }
        }
        $result.OldVersion = $currentVer

        # No filename-based version parsing.
        # We always install the package and determine the version afterward.

        Invoke-Command -Session $session -ScriptBlock { New-Item -Path $using:TempFolderOnTarget -ItemType Directory -Force | Out-Null }
        $remoteFile = Invoke-Command -Session $session -ScriptBlock { Join-Path $using:TempFolderOnTarget $using:MpamFileName }
        Copy-Item -Path $SourceFile -Destination $remoteFile -ToSession $session -Force

        $install = Invoke-Command -Session $session -ScriptBlock {
            $log = Join-Path $using:TempFolderOnTarget "install_$(Get-Date -f 'yyyyMMdd_HHmmss').log"
            $p = Start-Process -FilePath $using:remoteFile -ArgumentList '/q' -Wait -PassThru -NoNewWindow `
                       -RedirectStandardOutput $log -RedirectStandardError ($log + '.err')
            [pscustomobject]@{ ExitCode = $p.ExitCode; Log = $log }
        }

        $finalVer = Invoke-Command -Session $session -ScriptBlock {
            try { (Get-MpComputerStatus).AntivirusSignatureVersion } catch { $null }
        }
        $result.NewVersion = $finalVer

        # Collect remote logs if LogSharePath is provided and accessible
        if ($LogSharePath -and (Test-Path $LogSharePath -ErrorAction SilentlyContinue)) {
            try {
                $targetLogDir = Join-Path $LogSharePath $Computer
                if (-not (Test-Path $targetLogDir)) { New-Item -Path $targetLogDir -ItemType Directory -Force | Out-Null }
                Copy-Item -Path "$remoteFile.*" -Destination $targetLogDir -FromSession $session -Force -ErrorAction SilentlyContinue
            }
            catch {
                # Silently skip log collection if it fails - don't fail the entire update
            }
        }

        Invoke-Command -Session $session -ScriptBlock { Remove-Item (Split-Path $using:remoteFile -Parent) -Recurse -Force -ErrorAction SilentlyContinue }
        Remove-PSSession $session

        if ($install.ExitCode -eq 0 -and $finalVer) {
            $result.Status = 'Success'
            $result.Details = "$currentVer → $finalVer"
        } else {
            $result.Status = 'Failed'
            $result.Details = "Installer exit code: $($install.ExitCode)"
        }
    }
    catch {
        $result.Status = 'Failed'
        $result.Details = $_.Exception.Message -replace "`r`n", " "
    }
    finally {
        $sw.Stop()
        $result.DurationSec = [math]::Round($sw.Elapsed.TotalSeconds, 2)
        $result
    }
}

# ===================================================================
# Execution: Start-ThreadJob Parallel Engine (PS7+) or Serial (PS5.1)
# ===================================================================

$Results = @()
$MaxConcurrent = $ParallelThreads
$TimeoutSeconds = 300  # 5 minutes
$RetryLimit = 2

# Prepare per-host log directory
$PerHostLogDir = Join-Path $LogPath "PerHost"
if (-not (Test-Path $PerHostLogDir)) {
    New-Item -Path $PerHostLogDir -ItemType Directory -Force | Out-Null
}

# Build job queue
$Queue = foreach ($comp in $TargetComputers) {
    [pscustomobject]@{
        Computer = $comp
        Attempt  = 1
        Status   = 'Pending'
    }
}

$ActiveJobs = @()
$Completed = @()
$StartTimes = @{}
$JobMeta = @{}  # Key: Job.Id, Value: @{ Computer = <string>; Attempt = <int> }

Write-Log "Executing in THREADJOB mode ($MaxConcurrent concurrent jobs)" 'HEADER'

$DashboardTimer = [System.Diagnostics.Stopwatch]::StartNew()

while ($Queue.Count -gt 0 -or $ActiveJobs.Count -gt 0) {

    # Launch new jobs if capacity available
    while ($ActiveJobs.Count -lt $MaxConcurrent -and $Queue.Count -gt 0) {
    $item = $Queue[0]
    $Queue = if ($Queue.Count -gt 1) { $Queue[1..($Queue.Count - 1)] } else { @() }

    # Extract values into simple variables for Start-ThreadJob
    $comp    = $item.Computer
    $attempt = $item.Attempt

    $job = Start-ThreadJob -ScriptBlock {
        Invoke-DefenderUpdate `
            -Computer $using:comp `
            -SourceFile $using:SourceFile `
            -TempFolderOnTarget $using:TempFolderOnTarget `
            -MpamFileName $using:MpamFileName `
            -WhatIfMode:$using:WhatIfMode `
            -LogSharePath $using:LogSharePath
    }

    # Track start time and metadata (do NOT use Add-Member on the job)
    $StartTimes[$job.Id] = Get-Date
    $JobMeta[$job.Id] = @{
        Computer = $comp
        Attempt  = $attempt
    }

    $ActiveJobs = @($ActiveJobs) + $job
}

    # Check for completed jobs
    foreach ($job in @($ActiveJobs)) {
    if ($job.State -in 'Completed','Failed','Stopped') {

        # Look up metadata for this job
        $meta = $JobMeta[$job.Id]
        $computer = $meta.Computer
        $attempt  = $meta.Attempt

        $result = Receive-Job $job -ErrorAction SilentlyContinue

        # Normalize result: handle null or arrays
        if (-not $result) {
            $result = [pscustomobject]@{
                ComputerName = $computer
                Status       = 'Failed'
                OldVersion   = ''
                NewVersion   = ''
                DurationSec  = 0
                Details      = 'Job produced no output'
                Attempt      = $attempt
                Timeout      = $false
            }
        }
        elseif ($result -is [array]) {
            $result = $result[-1]
        }

        # Ensure Attempt is present and aligned with metadata
        if ($result.PSObject.Properties.Name -contains 'Attempt') {
            $result.Attempt = $attempt
        } else {
            $result | Add-Member -NotePropertyName Attempt -NotePropertyValue $attempt -Force
        }

        # Write per-host log
        $hostLog = Join-Path $PerHostLogDir "$computer.log"
        Add-Content -Path $hostLog -Value "$(Get-Date) Attempt $attempt: $($result.Status) - $($result.Details)"

        # -------------------------------
        # FAILURE CLASSIFICATION LOGIC
        # -------------------------------

        $details = $result.Details.ToLower()

        $isHardFail =
            $details -match 'winrm' -or
            $details -match 'not reachable' -or
            $details -match 'offline' -or
            $details -match 'unreachable' -or
            $details -match 'access denied' -or
            $details -match 'authentication' -or
            $details -match 'cannot find' -or
            $details -match 'dns' -or
            $result.Timeout

        if ($result.Status -eq 'Failed' -and -not $isHardFail -and $attempt -lt 3) {
            # Retryable failure
            Write-Log "Retry scheduled for $computer (attempt $attempt → $($attempt + 1))" 'WARN'
            $Queue += [pscustomobject]@{
                Computer = $computer
                Attempt  = $attempt + 1
                Status   = 'Pending'
            }
        }
        else {
            # Hard fail OR retry limit reached OR success
            $Completed += $result
        }

        # Cleanup
        Remove-Job $job -Force
        $ActiveJobs = @($ActiveJobs | Where-Object Id -ne $job.Id)
        $JobMeta.Remove($job.Id) | Out-Null
        $StartTimes.Remove($job.Id) | Out-Null
        
        }
    }

    # Timeout handling
    foreach ($job in @($ActiveJobs)) {
    $elapsed = (Get-Date) - $StartTimes[$job.Id]

    if ($elapsed.TotalSeconds -gt $TimeoutSeconds) {
        $meta = $JobMeta[$job.Id]
        $computer = $meta.Computer
        $attempt  = $meta.Attempt

        Write-Log "TIMEOUT: $computer exceeded $TimeoutSeconds seconds (attempt $attempt)" 'ERROR'
        Stop-Job $job -Force

        $Completed += [pscustomobject]@{
            ComputerName = $computer
            Status       = 'Failed'
            OldVersion   = ''
            NewVersion   = ''
            DurationSec  = [math]::Round($elapsed.TotalSeconds,2)
            Details      = 'Timeout'
            Attempt      = $attempt
            Timeout      = $true
        }

        Remove-Job $job -Force
        $ActiveJobs = @($ActiveJobs | Where-Object Id -ne $job.Id)
        $JobMeta.Remove($job.Id) | Out-Null
        $StartTimes.Remove($job.Id) | Out-Null

        }
    }


    # Dashboard
    if ($DashboardTimer.Elapsed.TotalSeconds -ge 5) {
        $DashboardTimer.Restart()
        $running = $ActiveJobs.Count
        $pending = $Queue.Count
        $done    = $Completed.Count

        Write-Host ""
        Write-Host "=== Defender Update Dashboard ===" -ForegroundColor Cyan
        Write-Host "Running:     $running"
        Write-Host "Pending:     $pending"
        Write-Host "Completed:   $done"
        Write-Host "Elapsed:     $([string]::Format('{0:hh\:mm\:ss}', (Get-Date) - $ScriptStartTime))"
        Write-Host ""
    }

    Start-Sleep -Milliseconds 500
}

$Results = $Completed
Write-Log "ThreadJob execution completed: $($Results.Count) computers processed" 'SUCCESS'

else {
    # SERIAL MODE (PowerShell 5.1)
    Write-Log "Executing in SERIAL mode (PowerShell 5.1)" 'WARN'
    $i = 0
    foreach ($comp in $TargetComputers) {
        $i++
        $pct = [math]::Round(($i / $TargetComputers.Count) * 100, 1)
        Write-Progress -Activity "Updating Defender Definitions" -Status "Processing $i of $($TargetComputers.Count)" -CurrentOperation $comp -PercentComplete $pct
        $Results += & $UpdateScriptBlock $comp $SourceFile $TempFolderOnTarget $MpamFileName $WhatIfMode $LogSharePath
    }
    Write-Progress -Completed -Activity "Done"
}

# ===================================================================
# Generate Beautiful HTML Report (with readable multi-line CSS)
# ===================================================================
function New-HtmlReport {
    param($Data, $RunTime)

    $css = @"
<style>
    body {
        font-family: 'Segoe UI', Arial, sans-serif;
        margin: 40px;
        background: #f5f7fa;
        color: #333;
    }
    h1 {
        color: #0078d4;
        border-bottom: 3px solid #0078d4;
        padding-bottom: 10px;
    }
    h2 {
        color: #005a9e;
    }
    table {
        width: 100%;
        border-collapse: collapse;
        margin: 25px 0;
        background: white;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        border-radius: 8px;
        overflow: hidden;
    }
    th {
        background: #0078d4;
        color: white;
        padding: 14px 12px;
        text-align: left;
        font-weight: 600;
    }
    td {
        padding: 12px;
        border-bottom: 1px solid #ddd;
    }
    tr:nth-child(even) {
        background: #f9f9f9;
    }
    .success { color: #107c10; font-weight: bold; }
    .failed  { color: #d13438; font-weight: bold; }
    .skipped { color: #9c5100; font-weight: bold; }
    .footer {
        margin-top: 50px;
        color: #666;
        font-size: 0.9em;
        text-align: center;
    }
    p {
        font-size: 1.1em;
        line-height: 1.6;
    }
</style>
"@

    $body = $Data | ConvertTo-Html -Fragment -Property ComputerName,Status,OldVersion,NewVersion,DurationSec,Details
    foreach ($i in 0..($body.Count-1)) {
        $body[$i] = $body[$i] -replace '<td>Success</td>', '<td class="success">Success</td>'
        $body[$i] = $body[$i] -replace '<td>Failed</td>', '<td class="failed">Failed</td>'
        $body[$i] = $body[$i] -replace '<td>No Update Needed</td>', '<td class="skipped">No Update Needed</td>'
    }

    $summary = "$($Data.Where{$_.Status -eq 'Success'}.Count) succeeded • $($Data.Where{$_.Status -eq 'Failed'}.Count) failed • $($Data.Where{$_.Status -eq 'No Update Needed'}.Count) already current"

    return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>Microsoft Defender Definitions Update Report – $(Get-Date -f 'yyyy-MM-dd')</title>
    $css
</head>
<body>
    <h1>Microsoft Defender Antivirus Definitions Update</h1>
    <p><strong>Run Date:</strong> $ScriptStartTime<br>
       <strong>Source File:</strong> $MpamFileName<br>
       <strong>Total Duration:</strong> $($RunTime.TotalMinutes.ToString('N2')) minutes</p>
       <h2>Version History Summary</h2>
    <p>
    <strong>Oldest Version Found:</strong> $OldestVersion<br>
    <strong>Newest Version Applied:</strong> $NewestVersion<br>
    <strong>Average Delta:</strong> $AverageDelta versions<br>
    <strong>Hosts Already Current:</strong> $($Data.Where{$_.Status -eq 'No Update Needed'}.Count)<br>
    <strong>Hosts Updated:</strong> $($Data.Where{$_.Status -eq 'Success'}.Count)<br>
    <strong>Hosts Failed:</strong> $($Data.Where{$_.Status -eq 'Failed'}.Count)
    </p>

    <h2>Version History Details</h2>
    <table>
        <tr>
            <th>Computer</th>
            <th>Old Version</th>
            <th>New Version</th>
            <th>Delta</th>
        </tr>
        $(
            $Data | ForEach-Object {
                "<tr>
                    <td>$($_.ComputerName)</td>
                    <td>$($_.OldVersion)</td>
                    <td>$($_.NewVersion)</td>
                    <td>$($_.Delta)</td>
                </tr>"
            } -join "`n"
        )
    </table>

    <h2>Summary: $summary</h2>

    <table>
        $($body -join "`n")
    </table>

    <div class="footer">
        Generated by Update-DefenderOffline.ps1 v$ScriptVersion<br>
        <a href="file://$LogFile">View Full Log</a>
    </div>
</body>
</html>
"@
}

$TotalDuration = (Get-Date) - $ScriptStartTime
# ===================================================================
# Version History Calculations
# ===================================================================
foreach ($r in $Results) {
    if ($r.OldVersion -and $r.NewVersion) {
        try {
            $r | Add-Member -NotePropertyName Delta -NotePropertyValue ([version]$r.NewVersion - [version]$r.OldVersion).Build -Force
        } catch {
            $r | Add-Member -NotePropertyName Delta -NotePropertyValue 'Unknown' -Force
        }
    } else {
        $r | Add-Member -NotePropertyName Delta -NotePropertyValue 'Unknown' -Force
    }

    # Log raw version history
    Write-Log "VersionHistory: $($r.ComputerName) Old=$($r.OldVersion) New=$($r.NewVersion) Delta=$($r.Delta)" 'INFO'
}

# Fleet-wide analytics
$validDeltas = $Results | Where-Object { $_.Delta -is [int] }
$OldestVersion = ($Results | Where-Object OldVersion | Sort-Object OldVersion | Select-Object -First 1).OldVersion
$NewestVersion = ($Results | Where-Object NewVersion | Sort-Object NewVersion -Descending | Select-Object -First 1).NewVersion
$AverageDelta  = if ($validDeltas) { [math]::Round(($validDeltas.Delta | Measure-Object -Average).Average, 1) } else { 'Unknown' }

$HtmlReport = New-HtmlReport -Data $Results -RunTime $TotalDuration
$ReportFile = Join-Path $ReportPath "DefenderUpdateReport_$(Get-Date -f 'yyyyMMdd_HHmmss').html"
$HtmlReport | Out-File -FilePath $ReportFile -Encoding utf8

Write-Log "UPDATE COMPLETE in $($TotalDuration.ToString('hh\:mm\:ss'))" 'HEADER'
Write-Log "Success: $($Results.Where{$_.Status -eq 'Success'}.Count) | Failed: $($Results.Where{$_.Status -eq 'Failed'}.Count) | Skipped: $($Results.Where{$_.Status -eq 'No Update Needed'}.Count)" 'HEADER'
Write-Log "Report saved: $ReportFile" 'SUCCESS'
Write-Host ""
Write-Host "=== Version Summary ===" -ForegroundColor Cyan
Write-Host "Oldest version found : $OldestVersion"
Write-Host "Newest version applied: $NewestVersion"
Write-Host "Average delta        : $AverageDelta versions"
Write-Host ""

# ===================================================================
# Optional Email Notification
# ===================================================================
if ($SendEmail -and $To -and $SmtpServer) {
    $mailParams = @{
        From       = $From
        To         = $To
        Subject    = "Defender Update $(Get-Date -f 'yyyy-MM-dd') – $($Results.Where{$_.Status -eq 'Success'}.Count)/$($Results.Count) OK"
        Body       = $HtmlReport
        BodyAsHtml = $true
        SmtpServer = $SmtpServer
        Port       = $SmtpPort
        UseSsl     = $SmtpUseSsl
        Attachments= $ReportFile
    }
    if ($SmtpCredential) { $mailParams.Credential = $SmtpCredential }

    try {
        Send-MailMessage @mailParams
        Write-Log "Email notification sent successfully" 'SUCCESS'
    }
    catch {
        Write-Log "Failed to send email: $($_.Exception.Message)" 'ERROR'
    }
}

# Final table output
$Results | Sort-Object ComputerName | Format-Table -AutoSize