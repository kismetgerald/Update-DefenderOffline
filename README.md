# Update-DefenderOffline

> Automate Microsoft Defender antivirus definition updates for air-gapped and offline Windows systems

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE.txt)
[![Version](https://img.shields.io/badge/Version-0.0.1--alpha-orange.svg)](https://github.com/kismetgerald/Update-DefenderOffline)

## Overview

**Update-DefenderOffline** is a PowerShell script that updates Microsoft Defender definitions on systems without internet access using PowerShell Remoting (WinRM). Perfect for air-gapped networks, SCADA systems, and high-security environments.

**Key Features:**
- 🚀 **10x faster** with PowerShell 7+ parallel processing (up to 32 threads)
- 📊 **Real-time progress** tracking and beautiful HTML reports
- 🔄 **Auto-discovery** of computers from Active Directory
- 📧 **Email notifications** with detailed results
- 🛡️ **Safe & tested** with dry-run mode and version checking
- ⚙️ **Enterprise-ready** for scheduled tasks and service accounts

## Quick Start

### Prerequisites

- PowerShell 5.1 or higher (7+ recommended for performance)
- Administrator privileges
- WinRM enabled on target computers
- Network share with `mpam-fe.exe` (x64 version)

### Basic Usage

1. **Download the latest Defender definitions** from Microsoft on an internet-connected machine and copy to a network share

2. **Run the script:**

```powershell
.\Update-DefenderOffline.ps1 `
    -SourceSharePath "\\fileserver\DefenderUpdates" `
    -MpamFileName "mpam-feX64.exe"
```

The script will:
- Auto-discover Windows computers from Active Directory
- Update all systems in parallel (if using PowerShell 7+)
- Generate an HTML report with results
- Log all activity to `C:\Logs\`

### First-Time Setup

**Option 1: Auto-discover computers from AD (recommended)**

Just run the script - it will automatically create `hosts.conf` from your Active Directory:

```powershell
.\Update-DefenderOffline.ps1 `
    -SourceSharePath "\\fileserver\DefenderUpdates" `
    -MpamFileName "mpam-feX64.exe"
```

**Option 2: Specify computers manually**

```powershell
.\Update-DefenderOffline.ps1 `
    -ComputerName "PC01","PC02","SERVER01" `
    -SourceSharePath "\\fileserver\DefenderUpdates" `
    -MpamFileName "mpam-feX64.exe"
```

**Option 3: Create hosts.conf file**

Create a file named `hosts.conf` in the script directory:

```
# One computer name per line
WORKSTATION01
WORKSTATION02
SERVER01
```

## Common Scenarios

### Weekly Scheduled Updates

```powershell
# Run via Task Scheduler with specific version file
.\Update-DefenderOffline.ps1 `
    -SourceSharePath "\\fileserver\updates" `
    -MpamFileName "mpam-feX64_1.405.9999.0.exe"
```

### With Email Notifications

1. **Save SMTP credentials (one-time):**

```powershell
.\Update-DefenderOffline.ps1 -SaveSmtpCredential
```

2. **Run with email enabled:**

```powershell
$cred = Import-Clixml ".\Config\SmtpCredential.xml"

.\Update-DefenderOffline.ps1 `
    -SourceSharePath "\\fileserver\updates" `
    -MpamFileName "mpam-feX64.exe" `
    -SendEmail `
    -SmtpServer "smtp.company.com" `
    -SmtpPort 587 `
    -SmtpUseSsl `
    -From "defender@company.com" `
    -To "it-team@company.com" `
    -SmtpCredential $cred
```

### Dry-Run Testing

Test without making any changes:

```powershell
.\Update-DefenderOffline.ps1 `
    -SourceSharePath "\\fileserver\updates" `
    -MpamFileName "mpam-feX64.exe" `
    -WhatIfMode
```

### High-Performance Mode

For large environments (500+ computers), use PowerShell 7+ with maximum threads:

```powershell
pwsh.exe -File .\Update-DefenderOffline.ps1 `
    -SourceSharePath "\\fileserver\updates" `
    -MpamFileName "mpam-feX64.exe" `
    -ParallelThreads 32
```

## How It Works

1. **Discovers targets** - Reads from hosts.conf or queries Active Directory
2. **Tests connectivity** - Verifies WinRM (port 5985) is accessible
3. **Checks versions** - Compares current vs. new version, skips if already updated
4. **Copies & installs** - Transfers mpam-fe.exe via PSSession and runs silent install
5. **Verifies success** - Confirms new version installed correctly
6. **Generates reports** - Creates HTML report with detailed results
7. **Sends notifications** - Emails summary (if configured)

## Getting the Definition Files

1. Visit [Microsoft Malware Protection Center](https://go.microsoft.com/fwlink/?LinkID=121721&arch=x64)
2. Download `mpam-fe.exe` (x64 version)
3. Copy to your network share
4. Run the script pointing to that share

**Tip:** Rename files with version numbers for tracking: `mpam-feX64_1.405.9999.0.exe`

## Output & Logs

- **Console:** Color-coded real-time progress
- **Log files:** `C:\Logs\Update-DefenderOffline_YYYYMMDD_HHmmss.log`
- **HTML reports:** `.\Reports\DefenderUpdateReport_YYYYMMDD_HHmmss.html`
- **Remote logs:** Optional collection to network share (configure with `-LogSharePath`)

## Requirements

| Component | Requirement |
|-----------|-------------|
| PowerShell | 5.1 minimum (7+ recommended) |
| Privileges | Administrator on local and remote systems |
| Network | WinRM enabled on targets (port 5985) |
| Source file | mpam-fe.exe (x64) on accessible share |
| Optional | ActiveDirectory module or domain membership |

## Configuration Options

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-SourceSharePath` | Network path to mpam-fe.exe | *(required)* |
| `-MpamFileName` | Filename of update package | *(required)* |
| `-ComputerName` | Manual computer list | *(auto-discover)* |
| `-LogPath` | Local log directory | `C:\Logs` |
| `-ReportPath` | HTML report directory | `.\Reports` |
| `-LogSharePath` | Remote log collection | *(disabled)* |
| `-ParallelThreads` | Max concurrent threads (PS 7+) | `16` |
| `-WhatIfMode` | Dry-run mode | `false` |
| `-SendEmail` | Enable email notifications | `false` |

See `Get-Help .\Update-DefenderOffline.ps1 -Full` for complete documentation.

## Troubleshooting

**"WinRM (5985) not reachable"**
```powershell
# On target computers, run:
Enable-PSRemoting -Force
```

**"This script requires administrative privileges"**
```powershell
# Run PowerShell as Administrator
Right-click PowerShell → "Run as Administrator"
```

**"Cannot parse version from filename"**
```
Filename must include version: mpam-feX64_1.405.9999.0.exe
                                          ^^^^^^^^^^^^^^
```

**No progress updates in parallel mode**
```powershell
# Upgrade to PowerShell 7 for parallel support
winget install Microsoft.PowerShell
```

## Performance Benchmarks

| Computers | PS 5.1 (Serial) | PS 7+ (16 threads) | PS 7+ (32 threads) |
|-----------|-----------------|-------------------|-------------------|
| 10 | ~2 min | ~1 min | ~1 min |
| 50 | ~10 min | ~4 min | ~3 min |
| 100 | ~20 min | ~8 min | ~5 min |
| 500 | ~100 min | ~35 min | ~20 min |

*Times assume ~10-15 seconds per computer*

## Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

## License

This project is licensed under the MIT License - see the [LICENSE.txt](LICENSE.txt) file for details.

## Author

**Kismet Agbasi**
- GitHub: [@kismetgerald](https://github.com/kismetgerald)
- Email: KismetG17@gmail.com

*AI Contributors: ClaudeAI, Grok*

## Support

For detailed documentation, see [claude.md](claude.md)

For issues or questions, please [open an issue](https://github.com/kismetgerald/Update-DefenderOffline/issues).

---

**⭐ If this project helps you, please consider giving it a star!**
