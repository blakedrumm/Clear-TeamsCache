# 🧹 Clear-TeamsCache

PowerShell script to safely clear Microsoft Teams cache for **Classic Teams** and **New Teams** on Windows.

This script is designed for support engineers, administrators, and end users who need to clear Microsoft Teams client cache so Teams can rebuild fresh local data, including messages, images, icons, thumbnails, web cache, IndexedDB data, GPU cache, local storage, service worker cache, and other cached client content.

**Repository:** https://github.com/blakedrumm/Clear-TeamsCache

---

## 📚 Table of Contents

* [🧹 Clear-TeamsCache](#-clear-teamscache)

  * [📌 Overview](#-overview)
  * [✨ Features](#-features)
  * [💻 Compatibility](#-compatibility)
  * [📁 Supported Cache Locations](#-supported-cache-locations)
  * [🧼 What This Script Clears](#-what-this-script-clears)
  * [🚫 What This Script Does Not Clear](#-what-this-script-does-not-clear)
  * [✅ Requirements](#-requirements)
  * [⬇️ Download](#️-download)
  * [🚀 Quick Start](#-quick-start)
  * [🧪 Usage Examples](#-usage-examples)
  * [⚙️ Parameters](#️-parameters)
  * [📝 Logging](#-logging)
  * [💾 Backup Behavior](#-backup-behavior)
  * [🛡️ Safety Design](#️-safety-design)
  * [🧰 Recommended Support Workflow](#-recommended-support-workflow)
  * [🔎 Troubleshooting](#-troubleshooting)
  * [📌 Notes](#-notes)
  * [👤 Author](#-author)
  * [📦 Repository](#-repository)
  * [📄 License](#-license)
  * [⚠️ Disclaimer](#️-disclaimer)

---

## 📌 Overview

`Clear-TeamsCache.ps1` clears Microsoft Teams cache for the currently signed-in user by default.

It supports both:

* **Classic Teams**
* **New Teams**

The script includes safety checks to help ensure only known Microsoft Teams cache paths are cleared.

---

## ✨ Features

* 🧹 Clears Classic Teams cache
* 🧹 Clears New Teams cache
* 💻 Supports Windows PowerShell 5.1
* 💻 Supports PowerShell 7+ on Windows
* 🔘 Uses PowerShell switch parameters
* 🧪 Supports dry-run mode
* 👥 Supports optional all-user-profile cleanup
* 💾 Supports optional backup before cleanup
* 🛑 Stops Teams-related processes before cleanup
* 🤝 Attempts graceful process close before force stop
* 🔁 Includes retry logic for locked files
* 🧰 Includes optional robocopy fallback for stubborn folders
* 🚀 Supports optional Teams restart after cleanup
* 📦 Supports optional New Teams AppX package reset
* 📝 Writes timestamped console output
* 📄 Writes an optional log file
* 🛡️ Includes safety validation to prevent accidental deletion outside known Teams cache paths

---

## 💻 Compatibility

This script is compatible with:

| Platform                 | Supported |
| ------------------------ | --------: |
| Windows PowerShell 5.1   |     ✅ Yes |
| PowerShell 7+ on Windows |     ✅ Yes |
| macOS                    |      ❌ No |
| Linux                    |      ❌ No |

---

## 📁 Supported Cache Locations

The script targets the following Microsoft Teams cache roots.

### 🟦 Classic Teams

```text
%AppData%\Microsoft\Teams
```

Example expanded path:

```text
C:\Users\<UserName>\AppData\Roaming\Microsoft\Teams
```

### 🟪 New Teams

```text
%LocalAppData%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams
```

Example expanded path:

```text
C:\Users\<UserName>\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams
```

---

## 🧼 What This Script Clears

This script clears Teams client cache data from the supported Teams cache roots.

This can include cached items such as:

* 💬 Messages
* 🖼️ Images
* 🧩 Icons
* 🖼️ Thumbnails
* 🌐 Web cache
* 🗃️ IndexedDB data
* 🎮 GPU cache
* 📦 Local storage
* ⚙️ Service worker cache
* 🧹 Other Teams client cached content

---

## 🚫 What This Script Does Not Clear

This script is intentionally scoped to Teams client cache paths.

It does **not** clear:

* 🔐 Windows credentials
* 🔐 Browser credentials
* 🪪 Office identity cache
* 🔑 WAM tokens
* 🔑 OneAuth tokens
* ☁️ Microsoft 365 identity data
* 📄 User documents
* 💬 Chat history stored in Microsoft 365
* ☁️ Teams data stored in the cloud

---

## ✅ Requirements

* Windows operating system
* Windows PowerShell 5.1 or PowerShell 7+
* Local user permissions for current-user cache cleanup
* Administrator permissions when using `-ClearAllUserProfiles`

---

## ⬇️ Download

Clone the repository:

```powershell
git clone https://github.com/blakedrumm/Clear-TeamsCache.git
```

Change into the repository folder:

```powershell
cd .\Clear-TeamsCache
```

---

## 🚀 Quick Start

Run the script with default options:

```powershell
.\Clear-TeamsCache.ps1
```

By default, the script will:

* ✅ Clear Teams cache for the current user
* ✅ Include Classic Teams cache
* ✅ Include New Teams cache
* ✅ Stop Teams-related processes before cleanup
* ✅ Restart Teams after cleanup
* ✅ Write a log file to the current user's temp folder

---

## 🧪 Usage Examples

### 🧹 Clear Teams cache for the current user

```powershell
.\Clear-TeamsCache.ps1
```

---

### 🔎 Preview what would be removed without deleting anything

```powershell
.\Clear-TeamsCache.ps1 -DryRunOnly
```

---

### 👥 Clear Teams cache for all local user profiles

Run PowerShell as Administrator first.

```powershell
.\Clear-TeamsCache.ps1 -ClearAllUserProfiles
```

---

### 💾 Clear Teams cache and create a backup first

```powershell
.\Clear-TeamsCache.ps1 -CreateBackupBeforeCleanup
```

---

### 🚫 Clear Teams cache without restarting Teams afterward

```powershell
.\Clear-TeamsCache.ps1 -SkipRestartTeamsAfterCleanup
```

---

### 🟦 Clear only Classic Teams cache

```powershell
.\Clear-TeamsCache.ps1 -SkipNewTeams
```

---

### 🟪 Clear only New Teams cache

```powershell
.\Clear-TeamsCache.ps1 -SkipClassicTeams
```

---

### 🛑 Skip stopping Teams before cleanup

```powershell
.\Clear-TeamsCache.ps1 -SkipStopTeamsBeforeCleanup
```

> ⚠️ This is not recommended because locked cache files may remain.

---

### 📝 Disable log file creation

```powershell
.\Clear-TeamsCache.ps1 -DisableLogFile
```

---

### 📦 Reset the New Teams AppX package

```powershell
.\Clear-TeamsCache.ps1 -ResetNewTeamsAppPackage
```

> ⚠️ This option is disabled by default because resetting the New Teams AppX package can remove app data and personalization settings for the current user.

---

## ⚙️ Parameters

| Parameter                                    |    Type | Description                                                                                            |
| -------------------------------------------- | ------: | ------------------------------------------------------------------------------------------------------ |
| `-ClearAllUserProfiles`                      |  Switch | Clears Teams cache for all local user profiles. Requires Administrator.                                |
| `-SkipCurrentUserProfile`                    |  Switch | Skips cleanup for the currently signed-in user. Useful with `-ClearAllUserProfiles`.                   |
| `-SkipClassicTeams`                          |  Switch | Skips Classic Teams cache cleanup.                                                                     |
| `-SkipNewTeams`                              |  Switch | Skips New Teams cache cleanup.                                                                         |
| `-SkipStopTeamsBeforeCleanup`                |  Switch | Skips stopping Teams-related processes before cleanup.                                                 |
| `-SkipForceStopTeamsProcesses`               |  Switch | Skips force-stopping Teams processes if graceful close does not work.                                  |
| `-StopTeamsGracefulTimeoutSeconds`           | Integer | Number of seconds to wait after graceful close before force-stopping Teams processes. Default is `15`. |
| `-SkipTeamsWebView2Processes`                |  Switch | Skips stopping Teams-related WebView2 processes.                                                       |
| `-SkipRestartTeamsAfterCleanup`              |  Switch | Skips restarting Teams after cleanup completes.                                                        |
| `-DryRunOnly`                                |  Switch | Shows what would be removed without deleting anything.                                                 |
| `-CreateBackupBeforeCleanup`                 |  Switch | Creates a backup of Teams cache contents before cleanup.                                               |
| `-BackupRoot`                                |  String | Folder where backups are stored when backup is enabled.                                                |
| `-DisableLogFile`                            |  Switch | Disables writing output to a log file.                                                                 |
| `-LogFilePath`                               |  String | Full path to the log file.                                                                             |
| `-RetryCount`                                | Integer | Number of retry attempts for file removal and process stop actions. Default is `3`.                    |
| `-RetryDelaySeconds`                         | Integer | Number of seconds to wait between retry attempts. Default is `2`.                                      |
| `-DisableRobocopyFallbackForStubbornFolders` |  Switch | Disables robocopy mirror fallback for stubborn cache folders.                                          |
| `-ResetNewTeamsAppPackage`                   |  Switch | Attempts to reset the New Teams AppX package for the current user. Disabled by default.                |

---

## 📝 Logging

By default, the script writes output to the console and to a log file in the current user's temp folder.

Example log path:

```text
C:\Users\<UserName>\AppData\Local\Temp\Clear-TeamsCache_<ComputerName>_<Timestamp>.log
```

Disable log file creation:

```powershell
.\Clear-TeamsCache.ps1 -DisableLogFile
```

Specify a custom log file path:

```powershell
.\Clear-TeamsCache.ps1 -LogFilePath "C:\Temp\Clear-TeamsCache.log"
```

---

## 💾 Backup Behavior

Backups are disabled by default.

Create a backup before clearing cache:

```powershell
.\Clear-TeamsCache.ps1 -CreateBackupBeforeCleanup
```

By default, backups are created under the current user's temp folder.

Example backup path:

```text
C:\Users\<UserName>\AppData\Local\Temp\TeamsCacheBackup_<Timestamp>
```

Specify a custom backup root:

```powershell
.\Clear-TeamsCache.ps1 -CreateBackupBeforeCleanup -BackupRoot "C:\Temp\TeamsCacheBackup"
```

---

## 🛡️ Safety Design

The script includes safety checks before deleting cache data.

It only allows cleanup under known Teams cache root paths:

```text
\AppData\Roaming\Microsoft\Teams
```

```text
\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams
```

The script also blocks wildcard deletion paths and refuses to clear unexpected locations.

---

## 🧰 Recommended Support Workflow

### 1️⃣ Run a dry run first

```powershell
.\Clear-TeamsCache.ps1 -DryRunOnly
```

### 2️⃣ Run the normal cleanup

```powershell
.\Clear-TeamsCache.ps1
```

### 3️⃣ For shared or multi-user devices

Run PowerShell as Administrator.

```powershell
.\Clear-TeamsCache.ps1 -ClearAllUserProfiles
```

---

## 🔎 Troubleshooting

### ❓ Script says Teams files could not be removed

Some files may still be locked by Teams, WebView2, antivirus, indexing, or another process.

Try running the script again:

```powershell
.\Clear-TeamsCache.ps1
```

If files still remain, close Teams manually and run again.

---

### ❓ All-user cleanup fails

Make sure PowerShell is running as Administrator:

```powershell
.\Clear-TeamsCache.ps1 -ClearAllUserProfiles
```

---

### ❓ Teams does not restart automatically

Open Microsoft Teams manually from the Start menu.

The script attempts multiple restart methods, but restart behavior can vary depending on Teams installation state and user context.

---

### ❓ New Teams AppX reset does not work in PowerShell 7+

The script attempts to use Windows PowerShell 5.1 internally as a fallback for AppX operations when needed.

If reset still fails, run the script from Windows PowerShell 5.1 or reset Microsoft Teams from Windows Settings.

---

## 📌 Notes

After Teams cache is cleared, Microsoft Teams may take longer to open the first time while it rebuilds cache.

Users may need to sign in again depending on the Teams client state, tenant policies, and local identity/session state.

---

## 👤 Author

**Blake Drumm**
[blakedrumm@microsoft.com](mailto:blakedrumm@microsoft.com)

---

## 📦 Repository

https://github.com/blakedrumm/Clear-TeamsCache

---

## 📄 License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

## ⚠️ Disclaimer

This script is provided as-is with no warranties or guarantees. Review and test before using in production environments.
