# Control My Update — Documentation

**Version 3.0**

Complete reference for deploying, configuring, and operating Control My Update (CMU).

## Table of contents

- [1. Overview](#1-overview)
- [2. Architecture](#2-architecture)
- [3. Requirements](#3-requirements)
- [4. Components](#4-components)
- [5. Deployment](#5-deployment)
- [6. Configuration reference](#6-configuration-reference)
- [7. Maintenance windows](#7-maintenance-windows)
- [8. Reboot behaviour](#8-reboot-behaviour)
- [9. End-user notifications](#9-end-user-notifications)
- [10. Update categories](#10-update-categories)
- [11. Reporting](#11-reporting)
- [12. The Profile Generator](#12-the-profile-generator)
- [13. Windows Update & Delivery Optimization CSP](#13-windows-update--delivery-optimization-csp)
- [14. Logging & troubleshooting](#14-logging--troubleshooting)
- [15. Version history](#15-version-history)
- [16. Uninstall](#16-uninstall)
- [17. Disclaimer](#17-disclaimer)

---

## 1. Overview

CMU replaces the opaque timing of Windows Update with an explicit, registry-driven engine. A scheduled task runs `ControlMyUpdate.ps1` as **SYSTEM** on an interval; the script reads its configuration from `HKLM:\SOFTWARE\ControlMyUpdate\Settings`, controls the Windows Update Agent through its native COM API, and writes detailed status back to `HKLM:\SOFTWARE\ControlMyUpdate\Status`. That status is surfaced to administrators through Workspace ONE Sensors and an Intelligence dashboard.

CMU is designed for Workspace ONE but is MDM-agnostic — the configuration is ultimately just registry values, which any tool (Intune, ConfigMgr, GPO, an RMM, or a manual script) can write.

## 2. Architecture

| Layer | Artifact | Responsibility |
|-------|----------|----------------|
| **Configure** | Profile Generator → PowerShell / SyncML / .reg | Writes `…\Settings` |
| **Engine** | `ControlMyUpdate.ps1` (scheduled task) | Scans, downloads, installs, reboots, reports |
| **Contract** | `HKLM:\SOFTWARE\ControlMyUpdate` | `\Settings` in, `\Status` out |
| **Report** | `Sensors\*.ps1` → Workspace ONE | Read `\Status`, publish as device attributes |
| **Visualize** | Intelligence dashboard | 20 widgets over the sensor attributes |

The engine also auto-detects the 32-bit vs. 64-bit registry hive (`SOFTWARE\ControlMyUpdate` vs. `SOFTWARE\WOW6432Node\ControlMyUpdate`).

## 3. Requirements

- Windows 10 20H2+ or Windows 11.
- Windows PowerShell 5.1 or later (`#Requires -Version 5.1`).
- Local administrator / SYSTEM rights (the scheduled task runs as SYSTEM).
- For WUfB deferrals: devices must be configured for Windows Update for Business.
- For reporting: Workspace ONE UEM + Intelligence (sensors and dashboard are WS1-specific; the `\Status` registry data is readable by any RMM).

## 4. Components

| File | Purpose |
|------|---------|
| `Application\ControlMyUpdate\ControlMyUpdate.ps1` | The engine. |
| `Application\ControlMyUpdate\install.ps1` | Creates two scheduled tasks (**Control My Update**, **Control My Update - Reboot Notification**) and copies files to `C:\Windows\ControlMyUpdate`. |
| `Application\ControlMyUpdate\Reboot_Notification.ps1` | Builds the toast (runs in user context). |
| `Application\ControlMyUpdate\HiddenPowerShell.vbs` | Launches the toast script with no visible window. |
| `Application\ControlMyUpdate\RestartScript.cmd` | `shutdown /r /f /t 120` helper. |
| `Application\ControlMyUpdate\ConnectionCheck.csv` | Endpoints tested when `RunConnectionTests = True`. |
| `Application\Detection.ps1` | Install detection: compares reported version and SHA-256 hash. |
| `Application\uninstall.ps1` | Removes tasks and the install folder. |
| `Profile Generator\ControlMyUpdate-ProfileGenerator.html` | Current generator (offline, single file). |
| `Sensors\*.ps1` | Workspace ONE sensors + `upload_sensors.ps1` / `delete_sensors.ps1`. |
| `ControlMyUpdate_Dashboard_2.1.4.json` | Workspace ONE Intelligence dashboard export. |

## 5. Deployment

### 5.1 Package & install the engine

Package the contents of `Application\ControlMyUpdate\` (plus `Detection.ps1` / `uninstall.ps1` alongside). Install command:

```
powershell -executionpolicy bypass -file install.ps1 -LogPath "C:\Temp\Logs" -UpdateInterval 30
```

- `-LogPath` — folder for the CMTrace log (`WindowsUpdate_YYYYMM.log`). Omit to disable logging.
- `-UpdateInterval` — minutes between engine runs (default 30).
- `-UpdateScheduledTask $true` — re-create the tasks (use when changing the interval).

`install.ps1` registers:
- **Control My Update** — SYSTEM, at startup and every `UpdateInterval` minutes.
- **Control My Update - Reboot Notification** — runs in the interactive user's context to show toasts.

### 5.2 Apply configuration

Generate a config with the [Profile Generator](#12-the-profile-generator) and deploy it as one of:

- **PowerShell script** — Intune *Scripts*, Workspace ONE *Scripts*, ConfigMgr, an RMM, or run manually as SYSTEM. Universal.
- **Workspace ONE Custom Settings (SyncML)** — a Custom Settings profile using the `com.airwatch.winrt.powershellcommand` CSP.
- **.reg file** — manual apply / lab.

Order does not matter — the engine reads whatever is present at run time. If no `\Settings` key exists, the engine logs an error and exits.

### 5.3 Detection

`Detection.ps1` returns exit `0` only when both the installed script version and its SHA-256 hash match. Example (v3.0):

```
powershell -executionpolicy bypass -file detection.ps1 -FileHash E2D14595D5DCEB152DD3AB5CCA44B24B34941D39789822B72F2EE75E235E26A5 -ScriptExpectedVersion "3.0"
```

> The hash is the SHA-256 of `ControlMyUpdate.ps1`. If you edit the engine, recompute it: `(Get-FileHash -Algorithm SHA256 .\ControlMyUpdate.ps1).Hash`.

### 5.4 Reporting (optional, Workspace ONE)

```
powershell -file .\Sensors\upload_sensors.ps1 -SourcePath "<folder with sensor .ps1>" `
  -APIEndpoint "asXXX.awmdm.com" -APIUser "<api user>" -APIPassword "<pw>" -APIKey "<aw-tenant-code>" `
  -OGID "<org group id>" -SmartGroupName "<smart group>"
```

Then import `ControlMyUpdate_Dashboard_2.1.4.json` into Workspace ONE Intelligence. The response type of each sensor is inferred from its filename suffix (`_bool` → BOOLEAN, `_count` → INTEGER, `_date` → DATETIME, otherwise STRING).

## 6. Configuration reference

All values are stored as **REG_SZ (String)**. Booleans are the literal strings `True` / `False`.

### Root — `HKLM:\SOFTWARE\ControlMyUpdate`

| Name | Values | Default | Description |
|------|--------|---------|-------------|
| `ScriptLogLevel` | `Info` / `Debug` / `Trace` | `Info` | Log verbosity (CMTrace format). |

### Settings — `HKLM:\SOFTWARE\ControlMyUpdate\Settings`

**Source & behaviour**

| Name | Values | Default | Description |
|------|--------|---------|-------------|
| `UpdateSource` | `Default` / `MU` / `WSUS` | `Default` | Where updates are searched. `MU` forces Microsoft Update; `WSUS` uses the configured WSUS server. |
| `ReportOnly` | `True` / `False` | `False` | Let Windows install normally; CMU only records status (no timing/reboot control). |
| `DirectDownload` | `True` / `False` | `True` | Pre-download updates as soon as they are found. |
| `RunConnectionTests` | `True` / `False` | `False` | Test the WU/DO endpoints (`ConnectionCheck.csv`) before scanning. |
| `RetryCount` | `0`–`99` | `3` | Download/install retry attempts. |
| `InstallDrivers` | `True` / `False` | `False` | Include driver updates. |

**Scanning**

| Name | Values | Default | Description |
|------|--------|---------|-------------|
| `ScanInterval` | hours | `1` | How often CMU scans (independent of the task interval). |
| `ScanRandomization` | minutes | `0` | Random 0–N delay added to each scan. |
| `LastScanTime` / `NextScanTime` / `LastInstallationDate` | (managed) | — | Written by the engine; initialise empty. |

**Update selection**

| Name | Values | Default | Description |
|------|--------|---------|-------------|
| `UpdateCategories` | `All` or comma-separated GUIDs | `All` | Restrict to specific categories ([GUID table](#10-update-categories)). |
| `HiddenUpdates` | `KB…,KB…` | `""` | KBs to hide (never install). |
| `UnHiddenUpdates` | `KB…,KB…` | `""` | KBs to un-hide. |
| `UninstallKBs` | `True` / `False` | `False` | Remove a blocked KB if already installed. |
| `EmergencyKB` | `KB…` | `""` | Install immediately, ignoring window and scan interval. |

**Maintenance window** — see [section 7](#7-maintenance-windows)

| Name | Values | Default | Description |
|------|--------|---------|-------------|
| `MaintenanceWindow` | `True` / `False` | `False` | Only install during a window. |
| `EnablePerDayMW` | `True` / `False` | `False` | Use per-day windows instead of the simple window. |
| `MWDay` | `0`–`6` comma list (Sun=0) | `""` | Days for the simple window (empty = every day). |
| `MWStartTime` / `MWStopTime` | `HH:mm` (24h) | `""` | Simple window bounds (may cross midnight). |
| `MWPerDay<Day>StartTime` / `…EndTime` | `HH:mm` | — | Per-day bounds, `<Day>` = `Monday`…`Sunday`. |

**Reboot** — see [section 8](#8-reboot-behaviour)

| Name | Values | Default | Description |
|------|--------|---------|-------------|
| `NoReboot` | `True` / `False` | `False` | Never reboot (notification only). |
| `RebootGracePeriod` | `0`–`30` (days) | `3` | Days a pending reboot is deferred while a user is present. `0` = no grace. |
| `BlockRebootWithUser` | `True` / `False` | `True` | Do not reboot while a user is signed in (until grace ends). |
| `ForceReboot` | `True` / `False` | `False` | After grace ends, reboot regardless of the window. |
| `ForceRebootwithNoUser` | `True` / `False` | `False` | Reboot as soon as no user is signed in (non-window devices). |
| `MWAutomaticReboot` | `True` / `False` | `False` | Automatically reboot inside the window when a restart is pending. |

**Notifications** — see [section 9](#9-end-user-notifications)

| Name | Values | Default | Description |
|------|--------|---------|-------------|
| `NotifyUser` | `True` / `False` | `True` | Show reboot toasts. |
| `NotificationInterval` | hours | `4` | Hours between reminders. |
| `ToastTitle` / `ToastMessage` / `ToastAdvice` | text | (samples) | Toast content. |

> **Removed in 3.0** (no longer read by the engine — do not set): `NoMWAutomaticReboot`, `MWBlockRebootWithUser`, `MWForceRebootOnlyDuringMW`, `NotifyEnduserOutsideOfMW`, `MWAutoRebootInterval`, `NoMWAutoRebootInterval`, `CMUAutoRebootInterval_OnlyMW`.

## 7. Maintenance windows

- **Disabled** (`MaintenanceWindow = False`): updates install whenever they are found.
- **Simple** (`EnablePerDayMW = False`): `MWDay` (numbers, Sunday = 0) plus `MWStartTime`/`MWStopTime`. An empty `MWDay` means every day. If the stop time is earlier than the start time, the window is treated as crossing midnight.
- **Per-day** (`EnablePerDayMW = True`): a distinct `MWPerDay<Day>StartTime`/`EndTime` for each weekday. When per-day is enabled, all seven day pairs must be present (the generator always writes them).

Day-number mapping: Sunday = 0, Monday = 1, Tuesday = 2, Wednesday = 3, Thursday = 4, Friday = 5, Saturday = 6.

## 8. Reboot behaviour

When an install leaves a pending reboot, the engine (`Test-PendingReboot` → `Start-RebootExecution`) decides whether to restart. Detection uses three signals: the CBS `RebootPending` key, the WU `RebootRequired` key, and `Microsoft.Update.SystemInfo.RebootRequired`.

Decision summary (v3.0):

1. **`NoReboot = True`** → never reboot; a notification may still be shown.
2. **Grace period** (`Test-GracePeriod`): with `RebootGracePeriod = 0` the grace is considered ended immediately. Otherwise the first pending reboot stamps `RebootDetectionDate`, and grace ends after `RebootGracePeriod` days.
3. **Window devices** (`MaintenanceWindow = True`): inside the window, with `MWAutomaticReboot = True` and grace not yet ended, the device reboots when no user is present, or when `BlockRebootWithUser = False`. Once grace ends the device reboots at the next window (or immediately if `ForceReboot = True`). `ForceRebootwithNoUser = True` reboots inside the window whenever no user is present.
4. **Non-window devices**: reboot when grace has ended and either no user is present (with `BlockRebootWithUser = True`) or `BlockRebootWithUser = False`. `ForceRebootwithNoUser = True` reboots as soon as no user is present.

A triggered reboot runs `shutdown /r /f /t 120` and shows the final "device will restart" notification. "User present" is detected via the `explorer` process.

## 9. End-user notifications

When `NotifyUser = True`, the **Reboot Notification** task shows a toast built from `ToastTitle`, `ToastMessage`, and `ToastAdvice`. Reminders repeat no more often than `NotificationInterval` hours (tracked via `\Status\RebootNotificationDate`). While within the grace period the toast offers a **Dismiss** button; once grace ends the dismiss option is removed. If `NoReboot = True`, notifications are suppressed.

## 10. Update categories

Set `UpdateCategories` to `All`, or to a comma-separated list of these category GUIDs:

| Category | GUID |
|----------|------|
| Application | `5C9376AB-8CE6-464A-B136-22113DD69801` |
| Connectors | `434DE588-ED14-48F5-8EED-A15E09A991F6` |
| Critical Updates | `E6CF1350-C01B-414D-A61F-263D14D133B4` |
| Definition Updates | `E0789628-CE08-4437-BE74-2495B842F43B` |
| Developer Kits | `E140075D-8433-45C3-AD87-E72345B36078` |
| Feature Packs | `B54E7D24-7ADD-428F-8B75-90A396FA584F` |
| Guidance | `9511D615-35B2-47BB-927F-F73D8E9260BB` |
| Service Packs | `68C5B0A3-D1A6-4553-AE49-01D3A7827828` |
| Tools | `B4832BD8-E735-4761-8DAF-37F882276DAB` |
| Update Rollups | `28BC880E-0592-4CBF-8F95-C79B17911D5F` |
| Updates | `CD5FFD1E-E932-4E3A-BF74-18BF0B1BBD83` |
| Security Updates | `0FA1201D-4330-4FA8-8AE9-B877473B6441` |
| Microsoft Defender | `8c3fcc84-7410-4a95-8b89-a166a0190486` |
| Windows 10 (1903+) | `b3c75dc1-155f-4be4-b015-3f1a91758e52` |
| Windows 11 | `72e7624a-5b00-45d2-b92f-e561c0a6a160` |
| Windows 10 LTSB | `d2085b71-5f1f-43a9-880d-ed159016d5c6` |
| Windows 10 | `a3c2375d-0c8a-42f9-bce0-28333e198407` |

## 11. Reporting

### 11.1 Status registry keys (`…\Status`)

The engine writes, among others: `Open Pending Updates`, `Total Missing Updates`, `Pending Updates`, `Pending Critical Updates`, `Pending Security Updates`, `Pending Definition Updates`, `Pending Feature Upgrades`, `Pending Update Rollups`, `PendingReboot`, `Installed KBs`, `Total Installed KBs`, per-KB detail under `…\Status\KBs\<KB>`, connectivity results under `…\Status\Connection`, and Delivery Optimization metrics under `…\Status\DO` (total & monthly MB / percentage by source).

### 11.2 Workspace ONE sensors (`Sensors\`)

41 files: 38 sensors plus `upload_sensors.ps1`, `delete_sensors.ps1`, and `hostname.ps1`. Each sensor is a thin reader of a `\Status` value:

- `wufb_*` (17) — pending counts/booleans, scan/install dates, update source, installed/missing KBs, latest CU install status.
- `do_total_*` (12) and `do_monthly_*` (8) — Delivery Optimization MB and percentage by source (peer / HTTP / cache host / group / internet), plus totals and uploads.
- `do_connection_test_bool` — Delivery Optimization connectivity.

`upload_sensors.ps1` publishes them to Workspace ONE UEM (`/API/mdm/devicesensors`, Basic auth + `aw-tenant-code`) as `WIN_RT` / `POWERSHELL` / SYSTEM / 64-bit sensors and assigns them to a Smart Group; `delete_sensors.ps1` removes them.

### 11.3 Intelligence dashboard

`ControlMyUpdate_Dashboard_2.1.4.json` provides 20 widgets bound to the sensor attributes (`airwatch.devicesensors._ca_*`): pending counts, pending reboot, update source, fully-patched, "no install over 1 month", DO download breakdown, and more.

## 12. The Profile Generator

Open `Profile Generator\ControlMyUpdate-ProfileGenerator.html` in any browser (offline, no dependencies, nothing leaves the page).

- **Control My Update** tab — the full v3.0 setting surface, with live preview and Copy/Download. Output formats: **PowerShell (.ps1)**, **Workspace ONE (SyncML)**, **Registry (.reg)**; each with a **Configure / Uninstall** switch. Fields show/hide contextually (e.g. report-only collapses the rest; per-day reveals the weekly grid; `NoReboot` hides reboot/notification options).
- **Windows Update** tab — see section 13.
- **Delivery Optimization** tab — see section 13.

The legacy WPF generator (`Update_Profile_creator.ps1`) is **deprecated** — it predates the v2.3 reboot model and no longer emits the current setting surface.

## 13. Windows Update & Delivery Optimization CSP

These tabs produce **OMA-URI SyncML** command blocks (`<Replace>` to install, `<Delete>` to remove) for a Workspace ONE / Intune **Custom Settings** profile — not CMU engine settings.

- **Windows Update** — one profile per ring; quality/feature deferrals scale by ring number. With more than 3 rings, ring 1 is placed on the Insider (Fast) channel. Covers auto-update behaviour, telemetry, channel/branch readiness, restart deadline & grace, target release version, WSUS server + dual-scan sources, and optional best-practice hidden settings.
- **Delivery Optimization** — download mode, group ID / source, min file size, bandwidth caps, Connected Cache host, and best-practice hidden settings.

## 14. Logging & troubleshooting

- **Log** — `<LogPath>\WindowsUpdate_YYYYMM.log`, CMTrace format. Raise detail with `ScriptLogLevel = Debug` or `Trace`.
- **No `\Settings` key** — the engine logs an error and exits; deploy a configuration first.
- **Nothing installing** — check `MaintenanceWindow`/window times, `ReportOnly`, and `NextScanTime` in `\Settings`; confirm the scheduled task is running as SYSTEM.
- **Unexpected reboots** — review `NoReboot`, `RebootGracePeriod`, `BlockRebootWithUser`, `ForceReboot`, `ForceRebootwithNoUser`, `MWAutomaticReboot` (see section 8).
- **Detection fails** — recompute the SHA-256 hash after any engine edit.
- **Connectivity** — enable `RunConnectionTests` and inspect `…\Status\Connection`.

## 15. Version history

**3.0** — Maintenance & efficiency release.
- *Fixes:* `NoReboot` is now honoured (was hard-coded off); post-install status is recorded against the correct KB (previously wrong for emergency/failed installs); 32/64-bit registry hive detection corrected; a redundant online update scan removed; a malformed pending-reboot check fixed.
- *Performance:* the installed-update list (WMI/DISM) is cached per run; `Get-WmiObject` → `Get-CimInstance`; buffered log writes; data-driven Delivery Optimization statistics.
- *Cleanup:* removed dead settings so the config matches the engine; added `#Requires -Version 5.1`.
- *Tooling:* new single-file HTML **Profile Generator** with the full setting surface (including `RebootGracePeriod`); legacy WPF GUI deprecated.

**2.3.1** — Fixed pending reboot registry status.
**2.3** — Redesigned reboot handler.
**2.2.x** — Update categories, per-day windows, block-reboot-with-user, driver toggle, configurable notification interval.
**2.1.x** — Retry count, connection tests, force reboot for no-user devices, update rollback.
**2.0** — Removed the PSWindowsUpdate dependency (native WU COM API).
**1.x** — Delivery Optimization stats, reporting mode, scan randomization.

## 16. Uninstall

```
powershell -executionpolicy bypass -file uninstall.ps1 -InstallDir "C:\Windows\ControlMyUpdate"
```

Removes both scheduled tasks and the install folder. To also clear configuration and status, delete the registry hive:

```
Remove-Item -Path 'HKLM:\SOFTWARE\ControlMyUpdate' -Recurse -Force
```

(The Profile Generator's **Uninstall** output produces the same removal as a script, SyncML, or .reg.)

## 17. Disclaimer

This solution is provided **AS-IS**, without warranty of any kind; the author disclaims all implied warranties and is not liable for any damages arising from its use. You assume the entire risk. Test on a pilot ring before production deployment.
