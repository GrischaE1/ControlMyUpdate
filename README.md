# Control My Update (CMU)

![Version](https://img.shields.io/badge/version-3.0-2563eb) ![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-0a6cff) ![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE) ![License](https://img.shields.io/badge/license-Free%20%2F%20AS--IS-lightgrey)

**Control My Update** takes granular control over Windows Update on managed Windows 10/11 devices — especially with **Windows Update for Business (WUfB)** — and provides an end-to-end patching **and reporting** solution. It is free, open-source, and built for the **Omnissa / VMware Workspace ONE** ecosystem (it also works standalone with Intune, ConfigMgr, an RMM, or GPO).

> 📺 Video walkthroughs: <https://www.youtube.com/playlist?list=PLxErG_Py3O2FnQpBfxT8n7erAMz1Ivcuq>
> 📖 Full reference: [Documentation.md](Documentation.md)

---

## What it does

- **Choose the update source** — Microsoft Update, WSUS, or device default.
- **Download ahead of time** — pre-stage updates so installation is fast.
- **Maintenance windows** — install only during a weekly window, per-day windows, or over-midnight windows.
- **Controlled reboots** — grace periods, block-while-user-signed-in, force-after-deadline, and a "never reboot" mode.
- **Block / uninstall specific KBs** and push **emergency KBs** that ignore the window.
- **Custom toast notifications** for end users.
- **Category & driver filtering.**
- **Rich reporting** into the registry → surfaced through Workspace ONE **Sensors** and a Workspace ONE **Intelligence** dashboard.

## How it works

Everything is coordinated through one registry hive, `HKLM:\SOFTWARE\ControlMyUpdate` — `\Settings` is configuration *in*, `\Status` is state *out*.

```
  ┌──────────────────────────┐  generates   ┌───────────────────────────┐
  │  Profile Generator (HTML) │ ───────────▶ │  PowerShell / SyncML /.reg │
  └──────────────────────────┘              └────────────┬──────────────┘
                                                          │ writes
                                                          ▼
                              HKLM:\SOFTWARE\ControlMyUpdate\Settings
                                                          │ reads
                                                          ▼
   ┌───────────────────────────────────────────────────────────────────┐
   │  ENGINE  ControlMyUpdate.ps1  (scheduled task, ~every 30 min, SYSTEM)│
   │  Windows Update COM API · maintenance window · reboot handler · DO   │
   └───────────────────────────────┬───────────────────────────────────┘
                                    │ writes
                                    ▼
             HKLM:\SOFTWARE\ControlMyUpdate\Status (+ \KBs \DO \Connection)
                                    │ read by
                                    ▼
        Workspace ONE Sensors ──▶ Intelligence custom attributes ──▶ Dashboard
```

## Repository layout

```
Application/
  ControlMyUpdate/
    ControlMyUpdate.ps1        # the engine (v3.0)
    install.ps1               # registers the scheduled tasks + copies files
    Reboot_Notification.ps1   # toast notification (user context)
    HiddenPowerShell.vbs      # launches the toast hidden
    RestartScript.cmd         # fallback restart helper
    ConnectionCheck.csv       # endpoints tested by RunConnectionTests
  Detection.ps1              # install detection (version + file hash)
  uninstall.ps1
  Installation.txt           # deployment commands (incl. detection hash)
Profile Generator/
  ControlMyUpdate-ProfileGenerator.html   # ⭐ current generator (offline, single file)
  Update_Profile_creator.ps1              # legacy WPF GUI (DEPRECATED in 3.0)
  resources/                              # legacy GUI resources
Sensors/                     # Workspace ONE sensors + upload/delete helpers
ControlMyUpdate_Dashboard_2.1.4.json     # Workspace ONE Intelligence dashboard
Documentation.md             # full reference
```

## Quick start

1. **Generate a configuration** — open [Profile Generator/ControlMyUpdate-ProfileGenerator.html](Profile%20Generator/ControlMyUpdate-ProfileGenerator.html) in any browser (works offline). Configure the **Control My Update** tab and download the output as a **PowerShell script**, a **Workspace ONE SyncML** profile, or a **.reg** file.
2. **Deploy the engine** — package `Application\ControlMyUpdate\` and run:
   ```
   powershell -executionpolicy bypass -file install.ps1 -LogPath "C:\Temp\Logs" -UpdateInterval 30
   ```
3. **Apply the configuration** from step 1 (writes `HKLM:\SOFTWARE\ControlMyUpdate\Settings`).
4. **(Optional) Reporting** — upload the `Sensors\` scripts to Workspace ONE and import the Intelligence dashboard.

Detection (version + hash) and uninstall commands are in [Application/Installation.txt](Application/Installation.txt). See [Documentation.md](Documentation.md) for step-by-step, per-platform guidance.

## The Profile Generator (v3.0)

A single self-contained HTML file — no install, no dependencies, nothing leaves the browser. Three tabs:

| Tab | Produces |
|-----|----------|
| **Control My Update** | The full v3.0 engine setting surface → PowerShell `.ps1`, Workspace ONE SyncML, or `.reg` |
| **Windows Update** | Ring-based WUfB / WSUS OMA-URI CSP (`<Replace>`/`<Delete>`) |
| **Delivery Optimization** | DO OMA-URI CSP |

It replaces the old WPF GUI (`Update_Profile_creator.ps1`), which is deprecated because it predated the v2.3 reboot model and no longer matched the engine.

## What's new in 3.0

- **Engine fixes:** `NoReboot` is now honoured; install status is recorded against the correct KB; 32/64-bit registry detection fixed; a redundant online scan and a malformed reboot check removed.
- **Engine performance:** the installed-update list (WMI/DISM) is cached per run, `Get-WmiObject` → `Get-CimInstance`, buffered logging, data-driven Delivery Optimization stats.
- **Cleaned configuration surface:** dead settings removed so the config matches the engine exactly.
- **New HTML Profile Generator** with the complete, current setting set (including `RebootGracePeriod`, which the old GUI could never set).

Full details: [Documentation.md → Version history](Documentation.md#version-history).

## Requirements

- Windows 10 (20H2+) or Windows 11
- Windows PowerShell 5.1+
- The engine runs as **SYSTEM** via scheduled task (created by `install.ps1`)

## Disclaimer

Provided **AS-IS**, without warranty of any kind. You assume all risk of use. Review every generated profile and test on a pilot ring before deploying to production fleets.

## Credits

Created by **Grischa Ernst**. Toast notification adapted from Damien Van Robaeys. Contributions and feature requests welcome via issues.
