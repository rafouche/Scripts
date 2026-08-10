# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this folder is

A flat collection of independent, standalone PowerShell/cmd scripts used by an MSP (Altec Solutions Group) for Windows endpoint and Active Directory/M365 administration — onboarding/offboarding, AD-to-Entra identity sync, VPN provisioning, agent (RMM/EDR) removal, and workstation cleanup/hardening. There is no project structure, no package manifest, no build step, and no shared library — each `.ps1`/`.cmd` file is self-contained and run directly (double-clicked, run from an elevated PowerShell prompt, or invoked by NinjaRMM).

This folder **is** a git repository (`master` branch) with `origin` set to `https://github.com/rafouche/Scripts.git` — check `git log`/`git status` before editing, and commit (and push, once the local GitHub credential is cached) after any script change so history stays on GitHub rather than only on disk. Pushing may require running `git push` from an interactive terminal the first time, since this machine has no cached GitHub credential or SSH key yet — Git Credential Manager will prompt via browser.

Several of these scripts are designed to be triggered by NinjaRMM as automation scripts (see `mcp__88622519...__run_script_on_device` / `list_automation_scripts` if working through the NinjaRMM MCP) — CLI parameters exist specifically for that non-interactive invocation path, separate from each script's GUI mode.

## Running / testing scripts

There is no test suite, linter, or CI in this folder. To validate a script:

```powershell
# Syntax-check without executing
[System.Management.Automation.Language.Parser]::ParseFile("path\to\Script.ps1", [ref]$null, [ref]$errors); $errors

# Dry run (for scripts that implement -WhatIf, see table below)
.\Onboard-ADUser.ps1 -WhatIf
```

Most scripts require **Administrator** privileges and self-elevate via a `Start-Process -Verb RunAs` relaunch at the top of the file if not already elevated — expect a UAC prompt / new process when testing interactively. The AD/M365 tools (`Onboard-ADUser.ps1`, `Offboard-ADUser.ps1`, `Set-ADEntraHardMatch.ps1`) additionally require PowerShell 7 and will auto-install it (via winget, falling back to a GitHub release MSI download) and relaunch under `pwsh.exe` if invoked from Windows PowerShell 5.1.

## Script categories

| Script | Purpose | Notes |
|---|---|---|
| `Onboard-ADUser.ps1` | Creates an AD user, adds to groups, sets OU, triggers Entra delta sync, assigns M365 license, optionally emails credentials | GUI when run with no args; CLI mode via `-FirstName`/`-LastName` for NinjaRMM. Reads `ADM365Config.json` (not present in this folder — created/supplied per-machine) for AppId/TenantId/CertThumbprint |
| `Offboard-ADUser.ps1` | Renames to "Historical-", disables + moves OU, converts mailbox to shared, grants delegate access, forwards mail, removes licenses | GUI or CLI via `-SamAccountName`. Same `ADM365Config.json` dependency |
| `Set-ADEntraHardMatch.ps1` | Sets `OnPremisesImmutableId` on an Entra user from the AD `ObjectGUID` to force a hard match | Single-user or batch-by-OU mode; cert-based Graph auth via `ADM365Config.json` |
| `JumpCloud-DeleteUser.ps1` | WinForms GUI for browsing/deleting JumpCloud users, groups, and devices across orgs | **Contains a hardcoded live JumpCloud API key** as `$DefaultApiKey` (line 6) — see Secrets below |
| `Connect-O365-MFA-v2-5.ps1` | Legacy (2017, third-party author) WinForms tool to open MFA-authenticated PowerShell sessions to Exchange Online/SharePoint Online/Compliance Center/Skype for Business | Hardcoded `$Tenant`/`$UPN` defaults (`fouche` tenant) near the top — check before reusing against a different tenant |
| `Setup-VPN.ps1`, `SetupGoldVPN.ps1`, `SetupMagersVPN.ps1`, `Setup-IKEv2-Windows.ps1` | Create/repair Windows VPN connection profiles (IKEv2 or L2TP/IPsec), patch `rasphone.pbk`, fix the Error-809 NAT-T registry key | Client-specific: `SetupGoldVPN.ps1`/`SetupMagersVPN.ps1` are hardcoded single-purpose variants of the parameterized `Setup-VPN.ps1`. **Hardcoded L2TP pre-shared keys** in script defaults — see Secrets below |
| `ForceRemoveHuntressAgents.ps1`, `ForceRemoveNinjaRMM.ps1`, `Remove-Atera.ps1`, `Remove-WolfSecurity.ps1`, `Uninstall_AE_with_Logging.ps1` | Force-uninstall a specific RMM/EDR/security agent (Huntress, NinjaRMM, Atera, HP Wolf Security, AutoElevate) via registry uninstall strings, service/process kill, and leftover file/registry cleanup | `Remove-WolfSecurity.ps1` deliberately does **not** brute-force-kill Wolf's drivers (would break the boot-security path) — it orchestrates HP's supported uninstaller instead; `-AttemptUninstall` is off by default (diagnose-only) |
| `Enable-Defender.ps1` | Clears registry policy keys that disable Windows Defender, re-enables Security Center/firewall, reports live (non-tombstoned) third-party AV | Supports `-WhatIf` via `SupportsShouldProcess` |
| `RemoveGhostDevices.ps1` | Removes non-present ("ghost") Plug and Play devices, with class/friendly-name include/exclude filtering | |
| `Remove-DeadDomainController.ps1` | Forces AD metadata cleanup of an offline DC: removes the computer account, DNS records, site/replication links, FRS/DFSR objects, seizes FSMO roles if held | Must run on the surviving DC as Domain Admin+; supports `-WhatIf` |
| `Clean-Image.ps1` | Strips OEM bloatware (HP/Dell/Lenovo/etc.) and Windows Appx packages via an allow/deny list, removes OOBE branding XML | `-Online` vs offline-image (`-OfflinePath`) modes |
| `Disable-AdapterPowerSaving.ps1` | Sets High Performance power plan, disables hibernate, disables all NIC power-saving/idle features, enables Wake-on-LAN | Writes a timestamped log file next to itself (`Disable-AdapterPowerSaving_<timestamp>.log`) |
| `Windows11-Enable-Upgrade.ps1` | Bypasses Windows 11 TPM/CPU upgrade eligibility checks via registry | Third-party script (credited to `asheroto` in header) — verify against current Microsoft bypass policy before relying on it, since Microsoft has changed these registry gates across releases |
| `Set-Computer-Name.ps1` | GUI prompt (`InputBox`) to rename the computer and reboot | |
| `AltecWindowsDefenderATPOnboardingScript.cmd` | Standard Microsoft-generated Defender for Endpoint onboarding `.cmd` (tenant-specific onboarding blob) | Not hand-written — treat as a vendor artifact, don't hand-edit the embedded onboarding payload |
| `Inst-GenApps.cmd` | One-line NiniteOne + Office deployment config wrapper | Depends on a NiniteOne install and a `64-Bit Microsoft Office.xml` config file that are not in this folder |

## Conventions across the AD/M365 tools

`Onboard-ADUser.ps1`, `Offboard-ADUser.ps1`, and `Set-ADEntraHardMatch.ps1` (the three large, actively-maintained "self-contained tool" scripts, all copyright Roger Fouche / Fouche Enterprises) share a common structure — when editing one, check whether the same fix applies to the others:

1. **Self-elevation block** at the very top (`WindowsPrincipal`/`RunAs` relaunch) — must run before anything else, including the PS7 check.
2. **PowerShell 7 bootstrap** — checks `$PSVersionTable.PSVersion.Major`, installs PS7 via winget or a GitHub-releases MSI download if missing, relaunches under `pwsh.exe`. Runs *before* `Set-StrictMode` since the bootstrap code itself needs permissive mode.
3. **Auto-discovery from the DC/AD environment** (target OU, AADConnect server, SharePoint Admin URL) rather than requiring them as required parameters — parameters exist only to *override* the auto-discovered value.
4. **`ADM365Config.json`** next to the script holds `AppId`/`TenantId`/`CertThumbprint` for cert-based Graph auth — this file is not checked into this folder (per-machine/per-tenant secret material) and must exist locally for these three scripts to authenticate. If asked to add tenant-specific config, it belongs in that JSON file, not hardcoded into the script.
5. **Dual GUI/CLI mode**: no parameters (or missing required ones) → WinForms GUI; supplying the documented CLI parameters (e.g. `-FirstName`/`-LastName`, `-SamAccountName`) → headless, for NinjaRMM.
6. **`-WhatIf`** simulates every write without committing changes — always available on these three tools; prefer it when validating changes to onboarding/offboarding logic before running against a real account.

## Secrets — do not commit real values

This folder is now a git repo (initialized 2026-08-07); a `.gitignore` excludes `ADM365Config.json`, `*.local.json`, and files matching `*secret*`/`*credentials*` so per-machine/per-tenant material never gets committed.

As of 2026-08-07, the following hardcoded live secrets were scrubbed from script defaults before the first full commit — the real values were relayed to Roger directly (chat, not committed) to move into a NinjaRMM script parameter/secure custom field or a password manager, and should be treated as needing rotation since they'd been sitting in plaintext on disk:

- `JumpCloud-DeleteUser.ps1` line 6 — `$DefaultApiKey` no longer hardcodes the live JumpCloud API key; it now defaults to `""` unless the `JC_API_KEY` environment variable is set, or the key is pasted into the GUI field at runtime.
- `Setup-VPN.ps1`, `SetupMagersVPN.ps1` — `$L2tpPsk` no longer defaults to a real pre-shared key; it must be passed via `-L2tpPsk` (e.g. a NinjaRMM script parameter). Both scripts now hard-fail with a clear error if `-VpnType L2TP` is used without `-L2tpPsk`.

**Left as-is (flagged, not scrubbed):** `Connect-O365-MFA-v2-5.ps1`'s hardcoded `$Tenant`/`$UPN` (Roger's own `fouche` tenant/admin UPN) — this is an identity default, not a credential, and the script has no alternate input path (legacy third-party GUI tool, no CLI override), so blanking it would just break the tool with no fallback. Still worth hand-editing before pointing it at a different tenant.

If asked to touch credentials in any of these files again, flag it rather than silently committing/propagating a new plaintext secret — prefer moving the value to `ADM365Config.json`-style local config (gitignored), a NinjaRMM script parameter/custom field, or a secret store, and don't paste real key material into new scripts.

## Copyright / licensing notes

Several scripts (notably `Set-ADEntraHardMatch.ps1`, `Onboard-ADUser.ps1`, `Offboard-ADUser.ps1`) carry a header asserting the script was authored by Roger Fouche independently (not work-for-hire) and licensed to Altec Solutions Group, Inc. with a termination-on-departure clause. Preserve these headers when editing — don't strip or alter the attribution/license text as a side effect of a code change.
