# AVD UAT automation

## Script comparison

| Script | What it does | Reporting | Unattended / CLM |
| --- | --- | --- | --- |
| `UAT_helper.ps1` | Opens applications and websites, pauses repeatedly for operator review, and runs several machine checks. | Console only; does not update the checklist. | No. It requires a desktop user and `Read-Host`. |
| `automate-avd-checks.ps1` | Reviews results already present in Excel and labels completed/pending rows. It does not execute the UAT checks. | Copies and updates XLSX through Excel COM; reports passing rows in HTML. | No. Excel COM is blocked by Constrained Language Mode and requires Office. |
| `AVD-UAT-Checklist-Automated.ps1` | Unattended entry point that invokes the check engine without exposing interactive checks. | Produces prepopulated HTML, CSV evidence, and a completed XLSX copy. | Yes. Intended for SYSTEM startup execution. |
| `AVD-UAT-Checklist-Automated.exe` | Executable form of the unattended entry point. | Produces the same outputs as the PowerShell entry point. | Yes. Uses the same startup task trigger. |
| `AVD-UAT-Checklist-Manual.ps1` | Loads automated evidence and prompts only for `Partial Pass` and `Manual` rows. | Produces separate final HTML, CSV, and XLSX files while preserving automated notes. | No. Run in the UAT user's interactive session. |
| `AVD-UAT-Checklist-Manual.exe` | Executable form of the interactive manual review entry point. | Produces the same final report set as the manual PowerShell script. | No. Run in the UAT user's interactive session. |
| `AVD-UAT-Checklist.ps1` | Shared check and report engine used by the automated entry point. | Produces the automated report set. | Do not invoke with interactive options from a scheduled task. |

## Design

`AVD-UAT-Checklist-Automated.ps1` calls the shared engine using built-in PowerShell commands for collection and HTML/CSV output. Excel COM and third-party modules are not used for workbook I/O. The small `tools\AvdChecklistTool\AvdChecklistTool.exe` utility reads and updates the XLSX package directly, preserving the source workbook and writing results into a timestamped copy.

Unattended runs do not launch GUI applications, URLs, settings pages, or interactive prompts. Checks that require a visible application or user action are recorded as `Manual`. Detailed timestamped progress, check results, exceptions, and output paths are written only to the automated run folder's `AVD UAT Automated Log - <timestamp>.log`; they are not displayed on screen.

Application startup probes wait no more than 10 seconds. The engine records existing process IDs before launch and stops only processes created by that probe. Guided browser, authentication, Store, installer, and visual checks are not reachable from the unattended entry point.

Checklist IDs are scoped by worksheet because the workbook reuses IDs across the Azure Virtual Desktop and Remote Apps sheets. Automated prerequisite evidence receives `Partial Pass` when the documented test also requires authentication, an upload/download, user observation, or comparison with an approved inventory.

The current template adds the manual Adobe Acrobat sensitivity-label workflow at
Azure Virtual Desktop **AVD-05**. AWS WickrGov is **AVD-06**, and subsequent checks
use their shifted worksheet IDs. Remote Apps numbering is unchanged. The final
rows are **AVD-80 (Certificate Verification)**, reviewed manually, and
**AVD-81 (Protected Roots Verification)**, checked automatically. This template
has no Fiery updater or IPv6 row. Use the matching updated template and engine
together; rerun automated collection rather than reusing evidence with older IDs.

The following checks are manual-only in every execution context. The automated
run records `Manual` without running their automated probes or roll-ups, and the
manual launcher prompts for their final results:

| Worksheet | ID | Manual verification |
| --- | --- | --- |
| Azure Virtual Desktop Tasks | AVD-05 | In Adobe Acrobat, apply a Microsoft Sensitivity Label to a PDF, then save, close, and reopen it without errors. |
| Azure Virtual Desktop Tasks | AVD-21 | About Teams displays AVD SlimCore Media Optimized or AVD Media Optimized. |
| Azure Virtual Desktop Tasks | AVD-25 | Complete the documented Teams activities and confirm Teams functions correctly. |
| Azure Virtual Desktop Tasks | AVD-41 | Outlook Cached Exchange Mode is checked and the download period is 3 days. |
| Azure Virtual Desktop Tasks | AVD-54 | As the signed-in UAT user, launch Snipping Tool, perform both documented snip tests, and confirm each snip opens in Snipping Tool without a default-application prompt. |
| Azure Virtual Desktop Tasks | AVD-55 | As the signed-in UAT user, capture and open a snip and confirm the image is visible rather than black or blank. |
| Azure Virtual Desktop Tasks | AVD-58 | Complete the IDCO CAC authentication workflow without errors or blocking actions. |
| Azure Virtual Desktop Tasks | AVD-59 | Visually confirm the Army wallpaper is applied. |
| Azure Virtual Desktop Tasks | AVD-61 | As the standard UAT user without elevation, try to create a folder in C:\ and confirm access is denied. |
| Azure Virtual Desktop Tasks | AVD-62 | As the standard UAT user without elevation, try to create a folder in D:\ and confirm access is denied. |
| Azure Virtual Desktop Tasks | AVD-69 | Confirm OneDrive is accessible and syncing. |
| Azure Virtual Desktop Tasks | AVD-80 | Verify the required DoD and ECMA certificates exist and are not expired. |
| Remote Apps Tasks | AVD-15 | Complete the IDCO CAC authentication workflow without errors or blocking actions. |

AVD-54 and AVD-55 never search current-user application aliases or launch Snipping
Tool during automated collection, including SYSTEM startup with no user logged in.
Both are recorded as `Manual` and presented by the manual launcher for review in
the signed-in UAT user's session.

AVD-61 and AVD-62 never attempt automated writes. Review them in the signed-in
standard UAT user's session, not as SYSTEM or an elevated administrator. A missing
drive or another error is not evidence of access denied; only the expected
permission denial satisfies these checks.

Each result follows a strict evidence rule:

- `Pass`: automation fully proves the documented expected result.
- `Partial Pass`: automation proves part of the result, but a named manual task remains.
- `Fail`: automation ran and found negative evidence, such as missing software, missing configuration, or an unsuccessful operation.
- `Manual`: no reliable automated action can prove any part of the expected result in the current execution context.

Azure Virtual Desktop Tasks **AVD-60 (Drive Encryption)** is included in automated
runs and checks BitLocker `ProtectionStatus` on both `C:` and `D:`. It passes
when protection is `Off` on both drives, even if either drive is still encrypted.
Protection `On` on either drive fails the check. Missing drives, unavailable
BitLocker cmdlets, unreadable status, or unknown status require manual review
unless another drive already proves a failure. The check does not change
BitLocker protection or encrypt/decrypt any drive.

HTML rows and workbook Result cells use green for `Pass`, yellow for `Partial Pass` and `Manual`, and red for `Fail`.

Azure Virtual Desktop Tasks **AVD-81 (Protected Roots Verification)** reads
`HKLM:\Software\Policies\Microsoft\SystemCertificates\Root\ProtectedRoots`
in both SYSTEM and user runs. It requires `Flags=65537` as `REG_DWORD` and
`PeerUsages` as `REG_MULTI_SZ` with exactly these entries (order does not matter):
`1.3.6.1.5.5.7.3.2`, `1.3.6.1.5.5.7.3.4`, and `1.3.6.1.4.1.311.10.3.4`.
Missing or unreadable policy values, incorrect types, or mismatched values fail
with diagnostic evidence. The check is read-only; it does not attempt certificate
imports. It is separate from the manual DoD/ECMA certificate inventory at AVD-80.

## Mandatory prerequisite gate

Both the automated and manual launchers invoke `AVD-UAT-Prerequisite-Check.ps1` before any checklist workbook is read, any UAT check runs, or any manual prompt is displayed. The gate runs fresh in the launcher's current execution context; a prior result is never reused.

The checker is compatible with PowerShell Constrained Language Mode. It uses native `dsregcmd.exe` and `schtasks.exe`, registry and service cmdlets, hashtables, operators, and file cmdlets. It does not require custom object creation, non-core static method invocation, `PSObject` reflection, `Add-Type`, `New-Object`, or the ScheduledTasks PowerShell module.

All six requirements must pass:

1. `dsregcmd.exe /status` reports `AzureAdJoined : YES`.
2. An `MS DM Server` enrollment exists under `HKLM:\SOFTWARE\Microsoft\Enrollments`, and its matching `\Microsoft\Windows\EnterpriseMgmt\<enrollment ID>\` namespace contains scheduled tasks.
3. Service `IntuneManagementExtension` exists and is `Running`.
4. FSLogix service `frxsvc` exists and is `Running`.
5. Microsoft Dependency Agent service `MicrosoftDependencyAgent` exists and is `Running`. This is a prerequisite-only check, not a UAT checklist row.
6. **Host Pool Status:** `SessionHostPool` under `HKLM:\SOFTWARE\Microsoft\RDMonitoringAgent` matches the entire expected pool name derived from the device name.

Host Pool Status runs automatically in both launchers; no separate test script,
Azure sign-in, or Azure query is needed. The device naming format is
`AVD<region><ring><two-digit group><slice><host-pool digit><three-digit host sequence>`.
Regions are AZ, VA, TX, or DE; rings are INT, PLT, BRD, or TST; slices are A or B.
The first digit after the slice selects the host-pool number (1-9); the remaining
three digits identify the host and are not part of the pool name.
AZ and TX devices require AZ host pools; VA and DE devices require VA host pools.
For example:

| Device | Required SessionHostPool |
| --- | --- |
| `AVDAZINT01A1001` | `P-ARMY-AVD-AZ-INT01A-1` |
| `AVDTXPLT02B2001` | `P-ARMY-AVD-AZ-PLT02B-2` |
| `AVDDEBRD03A1001` | `P-ARMY-AVD-VA-BRD03A-1` |

Comparison is case-insensitive but requires the complete name: added prefixes,
suffixes, or whitespace do not pass. Missing, blank, unreadable, non-string, or
unrecognized values and unrecognized device names fail and block UAT. The
prerequisite HTML report and log include the device name, actual pool, expected
pool, and mismatch reason (region, ring, group, slice, and/or host-pool number).
This adds a prerequisite only; it does not add a checklist workbook row or change
retry behavior.

The `tests\Test-HostPoolPrerequisite.ps1` development test validates mappings with
simulated registry data. It is not needed on session hosts and is not included in
the live-test ZIP. To run it from the source folder:

```powershell
powershell.exe -NoProfile -File .\tests\Test-HostPoolPrerequisite.ps1
```

For isolated gate/report and launcher-blocking tests in PowerShell 5.1 and 7,
including Constrained Language Mode, run
`powershell.exe -NoProfile -File .\tests\Test-PrerequisiteGate.ps1`.
These development tests simulate host dependencies without changing the live
registry or machine configuration.

Every gate execution creates an HTML report and troubleshooting log. Automated runs place them in the new `Automated_Results_<timestamp>` folder. Manual runs place them under `%TEMP%\ArmyAVD\UAT-Prerequisites`.

- Exit code `0`: all prerequisites passed; UAT may continue.
- Exit code `2`: one or more prerequisites failed; UAT did not start.
- Exit code `3`: the checker was missing or prerequisite evaluation could not complete; UAT did not start.

The checker can also be run by itself:

```powershell
powershell.exe -NoProfile -NonInteractive `
  -File .\AVD-UAT-Prerequisite-Check.ps1
```

The finalized scripts and rebuilt launchers must have valid organizational
code signatures before deployment under AllSigned, WDAC, or AppLocker.
Do not relax application-control or execution policy to deploy the update.
Verify the approved release before live testing:

```powershell
Get-AuthenticodeSignature .\*.ps1, .\*.exe, .\tools\AvdChecklistTool\AvdChecklistTool.exe |
    Format-Table Path, Status
```

Every required runtime signature must be `Valid` under the deployment policy.
Recalculate `SHA256SUMS.txt` and refresh the ZIP after signing. The included
Word runbook also documents the updated path and signing requirements.

## Check configuration

Some checks require an environment-approved baseline or test input. If an applicable value is omitted, that check reports `Manual` rather than guessing an expected result.

Deploy the complete folder contents to `C:\Program Files\ArmyAVD\UAT`.
When `-InputFile` is omitted, automation selects the newest `.xlsx` or `.xlsm`
workbook in that installation directory whose name begins with `AVD Test Plan`.
It does not search OneDrive Documents or the current working directory.
`-UatDirectory` selects an alternate installation root for testing.
`-InputFile` remains an explicit automated template override and accepts wildcards.

| Parameter | Purpose |
| --- | --- |
| `-PurviewExtensionId` | Required Chrome Purview extension ID. |
| `-EdgePurviewExtensionId` | Retained for configuration compatibility; AVD-29 now detects Purview, MIP, or DLP Windows services. |
| `-ChromeBookmarksBaselineFile` | Retained for configuration compatibility; AVD-15 verifies that managed Chrome policy exposes Army Bookmarks. |
| `-EdgeFavoritesBaselineFile` | Retained for configuration compatibility; Remote Apps AVD-10 and Azure Virtual Desktop AVD-33 verify that managed Edge policy exposes Army Bookmarks. |
| `-CopilotArtifactName`, `-CopilotAppId`, `-CopilotUrl` | Approved Copilot desktop, Start menu, and favorites artifacts. |
| `-ExeInstallerTestFile`, `-MsiInstallerTestFile` | Approved benign installers for application-control block tests. |
| `-RequiredLanguages` | Required language objects identified by `LanguageTag`. AVD-67 requires LpCab for each language; `LanguagePacks` and `LanguageFeatures` fields are retained for compatibility but do not control validation. |
| `-RequiredLanguageTags` | Legacy required BCP-47 language tags used when `RequiredLanguages` is not configured; each also requires LpCab. |
| `-WindowsUpdateHistoryUrl` | Official Microsoft update-history page used to compare the local build with the latest build from the previous month. |
| `-ApprovedOsBuild` | Optional fallback minimum build when no update-history URL is configured. |
| `-ExpectedRemoteAppHostPattern` | Regex matching approved Remote Apps ring and slice host names. |
| `-OwaUrl` | Approved absolute HTTPS OWA endpoint for the guided encryption test. |
| `-RunInteractiveChecks` | Enables guided user-interface and authentication checks; never supplies credentials or secrets. |

Use `avd-uat.config.example.json` as the schema for deployment configuration. `-ConfigurationFile` accepts that JSON object; explicit command-line check parameters override matching file values. Relative baseline and test-file paths resolve relative to the configuration file. Do not store credentials, PINs, tokens, or other secrets in it.

## Finalized live-test behavior

- Azure Virtual Desktop AVD-73 requires `SetDisableUXWUAccess=1`, `DeferFeatureUpdates=1`, and `DeferQualityUpdates=1` under `HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate`. It no longer reads or requires `SetDisablePauseUXAccess`. All three matching policies produce `Partial Pass`, retaining manual confirmation that Check for updates and Pause updates are greyed out and the organization-policy message is visible. Missing, unreadable, or mismatched required policies produce `Fail`.
- Azure Virtual Desktop AVD-67 requires installed **LpCab** packages for every configured language (de-DE, en-US, fr-FR, ja-JP, ko-KR by default). LXP packages and language features are not required. An LXP-only language or a missing required language fails the check.
- AVD-00 defaults to a 40% average CPU limit and an 80% maximum memory-use limit.
- Azure Virtual Desktop AVD-15 recursively queries `HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\managedbookmarks` and passes when any key name, value name, or value data contains `Army Bookmarks`.
- Azure Virtual Desktop AVD-29 passes when at least one Windows service display name matches Microsoft Purview, MIP, or DLP.
- Remote Apps AVD-11 and Azure Virtual Desktop AVD-34 read only the HKLM Edge `NewTabPageLocation` policy and require it to contain `https://armyeitaas.sharepoint-mil.us`.
- Remote Apps AVD-04 and AVD-05 are manual checks.
- Remote Apps AVD-09 and Azure Virtual Desktop AVD-32 read only HKLM Edge policy. `PromptForDownloadLocation` must equal `1`, and `DownloadDirectory` must target `C:\Users\${user_name}\OneDrive - US Army\Downloads` (a deployed username and typographic dash are also accepted).
- Remote Apps AVD-10 and Azure Virtual Desktop AVD-33 recursively query `HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge\managedfavorites` and pass when any key name, value name, or value data contains `Army Bookmarks`. No managed-policy JSON or object structure is required.

## Run once

```powershell
powershell.exe -NoProfile -NonInteractive -File .\AVD-UAT-Checklist-Automated.ps1
```

The default automated output is
`C:\Program Files\ArmyAVD\UAT\Automated_Results_<yyyyMMdd-HHmmss-fff>`.
The elevated/SYSTEM account must have write access there; manual users need
read access. Each invocation creates one new folder for the XLSX, CSV, HTML,
and troubleshooting log. `-OutputDirectory` overrides the results root, not
the per-run folder. Even an explicit `-LogPath` uses only its filename inside
that run folder.

For an alternate installation:

```powershell
powershell.exe -NoProfile -NonInteractive -File .\AVD-UAT-Checklist-Automated.ps1 `
  -UatDirectory 'C:\AVD-UAT-Test' `
  -ConfigurationFile '.\avd-uat.config.json'
```

Each run creates:

- `AVD UAT Automated Log - <timestamp>.log`: verbose file-only troubleshooting log.
- `AVD UAT Report - <timestamp>.html`: successful automated checks, checks requiring attention, and remaining manual checks.
- `AVD Checklist Results - <timestamp>.csv`: machine-readable evidence.
- `AVD Checklist - Completed - <timestamp>.xlsx`: copied checklist with Result, Issue Reference, Notes, and Automation Review populated.

## Manual completion

Run the follow-on workflow in the UAT user's interactive session, not SYSTEM.
It selects the newest timestamped `Automated_Results_*` folder under
`C:\Program Files\ArmyAVD\UAT`, and reads the completed XLSX and its matching
CSV from that same run. It never starts from the template or a Documents
workbook. An incomplete newest run raises an explicit error rather than
silently selecting stale evidence.

```powershell
.\AVD-UAT-Checklist-Manual.ps1
```

The rebuilt manual EXE uses the same selection rules. Existing explicit
`-InputFile`/`-AutomatedResultsCsv` values must match the latest run.
Only final manual outputs (XLSX, CSV, and HTML) go to
`%USERPROFILE%\OneDrive - US Army\Documents`; automation never writes there.
The automated source files remain unchanged.

The script prompts only for `Partial Pass` and `Manual` rows. `Skip` preserves the automated status. Completed review creates `AVD Checklist Final Results`, `AVD Checklist - Final`, and `AVD UAT Final Report` files without overwriting the automated outputs.

The same workflow can be started from `AVD-UAT-Checklist-Manual.exe`. Keep the executable beside the manual script, workbook helper, and matching `AVD Test Plan` workbook.

AVD-16 now validates the enforced Chrome `AIModeSettings=1` policy. AVD-18 runs in the signed-in user context, verifies the Copilot desktop and Start menu shortcuts, repairs either location by copying an existing approved shortcut, or creates `.url` shortcuts when `CopilotUrl` is configured. It reports `Pass` only when a matching Chrome or Edge favorite is also found; otherwise the browser favorite remains in manual review as `Partial Pass`.

## SYSTEM scheduled task

Startup tasks wait **5 minutes after boot** before launching UAT, allowing the
Intune Management Extension (IME) and other services time to initialize. This
delay applies to both PowerShell and executable task actions; daily triggers
are unchanged. The prerequisite gate still runs normally and stops UAT if IME
or another required prerequisite is not ready. A fixed delay is a startup
buffer, not a guarantee of service readiness.

Run the installer from an elevated PowerShell session:

```powershell
.\Install-AVDUATScheduledTask.ps1 -ExecutionMode PowerShell
```

To supply approved check configuration to the SYSTEM run:

```powershell
.\Install-AVDUATScheduledTask.ps1 `
  -ConfigurationFile '.\avd-uat.config.json' `
  -ExecutionMode PowerShell
```

For startup execution:

```powershell
.\Install-AVDUATScheduledTask.ps1 -ExecutionMode PowerShell
```

To register the compiled executable with the same startup trigger:

```powershell
.\Install-AVDUATScheduledTask.ps1 -ExecutionMode Executable
```

Preview registration without changing the machine:

```powershell
.\Install-AVDUATScheduledTask.ps1 -WhatIf
```

SYSTEM is appropriate for machine policy, registry, installation, DNS, and OS checks. It cannot validate the signed-in user's mapped drives, CurrentUser certificate store, OneDrive profile, GUI behavior, smart-card authentication, media playback, or application workflows. Those rows become `Manual` tasks in the report. Running the same script in the UAT user's context collects the additional user-scoped automated evidence; observation and authentication checks still remain manual.

After replacing an older deployment, rerun the installer with the original
task name to apply the five-minute startup delay and replace stale arguments
that point at Documents or an old output directory. Editing the installer alone
does not change an already registered task. Automated artifacts remain under the installation root; only final
manual outputs go to the UAT user's OneDrive Documents.
