# AVD Automation

Automation and supporting resources for Army Azure Virtual Desktop (AVD) Windows 11 environments. This repository brings together image-building infrastructure, application installation and image configuration scripts, exported security policies, and automated/manual user acceptance testing (UAT).

## Repository at a glance

| Area | Contents | Start here |
| --- | --- | --- |
| Image engineering | Packer templates, Terraform infrastructure, image distribution workflows, provisioning scripts, and Evaluate-STIG resources. | [Windows 11 image project](AVD_Repo_Files/army-avd-image-win11/) |
| Security policy exports | Device/user configuration JSON, STIG settings, WDAC/App Control XML, and certificate detection/remediation scripts. | [Policy exports](AVD_Repo_Files/_exports/) |
| Software deployment | A package installation orchestrator and individual application, certificate, language, optimization, and maintenance scripts. | [Software installer](Software_Installer/Install-Software.ps1) |
| Acceptance testing | Automated checks, guided manual review, scheduled task installation, configuration, workbook template, reports, and packaged launchers. | [UAT documentation](UAT/README.md) |

## Repository layout

```text
AVD_Automation/
|-- README.md
|-- AVD_Repo_Files/
|   |-- _exports/                    Security policies and certificate scripts
|   `-- army-avd-image-win11/         Nested image-build Git project
|       |-- .github/workflows/       Build, infrastructure, and distribution workflows
|       |-- packer/                  Image templates and Office configuration
|       |-- terraform/               Image-based VM and networking resources
|       |-- terraform-infra/         Resource group and storage infrastructure
|       |-- boostrap_tf/             Terraform state bootstrap resources
|       |-- scripts/                 Image provisioning and maintenance scripts
|       `-- Evaluate-STIG/           Assessment tooling and supporting resources
|-- Software_Installer/
|   |-- Install-Software.ps1         Canonical package installation entry point
|   `-- scripts/                     Individual installers and configuration helpers
|       `-- Logs/                    Existing installation summaries and transcripts
`-- UAT/
    |-- README.md                    Detailed deployment and operator instructions
    |-- AVD-UAT-Checklist*.ps1        Check engine and automated/manual launchers
    |-- AVD-UAT-Checklist*.exe        Packaged automated/manual launchers
    |-- AVD-UAT-Prerequisite-Check.ps1
    |-- AVD-UAT-Paths.ps1
    |-- Install-AVDUATScheduledTask.ps1
    |-- Pilot_Run.ps1
    |-- avd-uat.config*.json
    |-- army-bookmarks.json
    |-- tools/AvdChecklistTool/       Workbook helper executable
    `-- ...                          Workbook, runbook, certificate, ZIP, and hashes
```

## 1. Windows 11 image engineering

The [image-build project](AVD_Repo_Files/army-avd-image-win11/) contains:

- **Packer:** Windows image definitions, shared variables, Microsoft 365 Office configuration, and separate `base`, `general`, `inject`, `rxapps`, `STIGS`, and `updates` directories.
- **Terraform:** VM deployment from an Azure Compute Gallery image, networking and security resources, plus a separate infrastructure configuration for resource groups and storage.
- **State bootstrap:** [BOOTSTRAP-TFSTATE.ps1](AVD_Repo_Files/army-avd-image-win11/BOOTSTRAP-TFSTATE.ps1), the [boostrap_tf directory](AVD_Repo_Files/army-avd-image-win11/boostrap_tf/), and backend configuration resources.
- **GitHub Actions:** Build/test, development infrastructure, sandbox distribution, development distribution, and production base/general/update distribution and replication workflows.
- **Provisioning scripts:** Application installation, certificate staging, local policy, FSLogix, language packs, Windows updates, branding, AppX maintenance, and desktop optimization.
- **Evaluate-STIG:** PowerShell and Bash entry points, answer-file management, modules, checklist templates, certificates, prerequisites, XML resources, and bundled reference documentation.
- **Build metadata:** A `compute_build_number` helper and project-specific licensing.

The included workflows target Azure Government and reference environment-specific subscriptions, galleries, storage, credentials, and external artifacts. Review those dependencies and the selected Packer/Terraform variables before building or deploying. They are not a ready-to-run configuration for an arbitrary Azure tenant.

**Checkout note:** This directory is recorded in the parent repository as a Git link to a separate commit. This checkout contains its files, but there is no root `.gitmodules` mapping. A fresh clone may not populate it automatically; obtain the matching image project from the repository maintainer rather than assuming the parent clone includes all image-build content. Changes inside that project are managed separately from changes in the parent repository.

The [image project README](AVD_Repo_Files/army-avd-image-win11/readme.md) provides background; use the actual HCL templates and workflow definitions for current paths and execution settings.

## 2. Security policies and configuration exports

The [exports directory](AVD_Repo_Files/_exports/) contains JSON policy exports for:

- Microsoft Defender Antivirus, Defender Firewall, firewall rules, and Attack Surface Reduction (ASR).
- Microsoft Edge and Google Chrome.
- Microsoft Office 2019/Microsoft 365 and OneDrive.
- Windows 11 operating system settings.
- General device and user configuration, including user-scoped Office and Windows settings.

It also includes four XML application-control policy files covering a base policy, additional policy, and AVD product base/supplemental policies, plus [certificate detection](AVD_Repo_Files/_exports/EMDVTLSCACert-Detect.ps1) and [certificate remediation](AVD_Repo_Files/_exports/EMDVTLSCACert-Remediate.ps1) scripts.

Export filenames identify policy scope and, where applicable, STIG revisions or policy versions. Treat these files as configuration snapshots, not proof of current compliance. Review policy settings, assignments, certificate trust decisions, and compatibility with the target environment before import or deployment.

## 3. Software installation and image configuration

### Package orchestrator

[Install-Software.ps1](Software_Installer/Install-Software.ps1) requires Windows PowerShell 5.1 or later and coordinates a defined package catalog with vendor-specific silent installation arguments. It supports:

- Package selection with `-Include` and `-Exclude`.
- Catalog inspection with `-ListOnly` and simulated execution with `-WhatIf`.
- Automatic elevation for installation, unless `-NoAutoElevate` is specified.
- EXE/MSI, Office, PowerShell, batch, portable application, and package-specific deployment handling.
- Timestamped CSV summaries and installation transcripts under `Software_Installer\Logs`.
- Recognition of successful installer exit codes `0`, `1641`, and `3010`, including reboot-required outcomes.

From the repository root, inspect the catalog before installing:

```powershell
.\Software_Installer\Install-Software.ps1 -ListOnly
.\Software_Installer\Install-Software.ps1 -WhatIf
.\Software_Installer\Install-Software.ps1 -Include '7-Zip','Mozilla Firefox' -WhatIf
```

After staging and approving the required packages, run a selected installation:

```powershell
.\Software_Installer\Install-Software.ps1 -Include '7-Zip','Mozilla Firefox'
```

**Installer payloads are not a complete part of this repository.** The orchestrator resolves its software root to its own directory when `7z2603-x64.exe` is present there; otherwise, it uses the parent directory. Review the package catalog, paths, licensing requirements, and silent switches, then stage the expected payloads before execution. Some individual scripts also depend on external storage, network access, or supporting tools.

The [entry point under scripts](Software_Installer/scripts/Install-Software.ps1) forwards arguments to the canonical orchestrator; it is not a separate package catalog.

### Individual scripts

The [scripts directory](Software_Installer/scripts/) also provides standalone helpers in these categories:

| Category | Included coverage |
| --- | --- |
| Browsers and utilities | Chrome, Firefox, 7-Zip, Notepad/Notepad++, VLC, PowerShell 7, Windows App Runtime, Visual C++ redistributables, and WinGet preparation. |
| Microsoft productivity | Microsoft 365 Apps, Teams and its Outlook add-in, Copilot, Power Automate, Power BI Desktop, Power BI Report Builder, Store updates, and Photos. |
| AVD and profiles | FSLogix installation and redirections, WebRTC redirection, Remote Desktop installer staging, user logon updates, and virtual desktop optimization. |
| Army/mission applications | AWS WickrGov, Cisco Jabber, DAX Studio, Foundry ODBC, FedLog, IADS4, DEPTH, DTS Lite, OCE Tool, SAP GUI, and self-help tooling. |
| Security and management | Certificates and trust-chain staging, Purebred registration, YubiKey, S/MIME, Tychon, PrinterLogic, STIG Viewer, classification banners, LGPO, ADMX, and DSC configuration. |
| Language and appearance | Language pack installation, including German, French, Japanese, and Korean helpers, themes/wallpapers, image branding, and public desktop icon cleanup. |
| Image maintenance | Blob downloads and installer pinning, gallery image/version lookups, Windows Update controls, Secure Boot certificate updates, BitLocker decryption, AppX State Repository repair, and cleanup. |

Versioned alternatives are retained for some installers and staging helpers. Select the script appropriate to the intended build rather than executing every file. Existing [historical logs](Software_Installer/scripts/Logs/) are prior run artifacts, not current deployment results.

## 4. User acceptance testing

The [UAT suite](UAT/) separates unattended machine checks from interactive user verification. See the [UAT README](UAT/README.md) for configuration parameters, detailed check behavior, deployment instructions, output naming, and execution-context limitations.

### Main components

| Component | Purpose |
| --- | --- |
| [Automated launcher](UAT/AVD-UAT-Checklist-Automated.ps1) | Runs unattended checks through the shared engine and produces automated evidence. An EXE launcher is also included. |
| [Manual launcher](UAT/AVD-UAT-Checklist-Manual.ps1) | Opens the latest automated evidence and prompts for remaining manual/partial results in the UAT user's session. An EXE launcher is also included. |
| [Shared check engine](UAT/AVD-UAT-Checklist.ps1) | Collects evidence and produces HTML, CSV, and completed workbook reports. |
| [Prerequisite checker](UAT/AVD-UAT-Prerequisite-Check.ps1) | Gates both launchers on Entra join, MDM enrollment/tasks, running Intune Management Extension, and running FSLogix. |
| [Path helper](UAT/AVD-UAT-Paths.ps1) | Shared UAT path-resolution logic. |
| [Scheduled task installer](UAT/Install-AVDUATScheduledTask.ps1) | Registers SYSTEM execution using PowerShell or the executable launcher; startup execution waits five minutes after boot. |
| [Pilot runner](UAT/Pilot_Run.ps1) | Invokes the automated launcher with the included configuration and a short performance sampling interval. |
| [Configuration example](UAT/avd-uat.config.example.json) and [deployment configuration](UAT/avd-uat.config.json) | Environment-specific check settings and approved artifact paths. |
| [Army bookmarks](UAT/army-bookmarks.json) | Supporting bookmark baseline data; see the UAT documentation for current managed-policy checks. |
| [Workbook helper](UAT/tools/AvdChecklistTool/AvdChecklistTool.exe) | Reads and updates XLSX packages without requiring Excel COM or an Office installation for workbook I/O. |

### Supporting release resources

- [AVD test plan template](UAT/AVD%20Test%20Plan%20Template%20-%20DO%20NOT%20DELETE%20OR%20MOVE.xlsx): source checklist; preserve its filename and placement.
- [Deployment and operator runbook](UAT/AVD-UAT-Deployment-and-Operator-Runbook.docx): Word-format operating instructions.
- [Live-test ZIP](UAT/AVD-UAT-Live-Test-Final.zip): packaged deployment bundle.
- [Code-signing certificate](UAT/AVD-CodeSigning.cer) and [SHA-256 manifest](UAT/SHA256SUMS.txt): release verification resources. Their presence alone does not establish organizational trust or prove that the bundle matches the current scripts.

### Execution and results

The normal deployment root is `C:\Program Files\ArmyAVD\UAT`. Stage the full matching runtime set, including the workbook, helper, configuration, and prerequisite/path scripts, before running the launchers.

For an alternate checkout-based test, from the repository root:

```powershell
.\UAT\AVD-UAT-Checklist-Automated.ps1 `
    -UatDirectory "$PWD\UAT" `
    -ConfigurationFile "$PWD\UAT\avd-uat.config.json"
```

This executes real checks and requires the documented prerequisites and permissions; it is not a dry run. Follow the UAT documentation for production deployment and manual completion.

- Automated runs create a new `Automated_Results_<timestamp>` directory beneath the installation root by default, containing HTML, CSV, XLSX, and a troubleshooting log, plus prerequisite evidence.
- Manual review uses the newest automated result set and preserves its source files. Final manual HTML, CSV, and XLSX outputs go to the UAT user's `OneDrive - US Army\Documents`.
- Results distinguish `Pass`, `Partial Pass`, `Fail`, and `Manual`. Authentication, visual checks, and user workflows still require an interactive standard-user session where documented.
- SYSTEM runs cannot establish every user-scoped result. Do not treat automated prerequisite evidence as a substitute for completing the manual checklist.

## Operational guidance

- Use Windows/PowerShell and the privileges required by the selected script. Azure image builds additionally require Packer, Terraform, appropriate Azure tooling/authentication, and access to configured dependencies.
- Review all environment-specific paths, URLs, application versions, storage references, policy settings, and deployment configuration before use.
- Keep credentials, smart-card PINs, tokens, private keys, and other secrets out of scripts, JSON configuration, logs, and source control.
- Validate signatures and hashes against an approved release. Follow organizational AllSigned, WDAC, and AppLocker requirements; do not bypass controls or automatically trust the bundled certificate.
- After changing packaged UAT code, rebuild/sign the affected runtime artifacts and refresh the ZIP/hash manifest as described in the [UAT deployment guidance](UAT/README.md).
- Scripts can install software, change certificate stores or security settings, alter update behavior, decrypt drives, or modify infrastructure. Test selected workflows on a non-production image before broad deployment.
- Review applicable license files and vendor terms in the nested image project and bundled tooling. Inclusion in this repository does not grant licenses for third-party software or installation payloads.
