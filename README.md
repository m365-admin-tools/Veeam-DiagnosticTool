# Veeam-DiagnosticTool
Veeam Diagnostic Toolkit (PowerShell GUI)

A read-only health check for Veeam Backup & Replication with a point-and-click interface. Pick the checks you want, set your thresholds, run, and get a colour-coded results grid where every warning and failure carries a specific recommendation for what to investigate and how to fix it.

The output doubles as a punch list. Export it as CSV for your own tracking, or as a self-contained HTML report for the customer or the ticket.

Please visit https://m365admintools.com for more IT engineering tools and information

<img width="771" height="679" alt="VeeamDiagToolsm" src="https://github.com/user-attachments/assets/26f08047-752c-462d-a255-19d56975f22b" />


## Requirements

| Item | Requirement |
|---|---|
| Veeam Backup & Replication | 12.x or 13.x |
| PowerShell module | `Veeam.Backup.PowerShell`, or the legacy `VeeamPSSnapIn`. Both ship with the VBR console |
| Where to run it | On the VBR server, or on a machine with the VBR console installed that can reach it |
| PowerShell edition | Chosen automatically. See below |
| Rights | Runs in the current user's security context. An account holding a Veeam Backup role. Local administrator rights are needed for the Veeam Services check |

**PowerShell edition is handled for you.** VBR 12.x and earlier ship a Windows PowerShell 5.1 module. VBR 13.0 and later re-platformed onto PowerShell 7 and the module refuses to load on 5.1. The script reads the installed VBR version from the registry, then relaunches itself under the correct edition on an STA thread, which WinForms requires. A second window opening briefly on start is that switch happening, not a fault.

## Quick start

```powershell
# Local VBR server
.\Veeam-DiagnosticTool.ps1

# Point it at another server
.\Veeam-DiagnosticTool.ps1 -Server veeam01.contoso.com
```

If the script is blocked on first run:

```powershell
Unblock-File .\Veeam-DiagnosticTool.ps1
```

Then, in the window:

1. Connect to the Veeam server. Localhost by default, or type another name and supply alternate credentials.
2. Choose checks from the categorized checklist, or use a preset: Quick Health Check, Select All, Select None.
3. Set the thresholds.
4. Run. Each check returns Pass, Warn, Fail, or Error with a one-line summary, a recommendation, and full detail in the pane below the grid.
5. Export CSV, HTML, or both.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Server` | string | `localhost` | VBR server to connect to. Can also be set in the window |

## Checks

| Category | Check | In Quick Health Check |
|---|---|---|
| Server & License | Backup server version and patch level | Yes |
| Server & License | License status, expiry, and usage | Yes |
| Jobs & Sessions | Job last-run results | Yes |
| Jobs & Sessions | RPO compliance, meaning backups older than the threshold | Yes |
| Jobs & Sessions | Recent session failures and warnings | Yes |
| Jobs & Sessions | Disabled jobs | Yes |
| Jobs & Sessions | Currently running and long-running jobs | No |
| Repositories | Repository free space | Yes |
| Repositories | Scale-out repository extents | Yes |
| Repositories | Immutability coverage | Yes |
| Infrastructure | Veeam services | Yes |
| Infrastructure | Backup proxies | Yes |
| Configuration Backup | Configuration backup health | Yes |
| Security & Best Practice | 3-2-1, backup copy jobs present | Yes |
| Security & Best Practice | Job encryption coverage | No |

The two checks outside the Quick Health Check preset are the slower or noisier ones. Select them individually when you need them.

## Thresholds

| Threshold | Default | Range | Used by |
|---|---|---|---|
| RPO hours | 24 | 1 to 720 | RPO compliance. A backup older than this is flagged |
| Session lookback hours | 24 | 1 to 720 | Recent session failures and warnings |
| Repository warning percent used | 85 | 50 to 99 | Repository free space, scale-out extents |
| Repository failure percent used | 95 | 51 to 100 | Repository free space, scale-out extents |
| License expiry warning days | 30 | 1 to 365 | License status |

The failure percentage must be higher than the warning percentage. The window validates this before running.

Set these to the customer's actual SLA rather than leaving the defaults. An environment with a 4 hour RPO and one with a weekly RPO both look healthy at 24 hours, which tells you nothing.

## Output

**Results grid.** Status, check name, summary, and recommendation per row, with the full detail for the selected row shown underneath.

**CSV export (F7).** The grid as data, for tracking findings across sites or over time.

**HTML report (F8).** A self-contained file suitable for sending to a customer or attaching to a ticket.

**Connect log.** A diagnostic log written to `Documents\VeeamDiagnosticTool.connect.log`, or to the temp folder if Documents is unavailable. It records the host PowerShell version, edition, apartment state, process ID, OS, and the account in use. The path is shown in the status bar at startup. Include this file when reporting a connection problem.

## Keyboard shortcuts

| Key | Action |
|---|---|
| F5 | Run |
| F6 | Cancel |
| F7 | Export CSV |
| F8 | Export HTML report |
| Esc | Close |

## What it changes

Nothing. Every check reads configuration and status. No job is created, modified, started, or stopped, no retention is altered, and no backup file is touched. The tool disconnects from the VBR server when the window closes.

## How to read the results

Status is a heuristic based on what Veeam reports plus your threshold settings. It is a triage aid, not a verdict.

- A **Fail** means something crossed a threshold or Veeam reported a problem state. Confirm it in the Veeam console before acting on it.
- A **Pass** means nothing obviously wrong was found in that check. It is not a certification of health, and it is not evidence that a restore works.
- An **Error** means the check itself could not complete, usually a rights or connectivity problem rather than a finding about the environment.

## Troubleshooting

**A second window opens briefly at launch**

Expected. The script is relaunching itself under the PowerShell edition your Veeam build requires, on an STA thread.

**`The version of PowerShell on this computer is '5.1...'. The module requires 7.0 to run`**

The VBR 13 module loaded under Windows PowerShell. This is what the automatic relaunch exists to prevent, so it usually means the version detection failed. Confirm that PowerShell 7 is installed, and start the script from `pwsh` directly.

**Cannot load the Veeam module**

Install the VBR console on this machine, or run the tool on the VBR server. Confirm with:

```powershell
Get-Module -ListAvailable Veeam.Backup.PowerShell
```

**Connection to the Veeam server fails**

Check that the Veeam Backup Service is running, that TCP 9392 is reachable, and that the account holds a Veeam Backup role. Supply alternate credentials in the window for a remote server. The connect log records each attempt with the account in use.

**The Veeam Services check returns an error for a remote server**

That check queries Windows services on the target, which needs remote service query rights and the relevant firewall rules. Run the tool on the VBR server to avoid it, or deselect that check.

**A check returns Error rather than a result**

The account lacks rights to that object type, or the cmdlet is not present on that VBR build. Other checks still run and the grid shows which one failed.

## Known limitations

- Immutability coverage reports what the repositories are configured for. It does not verify that existing restore points on disk are actually immutable.
- The 3-2-1 check confirms that backup copy jobs exist. It does not confirm that the copies are current, complete, or genuinely offsite.
- All findings come from the VBR configuration database at the moment of the run. This is a point-in-time triage tool, not monitoring, and no check substitutes for a tested restore.
- The interface is WinForms, so it is Windows only.

## Related

- Free Microsoft 365, Active Directory, and Veeam tools at [m365admintools.com](https://m365admintools.com)

## Author

Charles Arconi, [m365admintools.com](https://m365admintools.com)

Not affiliated with, endorsed by, or supported by Veeam Software. Veeam is a trademark of Veeam Software Group GmbH.

## License

MIT. See [LICENSE](LICENSE).
