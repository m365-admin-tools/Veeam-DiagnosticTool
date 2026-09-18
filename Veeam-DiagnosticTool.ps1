<#
.SYNOPSIS
    M365admintools.com - Author Charles Arconi - created 8/12/2026
    Veeam Diagnostic Toolkit 1.0 - read-only health check for Veeam Backup & Replication
    12.x / 13.x with a checklist-driven GUI, a color-coded results grid, per-finding
    recommendations and an HTML customer report.

.DESCRIPTION
    Companion to the AD Diagnostic Toolkit. Connects to a Veeam Backup & Replication server
    through the Veeam PowerShell module and runs a catalog of read-only health checks:
    server version / patch level, licensing, job last-run results, RPO compliance, recent
    session failures, repositories and scale-out free space, immutability coverage, backup
    proxies, Veeam services, configuration backup, and a set of 3-2-1 / security best-
    practice flags. Every Warn or Fail carries a concrete "what to investigate / how to
    fix" recommendation, so the output doubles as a punch list.

    Workflow:
      1. Connect to the Veeam server (localhost by default, or type another server name;
         optional alternate credentials).
      2. Pick which checks to run from the categorized checklist, or use a preset:
         Quick Health Check, Select All, Select None.
      3. Set thresholds (RPO hours, repo free-space warn/fail %, license expiry warning).
      4. Run. Each result gets a Pass / Warn / Fail / Error status, a one-line summary, a
         recommendation, and full detail in the pane below the grid.
      5. Export a CSV of the grid and/or a self-contained HTML report for the customer or
         the ticket.

    Runs in the current user's security context. Install and run it ON the VBR server, or
    on a machine with the Veeam Backup & Replication console (which supplies the PowerShell
    module) that can reach the VBR server. No configuration is changed -- this tool only
    reads.

    Status is a HEURISTIC based on Veeam-reported state plus threshold comparisons. It is a
    triage aid, not a verdict -- review the detail and confirm in the Veeam console before
    acting on a Fail, and treat a Pass as "nothing obviously wrong found here", not
    "certified healthy".

.PARAMETER Server
    Optional VBR server name to connect to. Defaults to localhost.

.EXAMPLE
    .\Veeam-DiagnosticTool.ps1

.EXAMPLE
    .\Veeam-DiagnosticTool.ps1 -Server veeam01.contoso.com

.NOTES
    PowerShell edition is chosen automatically to match the installed Veeam build:
      * VBR 12.x and earlier -> Windows PowerShell 5.1
      * VBR 13.0+           -> PowerShell 7.x (the v13 module requires it)
    The script detects the VBR version and relaunches itself under the correct edition on
    an STA thread, so it can be started from either PowerShell. Requires the Veeam Backup &
    Replication PowerShell module (Veeam.Backup.PowerShell, or the legacy VeeamPSSnapIn),
    i.e. run it on the VBR server or a machine with the Veeam console installed.
    Read-only. Safe to run against production. Always confirm findings in the console
    before making changes.
#>

[CmdletBinding()]
param(
    [string]$Server = 'localhost'
)

# ---------------------------------------------------------------------------
# Launch under the PowerShell edition the installed Veeam module requires, on an
# STA thread (WinForms requires STA).
#
#   * Veeam VBR 12.x and earlier ship a *Windows PowerShell 5.1* module.
#   * Veeam VBR 13.0+ re-platformed onto .NET / PowerShell 7, so its module
#     requires *PowerShell 7.x* and refuses to load on Windows PowerShell 5.1
#     ("The version of PowerShell on this computer is '5.1...'. The module
#     requires 7.0 to run").
#
# We detect the installed VBR major version from the registry and, if the
# current host is the wrong edition or not STA, relaunch this same script with
# the correct powershell/pwsh and -STA, then stop this instance. That is why a
# second window may briefly open on start -- it is switching to the edition your
# Veeam build needs.
# ---------------------------------------------------------------------------
$reqEdition = $null       # 'Core' (PowerShell 7) or 'Desktop' (Windows PowerShell 5.1)
$vbrMajor   = $null
foreach ($key in @('HKLM:\SOFTWARE\Veeam\Veeam Backup and Replication',
                   'HKLM:\SOFTWARE\Wow6432Node\Veeam\Veeam Backup and Replication')) {
    try {
        $core = (Get-ItemProperty -Path $key -ErrorAction Stop).CorePath
        if ($core) {
            $dll = Join-Path $core 'Veeam.Backup.Core.dll'
            if (Test-Path -LiteralPath $dll) {
                $pv = (Get-Item -LiteralPath $dll).VersionInfo.ProductVersion
                if ($pv) {
                    $m = [regex]::Match($pv, '^\d+(\.\d+){0,3}')
                    if ($m.Success) {
                        $vs = $m.Value; if ($vs -notmatch '\.') { $vs = "$vs.0" }
                        $vbrMajor = ([version]$vs).Major; break
                    }
                }
            }
        }
    } catch { }
}
if     ($vbrMajor -ge 13) { $reqEdition = 'Core' }
elseif ($vbrMajor -ge 1)  { $reqEdition = 'Desktop' }

# Fallback if the product version was not readable: ask the module manifest.
if (-not $reqEdition) {
    try {
        $mod = Get-Module -ListAvailable -Name Veeam.Backup.PowerShell -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($mod) {
            $ce = @($mod.CompatiblePSEditions)
            if (($mod.PowerShellVersion -and $mod.PowerShellVersion.Major -ge 7) -or
                (($ce -contains 'Core') -and ($ce -notcontains 'Desktop'))) { $reqEdition = 'Core' }
            else { $reqEdition = 'Desktop' }
        }
    } catch { }
}

$curEdition = if ($PSVersionTable.PSEdition -eq 'Core') { 'Core' } else { 'Desktop' }
if (-not $reqEdition) { $reqEdition = $curEdition }   # unknown: keep edition, only fix STA
$curSTA = ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA)

if (($curEdition -ne $reqEdition) -or (-not $curSTA)) {
    $exe = $null
    if ($reqEdition -eq 'Core') {
        $exe = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source
        if (-not $exe) {
            foreach ($p in @("$env:ProgramFiles\PowerShell\7\pwsh.exe",
                             "$env:LOCALAPPDATA\Microsoft\WindowsApps\pwsh.exe")) {
                if (Test-Path -LiteralPath $p) { $exe = $p; break }
            }
        }
    } else {
        $exe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    }

    if ($exe -and (Test-Path -LiteralPath $exe) -and $PSCommandPath) {
        # Quote explicitly: Start-Process does not quote ArgumentList elements and
        # this script's path contains spaces ("OneDrive - Arconi ...").
        $argv = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File', ('"{0}"' -f $PSCommandPath))
        if ($PSBoundParameters.ContainsKey('Server')) { $argv += @('-Server', ('"{0}"' -f $Server)) }
        Start-Process -FilePath $exe -ArgumentList $argv | Out-Null
        return
    } else {
        $need = if ($reqEdition -eq 'Core') { 'PowerShell 7.x' } else { 'Windows PowerShell 5.1' }
        $msg  = "This tool must run in $need on an STA thread for your Veeam build" +
                $(if ($vbrMajor) { " (VBR major version $vbrMajor)" } else { '' }) +
                ", but the matching PowerShell executable could not be located.`r`n`r`nStart the script from $need directly."
        try {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show($msg,'PowerShell version','OK','Error') | Out-Null
        } catch { Write-Warning $msg }
        return
    }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
try { Add-Type -AssemblyName Microsoft.VisualBasic } catch { }
[System.Windows.Forms.Application]::EnableVisualStyles()

# ===========================================================================
# State
# ===========================================================================

$script:Ctx = [pscustomobject]@{
    Server        = $Server
    Connected     = $false
    Credential    = $null
    ModuleKind    = $null      # 'module' | 'snapin'
    ProductName   = $null
    ProductBuild  = $null
    PatchLevel    = $null
    Edition       = $null
    LicenseTo     = $null
}

$script:Results         = New-Object System.Collections.Generic.List[object]
$script:CancelRequested = $false
$script:SuppressChecks  = $false
$script:Cache           = @{}   # per-run cache of expensive Get-VBR* calls

# ---------------------------------------------------------------------------
# Header logo (ScriptLogo.png, 273x35 32-bit ARGB, transparent background)
# embedded as Base64 so the script stays self-contained and portable -- it runs
# on customer / VBR servers where the original image path does not exist.
# ---------------------------------------------------------------------------
$script:LogoBase64 = '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'

function Get-LogoImage {
    <# Decodes the embedded header logo into a standalone Bitmap. A full copy is
       returned so the backing MemoryStream can be disposed immediately -- an Image
       created directly from a stream requires that stream to stay open for the
       image's entire lifetime, which is an easy WinForms leak/AV to trip over. #>
    $bytes = [Convert]::FromBase64String($script:LogoBase64)
    $ms  = New-Object System.IO.MemoryStream(,$bytes)
    try {
        $src = [System.Drawing.Image]::FromStream($ms)
        try { return (New-Object System.Drawing.Bitmap $src) }
        finally { $src.Dispose() }
    } finally { $ms.Dispose() }
}

# ===========================================================================
# Check catalog
#
# Id       : stable key, stashed in the TreeNode.Tag for the checklist
# Category : groups checks under a TreeView parent node
# Label    : what the technician sees
# Action   : dispatch key into Invoke-VeeamCheck
# Default  : checked out of the box / included in the "Quick Health Check" preset
#
# A single check may return MANY result rows (e.g. one per job or per repository).
# ===========================================================================

$script:CheckCatalog = @(
    @{ Id='sv-version';   Category='Server & License';        Label='Backup Server Version & Patch Level'; Action='Version';        Default=$true  }
    @{ Id='sv-license';   Category='Server & License';        Label='License Status, Expiry & Usage';      Action='License';        Default=$true  }

    @{ Id='jb-results';   Category='Jobs & Sessions';         Label='Job Last-Run Results';                Action='JobResults';     Default=$true  }
    @{ Id='jb-rpo';       Category='Jobs & Sessions';         Label='RPO Compliance (stale backups)';      Action='RpoCompliance';  Default=$true  }
    @{ Id='jb-recent';    Category='Jobs & Sessions';         Label='Recent Session Failures / Warnings';  Action='RecentSessions'; Default=$true  }
    @{ Id='jb-disabled';  Category='Jobs & Sessions';         Label='Disabled Jobs';                       Action='DisabledJobs';   Default=$true  }
    @{ Id='jb-running';   Category='Jobs & Sessions';         Label='Currently Running / Long-Running';    Action='RunningJobs';    Default=$false }

    @{ Id='rp-space';     Category='Repositories';            Label='Repository Free Space';               Action='RepoFreeSpace';  Default=$true  }
    @{ Id='rp-sobr';      Category='Repositories';            Label='Scale-Out Repository Extents';        Action='SobrHealth';     Default=$true  }
    @{ Id='rp-immut';     Category='Repositories';            Label='Immutability Coverage';               Action='Immutability';   Default=$true  }

    @{ Id='in-services';  Category='Infrastructure';          Label='Veeam Services';                      Action='Services';       Default=$true  }
    @{ Id='in-proxies';   Category='Infrastructure';          Label='Backup Proxies';                      Action='Proxies';        Default=$true  }

    @{ Id='cb-config';    Category='Configuration Backup';    Label='Configuration Backup Health';         Action='ConfigBackup';   Default=$true  }

    @{ Id='bp-321';       Category='Security & Best Practice'; Label='3-2-1: Backup Copy Jobs Present';    Action='BackupCopy';     Default=$true  }
    @{ Id='bp-encrypt';   Category='Security & Best Practice'; Label='Job Encryption Coverage';           Action='JobEncryption';  Default=$false }
) | ForEach-Object {
    [pscustomobject]@{
        Id       = $_['Id']
        Category = $_['Category']
        Label    = $_['Label']
        Action   = $_['Action']
        Default  = [bool]$_['Default']
    }
}

# ===========================================================================
# Small helpers
# ===========================================================================

function Get-Prop {
    <# Safe property read. Under Set-StrictMode -Version Latest, touching a property
       that does not exist on an object throws; Veeam object shapes drift between
       v12 and v13, so every uncertain read goes through here and yields $null on
       failure instead of aborting the whole check. #>
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    try { return $Obj.$Name } catch { return $null }
}

function Invoke-Safe {
    <# Runs a scriptblock and returns its result, or $null if it throws. Used to wrap
       method calls (job.GetLastResult(), repo.GetContainer(), ...) and every Veeam
       cmdlet read that may not exist or may fail on a given build.

       It also strips the thread's SynchronizationContext for the duration. Veeam
       v13 cmdlets use async/await internally; run on the WinForms UI thread (which
       carries a WindowsFormsSynchronizationContext) their await continuations get
       posted back to that thread while it is blocked waiting -> permanent deadlock
       (the End-Task hang). A plain console has no such context, which is why the
       same cmdlet completes there in ~20-30s. Nulling the context reproduces the
       console's behaviour: continuations run on the thread pool and the call
       returns. #>
    param([scriptblock]$Script)
    $prevCtx = [System.Threading.SynchronizationContext]::Current
    try {
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
        return (& $Script)
    } catch { return $null }
    finally { [System.Threading.SynchronizationContext]::SetSynchronizationContext($prevCtx) }
}

function Format-Bytes {
    param($Bytes)
    if ($null -eq $Bytes) { return 'n/a' }
    try { $b = [double]$Bytes } catch { return 'n/a' }
    if ($b -le 0) { return '0 B' }
    $units = 'B','KB','MB','GB','TB','PB'
    $i = 0
    while ($b -ge 1024 -and $i -lt ($units.Count-1)) { $b /= 1024; $i++ }
    return ('{0:N1} {1}' -f $b, $units[$i])
}

function Get-VbData {
    <# Lazily loads and caches an expensive collection for the duration of one run.
       The cache is cleared at the start of every Run so a re-run always sees fresh
       data. #>
    param([string]$Key, [scriptblock]$Loader)
    if (-not $script:Cache.ContainsKey($Key)) {
        $script:Cache[$Key] = & $Loader
    }
    return $script:Cache[$Key]
}

function Get-CachedJobs        { Get-VbData 'jobs'      { @(Invoke-Safe { Get-VBRJob }) | Where-Object { $_ } } }
function Get-CachedSessions    { Get-VbData 'sessions'  { @(Invoke-Safe { Get-VBRBackupSession }) | Where-Object { $_ } } }
function Get-CachedRepos       { Get-VbData 'repos'     { @(Invoke-Safe { Get-VBRBackupRepository }) | Where-Object { $_ } } }
function Get-CachedSobrs       { Get-VbData 'sobrs'     { @(Invoke-Safe { Get-VBRBackupRepository -ScaleOut }) | Where-Object { $_ } } }
function Get-CachedProxies     {
    Get-VbData 'proxies' {
        $p = @()
        $p += @(Invoke-Safe { Get-VBRViProxy }) | Where-Object { $_ }
        $p += @(Invoke-Safe { Get-VBRHvProxy }) | Where-Object { $_ }
        ,$p
    }
}
function Get-CachedConfigJob   { Get-VbData 'configjob' { Invoke-Safe { Get-VBRConfigurationBackupJob } } }
function Get-CachedLicense     { Get-VbData 'license'   { Invoke-Safe { Get-VBRInstalledLicense } } }

# ===========================================================================
# Connection layer
# ===========================================================================

$script:ConnectLogPath = $null

function Write-ConnectLog {
    <# Writes one timestamped line to the connect diagnostic log and closes the
       file each call, so the line is flushed to disk immediately. That is the
       whole point: if Connect hangs hard enough to need End Task, whichever line
       is LAST in this file is exactly where it froze. #>
    param([string]$Message)
    if (-not $script:ConnectLogPath) { return }
    $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Message
    try { Add-Content -LiteralPath $script:ConnectLogPath -Value $line -Encoding UTF8 } catch { }
}

function Test-VbrPort {
    <# Bounded TCP reachability test for the Veeam Backup Service port. Never
       blocks longer than TimeoutMs, so it can flag an unreachable service before
       Connect-VBRServer (which can otherwise wait a very long time). #>
    param([string]$ComputerName, [int]$Port = 9392, [int]$TimeoutMs = 5000)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($ComputerName, $Port, $null, $null)
        $ok  = $iar.AsyncWaitHandle.WaitOne($TimeoutMs)
        if ($ok -and $client.Connected) { $client.EndConnect($iar); return $true }
        return $false
    } catch { return $false } finally { $client.Close() }
}

function Initialize-ConnectLog {
    <# Establishes the connect diagnostic log and writes a startup banner. Called
       at launch so the file is never blank once the app is running, and the exact
       path is echoed to the console window behind the form. #>
    if (-not $script:ConnectLogPath) {
        $dir = Join-Path $env:USERPROFILE 'Documents'
        try { if (-not (Test-Path -LiteralPath $dir)) { $dir = $env:TEMP } } catch { $dir = $env:TEMP }
        $script:ConnectLogPath = Join-Path $dir 'VeeamDiagnosticTool.connect.log'
    }
    Write-ConnectLog '==================================================================='
    Write-ConnectLog 'STARTUP'
    Write-ConnectLog ('Script : {0}' -f $PSCommandPath)
    Write-ConnectLog ('Host   : PS {0}  Edition={1}  Apartment={2}  PID={3}' -f `
        $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, [System.Threading.Thread]::CurrentThread.GetApartmentState(), $PID)
    Write-ConnectLog ('OS     : {0}' -f [System.Environment]::OSVersion.VersionString)
    Write-ConnectLog ('RunAs  : {0}' -f [System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    try { Write-Host "Veeam Diagnostic -- connect log: $script:ConnectLogPath" -ForegroundColor Cyan } catch { }
}

function Import-VeeamModule {
    if ($script:Ctx.ModuleKind) { Write-ConnectLog "Import-VeeamModule: already loaded (kind=$($script:Ctx.ModuleKind))"; return }
    Write-ConnectLog 'Import-VeeamModule: check for already-imported module'
    if (Get-Module -Name Veeam.Backup.PowerShell) { $script:Ctx.ModuleKind = 'module'; Write-ConnectLog 'Import-VeeamModule: already imported'; return }

    Write-ConnectLog 'Import-VeeamModule: Get-Module -ListAvailable BEGIN'
    $avail = Get-Module -ListAvailable -Name Veeam.Backup.PowerShell | Select-Object -First 1
    Write-ConnectLog 'Import-VeeamModule: Get-Module -ListAvailable DONE'
    if ($avail) {
        Write-ConnectLog ("Import-VeeamModule: found v{0} at '{1}' (ManifestPSVer={2}; Editions={3})" -f `
            $avail.Version, $avail.ModuleBase, $avail.PowerShellVersion, (@($avail.CompatiblePSEditions) -join ','))
        Write-ConnectLog 'Import-VeeamModule: Import-Module BEGIN'
        Import-Module Veeam.Backup.PowerShell -ErrorAction Stop -WarningAction SilentlyContinue
        Write-ConnectLog 'Import-VeeamModule: Import-Module DONE'
        $script:Ctx.ModuleKind = 'module'; return
    }

    # Legacy snap-in path. Get-PSSnapin does not exist in PowerShell 7, so guard it.
    Write-ConnectLog 'Import-VeeamModule: module not available; trying legacy VeeamPSSnapIn'
    if (Get-Command Get-PSSnapin -ErrorAction SilentlyContinue) {
        $snap = Get-PSSnapin -Registered -Name VeeamPSSnapIn -ErrorAction SilentlyContinue
        if ($snap) {
            if (-not (Get-PSSnapin -Name VeeamPSSnapIn -ErrorAction SilentlyContinue)) {
                Write-ConnectLog 'Import-VeeamModule: Add-PSSnapin BEGIN'
                Add-PSSnapin VeeamPSSnapIn -ErrorAction Stop
                Write-ConnectLog 'Import-VeeamModule: Add-PSSnapin DONE'
            }
            $script:Ctx.ModuleKind = 'snapin'; return
        }
    }
    Write-ConnectLog 'Import-VeeamModule: NO Veeam module or snap-in found'
    throw "The Veeam PowerShell module was not found. Install the Veeam Backup & Replication console (or run this on the VBR server) so that 'Veeam.Backup.PowerShell' or the 'VeeamPSSnapIn' snap-in is available."
}

function Connect-Veeam {
    param([string]$TargetServer, [pscustomobject]$Credential)

    $srv = if ([string]::IsNullOrWhiteSpace($TargetServer)) { 'localhost' } else { $TargetServer.Trim() }

    Write-ConnectLog 'Connect-Veeam: Import-VeeamModule BEGIN'
    Import-VeeamModule
    Write-ConnectLog "Connect-Veeam: Import-VeeamModule DONE (kind=$($script:Ctx.ModuleKind))"

    # Bounded reachability check for the Veeam Backup Service port. This never
    # hangs, so even if it fails we learn something. (Default port 9392; a custom
    # port would show False here yet may still connect, so this is a clue, not a
    # gate.)
    Write-ConnectLog "Connect-Veeam: TCP preflight ${srv}:9392 BEGIN"
    $portOk = Test-VbrPort -ComputerName $srv -Port 9392 -TimeoutMs 5000
    Write-ConnectLog "Connect-Veeam: TCP preflight ${srv}:9392 = $portOk"

    # Drop any prior session so a re-connect to a different server is clean.
    Write-ConnectLog 'Connect-Veeam: Disconnect prior session BEGIN'
    Invoke-Safe { Disconnect-VBRServer -ErrorAction SilentlyContinue } | Out-Null
    Write-ConnectLog 'Connect-Veeam: Disconnect prior session DONE'

    Write-ConnectLog "Connect-Veeam: Connect-VBRServer BEGIN (server='$srv'; creds=$(if ($Credential) { $Credential.UserName } else { '(current identity)' }))"
    # Strip the WinForms SynchronizationContext for the duration of the connect.
    # v13's Connect-VBRServer is async internally; on the UI thread its await
    # continuations post back to this (blocked) thread and deadlock forever. With
    # no context they run on the thread pool, so it completes as it does in a plain
    # console (~20-30s). The window is briefly unresponsive during the call --
    # expected, not a hang.
    $prevCtx = [System.Threading.SynchronizationContext]::Current
    try {
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
        if ($Credential) {
            Connect-VBRServer -Server $srv -Credential $Credential -ErrorAction Stop
        } else {
            Connect-VBRServer -Server $srv -ErrorAction Stop
        }
    } finally {
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($prevCtx)
    }
    Write-ConnectLog 'Connect-Veeam: Connect-VBRServer DONE'

    Write-ConnectLog 'Connect-Veeam: Get-VBRServerSession BEGIN'
    $session = Invoke-Safe { Get-VBRServerSession }
    Write-ConnectLog "Connect-Veeam: Get-VBRServerSession DONE (session present=$([bool]$session))"
    if (-not $session) { throw "Connected to '$srv' but no active Veeam server session was returned." }

    $script:Ctx.Server    = $srv
    $script:Ctx.Connected = $true
    $script:Cache         = @{}

    # Version / build. Get-VBRBackupServerInfo exists on newer builds; fall back to
    # the registry CorePath DLL's file version when it does not.
    Write-ConnectLog 'Connect-Veeam: reading server info / license'
    $info = Invoke-Safe { Get-VBRBackupServerInfo }
    if ($info) {
        $script:Ctx.ProductName  = Get-Prop $info 'Name'
        $script:Ctx.ProductBuild = Get-Prop $info 'Build'
        $script:Ctx.PatchLevel   = Get-Prop $info 'PatchLevel'
    }
    if (-not $script:Ctx.ProductBuild) {
        $ver = Invoke-Safe {
            $core = (Get-ItemProperty 'HKLM:\SOFTWARE\Veeam\Veeam Backup and Replication' -ErrorAction Stop).CorePath
            $dll  = Join-Path $core 'Veeam.Backup.Core.dll'
            (Get-Item $dll -ErrorAction Stop).VersionInfo.ProductVersion
        }
        if ($ver) { $script:Ctx.ProductBuild = $ver }
    }

    $lic = Get-CachedLicense
    if ($lic) {
        $script:Ctx.Edition   = Get-Prop $lic 'Edition'
        $script:Ctx.LicenseTo = Get-Prop $lic 'LicensedTo'
    }
    Write-ConnectLog "Connect-Veeam: OK (server='$($script:Ctx.Server)'; build='$($script:Ctx.ProductBuild)'; edition='$($script:Ctx.Edition)')"
}

function Disconnect-Veeam {
    Invoke-Safe { Disconnect-VBRServer -ErrorAction SilentlyContinue } | Out-Null
    $script:Ctx.Connected = $false
}

# ===========================================================================
# Result row helper
# ===========================================================================

function New-Row {
    param(
        [string]$Category, [string]$Check, [string]$Object = '',
        [ValidateSet('Pass','Warn','Fail','Error','Info')][string]$Status,
        [string]$Summary = '', [string]$Recommendation = '', [string]$Details = ''
    )
    [pscustomobject]@{
        Category       = $Category
        Check          = $Check
        Object         = $Object
        Status         = $Status
        Summary        = $Summary
        Recommendation = $Recommendation
        Details        = $Details
    }
}

# ===========================================================================
# Checks -- each returns an array of result rows
# ===========================================================================

function Check-Version {
    param($Def, $Options)
    $build = $script:Ctx.ProductBuild
    $patch = $script:Ctx.PatchLevel
    if (-not $build) {
        return New-Row -Category $Def.Category -Check $Def.Label -Object $script:Ctx.Server -Status 'Info' `
            -Summary 'Version could not be determined from this build.' `
            -Recommendation 'Confirm the installed version and patch level in the Veeam console (Help > About) and compare against the latest cumulative patch on veeam.com.'
    }
    $verText = if ($patch) { "$build (patch: $patch)" } else { "$build" }
    $summary = "VBR version $verText."
    $rec = 'Compare this build against the latest cumulative patch for your major version on the Veeam site. Running the newest patch is the single highest-value, lowest-risk maintenance item for a backup server.'
    return New-Row -Category $Def.Category -Check $Def.Label -Object $script:Ctx.Server -Status 'Info' `
        -Summary $summary -Recommendation $rec -Details "Product: $($script:Ctx.ProductName)`r`nBuild: $build`r`nPatch level: $patch"
}

function Check-License {
    param($Def, $Options)
    $lic = Get-CachedLicense
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $lic) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $script:Ctx.Server -Status 'Warn' `
            -Summary 'License information could not be read.' `
            -Recommendation 'Verify licensing in the Veeam console (main menu > License). An unreadable license can indicate an expired evaluation or a community edition.'))
        return $rows
    }

    $status  = Get-Prop $lic 'Status'          # Valid / Expired / Invalid / Warning
    $expiry  = Get-Prop $lic 'ExpirationDate'
    $support = Get-Prop $lic 'SupportExpirationDate'
    $edition = Get-Prop $lic 'Edition'
    $type    = Get-Prop $lic 'Type'
    $to      = Get-Prop $lic 'LicensedTo'

    # --- Status + expiry ---
    $warnDays = [int]$Options.LicenseWarnDays
    $st = 'Pass'; $sum = "License status: $status ($edition $type)."
    $rec = ''
    if ("$status" -match '(?i)expired|invalid') {
        $st = 'Fail'; $sum = "License status is '$status'."
        $rec = 'The license is expired or invalid. Backups may run in a restricted mode or stop. Apply a current license file immediately.'
    } elseif ($expiry -is [datetime]) {
        $days = [int]([Math]::Floor(($expiry - (Get-Date)).TotalDays))
        $sum = "License '$status', expires $($expiry.ToString('yyyy-MM-dd')) ($days day(s))."
        if ($days -lt 0)          { $st='Fail'; $rec='License has expired. Apply a renewed license file now.' }
        elseif ($days -le $warnDays) { $st='Warn'; $rec="License expires in $days day(s). Begin the renewal process and apply the new license before expiry to avoid job restrictions." }
    }
    $det = "Status: $status`r`nEdition: $edition`r`nType: $type`r`nLicensed to: $to`r`nExpiration: $expiry`r`nSupport expiration: $support"
    $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object 'Status / Expiry' -Status $st -Summary $sum -Recommendation $rec -Details $det))

    # --- Support contract expiry ---
    if ($support -is [datetime]) {
        $sdays = [int]([Math]::Floor(($support - (Get-Date)).TotalDays))
        $sst = 'Pass'; $srec = ''
        $ssum = "Support expires $($support.ToString('yyyy-MM-dd')) ($sdays day(s))."
        if ($sdays -lt 0)            { $sst='Warn'; $srec='Production support has lapsed. Without active support you cannot open Veeam support cases or (depending on edition) upgrade to new major versions. Renew maintenance.' }
        elseif ($sdays -le $warnDays) { $sst='Warn'; $srec="Support expires in $sdays day(s). Renew maintenance to keep upgrade rights and support access." }
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object 'Support Contract' -Status $sst -Summary $ssum -Recommendation $srec))
    }

    # --- Instance / socket usage ---
    $used  = Get-Prop $lic 'UsedInstancesNumber'
    $total = Get-Prop $lic 'LicensedInstancesNumber'
    if ($null -eq $used)  { $used  = Get-Prop $lic 'UsedInstances' }
    if ($null -eq $total) { $total = Get-Prop $lic 'LicensedInstances' }
    if (($null -ne $used) -and ($null -ne $total) -and ([double]$total -gt 0)) {
        $pct = [Math]::Round(([double]$used / [double]$total) * 100, 0)
        $ust = 'Pass'; $urec = ''
        $usum = "Instance usage: $used of $total ($pct%)."
        if ($pct -ge 100)   { $ust='Fail'; $urec='All licensed instances are consumed. New or protected workloads may fail to process. Reclaim unused instances or purchase additional capacity.' }
        elseif ($pct -ge 90) { $ust='Warn'; $urec="Instance usage is at $pct%. Plan capacity: review protected workloads and add instances before you hit the limit." }
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object 'Instance Usage' -Status $ust -Summary $usum -Recommendation $urec))
    }

    return $rows
}

function Check-JobResults {
    param($Def, $Options)
    $jobs = Get-CachedJobs
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $jobs -or $jobs.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Warn' `
            -Summary 'No backup jobs were found on this server.' `
            -Recommendation 'A VBR server with no jobs is either brand new or misconfigured. Confirm this is expected.'))
        return $rows
    }
    foreach ($job in $jobs) {
        if ($script:CancelRequested) { break }
        $name = Get-Prop $job 'Name'
        $enabled = Get-Prop $job 'IsScheduleEnabled'
        $last = Invoke-Safe { $job.GetLastResult() }   # Success / Warning / Failed / None
        $lastText = "$last"
        $st = 'Info'; $rec = ''
        switch -Regex ($lastText) {
            '(?i)success' { $st='Pass' }
            '(?i)warning' { $st='Warn'; $rec='Last run completed with warnings. Open the session details for this job to see which objects warned (often stale VM tools, skipped disks, or retry-on-first-attempt).' }
            '(?i)failed'  { $st='Fail'; $rec='Last run FAILED. Open the job session log, resolve the underlying error (credentials, snapshot, repository, network), and re-run. A failed job means no fresh restore point.' }
            default       { $st='Info'; $rec='Job has no completed sessions yet, or its last state is unknown. Verify it has run at least once.' }
        }
        $enabledText = if ($enabled -eq $false) { ' [schedule disabled]' } else { '' }
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status $st `
            -Summary "Last result: $lastText$enabledText" -Recommendation $rec))
    }
    return $rows
}

function Check-RpoCompliance {
    param($Def, $Options)
    $jobs = Get-CachedJobs
    $sessions = Get-CachedSessions
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $jobs -or $jobs.Count -eq 0) { return $rows }

    $rpoHours = [int]$Options.RpoHours
    $threshold = (Get-Date).AddHours(-$rpoHours)

    # Index the most recent SUCCESSFUL (or warning) session end-time per job name.
    $lastGood = @{}
    foreach ($s in $sessions) {
        $jn  = Get-Prop $s 'JobName'
        if (-not $jn) { continue }
        $res = "$(Get-Prop $s 'Result')"
        if ($res -match '(?i)success|warning') {
            $end = Get-Prop $s 'EndTime'
            if (-not ($end -is [datetime])) { $end = Get-Prop $s 'CreationTime' }
            if ($end -is [datetime]) {
                if (-not $lastGood.ContainsKey($jn) -or $end -gt $lastGood[$jn]) { $lastGood[$jn] = $end }
            }
        }
    }

    foreach ($job in $jobs) {
        if ($script:CancelRequested) { break }
        $name = Get-Prop $job 'Name'
        $enabled = Get-Prop $job 'IsScheduleEnabled'
        if ($enabled -eq $false) { continue }   # do not RPO-flag intentionally disabled jobs
        if ($lastGood.ContainsKey($name)) {
            $when = $lastGood[$name]
            $age  = [Math]::Round(((Get-Date) - $when).TotalHours, 1)
            if ($when -lt $threshold) {
                $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Fail' `
                    -Summary "Last good backup was $age h ago (RPO target $rpoHours h)." `
                    -Recommendation "No successful restore point within the $rpoHours-hour RPO. Investigate why the job is not completing successfully on schedule and restore protection for this workload." `
                    -Details "Last successful/warning session end: $when"))
            } else {
                $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Pass' `
                    -Summary "Last good backup $age h ago (within $rpoHours h)."))
            }
        } else {
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Warn' `
                -Summary 'No successful session found in history.' `
                -Recommendation 'This enabled job has no recorded successful run. Confirm it has ever completed and produced a restore point.'))
        }
    }
    return $rows
}

function Check-RecentSessions {
    param($Def, $Options)
    $sessions = Get-CachedSessions
    $rows = New-Object System.Collections.Generic.List[object]
    $hours = [int]$Options.SessionLookbackHours
    $since = (Get-Date).AddHours(-$hours)

    $recent = foreach ($s in $sessions) {
        $ct = Get-Prop $s 'CreationTime'
        if ($ct -is [datetime] -and $ct -ge $since) { $s }
    }
    $recent = @($recent)
    if ($recent.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Info' `
            -Summary "No sessions in the last $hours h."))
        return $rows
    }

    $failed  = @($recent | Where-Object { "$(Get-Prop $_ 'Result')" -match '(?i)failed' })
    $warned  = @($recent | Where-Object { "$(Get-Prop $_ 'Result')" -match '(?i)warning' })

    foreach ($s in $failed) {
        if ($script:CancelRequested) { break }
        $jn = Get-Prop $s 'JobName'; $ct = Get-Prop $s 'CreationTime'
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $jn -Status 'Fail' `
            -Summary "Failed session at $ct." `
            -Recommendation 'Open this session in History > Jobs, read the error, fix the root cause and re-run. Recurring failures on the same job point at a persistent config/infrastructure problem.'))
    }
    foreach ($s in $warned) {
        if ($script:CancelRequested) { break }
        $jn = Get-Prop $s 'JobName'; $ct = Get-Prop $s 'CreationTime'
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $jn -Status 'Warn' `
            -Summary "Warning session at $ct." `
            -Recommendation 'Review the warning detail. Persistent warnings degrade confidence in restores even though the job "succeeded".'))
    }
    if ($failed.Count -eq 0 -and $warned.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Pass' `
            -Summary "$($recent.Count) session(s) in last $hours h, all successful."))
    }
    return $rows
}

function Check-DisabledJobs {
    param($Def, $Options)
    $jobs = Get-CachedJobs
    $rows = New-Object System.Collections.Generic.List[object]
    $disabled = @($jobs | Where-Object { (Get-Prop $_ 'IsScheduleEnabled') -eq $false })
    if ($disabled.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Pass' -Summary 'No disabled jobs.'))
        return $rows
    }
    foreach ($job in $disabled) {
        $name = Get-Prop $job 'Name'
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Warn' `
            -Summary 'Job schedule is disabled.' `
            -Recommendation 'Confirm this job is disabled on purpose. A forgotten disabled job silently leaves a workload unprotected with no restore points being created.'))
    }
    return $rows
}

function Check-RunningJobs {
    param($Def, $Options)
    $sessions = Get-CachedSessions
    $rows = New-Object System.Collections.Generic.List[object]
    $running = @($sessions | Where-Object { "$(Get-Prop $_ 'State')" -match '(?i)working|running' })
    if ($running.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Info' -Summary 'No sessions currently running.'))
        return $rows
    }
    foreach ($s in $running) {
        $jn = Get-Prop $s 'JobName'; $ct = Get-Prop $s 'CreationTime'
        $hrs = if ($ct -is [datetime]) { [Math]::Round(((Get-Date) - $ct).TotalHours,1) } else { $null }
        $st = 'Info'; $rec = ''
        if ($null -ne $hrs -and $hrs -ge 12) {
            $st = 'Warn'
            $rec = 'This session has been running for an unusually long time. Check for a stuck task, an overloaded proxy/repository, or a network bottleneck.'
        }
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $jn -Status $st `
            -Summary "Running since $ct$(if ($null -ne $hrs) { " (${hrs}h)" })." -Recommendation $rec))
    }
    return $rows
}

function Check-RepoFreeSpace {
    param($Def, $Options)
    $repos = Get-CachedRepos
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $repos -or $repos.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Warn' -Summary 'No standard backup repositories found.' `
            -Recommendation 'Confirm repositories are configured (they may all be extents of a scale-out repository - see that check).'))
        return $rows
    }
    $warnPct = [double]$Options.RepoWarnPct
    $failPct = [double]$Options.RepoFailPct
    foreach ($repo in $repos) {
        if ($script:CancelRequested) { break }
        $name = Get-Prop $repo 'Name'
        $container = Invoke-Safe { $repo.GetContainer() }
        $free  = Get-Prop (Get-Prop $container 'CachedFreeSpace')  'InBytes'
        $total = Get-Prop (Get-Prop $container 'CachedTotalSpace') 'InBytes'
        if (($null -eq $free) -or ($null -eq $total) -or ([double]$total -le 0)) {
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Info' `
                -Summary 'Free space not reported (object storage / cloud / unavailable).' `
                -Recommendation 'Object-storage and some cloud repositories do not report free space here. Verify capacity in the provider console.'))
            continue
        }
        $usedPct = [Math]::Round(((([double]$total - [double]$free) / [double]$total) * 100), 0)
        $st = 'Pass'; $rec = ''
        $sum = "$(Format-Bytes $free) free of $(Format-Bytes $total) ($usedPct% used)."
        if ($usedPct -ge $failPct) {
            $st='Fail'; $rec="Repository is $usedPct% full (>= $failPct%). Backups will soon fail. Add capacity, offload/age out old restore points, or move workloads to another repository now."
        } elseif ($usedPct -ge $warnPct) {
            $st='Warn'; $rec="Repository is $usedPct% full (>= $warnPct%). Plan capacity: review retention and growth trend before it fills."
        }
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status $st -Summary $sum -Recommendation $rec))
    }
    return $rows
}

function Check-SobrHealth {
    param($Def, $Options)
    $sobrs = Get-CachedSobrs
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $sobrs -or $sobrs.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Info' -Summary 'No scale-out backup repositories configured.'))
        return $rows
    }
    $warnPct = [double]$Options.RepoWarnPct
    $failPct = [double]$Options.RepoFailPct
    foreach ($sobr in $sobrs) {
        if ($script:CancelRequested) { break }
        $sname = Get-Prop $sobr 'Name'
        $extents = @(Invoke-Safe { Get-VBRRepositoryExtent -Repository $sobr }) | Where-Object { $_ }
        if ($extents.Count -eq 0) {
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $sname -Status 'Warn' `
                -Summary 'Scale-out repository has no readable extents.' `
                -Recommendation 'A SOBR with no accessible extents cannot store backups. Check the underlying repositories and their status in the console.'))
            continue
        }
        foreach ($ext in $extents) {
            $extRepo = Get-Prop $ext 'Repository'
            $ename = Get-Prop $extRepo 'Name'
            if (-not $ename) { $ename = Get-Prop $ext 'Name' }
            $status = Get-Prop $ext 'Status'   # e.g. Normal / Maintenance / Evacuating
            $container = Invoke-Safe { $extRepo.GetContainer() }
            $free  = Get-Prop (Get-Prop $container 'CachedFreeSpace')  'InBytes'
            $total = Get-Prop (Get-Prop $container 'CachedTotalSpace') 'InBytes'

            $st = 'Pass'; $rec = ''; $sum = "Extent status: $status."
            if ("$status" -match '(?i)maintenance|evacuat|error|down|offline') {
                $st = 'Warn'
                $rec = "Extent is in '$status' state. Confirm this is intentional (e.g. planned maintenance/evacuation). Extents out of the Normal state reduce SOBR capacity and placement options."
            }
            if (($null -ne $free) -and ($null -ne $total) -and ([double]$total -gt 0)) {
                $usedPct = [Math]::Round(((([double]$total - [double]$free) / [double]$total) * 100), 0)
                $sum = "$sum  $(Format-Bytes $free) free of $(Format-Bytes $total) ($usedPct% used)."
                if ($usedPct -ge $failPct -and $st -ne 'Warn') { $st='Fail'; $rec="Extent is $usedPct% full. Add capacity or rebalance the SOBR." }
                elseif ($usedPct -ge $warnPct -and $st -eq 'Pass') { $st='Warn'; $rec="Extent is $usedPct% full. Watch capacity." }
            }
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object "$sname / $ename" -Status $st -Summary $sum -Recommendation $rec))
        }
    }
    return $rows
}

function Check-Immutability {
    param($Def, $Options)
    $repos = @(Get-CachedRepos) + @(Invoke-Safe { Get-VBRRepositoryExtent -Repository (Get-CachedSobrs) } | ForEach-Object { Get-Prop $_ 'Repository' })
    $repos = @($repos | Where-Object { $_ })
    $rows = New-Object System.Collections.Generic.List[object]
    if ($repos.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Info' -Summary 'No repositories to evaluate for immutability.'))
        return $rows
    }
    $immutableFound = $false
    foreach ($repo in $repos) {
        $name = Get-Prop $repo 'Name'
        # Immutability is exposed differently across builds; probe several shapes.
        $days = Get-Prop $repo 'ImmutabilityPeriod'
        if ($null -eq $days) { $days = Get-Prop (Invoke-Safe { $repo.GetImmutabilitySettings() }) 'ImmutabilityPeriod' }
        $enabledFlag = Get-Prop $repo 'IsImmutabilitySupported'
        $typeText = "$(Get-Prop $repo 'Type')"
        $isHardened = $typeText -match '(?i)hardened|linuxlocal'
        $hasImmut = ($null -ne $days -and [int64]$days -gt 0)
        if ($hasImmut) {
            $immutableFound = $true
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Pass' `
                -Summary "Immutability enabled ($days day(s))."))
        } else {
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Info' `
                -Summary "No immutability period reported$(if ($isHardened) { ' (hardened repo type)' } else { '' })." `
                -Recommendation 'This repository does not report an immutability window. For ransomware resilience, at least one copy of backups should be immutable (hardened Linux repo, object storage with lock, or tape).'))
        }
    }
    if (-not $immutableFound) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object '(overall)' -Status 'Warn' `
            -Summary 'No immutable repository detected on this server.' `
            -Recommendation 'None of the repositories report immutability. This is a major ransomware-resilience gap - implement a hardened Linux repository, object storage with object-lock, or an air-gapped/tape copy so backups cannot be deleted or encrypted by an attacker.'))
    }
    return $rows
}

function Check-Services {
    param($Def, $Options)
    $rows = New-Object System.Collections.Generic.List[object]
    $target = $script:Ctx.Server
    $isLocal = ($target -in @('localhost','127.0.0.1',$env:COMPUTERNAME,"$env:COMPUTERNAME.$env:USERDNSDOMAIN"))
    $services = Invoke-Safe {
        if ($isLocal) { Get-Service -Name 'Veeam*' -ErrorAction Stop }
        else { Get-Service -ComputerName $target -Name 'Veeam*' -ErrorAction Stop }
    }
    if (-not $services) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $target -Status 'Info' `
            -Summary 'Could not enumerate Veeam services on this host.' `
            -Recommendation "If the VBR server is remote, service enumeration needs RPC/remote-registry access and rights. Run the tool on the VBR server itself, or verify Veeam services in services.msc."))
        return $rows
    }
    $services = @($services)
    $stopped = @($services | Where-Object { "$($_.Status)" -ne 'Running' })
    # Some Veeam services are stopped by design (e.g. certain broker/mount helpers start on demand),
    # so flag stopped-but-Automatic as the meaningful signal where StartType is available.
    foreach ($svc in $stopped) {
        $startType = Invoke-Safe { $svc.StartType }
        $st = 'Warn'; $rec = 'This Veeam service is not running. If its start type is Automatic, start it and investigate why it stopped - a stopped core service can halt all backups.'
        if ("$startType" -match '(?i)manual|disabled') {
            $st = 'Info'; $rec = 'Service is stopped but set to Manual/Disabled - often normal for on-demand helper services. Confirm it is not a core service.'
        }
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $svc.Name -Status $st `
            -Summary "$($svc.DisplayName): $($svc.Status) (start: $startType)." -Recommendation $rec))
    }
    if ($stopped.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $target -Status 'Pass' `
            -Summary "All $($services.Count) Veeam service(s) running."))
    }
    return $rows
}

function Check-Proxies {
    param($Def, $Options)
    $proxies = Get-CachedProxies
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $proxies -or $proxies.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Info' `
            -Summary 'No VMware/Hyper-V backup proxies found.' `
            -Recommendation 'If this environment backs up VMs, at least one proxy is expected. Agent-only or direct-storage setups legitimately have none.'))
        return $rows
    }
    foreach ($px in $proxies) {
        $name = Get-Prop $px 'Name'
        $disabled = Get-Prop $px 'IsDisabled'
        # NB: do not name this $host -- that is a PowerShell automatic variable
        # (the host UI object); clobbering it can break the console session.
        $proxyHost = Get-Prop $px 'Host'
        $hostName = Get-Prop $proxyHost 'Name'
        if ($disabled -eq $true) {
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Warn' `
                -Summary 'Proxy is disabled.' `
                -Recommendation 'Confirm this proxy is disabled intentionally. Disabled proxies reduce backup concurrency and can bottleneck the backup window.'))
        } else {
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Pass' `
                -Summary "Enabled$(if ($hostName) { " (host: $hostName)" })."))
        }
    }
    return $rows
}

function Check-ConfigBackup {
    param($Def, $Options)
    $cfg = Get-CachedConfigJob
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $cfg) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Fail' `
            -Summary 'Configuration backup job could not be read.' `
            -Recommendation 'Configuration backup protects the VBR database itself (jobs, credentials, encryption keys). Verify it exists and is enabled under main menu > Configuration Backup.'))
        return $rows
    }

    # --- Enabled + schedule ---
    $enabled = Get-Prop $cfg 'Enabled'
    if ($enabled -eq $false) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object 'Enabled' -Status 'Fail' `
            -Summary 'Configuration backup is DISABLED.' `
            -Recommendation 'Enable scheduled configuration backup. Without it, a lost VBR server means rebuilding all jobs, credentials and (critically) encryption keys from scratch - potentially orphaning encrypted backups.'))
    } else {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object 'Enabled' -Status 'Pass' -Summary 'Configuration backup is enabled and scheduled.'))
    }

    # --- Encryption ---
    $enc = Get-Prop (Get-Prop $cfg 'EncryptionOptions') 'Enabled'
    if ($enc -eq $true) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object 'Encryption' -Status 'Pass' -Summary 'Configuration backup is encrypted.'))
    } else {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object 'Encryption' -Status 'Warn' `
            -Summary 'Configuration backup is NOT encrypted.' `
            -Recommendation 'Enable encryption on the configuration backup. It contains stored credentials and backup encryption keys; an unencrypted copy is a serious exposure if the file is stolen.'))
    }

    # --- Last result + age ---
    $lastResult = Get-Prop $cfg 'LastResult'
    $lastTime   = Get-Prop $cfg 'LastRun'
    if ($null -eq $lastTime) { $lastTime = Get-Prop $cfg 'LatestFinishTime' }
    $st = 'Info'; $rec = ''; $sum = "Last result: $lastResult."
    if ("$lastResult" -match '(?i)failed') { $st='Fail'; $rec='The most recent configuration backup failed. Fix it now - you may currently have no usable VBR configuration backup.' }
    elseif ("$lastResult" -match '(?i)warning') { $st='Warn'; $rec='Last configuration backup warned. Review the session.' }
    elseif ("$lastResult" -match '(?i)success') { $st='Pass' }
    if ($lastTime -is [datetime]) {
        $ageDays = [Math]::Round(((Get-Date) - $lastTime).TotalDays,1)
        $sum = "$sum Last run $($lastTime.ToString('yyyy-MM-dd HH:mm')) (${ageDays}d ago)."
        if ($ageDays -gt 7 -and $st -ne 'Fail') { $st='Warn'; $rec='Configuration backup has not run in over a week. Confirm the schedule is active and the target is reachable.' }
    }
    $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object 'Last Run' -Status $st -Summary $sum -Recommendation $rec))

    # --- Target ---
    $target = Get-Prop $cfg 'Target'
    if (-not $target) { $target = Get-Prop (Get-Prop $cfg 'RepositoryObject') 'Name' }
    if ($target) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object 'Target' -Status 'Info' `
            -Summary "Target repository: $target." `
            -Recommendation 'Best practice: store the configuration backup on a repository SEPARATE from your primary production backups (ideally off the VBR server), so a disaster that takes out production backups does not also take out the config backup needed to recover.'))
    }

    return $rows
}

function Check-BackupCopy {
    param($Def, $Options)
    $jobs = Get-CachedJobs
    $rows = New-Object System.Collections.Generic.List[object]
    $copyJobs = @($jobs | Where-Object {
        ("$(Get-Prop $_ 'JobType')" -match '(?i)copy|sync') -or ((Get-Prop $_ 'IsBackupCopy') -eq $true)
    })
    if ($copyJobs.Count -gt 0) {
        # Compute the detail string first: putting "($copyJobs | ...) -join" inline as
        # a parameter value makes the parser read -join as a New-Row parameter name.
        $copyNames = ($copyJobs | ForEach-Object { Get-Prop $_ 'Name' }) -join "`r`n"
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Pass' `
            -Summary "$($copyJobs.Count) backup copy job(s) present." `
            -Details $copyNames))
    } else {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Warn' `
            -Summary 'No backup copy jobs found.' `
            -Recommendation 'The 3-2-1 rule wants a second copy on different media, ideally offsite. With no backup copy job, all restore points may live in one place. Add a backup copy job to a secondary/offsite/object-storage repository.'))
    }
    return $rows
}

function Check-JobEncryption {
    param($Def, $Options)
    $jobs = Get-CachedJobs
    $rows = New-Object System.Collections.Generic.List[object]
    $backupJobs = @($jobs | Where-Object { "$(Get-Prop $_ 'JobType')" -match '(?i)backup' })
    if ($backupJobs.Count -eq 0) {
        $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Status 'Info' -Summary 'No backup jobs to evaluate for encryption.'))
        return $rows
    }
    foreach ($job in $backupJobs) {
        $name = Get-Prop $job 'Name'
        $opts = Invoke-Safe { $job.GetOptions() }
        $encEnabled = Get-Prop (Get-Prop $opts 'BackupStorageOptions') 'StorageEncryptionEnabled'
        if ($null -eq $encEnabled) {
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Info' `
                -Summary 'Encryption state could not be read for this job.'))
        } elseif ($encEnabled -eq $true) {
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Pass' -Summary 'At-rest encryption enabled.'))
        } else {
            $rows.Add((New-Row -Category $Def.Category -Check $Def.Label -Object $name -Status 'Info' `
                -Summary 'At-rest encryption not enabled.' `
                -Recommendation 'Consider enabling job encryption where backups leave your trusted premises (offsite copy, tape, object storage) or where the repository is not otherwise encrypted. Store the password safely - lost passwords mean unrecoverable backups.'))
        }
    }
    return $rows
}

function Invoke-VeeamCheck {
    param($Def, $Options)
    try {
        switch ($Def.Action) {
            'Version'        { return @(Check-Version        -Def $Def -Options $Options) }
            'License'        { return @(Check-License        -Def $Def -Options $Options) }
            'JobResults'     { return @(Check-JobResults     -Def $Def -Options $Options) }
            'RpoCompliance'  { return @(Check-RpoCompliance  -Def $Def -Options $Options) }
            'RecentSessions' { return @(Check-RecentSessions -Def $Def -Options $Options) }
            'DisabledJobs'   { return @(Check-DisabledJobs   -Def $Def -Options $Options) }
            'RunningJobs'    { return @(Check-RunningJobs    -Def $Def -Options $Options) }
            'RepoFreeSpace'  { return @(Check-RepoFreeSpace  -Def $Def -Options $Options) }
            'SobrHealth'     { return @(Check-SobrHealth     -Def $Def -Options $Options) }
            'Immutability'   { return @(Check-Immutability   -Def $Def -Options $Options) }
            'Services'       { return @(Check-Services       -Def $Def -Options $Options) }
            'Proxies'        { return @(Check-Proxies        -Def $Def -Options $Options) }
            'ConfigBackup'   { return @(Check-ConfigBackup   -Def $Def -Options $Options) }
            'BackupCopy'     { return @(Check-BackupCopy     -Def $Def -Options $Options) }
            'JobEncryption'  { return @(Check-JobEncryption  -Def $Def -Options $Options) }
            default          { throw "Unknown check action '$($Def.Action)'." }
        }
    } catch {
        return @(New-Row -Category $Def.Category -Check $Def.Label -Status 'Error' `
            -Summary $_.Exception.Message -Details $_.Exception.ToString())
    }
}

# ===========================================================================
# Logging + report export
# ===========================================================================

function Write-DiagLog {
    param([string]$Path, $Result)
    $sb = New-Object System.Text.StringBuilder
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine("**RESULT--$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $null = $sb.AppendLine("Category=$($Result.Category)")
    $null = $sb.AppendLine("Check=$($Result.Check)")
    $null = $sb.AppendLine("Object=$($Result.Object)")
    $null = $sb.AppendLine("Status=$($Result.Status)")
    $null = $sb.AppendLine("Summary=$($Result.Summary)")
    $null = $sb.AppendLine("Recommendation=$($Result.Recommendation)")
    try { Add-Content -LiteralPath $Path -Value $sb.ToString() -Encoding UTF8 } catch { }
}

function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return ($Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;')
}

function Export-HtmlReport {
    param([string]$Path)

    $server    = $script:Ctx.Server
    $version   = if ($script:Ctx.ProductBuild) { $script:Ctx.ProductBuild } else { 'Unknown' }
    $edition   = $script:Ctx.Edition
    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $total = $script:Results.Count
    $pass  = @($script:Results | Where-Object { $_.Status -eq 'Pass' }).Count
    $warn  = @($script:Results | Where-Object { $_.Status -eq 'Warn' }).Count
    $fail  = @($script:Results | Where-Object { $_.Status -eq 'Fail' }).Count
    $err   = @($script:Results | Where-Object { $_.Status -eq 'Error' }).Count

    $css = @'
body{font-family:Segoe UI,Arial,sans-serif;margin:0;padding:0;background:#f4f5f7;color:#222}
header{background:#0a2f5c;color:#fff;padding:20px 30px}
header h1{margin:0;font-size:22px}
header p{margin:4px 0 0;font-size:13px;color:#cfe0f5}
.wrap{padding:20px 30px}
.summary{display:flex;gap:14px;margin-bottom:24px;flex-wrap:wrap}
.card{background:#fff;border-radius:6px;box-shadow:0 1px 3px rgba(0,0,0,.15);padding:14px 20px;min-width:110px;text-align:center}
.card .n{font-size:26px;font-weight:700}
.card .l{font-size:12px;color:#666;text-transform:uppercase}
.c-pass{color:#0a8a3f}.c-warn{color:#c07a00}.c-fail{color:#c0392b}.c-err{color:#c0392b}
h2{border-bottom:2px solid #0a2f5c;padding-bottom:4px;margin-top:32px;font-size:18px;color:#0a2f5c}
table{width:100%;border-collapse:collapse;background:#fff;margin-top:8px}
th,td{text-align:left;padding:8px 10px;border-bottom:1px solid #e2e2e2;font-size:13px;vertical-align:top}
th{background:#eef2f7;font-size:12px;text-transform:uppercase;color:#444}
.badge{display:inline-block;padding:2px 10px;border-radius:10px;font-size:11px;font-weight:600;color:#fff;white-space:nowrap}
.b-pass{background:#0a8a3f}.b-warn{background:#c07a00}.b-fail{background:#c0392b}.b-err{background:#c0392b}.b-info{background:#777}
.rec{color:#8a3f00;font-size:12px}
.action{background:#fff8ec;border-left:4px solid #c07a00;padding:10px 14px;margin:6px 0;font-size:13px}
.action .t{font-weight:600;color:#8a3f00}
details summary{cursor:pointer;color:#0a2f5c;font-size:12px;margin-top:4px}
pre{white-space:pre-wrap;word-break:break-word;background:#f8f9fb;border:1px solid #e2e2e2;padding:8px;margin:6px 0 0;font-size:12px;max-height:320px;overflow:auto}
footer{padding:20px 30px;color:#888;font-size:12px}
'@

    $sb = New-Object System.Text.StringBuilder
    $null = $sb.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8">')
    $null = $sb.AppendLine("<title>Veeam Health Report - $(ConvertTo-HtmlSafe $server)</title>")
    $null = $sb.AppendLine("<style>$css</style></head><body>")
    $null = $sb.AppendLine('<header><h1>Veeam Backup &amp; Replication Health Report</h1>')
    $null = $sb.AppendLine("<p>Server: $(ConvertTo-HtmlSafe $server) &nbsp;|&nbsp; Version: $(ConvertTo-HtmlSafe $version) &nbsp;|&nbsp; Edition: $(ConvertTo-HtmlSafe $edition) &nbsp;|&nbsp; Generated: $generated</p></header>")
    $null = $sb.AppendLine('<div class="wrap">')
    $null = $sb.AppendLine('<div class="summary">')
    $null = $sb.AppendLine("<div class='card'><div class='n'>$total</div><div class='l'>Checks</div></div>")
    $null = $sb.AppendLine("<div class='card'><div class='n c-pass'>$pass</div><div class='l'>Pass</div></div>")
    $null = $sb.AppendLine("<div class='card'><div class='n c-warn'>$warn</div><div class='l'>Warn</div></div>")
    $null = $sb.AppendLine("<div class='card'><div class='n c-fail'>$fail</div><div class='l'>Fail</div></div>")
    $null = $sb.AppendLine("<div class='card'><div class='n c-err'>$err</div><div class='l'>Error</div></div>")
    $null = $sb.AppendLine('</div>')

    # --- Priority action list: every Fail/Warn that carries a recommendation ---
    $actions = @($script:Results | Where-Object { ($_.Status -in @('Fail','Warn')) -and $_.Recommendation })
    if ($actions.Count -gt 0) {
        $null = $sb.AppendLine('<h2>Recommended Actions</h2>')
        foreach ($a in ($actions | Sort-Object @{E={ if ($_.Status -eq 'Fail') {0} else {1} }})) {
            $tag = "$($a.Category) / $($a.Check)$(if ($a.Object) { " - $($a.Object)" })"
            $null = $sb.AppendLine("<div class='action'><div class='t'>[$($a.Status)] $(ConvertTo-HtmlSafe $tag)</div>$(ConvertTo-HtmlSafe $a.Recommendation)</div>")
        }
    }

    foreach ($cat in ($script:Results | Group-Object Category)) {
        $null = $sb.AppendLine("<h2>$(ConvertTo-HtmlSafe $cat.Name)</h2>")
        $null = $sb.AppendLine('<table><tr><th>Check</th><th>Object</th><th>Status</th><th>Summary</th><th>Recommendation</th></tr>')
        foreach ($r in $cat.Group) {
            $badgeClass = switch ($r.Status) { 'Pass' {'b-pass'} 'Warn' {'b-warn'} 'Fail' {'b-fail'} 'Error' {'b-err'} default {'b-info'} }
            $null = $sb.AppendLine('<tr>')
            $null = $sb.AppendLine("<td>$(ConvertTo-HtmlSafe $r.Check)</td>")
            $null = $sb.AppendLine("<td>$(ConvertTo-HtmlSafe $r.Object)</td>")
            $null = $sb.AppendLine("<td><span class='badge $badgeClass'>$($r.Status)</span></td>")
            $detBlock = if ($r.Details) { "<details><summary>Detail</summary><pre>$(ConvertTo-HtmlSafe $r.Details)</pre></details>" } else { '' }
            $null = $sb.AppendLine("<td>$(ConvertTo-HtmlSafe $r.Summary)$detBlock</td>")
            $null = $sb.AppendLine("<td class='rec'>$(ConvertTo-HtmlSafe $r.Recommendation)</td>")
            $null = $sb.AppendLine('</tr>')
        }
        $null = $sb.AppendLine('</table>')
    }

    $null = $sb.AppendLine('</div>')
    $null = $sb.AppendLine('<footer>Generated by Veeam Diagnostic Toolkit &mdash; m365admintools.com &mdash; Status is a heuristic based on Veeam-reported state and threshold comparisons. Confirm findings in the Veeam console before making changes.</footer>')
    $null = $sb.AppendLine('</body></html>')

    Set-Content -LiteralPath $Path -Value $sb.ToString() -Encoding UTF8
}

# ===========================================================================
# UI
# ===========================================================================

$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Veeam Diagnostic Toolkit 1.0'
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
$form.MinimumSize   = New-Object System.Drawing.Size(620, 340)
$form.FormBorderStyle = 'Sizable'
$form.MaximizeBox   = $true
$form.KeyPreview    = $true
$form.AutoScaleMode = 'None'

$fTitle = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
$fBold  = New-Object System.Drawing.Font('Segoe UI', 9,  [System.Drawing.FontStyle]::Bold)
$fSmall = New-Object System.Drawing.Font('Segoe UI', 8)
$fMono  = New-Object System.Drawing.Font('Consolas', 9)
$cBlue   = [System.Drawing.Color]::FromArgb(0,0,190)
$cRed    = [System.Drawing.Color]::FromArgb(190,0,0)
$cGreen  = [System.Drawing.Color]::FromArgb(0,120,0)
$cOrange = [System.Drawing.Color]::FromArgb(190,120,0)
$cGray   = [System.Drawing.Color]::FromArgb(110,110,110)

function New-Lbl {
    param($Text,$X,$Y,$W=120,$H=18,$Font=$null,$Color=$null)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X,$Y)
    $l.Size = New-Object System.Drawing.Size($W,$H)
    if ($Font)  { $l.Font = $Font }
    if ($Color) { $l.ForeColor = $Color }
    return $l
}

$tbl = New-Object System.Windows.Forms.TableLayoutPanel
$tbl.Dock = 'Fill'
$tbl.ColumnCount = 1
$tbl.RowCount = 2
$null = $tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 118)))
$form.Controls.Add($tbl)

$pnlMain = New-Object System.Windows.Forms.Panel
$pnlMain.Dock = 'Fill'
$pnlMain.AutoScroll = $true
$tbl.Controls.Add($pnlMain, 0, 0)

$pnlBottom = New-Object System.Windows.Forms.Panel
$pnlBottom.Dock = 'Fill'
$tbl.Controls.Add($pnlBottom, 0, 1)

# --- Header ---
# Brand logo (embedded). Falls back to the original text label if the image
# cannot be decoded, so the header is never left empty.
try {
    $picLogo = New-Object System.Windows.Forms.PictureBox
    $picLogo.Location  = New-Object System.Drawing.Point(12,7)
    $picLogo.Size      = New-Object System.Drawing.Size(219,28)   # native size = crisp
    $picLogo.SizeMode  = 'Zoom'
    $picLogo.BackColor = [System.Drawing.Color]::Transparent
    $picLogo.Image     = Get-LogoImage
    $pnlMain.Controls.Add($picLogo)
} catch {
    $pnlMain.Controls.Add((New-Lbl 'M365 Admin Tools' 12 6 250 24 $fTitle $cBlue))
}
$pnlMain.Controls.Add((New-Lbl 'Veeam Backup and Replication Diagnostics' 520 6 428 24 $fTitle))
$lnkSite = New-Object System.Windows.Forms.LinkLabel
$lnkSite.Text             = 'Charles Arconi  |  m365admintools.com'
$lnkSite.Location         = New-Object System.Drawing.Point(520,32)
$lnkSite.Size             = New-Object System.Drawing.Size(428,16)
$lnkSite.Font             = $fSmall
$lnkSite.ForeColor        = $cGray
$lnkSite.LinkColor        = $cBlue
$lnkSite.ActiveLinkColor  = $cBlue
$lnkSite.VisitedLinkColor = $cBlue
$lnkSite.LinkBehavior     = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$lnkSite.AutoSize         = $false

$siteLinkText = 'm365admintools.com'
$lnkSite.LinkArea = New-Object System.Windows.Forms.LinkArea(
    $lnkSite.Text.IndexOf($siteLinkText), $siteLinkText.Length)

$lnkSite.Add_LinkClicked({
    try {
        $lnkSite.LinkVisited = $true
        Start-Process 'https://m365admintools.com'
    } catch {
        Write-ErrorLog -Context 'Open m365admintools.com link' -ErrorRecord $_
        Set-Status "Could not open browser: $($_.Exception.Message)"
    }
})
$pnlMain.Controls.Add($lnkSite)

# --- 1. Target ---
$grpTarget = New-Object System.Windows.Forms.GroupBox
$grpTarget.Text = ' 1. Target '
$grpTarget.Location = New-Object System.Drawing.Point(12,52)
$grpTarget.Size = New-Object System.Drawing.Size(936,80)
$pnlMain.Controls.Add($grpTarget)

$grpTarget.Controls.Add((New-Lbl 'Veeam server' 12 25 78))
$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Location = New-Object System.Drawing.Point(94,22)
$txtServer.Size = New-Object System.Drawing.Size(230,22)
$txtServer.Text = $script:Ctx.Server
$grpTarget.Controls.Add($txtServer)

$btnCreds = New-Object System.Windows.Forms.Button
$btnCreds.Text = 'Credentials...'
$btnCreds.Location = New-Object System.Drawing.Point(332,21)
$btnCreds.Size = New-Object System.Drawing.Size(96,24)
$grpTarget.Controls.Add($btnCreds)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = 'Connect'
$btnConnect.Location = New-Object System.Drawing.Point(432,21)
$btnConnect.Size = New-Object System.Drawing.Size(96,24)
$btnConnect.Font = $fBold
$grpTarget.Controls.Add($btnConnect)

$lblConnInfo = New-Lbl 'Not connected.' 540 26 388 16 $fSmall $cGray
$grpTarget.Controls.Add($lblConnInfo)

$lblConnInfo2 = New-Lbl '' 12 55 916 16 $fSmall $cGray
$grpTarget.Controls.Add($lblConnInfo2)

# --- 2. Diagnostic Checks ---
$grpChecks = New-Object System.Windows.Forms.GroupBox
$grpChecks.Text = ' 2. Diagnostic Checks '
$grpChecks.Location = New-Object System.Drawing.Point(12,140)
$grpChecks.Size = New-Object System.Drawing.Size(936,300)
$pnlMain.Controls.Add($grpChecks)

$treeChecks = New-Object System.Windows.Forms.TreeView
$treeChecks.Location = New-Object System.Drawing.Point(10,20)
$treeChecks.Size = New-Object System.Drawing.Size(578,266)
$treeChecks.CheckBoxes = $true
$grpChecks.Controls.Add($treeChecks)

$btnPresetQuick = New-Object System.Windows.Forms.Button
$btnPresetQuick.Text = 'Quick Health Check'
$btnPresetQuick.Location = New-Object System.Drawing.Point(602,20)
$btnPresetQuick.Size = New-Object System.Drawing.Size(324,26)
$grpChecks.Controls.Add($btnPresetQuick)

$btnPresetFull = New-Object System.Windows.Forms.Button
$btnPresetFull.Text = 'Select All'
$btnPresetFull.Location = New-Object System.Drawing.Point(602,50)
$btnPresetFull.Size = New-Object System.Drawing.Size(158,26)
$grpChecks.Controls.Add($btnPresetFull)

$btnPresetNone = New-Object System.Windows.Forms.Button
$btnPresetNone.Text = 'Select None'
$btnPresetNone.Location = New-Object System.Drawing.Point(768,50)
$btnPresetNone.Size = New-Object System.Drawing.Size(158,26)
$grpChecks.Controls.Add($btnPresetNone)

$grpRunOpts = New-Object System.Windows.Forms.GroupBox
$grpRunOpts.Text = ' Thresholds '
$grpRunOpts.Location = New-Object System.Drawing.Point(602,88)
$grpRunOpts.Size = New-Object System.Drawing.Size(324,198)
$grpChecks.Controls.Add($grpRunOpts)

$grpRunOpts.Controls.Add((New-Lbl 'RPO target (hours)' 12 26 160 18))
$numRpo = New-Object System.Windows.Forms.NumericUpDown
$numRpo.Location = New-Object System.Drawing.Point(220,24)
$numRpo.Size = New-Object System.Drawing.Size(90,22)
$numRpo.Minimum = 1; $numRpo.Maximum = 720; $numRpo.Value = 24
$grpRunOpts.Controls.Add($numRpo)

$grpRunOpts.Controls.Add((New-Lbl 'Session lookback (hours)' 12 54 200 18))
$numLookback = New-Object System.Windows.Forms.NumericUpDown
$numLookback.Location = New-Object System.Drawing.Point(220,52)
$numLookback.Size = New-Object System.Drawing.Size(90,22)
$numLookback.Minimum = 1; $numLookback.Maximum = 720; $numLookback.Value = 24
$grpRunOpts.Controls.Add($numLookback)

$grpRunOpts.Controls.Add((New-Lbl 'Repo free-space Warn at % used' 12 82 200 18))
$numWarnPct = New-Object System.Windows.Forms.NumericUpDown
$numWarnPct.Location = New-Object System.Drawing.Point(220,80)
$numWarnPct.Size = New-Object System.Drawing.Size(90,22)
$numWarnPct.Minimum = 50; $numWarnPct.Maximum = 99; $numWarnPct.Value = 85
$grpRunOpts.Controls.Add($numWarnPct)

$grpRunOpts.Controls.Add((New-Lbl 'Repo free-space Fail at % used' 12 110 200 18))
$numFailPct = New-Object System.Windows.Forms.NumericUpDown
$numFailPct.Location = New-Object System.Drawing.Point(220,108)
$numFailPct.Size = New-Object System.Drawing.Size(90,22)
$numFailPct.Minimum = 51; $numFailPct.Maximum = 100; $numFailPct.Value = 95
$grpRunOpts.Controls.Add($numFailPct)

$grpRunOpts.Controls.Add((New-Lbl 'License expiry warning (days)' 12 138 200 18))
$numLicDays = New-Object System.Windows.Forms.NumericUpDown
$numLicDays.Location = New-Object System.Drawing.Point(220,136)
$numLicDays.Size = New-Object System.Drawing.Size(90,22)
$numLicDays.Minimum = 1; $numLicDays.Maximum = 365; $numLicDays.Value = 30
$grpRunOpts.Controls.Add($numLicDays)

$grpRunOpts.Controls.Add((New-Lbl 'Status is a heuristic. Confirm in the Veeam console before changing anything.' 12 166 300 28 $fSmall $cGray))

# --- 3. Results ---
$grpResults = New-Object System.Windows.Forms.GroupBox
$grpResults.Text = ' 3. Results '
$grpResults.Location = New-Object System.Drawing.Point(12,448)
$grpResults.Size = New-Object System.Drawing.Size(936,220)
$pnlMain.Controls.Add($grpResults)

$lvResults = New-Object System.Windows.Forms.ListView
$lvResults.Location = New-Object System.Drawing.Point(10,20)
$lvResults.Size = New-Object System.Drawing.Size(916,164)
$lvResults.View = 'Details'
$lvResults.FullRowSelect = $true
$lvResults.GridLines = $true
$null = $lvResults.Columns.Add('Category', 110)
$null = $lvResults.Columns.Add('Check',    150)
$null = $lvResults.Columns.Add('Object',   120)
$null = $lvResults.Columns.Add('Status',    60)
$null = $lvResults.Columns.Add('Summary',  250)
$null = $lvResults.Columns.Add('Recommendation', 210)
$grpResults.Controls.Add($lvResults)

$lblSummary = New-Lbl '' 10 190 916 20 $fBold
$grpResults.Controls.Add($lblSummary)

# --- 4. Details ---
$grpDetail = New-Object System.Windows.Forms.GroupBox
$grpDetail.Text = ' 4. Details (select a result row above) '
$grpDetail.Location = New-Object System.Drawing.Point(12,676)
$grpDetail.Size = New-Object System.Drawing.Size(936,160)
$pnlMain.Controls.Add($grpDetail)

$txtDetail = New-Object System.Windows.Forms.TextBox
$txtDetail.Location = New-Object System.Drawing.Point(10,20)
$txtDetail.Size = New-Object System.Drawing.Size(916,130)
$txtDetail.Multiline = $true
$txtDetail.ReadOnly = $true
$txtDetail.ScrollBars = 'Vertical'
$txtDetail.Font = $fMono
$grpDetail.Controls.Add($txtDetail)

# --- Action bar ---
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(12,4)
$progress.Size = New-Object System.Drawing.Size(936,12)
$progress.Anchor = 'Top,Left,Right'
$pnlBottom.Controls.Add($progress)

$pnlBottom.Controls.Add((New-Lbl 'Log' 12 24 26 16 $fSmall))
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(42,21)
$txtLog.Size = New-Object System.Drawing.Size(380,22)
$txtLog.Anchor = 'Top,Left'
$txtLog.Text = Join-Path $env:USERPROFILE 'Documents\VeeamDiagnosticTool.log'
$pnlBottom.Controls.Add($txtLog)

$btnOpenLog = New-Object System.Windows.Forms.Button
$btnOpenLog.Text = 'Open'
$btnOpenLog.Location = New-Object System.Drawing.Point(428,20)
$btnOpenLog.Size = New-Object System.Drawing.Size(58,24)
$btnOpenLog.Anchor = 'Top,Left'
$pnlBottom.Controls.Add($btnOpenLog)

$lblStatus = New-Lbl 'Ready. Connect to a Veeam server to begin.' 12 48 700 16 $fSmall
$lblStatus.Anchor = 'Top,Left,Right'
$pnlBottom.Controls.Add($lblStatus)

$flowBtn = New-Object System.Windows.Forms.FlowLayoutPanel
$flowBtn.Dock = 'Bottom'
$flowBtn.Height = 56
$flowBtn.FlowDirection = 'RightToLeft'
$flowBtn.WrapContents = $true
$flowBtn.Padding = New-Object System.Windows.Forms.Padding(8,6,10,6)
$pnlBottom.Controls.Add($flowBtn)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Size = New-Object System.Drawing.Size(86,28)
$flowBtn.Controls.Add($btnClose)

$btnExportHtml = New-Object System.Windows.Forms.Button
$btnExportHtml.Text = 'Export HTML Report  (F8)'
$btnExportHtml.Size = New-Object System.Drawing.Size(168,28)
$flowBtn.Controls.Add($btnExportHtml)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = 'Export CSV  (F7)'
$btnExportCsv.Size = New-Object System.Drawing.Size(120,28)
$flowBtn.Controls.Add($btnExportCsv)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Cancel  (F6)'
$btnCancel.Size = New-Object System.Drawing.Size(104,28)
$btnCancel.ForeColor = $cRed
$btnCancel.Enabled = $false
$flowBtn.Controls.Add($btnCancel)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'RUN CHECKS  (F5)'
$btnRun.Size = New-Object System.Drawing.Size(150,28)
$btnRun.Font = $fBold
$btnRun.ForeColor = $cBlue
$btnRun.Enabled = $false
$flowBtn.Controls.Add($btnRun)

# --- Size the window to fit the screen ---
$designW = 1000
$designH = 880
try {
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $w = [Math]::Min($designW, [Math]::Max(620, $wa.Width  - 40))
    $h = [Math]::Min($designH, [Math]::Max(340, $wa.Height - 40))
} catch { $w = $designW; $h = $designH }
$form.Size = New-Object System.Drawing.Size($w, $h)

$form.Add_Shown({
    try {
        $sc = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
        $nw = [Math]::Min($form.Width,  [Math]::Max(620, $sc.Width  - 40))
        $nh = [Math]::Min($form.Height, [Math]::Max(340, $sc.Height - 40))
        if ($nw -ne $form.Width -or $nh -ne $form.Height) { $form.Size = New-Object System.Drawing.Size($nw, $nh) }
        if ($form.Left -lt $sc.Left) { $form.Left = $sc.Left + 10 }
        if ($form.Top  -lt $sc.Top)  { $form.Top  = $sc.Top  + 10 }
    } catch { }
})

# ===========================================================================
# UI helpers
# ===========================================================================

function Set-Status {
    param([string]$Text)
    $lblStatus.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

function Build-CheckTree {
    $treeChecks.BeginUpdate()
    $treeChecks.Nodes.Clear()
    foreach ($cat in ($script:CheckCatalog | Group-Object Category)) {
        $catNode = New-Object System.Windows.Forms.TreeNode($cat.Name)
        foreach ($c in $cat.Group) {
            $leaf = New-Object System.Windows.Forms.TreeNode($c.Label)
            $leaf.Tag = $c.Id
            $leaf.Checked = [bool]$c.Default
            $null = $catNode.Nodes.Add($leaf)
        }
        $catNode.Checked = (@($catNode.Nodes | Where-Object { -not $_.Checked })).Count -eq 0
        $null = $treeChecks.Nodes.Add($catNode)
        $catNode.Expand()
    }
    $treeChecks.EndUpdate()
}

$treeChecks.Add_AfterCheck({
    param($sender,$e)
    if ($script:SuppressChecks) { return }
    $script:SuppressChecks = $true
    try {
        if ($e.Node.Nodes.Count -gt 0) {
            foreach ($child in $e.Node.Nodes) { $child.Checked = $e.Node.Checked }
        } else {
            $parent = $e.Node.Parent
            if ($parent) { $parent.Checked = (@($parent.Nodes | Where-Object { -not $_.Checked })).Count -eq 0 }
        }
    } finally { $script:SuppressChecks = $false }
})

function Set-CheckChecks {
    param([scriptblock]$Predicate)
    $script:SuppressChecks = $true
    try {
        foreach ($catNode in $treeChecks.Nodes) {
            foreach ($leaf in $catNode.Nodes) {
                $c = $script:CheckCatalog | Where-Object { $_.Id -eq [string]$leaf.Tag } | Select-Object -First 1
                $leaf.Checked = [bool](& $Predicate $c)
            }
            $catNode.Checked = (@($catNode.Nodes | Where-Object { -not $_.Checked })).Count -eq 0
        }
    } finally { $script:SuppressChecks = $false }
}

function Get-SelectedCheckDefs {
    $ids = New-Object System.Collections.Generic.List[string]
    foreach ($catNode in $treeChecks.Nodes) {
        foreach ($leaf in $catNode.Nodes) {
            if ($leaf.Checked) { $ids.Add([string]$leaf.Tag) }
        }
    }
    return @($script:CheckCatalog | Where-Object { $ids -contains $_.Id })
}

function Add-ResultRow {
    param($r)
    $it = New-Object System.Windows.Forms.ListViewItem($r.Category)
    $null = $it.SubItems.Add($r.Check)
    $null = $it.SubItems.Add($r.Object)
    $null = $it.SubItems.Add($r.Status)
    $null = $it.SubItems.Add($r.Summary)
    $null = $it.SubItems.Add($r.Recommendation)
    $it.Tag = $r
    switch ($r.Status) {
        'Fail'  { $it.ForeColor = $cRed }
        'Error' { $it.ForeColor = $cRed; $it.Font = $fBold }
        'Warn'  { $it.ForeColor = $cOrange }
        'Pass'  { $it.ForeColor = $cGreen }
        default { $it.ForeColor = $cGray }
    }
    $null = $lvResults.Items.Add($it)
    $it.EnsureVisible()
}

function Connect-Session {
    param([string]$TargetServer)

    Initialize-ConnectLog   # ensures $script:ConnectLogPath exists (also called at startup)

    Write-ConnectLog '==================================================================='
    Write-ConnectLog 'CONNECT attempt'
    Write-ConnectLog ('Host: PS {0}  Edition={1}  Apartment={2}' -f `
        $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, [System.Threading.Thread]::CurrentThread.GetApartmentState())
    Write-ConnectLog ('OS: {0}' -f [System.Environment]::OSVersion.VersionString)
    Write-ConnectLog ('RunAs: {0}' -f [System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    Write-ConnectLog ("Target: '{0}'  Creds: {1}" -f $TargetServer, $(if ($script:Ctx.Credential) { $script:Ctx.Credential.UserName } else { '(current Windows identity)' }))

    try {
        Set-Status "Connecting to '$TargetServer' ... this can take 20-60s; the window may look busy until it completes."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Connect-Veeam -TargetServer $TargetServer -Credential $script:Ctx.Credential
        Write-ConnectLog 'CONNECT SUCCESS'
        $verText = if ($script:Ctx.ProductBuild) { $script:Ctx.ProductBuild } else { 'version unknown' }
        $lblConnInfo.Text = "Connected: $($script:Ctx.Server)  ($verText)"
        $lblConnInfo.ForeColor = $cGreen
        $lblConnInfo2.Text = "Edition: $($script:Ctx.Edition)   |   Licensed to: $($script:Ctx.LicenseTo)   |   Module: $($script:Ctx.ModuleKind)"
        $btnRun.Enabled = $true
        Set-Status "Connected to $($script:Ctx.Server). Select checks and Run."
    } catch {
        Write-ConnectLog "CONNECT EXCEPTION: $($_.Exception.GetType().FullName): $($_.Exception.Message)"
        Write-ConnectLog "CONNECT EXCEPTION detail: $($_.Exception.ToString())"
        Write-ConnectLog "CONNECT script stack: $($_.ScriptStackTrace)"
        $script:Ctx.Connected = $false
        $btnRun.Enabled = $false
        $lblConnInfo.Text = 'Not connected.'
        $lblConnInfo.ForeColor = $cRed
        Set-Status "Connect failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Could not connect to Veeam.`r`n`r`n$($_.Exception.Message)`r`n`r`nDiagnostic log:`r`n$script:ConnectLogPath",
            'Not connected','OK','Warning') | Out-Null
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
}

# ===========================================================================
# Events
# ===========================================================================

$btnPresetQuick.Add_Click({ Set-CheckChecks -Predicate { param($c) [bool]$c.Default } })
$btnPresetFull.Add_Click({ Set-CheckChecks -Predicate { $true } })
$btnPresetNone.Add_Click({ Set-CheckChecks -Predicate { $false } })

function Show-CredentialDialog {
    <# A self-contained WinForms credential prompt. Get-Credential is avoided on
       purpose: under PowerShell 7 (and some hosts) it prompts in the console
       window behind the GUI instead of showing a dialog. This always shows a
       real dialog and returns a PSCredential, or $null if cancelled/blank. #>
    param([string]$Message = 'Enter credentials for the Veeam server', [string]$UserName = '')

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Veeam Credentials'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.StartPosition = 'CenterParent'
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.ClientSize = New-Object System.Drawing.Size(390,176)
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI',9)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Message
    $lbl.Location = New-Object System.Drawing.Point(12,12)
    $lbl.Size = New-Object System.Drawing.Size(366,34)
    $dlg.Controls.Add($lbl)

    $lblU = New-Object System.Windows.Forms.Label
    $lblU.Text = 'Username'; $lblU.Location = New-Object System.Drawing.Point(12,54); $lblU.Size = New-Object System.Drawing.Size(70,22)
    $dlg.Controls.Add($lblU)
    $txtU = New-Object System.Windows.Forms.TextBox
    $txtU.Location = New-Object System.Drawing.Point(90,52); $txtU.Size = New-Object System.Drawing.Size(288,22); $txtU.Text = $UserName
    $dlg.Controls.Add($txtU)

    $lblP = New-Object System.Windows.Forms.Label
    $lblP.Text = 'Password'; $lblP.Location = New-Object System.Drawing.Point(12,84); $lblP.Size = New-Object System.Drawing.Size(70,22)
    $dlg.Controls.Add($lblP)
    $txtP = New-Object System.Windows.Forms.TextBox
    $txtP.Location = New-Object System.Drawing.Point(90,82); $txtP.Size = New-Object System.Drawing.Size(288,22); $txtP.UseSystemPasswordChar = $true
    $dlg.Controls.Add($txtP)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = 'OK'; $btnOK.Location = New-Object System.Drawing.Point(210,130); $btnOK.Size = New-Object System.Drawing.Size(80,28)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOK)
    $btnCa = New-Object System.Windows.Forms.Button
    $btnCa.Text = 'Cancel'; $btnCa.Location = New-Object System.Drawing.Point(298,130); $btnCa.Size = New-Object System.Drawing.Size(80,28)
    $btnCa.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCa)
    $dlg.AcceptButton = $btnOK; $dlg.CancelButton = $btnCa

    $cred = $null
    try {
        $txtU.Select()
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txtU.Text.Trim()) {
            $sec = New-Object System.Security.SecureString
            foreach ($ch in $txtP.Text.ToCharArray()) { $sec.AppendChar($ch) }
            $sec.MakeReadOnly()
            $cred = New-Object System.Management.Automation.PSCredential($txtU.Text.Trim(), $sec)
        }
    } finally { $dlg.Dispose() }
    return $cred
}

$btnCreds.Add_Click({
    $existing = if ($script:Ctx.Credential) { $script:Ctx.Credential.UserName } else { '' }
    $c = Show-CredentialDialog -Message 'Alternate credentials for the Veeam server. Cancel = use your current Windows identity (recommended when running on the VBR server itself).' -UserName $existing
    if ($c) { $script:Ctx.Credential = $c; Set-Status "Alternate credentials set: $($c.UserName)" }
    else    { $script:Ctx.Credential = $null; Set-Status 'Using current Windows identity (credentials cleared).' }
})

$btnConnect.Add_Click({ Connect-Session -TargetServer $txtServer.Text.Trim() })

$btnOpenLog.Add_Click({
    $p = $txtLog.Text
    if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType File -Path $p -Force | Out-Null }
    Start-Process notepad.exe -ArgumentList $p
})

$btnClose.Add_Click({ $form.Close() })

$lvResults.Add_SelectedIndexChanged({
    if ($lvResults.SelectedItems.Count -eq 0) { return }
    $r = $lvResults.SelectedItems[0].Tag
    if ($r) {
        $txtDetail.Text = "Category      : $($r.Category)`r`nCheck         : $($r.Check)`r`nObject        : $($r.Object)`r`nStatus        : $($r.Status)`r`n`r`nSummary       : $($r.Summary)`r`n`r`nRecommendation: $($r.Recommendation)`r`n`r`n--- Detail ---`r`n$($r.Details)"
    }
})

$btnExportCsv.Add_Click({
    if ($script:Results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Run checks first.','Export','OK','Information') | Out-Null
        return
    }
    $d = New-Object System.Windows.Forms.SaveFileDialog
    $d.Filter = 'CSV (*.csv)|*.csv'
    $d.FileName = "VeeamDiag-$($script:Ctx.Server)-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
    if ($d.ShowDialog() -eq 'OK') {
        $script:Results | Select-Object Category,Check,Object,Status,Summary,Recommendation,Details |
            Export-Csv -LiteralPath $d.FileName -NoTypeInformation -Encoding UTF8
        Set-Status "Exported $($script:Results.Count) row(s) to CSV."
    }
})

$btnExportHtml.Add_Click({
    if ($script:Results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Run checks first.','Export','OK','Information') | Out-Null
        return
    }
    $d = New-Object System.Windows.Forms.SaveFileDialog
    $d.Filter = 'HTML report (*.html)|*.html'
    $d.FileName = "VeeamDiag-$($script:Ctx.Server)-$(Get-Date -Format yyyyMMdd-HHmmss).html"
    if ($d.ShowDialog() -eq 'OK') {
        Export-HtmlReport -Path $d.FileName
        Set-Status "HTML report written to $($d.FileName)."
        [System.Windows.Forms.MessageBox]::Show("Report saved.`r`n`r`n$($d.FileName)",'Export complete','OK','Information') | Out-Null
    }
})

$btnCancel.Add_Click({ $script:CancelRequested = $true; Set-Status 'Cancelling ...' })

$btnRun.Add_Click({
    if (-not $script:Ctx.Connected) {
        [System.Windows.Forms.MessageBox]::Show('Connect to a Veeam server first.','Run','OK','Warning') | Out-Null
        return
    }
    $checks = Get-SelectedCheckDefs
    if ($checks.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Select at least one check.','Run','OK','Warning') | Out-Null
        return
    }
    if ([int]$numFailPct.Value -le [int]$numWarnPct.Value) {
        [System.Windows.Forms.MessageBox]::Show('Repository Fail % must be greater than Warn %.','Run','OK','Warning') | Out-Null
        return
    }

    $options = [pscustomobject]@{
        RpoHours             = [int]$numRpo.Value
        SessionLookbackHours = [int]$numLookback.Value
        RepoWarnPct          = [int]$numWarnPct.Value
        RepoFailPct          = [int]$numFailPct.Value
        LicenseWarnDays      = [int]$numLicDays.Value
    }
    $logPath = $txtLog.Text.Trim()

    $script:CancelRequested = $false
    $script:Cache = @{}   # fresh data for this run
    $script:Results = New-Object System.Collections.Generic.List[object]
    $lvResults.Items.Clear()
    $txtDetail.Clear()
    $progress.Minimum = 0; $progress.Maximum = [Math]::Max($checks.Count,1); $progress.Value = 0
    $btnRun.Enabled = $false; $btnCancel.Enabled = $true
    $sumPass=0; $sumWarn=0; $sumFail=0; $sumErr=0

    try {
        foreach ($def in $checks) {
            if ($script:CancelRequested) { Set-Status 'Cancelled.'; break }
            $progress.Value = [Math]::Min($progress.Value+1, $progress.Maximum)
            Set-Status "[$($progress.Value)/$($progress.Maximum)] $($def.Category) - $($def.Label) ..."

            $rows = Invoke-VeeamCheck -Def $def -Options $options
            foreach ($r in $rows) {
                $script:Results.Add($r)
                Add-ResultRow $r
                switch ($r.Status) {
                    'Pass'  { $sumPass++ }
                    'Warn'  { $sumWarn++ }
                    'Fail'  { $sumFail++ }
                    'Error' { $sumErr++ }
                }
                if ($logPath) { Write-DiagLog -Path $logPath -Result $r }
            }
        }
    } finally {
        $btnRun.Enabled = $true; $btnCancel.Enabled = $false; $progress.Value = 0
    }

    $lblSummary.Text = "Completed: $($script:Results.Count) result(s).  Pass=$sumPass  Warn=$sumWarn  Fail=$sumFail  Error=$sumErr"
    $lblSummary.ForeColor = if ($sumFail -gt 0 -or $sumErr -gt 0) { $cRed } elseif ($sumWarn -gt 0) { $cOrange } else { $cGreen }
    Set-Status 'Health check complete. Review results and export a report.'
})

$form.Add_KeyDown({
    switch ($_.KeyCode) {
        'F5'     { if ($btnRun.Enabled)        { $btnRun.PerformClick() };        $_.Handled = $true }
        'F6'     { if ($btnCancel.Enabled)     { $btnCancel.PerformClick() };     $_.Handled = $true }
        'F7'     { if ($btnExportCsv.Enabled)  { $btnExportCsv.PerformClick() };  $_.Handled = $true }
        'F8'     { if ($btnExportHtml.Enabled) { $btnExportHtml.PerformClick() }; $_.Handled = $true }
        'Escape' { $form.Close(); $_.Handled = $true }
    }
})

$form.Add_FormClosed({ Disconnect-Veeam })

# ===========================================================================
# Go
# ===========================================================================

$form.Add_Shown({
    $form.Activate()
    Build-CheckTree
    Initialize-ConnectLog
    Set-Status "Ready. Connect log: $script:ConnectLogPath"
})
[void]$form.ShowDialog()
$form.Dispose()
