<#
.SYNOPSIS
    AD ReOrg - a WPF tool for Active Directory reorganisation.

.DESCRIPTION
    Connect to Active Directory and work with groups and users. A sidebar picks the
    tool; the group tools share one selection tree on the left.
      GROUPS
      * Create groups       - new groups from existing ones (copy members by category,
                              memberOf, attributes, description/notes, nest new in old).
      * Rename groups       - rename existing groups in place (CN, sAMAccountName,
                              displayName, description); SID, members and permissions
                              stay. Revert from a rename report.
      * Update attributes   - find/replace a text in any attribute of one group (its
                              attribute list shown) or many (add or replace values).
      * Copy members        - copy members from one group to another.
      COMPARE
      * Compare             - two groups or two users side by side, every attribute;
                              export the comparison (HTML/CSV/JSON).
      SNAPSHOTS
      * Snapshot users      - 1 export users, 2 compare two exports (what changed,
                              incl. SVG diagram), 3 add users back to lost groups.
      * Snapshot groups     - export the selected groups (OU, members, memberOf).
      ENVIRONMENT
      * OU structure        - export an OU tree in one domain, create it in another.
      * Schema attributes   - export the schema of user attributes (incl. Exchange and
                              custom), create them in a test forest's schema.
                              Both run the scripts in tools\ unchanged.

    The window is built in WPF/XAML (asset\xaml\MainWindow.xaml). All logic lives in
    functions\function\ and is dot-sourced. Application name and version are read
    from settings\config\version.json; settings from settings\config\config.json.

.NOTES
    Requires : Windows PowerShell 5.1 (STA) and the RSAT ActiveDirectory module.
    Run with : powershell.exe -STA -ExecutionPolicy Bypass -File ".\AD-ReOrg.ps1"
#>

# --- Root folder --------------------------------------------------------
if ($PSScriptRoot) { $script:ScriptRoot = $PSScriptRoot }
else               { $script:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition }

$FunctionsPath = Join-Path $script:ScriptRoot 'functions\function'

# --- Unblock files (Zone.Identifier) so nothing is blocked after unzip --
try {
    Get-ChildItem -Path $script:ScriptRoot -Include *.ps1, *.xaml, *.json, *.cmd -Recurse -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue
}
catch { Write-Warning "Could not run Unblock-File: $($_.Exception.Message)" }

# --- .NET/WPF assemblies + AD module + bindable types -------------------
# Loaded from the separate Modules.ps1 (dot-sourced below), then run via
# Import-AppModules once the function files are in place.

# --- Shared state (used by the dot-sourced function files) --------------
$script:App             = $null   # application info from version.json
$script:Config          = $null   # settings from config.json
$script:CheckMap        = @{}      # checkbox -> tree node
$script:Suppress        = 0
$script:GroupCache      = @{}
$script:Plans           = $null    # ObservableCollection[ReOrg.GroupPlan]
$script:TargetOuDN      = ''
$script:LastRunRecords  = @()
$script:CompareGroupA   = $null    # Groups Compare - loaded group A
$script:CompareGroupB   = $null    # Groups Compare - loaded group B
$script:CompareUserA    = $null    # User Compare  - loaded user A
$script:CompareUserB    = $null    # User Compare  - loaded user B
$script:SnapUsers       = $null
$script:SnapCheckMap    = @{}
$script:SnapDiff        = $null

# --- Load application info (settings\config\version.json) ---------------
function Get-AppInfo {
    $default = [pscustomobject]@{
        name = 'AD ReOrg'; shortName = 'ReOrg'; version = '0.0.0'
        company = ''; badge = 'WO'; description = ''
    }
    $path = Join-Path $script:ScriptRoot 'settings\config\version.json'
    if (-not (Test-Path -LiteralPath $path)) { return $default }
    try {
        $j = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
        foreach ($p in $default.PSObject.Properties.Name) {
            if ($null -ne $j.$p -and "$($j.$p)" -ne '') { $default.$p = $j.$p }
        }
    } catch { Write-Warning "Could not read settings\config\version.json: $($_.Exception.Message)" }
    return $default
}
$script:App = Get-AppInfo

# --- Load function files ------------------------------------------------
$requiredFunctionFiles = @(
    'Modules.ps1'
    'Config.ps1'
    'Ui-Common.ps1'
    'Report.ps1'
    'Tree.ps1'
    'Groups.ps1'
    'CopyMembers.ps1'
    'Compare.ps1'
    'Snapshot.ps1'
    'Analyse.ps1'
    'UrlMod.ps1'
    'Restore.ps1'
    'Rename.ps1'
    'Picker.ps1'
    'Environment.ps1'
    'Show-MainWindow.ps1'
)
foreach ($file in $requiredFunctionFiles) {
    $full = Join-Path $FunctionsPath $file
    if (-not (Test-Path -LiteralPath $full)) { throw "Missing function file: $full" }
    . $full
}

# --- Load .NET/WPF assemblies + AD module + bindable types --------------
Import-AppModules

# --- Load configuration + set up logging --------------------------------
$script:Config = Import-AppConfig -Path (Join-Path $script:ScriptRoot 'settings\config\config.json')

$logDir = Join-Path $script:ScriptRoot ($script:Config.logging.folder)
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
$script:LogFile = Join-Path $logDir ('ReOrg_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
# Logs older than logging.retentionDays are removed at start (0 or empty = keep everything).
$retention = 0; [void][int]::TryParse([string]$script:Config.logging.retentionDays, [ref]$retention)
if ($retention -gt 0) {
    $limit = (Get-Date).AddDays(-$retention)
    Get-ChildItem -LiteralPath $logDir -Filter 'ReOrg_*.log' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $limit } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

# Root DN: detected from AD at every start when config says "auto" (or is empty).
$script:DefaultSearchBase = Resolve-DefaultSearchBase -Configured ([string]$script:Config.domain.searchBase) -Server ([string]$script:Config.domain.server)

# --- Launch UI ----------------------------------------------------------
Show-MainWindow
