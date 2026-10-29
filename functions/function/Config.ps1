<#
    Loads settings\config\config.json and fills in default values for any missing
    keys. Logs are never deleted automatically.
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

function Import-AppConfig {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$Path)

    $default = [pscustomobject]@{
        domain   = [pscustomobject]@{ server = ''; searchBase = '' }
        defaults = [pscustomobject]@{
            namePrefix    = ''
            nameSuffix    = '-NEW'
            samSuffix     = ''
            nameFind      = ''
            nameReplace   = ''
            nameRegex     = $false
            description   = ''
            notes         = ''
            copyDescription = $true
            addDescription  = $false
            copyNotes       = $true
            addNotes        = $false
            groupScope    = 'Global'
            groupCategory = 'Security'
            copyUsers     = $true
            copyGroups    = $true
            copyComputers = $false
            copyMemberOf  = $false
            copyAttributes = ''
            onlyCopyExisting = $false
            nestNewInOld  = $true
            whatIf        = $true
            continueOnError = $false
        }
        report   = [pscustomobject]@{ folder = 'files'; fileName = 'report'; openAfter = $true }
        logging  = [pscustomobject]@{ folder = 'logs'; retentionDays = 30 }
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Warning "config.json missing ($Path) - using defaults."
        return $default
    }

    try {
        $loaded = (Get-Content -LiteralPath $Path -Raw -Encoding UTF8) | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not read config.json: $($_.Exception.Message) - using defaults."
        return $default
    }

    foreach ($section in 'domain','defaults','report','logging') {
        if ($null -eq $loaded.$section) { continue }
        foreach ($prop in $default.$section.PSObject.Properties.Name) {
            $val = $loaded.$section.$prop
            if ($null -ne $val) { $default.$section.$prop = $val }
        }
    }
    return $default
}

function Resolve-DefaultSearchBase {
    # Returns the root DN to start from. config.json domain.searchBase = "auto" (or
    # empty) -> read the domain's root DN from AD at every start; any other value
    # is used as-is. Falls back to an empty box if AD cannot be reached.
    param([string]$Configured, [string]$Server)
    $c = ([string]$Configured).Trim()
    if ($c -and $c -ne 'auto') { return $c }
    $p = @{ ErrorAction = 'Stop' }
    if ($Server) { $p.Server = $Server }
    try { $dn = [string](Get-ADRootDSE @p).defaultNamingContext; if ($dn) { return $dn } } catch { }
    try { $dn = [string](Get-ADDomain @p).DistinguishedName;   if ($dn) { return $dn } } catch { }
    Write-Warning 'Could not detect the domain root DN automatically - enter the SearchBase manually.'
    return ''
}
