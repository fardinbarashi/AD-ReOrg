<#
.SYNOPSIS
    Exports the schema definitions for a user's attributes from production to a JSON file.

.DESCRIPTION
    Run in PRODUCTION. Nothing in AD is changed; the script only reads the schema.

    Reads the attribute names from a UTF-8 CSV file (column "Attribute", or "Falt" from
    the earlier Swedish version) or directly from a user (-Identity). For each attribute
    the full definition is saved (OID, syntax, schemaIDGUID, link ID, ranges etc.) together
    with the class that holds the attribute. The classes in the user hierarchy (user,
    person, auxiliary classes added by Exchange etc.) are exported as well, so that
    Import-SchemaAttributes.ps1 can recreate them in the test environment.

    For every backward link (-BL) the matching forward link is exported too, since a
    backward link cannot be created without it.

.EXAMPLE
    .\Export-SchemaAttributes.ps1 -CsvPath C:\temp\UserAttributes.csv -Path C:\temp\SchemaAttributes.json

.EXAMPLE
    .\Export-SchemaAttributes.ps1 -Identity jdoe -Path C:\temp\SchemaAttributes.json
#>
#Requires -Modules ActiveDirectory
[CmdletBinding(DefaultParameterSetName = 'Csv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$CsvPath,

    [Parameter(Mandatory, ParameterSetName = 'User')]
    [string]$Identity,

    # Only attributes that have a value (InUse/Anvands = True in the CSV file)
    [Parameter(ParameterSetName = 'Csv')]
    [switch]$OnlyUsed,

    [string]$Path = '.\SchemaAttributes.json'
)

Import-Module ActiveDirectory -ErrorAction Stop

$schemaNC = (Get-ADRootDSE).schemaNamingContext

# ---------- 1. Which attributes to export ----------
if ($PSCmdlet.ParameterSetName -eq 'Csv') {
    $rows = @(Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8)
    if (-not $rows) { throw "The CSV file '$CsvPath' is empty." }
    $cols = $rows[0].PSObject.Properties.Name
    # Accept both the English and the earlier Swedish column names
    $nameCol  = @('Attribute', 'Falt') | Where-Object { $cols -contains $_ } | Select-Object -First 1
    $inUseCol = @('InUse', 'Anvands') | Where-Object { $cols -contains $_ } | Select-Object -First 1
    if (-not $nameCol) { throw "The CSV file must have a column named 'Attribute' (or 'Falt'). Columns found: $($cols -join ', ')" }
    if ($OnlyUsed) {
        if (-not $inUseCol) { throw "-OnlyUsed requires a column named 'InUse' (or 'Anvands')." }
        $rows = $rows | Where-Object { $_.$inUseCol -eq 'True' }
    }
    $names = @($rows | ForEach-Object { $_.$nameCol })
} else {
    $u  = Get-ADUser -Identity $Identity -ErrorAction Stop
    $de = [ADSI]"LDAP://$($u.DistinguishedName)"
    $de.RefreshCache([string[]]'allowedAttributes')
    $names = @($de.Properties['allowedAttributes'])
}
$names = $names | Where-Object { $_ } | Sort-Object -Unique
if (-not $names) { throw 'No attribute names found to export.' }

# ---------- 2. Read the whole schema once ----------
$attrProps = 'lDAPDisplayName', 'cn', 'attributeID', 'attributeSyntax', 'oMSyntax', 'oMObjectClass',
             'isSingleValued', 'rangeLower', 'rangeUpper', 'searchFlags', 'isMemberOfPartialAttributeSet',
             'adminDescription', 'adminDisplayName', 'schemaIDGUID', 'attributeSecurityGUID', 'linkID',
             'mAPIID', 'showInAdvancedViewOnly', 'systemOnly', 'systemFlags'

$classProps = 'lDAPDisplayName', 'cn', 'governsID', 'objectClassCategory', 'subClassOf',
              'mayContain', 'systemMayContain', 'mustContain', 'systemMustContain',
              'auxiliaryClass', 'systemAuxiliaryClass', 'possSuperiors', 'systemPossSuperiors',
              'schemaIDGUID', 'defaultObjectCategory', 'defaultSecurityDescriptor',
              'adminDescription', 'adminDisplayName', 'defaultHidingValue', 'systemOnly', 'systemFlags'

Write-Host "Reading schema: $schemaNC" -ForegroundColor Cyan
$attrByName = @{}
Get-ADObject -SearchBase $schemaNC -SearchScope OneLevel -LDAPFilter '(objectClass=attributeSchema)' -Properties $attrProps |
    ForEach-Object { $attrByName[$_.lDAPDisplayName] = $_ }
$classByName = @{}
Get-ADObject -SearchBase $schemaNC -SearchScope OneLevel -LDAPFilter '(objectClass=classSchema)' -Properties $classProps |
    ForEach-Object { $classByName[$_.lDAPDisplayName] = $_ }

# Convert AD values into something that can be stored in JSON (binary -> Base64)
function ConvertTo-Plain($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [byte[]]) { return [pscustomobject]@{ Base64 = [Convert]::ToBase64String($v) } }
    if ($v -is [System.Collections.ICollection] -and $v -isnot [string]) {
        return , @($v | ForEach-Object { ConvertTo-Plain $_ })
    }
    return $v
}
function ConvertTo-Definition($obj, $props) {
    $d = [ordered]@{}
    foreach ($p in $props) {
        $v = ConvertTo-Plain $obj.$p
        if ($null -ne $v -and "$v" -ne '') { $d[$p] = $v }
    }
    $d
}

# ---------- 3. Class hierarchy for user (superclasses + auxiliary classes) ----------
# Reason: Hierarchy = user or one of its superclasses
#         Aux       = auxiliary class attached to the class in Related
#         Super     = superclass of the auxiliary class in Related
$classInfo = [ordered]@{}
$queue = New-Object System.Collections.Queue
function Add-Class($name, $reason, $related) {
    if (-not $name -or $classInfo.Contains($name) -or -not $classByName[$name]) { return }
    $classInfo[$name] = [ordered]@{ Reason = $reason; Related = $related }
    $queue.Enqueue($name)
}
Add-Class 'user' 'Hierarchy' $null
while ($queue.Count) {
    $name = $queue.Dequeue()
    $c = $classByName[$name]
    $superReason = if ($classInfo[$name].Reason -eq 'Hierarchy') { 'Hierarchy' } else { 'Super' }
    Add-Class $c.subClassOf $superReason $name
    foreach ($aux in @($c.auxiliaryClass) + @($c.systemAuxiliaryClass)) { Add-Class $aux 'Aux' $name }
}

function Get-ClassAttributes($c) {
    @($c.mayContain) + @($c.systemMayContain) + @($c.mustContain) + @($c.systemMustContain) | Where-Object { $_ }
}

function Get-Category($o) {
    if ($o.cn -like 'ms-Exch-*') { 'Exchange' }
    elseif (([int]$o.systemFlags -band 16) -ne 0) { 'Base' }      # part of the Windows base schema
    else { 'Custom' }
}

# ---------- 4. Build the export ----------
$missing = @()
$attributes = foreach ($name in $names) {
    $a = $attrByName[$name]
    if (-not $a) { $missing += $name; continue }

    $holder = $null
    foreach ($cn in $classInfo.Keys) {
        if ((Get-ClassAttributes $classByName[$cn]) -contains $name) { $holder = $cn; break }
    }

    [ordered]@{
        Name       = $a.lDAPDisplayName
        Category   = Get-Category $a
        Role       = 'User'
        Holder     = $holder
        Definition = ConvertTo-Definition $a $attrProps
    }
}

# Backward links (odd linkID) require their forward link. The forward link usually lives
# on another object (e.g. an Exchange policy) and is therefore not among the user's
# attributes - include it.
$attrByLink = @{}
foreach ($a in $attrByName.Values) { if ($null -ne $a.linkID) { $attrByLink[[int]$a.linkID] = $a } }
$exported = @{}; foreach ($x in @($attributes)) { $exported[$x.Name] = $true }
$forwardLinks = foreach ($x in @($attributes)) {
    $lid = $x.Definition['linkID']
    if ($null -eq $lid -or ([int]$lid % 2) -ne 1) { continue }
    $fwd = $attrByLink[[int]$lid - 1]
    if (-not $fwd) { Write-Warning "No forward link found for $($x.Name) (linkID $lid)"; continue }
    if ($exported[$fwd.lDAPDisplayName]) { continue }
    $exported[$fwd.lDAPDisplayName] = $true
    [ordered]@{
        Name       = $fwd.lDAPDisplayName
        Category   = Get-Category $fwd
        Role       = 'ForwardLink'
        Holder     = $null
        Definition = ConvertTo-Definition $fwd $attrProps
    }
}
$attributes = @($attributes) + @($forwardLinks)

$classes = foreach ($cn in $classInfo.Keys) {
    $c = $classByName[$cn]
    [ordered]@{
        Name       = $c.lDAPDisplayName
        Reason     = $classInfo[$cn].Reason
        Related    = $classInfo[$cn].Related
        Category   = Get-Category $c
        Definition = ConvertTo-Definition $c $classProps
    }
}

$export = [ordered]@{
    ExportedFrom = (Get-ADDomain).DNSRoot
    ExportedAt   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    SchemaNC     = $schemaNC
    Attributes   = @($attributes)
    Classes      = @($classes)
}

$fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
[IO.File]::WriteAllText($fullPath, ($export | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding $true))

Write-Host "`nDone! Saved: $fullPath" -ForegroundColor Green
@($attributes) | Where-Object { $_.Role -eq 'User' } | Group-Object { $_.Category } | ForEach-Object { "  {0,-13} {1} attributes" -f $_.Name, $_.Count }
"  Forward links {0} (required by the backward links)" -f @($forwardLinks).Count
"  Classes       {0}" -f @($classes).Count
if ($missing) { Write-Warning "Not found in the schema (skipped): $($missing -join ', ')" }
