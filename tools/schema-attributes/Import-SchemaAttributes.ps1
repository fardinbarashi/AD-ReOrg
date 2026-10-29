<#
.SYNOPSIS
    Creates the attributes and classes from Export-SchemaAttributes.ps1 in the test environment's AD schema.

.DESCRIPTION
    Run in the TEST ENVIRONMENT by a member of Schema Admins. Changes are made on the schema master.

    WARNING: Schema changes can NOT be removed, only deactivated. Run with -WhatIf first.

    What it does:
      1. Compares the export with the test environment's schema.
      2. Creates the missing attributes with the same OID, syntax, schemaIDGUID, link ID etc.
         as in production. Linked attributes are created forward link first, backward link last.
      3. Creates the classes that hold the attributes in production (e.g. Exchange auxiliary
         classes) and attaches them to user the same way as in production.
      4. Adds attributes to existing classes (mayContain) when needed.
      5. Verifies that every attribute can now be used on a user object and saves a
         UTF-8 CSV report.

    Missing attributes from the Windows base schema mean the test environment has an
    older schema version. They cannot be created manually; run adprep from a newer
    Windows Server instead.

    Also reads JSON files created by the earlier Swedish version (Export-SchemaAttribut.ps1).

.EXAMPLE
    .\Import-SchemaAttributes.ps1 -Path C:\temp\SchemaAttributes.json -WhatIf

.EXAMPLE
    .\Import-SchemaAttributes.ps1 -Path C:\temp\SchemaAttributes.json

.EXAMPLE
    .\Import-SchemaAttributes.ps1 -Path C:\temp\SchemaAttributes.json -Attribute spAlternateNumID,sisSchoolGrade
    Creates only the listed attributes.
#>
#Requires -Modules ActiveDirectory
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    # Limit to these attributes (lDAPDisplayName)
    [string[]]$Attribute,

    # Skip Exchange attributes and classes
    [switch]$ExcludeExchange,

    # Do not ask for confirmation
    [switch]$Force,

    [string]$ReportPath
)

Import-Module ActiveDirectory -ErrorAction Stop

# UTF-8 with BOM in both Windows PowerShell 5.1 and PowerShell 7 (so Excel shows å/ä/ö correctly)
$csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }

# ---------- Read the export ----------
$fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
$data     = [IO.File]::ReadAllText($fullPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
if (-not $ReportPath) { $ReportPath = Join-Path (Split-Path $fullPath) ("SchemaImport-Report_{0:yyyyMMdd-HHmm}.csv" -f (Get-Date)) }

# Files from the earlier Swedish version use Swedish category names
$categoryMap = @{ 'Bas' = 'Base'; 'Egen' = 'Custom' }
foreach ($o in @($data.Attributes) + @($data.Classes)) {
    if ($categoryMap.ContainsKey("$($o.Category)")) { $o.Category = $categoryMap["$($o.Category)"] }
}

$server   = (Get-ADForest).SchemaMaster
$schemaNC = (Get-ADRootDSE -Server $server).schemaNamingContext
$srcNC    = $data.SchemaNC

Write-Host "Schema master: $server" -ForegroundColor Cyan
Write-Host "Schema:        $schemaNC" -ForegroundColor Cyan

$srcAttr  = @{}; foreach ($a in $data.Attributes) { $srcAttr[$a.Name] = $a }
$srcClass = @{}; foreach ($c in $data.Classes)    { $srcClass[$c.Name] = $c }

# The user's attributes. Forward links (Role = ForwardLink) are only created so the backward links work.
$scope = @($data.Attributes | Where-Object { $_.Role -ne 'ForwardLink' })
if ($Attribute) {
    $unknown = $Attribute | Where-Object { -not $srcAttr.ContainsKey($_) }
    if ($unknown) { Write-Warning "Not found in the export: $($unknown -join ', ')" }
    $scope = @($scope | Where-Object { $Attribute -contains $_.Name })
}

# ---------- Functions ----------
$report = New-Object System.Collections.Generic.List[object]
function Add-Report($name, $type, $status, $message) {
    $report.Add([pscustomobject]@{ Name = $name; Type = $type; Status = $status; Message = $message })
}

function ConvertFrom-Plain($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [System.Management.Automation.PSCustomObject] -and $v.PSObject.Properties['Base64']) {
        return , [Convert]::FromBase64String($v.Base64)
    }
    if ($v -is [array]) { return , @($v | ForEach-Object { ConvertFrom-Plain $_ }) }
    return $v
}
function New-Set($items) {
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($i in @($items)) { if ($i) { [void]$set.Add("$i") } }
    , $set
}
function New-StrList($items) {
    $l = New-Object System.Collections.Generic.List[string]
    foreach ($i in @($items)) { if ($i) { $l.Add("$i") } }
    , $l
}
function Get-List($obj, [string[]]$props) {
    @(foreach ($p in $props) { $obj.$p }) | Where-Object { $_ } | ForEach-Object { "$_" }
}

$script:refreshNeeded = $false
function Update-SchemaCache {
    if ($WhatIfPreference -or -not $script:refreshNeeded) { return }
    Write-Host '  Refreshing the schema cache...' -ForegroundColor DarkGray
    $rootDse = [ADSI]"LDAP://$server/RootDSE"
    $rootDse.Put('schemaUpdateNow', 1)
    $rootDse.SetInfo()
    $script:refreshNeeded = $false
}

function Invoke-Step($target, $action, [scriptblock]$do) {
    if (-not $PSCmdlet.ShouldProcess($target, $action)) { return $true }   # -WhatIf: keep planning as if it succeeded
    try {
        & $do
        $script:refreshNeeded = $true
        return $true
    } catch {
        Write-Warning "$action failed: $target - $($_.Exception.Message)"
        return $false
    }
}

# ---------- Read the test environment's schema ----------
$attrSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$oidMap  = @{}
$linkMap = @{}
Get-ADObject -Server $server -SearchBase $schemaNC -SearchScope OneLevel -LDAPFilter '(objectClass=attributeSchema)' `
        -Properties lDAPDisplayName, attributeID, linkID |
    ForEach-Object {
        [void]$attrSet.Add($_.lDAPDisplayName); $oidMap[$_.attributeID] = $_.lDAPDisplayName
        if ($null -ne $_.linkID) { $linkMap[[int]$_.linkID] = $_.lDAPDisplayName }
    }

$classProps = 'lDAPDisplayName', 'subClassOf', 'mayContain', 'systemMayContain', 'mustContain', 'systemMustContain',
              'auxiliaryClass', 'systemAuxiliaryClass', 'governsID'
$tClass = @{}
$govMap = @{}
Get-ADObject -Server $server -SearchBase $schemaNC -SearchScope OneLevel -LDAPFilter '(objectClass=classSchema)' -Properties $classProps |
    ForEach-Object {
        $tClass[$_.lDAPDisplayName] = @{
            DN    = $_.DistinguishedName
            Super = "$($_.subClassOf)"
            Attrs = New-Set (Get-List $_ 'mayContain', 'systemMayContain', 'mustContain', 'systemMustContain')
            Aux   = New-StrList (Get-List $_ 'auxiliaryClass', 'systemAuxiliaryClass')
        }
        $govMap[$_.governsID] = $_.lDAPDisplayName
    }

# Classes that user has in the test environment (superclasses + auxiliary classes), and their attributes
function Get-UserClasses {
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $q = New-Object System.Collections.Queue; $q.Enqueue('user')
    while ($q.Count) {
        $n = $q.Dequeue()
        if (-not $n -or -not $tClass.ContainsKey($n) -or -not $seen.Add($n)) { continue }
        $q.Enqueue($tClass[$n].Super)
        foreach ($x in $tClass[$n].Aux) { $q.Enqueue($x) }
    }
    , $seen
}
function Get-UserAttributes {
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($c in (Get-UserClasses)) { $set.UnionWith($tClass[$c].Attrs) }
    , $set
}

# ---------- Planning ----------
$toCreate = New-Object System.Collections.Generic.List[object]
foreach ($a in $scope) {
    if ($attrSet.Contains($a.Name)) { continue }
    $oid = $a.Definition.attributeID
    if ($a.Category -eq 'Base') {
        Add-Report $a.Name 'Attribute' 'Not created' 'Part of the Windows base schema - the test environment has an older schema version (run adprep from a newer Windows Server)'
    } elseif ($ExcludeExchange -and $a.Category -eq 'Exchange') {
        Add-Report $a.Name 'Attribute' 'Skipped' 'Exchange attribute (-ExcludeExchange)'
    } elseif ($oidMap.ContainsKey($oid)) {
        Add-Report $a.Name 'Attribute' 'Conflict' "OID $oid is already used by '$($oidMap[$oid])'"
    } else {
        $toCreate.Add($a)
    }
}

# Backward links need their forward link - add it if it is missing in the test environment
function Test-BackLink($a) { $a.Definition.PSObject.Properties['linkID'] -and ([int]$a.Definition.linkID % 2 -eq 1) }
foreach ($a in @($toCreate | Where-Object { Test-BackLink $_ })) {
    $fwdId = [int]$a.Definition.linkID - 1
    if ($linkMap.ContainsKey($fwdId)) { continue }                       # forward link already exists in the test environment
    if ($toCreate | Where-Object { $_.Definition.PSObject.Properties['linkID'] -and [int]$_.Definition.linkID -eq $fwdId }) { continue }
    $fwd = $data.Attributes | Where-Object { $_.Definition.PSObject.Properties['linkID'] -and [int]$_.Definition.linkID -eq $fwdId } | Select-Object -First 1
    $why = $null
    if (-not $fwd)                                   { $why = 'The forward link is missing from the export - rerun Export-SchemaAttributes.ps1 in production' }
    elseif ($fwd.Category -eq 'Base')                { $why = "The forward link '$($fwd.Name)' is part of the Windows base schema - requires adprep" }
    elseif ($ExcludeExchange -and $fwd.Category -eq 'Exchange') { $why = "The forward link '$($fwd.Name)' is an Exchange attribute (-ExcludeExchange)" }
    elseif ($oidMap.ContainsKey($fwd.Definition.attributeID)) { $why = "The forward link's OID is already used by '$($oidMap[$fwd.Definition.attributeID])'" }
    if ($why) {
        Add-Report $a.Name 'Attribute' 'Not created' "Backward link: $why"
        [void]$toCreate.Remove($a)
    } else {
        $toCreate.Add($fwd)
    }
}

$exchCount = @($toCreate | Where-Object Category -eq 'Exchange').Count
Write-Host "`nAttributes in the export:       $($scope.Count)"
Write-Host "Already in the test environment: $(@($scope | Where-Object { $attrSet.Contains($_.Name) }).Count)"
Write-Host "To be created:                   $($toCreate.Count)  (Exchange: $exchCount, forward links: $(@($toCreate | Where-Object Role -eq 'ForwardLink').Count))" -ForegroundColor Yellow

if ($toCreate.Count -gt 0 -and -not $Force -and -not $WhatIfPreference) {
    $msg = "The schema in $schemaNC will be extended with $($toCreate.Count) attributes. This can NOT be undone. Continue?"
    if (-not $PSCmdlet.ShouldContinue($msg, 'Schema change')) { Write-Host 'Cancelled.'; return }
}

# ---------- 1. Create attributes (forward links and regular first, backward links last) ----------
$attrCopy = 'lDAPDisplayName', 'attributeID', 'attributeSyntax', 'oMSyntax', 'oMObjectClass', 'isSingleValued',
            'rangeLower', 'rangeUpper', 'searchFlags', 'isMemberOfPartialAttributeSet', 'adminDescription',
            'adminDisplayName', 'schemaIDGUID', 'attributeSecurityGUID', 'linkID', 'mAPIID', 'showInAdvancedViewOnly'

function New-SchemaAttribute($a) {
    $d = $a.Definition
    $other = @{}
    foreach ($p in $attrCopy) {
        if ($d.PSObject.Properties[$p]) { $other[$p] = ConvertFrom-Plain $d.$p }
    }
    $dn = "CN=$($d.cn),$schemaNC"
    $ok = Invoke-Step $dn 'Create attribute' {
        New-ADObject -Server $server -Name $d.cn -Type attributeSchema -Path $schemaNC -OtherAttributes $other -ErrorAction Stop
    }
    if ($ok) {
        [void]$attrSet.Add($a.Name)
        if ($d.PSObject.Properties['linkID']) { $linkMap[[int]$d.linkID] = $a.Name }
        $msg = if ($a.Role -eq 'ForwardLink') { "$($a.Category) - forward link required by the backward link" } else { $a.Category }
        Add-Report $a.Name 'Attribute' $(if ($WhatIfPreference) { 'WhatIf' } else { 'Created' }) $msg
    } else {
        Add-Report $a.Name 'Attribute' 'Error' 'Could not be created, see warning'
    }
}

$first = @($toCreate | Where-Object { -not (Test-BackLink $_) })
$back  = @($toCreate | Where-Object { Test-BackLink $_ })

if ($toCreate.Count) { Write-Host "`n--- Creating attributes ---" -ForegroundColor Cyan }
foreach ($a in $first) { New-SchemaAttribute $a }
Update-SchemaCache
foreach ($a in $back) {
    $fwdId = [int]$a.Definition.linkID - 1
    if (-not $linkMap.ContainsKey($fwdId)) {
        Add-Report $a.Name 'Attribute' 'Not created' "Backward link - the forward link (linkID $fwdId) could not be created"
        continue
    }
    New-SchemaAttribute $a
}
Update-SchemaCache

# ---------- 2. Create the classes that are needed ----------
$failedClass = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
function New-SchemaClass([string]$name) {
    if ($tClass.ContainsKey($name)) { return $true }
    if ($failedClass.Contains($name)) { return $false }
    $ok = New-SchemaClassCore $name
    if (-not $ok) { [void]$failedClass.Add($name) }
    $ok
}
function New-SchemaClassCore([string]$name) {
    $src = $srcClass[$name]
    if (-not $src) { Add-Report $name 'Class' 'Error' 'Not found in the export'; return $false }
    if ($src.Category -eq 'Base') { Add-Report $name 'Class' 'Not created' 'Part of the Windows base schema - requires adprep'; return $false }
    if ($ExcludeExchange -and $src.Category -eq 'Exchange') { Add-Report $name 'Class' 'Skipped' '-ExcludeExchange'; return $false }

    $d = $src.Definition
    if ($govMap.ContainsKey($d.governsID)) {
        Add-Report $name 'Class' 'Conflict' "OID $($d.governsID) is already used by '$($govMap[$d.governsID])'"; return $false
    }
    if ($d.subClassOf -and -not (New-SchemaClass $d.subClassOf)) {
        Add-Report $name 'Class' 'Error' "The superclass '$($d.subClassOf)' is missing"; return $false
    }
    # A newly created superclass/attribute must be in the schema cache before the class can reference it
    Update-SchemaCache

    $must = @(Get-List $d 'mustContain', 'systemMustContain')
    $lack = @($must | Where-Object { -not $attrSet.Contains($_) })
    if ($lack) { Add-Report $name 'Class' 'Error' "Mandatory attributes are missing: $($lack -join ', ')"; return $false }

    $may  = @(Get-List $d 'mayContain', 'systemMayContain' | Where-Object { $attrSet.Contains($_) } | Sort-Object -Unique)
    $skippedMay = @(Get-List $d 'mayContain', 'systemMayContain' | Where-Object { -not $attrSet.Contains($_) })
    $aux  = @(Get-List $d 'auxiliaryClass', 'systemAuxiliaryClass' | Where-Object { $tClass.ContainsKey($_) })
    $poss = @(Get-List $d 'possSuperiors', 'systemPossSuperiors' | Where-Object { $tClass.ContainsKey($_) })

    $other = @{}
    foreach ($p in 'lDAPDisplayName', 'governsID', 'objectClassCategory', 'subClassOf', 'schemaIDGUID',
                   'defaultSecurityDescriptor', 'adminDescription', 'adminDisplayName', 'defaultHidingValue') {
        if ($d.PSObject.Properties[$p]) { $other[$p] = ConvertFrom-Plain $d.$p }
    }
    # defaultObjectCategory normally points to the class itself; replace the production schema DN with the test one
    $self = "CN=$($d.cn),$schemaNC"
    $doc  = "$($d.defaultObjectCategory)"
    if ($doc -like "*,$srcNC") { $doc = $doc.Substring(0, $doc.Length - $srcNC.Length) + $schemaNC }
    if ($doc -ne $self -and -not ($tClass.Values | Where-Object { $_.DN -eq $doc })) { $doc = $self }
    $other.defaultObjectCategory = $doc
    if ($may)  { $other.mayContain     = [string[]]$may }
    if ($must) { $other.mustContain    = [string[]]$must }
    if ($aux)  { $other.auxiliaryClass = [string[]]$aux }
    if ($poss) { $other.possSuperiors  = [string[]]$poss }

    $dn = "CN=$($d.cn),$schemaNC"
    $ok = Invoke-Step $dn 'Create class' {
        New-ADObject -Server $server -Name $d.cn -Type classSchema -Path $schemaNC -OtherAttributes $other -ErrorAction Stop
    }
    if (-not $ok) { Add-Report $name 'Class' 'Error' 'Could not be created, see warning'; return $false }

    $tClass[$name] = @{
        DN    = $dn
        Super = "$($d.subClassOf)"
        Attrs = New-Set ($may + $must)
        Aux   = New-StrList $aux
    }
    $msg = $src.Category
    if ($skippedMay) { $msg += " (without $($skippedMay.Count) attributes that do not exist in the test environment)" }
    Add-Report $name 'Class' $(if ($WhatIfPreference) { 'WhatIf' } else { 'Created' }) $msg
    return $true
}

# The chain of classes needed for a class to reach user (e.g. auxiliary class -> user)
function Get-AttachChain([string]$name) {
    $chain = New-Object System.Collections.Generic.List[string]
    $userClasses = Get-UserClasses
    while ($name -and -not $userClasses.Contains($name) -and -not $chain.Contains($name)) {
        $chain.Add($name)
        $info = $srcClass[$name]
        if (-not $info -or $info.Reason -eq 'Hierarchy') { break }
        $name = $info.Related
    }
    , $chain
}

$userAttrs = Get-UserAttributes
$needAttach = @($scope | Where-Object { $attrSet.Contains($_.Name) -and -not $userAttrs.Contains($_.Name) })
# Windows base attributes are never changed - they should come from adprep
$baseNotOnUser = @($needAttach | Where-Object Category -eq 'Base')
foreach ($a in $baseNotOnUser) { Add-Report $a.Name 'Attribute' 'Not on user' 'Base attribute that is not available on user in the test environment - requires adprep, not changed' }
$needAttach = @($needAttach | Where-Object Category -ne 'Base')

if ($needAttach) { Write-Host "`n--- Creating and attaching classes ---" -ForegroundColor Cyan }
foreach ($a in $needAttach) {
    if ($a.Holder) { foreach ($c in (Get-AttachChain $a.Holder)) { [void](New-SchemaClass $c) } }
}
Update-SchemaCache

# ---------- 3. Attach classes to user and add attributes to existing classes ----------
function Connect-Class([string]$name) {
    if ((Get-UserClasses).Contains($name)) { return $true }
    if (-not $tClass.ContainsKey($name)) { return $false }
    $info = $srcClass[$name]
    if (-not $info) { return $false }
    switch ($info.Reason) {
        'Aux' {
            if (-not (Connect-Class $info.Related)) { return $false }
            $parent = $tClass[$info.Related]
            $ok = Invoke-Step $parent.DN "Add auxiliary class '$name'" {
                Set-ADObject -Server $server -Identity $parent.DN -Add @{ auxiliaryClass = $name } -ErrorAction Stop
            }
            if ($ok) { $parent.Aux.Add($name); Add-Report $name 'Class' 'Attached' "Auxiliary class of '$($info.Related)'" }
            return $ok
        }
        'Super' { return (Connect-Class $info.Related) }
        default { return $false }
    }
}

foreach ($a in $needAttach) {
    if ((Get-UserAttributes).Contains($a.Name)) { continue }
    $holder = if ($a.Holder -and $tClass.ContainsKey($a.Holder)) { $a.Holder } else { 'user' }

    if ($holder -ne 'user' -and -not (Connect-Class $holder)) { $holder = 'user' }
    if ($a.Holder -and $holder -ne $a.Holder) {
        Add-Report $a.Name 'Attribute' 'Deviation' "The class '$($a.Holder)' could not be created/attached - the attribute is added directly to user"
    }

    if (-not $tClass[$holder].Attrs.Contains($a.Name)) {
        $hc = $tClass[$holder]
        $ok = Invoke-Step $hc.DN "Add '$($a.Name)' to mayContain" {
            Set-ADObject -Server $server -Identity $hc.DN -Add @{ mayContain = $a.Name } -ErrorAction Stop
        }
        if ($ok) { [void]$hc.Attrs.Add($a.Name) }
    }
}
Update-SchemaCache

# ---------- 4. Verification ----------
$final = Get-UserAttributes
foreach ($a in $scope) {
    $existing = $report | Where-Object { $_.Name -eq $a.Name -and $_.Type -eq 'Attribute' }
    if ($final.Contains($a.Name)) {
        if (-not $existing) { Add-Report $a.Name 'Attribute' 'Already existed' $a.Category }
    } elseif (-not $existing -or $existing.Status -in 'Created', 'WhatIf') {
        Add-Report $a.Name 'Attribute' 'Not on user' 'The attribute exists but is not available on user objects'
    }
}

$report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding $csvEncoding -Delimiter ';' -WhatIf:$false

Write-Host "`n--- Summary ---" -ForegroundColor Cyan
$report | Group-Object Type, Status | Sort-Object Name | ForEach-Object { "  {0,-28} {1}" -f $_.Name, $_.Count }
$okCount = @($scope | Where-Object { $final.Contains($_.Name) }).Count
Write-Host "`nAvailable on user: $okCount of $($scope.Count)" -ForegroundColor $(if ($okCount -eq $scope.Count) { 'Green' } else { 'Yellow' })
Write-Host "Report: $ReportPath"
if ($WhatIfPreference) { Write-Host 'WhatIf - nothing has been changed.' -ForegroundColor Yellow }
