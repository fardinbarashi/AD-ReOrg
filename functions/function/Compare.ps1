<#
    Groups Compare and User Compare tabs.
    Pick two objects, load every populated attribute and show them side by side in a
    table (Attribute / A / B / Status) with a name filter and "Show only differences".
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

# Attributes that are system-managed / constructed and are left out of the comparison.
$script:CompareExclude = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($x in @(
    'DistinguishedName','CN','Name','ObjectGUID','ObjectSid','SID','SamAccountType',
    'ObjectClass','ObjectCategory','whenCreated','whenChanged','Created','Modified',
    'createTimeStamp','modifyTimeStamp','uSNCreated','uSNChanged','nTSecurityDescriptor',
    'member','memberOf','primaryGroupID','instanceType','dSCorePropagationData','sIDHistory',
    'CanonicalName','isCriticalSystemObject','systemFlags','groupType','ProtectedFromAccidentalDeletion',
    'sDRightsEffective','PropertyNames','PropertyCount','AddedProperties','RemovedProperties',
    'ModifiedProperties','Deleted','msDS-KeyCredentialLink','msExchMailboxGuid'
)) { [void]$script:CompareExclude.Add($x) }

function Convert-AttrToText {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [byte[]]) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable]) {
        $parts = @()
        foreach ($x in $Value) { if ($x -is [byte[]]) { continue }; $s = [string]$x; if ($s) { $parts += $s } }
        return ($parts -join "`r`n")
    }
    return [string]$Value
}

function Get-AttrValueText {
    param($Object, [string]$Name)
    if (-not $Object) { return '' }
    $p = $Object.PSObject.Properties[$Name]
    if (-not $p) { return '' }
    return (Convert-AttrToText $p.Value)
}

function Get-CompareAttributeNames {
    # Names of populated, editable-looking attributes on either object.
    param($ObjectA, $ObjectB)
    $set = New-Object 'System.Collections.Generic.SortedSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($obj in @($ObjectA, $ObjectB)) {
        if (-not $obj) { continue }
        foreach ($pr in $obj.PSObject.Properties) {
            $n = $pr.Name
            if ($script:CompareExclude.Contains($n)) { continue }
            $txt = Convert-AttrToText $pr.Value
            if ([string]::IsNullOrEmpty($txt)) { continue }
            [void]$set.Add($n)
        }
    }
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($n in $set) { $list.Add($n) }
    return @($list.ToArray())
}

# ---------- Compare functions (groups and users) ----------

function Read-CompareObject {
    param([ValidateSet('Group','User')] [string]$Kind, [string]$DN)
    $server = $null
    if ($script:Config -and $script:Config.domain.server) { $server = [string]$script:Config.domain.server }
    $p = @{ Identity = $DN; Properties = '*'; ErrorAction = 'Stop' }
    if ($server) { $p.Server = $server }
    if ($Kind -eq 'User') { return Get-ADUser @p } else { return Get-ADGroup @p }
}

# ===================== SHARED COMPARISON TABLE =====================
# Both compare tabs show a read-only table: one row per attribute, columns
# Attribute / A / B / Status (identical or differs). Bound as a DataTable so WPF
# displays and filters it reliably. $Kind = 'Group' or 'User'.

function ConvertTo-CmpNorm {
    # Normalises a value for comparison (line endings, trailing whitespace).
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return (($Text -replace "`r`n", "`n").TrimEnd())
}

function Get-CmpContext {
    # The controls and loaded objects that belong to one compare tab.
    param([ValidateSet('Group','User')] [string]$Kind)
    if ($Kind -eq 'User') {
        return @{ A = $script:CompareUserA; B = $script:CompareUserB; Grid = $script:GridUserCmp; Result = $script:TxtUserResult
                  Filter = $script:TxtUserAttrFilter; OnlyDiff = $script:ChkUserOnlyDiff }
    }
    return @{ A = $script:CompareGroupA; B = $script:CompareGroupB; Grid = $script:GridGroupCmp; Result = $script:TxtCompareResult
              Filter = $script:TxtCompareAttrFilter; OnlyDiff = $script:ChkCompareOnlyDiff }
}

function Get-CmpTable {
    # The leading comma keeps PowerShell from unrolling the DataTable into its rows.
    param([string]$Kind)
    if ($script:CmpTables -and $script:CmpTables[$Kind]) { return ,$script:CmpTables[$Kind] }
    return $null
}

function Build-CmpTable {
    # Fills the comparison table from the two loaded objects and remembers the
    # original values so edits (and what to save) can be detected.
    param([ValidateSet('Group','User')] [string]$Kind)
    $c = Get-CmpContext -Kind $Kind
    if ($null -eq $script:CmpTables) { $script:CmpTables = @{} }

    $dt = New-Object System.Data.DataTable
    foreach ($col in 'Attribute','ValueA','ValueB','Status') { [void]$dt.Columns.Add($col, [string]) }
    foreach ($name in @(Get-CompareAttributeNames -ObjectA $c.A -ObjectB $c.B)) {
        $va = Get-AttrValueText -Object $c.A -Name $name
        $vb = Get-AttrValueText -Object $c.B -Name $name
        $row = $dt.NewRow()
        $row['Attribute'] = $name; $row['ValueA'] = $va; $row['ValueB'] = $vb
        $row['Status'] = if ((ConvertTo-CmpNorm $va) -eq (ConvertTo-CmpNorm $vb)) { 'identical' } else { 'differs' }
        [void]$dt.Rows.Add($row)
    }
    $script:CmpTables[$Kind] = $dt
    if ($c.Grid) { $c.Grid.ItemsSource = $dt.DefaultView }
    Update-CmpFilter -Kind $Kind
    Update-CmpSummary -Kind $Kind
}

function Update-CmpSummary {
    param([ValidateSet('Group','User')] [string]$Kind)
    $c = Get-CmpContext -Kind $Kind; $dt = Get-CmpTable -Kind $Kind
    if (-not $c.Result -or -not $dt) { return }
    $rows = @($dt.Rows)
    $diff = @($rows | Where-Object { $_['Status'] -eq 'differs' }).Count
    $c.Result.Text = "{0} attributes  ·  {1} differ  ·  {2} identical" -f $rows.Count, $diff, ($rows.Count - $diff)
    $c.Result.Foreground = if ($diff -gt 0) { New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(180,35,42)) }
                           else { New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(27,122,61)) }
}

function Update-CmpFilter {
    # Applies the attribute-name filter and "Show only differences" to the DataView.
    param([ValidateSet('Group','User')] [string]$Kind)
    $c = Get-CmpContext -Kind $Kind; $dt = Get-CmpTable -Kind $Kind
    if (-not $dt) { return }
    $parts = @()
    if ($c.OnlyDiff -and $c.OnlyDiff.IsChecked) { $parts += "Status <> 'identical'" }
    $q = if ($c.Filter) { ([string]$c.Filter.Text).Trim() } else { '' }
    if ($q) {
        $esc = $q.Replace("'", "''").Replace('[', '[[]').Replace('%', '[%]').Replace('*', '[*]')
        $parts += "Attribute LIKE '%$esc%'"
    }
    try { $dt.DefaultView.RowFilter = ($parts -join ' AND ') } catch { $dt.DefaultView.RowFilter = '' }
}

function Export-CompareReport {
    # Exports the comparison shown in the grid (after the attribute filter and
    # "Only differences") to HTML / CSV / JSON in files\compare-groups or
    # files\compare-users.
    param([ValidateSet('Group','User')] [string]$Kind, [ValidateSet('All','Html','Csv','Json')] [string]$Format = 'All')
    $c = Get-CmpContext -Kind $Kind; $dt = Get-CmpTable -Kind $Kind
    $noun = if ($Kind -eq 'User') { 'users' } else { 'groups' }
    if (-not $dt -or -not $c.A -or -not $c.B) {
        Show-Message -Text ("Compare two {0} first: check exactly two in the tree and click Compare selected." -f $noun) -Icon Warning -Caption 'Export comparison'; return
    }
    $nameA = [string]$c.A.Name; $nameB = [string]$c.B.Name
    $records = @(foreach ($rv in $dt.DefaultView) {
        [pscustomobject]@{ Attribute = [string]$rv['Attribute']; ValueA = [string]$rv['ValueA']; ValueB = [string]$rv['ValueB']; Status = [string]$rv['Status'] }
    })
    if ($records.Count -eq 0) { Show-Message -Text 'No rows to export - the filter hides every row.' -Icon Warning -Caption 'Export comparison'; return }

    $cols = @(
        @{ Key = 'Status';    Label = 'Status' }
        @{ Key = 'Attribute'; Label = 'Attribute' }
        @{ Key = 'ValueA';    Label = ('A: {0}' -f $nameA) }
        @{ Key = 'ValueB';    Label = ('B: {0}' -f $nameB) }
    )
    $diff = @($records | Where-Object { $_.Status -eq 'differs' }).Count
    $settings = [ordered]@{}
    $settings['A'] = ('{0}  ({1})' -f $nameA, [string]$c.A.DistinguishedName)
    $settings['B'] = ('{0}  ({1})' -f $nameB, [string]$c.B.DistinguishedName)
    $settings['Rows'] = ('{0} exported - {1} differ, {2} identical' -f $records.Count, $diff, ($records.Count - $diff))
    $flt = if ($c.Filter) { ([string]$c.Filter.Text).Trim() } else { '' }
    $settings['Filter'] = ('Attribute name: {0}; only differences: {1}' -f $(if ($flt) { '"' + $flt + '"' } else { '(none)' }), $(if ($c.OnlyDiff -and $c.OnlyDiff.IsChecked) { 'Yes' } else { 'No' }))

    $folder = Get-ReportFolder -Sub ('compare-{0}' -f $noun)
    $fileName = { param($t) ($t -replace '[\\/:*?"<>|\s]+', '_').Trim('_') }
    $base = Join-Path $folder ('compare_{0}_{1}_vs_{2}_{3}' -f $noun, (& $fileName $nameA), (& $fileName $nameB), (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $made = @()
    $title = ('Compare {0} - {1} vs {2}' -f $noun, $nameA, $nameB)
    if ($Format -in 'All','Html') { $hp = $base + '.html'; New-HtmlReport -Title $title -Columns $cols -Records $records -Path $hp -StatusKey 'Status' -Unit 'attributes' -Settings $settings -HideChangedBox; $made += $hp }
    if ($Format -in 'All','Csv')  { $cp = $base + '.csv';  Export-DataCsv -Columns $cols -Records $records -Path $cp; $made += $cp }
    if ($Format -in 'All','Json') { $jp = $base + '.json'; Export-DataJson -Records $records -Path $jp -Meta @{ type = ('compare-{0}' -f $noun); a = $nameA; b = $nameB; aDN = [string]$c.A.DistinguishedName; bDN = [string]$c.B.DistinguishedName; differ = $diff }; $made += $jp }
    foreach ($m in $made) { Write-Log ('Comparison exported: {0}' -f $m) 'OK' }
    Set-Status ('Comparison exported ({0} attributes).' -f $records.Count)
    $html = @($made | Where-Object { $_ -like '*.html' }) | Select-Object -First 1
    if ($script:Config -and $script:Config.report.openAfter) { if ($html) { Start-Process $html } else { Start-Process (Split-Path $made[0] -Parent) } }
}

# ===================== GROUPS COMPARE =====================

function Update-CompareSelectedCount {
    if (-not $script:CompareTree) { return }
    $n = @(Get-CheckedGroups $script:CompareTree).Count
    if ($script:TxtCompareCount) { $script:TxtCompareCount.Text = "Selected groups: $n (pick 2 to compare)" }
}
function Update-CompareTreeFilter {
    Update-CheckGroupTreeFilter -Tree $script:CompareTree `
        -AttrBoxes @($script:TxtCmpAttr1, $script:TxtCmpAttr2, $script:TxtCmpAttr3) `
        -ValBoxes  @($script:TxtCmpVal1,  $script:TxtCmpVal2,  $script:TxtCmpVal3) -Kind 'Group'
    Update-CompareSelectedCount
}

function Invoke-CompareLoad {
    $nodes = @(Get-CheckedGroups $script:CompareTree)
    if ($nodes.Count -ne 2) { Show-Message -Text ("Check the check box of exactly TWO groups in the tree to compare.`n`nYou have checked {0}." -f $nodes.Count) -Icon Warning; return }
    Set-Status 'Reading group attributes ...' -Busy
    try { $a = Read-CompareObject -Kind Group -DN $nodes[0].DN; $b = Read-CompareObject -Kind Group -DN $nodes[1].DN }
    catch { Set-Status 'Error.'; Write-Log "Could not read groups: $($_.Exception.Message)" 'ERROR'; return }

    $script:CompareGroupA = $a; $script:CompareGroupB = $b
    $script:TxtCompareAName.Text = ('{0}  ({1})' -f $a.Name, $a.SamAccountName)
    $script:TxtCompareBName.Text = ('{0}  ({1})' -f $b.Name, $b.SamAccountName)

    Build-CmpTable -Kind Group
    $n = (Get-CmpTable -Kind Group).Rows.Count
    Set-Status ("Comparing '{0}' and '{1}' - {2} attributes." -f $a.Name, $b.Name, $n)
    Write-Log ("Compare groups: '{0}' vs '{1}' ({2} attributes)." -f $a.Name, $b.Name, $n) 'INFO'
}

function Update-GroupCmpFilter   { Update-CmpFilter   -Kind Group }

# ===================== USER COMPARE =====================

function Build-UserCompareTree {
    # OU -> user leaves (with checkboxes). Built from the users read at connect time.
    param([Parameter(Mandatory)] [string]$RootDN, [Parameter(Mandatory)] [string]$RootName, [object[]]$Users)
    if (-not $script:UserTree) { return }
    $script:UserTree.Items.Clear()
    $rootLower = $RootDN.ToLower()
    $rootNode = New-GroupCheckNode -Type 'OU' -DN $RootDN -Label $RootName -TreeKind 'UserCompare'
    [void]$script:UserTree.Items.Add($rootNode)
    $map = @{}; $map[$rootLower] = $rootNode

    function Get-OrCreateUserOu($dn) {
        $dnl = $dn.ToLower()
        if ($dnl -eq $rootLower -or -not $dnl.EndsWith($rootLower)) { return $rootNode }
        if ($map.ContainsKey($dnl)) { return $map[$dnl] }
        $parent = ($dn -split ',', 2)[1]
        $pNode  = Get-OrCreateUserOu $parent
        $label  = ($dn -split ',', 2)[0] -replace '^OU=', '' -replace '^CN=', ''
        $n = New-GroupCheckNode -Type 'OU' -DN $dn -Label $label -TreeKind 'UserCompare'
        [void]$pNode.Items.Add($n); $map[$dnl] = $n
        return $n
    }

    foreach ($u in ($Users | Sort-Object Surname, GivenName, SamAccountName)) {
        $parentDn = ($u.DistinguishedName -split ',', 2)[1]
        $ouNode = Get-OrCreateUserOu $parentDn
        $nm = (@($u.GivenName, $u.Surname) | Where-Object { $_ }) -join ' '; if (-not $nm) { $nm = $u.Name }
        $leaf = New-GroupCheckNode -Type 'GroupLeaf' -DN $u.DistinguishedName -Label ('{0}  ({1})  |  {2}' -f $nm, $u.SamAccountName, $u.UserPrincipalName) -TreeKind 'UserCompare'
        $leaf.Tag.Name = $nm; $leaf.Tag.Sam = $u.SamAccountName
        [void]$ouNode.Items.Add($leaf)
        $p = $ouNode
        while ($p -and $p.Tag -and $p.Tag.Type -eq 'OU') { $p.Tag.Count++; $p = $p.Parent }
    }
    foreach ($item in (Get-TreeItem -Parent $script:UserTree)) {
        if ($item.Tag.Type -eq 'OU') {
            $base = if ($item -eq $rootNode) { $RootName } else { ($item.Tag.DN -split ',', 2)[0] -replace '^OU=', '' }
            $item.Tag.Text.Text = '{0}  ({1})' -f $base, $item.Tag.Count
        }
    }
    $rootNode.IsExpanded = $true
    Update-UserCompareCount
}

function Update-UserCompareCount {
    if (-not $script:UserTree) { return }
    $n = @(Get-CheckedGroups $script:UserTree).Count
    if ($script:TxtUserCompareCount) { $script:TxtUserCompareCount.Text = "Selected users: $n (pick 2 to compare)" }
}
function Update-UserTreeFilter {
    Update-CheckGroupTreeFilter -Tree $script:UserTree -FilterBox $script:TxtUserFilter `
        -AttrBoxes @($script:TxtUsrAttr1, $script:TxtUsrAttr2, $script:TxtUsrAttr3) `
        -ValBoxes  @($script:TxtUsrVal1,  $script:TxtUsrVal2,  $script:TxtUsrVal3) -Kind 'User'
    Update-UserCompareCount
}

function Invoke-UserCompareLoad {
    $nodes = @(Get-CheckedGroups $script:UserTree)
    if ($nodes.Count -ne 2) { Show-Message -Text ("Check the check box of exactly TWO users in the tree to compare.`n`nYou have checked {0}." -f $nodes.Count) -Icon Warning; return }
    Set-Status 'Reading user attributes ...' -Busy
    try { $a = Read-CompareObject -Kind User -DN $nodes[0].DN; $b = Read-CompareObject -Kind User -DN $nodes[1].DN }
    catch { Set-Status 'Error.'; Write-Log "Could not read users: $($_.Exception.Message)" 'ERROR'; return }

    $script:CompareUserA = $a; $script:CompareUserB = $b
    $script:TxtUserAName.Text = ('{0}  ({1})' -f $a.Name, $a.SamAccountName)
    $script:TxtUserBName.Text = ('{0}  ({1})' -f $b.Name, $b.SamAccountName)

    Build-CmpTable -Kind User
    $n = (Get-CmpTable -Kind User).Rows.Count
    Set-Status ("Comparing '{0}' and '{1}' - {2} attributes." -f $a.Name, $b.Name, $n)
    Write-Log ("Compare users: '{0}' vs '{1}' ({2} attributes)." -f $a.Name, $b.Name, $n) 'INFO'
}

function Update-UserCmpFilter   { Update-CmpFilter   -Kind User }
