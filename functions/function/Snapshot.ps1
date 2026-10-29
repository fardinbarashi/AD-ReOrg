<#
    Snapshot users tab:
      1. Read users under the root and show them in a nested tree
         (OU -> Department -> Office -> user) with checkboxes. Export the checked
         users to HTML, CSV and JSON (including group membership).
      2. Compare two snapshot files (CSV or JSON, old/new): every field plus group
         membership is diffed per person, with optional name patterns to pair renamed
         groups, and an SVG connection diagram.
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

function Get-UserOuPath {
    param([string]$DN)
    if (-not $DN) { return '' }
    $parent = ($DN -split ',', 2)[1]
    if (-not $parent) { return '' }
    $ous = @()
    foreach ($part in ($parent -split ',')) { if ($part -match '^OU=(.+)$') { $ous += $Matches[1] } }
    [array]::Reverse($ous)
    return ($ous -join '/')
}
function Get-CnFromDn { param([string]$DN); if (-not $DN) { return '' }; return ((($DN -split ',', 2)[0]) -replace '^CN=', '') }

function Convert-ByPattern {
    # Applies the compare "find -> replace" pairs to a name (case-insensitive, literal),
    # so a renamed group/folder (e.g. OLD -> NEW) normalises to a common form and is
    # matched instead of showing up as one lost + one gained.
    param([string]$Value, $Pairs)
    $v = [string]$Value
    foreach ($p in @($Pairs)) {
        if ($p -and $p.Find) { $v = $v -replace [regex]::Escape([string]$p.Find), ([string]$p.Replace) }
    }
    return $v
}

function Get-DiffPatternPairs {
    # Reads the 3 find/replace pattern boxes on the Compare files tab.
    $pairs = @()
    foreach ($pair in @(@($script:TxtDiffFind1,$script:TxtDiffRepl1), @($script:TxtDiffFind2,$script:TxtDiffRepl2), @($script:TxtDiffFind3,$script:TxtDiffRepl3))) {
        if ($pair[0]) { $f = ([string]$pair[0].Text).Trim(); if ($f) { $r = if ($pair[1]) { [string]$pair[1].Text } else { '' }; $pairs += @{ Find = $f; Replace = $r } } }
    }
    return @($pairs)
}

# ---------- Load + tree ----------

function Invoke-SnapLoad {
    # Reads every user under the root and builds the snapshot tree. Runs as part of
    # "Connect & load AD" (-Quiet keeps the connect subtitle/status intact).
    param([switch]$Quiet)
    $root = ([string]$script:TxtRoot.Text).Trim()
    if ([string]::IsNullOrWhiteSpace($root)) { Show-Message -Text 'Enter a root (SearchBase) at the top first.' -Icon Warning; return }
    $server = $null
    if ($script:Config -and $script:Config.domain.server) { $server = [string]$script:Config.domain.server }

    Set-Status 'Reading users ...' -Busy
    Write-Log ("Snapshot: reading users under {0}" -f $root) 'INFO'
    $users = @()
    try {
        $p = @{ SearchBase = $root; SearchScope = 'Subtree'; Filter = '*'; Properties = 'GivenName','Surname','SamAccountName','UserPrincipalName','Department','physicalDeliveryOfficeName','Enabled','MemberOf','DistinguishedName'; ErrorAction = 'Stop' }
        if ($server) { $p.Server = $server }
        $users = @(Get-ADUser @p)
    }
    catch { Set-Status 'Could not read users.'; Write-Log ("Snapshot: could not read users: {0}" -f $_.Exception.Message) 'ERROR'; Show-Message -Text ("Could not read users:`n{0}" -f $_.Exception.Message) -Icon Error; return }

    $recs = New-Object System.Collections.Generic.List[object]
    $total = $users.Count; $i = 0
    $script:Pb.Visibility = [System.Windows.Visibility]::Visible
    $script:Pb.Maximum = [Math]::Max($total, 1); $script:Pb.Value = 0
    foreach ($u in ($users | Sort-Object Surname, GivenName)) {
        $i++; $script:Pb.Value = $i
        if ($i % 100 -eq 0 -or $i -eq $total) { Set-Status ("Processing {0} of {1} ..." -f $i, $total) -Busy }
        $groups = @()
        if ($u.MemberOf) { $groups = @($u.MemberOf | ForEach-Object { Get-CnFromDn $_ } | Where-Object { $_ } | Sort-Object) }
        $name = (@($u.GivenName, $u.Surname) | Where-Object { $_ }) -join ' '; if (-not $name) { $name = '(no name)' }
        $recs.Add([pscustomobject]@{
            Account    = $u.SamAccountName
            Name       = $name
            Surname    = [string]$u.Surname
            GivenName  = [string]$u.GivenName
            UPN        = $u.UserPrincipalName
            Enabled    = if ($u.Enabled) { 'Yes' } else { 'No' }
            OU         = Get-UserOuPath $u.DistinguishedName
            Department = if ($u.Department) { [string]$u.Department } else { '' }
            Office     = if ($u.physicalDeliveryOfficeName) { [string]$u.physicalDeliveryOfficeName } else { '' }
            GroupCount = $groups.Count
            Groups     = ($groups -join ' | ')
        })
    }
    $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed
    $script:SnapUsers = @($recs.ToArray())
    Build-SnapTree
    $act = @($script:SnapUsers | Where-Object { $_.Enabled -eq 'Yes' }).Count
    if (-not $Quiet) {
        Set-Status ("Snapshot: {0} users loaded." -f $script:SnapUsers.Count)
    }
    Write-Log ("Snapshot: {0} users loaded ({1} enabled)." -f $script:SnapUsers.Count, $act) 'OK'
}

function New-SnapCheckNode {
    param([Parameter(Mandatory)] [ValidateSet('SnapGroup','SnapUser')] [string]$Type, [Parameter(Mandatory)] [string]$Label, [object]$User)
    $item  = New-Object System.Windows.Controls.TreeViewItem
    $panel = New-Object System.Windows.Controls.StackPanel; $panel.Orientation = 'Horizontal'
    $cb = New-Object System.Windows.Controls.CheckBox; $cb.VerticalAlignment = 'Center'; $cb.Margin = '0,0,6,0'
    $text = New-Object System.Windows.Controls.TextBlock; $text.VerticalAlignment = 'Center'; $text.Text = $Label
    if ($Type -eq 'SnapGroup') { $text.FontWeight = 'SemiBold' }
    [void]$panel.Children.Add($cb); [void]$panel.Children.Add($text)
    $item.Header = $panel
    $item.Tag = @{ Type = $Type; CheckBox = $cb; Text = $text; User = $User }
    $script:SnapCheckMap[$cb] = $item
    $cb.Add_Click({
        param($s, $e)
        $owner = $script:SnapCheckMap[$s]
        if ($null -ne $owner) { Set-NodeCheckState -Item $owner -IsChecked ([bool]$s.IsChecked); Update-SnapSelectedCount }
        if ($e) { $e.Handled = $true }
    })
    return $item
}

function Add-SnapLevel {
    param([object]$ParentNode, [object[]]$Users, [string[]]$Levels, [int]$LevelIndex)
    if ($LevelIndex -ge $Levels.Count) {
        foreach ($u in ($Users | Sort-Object Surname, GivenName, Account)) {
            $lbl = '{0}  |  {1}  |  {2}' -f $u.Account, $u.Name, $u.UPN
            $un = New-SnapCheckNode -Type 'SnapUser' -Label $lbl -User $u
            if ($ParentNode) { [void]$ParentNode.Items.Add($un) } else { [void]$script:SnapTree.Items.Add($un) }
        }
        return
    }
    $lv = $Levels[$LevelIndex]
    $grouped = $Users | Group-Object {
        $u = $_
        $val = switch ($lv) { 'OU' { $u.OU } 'Department' { $u.Department } 'Office' { $u.Office } default { '' } }
        if ($val) { $val } else { switch ($lv) { 'OU' { '(no OU)' } 'Department' { '(no department)' } 'Office' { '(no office)' } default { '(other)' } } }
    } | Sort-Object Name
    foreach ($g in $grouped) {
        $node = New-SnapCheckNode -Type 'SnapGroup' -Label ('{0}  ({1})' -f $g.Name, $g.Count)
        if ($ParentNode) { [void]$ParentNode.Items.Add($node) } else { [void]$script:SnapTree.Items.Add($node) }
        Add-SnapLevel -ParentNode $node -Users $g.Group -Levels $Levels -LevelIndex ($LevelIndex + 1)
    }
}

function Add-SnapOuTree {
    # Builds a real nested OU hierarchy (OU -> sub-OU -> sub-OU) from each user's
    # OU path, and places the users under their deepest OU. Each OU node shows the
    # number of users beneath it (including sub-OUs).
    param([object[]]$Users)
    $nodeByPath  = @{}
    $countByPath = @{}

    function Get-OrCreateSnapOu {
        param([string[]]$Segs)
        $key = ($Segs -join '/')
        if ($nodeByPath.ContainsKey($key)) { return $nodeByPath[$key] }
        $node = New-SnapCheckNode -Type 'SnapGroup' -Label $Segs[-1]
        if ($Segs.Count -le 1) { [void]$script:SnapTree.Items.Add($node) }
        else { $parent = Get-OrCreateSnapOu -Segs @($Segs[0..($Segs.Count - 2)]); [void]$parent.Items.Add($node) }
        $nodeByPath[$key] = $node
        return $node
    }

    $byOu = $Users | Group-Object { if ($_.OU) { [string]$_.OU } else { '(no OU)' } }
    foreach ($g in ($byOu | Sort-Object Name)) {
        $segs = @([string]$g.Name -split '/' | Where-Object { $_ -ne '' })
        if ($segs.Count -eq 0) { $segs = @('(no OU)') }
        $leaf = Get-OrCreateSnapOu -Segs $segs
        foreach ($u in ($g.Group | Sort-Object Surname, GivenName, Account)) {
            $lbl = '{0}  |  {1}  |  {2}' -f $u.Account, $u.Name, $u.UPN
            [void]$leaf.Items.Add((New-SnapCheckNode -Type 'SnapUser' -Label $lbl -User $u))
        }
        # count this OU and every ancestor
        for ($i = 1; $i -le $segs.Count; $i++) {
            $pk = (@($segs[0..($i - 1)]) -join '/')
            $countByPath[$pk] = [int]$countByPath[$pk] + $g.Count
        }
    }
    foreach ($pk in $nodeByPath.Keys) {
        $seg = ($pk -split '/')[-1]
        $nodeByPath[$pk].Tag.Text.Text = '{0}  ({1})' -f $seg, ([int]$countByPath[$pk])
    }
    foreach ($it in $script:SnapTree.Items) { if ($it -is [System.Windows.Controls.TreeViewItem]) { $it.IsExpanded = $true } }
}

function Build-SnapTree {
    if (-not $script:SnapTree -or $null -eq $script:SnapUsers) { return }
    $script:SnapCheckMap = @{}
    $script:SnapTree.Items.Clear()
    $q = ''
    if ($script:TxtSnapFilter) { $q = ([string]$script:TxtSnapFilter.Text).Trim().ToLower() }
    $groupBy = 'OU'
    if ($script:CmbSnapGroupBy -and $script:CmbSnapGroupBy.SelectedItem) { $groupBy = [string]$script:CmbSnapGroupBy.SelectedItem.Content }
    $view = $script:SnapUsers
    if ($q) { $view = @($view | Where-Object { ("{0} {1} {2} {3} {4} {5}" -f $_.Account, $_.Name, $_.OU, $_.Department, $_.Office, $_.Groups).ToLower().Contains($q) }) }
    # AD attribute filter (up to three attribute/value pairs, AND, contains).
    $pairs = Get-PairsFromBoxes -AttrBoxes @($script:TxtSnapAttr1, $script:TxtSnapAttr2, $script:TxtSnapAttr3) -ValBoxes @($script:TxtSnapVal1, $script:TxtSnapVal2, $script:TxtSnapVal3)
    if ($pairs.Count -gt 0) {
        foreach ($pr in $pairs) { Read-UserAttributeValues -Attr $pr.Attr }
        $cache = $script:UserAttrFilterCache
        $view = @($view | Where-Object { Test-AttrPairs -Pairs $pairs -Cache $cache -Key ([string]$_.Account).ToLower() })
    }
    if ($groupBy -eq 'OU') {
        Add-SnapOuTree -Users @($view)
    } else {
        $levels = switch ($groupBy) { 'Department' { @('Department','Office') } 'Office' { @('Office') } default { @() } }
        Add-SnapLevel -ParentNode $null -Users @($view) -Levels $levels -LevelIndex 0
    }
    Update-SnapSelectedCount
}

function Update-SnapSelectedCount {
    if (-not $script:SnapTree) { return }
    $n = @(Get-CheckedSnapUsers).Count
    $tot = if ($script:SnapUsers) { $script:SnapUsers.Count } else { 0 }
    $groupBy = 'OU'
    if ($script:CmbSnapGroupBy -and $script:CmbSnapGroupBy.SelectedItem) { $groupBy = [string]$script:CmbSnapGroupBy.SelectedItem.Content }
    if ($script:TxtSnapCount) { $script:TxtSnapCount.Text = ("Selected: {0} of {1} users  •  grouped by {2}" -f $n, $tot, $groupBy) }
}

function Get-CheckedSnapUsers {
    $list = New-Object System.Collections.Generic.List[object]
    if (-not $script:SnapTree) { return @($list.ToArray()) }
    foreach ($item in (Get-TreeItem -Parent $script:SnapTree)) {
        $m = $item.Tag
        if ($m.Type -eq 'SnapUser' -and $m.CheckBox -and $m.CheckBox.IsChecked -eq $true) { $list.Add($m.User) }
    }
    return @($list.ToArray())
}

function Export-SnapshotReport {
    param([ValidateSet('Html','Csv','Json','All')] [string]$Kind = 'All')
    if ($null -eq $script:SnapUsers -or @($script:SnapUsers).Count -eq 0) { Show-Message -Text 'Load users first (the "Load users" button).' -Icon Warning; return }
    $records = @(Get-CheckedSnapUsers | Sort-Object OU, Department, Office, Surname, GivenName, Account)
    if ($records.Count -eq 0) { Show-Message -Text 'Check the check box of at least one user (or a group/OU) in the tree first.' -Icon Warning; return }
    $groupBy = 'OU'
    if ($script:CmbSnapGroupBy -and $script:CmbSnapGroupBy.SelectedItem) { $groupBy = [string]$script:CmbSnapGroupBy.SelectedItem.Content }

    $cols = @(
        @{ Key='Account';    Label='Account' }
        @{ Key='Name';       Label='Name' }
        @{ Key='UPN';        Label='UPN' }
        @{ Key='Enabled';    Label='Enabled' }
        @{ Key='OU';         Label='OU' }
        @{ Key='Department'; Label='Department' }
        @{ Key='Office';     Label='Office' }
        @{ Key='GroupCount'; Label='Group count' }
        @{ Key='Groups';     Label='Groups' }
    )
    $folder = Get-ReportFolder -Sub 'snapshot-users'; $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'; $base = Join-Path $folder ('snapshot_{0}' -f $stamp); $made = @()
    Set-Status 'Creating snapshot report ...' -Busy
    if ($Kind -eq 'Html' -or $Kind -eq 'All') { $hp = $base + '.html'; New-HtmlReport -Title ('Snapshot users - grouped by {0}' -f $groupBy) -Columns $cols -Records $records -Path $hp -Unit 'users'; $made += $hp; Write-Log ("Snapshot HTML: {0}" -f $hp) 'OK' }
    if ($Kind -eq 'Csv'  -or $Kind -eq 'All') { $cp = $base + '.csv';  Export-DataCsv -Columns $cols -Records $records -Path $cp; $made += $cp; Write-Log ("Snapshot CSV: {0}" -f $cp) 'OK' }
    if ($Kind -eq 'Json' -or $Kind -eq 'All') { $jp = $base + '.json'; Export-DataJson -Records $records -Path $jp -Meta @{ type='snapshot'; groupedBy=$groupBy; root=([string]$script:TxtRoot.Text).Trim() }; $made += $jp; Write-Log ("Snapshot JSON: {0}" -f $jp) 'OK' }
    Set-Status ("Snapshot report created ({0} users)." -f $records.Count)
    $htmlMade = @($made | Where-Object { $_ -like '*.html' }) | Select-Object -First 1
    if ($htmlMade -and $script:Config.report.openAfter) { Start-Process $htmlMade }
    elseif ($made.Count -gt 0 -and $script:Config.report.openAfter) { Start-Process (Split-Path $made[0] -Parent) }
}

# ---------- Compare old/new CSV ----------

function Invoke-PickCsv {
    param([ValidateSet('Old','New')] [string]$Which)
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Snapshot exports (*.csv;*.json)|*.csv;*.json|CSV files (*.csv)|*.csv|JSON files (*.json)|*.json|All files (*.*)|*.*'
    $dlg.Title  = if ($Which -eq 'Old') { 'Select the OLD (earlier) snapshot file (CSV or JSON)' } else { 'Select the NEW (later) snapshot file (CSV or JSON)' }
    $folder = Get-ReportFolder -Sub 'snapshot-users'; if (Test-Path -LiteralPath $folder) { $dlg.InitialDirectory = $folder }
    if ($dlg.ShowDialog() -eq $true) { if ($Which -eq 'Old') { $script:TxtOldCsv.Text = $dlg.FileName } else { $script:TxtNewCsv.Text = $dlg.FileName } }
}

function Convert-SnapshotRowsToMap {
    # Turns snapshot rows (from CSV or JSON) into a map keyed by account/UPN,
    # each holding Account, Name and the group list. Shared by CSV and JSON import.
    param($Rows)
    $map = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $props = $row.PSObject.Properties.Name
        $keyCol = @('Account','SamAccountName','Sam','UPN','UserPrincipalName','User') | Where-Object { $props -contains $_ } | Select-Object -First 1
        $grpCol = @('Groups','MemberOf','Group','Membership') | Where-Object { $props -contains $_ } | Select-Object -First 1
        $namCol = @('Name','DisplayName') | Where-Object { $props -contains $_ } | Select-Object -First 1
        if (-not $keyCol) { continue }
        $key = ([string]$row.$keyCol).Trim(); if (-not $key) { continue }
        $groups = @(); if ($grpCol) { $groups = @(([string]$row.$grpCol) -split '\s*\|\s*|\s*;\s*|\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique) }
        $name = if ($namCol) { [string]$row.$namCol } else { '' }
        # Keep the whole row so the comparison can compare every field, not only groups.
        $map[$key.ToLower()] = @{ Account = $key; Name = $name; Groups = $groups; Row = $row; GroupCol = $grpCol }
    }
    return $map
}

function Import-SnapshotCsv {
    param([Parameter(Mandatory)] [string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $firstLine = (Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8)
    $delim = if ($firstLine -match ';') { ';' } elseif ($firstLine -match "`t") { "`t" } else { ',' }
    $rows = Import-Csv -LiteralPath $Path -Delimiter $delim -Encoding UTF8
    return (Convert-SnapshotRowsToMap -Rows $rows)
}

function Import-SnapshotFile {
    # Reads a snapshot export - CSV or JSON (from Export-SnapshotReport) - into a map.
    param([Parameter(Mandatory)] [string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $ext = [System.IO.Path]::GetExtension($Path).ToLower()
    if ($ext -eq '.json') {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $rows = if ($null -ne $obj.data) { @($obj.data) } else { @($obj) }
        return (Convert-SnapshotRowsToMap -Rows $rows)
    }
    if ($ext -eq '.csv' -or $ext -eq '.tsv' -or $ext -eq '.txt') { return (Import-SnapshotCsv -Path $Path) }
    throw "Unsupported file type '$ext'. Use a CSV or JSON file exported in Snapshots - step 1 Export users."
}

function Get-SnapshotCompareFields {
    # Ordered list of scalar fields to compare across two snapshot maps (every field
    # found in either file, except the groups column which is compared separately).
    param([hashtable]$Old, [hashtable]$New)
    $preferred = @('Name','GivenName','Surname','UPN','Enabled','OU','Department','Office','GroupCount')
    $fields = New-Object System.Collections.Generic.List[string]
    $seenF  = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($f in $preferred) { if ($seenF.Add($f)) { $fields.Add($f) } }
    foreach ($side in @($Old, $New)) {
        foreach ($e in $side.Values) {
            $gc = $e.GroupCol
            if ($e.Row) { foreach ($pn in $e.Row.PSObject.Properties.Name) { if ($pn -ne $gc -and $seenF.Add($pn)) { $fields.Add($pn) } } }
        }
    }
    return @($fields.ToArray())
}

function Get-SnapshotDiffRecords {
    # Diffs two snapshot maps per user across EVERY field in the file (name, surname,
    # UPN, enabled, OU, department, office, ...) plus group membership, and returns
    # one record per person describing exactly what changed.
    param([hashtable]$Old, [hashtable]$New, $GroupNamePairs = @())
    $keys = @($Old.Keys + $New.Keys | Select-Object -Unique)
    $fields = Get-SnapshotCompareFields -Old $Old -New $New

    $recs = New-Object System.Collections.Generic.List[object]
    foreach ($k in $keys) {
        $inOld = $Old.ContainsKey($k); $inNew = $New.ContainsKey($k)
        $oE = if ($inOld) { $Old[$k] } else { $null }
        $nE = if ($inNew) { $New[$k] } else { $null }
        $oldGroups = if ($inOld) { @($oE.Groups) } else { @() }
        $newGroups = if ($inNew) { @($nE.Groups) } else { @() }
        $account = if ($inNew) { $nE.Account } else { $oE.Account }
        $name    = if ($inNew -and $nE.Name) { $nE.Name } elseif ($inOld) { $oE.Name } else { '' }
        # Directional pattern: the OLD group name is transformed (e.g. OLD -> NEW)
        # to find its NEW counterpart, so an old "..OLD.." pairs with a new "..NEW..".
        # New names are matched as they are.
        $newSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($g in $newGroups) { [void]$newSet.Add([string]$g) }
        $oldSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($g in $oldGroups) { [void]$oldSet.Add([string]$g) }
        $oldNormSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($g in $oldGroups) { [void]$oldNormSet.Add((Convert-ByPattern -Value $g -Pairs $GroupNamePairs)) }
        # Lost only if neither the exact name nor its transformed name exists in NEW.
        $removed = @($oldGroups | Where-Object { -not $newSet.Contains([string]$_) -and -not $newSet.Contains((Convert-ByPattern -Value $_ -Pairs $GroupNamePairs)) })
        # Gained only if it matches neither an exact OLD name nor a transformed OLD name.
        $added   = @($newGroups | Where-Object { -not $oldSet.Contains([string]$_) -and -not $oldNormSet.Contains([string]$_) })

        $changes = New-Object System.Collections.Generic.List[string]
        if ($inOld -and $inNew) {
            foreach ($f in $fields) {
                $ov = if ($oE.Row -and $oE.Row.PSObject.Properties[$f]) { ([string]$oE.Row.$f).Trim() } else { '' }
                $nv = if ($nE.Row -and $nE.Row.PSObject.Properties[$f]) { ([string]$nE.Row.$f).Trim() } else { '' }
                if ($ov -ne $nv) { $changes.Add(("{0}: '{1}' -> '{2}'" -f $f, $ov, $nv)) }
            }
            if ($removed.Count -gt 0) { $changes.Add(("Groups lost: {0}" -f ($removed -join ', '))) }
            if ($added.Count   -gt 0) { $changes.Add(("Groups gained: {0}" -f ($added -join ', '))) }
        }

        $status =
            if     (-not $inNew) { 'Removed account (only in old)' }
            elseif (-not $inOld) { 'New account (only in new)' }
            elseif ($changes.Count -gt 0) { 'Changed' }
            else                          { 'Unchanged' }

        $recs.Add([pscustomobject]@{
            Account     = $account
            Name        = $name
            Status      = $status
            ChangeCount = $changes.Count
            Changes     = ($changes -join '; ')
            OldGroups   = $oldGroups.Count
            NewGroups   = $newGroups.Count
            LostCount   = $removed.Count
            LostGroups  = ($removed -join ' | ')
            AddedGroups = ($added -join ' | ')
        })
    }
    # Changed / removed / new users first, then unchanged; most changes first.
    return @($recs.ToArray() | Sort-Object @{ Expression = { if ($_.Status -eq 'Unchanged') { 1 } else { 0 } } }, @{ Expression = 'ChangeCount'; Descending = $true }, Account)
}

function Invoke-CompareCsv {
    $oldPath = ([string]$script:TxtOldCsv.Text).Trim()
    $newPath = ([string]$script:TxtNewCsv.Text).Trim()
    if (-not $oldPath -or -not $newPath) { Show-Message -Text 'Select both an OLD and a NEW snapshot file (CSV or JSON) first.' -Icon Warning; return }
    Set-Status 'Comparing snapshot files ...' -Busy
    try { $old = Import-SnapshotFile -Path $oldPath; $new = Import-SnapshotFile -Path $newPath }
    catch { Set-Status 'Read error.'; Write-Log ("Comparison: {0}" -f $_.Exception.Message) 'ERROR'; Show-Message -Text ("Could not read the file:`n{0}`n`nUse a CSV or JSON exported in Snapshots - step 1 Export users." -f $_.Exception.Message) -Icon Error; return }

    # Group-name normalisation pattern (find -> replace, up to 3 rows).
    $pairs = Get-DiffPatternPairs
    $all = Get-SnapshotDiffRecords -Old $old -New $new -GroupNamePairs $pairs
    $script:SnapDiff = $all
    $script:SnapCmpOld = $old   # kept for the SVG connection diagram
    $script:SnapCmpNew = $new
    $script:SnapCmpPairs = $pairs
    if ($pairs.Count) { Write-Log ("Group-name pattern applied: {0}" -f (($pairs | ForEach-Object { "'$($_.Find)'->'$($_.Replace)'" }) -join ', ')) 'INFO' }
    if ($script:GridDiff) {
        $gridCols = @(@{Key='Account'},@{Key='Name'},@{Key='Status'},@{Key='Changes'},@{Key='OldGroups'},@{Key='NewGroups'},@{Key='LostGroups'})
        $script:GridDiff.ItemsSource = (ConvertTo-DataTable -Records $all -Columns $gridCols).DefaultView
    }
    $changed = @($all | Where-Object { $_.Status -ne 'Unchanged' }).Count
    if ($script:TxtDiffCount) { $script:TxtDiffCount.Text = ("{0} people compared - {1} changed" -f $all.Count, $changed) }
    Set-Status ("Comparison done: {0} people, {1} changed." -f $all.Count, $changed)
    Write-Log ("Comparison: {0} people, {1} changed (old={2}, new={3})." -f $all.Count, $changed, $old.Count, $new.Count) 'OK'
}

function Export-DiffReport {
    param([ValidateSet('Html','Csv','Json','Svg','All')] [string]$Kind = 'All')
    if (-not $script:SnapCmpOld -or -not $script:SnapCmpNew) { Show-Message -Text 'Run a comparison first (the "Compare" button).' -Icon Warning; return }
    # Re-read the pattern from the boxes and recompute, so the report/SVG always
    # reflect what is currently typed - even if Compare was not clicked again.
    $script:SnapCmpPairs = Get-DiffPatternPairs
    $script:SnapDiff = Get-SnapshotDiffRecords -Old $script:SnapCmpOld -New $script:SnapCmpNew -GroupNamePairs $script:SnapCmpPairs
    $records = @($script:SnapDiff)
    $cols = @(
        @{ Key='Account';     Label='Account' }
        @{ Key='Name';        Label='Name' }
        @{ Key='Status';      Label='Status' }
        @{ Key='Changes';     Label='What changed' }
        @{ Key='OldGroups';   Label='Groups (old)' }
        @{ Key='NewGroups';   Label='Groups (new)' }
        @{ Key='LostGroups';  Label='Lost groups' }
        @{ Key='AddedGroups'; Label='Added groups' }
    )
    $folder = Get-ReportFolder -Sub 'snapshot-compare'; $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'; $base = Join-Path $folder ('comparison_{0}' -f $stamp); $made = @()
    $changed = @($records | Where-Object { $_.Status -ne 'Unchanged' }).Count
    $banner = ('<div class="banner no">{0} of {1} people changed between old and new (name, attributes or group membership).</div>' -f $changed, $records.Count)
    Set-Status 'Creating comparison report ...' -Busy
    if ($Kind -eq 'Html' -or $Kind -eq 'All') { $hp = $base + '.html'; New-HtmlReport -Title 'Comparison old/new - what changed per user' -Columns $cols -Records $records -Path $hp -StatusKey 'Status' -Banner $banner -Unit 'people'; $made += $hp; Write-Log ("Comparison HTML: {0}" -f $hp) 'OK' }
    if ($Kind -eq 'Csv'  -or $Kind -eq 'All') { $cp = $base + '.csv';  Export-DataCsv -Columns $cols -Records $records -Path $cp; $made += $cp; Write-Log ("Comparison CSV: {0}" -f $cp) 'OK' }
    if ($Kind -eq 'Json' -or $Kind -eq 'All') { $jp = $base + '.json'; Export-DataJson -Records $records -Path $jp -Meta @{ type='comparison'; old=([string]$script:TxtOldCsv.Text).Trim(); new=([string]$script:TxtNewCsv.Text).Trim() }; $made += $jp; Write-Log ("Comparison JSON: {0}" -f $jp) 'OK' }

    # SVG connection diagram (with 'All', or on its own with 'Svg').
    $svgPath = $null
    if ($Kind -eq 'Svg' -or $Kind -eq 'All') {
        try { $svgPath = Export-DiffSvg -Path ($base + '.svg'); if ($svgPath) { $made += $svgPath; Write-Log ("Comparison SVG: {0}" -f $svgPath) 'OK' } }
        catch { Write-Log ("Could not create SVG diagram: {0}" -f $_.Exception.Message) 'ERROR' }
    }

    Set-Status ("Comparison report created ({0} people)." -f $records.Count)
    $htmlMade = @($made | Where-Object { $_ -like '*.html' }) | Select-Object -First 1
    if ($htmlMade -and $script:Config.report.openAfter) { Start-Process $htmlMade }
    elseif ($svgPath -and $script:Config.report.openAfter) { Start-Process $svgPath }
}

function Export-DiffSvg {
    # Builds an SVG connection diagram from the last comparison (old vs new).
    # One block per user that changed (plus new/removed accounts). Each field and
    # each group is a row: an OLD box on the left, a NEW box on the right, joined by
    # a dashed line - RED if that item changed, GREY if unchanged. Changed rows are
    # bold. Each category (field / groups) has its own accent colour.
    param([Parameter(Mandatory)] [string]$Path)
    if (-not $script:SnapCmpOld -or -not $script:SnapCmpNew) { Show-Message -Text 'Run a comparison first.' -Icon Warning; return $null }
    # SVG coordinates must use a dot as decimal separator. The -f operator formats
    # numbers with the current culture (sv-SE writes 3,5), which turns y="128,5" into
    # a per-glyph position list and scatters letters over the diagram - so every
    # number is formatted with the invariant culture while this function runs.
    $prevCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture
    try { return Build-DiffSvg -Path $Path }
    finally { [System.Threading.Thread]::CurrentThread.CurrentCulture = $prevCulture }
}

function Build-DiffSvg {
    param([Parameter(Mandatory)] [string]$Path)
    $Old = $script:SnapCmpOld; $New = $script:SnapCmpNew
    $fields = @(Get-SnapshotCompareFields -Old $Old -New $New)
    $keys = @($Old.Keys + $New.Keys | Select-Object -Unique)

    # category accent palette (one colour per field, cycled)
    $palette = @('#1E76BE','#2E9E7B','#B8860B','#8E44AD','#C0562B','#2C7A7A','#7A5C2E','#4A5568','#A63D57','#3B7A57')
    $catColor = @{}; $ci = 0
    foreach ($f in $fields) { $catColor[$f] = $palette[$ci % $palette.Count]; $ci++ }
    $catColor['Groups'] = '#5A67D8'
    $red  = '#C8322B'; $grey = '#9AA4B0'; $green = '#1F9D57'; $line1 = '#dfe3e8'; $text = '#1f2933'

    $esc = {
        param($s)
        ([string]$s) -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
    }
    $pairs = @($script:SnapCmpPairs)
    $norm  = { param($g) Convert-ByPattern -Value $g -Pairs $pairs }
    $trim  = { param($s,$n); $t=[string]$s; if ($t.Length -gt $n) { $t.Substring(0,[Math]::Max(1,$n-1))+'…' } else { $t } }

    # only users with a difference (changed / only-old / only-new), then build the
    # per-user rows up front so we can size the boxes to the longest value.
    $itemsByKey = [ordered]@{}
    $maxLen = 10
    foreach ($k in $keys) {
        $io = $Old.ContainsKey($k); $inw = $New.ContainsKey($k)
        $items = New-Object System.Collections.Generic.List[object]
        $any = $false
        foreach ($f in $fields) {
            $ov = if ($io  -and $Old[$k].Row -and $Old[$k].Row.PSObject.Properties[$f]) { ([string]$Old[$k].Row.$f).Trim() } else { '' }
            $nv = if ($inw -and $New[$k].Row -and $New[$k].Row.PSObject.Properties[$f]) { ([string]$New[$k].Row.$f).Trim() } else { '' }
            $ch = ($ov -ne $nv); if ($ch) { $any = $true }
            $items.Add([pscustomobject]@{ Label=$f; Cat=$f; Old=$ov; New=$nv; Changed=$ch; Line=$(if($ch){'red'}else{'grey'}) })
            if ($ov.Length -gt $maxLen) { $maxLen = $ov.Length }; if ($nv.Length -gt $maxLen) { $maxLen = $nv.Length }
        }
        # Pair groups in two passes:
        #   1) exact identical names            -> identical, hidden
        #   2) pattern pairs (old OLD -> new NEW) -> green
        # then anything left over is red (lost on the old side / gained on the new).
        $og = if ($io)  { @($Old[$k].Groups | Where-Object { $_ }) } else { @() }
        $ng = if ($inw) { @($New[$k].Groups | Where-Object { $_ }) } else { @() }
        $newExact = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($g in $ng) { [void]$newExact.Add([string]$g) }
        $usedNew = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $grp = New-Object System.Collections.Generic.List[object]
        foreach ($g in $og) { if ($newExact.Contains([string]$g)) { [void]$usedNew.Add([string]$g) } }   # identical -> hidden
        foreach ($g in $og) {
            $t = & $norm $g
            if (([string]$t -ne [string]$g) -and $newExact.Contains([string]$t) -and -not $usedNew.Contains([string]$t)) {
                $grp.Add([pscustomobject]@{ Old=$g; New=$t; Line='green' }); [void]$usedNew.Add([string]$t)
            }
        }
        foreach ($g in $og) {
            $t = & $norm $g
            $exact = $newExact.Contains([string]$g)
            $paired = (([string]$t -ne [string]$g) -and $newExact.Contains([string]$t))
            if (-not $exact -and -not $paired) { $grp.Add([pscustomobject]@{ Old=$g; New=''; Line='red' }) }
        }
        foreach ($g in $ng) { if (-not $usedNew.Contains([string]$g)) { $grp.Add([pscustomobject]@{ Old=''; New=$g; Line='red' }) } }
        foreach ($r in ($grp | Sort-Object @{Expression={$_.Line}}, @{Expression={[string]$_.Old + [string]$_.New}})) {
            $any = $true
            $items.Add([pscustomobject]@{ Label='Group'; Cat='Groups'; Old=$r.Old; New=$r.New; Changed=($r.Line -eq 'red'); Line=$r.Line })
            if (([string]$r.Old).Length -gt $maxLen) { $maxLen = ([string]$r.Old).Length }; if (([string]$r.New).Length -gt $maxLen) { $maxLen = ([string]$r.New).Length }
        }
        if (-not ($io -and $inw)) { $any = $true }
        if ($any) { $itemsByKey[$k] = $items }
    }

    # size cards to the content (wider for long values); tooltips carry the full text
    $Wc = [int][Math]::Min([Math]::Max(300, $maxLen * 7 + 26), 380)
    $charsFit = [int](($Wc - 22) / 6.6)
    $rowH = 26; $rowGap = 6; $headH = 34; $padIn = 14; $midGap = 110
    $blockW = $padIn + $Wc + $midGap + $Wc + $padIn
    $keysArr = @($itemsByKey.Keys)
    # 3 users per row keeps the whole diagram narrow enough to stay readable.
    $cols = [Math]::Min(3, [Math]::Max(1, $keysArr.Count))
    $colGap = 24; $rowGapY = 30; $startX = 20
    $svgW = $cols * $blockW + ($cols - 1) * $colGap + 2 * $startX
    $rows = New-Object System.Collections.Generic.List[string]

    function New-Box { param($x,$y,$w,$h,$fill,$stroke,$title,$val,$full,$bold)
        $fw = if ($bold) { 'bold' } else { 'normal' }
        $b = New-Object System.Text.StringBuilder
        [void]$b.Append(('<rect x="{0}" y="{1}" width="{2}" height="{3}" rx="5" fill="{4}" stroke="{5}" stroke-width="1.3"/>' -f $x,$y,$w,$h,$fill,$stroke))
        [void]$b.Append(('<rect x="{0}" y="{1}" width="4" height="{2}" rx="2" fill="{3}"/>' -f $x,$y,$h,$stroke))
        [void]$b.Append(('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="10" fill="#6b7280">{2}</text>' -f ($x+10),($y+11),$title))
        [void]$b.Append(('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="12" font-weight="{2}" fill="#1f2933">{3}<title>{4}</title></text>' -f ($x+10),($y+22),$fw,$val,$full))
        return $b.ToString()
    }

    if ($keysArr.Count -eq 0) {
        $svgW = 520
        $rows.Add(('<text x="40" y="46" font-family="Segoe UI,Arial" font-size="14" fill="#1f2933">No differences between the two files.</text>')) | Out-Null
        $svgH = 80
    }
    else {
        # per-block height + per-row height (grid: $cols per row)
        $blkH = @{}
        foreach ($k in $keysArr) { $blkH[$k] = $headH + $itemsByKey[$k].Count * ($rowH + $rowGap) + 12 }
        $rowCount = [int][Math]::Ceiling($keysArr.Count / $cols)
        $rowHeights = @(); for ($r = 0; $r -lt $rowCount; $r++) {
            $mx = 0; for ($c = 0; $c -lt $cols; $c++) { $idx = $r*$cols + $c; if ($idx -lt $keysArr.Count) { $h = $blkH[$keysArr[$idx]]; if ($h -gt $mx) { $mx = $h } } }
            $rowHeights += $mx
        }

        # --- clickable menu: plain list of names (click a name to jump) ------
        $menu = New-Object System.Collections.Generic.List[string]
        [void]$menu.Add(('<text x="{0}" y="24" font-family="Segoe UI,Arial" font-size="13" font-weight="bold" fill="#1f2933">Users (click a name to jump):</text>' -f $startX))
        $px = $startX; $py = 46; $lineH = 20
        for ($i = 0; $i -lt $keysArr.Count; $i++) {
            $k = $keysArr[$i]
            $nm = if ($New.ContainsKey($k) -and $New[$k].Name) { $New[$k].Name } elseif ($Old.ContainsKey($k) -and $Old[$k].Name) { $Old[$k].Name }
                  elseif ($New.ContainsKey($k)) { $New[$k].Account } else { $Old[$k].Account }
            $lbl = & $trim $nm 40
            $w = [int](([string]$lbl).Length * 6.8 + 22)
            if ($px + $w -gt $svgW - $startX) { $px = $startX; $py += $lineH }
            if ($px -ne $startX) { [void]$menu.Add(('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="12" fill="#c3c9d0">|</text>' -f ($px-11),$py)) }
            [void]$menu.Add(('<a href="#u_{0}"><text x="{1}" y="{2}" font-family="Segoe UI,Arial" font-size="12" fill="#1E76BE" text-decoration="underline">{3}</text></a>' -f $i,$px,$py,(& $esc $lbl)))
            $px += $w
        }
        $menuBottom = $py + 8

        # --- legend: how to read the diagram -------------------------------
        # Long sentences are word-wrapped to the diagram width (SVG has no auto wrap).
        $wrap = { param($txt, $maxChars)
            $lines = New-Object System.Collections.Generic.List[string]; $cur = ''
            foreach ($w in ([string]$txt -split ' ')) {
                if ($cur -and ($cur.Length + 1 + $w.Length) -gt $maxChars) { $lines.Add($cur); $cur = $w } else { $cur = if ($cur) { "$cur $w" } else { $w } }
            }
            if ($cur) { $lines.Add($cur) }
            return ,$lines
        }
        $legendTop = $menuBottom + 10
        $legendW   = $svgW - 2*$startX
        $maxChars  = [int](($legendW - 80) / 5.6)
        $legendBody = New-Object System.Collections.Generic.List[string]
        $ly = $legendTop + 18; $lx = $startX + 12
        [void]$legendBody.Add(('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="12" font-weight="bold" fill="{2}">How to read this diagram</text>' -f $lx,$ly,$text))
        $ly += 17
        foreach ($ln in (& $wrap 'One card per user that differs between the two snapshots. Left column = OLD snapshot, right column = NEW snapshot. Every row is one attribute or one group membership, joined by a dashed line. Bold text = the value changed. Hover a box or a label to see the full text.' $maxChars)) {
            [void]$legendBody.Add(('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="11" fill="#6b7280">{2}</text>' -f $lx,$ly,(& $esc $ln))); $ly += 14
        }
        $ly += 6
        $samples = @(
            @{ Col = $red;   Text = 'Changed - the attribute has a different value, or the group was lost (only on the OLD side) / gained (only on the NEW side).' },
            @{ Col = $green; Text = 'Renamed - the group exists on both sides under a name that matches the Pattern boxes (find -> replace), e.g. OLD -> NEW.' },
            @{ Col = $grey;  Text = 'Identical - same value in both snapshots (attributes only; identical group memberships are hidden to keep the cards short).' }
        )
        foreach ($smp in $samples) {
            $first = $true
            foreach ($ln in (& $wrap $smp.Text ($maxChars - 9))) {
                if ($first) { [void]$legendBody.Add(('<line x1="{0}" y1="{1}" x2="{2}" y2="{1}" stroke="{3}" stroke-width="2" stroke-dasharray="5 4"/>' -f $lx,($ly-4),($lx+40),$smp.Col)); $first = $false }
                [void]$legendBody.Add(('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="11" fill="{2}">{3}</text>' -f ($lx+50),$ly,$text,(& $esc $ln))); $ly += 14
            }
            $ly += 2
        }
        $legendH = ($ly - $legendTop) + 2
        $legend = New-Object System.Collections.Generic.List[string]
        [void]$legend.Add(('<rect x="{0}" y="{1}" width="{2}" height="{3}" rx="6" fill="#ffffff" stroke="#dfe3e8"/>' -f $startX,$legendTop,$legendW,$legendH))
        foreach ($r in $legendBody) { [void]$legend.Add($r) }
        $bodyTop = $legendTop + $legendH + 16

        # row top offsets
        $rowTop = @(); $acc = $bodyTop
        for ($r = 0; $r -lt $rowCount; $r++) { $rowTop += $acc; $acc += $rowHeights[$r] + $rowGapY }
        $svgH = $acc + 26

        foreach ($r in $menu)   { [void]$rows.Add($r) }
        foreach ($r in $legend) { [void]$rows.Add($r) }

        for ($i = 0; $i -lt $keysArr.Count; $i++) {
            $k = $keysArr[$i]
            $col = $i % $cols; $row = [int][Math]::Floor($i / $cols)
            $ox = $startX + $col * ($blockW + $colGap); $oy = $rowTop[$row]
            $io = $Old.ContainsKey($k); $inw = $New.ContainsKey($k)
            $acct = if ($inw) { $New[$k].Account } else { $Old[$k].Account }
            $nm   = if ($inw -and $New[$k].Name) { $New[$k].Name } elseif ($io) { $Old[$k].Name } else { '' }
            $hstat = if (-not $inw) { 'removed (only in old)' } elseif (-not $io) { 'new (only in new)' } else { 'changed' }
            $items = $itemsByKey[$k]
            $leftCardX = $ox + $padIn; $rightCardX = $ox + $padIn + $Wc + $midGap

            $blk = New-Object System.Collections.Generic.List[string]
            [void]$blk.Add(('<rect x="{0}" y="{1}" width="{2}" height="{3}" rx="6" fill="#f6f7f9" stroke="#dfe3e8"/>' -f $ox,$oy,$blockW,$blkH[$k]))
            [void]$blk.Add(('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="14" font-weight="bold" fill="#1f2933">{2}  ({3})  -  {4}</text>' -f $leftCardX,($oy+20),(& $esc $nm),(& $esc $acct),$hstat))
            [void]$blk.Add(('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="11" fill="#6b7280">OLD</text>' -f $leftCardX,($oy+34)))
            [void]$blk.Add(('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="11" fill="#6b7280">NEW</text>' -f $rightCardX,($oy+34)))
            $ry = $oy + $headH + 6
            foreach ($it in $items) {
                $ac = $catColor[$it.Cat]; if (-not $ac) { $ac = '#4A5568' }
                $lineCol = switch ($it.Line) { 'red' { $red } 'green' { $green } default { $grey } }
                $isRed = ($it.Line -eq 'red')
                $fill  = if ($isRed) { '#fbeaea' } elseif ($it.Line -eq 'green') { '#e9f6ef' } else { '#ffffff' }
                $oldFull = & $esc $it.Old; $newFull = & $esc $it.New
                $oldVal = if ($it.Old) { & $esc (& $trim $it.Old $charsFit) } else { '-' }
                $newVal = if ($it.New) { & $esc (& $trim $it.New $charsFit) } else { '-' }
                [void]$blk.Add((New-Box $leftCardX $ry $Wc $rowH $fill $ac (& $esc $it.Label) $oldVal $oldFull $isRed))
                [void]$blk.Add((New-Box $rightCardX $ry $Wc $rowH $fill $ac (& $esc $it.Label) $newVal $newFull $isRed))
                $ymid = [int]($ry + $rowH/2)
                [void]$blk.Add(('<line x1="{0}" y1="{2}" x2="{1}" y2="{2}" stroke="{3}" stroke-width="1.6" stroke-dasharray="5 4"/>' -f ($leftCardX+$Wc),$rightCardX,$ymid,$lineCol))
                $lblRaw = if ($it.Label -eq 'Group') { $(if($it.New){$it.New}else{$it.Old}) } else { $it.Label }
                $lbl = & $esc (& $trim $lblRaw 22)
                $lblW = [int](([string]$lbl).Length * 6 + 12)
                $lblX = $leftCardX + $Wc + $midGap/2
                [void]$blk.Add(('<rect x="{0}" y="{1}" width="{2}" height="15" rx="7" fill="#ffffff" stroke="{3}"/>' -f [int]($lblX-$lblW/2),($ymid-8),$lblW,$lineCol))
                [void]$blk.Add(('<text x="{0}" y="{1}" text-anchor="middle" font-family="Segoe UI,Arial" font-size="9" fill="{2}">{3}<title>{4}</title></text>' -f $lblX,($ymid+4),$lineCol,$lbl,(& $esc $lblRaw)))
                $ry += $rowH + $rowGap
            }
            [void]$rows.Add(('<g id="u_{0}">{1}</g>' -f $i, ($blk -join '')))
        }
    }

    $footer = ('<text x="{0}" y="{1}" font-family="Segoe UI,Arial" font-size="10" fill="#9AA4B0">Generated {2} by AD ReOrg - Compare Snapshots.</text>' -f $startX,($svgH-10),(Get-Date -Format 'yyyy-MM-dd HH:mm'))
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(('<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="{0}" height="{1}" viewBox="0 0 {0} {1}">' -f $svgW,($svgH+8)))
    [void]$sb.AppendLine('<rect width="100%" height="100%" fill="#ffffff"/>')
    foreach ($r in $rows) { [void]$sb.AppendLine($r) }
    [void]$sb.AppendLine($footer)
    [void]$sb.AppendLine('</svg>')
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    return $Path
}
