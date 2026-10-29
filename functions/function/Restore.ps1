<#
    Users : Restore lost groups.

    Reads a comparison file (comparison_*.csv / .json written by
    Snapshot Users : Compare Snapshots), takes every group listed in its
    "LostGroups" column and shows it as one row per user and group. The checked
    users are then added back to those groups in Active Directory - the group name
    is used exactly as it stands in the file. Nothing is ever removed and no group
    is created.

    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

function Invoke-PickComparisonFile {
    # File picker for the comparison export, starting in files\snapshot-compare.
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Comparison files (*.csv;*.json)|*.csv;*.json|CSV files (*.csv)|*.csv|JSON files (*.json)|*.json|All files (*.*)|*.*'
    $dlg.Title  = 'Select a comparison file from Compare Snapshots (comparison_*.csv or .json)'
    $folder = Get-ReportFolder -Sub 'snapshot-compare'
    if (Test-Path -LiteralPath $folder) { $dlg.InitialDirectory = $folder }
    if ($dlg.ShowDialog() -eq $true) {
        $script:TxtRestoreFile.Text = $dlg.FileName
        Invoke-RestoreLoad
    }
}

function Import-ComparisonRows {
    # Reads a comparison export - CSV or JSON - and returns its rows.
    param([Parameter(Mandatory)] [string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $ext = [System.IO.Path]::GetExtension($Path).ToLower()
    if ($ext -eq '.json') {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($null -ne $obj.data) { return @($obj.data) }
        return @($obj)
    }
    if ($ext -eq '.csv' -or $ext -eq '.tsv' -or $ext -eq '.txt') {
        $first = (Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8)
        $delim = if ($first -match ';') { ';' } elseif ($first -match "`t") { "`t" } else { ',' }
        return @(Import-Csv -LiteralPath $Path -Delimiter $delim -Encoding UTF8)
    }
    throw "Unsupported file type '$ext'. Use a comparison CSV or JSON file from Compare Snapshots."
}

function Split-GroupListText {
    # "GroupA | GroupB" (how the comparison writes LostGroups) -> one name per item.
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    return @(([string]$Text) -split '\s*\|\s*|\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Get-LostGroupsFromChanges {
    # For a file whose LostGroups column is empty or missing: the
    # "What changed" text holds "Groups lost: A, B" written by Compare Snapshots.
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $m = [regex]::Match([string]$Text, '(?i)Groups lost:\s*(.+?)(?:;\s*Groups gained:|$)')
    if (-not $m.Success) { return @() }
    return @($m.Groups[1].Value -split '\s*,\s*' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Invoke-RestoreLoad {
    # Fills the grid from the comparison file: one row per user and lost group.
    $path = if ($script:TxtRestoreFile) { ([string]$script:TxtRestoreFile.Text).Trim() } else { '' }
    if (-not $path) { Show-Message -Text 'Select a comparison file first (the ... button).' -Icon Warning -Caption 'Restore lost groups'; return }

    # The grid rows are a compiled .NET type. A .NET type can never be replaced in a
    # running PowerShell process, so a console that already held an older AD ReOrg
    # assembly can lack it - show a message instead of loading nothing.
    if (-not ('ReOrg.RestoreRow' -as [type])) {
        Write-Log 'Restore lost groups: the row type ReOrg.RestoreRow is missing in this PowerShell session.' 'ERROR'
        Show-Message -Text "This PowerShell session cannot build the grid rows (the type ReOrg.RestoreRow is not loaded).`n`nThat happens when an older version of AD ReOrg was started in the same PowerShell window: .NET keeps the first version of a type for the life of the process.`n`nClose this window and start AD-ReOrg.ps1 from a NEW PowerShell window." -Icon Error -Caption 'Restore lost groups'
        return
    }

    Set-Status 'Reading the comparison file ...' -Busy
    $rows = @()
    try { $rows = @(Import-ComparisonRows -Path $path) }
    catch {
        Set-Status 'Could not read the file.'
        Write-Log ("Restore lost groups: could not read '{0}': {1}" -f $path, $_.Exception.Message) 'ERROR'
        Show-Message -Text $_.Exception.Message -Icon Error -Caption 'Restore lost groups'
        return
    }

    # Column names are looked up across every row, not only the first, so a file
    # whose first row happens to miss a value is still understood.
    $props = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $rows) { if ($row) { foreach ($pn in $row.PSObject.Properties.Name) { [void]$props.Add($pn) } } }
    $lostCol = @('LostGroups','Lost groups','Lost') | Where-Object { $props.Contains($_) } | Select-Object -First 1
    $accCol  = @('Account','SamAccountName','Sam','UPN','UserPrincipalName','User') | Where-Object { $props.Contains($_) } | Select-Object -First 1
    $namCol  = @('Name','DisplayName') | Where-Object { $props.Contains($_) } | Select-Object -First 1
    $chgCol  = @('Changes','What changed') | Where-Object { $props.Contains($_) } | Select-Object -First 1
    $cntCol  = @('LostCount') | Where-Object { $props.Contains($_) } | Select-Object -First 1

    if (-not $accCol -or (-not $lostCol -and -not $chgCol)) {
        Set-Status 'Wrong file type.'
        Show-Message -Text ("The file has no '{0}' column.`n`nColumns found: {1}`n`nPick a comparison file written by Snapshots - step 2 Compare snapshots (comparison_*.csv or .json) - not a snapshot export of users." -f $(if (-not $accCol) { 'Account' } else { 'LostGroups' }), (@($props) -join ', ')) -Icon Warning -Caption 'Restore lost groups'
        return
    }

    $list  = New-Object System.Collections.Generic.List[object]
    $seen  = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $fromChanges = 0; $lostCountSum = 0
    foreach ($row in $rows) {
        if ($null -eq $row) { continue }
        $acc = ([string]$row.$accCol).Trim()
        if (-not $acc) { continue }
        if ($cntCol) { $n = 0; if ([int]::TryParse(([string]$row.$cntCol), [ref]$n)) { $lostCountSum += $n } }
        $lost = if ($lostCol) { @(Split-GroupListText -Text ([string]$row.$lostCol)) } else { @() }
        if ($lost.Count -eq 0 -and $chgCol) {
            # The LostGroups column is empty (or absent) - read the "What changed" text.
            $lost = @(Get-LostGroupsFromChanges -Text ([string]$row.$chgCol))
            if ($lost.Count -gt 0) { $fromChanges += $lost.Count }
        }
        foreach ($g in $lost) {
            if (-not $seen.Add(($acc + '|' + $g))) { continue }   # same user+group listed twice
            try { $r = New-Object ReOrg.RestoreRow -ErrorAction Stop }
            catch {
                Write-Log ("Restore lost groups: could not create a grid row: {0}" -f $_.Exception.Message) 'ERROR'
                Show-Message -Text ("Could not build the grid rows: {0}`n`nClose this window and start AD-ReOrg.ps1 from a new PowerShell window." -f $_.Exception.Message) -Icon Error -Caption 'Restore lost groups'
                return
            }
            $r.Account = $acc
            $r.Name    = if ($namCol) { [string]$row.$namCol } else { '' }
            $r.Group   = $g
            $r.Status  = 'Not checked'
            $r.Include = $true
            $r.add_PropertyChanged({ param($s, $e) if ($e.PropertyName -eq 'Include') { Update-RestoreCount } })
            $list.Add($r)
        }
    }

    $script:RestoreAll  = @($list.ToArray() | Sort-Object Account, Group)
    $script:RestoreFile = $path
    Update-RestoreFilter

    if ($script:RestoreAll.Count -eq 0) {
        Set-Status 'No lost groups in that file.'
        Write-Log ("Restore lost groups: '{0}' holds no lost group memberships ({1} row(s) read; columns used: account '{2}', lost groups '{3}'; LostCount total {4})." -f $path, $rows.Count, $accCol, $(if ($lostCol) { $lostCol } else { '(none)' }), $lostCountSum) 'WARN'
        $detail = @"
No user in that comparison lost a group membership - there is nothing to restore.

Rows read      : $($rows.Count)
Account column : $accCol
Lost groups    : $(if ($lostCol) { $lostCol } else { 'not in the file - the "What changed" text was used instead' })
LostCount total: $lostCountSum

A group counts as lost only when a user had it in the OLD snapshot and has neither that group nor its renamed counterpart in the NEW one. Two things usually explain an empty list:

  * The Pattern boxes on Compare Snapshots (find -> replace, e.g. OLD -> NEW) pair a renamed group with its new name, so it is unchanged instead of lost. Clear those boxes, run Compare again and export a new comparison file.
  * The OLD file was exported after the change, or the users kept their old groups and only gained new ones.
"@
        Show-Message -Text $detail -Icon Information -Caption 'Restore lost groups'
        return
    }
    $users = @($script:RestoreAll | ForEach-Object { $_.Account } | Sort-Object -Unique).Count
    Set-Status ("Loaded {0} lost membership(s) for {1} user(s)." -f $script:RestoreAll.Count, $users)
    Write-Log ("Restore lost groups: {0} lost membership(s) for {1} user(s) from '{2}' ({3} row(s) read, column '{4}')." -f $script:RestoreAll.Count, $users, $path, $rows.Count, $(if ($lostCol) { $lostCol } else { $chgCol })) 'OK'
    if ($fromChanges -gt 0) { Write-Log ("  {0} group name(s) were read from the '{1}' text because the lost-groups column was empty." -f $fromChanges, $chgCol) 'INFO' }
}

function Update-RestoreFilter {
    # Text filter over account, name and group; the checked state lives on the rows,
    # so it survives filtering.
    if ($null -eq $script:RestoreAll) { return }
    $f = if ($script:TxtRestoreFilter) { ([string]$script:TxtRestoreFilter.Text).Trim() } else { '' }
    $rows = @($script:RestoreAll)
    if ($f) { $rows = @($rows | Where-Object { $_.Account -like "*$f*" -or $_.Name -like "*$f*" -or $_.Group -like "*$f*" }) }
    $oc = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    foreach ($r in $rows) { $oc.Add($r) }
    $script:RestoreView = $oc
    if ($script:GridRestore) { $script:GridRestore.ItemsSource = $oc }
    Update-RestoreCount
}

function Update-RestoreCount {
    $all     = @($script:RestoreAll)
    $shown   = @($script:RestoreView)
    $checked = @($all | Where-Object { $_.Include })
    $users   = @($checked | ForEach-Object { $_.Account } | Sort-Object -Unique).Count
    $groups  = @($checked | ForEach-Object { $_.Group }   | Sort-Object -Unique).Count
    if ($script:TxtRestoreCount) {
        $script:TxtRestoreCount.Text = if ($all.Count -eq 0) { 'No comparison file loaded.' }
            else { ('{0} lost membership(s) in the file, {1} shown.{2}Checked: {3} ({4} user(s), {5} group(s)).' -f $all.Count, $shown.Count, [Environment]::NewLine, $checked.Count, $users, $groups) }
    }
}

function Set-RestoreChecked {
    # Checks / unchecks every row currently shown (the filter decides what is shown).
    param([bool]$Value)
    Complete-GridEdit $script:GridRestore
    foreach ($r in @($script:RestoreView)) { $r.Include = $Value }
    Update-RestoreCount
}

function Get-CheckedRestoreRows {
    return @(@($script:RestoreAll) | Where-Object { $_.Include })
}

function Resolve-RestoreUser {
    # Finds the user by sAMAccountName (or UPN) and keeps its memberOf list, cached
    # per run so a user with many lost groups is read only once.
    param([Parameter(Mandatory)] [string]$Account, [string]$Server, [hashtable]$Cache)
    $key = $Account.ToLower()
    if ($Cache.ContainsKey($key)) { return $Cache[$key] }
    $esc = ConvertTo-LdapFilterText $Account
    $filter = if ($Account -like '*@*') { "(|(sAMAccountName=$esc)(userPrincipalName=$esc))" } else { "(sAMAccountName=$esc)" }
    $p = @{ LDAPFilter = $filter; Properties = 'memberOf'; ErrorAction = 'Stop' }
    if ($Server) { $p.Server = $Server }
    $u = $null
    try { $u = Get-ADUser @p | Select-Object -First 1 } catch { Write-Log ("  Could not look up user '{0}': {1}" -f $Account, $_.Exception.Message) 'WARN' }
    $info = if ($u) {
        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($d in @($u.memberOf)) { [void]$set.Add([string]$d) }
        [pscustomobject]@{ Found = $true; DN = [string]$u.DistinguishedName; Name = [string]$u.Name; MemberOf = $set }
    } else {
        [pscustomobject]@{ Found = $false; DN = ''; Name = ''; MemberOf = $null }
    }
    $Cache[$key] = $info
    return $info
}

function Resolve-RestoreGroup {
    # Finds the group by name (CN), sAMAccountName or full DN, cached per run.
    param([Parameter(Mandatory)] [string]$Name, [string]$Server, [hashtable]$Cache)
    $key = $Name.ToLower()
    if ($Cache.ContainsKey($key)) { return $Cache[$key] }
    $g = $null
    $p = @{ ErrorAction = 'Stop' }
    if ($Server) { $p.Server = $Server }
    if ($Name -match '^(?i)CN=.+,.+') {
        try { $g = Get-ADGroup -Identity $Name @p } catch { }
    }
    if (-not $g) {
        $esc = ConvertTo-LdapFilterText $Name
        try { $g = Get-ADGroup -LDAPFilter "(|(name=$esc)(sAMAccountName=$esc))" @p | Select-Object -First 1 }
        catch { Write-Log ("  Could not look up group '{0}': {1}" -f $Name, $_.Exception.Message) 'WARN' }
    }
    $info = if ($g) {
        [pscustomobject]@{ Found = $true; DN = [string]$g.DistinguishedName; Name = [string]$g.Name }
    } else {
        [pscustomobject]@{ Found = $false; DN = ''; Name = '' }
    }
    $Cache[$key] = $info
    return $info
}

function Update-RestoreRowState {
    # Resolves one row against AD and sets its status, DN columns and what to do.
    # Returns 'add', 'member' or 'error'.
    param($Row, [string]$Server, [hashtable]$UserCache, [hashtable]$GroupCache)
    $u = Resolve-RestoreUser  -Account $Row.Account -Server $Server -Cache $UserCache
    if (-not $u.Found) { $Row.UserDN = ''; $Row.GroupDN = ''; $Row.Status = 'Error: user not found in AD'; return 'error' }
    $Row.UserDN = $u.DN
    $g = Resolve-RestoreGroup -Name $Row.Group -Server $Server -Cache $GroupCache
    if (-not $g.Found) { $Row.GroupDN = ''; $Row.Status = 'Error: group not found in AD'; return 'error' }
    $Row.GroupDN = $g.DN
    if ($u.MemberOf -and $u.MemberOf.Contains($g.DN)) { $Row.Status = 'Already member'; return 'member' }
    $Row.Status = 'Missing - will be added'
    return 'add'
}

function Invoke-RestoreCheck {
    Complete-GridEdit $script:GridRestore
    # Reads AD for every checked row without writing anything: does the user exist,
    # does the group exist, and is the membership already back?
    $rows = Get-CheckedRestoreRows
    if ($rows.Count -eq 0) { Show-Message -Text 'No row is checked.' -Icon Warning -Caption 'Restore lost groups'; return }

    $server = $null
    if ($script:Config -and $script:Config.domain.server) { $server = [string]$script:Config.domain.server }
    $userCache = @{}; $groupCache = @{}
    $add = 0; $member = 0; $errors = 0
    $i = 0; $total = $rows.Count
    $script:Pb.Visibility = [System.Windows.Visibility]::Visible
    $script:Pb.Maximum = [Math]::Max($total, 1); $script:Pb.Value = 0
    Set-Status 'Checking Active Directory ...' -Busy

    foreach ($r in $rows) {
        $i++; $script:Pb.Value = $i
        if ($i % 25 -eq 0 -or $i -eq $total) { Set-Status ("Checking {0} of {1} ..." -f $i, $total) -Busy }
        switch (Update-RestoreRowState -Row $r -Server $server -UserCache $userCache -GroupCache $groupCache) {
            'add'    { $add++ }
            'member' { $member++ }
            default  { $errors++ }
        }
    }

    $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed
    Update-GridSafe $script:GridRestore
    $msg = ('Checked {0}: {1} missing, {2} already member, {3} not found.' -f $total, $add, $member, $errors)
    Set-Status $msg
    if ($script:TxtRestoreResult) { $script:TxtRestoreResult.Text = $msg }
    Write-Log ("Restore lost groups: " + $msg) 'OK'
}

function Export-RestoreReport {
    param([object[]]$Records, [switch]$WhatIf, $Settings = $null)
    $folder = Get-ReportFolder -Sub 'restore-lost-groups'
    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $base   = Join-Path $folder ('restore_lost_groups_{0}' -f $stamp)
    $cols = @(
        @{ Key = 'Account';   Label = 'Account (sAMAccountName)' }
        @{ Key = 'Name';      Label = 'Name' }
        @{ Key = 'Group';     Label = 'Group (from the file)' }
        @{ Key = 'GroupDN';   Label = 'Group DN' }
        @{ Key = 'Status';    Label = 'Status' }
        @{ Key = 'Timestamp'; Label = 'Timestamp' }
    )
    $banner = if ($WhatIf) {
        '<div class="banner wi">SIMULATION (WhatIf) - nothing was written to Active Directory. Clear WhatIf and run again for a live run.</div>'
    } else {
        '<div class="banner ok">Live run - the memberships below were restored in Active Directory.</div>'
    }
    $hp = $base + '.html'; New-HtmlReport -Title 'Restore lost groups - result' -Columns $cols -Records $Records -Path $hp -StatusKey 'Status' -Banner $banner -Unit 'memberships' -Settings $Settings -HideChangedBox
    $cp = $base + '.csv';  Export-DataCsv  -Columns $cols -Records $Records -Path $cp
    $jp = $base + '.json'; Export-DataJson -Records $Records -Path $jp -Meta @{ type = 'restore-lost-groups'; whatIf = [bool]$WhatIf; rows = @($Records).Count }
    Write-Log ("Report: {0}" -f $hp) 'OK'
    if ($script:Config -and $script:Config.report.openAfter -and (Test-Path -LiteralPath $hp)) { Start-Process $hp }
}

function Invoke-RestoreApply {
    Complete-GridEdit $script:GridRestore
    # Adds every checked user back to its lost group. Additive only - no membership
    # is ever removed and no group is created.
    $rows = Get-CheckedRestoreRows
    if ($rows.Count -eq 0) { Show-Message -Text 'No row is checked - nothing to restore.' -Icon Warning -Caption 'Restore lost groups'; return }

    $whatIf = if ($script:ChkRestoreWhatIf) { [bool]$script:ChkRestoreWhatIf.IsChecked } else { $true }
    $mode   = if ($whatIf) { 'WhatIf (SIMULATION - nothing changes)' } else { 'LIVE RUN' }
    $users  = @($rows | ForEach-Object { $_.Account } | Sort-Object -Unique)
    $groups = @($rows | ForEach-Object { $_.Group }   | Sort-Object -Unique)
    $confirm = @"
Restore lost group memberships.

File       : $($script:RestoreFile)
Memberships: $($rows.Count)
Users      : $($users.Count)
Groups     : $($groups.Count)
Run mode   : $mode

The users are added back to the groups - nothing is removed. Continue?
"@
    if ((Show-Message -Text $confirm -Icon Question -Buttons YesNo -Caption 'Confirm') -ne [System.Windows.MessageBoxResult]::Yes) { Set-Status 'Cancelled.'; return }

    $server = $null
    if ($script:Config -and $script:Config.domain.server) { $server = [string]$script:Config.domain.server }
    $userCache = @{}; $groupCache = @{}
    $pending = @{}                      # group DN -> rows to add
    $added = 0; $already = 0; $errors = 0

    Write-Log ("=== Restore lost groups: {0} membership(s), {1} user(s), {2} group(s) [{3}] ===" -f $rows.Count, $users.Count, $groups.Count, $mode) 'INFO'
    Set-Status 'Reading Active Directory ...' -Busy
    $i = 0; $total = $rows.Count
    $script:Pb.Visibility = [System.Windows.Visibility]::Visible
    $script:Pb.Maximum = [Math]::Max($total, 1); $script:Pb.Value = 0

    foreach ($r in $rows) {
        $i++; $script:Pb.Value = $i
        if ($i % 25 -eq 0 -or $i -eq $total) { Set-Status ("Reading {0} of {1} ..." -f $i, $total) -Busy }
        switch (Update-RestoreRowState -Row $r -Server $server -UserCache $userCache -GroupCache $groupCache) {
            'add' {
                if ($whatIf) { $r.Status = 'Would add' }
                else {
                    if (-not $pending.ContainsKey($r.GroupDN)) { $pending[$r.GroupDN] = New-Object System.Collections.Generic.List[object] }
                    $pending[$r.GroupDN].Add($r)
                }
            }
            'member' { $already++ }
            default  { $errors++ }
        }
    }

    if (-not $whatIf -and $pending.Count -gt 0) {
        $g = 0
        $script:Pb.Maximum = [Math]::Max($pending.Count, 1); $script:Pb.Value = 0
        foreach ($gdn in @($pending.Keys)) {
            $g++; $script:Pb.Value = $g
            $items = @($pending[$gdn].ToArray())   # .ToArray() first - @() on a generic List can throw in some PowerShell builds
            $gname = $items[0].Group
            Set-Status ("Adding members to '{0}' ({1} of {2}) ..." -f $gname, $g, $pending.Count) -Busy
            $p = @{ Identity = $gdn; Members = @($items | ForEach-Object { $_.UserDN }); ErrorAction = 'Stop' }
            if ($server) { $p.Server = $server }
            try {
                Add-ADGroupMember @p
                foreach ($r in $items) { $r.Status = 'Added'; $added++ }
                Write-Log ("  Added {0} member(s) to '{1}'." -f $items.Count, $gname) 'OK'
            }
            catch {
                # One bad member must not sink the rest of the group - retry one by one.
                Write-Log ("  Group '{0}': batch add failed ({1}) - retrying one by one." -f $gname, $_.Exception.Message) 'WARN'
                foreach ($r in $items) {
                    $sp = @{ Identity = $gdn; Members = $r.UserDN; ErrorAction = 'Stop' }
                    if ($server) { $sp.Server = $server }
                    try { Add-ADGroupMember @sp; $r.Status = 'Added'; $added++ }
                    catch {
                        $msg = $_.Exception.Message
                        if ($msg -match '(?i)already a member') { $r.Status = 'Already member'; $already++ }
                        else { $r.Status = 'Error: ' + $msg; $errors++; Write-Log ("    {0} -> '{1}': {2}" -f $r.Account, $gname, $msg) 'ERROR' }
                    }
                }
            }
        }
    }

    $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed
    Update-GridSafe $script:GridRestore

    $wouldAdd = @($rows | Where-Object { $_.Status -eq 'Would add' }).Count
    $result = if ($whatIf) {
        ('Done (WhatIf): {0} would be added, {1} already member, {2} not found.' -f $wouldAdd, $already, $errors)
    } else {
        ('Done: {0} added, {1} already member, {2} error(s).' -f $added, $already, $errors)
    }
    Set-Status $result
    if ($script:TxtRestoreResult) { $script:TxtRestoreResult.Text = $result }
    Write-Log ("=== " + $result + " ===") 'OK'

    $stamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $records = @($rows | ForEach-Object {
        [pscustomobject]@{
            Account   = $_.Account
            Name      = $_.Name
            Group     = $_.Group
            GroupDN   = $_.GroupDN
            Status    = $_.Status
            Timestamp = $stamp
        }
    })
    $settings = [ordered]@{
        'Run mode'        = $(if ($whatIf) { 'WhatIf - simulation, nothing written' } else { 'Live run' })
        'Comparison file' = [string]$script:RestoreFile
        'Source column'   = 'LostGroups (groups a user lost between the OLD and NEW snapshot)'
        'Memberships'     = ('{0} checked of {1} in the file' -f $rows.Count, @($script:RestoreAll).Count)
        'Users / groups'  = ('{0} user(s), {1} group(s)' -f $users.Count, $groups.Count)
        'Result'          = $result
    }
    try { Export-RestoreReport -Records $records -WhatIf:$whatIf -Settings $settings }
    catch { Write-Log ("Restore report failed: {0}" -f $_.Exception.Message) 'WARN' }
}
