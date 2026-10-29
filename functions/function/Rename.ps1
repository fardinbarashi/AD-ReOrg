<#
    Groups : Rename Groups.

    Renames existing groups in place instead of creating new ones: the group keeps
    its SID, members, memberOf and every permission granted to it - only the names
    change. Per group the tab can change:
      * Name (CN)        - Rename-ADObject (the DN follows)
      * sAMAccountName   - unique in the whole domain, checked before the run
      * displayName
      * description      - the same Find -> Replace and/or a text added at the end
                           or the start
    The rule (Find -> Replace, prefix, suffix) is applied to each attribute's own
    current value. Every run writes HTML/CSV/JSON to files\rename-groups; a live
    run's JSON holds the old and new values, so it can be loaded again to rename
    the groups back (revert).

    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

# ---------- Tree ----------

function Update-RenameCount {
    if (-not $script:RenameTree) { return }
    $n = @(Get-CheckedGroups $script:RenameTree).Count
}


# ---------- Rule ----------

function Get-RenameRule {
    # Everything the user chose in the name-rule and attribute sections.
    $t = { param($c) if ($c) { [string]$c.Text } else { '' } }
    $b = { param($c) [bool]($c -and $c.IsChecked) }
    $noSep = & $b $script:ChkRenDescNoSep
    $pos = if ($script:CmbRenDescPos -and $script:CmbRenDescPos.SelectedItem) { [string]$script:CmbRenDescPos.SelectedItem.Content } else { 'At the end' }
    return [pscustomobject]@{
        Find        = (& $t $script:TxtRenFind)
        Replace     = (& $t $script:TxtRenReplace)
        Regex       = (& $b $script:ChkRenRegex)
        Prefix      = (& $t $script:TxtRenPrefix).Trim()
        Suffix      = (& $t $script:TxtRenSuffix).Trim()
        DoName      = (& $b $script:ChkRenName)
        DoSam       = (& $b $script:ChkRenSam)
        DoDisplay   = (& $b $script:ChkRenDisplay)
        DescFind    = (& $b $script:ChkRenDescFind)
        DescAdd     = (& $t $script:TxtRenDescAdd).Trim()
        DescSep     = $(if ($noSep) { '' } else { & $t $script:TxtRenDescSep })
        DescAtStart = ($pos -eq 'At the start')
    }
}

function Get-RenamedText {
    # Find -> Replace (case-insensitive literal text, or a regular expression), then
    # prefix + suffix. A prefix/suffix the value already carries is not added again,
    # so running the same rule twice does not give "NEW-NEW-Sales".
    param([string]$Value, $Rule, [switch]$NoAffix)
    $v = [string]$Value
    if ($Rule.Find) {
        try {
            if ($Rule.Regex) { $v = [regex]::Replace($v, $Rule.Find, $Rule.Replace) }
            else { $v = Get-ReplacedText -Text $v -Find $Rule.Find -Replace $Rule.Replace }
        } catch { Write-Log ("Rename: the Find pattern '{0}' is not valid: {1}" -f $Rule.Find, $_.Exception.Message) 'WARN' }
    }
    if (-not $NoAffix) {
        if ($Rule.Prefix -and -not $v.StartsWith($Rule.Prefix, [StringComparison]::OrdinalIgnoreCase)) { $v = $Rule.Prefix + $v }
        if ($Rule.Suffix -and -not $v.EndsWith($Rule.Suffix, [StringComparison]::OrdinalIgnoreCase))   { $v = $v + $Rule.Suffix }
    }
    return $v.Trim()
}

function Get-RenamedDescription {
    # Optional Find -> Replace in the description, then the "Add text" at the end or
    # the start (skipped when the description already contains that text).
    param([string]$Value, $Rule)
    $d = [string]$Value
    if ($Rule.DescFind -and $Rule.Find) { $d = Get-RenamedText -Value $d -Rule $Rule -NoAffix }
    $add = [string]$Rule.DescAdd
    if ($add -and $d.IndexOf($add, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        if (-not $d.Trim())       { $d = $add }
        elseif ($Rule.DescAtStart) { $d = $add + $Rule.DescSep + $d }
        else                       { $d = $d + $Rule.DescSep + $add }
    }
    return $d
}

function Test-RenameRowChanged {
    # Case-sensitive on purpose: "sales" -> "Sales" is a real rename.
    param($Row)
    return ($Row.NewName -cne $Row.OldName -or $Row.NewSam -cne $Row.OldSam -or
            $Row.NewDisplay -cne $Row.OldDisplay -or $Row.NewDesc -cne $Row.OldDesc)
}

function Get-RenameChangeText {
    param($Row)
    $c = @()
    if ($Row.NewName    -cne $Row.OldName)    { $c += 'Name' }
    if ($Row.NewSam     -cne $Row.OldSam)     { $c += 'sAMAccountName' }
    if ($Row.NewDisplay -cne $Row.OldDisplay) { $c += 'displayName' }
    if ($Row.NewDesc    -cne $Row.OldDesc)    { $c += 'description' }
    return ($c -join ', ')
}

function Set-RenameRowTargets {
    # Computes the new values of one row from its current values and the rule.
    param($Row, $Rule)
    $Row.NewName    = if ($Rule.DoName)    { Get-RenamedText -Value $Row.OldName -Rule $Rule } else { $Row.OldName }
    $Row.NewSam     = $Row.OldSam
    if ($Rule.DoSam) { $x = Get-RenamedText -Value $Row.OldSam -Rule $Rule; $Row.NewSam = if ($x) { ConvertTo-SamAccountName -Name $x } else { '' } }
    $Row.NewDisplay = if ($Rule.DoDisplay -and $Row.OldDisplay) { Get-RenamedText -Value $Row.OldDisplay -Rule $Rule } else { $Row.OldDisplay }
    $Row.NewDesc    = Get-RenamedDescription -Value $Row.OldDesc -Rule $Rule
    $Row.Status     = if (Test-RenameRowChanged $Row) { 'Not checked' } else { 'No change' }
}

function Set-RenameDescState {
    # The separator box is only editable while "No separator" is cleared.
    if ($script:TxtRenDescSep -and $script:ChkRenDescNoSep) { $script:TxtRenDescSep.IsEnabled = -not [bool]$script:ChkRenDescNoSep.IsChecked }
}

function Set-RenameMode {
    # 'Rename' (rows built from the tree) or 'Revert' (rows loaded from a report).
    param([ValidateSet('Rename','Revert')] [string]$Mode, [string]$File = '')
    $script:RenameMode = $Mode
    $script:RenameRevertFile = $File
    if (-not $script:TxtRenMode) { return }
    if ($Mode -eq 'Revert') {
        $script:TxtRenMode.Text = ("REVERT - the rows below rename the groups back to the old values from {0}. Check in AD, then run with WhatIf first. 'Build preview' leaves this mode." -f [System.IO.Path]::GetFileName($File))
        if ($script:BrdRenMode) { $script:BrdRenMode.Visibility = [System.Windows.Visibility]::Visible }
    } else {
        $script:TxtRenMode.Text = ''
        if ($script:BrdRenMode) { $script:BrdRenMode.Visibility = [System.Windows.Visibility]::Collapsed }
    }
}

# ---------- Preview ----------

function Test-RenameRowType {
    # The grid rows are a compiled .NET type; a PowerShell window that already ran an
    # older AD ReOrg cannot load a newer one, so show a message instead of failing without one.
    if ('ReOrg.RenameRow' -as [type]) { return $true }
    Write-Log 'Rename: the row type ReOrg.RenameRow is missing in this PowerShell session.' 'ERROR'
    Show-Message -Text "This PowerShell session cannot build the preview rows (the type ReOrg.RenameRow is not loaded).`n`nThat happens when an older version of AD ReOrg was started in the same PowerShell window: .NET keeps the first version of a type for the life of the process.`n`nClose this window and start AD-ReOrg.ps1 from a NEW PowerShell window." -Icon Error -Caption 'Rename groups'
    return $false
}

function Set-RenameRows {
    param([object[]]$Rows)
    $oc = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    foreach ($r in @($Rows)) {
        $r.add_PropertyChanged({
            param($s, $e)
            if ($e.PropertyName -eq 'Include') { Update-RenamePreviewCount; return }
            if ($e.PropertyName -in 'NewName','NewSam','NewDisplay','NewDesc') {
                # A value edited in the grid has to be checked again.
                if ($s.Status -eq 'Ready' -or $s.Status -eq 'No change' -or $s.Status -like 'Error:*') {
                    $s.Status = if (Test-RenameRowChanged $s) { 'Not checked' } else { 'No change' }
                }
                Update-RenamePreviewCount
            }
        })
        $oc.Add($r)
    }
    $script:RenameRows = $oc
    if ($script:GridRename) { $script:GridRename.ItemsSource = $oc }
    Update-RenameView
}

function Build-RenamePreview {
    if (-not (Test-RenameRowType)) { return }
    $nodes = @(Get-CheckedGroups $script:RenameTree)
    if ($nodes.Count -eq 0) { Show-Message -Text 'Check at least one group (or an OU) in the selection to the left first.' -Icon Warning -Caption 'Rename groups'; return }
    $rule = Get-RenameRule
    if (-not ($rule.DoName -or $rule.DoSam -or $rule.DoDisplay -or $rule.DescFind -or $rule.DescAdd)) {
        Show-Message -Text 'Nothing to change: check at least one of Name (CN), sAMAccountName, displayName, or fill in the description options.' -Icon Warning -Caption 'Rename groups'; return
    }
    Set-RenameMode -Mode 'Rename'
    $server = $script:LoadedServer
    $rows = New-Object System.Collections.Generic.List[object]
    $total = $nodes.Count; $i = 0
    $script:Pb.Visibility = [System.Windows.Visibility]::Visible
    $script:Pb.Maximum = [Math]::Max($total, 1); $script:Pb.Value = 0
    foreach ($n in $nodes) {
        $i++; $script:Pb.Value = $i
        if ($i % 10 -eq 0 -or $i -eq $total) { Set-Status ("Reading groups {0} of {1} ..." -f $i, $total) -Busy }
        try {
            $p = @{ Identity = $n.DN; Properties = 'displayName','description'; ErrorAction = 'Stop' }
            if ($server) { $p.Server = $server }
            $g = Get-ADGroup @p
        }
        catch { Write-Log ("Rename: could not read '{0}': {1}" -f $n.Name, $_.Exception.Message) 'WARN'; continue }
        $r = New-Object ReOrg.RenameRow
        $r.ObjectGuid = [string]$g.ObjectGUID
        $r.OldDN      = [string]$g.DistinguishedName
        $r.OldName    = [string]$g.Name
        $r.OldSam     = [string]$g.SamAccountName
        $r.OldDisplay = [string]$g.displayName
        $r.OldDesc    = [string]$g.Description
        Set-RenameRowTargets -Row $r -Rule $rule
        $rows.Add($r)
    }
    $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed
    Set-RenameRows -Rows @($rows.ToArray() | Sort-Object OldName)
    $chg = @($script:RenameRows | Where-Object { $_.Status -ne 'No change' }).Count
    Set-Status ("Preview: {0} group(s), {1} with a change." -f $rows.Count, $chg)
    Write-Log  ("Rename preview: {0} group(s), {1} with a change." -f $rows.Count, $chg) 'INFO'
}

function Update-RenamePreviewNames {
    Complete-GridEdit $script:GridRename
    # "Recalculate": applies the current rule again to the rows already in the grid.
    if (-not $script:RenameRows -or $script:RenameRows.Count -eq 0) { Build-RenamePreview; return }
    if ($script:RenameMode -eq 'Revert') {
        Show-Message -Text "The grid holds a revert loaded from a report - its values come from that report, not from the rule.`n`nClick 'Build preview' to start a new rename from the tree instead." -Icon Info -Caption 'Rename groups'; return
    }
    $rule = Get-RenameRule
    foreach ($r in $script:RenameRows) { Set-RenameRowTargets -Row $r -Rule $rule }
    Update-RenameView
    Set-Status 'New values recalculated from the rule.'
}

function Update-RenameView {
    # Two check boxes decide which rows are shown: groups the rule changes and/or
    # groups it does not change (status No change - never renamed either way).
    if (-not $script:GridRename -or -not $script:RenameRows) { return }
    try {
        [void]$script:GridRename.CommitEdit()
        $showChg  = (-not $script:ChkRenShowChanged) -or [bool]$script:ChkRenShowChanged.IsChecked
        $showSame = [bool]($script:ChkRenShowSame -and $script:ChkRenShowSame.IsChecked)
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:RenameRows)
        if ($showChg -and $showSame) { $view.Filter = $null }
        else {
            $script:RenameShowChg = $showChg; $script:RenameShowSame = $showSame
            $view.Filter = [Predicate[object]]{
                param($r)
                $c = ($r.NewName -cne $r.OldName -or $r.NewSam -cne $r.OldSam -or $r.NewDisplay -cne $r.OldDisplay -or $r.NewDesc -cne $r.OldDesc)
                if ($c) { return [bool]$script:RenameShowChg } else { return [bool]$script:RenameShowSame }
            }
        }
        $view.Refresh()
    } catch { Write-Log ("Rename: could not filter the preview: {0}" -f $_.Exception.Message) 'WARN' }
    Update-RenamePreviewCount
}

function Update-RenamePreviewCount {
    if (-not $script:TxtRenCount) { return }
    if (-not $script:RenameRows -or $script:RenameRows.Count -eq 0) { $script:TxtRenCount.Text = 'No preview built yet - click "Build preview".'; return }
    $all    = @($script:RenameRows)
    $chgAll = @($all | Where-Object { Test-RenameRowChanged $_ })
    $chgInc = @($chgAll | Where-Object { $_.Include }).Count
    $same   = $all.Count - $chgAll.Count
    $verb   = if ($script:RenameMode -eq 'Revert') { 'Groups to rename back' } else { 'Groups to rename' }
    $text   = ("{0}: {1} (checked {1} of {2} with a change)" -f $verb, $chgInc, $chgAll.Count)
    if ($same -gt 0) {
        $shown = [bool]($script:ChkRenShowSame -and $script:ChkRenShowSame.IsChecked)
        $text += ("  ·  {0} selected group(s) will not change{1}" -f $same, $(if ($shown) { ' - shown as No change, skipped' } else { ' - hidden' }))
    }
    $script:TxtRenCount.Text = $text
}

function Set-RenameRowsChecked {
    param([bool]$Value)
    Complete-GridEdit $script:GridRename
    if (-not $script:RenameRows) { return }
    foreach ($r in $script:RenameRows) { $r.Include = $Value }
    Update-RenamePreviewCount
}

# ---------- Checks ----------

function Get-ParentDN {
    # Parent container of a DN; the first comma that is not escaped (CN=Smith\, J,OU=...).
    param([string]$DN)
    $parts = ([regex]'(?<!\\),').Split([string]$DN, 2)
    if ($parts.Count -gt 1) { return $parts[1] }
    return ''
}

function Test-RenameRows {
    # Validates every row that would change, against the batch and against AD, and
    # sets its Status to 'Ready' or 'Error: ...'. Nothing is written.
    param($Rows, [string]$Server)
    $rows = @($Rows)
    $samSeen = @{}; $nameSeen = @{}
    foreach ($r in $rows) {
        if (-not (Test-RenameRowChanged $r)) { continue }
        $k = ([string]$r.NewSam).ToLower();  $samSeen[$k]  = [int]$samSeen[$k] + 1
        $k = ((Get-ParentDN $r.OldDN) + '|' + $r.NewName).ToLower(); $nameSeen[$k] = [int]$nameSeen[$k] + 1
    }
    $total = $rows.Count; $i = 0
    foreach ($r in $rows) {
        $i++
        if ($i % 10 -eq 0 -or $i -eq $total) { Set-Status ("Checking {0} of {1} ..." -f $i, $total) -Busy }
        $r.NewName = ([string]$r.NewName).Trim(); $r.NewSam = ([string]$r.NewSam).Trim()
        if (-not (Test-RenameRowChanged $r)) { $r.Status = 'No change'; continue }
        if (-not $r.NewName) { $r.Status = 'Error: the new name is empty'; continue }
        if ($r.NewName.Length -gt 64) { $r.Status = ('Error: the new name has {0} characters - a CN allows 64' -f $r.NewName.Length); continue }
        if (-not $r.NewSam) { $r.Status = 'Error: the new sAMAccountName is empty'; continue }
        if ($r.NewSam -match '[\\/\[\]:;\|=\+\*\?<>",@]') { $r.Status = 'Error: the new sAMAccountName holds a character AD does not allow'; continue }
        if ($r.NewSam.Length -gt 256) { $r.Status = 'Error: the new sAMAccountName is longer than 256 characters'; continue }
        if ($r.NewSam -cne $r.OldSam -and $samSeen[$r.NewSam.ToLower()] -gt 1) { $r.Status = 'Error: the same new sAMAccountName is used by more than one row'; continue }
        if ($r.NewName -cne $r.OldName -and $nameSeen[((Get-ParentDN $r.OldDN) + '|' + $r.NewName).ToLower()] -gt 1) { $r.Status = 'Error: two rows get the same new name in the same OU'; continue }

        $g = $null
        try {
            $p = @{ Identity = $r.ObjectGuid; Properties = 'displayName','description'; ErrorAction = 'Stop' }
            if ($Server) { $p.Server = $Server }
            $g = Get-ADGroup @p
        } catch { }
        if (-not $g) { $r.Status = 'Error: the group was not found in AD'; continue }
        if ([string]$g.Name -cne $r.OldName -or [string]$g.SamAccountName -cne $r.OldSam) {
            $r.Status = ("Error: renamed in AD since the preview (now '{0}' / '{1}') - build the preview again" -f $g.Name, $g.SamAccountName); continue
        }
        $r.OldDN = [string]$g.DistinguishedName

        if ($r.NewName -ine $r.OldName) {
            $esc = ConvertTo-LdapFilterText $r.NewName
            $hit = $null
            try {
                $p = @{ SearchBase = (Get-ParentDN $r.OldDN); SearchScope = 'OneLevel'; LDAPFilter = "(name=$esc)"; ErrorAction = 'Stop' }
                if ($Server) { $p.Server = $Server }
                $hit = Get-ADObject @p | Where-Object { [string]$_.ObjectGUID -ne $r.ObjectGuid } | Select-Object -First 1
            } catch { Write-Log ("Rename: could not check the name '{0}': {1}" -f $r.NewName, $_.Exception.Message) 'WARN' }
            if ($hit) { $r.Status = ("Error: '{0}' already exists in the same OU" -f $r.NewName); continue }
        }
        if ($r.NewSam -ine $r.OldSam) {
            $esc = ConvertTo-LdapFilterText $r.NewSam
            $hit = $null
            try {
                $p = @{ LDAPFilter = "(sAMAccountName=$esc)"; ErrorAction = 'Stop' }
                if ($Server) { $p.Server = $Server }
                $hit = Get-ADObject @p | Where-Object { [string]$_.ObjectGUID -ne $r.ObjectGuid } | Select-Object -First 1
            } catch { Write-Log ("Rename: could not check the sAMAccountName '{0}': {1}" -f $r.NewSam, $_.Exception.Message) 'WARN' }
            if ($hit) { $r.Status = ("Error: sAMAccountName '{0}' is already used by '{1}'" -f $r.NewSam, $hit.Name); continue }
        }
        $r.Status = 'Ready'
    }
}

function Invoke-RenameCheck {
    Complete-GridEdit $script:GridRename
    if (-not $script:RenameRows -or $script:RenameRows.Count -eq 0) { Show-Message -Text 'Build the preview first (the "Build preview" button).' -Icon Warning -Caption 'Rename groups'; return }
    $rows = @($script:RenameRows | Where-Object { $_.Include })
    if ($rows.Count -eq 0) { Show-Message -Text 'No row is checked in the preview.' -Icon Warning -Caption 'Rename groups'; return }
    Test-RenameRows -Rows $rows -Server $script:LoadedServer
    Update-GridSafe $script:GridRename
    $ready = @($rows | Where-Object { $_.Status -eq 'Ready' }).Count
    $err   = @($rows | Where-Object { $_.Status -like 'Error:*' }).Count
    $same  = @($rows | Where-Object { $_.Status -eq 'No change' }).Count
    $msg = ('Checked {0}: {1} ready, {2} error(s), {3} without a change.' -f $rows.Count, $ready, $err, $same)
    Set-Status $msg
    if ($script:TxtRenResult) { $script:TxtRenResult.Text = $msg }
    Write-Log ('Rename check: ' + $msg) $(if ($err) { 'WARN' } else { 'OK' })
}

# ---------- Run ----------

function Export-RenameReport {
    param([object[]]$Records, [switch]$WhatIf, $Settings = $null)
    $folder = Get-ReportFolder -Sub 'rename-groups'
    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $base   = Join-Path $folder ('rename_groups_{0}' -f $stamp)
    $cols = @(
        @{ Key = 'Status';         Label = 'Status' }
        @{ Key = 'OldName';        Label = 'Old name' }
        @{ Key = 'NewName';        Label = 'New name' }
        @{ Key = 'OldSam';         Label = 'Old sAMAccountName' }
        @{ Key = 'NewSam';         Label = 'New sAMAccountName' }
        @{ Key = 'OldDisplayName'; Label = 'Old displayName' }
        @{ Key = 'NewDisplayName'; Label = 'New displayName' }
        @{ Key = 'OldDescription'; Label = 'Old description' }
        @{ Key = 'NewDescription'; Label = 'New description' }
        @{ Key = 'Changes';        Label = 'Changed' }
        @{ Key = 'OldDN';          Label = 'Old DN' }
        @{ Key = 'NewDN';          Label = 'New DN' }
        @{ Key = 'ObjectGUID';     Label = 'objectGUID' }
        @{ Key = 'Timestamp';      Label = 'Timestamp' }
    )
    $banner = if ($WhatIf) {
        '<div class="banner wi">SIMULATION (WhatIf) - nothing was written to Active Directory. Clear WhatIf and run again for a live run.</div>'
    } else {
        '<div class="banner ok">Live run - the groups below were renamed in Active Directory. Open this report''s JSON with Revert from report... in Rename groups to rename them back.</div>'
    }
    $hp = $base + '.html'; New-HtmlReport -Title 'Rename groups - result' -Columns $cols -Records $Records -Path $hp -StatusKey 'Status' -Banner $banner -Unit 'groups' -Settings $Settings -HideChangedBox
    $cp = $base + '.csv';  Export-DataCsv  -Columns $cols -Records $Records -Path $cp
    $jp = $base + '.json'; Export-DataJson -Records $Records -Path $jp -Meta @{ type = 'rename-groups'; whatIf = [bool]$WhatIf; mode = [string]$script:RenameMode; rows = @($Records).Count }
    Write-Log ("Report: {0}" -f $hp) 'OK'
    if ($script:Config -and $script:Config.report.openAfter -and (Test-Path -LiteralPath $hp)) { Start-Process $hp }
    return $jp
}

function Invoke-RenameApply {
    Complete-GridEdit $script:GridRename
    if (-not $script:RenameRows -or $script:RenameRows.Count -eq 0) { Show-Message -Text 'Build the preview first (the "Build preview" button).' -Icon Warning -Caption 'Rename groups'; return }
    $rows = @($script:RenameRows | Where-Object { $_.Include })
    if ($rows.Count -eq 0) { Show-Message -Text 'No row is checked in the preview.' -Icon Warning -Caption 'Rename groups'; return }
    $server = $script:LoadedServer
    $whatIf = if ($script:ChkRenWhatIf) { [bool]$script:ChkRenWhatIf.IsChecked } else { $true }
    $revert = ($script:RenameMode -eq 'Revert')

    # Always validate again - the grid may have been edited since the last check.
    Test-RenameRows -Rows $rows -Server $server
    Update-GridSafe $script:GridRename
    $ready = @($rows | Where-Object { $_.Status -eq 'Ready' })
    $err   = @($rows | Where-Object { $_.Status -like 'Error:*' }).Count
    if ($ready.Count -eq 0) {
        Set-Status 'Nothing to rename.'
        Show-Message -Text ("No row is ready to rename.`n`n{0} row(s) have an error (see the Status column), the rest have no change." -f $err) -Icon Warning -Caption 'Rename groups'
        return
    }
    $mode = if ($whatIf) { 'WhatIf (SIMULATION - nothing changes)' } else { 'LIVE RUN' }
    $what = if ($revert) { 'Rename back (revert)' } else { 'Rename' }
    $confirm = @"
$what $($ready.Count) group(s).

Name (CN) changes : $(@($ready | Where-Object { $_.NewName -cne $_.OldName }).Count)
sAMAccountName    : $(@($ready | Where-Object { $_.NewSam -cne $_.OldSam }).Count)
displayName       : $(@($ready | Where-Object { $_.NewDisplay -cne $_.OldDisplay }).Count)
description       : $(@($ready | Where-Object { $_.NewDesc -cne $_.OldDesc }).Count)
Skipped (errors)  : $err
Run mode          : $mode

The groups keep their SID, members and permissions. Continue?
"@
    if ((Show-Message -Text $confirm -Icon Question -Buttons YesNo -Caption 'Confirm') -ne [System.Windows.MessageBoxResult]::Yes) { Set-Status 'Cancelled.'; return }

    Write-Log ("=== {0}: {1} group(s) [{2}] ===" -f $what, $ready.Count, $mode) 'INFO'
    $done = 0; $failed = 0; $i = 0
    $script:Pb.Visibility = [System.Windows.Visibility]::Visible
    $script:Pb.Maximum = [Math]::Max($ready.Count, 1); $script:Pb.Value = 0
    foreach ($r in $ready) {
        $i++; $script:Pb.Value = $i
        Set-Status ("{0} {1} of {2}: {3} ..." -f $what, $i, $ready.Count, $r.OldName) -Busy
        $changes = Get-RenameChangeText $r
        if ($whatIf) {
            $r.Status = if ($revert) { 'Would revert' } else { 'Would rename' }
            $r.NewDN = if ($r.NewName -cne $r.OldName) { 'CN=' + ($r.NewName -replace '([,\\#+<>;"=])', '\$1') + ',' + (Get-ParentDN $r.OldDN) } else { $r.OldDN }
            Write-Log ("  Would change {0} on '{1}' -> '{2}' (sAM '{3}')." -f $changes, $r.OldName, $r.NewName, $r.NewSam) 'INFO'
            $done++; continue
        }
        $attrsDone = $false
        try {
            $replace = @{}; $clear = New-Object System.Collections.Generic.List[string]
            if ($r.NewSam     -cne $r.OldSam)     { $replace['sAMAccountName'] = $r.NewSam }
            if ($r.NewDisplay -cne $r.OldDisplay) { if ($r.NewDisplay) { $replace['displayName'] = $r.NewDisplay } else { $clear.Add('displayName') } }
            if ($r.NewDesc    -cne $r.OldDesc)    { if ($r.NewDesc)    { $replace['description'] = $r.NewDesc }    else { $clear.Add('description') } }
            if ($replace.Count -gt 0 -or $clear.Count -gt 0) {
                $p = @{ Identity = $r.ObjectGuid; ErrorAction = 'Stop' }
                if ($replace.Count -gt 0) { $p.Replace = $replace }
                if ($clear.Count -gt 0)   { $p.Clear = $clear.ToArray() }
                if ($server) { $p.Server = $server }
                Set-ADGroup @p
            }
            $attrsDone = $true
            if ($r.NewName -cne $r.OldName) {
                $p = @{ Identity = $r.ObjectGuid; NewName = $r.NewName; ErrorAction = 'Stop' }
                if ($server) { $p.Server = $server }
                Rename-ADObject @p
            }
            try {
                $p = @{ Identity = $r.ObjectGuid; ErrorAction = 'Stop' }; if ($server) { $p.Server = $server }
                $r.NewDN = [string](Get-ADGroup @p).DistinguishedName
            } catch { $r.NewDN = '' }
            $r.Status = if ($revert) { 'Reverted' } else { 'Renamed' }
            $done++
            Write-Log ("  {0}: '{1}' -> '{2}' ({3})." -f $r.Status, $r.OldName, $r.NewName, $changes) 'OK'
        }
        catch {
            $failed++
            $msg = $_.Exception.Message
            $r.Status = if ($attrsDone) { 'Error: the attributes were updated but the name (CN) could not be changed: ' + $msg } else { 'Error: ' + $msg }
            Write-Log ("  Could not rename '{0}': {1}" -f $r.OldName, $msg) 'ERROR'
        }
    }
    $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed
    Update-GridSafe $script:GridRename

    $result = if ($whatIf) { ('Done (WhatIf): {0} group(s) would be renamed, {1} skipped with an error.' -f $done, $err) }
              else { ('Done: {0} group(s) renamed, {1} failed, {2} skipped with an error.' -f $done, $failed, $err) }
    Set-Status $result
    if ($script:TxtRenResult) { $script:TxtRenResult.Text = $result }
    Write-Log ('=== ' + $result + ' ===') 'OK'
    if (-not $whatIf -and $done -gt 0) { Write-Log 'Rename: click Reload (top right) to show the new names in every tree.' 'INFO' }

    $stamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $records = @($rows | ForEach-Object {
        [pscustomobject]@{
            Status = $_.Status; OldName = $_.OldName; NewName = $_.NewName; OldSam = $_.OldSam; NewSam = $_.NewSam
            OldDisplayName = $_.OldDisplay; NewDisplayName = $_.NewDisplay; OldDescription = $_.OldDesc; NewDescription = $_.NewDesc
            Changes = (Get-RenameChangeText $_); OldDN = $_.OldDN; NewDN = $_.NewDN; ObjectGUID = $_.ObjectGuid; Timestamp = $stamp
        }
    })
    $rule = Get-RenameRule
    $settings = [ordered]@{}
    $settings['Run mode'] = $(if ($whatIf) { 'WhatIf - simulation, nothing written' } else { 'Live run' })
    if ($revert) {
        $settings['Mode']          = 'Revert - rename back to the old values'
        $settings['Source report'] = [string]$script:RenameRevertFile
    } else {
        $attrs = @(); if ($rule.DoName) { $attrs += 'Name (CN)' }; if ($rule.DoSam) { $attrs += 'sAMAccountName' }; if ($rule.DoDisplay) { $attrs += 'displayName' }
        $settings['Mode']                    = 'Rename'
        $settings['Find -> Replace']         = $(if ($rule.Find) { ('"{0}" -> "{1}"{2}' -f $rule.Find, $rule.Replace, $(if ($rule.Regex) { '  (regular expression)' } else { '  (case-insensitive)' })) } else { '(none)' })
        $settings['Prefix / Suffix']         = ('{0} / {1}' -f $(if ($rule.Prefix) { $rule.Prefix } else { '-' }), $(if ($rule.Suffix) { $rule.Suffix } else { '-' }))
        $settings['Attributes renamed']      = $(if ($attrs) { $attrs -join ', ' } else { 'None' })
        $settings['Description']             = ('Find -> Replace: {0}; add text: {1}' -f $(if ($rule.DescFind) { 'Yes' } else { 'No' }), $(if ($rule.DescAdd) { ('"{0}" {1}, separator {2}' -f $rule.DescAdd, $(if ($rule.DescAtStart) { 'at the start' } else { 'at the end' }), $(if ($rule.DescSep) { '"' + $rule.DescSep + '"' } else { 'none' })) } else { 'none' }))
    }
    $settings['Groups'] = ('{0} checked, {1} with a change' -f $rows.Count, ($ready.Count + $err))
    $settings['Result'] = $result
    try { [void](Export-RenameReport -Records $records -WhatIf:$whatIf -Settings $settings) }
    catch { Write-Log ("Rename report failed: {0}" -f $_.Exception.Message) 'WARN' }
}

# ---------- Revert ----------

function Invoke-RenameRevertPick {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Rename reports (*.json;*.csv)|*.json;*.csv|JSON files (*.json)|*.json|CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dlg.Title  = 'Select a rename report (rename_groups_*.json or .csv) to rename the groups back'
    $folder = Get-ReportFolder -Sub 'rename-groups'
    if (Test-Path -LiteralPath $folder) { $dlg.InitialDirectory = $folder }
    if ($dlg.ShowDialog() -eq $true) { Import-RenameRevert -Path $dlg.FileName }
}

function Import-RenameRevert {
    # Fills the preview with the way back: current value = the report's new value,
    # target = the report's old value. Only rows the report marks as Renamed or
    # Reverted are taken - a WhatIf report changed nothing.
    param([Parameter(Mandatory)] [string]$Path)
    if (-not (Test-RenameRowType)) { return }
    $rows = @(); $meta = $null
    try {
        $ext = [System.IO.Path]::GetExtension($Path).ToLower()
        if ($ext -eq '.json') {
            $obj = (Get-Content -LiteralPath $Path -Raw -Encoding UTF8) | ConvertFrom-Json -ErrorAction Stop
            $meta = $obj.info
            $rows = if ($null -ne $obj.data) { @($obj.data) } else { @($obj) }
        } else {
            $first = (Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8)
            $delim = if ($first -match ';') { ';' } elseif ($first -match "`t") { "`t" } else { ',' }
            $rows = @(Import-Csv -LiteralPath $Path -Delimiter $delim -Encoding UTF8)
        }
    }
    catch { Show-Message -Text ("Could not read the file:`n{0}" -f $_.Exception.Message) -Icon Error -Caption 'Rename groups'; return }

    $need = 'ObjectGUID','OldName','NewName','OldSam','NewSam'
    $props = @(); if ($rows.Count -gt 0) { $props = @($rows[0].PSObject.Properties.Name) }
    $missing = @($need | Where-Object { $props -notcontains $_ })
    if ($missing.Count -gt 0) {
        Show-Message -Text ("This is not a rename report - it has no {0} column(s).`n`nPick a rename_groups_*.json or .csv from files\rename-groups." -f ($missing -join ', ')) -Icon Warning -Caption 'Rename groups'; return
    }
    $done = @($rows | Where-Object { [string]$_.Status -in 'Renamed','Reverted' })
    if ($done.Count -eq 0) {
        $why = if ($meta -and $meta.whatIf) { 'It was a WhatIf run, so nothing was renamed.' } else { 'No row in it has the status Renamed or Reverted.' }
        Show-Message -Text ("There is nothing to revert in this report.`n`n{0}" -f $why) -Icon Info -Caption 'Rename groups'; return
    }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($d in $done) {
        $r = New-Object ReOrg.RenameRow
        $r.ObjectGuid = [string]$d.ObjectGUID
        $r.OldDN      = [string]$d.NewDN
        $r.OldName    = [string]$d.NewName;        $r.NewName    = [string]$d.OldName
        $r.OldSam     = [string]$d.NewSam;         $r.NewSam     = [string]$d.OldSam
        $r.OldDisplay = [string]$d.NewDisplayName; $r.NewDisplay = [string]$d.OldDisplayName
        $r.OldDesc    = [string]$d.NewDescription; $r.NewDesc    = [string]$d.OldDescription
        $r.Status     = 'Not checked'
        $list.Add($r)
    }
    Set-RenameMode -Mode 'Revert' -File $Path
    Set-RenameRows -Rows @($list.ToArray() | Sort-Object OldName)
    $skipped = $rows.Count - $done.Count
    $msg = ("Revert loaded: {0} group(s) to rename back{1}. Check in AD, then run with WhatIf first." -f $done.Count, $(if ($skipped) { " ($skipped row(s) in the report were not renamed and are left out)" } else { '' }))
    Set-Status $msg
    if ($script:TxtRenResult) { $script:TxtRenResult.Text = $msg }
    Write-Log ("Rename: revert loaded from '{0}' - {1} group(s)." -f $Path, $done.Count) 'INFO'
}
