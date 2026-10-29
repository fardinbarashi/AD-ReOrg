<#
    Create tab logic: build the preview of the new groups and run the creation
    (copy members by category, copy memberOf, copy the url attribute, nest the new
    group into the source), all WhatIf-aware.
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

# ---------- Target selection ----------

function Get-SelectedTargetTag {
    if (-not $script:OuTree) { return $null }
    $sel = $script:OuTree.SelectedItem
    if ($sel -is [System.Windows.Controls.TreeViewItem] -and $sel.Tag -and $sel.Tag.DN) { return $sel.Tag }
    return $null
}
function Get-SelectedTargetOu {
    $t = Get-SelectedTargetTag
    if (-not $t) { return $null }
    if ($t.Type -eq 'GroupLeaf') { return (($t.DN -split ',', 2)[1]) }
    return [string]$t.DN
}
function Get-SelectedTargetGroup {
    $t = Get-SelectedTargetTag
    if ($t -and $t.Type -eq 'GroupLeaf') { return $t }
    return $null
}
function Update-TargetOuText {
    $t = Get-SelectedTargetTag
    if ($script:TxtTargetOu) {
        if (-not $t) { $script:TxtTargetOu.Text = '(no target selected)' }
        elseif ($t.Type -eq 'GroupLeaf') { $script:TxtTargetOu.Text = ('TARGET GROUP: {0}  ({1})' -f $t.Name, $t.DN) }
        else { $script:TxtTargetOu.Text = $t.DN }
    }
    $script:TargetOuDN = Get-SelectedTargetOu
}

function Update-DescNotesState {
    # Grey out the Description / Notes boxes when "Copy" is selected (they come from
    # the source group); enable them only when "Add" is selected.
    if ($script:TxtDescription -and $script:ChkAddDesc) {
        $addD     = [bool]$script:ChkAddDesc.IsChecked
        $copyAddD = [bool]($script:ChkCopyAddDesc -and $script:ChkCopyAddDesc.IsChecked)
        $script:TxtDescription.IsEnabled = ($addD -or $copyAddD)
        if ($script:TxtDescHint) {
            if ($copyAddD) {
                $script:TxtDescHint.Text = "Your text is added after the description copied from each source group."
                $script:TxtDescHint.Visibility = [System.Windows.Visibility]::Visible
            }
            elseif ($addD) { $script:TxtDescHint.Visibility = [System.Windows.Visibility]::Collapsed }
            else {
                $script:TxtDescHint.Text = if ([bool]$script:ChkCopyDesc.IsChecked) { "Locked - description is copied from each source group. Check the 'Add description' or 'Copy and add description' check box to type your own." } else { "Not applied - check one of the check boxes 'Copy description', 'Add description' or 'Copy and add description'." }
                $script:TxtDescHint.Visibility = [System.Windows.Visibility]::Visible
            }
        }
    }
    if ($script:TxtNotes -and $script:ChkAddNotes) {
        $addN     = [bool]$script:ChkAddNotes.IsChecked
        $copyAddN = [bool]($script:ChkCopyAddNotes -and $script:ChkCopyAddNotes.IsChecked)
        $script:TxtNotes.IsEnabled = ($addN -or $copyAddN)
        if ($script:TxtNotesHint) {
            if ($copyAddN) {
                $script:TxtNotesHint.Text = "Your text is added after the notes copied from each source group."
                $script:TxtNotesHint.Visibility = [System.Windows.Visibility]::Visible
            }
            elseif ($addN) { $script:TxtNotesHint.Visibility = [System.Windows.Visibility]::Collapsed }
            else {
                $script:TxtNotesHint.Text = if ([bool]$script:ChkCopyNotes.IsChecked) { "Locked - notes are copied from each source group. Check the 'Add notes' or 'Copy and add notes' check box to type your own." } else { "Not applied - check one of the check boxes 'Copy notes', 'Add notes' or 'Copy and add notes'." }
                $script:TxtNotesHint.Visibility = [System.Windows.Visibility]::Visible
            }
        }
    }
}

# ---------- Member UPN lookup ----------

function Get-MemberUpn {
    param($Member, [hashtable]$Cache, [string]$Server)
    if ([string]$Member.objectClass -ne 'user') { return '' }
    if ($Member.PSObject.Properties['UserPrincipalName'] -and $Member.UserPrincipalName) { return [string]$Member.UserPrincipalName }
    $dn = [string]$Member.distinguishedName
    if (-not $dn) { return '' }
    if ($Cache.ContainsKey($dn)) { return $Cache[$dn] }
    $upn = ''
    try {
        $p = @{ Identity = $dn; Properties = 'UserPrincipalName'; ErrorAction = 'Stop' }
        if ($Server) { $p.Server = $Server }
        $u = Get-ADUser @p
        $upn = [string]$u.UserPrincipalName
    } catch { }
    $Cache[$dn] = $upn
    return $upn
}

# ---------- Name building ----------

function New-NewGroupName {
    # Applies find/replace to the source name, then adds prefix/suffix.
    param(
        [Parameter(Mandatory)] [string]$OldName,
        [string]$Prefix = '', [string]$Suffix = '',
        [string]$Find = '', [string]$Replace = '', [switch]$UseRegex
    )
    $base = $OldName
    if ($Find) {
        try {
            if ($UseRegex) { $base = [regex]::Replace($base, $Find, $Replace) }
            else           { $base = Get-ReplacedText -Text $base -Find $Find -Replace $Replace }   # case-insensitive
        } catch { $base = $OldName }
    }
    return ('{0}{1}{2}' -f $Prefix, $base, $Suffix).Trim()
}

# ---------- Preview ----------

function Build-Preview {
    $nodes = Get-CheckedGroupNode
    if (@($nodes).Count -eq 0) { Show-Message -Text 'Check the check box of at least one source group in the tree first.' -Icon Warning; return }

    $prefix    = [string]$script:TxtPrefix.Text
    $suffix    = [string]$script:TxtSuffix.Text
    $samSuffix = if ($script:TxtSamSuffix) { [string]$script:TxtSamSuffix.Text } else { '' }
    $find      = if ($script:TxtFind)    { [string]$script:TxtFind.Text }    else { '' }
    $replace   = if ($script:TxtReplace) { [string]$script:TxtReplace.Text } else { '' }
    $useRegex  = if ($script:ChkRegex)   { [bool]$script:ChkRegex.IsChecked } else { $false }

    $prev = @{}
    if ($script:Plans) { foreach ($p in $script:Plans) { $prev[$p.SourceDN] = $p } }
    if (-not $script:Plans) { $script:Plans = New-Object 'System.Collections.ObjectModel.ObservableCollection[ReOrg.GroupPlan]' }
    $script:Plans.Clear()

    $total = @($nodes).Count; $i = 0
    $script:Pb.Visibility = [System.Windows.Visibility]::Visible
    $script:Pb.Maximum = [Math]::Max($total, 1); $script:Pb.Value = 0
    Set-Status 'Building preview ...' -Busy

    foreach ($n in $nodes) {
        $i++; $script:Pb.Value = $i
        if ($i % 5 -eq 0 -or $i -eq $total) { Set-Status "Reading members $i of $total ..." -Busy }

        $count = 0
        try { $count = @(Get-GroupMemberDNs -DN $n.DN -Server $script:LoadedServer).Count } catch { $count = 0 }

        $plan = New-Object ReOrg.GroupPlan
        $plan.Include     = $true
        $plan.SourceName  = $n.Name
        $plan.SourceSam   = $n.Sam
        $plan.SourceDN    = $n.DN
        $plan.MemberCount = $count
        if ($prev.ContainsKey($n.DN) -and $prev[$n.DN].NewName) {
            $plan.NewName = $prev[$n.DN].NewName
            $plan.NewSam  = if ($prev[$n.DN].NewSam) { $prev[$n.DN].NewSam } else { ConvertTo-SamAccountName -Name ($plan.NewName + $samSuffix) }
        } else {
            $plan.NewName = New-NewGroupName -OldName $n.Name -Prefix $prefix -Suffix $suffix -Find $find -Replace $replace -UseRegex:$useRegex
            $plan.NewSam  = ConvertTo-SamAccountName -Name ($plan.NewName + $samSuffix)
        }
        $plan.Status = 'New'
        $script:Plans.Add($plan)
    }

    $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed
    if ($script:GridPreview) { $script:GridPreview.ItemsSource = $script:Plans }
    Update-PreviewCount
    Set-Status ("Preview: {0} groups ready to create." -f @($nodes).Count)
    Write-Log  ("Preview built for {0} groups." -f @($nodes).Count) 'INFO'
}

function Update-PreviewCount {
    param([int]$LiveSel = -1)
    if (-not $script:TxtPreviewCount) { return }
    if (-not $script:Plans -or $script:Plans.Count -eq 0) { $script:TxtPreviewCount.Text = 'No preview built yet - click "Build preview".'; return }
    $sel = @($script:Plans | Where-Object { $_.Include }).Count
    $planTotal = $script:Plans.Count
    $txt = ("Groups to create: {0} of {1}" -f $sel, $planTotal)
    # Warn when the tree selection no longer matches the built preview.
    if ($LiveSel -lt 0) { $LiveSel = @(Get-CheckedGroups $script:GroupTree).Count }
    if ($LiveSel -ne $planTotal) { $txt += ("    —    selection changed to {0}. Click 'Build preview' to refresh." -f $LiveSel) }
    $script:TxtPreviewCount.Text = $txt
}

function Apply-NamePattern {
    Complete-GridEdit $script:GridPreview
    if (-not $script:Plans -or $script:Plans.Count -eq 0) { Build-Preview; return }
    $prefix    = [string]$script:TxtPrefix.Text
    $suffix    = [string]$script:TxtSuffix.Text
    $samSuffix = if ($script:TxtSamSuffix) { [string]$script:TxtSamSuffix.Text } else { '' }
    $find      = if ($script:TxtFind)    { [string]$script:TxtFind.Text }    else { '' }
    $replace   = if ($script:TxtReplace) { [string]$script:TxtReplace.Text } else { '' }
    $useRegex  = if ($script:ChkRegex)   { [bool]$script:ChkRegex.IsChecked } else { $false }
    foreach ($p in $script:Plans) {
        $p.NewName = New-NewGroupName -OldName $p.SourceName -Prefix $prefix -Suffix $suffix -Find $find -Replace $replace -UseRegex:$useRegex
        $p.NewSam  = ConvertTo-SamAccountName -Name ($p.NewName + $samSuffix)
    }
    Update-GridSafe $script:GridPreview
    Set-Status 'Names recalculated (name = find/replace + prefix/suffix, sAM = + sAM suffix).'
}

# ---------- Create ----------

function Get-GroupRunSettings {
    # Collects what the user chose on the Create Groups tab for the report header:
    # where the source groups come from, where the new groups go, the name rule and
    # every action option. Returned as an ordered label -> value dictionary.
    param($Rows, [string]$OuDN, $FixedTarget)
    $txt = { param($c) if ($c) { ([string]$c.Text).Trim() } else { '' } }
    $chk = { param($c) if ($c -and $c.IsChecked) { 'Yes' } else { 'No' } }
    $find = & $txt $script:TxtFind; $repl = & $txt $script:TxtReplace
    $sourceOus = @($Rows | ForEach-Object { ([string]$_.SourceDN -split ',', 2)[1] } | Where-Object { $_ } | Sort-Object -Unique)
    $onlyExist = [bool]$script:ChkOnlyExisting.IsChecked
    $target = if ($onlyExist) {
        if ($FixedTarget) { ('Copy to existing group only: {0}  ({1})' -f $FixedTarget.Name, $FixedTarget.DN) } else { 'Copy to existing groups only (matched per row by name / sAM)' }
    } else { $OuDN }
    $cats = @(); if ($script:ChkCopyUsers.IsChecked) { $cats += 'Users' }; if ($script:ChkCopyGroups.IsChecked) { $cats += 'Groups' }; if ($script:ChkCopyComputers.IsChecked) { $cats += 'Computers' }
    $descMode = if ($script:ChkCopyAddDesc -and $script:ChkCopyAddDesc.IsChecked) { 'Copy and add' } elseif ($script:ChkAddDesc.IsChecked) { 'Add' } elseif ($script:ChkCopyDesc.IsChecked) { 'Copy from source' } else { 'Not applied' }
    $notesMode = if ($script:ChkCopyAddNotes -and $script:ChkCopyAddNotes.IsChecked) { 'Copy and add' } elseif ($script:ChkAddNotes.IsChecked) { 'Add' } elseif ($script:ChkCopyNotes.IsChecked) { 'Copy from source' } else { 'Not applied' }
    $descText = & $txt $script:TxtDescription; $notesText = & $txt $script:TxtNotes
    $s = [ordered]@{}
    $s['Run mode']               = if ($script:ChkWhatIf.IsChecked) { 'WhatIf - simulation, nothing written' } else { 'Live run' }
    $s['Source groups']          = ('{0} group(s)' -f @($Rows).Count)
    $s['Source OU(s)']           = ($sourceOus -join ' | ')
    $s['Target']                 = $target
    $s['Find in name -> Replace']= if ($find) { ('"{0}" -> "{1}"{2}' -f $find, $repl, $(if ($script:ChkRegex.IsChecked) { '  (regular expression)' } else { '' })) } else { '(none)' }
    $s['Prefix / Suffix']        = ('{0} / {1}' -f $(if (& $txt $script:TxtPrefix) { & $txt $script:TxtPrefix } else { '-' }), $(if (& $txt $script:TxtSuffix) { & $txt $script:TxtSuffix } else { '-' }))
    $s['sAM suffix']             = & $txt $script:TxtSamSuffix
    $s['Scope / Category']       = ('{0} / {1}' -f [string]$script:CmbScope.SelectedItem.Content, [string]$script:CmbCategory.SelectedItem.Content)
    $s['Copy members']           = if ($cats.Count) { $cats -join ', ' } else { 'None' }
    $s['Copy memberOf']          = & $chk $script:ChkCopyMemberOf
    $s['Attributes to copy']     = if (& $txt $script:TxtCopyAttrs) { & $txt $script:TxtCopyAttrs } else { 'None' }
    $s['Description']            = if ($descText -and $descMode -ne 'Copy from source' -and $descMode -ne 'Not applied') { ('{0}: "{1}"' -f $descMode, $descText) } else { $descMode }
    $s['Notes (info)']           = if ($notesText -and $notesMode -ne 'Copy from source' -and $notesMode -ne 'Not applied') { ('{0}: "{1}"' -f $notesMode, $notesText) } else { $notesMode }
    $s['Nest new group in source'] = & $chk $script:ChkNest
    $s['Continue on errors']     = & $chk $script:ChkContinueOnError
    return $s
}

function Invoke-CreateGroups {
    Complete-GridEdit $script:GridPreview
    if (-not $script:Plans -or $script:Plans.Count -eq 0) {
        Show-Message -Text 'Build the preview first (the "Build preview" button).' -Icon Warning; return
    }

    $desc       = ([string]$script:TxtDescription.Text).Trim()
    $notes      = [string]$script:TxtNotes.Text
    $copyDesc    = [bool]$script:ChkCopyDesc.IsChecked
    $addDesc     = [bool]$script:ChkAddDesc.IsChecked
    $copyAddDesc = [bool]($script:ChkCopyAddDesc -and $script:ChkCopyAddDesc.IsChecked)
    $copyNotes   = [bool]$script:ChkCopyNotes.IsChecked
    $addNotes    = [bool]$script:ChkAddNotes.IsChecked
    $copyAddNotes= [bool]($script:ChkCopyAddNotes -and $script:ChkCopyAddNotes.IsChecked)
    $scope      = [string]($script:CmbScope.SelectedItem.Content)
    $category   = [string]($script:CmbCategory.SelectedItem.Content)
    $copyUsers  = [bool]$script:ChkCopyUsers.IsChecked
    $copyGroups = [bool]$script:ChkCopyGroups.IsChecked
    $copyComps  = [bool]$script:ChkCopyComputers.IsChecked
    $copy       = ($copyUsers -or $copyGroups -or $copyComps)
    $nest       = [bool]$script:ChkNest.IsChecked
    $copyMemOf  = [bool]$script:ChkCopyMemberOf.IsChecked
    # Attributes to copy from source -> new group (free-text list, e.g. "url, extensionAttribute1").
    $copyAttrsRaw = if ($script:TxtCopyAttrs) { [string]$script:TxtCopyAttrs.Text } else { '' }
    $blockedAttrs = @('samaccountname','distinguishedname','dn','cn','name','objectguid','objectsid','objectclass','grouptype','instancetype','member','memberof','description','info')
    $seenAttr   = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $copyAttrs  = New-Object System.Collections.Generic.List[string]
    foreach ($a in ($copyAttrsRaw -split '[,;\s]+')) {
        $a = $a.Trim()
        if (-not $a) { continue }
        if ($blockedAttrs -contains $a.ToLower()) { Write-Log ("Attribute '{0}' is skipped (name/membership/description/notes are handled elsewhere)." -f $a) 'WARN'; continue }
        if ($seenAttr.Add($a)) { $copyAttrs.Add($a) }
    }
    $copyAny    = ($copyAttrs.Count -gt 0)
    $onlyExist  = [bool]$script:ChkOnlyExisting.IsChecked
    $whatIf     = [bool]$script:ChkWhatIf.IsChecked
    $continueOnErr = if ($script:ChkContinueOnError) { [bool]$script:ChkContinueOnError.IsChecked } else { $false }

    Update-TargetOuText
    $fixedTarget = Get-SelectedTargetGroup
    $ouDN = ''
    if (-not $onlyExist) {
        $ouDN = $script:TargetOuDN
        if ([string]::IsNullOrWhiteSpace($ouDN)) { Show-Message -Text 'Select a target OU in the middle tree where the groups will be created.' -Icon Warning; return }
    }
    if ($onlyExist -and -not ($copy -or $nest -or $copyMemOf -or $copyAny)) {
        Show-Message -Text 'Copy-only mode is on but no action (members / nesting / memberOf / attributes) is selected. Nothing to copy.' -Icon Warning; return
    }

    $allowed = New-Object System.Collections.Generic.List[string]
    if ($copyUsers)  { $allowed.Add('user') }
    if ($copyGroups) { $allowed.Add('group') }
    if ($copyComps)  { $allowed.Add('computer') }
    $catText = if ($copy) { (@($allowed.ToArray()) -join ', ') } else { '(none)' }

    $server = $null
    if ($script:Config -and $script:Config.domain.server) { $server = [string]$script:Config.domain.server }

    $rows = @($script:Plans | Where-Object { $_.Include })
    if ($rows.Count -eq 0) { Show-Message -Text 'No row is checked in the preview.' -Icon Warning; return }

    $names = @($rows | ForEach-Object { ([string]$_.NewName).Trim() })
    if ($names | Where-Object { -not $_ }) { Show-Message -Text 'At least one row has no new group name. Fill them in or uncheck the row.' -Icon Warning; return }

    foreach ($r in $rows) {
        $ts = ([string]$r.NewSam).Trim()
        $r.NewSam = if ($ts) { ConvertTo-SamAccountName -Name $ts } else { ConvertTo-SamAccountName -Name $r.NewName }
    }

    if (-not $onlyExist -and -not $continueOnErr) {
        $dupe = $names | Group-Object | Where-Object { $_.Count -gt 1 } | Select-Object -First 1
        if ($dupe) { Show-Message -Text ("The new name '{0}' occurs more than once. Names must be unique within the target OU.`n`nTip: check the 'Continue on errors' check box to create the rest and get a report of the clashes." -f $dupe.Name) -Icon Warning; return }
        $sams = @($rows | ForEach-Object { [string]$_.NewSam })
        $dupSam = $sams | Group-Object | Where-Object { $_.Count -gt 1 } | Select-Object -First 1
        if ($dupSam) { Show-Message -Text ("sAMAccountName '{0}' occurs more than once. Each group needs a unique sAM (edit the 'New sAM' column)." -f $dupSam.Name) -Icon Warning; return }
        $samCollision = @($rows | Where-Object { ([string]$_.NewSam).ToLower() -eq ([string]$_.SourceSam).ToLower() })
        if ($samCollision.Count -gt 0) {
            $ex = ($samCollision | Select-Object -First 3 | ForEach-Object { $_.SourceSam }) -join ', '
            Show-Message -Text ("{0} group(s) have the same sAMAccountName as the source (e.g. {1}).`n`nsAMAccountName must be unique across the domain - two groups can never share it, even in different OUs. The name (CN) may be identical.`n`nTo keep the same name but get a unique sAM: type a value in the 'sAM suffix' field (e.g. '_NEW') and click 'Recalculate names', or edit the 'New sAM' column directly." -f $samCollision.Count, $ex) -Icon Warning -Caption 'sAM must be unique'
            Set-Status 'Cancelled: sAMAccountName collides with the source group.'
            return
        }
    }

    $mode = if ($whatIf) { 'WhatIf (SIMULATION - nothing changes)' } else { 'LIVE RUN' }
    $verb = if ($onlyExist) { 'copy to' } else { 'create' }
    $head = if ($onlyExist) {
        if ($fixedTarget) { "Mode       : Copy to existing only`nTARGET     : $($fixedTarget.Name)  ($($fixedTarget.Sam))" }
        else              { "Mode       : Copy to existing only (matched per row by name/sAM)" }
    } else { "Target OU  : $ouDN`nScope      : $scope`nCategory   : $category" }
    $confirm = @"
You are about to $verb $($rows.Count) group(s).

$head
Copy members : $catText
Nest new in source : $nest
Copy memberOf : $copyMemOf
Copy attributes : $(if ($copyAny) { $copyAttrs -join ', ' } else { '(none)' })
Run mode   : $mode

Continue?
"@
    if ((Show-Message -Text $confirm -Icon Question -Buttons YesNo -Caption 'Confirm') -ne [System.Windows.MessageBoxResult]::Yes) { Set-Status 'Cancelled.'; return }

    Write-Log ("=== Start: {0} groups [{1}] ===" -f $rows.Count, $mode) 'INFO'
    Set-Status 'Working ...' -Busy

    $records       = New-Object System.Collections.Generic.List[object]
    $memberRecords = New-Object System.Collections.Generic.List[object]
    $upnCache      = @{}
    $batchSams     = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $created = 0; $failed = 0; $copiedTotal = 0; $nested = 0; $memOfTotal = 0; $attrTotal = 0

    $total = $rows.Count; $i = 0
    $script:Pb.Visibility = [System.Windows.Visibility]::Visible
    $script:Pb.Maximum = [Math]::Max($total, 1); $script:Pb.Value = 0

    foreach ($row in $rows) {
        $i++; $script:Pb.Value = $i
        $newName = ([string]$row.NewName).Trim()
        $typedSam = ([string]$row.NewSam).Trim()
        $newSam = if ($typedSam) { ConvertTo-SamAccountName -Name $typedSam } else { ConvertTo-SamAccountName -Name $newName }
        $row.NewSam = $newSam
        Set-Status ("Processing {0} ({1} of {2}) ..." -f $newName, $i, $total) -Busy

        $rowStatus = ''; $copiedCount = 0; $nestedFlag = 'No'; $memOfFlag = 'No'; $attrFlag = 'No'
        $target = $null; $ok = $true

        # Resolve the description / notes to apply (from source when "Copy", from
        # the box when "Add", otherwise none).
        $srcCached = $script:GroupCache[$row.SourceDN]
        if (-not $srcCached -and ($copyDesc -or $copyNotes)) {
            try { $csp = @{ Identity = $row.SourceDN; Properties = 'description','info'; ErrorAction = 'Stop' }; if ($server) { $csp.Server = $server }; $srcCached = Get-ADGroup @csp } catch { }
        }
        $srcDesc = if ($srcCached) { [string]$srcCached.Description } else { '' }
        $srcInfo = if ($srcCached) { [string]$srcCached.info } else { '' }
        $descVal = if ($copyAddDesc) { @($srcDesc, $desc | Where-Object { $_ -and ([string]$_).Trim() }) -join "`n" }
                   elseif ($copyDesc) { $srcDesc } elseif ($addDesc) { $desc } else { '' }
        $notesVal = if ($copyAddNotes) { @($srcInfo, $notes | Where-Object { $_ -and ([string]$_).Trim() }) -join "`n" }
                    elseif ($copyNotes) { $srcInfo } elseif ($addNotes) { $notes } else { '' }

        # Stops two rows from resolving to the same sAM within this batch (common
        # when long names truncate to the same 20-char sAM) - it would otherwise copy
        # into the wrong group.
        $dupInBatch = $false
        if (-not $onlyExist) { if (-not $batchSams.Add($newSam)) { $dupInBatch = $true } }

        # 1) Create or find the target group.
        if ($dupInBatch) {
            $ok = $false; $failed++
            $rowStatus = 'Error: duplicate sAM in this batch'
            Write-Log ("Duplicate sAM '{0}' within this run - skipped '{1}' (give it a unique sAM)." -f $newSam, $newName) 'ERROR'
        }
        else {
        try {
          if ($onlyExist) {
            $found = $null
            if ($fixedTarget) {
                try { $gp = @{ Identity = $fixedTarget.DN; ErrorAction = 'Stop' }; if ($server) { $gp.Server = $server }; $found = Get-ADGroup @gp } catch { $found = $null }
            } else {
                try { $gp = @{ LDAPFilter = "(|(sAMAccountName=$newSam)(name=$newName))"; ErrorAction = 'Stop' }; if ($server) { $gp.Server = $server }; $found = Get-ADGroup @gp | Select-Object -First 1 } catch { $found = $null }
            }
            if (-not $found) { $ok = $false; $rowStatus = 'Not found (nothing created)'; Write-Log ("Copy mode: target group '{0}' (sAM '{1}') not found - skipped." -f $newName, $newSam) 'WARN' }
            elseif ($found.DistinguishedName -eq $row.SourceDN) { $ok = $false; $rowStatus = 'Error: target = source group'; Write-Log ("Copy mode: target equals source '{0}' - skipped." -f $row.SourceName) 'ERROR' }
            else { $target = $found; $rowStatus = 'Existing (copy)'; Write-Log ("Copy mode: copying into existing group '{0}' ({1})." -f $found.Name, $found.DistinguishedName) 'INFO' }
          }
          else {
            $existing = $null
            try { $gp = @{ LDAPFilter = "(sAMAccountName=$newSam)"; ErrorAction = 'Stop' }; if ($server) { $gp.Server = $server }; $existing = Get-ADGroup @gp | Select-Object -First 1 } catch { $existing = $null }
            if ($existing) {
                if ($existing.DistinguishedName -eq $row.SourceDN) { $ok = $false; $rowStatus = 'Error: sAM = source group (use a unique sAM)'; Write-Log ("Skipped '{0}': sAMAccountName '{1}' is the source group's own." -f $newName, $newSam) 'ERROR' }
                else { $target = $existing; $rowStatus = 'Already exists (sAM taken)'; Write-Log ("A group with sAM '{0}' already exists: {1}" -f $newSam, $existing.DistinguishedName) 'WARN' }
            }
            elseif ($whatIf) {
                $rowStatus = 'Would create'
                Write-Log ("Would create: '{0}' (sAM '{1}', {2}/{3}) in {4}" -f $newName, $newSam, $scope, $category, $ouDN) 'INFO'
            }
            else {
                $np = @{ Name = $newName; SamAccountName = $newSam; GroupScope = $scope; GroupCategory = $category; Path = $ouDN; PassThru = $true; ErrorAction = 'Stop' }
                if ($descVal)         { $np.Description = $descVal }
                if ($notesVal.Trim()) { $np.OtherAttributes = @{ info = $notesVal } }
                if ($server)          { $np.Server = $server }
                $target = New-ADGroup @np
                $created++; $rowStatus = 'Created'
                Write-Log ("Created group '{0}' in {1}" -f $newName, $ouDN) 'OK'
            }
          }
        }
        catch { $ok = $false; $failed++; $rowStatus = 'Error: ' + $_.Exception.Message; Write-Log ("Could not create/find '{0}': {1}" -f $newName, $_.Exception.Message) 'ERROR' }
        }

        # 1b) Apply description / notes to an EXISTING target (new groups already got
        #     them at creation time).
        if ($ok -and $target -and $rowStatus -ne 'Created') {
            if ($whatIf) {
                if ($descVal)  { Write-Log ("  Would set description on '{0}'." -f $newName) 'INFO' }
                if ($notesVal.Trim()) { Write-Log ("  Would set notes (info) on '{0}'." -f $newName) 'INFO' }
            } else {
                if ($descVal)  { try { $ap = @{ Identity = $target.DistinguishedName; Replace = @{ description = $descVal }; ErrorAction = 'Stop' }; if ($server) { $ap.Server = $server }; Set-ADGroup @ap; Write-Log ("  Set description on '{0}'." -f $newName) 'OK' } catch { Write-Log ("  Could not set description on '{0}': {1}" -f $newName, $_.Exception.Message) 'WARN' } }
                if ($notesVal.Trim()) { try { $ap = @{ Identity = $target.DistinguishedName; Replace = @{ info = $notesVal }; ErrorAction = 'Stop' }; if ($server) { $ap.Server = $server }; Set-ADGroup @ap; Write-Log ("  Set notes (info) on '{0}'." -f $newName) 'OK' } catch { Write-Log ("  Could not set notes on '{0}': {1}" -f $newName, $_.Exception.Message) 'WARN' } }
            }
        }

        # 2) Copy members (filtered by category).
        if ($ok -and $copy) {
            try {
                $allMembers = @(Get-GroupMemberObjects -DN $row.SourceDN -Server $server)
                $members = @($allMembers | Where-Object { $allowed.Contains([string]$_.objectClass) })
                $toAddDNs = @()
                if ($members.Count -eq 0) { Write-Log ("  No members in selected categories [{0}] in '{1}'." -f $catText, $row.SourceName) 'INFO' }
                elseif ($whatIf -or -not $target) { $copiedCount = $members.Count; Write-Log ("  Would copy {0} member(s) from '{1}' to '{2}'." -f $members.Count, $row.SourceName, $newName) 'INFO' }
                else {
                    $existingM = @(); try { $existingM = @(Get-GroupMemberDNs -DN $target.DistinguishedName -Server $server) } catch { }
                    $toAdd = @($members | Where-Object { $existingM -notcontains $_.distinguishedName })
                    $toAddDNs = @($toAdd | Select-Object -ExpandProperty distinguishedName)
                    if ($toAdd.Count -gt 0) { Add-ADGroupMember -Identity $target.DistinguishedName -Members $toAddDNs -ErrorAction Stop; $copiedCount = $toAdd.Count; Write-Log ("  Copied {0} member(s) to '{1}'." -f $toAdd.Count, $newName) 'OK' }
                    else { Write-Log ("  All members already present in '{0}'." -f $newName) 'INFO' }
                }
                $copiedTotal += $copiedCount
                foreach ($m in $members) {
                    $act = if ($whatIf -or -not $target) { 'Would add' } elseif ($toAddDNs -contains [string]$m.distinguishedName) { 'Added' } else { 'Already present' }
                    $memberRecords.Add([pscustomobject]@{ Account = $m.SamAccountName; UPN = (Get-MemberUpn -Member $m -Cache $upnCache -Server $server); Name = $m.name; Type = $m.objectClass; TargetGroup = $newName; TargetSam = $newSam; SourceGroup = $row.SourceName; Action = $act; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') })
                }
            }
            catch { Write-Log ("  Could not copy members for '{0}': {1}" -f $newName, $_.Exception.Message) 'ERROR'; if ($rowStatus -notlike 'Error:*') { $rowStatus += ' (member error)' } }
        }

        # 3) Nest the new group into the source (inherit permissions).
        if ($ok -and $nest) {
            try {
                if ($whatIf -or -not $target) { $nestedFlag = 'Would nest'; Write-Log ("  Would add '{0}' as a member of '{1}' (inherit permissions)." -f $newName, $row.SourceName) 'INFO' }
                elseif ($target.DistinguishedName -eq $row.SourceDN) { $nestedFlag = 'No (same group)'; Write-Log ("  Skipped nesting: target is the same group as the source ('{0}')." -f $row.SourceName) 'WARN' }
                else { Add-ADGroupMember -Identity $row.SourceDN -Members $target.DistinguishedName -ErrorAction Stop; $nestedFlag = 'Yes'; $nested++; Write-Log ("  Added '{0}' as a member of '{1}' (inherit permissions)." -f $newName, $row.SourceName) 'OK' }
            }
            catch { $nestedFlag = 'Error'; Write-Log ("  Could not nest '{0}' in '{1}': {2}" -f $newName, $row.SourceName, $_.Exception.Message) 'ERROR'; if ($rowStatus -notlike 'Error:*') { $rowStatus += ' (nesting error)' } }
        }

        # 4) Copy the source group's memberOf to the new group.
        if ($ok -and $copyMemOf) {
            try {
                $sp = @{ Identity = $row.SourceDN; Properties = 'memberOf'; ErrorAction = 'Stop' }; if ($server) { $sp.Server = $server }
                $srcGrp = Get-ADGroup @sp
                # Only real DN values; drop any null/empty entries.
                $parents = @($srcGrp.memberOf | Where-Object { $_ } | ForEach-Object { [string]$_ })
                if ($parents.Count -eq 0) { $memOfFlag = 'No (source has none)'; Write-Log ("  Source group '{0}' is not a member of any group - no memberOf to copy." -f $row.SourceName) 'INFO' }
                elseif ($whatIf -or -not $target) { $memOfFlag = ('Would copy {0}' -f $parents.Count); Write-Log ("  Would add '{0}' to {1} group(s) (source memberOf)." -f $newName, $parents.Count) 'INFO' }
                else {
                    $targetDn = [string]$target.DistinguishedName
                    $added = 0; $skipped = 0
                    foreach ($pdn in $parents) {
                        # Never add the group to itself.
                        if ($pdn -eq $targetDn) { $skipped++; continue }
                        try { Add-ADGroupMember -Identity $pdn -Members $targetDn -ErrorAction Stop; $added++ }
                        catch {
                            $msg = $_.Exception.Message
                            if ($msg -match '(?i)already a member') { $skipped++; Write-Log ("    Already a member of '{0}' - skipped." -f (($pdn -split ',',2)[0] -replace '^CN=','')) 'INFO' }
                            else { Write-Log ("    Could not add '{0}' to '{1}': {2}" -f $newName, (($pdn -split ',',2)[0] -replace '^CN=',''), $msg) 'WARN' }
                        }
                    }
                    $memOfFlag = ('Yes ({0}/{1})' -f $added, $parents.Count); $memOfTotal += $added
                    Write-Log ("  Added '{0}' to {1} of {2} group(s) (source memberOf; {3} skipped)." -f $newName, $added, $parents.Count, $skipped) 'OK'
                }
            }
            catch { $memOfFlag = 'Error'; Write-Log ("  Could not copy memberOf for '{0}': {1}" -f $newName, $_.Exception.Message) 'ERROR'; if ($rowStatus -notlike 'Error:*') { $rowStatus += ' (memberOf error)' } }
        }

        # 5) Copy the requested attributes (e.g. url) from the source group.
        if ($ok -and $copyAny) {
            $attrParts = New-Object System.Collections.Generic.List[string]
            foreach ($attr in $copyAttrs) {
                try {
                    $sp = @{ Identity = $row.SourceDN; Properties = $attr; ErrorAction = 'Stop' }; if ($server) { $sp.Server = $server }
                    $srcVals = @(@((Get-ADGroup @sp).$attr) | Where-Object { $null -ne $_ -and "$_" -ne '' })
                    if ($srcVals.Count -eq 0) { $attrParts.Add("${attr}: none"); Write-Log ("  Source '{0}' has no '{1}'." -f $row.SourceName, $attr) 'INFO'; continue }
                    if ($whatIf -or -not $target) { $attrParts.Add(("{0}: would({1})" -f $attr, $srcVals.Count)); Write-Log ("  Would copy {0} '{1}' value(s) to '{2}'." -f $srcVals.Count, $attr, $newName) 'INFO'; continue }
                    $tp = @{ Identity = $target.DistinguishedName; Properties = $attr; ErrorAction = 'Stop' }; if ($server) { $tp.Server = $server }
                    $tg = Get-ADGroup @tp
                    $merged = New-Object System.Collections.Generic.List[object]; $seenV = New-Object System.Collections.Generic.HashSet[string]
                    foreach ($v in @($tg.$attr)) { if ($null -ne $v -and "$v" -ne '' -and $seenV.Add([string]$v)) { $merged.Add($v) } }
                    $added = 0
                    foreach ($v in $srcVals) { if ($seenV.Add([string]$v)) { $merged.Add($v); $added++ } }
                    if ($added -gt 0) {
                        $up = @{ Identity = $target.DistinguishedName; Replace = @{ $attr = @($merged.ToArray()) }; ErrorAction = 'Stop' }; if ($server) { $up.Server = $server }
                        Set-ADGroup @up
                        $attrParts.Add(("{0}: {1}" -f $attr, $added)); $attrTotal += $added
                        Write-Log ("  Copied {0} '{1}' value(s) to '{2}'." -f $added, $attr, $newName) 'OK'
                    }
                    else { $attrParts.Add("${attr}: present"); Write-Log ("  All '{0}' values already present in '{1}'." -f $attr, $newName) 'INFO' }
                }
                catch { $attrParts.Add("${attr}: error"); Write-Log ("  Could not copy '{0}' for '{1}': {2}" -f $attr, $newName, $_.Exception.Message) 'ERROR'; if ($rowStatus -notlike 'Error:*') { $rowStatus += " ($attr error)" } }
            }
            $attrFlag = ($attrParts -join ', ')
        }

        $row.Status = $rowStatus
        $records.Add([pscustomobject]@{
            SourceGroup    = $row.SourceName
            NewGroup       = $newName
            NewSam         = $newSam
            Scope          = $scope
            Category       = $category
            Description    = $descVal
            Notes          = $notesVal
            TargetOu       = $ouDN
            SourceMembers  = $row.MemberCount
            CopiedMembers  = $copiedCount
            Categories     = $catText
            NestedInSource = $nestedFlag
            MemberOfCopied = $memOfFlag
            AttributesCopied = $attrFlag
            Status         = $rowStatus
            Timestamp      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        })
    }

    $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed
    Update-GridSafe $script:GridPreview
    $script:LastRunRecords = @($records.ToArray())
    Write-Log ("=== Done: {0} created, {1} errors, {2} members copied, {3} nested, {4} memberOf, {5} attribute value(s). [{6}] ===" -f $created, $failed, $copiedTotal, $nested, $memOfTotal, $attrTotal, $mode) 'OK'

    $runSettings = Get-GroupRunSettings -Rows $rows -OuDN $ouDN -FixedTarget $fixedTarget
    try {
        $paths = New-GroupCreationReport -Records $script:LastRunRecords -WhatIf:$whatIf -Settings $runSettings
        Write-Log ("Report (HTML): {0}" -f $paths.Html) 'OK'
        Write-Log ("Report (CSV) : {0}" -f $paths.Csv)  'OK'
        Set-Status ("Done: {0} created, {1} errors. Report: {2}" -f $created, $failed, $paths.Html)
        if ($script:Config.report.openAfter -and (Test-Path -LiteralPath $paths.Html)) { Start-Process $paths.Html }
    }
    catch { Write-Log ("Could not create report: {0}" -f $_.Exception.Message) 'ERROR'; Set-Status ("Done: {0} created, {1} errors. (Report error - see log.)" -f $created, $failed) }

    $memRecs = @($memberRecords.ToArray())
    if ($memRecs.Count -gt 0) {
        try {
            $memRecs = @($memRecs | Sort-Object TargetGroup, Type, Account)
            $folder = Get-ReportFolder -Sub 'create-groups'; $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $mHtml = Join-Path $folder ('members_{0}.html' -f $stamp); $mCsv = Join-Path $folder ('members_{0}.csv' -f $stamp)
            $mCols = @(
                @{ Key='Account';     Label='Account (sAMAccountName)' }
                @{ Key='UPN';         Label='UPN' }
                @{ Key='Name';        Label='Name' }
                @{ Key='Type';        Label='Type' }
                @{ Key='TargetGroup'; Label='Target group' }
                @{ Key='TargetSam';   Label='Target sAM' }
                @{ Key='SourceGroup'; Label='Source group' }
                @{ Key='Action';      Label='Action' }
                @{ Key='Timestamp';   Label='Timestamp' }
            )
            New-HtmlReport -Title 'Member report - copied members' -Columns $mCols -Records $memRecs -Path $mHtml -StatusKey 'Action' -Banner ('<div class="banner ok">{0} member rows (account + UPN).</div>' -f $memRecs.Count) -Unit 'member rows' -Settings $runSettings -HideChangedBox
            Export-DataCsv -Columns $mCols -Records $memRecs -Path $mCsv
            Write-Log ("Member report (HTML): {0}" -f $mHtml) 'OK'
            Write-Log ("Member report (CSV) : {0}" -f $mCsv)  'OK'
        }
        catch { Write-Log ("Could not create member report: {0}" -f $_.Exception.Message) 'ERROR' }
    }

    # Failures report - the groups that could not be created, for manual fixing.
    $failedRows = @($script:LastRunRecords | Where-Object { $_.Status -match '(?i)error' -or $_.Status -like 'Not found*' })
    $failedReportPath = ''
    if ($failedRows.Count -gt 0) {
        try {
            $folder = Get-ReportFolder -Sub 'create-groups'; $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $fHtml = Join-Path $folder ('failed_{0}.html' -f $stamp); $fCsv = Join-Path $folder ('failed_{0}.csv' -f $stamp)
            $fCols = @(
                @{ Key='SourceGroup'; Label='Source group' }
                @{ Key='NewGroup';    Label='New name' }
                @{ Key='NewSam';      Label='New sAMAccountName' }
                @{ Key='Scope';       Label='Scope' }
                @{ Key='Category';    Label='Category' }
                @{ Key='TargetOu';    Label='Target OU' }
                @{ Key='Status';      Label='Reason / status' }
                @{ Key='Timestamp';   Label='Timestamp' }
            )
            $fBanner = ('<div class="banner no">{0} group(s) could not be created and need manual attention (often duplicate sAMAccountName from truncated names - give each a unique New sAM).</div>' -f $failedRows.Count)
            New-HtmlReport -Title 'Failed groups - fix manually' -Columns $fCols -Records $failedRows -Path $fHtml -StatusKey 'Status' -Banner $fBanner -Unit 'failed groups' -Settings $runSettings -HideChangedBox
            Export-DataCsv -Columns $fCols -Records $failedRows -Path $fCsv
            $failedReportPath = $fHtml
            Write-Log ("Failures report (HTML): {0}" -f $fHtml) 'OK'
            Write-Log ("Failures report (CSV) : {0}" -f $fCsv)  'OK'
            if ($script:Config.report.openAfter -and (Test-Path -LiteralPath $fHtml)) { Start-Process $fHtml }
        }
        catch { Write-Log ("Could not create failures report: {0}" -f $_.Exception.Message) 'ERROR' }
    }

    $failLine = if ($failedRows.Count -gt 0) { "`nSee failures report: $failedReportPath" } else { '' }
    $summary = if ($whatIf) { "Simulation done (WhatIf). Nothing was changed in AD.`n`nClear 'WhatIf' and run again for a live run.$failLine" }
               elseif ($onlyExist) { "Done (copy-only, no new groups).`n`nErrors : $failed`nMembers copied : $copiedTotal`nNested : $nested`nmemberOf added : $memOfTotal`nattribute values copied : $attrTotal$failLine" }
               else { "Done.`n`nGroups created : $created`nErrors : $failed`nMembers copied : $copiedTotal`nNested : $nested`nmemberOf added : $memOfTotal`nattribute values copied : $attrTotal$failLine" }
    Show-Message -Text $summary -Icon Info -Caption 'Result'
}
