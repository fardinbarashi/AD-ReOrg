<#
    Copy members tab: copy members (and optionally the url attribute) from one
    existing group to another. Pick both groups by clicking in one tree: the first
    click fills SourceGroup, the next click TargetGroup. Nothing new is created.
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

# ---- One tree, click to pick: first click -> SourceGroup, next click -> TargetGroup ----

function Build-CopyPickTree {
    # OU -> group leaves (no checkboxes). The Tag carries the fields the shared
    # tree filter expects (Type, DN, Sam, Text) so name/attribute filters work.
    param([Parameter(Mandatory)] [string]$RootDN, [Parameter(Mandatory)] [string]$RootName, [object[]]$Groups)
    if (-not $script:CopyTree) { return }
    $script:CopyTree.Items.Clear()
    $rootLower = $RootDN.ToLower()
    $rootNode = New-OuNode -DN $RootDN -Label $RootName -Type 'OU'
    $rootNode.Tag.Text = $rootNode.Header
    [void]$script:CopyTree.Items.Add($rootNode)
    $map = @{}; $map[$rootLower] = $rootNode

    function Get-OrCreateCopyOu($dn) {
        $dnl = $dn.ToLower()
        if ($dnl -eq $rootLower -or -not $dnl.EndsWith($rootLower)) { return $rootNode }
        if ($map.ContainsKey($dnl)) { return $map[$dnl] }
        $parent = ($dn -split ',', 2)[1]
        $pNode  = Get-OrCreateCopyOu $parent
        $label  = ($dn -split ',', 2)[0] -replace '^OU=', '' -replace '^CN=', ''
        $n = New-OuNode -DN $dn -Label $label -Type 'OU'
        $n.Tag.Text = $n.Header
        [void]$pNode.Items.Add($n); $map[$dnl] = $n
        return $n
    }

    foreach ($g in ($Groups | Sort-Object Name)) {
        $parentDn = ($g.DistinguishedName -split ',', 2)[1]
        $ouNode = Get-OrCreateCopyOu $parentDn
        $leaf = New-OuNode -DN $g.DistinguishedName -Label (Get-GroupTreeLabel -Group $g) -Type 'GroupLeaf' -Name $g.Name -Sam $g.SamAccountName
        $leaf.Tag.Text = $leaf.Header
        if ($g.memberOf -and @($g.memberOf).Count -gt 0) { $leaf.Header.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(180,35,42)) }
        [void]$ouNode.Items.Add($leaf)
    }
    $rootNode.IsExpanded = $true
}

function Update-CMFilter {
    Update-CheckGroupTreeFilter -Tree $script:CopyTree -FilterBox $script:TxtCMFilter `
        -AttrBoxes @($script:TxtCMAttr1, $script:TxtCMAttr2, $script:TxtCMAttr3) `
        -ValBoxes  @($script:TxtCMVal1,  $script:TxtCMVal2,  $script:TxtCMVal3) -Kind 'Group'
}

function Select-CopyGroup {
    # Tree selection: a group leaf fills SourceGroup when it is empty, otherwise TargetGroup.
    $item = $script:CopyTree.SelectedItem
    if (-not $item -or -not $item.Tag -or $item.Tag.Type -ne 'GroupLeaf') { return }
    $pick = [pscustomobject]@{ Name = [string]$item.Tag.Name; Sam = [string]$item.Tag.Sam; DN = [string]$item.Tag.DN }
    if ([string]::IsNullOrWhiteSpace($script:TxtCMSrcName.Text)) {
        $script:CMSource = $pick; $script:TxtCMSrcName.Text = $pick.Name
        $script:TxtCMSrcInfo.Text = $pick.DN
        Set-Status ("SourceGroup: {0}. Click another group for TargetGroup." -f $pick.Name)
    }
    else {
        $script:CMTarget = $pick; $script:TxtCMTgtName.Text = $pick.Name
        $script:TxtCMTgtInfo.Text = $pick.DN
        Set-Status ("TargetGroup: {0}." -f $pick.Name)
    }
    # Deselect so the same group can be clicked again (e.g. after clearing a name).
    try { $item.IsSelected = $false } catch { }
}

function Clear-CopySelection {
    $script:CMSource = $null; $script:CMTarget = $null
    if ($script:TxtCMSrcName) { $script:TxtCMSrcName.Text = '' }
    if ($script:TxtCMTgtName) { $script:TxtCMTgtName.Text = '' }
    if ($script:TxtCMSrcInfo) { $script:TxtCMSrcInfo.Text = '' }
    if ($script:TxtCMTgtInfo) { $script:TxtCMTgtInfo.Text = '' }
}

function Resolve-CopyGroup {
    # Turns the text in a Name box into {Name, Sam, DN}: the picked group when the
    # name still matches, otherwise a lookup by Name / sAMAccountName in the loaded groups.
    param([string]$Text, $Picked)
    $t = ([string]$Text).Trim()
    if (-not $t) { return $null }
    if ($Picked -and ($Picked.Name -eq $t -or $Picked.Sam -eq $t)) { return $Picked }
    foreach ($g in $script:GroupCache.Values) {
        if ($g.Name -ieq $t -or $g.SamAccountName -ieq $t) { return [pscustomobject]@{ Name = [string]$g.Name; Sam = [string]$g.SamAccountName; DN = [string]$g.DistinguishedName } }
    }
    return $null
}

function Invoke-CopyMembersDirect {
    if (-not $script:LoadedRootDN) { Show-Message -Text 'Connect to Active Directory first (Connect & load AD).' -Icon Warning; return }
    $src = Resolve-CopyGroup -Text $script:TxtCMSrcName.Text -Picked $script:CMSource
    $tgt = Resolve-CopyGroup -Text $script:TxtCMTgtName.Text -Picked $script:CMTarget
    if (-not $src) { Show-Message -Text 'Pick a SourceGroup: click a group in the tree (or type its name).' -Icon Warning; return }
    if (-not $tgt) { Show-Message -Text 'Pick a TargetGroup: click a second group in the tree (or type its name).' -Icon Warning; return }
    if ($src.DN -eq $tgt.DN) { Show-Message -Text 'SourceGroup and TargetGroup are the same group.' -Icon Warning; return }
    $sources = @($src); $targets = @($tgt)

    $copyUsers  = [bool]$script:ChkCMUsers.IsChecked
    $copyGroups = [bool]$script:ChkCMGroups.IsChecked
    $copyComps  = [bool]$script:ChkCMComputers.IsChecked
    $copyUrl    = $false   # url copying lives on the Bulk-update / Single group tabs
    $whatIf     = [bool]$script:ChkCMWhatIf.IsChecked
    $allowed = New-Object System.Collections.Generic.List[string]
    if ($copyUsers)  { $allowed.Add('user') }
    if ($copyGroups) { $allowed.Add('group') }
    if ($copyComps)  { $allowed.Add('computer') }
    $doMembers = ($allowed.Count -gt 0)
    if (-not $doMembers) { Show-Message -Text 'Select at least one category (Users/Groups/Computers).' -Icon Warning; return }
    $catText = if ($doMembers) { (@($allowed.ToArray()) -join ', ') } else { '(none - url only)' }

    $server = $null
    if ($script:Config -and $script:Config.domain.server) { $server = [string]$script:Config.domain.server }

    $mode = if ($whatIf) { 'WhatIf (SIMULATION)' } else { 'LIVE RUN' }
    if ((Show-Message -Text ("Copy between existing groups.`n`nSourceGroup : {0}`nTargetGroup : {1}`nCategories  : {2}`nCopy url attribute: {3}`nRun mode    : {4}`n`nNothing new is created. Continue?" -f $src.Name, $tgt.Name, $catText, ($(if($copyUrl){'Yes'}else{'No'})), $mode) -Icon Question -Buttons YesNo -Caption 'Confirm') -ne [System.Windows.MessageBoxResult]::Yes) { Set-Status 'Cancelled.'; return }

    Write-Log ("=== Copy members: '{0}' -> '{1}' [{2}], categories: {3}, url: {4} ===" -f $src.Name, $tgt.Name, $mode, $catText, ($(if($copyUrl){'Yes'}else{'No'}))) 'INFO'
    Set-Status 'Reading source data ...' -Busy

    $allSrc = @()
    if ($doMembers) {
        $srcMembers = @{}
        foreach ($s in $sources) {
            try { foreach ($m in @(Get-GroupMemberObjects -DN $s.DN -Server $server)) { if ($allowed.Contains([string]$m.objectClass)) { $srcMembers[[string]$m.distinguishedName] = $m } } }
            catch { Write-Log ("Could not read members in source group '{0}': {1}" -f $s.Name, $_.Exception.Message) 'WARN' }
        }
        $allSrc = @($srcMembers.Values)
        Write-Log ("Sources hold {0} unique members in the selected categories." -f $allSrc.Count) 'INFO'
    }

    $srcUrls = New-Object System.Collections.Generic.List[string]
    if ($copyUrl) {
        $seenUrl = New-Object System.Collections.Generic.HashSet[string]
        foreach ($s in $sources) {
            try { $sp = @{ Identity = $s.DN; Properties = 'url'; ErrorAction = 'Stop' }; if ($server) { $sp.Server = $server }; foreach ($u in @((Get-ADGroup @sp).url)) { if ($u -and $seenUrl.Add([string]$u)) { $srcUrls.Add([string]$u) } } }
            catch { Write-Log ("Could not read url for source group '{0}': {1}" -f $s.Name, $_.Exception.Message) 'WARN' }
        }
        Write-Log ("Sources hold {0} unique url value(s) to copy." -f $srcUrls.Count) 'INFO'
    }

    if ($doMembers -and $allSrc.Count -eq 0 -and -not $copyUrl) { Set-Status 'Nothing to copy.'; Show-Message -Text 'The source groups have no members in the selected categories.' -Icon Information; return }

    $memberRecords = New-Object System.Collections.Generic.List[object]
    $upnCache = @{}; $addedTotal = 0; $failed = 0; $urlTotal = 0

    $total = $targets.Count; $i = 0
    $script:Pb.Visibility = [System.Windows.Visibility]::Visible
    $script:Pb.Maximum = [Math]::Max($total, 1); $script:Pb.Value = 0

    foreach ($t in $targets) {
        $i++; $script:Pb.Value = $i
        Set-Status ("Target {0} of {1}: {2} ..." -f $i, $total, $t.Name) -Busy

        $existing = @()
        if ($doMembers) {
            try { $existing = @(Get-GroupMemberDNs -DN $t.DN -Server $server) } catch { }
            $toAdd = @($allSrc | Where-Object { $_.distinguishedName -ne $t.DN -and $existing -notcontains $_.distinguishedName })
            if ($whatIf) { Write-Log ("  Would add {0} member(s) to '{1}'." -f $toAdd.Count, $t.Name) 'INFO' }
            elseif ($toAdd.Count -gt 0) {
                try { Add-ADGroupMember -Identity $t.DN -Members @($toAdd | Select-Object -ExpandProperty distinguishedName) -ErrorAction Stop; $addedTotal += $toAdd.Count; Write-Log ("  Added {0} member(s) to '{1}'." -f $toAdd.Count, $t.Name) 'OK' }
                catch { $failed++; Write-Log ("  Could not add members to '{0}': {1}" -f $t.Name, $_.Exception.Message) 'ERROR' }
            }
            else { Write-Log ("  '{0}': nothing new to add (all already present)." -f $t.Name) 'INFO' }

            foreach ($m in $allSrc) {
                $act = if ($m.distinguishedName -eq $t.DN) { 'Skipped (target itself)' } elseif ($existing -contains $m.distinguishedName) { 'Already present' } elseif ($whatIf) { 'Would add' } else { 'Added' }
                $memberRecords.Add([pscustomobject]@{ Account = $m.SamAccountName; UPN = (Get-MemberUpn -Member $m -Cache $upnCache -Server $server); Name = $m.name; Type = $m.objectClass; TargetGroup = $t.Name; TargetSam = $t.Sam; Action = $act; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') })
            }
        }

        if ($copyUrl -and $srcUrls.Count -gt 0) {
            try {
                $tp = @{ Identity = $t.DN; Properties = 'url'; ErrorAction = 'Stop' }; if ($server) { $tp.Server = $server }
                $tg = Get-ADGroup @tp
                $newUrl = New-Object System.Collections.Generic.List[string]; $seen2 = New-Object System.Collections.Generic.HashSet[string]
                foreach ($u in @($tg.url)) { if ($u -and $seen2.Add([string]$u)) { $newUrl.Add([string]$u) } }
                $addedUrls = 0
                foreach ($u in $srcUrls) { if ($seen2.Add($u)) { $newUrl.Add($u); $addedUrls++ } }
                if ($addedUrls -eq 0) { Write-Log ("  '{0}': all url values already present." -f $t.Name) 'INFO' }
                elseif ($whatIf) { Write-Log ("  Would add {0} url value(s) to '{1}'." -f $addedUrls, $t.Name) 'INFO' }
                else { $up = @{ Identity = $t.DN; Replace = @{ url = @($newUrl.ToArray()) }; ErrorAction = 'Stop' }; if ($server) { $up.Server = $server }; Set-ADGroup @up; $urlTotal += $addedUrls; Write-Log ("  Added {0} url value(s) to '{1}'." -f $addedUrls, $t.Name) 'OK' }
            }
            catch { $failed++; Write-Log ("  Could not copy url to '{0}': {1}" -f $t.Name, $_.Exception.Message) 'ERROR' }
        }
    }
    $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed
    Write-Log ("=== Done: {0} member additions, {1} url additions, {2} errors. [{3}] ===" -f $addedTotal, $urlTotal, $failed, $mode) 'OK'

    $memRecs = @($memberRecords.ToArray())
    if ($memRecs.Count -gt 0) {
        try {
            $memRecs = @($memRecs | Sort-Object TargetGroup, Type, Account)
            $folder = Get-ReportFolder -Sub 'copy-members'; $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $mHtml = Join-Path $folder ('copy_members_{0}.html' -f $stamp); $mCsv = Join-Path $folder ('copy_members_{0}.csv' -f $stamp)
            $cols = @(
                @{ Key='Account';     Label='Account (sAMAccountName)' }
                @{ Key='UPN';         Label='UPN' }
                @{ Key='Name';        Label='Name' }
                @{ Key='Type';        Label='Type' }
                @{ Key='TargetGroup'; Label='Target group' }
                @{ Key='TargetSam';   Label='Target sAM' }
                @{ Key='Action';      Label='Action' }
                @{ Key='Timestamp';   Label='Timestamp' }
            )
            $settings = [ordered]@{
                'Run mode'      = $(if ($whatIf) { 'WhatIf - simulation, nothing written' } else { 'Live run' })
                'Source group'  = ('{0}  ({1})' -f $src.Name, $src.DN)
                'Target group'  = ('{0}  ({1})' -f $tgt.Name, $tgt.DN)
                'Categories'    = $catText
                'Members added' = ('{0} added, {1} error(s); existing members kept' -f $addedTotal, $failed)
            }
            $banner = if ($whatIf) { '<div class="banner wi">SIMULATION (WhatIf) - nothing was written to Active Directory.</div>' } else { ('<div class="banner ok">{0} member rows (account + UPN).</div>' -f $memRecs.Count) }
            New-HtmlReport -Title 'Copy members - result' -Columns $cols -Records $memRecs -Path $mHtml -StatusKey 'Action' -Banner $banner -Unit 'member rows' -Settings $settings -HideChangedBox
            Export-DataCsv -Columns $cols -Records $memRecs -Path $mCsv
            Write-Log ("Member report (HTML): {0}" -f $mHtml) 'OK'
            Write-Log ("Member report (CSV) : {0}" -f $mCsv)  'OK'
            if ($script:Config.report.openAfter -and (Test-Path -LiteralPath $mHtml)) { Start-Process $mHtml }
        }
        catch { Write-Log ("Could not create member report: {0}" -f $_.Exception.Message) 'ERROR' }
    }

    Set-Status ("Done: {0} member additions, {1} url additions, {2} errors." -f $addedTotal, $urlTotal, $failed)
    $summary = if ($whatIf) { "Simulation done (WhatIf). Nothing was changed in AD.`n`nClear 'WhatIf' and run again for a live run." }
               else { "Done.`n`nMember additions : $addedTotal`nURL additions    : $urlTotal`nErrors           : $failed`nTarget group     : $($tgt.Name)" }
    Show-Message -Text $summary -Icon Info -Caption 'Result'
}
