<#
    Bulk-update Attributes tab: find a text pattern in ANY AD group attribute and
    write a transformed value. For every group whose attribute contains the source
    pattern, the value is copied with source -> target replaced (case-insensitive,
    literal text):
      * multi-valued attribute (e.g. url)      -> the new value is ADDED, old kept
      * single-valued attribute (e.g. info)    -> the value is REPLACED
    Whether an attribute is multi-valued is read from the AD schema (cached).
    A value that already exists on the group is skipped. Scan builds a preview;
    Apply writes it (or only simulates it when the WhatIf check box is checked) and always writes
    an HTML/CSV/JSON report to files\bulk-update or files\single-group-update.

    Typical use: attribute "url", source "OU=OLD," -> target "OU=NEW," adds a copy
    of every LDAP query string that points at the new OU.

    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

# Columns used by the preview grid and the exported report.
function Get-UrlModColumns {
    return @(
        @{ Key = 'Group';     Label = 'Group' }
        @{ Key = 'Status';    Label = 'Status' }
        @{ Key = 'SourceUrl'; Label = 'Current value (matched)' }
        @{ Key = 'NewUrl';    Label = 'New value' }
        @{ Key = 'DN';        Label = 'Distinguished name' }
    )
}

function ConvertTo-LdapFilterText {
    # Escapes a literal for use inside an LDAP filter (RFC 4515).
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return $Text.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace("`0", '\00')
}

function Get-UrlModContext {
    # The controls of the Update attributes tool. (The former "Single group" tab is
    # merged into it: check one group to see its attribute list, many for bulk.)
    # -Single is accepted for compatibility and ignored.
    param([switch]$Single)
    return @{ Single = $false; Key = 'Bulk'; Caption = 'Update attributes'
              Attr = $script:TxtUrlAttr; Src = $script:TxtUrlSourceOu; Dst = $script:TxtUrlTargetOu
              Grid = $script:GridUrl; Count = $script:TxtUrlCount; WhatIf = $script:ChkUrlWhatIf; AddNew = $script:ChkUrlAddNew; NoSep = $script:ChkUrlNoSep; Sep = $script:TxtUrlSep; ReplaceOld = $script:ChkUrlReplaceOld
              Tree = $script:SelTree; TreeCount = $null; AttrGrid = $script:GridUrlAttrs; NonEmpty = $script:ChkUrlNonEmpty }
}

function Get-UrlModInputs {
    # Reads the input boxes of one tab. Attribute and source pattern are required;
    # the target groups are the ones checked in the tab's tree (by DN).
    param($Ctx)
    $attr = if ($Ctx.Attr) { ([string]$Ctx.Attr.Text).Trim() } else { '' }
    $srcP = if ($Ctx.Src)  { ([string]$Ctx.Src.Text).Trim() }  else { '' }
    $dstP = if ($Ctx.Dst)  { ([string]$Ctx.Dst.Text).Trim() }  else { '' }
    $names = @(Get-CheckedGroups $Ctx.Tree | ForEach-Object { $_.DN })
    $noSep = if ($Ctx.NoSep) { [bool]$Ctx.NoSep.IsChecked } else { $true }
    $sep = if ($noSep -or -not $Ctx.Sep) { '' } else { [string]$Ctx.Sep.Text }
    return [pscustomobject]@{ Attribute = $attr; SourcePattern = $srcP; TargetPattern = $dstP; SearchBase = [string]$script:LoadedRootDN; Names = $names; Separator = $sep }
}

function Test-UrlModInputs {
    # Validates the inputs; shows a message and returns $false when something is missing.
    param($In, $Ctx, [switch]$Quiet)
    $cap = $Ctx.Caption
    if ($Quiet) {
        # Silent check used by the automatic re-scan (no message boxes).
        if (-not $script:LoadedRootDN -or -not $In.Attribute -or -not $In.SourcePattern) { return $false }
        if ($Ctx.Single -and $In.Names.Count -ne 1) { return $false }
        if (-not $Ctx.Single -and $In.Names.Count -eq 0) { return $false }
        if ($In.Attribute -match '[^A-Za-z0-9\-]') { return $false }
        return $true
    }
    if (-not $script:LoadedRootDN) { Show-Message -Text 'Connect to Active Directory first (Connect & load AD).' -Icon Warning; return $false }
    if ($Ctx.Single -and $In.Names.Count -ne 1) { Show-Message -Text ("Check the check box of exactly ONE group in the tree.`n`nYou have checked {0}." -f $In.Names.Count) -Icon Warning -Caption $cap; return $false }
    if (-not $Ctx.Single -and $In.Names.Count -eq 0) { Show-Message -Text 'Check the check box of at least one group (or an OU) in the tree.' -Icon Warning -Caption $cap; return $false }
    if (-not $In.Attribute)     { Show-Message -Text 'Enter the AD attribute to update (for example url, description, info) - or click a row in the attribute list.' -Icon Warning -Caption $cap; return $false }
    if (-not $In.SourcePattern) { Show-Message -Text 'Enter a Source Pattern - the text to look for in the attribute value.' -Icon Warning -Caption $cap; return $false }
    if ($In.Attribute -match '[^A-Za-z0-9\-]') { Show-Message -Text ("'{0}' is not a valid attribute name (letters, digits and - only)." -f $In.Attribute) -Icon Warning -Caption $cap; return $false }
    return $true
}

function Test-UrlModAttributeMultiValued {
    # $true when the attribute is multi-valued according to the AD schema (cached).
    # Falls back to multi-valued for 'url' and single-valued otherwise if the schema
    # cannot be read.
    param([string]$Attr)
    if ($null -eq $script:UrlModMultiCache) { $script:UrlModMultiCache = @{} }
    $k = $Attr.ToLower()
    if ($script:UrlModMultiCache.ContainsKey($k)) { return $script:UrlModMultiCache[$k] }
    # AD quirk: some attributes are multi-valued in the schema but AD enforces a
    # single value on group/user objects (description is the classic case).
    $forcedSingle = @('description','displayname','info','mail','managedby','samaccountname','name','cn')
    if ($forcedSingle -contains $k) { $script:UrlModMultiCache[$k] = $false; return $false }
    $multi = ($k -eq 'url')
    try {
        $server = $script:LoadedServer
        $rp = @{ ErrorAction = 'Stop' }; if ($server) { $rp.Server = $server }
        $schemaNc = (Get-ADRootDSE @rp).schemaNamingContext
        $sp = @{ SearchBase = $schemaNc; LDAPFilter = "(&(objectClass=attributeSchema)(lDAPDisplayName=$(ConvertTo-LdapFilterText $Attr)))"; Properties = 'isSingleValued'; ErrorAction = 'Stop' }
        if ($server) { $sp.Server = $server }
        $def = Get-ADObject @sp | Select-Object -First 1
        if ($def) { $multi = -not [bool]$def.isSingleValued }
        else { Write-Log ("Bulk-update: attribute '{0}' not found in the schema - assuming {1}." -f $Attr, $(if ($multi) { 'multi-valued' } else { 'single-valued' })) 'WARN' }
    }
    catch { Write-Log ("Bulk-update: could not read the schema for '{0}': {1}" -f $Attr, $_.Exception.Message) 'WARN' }
    $script:UrlModMultiCache[$k] = $multi
    return $multi
}

function Get-UrlModTargets {
    # Returns the AD group objects to process (with the attribute loaded), by DN.
    param([string[]]$Names, [string]$Attr)
    $server = $script:LoadedServer
    $groups = New-Object System.Collections.Generic.List[object]
    foreach ($n in @($Names)) {
        try {
            $p = @{ Identity = $n; Properties = $Attr; ErrorAction = 'Stop' }
            if ($server) { $p.Server = $server }
            $groups.Add((Get-ADGroup @p))
        }
        catch { Write-Log ("Bulk-update: group '{0}' not found: {1}" -f $n, $_.Exception.Message) 'WARN' }
    }
    return @($groups.ToArray())
}

function Get-UrlModPlan {
    # One record per attribute value that contains the source pattern.
    #   'New'        -> the new value will be ADDED, old kept   (Add new value checked, multi-valued attribute)
    #   'Replace'    -> the old value is REMOVED and the new one written
    #                   (Add new value cleared, or the attribute only allows one value)
    #   'Remove old' -> the new value already exists, only the old one is removed (Add new value cleared)
    #   'Already exists' / 'Duplicate in run' / 'Unchanged' -> skipped
    param([string]$Attr, [string]$SourcePattern, [string]$TargetPattern, [bool]$MultiValued, [bool]$AddMode, [string]$Separator = '', $Groups)

    $srcRegex = [regex]::Escape($SourcePattern)
    $recs = New-Object System.Collections.Generic.List[object]

    foreach ($g in @($Groups | Sort-Object Name)) {
        if (-not $g) { continue }
        $values = @($g.$Attr | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ })
        if ($values.Count -eq 0) { continue }

        $sourceEntries = @($values | Where-Object { $_ -imatch $srcRegex })
        if ($sourceEntries.Count -eq 0) { continue }

        $plannedForGroup = New-Object System.Collections.Generic.List[string]
        foreach ($src in $sourceEntries) {
            $newEntry = [regex]::Replace([string]$src, $srcRegex, $TargetPattern, 'IgnoreCase')
            $newValue = $newEntry
            if ($MultiValued) { $status = if ($AddMode) { 'New' } else { 'Replace' } }
            elseif ($AddMode) {
                # A single-valued attribute cannot hold two values: "keep old" means the
                # Target text is appended to the existing value (optionally with a separator),
                # e.g. "A; B" + "C" -> "A; BC" or, with separator "; ", "A; B; C".
                $status = 'Append'; $newValue = "$src$Separator$TargetPattern"
            }
            else { $status = 'Replace' }

            if ($status -ne 'Append' -and $newEntry -ieq $src) {
                $status = 'Unchanged'
            }
            elseif ($status -eq 'Append' -and $TargetPattern -and $src -imatch [regex]::Escape($TargetPattern)) {
                $status = 'Already exists'; $newValue = $src   # the target text is already part of the value
            }
            elseif ($status -ne 'Append' -and ($values | Where-Object { $_ -ieq $newEntry })) {
                $status = if ($MultiValued -and -not $AddMode) { 'Remove old' } else { 'Already exists' }
            }
            elseif ($status -ne 'Append' -and ($plannedForGroup | Where-Object { $_ -ieq $newEntry })) {
                $status = if ($MultiValued -and -not $AddMode) { 'Remove old' } else { 'Duplicate in run' }
            }
            else {
                [void]$plannedForGroup.Add($newEntry)
            }
            $recs.Add([pscustomobject]@{
                Group     = [string]$g.Name
                Status    = $status
                SourceUrl = [string]$src
                NewUrl    = [string]$newValue
                DN        = [string]$g.DistinguishedName
            })
        }
    }
    return @($recs.ToArray())
}

function Set-UrlModGrid {
    # Binds a record set to a preview grid via a DataTable/DataView.
    param($Records, $Grid)
    if (-not $Grid) { return }
    $view = (ConvertTo-DataTable -Records @($Records) -Columns (Get-UrlModColumns)).DefaultView
    $Grid.ItemsSource = $view
}

function Invoke-UrlModScan {
    # -Auto: silent re-scan after the Add/Replace or separator options change (only
    # when a scan has already been run on the tab).
    param([switch]$Single, [switch]$Auto)
    $ctx = Get-UrlModContext -Single:$Single
    if ($null -eq $script:UpdScanned) { $script:UpdScanned = @{} }
    if ($Auto -and -not $script:UpdScanned[$ctx.Key]) { return }
    $in = Get-UrlModInputs -Ctx $ctx
    if (-not (Test-UrlModInputs -In $in -Ctx $ctx -Quiet:$Auto)) { return }

    Set-Status ("Scanning '{0}' values ..." -f $in.Attribute) -Busy
    $addModeLog = if ($ctx.AddNew) { [bool]$ctx.AddNew.IsChecked } else { $true }
    Write-Log ("{0} scan: {1}  '{2}' -> '{3}'  ({4} checked group(s); mode: {5}; separator: '{6}')" -f $ctx.Caption, $in.Attribute, $in.SourcePattern, $in.TargetPattern, $in.Names.Count, $(if ($addModeLog) { 'Add new value (keep old)' } else { 'Remove old value and add new' }), $in.Separator) 'INFO'

    $multi = Test-UrlModAttributeMultiValued -Attr $in.Attribute
    $groups = @()
    try { $groups = @(Get-UrlModTargets -Names $in.Names -Attr $in.Attribute) }
    catch {
        Set-Status 'Scan failed.'
        Write-Log ("Bulk-update scan failed: {0}" -f $_.Exception.Message) 'ERROR'
        Show-Message -Text $_.Exception.Message -Icon Error -Caption $ctx.Caption
        return
    }

    $addMode = if ($ctx.AddNew) { [bool]$ctx.AddNew.IsChecked } else { $true }
    $plan = Get-UrlModPlan -Attr $in.Attribute -SourcePattern $in.SourcePattern -TargetPattern $in.TargetPattern -MultiValued $multi -AddMode $addMode -Separator $in.Separator -Groups $groups
    Set-UrlModGrid -Records $plan -Grid $ctx.Grid

    $todo = @($plan | Where-Object { $_.Status -in 'New','Replace','Remove old','Append' }).Count
    $skip = @($plan | Where-Object { $_.Status -eq 'Already exists' }).Count
    $mode = if (-not $multi) { if ($addMode) { 'single-valued: new text is appended to the existing value' } else { 'single-valued: the value is replaced' } } elseif ($addMode) { 'new value is added, old kept' } else { 'old value is removed, new value written' }
    if ($ctx.Count) {
        $ctx.Count.Text = ("'{0}' ({1})  ·  {2} group(s) match  ·  {3} to write  ·  {4} already present" -f $in.Attribute, $mode, $groups.Count, $todo, $skip)
    }
    $script:UpdScanned[$ctx.Key] = $true
    Set-Status ("Scan done: {0} change(s) proposed." -f $todo)
    Write-Log ("{0} scan: {1} group(s), {2} to write, {3} already present ({4})." -f $ctx.Caption, $groups.Count, $todo, $skip, $mode) 'OK'
}

function Export-UrlModReport {
    param([object[]]$Records, [string]$Attr, [switch]$WhatIf, [switch]$Single, $Settings = $null)
    $folder = Get-ReportFolder -Sub $(if ($Single) { 'single-group-update' } else { 'bulk-update' })
    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $prefix = if ($Single) { 'single_update' } else { 'bulk_update' }
    $base   = Join-Path $folder ('{0}_{1}_{2}' -f $prefix, ($Attr -replace '[^A-Za-z0-9\-]', ''), $stamp)
    $cols   = Get-UrlModColumns

    $banner = if ($WhatIf) {
        '<div class="banner wi">SIMULATION (WhatIf) - nothing was written to Active Directory. Clear WhatIf and run again for a live run.</div>'
    } else {
        '<div class="banner ok">Live run - the values below were written in Active Directory.</div>'
    }

    $hp = $base + '.html'; New-HtmlReport -Title ("{0} report - attribute '{1}'" -f $(if ($Single) { 'Single group update' } else { 'Bulk-update' }), $Attr) -Columns $cols -Records $Records -Path $hp -StatusKey 'Status' -Banner $banner -Unit 'rows' -Settings $Settings -HideChangedBox
    $cp = $base + '.csv';  Export-DataCsv  -Columns $cols -Records $Records -Path $cp
    $jp = $base + '.json'; Export-DataJson -Records $Records -Path $jp -Meta @{ type = 'bulk-update'; attribute = $Attr; whatIf = [bool]$WhatIf; rows = @($Records).Count }
    Write-Log ("Report: {0}" -f $hp) 'OK'
    if ($script:Config -and $script:Config.report.openAfter) { Start-Process $hp }
}

function Invoke-UrlModApply {
    param([switch]$Single)
    $ctx = Get-UrlModContext -Single:$Single
    $in = Get-UrlModInputs -Ctx $ctx
    if (-not (Test-UrlModInputs -In $in -Ctx $ctx)) { return }
    $whatIf = if ($ctx.WhatIf) { [bool]$ctx.WhatIf.IsChecked } else { $true }

    # Recompute from AD so the run always reflects the current state, not an earlier scan.
    Set-Status ("Reading '{0}' values ..." -f $in.Attribute) -Busy
    $multi = Test-UrlModAttributeMultiValued -Attr $in.Attribute
    $groups = @()
    try { $groups = @(Get-UrlModTargets -Names $in.Names -Attr $in.Attribute) }
    catch {
        Set-Status 'Apply failed.'
        Write-Log ("Bulk-update apply failed: {0}" -f $_.Exception.Message) 'ERROR'
        Show-Message -Text $_.Exception.Message -Icon Error -Caption $ctx.Caption
        return
    }
    $addMode = if ($ctx.AddNew) { [bool]$ctx.AddNew.IsChecked } else { $true }
    $plan = Get-UrlModPlan -Attr $in.Attribute -SourcePattern $in.SourcePattern -TargetPattern $in.TargetPattern -MultiValued $multi -AddMode $addMode -Separator $in.Separator -Groups $groups

    $todo = @($plan | Where-Object { $_.Status -in 'New','Replace','Remove old','Append' })
    if ($todo.Count -eq 0) {
        Set-UrlModGrid -Records $plan -Grid $ctx.Grid
        Set-Status 'Nothing to write.'
        Show-Message -Text ("No changes to write for '{0}'. Every value containing '{1}' already has its '{2}' counterpart, or the replacement changes nothing." -f $in.Attribute, $in.SourcePattern, $in.TargetPattern) -Icon Information -Caption $ctx.Caption
        return
    }

    if (-not $whatIf) {
        $verb = if ($multi) { 'Add' } else { 'Replace' }
        $ans = Show-Message -Text ("{0} {1} '{2}' value(s) on {3} group(s) in Active Directory?`n`n'{4}' -> '{5}'" -f $verb, $todo.Count, $in.Attribute, (@($todo | Select-Object -ExpandProperty DN -Unique).Count), $in.SourcePattern, $in.TargetPattern) -Icon Question -Buttons YesNo -Caption 'Confirm live run'
        if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { Set-Status 'Cancelled.'; return }
    }

    Set-Status $(if ($whatIf) { 'Simulating ...' } else { 'Writing ...' }) -Busy
    $server = $script:LoadedServer
    $done = 0; $errors = 0
    $result = New-Object System.Collections.Generic.List[object]

    foreach ($rec in $plan) {
        $kind = [string]$rec.Status
        if ($kind -notin 'New','Replace','Remove old','Append') { $result.Add($rec); continue }   # keep skipped rows visible in the report
        if ($whatIf) {
            $rec.Status = switch ($kind) { 'New' { 'Would add' } 'Replace' { 'Would replace' } 'Append' { 'Would append' } default { 'Would remove old' } }
            $done++
            Write-Log ("{0} on {1}: {2}" -f $rec.Status, $rec.Group, $rec.NewUrl) 'INFO'
        }
        else {
            try {
                $p = @{ Identity = $rec.DN; ErrorAction = 'Stop' }
                if ($server) { $p.Server = $server }
                switch ($kind) {
                    'New' {
                        $p.Add = @{ $in.Attribute = $rec.NewUrl }
                        try { Set-ADGroup @p; $rec.Status = 'Added' }
                        catch {
                            # The schema said multi-valued but AD enforces a single value:
                            # replace the value instead and remember the attribute as single-valued.
                            if ($_.Exception.Message -match 'only one value') {
                                Write-Log ("'{0}' only allows one value on {1} - replacing instead of adding." -f $in.Attribute, $rec.Group) 'WARN'
                                $script:UrlModMultiCache[$in.Attribute.ToLower()] = $false
                                $p.Remove('Add'); $p.Replace = @{ $in.Attribute = $rec.NewUrl }
                                Set-ADGroup @p; $rec.Status = 'Replaced'
                            } else { throw }
                        }
                    }
                    'Append' {
                        $p.Replace = @{ $in.Attribute = $rec.NewUrl }   # old + separator + new in the single value
                        Set-ADGroup @p; $rec.Status = 'Appended'
                    }
                    'Replace' {
                        if ($multi) { $p.Remove = @{ $in.Attribute = $rec.SourceUrl }; $p.Add = @{ $in.Attribute = $rec.NewUrl } }   # old value out, new value in
                        else        { $p.Replace = @{ $in.Attribute = $rec.NewUrl } }
                        Set-ADGroup @p; $rec.Status = 'Replaced'
                    }
                    default {
                        $p.Remove = @{ $in.Attribute = $rec.SourceUrl }
                        Set-ADGroup @p; $rec.Status = 'Removed old'
                    }
                }
                $done++
                Write-Log ("{0} on {1}: {2}" -f $rec.Status, $rec.Group, $rec.NewUrl) 'OK'
            }
            catch {
                $rec.Status = 'Error: ' + $_.Exception.Message
                $errors++
                Write-Log ("Could not write '{0}' on {1}: {2}" -f $in.Attribute, $rec.Group, $_.Exception.Message) 'ERROR'
            }
        }
        $result.Add($rec)
        Invoke-DoEvents
    }

    $records = @($result.ToArray())
    Set-UrlModGrid -Records $records -Grid $ctx.Grid

    if ($ctx.Count) { $ctx.Count.Text = '' }   # outcome is shown in the status bar and the log

    $settings = [ordered]@{
        'Run mode'                 = $(if ($whatIf) { 'WhatIf - simulation, nothing written' } else { 'Live run' })
        'Groups'                   = ('{0} selected, {1} holding the attribute' -f @($in.Names).Count, @($groups).Count)
        'AD-attribute'             = ('{0} ({1})' -f $in.Attribute, $(if ($multi) { 'multi-valued' } else { 'single-valued' }))
        'Source Pattern -> Target' = ('"{0}" -> "{1}"' -f $in.SourcePattern, $in.TargetPattern)
        'Mode'                     = $(if ($addMode) { 'Add new value (keep old)' } else { 'Remove old value and add new' })
        'Separator'                = $(if (-not $addMode -or $multi) { 'n/a' } elseif ($in.Separator) { ('"{0}"' -f $in.Separator) } else { 'None' })
        'Result'                   = ('{0} written, {1} error(s)' -f $done, $errors)
    }
    try { Export-UrlModReport -Records $records -Attr $in.Attribute -WhatIf:$whatIf -Single:$Single -Settings $settings }
    catch { Write-Log ("Bulk-update report failed: {0}" -f $_.Exception.Message) 'WARN' }

    Set-Status $(if ($whatIf) { ("Done (WhatIf): {0} would be written." -f $done) } else { ("Done: {0} written, {1} error(s)." -f $done, $errors) })
    Write-Log ("Bulk-update {0}: {1} {2}, {3} error(s)." -f $(if ($whatIf) { 'WHATIF' } else { 'COMMIT' }), $done, $(if ($whatIf) { 'would be written' } else { 'written' }), $errors) 'OK'
}

# ---------- Trees + "attributes of the selected group" (both tabs) ----------

function Update-UpdTreeCount {
    param($Ctx)
    if (-not $Ctx.Tree) { return }
    $n = @(Get-CheckedGroups $Ctx.Tree).Count
    if ($Ctx.TreeCount) {
        $Ctx.TreeCount.Text = "Selected groups: $n"
    }
    Update-UpdAttributes -Ctx $Ctx
}

function Update-UrlCount { Update-UpdTreeCount -Ctx (Get-UrlModContext) }


# PowerShell object artefacts that are not AD attributes.
$script:SgSkipProps = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($x in 'PropertyNames','PropertyCount','AddedProperties','RemovedProperties','ModifiedProperties') { [void]$script:SgSkipProps.Add($x) }

function Get-SgSchemaAttributes {
    # Every attribute a group object MAY or MUST hold according to the AD schema:
    # walks the class chain (group -> top) plus auxiliary classes. Cached. Returns
    # an empty list when the schema cannot be read.
    if ($null -ne $script:SgSchemaAttrs) { return $script:SgSchemaAttrs }
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    try {
        $server = $script:LoadedServer
        $rp = @{ ErrorAction = 'Stop' }; if ($server) { $rp.Server = $server }
        $schemaNc = (Get-ADRootDSE @rp).schemaNamingContext
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $queue = New-Object System.Collections.Generic.Queue[string]
        $queue.Enqueue('group')
        while ($queue.Count -gt 0) {
            $cls = $queue.Dequeue()
            if (-not $seen.Add($cls)) { continue }
            $cp = @{ SearchBase = $schemaNc; LDAPFilter = "(&(objectClass=classSchema)(lDAPDisplayName=$(ConvertTo-LdapFilterText $cls)))"
                     Properties = 'mayContain','systemMayContain','mustContain','systemMustContain','subClassOf','auxiliaryClass','systemAuxiliaryClass'; ErrorAction = 'Stop' }
            if ($server) { $cp.Server = $server }
            $def = Get-ADObject @cp | Select-Object -First 1
            if (-not $def) { continue }
            foreach ($prop in 'mayContain','systemMayContain','mustContain','systemMustContain') { foreach ($a in @($def.$prop)) { if ($a) { [void]$set.Add([string]$a) } } }
            foreach ($prop in 'subClassOf','auxiliaryClass','systemAuxiliaryClass') { foreach ($c in @($def.$prop)) { if ($c -and $c -ne $cls) { $queue.Enqueue([string]$c) } } }
        }
        Write-Log ("Attribute list: {0} attributes read from the schema for class 'group'." -f $set.Count) 'INFO'
    }
    catch { Write-Log ("Attribute list: could not read the schema: {0}" -f $_.Exception.Message) 'WARN' }
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($n in $set) { $list.Add($n) }
    $script:SgSchemaAttrs = @($list.ToArray() | Sort-Object)
    return $script:SgSchemaAttrs
}

function Update-UpdAttributes {
    # Fills the tab's "Attributes of the selected group" table from the FIRST checked
    # group (the only one on the single tab). "Ignore empty attributes" (default)
    # lists only attributes holding a value, otherwise every attribute the schema
    # allows for a group.
    param($Ctx)
    if (-not $Ctx.AttrGrid) { return }
    if ($null -eq $script:UpdObjects) { $script:UpdObjects = @{} }
    $nodes = @(Get-CheckedGroups $Ctx.Tree)
    if ($nodes.Count -eq 0 -or ($Ctx.Single -and $nodes.Count -ne 1)) { $script:UpdObjects[$Ctx.Key] = $null; $Ctx.AttrGrid.ItemsSource = $null; return }

    $dn = [string]$nodes[0].DN
    $obj = $script:UpdObjects[$Ctx.Key]
    if (-not $obj -or [string]$obj.DistinguishedName -ne $dn) {
        try { $obj = Read-CompareObject -Kind Group -DN $dn }
        catch { Write-Log ("Attribute list: could not read '{0}': {1}" -f $nodes[0].Name, $_.Exception.Message) 'WARN'; $obj = $null }
        $script:UpdObjects[$Ctx.Key] = $obj
    }
    if (-not $obj) { $Ctx.AttrGrid.ItemsSource = $null; return }
    $onlyFilled = if ($Ctx.NonEmpty) { [bool]$Ctx.NonEmpty.IsChecked } else { $true }

    $names = New-Object 'System.Collections.Generic.SortedSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($pr in $obj.PSObject.Properties) { if (-not $script:SgSkipProps.Contains($pr.Name)) { [void]$names.Add($pr.Name) } }
    if (-not $onlyFilled) { foreach ($n in (Get-SgSchemaAttributes)) { [void]$names.Add($n) } }

    $recs = New-Object System.Collections.Generic.List[object]
    foreach ($n in $names) {
        $txt = Get-AttrValueText -Object $obj -Name $n
        if ($onlyFilled -and [string]::IsNullOrEmpty($txt)) { continue }
        $recs.Add([pscustomobject]@{ Attribute = $n; Value = $txt })
    }
    $cols = @( @{ Key = 'Attribute'; Label = 'Attribute' }, @{ Key = 'Value'; Label = 'Value' } )
    $Ctx.AttrGrid.ItemsSource = (ConvertTo-DataTable -Records @($recs.ToArray()) -Columns $cols).DefaultView
}

function Select-UpdAttributeRow {
    # Clicking an attribute row puts its name in the tab's AD-attribute box.
    param($Ctx)
    $sel = $Ctx.AttrGrid.SelectedItem
    if ($sel -is [System.Data.DataRowView] -and $Ctx.Attr) { $Ctx.Attr.Text = [string]$sel.Row['Attribute'] }
}

function Update-UpdPreview {
    # Runs the scan again (without messages) when the Add/Replace or separator options change.
    param([switch]$Single)
    Invoke-UrlModScan -Single:$Single -Auto
}

function Update-UpdSepState {
    # The separator only matters in "Add new value (keep old)" on a single-valued
    # attribute: grey the separator controls out in "Remove old value and add new",
    # and the separator box while the "No separator" check box is checked.
    param([switch]$Single)
    $ctx = Get-UrlModContext -Single:$Single
    $addMode = if ($ctx.AddNew) { [bool]$ctx.AddNew.IsChecked } else { $true }
    if ($ctx.NoSep) { $ctx.NoSep.IsEnabled = $addMode }
    if ($ctx.Sep)   { $ctx.Sep.IsEnabled = ($addMode -and $ctx.NoSep -and -not [bool]$ctx.NoSep.IsChecked) }
}
