<#
    Loads the WPF window from asset\xaml\MainWindow.xaml, wires every control and
    shows the UI. Applies the application info (version.json) and defaults (config.json).
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

function Set-ComboByContent {
    param([System.Windows.Controls.ComboBox]$Combo, [string]$Content)
    if (-not $Combo -or [string]::IsNullOrWhiteSpace($Content)) { return }
    foreach ($it in $Combo.Items) { if ([string]$it.Content -eq $Content) { $Combo.SelectedItem = $it; return } }
}

function Show-MainWindow {
    [CmdletBinding()] param()

    $xamlPath = Join-Path $script:ScriptRoot 'asset\xaml\MainWindow.xaml'
    if (-not (Test-Path -LiteralPath $xamlPath)) { throw "XAML not found: $xamlPath" }
    [xml]$xaml = Get-Content -Path $xamlPath -Raw -Encoding UTF8
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    $script:Win = $window

    $names = @(
        # Shell
        'TitleText','TagLineText','SubtitleText','DotConn','BtnConnect','BtnReload',
        'TxtRoot','TxtStatus','Pb','LogBox','TglLog','BrdLog',
        'NavCreate','NavRename','NavUpdate','NavCopy','NavCompare','NavSnapUsers','NavSnapGroups','NavOu','NavSchema',
        'TxtToolTitle','TxtToolSub','BtnHelp','BrdHelp','TxtHelp','ToolTabs',
        # Shared selection (Create, Rename, Update attributes, Snapshot groups)
        'SelPane','SelTree','TxtSelCount','TxtSelFilter','ExpSelAttr','TxtSelAttr1','TxtSelVal1','TxtSelAttr2','TxtSelVal2','TxtSelAttr3','TxtSelVal3',
        'BtnSelAll','BtnSelNone','BtnSelExpand','BtnSelCollapse',
        # Create groups
        'GrdCreate','ScvCreateOpts','TxtTargetFilter','OuTree','TxtTargetOu','TglOuPicker','PnlOuPicker',
        'TxtFind','TxtReplace','ChkRegex','TxtPrefix','TxtSuffix','TxtSamSuffix','TxtDescription','TxtDescHint','TxtNotes','TxtNotesHint','CmbScope','CmbCategory',
        'ChkOnlyExisting','ChkCopyUsers','ChkCopyGroups','ChkCopyComputers','ChkCopyMemberOf','TxtCopyAttrs',
        'ChkCopyDesc','ChkAddDesc','ChkCopyAddDesc','ChkCopyNotes','ChkAddNotes','ChkCopyAddNotes','ChkNest','ChkWhatIf','ChkContinueOnError',
        'BtnBuildPreview','BtnApplyPattern','GridPreview','TxtPreviewCount','BtnCreate',
        # Rename groups
        'GrdRename','ScvRenameOpts','TxtRenFind','TxtRenReplace','ChkRenRegex','TxtRenPrefix','TxtRenSuffix','ChkRenName','ChkRenSam','ChkRenDisplay',
        'TxtRenDescAdd','CmbRenDescPos','ChkRenDescNoSep','TxtRenDescSep','ChkRenDescFind',
        'BtnRenPreview','BtnRenRecalc','BtnRenRevert','ChkRenShowChanged','ChkRenShowSame','BtnRenRowsAll','BtnRenRowsNone','BrdRenMode','TxtRenMode','GridRename',
        'TxtRenCount','BtnRenCheck','ChkRenWhatIf','BtnRenApply','TxtRenResult',
        # Update attributes (bulk + single group merged)
        'TxtUrlAttr','TxtUrlSourceOu','TxtUrlTargetOu','ChkUrlNoSep','TxtUrlSep','BtnUrlScan','TxtUrlCount','ChkUrlWhatIf','BtnUrlApply','ChkUrlAddNew','ChkUrlReplaceOld','GridUrl',
        'RbUpdChanges','RbUpdAttrs','ChkUrlNonEmpty','GridUrlAttrs',
        # Copy members
        'ChkCMUsers','ChkCMGroups','ChkCMComputers','ChkCMWhatIf','BtnCMRun',
        'CopyTree','TxtCMFilter','TxtCMAttr1','TxtCMVal1','TxtCMAttr2','TxtCMVal2','TxtCMAttr3','TxtCMVal3','BtnCMClear','BtnCMExpand','BtnCMCollapse',
        'TxtCMSrcName','TxtCMSrcInfo','TxtCMTgtName','TxtCMTgtInfo',
        # Snapshot groups (formerly Group analyser)
        'TxtAnalyseAttr','TxtAnalyseCount','BtnAnalyseAll','BtnAnalyseFmt',
        # Compare (groups / users)
        'RbCmpGroups','RbCmpUsers','PnlCmpGroups','PnlCmpUsers','BtnCmpExport','BtnCmpExportFmt',
        'TxtCmpAttr1','TxtCmpVal1','TxtCmpAttr2','TxtCmpVal2','TxtCmpAttr3','TxtCmpVal3','CompareTree','TxtCompareCount','BtnCompareUncheck','BtnCompareLoad',
        'TxtCompareResult','TxtCompareAName','TxtCompareBName','TxtCompareAttrFilter','ChkCompareOnlyDiff','GridGroupCmp',
        'TxtUserFilter','TxtUsrAttr1','TxtUsrVal1','TxtUsrAttr2','TxtUsrVal2','TxtUsrAttr3','TxtUsrVal3','UserTree','TxtUserCompareCount','BtnUserUncheck','BtnUserLoad',
        'TxtUserResult','TxtUserAName','TxtUserBName','TxtUserAttrFilter','ChkUserOnlyDiff','GridUserCmp',
        # Snapshots: 1 export, 2 compare, 3 restore lost groups
        'SnapSteps',
        'CmbSnapGroupBy','TxtSnapFilter','TxtSnapAttr1','TxtSnapVal1','TxtSnapAttr2','TxtSnapVal2','TxtSnapAttr3','TxtSnapVal3','TxtSnapCount','SnapTree',
        'BtnSnapCheckAll','BtnSnapUncheckAll','BtnSnapExpand','BtnSnapCollapse','BtnSnapAll','BtnSnapFmt',
        'TxtOldCsv','BtnPickOld','TxtNewCsv','BtnPickNew','TxtDiffFind1','TxtDiffRepl1','TxtDiffFind2','TxtDiffRepl2','TxtDiffFind3','TxtDiffRepl3','BtnCompareCsv','TxtDiffCount','GridDiff','BtnDiffAll','BtnDiffFmt',
        'TxtRestoreFile','BtnRestorePick','BtnRestoreLoad','GridRestore','TxtRestoreFilter','TxtRestoreCount',
        'BtnRestoreAll','BtnRestoreNone','BtnRestoreCheck','ChkRestoreWhatIf','BtnRestoreApply','TxtRestoreResult',
        # OU structure (bundled scripts in tools\ou-structure)
        'OuSteps','TxtOuExpRoot','BtnOuExpBrowse','BtnOuExpUseRoot','ChkOuExpSubOus','RbOuExpResult','RbOuExpOutput','GridOuExp','TxtOuExpOut','TxtOuExpResult','BtnOuExpFolder','BtnOuExport',
        'TxtOuImpFile','BtnOuImpPick','BtnOuImpLoad','TxtOuImpParent','BtnOuImpParentBrowse','TxtOuImpCount','RbOuImpResult','RbOuImpOutput','GridOuImp','TxtOuImpOut','TxtOuImpResult','ChkOuImpWhatIf','BtnOuImpFolder','BtnOuImport',
        # Schema attributes (bundled scripts in tools\schema-attributes)
        'SchSteps','RbSchFromUser','RbSchFromCsv','PnlSchUser','PnlSchCsv','TxtSchUser','BtnSchUserBrowse','TxtSchCsv','BtnSchCsvPick','ChkSchOnlyUsed',
        'RbSchExpResult','RbSchExpOutput','GridSchExp','TxtSchExpOut','TxtSchExpResult','BtnSchExpFolder','BtnSchExport',
        'TxtSchImpFile','BtnSchImpPick','BtnSchImpLoad','TxtSchOnly','ChkSchExclExchange','TxtSchImpCount',
        'RbSchImpResult','RbSchImpOutput','GridSchImp','TxtSchImpOut','TxtSchImpResult','ChkSchImpWhatIf','BtnSchImpFolder','BtnSchImport'
    )
    foreach ($n in $names) { Set-Variable -Name $n -Scope script -Value ($window.FindName($n)) }

    # The group tools share ONE selection tree. The functions still refer to the
    # tree by its old per-tool names, so those names all point at the shared tree.
    $script:GroupTree = $script:SelTree; $script:RenameTree = $script:SelTree
    $script:UrlTree   = $script:SelTree; $script:AnalyseTree = $script:SelTree

    # Size the window to the screen: no taller than the work area (laptops with
    # 125 % scaling only have ~860 px), maximised on small screens.
    $wa = [System.Windows.SystemParameters]::WorkArea
    $window.Width  = [Math]::Min(1560, [Math]::Floor($wa.Width  * 0.96))
    $window.Height = [Math]::Min(960,  [Math]::Floor($wa.Height * 0.96))
    if ($wa.Height -lt 760 -or $wa.Width -lt 1200) { $window.WindowState = [System.Windows.WindowState]::Maximized }

    $script:CheckMap = @{}
    $script:Suppress = 0

    # ---- Application info + defaults ----
    $a = $script:App
    if ($a) {
        $title = if ($a.version) { ('{0}  v{1}' -f $a.name, $a.version) } else { [string]$a.name }
        $script:TitleText.Text = $title; $script:Win.Title = [string]$a.name
        if ($a.description) { $script:TagLineText.Text = [string]$a.description }
    }
    $cfg = $script:Config
    if ($cfg) {
        $script:TxtRoot.Text        = [string]$script:DefaultSearchBase
        $script:TxtPrefix.Text      = [string]$cfg.defaults.namePrefix
        $script:TxtSuffix.Text      = [string]$cfg.defaults.nameSuffix
        if ($null -ne $cfg.defaults.samSuffix)   { $script:TxtSamSuffix.Text = [string]$cfg.defaults.samSuffix }
        if ($null -ne $cfg.defaults.nameFind)    { $script:TxtFind.Text    = [string]$cfg.defaults.nameFind }
        if ($null -ne $cfg.defaults.nameReplace) { $script:TxtReplace.Text = [string]$cfg.defaults.nameReplace }
        $script:ChkRegex.IsChecked = [bool]$cfg.defaults.nameRegex
        $script:TxtDescription.Text = [string]$cfg.defaults.description
        if ($null -ne $cfg.defaults.notes) { $script:TxtNotes.Text = [string]$cfg.defaults.notes }
        Set-ComboByContent -Combo $script:CmbScope    -Content ([string]$cfg.defaults.groupScope)
        Set-ComboByContent -Combo $script:CmbCategory -Content ([string]$cfg.defaults.groupCategory)
        $script:ChkCopyUsers.IsChecked     = [bool]$cfg.defaults.copyUsers
        $script:ChkCopyGroups.IsChecked    = [bool]$cfg.defaults.copyGroups
        $script:ChkCopyComputers.IsChecked = [bool]$cfg.defaults.copyComputers
        $script:ChkCopyMemberOf.IsChecked  = [bool]$cfg.defaults.copyMemberOf
        if ($null -ne $cfg.defaults.copyAttributes) { $script:TxtCopyAttrs.Text = [string]$cfg.defaults.copyAttributes }
        $script:ChkOnlyExisting.IsChecked  = [bool]$cfg.defaults.onlyCopyExisting
        $script:ChkNest.IsChecked          = [bool]$cfg.defaults.nestNewInOld
        $script:ChkWhatIf.IsChecked        = [bool]$cfg.defaults.whatIf
        if ($null -ne $cfg.defaults.continueOnError) { $script:ChkContinueOnError.IsChecked = [bool]$cfg.defaults.continueOnError }
        if ($null -ne $cfg.defaults.copyDescription)    { $script:ChkCopyDesc.IsChecked     = [bool]$cfg.defaults.copyDescription }
        if ($null -ne $cfg.defaults.addDescription)     { $script:ChkAddDesc.IsChecked      = [bool]$cfg.defaults.addDescription }
        if ($null -ne $cfg.defaults.copyAddDescription) { $script:ChkCopyAddDesc.IsChecked  = [bool]$cfg.defaults.copyAddDescription }
        if ($null -ne $cfg.defaults.copyNotes)          { $script:ChkCopyNotes.IsChecked    = [bool]$cfg.defaults.copyNotes }
        if ($null -ne $cfg.defaults.addNotes)           { $script:ChkAddNotes.IsChecked     = [bool]$cfg.defaults.addNotes }
        if ($null -ne $cfg.defaults.copyAddNotes)       { $script:ChkCopyAddNotes.IsChecked = [bool]$cfg.defaults.copyAddNotes }
    }
    Update-DescNotesState

    # ---- Local tree routines ----
    $checkAll = {
        param($tree, $countFn)
        foreach ($item in $tree.Items) { if ($item -is [System.Windows.Controls.TreeViewItem] -and $item.Tag.CheckBox) { $item.Tag.CheckBox.IsChecked = $true; Set-NodeCheckState -Item $item -IsChecked $true } }
        & $countFn
    }
    $uncheckAll = {
        param($tree, $countFn)
        $script:Suppress++
        try { foreach ($item in (Get-TreeItem -Parent $tree)) { if ($item.Tag.CheckBox) { $item.Tag.CheckBox.IsChecked = $false } } } finally { $script:Suppress-- }
        & $countFn
    }
    $expandAll   = { param($tree) foreach ($item in (Get-TreeItem -Parent $tree)) { if ($item.Items.Count -gt 0) { $item.IsExpanded = $true } } }
    $collapseTop = { param($tree) foreach ($item in $tree.Items) { if ($item -is [System.Windows.Controls.TreeViewItem]) { $item.IsExpanded = $false } } }

    # ---- Connection ----
    $connect = {
        try { Connect-AndLoadAD -RootDN $script:TxtRoot.Text.Trim() }
        catch { Set-Status 'Connection error.'; Write-Log "ERROR: $($_.Exception.Message)" 'ERROR'; Show-Message -Text $_.Exception.Message -Icon Error -Caption 'Error' }
    }
    $script:BtnConnect.Add_Click($connect)
    $script:BtnReload.Add_Click($connect)
    $script:TxtRoot.Add_KeyDown({ param($s, $e) Invoke-UiAction { if ($e.Key -eq [System.Windows.Input.Key]::Enter) { & $connect } } })

    # Attribute-filter wiring routine: attribute-name boxes apply on focus-leave
    # (never per keystroke - that would query AD for half-typed names); value
    # boxes apply as you type (the attribute is already cached by then).
    $wireAttrFilter = {
        param($AttrBoxes, $ValBoxes, $Action)
        foreach ($tb in @($AttrBoxes)) { if ($tb) { $tb.Add_LostFocus($Action) } }
        foreach ($tb in @($ValBoxes))  { if ($tb) { $tb.Add_TextChanged($Action) } }
    }

    # ---- Navigation: the sidebar picks the tool ----
    $script:ToolInfo = [ordered]@{
        create  = @{ Index = 0; Pane = $true;  Title = 'Create groups'
                     Sub   = 'New groups from existing ones - members, memberOf and attributes copied'
                     Help  = 'Check the source groups (or whole OUs) in the selection to the left. Choose the target OU, set the name rule (Find -> Replace, prefix, suffix) and what to copy. Build the preview - New name and New sAM can be edited per row - then run with WhatIf first. "Copy to existing groups only" creates nothing: members are copied to the group picked as target. Reports: files\create-groups.' }
        rename  = @{ Index = 1; Pane = $true;  Title = 'Rename groups'
                     Sub   = 'Same groups, new names - SID, members and permissions stay'
                     Help  = 'Check the groups in the selection to the left, set Find -> Replace and what it applies to (Name/CN, sAMAccountName, displayName). Build the preview, Check in AD (sAMAccountName free in the domain, name free in the OU, max 64 characters), then rename - WhatIf first. Every live run writes old and new values to files\rename-groups; "Revert from report..." renames the groups back.' }
        update  = @{ Index = 2; Pane = $true;  Title = 'Update attributes'
                     Sub   = 'Find -> replace a text in any attribute - one group or many'
                     Help  = 'Check one group to see all its attributes (click one to pick it), or many groups or a whole OU for a bulk update. Every value containing Find is written again with Find replaced (case-insensitive, literal text). Multi-valued attributes such as url get the new value added (or replaced); single-valued ones such as description are replaced or appended. Scan to preview, then apply - WhatIf first. Reports: files\bulk-update.' }
        copy    = @{ Index = 3; Pane = $false; Title = 'Copy members'
                     Sub   = 'Copy the members of one group into another'
                     Help  = 'Click a group in the tree: the first click fills Source, the next fills Target (clear a name and the next click fills that box again). The members of the source in the chosen types are added to the target. Additive - existing members are kept. A member report (account + UPN) is written to files\copy-members.' }
        snapgroups = @{ Index = 4; Pane = $true; Title = 'Snapshot groups'
                     Sub   = 'Export the selected groups - OU, members, memberOf'
                     Help  = 'Check groups or whole OUs in the selection to the left and export. The arrow next to Export groups exports a single format. An optional extra attribute adds a column - and then only groups that have a value are exported. Files: files\snapshot-groups.' }
        compare = @{ Index = 5; Pane = $false; Title = 'Compare'
                     Sub   = 'Two groups or two users side by side, every attribute'
                     Help  = 'Choose Groups or Users, check exactly two in the tree and click Compare selected. Rows that differ are red; "Only differences" hides the rest. Read-only - nothing is written to AD.' }
        snapusers = @{ Index = 6; Pane = $false; Title = 'Snapshot users'
                     Sub   = 'Export users, compare two exports, give back lost group memberships'
                     Help  = 'Step 1 exports the selected users with all attributes and group memberships - once before the move (OLD) and once after (NEW). Step 2 compares two exports and lists what changed per user, including lost and gained groups (pair renamed groups with Find -> Replace). Step 3 reads a comparison file and adds users back to the groups they lost - WhatIf first. Files: files\snapshot-users, files\snapshot-compare, files\restore-lost-groups.' }
        ou      = @{ Index = 7; Pane = $false; Title = 'OU structure'
                     Sub   = 'Copy an OU and all its sub-OUs from one domain to another'
                     Help  = 'Step 1 (in the source domain, read-only): exports an OU and, with Include all sub-OUs checked (default), every sub-OU below it - name, description and the protected flag - to a CSV. Step 2 (in the target domain): creates the structure from that CSV, parent before child, under the domain root or the parent DN you enter. OUs that already exist are skipped, so it can be run again. Only OUs are copied - no users, groups, computers or GPO links. Runs the scripts in tools\ou-structure unchanged; files: files\ou-structure.' }
        schema  = @{ Index = 8; Pane = $false; Title = 'Schema attributes'
                     Sub   = 'Copy the schema attributes production users have into a test forest'
                     Help  = 'Step 1 (in the source forest, read-only): exports the schema definitions of every attribute a user can have - from a user, or from a CSV of attribute names - including Exchange and custom attributes, their classes and the forward links that backward links need. Step 2 (in the target forest, as Schema Admins): creates the missing attributes and classes with the same OID, syntax and schemaIDGUID and attaches them to user. Schema changes are permanent - run WhatIf first. Runs the scripts in tools\schema-attributes unchanged; files: files\schema-attributes.' }
    }
    function script:Set-ActiveTool {
        param([string]$Id)
        $t = $script:ToolInfo[$Id]; if (-not $t) { return }
        $script:ActiveTool = $Id
        $script:ToolTabs.SelectedIndex = $t.Index
        $script:TxtToolTitle.Text = $t.Title
        $script:TxtToolSub.Text   = $t.Sub
        $script:TxtHelp.Text      = $t.Help
        $script:BtnHelp.IsChecked = $false
        $script:SelPane.Visibility = if ($t.Pane) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
        if ($Id -eq 'update') { Update-UrlCount }
    }
    $script:NavCreate.Add_Checked({  Invoke-UiAction { Set-ActiveTool 'create' } })
    $script:NavRename.Add_Checked({  Invoke-UiAction { Set-ActiveTool 'rename' } })
    $script:NavUpdate.Add_Checked({  Invoke-UiAction { Set-ActiveTool 'update' } })
    $script:NavCopy.Add_Checked({    Invoke-UiAction { Set-ActiveTool 'copy' } })
    $script:NavCompare.Add_Checked({ Invoke-UiAction { Set-ActiveTool 'compare' } })
    $script:NavSnapUsers.Add_Checked({  Invoke-UiAction { Set-ActiveTool 'snapusers' } })
    $script:NavSnapGroups.Add_Checked({ Invoke-UiAction { Set-ActiveTool 'snapgroups' } })
    $script:NavOu.Add_Checked({         Invoke-UiAction { Set-ActiveTool 'ou' } })
    $script:NavSchema.Add_Checked({     Invoke-UiAction { Set-ActiveTool 'schema' } })
    Set-ActiveTool 'create'

    # The options card above a preview grid may take what the grid does not need:
    # it scrolls instead of pushing the grid out of view (small screen, log open).
    $script:GrdCreate.Add_SizeChanged({ Invoke-UiAction { $script:ScvCreateOpts.MaxHeight = [Math]::Max(110, $script:GrdCreate.ActualHeight - 182) } })
    $script:GrdRename.Add_SizeChanged({ Invoke-UiAction { $script:ScvRenameOpts.MaxHeight = [Math]::Max(110, $script:GrdRename.ActualHeight - 182) } })

    # ---- Shared selection tree ----
    $script:TxtSelFilter.Add_TextChanged({ Invoke-UiAction { if ($script:SelTree) { Update-SelFilter } } })
    & $wireAttrFilter @($script:TxtSelAttr1,$script:TxtSelAttr2,$script:TxtSelAttr3) @($script:TxtSelVal1,$script:TxtSelVal2,$script:TxtSelVal3) { if ($script:SelTree) { Update-SelFilter } }
    $script:BtnSelAll.Add_Click({      Invoke-UiAction { & $checkAll   $script:SelTree { Update-SelCount } } })
    $script:BtnSelNone.Add_Click({     Invoke-UiAction { & $uncheckAll $script:SelTree { Update-SelCount } } })
    $script:BtnSelExpand.Add_Click({   Invoke-UiAction { & $expandAll  $script:SelTree } })
    $script:BtnSelCollapse.Add_Click({ Invoke-UiAction { & $collapseTop $script:SelTree } })

    # ---- Create groups ----
    $script:TxtTargetFilter.Add_TextChanged({ Invoke-UiAction { if ($script:OuTree) { Update-OuTreeFilter } } })
    # Picking a target closes the OU picker again.
    $script:OuTree.Add_SelectedItemChanged({ Invoke-UiAction { Update-TargetOuText; if ($script:OuTree.SelectedItem) { $script:TglOuPicker.IsChecked = $false } } })
    $script:BtnBuildPreview.Add_Click({ Invoke-UiAction { Build-Preview } })
    $script:BtnApplyPattern.Add_Click({ Invoke-UiAction { Apply-NamePattern } })
    $script:GridPreview.Add_CellEditEnding({ Invoke-UiAction { Invoke-DoEvents; Update-PreviewCount } })
    $script:BtnCreate.Add_Click({ Invoke-UiAction { Invoke-CreateGroups } })
    # Description: Copy / Add / Copy and add are mutually exclusive.
    $script:ChkCopyDesc.Add_Click({    Invoke-UiAction { if ($script:ChkCopyDesc.IsChecked)    { $script:ChkAddDesc.IsChecked = $false; $script:ChkCopyAddDesc.IsChecked = $false }; Update-DescNotesState } })
    $script:ChkAddDesc.Add_Click({     Invoke-UiAction { if ($script:ChkAddDesc.IsChecked)     { $script:ChkCopyDesc.IsChecked = $false; $script:ChkCopyAddDesc.IsChecked = $false }; Update-DescNotesState } })
    $script:ChkCopyAddDesc.Add_Click({ Invoke-UiAction { if ($script:ChkCopyAddDesc.IsChecked) { $script:ChkCopyDesc.IsChecked = $false; $script:ChkAddDesc.IsChecked = $false }; Update-DescNotesState } })
    # Notes: Copy / Add / Copy and add are mutually exclusive.
    $script:ChkCopyNotes.Add_Click({    Invoke-UiAction { if ($script:ChkCopyNotes.IsChecked)    { $script:ChkAddNotes.IsChecked = $false; $script:ChkCopyAddNotes.IsChecked = $false }; Update-DescNotesState } })
    $script:ChkAddNotes.Add_Click({     Invoke-UiAction { if ($script:ChkAddNotes.IsChecked)     { $script:ChkCopyNotes.IsChecked = $false; $script:ChkCopyAddNotes.IsChecked = $false }; Update-DescNotesState } })
    $script:ChkCopyAddNotes.Add_Click({ Invoke-UiAction { if ($script:ChkCopyAddNotes.IsChecked) { $script:ChkCopyNotes.IsChecked = $false; $script:ChkAddNotes.IsChecked = $false }; Update-DescNotesState } })

    # ---- Rename groups (in place - SID, members and permissions stay) ----
    $script:RenameRows = $null
    $script:RenameMode = 'Rename'
    $script:ChkRenDescNoSep.Add_Click({ Invoke-UiAction { Set-RenameDescState } })
    Set-RenameDescState
    $script:BtnRenPreview.Add_Click({  Invoke-UiAction { Build-RenamePreview } })
    $script:BtnRenRecalc.Add_Click({   Invoke-UiAction { Update-RenamePreviewNames } })
    $script:BtnRenRevert.Add_Click({   Invoke-UiAction { Invoke-RenameRevertPick } })
    $script:ChkRenShowChanged.Add_Click({ Invoke-UiAction { Update-RenameView } })
    $script:ChkRenShowSame.Add_Click({    Invoke-UiAction { Update-RenameView } })
    $script:BtnRenRowsAll.Add_Click({  Invoke-UiAction { Set-RenameRowsChecked -Value $true } })
    $script:BtnRenRowsNone.Add_Click({ Invoke-UiAction { Set-RenameRowsChecked -Value $false } })
    $script:BtnRenCheck.Add_Click({    Invoke-UiAction { Invoke-RenameCheck } })
    $script:BtnRenApply.Add_Click({    Invoke-UiAction { Invoke-RenameApply } })
    Update-RenamePreviewCount

    # ---- Update attributes ----
    $script:ChkUrlNonEmpty.Add_Click({ Invoke-UiAction { Update-UpdAttributes -Ctx (Get-UrlModContext) } })
    $script:RbUpdAttrs.Add_Checked({   Invoke-UiAction { Update-UpdAttributes -Ctx (Get-UrlModContext) } })
    $script:GridUrlAttrs.Add_SelectionChanged({ Invoke-UiAction { Select-UpdAttributeRow -Ctx (Get-UrlModContext) } })
    $script:ChkUrlNoSep.Add_Click({ Invoke-UiAction { Update-UpdSepState; Update-UpdPreview } })
    $script:BtnUrlScan.Add_Click({ Invoke-UiAction { $script:RbUpdChanges.IsChecked = $true; Invoke-UrlModScan } })
    $script:BtnUrlApply.Add_Click({ Invoke-UiAction { $script:RbUpdChanges.IsChecked = $true; Invoke-UrlModApply } })
    # Add / Replace are mutually exclusive - one of them is always checked.
    $script:ChkUrlAddNew.Add_Click({     Invoke-UiAction { if ($script:ChkUrlAddNew.IsChecked)     { $script:ChkUrlReplaceOld.IsChecked = $false } else { $script:ChkUrlReplaceOld.IsChecked = $true }; Update-UpdSepState; Update-UpdPreview } })
    $script:ChkUrlReplaceOld.Add_Click({ Invoke-UiAction { if ($script:ChkUrlReplaceOld.IsChecked) { $script:ChkUrlAddNew.IsChecked = $false }     else { $script:ChkUrlAddNew.IsChecked = $true }; Update-UpdSepState; Update-UpdPreview } })
    $script:TxtUrlSep.Add_TextChanged({ Invoke-UiAction { Update-UpdPreview } })

    # ---- Copy members (own tree; click = pick source, then target) ----
    $script:TxtCMFilter.Add_TextChanged({ Invoke-UiAction { if ($script:CopyTree) { Update-CMFilter } } })
    & $wireAttrFilter @($script:TxtCMAttr1,$script:TxtCMAttr2,$script:TxtCMAttr3) @($script:TxtCMVal1,$script:TxtCMVal2,$script:TxtCMVal3) { if ($script:CopyTree) { Update-CMFilter } }
    $script:CopyTree.Add_SelectedItemChanged({ Invoke-UiAction { Select-CopyGroup } })
    $script:BtnCMClear.Add_Click({ Invoke-UiAction { Clear-CopySelection } })
    $script:BtnCMExpand.Add_Click({ Invoke-UiAction { & $expandAll $script:CopyTree } })
    $script:BtnCMCollapse.Add_Click({ Invoke-UiAction { & $collapseTop $script:CopyTree } })
    $script:TxtCMSrcName.Add_TextChanged({ Invoke-UiAction { if ([string]::IsNullOrWhiteSpace($script:TxtCMSrcName.Text)) { $script:CMSource = $null; $script:TxtCMSrcInfo.Text = '' } } })
    $script:TxtCMTgtName.Add_TextChanged({ Invoke-UiAction { if ([string]::IsNullOrWhiteSpace($script:TxtCMTgtName.Text)) { $script:CMTarget = $null; $script:TxtCMTgtInfo.Text = '' } } })
    $script:BtnCMRun.Add_Click({ Invoke-UiAction { Invoke-CopyMembersDirect } })

    # ---- Snapshot groups (formerly Group analyser) ----
    $script:BtnAnalyseAll.Add_Click({ Invoke-UiAction { Export-AnalyseReport -Kind 'All' } })
    $script:BtnAnalyseFmt.Add_Click({
        Invoke-UiAction { Show-FormatMenu -Button $script:BtnAnalyseFmt -Items @(
            @('HTML only', { Export-AnalyseReport -Kind 'Html' }),
            @('CSV only',  { Export-AnalyseReport -Kind 'Csv' }),
            @('JSON only', { Export-AnalyseReport -Kind 'Json' })) }
    })

    # ---- Compare: groups ----
    & $wireAttrFilter @($script:TxtCmpAttr1,$script:TxtCmpAttr2,$script:TxtCmpAttr3) @($script:TxtCmpVal1,$script:TxtCmpVal2,$script:TxtCmpVal3) { if ($script:CompareTree) { Update-CompareTreeFilter } }
    $script:BtnCompareUncheck.Add_Click({ Invoke-UiAction { & $uncheckAll $script:CompareTree { Update-CompareSelectedCount } } })
    $script:BtnCompareLoad.Add_Click({ Invoke-UiAction { Invoke-CompareLoad } })
    $script:TxtCompareAttrFilter.Add_TextChanged({ Invoke-UiAction { Update-GroupCmpFilter } })
    $script:ChkCompareOnlyDiff.Add_Click({ Invoke-UiAction { Update-GroupCmpFilter } })
    # ---- Compare: users ----
    $script:TxtUserFilter.Add_TextChanged({ Invoke-UiAction { if ($script:UserTree) { Update-UserTreeFilter } } })
    & $wireAttrFilter @($script:TxtUsrAttr1,$script:TxtUsrAttr2,$script:TxtUsrAttr3) @($script:TxtUsrVal1,$script:TxtUsrVal2,$script:TxtUsrVal3) { if ($script:UserTree) { Update-UserTreeFilter } }
    $script:BtnUserUncheck.Add_Click({ Invoke-UiAction { & $uncheckAll $script:UserTree { Update-UserCompareCount } } })
    $script:BtnUserLoad.Add_Click({ Invoke-UiAction { Invoke-UserCompareLoad } })
    $script:TxtUserAttrFilter.Add_TextChanged({ Invoke-UiAction { Update-UserCmpFilter } })
    $script:ChkUserOnlyDiff.Add_Click({ Invoke-UiAction { Update-UserCmpFilter } })

    # ---- Compare: export the comparison shown ----
    $script:BtnCmpExport.Add_Click({ Invoke-UiAction { Export-CompareReport -Kind $(if ($script:RbCmpUsers.IsChecked) { 'User' } else { 'Group' }) -Format 'All' } })
    $script:BtnCmpExportFmt.Add_Click({
        Invoke-UiAction { Show-FormatMenu -Button $script:BtnCmpExportFmt -Items @(
            @('HTML only', { Export-CompareReport -Kind $(if ($script:RbCmpUsers.IsChecked) { 'User' } else { 'Group' }) -Format 'Html' }),
            @('CSV only',  { Export-CompareReport -Kind $(if ($script:RbCmpUsers.IsChecked) { 'User' } else { 'Group' }) -Format 'Csv' }),
            @('JSON only', { Export-CompareReport -Kind $(if ($script:RbCmpUsers.IsChecked) { 'User' } else { 'Group' }) -Format 'Json' })) }
    })

    # ---- OU structure ----
    $script:BtnOuExpUseRoot.Add_Click({ Invoke-UiAction { $script:TxtOuExpRoot.Text = ([string]$script:TxtRoot.Text).Trim() } })
    $script:BtnOuExpBrowse.Add_Click({ Invoke-UiAction {
        $dn = Select-AdObject -Kind OU -Title 'Select the root OU to export' -Current ([string]$script:TxtOuExpRoot.Text).Trim()
        if ($dn) { $script:TxtOuExpRoot.Text = $dn }
    } })
    $script:BtnOuImpParentBrowse.Add_Click({ Invoke-UiAction {
        $dn = Select-AdObject -Kind OU -AllowDomainRoot -Title 'Select where to create the OU structure' -Current ([string]$script:TxtOuImpParent.Text).Trim()
        if ($dn) {
            $script:TxtOuImpParent.Text = $dn
            # Show the new target DNs right away when a file is already loaded.
            if ($script:OuImpRows -and ([string]$script:TxtOuImpFile.Text).Trim()) { [void](Import-OuFile) }
        }
    } })
    $script:BtnOuExport.Add_Click({ Invoke-UiAction { Invoke-OuExport } })
    $script:BtnOuExpFolder.Add_Click({ Invoke-UiAction { Open-ToolFolder -Sub 'ou-structure' } })
    $script:BtnOuImpPick.Add_Click({
        Invoke-UiAction { $f = Select-ToolFile -Filter 'OU structure (*.csv)|*.csv|All files (*.*)|*.*' -Title 'Select an OU structure file (OU-Structure_*.csv)' -Sub 'ou-structure'
        if ($f) { $script:TxtOuImpFile.Text = $f; [void](Import-OuFile) } }
    })
    $script:BtnOuImpLoad.Add_Click({ Invoke-UiAction { [void](Import-OuFile) } })
    $script:BtnOuImport.Add_Click({ Invoke-UiAction { Invoke-OuImport } })
    $script:BtnOuImpFolder.Add_Click({ Invoke-UiAction { Open-ToolFolder -Sub 'ou-structure' } })

    # ---- Schema attributes ----
    $script:BtnSchUserBrowse.Add_Click({ Invoke-UiAction {
        $sam = Select-AdObject -Kind User -Title 'Select the user to read the attribute list from' -Current ([string]$script:TxtSchUser.Text).Trim()
        if ($sam) { $script:TxtSchUser.Text = $sam }
    } })
    $script:BtnSchCsvPick.Add_Click({
        Invoke-UiAction { $f = Select-ToolFile -Filter 'Attribute list (*.csv)|*.csv|All files (*.*)|*.*' -Title 'Select a CSV with a column Attribute' -Sub 'schema-attributes'
        if ($f) { $script:TxtSchCsv.Text = $f } }
    })
    $script:BtnSchExport.Add_Click({ Invoke-UiAction { Invoke-SchemaExport } })
    $script:BtnSchExpFolder.Add_Click({ Invoke-UiAction { Open-ToolFolder -Sub 'schema-attributes' } })
    $script:BtnSchImpPick.Add_Click({
        Invoke-UiAction { $f = Select-ToolFile -Filter 'Schema export (*.json)|*.json|All files (*.*)|*.*' -Title 'Select a schema export (SchemaAttributes_*.json)' -Sub 'schema-attributes'
        if ($f) { $script:TxtSchImpFile.Text = $f; [void](Import-SchemaFile) } }
    })
    $script:BtnSchImpLoad.Add_Click({ Invoke-UiAction { [void](Import-SchemaFile) } })
    $script:BtnSchImport.Add_Click({ Invoke-UiAction { Invoke-SchemaImport } })
    $script:BtnSchImpFolder.Add_Click({ Invoke-UiAction { Open-ToolFolder -Sub 'schema-attributes' } })

    # ---- Snapshots, step 1: export users ----
    $script:CmbSnapGroupBy.Add_SelectionChanged({ Invoke-UiAction { if ($script:SnapUsers) { Build-SnapTree } } })
    $script:TxtSnapFilter.Add_TextChanged({ Invoke-UiAction { if ($script:SnapUsers) { Build-SnapTree } } })
    & $wireAttrFilter @($script:TxtSnapAttr1,$script:TxtSnapAttr2,$script:TxtSnapAttr3) @($script:TxtSnapVal1,$script:TxtSnapVal2,$script:TxtSnapVal3) { if ($script:SnapUsers) { Build-SnapTree } }
    $script:BtnSnapCheckAll.Add_Click({ Invoke-UiAction { & $checkAll $script:SnapTree { Update-SnapSelectedCount } } })
    $script:BtnSnapUncheckAll.Add_Click({ Invoke-UiAction { & $uncheckAll $script:SnapTree { Update-SnapSelectedCount } } })
    $script:BtnSnapExpand.Add_Click({ Invoke-UiAction { & $expandAll $script:SnapTree } })
    $script:BtnSnapCollapse.Add_Click({ Invoke-UiAction { & $collapseTop $script:SnapTree } })
    $script:BtnSnapAll.Add_Click({ Invoke-UiAction { Export-SnapshotReport -Kind 'All' } })
    $script:BtnSnapFmt.Add_Click({
        Invoke-UiAction { Show-FormatMenu -Button $script:BtnSnapFmt -Items @(
            @('HTML only', { Export-SnapshotReport -Kind 'Html' }),
            @('CSV only',  { Export-SnapshotReport -Kind 'Csv' }),
            @('JSON only', { Export-SnapshotReport -Kind 'Json' })) }
    })
    # ---- Snapshots, step 2: compare ----
    $script:BtnPickOld.Add_Click({ Invoke-UiAction { Invoke-PickCsv -Which 'Old' } })
    $script:BtnPickNew.Add_Click({ Invoke-UiAction { Invoke-PickCsv -Which 'New' } })
    $script:BtnCompareCsv.Add_Click({ Invoke-UiAction { Invoke-CompareCsv } })
    $script:BtnDiffAll.Add_Click({ Invoke-UiAction { Export-DiffReport -Kind 'All' } })
    $script:BtnDiffFmt.Add_Click({
        Invoke-UiAction { Show-FormatMenu -Button $script:BtnDiffFmt -Items @(
            @('HTML only',        { Export-DiffReport -Kind 'Html' }),
            @('CSV only',         { Export-DiffReport -Kind 'Csv' }),
            @('JSON only',        { Export-DiffReport -Kind 'Json' }),
            @('SVG diagram only', { Export-DiffReport -Kind 'Svg' })) }
    })
    # ---- Snapshots, step 3: restore lost groups ----
    $script:BtnRestorePick.Add_Click({ Invoke-UiAction { Invoke-PickComparisonFile } })
    $script:BtnRestoreLoad.Add_Click({ Invoke-UiAction { Invoke-RestoreLoad } })
    $script:TxtRestoreFilter.Add_TextChanged({ Invoke-UiAction { Update-RestoreFilter } })
    $script:BtnRestoreAll.Add_Click({  Invoke-UiAction { Set-RestoreChecked -Value $true } })
    $script:BtnRestoreNone.Add_Click({ Invoke-UiAction { Set-RestoreChecked -Value $false } })
    $script:BtnRestoreCheck.Add_Click({ Invoke-UiAction { Invoke-RestoreCheck } })
    $script:BtnRestoreApply.Add_Click({ Invoke-UiAction { Invoke-RestoreApply } })

    Write-Log ('{0} v{1} ready. Click "Connect & load AD" to read groups.' -f $script:App.name, $script:App.version) 'INFO'
    [void]$window.ShowDialog()
}
