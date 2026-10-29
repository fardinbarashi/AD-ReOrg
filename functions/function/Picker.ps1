<#
    Browse dialog for Active Directory (asset\xaml\OuPicker.xaml).

    Select-AdObject -Kind OU   returns the DN of the OU picked (or of the domain
                               root with -AllowDomainRoot)
    Select-AdObject -Kind User returns the sAMAccountName of the user picked

    The tree starts at the domain root and loads the children of a node when it is
    expanded, so it opens fast in large domains. The search box finds OUs by name
    or users by name / account (ambiguous name resolution).

    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

function Get-PickerAdParams {
    # -Server for the AD cmdlets: the loaded server, else the one in config.json.
    $p = @{ ErrorAction = 'Stop' }
    if ($script:LoadedServer) { $p.Server = $script:LoadedServer }
    elseif ($script:Config -and $script:Config.domain.server) { $p.Server = [string]$script:Config.domain.server }
    return $p
}

function New-PickerItem {
    # One tree node. Containers get a placeholder child so the expand arrow shows.
    param([string]$DN, [string]$Label, [ValidateSet('Domain','OU','Container','User')] [string]$Type, [string]$Sam = '')
    $it = New-Object System.Windows.Controls.TreeViewItem
    $it.Header = $Label
    $it.Tag = @{ DN = $DN; Type = $Type; Sam = $Sam; Loaded = $false }
    if ($Type -eq 'User') { $it.Foreground = [System.Windows.Media.Brushes]::DarkSlateGray }
    else { [void]$it.Items.Add('Loading ...') }
    return $it
}

function Expand-PickerItem {
    # Loads the child OUs (and, when picking a user, containers and users) of a node once.
    param($Item, [ValidateSet('OU','User')] [string]$Kind)
    if (-not $Item -or -not $Item.Tag -or $Item.Tag.Loaded -or $Item.Tag.Type -eq 'User') { return }
    $Item.Tag.Loaded = $true
    $Item.Items.Clear()
    $p = Get-PickerAdParams
    $dn = [string]$Item.Tag.DN
    try {
        if ($Kind -eq 'User') {
            $kids = @(Get-ADObject @p -SearchBase $dn -SearchScope OneLevel -LDAPFilter '(|(objectClass=organizationalUnit)(objectClass=container))' |
                      Where-Object { [string]$_.Name -notin 'System', 'Program Data', 'ForeignSecurityPrincipals', 'Managed Service Accounts', 'LostAndFound', 'NTDS Quotas', 'TPM Devices', 'Keys', 'Infrastructure' } |
                      Sort-Object Name)
            foreach ($k in $kids) {
                $type = if ([string]$k.ObjectClass -eq 'organizationalUnit') { 'OU' } else { 'Container' }
                [void]$Item.Items.Add((New-PickerItem -DN $k.DistinguishedName -Label $k.Name -Type $type))
            }
            foreach ($u in @(Get-ADUser @p -SearchBase $dn -SearchScope OneLevel -Filter * | Sort-Object Name)) {
                [void]$Item.Items.Add((New-PickerItem -DN $u.DistinguishedName -Label ('{0}  ({1})' -f $u.Name, $u.SamAccountName) -Type 'User' -Sam $u.SamAccountName))
            }
        } else {
            foreach ($o in @(Get-ADOrganizationalUnit @p -SearchBase $dn -SearchScope OneLevel -Filter * | Sort-Object Name)) {
                [void]$Item.Items.Add((New-PickerItem -DN $o.DistinguishedName -Label $o.Name -Type 'OU'))
            }
        }
    } catch {
        [void]$Item.Items.Add(('Could not read: {0}' -f $_.Exception.Message))
    }
}

function Open-PickerPath {
    # Expands the tree down to $DN and selects it (used when the field already holds a DN).
    param($Root, [string]$DN, [string]$Kind)
    if (-not $Root -or -not $DN) { return }
    $target = $DN.ToLower()
    $node = $Root
    for ($guard = 0; $guard -lt 64 -and $node; $guard++) {
        if (([string]$node.Tag.DN).ToLower() -eq $target) { $node.IsSelected = $true; $node.BringIntoView(); return }
        Expand-PickerItem -Item $node -Kind $Kind
        $node.IsExpanded = $true
        $next = $null
        foreach ($c in $node.Items) {
            if ($c -is [System.Windows.Controls.TreeViewItem]) {
                $cdn = ([string]$c.Tag.DN).ToLower()
                if ($target -eq $cdn -or $target.EndsWith(',' + $cdn)) { $next = $c; break }
            }
        }
        $node = $next
    }
}

function Select-AdObject {
    # Shows the browse dialog. Returns the OU's DN / the user's sAMAccountName, or $null.
    param(
        [ValidateSet('OU','User')] [string]$Kind = 'OU',
        [string]$Title = '',
        [string]$Current = '',
        [switch]$AllowDomainRoot
    )
    $path = Join-Path $script:ScriptRoot 'asset\xaml\OuPicker.xaml'
    if (-not (Test-Path -LiteralPath $path)) { throw "XAML not found: $path" }
    [xml]$x = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    $dlg = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
    if ($script:Win) { $dlg.Owner = $script:Win }
    $dlg.Title = if ($Title) { $Title } elseif ($Kind -eq 'User') { 'Select a user' } else { 'Select an OU' }
    $tree   = $dlg.FindName('TreePick');   $list  = $dlg.FindName('LstPick')
    $dnText = $dlg.FindName('TxtPickDN');  $hint  = $dlg.FindName('TxtPickHint')
    $search = $dlg.FindName('TxtPickSearch'); $btnSearch = $dlg.FindName('BtnPickSearch'); $btnTree = $dlg.FindName('BtnPickTree')
    $ok     = $dlg.FindName('BtnPickOk')
    $hint.Text = if ($Kind -eq 'User') { 'Expand the tree and click a user, or search by name or account. Double-click selects.' }
                 else { 'Expand the tree and click an OU, or search by OU name. Double-click selects.' }

    $p = Get-PickerAdParams
    $domain = Get-ADDomain @p
    $rootItem = New-PickerItem -DN ([string]$domain.DistinguishedName) -Label ([string]$domain.DNSRoot) -Type 'Domain'
    [void]$tree.Items.Add($rootItem)
    $script:PickerKind = $Kind

    # The picked object: @{ DN; Type; Sam }
    $script:PickerChoice = $null
    $choose = {
        param($tag)
        $script:PickerChoice = $null
        if (-not $tag -or -not ($tag -is [hashtable])) { $dnText.Text = '(nothing selected)'; $ok.IsEnabled = $false; return }
        $dnText.Text = [string]$tag.DN
        $fits = if ($Kind -eq 'User') { $tag.Type -eq 'User' } else { $tag.Type -eq 'OU' -or ($tag.Type -eq 'Domain' -and $AllowDomainRoot) }
        if ($fits) { $script:PickerChoice = $tag }
        $ok.IsEnabled = [bool]$fits
    }

    $tree.AddHandler([System.Windows.Controls.TreeViewItem]::ExpandedEvent, [System.Windows.RoutedEventHandler]{
        param($s, $e) Invoke-UiAction { Expand-PickerItem -Item $e.OriginalSource -Kind $script:PickerKind }
    })
    $tree.Add_SelectedItemChanged({ Invoke-UiAction { $sel = $tree.SelectedItem; & $choose $(if ($sel -is [System.Windows.Controls.TreeViewItem]) { $sel.Tag } else { $null }) } })
    $tree.Add_MouseDoubleClick({ Invoke-UiAction { if ($ok.IsEnabled) { $dlg.DialogResult = $true } } })
    $list.Add_SelectionChanged({ Invoke-UiAction { $sel = $list.SelectedItem; & $choose $(if ($sel) { $sel.Tag } else { $null }) } })
    $list.Add_MouseDoubleClick({ Invoke-UiAction { if ($ok.IsEnabled) { $dlg.DialogResult = $true } } })
    $ok.Add_Click({ Invoke-UiAction { $dlg.DialogResult = $true } })

    $runSearch = {
        $q = ([string]$search.Text).Trim()
        if (-not $q) { return }
        $esc = ConvertTo-LdapFilterText $q
        $list.Items.Clear()
        try {
            $hits = if ($Kind -eq 'User') { @(Get-ADUser @p -LDAPFilter "(anr=$esc)" -ResultSetSize 200 | Sort-Object Name) }
                    else { @(Get-ADOrganizationalUnit @p -LDAPFilter "(name=*$esc*)" -ResultSetSize 200 | Sort-Object Name) }
        } catch { $hits = @(); [void]$list.Items.Add(('Search failed: {0}' -f $_.Exception.Message)) }
        foreach ($h in $hits) {
            $li = New-Object System.Windows.Controls.ListBoxItem
            if ($Kind -eq 'User') {
                $li.Content = '{0}  ({1})  -  {2}' -f $h.Name, $h.SamAccountName, (([regex]'(?<!\\),').Split([string]$h.DistinguishedName, 2)[1])
                $li.Tag = @{ DN = [string]$h.DistinguishedName; Type = 'User'; Sam = [string]$h.SamAccountName }
            } else {
                $li.Content = '{0}  -  {1}' -f $h.Name, $h.DistinguishedName
                $li.Tag = @{ DN = [string]$h.DistinguishedName; Type = 'OU'; Sam = '' }
            }
            [void]$list.Items.Add($li)
        }
        if ($hits.Count -eq 0 -and $list.Items.Count -eq 0) { [void]$list.Items.Add('No match.') }
        $tree.Visibility = [System.Windows.Visibility]::Collapsed
        $list.Visibility = [System.Windows.Visibility]::Visible
        $btnTree.Visibility = [System.Windows.Visibility]::Visible
        & $choose $null
    }
    $btnSearch.Add_Click({ Invoke-UiAction { & $runSearch } })
    # PreviewKeyDown: Enter in the search box searches instead of pressing the default Select button.
    $search.Add_PreviewKeyDown({ param($s, $e) Invoke-UiAction { if ($e.Key -eq [System.Windows.Input.Key]::Enter) { $e.Handled = $true; & $runSearch } } })
    $btnTree.Add_Click({ Invoke-UiAction {
        $list.Visibility = [System.Windows.Visibility]::Collapsed
        $tree.Visibility = [System.Windows.Visibility]::Visible
        $btnTree.Visibility = [System.Windows.Visibility]::Collapsed
        & $choose $(if ($tree.SelectedItem -is [System.Windows.Controls.TreeViewItem]) { $tree.SelectedItem.Tag } else { $null })
    } })

    # Open at the DN already in the field (an OU, or the OU a user sits in).
    Expand-PickerItem -Item $rootItem -Kind $Kind
    $rootItem.IsExpanded = $true
    $start = [string]$Current
    if ($Kind -eq 'User' -and $start -and $start -notmatch '=') {
        try { $start = [string](Get-ADUser @p -Identity $start).DistinguishedName } catch { $start = '' }
    }
    if ($start -match '=') { try { Open-PickerPath -Root $rootItem -DN $start -Kind $Kind } catch { } }

    if ($dlg.ShowDialog() -eq $true -and $script:PickerChoice) {
        if ($Kind -eq 'User') { return [string]$script:PickerChoice.Sam }
        return [string]$script:PickerChoice.DN
    }
    return $null
}
