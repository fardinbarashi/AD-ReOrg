<#
    Builds the tree views and loads Active Directory:
      * the source-group tree (OU -> groups, with checkboxes) for the Create tab
      * the target tree (OU + groups) used to pick where new groups go / which
        existing group to copy into
      * the checkbox trees on the other group tabs and the click-to-pick tree on
        Copy members (same data)
    Also holds the shared tree routines used by every tab.
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

# ---------- Shared tree routines ----------

function Get-TreeItem {
    # Enumerates every TreeViewItem under a parent (TreeView or TreeViewItem).
    param([Parameter(Mandatory)] [object]$Parent)
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($child in $Parent.Items) {
        if (-not ($child -is [System.Windows.Controls.TreeViewItem])) { continue }
        $result.Add($child)
        if ($child.Items.Count -gt 0) {
            foreach ($g in (Get-TreeItem -Parent $child)) { $result.Add($g) }
        }
    }
    return $result
}

function Set-NodeCheckState {
    # Propagates a check state down to every descendant node.
    param(
        [Parameter(Mandatory)] [System.Windows.Controls.TreeViewItem]$Item,
        [Parameter(Mandatory)] [bool]$IsChecked
    )
    $script:Suppress++
    try {
        foreach ($child in $Item.Items) {
            if (-not ($child -is [System.Windows.Controls.TreeViewItem])) { continue }
            $cmeta = $child.Tag
            if ($null -eq $cmeta) { continue }
            if ($cmeta.CheckBox) { $cmeta.CheckBox.IsChecked = $IsChecked }
            Set-NodeCheckState -Item $child -IsChecked $IsChecked
        }
    }
    finally { $script:Suppress-- }
}

function Get-GroupMemberCount {
    param($Group)
    $cnt = 0
    if ($Group -and $Group.member) { $cnt = ([array]$Group.member).Count }
    return [int]$cnt
}

function Get-GroupMemberDNs {
    # Returns EVERY member DN of a group by reading its 'member' attribute.
    # Unlike Get-ADGroupMember (capped at MaxGroupOrMemberEntries, default 5000),
    # the member attribute is returned in full via automatic range retrieval.
    param([Parameter(Mandatory)] [string]$DN, [string]$Server)
    $p = @{ Identity = $DN; Properties = 'member'; ErrorAction = 'Stop' }
    if ($Server) { $p.Server = $Server }
    $g = Get-ADGroup @p
    return @(@($g.member) | Where-Object { $_ })
}

function Get-GroupMemberObjects {
    # Returns EVERY member of a group as objects with distinguishedName, objectClass,
    # SamAccountName, name and UserPrincipalName. Reads the 'member' attribute (no
    # 5000-member cap); identity details come from the lookups built at connect time,
    # with a per-DN AD lookup for members outside the loaded root.
    param([Parameter(Mandatory)] [string]$DN, [string]$Server)
    if ($null -eq $script:MemberInfoCache) { $script:MemberInfoCache = @{} }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($mdn in (Get-GroupMemberDNs -DN $DN -Server $Server)) {
        $lower = ([string]$mdn).ToLower()
        $info = $null
        if     ($script:GroupInfoByDn    -and $script:GroupInfoByDn.ContainsKey($lower))    { $info = $script:GroupInfoByDn[$lower] }
        elseif ($script:ComputerInfoByDn -and $script:ComputerInfoByDn.ContainsKey($lower)) { $info = $script:ComputerInfoByDn[$lower] }
        elseif ($script:UserInfoByDn     -and $script:UserInfoByDn.ContainsKey($lower))     { $info = $script:UserInfoByDn[$lower] }
        elseif ($script:MemberInfoCache.ContainsKey($lower))                                 { $info = $script:MemberInfoCache[$lower] }
        if (-not $info) {
            try {
                $op = @{ Identity = $mdn; Properties = 'objectClass','sAMAccountName','name','userPrincipalName'; ErrorAction = 'Stop' }
                if ($Server) { $op.Server = $Server }
                $o = Get-ADObject @op
                $info = @{ Class = [string]$o.objectClass; Sam = [string]$o.sAMAccountName; Name = [string]$o.name; Upn = [string]$o.userPrincipalName }
            }
            catch { $info = @{ Class = 'user'; Sam = ''; Name = (($mdn -split ',', 2)[0] -replace '^CN=', ''); Upn = '' } }
            $script:MemberInfoCache[$lower] = $info
        }
        $out.Add([pscustomobject]@{ distinguishedName = [string]$mdn; objectClass = $info.Class; SamAccountName = $info.Sam; name = $info.Name; UserPrincipalName = $info.Upn })
    }
    return @($out.ToArray())
}

function Get-GroupCategoryTag {
    # Returns "(U)(C)(G)" showing which member object classes a group contains,
    # using the pre-built domain DN lookups.
    param($Group)
    if (-not $Group -or -not $Group.member) { return '' }
    $hasU = $false; $hasC = $false; $hasG = $false
    foreach ($dn in @($Group.member)) {
        $key = ([string]$dn).ToLower()
        if     ($script:GroupDnSet    -and $script:GroupDnSet.Contains($key))    { $hasG = $true }
        elseif ($script:ComputerDnSet -and $script:ComputerDnSet.Contains($key)) { $hasC = $true }
        elseif ($script:UserDnSet     -and $script:UserDnSet.Contains($key))     { $hasU = $true }
        else { $hasU = $true }   # members outside the search base are treated as users
        if ($hasU -and $hasC -and $hasG) { break }
    }
    $t = ''
    if ($hasU) { $t += '(U)' }
    if ($hasC) { $t += '(C)' }
    if ($hasG) { $t += '(G)' }
    return $t
}

function Get-GroupTreeLabel {
    # Label order:  Name  (U)(C)(G)  ·  12 mem.  ·  Global/Security  ·  sAMAccountName
    param([Parameter(Mandatory)] $Group, [string]$Marker = '')
    $scope = [string]$Group.GroupScope
    $cat   = [string]$Group.GroupCategory
    $cnt   = Get-GroupMemberCount -Group $Group
    $tags  = Get-GroupCategoryTag -Group $Group
    $tg    = if ($tags) { "  $tags" } else { '' }
    $sc    = if ($scope -or $cat) { ('   ·   {0}/{1}' -f $scope, $cat) } else { '' }
    return ('{0}{1}   ·   {2} mem.{3}   ·   {4}{5}' -f $Group.Name, $tg, $cnt, $sc, $Group.SamAccountName, $Marker)
}

# ---------- Checkbox group tree ----------

function New-GroupCheckNode {
    # A TreeViewItem with [checkbox] [text]. $TreeKind routes the checkbox count.
    param(
        [Parameter(Mandatory)] [ValidateSet('OU','GroupLeaf')] [string]$Type,
        [Parameter(Mandatory)] [string]$DN,
        [Parameter(Mandatory)] [string]$Label,
        [ValidateSet('Shared','Create','Compare','UserCompare','Analyse','Bulk','Rename')] [string]$TreeKind = 'Create'
    )
    $item  = New-Object System.Windows.Controls.TreeViewItem
    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Orientation = 'Horizontal'
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.VerticalAlignment = 'Center'; $cb.Margin = '0,0,6,0'
    $text = New-Object System.Windows.Controls.TextBlock
    $text.VerticalAlignment = 'Center'; $text.Text = $Label
    if ($Type -eq 'OU') { $text.FontWeight = 'SemiBold' }
    [void]$panel.Children.Add($cb); [void]$panel.Children.Add($text)
    $item.Header = $panel
    $item.Tag = @{ Type = $Type; DN = $DN; Text = $text; CheckBox = $cb; Count = 0; Name = ''; Sam = ''; HasUrl = $false; TreeKind = $TreeKind }
    $script:CheckMap[$cb] = $item
    $cb.Add_Click({
        param($s, $e)
        $owner = $script:CheckMap[$s]
        if ($null -ne $owner) {
            Set-NodeCheckState -Item $owner -IsChecked ([bool]$s.IsChecked)
            switch ($owner.Tag.TreeKind) {
                'Compare'     { Update-CompareSelectedCount }
                'UserCompare' { Update-UserCompareCount }
                'Analyse'     { Update-AnalyseCount }
                'Shared'      { Update-SelCount }
                'Bulk'        { Update-UrlCount }
                'Rename'      { Update-RenameCount }
                default       { Update-GroupSelectedCount }
            }
        }
        if ($e) { $e.Handled = $true }
    })
    return $item
}

function New-CheckGroupTree {
    # Builds an OU -> groups tree with checkboxes into the given TreeView.
    param(
        [Parameter(Mandatory)] $Tree,
        [Parameter(Mandatory)] [ValidateSet('Shared','Create','Compare','UserCompare','Analyse','Bulk','Rename')] [string]$TreeKind,
        [Parameter(Mandatory)] [string]$RootDN,
        [Parameter(Mandatory)] [string]$RootName,
        [object[]]$Groups
    )
    if (-not $Tree) { return }
    $Tree.Items.Clear()
    $rootLower = $RootDN.ToLower()

    $rootNode = New-GroupCheckNode -Type 'OU' -DN $RootDN -Label $RootName -TreeKind $TreeKind
    [void]$Tree.Items.Add($rootNode)
    $map = @{}; $map[$rootLower] = $rootNode

    function Get-OrCreateOu($dn) {
        $dnl = $dn.ToLower()
        if ($dnl -eq $rootLower -or -not $dnl.EndsWith($rootLower)) { return $rootNode }
        if ($map.ContainsKey($dnl)) { return $map[$dnl] }
        $parent = ($dn -split ',', 2)[1]
        $pNode  = Get-OrCreateOu $parent
        $label  = ($dn -split ',', 2)[0] -replace '^OU=', '' -replace '^CN=', ''
        $n = New-GroupCheckNode -Type 'OU' -DN $dn -Label $label -TreeKind $TreeKind
        [void]$pNode.Items.Add($n); $map[$dnl] = $n
        return $n
    }

    $isCompare = ($TreeKind -eq 'Compare')
    $isMain    = ($TreeKind -in 'Shared','Create')
    foreach ($g in ($Groups | Sort-Object Name)) {
        $parentDn = ($g.DistinguishedName -split ',', 2)[1]
        $ouNode = Get-OrCreateOu $parentDn
        $url = ($g.url) -and (@($g.url).Count -gt 0)
        if ($isCompare) {
            $label  = '{0}  ({1})' -f $g.Name, $g.SamAccountName
        } else {
            $label  = Get-GroupTreeLabel -Group $g
        }
        $leaf = New-GroupCheckNode -Type 'GroupLeaf' -DN $g.DistinguishedName -Label $label -TreeKind $TreeKind
        $leaf.Tag.Name   = $g.Name
        $leaf.Tag.Sam    = $g.SamAccountName
        $leaf.Tag.HasUrl = [bool]$url
        # Colour a group red when it is a member of other groups (nested), on the
        # Create / shared selection tree.
        if ($isMain -and $g.memberOf -and @($g.memberOf).Count -gt 0) {
            $leaf.Tag.Text.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(180,35,42))
            $leaf.Tag.HasMemberOf = $true
        }
        [void]$ouNode.Items.Add($leaf)
        # The group objects of the main (shared) tree feed Create groups and Snapshot groups.
        if ($isMain) { $script:GroupCache[$g.DistinguishedName] = $g }

        $p = $ouNode
        while ($p -and $p.Tag -and $p.Tag.Type -eq 'OU') { $p.Tag.Count++; $p = $p.Parent }
    }

    foreach ($item in (Get-TreeItem -Parent $Tree)) {
        if ($item.Tag.Type -eq 'OU') {
            $base = if ($item -eq $rootNode) { $RootName } else { ($item.Tag.DN -split ',', 2)[0] -replace '^OU=', '' }
            $item.Tag.Text.Text = '{0}  ({1})' -f $base, $item.Tag.Count
        }
    }
    $rootNode.IsExpanded = $true
}

function Get-CheckedGroups {
    # Returns {Name,Sam,DN} for checked + visible group leaves in a tree.
    param($Tree)
    $list = New-Object System.Collections.Generic.List[object]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    if (-not $Tree) { return @($list.ToArray()) }
    foreach ($item in (Get-TreeItem -Parent $Tree)) {
        $m = $item.Tag
        if ($m.Type -eq 'GroupLeaf' -and $item.Visibility -eq [System.Windows.Visibility]::Visible -and
            $m.CheckBox -and $m.CheckBox.IsChecked -eq $true) {
            if ($seen.Add($m.DN)) { $list.Add([pscustomobject]@{ Name = $m.Name; Sam = $m.Sam; DN = $m.DN }) }
        }
    }
    return @($list.ToArray())
}

function Update-CheckGroupTreeFilter {
    # Filters leaves by name text, optional "only url", and up to three AD
    # attribute/value pairs (AND, contains). Hides OUs with no visible leaf.
    #   $Kind = 'Group' -> match by DN via the group attribute cache
    #   $Kind = 'User'  -> match by sAMAccountName via the user attribute cache
    param($Tree, $FilterBox, [bool]$OnlyWithUrl = $false, $OnlyWithUrlBox = $null,
          $AttrBoxes = @(), $ValBoxes = @(), [string]$Kind = 'Group')
    if (-not $Tree) { return }
    $q = ''
    if ($FilterBox) { $q = ([string]$FilterBox.Text).Trim().ToLower() }
    $onlyUrl = $false
    if ($OnlyWithUrlBox) { $onlyUrl = [bool]$OnlyWithUrlBox.IsChecked }

    $pairs = Get-PairsFromBoxes -AttrBoxes $AttrBoxes -ValBoxes $ValBoxes
    $cache = $null
    if ($pairs.Count -gt 0) {
        if ($Kind -eq 'User') { foreach ($pr in $pairs) { Read-UserAttributeValues -Attr $pr.Attr }; $cache = $script:UserAttrFilterCache }
        else                  { foreach ($pr in $pairs) { Read-GroupAttributeValues -Attr $pr.Attr }; $cache = $script:AttrFilterCache }
    }

    foreach ($item in (Get-TreeItem -Parent $Tree)) {
        if ($item.Tag.Type -ne 'GroupLeaf') { continue }
        $visible = $true
        if ($onlyUrl -and -not $item.Tag.HasUrl) { $visible = $false }
        if ($visible -and $q -and ($item.Tag.Text.Text.ToLower().IndexOf($q) -lt 0)) { $visible = $false }
        if ($visible -and $pairs.Count -gt 0) {
            $key = if ($Kind -eq 'User') { ([string]$item.Tag.Sam).ToLower() } else { ([string]$item.Tag.DN).ToLower() }
            if (-not (Test-AttrPairs -Pairs $pairs -Cache $cache -Key $key)) { $visible = $false }
        }
        $item.Visibility = if ($visible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    }

    function Test-HasVisibleGroup($node) {
        foreach ($ch in $node.Items) {
            if (-not ($ch -is [System.Windows.Controls.TreeViewItem])) { continue }
            if ($ch.Tag.Type -eq 'GroupLeaf') {
                if ($ch.Visibility -eq [System.Windows.Visibility]::Visible) { return $true }
            }
            elseif (Test-HasVisibleGroup $ch) { return $true }
        }
        return $false
    }
    $active = ($q -or $onlyUrl -or $pairs.Count -gt 0)
    foreach ($item in (Get-TreeItem -Parent $Tree)) {
        if ($item.Tag.Type -ne 'OU') { continue }
        if ($active) {
            $has = Test-HasVisibleGroup $item
            $item.Visibility = if ($has) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
            if ($has -and ($q -or $pairs.Count -gt 0)) { $item.IsExpanded = $true }
        } else {
            $item.Visibility = [System.Windows.Visibility]::Visible
        }
    }
}

# ---------- Create tab: source tree + counters ----------

function Update-GroupSelectedCount {
    if (-not $script:GroupTree) { return }
    $n = @(Get-CheckedGroups $script:GroupTree).Count
    # Keep the preview counter in sync so an outdated "Groups to create" is flagged.
    if (Get-Command Update-PreviewCount -ErrorAction SilentlyContinue) { Update-PreviewCount -LiveSel $n }
}
function Get-CheckedGroupNode { return @(Get-CheckedGroups $script:GroupTree) }

function Read-GroupAttributeValues {
    # Loads (once, then caches) the values of an AD attribute for every group that
    # holds it, so the attribute filter can match without re-querying per keystroke.
    param([string]$Attr)
    if (-not $Attr) { return }
    $k = $Attr.ToLower()
    if ($null -eq $script:AttrFilterCache) { $script:AttrFilterCache = @{} }
    if ($script:AttrFilterCache.ContainsKey($k)) { return }
    $map = @{}
    try {
        $p = @{ SearchBase = $script:LoadedRootDN; SearchScope = 'Subtree'; LDAPFilter = "($Attr=*)"; Properties = $Attr; ErrorAction = 'Stop' }
        if ($script:LoadedServer) { $p.Server = $script:LoadedServer }
        foreach ($g in @(Get-ADGroup @p)) {
            $txt = Convert-AttrToText $g.$Attr
            if ($txt) { $map[([string]$g.DistinguishedName).ToLower()] = $txt }
        }
        Write-Log ("Attribute filter: '{0}' loaded ({1} group(s) have a value)." -f $Attr, $map.Count) 'INFO'
    } catch {
        Write-Log ("Attribute filter: could not load attribute '{0}': {1}" -f $Attr, $_.Exception.Message) 'WARN'
    }
    $script:AttrFilterCache[$k] = $map
}

function Read-UserAttributeValues {
    # Same as Read-GroupAttributeValues, for users; cache keyed by sAMAccountName.
    param([string]$Attr)
    if (-not $Attr) { return }
    $k = $Attr.ToLower()
    if ($null -eq $script:UserAttrFilterCache) { $script:UserAttrFilterCache = @{} }
    if ($script:UserAttrFilterCache.ContainsKey($k)) { return }
    $map = @{}
    $base = if ($script:LoadedRootDN) { $script:LoadedRootDN } elseif ($script:TxtRoot) { ([string]$script:TxtRoot.Text).Trim() } else { '' }
    if (-not $base) { $script:UserAttrFilterCache[$k] = $map; return }
    try {
        $p = @{ SearchBase = $base; SearchScope = 'Subtree'; LDAPFilter = "($Attr=*)"; Properties = @($Attr,'sAMAccountName'); ErrorAction = 'Stop' }
        if ($script:LoadedServer) { $p.Server = $script:LoadedServer }
        foreach ($u in @(Get-ADUser @p)) {
            $txt = Convert-AttrToText $u.$Attr
            if ($txt -and $u.SamAccountName) { $map[([string]$u.SamAccountName).ToLower()] = $txt }
        }
        Write-Log ("User attribute filter: '{0}' loaded ({1} user(s) have a value)." -f $Attr, $map.Count) 'INFO'
    } catch {
        Write-Log ("User attribute filter: could not load attribute '{0}': {1}" -f $Attr, $_.Exception.Message) 'WARN'
    }
    $script:UserAttrFilterCache[$k] = $map
}

function Get-PairsFromBoxes {
    # Reads attribute/value textbox pairs. $AttrBoxes/$ValBoxes are parallel arrays
    # of TextBox controls. A pair is active only when its attribute box has text.
    param($AttrBoxes, $ValBoxes)
    $pairs = New-Object System.Collections.Generic.List[object]
    $a = @($AttrBoxes); $v = @($ValBoxes)
    for ($i = 0; $i -lt $a.Count; $i++) {
        if (-not $a[$i]) { continue }
        $attr = ([string]$a[$i].Text).Trim()
        if (-not $attr) { continue }
        $val = if ($i -lt $v.Count -and $v[$i]) { ([string]$v[$i].Text).Trim().ToLower() } else { '' }
        $pairs.Add(@{ Attr = $attr; Value = $val })
    }
    return @($pairs.ToArray())
}

function Test-AttrPairs {
    # AND across pairs, case-insensitive "contains". Empty value = "has any value".
    # $Cache is attrLower -> (keyLower -> text); $Key is the object's DN or sAM (lower).
    param($Pairs, $Cache, [string]$Key)
    foreach ($pr in @($Pairs)) {
        $m = if ($Cache) { $Cache[$pr.Attr.ToLower()] } else { $null }
        $aval = ''
        if ($m -and $Key -and $m.ContainsKey($Key)) { $aval = [string]$m[$Key] }
        $ok = if ($pr.Value) { $aval.ToLower().Contains($pr.Value) } else { [bool]$aval }
        if (-not $ok) { return $false }
    }
    return $true
}

function Update-SelFilter {
    # Shared selection tree: name text + up to three AD attribute/value pairs (AND, contains).
    if (-not $script:SelTree) { return }
    Update-CheckGroupTreeFilter -Tree $script:SelTree -FilterBox $script:TxtSelFilter `
        -AttrBoxes @($script:TxtSelAttr1, $script:TxtSelAttr2, $script:TxtSelAttr3) `
        -ValBoxes  @($script:TxtSelVal1,  $script:TxtSelVal2,  $script:TxtSelVal3) `
        -Kind 'Group'
    $n = 0
    foreach ($b in @($script:TxtSelAttr1, $script:TxtSelAttr2, $script:TxtSelAttr3)) { if ($b -and ([string]$b.Text).Trim()) { $n++ } }
    if ($script:ExpSelAttr) { $script:ExpSelAttr.Header = $(if ($n) { "Attribute filter ($n active)" } else { 'Attribute filter' }) }
    Update-SelCount
}
function Update-GroupTreeFilter { Update-SelFilter }

function Update-SelCount {
    # The shared tree feeds every group tool: refresh their counters. The attribute
    # list of Update attributes reads AD, so it is only refreshed while that tool is
    # open (and when it is opened).
    if (-not $script:SelTree) { return }
    $n = @(Get-CheckedGroups $script:SelTree).Count
    if ($script:TxtSelCount) { $script:TxtSelCount.Text = [string]$n }
    Update-GroupSelectedCount
    Update-AnalyseCount
    if ($script:ActiveTool -eq 'update') { Update-UrlCount }
}

# ---------- Target tree (OU + groups, single select) ----------

function New-OuNode {
    param(
        [Parameter(Mandatory)] [string]$DN,
        [Parameter(Mandatory)] [string]$Label,
        [ValidateSet('OU','GroupLeaf')] [string]$Type = 'OU',
        [string]$Name = '', [string]$Sam = ''
    )
    $item = New-Object System.Windows.Controls.TreeViewItem
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Label; $tb.VerticalAlignment = 'Center'
    if ($Type -eq 'OU') { $tb.FontWeight = 'SemiBold' }
    $item.Header = $tb
    $item.Tag = @{ Type = $Type; DN = $DN; Name = $Name; Sam = $Sam }
    return $item
}

function Build-OuTree {
    # Target tree: the OU structure plus groups as leaves.
    param([Parameter(Mandatory)] [string]$RootDN, [Parameter(Mandatory)] [string]$RootName, [object[]]$Ous, [object[]]$Groups)
    if (-not $script:OuTree) { return }
    $script:OuTree.Items.Clear()
    $rootLower = $RootDN.ToLower()

    $rootNode = New-OuNode -DN $RootDN -Label $RootName -Type 'OU'
    [void]$script:OuTree.Items.Add($rootNode)
    $map = @{}; $map[$rootLower] = $rootNode

    function Get-OrCreateTargetOu($dn) {
        $dnl = $dn.ToLower()
        if ($dnl -eq $rootLower -or -not $dnl.EndsWith($rootLower)) { return $rootNode }
        if ($map.ContainsKey($dnl)) { return $map[$dnl] }
        $parent = ($dn -split ',', 2)[1]
        $pNode  = Get-OrCreateTargetOu $parent
        $label  = ($dn -split ',', 2)[0] -replace '^OU=', '' -replace '^CN=', ''
        $n = New-OuNode -DN $dn -Label $label -Type 'OU'
        [void]$pNode.Items.Add($n); $map[$dnl] = $n
        return $n
    }

    $sorted = $Ous | Sort-Object @{ Expression = { ($_.DistinguishedName -split ',').Count } }, DistinguishedName
    foreach ($ou in $sorted) {
        $dnl = $ou.DistinguishedName.ToLower()
        if ($map.ContainsKey($dnl)) { continue }
        $parentDn   = (($ou.DistinguishedName -split ',', 2)[1]).ToLower()
        $parentNode = if ($map.ContainsKey($parentDn)) { $map[$parentDn] } else { $rootNode }
        $n = New-OuNode -DN $ou.DistinguishedName -Label $ou.Name -Type 'OU'
        [void]$parentNode.Items.Add($n)
        $map[$dnl] = $n
    }

    if ($Groups) {
        foreach ($g in ($Groups | Sort-Object Name)) {
            $parentDn = ($g.DistinguishedName -split ',', 2)[1]
            $pNode = Get-OrCreateTargetOu $parentDn
            $leaf = New-OuNode -DN $g.DistinguishedName -Label ('{0}  ({1})' -f $g.Name, $g.SamAccountName) -Type 'GroupLeaf' -Name $g.Name -Sam $g.SamAccountName
            [void]$pNode.Items.Add($leaf)
        }
    }
    $rootNode.IsExpanded = $true
}

function Update-OuTreeFilter {
    if (-not $script:OuTree) { return }
    $q = ''
    if ($script:TxtTargetFilter) { $q = ([string]$script:TxtTargetFilter.Text).Trim().ToLower() }

    foreach ($item in (Get-TreeItem -Parent $script:OuTree)) {
        if ($item.Tag.Type -ne 'GroupLeaf') { continue }
        $label = if ($item.Header -is [System.Windows.Controls.TextBlock]) { [string]$item.Header.Text } else { [string]$item.Header }
        $visible = (-not $q) -or ($label.ToLower().IndexOf($q) -ge 0)
        $item.Visibility = if ($visible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    }
    function Test-HasVisibleTargetGroup($node) {
        foreach ($ch in $node.Items) {
            if (-not ($ch -is [System.Windows.Controls.TreeViewItem])) { continue }
            if ($ch.Tag.Type -eq 'GroupLeaf') { if ($ch.Visibility -eq [System.Windows.Visibility]::Visible) { return $true } }
            elseif (Test-HasVisibleTargetGroup $ch) { return $true }
        }
        return $false
    }
    foreach ($item in (Get-TreeItem -Parent $script:OuTree)) {
        if ($item.Tag.Type -ne 'OU') { continue }
        if ($q) {
            $has = Test-HasVisibleTargetGroup $item
            $item.Visibility = if ($has) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
            if ($has) { $item.IsExpanded = $true }
        } else {
            $item.Visibility = [System.Windows.Visibility]::Visible
        }
    }
}

# ---------- Load Active Directory ----------

function Connect-AndLoadAD {
    param([Parameter(Mandatory)] [string]$RootDN)
    if ([string]::IsNullOrWhiteSpace($RootDN)) { throw 'Enter a root (SearchBase) first.' }

    Set-Status 'Connecting to Active Directory ...' -Busy
    Write-Log "Connecting, reading $RootDN" 'INFO'

    $server = $null
    if ($script:Config -and $script:Config.domain.server) { $server = [string]$script:Config.domain.server }
    $script:LoadedRootDN = $RootDN
    $script:LoadedServer = $server
    $script:AttrFilterCache = @{}
    $script:UserAttrFilterCache = @{}

    $rootName = ($RootDN -split ',', 2)[0] -replace '^OU=', '' -replace '^DC=', ''
    try {
        $p = @{ Identity = $RootDN; Properties = 'name'; ErrorAction = 'Stop' }
        if ($server) { $p.Server = $server }
        $rootObj = Get-ADObject @p
        if ($rootObj.Name) { $rootName = $rootObj.Name }
    }
    catch {
        Set-Status 'Could not read the root.'
        Write-Log "Could not read the root: $($_.Exception.Message)" 'ERROR'
        throw "Could not read the root:`n$RootDN`n`n$($_.Exception.Message)"
    }

    Write-Log 'Reading OU structure ...' 'INFO'
    $ous = @()
    try {
        $p = @{ SearchBase = $RootDN; SearchScope = 'Subtree'; Filter = '*'; Properties = 'Name'; ErrorAction = 'Stop' }
        if ($server) { $p.Server = $server }
        $ous = @(Get-ADOrganizationalUnit @p)
    }
    catch { Write-Log "Could not read OUs: $($_.Exception.Message)" 'WARN' }

    Write-Log 'Reading groups ...' 'INFO'
    $groups = @()
    try {
        $p = @{ SearchBase = $RootDN; SearchScope = 'Subtree'; Filter = '*'; Properties = 'Name','SamAccountName','DistinguishedName','url','info','Description','GroupScope','GroupCategory','member','memberOf'; ErrorAction = 'Stop' }
        if ($server) { $p.Server = $server }
        $groups = @(Get-ADGroup @p)
    }
    catch {
        Set-Status 'Could not read groups.'
        Write-Log "Could not read groups: $($_.Exception.Message)" 'ERROR'
        throw "Could not read groups:`n$($_.Exception.Message)"
    }

    # DN lookups so each group can show (U)/(C)/(G) for its member object classes.
    Write-Log 'Reading user and computer identities (for member classification) ...' 'INFO'
    $script:UserDnSet     = New-Object System.Collections.Generic.HashSet[string]
    $script:ComputerDnSet = New-Object System.Collections.Generic.HashSet[string]
    $script:GroupDnSet    = New-Object System.Collections.Generic.HashSet[string]
    # DN -> identity details (sAM, name, UPN) so member lists and reports need no per-member lookups.
    $script:UserInfoByDn = @{}; $script:ComputerInfoByDn = @{}; $script:GroupInfoByDn = @{}; $script:MemberInfoCache = @{}
    foreach ($g in $groups) {
        $k = ([string]$g.DistinguishedName).ToLower()
        [void]$script:GroupDnSet.Add($k)
        $script:GroupInfoByDn[$k] = @{ Class = 'group'; Sam = [string]$g.SamAccountName; Name = [string]$g.Name; Upn = '' }
    }
    $userObjs = @()
    try {
        $up = @{ SearchBase = $RootDN; SearchScope = 'Subtree'; Filter = '*'; ErrorAction = 'Stop' }
        if ($server) { $up.Server = $server }
        $userObjs = @(Get-ADUser @up)
        foreach ($u in $userObjs) {
            $k = ([string]$u.DistinguishedName).ToLower()
            [void]$script:UserDnSet.Add($k)
            $script:UserInfoByDn[$k] = @{ Class = 'user'; Sam = [string]$u.SamAccountName; Name = [string]$u.Name; Upn = [string]$u.UserPrincipalName }
        }
    } catch { Write-Log "Could not read users: $($_.Exception.Message)" 'WARN' }
    try {
        $cp = @{ SearchBase = $RootDN; SearchScope = 'Subtree'; Filter = '*'; ErrorAction = 'Stop' }
        if ($server) { $cp.Server = $server }
        foreach ($c in @(Get-ADComputer @cp)) {
            $k = ([string]$c.DistinguishedName).ToLower()
            [void]$script:ComputerDnSet.Add($k)
            $script:ComputerInfoByDn[$k] = @{ Class = 'computer'; Sam = [string]$c.SamAccountName; Name = [string]$c.Name; Upn = '' }
        }
    } catch { Write-Log "Could not read computers: $($_.Exception.Message)" 'WARN' }

    $script:GroupCache.Clear()
    $script:CheckMap.Clear()

    # One shared selection tree for Create, Rename, Update attributes and Snapshot groups.
    New-CheckGroupTree -Tree $script:SelTree     -TreeKind 'Shared'  -RootDN $RootDN -RootName $rootName -Groups $groups
    Build-OuTree       -RootDN $RootDN -RootName $rootName -Ous $ous -Groups $groups
    New-CheckGroupTree -Tree $script:CompareTree -TreeKind 'Compare' -RootDN $RootDN -RootName $rootName -Groups $groups
    Build-CopyPickTree -RootDN $RootDN -RootName $rootName -Groups $groups
    Clear-CopySelection
    try { Build-UserCompareTree -RootDN $RootDN -RootName $rootName -Users $userObjs } catch { Write-Log "Could not build the user tree: $($_.Exception.Message)" 'WARN' }

    # Snapshot users tab: load the users (with group membership) as part of connecting.
    try { Invoke-SnapLoad -Quiet } catch { Write-Log "Could not load snapshot users: $($_.Exception.Message)" 'WARN' }

    Set-Subtitle ("Connected · {0} · {1} groups · {2} users" -f $(if ($server) { $server } else { $rootName }), $groups.Count, $script:UserDnSet.Count)
    Set-Status  ("Loaded: {0} groups in {1} OU." -f $groups.Count, $ous.Count)
    Write-Log   ("Done: {0} groups, {1} OU, {2} users, {3} computers." -f $groups.Count, $ous.Count, $script:UserDnSet.Count, $script:ComputerDnSet.Count) 'OK'

    Update-SelFilter
    Update-CompareSelectedCount
    Update-UserCompareCount
}
