<#
    Snapshot groups (formerly Group analyser): pick OUs/groups in the shared
    selection, optionally filter on an AD attribute, and export HTML / CSV / JSON. Each row holds the group name,
    its OU placement, its members, the groups it is a member of (memberOf) and the
    value of the attribute you filter on.
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

function Update-AnalyseCount {
    if (-not $script:AnalyseTree) { return }
    $n = @(Get-CheckedGroups $script:AnalyseTree).Count
    if ($script:TxtAnalyseCount) { $script:TxtAnalyseCount.Text = "Selected groups: $n" }
}

function Get-DnCnList {
    # Array of DNs -> sorted list of CN (short) names.
    param($DnValues)
    if (-not $DnValues) { return @() }
    return @(@($DnValues) | ForEach-Object { Get-CnFromDn $_ } | Where-Object { $_ } | Sort-Object)
}

function Export-AnalyseReport {
    param([ValidateSet('Html','Csv','Json','All')] [string]$Kind = 'All')
    $checked = @(Get-CheckedGroups $script:AnalyseTree)
    if ($checked.Count -eq 0) { Show-Message -Text 'Check at least one OU or group in the selection to the left first.' -Icon Warning; return }

    $attr = if ($script:TxtAnalyseAttr) { ([string]$script:TxtAnalyseAttr.Text).Trim() } else { '' }
    $attrMap = $null
    if ($attr) { Read-GroupAttributeValues -Attr $attr; $attrMap = $script:AttrFilterCache[$attr.ToLower()] }

    Set-Status 'Building the group snapshot ...' -Busy
    $recs = New-Object System.Collections.Generic.List[object]
    foreach ($c in $checked) {
        $g = $script:GroupCache[$c.DN]
        if (-not $g) { continue }
        $attrVal = ''
        if ($attr) { $dnl = ([string]$c.DN).ToLower(); if ($attrMap -and $attrMap.ContainsKey($dnl)) { $attrVal = $attrMap[$dnl] } }
        if ($attr -and -not $attrVal) { continue }   # filter: only groups that hold the attribute
        $members  = Get-DnCnList $g.member
        $memberOf = Get-DnCnList $g.memberOf
        $recs.Add([pscustomobject]@{
            Name           = $g.Name
            OU             = (Get-UserOuPath $c.DN)
            SamAccountName = $g.SamAccountName
            Scope          = [string]$g.GroupScope
            Category       = [string]$g.GroupCategory
            MemberCount    = $members.Count
            Members        = ($members -join ' | ')
            MemberOfCount  = $memberOf.Count
            MemberOf       = ($memberOf -join ' | ')
            FilterAttr     = $attrVal
        })
    }
    $records = @($recs.ToArray() | Sort-Object OU, Name)
    if ($records.Count -eq 0) {
        Set-Status 'Nothing to report.'
        Show-Message -Text ("No groups matched. {0}" -f ($(if ($attr) { "No selected group has a value in '$attr'." } else { 'Nothing is selected.' }))) -Icon Information
        return
    }

    $cols = @(
        @{ Key='Name';           Label='Group name' }
        @{ Key='OU';             Label='OU placement' }
        @{ Key='SamAccountName'; Label='sAMAccountName' }
        @{ Key='Scope';          Label='Scope' }
        @{ Key='Category';       Label='Category' }
        @{ Key='MemberCount';    Label='Member count' }
        @{ Key='Members';        Label='Members' }
        @{ Key='MemberOfCount';  Label='memberOf count' }
        @{ Key='MemberOf';       Label='Member of (groups)' }
    )
    if ($attr) { $cols += @{ Key='FilterAttr'; Label=$attr } }

    $folder = Get-ReportFolder -Sub 'snapshot-groups'; $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'; $base = Join-Path $folder ('snapshot_groups_{0}' -f $stamp); $made = @()
    $title = if ($attr) { "Snapshot groups - filtered on '$attr'" } else { 'Snapshot groups' }
    if ($Kind -eq 'Html' -or $Kind -eq 'All') { $hp = $base + '.html'; New-HtmlReport -Title $title -Columns $cols -Records $records -Path $hp -Unit 'groups'; $made += $hp; Write-Log ("Snapshot groups HTML: {0}" -f $hp) 'OK' }
    if ($Kind -eq 'Csv'  -or $Kind -eq 'All') { $cp = $base + '.csv';  Export-DataCsv -Columns $cols -Records $records -Path $cp; $made += $cp; Write-Log ("Snapshot groups CSV: {0}" -f $cp) 'OK' }
    if ($Kind -eq 'Json' -or $Kind -eq 'All') { $jp = $base + '.json'; Export-DataJson -Records $records -Path $jp -Meta @{ type='snapshot-groups'; attribute=$attr; groups=$records.Count }; $made += $jp; Write-Log ("Snapshot groups JSON: {0}" -f $jp) 'OK' }

    Set-Status ("Group snapshot created ({0} groups)." -f $records.Count)
    $htmlMade = @($made | Where-Object { $_ -like '*.html' }) | Select-Object -First 1
    if ($htmlMade -and $script:Config.report.openAfter) { Start-Process $htmlMade }
    elseif ($made.Count -gt 0 -and $script:Config.report.openAfter) { Start-Process (Split-Path $made[0] -Parent) }
}
