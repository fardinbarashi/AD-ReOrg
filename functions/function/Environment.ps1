<#
    ENVIRONMENT tools: OU structure and Schema attributes.

    Both are front-ends for the bundled scripts in tools\ (kept unchanged, so the
    same code runs as from the command line):
      tools\ou-structure\Export-OUStructure.ps1 / Import-OUStructure.ps1
      tools\schema-attributes\Export-SchemaAttributes.ps1 / Import-SchemaAttributes.ps1

    Each script runs in its own PowerShell process (tools\Invoke-ToolScript.ps1). Its
    output is shown live under "Script output"; the files it writes (CSV / JSON /
    report) are read back into the "Result" grid. Files go to files\ou-structure and
    files\schema-attributes.

    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

# ---------- Running a bundled script ----------

function Get-ToolScriptPath {
    param([Parameter(Mandatory)] [string]$Folder, [Parameter(Mandatory)] [string]$Name)
    $p = Join-Path $script:ScriptRoot (Join-Path 'tools' (Join-Path $Folder $Name))
    if (-not (Test-Path -LiteralPath $p)) { throw "The script is missing: $p" }
    return $p
}

function Format-ToolParams {
    # "-Path 'C:\x.csv' -WhatIf" - for the log and the output header.
    param($Params)
    $parts = foreach ($k in $Params.Keys) {
        $v = $Params[$k]
        if ($v -is [bool]) { if ($v) { "-$k" } }
        elseif ($v -is [System.Array]) { "-$k " + (($v | ForEach-Object { "'$_'" }) -join ',') }
        else { "-$k '$v'" }
    }
    return (@($parts) -join ' ')
}

function Read-SharedText {
    # Reads a file another process is still writing (UTF-8).
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        # ANSI colour codes (PowerShell 7) are removed so lines can be read as text.
        try { $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8); return ($sr.ReadToEnd() -replace "$([char]27)\[[0-9;]*[A-Za-z]", '') } finally { $fs.Dispose() }
    } catch { return '' }
}

function Invoke-ToolScript {
    # Runs one bundled script in its own PowerShell process and streams its output
    # into $OutBox. Returns ExitCode, Output (stdout) and Errors (stderr).
    param(
        [Parameter(Mandatory)] [string]$Script,
        [Parameter(Mandatory)] $Params,
        $OutBox,
        [string]$Title = 'Running script'
    )
    $runner = Get-ToolScriptPath -Folder '.' -Name 'Invoke-ToolScript.ps1'
    $work = Join-Path ([System.IO.Path]::GetTempPath()) ('ADReOrg_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $argsFile = Join-Path $work 'args.json'
    $outFile  = Join-Path $work 'out.txt'
    $errFile  = Join-Path $work 'err.txt'
    $plain = [ordered]@{}; foreach ($k in $Params.Keys) { $plain[$k] = $Params[$k] }
    [System.IO.File]::WriteAllText($argsFile, ($plain | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding $false))

    $exe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $argLine = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Script "{1}" -ArgsFile "{2}"' -f $runner, $Script, $argsFile
    $header = ('> {0} {1}' -f [System.IO.Path]::GetFileName($Script), (Format-ToolParams $Params)) + [Environment]::NewLine + [Environment]::NewLine
    Write-Log ('Running {0}' -f $header.Trim()) 'INFO'
    if ($OutBox) { $OutBox.Text = $header }

    $proc = Start-Process -FilePath $exe -ArgumentList $argLine -RedirectStandardOutput $outFile -RedirectStandardError $errFile -WindowStyle Hidden -PassThru
    $null = $proc.Handle          # keeps ExitCode readable after the process ends
    if ($script:Pb) { $script:Pb.IsIndeterminate = $true; $script:Pb.Visibility = [System.Windows.Visibility]::Visible }
    Set-Status ('{0} ...' -f $Title) -Busy
    $shown = ''
    try {
        while (-not $proc.HasExited) {
            Start-Sleep -Milliseconds 250
            $now = (Read-SharedText $outFile) + (Read-SharedText $errFile)
            if ($OutBox -and $now -ne $shown) { $OutBox.Text = $header + $now; $OutBox.ScrollToEnd(); $shown = $now }
            Invoke-DoEvents
        }
        $proc.WaitForExit()
    }
    finally {
        if ($script:Pb) { $script:Pb.IsIndeterminate = $false; $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed }
    }
    $out = Read-SharedText $outFile
    $err = Read-SharedText $errFile
    if ($OutBox) { $OutBox.Text = $header + $out + $(if ($err) { [Environment]::NewLine + $err } else { '' }); $OutBox.ScrollToEnd() }
    $code = $proc.ExitCode
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    Set-Status ('{0} - finished (exit code {1}).' -f $Title, $code)
    foreach ($l in @($err -split "`r?`n" | Where-Object { $_.Trim() })) { Write-Log ('  ' + $l) 'WARN' }
    return [pscustomobject]@{ ExitCode = $code; Output = $out; Errors = $err }
}

function Get-ToolFailureText {
    param($Result)
    $lines = @(([string]$Result.Errors + "`n" + [string]$Result.Output) -split "`r?`n" | Where-Object { $_ -match 'ERROR|Exception|not found|cannot|denied|failed' })
    if ($lines.Count -gt 0) { return ($lines | Select-Object -Last 3) -join "`n" }
    return ('The script ended with exit code {0}.' -f $Result.ExitCode)
}

function Open-ToolFolder {
    param([string]$Sub)
    $f = Get-ReportFolder -Sub $Sub
    Start-Process explorer.exe -ArgumentList ('"{0}"' -f $f)
}

function Select-ToolFile {
    param([string]$Filter, [string]$Title, [string]$Sub)
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = $Filter; $dlg.Title = $Title
    $folder = Get-ReportFolder -Sub $Sub
    if (Test-Path -LiteralPath $folder) { $dlg.InitialDirectory = $folder }
    if ($dlg.ShowDialog() -eq $true) { return $dlg.FileName }
    return $null
}

function Get-ToolDomainDN {
    # DN of the domain this computer (or the configured server) belongs to.
    $p = @{ ErrorAction = 'Stop' }
    if ($script:LoadedServer) { $p.Server = $script:LoadedServer }
    elseif ($script:Config -and $script:Config.domain.server) { $p.Server = [string]$script:Config.domain.server }
    return [string](Get-ADDomain @p).DistinguishedName
}

# =====================================================================
#  OU STRUCTURE
# =====================================================================

function Invoke-OuExport {
    $root = ([string]$script:TxtOuExpRoot.Text).Trim()
    if (-not $root) { Show-Message -Text 'Enter the root OU - its name (e.g. Contoso) or its full DN.' -Icon Warning -Caption 'Export OU structure'; return }
    $isDn = $root -match '^(?i)(OU|DC)='
    $name = if ($isDn) { ((([regex]'(?<!\\),').Split($root, 2)[0]) -replace '^(?i)OU=', '') -replace '\\(.)', '$1' } else { $root }
    $folder = Get-ReportFolder -Sub 'ou-structure'
    $file = Join-Path $folder ('OU-Structure_{0}_{1}.csv' -f ($name -replace '[\\/:*?"<>|\s]+', '_'), (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $params = [ordered]@{}
    if ($isDn) { $params.SearchBase = $root } else { $params.OUName = $root }
    $params.Path = $file

    $script:RbOuExpOutput.IsChecked = $true
    $r = Invoke-ToolScript -Script (Get-ToolScriptPath 'ou-structure' 'Export-OUStructure.ps1') -Params $params -OutBox $script:TxtOuExpOut -Title 'Exporting OU structure'
    if ($r.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $file)) {
        $script:TxtOuExpResult.Text = 'Export failed - see the script output.'
        Show-Message -Text ("The export did not finish.`n`n{0}" -f (Get-ToolFailureText $r)) -Icon Error -Caption 'Export OU structure'
        return
    }
    $rows = @(Import-Csv -LiteralPath $file -Delimiter ';' -Encoding UTF8)
    # "Include all sub-OUs" cleared: the script always exports the whole tree, so
    # the file is written again with only the root OU (the row with the lowest depth).
    $subs = (-not $script:ChkOuExpSubOus) -or [bool]$script:ChkOuExpSubOus.IsChecked
    $note = ''
    if (-not $subs -and $rows.Count -gt 1) {
        $top = ($rows | ForEach-Object { [int]$_.Depth } | Measure-Object -Minimum).Minimum
        $rows = @($rows | Where-Object { [int]$_.Depth -eq $top })
        $rows | Export-Csv -LiteralPath $file -Delimiter ';' -NoTypeInformation -Encoding UTF8
        $note = ' (sub-OUs not included)'
        if ($script:TxtOuExpOut) { $script:TxtOuExpOut.Text += ([Environment]::NewLine + 'AD ReOrg: "Include all sub-OUs" is cleared - the file now holds only the root OU.'); $script:TxtOuExpOut.ScrollToEnd() }
    }
    $cols = @( @{ Key = 'Depth' }, @{ Key = 'Name' }, @{ Key = 'RelativeDN' }, @{ Key = 'Description' }, @{ Key = 'Protected' }, @{ Key = 'SourceDN' } )
    $script:GridOuExp.ItemsSource = (ConvertTo-DataTable -Records $rows -Columns $cols).DefaultView
    $script:RbOuExpResult.IsChecked = $true
    $script:TxtOuExpResult.Text = ('{0} OU(s) exported{1} to {2}' -f $rows.Count, $note, $file)
    $script:TxtOuImpFile.Text = $file      # step 2 uses the latest export
    Write-Log ('OU structure: {0} OU(s) exported{1} to {2}' -f $rows.Count, $note, $file) 'OK'
}

function Get-OuImportParent {
    $t = ([string]$script:TxtOuImpParent.Text).Trim()
    if ($t) { return $t }
    return (Get-ToolDomainDN)
}

function Import-OuFile {
    # Reads the CSV into the grid with the DN each OU will get in this domain.
    $path = ([string]$script:TxtOuImpFile.Text).Trim().Trim('"')
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { Show-Message -Text 'Pick the CSV file from step 1 (OU-Structure_*.csv).' -Icon Warning -Caption 'Import OU structure'; return $false }
    $rows = @(Import-Csv -LiteralPath $path -Delimiter ';' -Encoding UTF8)
    $need = 'Name', 'RelativeDN', 'Depth'
    $have = if ($rows.Count) { @($rows[0].PSObject.Properties.Name) } else { @() }
    $miss = @($need | Where-Object { $have -notcontains $_ })
    if ($rows.Count -eq 0 -or $miss.Count) {
        Show-Message -Text ("This is not an OU structure file from step 1.`n`nMissing column(s): {0}" -f ($(if ($miss.Count) { $miss -join ', ' } else { '(the file is empty)' }))) -Icon Warning -Caption 'Import OU structure'
        return $false
    }
    $parent = ''
    try { $parent = Get-OuImportParent } catch { $parent = '(domain root - could not be read: ' + $_.Exception.Message + ')' }
    $script:OuImpParent = $parent
    $script:OuImpRows = @($rows | Sort-Object { [int]$_.Depth }, RelativeDN | ForEach-Object {
        [pscustomobject]@{
            Status      = 'Not run'
            Name        = $_.Name
            TargetDN    = ('{0},{1}' -f $_.RelativeDN, $parent)
            Depth       = $_.Depth
            Description = $_.Description
            Protected   = $_.Protected
        }
    })
    Set-OuImportGrid
    $script:RbOuImpResult.IsChecked = $true
    $script:TxtOuImpCount.Text = ('{0} OUs in the file  ·  created under {1}' -f $script:OuImpRows.Count, $parent)
    return $true
}

function Set-OuImportGrid {
    $cols = @( @{ Key = 'Status' }, @{ Key = 'Name' }, @{ Key = 'TargetDN' }, @{ Key = 'Depth' }, @{ Key = 'Description' }, @{ Key = 'Protected' } )
    $script:GridOuImp.ItemsSource = (ConvertTo-DataTable -Records $script:OuImpRows -Columns $cols).DefaultView
}

function Invoke-OuImport {
    if (-not (Import-OuFile)) { return }      # always re-read: the file or the parent may have changed
    $whatIf = [bool]$script:ChkOuImpWhatIf.IsChecked
    $parent = $script:OuImpParent
    if ($parent -like '(domain root*') { Show-Message -Text ('Could not read the domain root. Enter the parent DN to create the OUs under.') -Icon Warning -Caption 'Import OU structure'; return }
    $mode = if ($whatIf) { 'WhatIf (SIMULATION - nothing changes)' } else { 'LIVE RUN' }
    $q = "Create up to {0} OUs under:`n{1}`n`nOUs that already exist are skipped.`nRun mode: {2}`n`nContinue?" -f $script:OuImpRows.Count, $parent, $mode
    if ((Show-Message -Text $q -Icon Question -Buttons YesNo -Caption 'Confirm') -ne [System.Windows.MessageBoxResult]::Yes) { Set-Status 'Cancelled.'; return }

    $file = ([string]$script:TxtOuImpFile.Text).Trim().Trim('"')
    $params = [ordered]@{ Path = $file; TargetParentDN = $parent }
    if ($whatIf) { $params.WhatIf = $true }
    $script:RbOuImpOutput.IsChecked = $true
    $r = Invoke-ToolScript -Script (Get-ToolScriptPath 'ou-structure' 'Import-OUStructure.ps1') -Params $params -OutBox $script:TxtOuImpOut -Title 'Importing OU structure'

    # Statuses from the script's output, matched on the target DN.
    $byDn = @{}; foreach ($row in $script:OuImpRows) { $byDn[$row.TargetDN.ToLower()] = $row }
    foreach ($line in ([string]$r.Output + "`n" + [string]$r.Errors) -split "`r?`n") {
        $dn = $null; $st = $null
        if     ($line -match '^\s*Already exists:\s*(.+?)\s*$')                                { $dn = $Matches[1]; $st = 'Already exists' }
        elseif ($line -match '^\s*Created:\s*(.+?)\s*$')                                       { $dn = $Matches[1]; $st = 'Created' }
        elseif ($line -match '^What if: Performing the operation "Create OU" on target "(.+)"\.\s*$') { $dn = $Matches[1]; $st = 'Would create' }
        elseif ($line -match '^\s*WARNING:\s*Failed:\s*(.+?) - (.+)$')                         { $dn = $Matches[1]; $st = 'Error: ' + $Matches[2] }
        if ($dn -and $byDn.ContainsKey($dn.ToLower())) { $byDn[$dn.ToLower()].Status = $st }
    }
    foreach ($row in $script:OuImpRows) { if ($row.Status -eq 'Not run') { $row.Status = $(if ($r.ExitCode -ne 0) { 'Not run (the script stopped)' } else { 'No result' }) } }
    Set-OuImportGrid
    $script:RbOuImpResult.IsChecked = $true

    $cnt = { param($pat) @($script:OuImpRows | Where-Object { $_.Status -like $pat }).Count }
    $verb = if ($whatIf) { 'would be created' } else { 'created' }
    $made = if ($whatIf) { & $cnt 'Would create' } else { & $cnt 'Created' }
    $result = ('{0} {1}, {2} already existed, {3} error(s).' -f $made, $verb, (& $cnt 'Already exists'), (& $cnt 'Error*'))
    if ($r.ExitCode -ne 0) { $result = 'The script stopped: ' + (Get-ToolFailureText $r) }
    $script:TxtOuImpResult.Text = $result
    Write-Log ('OU import: ' + $result) $(if ($r.ExitCode -ne 0) { 'WARN' } else { 'OK' })

    # Report
    $cols = @( @{ Key = 'Status'; Label = 'Status' }, @{ Key = 'Name'; Label = 'Name' }, @{ Key = 'TargetDN'; Label = 'Target DN' }, @{ Key = 'Description'; Label = 'Description' }, @{ Key = 'Protected'; Label = 'Protected' } )
    $settings = [ordered]@{ 'Run mode' = $mode; 'CSV file' = $file; 'Created under' = $parent; 'Result' = $result }
    $banner = if ($whatIf) { '<div class="banner wi">SIMULATION (WhatIf) - nothing was created.</div>' } else { '<div class="banner ok">Live run - the OUs marked Created were created.</div>' }
    $base = Join-Path (Get-ReportFolder -Sub 'ou-structure') ('ou_import_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        New-HtmlReport -Title 'Import OU structure - result' -Columns $cols -Records $script:OuImpRows -Path ($base + '.html') -StatusKey 'Status' -Banner $banner -Unit 'OUs' -Settings $settings -HideChangedBox
        Export-DataCsv -Columns $cols -Records $script:OuImpRows -Path ($base + '.csv')
        Export-DataJson -Records $script:OuImpRows -Path ($base + '.json') -Meta @{ type = 'ou-import'; whatIf = $whatIf; parent = $parent; file = $file }
        Write-Log ('Report: {0}.html' -f $base) 'OK'
        if ($script:Config -and $script:Config.report.openAfter) { Start-Process ($base + '.html') }
    } catch { Write-Log ('OU import report failed: {0}' -f $_.Exception.Message) 'WARN' }
}

# =====================================================================
#  SCHEMA ATTRIBUTES
# =====================================================================

function New-UserAttributeList {
    # The attribute list of a user WITHOUT the values (no personal data): every
    # attribute a user object can have (allowedAttributes) and whether this user
    # has a value in it. Written as Attribute;InUse for Export-SchemaAttributes.ps1.
    param([Parameter(Mandatory)] [string]$Identity, [Parameter(Mandatory)] [string]$Path)
    $p = @{ Identity = $Identity; Properties = '*'; ErrorAction = 'Stop' }
    $server = $script:LoadedServer
    if ($server) { $p.Server = $server }
    $u = Get-ADUser @p
    $ldap = if ($server) { "LDAP://$server/$($u.DistinguishedName)" } else { "LDAP://$($u.DistinguishedName)" }
    $de = [ADSI]$ldap
    $de.RefreshCache([string[]]'allowedAttributes')
    $rows = @($de.Properties['allowedAttributes'] | Sort-Object | ForEach-Object {
        $v = $u.$_
        [pscustomobject]@{ Attribute = $_; InUse = (-not [string]::IsNullOrEmpty("$v")) }
    })
    $rows | Export-Csv -LiteralPath $Path -Delimiter ';' -NoTypeInformation -Encoding UTF8
    return $rows.Count
}

function Invoke-SchemaExport {
    $folder = Get-ReportFolder -Sub 'schema-attributes'
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $json = Join-Path $folder ('SchemaAttributes_{0}.json' -f $stamp)
    $onlyUsed = [bool]$script:ChkSchOnlyUsed.IsChecked
    $params = [ordered]@{}

    if ($script:RbSchFromUser.IsChecked) {
        $user = ([string]$script:TxtSchUser.Text).Trim()
        if (-not $user) { Show-Message -Text 'Enter the sAMAccountName of a user that has the attributes you want in the test environment.' -Icon Warning -Caption 'Export schema attributes'; return }
        if ($onlyUsed) {
            # The script's -OnlyUsed works on a CSV: build one (names + InUse, no values).
            $csv = Join-Path $folder ('UserAttributes_{0}_{1}.csv' -f ($user -replace '[\\/:*?"<>|\s]+', '_'), $stamp)
            Set-Status ('Reading the attributes of {0} ...' -f $user) -Busy
            try { $n = New-UserAttributeList -Identity $user -Path $csv }
            catch { Set-Status 'Error.'; Show-Message -Text ("Could not read the user '{0}':`n{1}" -f $user, $_.Exception.Message) -Icon Error -Caption 'Export schema attributes'; return }
            Write-Log ('Schema export: attribute list of {0} ({1} attributes, no values) -> {2}' -f $user, $n, $csv) 'INFO'
            $params.CsvPath = $csv; $params.OnlyUsed = $true
        } else {
            $params.Identity = $user
        }
    } else {
        $csv = ([string]$script:TxtSchCsv.Text).Trim().Trim('"')
        if (-not $csv -or -not (Test-Path -LiteralPath $csv)) { Show-Message -Text 'Pick the CSV file with the attribute names (column Attribute).' -Icon Warning -Caption 'Export schema attributes'; return }
        $params.CsvPath = $csv
        if ($onlyUsed) { $params.OnlyUsed = $true }
    }
    $params.Path = $json

    $script:RbSchExpOutput.IsChecked = $true
    $r = Invoke-ToolScript -Script (Get-ToolScriptPath 'schema-attributes' 'Export-SchemaAttributes.ps1') -Params $params -OutBox $script:TxtSchExpOut -Title 'Exporting schema attributes'
    if ($r.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $json)) {
        $script:TxtSchExpResult.Text = 'Export failed - see the script output.'
        Show-Message -Text ("The export did not finish.`n`n{0}" -f (Get-ToolFailureText $r)) -Icon Error -Caption 'Export schema attributes'
        return
    }
    $data = Read-SchemaExportFile -Path $json
    if (-not $data) { return }
    $script:GridSchExp.ItemsSource = (ConvertTo-DataTable -Records $data.Rows -Columns $data.Columns).DefaultView
    $script:RbSchExpResult.IsChecked = $true
    $script:TxtSchExpResult.Text = ('{0}  ·  saved to {1}' -f $data.Summary, $json)
    $script:TxtSchImpFile.Text = $json     # step 2 uses the latest export
    Write-Log ('Schema export: {0} -> {1}' -f $data.Summary, $json) 'OK'
}

function Read-SchemaExportFile {
    # Reads SchemaAttributes_*.json into grid rows (attributes, then classes).
    param([Parameter(Mandatory)] [string]$Path)
    try { $d = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json }
    catch { Show-Message -Text ("Could not read the file:`n{0}" -f $_.Exception.Message) -Icon Error -Caption 'Schema attributes'; return $null }
    if (-not $d.PSObject.Properties['Attributes']) { Show-Message -Text 'This is not a schema export from step 1 (no Attributes list).' -Icon Warning -Caption 'Schema attributes'; return $null }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($a in @($d.Attributes)) {
        $def = $a.Definition
        $rows.Add([pscustomobject]@{
            Type = 'Attribute'; Name = $a.Name; Category = $a.Category; Role = $a.Role; Holder = $a.Holder
            Syntax = $def.attributeSyntax; SingleValued = $def.isSingleValued; LinkID = $def.linkID; OID = $def.attributeID
        })
    }
    foreach ($c in @($d.Classes)) {
        $rows.Add([pscustomobject]@{
            Type = 'Class'; Name = $c.Name; Category = $c.Category; Role = $c.Reason; Holder = $c.Related
            Syntax = ''; SingleValued = ''; LinkID = ''; OID = $c.Definition.governsID
        })
    }
    $attrs = @($d.Attributes)
    $byCat = @($attrs | Where-Object { $_.Role -ne 'ForwardLink' } | Group-Object Category | Sort-Object Name | ForEach-Object { '{0} {1}' -f $_.Count, $_.Name })
    $fwd = @($attrs | Where-Object { $_.Role -eq 'ForwardLink' }).Count
    $summary = ('{0} attributes ({1}), {2} forward links, {3} classes' -f @($attrs | Where-Object { $_.Role -ne 'ForwardLink' }).Count, ($byCat -join ', '), $fwd, @($d.Classes).Count)
    $from = if ($d.ExportedFrom) { (' - exported from {0} at {1}' -f $d.ExportedFrom, $d.ExportedAt) } else { '' }
    $cols = @( @{ Key = 'Type' }, @{ Key = 'Name' }, @{ Key = 'Category' }, @{ Key = 'Role' }, @{ Key = 'Holder' }, @{ Key = 'Syntax' }, @{ Key = 'SingleValued' }, @{ Key = 'LinkID' }, @{ Key = 'OID' } )
    return [pscustomobject]@{ Rows = @($rows.ToArray()); Columns = $cols; Summary = $summary; From = $from; Data = $d }
}

function Import-SchemaFile {
    $path = ([string]$script:TxtSchImpFile.Text).Trim().Trim('"')
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { Show-Message -Text 'Pick the JSON file from step 1 (SchemaAttributes_*.json).' -Icon Warning -Caption 'Import schema attributes'; return $false }
    $data = Read-SchemaExportFile -Path $path
    if (-not $data) { return $false }
    $script:GridSchImp.ItemsSource = (ConvertTo-DataTable -Records $data.Rows -Columns $data.Columns).DefaultView
    $script:RbSchImpResult.IsChecked = $true
    $script:TxtSchImpCount.Text = ($data.Summary + $data.From)
    return $true
}

function Invoke-SchemaImport {
    if (-not (Import-SchemaFile)) { return }
    $path = ([string]$script:TxtSchImpFile.Text).Trim().Trim('"')
    $whatIf = [bool]$script:ChkSchImpWhatIf.IsChecked
    $only = @(([string]$script:TxtSchOnly.Text) -split '[,;\s]+' | Where-Object { $_ })
    $exclExch = [bool]$script:ChkSchExclExchange.IsChecked
    if ($null -eq $script:SchWhatIfDone) { $script:SchWhatIfDone = @{} }

    if ($whatIf) {
        $q = "Simulate the schema import (WhatIf) - nothing is changed.`n`nFile: {0}`nAttributes: {1}`nExchange: {2}`n`nContinue?" -f $path, $(if ($only) { $only -join ', ' } else { 'all in the file' }), $(if ($exclExch) { 'skipped' } else { 'included' })
        if ((Show-Message -Text $q -Icon Question -Buttons YesNo -Caption 'Confirm') -ne [System.Windows.MessageBoxResult]::Yes) { Set-Status 'Cancelled.'; return }
    } else {
        if (-not $script:SchWhatIfDone[$path.ToLower()]) {
            $q = "This file has not been run with WhatIf in this session.`n`nRun it with WhatIf first and read the result before extending the schema.`n`nContinue with a LIVE run anyway?"
            if ((Show-Message -Text $q -Icon Warning -Buttons YesNo -Caption 'No WhatIf run') -ne [System.Windows.MessageBoxResult]::Yes) { Set-Status 'Cancelled.'; return }
        }
        $q = "LIVE RUN - the schema of this forest will be extended.`n`nSchema changes are PERMANENT: attributes and classes can be deactivated but never deleted.`n`nFile: {0}`nAttributes: {1}`nExchange: {2}`n`nYou need to be a member of Schema Admins, and a snapshot/backup of the domain controller should exist.`n`nExtend the schema now?" -f $path, $(if ($only) { $only -join ', ' } else { 'all in the file' }), $(if ($exclExch) { 'skipped' } else { 'included' })
        if ((Show-Message -Text $q -Icon Warning -Buttons YesNo -Caption 'Extend the schema - permanent') -ne [System.Windows.MessageBoxResult]::Yes) { Set-Status 'Cancelled.'; return }
    }

    $report = Join-Path (Get-ReportFolder -Sub 'schema-attributes') ('SchemaImport-Report_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $params = [ordered]@{ Path = $path; ReportPath = $report; Force = $true }   # Force: the confirmation was given above
    if ($only)     { $params.Attribute = [string[]]$only }
    if ($exclExch) { $params.ExcludeExchange = $true }
    if ($whatIf)   { $params.WhatIf = $true }

    $script:RbSchImpOutput.IsChecked = $true
    $r = Invoke-ToolScript -Script (Get-ToolScriptPath 'schema-attributes' 'Import-SchemaAttributes.ps1') -Params $params -OutBox $script:TxtSchImpOut -Title $(if ($whatIf) { 'Simulating the schema import' } else { 'Extending the schema' })
    if ($whatIf -and $r.ExitCode -eq 0) { $script:SchWhatIfDone[$path.ToLower()] = $true }

    if (-not (Test-Path -LiteralPath $report)) {
        $script:TxtSchImpResult.Text = 'No report was written - see the script output.'
        Show-Message -Text ("The import did not finish.`n`n{0}" -f (Get-ToolFailureText $r)) -Icon Error -Caption 'Import schema attributes'
        return
    }
    $rows = @(Import-Csv -LiteralPath $report -Delimiter ';' -Encoding UTF8)
    $cols = @( @{ Key = 'Status'; Label = 'Status' }, @{ Key = 'Type'; Label = 'Type' }, @{ Key = 'Name'; Label = 'Name' }, @{ Key = 'Message'; Label = 'Message' } )
    $script:GridSchImp.ItemsSource = (ConvertTo-DataTable -Records $rows -Columns $cols).DefaultView
    $script:RbSchImpResult.IsChecked = $true
    $sum = @($rows | Group-Object Status | Sort-Object Count -Descending | ForEach-Object { '{0} {1}' -f $_.Count, $_.Name }) -join ', '
    $avail = [regex]::Match([string]$r.Output, 'Available on user:\s*(\d+ of \d+)').Groups[1].Value
    $result = ('{0}{1}{2}' -f $(if ($whatIf) { 'WhatIf: ' } else { '' }), $sum, $(if ($avail) { "  ·  available on user: $avail" } else { '' }))
    $script:TxtSchImpResult.Text = $result
    $script:TxtSchImpCount.Text = ('Report: {0}' -f $report)
    Write-Log ('Schema import: ' + $result) 'OK'

    $banner = if ($whatIf) { '<div class="banner wi">SIMULATION (WhatIf) - the schema was not changed.</div>' } else { '<div class="banner ok">Live run - the schema was extended as listed below.</div>' }
    $settings = [ordered]@{ 'Run mode' = $(if ($whatIf) { 'WhatIf' } else { 'Live run' }); 'File' = $path; 'Attributes' = $(if ($only) { $only -join ', ' } else { 'All in the file' }); 'Exchange' = $(if ($exclExch) { 'Skipped' } else { 'Included' }); 'Result' = $result }
    try {
        $hp = [System.IO.Path]::ChangeExtension($report, '.html')
        New-HtmlReport -Title 'Import schema attributes - result' -Columns $cols -Records $rows -Path $hp -StatusKey 'Status' -Banner $banner -Unit 'rows' -Settings $settings -HideChangedBox
        if ($script:Config -and $script:Config.report.openAfter) { Start-Process $hp }
    } catch { Write-Log ('Schema report (HTML) failed: {0}' -f $_.Exception.Message) 'WARN' }
}
