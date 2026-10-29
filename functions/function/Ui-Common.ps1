<#
    Shared UI routines: status text, logging, the header subtitle, a WPF DoEvents
    pump, and small string utilities.
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

function Invoke-DoEvents {
    # WPF equivalent of WinForms DoEvents - lets the UI repaint during long loops.
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    $cb = [System.Windows.Threading.DispatcherOperationCallback]{ param($f) $f.Continue = $false; $null }
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
        [System.Windows.Threading.DispatcherPriority]::Background, $cb, $frame) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Set-Status {
    param([string]$Text, [switch]$Busy)
    if ($script:TxtStatus) { $script:TxtStatus.Text = $Text }
    if ($script:Win) {
        $script:Win.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
    }
    Invoke-DoEvents
}

function Set-Subtitle {
    # Connection chip in the top bar: green dot when connected, amber otherwise.
    param([string]$Text)
    if ($script:SubtitleText) { $script:SubtitleText.Text = $Text }
    if ($script:DotConn) {
        $c = if ($Text -like 'Connected*') { '#4ADE80' } else { '#F59E0B' }
        $script:DotConn.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($c)
    }
}

function Invoke-UiAction {
    # Every button/checkbox handler runs through this. PowerShell swallows errors
    # thrown inside WPF event handlers, so without it a failing click looks like
    # "nothing happened". Here the error is logged (the log opens) and shown.
    param([scriptblock]$Action)
    try { & $Action }
    catch {
        $msg = $_.Exception.Message
        $pos = if ($_.InvocationInfo) { $_.InvocationInfo.PositionMessage } else { '' }
        try { if ($script:Pb) { $script:Pb.Visibility = [System.Windows.Visibility]::Collapsed }; if ($script:Win) { $script:Win.Cursor = $null } } catch { }
        if ($script:TxtStatus) { $script:TxtStatus.Text = 'Error - see the log.' }
        Write-Log ("ERROR: {0}`n{1}" -f $msg, $pos) 'ERROR'
        try { Show-Message -Text ("Something went wrong:`n`n{0}`n`nThe details are in the log." -f $msg) -Icon Error -Caption 'AD ReOrg' } catch { }
    }
}

function Complete-GridEdit {
    # Commits a pending cell/row edit, so a value typed or a box clicked right before
    # a button click is in the row objects when the button's code reads them.
    param($Grid)
    if (-not $Grid) { return }
    try { [void]$Grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) } catch { }
}

function Update-GridSafe {
    # Items.Refresh() throws while a row is still being edited - commit first.
    param($Grid)
    if (-not $Grid) { return }
    Complete-GridEdit $Grid
    try { $Grid.Items.Refresh() } catch { Write-Log ("Could not refresh the grid: {0}" -f $_.Exception.Message) 'WARN' }
}

function Show-FormatMenu {
    # Drop-down for the caret half of a split Export button.
    # $Items: @( @('HTML only', { ... }), @('CSV only', { ... }) )
    param($Button, [object[]]$Items)
    $menu = New-Object System.Windows.Controls.ContextMenu
    foreach ($it in $Items) {
        $mi = New-Object System.Windows.Controls.MenuItem
        $mi.Header = [string]$it[0]
        $mi.Tag = $it[1]
        $mi.Add_Click({ param($s, $e) Invoke-UiAction $s.Tag })
        [void]$menu.Items.Add($mi)
    }
    $menu.PlacementTarget = $Button
    $menu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Top
    $menu.IsOpen = $true
}

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','OK')] [string]$Level = 'INFO')
    $ts   = (Get-Date -Format 'HH:mm:ss')
    $line = '[{0}] {1,-5} {2}' -f $ts, $Level, $Message

    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch { }
    }
    if ($script:LogBox) {
        $script:LogBox.AppendText($line + [Environment]::NewLine)
        $script:LogBox.ScrollToEnd()
        # The log panel is folded away by default - open it on an error.
        if ($Level -eq 'ERROR' -and $script:TglLog) { $script:TglLog.IsChecked = $true }
        Invoke-DoEvents
    }
}

function ConvertTo-HtmlSafe {
    # Escapes a value for use in HTML.
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return ([string]$Text).
        Replace('&','&amp;').
        Replace('<','&lt;').
        Replace('>','&gt;').
        Replace('"','&quot;').
        Replace("'",'&#39;')
}

function Get-ReplacedText {
    # Find -> Replace of literal text, case-insensitive (like AD names). When the
    # replacement is typed in lower case only, each hit keeps the case style of the
    # text it replaces: ceasar -> beasar turns "Ceasar A" into "Beasar A" and
    # "CEASAR" into "BEASAR". A replacement with any capital letter is used as typed.
    param([string]$Text, [string]$Find, [string]$Replace)
    if (-not $Find -or $null -eq $Text) { return $Text }
    $repl  = [string]$Replace
    $adapt = ($repl -cmatch '\p{Ll}') -and -not ($repl -cmatch '\p{Lu}')
    $rx = New-Object System.Text.RegularExpressions.Regex ([regex]::Escape($Find)), ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    return $rx.Replace($Text, [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        if (-not $adapt) { return $repl }
        $hit = $m.Value
        if ($hit -cmatch '\p{Lu}' -and -not ($hit -cmatch '\p{Ll}')) { return $repl.ToUpper() }
        if ([char]::IsUpper($hit[0])) { return $repl.Substring(0, 1).ToUpper() + $repl.Substring(1) }
        return $repl
    })
}

function ConvertTo-SamAccountName {
    # Builds a valid sAMAccountName from a name. Group sAMAccountNames are NOT
    # bound by the 20-character logon limit (that only applies to user accounts),
    # so the full name is kept - up to the 256-character schema maximum - which
    # keeps every group unique instead of colliding on a 20-char truncation.
    param([Parameter(Mandatory)] [string]$Name)
    $sam = $Name -replace '[\\/\[\]:;\|=\+\*\?<>",@]', ''
    $sam = $sam -replace '\s+', ' '
    $sam = $sam.Trim()
    if ($sam.Length -gt 256) { $sam = $sam.Substring(0, 256).Trim() }
    return $sam
}

function Show-Message {
    param(
        [string]$Text,
        [string]$Caption = 'AD ReOrg',
        [ValidateSet('Info','Information','Warning','Error','Question')] [string]$Icon = 'Info',
        [ValidateSet('OK','YesNo','OKCancel')] [string]$Buttons = 'OK'
    )
    $img = switch ($Icon) {
        'Warning'  { [System.Windows.MessageBoxImage]::Warning }
        'Error'    { [System.Windows.MessageBoxImage]::Error }
        'Question' { [System.Windows.MessageBoxImage]::Question }
        default    { [System.Windows.MessageBoxImage]::Information }
    }
    $btn = switch ($Buttons) {
        'YesNo'    { [System.Windows.MessageBoxButton]::YesNo }
        'OKCancel' { [System.Windows.MessageBoxButton]::OKCancel }
        default    { [System.Windows.MessageBoxButton]::OK }
    }
    $answer = [System.Windows.MessageBox]::Show($Text, $Caption, $btn, $img)
    # Only a question returns its answer. A plain OK message returns nothing, so a
    # function that shows a warning and then returns $false returns only $false
    # (otherwise the MessageBoxResult would end up in its output as well).
    if ($Buttons -ne 'OK') { return $answer }
}
