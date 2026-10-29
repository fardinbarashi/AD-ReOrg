<#
    Report functions: standalone HTML reports with a per-column filter on every column
    plus a global search, sortable headers and client-side CSV export. Also CSV and
    JSON writers. Used by group creation, snapshot, comparison and copy operations.
    Dot-sourced by AD-ReOrg.ps1 and shares its scope.
#>

function Get-ReportFolder {
    # Report root (config report.folder, default 'files') plus one sub folder per
    # report type, so e.g. copy-members and snapshot-compare files never mix.
    param([string]$Sub = '')
    $rel = 'files'
    if ($script:Config -and $script:Config.report.folder) { $rel = [string]$script:Config.report.folder }
    $folder = if ([System.IO.Path]::IsPathRooted($rel)) { $rel } else { Join-Path $script:ScriptRoot $rel }
    if ($Sub) { $folder = Join-Path $folder $Sub }
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    return $folder
}

function Get-LogoDataUri {
    $logo = Join-Path $script:ScriptRoot 'asset\logo\logo.png'
    if (-not (Test-Path -LiteralPath $logo)) { return $null }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($logo)
        return 'data:image/png;base64,' + [System.Convert]::ToBase64String($bytes)
    } catch { return $null }
}

function Get-StatusBadgeClass {
    # Picks a badge colour for a status value (green / red / amber / grey), else ''.
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    if ($Value -match '(?i)^removed old$')                                 { return 'ok' }
    if ($Value -match '(?i)^differs$')                                    { return 'no' }
    if ($Value -match '(?i)^identical$|^no change$')                     { return 'nt' }
    if ($Value -match '(?i)^renamed$|^reverted$|^ready$|^exists$')       { return 'ok' }
    if ($Value -match '(?i)^attached$')                                   { return 'ok' }
    if ($Value -match '(?i)^whatif$')                                     { return 'wi' }
    if ($Value -match '(?i)conflict|not on user|deviation|not created|no result|not run') { return 'no' }
    if ($Value -match '(?i)error|fail|removed|lost|only in old|skipped') { return 'no' }
    if ($Value -match '(?i)^created$|^ok$|^added$|^replaced$|^appended$|^removed old$|unchanged|in both') { return 'ok' }
    if ($Value -match '(?i)would')                                        { return 'wi' }
    if ($Value -match '(?i)already|only in new')                          { return 'nt' }
    return ''
}

function Test-ReportRowUnchanged {
    # True when a status/action value means "nothing was (or would be) changed",
    # so the report's "Show only changed" box can hide the row.
    param([string]$Value)
    return [bool]($Value -match '(?i)^(unchanged|already|duplicate in run|skipped|not found)')
}

function Get-ReportStyleBlock {
    return @'
<style>
  :root {
    --bg: #f6f7f9; --panel: #ffffff; --border: #dfe3e8; --text: #1f2933;
    --muted: #8a94a0; --accent: #1E76BE; --accent-soft: #e8f0fe;
    --ok-bg: #e3f5e9; --ok-fg: #1b7a3d; --no-bg: #fdeaea; --no-fg: #b4232a;
    --wi-bg: #fff4e0; --wi-fg: #a76b00; --nt-bg: #eef1f5; --nt-fg: #55606e;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #14181d; --panel: #1c2127; --border: #2c333c; --text: #e6e9ed;
      --muted: #8a94a0; --accent: #6ea0ff; --accent-soft: #22304a;
      --ok-bg: #17321f; --ok-fg: #6cd48f; --no-bg: #38191c; --no-fg: #ff8b91;
      --wi-bg: #3a2e14; --wi-fg: #ffca6a; --nt-bg: #232a33; --nt-fg: #aeb6c0;
    }
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--text);
         font: 14px/1.45 "Segoe UI", system-ui, -apple-system, sans-serif; }
  header { padding: 18px 22px 12px; }
  .brand { display: flex; align-items: center; gap: 14px; }
  .logo { width: 42px; height: 42px; object-fit: contain; border-radius: 8px; background: #fff; padding: 4px; }
  .badge-box { width: 46px; height: 46px; border-radius: 8px; background: var(--accent);
               color: #fff; font-weight: 700; font-size: 18px; display: flex;
               align-items: center; justify-content: center; }
  .titles h1 { margin: 0 0 2px; font-size: 20px; font-weight: 600; }
  .meta { color: var(--muted); font-size: 12.5px; }
  .banner { margin: 14px 0 0; padding: 9px 13px; border-radius: 8px; font-size: 13px; }
  .banner.wi { background: var(--wi-bg); color: var(--wi-fg); }
  .banner.ok { background: var(--ok-bg); color: var(--ok-fg); }
  .banner.no { background: var(--no-bg); color: var(--no-fg); }
  .settings { margin: 12px 0 0; padding: 10px 14px 6px; border: 1px solid var(--border); border-radius: 8px;
              background: var(--panel); display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 2px 28px; }
  .settings h2 { grid-column: 1 / -1; margin: 0 0 6px; font-size: 13px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: .04em; }
  .settings .kv { display: flex; gap: 10px; padding: 4px 0; border-bottom: 1px dashed var(--border); font-size: 13px; }
  .settings .k { flex: 0 0 190px; color: var(--muted); }
  .settings .v { flex: 1; word-break: break-word; }
  .toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; padding: 14px 22px; }
  .toolbar input[type=text] { padding: 7px 10px; border: 1px solid var(--border);
             border-radius: 6px; background: var(--panel); color: var(--text); min-width: 260px; }
  .toolbar button { padding: 7px 13px; border: 1px solid var(--border); border-radius: 6px;
             background: var(--panel); color: var(--text); cursor: pointer; }
  .toolbar button:hover { background: var(--accent-soft); border-color: var(--accent); }
  .toolbar label { display: inline-flex; align-items: center; gap: 6px; cursor: pointer; user-select: none; }
  .count { color: var(--muted); font-size: 13px; margin-left: auto; }
  .wrap { margin: 0 22px 30px; background: var(--panel); border: 1px solid var(--border);
          border-radius: 8px; overflow: auto; max-height: 76vh; }
  table { border-collapse: separate; border-spacing: 0; width: 100%; font-size: 13px; }
  th, td { text-align: left; padding: 8px 10px; border-bottom: 1px solid var(--border); vertical-align: top; }
  thead th { position: sticky; background: var(--panel); z-index: 2; }
  tr.head th { top: 0; cursor: pointer; user-select: none; white-space: nowrap;
               font-weight: 600; border-bottom: 1px solid var(--border); }
  tr.head th:hover { color: var(--accent); }
  tr.filters th { top: 35px; padding: 4px 6px 6px; }
  tr.filters input { width: 100%; padding: 4px 6px; font-size: 12px;
           border: 1px solid var(--border); border-radius: 4px; background: var(--bg); color: var(--text); }
  tr.filters input:focus { outline: none; border-color: var(--accent); }
  tbody tr:nth-child(even) { background: rgba(127,127,127,.05); }
  tbody tr:hover { background: var(--accent-soft); }
  .muted { color: var(--muted); }
  .badge { display: inline-block; padding: 1px 8px; border-radius: 10px; font-size: 12px; white-space: nowrap; }
  .badge.ok { background: var(--ok-bg); color: var(--ok-fg); }
  .badge.no { background: var(--no-bg); color: var(--no-fg); }
  .badge.wi { background: var(--wi-bg); color: var(--wi-fg); }
  .badge.nt { background: var(--nt-bg); color: var(--nt-fg); }
  .arrow { font-size: 10px; margin-left: 5px; color: var(--muted); }
  td.hlmatch { background: rgba(255, 214, 0, .18); }
</style>
'@
}

function Get-ReportScriptBlock {
    return @'
<script>
(function () {
  var table = document.getElementById('tbl');
  var tbody = table.tBodies[0];
  var rows = Array.prototype.slice.call(tbody.rows);
  var headCells = Array.prototype.slice.call(table.querySelectorAll('tr.head th'));
  var filterInputs = Array.prototype.slice.call(table.querySelectorAll('input.colfilter'));
  var globalInput = document.getElementById('globalSearch');
  var countEl = document.getElementById('count');
  var onlyChanged = document.getElementById('onlyChanged');

  rows.forEach(function (r) {
    r._text = Array.prototype.map.call(r.cells, function (c) { return c.innerText.toLowerCase(); });
    r._all = r._text.join(' | ');
  });

  function applyFilters() {
    var global = globalInput.value.trim().toLowerCase();
    var active = filterInputs.map(function (inp) { return inp.value.trim().toLowerCase(); });
    var visible = 0;
    var hideUnchanged = onlyChanged && onlyChanged.checked;
    rows.forEach(function (r) {
      var show = true;
      if (hideUnchanged && r.getAttribute('data-unchanged') === '1') { show = false; }
      if (global && r._all.indexOf(global) === -1) { show = false; }
      if (show) {
        for (var i = 0; i < active.length; i++) {
          if (!active[i]) { continue; }
          var terms = active[i].split(/\s+/);
          for (var t = 0; t < terms.length; t++) {
            if (r._text[i].indexOf(terms[t]) === -1) { show = false; break; }
          }
          if (!show) { break; }
        }
      }
      r.style.display = show ? '' : 'none';
      if (show) { visible++; }
      for (var c = 0; c < r.cells.length; c++) {
        if (active[c] && show) { r.cells[c].classList.add('hlmatch'); }
        else { r.cells[c].classList.remove('hlmatch'); }
      }
    });
    countEl.textContent = visible + ' of ' + rows.length + ' rows';
  }

  filterInputs.forEach(function (inp) {
    inp.addEventListener('input', applyFilters);
    inp.addEventListener('click', function (e) { e.stopPropagation(); });
  });
  globalInput.addEventListener('input', applyFilters);
  if (onlyChanged) { onlyChanged.addEventListener('change', applyFilters); }

  document.getElementById('clearBtn').addEventListener('click', function () {
    filterInputs.forEach(function (i) { i.value = ''; });
    globalInput.value = '';
    if (onlyChanged) { onlyChanged.checked = false; }
    applyFilters();
  });

  var sortState = { index: -1, asc: true };
  headCells.forEach(function (th, index) {
    th.addEventListener('click', function () {
      var asc = (sortState.index === index) ? !sortState.asc : true;
      sortState = { index: index, asc: asc };
      headCells.forEach(function (h) { h.querySelector('.arrow').textContent = ''; });
      th.querySelector('.arrow').textContent = asc ? '▲' : '▼';
      var sorted = rows.slice().sort(function (a, b) {
        var x = a._text[index], y = b._text[index];
        var nx = parseFloat(x), ny = parseFloat(y);
        if (!isNaN(nx) && !isNaN(ny) && x !== '' && y !== '') { return asc ? nx - ny : ny - nx; }
        return asc ? x.localeCompare(y) : y.localeCompare(x);
      });
      sorted.forEach(function (r) { tbody.appendChild(r); });
    });
  });

  function visibleRows() {
    return rows.filter(function (r) { return r.style.display !== 'none'; });
  }
  function download(name, content, type) {
    var blob = new Blob([content], { type: type });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  document.getElementById('jsonBtn').addEventListener('click', function () {
    var keys = headCells.map(function (h) { return h.getAttribute('data-key') || h.textContent.replace(/[▲▼]/g, '').trim(); });
    var data = visibleRows().map(function (r) {
      var o = {};
      for (var c = 0; c < keys.length; c++) {
        var t = r.cells[c] ? r.cells[c].innerText.replace(/\s+/g, ' ').trim() : '';
        o[keys[c]] = (t === '-') ? '' : t;
      }
      return o;
    });
    var doc = { application: document.body.getAttribute('data-app'), version: document.body.getAttribute('data-ver'),
                title: document.title, generated: new Date().toISOString(), count: data.length, data: data };
    download('report.json', JSON.stringify(doc, null, 2), 'application/json;charset=utf-8;');
  });

  document.getElementById('csvBtn').addEventListener('click', function () {
    var lines = [];
    lines.push(headCells.map(function (h) {
      return '"' + h.textContent.replace(/[▲▼]/g, '').trim().replace(/"/g, '""') + '"';
    }).join(';'));
    visibleRows().forEach(function (r) {
      lines.push(Array.prototype.map.call(r.cells, function (c) {
        return '"' + c.innerText.replace(/\s+/g, ' ').trim().replace(/"/g, '""') + '"';
      }).join(';'));
    });
    download('report.csv', '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8;');
  });

  applyFilters();
})();
</script>
'@
}

function ConvertTo-DataTable {
    # Builds a System.Data.DataTable (all string columns) from records + column defs.
    # A DataTable's DefaultView binds reliably in a WPF DataGrid, unlike a raw
    # PSCustomObject array (which can render zero rows).
    param([object[]]$Records, [object[]]$Columns)
    $dt = New-Object System.Data.DataTable
    foreach ($c in $Columns) { [void]$dt.Columns.Add([string]$c.Key, [string]) }
    foreach ($r in @($Records)) {
        if ($null -eq $r) { continue }
        $row = $dt.NewRow()
        foreach ($c in $Columns) {
            $v = $r.($c.Key)
            $row[[string]$c.Key] = if ($null -eq $v) { '' } else { [string]$v }
        }
        [void]$dt.Rows.Add($row)
    }
    return ,$dt
}

function New-HtmlReport {
    <#
        Generic HTML report. Every column gets its own filter box (plus a global
        search). $Columns = @( @{Key='';Label=''}, ... ).
        -StatusKey colours that column as a badge and adds a "Show only changed" box
        that hides rows whose status means nothing changed (see Test-ReportRowUnchanged)
        unless -HideChangedBox is given; -Banner (html) shows in the header; -Settings
        (ordered label -> value) is rendered as a "Run settings" panel under the banner.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Title,
        [Parameter(Mandatory)] [object[]]$Columns,
        [Parameter(Mandatory)] $Records,
        [Parameter(Mandatory)] [string]$Path,
        [string]$StatusKey = '',
        [string]$Banner = '',
        [string]$Unit = 'rows',
        $Settings = $null,                 # ordered dictionary label -> value, shown as "Run settings" under the banner
        [switch]$HideChangedBox            # leave out the "Show only changed" box even when -StatusKey is given
    )
    $Records = @($Records)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<thead>')
    [void]$sb.AppendLine('<tr class="head">')
    foreach ($c in $Columns) {
        [void]$sb.AppendLine(('<th data-key="{0}">{1}<span class="arrow"></span></th>' -f $c.Key, (ConvertTo-HtmlSafe $c.Label)))
    }
    [void]$sb.AppendLine('</tr>')
    [void]$sb.AppendLine('<tr class="filters">')
    foreach ($c in $Columns) {
        [void]$sb.AppendLine(('<th><input type="text" class="colfilter" data-col="{0}" placeholder="filter..."></th>' -f $c.Key))
    }
    [void]$sb.AppendLine('</tr>')
    [void]$sb.AppendLine('</thead>')
    [void]$sb.AppendLine('<tbody>')
    foreach ($r in $Records) {
        $rowAttr = ''
        if ($StatusKey -and (Test-ReportRowUnchanged ([string]$r.$StatusKey))) { $rowAttr = ' data-unchanged="1"' }
        [void]$sb.AppendLine("<tr$rowAttr>")
        foreach ($c in $Columns) {
            $val = [string]$r.($c.Key)
            $enc = ConvertTo-HtmlSafe $val
            if ($StatusKey -and $c.Key -eq $StatusKey) {
                $cls = Get-StatusBadgeClass $val
                if ($cls) { $enc = ('<span class="badge {0}">{1}</span>' -f $cls, $enc) }
            }
            if (-not $enc) { $enc = '<span class="muted">-</span>' }
            [void]$sb.AppendLine("<td>$enc</td>")
        }
        [void]$sb.AppendLine('</tr>')
    }
    [void]$sb.AppendLine('</tbody>')

    $logoUri = Get-LogoDataUri
    $badge   = if ($script:App -and $script:App.badge) { [string]$script:App.badge } else { 'AO' }
    $brand   = if ($logoUri) {
        '<img class="logo" src="' + $logoUri + '" alt="logo">'
    } else {
        '<div class="badge-box">' + (ConvertTo-HtmlSafe $badge) + '</div>'
    }
    $appName = if ($script:App) { [string]$script:App.name } else { 'AD ReOrg' }
    $appVer  = if ($script:App) { [string]$script:App.version } else { '' }

    $template = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
__STYLE__
</head>
<body data-app="__APP__" data-ver="__VER__">
<header>
  <div class="brand">__BRAND__
    <div class="titles">
      <h1>__TITLE__</h1>
      <div class="meta">Generated __CREATED__ &middot; __APP__ v__VER__ &middot; __TOTAL__ __UNIT__</div>
    </div>
  </div>
  __BANNER__
__SETTINGS__
</header>

<div class="toolbar">
  <input type="text" id="globalSearch" placeholder="Search all columns...">
  <button id="clearBtn">Clear filters</button>
  <button id="csvBtn">Export visible rows (CSV)</button>
  <button id="jsonBtn">Export visible rows (JSON)</button>
__ONLYCHANGED__  <span class="count" id="count"></span>
</div>

<div class="wrap">
  <table id="tbl">
__TABLE__
  </table>
</div>

__SCRIPT__
</body>
</html>
'@

    $onlyChanged = if ($StatusKey -and -not $HideChangedBox) { '  <label><input type="checkbox" id="onlyChanged"> Show only changed</label>' + "`r`n" } else { '' }
    $settingsHtml = ''
    if ($Settings -and $Settings.Count -gt 0) {
        $sh = New-Object System.Text.StringBuilder
        [void]$sh.AppendLine('  <div class="settings"><h2>Run settings</h2>')
        foreach ($key in @($Settings.Keys)) {
            $v = [string]$Settings[$key]; if (-not $v) { $v = '-' }
            [void]$sh.AppendLine(('    <div class="kv"><span class="k">{0}</span><span class="v">{1}</span></div>' -f (ConvertTo-HtmlSafe ([string]$key)), (ConvertTo-HtmlSafe $v)))
        }
        [void]$sh.AppendLine('  </div>')
        $settingsHtml = $sh.ToString()
    }
    $html = $template.
        Replace('__ONLYCHANGED__', $onlyChanged).
        Replace('__SETTINGS__', $settingsHtml).
        Replace('__STYLE__',   (Get-ReportStyleBlock)).
        Replace('__SCRIPT__',  (Get-ReportScriptBlock)).
        Replace('__TABLE__',   $sb.ToString()).
        Replace('__BRAND__',   $brand).
        Replace('__BANNER__',  $Banner).
        Replace('__TITLE__',   (ConvertTo-HtmlSafe $Title)).
        Replace('__UNIT__',    $Unit).
        Replace('__CREATED__', (Get-Date -Format 'yyyy-MM-dd HH:mm')).
        Replace('__APP__',     (ConvertTo-HtmlSafe $appName)).
        Replace('__VER__',     (ConvertTo-HtmlSafe $appVer)).
        Replace('__TOTAL__',   [string]$Records.Count)

    [System.IO.File]::WriteAllText($Path, $html, (New-Object System.Text.UTF8Encoding($true)))
}

function Export-DataCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object[]]$Columns,
        [Parameter(Mandatory)] $Records,
        [Parameter(Mandatory)] [string]$Path
    )
    $Records = @($Records)
    $keys = $Columns | ForEach-Object { $_.Key }
    if ($Records.Count -gt 0) {
        $Records | Select-Object -Property $keys |
            Export-Csv -Path $Path -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    } else {
        ($keys -join ';') | Set-Content -LiteralPath $Path -Encoding UTF8
    }
}

function Export-DataJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Records,
        [Parameter(Mandatory)] [string]$Path,
        [object]$Meta
    )
    $Records = @($Records)
    $payload = [ordered]@{
        application = if ($script:App) { [string]$script:App.name } else { 'AD ReOrg' }
        version     = if ($script:App) { [string]$script:App.version } else { '' }
        generated   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        count       = $Records.Count
    }
    if ($Meta) { $payload.info = $Meta }
    $payload.data = $Records
    $json = $payload | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function New-GroupCreationReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Records,
        [switch]$WhatIf,
        $Settings = $null
    )
    $Records = @($Records)
    $columns = @(
        @{ Key = 'SourceGroup';    Label = 'Source group' }
        @{ Key = 'NewGroup';       Label = 'New group' }
        @{ Key = 'NewSam';         Label = 'New sAMAccountName' }
        @{ Key = 'Scope';          Label = 'Scope' }
        @{ Key = 'Category';       Label = 'Category' }
        @{ Key = 'Description';    Label = 'Description' }
        @{ Key = 'Notes';          Label = 'Notes (info)' }
        @{ Key = 'TargetOu';       Label = 'Target OU' }
        @{ Key = 'SourceMembers';  Label = 'Members (source)' }
        @{ Key = 'CopiedMembers';  Label = 'Copied' }
        @{ Key = 'Categories';     Label = 'Categories' }
        @{ Key = 'NestedInSource'; Label = 'Nested in source' }
        @{ Key = 'MemberOfCopied'; Label = 'memberOf copied' }
        @{ Key = 'AttributesCopied'; Label = 'Attributes copied' }
        @{ Key = 'Status';         Label = 'Status' }
        @{ Key = 'Timestamp';      Label = 'Timestamp' }
    )

    $folder   = Get-ReportFolder -Sub 'create-groups'
    $baseName = 'report'
    if ($script:Config -and $script:Config.report.fileName) { $baseName = [string]$script:Config.report.fileName }
    $stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
    $htmlPath = Join-Path $folder ('{0}_{1}.html' -f $baseName, $stamp)
    $csvPath  = Join-Path $folder ('{0}_{1}.csv'  -f $baseName, $stamp)

    $banner = if ($WhatIf) {
        '<div class="banner wi">SIMULATION (WhatIf) - nothing was written to Active Directory. Clear WhatIf and run again for a live run.</div>'
    } else {
        '<div class="banner ok">Live run - the groups below were created/updated in Active Directory.</div>'
    }

    New-HtmlReport -Title 'Group creation report' -Columns $columns -Records $Records `
                   -Path $htmlPath -StatusKey 'Status' -Banner $banner -Unit 'rows' -Settings $Settings -HideChangedBox
    Export-DataCsv -Columns $columns -Records $Records -Path $csvPath

    return [pscustomobject]@{ Html = $htmlPath; Csv = $csvPath }
}
