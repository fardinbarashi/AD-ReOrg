<#
    Runner used by AD ReOrg to start one of the bundled scripts in tools\ in its own
    PowerShell process. The scripts themselves are kept unchanged.

    -Script   full path of the script to run
    -ArgsFile JSON file with the parameters ({"Path": "...", "WhatIf": true, ...});
              true/false become switches, arrays become string arrays.

    Output is written as UTF-8 so AD ReOrg can show names with å/ä/ö correctly.
#>
param(
    [Parameter(Mandatory)] [string]$Script,
    [Parameter(Mandatory)] [string]$ArgsFile
)
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
# PowerShell 7 colours warnings with ANSI codes when output is redirected - keep it plain text.
if ($PSStyle) { try { $PSStyle.OutputRendering = 'PlainText' } catch { } }

$params = @{}
$obj = Get-Content -LiteralPath $ArgsFile -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($p in $obj.PSObject.Properties) {
    $v = $p.Value
    if ($v -is [bool])               { $params[$p.Name] = $v }               # switch: -Name:$true / -Name:$false
    elseif ($v -is [System.Array])   { $params[$p.Name] = [string[]]@($v) }
    elseif ($null -ne $v -and "$v")  { $params[$p.Name] = [string]$v }
}

Set-Location -LiteralPath (Split-Path -Parent $Script)
try {
    & $Script @params
    exit 0
}
catch {
    [Console]::Error.WriteLine('ERROR: ' + $_.Exception.Message)
    exit 1
}
