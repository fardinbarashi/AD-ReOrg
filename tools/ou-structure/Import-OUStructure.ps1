<#
.SYNOPSIS
    Reads a CSV file from Export-OUStructure.ps1 and creates the OU structure in AD.

.DESCRIPTION
    The OUs are created in the correct order (parent before child). OUs that already
    exist are skipped. By default the structure is created directly under the domain
    root. Use -TargetParentDN to place it somewhere else.

    The CSV file is read as UTF-8 (semicolon-delimited).

    Run with -WhatIf first to see what would be created.

.EXAMPLE
    .\Import-OUStructure.ps1 -Path .\OU-Structure_Contoso.csv -WhatIf

.EXAMPLE
    .\Import-OUStructure.ps1 -Path .\OU-Structure_Contoso.csv -TargetParentDN "OU=Regions,DC=test,DC=local"
#>
#Requires -Modules ActiveDirectory
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [string]$TargetParentDN
)

Import-Module ActiveDirectory -ErrorAction Stop

if (-not $TargetParentDN) {
    $TargetParentDN = (Get-ADDomain).DistinguishedName
}

$dnSplit = '(?<!\\),'

# Check that the target location exists
try {
    Get-ADObject -Identity $TargetParentDN -ErrorAction Stop | Out-Null
} catch {
    throw "The target location '$TargetParentDN' does not exist in AD."
}

$rows = Import-Csv -Path $Path -Delimiter ';' -Encoding UTF8 |
        Sort-Object { [int]$_.Depth }, RelativeDN

Write-Host "Creating $(@($rows).Count) OUs under: $TargetParentDN" -ForegroundColor Cyan

$created = 0; $skipped = 0; $failed = 0

foreach ($row in $rows) {
    $targetDN = "$($row.RelativeDN),$TargetParentDN"
    $parentDN = ($targetDN -split $dnSplit, 2)[1]

    # Does the OU already exist?
    $exists = $false
    try {
        Get-ADOrganizationalUnit -Identity $targetDN -ErrorAction Stop | Out-Null
        $exists = $true
    } catch { }

    if ($exists) {
        Write-Host "  Already exists: $targetDN" -ForegroundColor DarkGray
        $skipped++
        continue
    }

    $params = @{
        Name                            = $row.Name
        Path                            = $parentDN
        ProtectedFromAccidentalDeletion = ($row.Protected -eq 'True')
        ErrorAction                     = 'Stop'
    }
    if ($row.Description) { $params.Description = $row.Description }

    if ($PSCmdlet.ShouldProcess($targetDN, 'Create OU')) {
        try {
            New-ADOrganizationalUnit @params
            Write-Host "  Created:        $targetDN" -ForegroundColor Green
            $created++
        } catch {
            Write-Warning "  Failed: $targetDN - $($_.Exception.Message)"
            $failed++
        }
    }
}

Write-Host "`nDone. Created: $created  Already existed: $skipped  Errors: $failed" -ForegroundColor Cyan
