<#
.SYNOPSIS
    Exports an OU and all of its sub-OUs to a UTF-8 CSV file.

.DESCRIPTION
    Finds the OU (by name or full DN) and exports the entire structure below it,
    including the root OU itself. Each OU is stored with a relative path (RelativeDN)
    so the structure can be recreated anywhere with Import-OUStructure.ps1.

    The CSV file is semicolon-delimited and UTF-8 encoded.

.EXAMPLE
    .\Export-OUStructure.ps1 -OUName Contoso
    Exports the OU "Contoso" to .\OU-Structure_Contoso.csv

.EXAMPLE
    .\Export-OUStructure.ps1 -SearchBase "OU=Contoso,OU=Regions,DC=contoso,DC=local" -Path C:\Temp\ou.csv
    Use the full DN if more than one OU has the same name.
#>
#Requires -Modules ActiveDirectory
[CmdletBinding()]
param(
    # Name of the root OU
    [string]$OUName,

    # Full DN of the root OU (use when several OUs have the same name)
    [string]$SearchBase,

    # Output file (default: .\OU-Structure_<root OU name>.csv)
    [string]$Path
)

if (-not $OUName -and -not $SearchBase) {
    throw 'Specify the root OU with -OUName or -SearchBase.'
}

Import-Module ActiveDirectory -ErrorAction Stop

# Regex that splits a DN on commas that are NOT escaped (e.g. "OU=Building\, Tech")
$dnSplit = '(?<!\\),'

# Find the root OU if no full DN was given
if (-not $SearchBase) {
    $safeName = $OUName -replace "'", "''"
    $root = @(Get-ADOrganizationalUnit -Filter "Name -eq '$safeName'")
    if ($root.Count -eq 0) {
        throw "No OU named '$OUName' was found."
    }
    if ($root.Count -gt 1) {
        throw "More than one OU is named '$OUName':`n$($root.DistinguishedName -join "`n")`nSpecify -SearchBase with the full DN."
    }
    $SearchBase = $root[0].DistinguishedName
}

# Get the exact DN from AD (handles different casing/spaces in a manually entered -SearchBase)
try {
    $SearchBase = (Get-ADOrganizationalUnit -Identity $SearchBase -ErrorAction Stop).DistinguishedName
} catch {
    throw "No OU with DN '$SearchBase' was found."
}

# Parent of the root OU, e.g. "DC=contoso,DC=local"
$rootParent = ($SearchBase -split $dnSplit, 2)[1]

# Default file name from the root OU's name
if (-not $Path) {
    $rootName = (($SearchBase -split $dnSplit, 2)[0] -replace '^OU=', '') -replace '\\(.)', '$1'
    $Path = ".\OU-Structure_$($rootName -replace '[\\/:*?"<>|]', '_').csv"
}

# UTF-8 with BOM in both Windows PowerShell 5.1 and PowerShell 7 (so Excel shows å/ä/ö correctly)
$csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }

Write-Host "Exporting OU structure from: $SearchBase" -ForegroundColor Cyan

$ous = Get-ADOrganizationalUnit -SearchBase $SearchBase -SearchScope Subtree -Filter * `
        -Properties Description, ProtectedFromAccidentalDeletion

$result = foreach ($ou in $ous) {
    # Remove the parent part so only the structure from the root OU and down remains
    $relative = $ou.DistinguishedName.Substring(0, $ou.DistinguishedName.Length - $rootParent.Length - 1)

    [pscustomobject]@{
        Name        = $ou.Name
        RelativeDN  = $relative
        Depth       = ($relative -split $dnSplit).Count
        Description = $ou.Description
        Protected   = $ou.ProtectedFromAccidentalDeletion
        SourceDN    = $ou.DistinguishedName
    }
}

$result |
    Sort-Object Depth, RelativeDN |
    Export-Csv -Path $Path -NoTypeInformation -Encoding $csvEncoding -Delimiter ';'

Write-Host "Done! $(@($result).Count) OUs exported to $((Resolve-Path $Path).Path)" -ForegroundColor Green
