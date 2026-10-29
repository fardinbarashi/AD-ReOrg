# AD Schema Attribute Export / Import

PowerShell scripts that copy the schema attributes production users have, including Exchange and custom schema extensions, into another forest's schema. Typically used to build a test environment where users can hold the same data as in production, even without an Exchange server.

| Script | Run in | What it does |
|---|---|---|
| `Export-SchemaAttributes.ps1` | Source (e.g. production) | Exports the schema definitions of a user's attributes to a JSON file |
| `Import-SchemaAttributes.ps1` | Target (e.g. test) | Creates the missing attributes and classes in the target schema |

The attributes are recreated **as identically as possible**: same OID, syntax, `schemaIDGUID`, link ID, flags and the same class structure (e.g. Exchange auxiliary classes attached to `user`).

The export script only reads the schema; it never changes anything in the source forest.

> ⚠️ **Schema changes are permanent.** Attributes and classes can be deactivated but never deleted. Always run the import with `-WhatIf` first and take a snapshot/backup of the target domain controller.

---

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- The `ActiveDirectory` module (RSAT: *Active Directory Domain Services and Lightweight Directory Services Tools*)
- **Import:** membership in **Schema Admins** in the target forest. Changes are written to the schema master, which the script finds automatically.

CSV files are semicolon-delimited and **UTF-8** encoded (with BOM, so Excel shows characters like å/ä/ö correctly). The JSON file is UTF-8.

---

## Quick start

```powershell
# ---------- In the SOURCE forest ----------
# 1. List every attribute a user can have (see "Listing a user's attributes")
$u = Get-ADUser 'username' -Properties *; $de = [ADSI]"LDAP://$($u.DistinguishedName)"; $de.RefreshCache([string[]]'allowedAttributes'); $r = $de.Properties['allowedAttributes'] | Sort-Object | ForEach-Object { $v = $u.$_; [pscustomobject]@{ Attribute = $_; InUse = -not [string]::IsNullOrEmpty("$v"); Value = ($v -join '; ') } }; $r | Export-Csv 'C:\temp\UserAttributes.csv' -NoTypeInformation -Encoding UTF8 -Delimiter ';'; "Attributes: $(@($r).Count)"

# 2. Export the schema definitions
.\Export-SchemaAttributes.ps1 -CsvPath C:\temp\UserAttributes.csv -Path C:\temp\SchemaAttributes.json

# ---------- Copy SchemaAttributes.json to the TARGET forest ----------
.\Import-SchemaAttributes.ps1 -Path C:\temp\SchemaAttributes.json -WhatIf    # preview
.\Import-SchemaAttributes.ps1 -Path C:\temp\SchemaAttributes.json            # create
```

> **Tip:** PowerShell does not run scripts from the current folder without a path. Always start with `.\`, e.g. `.\Import-SchemaAttributes.ps1`.

---

## Listing a user's attributes

The export needs a list of attribute names. This one-liner lists **every** attribute a user object can have, whether it has a value or not, including Exchange and custom schema extensions. It uses the constructed attribute `allowedAttributes`:

```powershell
$u = Get-ADUser 'username' -Properties *; $de = [ADSI]"LDAP://$($u.DistinguishedName)"; $de.RefreshCache([string[]]'allowedAttributes'); $r = $de.Properties['allowedAttributes'] | Sort-Object | ForEach-Object { $v = $u.$_; [pscustomobject]@{ Attribute = $_; InUse = -not [string]::IsNullOrEmpty("$v"); Value = ($v -join '; ') } }; $r | Export-Csv 'C:\temp\UserAttributes.csv' -NoTypeInformation -Encoding UTF8 -Delimiter ';'; "Attributes: $(@($r).Count)"
```

**CSV columns:** `Attribute`, `InUse`, `Value`

> ⚠️ The `Value` column contains real user data, which may include personal identity numbers. The scripts only use the `Attribute` and `InUse` columns, so store this file accordingly and never commit it.

Alternatively, `Export-SchemaAttributes.ps1 -Identity username` reads the attribute list directly from a user without a CSV file.

---

## Export-SchemaAttributes.ps1

Reads the source schema once and saves, for every attribute:

- the full definition (OID, syntax, `schemaIDGUID`, link ID, ranges, search flags etc.)
- the class that holds the attribute on user objects
- a category: **Base** (Windows base schema), **Exchange** or **Custom**

It also exports:

- the **class hierarchy of `user`**: superclasses and auxiliary classes, such as Exchange's `msExchCustomAttributes` and custom auxiliary classes
- the **forward link** for every backward link attribute (`…BL`). Users only have the backward link, but it cannot be created without its forward link, which usually lives on another object type.

```powershell
.\Export-SchemaAttributes.ps1 -CsvPath C:\temp\UserAttributes.csv -Path C:\temp\SchemaAttributes.json
.\Export-SchemaAttributes.ps1 -CsvPath C:\temp\UserAttributes.csv -OnlyUsed      # only attributes with a value
.\Export-SchemaAttributes.ps1 -Identity username                                  # read the list from a user
```

| Parameter | Description |
|---|---|
| `-CsvPath` | CSV with a column `Attribute` (and optionally `InUse`) |
| `-Identity` | Read the attribute list from this user instead of a CSV file |
| `-OnlyUsed` | Only export attributes where `InUse` is `True` |
| `-Path` | Output JSON file (default `.\SchemaAttributes.json`) |

Example output:

```
Done! Saved: C:\temp\SchemaAttributes.json
  Base          402 attributes
  Custom        24 attributes
  Exchange      451 attributes
  Forward links 41 (required by the backward links)
  Classes       14
```

---

## Import-SchemaAttributes.ps1

The script:

1. Compares the export with the target schema.
2. Creates missing attributes with the **same** definition as the source. Forward links are created before backward links.
3. Creates the classes that hold the attributes in the source (e.g. Exchange auxiliary classes) and attaches them to `user` the same way.
4. Adds attributes to existing classes (`mayContain`) where needed.
5. Refreshes the schema cache between steps and verifies that every attribute is available on user objects.
6. Writes a report: `SchemaImport-Report_<date>.csv`.

It asks for confirmation once before changing anything (unless `-Force` is used). Attributes and classes that already exist are skipped, so it can be run again.

```powershell
.\Import-SchemaAttributes.ps1 -Path C:\temp\SchemaAttributes.json -WhatIf
.\Import-SchemaAttributes.ps1 -Path C:\temp\SchemaAttributes.json
.\Import-SchemaAttributes.ps1 -Path C:\temp\SchemaAttributes.json -Attribute spAlternateNumID,extensionAttribute1
.\Import-SchemaAttributes.ps1 -Path C:\temp\SchemaAttributes.json -ExcludeExchange
```

| Parameter | Description |
|---|---|
| `-Path` | JSON file from `Export-SchemaAttributes.ps1` |
| `-Attribute` | Only create these attributes (lDAPDisplayName) |
| `-ExcludeExchange` | Skip Exchange attributes and classes |
| `-Force` | Do not ask for confirmation |
| `-ReportPath` | Report file (default: next to the JSON file) |
| `-WhatIf` | Show what would be done without changing anything |

### Report

**Columns:** `Name`, `Type`, `Status`, `Message`

| Status | Meaning |
|---|---|
| `Created` | The attribute/class was created |
| `Already existed` | Already in the target schema, nothing done |
| `Attached` | An auxiliary class was attached to its parent class |
| `Not created` | See message, e.g. a Windows base attribute (requires adprep) or a backward link whose forward link is missing |
| `Conflict` | The OID is already used by another attribute in the target schema |
| `Not on user` | The attribute exists but is not available on user objects |
| `Deviation` | The holding class could not be created, so the attribute was added directly to `user` |
| `Error` | AD refused the change, see the warning in the console |

Show everything that needs attention:

```powershell
Import-Csv .\SchemaImport-Report_*.csv -Delimiter ';' -Encoding UTF8 |
    Where-Object Status -notin 'Created', 'Already existed', 'Attached'
```

---

## Notes and limitations

- **Windows base attributes** (category *Base*) are never created or modified. If they are missing, the target forest has an older schema version. Run `adprep /forestprep` from a newer Windows Server instead.
- **Exchange attributes** created this way can store data, but there is no Exchange server to use them. If Exchange will be installed in the target forest later, rebuilding the forest and letting Exchange's own `setup /PrepareSchema` create the schema is the safest option.
- **Backward links** (`…BL`) are filled in automatically by AD from their forward links. They cannot be written directly.
- **Running again:** existing attributes and classes are skipped. If an earlier run reported backward links without forward links, re-run the export with this version and import again.
- **Files from the earlier Swedish version** (`AnvandarFalt.csv` with columns `Falt`/`Anvands`, and `SchemaAttribut.json`) can still be used.
- **Export files are environment data.** Keep CSV and JSON files out of the repository (see `.gitignore`).

---

## Repository layout

```
.
├── Export-SchemaAttributes.ps1
├── Import-SchemaAttributes.ps1
├── README.md
└── .gitignore
```
