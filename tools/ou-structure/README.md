# AD OU Structure Export / Import

PowerShell scripts that copy an Active Directory OU structure, such as one top-level OU and all its sub-OUs, from one domain to another. Typically used to build a test environment that mirrors production.

| Script | Run in | What it does |
|---|---|---|
| `Export-OUStructure.ps1` | Source (e.g. production) | Exports an OU and all sub-OUs to a CSV file |
| `Import-OUStructure.ps1` | Target (e.g. test) | Creates the OU structure from the CSV file |

Only OUs are copied, not users, groups, computers or GPO links. The export script only reads from AD; it never changes anything in the source domain.

---

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- The `ActiveDirectory` module (RSAT: *Active Directory Domain Services and Lightweight Directory Services Tools*)
- Permission to create OUs in the target location (for the import)

The CSV file is semicolon-delimited and **UTF-8** encoded (with BOM, so Excel shows characters like å/ä/ö correctly).

---

## Quick start

```powershell
# In the SOURCE domain
.\Export-OUStructure.ps1 -OUName Contoso -Path C:\temp\OU-Structure_Contoso.csv

# Copy the CSV file to the TARGET domain, then:
.\Import-OUStructure.ps1 -Path C:\temp\OU-Structure_Contoso.csv -WhatIf    # preview
.\Import-OUStructure.ps1 -Path C:\temp\OU-Structure_Contoso.csv            # create
```

> **Tip:** PowerShell does not run scripts from the current folder without a path. Always start with `.\`, e.g. `.\Import-OUStructure.ps1`.

---

## Export-OUStructure.ps1

Finds the OU and exports it together with every sub-OU. Each OU is saved with a path *relative* to the root OU's parent, so the structure can be recreated anywhere, for example under the domain root of another domain or under a different parent OU.

```powershell
.\Export-OUStructure.ps1 -OUName Contoso                   # OU "Contoso" -> .\OU-Structure_Contoso.csv
.\Export-OUStructure.ps1 -OUName Fabrikam -Path C:\temp\fabrikam.csv
.\Export-OUStructure.ps1 -SearchBase "OU=Contoso,OU=Regions,DC=contoso,DC=local"
```

| Parameter | Description |
|---|---|
| `-OUName` | Name of the root OU. Either `-OUName` or `-SearchBase` is required. |
| `-SearchBase` | Full DN of the root OU. Use it if more than one OU has the same name. |
| `-Path` | Output CSV file (default `.\OU-Structure_<root OU name>.csv`) |

If more than one OU has the given name, the script stops and lists them so you can choose with `-SearchBase`.

**CSV columns**

| Column | Example |
|---|---|
| `Name` | `Users` |
| `RelativeDN` | `OU=Users,OU=Sales,OU=Contoso` |
| `Depth` | `3` |
| `Description` | `Sales users` |
| `Protected` | `True` |
| `SourceDN` | `OU=Users,OU=Sales,OU=Contoso,DC=contoso,DC=local` (for reference only) |

---

## Import-OUStructure.ps1

Creates the OUs in the right order, parent before child. OUs that already exist are skipped, so the script can be run again. The description and *Protect object from accidental deletion* setting are copied.

```powershell
.\Import-OUStructure.ps1 -Path .\OU-Structure_Contoso.csv -WhatIf
.\Import-OUStructure.ps1 -Path .\OU-Structure_Contoso.csv
.\Import-OUStructure.ps1 -Path .\OU-Structure_Contoso.csv -TargetParentDN "OU=Regions,DC=test,DC=local"
```

| Parameter | Description |
|---|---|
| `-Path` | CSV file from `Export-OUStructure.ps1` |
| `-TargetParentDN` | Where to create the structure (default: the domain root) |
| `-WhatIf` | Show what would be created without changing anything |

Example output:

```
Creating 208 OUs under: DC=test,DC=contoso,DC=local
  Already exists: OU=Contoso,DC=test,DC=contoso,DC=local
  Created:        OU=Sales,OU=Contoso,DC=test,DC=contoso,DC=local
  Created:        OU=Users,OU=Sales,OU=Contoso,DC=test,DC=contoso,DC=local
  ...
Done. Created: 206  Already existed: 2  Errors: 0
```

---

## Notes

- OU names with special characters (commas, å/ä/ö etc.) are handled.
- OUs created as *protected* cannot be deleted until the protection is removed (`Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $false`).
- Files created by the earlier Swedish version (`OU-Struktur_*.csv`) can still be imported, since the columns are the same.
- The CSV file describes your directory structure, so keep it out of the repository (see `.gitignore`).

---

## Repository layout

```
.
├── Export-OUStructure.ps1
├── Import-OUStructure.ps1
├── README.md
└── .gitignore
```
