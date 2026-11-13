# xmla-connection-migrator

> This is sample code, provided as is under the MIT License. It is not an AtScale product and AtScale does not support it. Read [DISCLAIMER.md](DISCLAIMER.md) before you use it.

A PowerShell tool that bulk-updates AtScale XMLA connection strings in Excel workbooks and Office Data Connection (.odc) files when AtScale moves from the installer version to the container version.

After the migration, users open their existing workbooks and refresh them against the new AtScale server without editing a connection by hand. The tool changes only the `Data Source` of AtScale (MSOLAP) connections, plus catalog and cube names if you ask for a rename.

## Contents

- [What the tool changes](#what-the-tool-changes)
- [Prerequisites](#prerequisites)
- [Before you run](#before-you-run)
- [Configure the URL mapping](#configure-the-url-mapping)
- [Run the migration](#run-the-migration)
- [Restore from backup](#restore-from-backup)
- [Rename catalogs and cubes](#rename-catalogs-and-cubes)
- [Parameters](#parameters)
- [Report columns and status values](#report-columns-and-status-values)
- [Run the tests](#run-the-tests)
- [Limits](#limits)
- [License](#license)

## What the tool changes

| File type | What happens |
|---|---|
| `.xlsx`, `.xlsm` | Updates the `connection` attribute of each `<dbPr>` in `xl/connections.xml`. Other ZIP entries keep their bytes and their order. |
| `.odc` | Updates the text of `<odc:ConnectionString>`. The file encoding and all other bytes stay the same. |
| `.xlsb`, `.xls` | Reported as "manual / COM needed" and skipped. With `-UseExcelCom` on Windows with Excel, Excel COM updates them. |
| Power Query (DataMashup) | Reported only. |
| Signed workbooks (digital or VBA signature) | Skipped, because the signature becomes invalid. `-Force` changes them anyway. |
| Read-only, locked, open in Excel, password-protected | Skipped and reported. |

A connection changes only when its provider is `MSOLAP` and its `Data Source` matches an old URL in the mapping. The match ignores case and a trailing slash. The tool parses the connection string as key/value pairs (quoted values included) and replaces only the matched values. All other keys stay as they are.

The tool is safe to run more than once. A file that already uses the new URL is reported as "already migrated" and is not touched. A file is written, and its timestamp changes, only when its content changes.

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7.x. The main script uses no external modules.
- Pester 5.x, only to run the tests.
- Write access to the folders with the workbooks, and disk space for the backups.
- Microsoft Excel on one test machine, to test the refresh against the new endpoint.
- For `-UseExcelCom` only: Windows with Excel installed.

## Before you run

Do these two checks first. They decide whether you need the rename parameters.

1. Edit one copy of a workbook by hand so that its connection points to the new endpoint (`Data` > `Queries & Connections` > `Properties` > `Definition`).
2. Refresh the workbook in Excel.
   Result: the PivotTables refresh without errors against the container version.
3. Compare the catalog and cube names in the installer version and the container version.
4. If the names are the same, run the tool without `-CatalogMap` and `-CubeMap`.
5. If the names differ, note each old and new name. Use them with `-CatalogMap` and `-CubeMap` (see [Rename catalogs and cubes](#rename-catalogs-and-cubes)).

## Configure the URL mapping

The URL mapping lives in `config/mapping.json`, not in the code. The file is in `.gitignore`, because it holds customer host names.

1. Copy `config/mapping.example.json` to `config/mapping.json`.
2. Set `old` to the installer URL and `new` to the container URL. Add one entry for each server.

```json
{
  "mappings": [
    {
      "old": "http://atscale.server.com:10502/xmla/default",
      "new": "https://atscale.server.com/engine/xmla"
    }
  ]
}
```

The script rejects a mapping where a new URL is also an old URL, because a second run would change the files again.

## Run the migration

Run all commands from the repository root.

1. Run a dry run on the full share. Dry run is the default and changes no files.

   ```powershell
   .\src\Invoke-XmlaMigration.ps1 -Path \\fileserver\finance -Recurse -ReportPath .\reports\dryrun.csv
   ```

   Result: the console shows a summary per status, and the CSV report lists every connection.

2. Open `reports\dryrun.csv` in Excel.
3. Check that each row with status `would change` shows the correct old and new data source.
4. Check each row with status `skipped` and its reason. Close open workbooks, or plan manual work for `.xlsb` and `.xls` files.
5. Copy a small set of workbooks to a pilot folder.
6. Run the migration with `-Apply` on the pilot folder.

   ```powershell
   .\src\Invoke-XmlaMigration.ps1 -Path D:\pilot -Recurse -Apply -BackupPath D:\xmla-backup-pilot -ReportPath .\reports\pilot.csv
   ```

   Result: rows have status `changed`, and `D:\xmla-backup-pilot` has a copy of each changed file.

7. Open each pilot workbook in Excel and refresh all data.
   Result: every PivotTable and cube formula refreshes against the new server.
8. If a refresh fails, restore the pilot files (see [Restore from backup](#restore-from-backup)) and fix the mapping or rename maps.
9. Tell the users that the change starts, and ask them to close their workbooks.
10. Run the migration with `-Apply` on the full share.

    ```powershell
    .\src\Invoke-XmlaMigration.ps1 -Path \\fileserver\finance -Recurse -Apply -BackupPath D:\xmla-backup-finance -ReportPath .\reports\rollout.csv
    ```

11. Check the rows with status `skipped` in `reports\rollout.csv`.
12. Run step 10 again after the skipped files are closed.
    Result: files done in the first run show `already migrated`, and the remaining files show `changed`.

## Restore from backup

The backup folder keeps the folder structure relative to `-Path`.

1. Find the backup folder in the console summary (`Backups :`) or in the `Reason` column of the report.
2. Ask the users to close the affected workbooks.
3. Copy the backup back over the original folder.

   ```powershell
   Copy-Item -Path D:\xmla-backup-finance\* -Destination \\fileserver\finance -Recurse -Force
   ```

   Result: the files have their original content and connection strings.

4. To restore one file only, copy that file from the same relative path in the backup folder.

## Rename catalogs and cubes

Use the rename parameters only if the names differ between the installer and the container version. Names match case-insensitively. A rename applies only to connections whose data source matches the URL mapping (old or new URL).

```powershell
.\src\Invoke-XmlaMigration.ps1 -Path \\fileserver\finance -Recurse -Apply `
    -BackupPath D:\xmla-backup-finance `
    -CatalogMap @{ 'Sales Catalog' = 'Sales_Container' } `
    -CubeMap @{ 'Sales Cube' = 'Sales Model'; 'Finance Cube' = 'Finance Model' }
```

`-CatalogMap` changes `Initial Catalog` in the connection string. `-CubeMap` changes the `Cube` key if it exists, and the cube name in the command text: the `command` attribute in `xl/connections.xml` (when `commandType="1"`) and `<odc:CommandText>` in ODC files (when `<odc:CommandType>` is `Cube`). Excel uses the command text to select the cube, so a cube rename without it would break the refresh.

## Parameters

| Parameter | Description |
|---|---|
| `-Path` | File or folder to process. Required. |
| `-Recurse` | Include subfolders. |
| `-Config` | Mapping file. Default: `config/mapping.json`. |
| `-CatalogMap` | Optional catalog renames, as a hashtable. |
| `-CubeMap` | Optional cube renames, as a hashtable. |
| `-Apply` | Write the changes. Without it, the run is a dry run. |
| `-BackupPath` | Backup folder. Default: `.\xmla-backup-<timestamp>`. Files in this folder are never scanned. |
| `-ReportPath` | CSV report. Default: `.\xmla-migration-report-<timestamp>.csv`. |
| `-UseExcelCom` | Use Excel COM for `.xlsb` and `.xls`. Off by default. Windows with Excel only. |
| `-Force` | Change signed workbooks. The signature becomes invalid. |
| `-PassThru` | Also return the report rows as objects. |

## Report columns and status values

Columns: `FilePath`, `FileType`, `ConnectionName`, `OldDataSource`, `NewDataSource`, `OldCatalog`, `NewCatalog`, `OldCube`, `NewCube`, `Status`, `Reason`.

There is one row per connection. A skipped file without readable connections gets one row with an empty `ConnectionName`.

| Status | Meaning |
|---|---|
| `changed` | The connection was updated (`-Apply`). `Reason` shows the backup path. |
| `would change` | Dry run: the connection would be updated. |
| `already migrated` | The connection already uses the new URL and no rename applies. |
| `no match` | Not an MSOLAP connection, or the data source is not in the mapping. |
| `skipped` | Not changed. `Reason` says why (binary format, signature, locked, Power Query, error). |

## Run the tests

The tests generate small `.xlsx` and `.odc` fixtures at run time. No customer workbooks are needed.

1. Install Pester 5 if it is not installed.

   ```powershell
   Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser -Force -SkipPublisherCheck
   ```

2. Run the tests from the repository root.

   ```powershell
   Invoke-Pester -Path .\tests -Output Detailed
   ```

   Result: all tests pass.

## Limits

- `.xlsb` and `.xls` need Excel COM or a manual change. The COM path is not covered by the automated tests.
- The `<meta name=Catalog>` and `<meta name=Table>` tags in the HTML part of an ODC file are display hints and are not changed. Excel reads the connection from the `<odc:OfficeDataConnection>` XML.
- `xl/connections.xml` is parsed and saved again as XML. The values stay the same, but formatting details can change, for example `<a/>` becomes `<a />`, and a line break before the root element becomes LF.
- Connection string keys that contain an escaped `=` (`==`) are not supported. AtScale connection strings do not use them.
- Workbooks that connect through an `.odc` file keep their own copy of the connection string. Run the tool on both the workbooks and the `.odc` files.

## License

MIT License, see [LICENSE](LICENSE). Read [DISCLAIMER.md](DISCLAIMER.md): this is sample code without support or warranty.
