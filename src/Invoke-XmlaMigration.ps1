<#
.SYNOPSIS
    Bulk-updates AtScale XMLA connection strings in Excel workbooks and ODC files.

.DESCRIPTION
    Rewrites the Data Source of MSOLAP connections from the AtScale installer URL
    to the AtScale container URL, as defined in config/mapping.json. Optionally
    renames catalogs and cubes. Dry run by default: nothing changes without -Apply.

    Inputs : -Path (file or folder) and config/mapping.json.
    Outputs: CSV report (-ReportPath), console summary, changed files and backups (with -Apply).

.PARAMETER Path
    A file or a folder to process.
.PARAMETER Recurse
    Include subfolders.
.PARAMETER Config
    URL mapping file. Default: config/mapping.json next to the src folder.
.PARAMETER CatalogMap
    Optional catalog renames, for example @{ 'OldCatalog' = 'NewCatalog' }.
.PARAMETER CubeMap
    Optional cube renames, for example @{ 'OldCube' = 'NewCube' }.
.PARAMETER Apply
    Write the changes. Without it, the script only reports what it would change.
.PARAMETER BackupPath
    Backup folder for changed files. Default: .\xmla-backup-<timestamp>.
.PARAMETER ReportPath
    CSV report path. Default: .\xmla-migration-report-<timestamp>.csv.
.PARAMETER UseExcelCom
    Use Excel COM (Windows with Excel only) for .xlsb and .xls files.
.PARAMETER Force
    Change signed workbooks. The signature becomes invalid.
.PARAMETER PassThru
    Return the report rows as objects.

.EXAMPLE
    .\src\Invoke-XmlaMigration.ps1 -Path \\fileserver\finance -Recurse

.EXAMPLE
    .\src\Invoke-XmlaMigration.ps1 -Path .\pilot -Recurse -Apply -CatalogMap @{ 'Sales' = 'Sales_Container' } -CubeMap @{ 'Sales Cube' = 'Sales' }
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [switch]$Recurse,
    [string]$Config,
    [hashtable]$CatalogMap,
    [hashtable]$CubeMap,
    [switch]$Apply,
    [string]$BackupPath,
    [string]$ReportPath,
    [switch]$UseExcelCom,
    [switch]$Force,
    [switch]$PassThru
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'XmlaConnectionMigrator.psm1') -Force

if (-not $Config) {
    $Config = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config/mapping.json'
}

$options = @{
    Path        = $Path
    Recurse     = $Recurse.IsPresent
    ConfigPath  = $Config
    CatalogMap  = $CatalogMap
    CubeMap     = $CubeMap
    Apply       = $Apply.IsPresent
    BackupPath  = $BackupPath
    ReportPath  = $ReportPath
    UseExcelCom = $UseExcelCom.IsPresent
    Force       = $Force.IsPresent
}

$result = Start-XmlaMigration -Options $options
$result.Rows | Format-Table -Property FilePath, ConnectionName, Status, Reason -AutoSize | Out-Host

if ($PassThru) { $result.Rows }
