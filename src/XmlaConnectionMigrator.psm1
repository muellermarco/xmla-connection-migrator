<#
    XmlaConnectionMigrator module

    Purpose : Functions that rewrite AtScale XMLA connection strings in Excel
              Open XML workbooks (.xlsx/.xlsm) and Office Data Connection
              (.odc) files. Entry point is src/Invoke-XmlaMigration.ps1.
    Inputs  : See Start-XmlaMigration.
    Outputs : Report rows (PSCustomObject) and, with -Apply, changed files.
    Runs on : Windows PowerShell 5.1 and PowerShell 7+. No external modules.
#>
Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$privateFolder = Join-Path -Path $PSScriptRoot -ChildPath 'Private'
foreach ($scriptFile in (Get-ChildItem -LiteralPath $privateFolder -Filter '*.ps1' | Sort-Object -Property Name)) {
    . $scriptFile.FullName
}

Export-ModuleMember -Function @(
    'ConvertFrom-ConnectionString'
    'Get-ConnectionStringValue'
    'Set-ConnectionStringValue'
    'Read-MigrationConfig'
    'Get-ConnectionUpdate'
    'Start-XmlaMigration'
)
