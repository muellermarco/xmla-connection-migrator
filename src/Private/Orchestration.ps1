<#
    Run orchestration: build the context, find files, dispatch by file type.
#>

$script:SupportedExtensions = @('.xlsx', '.xlsm', '.odc', '.xlsb', '.xls')

function New-MigrationContext {
    param([Parameter(Mandatory)][hashtable]$Options)

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupPath = $Options.BackupPath
    if (-not $backupPath) { $backupPath = Join-Path -Path (Get-Location).ProviderPath -ChildPath "xmla-backup-$timestamp" }
    return [pscustomobject]@{
        RootPath    = $Options.RootPath
        Config      = Read-MigrationConfig -Path (Resolve-FullPath -Path $Options.ConfigPath)
        CatalogMap  = ConvertTo-RenameMap -Map $Options.CatalogMap -ParameterName 'CatalogMap'
        CubeMap     = ConvertTo-RenameMap -Map $Options.CubeMap -ParameterName 'CubeMap'
        Apply       = [bool]$Options.Apply
        Force       = [bool]$Options.Force
        UseExcelCom = [bool]$Options.UseExcelCom
        BackupPath  = Resolve-FullPath -Path $backupPath
        Timestamp   = $timestamp
    }
}

function Resolve-FullPath {
    # Resolves a relative path against the PowerShell location, not the process directory.
    param([Parameter(Mandatory)][string]$Path)
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function Test-IsUnderPath {
    param([string]$Path, [string]$Folder)
    $prefix = $Folder.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    return $Path.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-MigrationCandidateFile {
    # Lists supported files. Skips Excel lock files (~$*) and anything inside the backup folder.
    param([Parameter(Mandatory)][System.IO.FileSystemInfo]$Item, [bool]$Recurse, [string]$BackupPath)

    if ($Item -is [System.IO.FileInfo]) {
        if ($script:SupportedExtensions -notcontains $Item.Extension.ToLowerInvariant()) {
            throw "Unsupported file type '$($Item.Extension)'. Supported: $($script:SupportedExtensions -join ', ')."
        }
        return $Item
    }
    Get-ChildItem -LiteralPath $Item.FullName -File -Recurse:$Recurse |
        Where-Object {
            ($script:SupportedExtensions -contains $_.Extension.ToLowerInvariant()) -and
            -not $_.Name.StartsWith('~$') -and
            -not (Test-IsUnderPath -Path $_.FullName -Folder $BackupPath)
        } |
        Sort-Object -Property FullName
}

function Invoke-FileMigration {
    # Dispatches one file to its handler. An unexpected error skips the file, not the run.
    param([Parameter(Mandatory)][System.IO.FileInfo]$File, [Parameter(Mandatory)][psobject]$Context)

    Write-Verbose "Processing '$($File.FullName)'"
    try {
        switch ($File.Extension.ToLowerInvariant()) {
            { $_ -in '.xlsx', '.xlsm' } { Invoke-OpenXmlMigration -File $File -Context $Context }
            '.odc' { Invoke-OdcMigration -File $File -Context $Context }
            { $_ -in '.xlsb', '.xls' } { Invoke-BinaryWorkbookMigration -File $File -Context $Context }
        }
    }
    catch {
        Write-Warning "Skipped '$($File.FullName)': $($_.Exception.Message)"
        New-ReportRow -File $File -Status 'skipped' -Reason "Error: $($_.Exception.Message)"
    }
}

function Start-XmlaMigration {
    <#
    .SYNOPSIS
        Runs the migration for a file or folder and returns the report rows.
    .PARAMETER Options
        Hashtable with Path, Recurse, ConfigPath, CatalogMap, CubeMap, Apply,
        BackupPath, ReportPath, UseExcelCom, Force.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Options)

    $item = Get-Item -LiteralPath $Options.Path -ErrorAction Stop
    $Options.RootPath = if ($item -is [System.IO.FileInfo]) { $item.DirectoryName } else { $item.FullName }
    $context = New-MigrationContext -Options $Options

    $rows = New-Object System.Collections.Generic.List[psobject]
    $files = @(Get-MigrationCandidateFile -Item $item -Recurse ([bool]$Options.Recurse) -BackupPath $context.BackupPath)
    foreach ($file in $files) {
        foreach ($row in @(Invoke-FileMigration -File $file -Context $context)) { $rows.Add($row) }
    }
    return [pscustomobject]@{ Context = $context; FileCount = $files.Count; Rows = $rows.ToArray() }
}
