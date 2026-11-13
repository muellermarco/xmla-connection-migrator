<#
    File checks, backups, and safe replacement of changed files.
#>

function Get-FileSkipReason {
    # Returns a reason when the file must not be touched (read-only, open in Excel, locked), else $null.
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)

    if ($File.IsReadOnly) { return 'Read-only file' }
    $lockFile = Join-Path -Path $File.DirectoryName -ChildPath ('~$' + $File.Name)
    if (Test-Path -LiteralPath $lockFile) { return "File is open in Excel (lock file '~`$$($File.Name)' exists)" }
    try {
        $stream = [System.IO.File]::Open($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Dispose()
    }
    catch [System.UnauthorizedAccessException] { return 'No write access to the file' }
    catch [System.IO.IOException] { return "File is locked by another process: $($_.Exception.Message)" }
    return $null
}

function Get-FileContainerKind {
    # Reads the file header. Zip = Open XML package, Cfb = OLE compound file (encrypted or binary .xls).
    param([Parameter(Mandatory)][string]$Path)

    $header = New-Object byte[] 8
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($header, 0, 8) } finally { $stream.Dispose() }
    if ($read -ge 4 -and $header[0] -eq 0x50 -and $header[1] -eq 0x4B -and $header[2] -eq 0x03 -and $header[3] -eq 0x04) { return 'Zip' }
    $cfb = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    if ($read -lt 8) { return 'Unknown' }
    for ($index = 0; $index -lt 8; $index++) {
        if ($header[$index] -ne $cfb[$index]) { return 'Unknown' }
    }
    return 'Cfb'
}

function Get-RelativeChildPath {
    # Path of $FullPath relative to $Root. PowerShell 5.1 has no Path.GetRelativePath.
    param([string]$Root, [string]$FullPath)

    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $prefix = $rootFull + [System.IO.Path]::DirectorySeparatorChar
    if ($FullPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $FullPath.Substring($prefix.Length)
    }
    return [System.IO.Path]::GetFileName($FullPath)
}

function Backup-MigrationFile {
    # Copies the original file to the backup folder, keeping the relative folder structure.
    param([Parameter(Mandatory)][System.IO.FileInfo]$File, [Parameter(Mandatory)][psobject]$Context)

    $relative = Get-RelativeChildPath -Root $Context.RootPath -FullPath $File.FullName
    $target = Join-Path -Path $Context.BackupPath -ChildPath $relative
    if (Test-Path -LiteralPath $target) {
        # Never overwrite an earlier backup: it may be the only copy of the original.
        $target = '{0}.{1:yyyyMMdd-HHmmss-fff}' -f $target, (Get-Date)
    }
    $targetFolder = Split-Path -Path $target -Parent
    if (-not (Test-Path -LiteralPath $targetFolder)) { [void](New-Item -ItemType Directory -Path $targetFolder -Force) }
    [System.IO.File]::Copy($File.FullName, $target, $false)
    if ((Get-Item -LiteralPath $target).Length -ne $File.Length) { throw "Backup of '$($File.FullName)' is incomplete." }
    return $target
}

function New-MigrationTempPath {
    param([System.IO.FileInfo]$File)
    $name = 'xmla-migrator-{0}{1}' -f [guid]::NewGuid().ToString('N'), $File.Extension
    return Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath $name
}

function Complete-FileReplacement {
    # Backs up the original, then copies the prepared temp file over it. Returns the backup path.
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][string]$TempPath,
        [Parameter(Mandatory)][psobject]$Context
    )
    $backup = Backup-MigrationFile -File $File -Context $Context
    [System.IO.File]::Copy($TempPath, $File.FullName, $true)
    Write-Verbose "Updated '$($File.FullName)'. Backup: '$backup'."
    return $backup
}
