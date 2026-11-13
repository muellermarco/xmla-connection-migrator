<#
    Report rows and the change decision shared by all file handlers.
    Status values: changed, would change, already migrated, no match, skipped.
#>

function Join-Reason {
    param([AllowNull()][string]$First, [AllowNull()][string]$Second)
    $parts = @(@($First, $Second) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    return ($parts -join '; ')
}

function New-ReportRow {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [string]$ConnectionName = '',
        [psobject]$Update,
        [Parameter(Mandatory)][string]$Status,
        [string]$Reason = ''
    )
    $row = [ordered]@{
        FilePath       = $File.FullName
        FileType       = $File.Extension.TrimStart('.').ToLowerInvariant()
        ConnectionName = $ConnectionName
        OldDataSource  = ''
        NewDataSource  = ''
        OldCatalog     = ''
        NewCatalog     = ''
        OldCube        = ''
        NewCube        = ''
        Status         = $Status
        Reason         = $Reason
    }
    if ($null -ne $Update) {
        foreach ($field in 'OldDataSource', 'NewDataSource', 'OldCatalog', 'NewCatalog', 'OldCube', 'NewCube') {
            $row[$field] = [string]$Update.$field
        }
    }
    return [pscustomobject]$row
}

function Get-ChangeDecision {
    # Decides what happens to planned changes in one file: write, report only, or skip.
    param([int]$ChangeCount, [AllowEmptyString()][string]$SignatureNote, [Parameter(Mandatory)][psobject]$Context)

    $decision = [pscustomobject]@{ Write = $false; Status = 'would change'; Reason = '' }
    if ($ChangeCount -eq 0) { return $decision }
    if ($SignatureNote) {
        if (-not $Context.Force) {
            $decision.Status = 'skipped'
            $decision.Reason = "$SignatureNote. Use -Force to change it anyway"
            return $decision
        }
        $decision.Reason = "$SignatureNote (-Force given)"
    }
    $decision.Write = [bool]$Context.Apply
    return $decision
}

function Invoke-DecisionWrite {
    # Runs the write step for a decision and records the outcome. The file stays unchanged on error.
    param([Parameter(Mandatory)][psobject]$Decision, [Parameter(Mandatory)][scriptblock]$WriteStep, [string]$FilePath)

    if (-not $Decision.Write) { return }
    try {
        $backupPath = & $WriteStep
        $Decision.Status = 'changed'
        $Decision.Reason = Join-Reason $Decision.Reason "Backup: $backupPath"
    }
    catch {
        Write-Warning "Could not update '$FilePath': $($_.Exception.Message)"
        $Decision.Status = 'skipped'
        $Decision.Reason = "Write failed, file unchanged: $($_.Exception.Message)"
    }
}

function ConvertTo-MigrationRows {
    # Turns connection plans (Name + Update) and the file decision into report rows.
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][AllowEmptyCollection()][psobject[]]$Plans,
        [Parameter(Mandatory)][psobject]$Decision
    )
    foreach ($plan in $Plans) {
        $update = $plan.Update
        switch ($update.Action) {
            'Change' { $status = $Decision.Status; $reason = Join-Reason $update.Reason $Decision.Reason }
            'AlreadyMigrated' { $status = 'already migrated'; $reason = $update.Reason }
            'NoMatch' { $status = 'no match'; $reason = $update.Reason }
            default { $status = 'skipped'; $reason = $update.Reason }
        }
        New-ReportRow -File $File -ConnectionName $plan.Name -Update $update -Status $status -Reason $reason
    }
}

function ConvertTo-SafeCsvValue {
    # Prefixes values that Excel would read as a formula (CSV injection, CWE-1236).
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -match '^[=+\-@\t\r]') { return "'" + $Value }
    return $Value
}

function Export-MigrationReport {
    <#
    .SYNOPSIS
        Writes the report rows to a CSV file. Returns the full path.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][psobject[]]$Rows, [Parameter(Mandatory)][string]$Path)

    $folder = Split-Path -Path $Path -Parent
    if ($folder -and -not (Test-Path -LiteralPath $folder)) { [void](New-Item -ItemType Directory -Path $folder -Force) }
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $safeRows = foreach ($row in $Rows) {
        $copy = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) { $copy[$property.Name] = ConvertTo-SafeCsvValue -Value ([string]$property.Value) }
        [pscustomobject]$copy
    }
    if (@($safeRows).Count -eq 0) {
        # Export-Csv writes nothing for zero rows. Write the header so the file is still valid.
        $header = '"FilePath","FileType","ConnectionName","OldDataSource","NewDataSource","OldCatalog","NewCatalog","OldCube","NewCube","Status","Reason"'
        Set-Content -LiteralPath $Path -Value $header -Encoding $encoding
    }
    else {
        $safeRows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding $encoding
    }
    return $Path
}

function Write-MigrationSummary {
    # Prints a short summary of the run to the console.
    param([Parameter(Mandatory)][psobject]$Result, [Parameter(Mandatory)][string]$ReportPath)

    $rows = @($Result.Rows)
    $mode = if ($Result.Context.Apply) { 'APPLY (files changed)' } else { 'DRY RUN (no files changed, use -Apply to write)' }
    Write-Host ''
    Write-Host "XMLA connection migration: $mode"
    Write-Host ("Files scanned : {0}" -f $Result.FileCount)
    Write-Host ("Connections   : {0}" -f @($rows | Where-Object { $_.ConnectionName }).Count)
    foreach ($status in 'changed', 'would change', 'already migrated', 'no match', 'skipped') {
        $count = @($rows | Where-Object { $_.Status -eq $status }).Count
        Write-Host ("  {0,-17}: {1}" -f $status, $count)
    }
    if ($Result.Context.Apply -and @($rows | Where-Object { $_.Status -eq 'changed' }).Count -gt 0) {
        Write-Host "Backups       : $($Result.Context.BackupPath)"
    }
    Write-Host "Report        : $ReportPath"
}
