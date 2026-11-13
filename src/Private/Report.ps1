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
