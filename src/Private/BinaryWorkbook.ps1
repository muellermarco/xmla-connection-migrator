<#
    .xlsb / .xls handling. These binary formats cannot be edited as XML.
    By default they are reported as "manual / COM needed" and skipped.
    With -UseExcelCom (Windows with Excel only) the script uses Excel COM:
    macros are disabled, links are not updated, and the workbook is saved only
    when a connection changed.
#>

function Test-ExcelComAvailable {
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) { return $false }
    return Test-Path -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Excel.Application'
}

function Get-ExcelComConnectionPlan {
    # Plans updates for all OLE DB connections in an open workbook. Does not change anything.
    param([Parameter(Mandatory)]$Workbook, [Parameter(Mandatory)][psobject]$Context)

    $xlConnectionTypeOLEDB = 1
    $xlCmdCube = 1
    $plans = New-Object System.Collections.Generic.List[psobject]
    foreach ($connection in $Workbook.Connections) {
        $name = [string]$connection.Name
        if ($connection.Type -ne $xlConnectionTypeOLEDB) {
            $plans.Add([pscustomobject]@{ Name = $name; Update = (New-ConnectionUpdate 'NoMatch' 'Not an OLE DB connection'); Com = $null })
            continue
        }
        $oleDb = $connection.OLEDBConnection
        $connectionString = ([string]$oleDb.Connection) -replace '^OLEDB;', ''
        $command = [string]$oleDb.CommandText
        $update = Get-ConnectionUpdate -ConnectionString $connectionString -CommandText $command -CommandIsCube ($oleDb.CommandType -eq $xlCmdCube) -Context $Context
        $plans.Add([pscustomobject]@{ Name = $name; Update = $update; Com = $oleDb; Command = $command })
    }
    return , $plans.ToArray()
}

function Get-ExcelComSignatureNote {
    param([Parameter(Mandatory)]$Workbook)
    $notes = @()
    if ($Workbook.Signatures.Count -gt 0) { $notes += 'Digital signature will become invalid' }
    if ($Workbook.VBASigned) { $notes += 'VBA project signature will become invalid' }
    return ($notes -join '; ')
}

function Save-ExcelComChange {
    # Backs up the file, writes the planned connection changes, and saves the workbook. Returns the backup path.
    param([System.IO.FileInfo]$File, $Workbook, [psobject[]]$Plans, [psobject]$Context)

    $backup = Backup-MigrationFile -File $File -Context $Context
    foreach ($plan in ($Plans | Where-Object { $_.Update.Action -eq 'Change' })) {
        $plan.Com.Connection = 'OLEDB;' + $plan.Update.NewConnectionString
        if ($plan.Update.NewCommandText -cne $plan.Command) { $plan.Com.CommandText = $plan.Update.NewCommandText }
    }
    $Workbook.Save()
    return $backup
}

function Close-ExcelCom {
    param($Excel, $Workbook)
    if ($null -ne $Workbook) {
        $Workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)
    }
    if ($null -ne $Excel) {
        $Excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Invoke-ExcelComMigration {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File, [Parameter(Mandatory)][psobject]$Context)

    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: no macros run
        try {
            # A dummy password makes Excel fail on protected files instead of showing a prompt.
            $workbook = $excel.Workbooks.Open($File.FullName, 0, (-not $Context.Apply), [Type]::Missing, 'xmla-migrator-no-password')
        }
        catch [System.Runtime.InteropServices.COMException] {
            return New-ReportRow -File $File -Status 'skipped' -Reason "Excel cannot open the file (password-protected?): $($_.Exception.Message)"
        }
        $plans = Get-ExcelComConnectionPlan -Workbook $workbook -Context $Context
        $changeCount = @($plans | Where-Object { $_.Update.Action -eq 'Change' }).Count
        $decision = Get-ChangeDecision -ChangeCount $changeCount -SignatureNote (Get-ExcelComSignatureNote -Workbook $workbook) -Context $Context
        Invoke-DecisionWrite -Decision $decision -FilePath $File.FullName -WriteStep {
            Save-ExcelComChange -File $File -Workbook $workbook -Plans $plans -Context $Context
        }
        ConvertTo-MigrationRows -File $File -Plans $plans -Decision $decision
    }
    finally {
        Close-ExcelCom -Excel $excel -Workbook $workbook
    }
}

function Invoke-BinaryWorkbookMigration {
    <#
    .SYNOPSIS
        Reports .xlsb/.xls files, or migrates them with Excel COM when -UseExcelCom is given.
    #>
    param([Parameter(Mandatory)][System.IO.FileInfo]$File, [Parameter(Mandatory)][psobject]$Context)

    if (-not $Context.UseExcelCom) {
        return New-ReportRow -File $File -Status 'skipped' -Reason 'Binary format: manual / COM needed (use -UseExcelCom on Windows with Excel)'
    }
    if (-not (Test-ExcelComAvailable)) {
        return New-ReportRow -File $File -Status 'skipped' -Reason 'Binary format: -UseExcelCom needs Windows with Excel installed'
    }
    $skipReason = Get-FileSkipReason -File $File
    if ($skipReason) { return New-ReportRow -File $File -Status 'skipped' -Reason $skipReason }
    return Invoke-ExcelComMigration -File $File -Context $Context
}
