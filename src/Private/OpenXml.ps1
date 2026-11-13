<#
    .xlsx / .xlsm handling. The workbook is a ZIP package. Only the entry
    xl/connections.xml changes. All other entries keep their compressed bytes
    and their order, because ZipArchive in Update mode copies untouched entries
    as they are.
#>

$script:MaxXmlEntryBytes = 50MB

function Read-ZipEntryBytes {
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchiveEntry]$Entry)

    if ($Entry.Length -gt $script:MaxXmlEntryBytes) { throw "ZIP entry '$($Entry.FullName)' is larger than $($script:MaxXmlEntryBytes) bytes." }
    $source = $Entry.Open()
    $buffer = New-Object System.IO.MemoryStream
    try {
        $source.CopyTo($buffer)
        return , $buffer.ToArray()
    }
    finally {
        $source.Dispose()
        $buffer.Dispose()
    }
}

function Test-DataMashupEntry {
    # Power Query stores its queries in a customXml item with a DataMashup root element.
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchiveEntry]$Entry)

    $reader = New-Object System.IO.StreamReader ($Entry.Open(), [System.Text.Encoding]::UTF8, $true)
    try {
        $head = New-Object char[] 1024
        $count = $reader.Read($head, 0, $head.Length)
        return (New-Object string ($head, 0, $count)).Contains('DataMashup')
    }
    finally { $reader.Dispose() }
}

function Get-PackageSignatureNote {
    param([string[]]$EntryNames)

    $notes = @()
    if (@($EntryNames | Where-Object { $_ -like '_xmlsignatures/*' }).Count -gt 0) {
        $notes += 'Digital signature will become invalid'
    }
    if (@($EntryNames | Where-Object { $_ -like 'xl/vbaProjectSignature*' }).Count -gt 0) {
        $notes += 'VBA project signature will become invalid'
    }
    return ($notes -join '; ')
}

function Read-OpenXmlPackage {
    # Reads what the migration needs from the package without changing it.
    param([Parameter(Mandatory)][string]$Path)

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entries = @($zip.Entries)
        $names = [string[]]@($entries | ForEach-Object { $_.FullName })
        $info = [pscustomobject]@{
            EntryNames           = $names
            ConnectionsEntryName = $null
            ConnectionsBytes     = $null
            SignatureNote        = Get-PackageSignatureNote -EntryNames $names
            HasDataMashup        = $false
        }
        foreach ($entry in $entries) {
            if ($entry.FullName -ieq 'xl/connections.xml') {
                $info.ConnectionsEntryName = $entry.FullName
                $info.ConnectionsBytes = Read-ZipEntryBytes -Entry $entry
            }
            elseif ($entry.FullName -imatch '^customXml/item\d+\.xml$' -and -not $info.HasDataMashup) {
                $info.HasDataMashup = Test-DataMashupEntry -Entry $entry
            }
        }
        return $info
    }
    finally { $zip.Dispose() }
}

function Get-OpenXmlConnectionPlan {
    # Plans the update of every <connection> and applies changes to the in-memory document.
    param([Parameter(Mandatory)][System.Xml.XmlDocument]$Document, [Parameter(Mandatory)][psobject]$Context)

    $plans = New-Object System.Collections.Generic.List[psobject]
    # local-name() supports both transitional and strict Open XML namespaces.
    $connections = $Document.SelectNodes("/*[local-name()='connections']/*[local-name()='connection']")
    foreach ($connection in $connections) {
        $name = $connection.GetAttribute('name')
        $dbPr = $connection.SelectSingleNode("*[local-name()='dbPr']")
        if ($null -eq $dbPr -or -not $dbPr.HasAttribute('connection')) {
            $plans.Add([pscustomobject]@{ Name = $name; Update = (New-ConnectionUpdate 'NoMatch' 'Not an OLE DB connection') })
            continue
        }
        $command = if ($dbPr.HasAttribute('command')) { $dbPr.GetAttribute('command') } else { $null }
        $isCube = $dbPr.GetAttribute('commandType') -eq '1'
        $update = Get-ConnectionUpdate -ConnectionString $dbPr.GetAttribute('connection') -CommandText $command -CommandIsCube $isCube -Context $Context
        if ($update.Action -eq 'Change') {
            $dbPr.SetAttribute('connection', $update.NewConnectionString)
            if ($null -ne $command -and $update.NewCommandText -cne $command) { $dbPr.SetAttribute('command', $update.NewCommandText) }
        }
        $plans.Add([pscustomobject]@{ Name = $name; Update = $update })
    }
    return , $plans.ToArray()
}

function Update-ZipEntry {
    # Replaces the content of one entry in place, so the entry keeps its position.
    param([string]$ZipPath, [string]$EntryName, [byte[]]$Bytes)

    $zip = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry($EntryName)
        if ($null -eq $entry) { throw "Entry '$EntryName' not found in '$ZipPath'." }
        $stream = $entry.Open()
        try {
            $stream.SetLength(0)
            $stream.Write($Bytes, 0, $Bytes.Length)
        }
        finally { $stream.Dispose() }
    }
    finally { $zip.Dispose() }
}

function Test-ZipPackage {
    # Reads every entry of the new package and checks that names and order did not change.
    param([string]$Path, [string[]]$ExpectedEntryNames)

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $names = [string[]]@($zip.Entries | ForEach-Object { $_.FullName })
        if (($names -join "`n") -cne ($ExpectedEntryNames -join "`n")) { throw 'ZIP entry list changed during the update.' }
        foreach ($entry in $zip.Entries) {
            $stream = $entry.Open()
            try { $stream.CopyTo([System.IO.Stream]::Null) } finally { $stream.Dispose() }
        }
    }
    finally { $zip.Dispose() }
}

function Write-OpenXmlPackage {
    # Writes the change to a temp copy, validates it, backs up the original, then replaces it.
    param([System.IO.FileInfo]$File, [psobject]$Package, [byte[]]$Bytes, [psobject]$Context)

    $tempPath = New-MigrationTempPath -File $File
    try {
        [System.IO.File]::Copy($File.FullName, $tempPath, $false)
        Update-ZipEntry -ZipPath $tempPath -EntryName $Package.ConnectionsEntryName -Bytes $Bytes
        Test-ZipPackage -Path $tempPath -ExpectedEntryNames $Package.EntryNames
        return Complete-FileReplacement -File $File -TempPath $tempPath -Context $Context
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
    }
}

function Get-OpenXmlSkipReason {
    param([System.IO.FileInfo]$File)

    $reason = Get-FileSkipReason -File $File
    if ($reason) { return $reason }
    switch (Get-FileContainerKind -Path $File.FullName) {
        'Zip' { return $null }
        'Cfb' { return 'Password-protected (encrypted) workbook' }
        default { return 'Not a valid Open XML (ZIP) package' }
    }
}

function Invoke-OpenXmlMigration {
    <#
    .SYNOPSIS
        Migrates the XMLA connections of one .xlsx or .xlsm file. Returns report rows.
    #>
    param([Parameter(Mandatory)][System.IO.FileInfo]$File, [Parameter(Mandatory)][psobject]$Context)

    $skipReason = Get-OpenXmlSkipReason -File $File
    if ($skipReason) { return New-ReportRow -File $File -Status 'skipped' -Reason $skipReason }

    try { $package = Read-OpenXmlPackage -Path $File.FullName }
    catch [System.IO.InvalidDataException] { return New-ReportRow -File $File -Status 'skipped' -Reason "Damaged ZIP package: $($_.Exception.Message)" }

    if ($package.HasDataMashup) {
        New-ReportRow -File $File -ConnectionName '(DataMashup)' -Status 'skipped' -Reason 'Power Query (DataMashup) found: report only'
    }
    if ($null -eq $package.ConnectionsEntryName) {
        return New-ReportRow -File $File -Status 'no match' -Reason 'No external connections (no xl/connections.xml)'
    }

    $xml = Read-XmlBytes -Bytes $package.ConnectionsBytes
    $plans = Get-OpenXmlConnectionPlan -Document $xml.Document -Context $Context
    $changeCount = @($plans | Where-Object { $_.Update.Action -eq 'Change' }).Count
    $decision = Get-ChangeDecision -ChangeCount $changeCount -SignatureNote $package.SignatureNote -Context $Context
    Invoke-DecisionWrite -Decision $decision -FilePath $File.FullName -WriteStep {
        $newBytes = ConvertTo-XmlBytes -Document $xml.Document -WithBom $xml.HasBom
        Write-OpenXmlPackage -File $File -Package $package -Bytes $newBytes -Context $Context
    }
    ConvertTo-MigrationRows -File $File -Plans $plans -Decision $decision
}
