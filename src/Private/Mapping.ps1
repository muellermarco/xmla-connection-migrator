<#
    URL mapping configuration, rename maps, and the per-connection update plan.
#>

function ConvertTo-NormalizedUrl {
    # Normalises a data source URL for matching: trimmed, no trailing slash, lower case.
    param([AllowNull()][string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    return $Url.Trim().TrimEnd('/').ToLowerInvariant()
}

function Assert-HttpUrl {
    param([string]$Url, [string]$Label)
    $uri = $null
    if ([string]::IsNullOrWhiteSpace($Url) -or -not [System.Uri]::TryCreate($Url.Trim(), [System.UriKind]::Absolute, [ref]$uri) -or
        ($uri.Scheme -ne 'http' -and $uri.Scheme -ne 'https')) {
        throw "Mapping config: '$Url' in $Label is not an absolute http or https URL."
    }
}

function Read-MigrationConfig {
    <#
    .SYNOPSIS
        Reads and validates the URL mapping file (config/mapping.json).
    .OUTPUTS
        PSCustomObject with Mappings, OldIndex (normalised old URL -> mapping) and NewIndex.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Mapping config not found: '$Path'. Copy config/mapping.example.json to config/mapping.json and edit it, or pass -Config."
    }
    try {
        $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "Mapping config '$Path' is not valid JSON: $($_.Exception.Message)"
    }
    if ($null -eq $json -or -not ($json.PSObject.Properties.Name -contains 'mappings') -or @($json.mappings).Count -eq 0) {
        throw "Mapping config '$Path' must contain a non-empty 'mappings' array."
    }

    $oldIndex = @{}
    $newIndex = @{}
    $mappings = New-Object System.Collections.Generic.List[psobject]
    $number = 0
    foreach ($entry in @($json.mappings)) {
        $number++
        $names = @($entry.PSObject.Properties.Name)
        if (-not ($names -contains 'old') -or -not ($names -contains 'new')) {
            throw "Mapping config: entry $number must have 'old' and 'new'."
        }
        Assert-HttpUrl -Url $entry.old -Label "entry $number 'old'"
        Assert-HttpUrl -Url $entry.new -Label "entry $number 'new'"
        $mapping = [pscustomobject]@{ Old = $entry.old.Trim(); New = $entry.new.Trim() }
        $oldKey = ConvertTo-NormalizedUrl $mapping.Old
        if ($oldIndex.ContainsKey($oldKey)) { throw "Mapping config: old URL '$($mapping.Old)' appears more than once." }
        $oldIndex[$oldKey] = $mapping
        $newIndex[(ConvertTo-NormalizedUrl $mapping.New)] = $mapping
        $mappings.Add($mapping)
    }
    foreach ($key in $newIndex.Keys) {
        # A new URL that is also an old URL would make a second run change the file again.
        if ($oldIndex.ContainsKey($key)) { throw "Mapping config: URL '$key' is both an old and a new URL. Chained mappings are not allowed." }
    }
    return [pscustomobject]@{ Mappings = $mappings.ToArray(); OldIndex = $oldIndex; NewIndex = $newIndex }
}

function ConvertTo-RenameMap {
    # Converts a -CatalogMap / -CubeMap hashtable into a case-insensitive dictionary.
    param([AllowNull()][hashtable]$Map, [string]$ParameterName)

    $dictionary = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($null -eq $Map) { return , $dictionary }
    foreach ($key in $Map.Keys) {
        $value = $Map[$key]
        if ([string]::IsNullOrWhiteSpace([string]$key) -or [string]::IsNullOrWhiteSpace([string]$value)) {
            throw "-$ParameterName contains an empty name. Use @{ 'OldName' = 'NewName' }."
        }
        if ($dictionary.ContainsKey([string]$key)) { throw "-$ParameterName contains '$key' more than once (names are case-insensitive)." }
        $dictionary.Add(([string]$key).Trim(), ([string]$value).Trim())
    }
    return , $dictionary
}

function Resolve-RenamedName {
    # No [string] type on $Name: PowerShell would turn $null into '' and create a false change.
    param($Name, $Map)
    if ($null -ne $Name -and $Map.ContainsKey([string]$Name)) { return $Map[[string]$Name] }
    return $Name
}

function New-ConnectionUpdate {
    param([string]$Action, [string]$Reason)
    return [pscustomobject]@{
        Action = $Action; Reason = $Reason
        OldDataSource = $null; NewDataSource = $null; OldCatalog = $null; NewCatalog = $null
        OldCube = $null; NewCube = $null; NewConnectionString = $null; NewCommandText = $null
    }
}

function Get-ConnectionUpdate {
    <#
    .SYNOPSIS
        Decides what to change in one connection.
    .OUTPUTS
        PSCustomObject. Action is Change, AlreadyMigrated, NoMatch, ReportOnly or Skip.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ConnectionString,
        [AllowNull()][object]$CommandText,
        [bool]$CommandIsCube,
        [Parameter(Mandatory)][psobject]$Context
    )

    try { $segments = ConvertFrom-ConnectionString -ConnectionString $ConnectionString }
    catch [System.FormatException] { return New-ConnectionUpdate 'Skip' "Connection string cannot be parsed: $($_.Exception.Message)" }

    $provider = [string](Get-ConnectionStringValue -Segments $segments -Key 'Provider')
    if ($provider -match '^\s*Microsoft\.Mashup\.OleDb') { return New-ConnectionUpdate 'ReportOnly' 'Power Query connection: report only' }
    if ($provider -notmatch '^\s*MSOLAP(\.\d+)?\s*$') { return New-ConnectionUpdate 'NoMatch' "Provider is not MSOLAP ('$provider')" }

    $update = New-ConnectionUpdate 'NoMatch' 'Data Source is not in the URL mapping'
    $update.OldDataSource = Get-ConnectionStringValue -Segments $segments -Key 'Data Source'
    $update.OldCatalog = Get-ConnectionStringValue -Segments $segments -Key 'Initial Catalog'
    $cubeInString = Get-ConnectionStringValue -Segments $segments -Key 'Cube'
    $update.OldCube = if ($null -ne $cubeInString) { $cubeInString } elseif ($CommandIsCube) { $CommandText } else { $null }
    $update.NewDataSource = $update.OldDataSource
    $update.NewCatalog = $update.OldCatalog
    $update.NewCube = $update.OldCube

    $urlKey = ConvertTo-NormalizedUrl $update.OldDataSource
    if ($Context.Config.OldIndex.ContainsKey($urlKey)) {
        $update.NewDataSource = $Context.Config.OldIndex[$urlKey].New
    }
    elseif (-not $Context.Config.NewIndex.ContainsKey($urlKey)) {
        return $update
    }

    $update.NewCatalog = Resolve-RenamedName -Name $update.OldCatalog -Map $Context.CatalogMap
    $update.NewCube = Resolve-RenamedName -Name $update.OldCube -Map $Context.CubeMap

    $values = @{}
    if ($update.NewDataSource -cne $update.OldDataSource) { $values['Data Source'] = $update.NewDataSource }
    if ($update.NewCatalog -cne $update.OldCatalog) { $values['Initial Catalog'] = $update.NewCatalog }
    if ($null -ne $cubeInString -and $update.NewCube -cne $cubeInString) { $values['Cube'] = $update.NewCube }
    $update.NewCommandText = $CommandText
    if ($CommandIsCube) { $update.NewCommandText = Resolve-RenamedName -Name $CommandText -Map $Context.CubeMap }

    if ($values.Count -eq 0 -and $update.NewCommandText -ceq $CommandText) {
        $update.Action = 'AlreadyMigrated'
        $update.Reason = 'Data Source already points to the new URL'
        return $update
    }
    $update.Action = 'Change'
    $update.Reason = if ($values.ContainsKey('Data Source')) { '' } else { 'Rename on a connection that already uses the new URL' }
    $update.NewConnectionString = Set-ConnectionStringValue -ConnectionString $ConnectionString -Values $values
    return $update
}
