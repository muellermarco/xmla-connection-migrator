<#
    .odc (Office Data Connection) handling.

    An .odc file is an HTML page with an embedded XML island
    (<odc:OfficeDataConnection>). The HTML part is not well-formed XML, so the
    script parses only the island with XmlDocument to read the decoded values.
    It then replaces only the text content of <odc:ConnectionString> (and of
    <odc:CommandText> for a cube rename), XML-escaped, at the exact position in
    the file. All other bytes and the file encoding stay the same.
#>

$script:OdcRootPattern = '<(?<p>(?:[\w-]+:)?)OfficeDataConnection(?:\s[^>]*)?(?<!/)>.*?</\k<p>OfficeDataConnection\s*>'
$script:OdcConnectionPattern = '<(?<p>(?:[\w-]+:)?)Connection(?:\s[^>]*)?(?<!/)>(?<body>.*?)</\k<p>Connection\s*>'

function Get-OdcElementMatch {
    # Finds a child element with text content in a <Connection> body. Returns the regex match or $null.
    param([string]$Body, [string]$LocalName)

    $pattern = '<(?<p>(?:[\w-]+:)?)' + $LocalName + '(?:\s[^>]*)?(?<!/)>(?<value>.*?)</\k<p>' + $LocalName + '\s*>'
    $match = [regex]::Match($Body, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $match.Success) { return $null }
    if ($match.Groups['value'].Value.Contains('<')) { throw "Unsupported markup (CDATA or comment) inside <$LocalName>." }
    return $match
}

function Read-EncodedTextFile {
    # Reads a text file and remembers its encoding and byte order mark, so it can be written back the same way.
    param([Parameter(Mandatory)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $preambleLength = 0
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $encoding = New-Object System.Text.UTF8Encoding $false; $preambleLength = 3
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encoding = New-Object System.Text.UnicodeEncoding ($false, $false); $preambleLength = 2
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $encoding = New-Object System.Text.UnicodeEncoding ($true, $false); $preambleLength = 2
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -ne 0 -and $bytes[1] -eq 0) {
        $encoding = New-Object System.Text.UnicodeEncoding ($false, $false)
    }
    else {
        $encoding = New-Object System.Text.UTF8Encoding ($false, $true)
        try { [void]$encoding.GetString($bytes) }
        catch [System.Text.DecoderFallbackException] {
            # Not UTF-8. Latin-1 maps every byte to one character, so unchanged bytes round-trip exactly.
            $encoding = [System.Text.Encoding]::GetEncoding(28591)
        }
    }
    $preamble = New-Object byte[] $preambleLength
    [System.Array]::Copy($bytes, $preamble, $preambleLength)
    return [pscustomobject]@{
        Text     = $encoding.GetString($bytes, $preambleLength, $bytes.Length - $preambleLength)
        Encoding = $encoding
        Preamble = $preamble
    }
}

function ConvertTo-EncodedBytes {
    param([Parameter(Mandatory)][psobject]$Content, [Parameter(Mandatory)][string]$Text)

    $body = $Content.Encoding.GetBytes($Text)
    if ($Content.Encoding.GetString($body) -cne $Text) { throw "The new text cannot be stored in the file encoding ($($Content.Encoding.WebName))." }
    $bytes = New-Object byte[] ($Content.Preamble.Length + $body.Length)
    [System.Array]::Copy($Content.Preamble, 0, $bytes, 0, $Content.Preamble.Length)
    [System.Array]::Copy($body, 0, $bytes, $Content.Preamble.Length, $body.Length)
    return , $bytes
}

function Get-OdcReplacement {
    # Returns the text replacement (absolute position in the island) for one element value.
    param([System.Text.RegularExpressions.Match]$ConnectionMatch, [string]$LocalName, [string]$NewValue)

    $body = $ConnectionMatch.Groups['body']
    $match = Get-OdcElementMatch -Body $body.Value -LocalName $LocalName
    if ($null -eq $match) { throw "Element <$LocalName> not found in the ODC text." }
    $value = $match.Groups['value']
    return [pscustomobject]@{ Start = $body.Index + $value.Index; Length = $value.Length; Text = ConvertTo-XmlText $NewValue }
}

function Get-OdcConnectionPlan {
    # Plans the update of every <odc:Connection> in the island. Returns Plans and Replacements.
    param([string]$IslandText, [System.Xml.XmlDocument]$Document, [string]$BaseName, [psobject]$Context)

    $nodes = @($Document.SelectNodes("//*[local-name()='Connection']"))
    $connectionMatches = @([regex]::Matches($IslandText, $script:OdcConnectionPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline))
    if ($nodes.Count -ne $connectionMatches.Count) { throw 'Unexpected ODC structure: <Connection> elements could not be matched.' }

    $plans = New-Object System.Collections.Generic.List[psobject]
    $replacements = New-Object System.Collections.Generic.List[psobject]
    for ($index = 0; $index -lt $nodes.Count; $index++) {
        $name = if ($nodes.Count -gt 1) { "$BaseName #$($index + 1)" } else { $BaseName }
        $connectionString = $nodes[$index].SelectSingleNode("*[local-name()='ConnectionString']")
        if ($null -eq $connectionString) {
            $plans.Add([pscustomobject]@{ Name = $name; Update = (New-ConnectionUpdate 'NoMatch' 'No connection string') })
            continue
        }
        $commandType = $nodes[$index].SelectSingleNode("*[local-name()='CommandType']")
        $commandNode = $nodes[$index].SelectSingleNode("*[local-name()='CommandText']")
        $command = if ($null -ne $commandNode) { $commandNode.InnerText } else { $null }
        $isCube = $null -ne $commandType -and $commandType.InnerText.Trim() -ieq 'Cube'
        $update = Get-ConnectionUpdate -ConnectionString $connectionString.InnerText -CommandText $command -CommandIsCube $isCube -Context $Context
        if ($update.Action -eq 'Change') {
            $replacements.Add((Get-OdcReplacement -ConnectionMatch $connectionMatches[$index] -LocalName 'ConnectionString' -NewValue $update.NewConnectionString))
            if ($null -ne $command -and $update.NewCommandText -cne $command) {
                $replacements.Add((Get-OdcReplacement -ConnectionMatch $connectionMatches[$index] -LocalName 'CommandText' -NewValue $update.NewCommandText))
            }
        }
        $plans.Add([pscustomobject]@{ Name = $name; Update = $update })
    }
    return [pscustomobject]@{ Plans = $plans.ToArray(); Replacements = $replacements.ToArray() }
}

function Set-TextReplacement {
    # Applies replacements (positions relative to $Offset) from the end, so earlier positions stay valid.
    param([string]$Text, [int]$Offset, [psobject[]]$Replacements)

    $result = $Text
    foreach ($replacement in ($Replacements | Sort-Object -Property Start -Descending)) {
        $start = $Offset + $replacement.Start
        $result = $result.Substring(0, $start) + $replacement.Text + $result.Substring($start + $replacement.Length)
    }
    return $result
}

function Write-OdcFile {
    param([System.IO.FileInfo]$File, [byte[]]$Bytes, [psobject]$Context)

    $tempPath = New-MigrationTempPath -File $File
    try {
        [System.IO.File]::WriteAllBytes($tempPath, $Bytes)
        return Complete-FileReplacement -File $File -TempPath $tempPath -Context $Context
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
    }
}

function Invoke-OdcMigration {
    <#
    .SYNOPSIS
        Migrates the XMLA connection of one .odc file. Returns report rows.
    #>
    param([Parameter(Mandatory)][System.IO.FileInfo]$File, [Parameter(Mandatory)][psobject]$Context)

    $skipReason = Get-FileSkipReason -File $File
    if ($skipReason) { return New-ReportRow -File $File -Status 'skipped' -Reason $skipReason }

    $content = Read-EncodedTextFile -Path $File.FullName
    $island = [regex]::Match($content.Text, $script:OdcRootPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $island.Success) { return New-ReportRow -File $File -Status 'no match' -Reason 'No <odc:OfficeDataConnection> element found' }

    try { $document = Read-XmlString -Text $island.Value }
    catch [System.Xml.XmlException] { return New-ReportRow -File $File -Status 'skipped' -Reason "ODC connection XML cannot be parsed: $($_.Exception.Message)" }

    $plan = Get-OdcConnectionPlan -IslandText $island.Value -Document $document -BaseName $File.BaseName -Context $Context
    $decision = Get-ChangeDecision -ChangeCount $plan.Replacements.Count -SignatureNote '' -Context $Context
    Invoke-DecisionWrite -Decision $decision -FilePath $File.FullName -WriteStep {
        $newText = Set-TextReplacement -Text $content.Text -Offset $island.Index -Replacements $plan.Replacements
        $newBytes = ConvertTo-EncodedBytes -Content $content -Text $newText
        Write-OdcFile -File $File -Bytes $newBytes -Context $Context
    }
    ConvertTo-MigrationRows -File $File -Plans $plan.Plans -Decision $decision
}
