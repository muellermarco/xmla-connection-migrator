<#
    OLE DB connection string parsing.

    The parser keeps the position of every value in the original string, so a
    caller can replace one value and leave every other character unchanged.
    Rules: pairs are separated by ';', key and value by the first '='. A value
    may be quoted with " or '. Inside quotes, a doubled quote is a literal quote
    and ';' is part of the value.
#>

function Read-ConnectionStringValue {
    # Reads one value that starts at $Start. Returns value text, raw span, quote char and next index.
    param([string]$Text, [int]$Start)

    $length = $Text.Length
    $first = if ($Start -lt $length) { $Text[$Start] } else { [char]0 }

    if ($first -ne '"' -and $first -ne "'") {
        $end = $Text.IndexOf(';', $Start)
        if ($end -lt 0) { $end = $length }
        $raw = $Text.Substring($Start, $end - $Start).TrimEnd()
        return [pscustomobject]@{ Value = $raw; ValueLength = $raw.Length; QuoteChar = $null; Next = $end + 1 }
    }

    $builder = New-Object System.Text.StringBuilder
    $index = $Start + 1
    while ($true) {
        if ($index -ge $length) {
            throw [System.FormatException]::new("Unterminated quoted value at position $Start.")
        }
        $char = $Text[$index]
        if ($char -eq $first) {
            if (($index + 1) -lt $length -and $Text[$index + 1] -eq $first) {
                [void]$builder.Append($first)
                $index += 2
                continue
            }
            break
        }
        [void]$builder.Append($char)
        $index++
    }

    $next = $Text.IndexOf(';', $index + 1)
    if ($next -lt 0) { $next = $length }
    return [pscustomobject]@{
        Value       = $builder.ToString()
        ValueLength = $index + 1 - $Start
        QuoteChar   = [string]$first
        Next        = $next + 1
    }
}

function ConvertFrom-ConnectionString {
    <#
    .SYNOPSIS
        Parses an OLE DB connection string into ordered key/value segments.
    .OUTPUTS
        PSCustomObject[] with Key, Value, ValueStart, ValueLength, QuoteChar.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$ConnectionString)

    $text = $ConnectionString
    $segments = New-Object System.Collections.Generic.List[psobject]
    $position = 0
    while ($position -lt $text.Length) {
        $equals = $text.IndexOf('=', $position)
        $semicolon = $text.IndexOf(';', $position)
        if ($equals -lt 0 -or ($semicolon -ge 0 -and $semicolon -lt $equals)) {
            # Segment without '=' (empty or malformed). Keep it untouched and move on.
            if ($semicolon -lt 0) { break }
            $position = $semicolon + 1
            continue
        }

        $key = $text.Substring($position, $equals - $position).Trim()
        $valueStart = $equals + 1
        while ($valueStart -lt $text.Length -and [char]::IsWhiteSpace($text[$valueStart])) { $valueStart++ }

        $value = Read-ConnectionStringValue -Text $text -Start $valueStart
        $segments.Add([pscustomobject]@{
                Key         = $key
                Value       = $value.Value
                ValueStart  = $valueStart
                ValueLength = $value.ValueLength
                QuoteChar   = $value.QuoteChar
            })
        $position = $value.Next
    }
    return , $segments.ToArray()
}

function Get-ConnectionStringValue {
    <#
    .SYNOPSIS
        Returns the value of the first segment whose key matches (case-insensitive), or $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][psobject[]]$Segments,
        [Parameter(Mandatory)][string]$Key
    )
    foreach ($segment in $Segments) {
        if ([string]::Equals($segment.Key, $Key, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $segment.Value
        }
    }
    return $null
}

function Format-ConnectionStringValue {
    param([string]$Value, [string]$QuoteChar)

    if (-not $QuoteChar) {
        $needsQuotes = $Value.Contains(';') -or $Value.StartsWith('"') -or $Value.StartsWith("'") -or
        ($Value.Trim() -ne $Value)
        if (-not $needsQuotes) { return $Value }
        $QuoteChar = '"'
    }
    return $QuoteChar + $Value.Replace($QuoteChar, $QuoteChar + $QuoteChar) + $QuoteChar
}

function Set-ConnectionStringValue {
    <#
    .SYNOPSIS
        Replaces the values of existing keys and returns the new connection string.
    .DESCRIPTION
        Only the value spans of the given keys change. Keys that are not in the
        string are not added. Keys are matched case-insensitively. A value that
        was quoted keeps its quote character.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ConnectionString,
        [Parameter(Mandatory)][hashtable]$Values
    )

    $segments = ConvertFrom-ConnectionString -ConnectionString $ConnectionString
    $replacements = New-Object System.Collections.Generic.List[psobject]
    foreach ($segment in $segments) {
        foreach ($key in $Values.Keys) {
            if ([string]::Equals($segment.Key, [string]$key, [System.StringComparison]::OrdinalIgnoreCase)) {
                $replacements.Add([pscustomobject]@{
                        Start  = $segment.ValueStart
                        Length = $segment.ValueLength
                        Text   = Format-ConnectionStringValue -Value ([string]$Values[$key]) -QuoteChar $segment.QuoteChar
                    })
            }
        }
    }

    $result = $ConnectionString
    foreach ($replacement in ($replacements | Sort-Object -Property Start -Descending)) {
        $result = $result.Substring(0, $replacement.Start) + $replacement.Text +
        $result.Substring($replacement.Start + $replacement.Length)
    }
    return $result
}
