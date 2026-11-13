<#
    Safe XML load and save. DTDs and external resolvers are off (no XXE).
    Whitespace is preserved, and attribute values with line breaks are written
    as character references, so a save does not change values.
#>

function New-SafeXmlReaderSettings {
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    return $settings
}

function Read-XmlBytes {
    # Returns the parsed document and whether the bytes started with a UTF-8 byte order mark.
    param([Parameter(Mandatory)][byte[]]$Bytes)

    $hasBom = $Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF
    $document = New-Object System.Xml.XmlDocument
    $document.PreserveWhitespace = $true
    $document.XmlResolver = $null
    $stream = New-Object System.IO.MemoryStream (, $Bytes)
    try {
        $reader = [System.Xml.XmlReader]::Create($stream, (New-SafeXmlReaderSettings))
        try { $document.Load($reader) } finally { $reader.Dispose() }
    }
    finally { $stream.Dispose() }
    return [pscustomobject]@{ Document = $document; HasBom = $hasBom }
}

function Read-XmlString {
    param([Parameter(Mandatory)][string]$Text)

    $document = New-Object System.Xml.XmlDocument
    $document.PreserveWhitespace = $true
    $document.XmlResolver = $null
    $stringReader = New-Object System.IO.StringReader $Text
    try {
        $reader = [System.Xml.XmlReader]::Create($stringReader, (New-SafeXmlReaderSettings))
        try { $document.Load($reader) } finally { $reader.Dispose() }
    }
    finally { $stringReader.Dispose() }
    return $document
}

function ConvertTo-XmlBytes {
    # Serialises the document as UTF-8, with a byte order mark only if the source had one.
    param([Parameter(Mandatory)][System.Xml.XmlDocument]$Document, [bool]$WithBom)

    # Entitize would write a CR outside the root element as &#xD;, which is not valid there.
    # Line breaks between the declaration and the root element therefore become LF only.
    foreach ($node in @($Document.ChildNodes)) {
        if ($node.NodeType -eq [System.Xml.XmlNodeType]::Whitespace -or $node.NodeType -eq [System.Xml.XmlNodeType]::SignificantWhitespace) {
            $node.Value = $node.Value.Replace("`r", '')
        }
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding $WithBom
    $settings.Indent = $false
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::Entitize
    $stream = New-Object System.IO.MemoryStream
    try {
        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        try { $Document.Save($writer) } finally { $writer.Dispose() }
        return , $stream.ToArray()
    }
    finally { $stream.Dispose() }
}

function ConvertTo-XmlText {
    # Escapes a string for XML element content.
    param([AllowEmptyString()][string]$Value)
    return $Value.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}
