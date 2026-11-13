<#
    Test helpers: generate small .xlsx and .odc fixtures at test time and read
    them back. No real customer workbooks are used or committed.
#>
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$OldUrl = 'http://atscale.server.com:10502/xmla/default'
$NewUrl = 'https://atscale.server.com/engine/xmla'

function New-AtScaleConnectionString {
    param(
        [string]$DataSource = $OldUrl,
        [string]$Catalog = 'Sales',
        [string]$Cube,
        [string]$Provider = 'MSOLAP.8',
        [string]$Extra
    )
    $text = "Provider=$Provider;Integrated Security=ClaimsToken;Persist Security Info=True;Initial Catalog=$Catalog;" +
    "Data Source=$DataSource;MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Update Isolation Level=2"
    if ($Cube) { $text += ";Cube=$Cube" }
    if ($Extra) { $text += ";$Extra" }
    return $text
}

function ConvertTo-XmlAttributeText {
    param([string]$Value)
    return [System.Security.SecurityElement]::Escape($Value)
}

function ConvertTo-TestXmlText {
    # Same escaping as the tool uses for element text.
    param([string]$Value)
    return $Value.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function New-ConnectionsXml {
    # Each connection: @{ Name; ConnectionString; Command = 'Sales Cube'; CommandType = '1' }
    param([hashtable[]]$Connections)

    $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n" +
    '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr16" ' +
    'xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16">'
    $id = 1
    foreach ($connection in $Connections) {
        $command = if ($connection.ContainsKey('Command')) { $connection.Command } else { 'Sales Cube' }
        $commandType = if ($connection.ContainsKey('CommandType')) { $connection.CommandType } else { '1' }
        $xml += ('<connection id="{0}" keepAlive="1" name="{1}" type="5" refreshedVersion="8" background="1">' -f $id, (ConvertTo-XmlAttributeText $connection.Name)) +
        ('<dbPr connection="{0}" command="{1}" commandType="{2}"/>' -f (ConvertTo-XmlAttributeText $connection.ConnectionString), (ConvertTo-XmlAttributeText $command), $commandType) +
        '<olapPr sendLocale="1" rowDrillCount="1000"/></connection>'
        $id++
    }
    return $xml + '</connections>'
}

function New-TestXlsx {
    <#
        Creates a minimal workbook package. -Connections builds xl/connections.xml.
        -ExtraEntries (ordered hashtable name -> text) adds more parts at the end.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [hashtable[]]$Connections,
        [System.Collections.Specialized.OrderedDictionary]$ExtraEntries
    )
    $main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $entries = [ordered]@{
        '[Content_Types].xml'        = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/></Types>'
        '_rels/.rels'                = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"$rel/officeDocument`" Target=`"xl/workbook.xml`"/></Relationships>"
        'xl/workbook.xml'            = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook xmlns=`"$main`" xmlns:r=`"$rel`"><sheets><sheet name=`"Sheet1`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
        'xl/_rels/workbook.xml.rels' = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"$rel/worksheet`" Target=`"worksheets/sheet1.xml`"/><Relationship Id=`"rId2`" Type=`"$rel/connections`" Target=`"connections.xml`"/></Relationships>"
        'xl/worksheets/sheet1.xml'   = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"$main`"><sheetData><row r=`"1`"><c r=`"A1`" t=`"inlineStr`"><is><t>Fixture</t></is></c></row></sheetData></worksheet>"
    }
    if ($Connections) { $entries['xl/connections.xml'] = New-ConnectionsXml -Connections $Connections }
    if ($ExtraEntries) { foreach ($key in $ExtraEntries.Keys) { $entries[$key] = $ExtraEntries[$key] } }

    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { [void](New-Item -ItemType Directory -Path $folder -Force) }
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $entries.Keys) {
            $writer = New-Object System.IO.StreamWriter ($zip.CreateEntry($name).Open(), (New-Object System.Text.UTF8Encoding $false))
            try { $writer.Write($entries[$name]) } finally { $writer.Dispose() }
        }
    }
    finally { $zip.Dispose() }
    return Get-Item -LiteralPath $Path
}

function Get-ZipSnapshot {
    # Returns "name|sha256" for every entry, in archive order.
    param([Parameter(Mandatory)][string]$Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in $zip.Entries) {
            $stream = $entry.Open()
            try { $hash = [System.BitConverter]::ToString($sha.ComputeHash($stream)) } finally { $stream.Dispose() }
            '{0}|{1}' -f $entry.FullName, $hash
        }
    }
    finally {
        $zip.Dispose()
        $sha.Dispose()
    }
}

function Get-XlsxConnectionsText {
    param([Parameter(Mandatory)][string]$Path)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $reader = New-Object System.IO.StreamReader ($zip.GetEntry('xl/connections.xml').Open())
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    finally { $zip.Dispose() }
}

function Get-XlsxDbPr {
    # Returns the <dbPr> element of the named connection.
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Name)
    [xml]$document = Get-XlsxConnectionsText -Path $Path
    foreach ($connection in $document.DocumentElement.ChildNodes) {
        if ($connection.LocalName -eq 'connection' -and $connection.GetAttribute('name') -eq $Name) {
            return $connection.SelectSingleNode("*[local-name()='dbPr']")
        }
    }
    throw "Connection '$Name' not found in '$Path'."
}

function Get-OdcText {
    param([string]$ConnectionString, [string]$CommandText = 'Sales Cube')
    $connectionText = ConvertTo-TestXmlText $ConnectionString
    $commandText = ConvertTo-TestXmlText $CommandText
    return @"
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/x-ms-odc; charset=utf-8">
<meta name=ProgId content=ODC.Cube>
<meta name=SourceType content=OLEDB>
<title>AtScale Sales</title>
<xml id=docprops><o:DocumentProperties
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns="http://www.w3.org/TR/REC-html40">
  <o:Name>AtScale Sales</o:Name>
 </o:DocumentProperties>
</xml><xml id=msodc><odc:OfficeDataConnection
  xmlns:odc="urn:schemas-microsoft-com:office:odc"
  xmlns="http://www.w3.org/TR/REC-html40">
  <odc:Connection odc:Type="OLEDB">
   <odc:ConnectionString>$connectionText</odc:ConnectionString>
   <odc:CommandType>Cube</odc:CommandType>
   <odc:CommandText>$commandText</odc:CommandText>
  </odc:Connection>
 </odc:OfficeDataConnection>
</xml>
</head>
<body></body>
</html>
"@
}

function New-TestOdc {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ConnectionString,
        [string]$CommandText = 'Sales Cube',
        [System.Text.Encoding]$Encoding = (New-Object System.Text.UTF8Encoding $true)
    )
    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { [void](New-Item -ItemType Directory -Path $folder -Force) }
    [System.IO.File]::WriteAllText($Path, (Get-OdcText -ConnectionString $ConnectionString -CommandText $CommandText), $Encoding)
    return Get-Item -LiteralPath $Path
}

function Get-FileState {
    # Content hash and last write time, to prove that a file did not change.
    param([Parameter(Mandatory)][string]$Path)
    $item = Get-Item -LiteralPath $Path
    return '{0}|{1}' -f (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash, $item.LastWriteTimeUtc.Ticks
}
