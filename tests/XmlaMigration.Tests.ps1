<#
    Pester 5 tests for xmla-connection-migrator.
    Run: Invoke-Pester -Path ./tests -Output Detailed
#>

BeforeAll {
    . (Join-Path -Path $PSScriptRoot -ChildPath 'TestHelpers.ps1')
    $RepoRoot = Split-Path -Path $PSScriptRoot -Parent
    $EntryScript = Join-Path -Path $RepoRoot -ChildPath 'src/Invoke-XmlaMigration.ps1'
    Import-Module -Name (Join-Path -Path $RepoRoot -ChildPath 'src/XmlaConnectionMigrator.psm1') -Force

    $ConfigPath = Join-Path -Path $TestDrive -ChildPath 'mapping.json'
    @{ mappings = @(@{ old = $OldUrl; new = $NewUrl }) } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8

    function New-CaseFolder {
        $folder = Join-Path -Path $TestDrive -ChildPath ([guid]::NewGuid().ToString('N'))
        [void](New-Item -ItemType Directory -Path (Join-Path $folder 'data') -Force)
        return $folder
    }

    function Invoke-TestMigration {
        # Runs the entry script on <case>/data and returns the CSV report rows.
        param([string]$Case, [switch]$Apply, [switch]$Force, [switch]$Recurse, [hashtable]$CatalogMap, [hashtable]$CubeMap, [string]$Path)
        if (-not $Path) { $Path = Join-Path $Case 'data' }
        $report = Join-Path $Case ('report-{0}.csv' -f [guid]::NewGuid().ToString('N'))
        $parameters = @{
            Path = $Path; Config = $ConfigPath; ReportPath = $report; BackupPath = (Join-Path $Case 'backup')
            Apply = $Apply; Force = $Force; Recurse = $Recurse
        }
        if ($CatalogMap) { $parameters.CatalogMap = $CatalogMap }
        if ($CubeMap) { $parameters.CubeMap = $CubeMap }
        & $EntryScript @parameters 6> $null
        return @(Import-Csv -LiteralPath $report)
    }

    function Get-DbPrValue {
        param([string]$Path, [string]$Name, [string]$Key)
        $dbPr = Get-XlsxDbPr -Path $Path -Name $Name
        $segments = ConvertFrom-ConnectionString -ConnectionString $dbPr.GetAttribute('connection')
        return Get-ConnectionStringValue -Segments $segments -Key $Key
    }
}

Describe 'Connection string parser' {
    It 'reads keys case-insensitively and decodes quoted values' {
        $segments = ConvertFrom-ConnectionString -ConnectionString 'Provider=MSOLAP.8;Password="a;b""c";Data Source=http://x/'
        Get-ConnectionStringValue -Segments $segments -Key 'password' | Should -BeExactly 'a;b"c'
        Get-ConnectionStringValue -Segments $segments -Key 'DATA SOURCE' | Should -BeExactly 'http://x/'
    }

    It 'replaces only the Data Source value and keeps spacing and other keys' {
        $original = "Provider=MSOLAP.8; Data Source = $OldUrl ;Extended Properties='Locale=1033;x=y'"
        $expected = "Provider=MSOLAP.8; Data Source = $NewUrl ;Extended Properties='Locale=1033;x=y'"
        Set-ConnectionStringValue -ConnectionString $original -Values @{ 'Data Source' = $NewUrl } | Should -BeExactly $expected
    }

    It 'keeps the quote character of a quoted value' {
        $result = Set-ConnectionStringValue -ConnectionString "Data Source=`"$OldUrl`";Provider=MSOLAP" -Values @{ 'data source' = $NewUrl }
        $result | Should -BeExactly "Data Source=`"$NewUrl`";Provider=MSOLAP"
    }

    It 'quotes a new value that contains a semicolon' {
        Set-ConnectionStringValue -ConnectionString 'Initial Catalog=Sales;Provider=MSOLAP' -Values @{ 'Initial Catalog' = 'A;B' } |
            Should -BeExactly 'Initial Catalog="A;B";Provider=MSOLAP'
    }

    It 'throws on an unterminated quoted value' {
        { ConvertFrom-ConnectionString -ConnectionString 'Password="abc' } | Should -Throw
    }
}

Describe 'Mapping config' {
    It 'rejects a URL that is both an old and a new URL' {
        $path = Join-Path $TestDrive 'chained.json'
        '{"mappings":[{"old":"http://a/x","new":"http://b/x"},{"old":"http://b/x","new":"http://c/x"}]}' | Set-Content -LiteralPath $path
        { Read-MigrationConfig -Path $path } | Should -Throw '*Chained*'
    }

    It 'gives a clear message when the config file is missing' {
        { Read-MigrationConfig -Path (Join-Path $TestDrive 'missing.json') } | Should -Throw '*mapping.example.json*'
    }
}

Describe 'Excel .xlsx migration' {
    It 'replaces a matching URL and changes nothing else in the connection string' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/sales.xlsx" -Connections @(@{ Name = 'AtScale Sales'; ConnectionString = (New-AtScaleConnectionString) })

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows.Count | Should -Be 1
        $rows[0].Status | Should -Be 'changed'
        $rows[0].OldDataSource | Should -Be $OldUrl
        $rows[0].NewDataSource | Should -Be $NewUrl
        (Get-XlsxDbPr -Path $file.FullName -Name 'AtScale Sales').GetAttribute('connection') |
            Should -BeExactly (New-AtScaleConnectionString -DataSource $NewUrl)
    }

    It 'matches the old URL case-insensitively and with a trailing slash' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/upper.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString -DataSource 'HTTP://ATSCALE.server.com:10502/XMLA/default/') })

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'changed'
        Get-DbPrValue -Path $file.FullName -Name 'C1' -Key 'Data Source' | Should -BeExactly $NewUrl
    }

    It 'reports no match and leaves the file unchanged for an unknown URL' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/other.xlsx" -Connections @(@{ Name = 'Other'; ConnectionString = (New-AtScaleConnectionString -DataSource 'http://other.server.com/xmla') })
        $before = Get-FileState -Path $file.FullName

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'no match'
        Get-FileState -Path $file.FullName | Should -Be $before
    }

    It 'reports already migrated on a second run and does not touch the file' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/twice.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString) })
        [void](Invoke-TestMigration -Case $case -Apply)
        $afterFirstRun = Get-FileState -Path $file.FullName

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'already migrated'
        Get-FileState -Path $file.FullName | Should -Be $afterFirstRun
    }

    It 'leaves catalog and cube unchanged without rename maps' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/norename.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString -Cube 'Sales Cube') })

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'changed'
        $rows[0].NewCatalog | Should -Be 'Sales'
        $rows[0].NewCube | Should -Be 'Sales Cube'
        Get-DbPrValue -Path $file.FullName -Name 'C1' -Key 'Initial Catalog' | Should -BeExactly 'Sales'
        Get-DbPrValue -Path $file.FullName -Name 'C1' -Key 'Cube' | Should -BeExactly 'Sales Cube'
        (Get-XlsxDbPr -Path $file.FullName -Name 'C1').GetAttribute('command') | Should -BeExactly 'Sales Cube'
    }

    It 'renames catalog and cube with -CatalogMap and -CubeMap' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/rename.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString -Cube 'Sales Cube') })

        $rows = Invoke-TestMigration -Case $case -Apply -CatalogMap @{ 'sales' = 'Sales_Container' } -CubeMap @{ 'Sales Cube' = 'Sales Model' }

        $rows[0].Status | Should -Be 'changed'
        $rows[0].OldCatalog | Should -Be 'Sales'
        $rows[0].NewCatalog | Should -Be 'Sales_Container'
        $rows[0].NewCube | Should -Be 'Sales Model'
        Get-DbPrValue -Path $file.FullName -Name 'C1' -Key 'Data Source' | Should -BeExactly $NewUrl
        Get-DbPrValue -Path $file.FullName -Name 'C1' -Key 'Initial Catalog' | Should -BeExactly 'Sales_Container'
        Get-DbPrValue -Path $file.FullName -Name 'C1' -Key 'Cube' | Should -BeExactly 'Sales Model'
        (Get-XlsxDbPr -Path $file.FullName -Name 'C1').GetAttribute('command') | Should -BeExactly 'Sales Model'
    }

    It 'does not rename catalogs on connections whose data source is not mapped' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/unmapped.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString -DataSource 'http://other/xmla') })
        $before = Get-FileState -Path $file.FullName

        $rows = Invoke-TestMigration -Case $case -Apply -CatalogMap @{ 'Sales' = 'Sales_Container' }

        $rows[0].Status | Should -Be 'no match'
        Get-FileState -Path $file.FullName | Should -Be $before
    }

    It 'keeps quoted values in the connection string' {
        $case = New-CaseFolder
        $extra = "Password=`"p;w`"`"d`";Extended Properties='Locale Identifier=1033;Timeout=60'"
        $file = New-TestXlsx -Path "$case/data/quoted.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString -Extra $extra) })

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'changed'
        (Get-XlsxDbPr -Path $file.FullName -Name 'C1').GetAttribute('connection') |
            Should -BeExactly (New-AtScaleConnectionString -DataSource $NewUrl -Extra $extra)
        Get-DbPrValue -Path $file.FullName -Name 'C1' -Key 'Password' | Should -BeExactly 'p;w"d'
    }

    It 'keeps XML-escaped characters valid' {
        $case = New-CaseFolder
        $extra = 'Application Name="R&D <Finance>"'
        $file = New-TestXlsx -Path "$case/data/escaped.xlsx" -Connections @(@{ Name = 'R&D "Sales"'; ConnectionString = (New-AtScaleConnectionString -Extra $extra) })

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'changed'
        $rows[0].ConnectionName | Should -BeExactly 'R&D "Sales"'
        $rawXml = Get-XlsxConnectionsText -Path $file.FullName
        $rawXml | Should -Match '&amp;'
        $rawXml | Should -Match '&lt;Finance&gt;'
        (Get-XlsxDbPr -Path $file.FullName -Name 'R&D "Sales"').GetAttribute('connection') |
            Should -BeExactly (New-AtScaleConnectionString -DataSource $NewUrl -Extra $extra)
    }

    It 'leaves non-MSOLAP connections untouched' {
        $case = New-CaseFolder
        $sql = "Provider=SQLOLEDB.1;Data Source=$OldUrl;Initial Catalog=Sales"
        $file = New-TestXlsx -Path "$case/data/mixed.xlsx" -Connections @(
            @{ Name = 'SQL'; ConnectionString = $sql; Command = 'SELECT 1'; CommandType = '2' }
            @{ Name = 'Cube'; ConnectionString = (New-AtScaleConnectionString) }
        )

        $rows = Invoke-TestMigration -Case $case -Apply

        ($rows | Where-Object ConnectionName -EQ 'SQL').Status | Should -Be 'no match'
        ($rows | Where-Object ConnectionName -EQ 'Cube').Status | Should -Be 'changed'
        (Get-XlsxDbPr -Path $file.FullName -Name 'SQL').GetAttribute('connection') | Should -BeExactly $sql
    }

    It 'changes nothing in a dry run' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/dry.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString) })
        $before = Get-FileState -Path $file.FullName

        $rows = Invoke-TestMigration -Case $case

        $rows[0].Status | Should -Be 'would change'
        $rows[0].NewDataSource | Should -Be $NewUrl
        Get-FileState -Path $file.FullName | Should -Be $before
        Join-Path $case 'backup' | Should -Not -Exist
    }

    It 'creates a backup with the same relative path before changing a file' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/finance/2024/deep.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString) })
        $originalHash = (Get-FileHash -LiteralPath $file.FullName).Hash

        $rows = Invoke-TestMigration -Case $case -Apply -Recurse

        $rows[0].Status | Should -Be 'changed'
        $backup = Join-Path $case 'backup/finance/2024/deep.xlsx'
        $backup | Should -Exist
        (Get-FileHash -LiteralPath $backup).Hash | Should -Be $originalHash
        (Get-FileHash -LiteralPath $file.FullName).Hash | Should -Not -Be $originalHash
    }

    It 'keeps the ZIP valid and all other entries unchanged and in order' {
        $case = New-CaseFolder
        $extra = [ordered]@{
            'xl/pivotCache/pivotCacheDefinition1.xml' = '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" connectionId="1"><cacheSource type="external" connectionId="1"/></pivotCacheDefinition>'
            'docProps/app.xml'                        = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>'
        }
        $file = New-TestXlsx -Path "$case/data/zip.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString) }) -ExtraEntries $extra
        $before = @(Get-ZipSnapshot -Path $file.FullName)

        [void](Invoke-TestMigration -Case $case -Apply)

        $after = @(Get-ZipSnapshot -Path $file.FullName)
        $after.Count | Should -Be $before.Count
        ($after | ForEach-Object { $_.Split('|')[0] }) -join ',' | Should -BeExactly (($before | ForEach-Object { $_.Split('|')[0] }) -join ',')
        for ($index = 0; $index -lt $before.Count; $index++) {
            if ($before[$index] -like 'xl/connections.xml|*') {
                $after[$index] | Should -Not -Be $before[$index]
            }
            else {
                $after[$index] | Should -BeExactly $before[$index]
            }
        }
    }

    It 'reports Power Query connections only' {
        $case = New-CaseFolder
        $mashup = '<?xml version="1.0" encoding="utf-8"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">AAAA</DataMashup>'
        $file = New-TestXlsx -Path "$case/data/pq.xlsx" -ExtraEntries ([ordered]@{ 'customXml/item1.xml' = $mashup }) -Connections @(
            @{ Name = 'Query - Sales'; ConnectionString = 'Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Sales'; Command = 'SELECT * FROM [Sales]'; CommandType = '2' }
        )
        $before = Get-FileState -Path $file.FullName

        $rows = Invoke-TestMigration -Case $case -Apply

        @($rows | Where-Object Reason -Like '*Power Query*').Count | Should -Be 2
        @($rows | Where-Object Status -NE 'skipped').Count | Should -Be 0
        Get-FileState -Path $file.FullName | Should -Be $before
    }

    It 'skips a signed workbook unless -Force is given' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/signed.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString) }) -ExtraEntries ([ordered]@{ '_xmlsignatures/sig1.xml' = '<Signature/>' })
        $before = Get-FileState -Path $file.FullName

        $rows = Invoke-TestMigration -Case $case -Apply
        $rows[0].Status | Should -Be 'skipped'
        $rows[0].Reason | Should -Match 'signature'
        Get-FileState -Path $file.FullName | Should -Be $before

        $forced = Invoke-TestMigration -Case $case -Apply -Force
        $forced[0].Status | Should -Be 'changed'
        $forced[0].Reason | Should -Match 'signature will become invalid'
    }

    It 'skips read-only files' {
        $case = New-CaseFolder
        $file = New-TestXlsx -Path "$case/data/readonly.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString) })
        $file.IsReadOnly = $true
        try {
            $rows = Invoke-TestMigration -Case $case -Apply
            $rows[0].Status | Should -Be 'skipped'
            $rows[0].Reason | Should -Match 'Read-only'
        }
        finally { $file.IsReadOnly = $false }
    }

    It 'skips files that are open in Excel (lock file present)' {
        $case = New-CaseFolder
        [void](New-TestXlsx -Path "$case/data/open.xlsx" -Connections @(@{ Name = 'C1'; ConnectionString = (New-AtScaleConnectionString) }))
        Set-Content -LiteralPath (Join-Path $case 'data/~$open.xlsx') -Value 'lock'

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows.Count | Should -Be 1
        $rows[0].Status | Should -Be 'skipped'
        $rows[0].Reason | Should -Match 'open in Excel'
    }

    It 'skips password-protected (encrypted) workbooks' {
        $case = New-CaseFolder
        $path = Join-Path $case 'data/protected.xlsx'
        $bytes = New-Object byte[] 512
        [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1) | ForEach-Object -Begin { $i = 0 } -Process { $bytes[$i++] = $_ }
        [System.IO.File]::WriteAllBytes($path, $bytes)

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'skipped'
        $rows[0].Reason | Should -Match 'Password-protected'
    }
}

Describe 'ODC migration' {
    It 'replaces the URL and keeps every other byte, including the UTF-8 byte order mark' {
        $case = New-CaseFolder
        $file = New-TestOdc -Path "$case/data/sales.odc" -ConnectionString (New-AtScaleConnectionString)

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'changed'
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
        ($bytes[0..2] -join ',') | Should -Be '239,187,191'
        $expected = Get-OdcText -ConnectionString (New-AtScaleConnectionString -DataSource $NewUrl)
        [System.IO.File]::ReadAllText($file.FullName) | Should -BeExactly $expected
    }

    It 'renames catalog and cube in the connection string and command text' {
        $case = New-CaseFolder
        $file = New-TestOdc -Path "$case/data/rename.odc" -ConnectionString (New-AtScaleConnectionString)

        $rows = Invoke-TestMigration -Case $case -Apply -CatalogMap @{ 'Sales' = 'Sales_Container' } -CubeMap @{ 'Sales Cube' = 'Sales Model' }

        $rows[0].NewCatalog | Should -Be 'Sales_Container'
        $rows[0].NewCube | Should -Be 'Sales Model'
        $expected = Get-OdcText -ConnectionString (New-AtScaleConnectionString -DataSource $NewUrl -Catalog 'Sales_Container') -CommandText 'Sales Model'
        [System.IO.File]::ReadAllText($file.FullName) | Should -BeExactly $expected
    }

    It 'keeps XML escaping and the UTF-16 encoding' {
        $case = New-CaseFolder
        $connectionString = New-AtScaleConnectionString -Extra 'Application Name="R&D <Finance>"'
        $file = New-TestOdc -Path "$case/data/utf16.odc" -ConnectionString $connectionString -Encoding ([System.Text.Encoding]::Unicode)

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'changed'
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
        ($bytes[0..1] -join ',') | Should -Be '255,254'
        $text = [System.IO.File]::ReadAllText($file.FullName)
        $text | Should -BeExactly (Get-OdcText -ConnectionString (New-AtScaleConnectionString -DataSource $NewUrl -Extra 'Application Name="R&D <Finance>"'))
        $text | Should -Match 'R&amp;D &lt;Finance&gt;'
    }

    It 'reports already migrated for a migrated ODC file' {
        $case = New-CaseFolder
        $file = New-TestOdc -Path "$case/data/done.odc" -ConnectionString (New-AtScaleConnectionString -DataSource $NewUrl)
        $before = Get-FileState -Path $file.FullName

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].Status | Should -Be 'already migrated'
        Get-FileState -Path $file.FullName | Should -Be $before
    }
}

Describe 'Binary workbooks' {
    It 'reports .xlsb as manual / COM needed and skips it' {
        $case = New-CaseFolder
        $path = Join-Path $case 'data/binary.xlsb'
        [System.IO.File]::WriteAllBytes($path, [byte[]](1..64))
        $before = Get-FileState -Path $path

        $rows = Invoke-TestMigration -Case $case -Apply

        $rows[0].FileType | Should -Be 'xlsb'
        $rows[0].Status | Should -Be 'skipped'
        $rows[0].Reason | Should -Match 'manual / COM needed'
        Get-FileState -Path $path | Should -Be $before
    }
}
