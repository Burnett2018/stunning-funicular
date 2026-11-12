<#
.SYNOPSIS
    Joins the SAP contract exports and writes the supplier, item and price inputs of
    Ultra-Procurement.ps1 (Phases D and E of docs/sap_export_contracts.md).

.DESCRIPTION
    1. Reads every .xlsx or .csv file in data/sap_exports: exports of ZCONTRACTS_REPORT,
       ME3N, ME3M, ME3L, ZPHA_CONTRACTS, Z_CONTR2, ZCONTR or Z_VENDOR_MATERIAL.
       Column headers are recognised with data/sap_column_map.json; add a header there
       when a report names a column differently. Title rows above the header are skipped.
    2. Classifies each file by its columns:
         contract lines  - contract number and item (position) columns. All such files
                           are joined on contract number + item.
         vendor material - supplier and material columns, no contract column
                           (Z_VENDOR_MATERIAL). Joined on supplier + material.
    3. Checks every joined contract line (Phase E):
         - unit price = Net Price / Per (price unit); the currency must be EUR;
         - if the order price unit is exported, it must equal the order unit;
         - a field may not have different values in two exports;
         - deleted or blocked items, text items without a material number and items
           with no price in any export are excluded and listed;
         - a material must have a single order unit across all contracts;
         - a supplier may not have two different unit prices for the same material in
           validity periods that overlap on or after the as-of date.
    4. Writes, only if no error was found:
         data/suppliers.xlsx                  Supplier Code, Supplier Name
         data/items_catalog.xlsx              Item Code, Description, Unit
         data/price_lists/sap_contracts.xlsx  one row per included contract line
       Existing files are replaced only with -Force, after a copy is kept in the run's
       backup folder.
    5. Always writes, for checking and archiving:
         output/sap_merge_<timestamp>/merged_contract_lines.xlsx  every line, status, reason
         output/sap_merge_<timestamp>/merge_summary.json          inputs (SHA-256), counts,
                                                                  errors, ME33K sample
         output/sap_merge_<timestamp>/generated/                  copies of the data files
         logs/sap_merge.log

    Nothing is invented: suppliers, materials, prices and dates come only from the
    exports. No Excel or extra modules are needed. Works on Windows PowerShell 5.1 and 7+.

.PARAMETER Root
    Project root (the folder with data/ and scripts/). Defaults to the parent of the
    folder this script is in.

.PARAMETER InputDir
    Folder with the SAP exports. Defaults to data/sap_exports.

.PARAMETER AsOfDate
    Date for the "valid on" column, the overlapping-price check and the ME33K sample.
    Defaults to today.

.PARAMETER Force
    Replace existing data/suppliers.xlsx, data/items_catalog.xlsx and
    data/price_lists/sap_contracts.xlsx. Backups are kept in the run folder.

.EXAMPLE
    powershell.exe -File scripts/Merge-SapContracts.ps1

.EXAMPLE
    powershell.exe -File scripts/Merge-SapContracts.ps1 -Force
#>
[CmdletBinding()]
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$InputDir,
    [datetime]$AsOfDate = (Get-Date).Date,
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
$Inv       = [System.Globalization.CultureInfo]::InvariantCulture
$Euro      = [string][char]0x20AC
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$RunStamp  = (Get-Date).ToString('yyyyMMdd-HHmmss', $Inv)
$RunId     = 'SAPMERGE-' + $RunStamp
$AsOfDate  = $AsOfDate.Date

# Logical fields read from the exports. Their header names are in data/sap_column_map.json.
$IdFields       = @('contract', 'item', 'supplier', 'material')
$NumberFields   = @('net_price', 'price_unit')
$DateFields     = @('valid_from', 'valid_to')
$TextFields     = @('supplier_name', 'description', 'supplier_material', 'currency', 'order_unit', 'order_price_unit',
                    'deletion_indicator', 'manufacturer', 'manufacturer_part')
$AllFields      = @($IdFields + $NumberFields + $DateFields + $TextFields)
$RequiredFields = @('contract', 'item', 'supplier', 'material', 'net_price', 'price_unit', 'currency', 'order_unit', 'valid_to')
# Free-text fields: different values in two exports are a warning, not an error.
$SoftFields     = @('supplier_name', 'description', 'supplier_material', 'manufacturer', 'manufacturer_part')
# Fields taken from vendor material exports (Z_VENDOR_MATERIAL).
$VmFields       = @('supplier_name', 'supplier_material', 'manufacturer', 'manufacturer_part')
# Deletion indicator values that mean "not deleted".
$NotDeleted     = @('0', 'false', 'no', 'n', 'οχι')
$SampleSize     = 5

# ---------------------------------------------------------------------------
# Logging and issue tracking
# ---------------------------------------------------------------------------
$script:LogFile  = $null
$script:Errors   = New-Object System.Collections.Generic.List[string]
$script:Warnings = New-Object System.Collections.Generic.List[string]

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'STEP', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffzzz', $Inv), $RunId, $Level.PadRight(5), $Message
    if ($script:LogFile) { [System.IO.File]::AppendAllText($script:LogFile, $line + [Environment]::NewLine, $Utf8NoBom) }
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'STEP'  { Write-Host $line -ForegroundColor Cyan }
        default { Write-Host $line }
    }
}

function Add-Issue {
    param([string]$Message, [switch]$Warning)
    if ($Warning) { $script:Warnings.Add($Message); Write-Log $Message 'WARN' }
    else { $script:Errors.Add($Message); Write-Log $Message 'ERROR' }
}

function Stop-IfErrors([string]$Stage) {
    if ($script:Errors.Count -gt 0) {
        throw "Validation failed during '$Stage' with $($script:Errors.Count) error(s). See logs/sap_merge.log and the run folder."
    }
}

# ---------------------------------------------------------------------------
# Value helpers
# ---------------------------------------------------------------------------
function ConvertTo-Key([string]$Text) {
    if ($null -eq $Text) { return '' }
    $d = $Text.Normalize([System.Text.NormalizationForm]::FormD)
    $d = [regex]::Replace($d, '\p{Mn}', '')
    return ([regex]::Replace($d, '[^\p{L}\p{Nd}]', '')).ToLowerInvariant()
}

# SAP numbers (contract, item, supplier, material): leading zeros are dropped from
# all-digit values, so 0000100023 and 100023 are the same supplier.
function ConvertTo-SapId([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $t = $Text.Trim()
    if ($t -match '^\d+\.0+$') { $t = $t.Substring(0, $t.IndexOf('.')) }
    if ($t -match '^\d+$') { $t = $t.TrimStart('0'); if ($t -eq '') { $t = '0' } }
    return $t.ToUpperInvariant()
}

function ConvertTo-Decimal([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = [regex]::Replace($Text, '[\s\u00A0' + $Euro + ']|EUR', '')
    if ($t -match '^-?\d{1,3}(\.\d{3})+,\d+$') { $t = $t.Replace('.', '').Replace(',', '.') }      # 1.234,56
    elseif ($t -match '^-?\d+,\d+$') { $t = $t.Replace(',', '.') }                                 # 12,50
    elseif ($t -match '^-?\d{1,3}(,\d{3})+(\.\d+)?$') { $t = $t.Replace(',', '') }                 # 1,234.56
    $d = [decimal]0
    if ([decimal]::TryParse($t, [System.Globalization.NumberStyles]::Float, $Inv, [ref]$d)) {
        # Excel stores numbers as binary floating point (0.1 can come back as 0.10000000000000001).
        return [Math]::Round($d, 6, [MidpointRounding]::AwayFromZero)
    }
    return $null
}

# Returns $null for an empty value, a [datetime], or the string 'INVALID'.
function ConvertTo-DateValue([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim()
    $dt = [datetime]::MinValue
    if ($t -match '^(19|20|99)\d{6}$') {
        if ([datetime]::TryParseExact($t, 'yyyyMMdd', $Inv, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt.Date }
        return 'INVALID'
    }
    if ($t -match '^\d+(\.\d+)?$') {
        $n = [double]::Parse($t, $Inv)
        if ($n -ge 1 -and $n -lt 2958466) { return [DateTime]::FromOADate($n).Date }
        return 'INVALID'
    }
    $formats = [string[]]@('dd.MM.yyyy', 'd.M.yyyy', 'dd/MM/yyyy', 'd/M/yyyy', 'yyyy-MM-dd', 'dd-MM-yyyy', 'yyyy-MM-ddTHH:mm:ss',
        'dd.MM.yyyy HH:mm:ss', 'dd/MM/yyyy HH:mm:ss', 'yyyy-MM-dd HH:mm:ss')
    if ([datetime]::TryParseExact($t, $formats, $Inv, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt.Date }
    return 'INVALID'
}

# "100023 ACME LTD" (ME3N "Supplier/Supplying Plant") -> code 100023, name ACME LTD.
function Split-SupplierValue([string]$Raw) {
    $t = $Raw.Trim()
    if ($t -match '^(\d+)\s+(\S.*)$') { return @{ Code = (ConvertTo-SapId $Matches[1]); Name = $Matches[2].Trim(); Valid = $true } }
    if ($t -match '\s') { return @{ Code = ''; Name = $null; Valid = $false } }
    return @{ Code = (ConvertTo-SapId $t); Name = $null; Valid = $true }
}

function Format-Number($Value) { if ($null -eq $Value) { return '' }; return ([decimal]$Value).ToString('0.######', $Inv) }
function Format-DateText($Value) { if ($null -eq $Value) { return '' }; return ([datetime]$Value).ToString('dd/MM/yyyy', $Inv) }

function Get-RelPath([string]$Path) {
    $full = [System.IO.Path]::GetFullPath($Path)
    if ($full.StartsWith($RootPath, [StringComparison]::OrdinalIgnoreCase)) { $full = $full.Substring($RootPath.Length).TrimStart('\', '/') }
    return $full.Replace('\', '/')
}

function Get-Sha256([string]$Path) { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }

function ConvertTo-ColumnName([int]$Index) {
    $n = $Index + 1; $s = ''
    while ($n -gt 0) { $m = ($n - 1) % 26; $s = ([string][char](65 + $m)) + $s; $n = [int][Math]::Floor(($n - 1) / 26) }
    return $s
}

# ---------------------------------------------------------------------------
# Readers (.xlsx without Excel, .csv). Each returns @{ Rows = list of @{ Row; Cells } }
# where Cells maps a 0-based column index to the trimmed cell text.
# ---------------------------------------------------------------------------
function Get-ZipXml($Zip, [string]$Name) {
    $entry = $null
    foreach ($e in $Zip.Entries) { if ($e.FullName.Replace('\', '/') -ieq $Name) { $entry = $e; break } }
    if ($null -eq $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try {
        $xml = New-Object System.Xml.XmlDocument
        $xml.XmlResolver = $null
        $xml.LoadXml($reader.ReadToEnd())
        return , $xml
    } finally { $reader.Dispose() }
}

function ConvertFrom-CellRef([string]$Ref) {
    $n = 0
    foreach ($ch in ($Ref -replace '[0-9]', '').ToUpperInvariant().ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n - 1
}

function Read-XlsxRows([string]$Path) {
    $ns  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $rns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $rows = New-Object System.Collections.Generic.List[object]
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $wb = Get-ZipXml $zip 'xl/workbook.xml'
            if ($null -eq $wb) { throw "'$Path' is not a valid .xlsx workbook." }
            $sheet = $wb.GetElementsByTagName('sheet', $ns)[0]
            if ($null -eq $sheet) { throw "'$Path' has no worksheets." }
            $rid = $sheet.GetAttribute('id', $rns)
            $target = $null
            $rels = Get-ZipXml $zip 'xl/_rels/workbook.xml.rels'
            foreach ($r in $rels.DocumentElement.ChildNodes) {
                if ($r -is [System.Xml.XmlElement] -and $r.GetAttribute('Id') -eq $rid) { $target = $r.GetAttribute('Target') }
            }
            if (-not $target) { throw "'$Path': first worksheet not found." }
            if ($target.StartsWith('/')) { $target = $target.TrimStart('/') } else { $target = 'xl/' + $target }

            $shared = New-Object System.Collections.Generic.List[string]
            $sst = Get-ZipXml $zip 'xl/sharedStrings.xml'
            if ($null -ne $sst) {
                foreach ($si in $sst.GetElementsByTagName('si', $ns)) {
                    $sb = New-Object System.Text.StringBuilder
                    foreach ($t in $si.GetElementsByTagName('t', $ns)) { if ($t.ParentNode.LocalName -ne 'rPh') { [void]$sb.Append($t.InnerText) } }
                    $shared.Add($sb.ToString())
                }
            }

            $ws = Get-ZipXml $zip $target
            if ($null -eq $ws) { throw "'$Path': worksheet '$target' is missing." }
            $rowCounter = 0
            foreach ($row in $ws.GetElementsByTagName('row', $ns)) {
                $rowCounter++
                $rowNum = $row.GetAttribute('r'); if (-not $rowNum) { $rowNum = $rowCounter }
                $cells = @{}; $seq = 0
                foreach ($c in $row.ChildNodes) {
                    if ($c.LocalName -ne 'c') { continue }
                    $ref = $c.GetAttribute('r')
                    $idx = if ($ref) { ConvertFrom-CellRef $ref } else { $seq }
                    $seq = $idx + 1
                    $type = $c.GetAttribute('t')
                    $v = $null
                    if ($type -eq 'inlineStr') { $v = $c.InnerText }
                    else {
                        $vn = $c.GetElementsByTagName('v', $ns)
                        if ($vn.Count -gt 0) {
                            $raw = $vn[0].InnerText
                            if ($type -eq 's') { $v = $shared[[int]$raw] }
                            elseif ($type -eq 'b') { $v = if ($raw -eq '1') { 'TRUE' } else { 'FALSE' } }
                            else { $v = $raw }
                        }
                    }
                    if ($null -ne $v -and "$v".Trim() -ne '') { $cells[[int]$idx] = "$v".Trim() }
                }
                if ($cells.Count -gt 0) { $rows.Add(@{ Row = [int]$rowNum; Cells = $cells }) }
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
    return @{ Rows = $rows }
}

function Split-CsvLine([string]$Line, [char]$Delim) {
    $out = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $inQuotes = $false
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($inQuotes) {
            if ($ch -eq [char]'"') {
                if ($i + 1 -lt $Line.Length -and $Line[$i + 1] -eq [char]'"') { [void]$sb.Append('"'); $i++ }
                else { $inQuotes = $false }
            } else { [void]$sb.Append($ch) }
        }
        elseif ($ch -eq [char]'"') { $inQuotes = $true }
        elseif ($ch -eq $Delim) { $out.Add($sb.ToString()); [void]$sb.Clear() }
        else { [void]$sb.Append($ch) }
    }
    $out.Add($sb.ToString())
    return , $out.ToArray()
}

function Read-CsvRows([string]$Path) {
    $lines = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    # Delimiter: the one that splits the first 20 non-empty lines into the most columns.
    $best = [char]';'; $bestCount = 0
    foreach ($cand in @([char]';', [char]"`t", [char]',')) {
        $max = 0; $seen = 0
        foreach ($l in $lines) {
            if ([string]::IsNullOrWhiteSpace($l)) { continue }
            $n = $l.Split($cand).Count
            if ($n -gt $max) { $max = $n }
            $seen++; if ($seen -ge 20) { break }
        }
        if ($max -gt $bestCount) { $best = $cand; $bestCount = $max }
    }
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ([string]::IsNullOrWhiteSpace($lines[$i])) { continue }
        $vals = Split-CsvLine $lines[$i] $best
        $cells = @{}
        for ($k = 0; $k -lt $vals.Length; $k++) { $v = $vals[$k].Trim(); if ($v -ne '') { $cells[[int]$k] = $v } }
        if ($cells.Count -gt 0) { $rows.Add(@{ Row = $i + 1; Cells = $cells }) }
    }
    return @{ Rows = $rows }
}

function Get-Cell($Cells, $Col) {
    if ($null -eq $Col) { return $null }
    if ($Cells.ContainsKey($Col.Index)) { return [string]$Cells[$Col.Index] }
    return $null
}

# The header is the first of the first 20 rows with at least two known column names.
function Find-HeaderRow($Rows) {
    $limit = [Math]::Min(20, $Rows.Count)
    for ($i = 0; $i -lt $limit; $i++) {
        $hits = 0
        foreach ($v in $Rows[$i].Cells.Values) { if ($AliasIndex.ContainsKey((ConvertTo-Key $v))) { $hits++ } }
        if ($hits -ge 2) { return $i }
    }
    return -1
}

# Maps each logical field to the first matching column: @{ field = @{ Index; Header } }.
function Get-ColumnMap([string]$FileName, $HeaderCells) {
    $byKey = @{}
    foreach ($idx in @($HeaderCells.Keys | Sort-Object)) {
        $text = [string]$HeaderCells[$idx]
        $k = ConvertTo-Key $text
        if ($k -eq '') { continue }
        if ($byKey.ContainsKey($k)) {
            if ($AliasIndex.ContainsKey($k)) {
                Add-Issue "${FileName}: column '$text' appears more than once; the first one (column $(ConvertTo-ColumnName $byKey[$k])) is used." -Warning
            }
            continue
        }
        $byKey[$k] = [int]$idx
    }
    $cols = @{}
    foreach ($field in $AllFields) {
        if (-not $Aliases.ContainsKey($field)) { continue }
        foreach ($a in $Aliases[$field]) {
            $k = ConvertTo-Key $a
            if ($byKey.ContainsKey($k)) { $cols[$field] = @{ Index = $byKey[$k]; Header = [string]$HeaderCells[$byKey[$k]] }; break }
        }
    }
    return $cols
}

# ---------------------------------------------------------------------------
# Records
# ---------------------------------------------------------------------------
function New-LineRecord([string]$Contract, [string]$Item) {
    return @{
        Contract = $Contract; Item = $Item; F = @{}
        Sources = (New-Object System.Collections.Generic.List[string])
        Problems = (New-Object System.Collections.Generic.List[string])
        Reasons = (New-Object System.Collections.Generic.List[string])
        Status = $null; UnitPrice = $null; ValidOn = $false; Sample = $false
    }
}

function Get-Raw($Record, [string]$Field) { if ($Record.F.ContainsKey($Field)) { return $Record.F[$Field].Raw }; return $null }
function Get-Norm($Record, [string]$Field) { if ($Record.F.ContainsKey($Field)) { return $Record.F[$Field].Norm }; return $null }

# Stores a field value; a different value from another export is an error
# (a warning for free-text fields). Supplier names are compared per supplier in
# step 5 (Add-SupplierName), so they are not reported here line by line.
function Set-RecordField($Record, [string]$Field, [string]$Raw, $Norm, [string]$Src, [string]$Label) {
    if (-not $Record.F.ContainsKey($Field)) { $Record.F[$Field] = @{ Raw = $Raw; Norm = $Norm; Src = $Src }; return }
    $old = $Record.F[$Field]
    if ($old.Norm -eq $Norm -or $Field -eq 'supplier_name') { return }
    $msg = "${Label}: $Field is '$($old.Raw)' in $($old.Src) but '$Raw' in $Src."
    if ($SoftFields -contains $Field) { Add-Issue $msg -Warning }
    else {
        Add-Issue $msg
        $Record.Problems.Add("$Field differs between exports: '$($old.Raw)' ($($old.Src)) / '$Raw' ($Src)")
    }
}

# Every supplier name seen in contract exports, per supplier code, in file order.
function Add-SupplierName([string]$Code, [string]$Name) {
    $k = ConvertTo-Key $Name
    if (-not $Code -or $k -eq '') { return }
    if (-not $NamesSeen.ContainsKey($Code)) { $NamesSeen[$Code] = [ordered]@{} }
    if (-not $NamesSeen[$Code].Contains($k)) { $NamesSeen[$Code][$k] = $Name.Trim() }
}

function Set-LineError($Record, [string]$Reason) {
    $Record.Status = 'ERROR'
    if (-not $Record.Reasons.Contains($Reason)) { $Record.Reasons.Add($Reason) }
}

# ---------------------------------------------------------------------------
# .xlsx writer (one worksheet, header row in bold, frozen)
# ---------------------------------------------------------------------------
function ConvertTo-XmlText([string]$Text) {
    if ($null -eq $Text) { return '' }
    $clean = [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return [System.Security.SecurityElement]::Escape($clean)
}

function Write-XlsxFile([string]$Path, [string]$SheetName, [string[]]$Headers, $Rows) {
    $widths = New-Object 'int[]' $Headers.Count
    $all = New-Object System.Collections.Generic.List[object]
    $all.Add([object[]]$Headers)
    foreach ($r in $Rows) { $all.Add([object[]]$r) }
    $sd = New-Object System.Text.StringBuilder
    for ($r = 0; $r -lt $all.Count; $r++) {
        $rowNum = $r + 1
        $cells = $all[$r]
        [void]$sd.Append("<row r=`"$rowNum`">")
        for ($c = 0; $c -lt $Headers.Count; $c++) {
            $v = if ($c -lt $cells.Count) { $cells[$c] } else { $null }
            if ($null -eq $v) { continue }
            $ref = (ConvertTo-ColumnName $c) + $rowNum
            if ($v -is [datetime]) {
                [void]$sd.Append("<c r=`"$ref`" s=`"2`"><v>$($v.Date.ToOADate().ToString('0', $Inv))</v></c>")
                $len = 10
            } elseif ($v -is [decimal] -or $v -is [int] -or $v -is [long] -or $v -is [double]) {
                $txt = ([decimal]$v).ToString('0.######', $Inv)
                [void]$sd.Append("<c r=`"$ref`"><v>$txt</v></c>")
                $len = $txt.Length
            } else {
                $txt = [string]$v
                if ($txt -eq '') { continue }
                $style = if ($r -eq 0) { ' s="1"' } else { '' }
                [void]$sd.Append("<c r=`"$ref`"$style t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText $txt)</t></is></c>")
                $len = $txt.Length
            }
            if ($len -gt $widths[$c]) { $widths[$c] = $len }
        }
        [void]$sd.Append('</row>')
    }
    $colsXml = New-Object System.Text.StringBuilder
    for ($c = 0; $c -lt $Headers.Count; $c++) {
        $w = [Math]::Min(60, [Math]::Max(8, $widths[$c] + 2))
        [void]$colsXml.Append("<col min=`"$($c + 1)`" max=`"$($c + 1)`" width=`"$w`" customWidth=`"1`"/>")
    }
    $decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    $main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $rel  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $pkg  = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $ct   = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
    $entries = [ordered]@{}
    $entries['[Content_Types].xml'] = $decl + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        "<Override PartName=`"/xl/workbook.xml`" ContentType=`"$ct.sheet.main+xml`"/>" +
        "<Override PartName=`"/xl/worksheets/sheet1.xml`" ContentType=`"$ct.worksheet+xml`"/>" +
        "<Override PartName=`"/xl/styles.xml`" ContentType=`"$ct.styles+xml`"/></Types>"
    $entries['_rels/.rels'] = $decl + "<Relationships xmlns=`"$pkg`">" +
        "<Relationship Id=`"rId1`" Type=`"$rel/officeDocument`" Target=`"xl/workbook.xml`"/></Relationships>"
    $entries['xl/workbook.xml'] = $decl + "<workbook xmlns=`"$main`" xmlns:r=`"$rel`"><sheets>" +
        "<sheet name=`"$(ConvertTo-XmlText $SheetName)`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
    $entries['xl/_rels/workbook.xml.rels'] = $decl + "<Relationships xmlns=`"$pkg`">" +
        "<Relationship Id=`"rId1`" Type=`"$rel/worksheet`" Target=`"worksheets/sheet1.xml`"/>" +
        "<Relationship Id=`"rId2`" Type=`"$rel/styles`" Target=`"styles.xml`"/></Relationships>"
    $entries['xl/styles.xml'] = $decl + "<styleSheet xmlns=`"$main`">" +
        '<numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    $entries['xl/worksheets/sheet1.xml'] = $decl + "<worksheet xmlns=`"$main`" xmlns:r=`"$rel`">" +
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        '<sheetFormatPr defaultRowHeight="15"/>' + "<cols>$($colsXml.ToString())</cols>" +
        "<sheetData>$($sd.ToString())</sheetData></worksheet>"

    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $entries.Keys) {
                $w = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open(), $Utf8NoBom)
                try { $w.Write($entries[$name]) } finally { $w.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

function Write-JsonFile([string]$Path, $Object) {
    [System.IO.File]::WriteAllText($Path, ($Object | ConvertTo-Json -Depth 10), $Utf8NoBom)
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$RootPath  = [System.IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
$DataDir   = Join-Path $RootPath 'data'
if (-not $InputDir) { $InputDir = Join-Path $DataDir 'sap_exports' }
$InputDir  = [System.IO.Path]::GetFullPath($InputDir)
$MapFile   = Join-Path $DataDir 'sap_column_map.json'
$LogDir    = Join-Path $RootPath 'logs'
$RunDir    = Join-Path (Join-Path $RootPath 'output') ('sap_merge_' + $RunStamp)
$StageDir  = Join-Path $RunDir 'generated'
$BackupDir = Join-Path $RunDir 'backup'
$ReportOut = Join-Path $RunDir 'merged_contract_lines.xlsx'
$JsonOut   = Join-Path $RunDir 'merge_summary.json'
$Targets   = [ordered]@{
    suppliers     = Join-Path $DataDir 'suppliers.xlsx'
    items_catalog = Join-Path $DataDir 'items_catalog.xlsx'
    price_list    = Join-Path (Join-Path $DataDir 'price_lists') 'sap_contracts.xlsx'
}

New-Item -ItemType Directory -Force -Path $LogDir, $RunDir | Out-Null
$script:LogFile = Join-Path $LogDir 'sap_merge.log'
$Inputs     = New-Object System.Collections.Generic.List[object]
$Published  = New-Object System.Collections.Generic.List[object]
$Backups    = New-Object System.Collections.Generic.List[object]
$Lines      = @{}
$VM         = @{}
$VmNames    = @{}
$NamesSeen  = @{}
$Aliases    = @{}
$AliasIndex = @{}
$Sorted     = @()
$Counts     = [ordered]@{}
$Excluded   = [ordered]@{}
$reportWritten = $false
$exitCode   = 0

function New-Summary([string]$Status) {
    return [ordered]@{
        run_id       = $RunId
        status       = $Status
        generated_at = (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz', $Inv)
        as_of_date   = $AsOfDate.ToString('yyyy-MM-dd', $Inv)
        input_folder = (Get-RelPath $InputDir)
        inputs       = $Inputs.ToArray()
        counts       = $Counts
        excluded_reasons = $Excluded
        me33k_sample = @($Sorted | Where-Object { $_.Sample } | ForEach-Object {
            [ordered]@{
                contract = $_.Contract; item = $_.Item; supplier = (Get-Raw $_ 'supplier'); material = (Get-Raw $_ 'material')
                net_price = (Get-Norm $_ 'net_price'); per = (Get-Norm $_ 'price_unit'); unit_price = $_.UnitPrice
                currency = (Get-Raw $_ 'currency'); order_unit = (Get-Raw $_ 'order_unit')
                valid_from = (Format-DateText (Get-Norm $_ 'valid_from')); valid_to = (Format-DateText (Get-Norm $_ 'valid_to'))
            }
        })
        outputs = [ordered]@{
            report     = $(if ($reportWritten) { Get-RelPath $ReportOut } else { $null })
            summary    = (Get-RelPath $JsonOut)
            log        = (Get-RelPath $script:LogFile)
            data_files = $Published.ToArray()
            backups    = $Backups.ToArray()
        }
        warnings = $script:Warnings.ToArray()
        errors   = $script:Errors.ToArray()
    }
}

try {
    Write-Log '===== SAP contract merge started =====' 'STEP'
    Write-Log "Root: $RootPath | Input: $(Get-RelPath $InputDir) | As-of date: $($AsOfDate.ToString('yyyy-MM-dd', $Inv)) | Host: PowerShell $($PSVersionTable.PSVersion)"

    # --- Step 1: column map --------------------------------------------------
    Write-Log 'Step 1/6: Loading column names (data/sap_column_map.json)' 'STEP'
    if (-not (Test-Path -LiteralPath $MapFile)) { Add-Issue 'Required file not found: data/sap_column_map.json' }
    Stop-IfErrors 'column map'
    $mapJson = [System.IO.File]::ReadAllText($MapFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    $Inputs.Add([ordered]@{ role = 'column_map'; file = (Get-RelPath $MapFile); sha256 = (Get-Sha256 $MapFile) })
    foreach ($f in $AllFields) {
        $p = $mapJson.PSObject.Properties[$f]
        $names = @()
        if ($null -ne $p -and $null -ne $p.Value) { $names = @($p.Value | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() -ne '' }) }
        if ($names.Count -eq 0) {
            if ($RequiredFields -contains $f) { Add-Issue "sap_column_map.json: no column names for required field '$f'." }
            continue
        }
        $Aliases[$f] = $names
        foreach ($a in $names) {
            $k = ConvertTo-Key $a
            if ($k -eq '') { Add-Issue "sap_column_map.json: column name '$a' in '$f' has no letters or digits."; continue }
            if ($AliasIndex.ContainsKey($k) -and $AliasIndex[$k] -ne $f) { Add-Issue "sap_column_map.json: '$a' is listed under both '$($AliasIndex[$k])' and '$f'."; continue }
            $AliasIndex[$k] = $f
        }
    }
    foreach ($p in $mapJson.PSObject.Properties) {
        if ($p.Name -notlike '_*' -and $AllFields -notcontains $p.Name) { Add-Issue "sap_column_map.json: unknown field '$($p.Name)' ignored." -Warning }
    }
    Stop-IfErrors 'column map'

    # --- Step 2: read exports ------------------------------------------------
    Write-Log "Step 2/6: Reading SAP exports ($(Get-RelPath $InputDir))" 'STEP'
    if (-not (Test-Path -LiteralPath $InputDir)) { Add-Issue "Input folder not found: $(Get-RelPath $InputDir)" }
    Stop-IfErrors 'input discovery'
    $files = @(Get-ChildItem -LiteralPath $InputDir -File | Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '.*' } | Sort-Object Name)
    $contractFiles = 0
    $opuSeen = $false
    foreach ($f in $files) {
        if ($f.Extension -notmatch '^\.(xlsx|csv)$') { Add-Issue "$($f.Name) skipped: export to .xlsx (Phase C) or save as CSV UTF-8." -Warning; continue }
        try { $t = if ($f.Extension -ieq '.xlsx') { Read-XlsxRows $f.FullName } else { Read-CsvRows $f.FullName } }
        catch {
            Add-Issue "$($f.Name): cannot be read ($($_.Exception.Message)). Is it open in Excel, or not a real $($f.Extension) file? Export it again."
            $Inputs.Add([ordered]@{ role = 'unreadable'; file = (Get-RelPath $f.FullName); sha256 = (Get-Sha256 $f.FullName) })
            continue
        }
        $rows = $t.Rows
        $h = Find-HeaderRow $rows
        if ($h -lt 0) {
            $first = if ($rows.Count -gt 0) { (@($rows[0].Cells.Keys | Sort-Object | ForEach-Object { $rows[0].Cells[$_] }) -join ' | ') } else { '(empty file)' }
            Add-Issue "$($f.Name): no header row with known column names in the first 20 rows. First row: $first"
            $Inputs.Add([ordered]@{ role = 'unrecognised'; file = (Get-RelPath $f.FullName); sha256 = (Get-Sha256 $f.FullName) })
            continue
        }
        $cols = Get-ColumnMap $f.Name $rows[$h].Cells
        $headerText = (@($rows[$h].Cells.Keys | Sort-Object | ForEach-Object { $rows[$h].Cells[$_] }) -join ' | ')
        $kind = $null
        if ($cols.ContainsKey('contract') -and $cols.ContainsKey('item')) { $kind = 'contract_lines' }
        elseif ($cols.ContainsKey('contract')) {
            Add-Issue "$($f.Name): has a contract number column but no item (position) column. Add the Item column with Change Layout and export again, or add its name to data/sap_column_map.json. Headers: $headerText"
        }
        elseif ($cols.ContainsKey('supplier') -and $cols.ContainsKey('material')) { $kind = 'vendor_material' }
        else {
            Add-Issue "$($f.Name): not recognised as contract lines (contract + item) or vendor material (supplier + material). Headers: $headerText. Add the missing column names to data/sap_column_map.json."
        }
        $mapped = [ordered]@{}
        foreach ($field in $AllFields) { if ($cols.ContainsKey($field)) { $mapped[$field] = $cols[$field].Header } }
        $info = [ordered]@{ role = $(if ($kind) { $kind } else { 'unrecognised' }); file = (Get-RelPath $f.FullName); sha256 = (Get-Sha256 $f.FullName)
                            header_row = $rows[$h].Row; data_rows = ($rows.Count - $h - 1); columns = $mapped }
        $Inputs.Add($info)
        if (-not $kind) { continue }
        Write-Log ("{0}: {1}, header on row {2}, {3} data row(s); columns: {4}" -f $f.Name, $kind, $rows[$h].Row, ($rows.Count - $h - 1),
            ((@($mapped.Keys) | ForEach-Object { "$_='$($mapped[$_])'" }) -join ', '))

        $hasNameCol = $cols.ContainsKey('supplier_name')
        $blank = 0; $used = 0
        $skipped = New-Object System.Collections.Generic.List[string]
        for ($i = $h + 1; $i -lt $rows.Count; $i++) {
            $cells = $rows[$i].Cells
            $src = "$($f.Name) row $($rows[$i].Row)"
            if ($kind -eq 'contract_lines') {
                $c  = ConvertTo-SapId (Get-Cell $cells $cols['contract'])
                $it = ConvertTo-SapId (Get-Cell $cells $cols['item'])
                if ($c -eq '' -and $it -eq '') { $blank++; continue }
                if ($c -notmatch '^\d+$' -or $it -notmatch '^\d+$') { $skipped.Add("$src (contract '$c', item '$it')"); continue }
                $key = "$c|$it"
                if (-not $Lines.ContainsKey($key)) { $Lines[$key] = New-LineRecord $c $it }
                $rec = $Lines[$key]
                $rec.Sources.Add($src)
                $label = "Contract $c item $it"
                $rowSupplier = $null
                $rowNames = New-Object System.Collections.Generic.List[string]
                foreach ($field in @($cols.Keys)) {
                    if ($field -eq 'contract' -or $field -eq 'item') { continue }
                    $raw = Get-Cell $cells $cols[$field]
                    if ([string]::IsNullOrWhiteSpace($raw)) { continue }
                    $raw = $raw.Trim()
                    $norm = $null
                    if ($field -eq 'supplier') {
                        $sp = Split-SupplierValue $raw
                        if (-not $sp.Valid) { Add-Issue "${src}: supplier '$raw' is not a supplier number."; $rec.Problems.Add("Supplier '$raw' is not a supplier number ($src)"); continue }
                        if ($sp.Name -and -not $hasNameCol) {
                            $rowNames.Add($sp.Name)
                            Set-RecordField $rec 'supplier_name' $sp.Name (ConvertTo-Key $sp.Name) $src $label
                        }
                        $raw = $sp.Code; $norm = $sp.Code; $rowSupplier = $sp.Code
                    }
                    elseif ($IdFields -contains $field) { $norm = ConvertTo-SapId $raw; $raw = $norm }
                    elseif ($NumberFields -contains $field) {
                        $norm = ConvertTo-Decimal $raw
                        if ($null -eq $norm) { Add-Issue "${src}: unreadable $field '$raw'."; $rec.Problems.Add("Unreadable $field '$raw' ($src)"); continue }
                        $raw = Format-Number $norm
                    }
                    elseif ($DateFields -contains $field) {
                        $norm = ConvertTo-DateValue $raw
                        if ($norm -is [string]) { Add-Issue "${src}: unreadable $field '$raw' (use dd.MM.yyyy or an Excel date)."; $rec.Problems.Add("Unreadable $field '$raw' ($src)"); continue }
                        $raw = Format-DateText $norm
                    }
                    elseif ($field -eq 'currency') {
                        $norm = $raw.ToUpperInvariant(); if ($norm -eq $Euro) { $norm = 'EUR' }; $raw = $norm
                    }
                    else { $norm = ConvertTo-Key $raw }
                    if ($field -eq 'supplier_name') { $rowNames.Add($raw) }
                    Set-RecordField $rec $field $raw $norm $src $label
                }
                $code = if ($rowSupplier) { $rowSupplier } else { Get-Norm $rec 'supplier' }
                foreach ($n in $rowNames) { Add-SupplierName $code $n }
            } else {
                $sRaw = Get-Cell $cells $cols['supplier']
                $m = ConvertTo-SapId (Get-Cell $cells $cols['material'])
                if ([string]::IsNullOrWhiteSpace($sRaw) -and $m -eq '') { $blank++; continue }
                $sp = $null
                if (-not [string]::IsNullOrWhiteSpace($sRaw)) { $sp = Split-SupplierValue $sRaw }
                if ($null -eq $sp -or -not $sp.Valid -or $sp.Code -eq '' -or $m -eq '') { $skipped.Add("$src (supplier '$sRaw', material '$m')"); continue }
                $key = "$($sp.Code)|$m"
                if (-not $VM.ContainsKey($key)) { $VM[$key] = @{ Supplier = $sp.Code; Material = $m; Vals = @{}; Sources = (New-Object System.Collections.Generic.List[string]) } }
                $rec = $VM[$key]
                $rec.Sources.Add($src)
                $nameVals = New-Object System.Collections.Generic.List[string]
                if ($sp.Name -and -not $hasNameCol) { $nameVals.Add($sp.Name) }
                foreach ($field in $VmFields) {
                    $raw = $null
                    if ($field -eq 'supplier_name' -and $nameVals.Count -gt 0) { $raw = $nameVals[0] }
                    elseif ($cols.ContainsKey($field)) { $raw = Get-Cell $cells $cols[$field] }
                    if ([string]::IsNullOrWhiteSpace($raw)) { continue }
                    $raw = $raw.Trim()
                    if ($field -eq 'supplier_name') {
                        if (-not $VmNames.ContainsKey($sp.Code)) { $VmNames[$sp.Code] = [ordered]@{} }
                        $nk = ConvertTo-Key $raw
                        if (-not $VmNames[$sp.Code].Contains($nk)) { $VmNames[$sp.Code][$nk] = $raw }
                        continue
                    }
                    if (-not $rec.Vals.ContainsKey($field)) { $rec.Vals[$field] = [ordered]@{} }
                    $nk = ConvertTo-Key $raw
                    if (-not $rec.Vals[$field].Contains($nk)) { $rec.Vals[$field][$nk] = $raw }
                }
            }
            $used++
        }
        $info['rows_used'] = $used
        $info['rows_blank_keys'] = $blank
        $info['rows_skipped'] = $skipped.Count
        if ($kind -eq 'contract_lines') { $contractFiles++; if ($cols.ContainsKey('order_price_unit')) { $opuSeen = $true } }
        if ($skipped.Count -gt 0) {
            Add-Issue ("{0}: {1} row(s) skipped because the key columns are not SAP numbers, e.g. {2}." -f $f.Name, $skipped.Count,
                ((@($skipped) | Select-Object -First 3) -join '; ')) -Warning
        }
        if ($blank -gt 0) { Write-Log "$($f.Name): $blank row(s) without key values ignored (blank or total rows)." }
    }
    # Value errors found so far are kept on their lines and stop the run after the
    # report is written; only a missing contract export stops it here.
    if ($files.Count -eq 0) { Add-Issue "No export files found in $(Get-RelPath $InputDir)."; Stop-IfErrors 'reading exports' }
    if ($contractFiles -eq 0) { Add-Issue "No contract line export (contract number + item columns) found in $(Get-RelPath $InputDir)."; Stop-IfErrors 'reading exports' }
    if (-not $opuSeen) {
        Add-Issue 'No contract export has an order price unit column, so prices could not be checked to be per order unit. Add Order Price Unit to the ME3N layout, or check this by hand.' -Warning
    }
    Write-Log "Contract lines joined: $($Lines.Count); vendor material records: $($VM.Count)"

    # --- Step 3: vendor material data ----------------------------------------
    Write-Log 'Step 3/6: Adding vendor material data (supplier + material)' 'STEP'
    $vmUsed = @{}
    foreach ($rec in $Lines.Values) {
        $s = Get-Norm $rec 'supplier'; $m = Get-Norm $rec 'material'
        if (-not $s -or -not $m) { continue }
        $k = "$s|$m"
        if (-not $VM.ContainsKey($k)) { continue }
        $v = $VM[$k]
        $first = -not $vmUsed.ContainsKey($k)
        $vmUsed[$k] = $true
        foreach ($field in @('supplier_material', 'manufacturer', 'manufacturer_part')) {
            if (-not $v.Vals.ContainsKey($field)) { continue }
            $vals = @($v.Vals[$field].Values)
            $joined = $vals -join ' / '
            if ($vals.Count -gt 1 -and $first) {
                Add-Issue "Supplier $s material ${m}: the vendor material export has several $field values ($joined); all are shown." -Warning
            }
            if ($rec.F.ContainsKey($field)) {
                if (-not $v.Vals[$field].Contains($rec.F[$field].Norm)) {
                    Add-Issue ("Contract {0} item {1}: {2} is '{3}' in {4} but '{5}' in the vendor material export." -f $rec.Contract, $rec.Item, $field,
                        $rec.F[$field].Raw, $rec.F[$field].Src, $joined) -Warning
                }
            } else {
                $rec.F[$field] = @{ Raw = $joined; Norm = (ConvertTo-Key $joined); Src = ($v.Sources | Select-Object -First 1) }
            }
        }
    }
    Write-Log "Vendor material records matched to contract lines: $($vmUsed.Count) of $($VM.Count)"

    # --- Step 4: line checks -------------------------------------------------
    Write-Log 'Step 4/6: Checking each contract line' 'STEP'
    $Sorted = @($Lines.Values | Sort-Object @{ e = { $_.Contract.PadLeft(20, '0') } }, @{ e = { $_.Item.PadLeft(10, '0') } })
    foreach ($rec in $Sorted) {
        $label = "Contract $($rec.Contract) item $($rec.Item)"
        $errs = New-Object System.Collections.Generic.List[string]
        foreach ($p in $rec.Problems) { $errs.Add($p) }
        $exclude = $null
        $del = Get-Norm $rec 'deletion_indicator'
        $mat = Get-Norm $rec 'material'
        $net = Get-Norm $rec 'net_price'
        if ($del -and $NotDeleted -notcontains $del) { $exclude = "Deleted or blocked in SAP (deletion indicator '$(Get-Raw $rec 'deletion_indicator')')" }
        elseif (-not $mat) { $exclude = 'No material number (text item)' }
        elseif ($null -eq $net) { $exclude = 'No net price in any export' }
        else {
            $line = New-Object System.Collections.Generic.List[string]
            if (-not (Get-Norm $rec 'supplier')) { $line.Add('No supplier code') }
            if ($net -le 0) { $line.Add("Net price must be greater than zero (found $(Format-Number $net))") }
            $per = Get-Norm $rec 'price_unit'
            if ($null -eq $per) { $line.Add('No price unit (Per)') }
            elseif ($per -le 0) { $line.Add("Price unit (Per) must be greater than zero (found $(Format-Number $per))") }
            $cur = Get-Norm $rec 'currency'
            if (-not $cur) { $line.Add('No currency') }
            elseif ($cur -ne 'EUR') { $line.Add("Currency is $cur, not EUR") }
            $ou = Get-Norm $rec 'order_unit'
            $opu = Get-Norm $rec 'order_price_unit'
            if (-not $ou) { $line.Add('No order unit') }
            elseif ($opu -and $opu -ne $ou) {
                $line.Add("Price is per order price unit '$(Get-Raw $rec 'order_price_unit')' but the order unit is '$(Get-Raw $rec 'order_unit')'; the conversion factor is not in the exports")
            }
            $vf = Get-Norm $rec 'valid_from'; $vt = Get-Norm $rec 'valid_to'
            if ($null -eq $vt) { $line.Add('No validity end date') }
            elseif ($null -ne $vf -and $vf -gt $vt) { $line.Add('Validity start is after validity end') }
            foreach ($e in $line) { Add-Issue "${label}: $e."; $errs.Add($e) }
            if ($line.Count -eq 0) { $rec.UnitPrice = [Math]::Round($net / $per, 6, [MidpointRounding]::AwayFromZero) }
        }
        $vf = Get-Norm $rec 'valid_from'; $vt = Get-Norm $rec 'valid_to'
        $rec.ValidOn = ($null -ne $vt -and $vt -ge $AsOfDate -and ($null -eq $vf -or $vf -le $AsOfDate))
        if ($errs.Count -gt 0) { $rec.Status = 'ERROR'; foreach ($e in $errs) { $rec.Reasons.Add($e) } }
        elseif ($exclude) {
            $rec.Status = 'EXCLUDED'; $rec.Reasons.Add($exclude)
            $reasonKey = $exclude -replace " \(deletion indicator '.*'\)$", ''
            if (-not $Excluded.Contains($reasonKey)) { $Excluded[$reasonKey] = 0 }
            $Excluded[$reasonKey]++
        }
        else { $rec.Status = 'INCLUDED' }
    }
    foreach ($k in @($Excluded.Keys)) { Add-Issue "$($Excluded[$k]) contract line(s) excluded: $k. See the report." -Warning }

    # --- Step 5: cross-line checks -------------------------------------------
    Write-Log 'Step 5/6: Cross-checking units, overlapping prices and supplier names' 'STEP'
    # Every line with a unit price takes part, including lines already in error, so
    # that one run reports all problems.
    $priced = @($Sorted | Where-Object { $null -ne $_.UnitPrice })
    foreach ($g in @($priced | Group-Object { Get-Norm $_ 'material' })) {
        $units = @($g.Group | ForEach-Object { Get-Norm $_ 'order_unit' } | Select-Object -Unique)
        if ($units.Count -gt 1) {
            $detail = (@($g.Group | ForEach-Object { "$(Get-Raw $_ 'order_unit') in $($_.Contract)/$($_.Item)" }) -join ', ')
            Add-Issue "Material $(Get-Raw $g.Group[0] 'material') has different order units: $detail. The catalog needs one unit per item."
            foreach ($r in $g.Group) { Set-LineError $r "Material has different order units across contract lines: $detail" }
        }
    }

    foreach ($g in @($priced | Group-Object { "$(Get-Norm $_ 'supplier')|$(Get-Norm $_ 'material')" })) {
        $grp = @($g.Group)
        for ($i = 0; $i -lt $grp.Count; $i++) {
            for ($j = $i + 1; $j -lt $grp.Count; $j++) {
                $a = $grp[$i]; $b = $grp[$j]
                if ($a.UnitPrice -eq $b.UnitPrice) { continue }
                $af = Get-Norm $a 'valid_from'; if ($null -eq $af) { $af = [datetime]::MinValue }
                $bf = Get-Norm $b 'valid_from'; if ($null -eq $bf) { $bf = [datetime]::MinValue }
                $at = Get-Norm $a 'valid_to'; $bt = Get-Norm $b 'valid_to'
                if ($af -gt $bt -or $bf -gt $at) { continue }
                $end = if ($at -lt $bt) { $at } else { $bt }
                $msg = "Supplier {0}, material {1}: unit price {2} in {3}/{4} and {5} in {6}/{7} with overlapping validity" -f (Get-Raw $a 'supplier'), (Get-Raw $a 'material'),
                    (Format-Number $a.UnitPrice), $a.Contract, $a.Item, (Format-Number $b.UnitPrice), $b.Contract, $b.Item
                if ($end -ge $AsOfDate) {
                    Add-Issue "$msg. Export each hospital or plant separately, or correct the contracts, so that one price applies."
                    Set-LineError $a "$msg (other line $($b.Contract)/$($b.Item))"
                    Set-LineError $b "$msg (other line $($a.Contract)/$($a.Item))"
                } else {
                    Add-Issue "$msg; the overlap ended on $(Format-DateText $end), before the as-of date." -Warning
                }
            }
        }
    }

    # A name read before its line's supplier code was known (code from a later file).
    foreach ($r in $Sorted) { Add-SupplierName (Get-Norm $r 'supplier') (Get-Raw $r 'supplier_name') }
    $SupplierNames = @{}
    foreach ($code in @($priced | ForEach-Object { Get-Norm $_ 'supplier' } | Select-Object -Unique)) {
        $list = @()
        if ($NamesSeen.ContainsKey($code)) { $list = @($NamesSeen[$code].Values) }
        elseif ($VmNames.ContainsKey($code)) { $list = @($VmNames[$code].Values) }
        if ($list.Count -eq 0) {
            Add-Issue "Supplier $code has no name in any export. Add the supplier name column to the layout and export again."
            foreach ($r in $priced) { if ((Get-Norm $r 'supplier') -eq $code) { Set-LineError $r 'Supplier has no name in any export' } }
            continue
        }
        if ($list.Count -gt 1) { Add-Issue "Supplier ${code}: several names in the exports ($($list -join ' | ')); '$($list[0])' is used." -Warning }
        $SupplierNames[$code] = $list[0]
    }

    $inc = @($Sorted | Where-Object { $_.Status -eq 'INCLUDED' })
    $valid = @($inc | Where-Object { $_.ValidOn })
    if ($valid.Count -gt 0) {
        $n = [Math]::Min($SampleSize, $valid.Count)
        for ($i = 0; $i -lt $n; $i++) { $valid[[int][Math]::Floor($i * $valid.Count / $n)].Sample = $true }
    }

    foreach ($k in $Targets.Keys) {
        if ((Test-Path -LiteralPath $Targets[$k]) -and -not $Force) {
            Add-Issue "$(Get-RelPath $Targets[$k]) already exists. Run again with -Force to replace it (a backup is kept in the run folder)."
        }
    }

    # --- Step 6: report, data files ------------------------------------------
    Write-Log 'Step 6/6: Writing the report and the data files' 'STEP'
    $materials = [ordered]@{}
    foreach ($r in $inc) {
        $m = Get-Norm $r 'material'
        if (-not $materials.Contains($m)) { $materials[$m] = $r }
    }
    $Counts['contract_lines']        = $Sorted.Count
    $Counts['included']              = $inc.Count
    $Counts['excluded']              = @($Sorted | Where-Object { $_.Status -eq 'EXCLUDED' }).Count
    $Counts['error_lines']           = @($Sorted | Where-Object { $_.Status -eq 'ERROR' }).Count
    $Counts['valid_on_as_of_date']   = @($inc | Where-Object { $_.ValidOn }).Count
    $Counts['suppliers']             = $SupplierNames.Count
    $Counts['items']                 = $materials.Count
    $Counts['vendor_material_records'] = $VM.Count
    $Counts['vendor_material_matched'] = $vmUsed.Count

    $validHeader = 'Valid On ' + $AsOfDate.ToString('dd/MM/yyyy', $Inv)
    $reportRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $Sorted) {
        $code = Get-Norm $r 'supplier'
        $sName = if ($code -and $SupplierNames.ContainsKey($code)) { $SupplierNames[$code] } else { Get-Raw $r 'supplier_name' }
        $reportRows.Add([object[]]@(
            $r.Status, ($r.Reasons -join '; '), $r.Contract, $r.Item, (Get-Raw $r 'supplier'), $sName,
            (Get-Raw $r 'material'), (Get-Raw $r 'description'), (Get-Raw $r 'order_unit'), (Get-Raw $r 'order_price_unit'),
            (Get-Norm $r 'net_price'), (Get-Norm $r 'price_unit'), $r.UnitPrice, (Get-Raw $r 'currency'),
            (Get-Norm $r 'valid_from'), (Get-Norm $r 'valid_to'), $(if ($r.ValidOn) { 'YES' } else { 'NO' }),
            (Get-Raw $r 'deletion_indicator'), (Get-Raw $r 'supplier_material'), (Get-Raw $r 'manufacturer'), (Get-Raw $r 'manufacturer_part'),
            $(if ($r.Sample) { 'YES' } else { '' }), ($r.Sources -join '; ')))
    }
    Write-XlsxFile $ReportOut 'Contract lines' @('Status', 'Reason', 'SAP Contract', 'Contract Item', 'Supplier Code', 'Supplier Name',
        'Material', 'Description', 'Order Unit', 'Order Price Unit', 'Net Price', 'Per', 'Unit Price', 'Currency', 'Valid From', 'Valid To',
        $validHeader, 'Deletion Indicator', 'Supplier Material No.', 'Manufacturer', 'Manufacturer Part No.', 'ME33K Check', 'Sources') $reportRows
    $reportWritten = $true
    Write-Log "Written $(Get-RelPath $ReportOut) ($($Sorted.Count) line(s))"
    Write-Log ("Lines: {0} included, {1} excluded, {2} with errors; {3} supplier(s), {4} item(s)" -f $Counts['included'], $Counts['excluded'],
        $Counts['error_lines'], $Counts['suppliers'], $Counts['items'])
    Stop-IfErrors 'checks'
    if ($inc.Count -eq 0) { Add-Issue 'No contract line passed the checks; nothing to write.' }
    Stop-IfErrors 'checks'

    New-Item -ItemType Directory -Force -Path $StageDir | Out-Null
    $staged = [ordered]@{
        suppliers     = Join-Path $StageDir 'suppliers.xlsx'
        items_catalog = Join-Path $StageDir 'items_catalog.xlsx'
        price_list    = Join-Path $StageDir 'sap_contracts.xlsx'
    }
    $supRows = New-Object System.Collections.Generic.List[object]
    foreach ($code in @($SupplierNames.Keys | Sort-Object { $_.PadLeft(20, '0') })) { $supRows.Add([object[]]@($code, $SupplierNames[$code])) }
    Write-XlsxFile $staged['suppliers'] 'Suppliers' @('Supplier Code', 'Supplier Name') $supRows

    $catRows = New-Object System.Collections.Generic.List[object]
    foreach ($m in @($materials.Keys | Sort-Object { $_.PadLeft(40, '0') })) {
        $r = $materials[$m]
        $catRows.Add([object[]]@((Get-Raw $r 'material'), (Get-Raw $r 'description'), (Get-Raw $r 'order_unit')))
    }
    Write-XlsxFile $staged['items_catalog'] 'Items' @('Item Code', 'Description', 'Unit') $catRows

    $plRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $inc) {
        $code = Get-Norm $r 'supplier'
        $plRows.Add([object[]]@($code, $SupplierNames[$code], (Get-Raw $r 'material'), (Get-Raw $r 'description'), (Get-Raw $r 'order_unit'),
            $r.UnitPrice, 'EUR', (Get-Norm $r 'valid_from'), (Get-Norm $r 'valid_to'), $r.Contract, $r.Item,
            (Get-Norm $r 'net_price'), (Get-Norm $r 'price_unit'), (Get-Raw $r 'supplier_material'), (Get-Raw $r 'manufacturer'),
            (Get-Raw $r 'manufacturer_part')))
    }
    Write-XlsxFile $staged['price_list'] 'SAP contracts' @('Supplier Code', 'Supplier Name', 'Item Code', 'Description', 'Unit', 'Unit Price',
        'Currency', 'Valid From', 'Valid To', 'SAP Contract', 'Contract Item', 'Net Price', 'Price Unit (Per)', 'Supplier Material No.',
        'Manufacturer', 'Manufacturer Part No.') $plRows

    $stagedRows = @{ suppliers = $supRows.Count; items_catalog = $catRows.Count; price_list = $plRows.Count }
    foreach ($k in $Targets.Keys) {
        $dest = $Targets[$k]
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
        if (Test-Path -LiteralPath $dest) {
            New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
            $bk = Join-Path $BackupDir ([System.IO.Path]::GetFileName($dest))
            Copy-Item -LiteralPath $dest -Destination $bk -Force
            $Backups.Add([ordered]@{ file = (Get-RelPath $dest); backup = (Get-RelPath $bk); sha256 = (Get-Sha256 $bk) })
            Write-Log "Backed up $(Get-RelPath $dest) to $(Get-RelPath $bk)"
        }
        Copy-Item -LiteralPath $staged[$k] -Destination $dest -Force
        $Published.Add([ordered]@{ file = (Get-RelPath $dest); rows = $stagedRows[$k]; sha256 = (Get-Sha256 $dest) })
        Write-Log "Written $(Get-RelPath $dest) ($($stagedRows[$k]) row(s))"
    }

    Write-JsonFile $JsonOut (New-Summary 'COMPLETED')
    Write-Log "Written $(Get-RelPath $JsonOut)"
    Write-Log ("===== Completed: {0} contract line(s) written for {1} supplier(s) and {2} item(s); {3} excluded; {4} warning(s) =====" -f
        $plRows.Count, $supRows.Count, $catRows.Count, $Counts['excluded'], $script:Warnings.Count) 'STEP'
    if ($Counts['valid_on_as_of_date'] -gt 0) {
        Write-Log "Phase E: check the $(@($Sorted | Where-Object { $_.Sample }).Count) line(s) marked 'ME33K Check' in the report against ME33K > Conditions." 'STEP'
    }
}
catch {
    $unexpected = ($script:Errors.Count -eq 0)
    $exitCode = if ($unexpected) { 1 } else { 2 }
    if ($unexpected) { $script:Errors.Add($_.Exception.Message) }
    Write-Log "Run FAILED: $($_.Exception.Message)" 'ERROR'
    if ($unexpected -and $_.InvocationInfo) { Write-Log "At: $($_.InvocationInfo.PositionMessage -replace '\s+', ' ')" 'ERROR' }
    try {
        Write-JsonFile $JsonOut (New-Summary 'FAILED')
        $changed = if ($Published.Count -gt 0) {
            'these data files were already replaced: ' + ((@($Published) | ForEach-Object { $_.file }) -join ', ') + ' (backups in ' + (Get-RelPath $BackupDir) + ')'
        } else { 'nothing in data/ was changed' }
        Write-Log "Failure summary written to $(Get-RelPath $JsonOut); $changed." 'ERROR'
    } catch { Write-Log "Could not write failure summary: $($_.Exception.Message)" 'ERROR' }
}
exit $exitCode
