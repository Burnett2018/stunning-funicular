<#
.SYNOPSIS
    Unified procurement automation for the kitchens of Limassol General Hospital
    and Troodos Hospital (Law 73(I)/2016 as amended by Law 113(I)/2025, Article 90(1)).

.DESCRIPTION
    1. Loads data/kitchens/*.xlsx, data/suppliers.xlsx, data/items_catalog.xlsx,
       data/price_lists/* (.xlsx or .csv) and data/legal/*.pdf.
    2. Builds the full item -> supplier -> price matrix.
    3. Applies Article 90(1)(a) (simple acceptance of offer) when the total cost is
       <= the configured threshold, otherwise Article 90(1)(b) (summary award).
    4. Produces ONE unified procurement package from templates/unified_template.docx:
       price comparison, supplier selection justification, acceptance-of-offer or
       summary award text, final order text, plus a JSON summary.
    5. Logs every step to logs/procurement.log.

    Outputs (in output/<RequestId>/):
        unified_order.docx
        unified_summary.json

    The script never invents data: every supplier, item and price must come from the
    files in data/. Any missing or invalid value stops the run with a logged error
    and a FAILED unified_summary.json; no order document is produced in that case.

    The Article 90(1)(a) threshold is NOT hard-coded. Set it in
    data/procurement_config.json ("threshold_90_1_a_eur") or pass -Threshold90a.

    No Excel, Word or extra PowerShell modules are needed; .xlsx and .docx files are
    read and written directly. Works on Windows PowerShell 5.1 and PowerShell 7+.

.PARAMETER Root
    Project root (the folder containing data/, templates/, scripts/). Defaults to the
    parent of the folder this script is in.

.PARAMETER Threshold90a
    Article 90(1)(a) threshold in EUR, excluding VAT. Overrides the config file.

.PARAMETER Threshold90b
    Optional upper limit in EUR (excl. VAT) for Article 90(1)(b). If set and the total
    exceeds it, the run stops: the request falls outside Article 90(1).

.PARAMETER AsOfDate
    Date used to check price-list validity (ValidFrom / ValidTo). Defaults to today.

.EXAMPLE
    powershell.exe -File scripts/Ultra-Procurement.ps1

.EXAMPLE
    pwsh -File scripts/Ultra-Procurement.ps1 -Threshold90a 10000
#>
[CmdletBinding()]
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [Nullable[decimal]]$Threshold90a = $null,
    [Nullable[decimal]]$Threshold90b = $null,
    [datetime]$AsOfDate = (Get-Date).Date,
    [string]$OutputDir
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
$LawText   = 'Law 73(I)/2016, as amended by Law 113(I)/2025'
$Authority = 'Procurement Office of Limassol General Hospital and Troodos Hospital'
$RequestId = 'PRC-' + (Get-Date).ToString('yyyyMMdd-HHmmss', $Inv)

# Accepted column headers (case, spaces, punctuation and Greek accents are ignored).
$Cols = @{
    SupplierId   = @('Supplier Id', 'Supplier Code', 'Vendor Id', 'Vendor Code', 'Supplier', 'Κωδικός Προμηθευτή', 'Προμηθευτής')
    SupplierKey  = @('Supplier Id', 'Supplier Code', 'Vendor Id', 'Vendor Code', 'Code', 'Id', 'Κωδικός Προμηθευτή', 'Κωδικός')
    SupplierName = @('Supplier Name', 'Name', 'Company', 'Company Name', 'Επωνυμία', 'Όνομα')
    Vat          = @('VAT', 'VAT Number', 'VAT No', 'TIC', 'Tax Id', 'ΑΦΜ', 'Αριθμός ΦΠΑ')
    Email        = @('Email', 'E-mail', 'Email Address')
    Phone        = @('Phone', 'Telephone', 'Tel', 'Τηλέφωνο')
    Address      = @('Address', 'Διεύθυνση')
    Active       = @('Active', 'Status', 'Ενεργός', 'Κατάσταση')
    ItemCode     = @('Item Code', 'Item Id', 'Code', 'SKU', 'Κωδικός Είδους', 'Κωδικός')
    Description  = @('Description', 'Item Description', 'Item Name', 'Item', 'Name', 'Περιγραφή', 'Είδος')
    Unit         = @('Unit', 'UoM', 'Unit of Measure', 'Μονάδα', 'Μονάδα Μέτρησης')
    Quantity     = @('Quantity', 'Qty', 'Requested Quantity', 'Ποσότητα')
    UnitPrice    = @('Unit Price', 'Price', 'Unit Price excl VAT', 'Price excl VAT', 'Τιμή Μονάδας', 'Τιμή')
    ValidFrom    = @('Valid From', 'From', 'Ισχύει Από')
    ValidTo      = @('Valid To', 'Valid Until', 'To', 'Ισχύει Έως')
    Currency     = @('Currency', 'Νόμισμα')
    Hospital     = @('Hospital', 'Νοσοκομείο')
}

# ---------------------------------------------------------------------------
# Logging and issue tracking
# ---------------------------------------------------------------------------
$script:LogFile  = $null
$script:Errors   = New-Object System.Collections.Generic.List[string]
$script:Warnings = New-Object System.Collections.Generic.List[string]

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'STEP', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffzzz', $Inv), $RequestId, $Level.PadRight(5), $Message
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
        throw "Validation failed during '$Stage' with $($script:Errors.Count) error(s). See the log."
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

function ConvertTo-Decimal([string]$Text) {
    if ($null -eq $Text) { return $null }
    $t = [regex]::Replace($Text, '[\s ' + $Euro + ']|EUR', '')
    if ($t -match '^-?\d{1,3}(\.\d{3})+,\d+$') { $t = $t.Replace('.', '').Replace(',', '.') }      # 1.234,56
    elseif ($t -match '^-?\d+,\d+$') { $t = $t.Replace(',', '.') }                                 # 12,50
    elseif ($t -match '^-?\d{1,3}(,\d{3})+(\.\d+)?$') { $t = $t.Replace(',', '') }                 # 1,234.56
    $d = [decimal]0
    if ([decimal]::TryParse($t, [System.Globalization.NumberStyles]::Float, $Inv, [ref]$d)) { return $d }
    return $null
}

function ConvertTo-DateValue([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    if ($Text -match '^\d+(\.\d+)?$') { return [DateTime]::FromOADate([double]::Parse($Text, $Inv)).Date }
    $formats = [string[]]@('yyyy-MM-dd', 'dd/MM/yyyy', 'd/M/yyyy', 'dd.MM.yyyy', 'd.M.yyyy', 'dd-MM-yyyy', 'yyyy-MM-ddTHH:mm:ss')
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact($Text.Trim(), $formats, $Inv, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt.Date }
    return 'INVALID'
}

function Round-Money([decimal]$Value) { return [Math]::Round($Value, 2, [MidpointRounding]::AwayFromZero) }
function Format-Money([decimal]$Value) { return $Euro + ' ' + $Value.ToString('#,##0.00', $Inv) }
function Format-Price([decimal]$Value) { return $Euro + ' ' + $Value.ToString('#,##0.00##', $Inv) }
function Format-Qty([decimal]$Value) { return $Value.ToString('#,##0.###', $Inv) }

function Get-RelPath([string]$Path) {
    $full = [System.IO.Path]::GetFullPath($Path)
    if ($full.StartsWith($RootPath, [StringComparison]::OrdinalIgnoreCase)) { $full = $full.Substring($RootPath.Length).TrimStart('\', '/') }
    return $full.Replace('\', '/')
}

function Get-Sha256([string]$Path) { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }

# ---------------------------------------------------------------------------
# Table readers (.xlsx without Excel, .csv)
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

function Read-XlsxTable([string]$Path) {
    $ns  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $rns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
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
            $header = $null
            $rows = New-Object System.Collections.Generic.List[object]
            $rowCounter = 0
            foreach ($row in $ws.GetElementsByTagName('row', $ns)) {
                $rowCounter++
                $rowNum = $row.GetAttribute('r'); if (-not $rowNum) { $rowNum = $rowCounter }
                $vals = @{}; $seq = 0
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
                    if ($null -ne $v -and "$v".Trim() -ne '') { $vals[$idx] = "$v".Trim() }
                }
                if ($vals.Count -eq 0) { continue }
                if ($null -eq $header) { $header = $vals; continue }
                $rec = @{ '__row' = [int]$rowNum }
                foreach ($k in $vals.Keys) { if ($header.ContainsKey($k)) { $rec[(ConvertTo-Key $header[$k])] = $vals[$k] } }
                $rows.Add($rec)
            }
            $headers = @()
            if ($null -ne $header) { $headers = @($header.Keys | Sort-Object | ForEach-Object { $header[$_] }) }
            return @{ Headers = $headers; Rows = $rows }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

function Read-CsvTable([string]$Path) {
    $first = (Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8)
    $delim = if (($first.Split(';').Count) -gt ($first.Split(',').Count)) { ';' } else { ',' }
    $data = @(Import-Csv -LiteralPath $Path -Delimiter $delim -Encoding UTF8)
    $rows = New-Object System.Collections.Generic.List[object]
    $line = 1
    foreach ($d in $data) {
        $line++
        $rec = @{ '__row' = $line }; $any = $false
        foreach ($p in $d.PSObject.Properties) {
            if ($null -ne $p.Value -and "$($p.Value)".Trim() -ne '') { $rec[(ConvertTo-Key $p.Name)] = "$($p.Value)".Trim(); $any = $true }
        }
        if ($any) { $rows.Add($rec) }
    }
    $headers = @()
    if ($data.Count -gt 0) { $headers = @($data[0].PSObject.Properties | ForEach-Object { $_.Name }) }
    return @{ Headers = $headers; Rows = $rows }
}

function Read-Table([string]$Path) {
    switch -Regex ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '^\.xlsx$' { return Read-XlsxTable $Path }
        '^\.csv$'  { return Read-CsvTable $Path }
        default    { throw "Unsupported file type '$([System.IO.Path]::GetFileName($Path))' (use .xlsx or .csv)." }
    }
}

function Get-Field($Rec, [string[]]$Aliases) {
    foreach ($a in $Aliases) {
        $k = ConvertTo-Key $a
        if ($Rec.ContainsKey($k) -and "$($Rec[$k])".Trim() -ne '') { return "$($Rec[$k])".Trim() }
    }
    return $null
}

function Find-Column($Table, [string[]]$Aliases) {
    $keys = @($Table.Headers | ForEach-Object { ConvertTo-Key $_ })
    foreach ($a in $Aliases) { if ($keys -contains (ConvertTo-Key $a)) { return $a } }
    return $null
}

function Register-Input([string]$Role, [string]$Path, [int]$Rows) {
    $script:Inputs.Add([ordered]@{ role = $Role; file = (Get-RelPath $Path); sha256 = (Get-Sha256 $Path); rows = $Rows })
    Write-Log ("Loaded {0}: {1} ({2} data rows, SHA-256 {3})" -f $Role, (Get-RelPath $Path), $Rows, (Get-Sha256 $Path))
}

function Get-HospitalName([string]$Kitchen) {
    $k = ConvertTo-Key $Kitchen
    if ($k -match 'lemesos|limassol|λεμεσ|gnl') { return 'Limassol General Hospital' }
    if ($k -match 'troodos|τροοδ|gnt') { return 'Troodos Hospital' }
    return $null
}

# ---------------------------------------------------------------------------
# WordprocessingML builders
# ---------------------------------------------------------------------------
function ConvertTo-XmlText([string]$Text) {
    if ($null -eq $Text) { return '' }
    $clean = [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return [System.Security.SecurityElement]::Escape($clean)
}

function New-WRun([string]$Text, [switch]$Bold, [switch]$Italic, [int]$Size = 22) {
    $rpr = '<w:rPr>' + $(if ($Bold) { '<w:b/>' }) + $(if ($Italic) { '<w:i/>' }) + "<w:sz w:val=`"$Size`"/><w:szCs w:val=`"$Size`"/></w:rPr>"
    return "<w:r>$rpr<w:t xml:space=`"preserve`">$(ConvertTo-XmlText $Text)</w:t></w:r>"
}

# A paragraph. $Text may be a string, or an array of @{ Text; Bold; Italic } segments.
function New-WPara($Text, [switch]$Bold, [switch]$Italic, [int]$Size = 22, [int]$Indent = 0, [int]$After = 100, [switch]$KeepNext) {
    $ppr = '<w:pPr>' + $(if ($KeepNext) { '<w:keepNext/>' }) + "<w:spacing w:before=`"0`" w:after=`"$After`"/>" +
        $(if ($Indent -gt 0) { "<w:ind w:left=`"$Indent`" w:hanging=`"240`"/>" }) + '</w:pPr>'
    $runs = ''
    if ($Text -is [string]) { $runs = New-WRun $Text -Bold:$Bold -Italic:$Italic -Size $Size }
    else { foreach ($seg in $Text) { $runs += New-WRun $seg['Text'] -Bold:([bool]$seg['Bold']) -Italic:([bool]$seg['Italic']) -Size $Size } }
    return "<w:p>$ppr$runs</w:p>"
}

# A table. Cells are strings or @{ Text; Bold; Fill }. $Right lists right-aligned column indexes.
function New-WTable([string[]]$Headers, $Rows, [int[]]$Widths, [int[]]$Right = @(), [int]$Size = 18) {
    $total = ($Widths | Measure-Object -Sum).Sum
    $b = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
    $borders = '<w:tblBorders>' + (($('top', 'left', 'bottom', 'right', 'insideH', 'insideV') | ForEach-Object { $b -f $_ }) -join '') + '</w:tblBorders>'
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<w:tbl><w:tblPr><w:tblW w:w=`"$total`" w:type=`"dxa`"/>$borders<w:tblLayout w:type=`"fixed`"/>")
    [void]$sb.Append('<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>')
    foreach ($w in $Widths) { [void]$sb.Append("<w:gridCol w:w=`"$w`"/>") }
    [void]$sb.Append('</w:tblGrid>')
    $allRows = New-Object System.Collections.Generic.List[object]
    $allRows.Add([object[]]$Headers)
    foreach ($row in $Rows) { $allRows.Add($row) }
    for ($r = 0; $r -lt $allRows.Count; $r++) {
        $cells = @($allRows[$r])
        [void]$sb.Append('<w:tr>' + $(if ($r -eq 0) { '<w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>' } else { '<w:trPr><w:cantSplit/></w:trPr>' }))
        for ($c = 0; $c -lt $Widths.Count; $c++) {
            $cell = if ($c -lt $cells.Count) { $cells[$c] } else { '' }
            $text = ''; $bold = ($r -eq 0); $fill = if ($r -eq 0) { 'D9E2F3' } else { $null }
            if ($cell -is [hashtable]) {
                $text = [string]$cell.Text
                if ($cell.ContainsKey('Bold') -and $cell.Bold) { $bold = $true }
                if ($cell.ContainsKey('Fill') -and $cell.Fill) { $fill = $cell.Fill }
            } else { $text = [string]$cell }
            $shd = if ($fill) { "<w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$fill`"/>" } else { '' }
            $jc = if ($r -gt 0 -and $Right -contains $c) { '<w:jc w:val="right"/>' } else { '' }
            [void]$sb.Append("<w:tc><w:tcPr><w:tcW w:w=`"$($Widths[$c])`" w:type=`"dxa`"/>$shd</w:tcPr>")
            [void]$sb.Append("<w:p><w:pPr><w:spacing w:before=`"0`" w:after=`"0`"/>$jc</w:pPr>$(New-WRun $text -Bold:$bold -Size $Size)</w:p></w:tc>")
        }
        [void]$sb.Append('</w:tr>')
    }
    [void]$sb.Append('</w:tbl><w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr></w:p>')
    return $sb.ToString()
}

# Fills the template: block placeholders replace their whole paragraph with generated
# XML; inline placeholders are replaced in place with escaped text.
function Write-DocxFromTemplate([string]$Template, [string]$OutFile, [hashtable]$Inline, [hashtable]$Blocks) {
    Copy-Item -LiteralPath $Template -Destination $OutFile -Force
    $fs = [System.IO.File]::Open($OutFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $parts = @($zip.Entries | Where-Object { $_.FullName -match '^word/(document|header\d*|footer\d*)\.xml$' } | ForEach-Object { $_.FullName })
            $found = @{}
            foreach ($name in $parts) {
                $entry = $zip.GetEntry($name)
                $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
                $xml = $reader.ReadToEnd(); $reader.Dispose()
                foreach ($k in $Blocks.Keys) {
                    $pattern = '<w:p\b(?:(?!</w:p>).)*?\{\{' + [regex]::Escape($k) + '\}\}(?:(?!</w:p>).)*?</w:p>'
                    $m = [regex]::Match($xml, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
                    if ($m.Success) { $xml = $xml.Substring(0, $m.Index) + $Blocks[$k] + $xml.Substring($m.Index + $m.Length); $found[$k] = $true }
                }
                foreach ($k in $Inline.Keys) {
                    $ph = '{{' + $k + '}}'
                    if ($xml.Contains($ph)) { $xml = $xml.Replace($ph, (ConvertTo-XmlText $Inline[$k])); $found[$k] = $true }
                }
                $left = [regex]::Matches($xml, '\{\{[A-Za-z0-9_]+\}\}') | ForEach-Object { $_.Value }
                foreach ($l in @($left)) { Add-Issue "Template placeholder $l in $name has no value and was left as is." -Warning }
                $entry.Delete()
                $writer = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open(), $Utf8NoBom)
                $writer.Write($xml); $writer.Dispose()
            }
            foreach ($k in @($Blocks.Keys) + @($Inline.Keys)) {
                if (-not $found.ContainsKey($k)) { Add-Issue "Template has no {{$k}} placeholder; that content is only in the JSON summary." -Warning }
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

function Write-JsonFile([string]$Path, $Object) {
    [System.IO.File]::WriteAllText($Path, ($Object | ConvertTo-Json -Depth 12), $Utf8NoBom)
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$RootPath      = [System.IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
$DataDir       = Join-Path $RootPath 'data'
$KitchensDir   = Join-Path $DataDir 'kitchens'
$PriceListsDir = Join-Path $DataDir 'price_lists'
$LegalDir      = Join-Path $DataDir 'legal'
$SuppliersFile = Join-Path $DataDir 'suppliers.xlsx'
$CatalogFile   = Join-Path $DataDir 'items_catalog.xlsx'
$ConfigFile    = Join-Path $DataDir 'procurement_config.json'
$TemplateFile  = Join-Path (Join-Path $RootPath 'templates') 'unified_template.docx'
$LogDir        = Join-Path $RootPath 'logs'
if (-not $OutputDir) { $OutputDir = Join-Path (Join-Path $RootPath 'output') $RequestId }
$DocxOut = Join-Path $OutputDir 'unified_order.docx'
$JsonOut = Join-Path $OutputDir 'unified_summary.json'

New-Item -ItemType Directory -Force -Path $LogDir, $OutputDir | Out-Null
$script:LogFile = Join-Path $LogDir 'procurement.log'
$script:Inputs  = New-Object System.Collections.Generic.List[object]
$legalRefs      = New-Object System.Collections.Generic.List[object]
$exitCode       = 0

try {
    Write-Log "===== Unified procurement run started =====" 'STEP'
    Write-Log "Root: $RootPath | As-of date: $($AsOfDate.ToString('yyyy-MM-dd', $Inv)) | Host: PowerShell $($PSVersionTable.PSVersion)"
    Write-Log "PROCUREMENT_SYSTEM=$env:PROCUREMENT_SYSTEM LEGAL_FRAMEWORK=$env:LEGAL_FRAMEWORK"
    if ($env:LEGAL_FRAMEWORK -and $env:LEGAL_FRAMEWORK -ne 'ARTICLE_90') {
        Add-Issue "LEGAL_FRAMEWORK is '$env:LEGAL_FRAMEWORK'; this script only applies Article 90(1)(a)/(b)." -Warning
    }

    # --- Step 1: configuration -------------------------------------------------
    Write-Log 'Step 1/8: Loading configuration' 'STEP'
    if (Test-Path -LiteralPath $ConfigFile) {
        $cfg = Get-Content -LiteralPath $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
        Register-Input 'configuration' $ConfigFile 1
        if ($null -eq $Threshold90a -and $cfg.PSObject.Properties['threshold_90_1_a_eur'] -and $null -ne $cfg.threshold_90_1_a_eur) {
            $Threshold90a = ConvertTo-Decimal ([string]$cfg.threshold_90_1_a_eur)
        }
        if ($null -eq $Threshold90b -and $cfg.PSObject.Properties['threshold_90_1_b_max_eur'] -and $null -ne $cfg.threshold_90_1_b_max_eur) {
            $Threshold90b = ConvertTo-Decimal ([string]$cfg.threshold_90_1_b_max_eur)
        }
    }
    if ($null -eq $Threshold90a -or $Threshold90a -le 0) {
        Add-Issue ("Article 90(1)(a) threshold is not set. Add ""threshold_90_1_a_eur"" (EUR, excl. VAT) to " +
            "data/procurement_config.json or pass -Threshold90a. It is never assumed.")
    }
    if ($null -ne $Threshold90b -and $null -ne $Threshold90a -and $Threshold90b -le $Threshold90a) {
        Add-Issue "threshold_90_1_b_max_eur ($Threshold90b) must be greater than threshold_90_1_a_eur ($Threshold90a)."
    }
    Stop-IfErrors 'configuration'
    Write-Log ("Article 90(1)(a) threshold: {0} excl. VAT{1}" -f (Format-Money $Threshold90a),
        $(if ($null -ne $Threshold90b) { "; Article 90(1)(b) ceiling: $(Format-Money $Threshold90b)" } else { '' }))

    foreach ($req in @($SuppliersFile, $CatalogFile, $TemplateFile)) {
        if (-not (Test-Path -LiteralPath $req)) { Add-Issue "Required file not found: $(Get-RelPath $req)" }
    }
    if (-not (Test-Path -LiteralPath $KitchensDir)) { Add-Issue 'Required folder not found: data/kitchens' }
    Stop-IfErrors 'input discovery'

    # --- Step 2: suppliers -----------------------------------------------------
    Write-Log 'Step 2/8: Loading suppliers (data/suppliers.xlsx)' 'STEP'
    $t = Read-Table $SuppliersFile
    Register-Input 'suppliers' $SuppliersFile $t.Rows.Count
    if (-not (Find-Column $t $Cols.SupplierKey))  { Add-Issue "suppliers.xlsx: no supplier code column. Headers found: $($t.Headers -join ', ')" }
    if (-not (Find-Column $t $Cols.SupplierName)) { Add-Issue "suppliers.xlsx: no supplier name column. Headers found: $($t.Headers -join ', ')" }
    Stop-IfErrors 'suppliers'
    $Suppliers = [ordered]@{}
    $SupplierByName = @{}
    foreach ($r in $t.Rows) {
        $id = Get-Field $r $Cols.SupplierKey; $name = Get-Field $r $Cols.SupplierName
        if (-not $id -or -not $name) { Add-Issue "suppliers.xlsx row $($r.__row): missing supplier code or name."; continue }
        $key = $id.ToUpperInvariant()
        if ($Suppliers.Contains($key)) { Add-Issue "suppliers.xlsx row $($r.__row): duplicate supplier code '$id'."; continue }
        $activeRaw = Get-Field $r $Cols.Active
        $active = -not ($activeRaw -and (ConvertTo-Key $activeRaw) -match '^(no|false|0|inactive|n|οχι|ανενεργος|ανενεργη)$')
        $Suppliers[$key] = [ordered]@{
            Id = $id; Name = $name; Vat = Get-Field $r $Cols.Vat; Email = Get-Field $r $Cols.Email
            Phone = Get-Field $r $Cols.Phone; Address = Get-Field $r $Cols.Address; Active = $active
        }
        $SupplierByName[(ConvertTo-Key $name)] = $key
        if (-not $active) { Write-Log "Supplier $id ($name) is marked inactive; its offers will be excluded." 'WARN' }
    }
    if ($Suppliers.Count -eq 0) { Add-Issue 'suppliers.xlsx contains no suppliers.' }
    Stop-IfErrors 'suppliers'
    Write-Log "Suppliers registered: $($Suppliers.Count)"

    function Resolve-Supplier([string]$Value) {
        if (-not $Value) { return $null }
        $k = $Value.Trim().ToUpperInvariant()
        if ($Suppliers.Contains($k)) { return $k }
        $n = ConvertTo-Key $Value
        if ($SupplierByName.ContainsKey($n)) { return $SupplierByName[$n] }
        return $null
    }

    # --- Step 3: item catalog --------------------------------------------------
    Write-Log 'Step 3/8: Loading item catalog (data/items_catalog.xlsx)' 'STEP'
    $t = Read-Table $CatalogFile
    Register-Input 'items_catalog' $CatalogFile $t.Rows.Count
    if (-not (Find-Column $t $Cols.ItemCode)) { Add-Issue "items_catalog.xlsx: no item code column. Headers found: $($t.Headers -join ', ')" }
    Stop-IfErrors 'catalog'
    $Catalog = [ordered]@{}
    $Offers = New-Object System.Collections.Generic.List[object]
    $catalogHasPrices = (Find-Column $t $Cols.UnitPrice) -and (Find-Column $t $Cols.SupplierId)
    foreach ($r in $t.Rows) {
        $code = Get-Field $r $Cols.ItemCode
        if (-not $code) { Add-Issue "items_catalog.xlsx row $($r.__row): missing item code."; continue }
        $key = $code.ToUpperInvariant()
        if (-not $Catalog.Contains($key)) {
            $desc = Get-Field $r $Cols.Description
            if (-not $desc) { Add-Issue "items_catalog.xlsx row $($r.__row): item '$code' has no description." -Warning; $desc = '' }
            $Catalog[$key] = [ordered]@{ Code = $code; Description = $desc; Unit = (Get-Field $r $Cols.Unit) }
        } elseif (-not $catalogHasPrices) {
            Add-Issue "items_catalog.xlsx row $($r.__row): duplicate item code '$code'."
        }
        if ($catalogHasPrices) {
            $Offers.Add(@{ Rec = $r; Source = 'items_catalog.xlsx'; SupplierValue = (Get-Field $r $Cols.SupplierId); ItemCode = $code })
        }
    }
    Stop-IfErrors 'catalog'
    Write-Log "Catalog items: $($Catalog.Count)$(if ($catalogHasPrices) { ' (catalog also carries supplier prices)' })"

    # --- Step 4: price lists ---------------------------------------------------
    Write-Log 'Step 4/8: Loading price lists (data/price_lists/*)' 'STEP'
    $plFiles = @()
    if (Test-Path -LiteralPath $PriceListsDir) {
        $plFiles = @(Get-ChildItem -LiteralPath $PriceListsDir -File | Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '.*' } | Sort-Object Name)
    }
    foreach ($f in $plFiles) {
        if ($f.Extension -notmatch '^\.(xlsx|csv)$') { Add-Issue "Price list '$($f.Name)' skipped: only .xlsx and .csv are supported." -Warning; continue }
        $t = Read-Table $f.FullName
        Register-Input 'price_list' $f.FullName $t.Rows.Count
        $hasSupplierCol = [bool](Find-Column $t $Cols.SupplierId)
        $fileSupplier = $null
        if (-not $hasSupplierCol) {
            $fileSupplier = $f.BaseName -replace '(?i)[_\- ]?(price[_\- ]?list|prices|τιμοκαταλογος)$', ''
            if (-not (Resolve-Supplier $fileSupplier)) {
                Add-Issue "Price list '$($f.Name)': no supplier column, and the file name '$fileSupplier' matches no supplier in suppliers.xlsx."
                continue
            }
            Write-Log "Price list '$($f.Name)' has no supplier column; attributed to supplier '$fileSupplier' by file name."
        }
        if (-not (Find-Column $t $Cols.ItemCode))  { Add-Issue "Price list '$($f.Name)': no item code column. Headers found: $($t.Headers -join ', ')"; continue }
        if (-not (Find-Column $t $Cols.UnitPrice)) { Add-Issue "Price list '$($f.Name)': no unit price column. Headers found: $($t.Headers -join ', ')"; continue }
        foreach ($r in $t.Rows) {
            $sv = if ($hasSupplierCol) { Get-Field $r $Cols.SupplierId } else { $fileSupplier }
            $Offers.Add(@{ Rec = $r; Source = $f.Name; SupplierValue = $sv; ItemCode = (Get-Field $r $Cols.ItemCode) })
        }
    }
    if ($plFiles.Count -eq 0 -and -not $catalogHasPrices) { Add-Issue 'No price lists found in data/price_lists and the catalog has no supplier prices.' }
    Stop-IfErrors 'price lists'

    # Validate offers. Key: SUPPLIER|ITEM -> offer
    $ValidOffers = @{}
    $ignored = 0
    foreach ($o in $Offers) {
        $r = $o.Rec; $where = "$($o.Source) row $($r.__row)"
        if (-not $o.ItemCode) { Add-Issue "${where}: missing item code."; continue }
        $sKey = Resolve-Supplier $o.SupplierValue
        if (-not $sKey) { Add-Issue "${where}: supplier '$($o.SupplierValue)' is not registered in suppliers.xlsx."; continue }
        $iKey = $o.ItemCode.ToUpperInvariant()
        $priceRaw = Get-Field $r $Cols.UnitPrice
        $price = ConvertTo-Decimal $priceRaw
        if ($null -eq $price -or $price -le 0) { Add-Issue "${where}: invalid unit price '$priceRaw' for item '$($o.ItemCode)'."; continue }
        $cur = Get-Field $r $Cols.Currency
        if ($cur -and $cur.ToUpperInvariant() -notin @('EUR', $Euro)) { Add-Issue "${where}: currency '$cur' is not EUR."; continue }
        $from = ConvertTo-DateValue (Get-Field $r $Cols.ValidFrom)
        $to   = ConvertTo-DateValue (Get-Field $r $Cols.ValidTo)
        if ($from -eq 'INVALID' -or $to -eq 'INVALID') { Add-Issue "${where}: unreadable ValidFrom/ValidTo date."; continue }
        if (($null -ne $from -and $from -gt $AsOfDate) -or ($null -ne $to -and $to -lt $AsOfDate)) {
            Write-Log "${where}: offer from $($Suppliers[$sKey].Id) for '$($o.ItemCode)' not valid on $($AsOfDate.ToString('yyyy-MM-dd', $Inv)); ignored."
            $ignored++; continue
        }
        if (-not $Suppliers[$sKey].Active) { $ignored++; continue }
        if (-not $Catalog.Contains($iKey)) { Add-Issue "${where}: item '$($o.ItemCode)' is not in items_catalog.xlsx; offer ignored." -Warning; $ignored++; continue }
        $k = "$sKey|$iKey"
        if ($ValidOffers.ContainsKey($k)) {
            if ($ValidOffers[$k].UnitPrice -ne $price) {
                Add-Issue ("{0}: conflicting prices for supplier {1}, item {2}: {3} vs {4} ({5})." -f $where, $Suppliers[$sKey].Id, $o.ItemCode,
                    (Format-Price $price), (Format-Price $ValidOffers[$k].UnitPrice), $ValidOffers[$k].Source)
            }
            continue
        }
        $ValidOffers[$k] = @{ SupplierKey = $sKey; ItemKey = $iKey; UnitPrice = $price; Source = "$where" }
    }
    Stop-IfErrors 'price validation'
    Write-Log "Valid offers: $($ValidOffers.Count); offers ignored (expired, inactive supplier or unknown item): $ignored"

    # --- Step 5: kitchen requirements -----------------------------------------
    Write-Log 'Step 5/8: Loading kitchen requirements (data/kitchens/*.xlsx)' 'STEP'
    $kFiles = @(Get-ChildItem -LiteralPath $KitchensDir -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name)
    if ($kFiles.Count -eq 0) { Add-Issue 'No kitchen requirement files (*.xlsx) found in data/kitchens.' }
    Stop-IfErrors 'kitchens'
    $Demand = [ordered]@{}
    $Kitchens = New-Object System.Collections.Generic.List[object]
    foreach ($f in $kFiles) {
        $t = Read-Table $f.FullName
        Register-Input 'kitchen_requirements' $f.FullName $t.Rows.Count
        $kitchen = $f.BaseName -replace '(?i)[_\- ]?requirements?$', ''
        if (-not (Find-Column $t $Cols.ItemCode)) { Add-Issue "$($f.Name): no item code column. Headers found: $($t.Headers -join ', ')"; continue }
        if (-not (Find-Column $t $Cols.Quantity)) { Add-Issue "$($f.Name): no quantity column. Headers found: $($t.Headers -join ', ')"; continue }
        $lines = 0; $hospital = $null
        foreach ($r in $t.Rows) {
            $where = "$($f.Name) row $($r.__row)"
            $code = Get-Field $r $Cols.ItemCode
            $qRaw = Get-Field $r $Cols.Quantity
            if (-not $code -and -not $qRaw) { continue }
            if (-not $code) { Add-Issue "${where}: quantity given without an item code."; continue }
            $iKey = $code.ToUpperInvariant()
            if (-not $Catalog.Contains($iKey)) { Add-Issue "${where}: item '$code' is not in items_catalog.xlsx."; continue }
            $q = ConvertTo-Decimal $qRaw
            if ($null -eq $q -or $q -le 0) { Add-Issue "${where}: invalid quantity '$qRaw' for item '$code' (must be a number > 0)."; continue }
            $unit = Get-Field $r $Cols.Unit
            $catUnit = $Catalog[$iKey].Unit
            if ($unit -and $catUnit -and (ConvertTo-Key $unit) -ne (ConvertTo-Key $catUnit)) {
                Add-Issue "${where}: unit '$unit' for item '$code' differs from catalog unit '$catUnit'."; continue
            }
            $h = Get-Field $r $Cols.Hospital
            if ($h) { $hospital = $h }
            if (-not $Demand.Contains($iKey)) { $Demand[$iKey] = @{ Quantity = [decimal]0; Split = [ordered]@{} } }
            $Demand[$iKey].Quantity += $q
            if ($Demand[$iKey].Split.Contains($kitchen)) {
                Add-Issue "${where}: item '$code' listed more than once for kitchen $kitchen; quantities added together." -Warning
                $Demand[$iKey].Split[$kitchen] += $q
            } else { $Demand[$iKey].Split[$kitchen] = $q }
            $lines++
        }
        if (-not $hospital) { $hospital = Get-HospitalName $kitchen }
        if (-not $hospital) { Add-Issue "$($f.Name): hospital could not be determined from the file name; add a Hospital column." -Warning; $hospital = $kitchen }
        $Kitchens.Add([ordered]@{ kitchen = $kitchen; hospital = $hospital; file = (Get-RelPath $f.FullName); lines = $lines })
        Write-Log "Kitchen ${kitchen} ($hospital): $lines requirement line(s)"
    }
    if ($Demand.Count -eq 0) { Add-Issue 'Kitchen files contain no valid requirement lines.' }
    Stop-IfErrors 'kitchen requirements'

    # --- Step 6: legal references ---------------------------------------------
    Write-Log 'Step 6/8: Registering legal references (data/legal/*.pdf)' 'STEP'
    if (Test-Path -LiteralPath $LegalDir) {
        foreach ($p in @(Get-ChildItem -LiteralPath $LegalDir -Filter '*.pdf' -File | Sort-Object Name)) {
            $legalRefs.Add([ordered]@{ file = (Get-RelPath $p.FullName); sha256 = (Get-Sha256 $p.FullName); bytes = $p.Length })
            Write-Log "Legal reference: $(Get-RelPath $p.FullName) (SHA-256 $(Get-Sha256 $p.FullName))"
        }
    }
    if ($legalRefs.Count -eq 0) { Add-Issue 'No legal PDFs found in data/legal; the package cites the law without attached references.' -Warning }

    # --- Step 7: price matrix, selection and decision --------------------------
    Write-Log 'Step 7/8: Building item -> supplier -> price matrix and selecting suppliers' 'STEP'
    $Matrix = New-Object System.Collections.Generic.List[object]
    $quotingSuppliers = [ordered]@{}
    foreach ($iKey in $Demand.Keys) {
        $item = $Catalog[$iKey]
        $offers = @($ValidOffers.Values | Where-Object { $_.ItemKey -eq $iKey } | Sort-Object @{ e = { $_.UnitPrice } }, @{ e = { $Suppliers[$_.SupplierKey].Id } })
        if ($offers.Count -eq 0) { Add-Issue "Item $($item.Code) ($($item.Description)): no valid offer from any registered active supplier."; continue }
        foreach ($o in $offers) { $quotingSuppliers[$o.SupplierKey] = $true }
        $best = $offers[0]
        $ties = @($offers | Where-Object { $_.UnitPrice -eq $best.UnitPrice })
        $qty = $Demand[$iKey].Quantity
        $lineTotal = Round-Money ($qty * $best.UnitPrice)
        $s = $Suppliers[$best.SupplierKey]
        if ($offers.Count -eq 1) {
            $why = "only valid offer, from $($s.Name), at $(Format-Price $best.UnitPrice)"
        } elseif ($ties.Count -gt 1) {
            $why = "lowest price $(Format-Price $best.UnitPrice) offered by $($ties.Count) suppliers (" +
                (($ties | ForEach-Object { $Suppliers[$_.SupplierKey].Id }) -join ', ') + "); $($s.Id) selected by lowest supplier code. Confirm manually."
            Add-Issue "Item $($item.Code): price tie between $(($ties | ForEach-Object { $Suppliers[$_.SupplierKey].Id }) -join ', '); selected $($s.Id) by lowest supplier code." -Warning
        } else {
            $next = $offers[1]
            $pct = (($next.UnitPrice - $best.UnitPrice) / $best.UnitPrice * 100).ToString('0.0', $Inv)
            $why = "lowest of $($offers.Count) valid offers at $(Format-Price $best.UnitPrice); next lowest $($Suppliers[$next.SupplierKey].Name) at $(Format-Price $next.UnitPrice) (+$pct%)"
        }
        $Matrix.Add([ordered]@{
            ItemKey = $iKey; Code = $item.Code; Description = $item.Description; Unit = $item.Unit; Quantity = $qty
            Split = $Demand[$iKey].Split; Offers = $offers; SupplierKey = $best.SupplierKey; UnitPrice = $best.UnitPrice
            LineTotal = $lineTotal; Tie = ($ties.Count -gt 1); Justification = $why
        })
        Write-Log "Item $($item.Code): $($offers.Count) offer(s); selected $($s.Id) at $(Format-Price $best.UnitPrice) x $(Format-Qty $qty) = $(Format-Money $lineTotal)"
    }
    Stop-IfErrors 'price matrix'

    $GrandTotal = [decimal]0
    foreach ($m in $Matrix) { $GrandTotal += $m.LineTotal }
    $Selected = [ordered]@{}
    foreach ($m in $Matrix) {
        if (-not $Selected.Contains($m.SupplierKey)) { $Selected[$m.SupplierKey] = @{ Lines = New-Object System.Collections.Generic.List[object]; Total = [decimal]0 } }
        $Selected[$m.SupplierKey].Lines.Add($m)
        $Selected[$m.SupplierKey].Total += $m.LineTotal
    }

    if ($null -ne $Threshold90b -and $GrandTotal -gt $Threshold90b) {
        Add-Issue ("Total {0} exceeds the configured Article 90(1)(b) ceiling of {1}; this request cannot be processed under Article 90(1)." -f (Format-Money $GrandTotal), (Format-Money $Threshold90b))
    }
    Stop-IfErrors 'decision'
    $isA = ($GrandTotal -le $Threshold90a)
    $Article   = if ($isA) { '90(1)(a)' } else { '90(1)(b)' }
    $Procedure = if ($isA) { 'Simple acceptance of offer' } else { 'Summary award procedure' }
    Write-Log ("Decision: total {0} excl. VAT {1} threshold {2} -> Article {3} ({4})" -f (Format-Money $GrandTotal),
        $(if ($isA) { '<=' } else { '>' }), (Format-Money $Threshold90a), $Article, $Procedure) 'STEP'

    # --- Step 8: documents ---------------------------------------------------
    Write-Log 'Step 8/8: Generating unified_order.docx and unified_summary.json' 'STEP'
    $today = (Get-Date).ToString('dd/MM/yyyy', $Inv)
    $n = 0
    foreach ($sKey in @($Selected.Keys)) { $n++; $Selected[$sKey].OrderNo = "$RequestId-O$n"; $Selected[$sKey].AcceptNo = "$RequestId-A$n" }
    $W = 9026  # usable A4 width in twentieths of a point (1" margins)

    # Items ordered
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $Matrix) {
        $split = ($m.Split.Keys | ForEach-Object { "${_}: $(Format-Qty $m.Split[$_])" }) -join '; '
        $rows.Add(@($m.Code, $m.Description, [string]$m.Unit, (Format-Qty $m.Quantity), $split, $Suppliers[$m.SupplierKey].Id))
    }
    $itemsXml = New-WTable @('Code', 'Description', 'Unit', 'Total qty', 'Per kitchen', 'Supplier') $rows @(1000, 2700, 800, 1000, 2326, 1200) -Right @(3)

    # Price comparison matrix + justification
    $supCols = @($quotingSuppliers.Keys)
    $fixed = 900 + 600 + 750 + 1100 + 1050
    $supW = [Math]::Min(1100, [Math]::Floor(($W - $fixed - 1400) / [Math]::Max(1, $supCols.Count)))
    $descW = $W - $fixed - $supW * $supCols.Count
    $headers = @('Code', 'Description', 'Unit', 'Qty') + @($supCols | ForEach-Object { $Suppliers[$_].Id }) + @('Selected', 'Line total')
    $widths  = @(900, $descW, 600, 750) + @($supCols | ForEach-Object { [int]$supW }) + @(1100, 1050)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $Matrix) {
        $cells = @($m.Code, $m.Description, [string]$m.Unit, (Format-Qty $m.Quantity))
        foreach ($sKey in $supCols) {
            $o = @($m.Offers | Where-Object { $_.SupplierKey -eq $sKey })
            if ($o.Count -eq 0) { $cells += '-' }
            elseif ($sKey -eq $m.SupplierKey) { $cells += @{ Text = (Format-Price $o[0].UnitPrice); Bold = $true; Fill = 'E2EFDA' } }
            else { $cells += (Format-Price $o[0].UnitPrice) }
        }
        $cells += $Suppliers[$m.SupplierKey].Id
        $cells += (Format-Money $m.LineTotal)
        $rows.Add($cells)
    }
    $totalRow = @(@{ Text = 'TOTAL'; Bold = $true }) + @(1..(3 + $supCols.Count + 1) | ForEach-Object { '' }) + @(@{ Text = (Format-Money $GrandTotal); Bold = $true })
    $rows.Add($totalRow)
    $right = @(3) + @(0..($supCols.Count + 1) | ForEach-Object { 4 + $_ })
    $cmpXml = New-WTable $headers $rows $widths -Right $right -Size 16
    $cmpXml += New-WPara ('Supplier codes: ' + (($supCols | ForEach-Object { "$($Suppliers[$_].Id) = $($Suppliers[$_].Name)" }) -join '; ') + '. Unit prices in EUR excluding VAT; selected offer highlighted.') -Italic -Size 18 -After 200
    $cmpXml += New-WPara 'Supplier selection justification' -Bold -KeepNext
    $cmpXml += New-WPara ("Method: for each item, the lowest unit price (excluding VAT) among valid offers from active suppliers registered in suppliers.xlsx, " +
        "taken from the price lists in force on $($AsOfDate.ToString('dd/MM/yyyy', $Inv)). Equal prices are resolved by lowest supplier code and flagged for manual confirmation.")
    foreach ($m in $Matrix) { $cmpXml += New-WPara @(@{ Text = "- $($m.Code) $($m.Description): "; Bold = $true }, @{ Text = $m.Justification + '.' }) -Indent 360 -After 60 }

    # Legal basis and decision
    $legalXml = New-WPara "Article $Article of $LawText - $Procedure."
    $legalXml += New-WPara ("Article 90(1)(a): simple acceptance of offer (no award document). Article 90(1)(b): summary award procedure (formal award document).") -Size 20
    if ($legalRefs.Count -gt 0) {
        $legalXml += New-WPara 'Reference documents:' -Size 20 -After 40
        foreach ($l in $legalRefs) { $legalXml += New-WPara "- $($l.file) (SHA-256 $($l.sha256.Substring(0, 16))...)" -Size 20 -Indent 360 -After 40 }
    }
    $decisionXml = if ($isA) {
        New-WPara ("The total value of this request is $(Format-Money $GrandTotal) excluding VAT, which does not exceed the Article 90(1)(a) threshold of " +
            "$(Format-Money $Threshold90a). The procurement is carried out by simple acceptance of offer under Article 90(1)(a) of $LawText. No award document is issued.")
    } else {
        New-WPara ("The total value of this request is $(Format-Money $GrandTotal) excluding VAT, which exceeds the Article 90(1)(a) threshold of " +
            "$(Format-Money $Threshold90a)$(if ($null -ne $Threshold90b) { " and does not exceed the Article 90(1)(b) ceiling of $(Format-Money $Threshold90b)" }). " +
            "The procurement is carried out by the summary award procedure under Article 90(1)(b) of $LawText, and a formal award decision is issued below.")
    }

    # Acceptance of offer (a) or summary award decision (b)
    $awardXml = ''
    if ($isA) {
        foreach ($sKey in $Selected.Keys) {
            $s = $Suppliers[$sKey]; $sel = $Selected[$sKey]
            $awardXml += New-WPara "ACCEPTANCE OF OFFER No. $($sel.AcceptNo) - Article 90(1)(a)" -Bold -KeepNext
            $awardXml += New-WPara ("To: $($s.Name)" + $(if ($s.Vat) { " (VAT No. $($s.Vat))" }))
            $awardXml += New-WPara ("The $Authority accepts your offer for the $($sel.Lines.Count) item(s) listed in order $($sel.OrderNo), " +
                "for a total of $(Format-Money $sel.Total) excluding VAT, under Article 90(1)(a) of $LawText.") -After 200
        }
        $awardXml += New-WPara 'Under Article 90(1)(a) no award document is issued; this acceptance and the final order below constitute the procurement record.' -Italic -Size 20
    } else {
        $awardXml += New-WPara "SUMMARY AWARD DECISION No. $RequestId-AW - Article 90(1)(b)" -Bold -KeepNext
        $awardXml += New-WPara ("Having regard to Article 90(1)(b) of $LawText, the price comparison and the supplier selection justification above, " +
            "the $Authority awards the following contract(s) under the summary award procedure:")
        foreach ($sKey in $Selected.Keys) {
            $s = $Suppliers[$sKey]; $sel = $Selected[$sKey]
            $awardXml += New-WPara ("- $($s.Name)" + $(if ($s.Vat) { " (VAT No. $($s.Vat))" }) + ": $($sel.Lines.Count) item(s), $(Format-Money $sel.Total) excluding VAT - order $($sel.OrderNo)") -Indent 360 -After 60
        }
        $awardXml += New-WPara "Total awarded value: $(Format-Money $GrandTotal) excluding VAT." -Bold
        $awardXml += New-WPara 'This decision takes effect upon signature by the authorising officer.' -Italic -Size 20
    }

    # Final order(s)
    $orderXml = ''
    foreach ($sKey in $Selected.Keys) {
        $s = $Suppliers[$sKey]; $sel = $Selected[$sKey]
        $orderXml += New-WPara "ORDER No. $($sel.OrderNo) - $($s.Name)" -Bold -KeepNext
        $contact = @()
        if ($s.Vat) { $contact += "VAT No. $($s.Vat)" }
        if ($s.Address) { $contact += $s.Address }
        if ($s.Phone) { $contact += "Tel. $($s.Phone)" }
        if ($s.Email) { $contact += $s.Email }
        if ($contact.Count -gt 0) { $orderXml += New-WPara ($contact -join ' | ') -Size 20 }
        $orderXml += New-WPara "Legal basis: Article $Article of $LawText. Reference: $(if ($isA) { $sel.AcceptNo } else { "$RequestId-AW" })." -Size 20
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($m in $sel.Lines) {
            $split = ($m.Split.Keys | ForEach-Object { "${_}: $(Format-Qty $m.Split[$_])" }) -join '; '
            $rows.Add(@($m.Code, $m.Description, [string]$m.Unit, (Format-Qty $m.Quantity), $split, (Format-Price $m.UnitPrice), (Format-Money $m.LineTotal)))
        }
        $rows.Add(@(@{ Text = 'TOTAL'; Bold = $true }, '', '', '', '', '', @{ Text = (Format-Money $sel.Total); Bold = $true }))
        $orderXml += New-WTable @('Code', 'Description', 'Unit', 'Qty', 'Deliver to', 'Unit price', 'Total') $rows @(900, 2300, 700, 800, 1826, 1200, 1300) -Right @(3, 5, 6)
    }
    $orderXml += New-WPara ("Grand total: $(Format-Money $GrandTotal) excluding VAT. Prices exclude VAT. Deliver each quantity to the kitchen shown. " +
        "Package $RequestId; machine-readable summary: unified_summary.json.") -Size 20

    $kitchenNames = ($Kitchens | ForEach-Object { "$($_.kitchen) ($($_.hospital))" }) -join ', '
    $supplierNames = ($Selected.Keys | ForEach-Object { $Suppliers[$_].Name }) -join ', '
    Write-DocxFromTemplate $TemplateFile $DocxOut `
        @{ supplier_name = $supplierNames; date = $today; kitchen_name = $kitchenNames } `
        @{ items_table = $itemsXml; price_comparison = $cmpXml; legal_basis = $legalXml; decision_text = $decisionXml
           acceptance_or_award_text = $awardXml; final_order_text = $orderXml }
    Write-Log "Written $(Get-RelPath $DocxOut)"

    $summary = [ordered]@{
        request_id   = $RequestId
        status       = 'COMPLETED'
        generated_at = (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz', $Inv)
        as_of_date   = $AsOfDate.ToString('yyyy-MM-dd', $Inv)
        authority    = $Authority
        legal_framework = [ordered]@{
            law = $LawText; article = $Article; procedure = $Procedure
            award_document_issued = (-not $isA)
            threshold_90_1_a_eur = $Threshold90a; threshold_90_1_b_max_eur = $Threshold90b
        }
        decision = [ordered]@{
            article = $Article; procedure = $Procedure; total_excl_vat_eur = $GrandTotal
            rule = "total $(if ($isA) { '<=' } else { '>' }) Article 90(1)(a) threshold"
            reference = $(if ($isA) { @($Selected.Values | ForEach-Object { $_.AcceptNo }) } else { "$RequestId-AW" })
        }
        kitchens = $Kitchens.ToArray()
        price_matrix = @($Matrix | ForEach-Object {
            $m = $_
            [ordered]@{
                item_code = $m.Code; description = $m.Description; unit = $m.Unit; quantity = $m.Quantity
                per_kitchen = $m.Split
                offers = @($m.Offers | ForEach-Object { [ordered]@{ supplier_id = $Suppliers[$_.SupplierKey].Id; unit_price_eur = $_.UnitPrice; source = $_.Source } })
                selected_supplier_id = $Suppliers[$m.SupplierKey].Id; selected_unit_price_eur = $m.UnitPrice
                line_total_eur = $m.LineTotal; tie = $m.Tie; justification = $m.Justification
            }
        })
        orders = @($Selected.Keys | ForEach-Object {
            $s = $Suppliers[$_]; $sel = $Selected[$_]
            [ordered]@{
                order_number = $sel.OrderNo; acceptance_number = $(if ($isA) { $sel.AcceptNo } else { $null })
                supplier_id = $s.Id; supplier_name = $s.Name; vat_number = $s.Vat
                items = @($sel.Lines | ForEach-Object { $_.Code }); total_excl_vat_eur = $sel.Total
            }
        })
        totals = [ordered]@{ items = $Matrix.Count; suppliers_compared = $supCols.Count; suppliers_selected = $Selected.Count; grand_total_excl_vat_eur = $GrandTotal }
        inputs = $script:Inputs.ToArray()
        legal_references = $legalRefs.ToArray()
        outputs = [ordered]@{ document = (Get-RelPath $DocxOut); summary = (Get-RelPath $JsonOut); log = (Get-RelPath $script:LogFile) }
        warnings = $script:Warnings.ToArray()
        errors = @()
    }
    Write-JsonFile $JsonOut $summary
    Write-Log "Written $(Get-RelPath $JsonOut)"
    Write-Log ("===== Completed: Article {0}, {1} item(s), {2} supplier(s), total {3} excl. VAT, {4} warning(s) =====" -f $Article, $Matrix.Count, $Selected.Count, (Format-Money $GrandTotal), $script:Warnings.Count) 'STEP'
}
catch {
    $exitCode = if ($script:Errors.Count -gt 0) { 2 } else { 1 }
    if ($script:Errors.Count -eq 0) { $script:Errors.Add($_.Exception.Message) }
    Write-Log "Run FAILED: $($_.Exception.Message)" 'ERROR'
    if ($_.InvocationInfo) { Write-Log "At: $($_.InvocationInfo.PositionMessage -replace '\s+', ' ')" 'ERROR' }
    try {
        Write-JsonFile $JsonOut ([ordered]@{
            request_id = $RequestId; status = 'FAILED'; generated_at = (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz', $Inv)
            inputs = $script:Inputs.ToArray(); legal_references = $legalRefs.ToArray()
            warnings = $script:Warnings.ToArray(); errors = $script:Errors.ToArray()
            outputs = [ordered]@{ document = $null; summary = (Get-RelPath $JsonOut); log = (Get-RelPath $script:LogFile) }
        })
        Write-Log "Failure summary written to $(Get-RelPath $JsonOut); no order document was produced." 'ERROR'
    } catch { Write-Log "Could not write failure summary: $($_.Exception.Message)" 'ERROR' }
}
exit $exitCode
