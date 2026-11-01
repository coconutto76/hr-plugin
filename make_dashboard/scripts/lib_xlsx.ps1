# xlsx 읽기 도우미 (hr-risk/scripts/lib_xlsx.ps1 에서 복사하고 Grid 출력을 더함).
# Excel 없이 파일 구조(zip+XML)를 직접 읽는다. 압축을 디스크에 풀지 않고 메모리에서만 처리한다.

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntryText($Zip, [string]$Name) {
    $entry = $Zip.GetEntry($Name)
    if (-not $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-ColumnIndex([string]$CellRef) {
    # "AB12" -> 27 (0부터). 비어 있는 셀은 XML에서 빠질 수 있어서 위치를 열 문자로 계산한다.
    $n = 0
    foreach ($ch in ($CellRef -replace '\d', '').ToUpper().ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n - 1
}

function Read-XlsxWorkbook {
    # 반환: 시트 이름 -> @{ Headers = 첫 행의 열 이름; Rows = 둘째 행부터의 행(hashtable); Grid = 모든 행(값 배열) }
    # Grid는 표가 여러 개 들어 있거나 열 수가 들쭉날쭉한 시트(요약, 검증 등)를 읽을 때 쓴다.
    param([Parameter(Mandatory)][string]$Path)

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $shared = @()
        $sstText = Read-ZipEntryText $zip 'xl/sharedStrings.xml'
        if ($sstText) {
            $sst = [xml]$sstText
            foreach ($si in $sst.sst.si) { $shared += [string]$si.InnerText }
        }

        $wb = [xml](Read-ZipEntryText $zip 'xl/workbook.xml')
        $rels = [xml](Read-ZipEntryText $zip 'xl/_rels/workbook.xml.rels')
        $relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

        $result = [ordered]@{}
        foreach ($sheet in $wb.workbook.sheets.sheet) {
            $rid = $sheet.GetAttribute('id', $relNs)
            $target = ($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target
            $target = $target -replace '^/', ''
            if ($target -notmatch '^xl/') { $target = "xl/$target" }
            $doc = [xml](Read-ZipEntryText $zip $target)

            $headers = @()
            $rows = New-Object System.Collections.Generic.List[object]
            $grid = New-Object System.Collections.Generic.List[object]
            $first = $true
            foreach ($row in $doc.worksheet.sheetData.row) {
                $cells = @{}
                foreach ($c in $row.c) {
                    $idx = Get-ColumnIndex $c.r
                    $val = $null
                    switch ($c.t) {
                        's'         { $val = $shared[[int]$c.v] }
                        'inlineStr' { $val = [string]$c.is.InnerText }
                        'str'       { $val = [string]$c.v }
                        'b'         { $val = ($c.v -eq '1') }
                        'e'         { $val = $null }
                        default     {
                            if ($null -ne $c.v -and "$($c.v)" -ne '') {
                                $val = [double]::Parse([string]$c.v, [System.Globalization.CultureInfo]::InvariantCulture)
                            }
                        }
                    }
                    $cells[$idx] = $val
                }
                $max = if ($cells.Count) { ($cells.Keys | Measure-Object -Maximum).Maximum } else { -1 }
                $arr = New-Object object[] ($max + 1)
                for ($i = 0; $i -le $max; $i++) { $arr[$i] = $cells[$i] }
                $grid.Add($arr)
                if ($first) {
                    for ($i = 0; $i -le $max; $i++) { $headers += [string]$cells[$i] }
                    $first = $false
                } else {
                    $obj = [ordered]@{}
                    for ($i = 0; $i -lt $headers.Count; $i++) {
                        if ($headers[$i]) { $obj[$headers[$i]] = $cells[$i] }
                    }
                    $rows.Add($obj)
                }
            }
            $result[[string]$sheet.name] = @{ Headers = $headers; Rows = $rows; Grid = $grid }
        }
        return $result
    } finally { $zip.Dispose() }
}

function ConvertTo-DateValue($Value) {
    # 엑셀 일련번호(숫자) 또는 날짜 문자열을 [datetime]으로. 비었거나 해석 불가면 $null.
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    if ($Value -is [double]) { return ([datetime]'1899-12-30').AddDays($Value) }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$parsed)) { return $parsed }
    return $null
}
