# 결과 엑셀 쓰기 도우미. Excel COM을 사용하므로 Excel이 설치된 Windows에서만 동작한다.

function Start-ExcelApp {
    try { $xl = New-Object -ComObject Excel.Application }
    catch { throw "Excel을 실행할 수 없습니다. 결과 파일은 Excel이 설치된 Windows PC에서만 만들 수 있습니다. ($($_.Exception.Message))" }
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.SheetsInNewWorkbook = 1
    return $xl
}

function Stop-ExcelApp($Xl, $Workbook) {
    if ($Workbook) { try { $Workbook.Close($false) } catch {} ; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) }
    if ($Xl) { try { $Xl.Quit() } catch {} ; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Xl) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

function ConvertTo-SafeCell($Value) {
    # 문자열이 = + - @ 로 시작하면 Excel이 수식으로 해석하므로 앞에 ' 를 붙여 글자로 고정한다.
    # 파이프라인을 거친 값은 PSObject로 감싸져 있을 수 있고, 이 상태로 배열째 Excel에 넘기면 메모리 오류가 난다. 순수한 값으로 풀어 준다.
    if ($null -eq $Value) { return $null }
    $base = $Value.PSObject.BaseObject
    if ($null -ne $base) { $Value = $base }
    if ($Value -is [string]) {
        $text = [string]$Value
        if ($text -match '^[=+\-@]') { return "'" + $text }
        return $text
    }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return [double]$Value }
    if ($Value -is [bool]) { return $Value }
    return [string]$Value
}

function ConvertTo-Grid($Rows, [int]$Width) {
    $n = @($Rows).Count
    $grid = New-Object 'object[,]' $n, $Width
    for ($i = 0; $i -lt $n; $i++) {
        $row = @($Rows)[$i]
        for ($j = 0; $j -lt $Width; $j++) {
            $v = if ($j -lt @($row).Count) { @($row)[$j] } else { $null }
            $grid[$i, $j] = ConvertTo-SafeCell $v
        }
    }
    return , $grid
}

function New-Sheet($Workbook, [string]$Name) {
    if ($Workbook.Worksheets.Count -eq 1 -and -not $script:FirstSheetUsed) {
        $ws = $Workbook.Worksheets.Item(1)
        $script:FirstSheetUsed = $true
    } else {
        $ws = $Workbook.Worksheets.Add([Type]::Missing, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
    }
    $ws.Name = $Name
    return $ws
}

function Write-GridToSheet($Ws, $Rows, [int]$Width, [int]$StartRow = 1) {
    if (@($Rows).Count -eq 0) { return }
    if ($env:HR_RISK_DEBUG) { Write-Host ("DEBUG grid rows={0} width={1}" -f @($Rows).Count, $Width) }
    $grid = ConvertTo-Grid $Rows $Width
    $range = $Ws.Range($Ws.Cells($StartRow, 1), $Ws.Cells($StartRow + @($Rows).Count - 1, $Width))
    try { $range.Value2 = $grid } catch {
        if ($env:HR_RISK_DEBUG) { for ($i = 0; $i -lt @($Rows).Count; $i++) { for ($j = 0; $j -lt $Width; $j++) { try { $Ws.Cells($StartRow + $i, $j + 1).Value2 = $grid[$i, $j] } catch { Write-Host ("DEBUG bad cell r{0} c{1} len={2} head=[{3}]" -f $i, $j, ([string]$grid[$i, $j]).Length, ([string]$grid[$i, $j]).Substring(0, [math]::Min(40, ([string]$grid[$i, $j]).Length))) } } } }
        throw
    }
}

function Set-HeaderStyle($Xl, $Ws, [int]$Width, [int]$Row = 1, [bool]$Freeze = $true) {
    $hdr = $Ws.Range($Ws.Cells($Row, 1), $Ws.Cells($Row, $Width))
    $hdr.Font.Bold = $true
    $hdr.Interior.Color = 15921906   # 연한 회색
    $hdr.WrapText = $true
    $hdr.VerticalAlignment = -4108
    if ($Freeze) {
        try { $Ws.Activate(); $Xl.ActiveWindow.FreezePanes = $false; $Xl.ActiveWindow.SplitColumn = 0; $Xl.ActiveWindow.SplitRow = $Row; $Xl.ActiveWindow.FreezePanes = $true } catch {}
    }
}

function Set-ColumnWidths($Ws, [double[]]$Widths) {
    for ($i = 0; $i -lt $Widths.Count; $i++) { $Ws.Columns.Item($i + 1).ColumnWidth = $Widths[$i] }
}

function Set-GradeColors($Ws, [int]$Col, [int]$FirstRow, [int]$LastRow) {
    # 등급 칸 색: 상=붉은색, 중=노란색, 하=초록색, 판단불가=회색 (RGB를 BGR 정수로 변환한 값)
    $map = @{ '상' = 13421823; '중' = 10284031; '하' = 13561798; '판단불가' = 14277081 }
    for ($r = $FirstRow; $r -le $LastRow; $r++) {
        $cell = $Ws.Cells($r, $Col)
        $t = [string]$cell.Value2
        if ($map.ContainsKey($t)) { $cell.Interior.Color = $map[$t] }
    }
}

function Save-Workbook($Workbook, [string]$Path) {
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $Workbook.SaveAs($Path, 51)   # 51 = xlOpenXMLWorkbook(.xlsx)
}
