# make_dashboard 본체: hr-risk 결과 엑셀(+선택: 원본 인사데이터, ID매핑) -> dashboard.html, dashboard_share.html
# 사용: powershell -NoProfile -ExecutionPolicy Bypass -File build_dashboard.ps1 -ResultPath <퇴직위험_결과_YYYYMM.xlsx> [-SourcePath <인사데이터.xlsx>] [-MappingPath <ID매핑_YYYYMM.xlsx>] [-OutputDir <폴더>]
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ResultPath,
    [string]$SourcePath,
    [string]$MappingPath,
    [string]$OutputDir,
    [switch]$AllowSkillDir,
    [switch]$Open
)
$ErrorActionPreference = 'Stop'
$inv = [System.Globalization.CultureInfo]::InvariantCulture

# 스킬 폴더 기준 상대 경로: 어디에 설치돼도 이 스크립트 위치에서 보조 파일(assets, lib)을 찾는다.
$SkillRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'lib_xlsx.ps1')

function Stop-Build([string]$Message) {
    Write-Host "중단: $Message"
    Write-Host '(대시보드 파일은 만들지 않았습니다)'
    exit 2
}
function Resolve-InputPath([string]$Path, [string]$Label) {
    if (-not [System.IO.Path]::IsPathRooted($Path)) { $Path = Join-Path (Get-Location).Path $Path }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { Stop-Build "$Label 파일을 찾을 수 없습니다: $Path" }
    return (Resolve-Path -LiteralPath $Path).Path
}

# ---------------------------------------------------------------- 출력 폴더 (설치된 스킬 안에는 쓰지 않는다)
if (-not $OutputDir) { $OutputDir = (Get-Location).Path }
elseif (-not [System.IO.Path]::IsPathRooted($OutputDir)) { $OutputDir = Join-Path (Get-Location).Path $OutputDir }
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
$skillFull = [System.IO.Path]::GetFullPath($SkillRoot).TrimEnd('\') + '\'
if (($OutputDir.TrimEnd('\') + '\').StartsWith($skillFull, [System.StringComparison]::OrdinalIgnoreCase) -and -not $AllowSkillDir) {
    Stop-Build "출력 폴더가 스킬이 설치된 폴더 안입니다 ($OutputDir). 설치된 스킬 안에는 결과를 쓰지 않습니다. 작업 폴더를 -OutputDir 로 지정하세요."
}

$ResultPath = Resolve-InputPath $ResultPath '결과'
if ($SourcePath) { $SourcePath = Resolve-InputPath $SourcePath '원본' }
if ($MappingPath) { $MappingPath = Resolve-InputPath $MappingPath 'ID매핑' }

# ---------------------------------------------------------------- 공통 도우미
function Get-Key($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) { return [string][int64]$v }
    return ([string]$v).Trim()
}
function Mean($a) {
    $arr = @($a | Where-Object { $null -ne $_ })
    if ($arr.Count -eq 0) { return $null }
    return ($arr | Measure-Object -Average).Average
}
function Round-N($x, [int]$d = 2) { if ($null -eq $x) { return $null }; return [math]::Round([double]$x, $d) }

function ConvertTo-JsonString([string]$s) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    foreach ($ch in $s.ToCharArray()) {
        $c = [int]$ch
        if ($ch -ceq '"') { [void]$sb.Append('\"') }
        elseif ($ch -ceq '\') { [void]$sb.Append('\\') }
        elseif ($c -lt 32 -or $c -eq 60 -or $c -eq 62 -or $c -eq 38 -or $c -eq 0x2028 -or $c -eq 0x2029) { [void]$sb.Append(('\u{0:x4}' -f $c)) }
        else { [void]$sb.Append($ch) }
    }
    [void]$sb.Append('"')
    return $sb.ToString()
}
function ConvertTo-JsonText($o) {
    # HTML 안에 넣어도 </script> 등으로 깨지지 않도록 < > & 를 \uXXXX 로 바꾼다.
    if ($null -eq $o) { return 'null' }
    if ($o -is [bool]) { if ($o) { return 'true' } else { return 'false' } }
    if ($o -is [string]) { return (ConvertTo-JsonString $o) }
    if ($o -is [double] -or $o -is [single] -or $o -is [decimal]) {
        if ([double]::IsNaN([double]$o) -or [double]::IsInfinity([double]$o)) { return 'null' }
        return ([double]$o).ToString('R', $inv)
    }
    if ($o -is [int] -or $o -is [long] -or $o -is [int16] -or $o -is [byte]) { return ([long]$o).ToString($inv) }
    if ($o -is [System.Collections.IDictionary]) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($k in $o.Keys) { $parts.Add((ConvertTo-JsonString ([string]$k)) + ':' + (ConvertTo-JsonText $o[$k])) }
        return '{' + ($parts -join ',') + '}'
    }
    if ($o -is [System.Collections.IEnumerable]) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($x in $o) { $parts.Add((ConvertTo-JsonText $x)) }
        return '[' + ($parts -join ',') + ']'
    }
    return (ConvertTo-JsonString ([string]$o))
}

function Format-CellText($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) {
        if ($v -eq [math]::Floor($v)) { return ([long]$v).ToString($inv) }
        return ([math]::Round($v, 2)).ToString($inv)
    }
    return [string]$v
}
function Get-TrimmedRow($row) {
    $vals = @($row | ForEach-Object { Format-CellText $_ })
    $last = -1
    for ($i = 0; $i -lt $vals.Count; $i++) { if ($vals[$i] -ne '') { $last = $i } }
    if ($last -lt 0) { return @() }
    return @($vals[0..$last])
}
function ConvertTo-Blocks($grid) {
    # 검증 시트처럼 제목-표-제목-표가 이어진 시트를 {title, lines, tables} 덩어리로 바꾼다.
    $blocks = New-Object System.Collections.Generic.List[object]
    $cur = $null; $tbl = $null; $pct = @()
    foreach ($row in $grid) {
        $vals = @($row)
        $filled = @($vals | Where-Object { $null -ne $_ -and "$_" -ne '' })
        if ($filled.Count -eq 0) { $tbl = $null; continue }
        if ($filled.Count -eq 1 -and $vals.Count -ge 1 -and $null -ne $vals[0] -and "$($vals[0])" -ne '') {
            $text = [string]$vals[0]
            if ($text -match '^(검증 \d|분석의 한계)') {
                $cur = [ordered]@{ title = $text; lines = (New-Object System.Collections.Generic.List[object]); tables = (New-Object System.Collections.Generic.List[object]) }
                $blocks.Add($cur); $tbl = $null
            } elseif ($cur) { $cur.lines.Add($text); $tbl = $null }
            continue
        }
        if (-not $cur) { continue }
        if (-not $tbl) {
            $tbl = [ordered]@{ head = (Get-TrimmedRow $vals); rows = (New-Object System.Collections.Generic.List[object]) }
            $cur.tables.Add($tbl)
            $pct = @(); for ($i = 0; $i -lt @($tbl.head).Count; $i++) { if (@($tbl.head)[$i] -like '*비율*') { $pct += $i } }
        } else {
            $cells = @()
            for ($i = 0; $i -lt $vals.Count; $i++) {
                if (($pct -contains $i) -and $vals[$i] -is [double]) { $cells += ([double]$vals[$i]).ToString('P0', $inv) } else { $cells += (Format-CellText $vals[$i]) }
            }
            $last = -1; for ($i = 0; $i -lt $cells.Count; $i++) { if ($cells[$i] -ne '') { $last = $i } }
            $tbl.rows.Add(@($cells[0..$last]))
        }
    }
    return $blocks
}
function ConvertTo-SimpleTable($grid) {
    $head = @(); $rows = New-Object System.Collections.Generic.List[object]; $first = $true
    foreach ($row in $grid) {
        $cells = Get-TrimmedRow $row
        if ($cells.Count -eq 0) { continue }
        if ($first) { $head = $cells; $first = $false } else { $rows.Add($cells) }
    }
    return [ordered]@{ head = $head; rows = $rows }
}
function Get-SectionLines($grid, [string]$marker, [int]$max = 99) {
    $out = New-Object System.Collections.Generic.List[string]; $on = $false
    foreach ($row in $grid) {
        $v = if (@($row).Count) { Format-CellText @($row)[0] } else { '' }
        if (-not $on) { if ($v -eq $marker) { $on = $true }; continue }
        if ($v -eq '' -or $v -in @('요약 3줄', '분석 조건', '분석의 한계')) { break }
        $out.Add($v); if ($out.Count -ge $max) { break }
    }
    return $out
}

# ---------------------------------------------------------------- 1) 결과 파일 읽기
Write-Host "결과 읽는 중: $ResultPath"
$res = Read-XlsxWorkbook $ResultPath
foreach ($s in '요약', '개인결과', '조직결과') {
    if (-not $res.Contains($s)) { Stop-Build "결과 파일에 '$s' 시트가 없습니다. hr-risk가 만든 퇴직위험_결과_YYYYMM.xlsx인지 확인하세요." }
}
$title = [string]@($res['요약'].Grid[0])[0]
if ($title -notmatch '기준일\s*(\d{4}-\d{2}-\d{2})') { Stop-Build "결과 파일의 '요약' 시트에서 기준일을 읽지 못했습니다." }
$baseDate = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd', $inv)
$baseStr = $baseDate.ToString('yyyy-MM-dd'); $ym = $baseDate.ToString('yyyyMM')
$resultFile = Split-Path -Leaf $ResultPath
$notices = New-Object System.Collections.Generic.List[object]
if ($resultFile -match '(\d{6})\.xlsx$' -and $Matches[1] -ne $ym) {
    $notices.Add(@{ level = 'warn'; text = "결과 파일 이름의 월($($Matches[1]))과 시트의 기준일($baseStr)이 다릅니다. 기준일은 시트 값을 따랐습니다." })
}

$hdr = @($res['개인결과'].Headers)
foreach ($forbidden in '이름', '사번', 'name', 'employee_id') {
    if ($hdr -contains $forbidden) { Stop-Build "결과 파일의 개인결과에 '$forbidden' 열이 있습니다. 이름·사번이 들어 있는 파일로는 대시보드를 만들지 않습니다." }
}
foreach ($c in '분석용 ID', '등급', '본부', '조직', '직급', '판단 근거', '추가 확인 사항') {
    if ($hdr -notcontains $c) { Stop-Build "개인결과 시트에 필수 열 '$c'이(가) 없습니다." }
}
$scoreCol = $hdr | Where-Object { $_ -like '점수*' } | Select-Object -First 1
if (-not $scoreCol) { Stop-Build "개인결과 시트에서 점수 열을 찾지 못했습니다." }
$iw = [array]::IndexOf($hdr, '계산에 쓴 가중치 합'); $ina = [array]::IndexOf($hdr, '판단불가 신호')
$sigCols = @(); if ($iw -ge 0 -and $ina -gt $iw) { $sigCols = @($hdr[($iw + 1)..($ina - 1)]) }

$people = New-Object System.Collections.Generic.List[object]
$seen = @{}
foreach ($r in $res['개인결과'].Rows) {
    $id = [string]$r['분석용 ID']
    if ($id -notmatch '^R\d+$') { continue }
    if ($seen.ContainsKey($id)) { Stop-Build "개인결과에 분석용 ID가 중복됩니다: $id" }
    $seen[$id] = $true
    $sig = @(); foreach ($c in $sigCols) { $sig += , $(if ($r[$c] -is [double]) { Round-N $r[$c] 3 } else { $null }) }
    $na = [string]$r['판단불가 신호']
    $people.Add([ordered]@{
        id = $id; grade = [string]$r['등급']
        score = $(if ($r[$scoreCol] -is [double]) { Round-N $r[$scoreCol] 1 } else { $null })
        div = [string]$r['본부']; org = [string]$r['조직']; job = [string]$r['직군']; level = $r['직급']
        tenure = $(if ($r['근속(년)'] -is [double]) { Round-N $r['근속(년)'] 1 } else { $null })
        availW = $r['계산에 쓴 가중치 합']; sig = $sig; na = $na
        reason = [string]$r['판단 근거']; action = [string]$r['추가 확인 사항']; ref = [string]$r['참고']; m = $null
    })
}
if ($people.Count -eq 0) { Stop-Build "개인결과 시트에서 분석용 ID(R001 형식) 행을 찾지 못했습니다." }
$gradeNames = @('상', '중', '하', '판단불가')
foreach ($p in $people) { if ($gradeNames -notcontains $p.grade) { Stop-Build "알 수 없는 등급 값입니다: '$($p.grade)' ($($p.id))" } }
$total = $people.Count
$gradeCounts = [ordered]@{}; foreach ($g in $gradeNames) { $gradeCounts[$g] = @($people | Where-Object { $_.grade -eq $g }).Count }

# 개인결과와 조직결과가 서로 맞는지 확인
$orgMismatch = 0
foreach ($r in $res['조직결과'].Rows) {
    if ($r['조직'] -eq '(본부 전체)' -or $r['인원'] -isnot [double]) { continue }
    $mem = @($people | Where-Object { $_.div -eq $r['본부'] -and $_.org -eq $r['조직'] })
    if ($mem.Count -ne [int]$r['인원']) { $orgMismatch++; continue }
    if ($r['상'] -is [double] -and @($mem | Where-Object { $_.grade -eq '상' }).Count -ne [int]$r['상']) { $orgMismatch++ }
}
if ($orgMismatch) { $notices.Add(@{ level = 'warn'; text = "결과 파일의 개인결과와 조직결과 시트가 서로 맞지 않는 조직이 $orgMismatch 곳 있습니다. 개인결과 값을 기준으로 집계했습니다." }) }

# 요약·상세 설명 자료
$summaryLines = @(Get-SectionLines $res['요약'].Grid '요약 3줄' 3)
$conditions = @(Get-SectionLines $res['요약'].Grid '분석 조건')
$limits = @(Get-SectionLines $res['요약'].Grid '분석의 한계')
$validation = @(); if ($res.Contains('검증')) { $validation = @(ConvertTo-Blocks $res['검증'].Grid) }
$criteria = $null; if ($res.Contains('산정기준')) { $criteria = ConvertTo-SimpleTable $res['산정기준'].Grid }
$dataChecks = $null; if ($res.Contains('데이터점검')) { $dataChecks = ConvertTo-SimpleTable $res['데이터점검'].Grid }
$hypMap = @{
    '초과근무 증가' = '최근 3개월 초과근무가 직전 3개월보다 늘면 업무 부담이 커졌다는 신호일 수 있다.'
    '평가 변화'     = '평가가 하락하거나 낮은 등급에서 정체되면 이탈 위험 신호일 수 있다.'
    '근속 구간'     = '근속 1~3년처럼 특정 구간에서 이동 가능성이 커질 수 있다.'
    '연차 소진 낮음' = '연차를 거의 쓰지 못하면 업무 과중의 보조 신호일 수 있다.'
    '승진 보류'     = '최근 승진 심사가 보류되면 성장 기회에 대한 불만 신호일 수 있다.'
    '급여 밴드 하단' = '급여가 밴드 하단에 머물면 보상 수준에 대한 불만 신호일 수 있다.'
}
$hypotheses = @(); foreach ($c in $sigCols) { if ($hypMap.ContainsKey($c)) { $hypotheses += , @($c, $hypMap[$c]) } else { $hypotheses += , @($c, '') } }
$detail = [ordered]@{ hypotheses = $hypotheses; validation = $validation; criteria = $criteria; dataChecks = $dataChecks; limits = $limits; conditions = $conditions }

# ---------------------------------------------------------------- 2) 원본 읽기 (선택)
$hasSource = $false; $linked = $false
$srcMetrics = $null; $company = $null; $srcGroups = $null; $monthly = $null
$reasons = [ordered]@{}
$noSource = '원본 인사데이터를 함께 넣지 않아 표시할 수 없습니다. 같은 기준일의 원본(-SourcePath)을 넣으면 표시됩니다.'
$reasons.metrics = $noSource; $reasons.orgMetrics = $noSource; $reasons.monthly = $noSource

$ratingOrder = @('D', 'C', 'B', 'A', 'S')
if ($SourcePath) {
    Write-Host "원본 읽는 중: $SourcePath"
    $src = Read-XlsxWorkbook $SourcePath
    foreach ($s in 'v_employees_current', 'v_attendance_monthly', 'hr_performance_reviews') {
        if (-not $src.Contains($s)) { Stop-Build "원본에 '$s' 시트가 없습니다." }
    }
    $emp = $src['v_employees_current'].Rows
    $asOf = @($emp | ForEach-Object { ConvertTo-DateValue $_['as_of'] } | Where-Object { $_ } | ForEach-Object { $_.Date } | Sort-Object -Unique)
    if ($asOf.Count -ne 1) { Stop-Build "원본의 as_of가 한 날짜가 아니어서 기준일을 확인할 수 없습니다." }
    if ($asOf[0] -ne $baseDate) { Stop-Build "기준일이 다릅니다. 결과 파일 $baseStr / 원본 $($asOf[0].ToString('yyyy-MM-dd')). 같은 기준일의 원본을 넣어 주세요." }
    if ($src.Contains('개요')) {
        foreach ($r in $src['개요'].Rows) {
            if ($r['항목'] -eq '기준일') {
                $ob = ConvertTo-DateValue $r['내용']
                if ($ob -and $ob.Date -ne $baseDate) { Stop-Build "기준일이 다릅니다. 결과 파일 $baseStr / 원본 개요 시트 $($ob.ToString('yyyy-MM-dd'))." }
            }
        }
    }
    $hasSource = $true

    $lastFull = if ($baseDate.AddDays(1).Day -eq 1) { [datetime]::new($baseDate.Year, $baseDate.Month, 1) } else { [datetime]::new($baseDate.Year, $baseDate.Month, 1).AddMonths(-1) }
    $lastFullKey = $lastFull.ToString('yyyy-MM')
    $recentKeys = 0..2 | ForEach-Object { $lastFull.AddMonths(-$_).ToString('yyyy-MM') }

    # 대상: 기준일 이전 입사자 (기준일 이후 입사자는 hr-risk와 같이 제외)
    $pop = New-Object System.Collections.Generic.List[object]; $popKeys = @{}
    foreach ($e in $emp) {
        $h = ConvertTo-DateValue $e['hire_date']
        if ($h -and $h.Date -gt $baseDate) { continue }
        $pop.Add($e); $popKeys[(Get-Key $e['employee_id'])] = $true
    }

    # 근태: 기준일 이후 월은 쓰지 않는다
    $att = @{}; $attAfter = 0; $monthSet = @{}
    foreach ($r in $src['v_attendance_monthly'].Rows) {
        $k = Get-Key $r['employee_id']; if (-not $popKeys.ContainsKey($k)) { continue }
        $m = $null; $pm = $r['pay_month']
        if ($pm -is [string] -and $pm -match '^(\d{4})[-/.](\d{1,2})') { $m = ('{0}-{1:00}' -f $Matches[1], [int]$Matches[2]) }
        else { $d = ConvertTo-DateValue $pm; if ($d) { $m = $d.ToString('yyyy-MM') } }
        if (-not $m) { continue }
        if ($m -gt $lastFullKey) { $attAfter++; continue }
        $ot = $r['overtime_hours']
        if ($ot -isnot [double] -or $ot -lt 0 -or $ot -gt 200) { continue }
        if (-not $att.ContainsKey($k)) { $att[$k] = @{} }
        $att[$k][$m] = $ot; $monthSet[$m] = $true
    }
    # 평가: 기준일 이후 확정분·미확정 제외
    $rev = @{}; $revAfter = 0
    foreach ($r in $src['hr_performance_reviews'].Rows) {
        $k = Get-Key $r['employee_id']; if (-not $popKeys.ContainsKey($k)) { continue }
        if ($r['status'] -ne '확정') { continue }
        $fd = ConvertTo-DateValue $r['finalized_at']; if (-not $fd) { continue }
        if ($fd.Date -gt $baseDate) { $revAfter++; continue }
        $ri = [array]::IndexOf($ratingOrder, [string]$r['final_rating']); if ($ri -lt 0) { continue }
        if (-not $rev.ContainsKey($k)) { $rev[$k] = New-Object System.Collections.Generic.List[object] }
        $rev[$k].Add(@{ At = $fd; R = $ri + 1 })
    }
    # 인상률: 기준일 이전 최근 급여 변동률
    $raise = @{}
    if ($src.Contains('hr_compensation_history')) {
        $best = @{}
        foreach ($r in $src['hr_compensation_history'].Rows) {
            $k = Get-Key $r['employee_id']; if (-not $popKeys.ContainsKey($k)) { continue }
            $from = ConvertTo-DateValue $r['effective_from']; if (-not $from -or $from.Date -gt $baseDate) { continue }
            if ($r['change_pct'] -isnot [double]) { continue }
            if (-not $best.ContainsKey($k) -or $best[$k].From -lt $from) { $best[$k] = @{ From = $from; Pct = $r['change_pct'] } }
        }
        foreach ($k in $best.Keys) { $raise[$k] = $best[$k].Pct }
    }

    $srcMetrics = @{}   # 사번 -> 지표. 사번은 메모리에서만 쓰고 출력하지 않는다.
    foreach ($e in $pop) {
        $k = Get-Key $e['employee_id']
        $vals = @(); if ($att.ContainsKey($k)) { $vals = @($recentKeys | Where-Object { $att[$k].ContainsKey($_) } | ForEach-Object { $att[$k][$_] }) }
        $ot3 = if ($vals.Count -ge 2) { Mean $vals } else { $null }
        $rn = $null; $rp = $null
        if ($rev.ContainsKey($k)) {
            $sorted = @($rev[$k] | Sort-Object { $_.At })
            $rn = $sorted[-1].R; if ($sorted.Count -ge 2) { $rp = $sorted[-2].R }
        }
        $srcMetrics[$k] = @{
            div = [string]$e['division_name']; org = [string]$e['org_unit_name']; job = [string]$e['job_family']; level = $e['grade_level']
            ot3 = $ot3; rn = $rn; rp = $rp; raise = $(if ($raise.ContainsKey($k)) { $raise[$k] } else { $null })
        }
    }
    $allM = @($srcMetrics.Values)
    $company = [ordered]@{
        ot3 = Round-N (Mean ($allM | ForEach-Object { $_.ot3 })) 2
        rating = Round-N (Mean ($allM | ForEach-Object { $_.rn })) 2
        ratingDelta = Round-N (Mean ($allM | Where-Object { $null -ne $_.rn -and $null -ne $_.rp } | ForEach-Object { $_.rn - $_.rp })) 2
        raise = Round-N (Mean ($allM | ForEach-Object { $_.raise })) 2
        n = $allM.Count
    }

    # 부서 단위 원본 집계 (직원 연결이 없어도 계산 가능)
    $minGroup = 5
    $srcGroups = New-Object System.Collections.Generic.List[object]
    $divNames = @($allM | ForEach-Object { $_.div } | Sort-Object -Unique)
    foreach ($dv in $divNames) {
        $inDiv = @($allM | Where-Object { $_.div -eq $dv })
        $srcGroups.Add([ordered]@{ div = $dv; org = $null; n = $inDiv.Count
            ot3 = Round-N (Mean ($inDiv | ForEach-Object { $_.ot3 })) 2; rn = Round-N (Mean ($inDiv | ForEach-Object { $_.rn })) 2; raise = Round-N (Mean ($inDiv | ForEach-Object { $_.raise })) 2 })
        foreach ($og in @($inDiv | ForEach-Object { $_.org } | Sort-Object -Unique)) {
            $inOrg = @($inDiv | Where-Object { $_.org -eq $og })
            $srcGroups.Add([ordered]@{ div = $dv; org = $og; n = $inOrg.Count
                ot3 = Round-N (Mean ($inOrg | ForEach-Object { $_.ot3 })) 2; rn = Round-N (Mean ($inOrg | ForEach-Object { $_.rn })) 2; raise = Round-N (Mean ($inOrg | ForEach-Object { $_.raise })) 2 })
        }
    }

    # 월간 변화: 기준일이 속한 달(또는 직전 달)까지의 월별 1인당 평균 초과근무. 월 인원이 5명 미만이면 값을 비운다.
    $months = @($monthSet.Keys | Sort-Object)
    $groupMembers = @{}   # 그룹키 -> 사번 목록
    $groupMembers['company'] = @($srcMetrics.Keys)
    foreach ($k in $srcMetrics.Keys) {
        $a = 'div:' + $srcMetrics[$k].div; $b = 'org:' + $srcMetrics[$k].div + '|' + $srcMetrics[$k].org
        foreach ($gk in $a, $b) { if (-not $groupMembers.ContainsKey($gk)) { $groupMembers[$gk] = New-Object System.Collections.Generic.List[string] }; $groupMembers[$gk].Add($k) }
    }
    $series = [ordered]@{}
    foreach ($gk in $groupMembers.Keys) {
        $line = @()
        foreach ($mo in $months) {
            $v = @($groupMembers[$gk] | Where-Object { $att.ContainsKey($_) -and $att[$_].ContainsKey($mo) } | ForEach-Object { $att[$_][$mo] })
            $line += , $(if ($v.Count -ge $minGroup) { Round-N (Mean $v) 2 } else { $null })
        }
        $series[$gk] = @{ size = @($groupMembers[$gk]).Count; values = $line }
    }
    $monthly = [ordered]@{ months = $months; series = $series; unit = '시간/월(1인당 평균)' }
    $reasons.orgMetrics = $null; $reasons.monthly = $null
    $reasons.metrics = '원본은 있지만 분석용 ID와 직원을 연결하지 않았습니다. ID매핑 파일(-MappingPath)을 함께 넣으면 직원별로 표시됩니다.'
    Write-Host "원본 기준일 확인: $baseStr 일치 / 기준일 이후 제외 - 근태 $attAfter 행, 평가 $revAfter 건"

    # 직원 연결 (분석용 ID 기준)
    if ($MappingPath) {
        $map = Read-XlsxWorkbook $MappingPath
        $mapName = Split-Path -Leaf $MappingPath
        $why = $null
        if (-not $map.Contains('매핑') -or @($map['매핑'].Headers) -notcontains '분석용 ID' -or @($map['매핑'].Headers) -notcontains '사번') {
            $why = "ID매핑 파일에서 '매핑' 시트의 '분석용 ID'·'사번' 열을 찾지 못했습니다."
        } elseif ($mapName -match '(\d{6})\.xlsx$' -and $Matches[1] -ne $ym) {
            $why = "ID매핑 파일의 기준 월($($Matches[1]))이 결과 파일($ym)과 다릅니다."
        } else {
            $idToKey = @{}; $dup = $false
            foreach ($r in $map['매핑'].Rows) { $i = [string]$r['분석용 ID']; if ($i -notmatch '^R\d+$') { continue }; if ($idToKey.ContainsKey($i)) { $dup = $true }; $idToKey[$i] = Get-Key $r['사번'] }
            if ($dup) { $why = 'ID매핑 파일에 분석용 ID가 중복됩니다.' }
            elseif ($idToKey.Count -ne $people.Count) { $why = "ID매핑 파일의 행 수($($idToKey.Count))가 결과의 직원 수($($people.Count))와 다릅니다." }
            else {
                $bad = 0
                foreach ($p in $people) {
                    if (-not $idToKey.ContainsKey($p.id)) { $bad++; continue }
                    $s = $srcMetrics[$idToKey[$p.id]]
                    if (-not $s -or $s.div -ne $p.div -or $s.org -ne $p.org -or $s.job -ne $p.job -or "$($s.level)" -ne "$($p.level)") { $bad++ }
                }
                if ($bad) { $why = "결과와 원본의 소속·직군·직급이 어긋나는 분석용 ID가 $bad 건 있어 같은 직원임을 확인할 수 없습니다." }
                else {
                    foreach ($p in $people) {
                        $s = $srcMetrics[$idToKey[$p.id]]
                        $p.m = @((Round-N $s.ot3 2), $s.rn, $s.rp, (Round-N $s.raise 2))
                    }
                    $linked = $true
                }
            }
        }
        if ($linked) { $reasons.metrics = $null; Write-Host "직원 연결: 분석용 ID $($people.Count)건 모두 소속·직군·직급 대조 일치" }
        else { $reasons.metrics = "$why 그래서 직원별로 연결하지 않았습니다."; $notices.Add(@{ level = 'warn'; text = $reasons.metrics }); Write-Host "직원 연결 안 함: $why" }
        $map = $null; $idToKey = $null
    }
    $srcMetrics = $null   # 사번이 든 자료는 여기서 버린다
} elseif ($MappingPath) {
    $notices.Add(@{ level = 'info'; text = 'ID매핑 파일은 원본 인사데이터와 함께 있을 때만 쓰므로 사용하지 않았습니다.' })
}

# ---------------------------------------------------------------- 3) 조직별 집계 (공유용) + 5명 미만 보호
$minGroup = 5
function New-GroupRow($dv, $og, $members) {
    $n = @($members).Count
    $scored = @($members | Where-Object { $null -ne $_.score })
    $sigAvg = @()
    for ($i = 0; $i -lt $sigCols.Count; $i++) {
        $vals = @($members | ForEach-Object { @($_.sig)[$i] } | Where-Object { $null -ne $_ })
        $sigAvg += , $(if ($vals.Count) { Round-N (Mean $vals) 3 } else { $null })
    }
    $sg = $null; if ($srcGroups) { $sg = $srcGroups | Where-Object { $_.div -eq $dv -and $_.org -eq $og } | Select-Object -First 1 }
    return [ordered]@{
        div = $dv; org = $og; n = $n; suppressed = $false; protect = $false
        avgScore = $(if ($scored.Count) { Round-N (Mean ($scored | ForEach-Object { $_.score })) 1 } else { $null })
        counts = @($gradeNames | ForEach-Object { $g = $_; @($members | Where-Object { $_.grade -eq $g }).Count })
        sig = $sigAvg
        src = $(if ($sg) { [ordered]@{ ot3 = $sg.ot3; rn = $sg.rn; raise = $sg.raise } } else { $null })
    }
}
function Hide-GroupRow($row, [bool]$protect) {
    $row.suppressed = $true; $row.protect = $protect
    $row.n = $null; $row.avgScore = $null; $row.counts = $null; $row.sig = $null; $row.src = $null
}
$shareOrgs = New-Object System.Collections.Generic.List[object]
foreach ($dv in @($people | ForEach-Object { $_.div } | Sort-Object -Unique)) {
    $inDiv = @($people | Where-Object { $_.div -eq $dv })
    $divRow = New-GroupRow $dv $null $inDiv
    if ($divRow.n -lt $minGroup) { Hide-GroupRow $divRow $false }
    $shareOrgs.Add($divRow)
    $rows = @(); foreach ($og in @($inDiv | ForEach-Object { $_.org } | Sort-Object -Unique)) { $rows += , (New-GroupRow $dv $og @($inDiv | Where-Object { $_.org -eq $og })) }
    # 5명 미만 조직은 값을 뺀다. 빼고 남은 합계에서 역산되지 않도록, 가려진 인원이 5명 미만이면 가장 작은 조직을 더 가린다.
    $hiddenN = 0; foreach ($r in $rows) { if ($r.n -lt $minGroup) { $hiddenN += $r.n } }
    $anyHidden = @($rows | Where-Object { $_.n -lt $minGroup }).Count -gt 0
    $hiddenKeys = @{}; foreach ($r in $rows) { if ($r.n -lt $minGroup) { $hiddenKeys[$r.org] = $false } }
    if ($anyHidden) {
        foreach ($r in @($rows | Where-Object { -not $hiddenKeys.ContainsKey($_.org) } | Sort-Object { $_.n })) {
            if ($hiddenN -ge $minGroup) { break }
            $hiddenKeys[$r.org] = $true; $hiddenN += $r.n
        }
    }
    foreach ($r in $rows) { if ($hiddenKeys.ContainsKey($r.org)) { Hide-GroupRow $r ([bool]$hiddenKeys[$r.org]) }; $shareOrgs.Add($r) }
}
$hiddenOrgNames = @{}; foreach ($r in $shareOrgs) { if ($r.suppressed -and $r.org) { $hiddenOrgNames["org:$($r.div)|$($r.org)"] = $true } }

# ---------------------------------------------------------------- 4) 화면 데이터 만들기
$meta = [ordered]@{
    baseDate = $baseStr; ym = $ym; resultFile = $resultFile; sourceFile = $(if ($hasSource) { Split-Path -Leaf $SourcePath } else { $null })
    hasSource = $hasSource; linked = $linked; minGroup = $minGroup; notices = $notices
}
$signals = @($sigCols)
$summary = [ordered]@{ total = $total; counts = $gradeCounts; lines = $summaryLines }

$internal = [ordered]@{
    meta = $meta; summary = $summary; signals = $signals; people = $people; company = $company; srcGroups = $srcGroups
    monthly = $monthly; reasons = $reasons; detail = $detail
}

# 공유용: 개인 행·ID가 없는 집계만 넣는다. 5명 미만 조직(및 역산 방지 조직)의 월간 계열도 뺀다.
$shareMonthly = $null
if ($monthly) {
    $sSeries = [ordered]@{}
    foreach ($k in $monthly.series.Keys) { if (-not $hiddenOrgNames.ContainsKey($k) -and -not ($k -like 'org:*' -and $monthly.series[$k].size -lt $minGroup)) { $sSeries[$k] = $monthly.series[$k] } }
    $shareMonthly = [ordered]@{ months = $monthly.months; series = $sSeries; unit = $monthly.unit }
}
$shareMeta = [ordered]@{ baseDate = $baseStr; ym = $ym; resultFile = $resultFile; hasSource = $hasSource; minGroup = $minGroup; notices = @($notices | Where-Object { $_.text -notlike '*분석용 ID*' }) }
$shareReasons = [ordered]@{ orgMetrics = $reasons.orgMetrics; monthly = $reasons.monthly }
$share = [ordered]@{
    meta = $shareMeta; summary = $summary; signals = $signals; orgs = $shareOrgs; company = $company
    monthly = $shareMonthly; reasons = $shareReasons; detail = $detail
}

# ---------------------------------------------------------------- 5) HTML 만들기
$assets = Join-Path $SkillRoot 'assets'
function Read-Asset([string]$name) {
    $p = Join-Path $assets $name
    if (-not (Test-Path -LiteralPath $p)) { Stop-Build "스킬 파일이 없습니다: $p (스킬 폴더가 온전한지 확인하세요)" }
    return [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8)
}
$css = Read-Asset 'common.css'; $commonJs = Read-Asset 'common.js'
function Build-Html([string]$templateName, [string]$appName, $data, [string]$pageTitle) {
    $html = Read-Asset $templateName
    $html = $html.Replace('{{TITLE}}', $pageTitle).Replace('{{CSS}}', $css).Replace('{{DATA_JSON}}', (ConvertTo-JsonText $data))
    $html = $html.Replace('{{COMMON_JS}}', $commonJs).Replace('{{APP_JS}}', (Read-Asset $appName))
    return $html
}
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding $false
$dashPath = Join-Path $OutputDir 'dashboard.html'; $sharePath = Join-Path $OutputDir 'dashboard_share.html'
[System.IO.File]::WriteAllText($dashPath, (Build-Html 'dashboard.template.html' 'dashboard.js' $internal "퇴직 위험 신호 대시보드 $baseStr"), $utf8)
[System.IO.File]::WriteAllText($sharePath, (Build-Html 'share.template.html' 'share.js' $share "퇴직 위험 신호 부서별 요약 $baseStr"), $utf8)

# ---------------------------------------------------------------- 6) 콘솔 출력
Write-Host ''
Write-Host "기준일: $baseStr / 총인원 $total`명 (상 $($gradeCounts['상']), 중 $($gradeCounts['중']), 하 $($gradeCounts['하']), 판단불가 $($gradeCounts['판단불가']))"
Write-Host "원본: $(if ($hasSource) { '사용' } else { '없음' }) / 직원 연결: $(if ($linked) { '확인됨' } else { '안 함' })"
foreach ($n in $notices) { Write-Host "[$($n.level)] $($n.text)" }
Write-Host ''
Write-Host "내부용: $dashPath"
Write-Host "공유용: $sharePath"
$dashUrl = ([System.Uri]$dashPath).AbsoluteUri
Write-Host ''
Write-Host "열려면: 탐색기에서 dashboard.html을 더블클릭하거나, 브라우저 주소창에 아래 주소를 붙여 넣으세요."
Write-Host "  $dashUrl"
if ($Open) {
    try { Start-Process -FilePath $dashPath -ErrorAction Stop; Write-Host '기본 브라우저로 dashboard.html을 열었습니다.' }
    catch { Write-Host "자동으로 열지 못했습니다 ($($_.Exception.Message)). 위 경로의 파일을 직접 열어 주세요." }
}
