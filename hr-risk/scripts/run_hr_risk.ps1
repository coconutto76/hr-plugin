# hr-risk 본체: 인사데이터 엑셀 -> 퇴직위험_결과_YYYYMM.xlsx (+ ID매핑_YYYYMM.xlsx)
# 사용: powershell -ExecutionPolicy Bypass -File run_hr_risk.ps1 -InputPath <인사데이터.xlsx> [-BaseDate 2026-04-30] [-OutputDir <폴더>] [-MappingDir <폴더>]
[CmdletBinding()]
param(
    [string]$InputPath,
    [string]$BaseDate,
    [string]$OutputDir,
    [string]$MappingDir,
    [string]$ConfigPath
)
$ErrorActionPreference = 'Stop'
$inv = [System.Globalization.CultureInfo]::InvariantCulture

# 스킬 폴더 기준 상대 경로: 설치 위치가 바뀌어도 이 스크립트 위치에서 보조 파일을 찾는다.
$SkillRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'lib_xlsx.ps1')
. (Join-Path $PSScriptRoot 'lib_excel_out.ps1')
if (-not $ConfigPath) { $ConfigPath = Join-Path $SkillRoot 'config\risk_config.json' }
$cfg = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $InputPath) { $InputPath = Join-Path (Get-Location) '인사데이터_샘플.xlsx' }
$InputPath = (Resolve-Path -LiteralPath $InputPath).Path
if (-not $OutputDir) { $OutputDir = Split-Path -Parent $InputPath }
if (-not $MappingDir) { $MappingDir = $OutputDir }

# ---------------------------------------------------------------- 공통 도우미
$Issues = New-Object System.Collections.Generic.List[object]
function Add-Issue([string]$Level, [string]$Area, [string]$Item, [string]$Detail, $Count = '') {
    $Issues.Add(@($Level, $Area, $Item, $Detail, $Count)) | Out-Null
}
function Stop-WithIssues {
    $stops = @($Issues | Where-Object { $_[0] -eq '중단' })
    Write-Host '--- 중단: 아래 문제를 고친 뒤 다시 실행하세요 (결과 파일은 만들지 않았습니다) ---'
    foreach ($s in $stops) { Write-Host ("[{0}] {1} - {2}: {3} {4}" -f $s[0], $s[1], $s[2], $s[3], $s[4]) }
    exit 2
}
function Get-Key($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) { return [string][int64]$v }
    return ([string]$v).Trim()
}
function Clip([double]$x, [double]$lo = 0, [double]$hi = 1) { return [math]::Max($lo, [math]::Min($hi, $x)) }
function Mean($a) { $arr = @($a | Where-Object { $null -ne $_ }); if ($arr.Count -eq 0) { return $null }; return ($arr | Measure-Object -Average).Average }
function ConvertTo-MonthKey($v) {
    if ($null -eq $v -or "$v" -eq '') { return $null }
    if ($v -is [string] -and $v -match '^(\d{4})[-/.](\d{1,2})') { return ('{0}-{1:00}' -f $Matches[1], [int]$Matches[2]) }
    $d = ConvertTo-DateValue $v
    if ($d) { return $d.ToString('yyyy-MM') }
    return $null
}
function Fmt([double]$x, [string]$f = '0.0') { return $x.ToString($f, $inv) }
function Group-ByKey($Rows, [string]$Col) {
    $h = @{}
    foreach ($r in $Rows) {
        $k = Get-Key $r[$Col]
        if (-not $h.ContainsKey($k)) { $h[$k] = New-Object System.Collections.Generic.List[object] }
        $h[$k].Add($r)
    }
    return $h
}

# ---------------------------------------------------------------- 2단계: 데이터 읽기·점검
Write-Host "입력 읽는 중: $InputPath"
$wb = Read-XlsxWorkbook $InputPath

$core = [ordered]@{
    'v_employees_current'    = @('employee_id', 'as_of', 'division_name', 'org_unit_name', 'job_family', 'grade_level', 'employment_status', 'employment_type', 'contract_end_date', 'hire_date', 'annual_leave_granted', 'annual_leave_used')
    'v_attendance_monthly'   = @('employee_id', 'pay_month', 'overtime_hours')
    'hr_performance_reviews' = @('employee_id', 'cycle_id', 'status', 'final_rating', 'finalized_at')
}
$optional = [ordered]@{
    'hr_employees'           = @('employee_id', 'hire_date')
    'hr_compensation_history' = @('employee_id', 'effective_from', 'effective_to', 'base_annual_krw10k', 'band_min_krw10k', 'band_max_krw10k')
    'hr_promotion_reviews'   = @('employee_id', 'result', 'decided_at')
    'hr_personnel_actions'   = @('employee_id', 'effective_date', 'action_type')
}
foreach ($s in $core.Keys) {
    if (-not $wb.Contains($s)) { Add-Issue '중단' $s '시트 없음' '필수 시트가 없습니다'; continue }
    foreach ($c in $core[$s]) { if ($wb[$s].Headers -notcontains $c) { Add-Issue '중단' $s "열 없음: $c" '필수 열 이름이 다릅니다' } }
}
foreach ($s in $optional.Keys) {
    if (-not $wb.Contains($s)) { Add-Issue '경고' $s '시트 없음' '이 시트에서 계산하는 신호는 판단불가로 처리합니다'; continue }
    foreach ($c in $optional[$s]) { if ($wb[$s].Headers -notcontains $c) { Add-Issue '경고' $s "열 없음: $c" '이 시트의 신호는 판단불가로 처리합니다'; $wb.Remove($s); break } }
}
if (@($Issues | Where-Object { $_[0] -eq '중단' }).Count) { Stop-WithIssues }

Write-Host '[2/6] 데이터 점검 중'
# 기준일: 인자 > 개요 시트 > as_of. 서로 다르면 중단한다.
$emp = $wb['v_employees_current'].Rows
$asOfValues = @($emp | ForEach-Object { (ConvertTo-DateValue $_['as_of']) } | Where-Object { $_ } | ForEach-Object { $_.Date } | Sort-Object -Unique)
if ($asOfValues.Count -ne 1) { Add-Issue '중단' 'v_employees_current' 'as_of 불일치' '재직자 명단의 as_of가 한 날짜가 아니거나 비어 있습니다' $asOfValues.Count }
$overviewBase = $null
if ($wb.Contains('개요')) {
    foreach ($r in $wb['개요'].Rows) { if ($r['항목'] -eq '기준일') { $overviewBase = ConvertTo-DateValue $r['내용'] } }
}
if ($BaseDate) { $base = ([datetime]::Parse($BaseDate, $inv)).Date }
elseif ($asOfValues.Count -eq 1) { $base = $asOfValues[0] }
else { $base = $null }
if ($base) {
    if ($asOfValues.Count -eq 1 -and $asOfValues[0] -ne $base) { Add-Issue '중단' '기준일' '기준일 불일치' "지정/계산한 기준일 $($base.ToString('yyyy-MM-dd')) 와 as_of $($asOfValues[0].ToString('yyyy-MM-dd')) 가 다릅니다" }
    if ($overviewBase -and $overviewBase.Date -ne $base) { Add-Issue '중단' '기준일' '기준일 불일치' "개요 시트 기준일 $($overviewBase.ToString('yyyy-MM-dd')) 와 다릅니다" }
}
if (@($Issues | Where-Object { $_[0] -eq '중단' }).Count) { Stop-WithIssues }
$baseStr = $base.ToString('yyyy-MM-dd')

# 사번 중복·공백
$blank = @($emp | Where-Object { (Get-Key $_['employee_id']) -eq '' }).Count
if ($blank) { Add-Issue '중단' 'v_employees_current' '사번 공백' '사번이 비어 있는 행이 있습니다' $blank }
$dupIds = @($emp | Group-Object { Get-Key $_['employee_id'] } | Where-Object { $_.Count -gt 1 })
if ($dupIds.Count) { Add-Issue '중단' 'v_employees_current' '사번 중복' '같은 사번이 여러 행입니다' $dupIds.Count }
if (@($Issues | Where-Object { $_[0] -eq '중단' }).Count) { Stop-WithIssues }

# 대상자: 기준일 이전 입사자만. 기준일 이후 입사자는 정보가 없는 시점이므로 제외한다.
$pop = New-Object System.Collections.Generic.List[object]
$lateHire = 0
foreach ($e in $emp) {
    $h = ConvertTo-DateValue $e['hire_date']
    if ($h -and $h.Date -gt $base) { $lateHire++; continue }
    $pop.Add($e)
}
if ($lateHire) { Add-Issue '경고' 'v_employees_current' '기준일 이후 입사자 제외' '입사일이 기준일보다 늦은 직원은 분석에서 뺐습니다' $lateHire }
$popKeys = @{}
foreach ($e in $pop) { $popKeys[(Get-Key $e['employee_id'])] = $true }
$onLeave = @($pop | Where-Object { $_['employment_status'] -eq '휴직' }).Count
if ($onLeave) { Add-Issue '정보' 'v_employees_current' '휴직자' '휴직 중인 직원은 근태 신호가 적을 수 있습니다' $onLeave }

# 입사일 교차 확인
if ($wb.Contains('hr_employees')) {
    $he = @{}; foreach ($r in $wb['hr_employees'].Rows) { $he[(Get-Key $r['employee_id'])] = ConvertTo-DateValue $r['hire_date'] }
    $mis = 0; $miss = 0
    foreach ($e in $pop) {
        $k = Get-Key $e['employee_id']
        if (-not $he.ContainsKey($k)) { $miss++; continue }
        $d1 = ConvertTo-DateValue $e['hire_date']
        if ($d1 -and $he[$k] -and $d1.Date -ne $he[$k].Date) { $mis++ }
    }
    if ($miss) { Add-Issue '경고' 'hr_employees' '입사 기록 없음' '재직자 명단에는 있으나 입사 기록 시트에 없는 직원' $miss }
    if ($mis) { Add-Issue '경고' 'hr_employees' '입사일 불일치' '두 시트의 입사일이 다른 직원 (재직자 명단의 입사일 사용)' $mis }
}

# 근태: 기준일 이후 월 제외, 중복 점검
$lastFull = if ($base.AddDays(1).Day -eq 1) { [datetime]::new($base.Year, $base.Month, 1) } else { [datetime]::new($base.Year, $base.Month, 1).AddMonths(-1) }
$lastFullKey = $lastFull.ToString('yyyy-MM')
$att = @{}
$attAfter = 0; $attDup = 0; $attBadMonth = 0; $attBadValue = 0; $attNotPop = 0
foreach ($r in $wb['v_attendance_monthly'].Rows) {
    $k = Get-Key $r['employee_id']
    $m = ConvertTo-MonthKey $r['pay_month']
    if (-not $m) { $attBadMonth++; continue }
    if ($m -gt $lastFullKey) { $attAfter++; continue }
    if (-not $popKeys.ContainsKey($k)) { $attNotPop++; continue }
    $ot = $r['overtime_hours']
    if ($null -eq $ot -or $ot -isnot [double] -or $ot -lt 0 -or $ot -gt 200) { $attBadValue++; continue }
    if (-not $att.ContainsKey($k)) { $att[$k] = @{} }
    if ($att[$k].ContainsKey($m)) { $attDup++ }
    $att[$k][$m] = $ot
}
Add-Issue '정보' 'v_attendance_monthly' '기준일 이후 월 제외' "기준일 이후에 생긴 근태 행은 사용하지 않았습니다 (사용 가능한 마지막 월 $lastFullKey)" $attAfter
if ($attBadMonth) { Add-Issue '경고' 'v_attendance_monthly' '월 해석 불가' 'pay_month를 해석할 수 없는 행을 뺐습니다' $attBadMonth }
if ($attBadValue) { Add-Issue '경고' 'v_attendance_monthly' '초과근무 값 이상' '비었거나 음수/200 초과인 초과근무 값을 뺐습니다' $attBadValue }
if ($attDup) { Add-Issue '경고' 'v_attendance_monthly' '직원-월 중복' '같은 직원·월 행이 여러 개라 마지막 값을 썼습니다' $attDup }
if ($attNotPop) { Add-Issue '경고' 'v_attendance_monthly' '대상 아닌 직원' '재직자 명단에 없는 직원의 근태 행을 뺐습니다' $attNotPop }

# 평가: 기준일 이후 확정분 제외, 미확정 제외, 중복 점검
$rev = @{}
$revAfter = 0; $revNotFinal = 0; $revDup = 0; $revBadRating = 0
$ratingOrder = @($cfg.rating.order)
foreach ($r in $wb['hr_performance_reviews'].Rows) {
    $k = Get-Key $r['employee_id']
    if (-not $popKeys.ContainsKey($k)) { continue }
    if ($r['status'] -ne '확정') { $revNotFinal++; continue }
    $fd = ConvertTo-DateValue $r['finalized_at']
    if (-not $fd) { $revNotFinal++; continue }
    if ($fd.Date -gt $base) { $revAfter++; continue }
    if ($ratingOrder -notcontains [string]$r['final_rating']) { $revBadRating++; continue }
    if (-not $rev.ContainsKey($k)) { $rev[$k] = @{} }
    if ($rev[$k].ContainsKey([string]$r['cycle_id'])) { $revDup++ }
    $rev[$k][[string]$r['cycle_id']] = @{ At = $fd; Rating = [string]$r['final_rating'] }
}
Add-Issue '정보' 'hr_performance_reviews' '기준일 이후 확정 평가 제외' '기준일 이후에 확정된 평가는 사용하지 않았습니다' $revAfter
if ($revNotFinal) { Add-Issue '경고' 'hr_performance_reviews' '미확정 평가 제외' '확정이 아니거나 확정일이 없는 평가를 뺐습니다' $revNotFinal }
if ($revBadRating) { Add-Issue '경고' 'hr_performance_reviews' '등급 값 이상' '설정에 없는 평가 등급을 뺐습니다' $revBadRating }
if ($revDup) { Add-Issue '경고' 'hr_performance_reviews' '직원-주기 중복' '같은 직원·평가주기 행이 여러 개라 마지막 값을 썼습니다' $revDup }

# 보상 이력: 기준일 이후 시작분 제외
$comp = @{}
if ($wb.Contains('hr_compensation_history')) {
    $cAfter = 0
    foreach ($r in $wb['hr_compensation_history'].Rows) {
        $k = Get-Key $r['employee_id']
        if (-not $popKeys.ContainsKey($k)) { continue }
        $from = ConvertTo-DateValue $r['effective_from']
        if (-not $from) { continue }
        if ($from.Date -gt $base) { $cAfter++; continue }
        if (-not $comp.ContainsKey($k) -or $comp[$k].From -lt $from) { $comp[$k] = @{ From = $from; Pay = $r['base_annual_krw10k']; Min = $r['band_min_krw10k']; Max = $r['band_max_krw10k'] } }
    }
    Add-Issue '정보' 'hr_compensation_history' '기준일 이후 시작 이력 제외' '기준일 이후에 시작하는 보상 이력은 사용하지 않았습니다' $cAfter
}

# 승진 심사: 기준일 이후 결정분 제외
$promoHold = @{}
$promoOn = $false
if ($wb.Contains('hr_promotion_reviews')) {
    $promoOn = $true
    $pAfter = 0
    $lookback = $base.AddMonths(-[int]$cfg.promotion_hold.lookback_months)
    foreach ($r in $wb['hr_promotion_reviews'].Rows) {
        $k = Get-Key $r['employee_id']
        $d = ConvertTo-DateValue $r['decided_at']
        if (-not $d) { continue }
        if ($d.Date -gt $base) { $pAfter++; continue }
        if ($r['result'] -eq '보류' -and $d.Date -ge $lookback) { $promoHold[$k] = $true }
    }
    Add-Issue '정보' 'hr_promotion_reviews' '기준일 이후 결정 제외' '기준일 이후에 결정된 승진 심사는 사용하지 않았습니다' $pAfter
}

# 인사발령: 점수에는 쓰지 않는다. 기준일 이후 퇴사 기록이 점수로 새지 않았는지 확인하는 용도.
if ($wb.Contains('hr_personnel_actions')) {
    $paAfter = 0; $paBefore = 0
    foreach ($r in $wb['hr_personnel_actions'].Rows) {
        $d = ConvertTo-DateValue $r['effective_date']
        if (-not $d) { continue }
        if ($d.Date -gt $base) { $paAfter++ }
        elseif ($r['action_type'] -eq '퇴사' -and $popKeys.ContainsKey((Get-Key $r['employee_id']))) { $paBefore++ }
    }
    Add-Issue '정보' 'hr_personnel_actions' '기준일 이후 발령 제외' '점수 계산에 쓰지 않았습니다 (퇴사 기록이 점수로 들어가는 것을 막기 위함)' $paAfter
    if ($paBefore) { Add-Issue '경고' 'hr_personnel_actions' '기준일 이전 퇴사 기록이 있는 재직자' '명단과 인사발령이 맞지 않습니다' $paBefore }
}

# 점수에 쓰지 않는 시트/열을 명시한다.
Add-Issue '정보' '제외 정보' '점수 산정에서 제외' '이름, 사번(분석용 ID만 사용), 병가(sick_days), 설문 응답·미응답, 성별, 나이·생년월일, 혼인 여부, 국적, 거주지, 퇴사라벨' ''

# ---------------------------------------------------------------- 3단계: 신호 계산
Write-Host '[3/6] 신호 계산 중'
$W = $cfg.weights
$signalDefs = @(
    @{ Key = 'overtime';       Label = '초과근무 증가'; Weight = [double]$W.overtime }
    @{ Key = 'rating';         Label = '평가 변화';     Weight = [double]$W.rating }
    @{ Key = 'tenure';         Label = '근속 구간';     Weight = [double]$W.tenure }
    @{ Key = 'leave_use';      Label = '연차 소진 낮음'; Weight = [double]$W.leave_use }
    @{ Key = 'promotion_hold'; Label = '승진 보류';     Weight = [double]$W.promotion_hold }
    @{ Key = 'pay_band';       Label = '급여 밴드 하단'; Weight = [double]$W.pay_band }
)
$recentN = [int]$cfg.overtime.recent_months; $prevN = [int]$cfg.overtime.previous_months
$minM = [int]$cfg.overtime.min_months_per_window
$recentKeys = 0..($recentN - 1) | ForEach-Object { $lastFull.AddMonths(-$_).ToString('yyyy-MM') }
$prevKeys = $recentN..($recentN + $prevN - 1) | ForEach-Object { $lastFull.AddMonths(-$_).ToString('yyyy-MM') }

$people = New-Object System.Collections.Generic.List[object]
foreach ($e in $pop) {
    $k = Get-Key $e['employee_id']
    $p = [ordered]@{ Key = $k; Div = $e['division_name']; Org = $e['org_unit_name']; JobFamily = $e['job_family']; Grade = $e['grade_level']; S = @{}; Note = @{} }
    $hire = ConvertTo-DateValue $e['hire_date']
    $tenure = if ($hire) { ($base - $hire.Date).TotalDays / 365.25 } else { $null }
    $p.Tenure = $tenure

    # 초과근무 증가: 최근 3개월 평균 - 직전 3개월 평균
    $rv = @(); $pv = @()
    if ($att.ContainsKey($k)) {
        $rv = @($recentKeys | Where-Object { $att[$k].ContainsKey($_) } | ForEach-Object { $att[$k][$_] })
        $pv = @($prevKeys | Where-Object { $att[$k].ContainsKey($_) } | ForEach-Object { $att[$k][$_] })
    }
    if ($rv.Count -ge $minM -and $pv.Count -ge $minM) {
        $ra = Mean $rv; $pa = Mean $pv; $delta = $ra - $pa
        $p.S.overtime = Clip ($delta / [double]$cfg.overtime.delta_cap_hours)
        $p.Note.overtime = "최근 ${recentN}개월 평균 $(Fmt $ra)h, 직전 ${prevN}개월 평균 $(Fmt $pa)h ($(if ($delta -ge 0) { '+' })$(Fmt $delta)h)"
    } else { $p.S.overtime = $null; $p.Note.overtime = '근태 월 부족' }

    # 평가 변화: 최근 확정 평가 vs 직전 확정 평가
    $p.S.rating = $null; $p.Note.rating = '확정 평가 1건 이하'
    if ($rev.ContainsKey($k) -and $rev[$k].Count -ge 2) {
        $sorted = @($rev[$k].Values | Sort-Object { $_.At })
        $latest = $sorted[-1]; $prior = $sorted[-2]
        $diff = [array]::IndexOf($ratingOrder, $latest.Rating) - [array]::IndexOf($ratingOrder, $prior.Rating)
        $flatLowIdx = [array]::IndexOf($ratingOrder, [string]$cfg.rating.flat_at_or_below)
        if ($diff -le -2) { $s = [double]$cfg.rating.drop_2_or_more }
        elseif ($diff -eq -1) { $s = [double]$cfg.rating.drop_1 }
        elseif ($diff -eq 0 -and [array]::IndexOf($ratingOrder, $latest.Rating) -le $flatLowIdx) { $s = [double]$cfg.rating.flat_low_score }
        else { $s = 0.0 }
        $p.S.rating = $s
        $p.Note.rating = "평가 $($prior.Rating)→$($latest.Rating)"
    }

    # 근속 구간
    if ($null -ne $tenure) {
        foreach ($b in $cfg.tenure_bins) { if ($tenure -lt [double]$b.upto) { $p.S.tenure = [double]$b.score; $p.Note.tenure = "근속 $(Fmt $tenure)년($($b.label))"; break } }
    } else { $p.S.tenure = $null; $p.Note.tenure = '입사일 없음' }

    # 연차 소진 낮음
    $g = $e['annual_leave_granted']; $u = $e['annual_leave_used']
    if ($g -is [double] -and $g -gt 0 -and $u -is [double]) {
        $rate = $u / $g
        $p.S.leave_use = if ($rate -lt [double]$cfg.leave_use.low_rate) { 1.0 } elseif ($rate -lt [double]$cfg.leave_use.mid_rate) { 0.5 } else { 0.0 }
        $p.Note.leave_use = "연차 소진율 $(Fmt ($rate * 100) '0')%"
    } else { $p.S.leave_use = $null; $p.Note.leave_use = '연차 정보 없음' }

    # 승진 보류
    if ($promoOn) {
        $p.S.promotion_hold = if ($promoHold.ContainsKey($k)) { 1.0 } else { 0.0 }
        $p.Note.promotion_hold = '최근 승진 심사 보류'
    } else { $p.S.promotion_hold = $null; $p.Note.promotion_hold = '승진 심사 시트 없음' }

    # 급여 밴드 내 위치
    $p.S.pay_band = $null; $p.Note.pay_band = '보상 이력 없음'
    if ($comp.ContainsKey($k)) {
        $c = $comp[$k]
        if ($c.Max -is [double] -and $c.Min -is [double] -and $c.Pay -is [double] -and $c.Max -gt $c.Min) {
            $pos = ($c.Pay - $c.Min) / ($c.Max - $c.Min)
            $mid = [double]$cfg.pay_band.midpoint
            $p.S.pay_band = Clip (($mid - $pos) / $mid)
            $p.Note.pay_band = "급여 밴드 내 위치 $(Fmt ($pos * 100) '0')%"
        }
    }

    # 합산: 계산 가능한 신호의 가중치만 사용해 0~100으로 환산한다.
    $availW = 0.0; $sum = 0.0; $na = @()
    foreach ($d in $signalDefs) {
        $s = $p.S[$d.Key]
        if ($null -eq $s) { $na += $d.Label } else { $availW += $d.Weight; $sum += $d.Weight * $s }
    }
    $p.AvailW = $availW
    $p.NA = $na
    if ($availW -ge [double]$cfg.min_available_weight) {
        $p.Score = [math]::Round(100 * $sum / $availW, 1)
        $p.Grade2 = if ($p.Score -ge [double]$cfg.grade_cutoffs.high) { '상' } elseif ($p.Score -ge [double]$cfg.grade_cutoffs.mid) { '중' } else { '하' }
    } else { $p.Score = $null; $p.Grade2 = '판단불가' }

    $hits = @($signalDefs | Where-Object { $null -ne $p.S[$_.Key] -and $p.S[$_.Key] -ge [double]$cfg.flag_threshold })
    $p.Hits = @($hits | ForEach-Object { $_.Key })
    $p.Reason = if ($hits.Count) { ($hits | ForEach-Object { "$($_.Label): $($p.Note[$_.Key])" }) -join ' / ' } else { '두드러진 신호 없음' }
    $p.Action = if ($hits.Count) { (@($hits | ForEach-Object { $cfg.actions.($_.Key) } | Select-Object -Unique)) -join ' / ' } else { '정기 면담 유지' }
    $ce = ConvertTo-DateValue $e['contract_end_date']
    $p.Ref = ''
    if ($ce -and $ce.Date -ge $base -and $ce.Date -le $base.AddDays([int]$cfg.observation.primary_days)) { $p.Ref = "계약 만료 예정(관찰 기간 내) - 자발적 퇴사 위험과 구분해서 보세요" }
    $people.Add([pscustomobject]$p)
}

$noOt = @($people | Where-Object { $null -eq $_.S.overtime }).Count
$noRv = @($people | Where-Object { $null -eq $_.S.rating }).Count
if ($noOt) { Add-Issue '경고' 'v_attendance_monthly' '초과근무 신호 판단불가' "최근 ${recentN}개월/직전 ${prevN}개월 중 각각 ${minM}개월 이상 근태가 없는 직원" $noOt }
if ($noRv) { Add-Issue '경고' 'hr_performance_reviews' '평가 변화 판단불가' '기준일 이전 확정 평가가 2건 미만인 직원' $noRv }
$noGrade = @($people | Where-Object { $_.Grade2 -eq '판단불가' }).Count
if ($noGrade) { Add-Issue '경고' '점수' '등급 판단불가' "계산 가능한 신호의 가중치 합이 $($cfg.min_available_weight) 미만이라 등급을 매기지 않았습니다" $noGrade }

# 분석용 ID: 기준일을 씨앗으로 섞어 번호를 매긴다. 같은 입력·기준일이면 같은 ID가 나오고, 사번 순서는 드러나지 않는다.
$rng = New-Object System.Random([int]$base.ToString('yyyyMMdd'))
$order = 0..($people.Count - 1)
for ($i = $order.Count - 1; $i -gt 0; $i--) { $j = $rng.Next($i + 1); $t = $order[$i]; $order[$i] = $order[$j]; $order[$j] = $t }
for ($i = 0; $i -lt $people.Count; $i++) { $people[$i] | Add-Member -NotePropertyName AnalysisId -NotePropertyValue ('R{0:000}' -f ($order[$i] + 1)) }

# ---------------------------------------------------------------- 4단계: 검증 (퇴사라벨은 여기서만 사용)
Write-Host '[4/6] 검증 중'
$validationRows = New-Object System.Collections.Generic.List[object]
$validationTitles = New-Object System.Collections.Generic.List[int]
$eventCount = $null
function Add-V($row, [bool]$title = $false) { if ($title) { $validationTitles.Add($validationRows.Count) }; $validationRows.Add($row) | Out-Null }
$labelOk = $false
if ($wb.Contains('퇴사라벨') -and $wb['퇴사라벨'].Headers -contains 'left') {
    $lab = @{}
    foreach ($r in $wb['퇴사라벨'].Rows) { $lab[(Get-Key $r['employee_id'])] = $r }
    $obsOk = @($lab.Values | Where-Object { $d = ConvertTo-DateValue $_['observed_until']; $d -and $d.Date -lt $base.AddDays([int]$cfg.observation.primary_days) }).Count
    if ($obsOk) { Add-Issue '경고' '퇴사라벨' '관측 기간 부족' "관측 종료일이 기준일+$($cfg.observation.primary_days)일보다 이른 직원이 있어 검증 비율이 낮게 나올 수 있습니다" $obsOk }
    $labelOk = $true
}
$gradeOrder = @('상', '중', '하', '판단불가')
if ($labelOk) {
    $eventCount = @($people | Where-Object { $lab.ContainsKey($_.Key) -and $lab[$_.Key]['left'] -eq 1 }).Count
    Add-V @('검증 1. 등급별 실제 결과 (사후 점검, 점수 계산에는 쓰지 않음)') $true
    Add-V @('등급', '인원', "자발적 퇴사 $($cfg.observation.primary_days)일 이내", '비율', "퇴사 $($cfg.observation.reference_days)일 이내(유형 불문, 참고)", '비율')
    foreach ($gname in $gradeOrder) {
        $grp = @($people | Where-Object { $_.Grade2 -eq $gname })
        if ($grp.Count -eq 0) { continue }
        $l90 = @($grp | Where-Object { $lab.ContainsKey($_.Key) -and $lab[$_.Key]['left'] -eq 1 }).Count
        $l180 = @($grp | Where-Object { $lab.ContainsKey($_.Key) -and $lab[$_.Key]['left_any_180d'] -eq 1 }).Count
        Add-V @($gname, $grp.Count, $l90, ('{0:P0}' -f ($l90 / $grp.Count)), $l180, ('{0:P0}' -f ($l180 / $grp.Count)))
    }
    $totAll = $people.Count
    Add-V @('전체', $totAll, $eventCount, ('{0:P0}' -f ($eventCount / [math]::Max(1, $totAll))), @($people | Where-Object { $lab.ContainsKey($_.Key) -and $lab[$_.Key]['left_any_180d'] -eq 1 }).Count, '')
    Add-V @('')
    Add-V @('검증 2. 신호별 반례 확인 (신호 점수가 기준 이상이면 "신호 있음")') $true
    Add-V @('신호', '신호 있음 인원', '그중 퇴사', '퇴사 비율', '신호 없음 인원', '그중 퇴사', '퇴사 비율', '반례 A: 신호 있는데 퇴사 안 함', '반례 B: 신호 없는데 퇴사', '판단불가 인원')
    foreach ($d in $signalDefs) {
        $has = @($people | Where-Object { $null -ne $_.S[$d.Key] -and $_.S[$d.Key] -ge [double]$cfg.flag_threshold })
        $not = @($people | Where-Object { $null -ne $_.S[$d.Key] -and $_.S[$d.Key] -lt [double]$cfg.flag_threshold })
        $na = @($people | Where-Object { $null -eq $_.S[$d.Key] }).Count
        $hl = @($has | Where-Object { $lab.ContainsKey($_.Key) -and $lab[$_.Key]['left'] -eq 1 }).Count
        $nl = @($not | Where-Object { $lab.ContainsKey($_.Key) -and $lab[$_.Key]['left'] -eq 1 }).Count
        $r1 = if ($has.Count) { '{0:P0}' -f ($hl / $has.Count) } else { '-' }
        $r2 = if ($not.Count) { '{0:P0}' -f ($nl / $not.Count) } else { '-' }
        Add-V @($d.Label, $has.Count, $hl, $r1, $not.Count, $nl, $r2, ($has.Count - $hl), $nl, $na)
    }
    Add-V @('')
} else {
    Add-V @('검증 1~2. 퇴사 이력(퇴사라벨 시트)이 없어 등급별·신호별 사후 점검을 하지 않았습니다.') $true
    Add-V @('')
}
Add-V @('검증 3. 기준일 이후 정보 사용 여부') $true
Add-V @('사용하지 않은 정보', '건수')
foreach ($i in @($Issues | Where-Object { $_[2] -like '*기준일 이후*' -and $_[0] -eq '정보' })) { Add-V @("$($i[1]): $($i[2])", $i[4]) }
Add-V @('')
$limitLines = New-Object System.Collections.Generic.List[string]
$limitLines.Add('점수는 직원을 우선 확인할 순서를 나타내는 지표이며 퇴사 확률이 아닙니다.')
if ($labelOk) {
    $limitLines.Add("검증에 쓸 수 있는 자발적 퇴사가 $eventCount 건뿐이고 기준일이 하나라 시간 흐름에 따른 검증을 할 수 없습니다. 예측 성능을 확인할 수 없습니다.")
    if ($eventCount -lt [int]$cfg.min_events_for_validation) { $limitLines.Add("퇴사 건수가 기준($($cfg.min_events_for_validation)건)보다 적어 위 사후 점검 비율은 우연의 영향이 큽니다. 참고로만 보세요.") }
} else {
    $limitLines.Add('퇴사 이력이 없어 예측 성능을 확인할 수 없습니다.')
}
$limitLines.Add('가중치와 등급 경계는 초기 설정값이며, 신호를 합산하는 규칙 기반 점수입니다. 입사·근태·평가 데이터에 없는 사정(개인 사유, 외부 제안 등)은 반영하지 못합니다.')
$limitLines.Add('유지 우선순위는 위험 점수와 별도로 HR이 판단합니다. 이 점수에는 핵심 인재 여부·성과 가치 같은 유지 판단이 섞여 있지 않고, 점수와 합친 순위도 만들지 않았습니다.')
Add-V @('분석의 한계') $true
foreach ($l in $limitLines) { Add-V @($l) }

# ---------------------------------------------------------------- 조직 집계
$minGroup = [int]$cfg.min_group_size
function Get-GroupRow($label1, $label2, $members) {
    $n = @($members).Count
    if ($n -lt $minGroup) { return @($label1, $label2, $n, '표본 부족', '', '', '', '', '') }
    $scored = @($members | Where-Object { $null -ne $_.Score })
    $avg = if ($scored.Count) { Fmt (Mean ($scored | ForEach-Object { $_.Score })) } else { '' }
    $hi = @($members | Where-Object { $_.Grade2 -eq '상' }).Count
    $mi = @($members | Where-Object { $_.Grade2 -eq '중' }).Count
    $lo = @($members | Where-Object { $_.Grade2 -eq '하' }).Count
    $na = @($members | Where-Object { $_.Grade2 -eq '판단불가' }).Count
    $top = ($signalDefs | ForEach-Object { $d = $_; [pscustomobject]@{ L = $d.Label; C = @($members | Where-Object { $_.Hits -contains $d.Key }).Count } } | Sort-Object C -Descending | Select-Object -First 1)
    return @($label1, $label2, $n, $avg, $hi, $mi, $lo, $na, $(if ($top.C -gt 0) { "$($top.L) ($($top.C)명)" } else { '-' }))
}
$orgRows = New-Object System.Collections.Generic.List[object]
foreach ($dv in ($people | Group-Object Div | Sort-Object Name)) {
    $orgRows.Add((Get-GroupRow $dv.Name '(본부 전체)' $dv.Group))
    foreach ($og in ($dv.Group | Group-Object Org | Sort-Object Name)) { $orgRows.Add((Get-GroupRow $dv.Name $og.Name $og.Group)) }
}

# ---------------------------------------------------------------- 요약 3줄
$cnt = @{}; foreach ($gname in $gradeOrder) { $cnt[$gname] = @($people | Where-Object { $_.Grade2 -eq $gname }).Count }
$total = $people.Count
$pct = { param($x) '{0:P0}' -f ($x / [math]::Max(1, $total)) }
$line1 = "기준일 $baseStr 재직자 $total`명 중 상 $($cnt['상'])명($(& $pct $cnt['상'])), 중 $($cnt['중'])명, 하 $($cnt['하'])명, 판단불가 $($cnt['판단불가'])명입니다. 점수는 우선 확인 순서를 나타내며 퇴사 확률이 아닙니다."
# 근속 구간은 모든 직원에게 값이 있는 '프로필' 성격이라 "많이 나타난 신호" 순위에서는 뺀다.
$sigCounts = $signalDefs | Where-Object { $_.Key -ne 'tenure' } | ForEach-Object { $d = $_; [pscustomobject]@{ L = $d.Label; C = @($people | Where-Object { $_.Hits -contains $d.Key }).Count } } | Sort-Object C -Descending
$orgOnly = @($orgRows | Where-Object { $_[1] -ne '(본부 전체)' -and $_[3] -ne '표본 부족' -and $_[4] -ne '' })
$topOrg = $orgOnly | Sort-Object @{ Expression = { [double]$_[4] / [double]$_[2] }; Descending = $true }, @{ Expression = { [double]$_[3] }; Descending = $true } | Select-Object -First 1
$orgText = if ($topOrg -and [int]$topOrg[4] -gt 0) { "상 등급 비율이 가장 높은 조직은 $($topOrg[0]) $($topOrg[1])($($topOrg[4])/$($topOrg[2])명)" } else { '상 등급이 몰린 조직은 없습니다' }
$line2 = "$orgText 이고, 가장 많이 나타난 신호는 $($sigCounts[0].L) $($sigCounts[0].C)명, 다음은 $($sigCounts[1].L) $($sigCounts[1].C)명입니다."
$evText = if ($labelOk) { "자발적 퇴사 $eventCount`건·기준일 1개로 예측 성능은 확인할 수 없습니다" } else { '퇴사 이력이 없어 예측 성능은 확인할 수 없습니다' }
$line3 = "상 등급부터 업무량 점검, 면담, 성장 기회 논의로 확인을 시작하세요. $evText. 유지 우선순위는 HR이 따로 판단합니다."
$summaryLines = @($line1, $line2, $line3)

# ---------------------------------------------------------------- 결과 파일 쓰기
Write-Host '[5/6] 결과 파일 쓰는 중'
$ym = $base.ToString('yyyyMM')
$resultPath = Join-Path $OutputDir "퇴직위험_결과_$ym.xlsx"
$mapPath = Join-Path $MappingDir "ID매핑_$ym.xlsx"
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $MappingDir | Out-Null

$sortedPeople = @($people | Sort-Object @{ Expression = { if ($null -eq $_.Score) { -1 } else { $_.Score } }; Descending = $true })
function Get-SignalCell($p, $key) { $s = $p.S[$key]; if ($null -eq $s) { return '판단불가' } else { return [math]::Round($s, 2) } }

$xl = $null; $book = $null; $mapBook = $null
$script:FirstSheetUsed = $false
try {
    $xl = Start-ExcelApp
    $book = $xl.Workbooks.Add()

    # 요약
    $ws = New-Sheet $book '요약'
    $rows = @(
        @("퇴직 위험 신호 검토 결과 - 기준일 $baseStr"),
        @(''),
        @('요약 3줄'),
        @($line1), @($line2), @($line3),
        @(''),
        @('분석 조건'),
        @("기준일: $baseStr / 관찰 기간: 기준일 이후 $($cfg.observation.primary_days)일 이내 자발적 퇴사 (참고: $($cfg.observation.reference_days)일)"),
        @("대상: 기준일 재직자 $total`명 / 입력 파일: $(Split-Path -Leaf $InputPath)"),
        @('이 파일에는 이름·사번이 없습니다. 행은 분석용 ID로 구분하며, 대응표는 별도 파일(ID매핑)에 있습니다.'),
        @(''),
        @('분석의 한계')
    ) + @($limitLines | ForEach-Object { , @($_) })
    Write-GridToSheet $ws $rows 1
    $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Size = 14
    foreach ($r in 3, 8, 13) { $ws.Cells($r, 1).Font.Bold = $true }
    Set-ColumnWidths $ws @(130)
    $ws.Range("A1:A$($rows.Count)").WrapText = $true

    # 개인결과
    $ws = New-Sheet $book '개인결과'
    $hdr = @('분석용 ID', '등급', '점수(0~100, 우선 확인 순서)', '본부', '조직', '직군', '직급', '근속(년)', '계산에 쓴 가중치 합', '초과근무 증가', '평가 변화', '근속 구간', '연차 소진 낮음', '승진 보류', '급여 밴드 하단', '판단불가 신호', '판단 근거', '추가 확인 사항', '참고')
    $rows = @(, $hdr)
    foreach ($p in $sortedPeople) {
        $rows += , @($p.AnalysisId, $p.Grade2, $(if ($null -eq $p.Score) { '' } else { $p.Score }), $p.Div, $p.Org, $p.JobFamily, $p.Grade,
            $(if ($null -eq $p.Tenure) { '' } else { [math]::Round($p.Tenure, 1) }), $p.AvailW,
            (Get-SignalCell $p 'overtime'), (Get-SignalCell $p 'rating'), (Get-SignalCell $p 'tenure'), (Get-SignalCell $p 'leave_use'), (Get-SignalCell $p 'promotion_hold'), (Get-SignalCell $p 'pay_band'),
            ($p.NA -join ', '), $p.Reason, $p.Action, $p.Ref)
    }
    Write-GridToSheet $ws $rows $hdr.Count
    Set-HeaderStyle $xl $ws $hdr.Count
    Set-ColumnWidths $ws @(10, 9, 16, 12, 12, 12, 7, 9, 12, 10, 10, 10, 11, 9, 11, 22, 70, 50, 40)
    Set-GradeColors $ws 2 2 $rows.Count
    $ws.Cells($rows.Count + 2, 1).Value2 = '점수는 우선 확인할 순서를 나타내며 퇴사 확률이 아닙니다. 유지 우선순위는 이 점수와 합치지 않고 HR이 따로 판단합니다. 이름·사번은 이 파일에 없습니다.'
    $ws.Range($ws.Cells(1, 1), $ws.Cells($rows.Count, $hdr.Count)).AutoFilter(1) | Out-Null

    # 조직결과
    $ws = New-Sheet $book '조직결과'
    $hdr = @('본부', '조직', '인원', '평균 점수', '상', '중', '하', '판단불가', '가장 많은 신호')
    $rows = @(, $hdr) + @($orgRows | ForEach-Object { , $_ })
    Write-GridToSheet $ws $rows $hdr.Count
    Set-HeaderStyle $xl $ws $hdr.Count
    Set-ColumnWidths $ws @(14, 16, 8, 11, 7, 7, 7, 10, 28)
    $ws.Cells($rows.Count + 2, 1).Value2 = "인원이 $minGroup`명 미만인 조직은 개인이 드러날 수 있어 '표본 부족'으로 표시하고 수치를 숨겼습니다."

    # 검증
    $ws = New-Sheet $book '검증'
    $vWidth = [int](($validationRows.ToArray() | ForEach-Object { @($_).Count } | Measure-Object -Maximum).Maximum)
    Write-GridToSheet $ws $validationRows.ToArray() $vWidth
    foreach ($t in $validationTitles) { $ws.Cells($t + 1, 1).Font.Bold = $true }
    Set-ColumnWidths $ws (@(34) + @(1..($vWidth - 1) | ForEach-Object { 18 }))

    # 산정기준
    $ws = New-Sheet $book '산정기준'
    $rows = @(
        @('신호', '가중치', '계산 방식', '점수 범위'),
        @('초과근무 증가', $W.overtime, "최근 ${recentN}개월 평균 초과근무 - 직전 ${prevN}개월 평균 (각 구간 ${minM}개월 이상 있어야 계산). 증가분 $($cfg.overtime.delta_cap_hours)시간 이상이면 1", '0~1'),
        @('평가 변화', $W.rating, "최근 확정 평가 vs 직전 확정 평가. 2단계 이상 하락 $($cfg.rating.drop_2_or_more), 1단계 하락 $($cfg.rating.drop_1), 같은 등급이 $($cfg.rating.flat_at_or_below) 이하이면 $($cfg.rating.flat_low_score), 그 밖에 0", '0~1'),
        @('근속 구간', $W.tenure, (($cfg.tenure_bins | ForEach-Object { "$($_.label) $($_.score)" }) -join ', '), '0~1'),
        @('연차 소진 낮음', $W.leave_use, "연차 사용/부여 비율이 $($cfg.leave_use.low_rate) 미만 1, $($cfg.leave_use.mid_rate) 미만 0.5", '0~1'),
        @('승진 보류', $W.promotion_hold, "최근 $($cfg.promotion_hold.lookback_months)개월 안에 승진 심사 결과가 보류", '0 또는 1'),
        @('급여 밴드 하단', $W.pay_band, '급여 밴드 안 위치가 중간보다 낮을수록 높음 (최하단 1)', '0~1'),
        @(''),
        @('점수 환산', '', '계산 가능한 신호의 (가중치 x 점수) 합 / 계산 가능한 가중치 합 x 100. 신호가 부족한 직원이 불리해지지 않게 하기 위함', ''),
        @('판단불가', '', "계산 가능한 가중치 합이 $($cfg.min_available_weight) 미만이면 점수와 등급을 매기지 않음", ''),
        @('등급', '', "상: $($cfg.grade_cutoffs.high) 이상 / 중: $($cfg.grade_cutoffs.mid) 이상 $($cfg.grade_cutoffs.high) 미만 / 하: $($cfg.grade_cutoffs.mid) 미만 (고정 점수 구간)", ''),
        @('신호 있음 기준', '', "신호 점수 $($cfg.flag_threshold) 이상이면 판단 근거에 표시하고 추가 확인 사항을 붙임", ''),
        @(''),
        @('점수 산정에서 제외한 정보', '', '이름, 사번, 병가, 설문 응답/미응답, 성별, 나이·생년월일, 혼인 여부, 국적, 거주지, 퇴사라벨, 인사발령(퇴사 기록)', ''),
        @('기준일 이후 정보', '', "기준일($baseStr) 이후에 생긴 근태·평가·보상·승진 기록과 퇴사일(인사발령의 퇴사 기록, 퇴사라벨)은 점수 입력에서 모두 제외", ''),
        @(''),
        @('점수의 의미', '', '점수는 우선 확인할 순서를 나타내는 지표이며 퇴사 확률이 아닙니다.', ''),
        @('유지 우선순위', '', '점수와 합치지 않습니다. 핵심 인재 여부·성과 가치 같은 유지 판단은 점수에 넣지 않았고, 누구를 붙잡을지는 HR이 별도로 판단합니다.', '')
    )
    Write-GridToSheet $ws $rows 4
    Set-HeaderStyle $xl $ws 4
    Set-ColumnWidths $ws @(26, 9, 110, 12)
    $ws.Range("C1:C$($rows.Count)").WrapText = $true

    # 데이터점검
    $ws = New-Sheet $book '데이터점검'
    $hdr = @('구분', '대상', '항목', '내용', '건수')
    $rows = @(, $hdr) + @($Issues | ForEach-Object { , $_ })
    Write-GridToSheet $ws $rows $hdr.Count
    Set-HeaderStyle $xl $ws $hdr.Count
    Set-ColumnWidths $ws @(8, 26, 30, 90, 8)

    ($book.Worksheets.Item('요약')).Activate()
    Save-Workbook $book $resultPath

    # ID 매핑 (결과 파일과 분리, HR만 보관)
    $mapBook = $xl.Workbooks.Add()
    $script:FirstSheetUsed = $false
    $ws = New-Sheet $mapBook '매핑'
    $rows = @(, @('분석용 ID', '사번')) + @($people | Sort-Object AnalysisId | ForEach-Object { , @($_.AnalysisId, $_.Key) })
    Write-GridToSheet $ws $rows 2
    Set-HeaderStyle $xl $ws 2
    Set-ColumnWidths $ws @(12, 14)
    $ws.Range($ws.Cells(2, 2), $ws.Cells($rows.Count, 2)).NumberFormat = '@'
    $ws.Cells(1, 4).Value2 = '이 파일은 HR 담당자만 보관하세요. 결과 파일과 함께 공유하면 분석용 ID가 사번으로 연결됩니다.'
    Save-Workbook $mapBook $mapPath
} finally {
    if ($mapBook) { try { $mapBook.Close($false) } catch {} }
    Stop-ExcelApp $xl $book
}

# ---------------------------------------------------------------- 콘솔 출력
Write-Host ''
Write-Host '=== 요약 3줄 ==='
$summaryLines | ForEach-Object { Write-Host $_ }
Write-Host ''
Write-Host "결과 파일: $resultPath"
Write-Host "ID 매핑 파일(HR만 보관): $mapPath"
Write-Host ''
Write-Host '=== 데이터 점검 (경고 이상) ==='
foreach ($i in @($Issues | Where-Object { $_[0] -ne '정보' })) { Write-Host ("[{0}] {1} - {2}: {3} {4}" -f $i[0], $i[1], $i[2], $i[3], $i[4]) }
Write-Host ''
Write-Host '=== 분석의 한계 ==='
$limitLines | ForEach-Object { Write-Host "- $_" }
