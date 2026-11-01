# 공유용 dashboard_share.html에 개인 정보가 들어갔는지 파일 전체를 검사한다. 공유하기 전에 실행한다.
# 사용: powershell -NoProfile -ExecutionPolicy Bypass -File check_share.ps1 -SharePath <dashboard_share.html> [-SourcePath <인사데이터.xlsx>]
# -SourcePath를 주면 원본의 이름·사번이 파일 안에 있는지도 확인한다 (이름·사번은 화면에 출력하지 않는다).
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SharePath,
    [string]$SourcePath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib_xlsx.ps1')

if (-not [System.IO.Path]::IsPathRooted($SharePath)) { $SharePath = Join-Path (Get-Location).Path $SharePath }
if (-not (Test-Path -LiteralPath $SharePath -PathType Leaf)) { Write-Host "파일을 찾을 수 없습니다: $SharePath"; exit 2 }
$text = [System.IO.File]::ReadAllText($SharePath, [System.Text.Encoding]::UTF8)

$findings = New-Object System.Collections.Generic.List[string]
$m = [regex]::Matches($text, 'R\d{3}')
if ($m.Count) { $findings.Add("분석용 ID 형태(R000)가 $($m.Count)건 있습니다.") }
$m = [regex]::Matches($text, '(?<![\d.])\d{7}(?![\d.])')
if ($m.Count) { $findings.Add("사번 형태(7자리 숫자)가 $($m.Count)건 있습니다.") }
$m = [regex]::Matches($text, 'https?://')
if ($m.Count) { $findings.Add("외부 주소(http/https)가 $($m.Count)건 있습니다.") }
foreach ($k in '담당자', '확인 상태', '다음 확인일', 'localStorage', 'hr-dashboard-notes', '"people"') {
    if ($text.Contains($k)) { $findings.Add("개인별 기능·자료를 가리키는 문자열 '$k'이(가) 있습니다.") }
}
if ($SourcePath) {
    if (-not [System.IO.Path]::IsPathRooted($SourcePath)) { $SourcePath = Join-Path (Get-Location).Path $SourcePath }
    $src = Read-XlsxWorkbook $SourcePath
    if ($src.Contains('v_employees_current')) {
        $nameHits = 0; $idHits = 0
        foreach ($r in $src['v_employees_current'].Rows) {
            $n = [string]$r['name']; if ($n -and $text.Contains($n)) { $nameHits++ }
            $i = $r['employee_id']; $k = if ($i -is [double]) { [string][int64]$i } else { [string]$i }
            if ($k -and $text.Contains($k)) { $idHits++ }
        }
        if ($nameHits) { $findings.Add("원본의 직원 이름이 $nameHits 건 들어 있습니다.") }
        if ($idHits) { $findings.Add("원본의 사번이 $idHits 건 들어 있습니다.") }
    } else { $findings.Add("원본에 v_employees_current 시트가 없어 이름·사번 대조는 하지 못했습니다.") }
}

Write-Host "검사 대상: $SharePath ($([math]::Round($text.Length / 1KB))KB)"
if ($findings.Count -eq 0) {
    Write-Host "결과: 통과 - 개인 식별 정보와 외부 주소가 발견되지 않았습니다.$(if (-not $SourcePath) { ' (원본을 주지 않아 이름·사번 대조는 하지 않았습니다)' })"
    exit 0
}
Write-Host '결과: 실패 - 공유하지 마세요.'
$findings | ForEach-Object { Write-Host "- $_" }
exit 1
