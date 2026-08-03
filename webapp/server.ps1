# AI 안전일지 로컬 서버
# - 정적 폼(index.html) 서빙
# - /api/weather   : 평택 팽성읍 당일 날씨 자동조회
# - /api/context   : 직전 시트 이름 / 전일 무재해시간 조회
# - /api/submit    : 폼 데이터를 받아 마스터 엑셀에 오늘자 시트 생성/갱신

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Weather.ps1")

$root = Split-Path -Parent $PSScriptRoot
$masterPath = Join-Path $root "안전일지(교도소 신축시설 현장) 남동현.xlsx"
$indexPath = Join-Path $PSScriptRoot "index.html"
$prefix = "http://localhost:8787/"

if (-not (Test-Path $masterPath)) { throw "마스터 파일을 찾을 수 없습니다: $masterPath" }

# ---------- Excel 연결 (서버 수명 동안 하나의 인스턴스를 유지) ----------
$script:Excel = New-Object -ComObject Excel.Application
$script:Excel.Visible = $false
$script:Excel.DisplayAlerts = $false
$script:MasterWb = $script:Excel.Workbooks.Open($masterPath)

if ($script:MasterWb.ReadOnly) {
    $script:Excel.Quit()
    throw "마스터 파일이 다른 프로그램(Excel)에서 이미 열려 있어 읽기전용으로만 열렸습니다. 엑셀에서 해당 파일을 닫은 뒤 서버를 다시 시작해주세요."
}

Write-Host "마스터 워크북 연결됨: $masterPath"

$hasTemplate = $false
foreach ($ws in $script:MasterWb.Worksheets) {
    if ($ws.Name -eq "TEMPLATE") { $hasTemplate = $true }
}
if (-not $hasTemplate) {
    $script:MasterWb.Close($false)
    $script:Excel.Quit()
    throw "마스터 워크북에 TEMPLATE 시트가 없습니다. 먼저 PrepareMasterTemplate.ps1 을 실행하세요."
}

# ---------- 유틸 ----------
function Get-KoreanDateLabel([datetime]$d) {
    return "$($d.Month)월 $($d.Day)일"
}

function Parse-SheetDateLabel([string]$name, [datetime]$ref) {
    $m = [regex]::Match($name, '(\d{1,2})월\s*(\d{1,2})일')
    if (-not $m.Success) { return $null }
    $month = [int]$m.Groups[1].Value
    $day = [int]$m.Groups[2].Value
    try {
        $candidate = Get-Date -Year $ref.Year -Month $month -Day $day
    } catch { return $null }
    # 새해가 지난 직후(1월) 작성된 지난해 12월 시트 등을 대비해, 오늘보다 60일 이상 미래면 작년으로 간주
    if (($candidate - $ref).TotalDays -gt 60) {
        $candidate = $candidate.AddYears(-1)
    }
    return $candidate
}

function Get-DateSheetsSorted {
    # 반환: @({Sheet=<worksheet COM>; Date=<datetime>; Name=<string>}, ...) 최신순 정렬
    $today = Get-Date
    $items = @()
    foreach ($ws in $script:MasterWb.Worksheets) {
        if ($ws.Name -match '분석') { continue }
        $d = Parse-SheetDateLabel -name $ws.Name -ref $today
        if ($null -ne $d) {
            $items += [PSCustomObject]@{ Sheet = $ws; Date = $d; Name = $ws.Name }
        }
    }
    return $items | Sort-Object -Property Date -Descending
}

function Clear-CutCopyMode {
    [void]$script:Excel.GetType().InvokeMember("CutCopyMode", [System.Reflection.BindingFlags]::SetProperty, $null, $script:Excel, @($false))
}

# ---------- HTTP 응답 헬퍼 ----------
function Send-Json($response, $obj, [int]$status = 200) {
    $json = $obj | ConvertTo-Json -Depth 10 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $response.StatusCode = $status
    $response.ContentType = "application/json; charset=utf-8"
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
    $response.OutputStream.Close()
}

function Send-Html($response, [string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $response.StatusCode = 200
    $response.ContentType = "text/html; charset=utf-8"
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
    $response.OutputStream.Close()
}

function Read-JsonBody($request) {
    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
    $text = $reader.ReadToEnd()
    $reader.Close()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return $text | ConvertFrom-Json
}

# ---------- 핵심 로직 ----------
function Get-ContextInfo {
    $sorted = Get-DateSheetsSorted
    if ($sorted.Count -eq 0) {
        return [PSCustomObject]@{ ok = $true; lastSheetName = "(없음)"; prevAccidentFreeHours = 0 }
    }
    $last = $sorted[0]
    $hours = $last.Sheet.Range("C5").Value2
    if ($null -eq $hours) { $hours = 0 }
    return [PSCustomObject]@{ ok = $true; lastSheetName = $last.Name; prevAccidentFreeHours = [double]$hours }
}

function Invoke-Submit($payload) {
    try { $payload | ConvertTo-Json -Depth 10 | Out-File -FilePath (Join-Path $PSScriptRoot "last_payload.json") -Encoding utf8 } catch {}
    $today = Get-Date
    $todayLabel = Get-KoreanDateLabel $today
    $sorted = Get-DateSheetsSorted

    $existing = $sorted | Where-Object { $_.Name.Trim() -eq $todayLabel.Trim() } | Select-Object -First 1
    $prevCandidates = $sorted | Where-Object { $_.Name.Trim() -ne $todayLabel.Trim() }
    $prevHours = 0
    if ($prevCandidates.Count -gt 0) {
        $v = $prevCandidates[0].Sheet.Range("C5").Value2
        if ($null -ne $v) { $prevHours = [double]$v }
    }

    $createdNew = -not $existing
    if ($existing) {
        $ws = $existing.Sheet
    }
    else {
        $tplSheet = $script:MasterWb.Worksheets.Item("TEMPLATE")
        $tplSheet.Visible = -1   # xlSheetVisible - 매우 숨김 상태에서는 Copy()가 동작하지 않음
        $tplSheet.Copy($script:MasterWb.Worksheets.Item(1), [System.Type]::Missing)   # 마스터의 맨 앞에 복사 (같은 워크북 내 이동, SaveAs 불필요)
        Clear-CutCopyMode
        $tplSheet.Visible = 2   # 원본 TEMPLATE는 다시 매우 숨김 처리
        $ws = $script:MasterWb.Worksheets.Item(1)
        $ws.Name = $todayLabel
        $ws.Visible = -1   # xlSheetVisible (TEMPLATE는 매우 숨김 상태였으므로 복사본은 보이게 전환)
    }

    try {
    # 기본 정보
    $ws.Range("C4").Value2 = [string]$payload.weather
    $ws.Range("I6").Value2 = [double]$payload.prevWorkHours
    $ws.Range("W16").Value2 = [double]$payload.excludeAm
    $ws.Range("Y16").Value2 = [double]$payload.excludePm
    $ws.Range("T20").Value2 = $prevHours

    # 공종별 인원/작업내용 (8~15행 슬롯에 순서대로 매핑)
    $row = 8
    foreach ($t in $payload.trades) {
        if ($row -gt 15) { break }
        $ws.Range("A$row").Value2 = [string]$t.name
        $ws.Range("C$row").Value2 = [double]$t.count
        $ws.Range("F$row").Value2 = [string]$t.work
        $row++
    }
    while ($row -le 15) {
        $ws.Range("A$row").Value2 = $null
        $ws.Range("C$row").Value2 = $null
        $ws.Range("F$row").Value2 = $null
        $row++
    }

    # 일일 위험성평가
    $c = $payload.checks

    [string]$v17 = if ($c.safety.ok) { "일일 안전점검 결과 이상 무" } else { [string]$c.safety.note }
    $ws.Range("B17").Value2 = $v17

    [string]$v18 = if ($c.hazmat.ok) { "위험물 저장소 점검 결과 이상 무" } else { [string]$c.hazmat.note }
    $ws.Range("B18").Value2 = $v18

    [string]$v19 = if ($c.sl.ok) { "S/L ($([int]$c.sl.count)대) 점검 결과 이상 무" } else { [string]$c.sl.note }
    $ws.Range("B19").Value2 = $v19

    [string]$v20 = if ($c.heat.ok) { "온열질환 자율 점검 결과 이상 무" } else { [string]$c.heat.note }
    $ws.Range("B20").Value2 = $v20

    if ($c.equip.ok) {
        if ($c.equip.equip1.name) {
            [string]$v21 = "장비 ($([string]$c.equip.equip1.name) $([int]$c.equip.equip1.qty)대)  점검 결과 이상 무"
            $ws.Range("B21").Value2 = $v21
        } else { $ws.Range("B21").Value2 = $null }
        if ($c.equip.equip2.name) {
            [string]$vs21 = "장비 ($([string]$c.equip.equip2.name) $([int]$c.equip.equip2.qty)대) 점검 결과 이상 무"
            $ws.Range("S21").Value2 = $vs21
        } else { $ws.Range("S21").Value2 = $null }
    } else {
        [string]$v21b = $c.equip.note
        $ws.Range("B21").Value2 = $v21b
        $ws.Range("S21").Value2 = $null
    }

    [string]$vRisk = $payload.riskAction
    $ws.Range("I16").Value2 = $vRisk

    # 안전관계기록
    [string]$vEduDaily = $payload.edu.daily
    $ws.Range("C23").Value2 = $vEduDaily
    [string]$vEduElec = $payload.edu.elec
    $ws.Range("C24").Value2 = $vEduElec
    [string]$vEduMech = $payload.edu.mech
    $ws.Range("C25").Value2 = $vEduMech
    [string]$vEduPaint = $payload.edu.paint
    $ws.Range("C26").Value2 = $vEduPaint
    [string]$vEduInterior = $payload.edu.interior
    $ws.Range("C27").Value2 = $vEduInterior
    [string]$vEduCivil = $payload.edu.civil
    $ws.Range("C28").Value2 = $vEduCivil
    [string]$vEduFence = $payload.edu.fence
    $ws.Range("C29").Value2 = $vEduFence

    [string]$vProgToday = $payload.progressToday
    $ws.Range("B30").Value2 = $vProgToday
    [string]$vProgPlan = $payload.progressPlan
    $ws.Range("B31").Value2 = $vProgPlan

    $script:MasterWb.Save()

    return [PSCustomObject]@{ ok = $true; sheetName = $ws.Name }
    }
    catch {
        if ($createdNew) {
            try {
                $script:Excel.DisplayAlerts = $false
                $ws.Delete()
            } catch {}
        }
        throw
    }
}

# ---------- HTTP 서버 ----------
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "서버 시작됨: $prefix  (종료하려면 Ctrl+C)"

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        try {
            $path = $request.Url.AbsolutePath

            if ($request.HttpMethod -eq "GET" -and ($path -eq "/" -or $path -eq "/index.html")) {
                Send-Html $response $indexPath
            }
            elseif ($request.HttpMethod -eq "GET" -and $path -eq "/api/weather") {
                try {
                    $w = Get-TodayWeather
                    Send-Json $response ([PSCustomObject]@{ ok = $true; text = $w.Text; tmax = $w.TMax; tmin = $w.TMin })
                } catch {
                    Send-Json $response ([PSCustomObject]@{ ok = $false; error = $_.Exception.Message })
                }
            }
            elseif ($request.HttpMethod -eq "GET" -and $path -eq "/api/context") {
                try {
                    Send-Json $response (Get-ContextInfo)
                } catch {
                    Send-Json $response ([PSCustomObject]@{ ok = $false; error = $_.Exception.Message })
                }
            }
            elseif ($request.HttpMethod -eq "POST" -and $path -eq "/api/submit") {
                try {
                    $payload = Read-JsonBody $request
                    $result = Invoke-Submit $payload
                    Send-Json $response $result
                } catch {
                    $detail = "$($_.Exception.Message) @ line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
                    Send-Json $response ([PSCustomObject]@{ ok = $false; error = $detail }) 500
                }
            }
            else {
                Send-Json $response ([PSCustomObject]@{ ok = $false; error = "not found" }) 404
            }
        }
        catch {
            try { Send-Json $response ([PSCustomObject]@{ ok = $false; error = $_.Exception.Message }) 500 } catch {}
        }
    }
}
finally {
    $listener.Stop()
    $script:MasterWb.Close($false)
    $script:Excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
