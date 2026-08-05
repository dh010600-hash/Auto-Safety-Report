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

# ---------- 안전 백업: 저장 직전 현재 상태를 보관 ----------
# (엑셀 COM 자동화 중 강제종료 등으로 파일이 손상/데이터 유실될 경우를 대비)
$backupDir = Join-Path $root "backups"
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
function Backup-MasterFile {
    $backupName = "안전일지_백업_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss")
    Copy-Item -Path $masterPath -Destination (Join-Path $backupDir $backupName) -Force
    # 오래된 백업 정리 (최근 30개만 유지 - 원본이 25MB 안팎이라 무한정 쌓이지 않도록)
    Get-ChildItem $backupDir -Filter "안전일지_백업_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 | Remove-Item -Force -ErrorAction SilentlyContinue
}
Backup-MasterFile

# ---------- Excel 연결: 요청이 있을 때만 열고, 끝나면 바로 닫는다 ----------
# (서버가 파일을 계속 붙잡고 있으면 그 사이엔 사용자가 엑셀에서 직접 파일을
#  열어 수정할 수 없으므로, 실제 처리하는 짧은 순간에만 열었다 닫는다)
$script:Excel = $null
$script:MasterWb = $null
$script:ExcelPid = $null

function Open-MasterForOp {
    # Quit() 만으로는 EXCEL.EXE 프로세스가 완전히 안 죽는 경우가 있어(COM 참조 잔류),
    # 새로 뜨는 프로세스의 PID를 직접 추적해서 나중에 확실히 종료시킨다.
    $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Start-Sleep -Milliseconds 300
    $after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $script:ExcelPid = $after | Where-Object { $before -notcontains $_ } | Select-Object -First 1

    $wb = $excel.Workbooks.Open($masterPath)
    if ($wb.ReadOnly) {
        try { $wb.Close($false) } catch {}
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        if ($script:ExcelPid) { Start-Sleep -Milliseconds 300; Stop-Process -Id $script:ExcelPid -Force -ErrorAction SilentlyContinue }
        $script:ExcelPid = $null
        throw "마스터 파일이 다른 프로그램(엑셀)에서 열려 있어 지금은 처리할 수 없습니다. 엑셀에서 파일을 닫고 다시 시도해주세요."
    }
    $script:Excel = $excel
    $script:MasterWb = $wb
}

function Close-MasterForOp {
    if ($script:MasterWb) { try { $script:MasterWb.Close($false) } catch {} }
    if ($script:Excel) {
        try { $script:Excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Excel) | Out-Null
    }
    $script:MasterWb = $null
    $script:Excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($script:ExcelPid) {
        Start-Sleep -Milliseconds 500
        if (Get-Process -Id $script:ExcelPid -ErrorAction SilentlyContinue) {
            Stop-Process -Id $script:ExcelPid -Force -ErrorAction SilentlyContinue
        }
    }
    $script:ExcelPid = $null
}

# 시작할 때 한 번만 열어서 TEMPLATE 시트 존재 여부만 확인하고 바로 닫는다.
Open-MasterForOp
$hasTemplate = $false
foreach ($ws in $script:MasterWb.Worksheets) {
    if ($ws.Name -eq "TEMPLATE") { $hasTemplate = $true }
}
Close-MasterForOp
if (-not $hasTemplate) {
    throw "마스터 워크북에 TEMPLATE 시트가 없습니다. 먼저 PrepareMasterTemplate.ps1 을 실행하세요."
}
Write-Host "마스터 파일 확인 완료: $masterPath (평소엔 닫아둠 - 필요할 때만 열었다 닫습니다)"

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

# 일주일(7일) 넘은 날짜 시트를 월별 보관 파일로 옮기고 메인 파일에서는 지운다.
# 가장 최근 시트(전일 무재해시간 연동에 필요)는 아무리 오래됐어도 절대 옮기지 않는다.
function Invoke-ArchiveOldSheets {
    $sorted = Get-DateSheetsSorted
    if ($sorted.Count -le 1) { return }
    $today = Get-Date
    $toArchive = $sorted | Select-Object -Skip 1 | Where-Object { ($today - $_.Date).TotalDays -gt 7 }
    if (-not $toArchive -or $toArchive.Count -eq 0) { return }

    foreach ($item in $toArchive) {
        $archiveName = "안전일지_보관_{0}년 {1:D2}월.xlsx" -f $item.Date.Year, $item.Date.Month
        $archivePath = Join-Path $root $archiveName
        $isNewArchive = -not (Test-Path $archivePath)
        $archiveWb = $null
        try {
            if ($isNewArchive) {
                $archiveWb = $script:Excel.Workbooks.Add()
            } else {
                $archiveWb = $script:Excel.Workbooks.Open($archivePath)
            }

            $sheetToMove = $item.Sheet
            $sheetToMove.Visible = -1
            $destAfter = $archiveWb.Worksheets.Item($archiveWb.Worksheets.Count)
            $sheetToMove.Copy([System.Type]::Missing, $destAfter)
            Clear-CutCopyMode

            if ($isNewArchive) {
                # 새로 만든 워크북의 기본 빈 시트(Sheet1 등)는 방금 옮긴 시트만 남기고 정리
                for ($i = $archiveWb.Worksheets.Count; $i -ge 1; $i--) {
                    if ($archiveWb.Worksheets.Item($i).Name -ne $sheetToMove.Name) {
                        $archiveWb.Worksheets.Item($i).Delete()
                    }
                }
                $archiveWb.SaveAs($archivePath, 51)
            } else {
                $archiveWb.Save()
            }
            $archiveWb.Close($false)
            $archiveWb = $null

            $script:MasterWb.Worksheets.Item($item.Name).Delete()
            Write-Host "보관 완료: '$($item.Name)' -> $archiveName"
        }
        finally {
            if ($archiveWb) { try { $archiveWb.Close($false) } catch {} }
        }
    }
    $script:MasterWb.Save()
}

# 복제된 시트에서는 Range.Value/.Value2 에 숫자(double)를 대입할 때 "Double을 String으로
# 캐스팅할 수 없음" 오류가 재현성 있게 발생한다 (TEMPLATE 원본 시트에서는 발생하지 않음).
# 반면 숫자를 문자열로 만들어 .Formula 로 대입하면(사용자가 셀에 직접 입력하는 것과 동일한
# 경로) 문제없이 동작하므로, 숫자 값은 Formula 경로를 쓰고 문자열/공백 값은 Value 경로를 쓴다.
function Set-CellValue($ws, [string]$addr, $value) {
    $lastErr = $null
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            if ($null -eq $value) {
                $ws.Range($addr).Value = $null
            }
            elseif ($value -is [double] -or $value -is [int] -or $value -is [int64] -or $value -is [decimal]) {
                $ws.Range($addr).Formula = $value.ToString([System.Globalization.CultureInfo]::InvariantCulture)
            }
            else {
                $ws.Range($addr).Value = $value
            }
            return
        } catch {
            $lastErr = $_
            Start-Sleep -Milliseconds (100 * $attempt)
        }
    }
    throw "셀 $addr 쓰기 실패 (4회 재시도 후): $($lastErr.Exception.Message)"
}

# 다른 시트를 참조하는 수식을 직접 넣을 때 사용 (예: 전일 무재해시간을 전날 시트의
# C5 셀에 실시간으로 연동 - 전날 값이 나중에 수정되어도 자동으로 반영됨)
function Set-CellFormula($ws, [string]$addr, [string]$formula) {
    $lastErr = $null
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            $ws.Range($addr).Formula = $formula
            return
        } catch {
            $lastErr = $_
            Start-Sleep -Milliseconds (100 * $attempt)
        }
    }
    throw "셀 $addr 수식 쓰기 실패 (4회 재시도 후): $($lastErr.Exception.Message)"
}

# 사진대지 3칸 위치/크기 (실제 사람이 손으로 붙였던 사진의 좌표를 참고한 고정값, 단위: 포인트)
$script:PhotoSlots = @(
    @{ Left = 33;    Top = 780.8; Width = 186.8; Height = 135 },
    @{ Left = 219.8; Top = 780.8; Width = 181.5; Height = 135 },
    @{ Left = 401.2; Top = 780.8; Width = 176.3; Height = 135 }
)

# data:image/...;base64,... 형태의 사진을 사진대지 슬롯에 삽입.
# 슬롯별로 이름표(AutoPhoto1~3)를 붙여, 재제출 시 우리가 넣은 사진만 교체하고
# 사용자가 엑셀에서 직접 붙여넣은 사진은 건드리지 않는다.
function Set-SheetPhotos($ws, $photos) {
    if (-not $photos) { return }
    for ($i = 0; $i -lt $photos.Count -and $i -lt $script:PhotoSlots.Count; $i++) {
        $dataUrl = $photos[$i]
        if (-not $dataUrl) { continue }
        $shapeName = "AutoPhoto$($i + 1)"
        for ($j = $ws.Shapes.Count; $j -ge 1; $j--) {
            if ($ws.Shapes.Item($j).Name -eq $shapeName) { $ws.Shapes.Item($j).Delete() }
        }
        $base64 = $dataUrl -replace '^data:image/\w+;base64,', ''
        $bytes = [System.Convert]::FromBase64String($base64)
        $tempFile = [System.IO.Path]::Combine($env:TEMP, "safety_photo_$([guid]::NewGuid().ToString('N')).jpg")
        [System.IO.File]::WriteAllBytes($tempFile, $bytes)
        try {
            $slot = $script:PhotoSlots[$i]
            $lastErr = $null
            $done = $false
            for ($attempt = 1; $attempt -le 3 -and -not $done; $attempt++) {
                try {
                    $shape = $ws.Shapes.AddPicture($tempFile, $false, $true, $slot.Left, $slot.Top, $slot.Width, $slot.Height)
                    $shape.Name = $shapeName
                    $done = $true
                } catch {
                    $lastErr = $_
                    Start-Sleep -Milliseconds (150 * $attempt)
                }
            }
            if (-not $done) { throw "사진 $($i + 1) 삽입 실패: $($lastErr.Exception.Message)" }
        } finally {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
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
    $today = Get-Date
    $todayLabel = Get-KoreanDateLabel $today
    $sorted = Get-DateSheetsSorted

    $existing = $sorted | Where-Object { $_.Name.Trim() -eq $todayLabel.Trim() } | Select-Object -First 1
    $prevCandidates = $sorted | Where-Object { $_.Name.Trim() -ne $todayLabel.Trim() }
    $prevSheetName = $null
    if ($prevCandidates.Count -gt 0) {
        $prevSheetName = $prevCandidates[0].Name
    }
    try {
        $logLine = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  today=[$todayLabel] sorted=[$($sorted.Name -join ', ')] prevSheetName=[$prevSheetName]"
        $logLine | Out-File -FilePath (Join-Path $PSScriptRoot "submit_debug.log") -Append -Encoding utf8
    } catch {}

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
    Set-CellValue $ws "C4" ([string]$payload.weather)
    Set-CellValue $ws "I6" ([double]$payload.prevWorkHours)
    Set-CellValue $ws "W16" ([double]$payload.excludeAm)
    Set-CellValue $ws "Y16" ([double]$payload.excludePm)
    if ($prevSheetName) {
        # 전일 시트의 C5(무재해달성시간)를 실시간으로 참조 - 전날 값이 나중에 수정돼도 자동 반영
        $escapedName = $prevSheetName.Trim().Replace("'", "''")
        Set-CellFormula $ws "T20" "='$escapedName'!C5"
    } else {
        Set-CellValue $ws "T20" 0
    }

    # 공종별 인원/작업내용 (8~15행 슬롯에 순서대로 매핑)
    $row = 8
    foreach ($t in $payload.trades) {
        if ($row -gt 15) { break }
        Set-CellValue $ws "A$row" ([string]$t.name)
        Set-CellValue $ws "C$row" ([double]$t.count)
        Set-CellValue $ws "F$row" ([string]$t.work)
        $row++
    }
    while ($row -le 15) {
        Set-CellValue $ws "A$row" $null
        Set-CellValue $ws "C$row" $null
        Set-CellValue $ws "F$row" $null
        $row++
    }

    # 일일 위험성평가
    $c = $payload.checks

    [string]$v17 = if ($c.safety.ok) { "일일 안전점검 결과 이상 무" } else { [string]$c.safety.note }
    Set-CellValue $ws "B17" $v17

    [string]$v18 = if ($c.hazmat.ok) { "위험물 저장소 점검 결과 이상 무" } else { [string]$c.hazmat.note }
    Set-CellValue $ws "B18" $v18

    [string]$v19 = if ($c.sl.ok) { "S/L ($([int]$c.sl.count)대) 점검 결과 이상 무" } else { [string]$c.sl.note }
    Set-CellValue $ws "B19" $v19

    [string]$v20 = if ($c.heat.ok) { "온열질환 자율 점검 결과 이상 무" } else { [string]$c.heat.note }
    Set-CellValue $ws "B20" $v20

    # 장비 여러 대는 사용자가 실제로 쓰던 형식대로 "(장비명 1대, 장비명 1대, ...)" 한 칸에 콤마로 묶는다.
    if ($c.equip.ok) {
        $equipItems = @($c.equip.items) | Where-Object { $_ -and $_.name }
        if ($equipItems.Count -gt 0) {
            $parts = $equipItems | ForEach-Object { "$($_.name) $([int]$_.qty)대" }
            [string]$v21 = "장비 ($($parts -join ', ')) 점검 결과 이상 무"
            Set-CellValue $ws "B21" $v21
        } else {
            Set-CellValue $ws "B21" $null
        }
        Set-CellValue $ws "S21" $null
    } else {
        [string]$v21b = $c.equip.note
        Set-CellValue $ws "B21" $v21b
        Set-CellValue $ws "S21" $null
    }

    [string]$vRisk = $payload.riskAction
    Set-CellValue $ws "I16" $vRisk

    # 안전관계기록
    Set-CellValue $ws "C23" ([string]$payload.edu.daily)
    Set-CellValue $ws "C24" ([string]$payload.edu.elec)
    Set-CellValue $ws "C25" ([string]$payload.edu.mech)
    Set-CellValue $ws "C26" ([string]$payload.edu.paint)
    Set-CellValue $ws "C27" ([string]$payload.edu.interior)
    Set-CellValue $ws "C28" ([string]$payload.edu.civil)
    Set-CellValue $ws "C29" ([string]$payload.edu.fence)

    # 금일 진행사항 / 금주·내주 예정사항을 같은 스타일 셀에 그대로 적으면 구분이
    # 안 되므로, 내용 앞에 항목명을 자동으로 붙여준다.
    [string]$progressTodayRaw = $payload.progressToday
    [string]$progressPlanRaw = $payload.progressPlan
    $progressTodayText = if ($progressTodayRaw.Trim()) { "금일 진행사항: $($progressTodayRaw.Trim())" } else { "" }
    $progressPlanText = if ($progressPlanRaw.Trim()) { "금주 및 내주 예정사항 : $($progressPlanRaw.Trim())" } else { "" }
    Set-CellValue $ws "B30" $progressTodayText
    Set-CellValue $ws "B31" $progressPlanText

    Set-SheetPhotos $ws $payload.photos

    Backup-MasterFile
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

# 서버 시작할 때마다 일주일 지난 시트를 자동으로 보관 파일로 정리
try {
    Open-MasterForOp
    try { Invoke-ArchiveOldSheets }
    finally { Close-MasterForOp }
} catch {
    Write-Host "시트 자동 보관 중 오류 (무시하고 계속 진행): $($_.Exception.Message)"
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
                    Open-MasterForOp
                    try { Send-Json $response (Get-ContextInfo) }
                    finally { Close-MasterForOp }
                } catch {
                    Send-Json $response ([PSCustomObject]@{ ok = $false; error = $_.Exception.Message })
                }
            }
            elseif ($request.HttpMethod -eq "POST" -and $path -eq "/api/submit") {
                try {
                    $payload = Read-JsonBody $request
                    Open-MasterForOp
                    try {
                        $result = Invoke-Submit $payload
                        Send-Json $response $result
                    }
                    finally { Close-MasterForOp }
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
    Close-MasterForOp
}
