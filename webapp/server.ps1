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

# 일일 위험성평가의 단순 체크 항목(안전점검/위험물저장소/S-L/온열질환/순회점검) 공통 처리.
# B열에는 "{라벨} 이상 무" / "{라벨} 이상 유"를 항상 명시적으로 적고, I열("위험성평가에
# 관한 조치")에는 유(이상 있음)일 때만 사용자가 입력한 조치/코멘트를 적는다.
function Write-CheckRow($ws, [string]$bAddr, [string]$iAddr, [string]$label, $chk) {
    $suffix = if ($chk.ok) { "이상 무" } else { "이상 유" }
    Set-CellValue $ws $bAddr "$label $suffix"
    Set-CellValue $ws $iAddr (if ($chk.ok) { $null } else { [string]$chk.note })
}

# 사진대지 3칸 위치/크기 (B32:G37, H32:L37, M32:R37 병합 셀의 실제 좌표, 단위: 포인트)
$script:PhotoSlots = @(
    @{ Left = 33;     Top = 780.75; Width = 186.75; Height = 135 },
    @{ Left = 219.75; Top = 780.75; Width = 181.5;  Height = 135 },
    @{ Left = 401.25; Top = 780.75; Width = 176.25; Height = 135 }
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

            # 박스에 Width/Height를 그대로 지정하면 사진이 늘어나거나 찌그러져 보이므로,
            # 원본 픽셀 크기를 먼저 읽어서 비율을 유지한 채(contain) 박스 중앙에 들어갈
            # 최종 크기/위치를 계산한 뒤 한 번에 그 크기로 삽입한다.
            Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
            $img = [System.Drawing.Image]::FromFile($tempFile)
            $pxW = $img.Width
            $pxH = $img.Height
            $img.Dispose()
            $nativeW = $pxW * 0.75   # 96 DPI 기준 픽셀 -> 포인트 변환
            $nativeH = $pxH * 0.75
            $finalW = $slot.Width
            $finalH = $slot.Height
            $finalLeft = $slot.Left
            $finalTop = $slot.Top
            if ($nativeW -gt 0 -and $nativeH -gt 0) {
                $scale = [Math]::Min($slot.Width / $nativeW, $slot.Height / $nativeH)
                $finalW = $nativeW * $scale
                $finalH = $nativeH * $scale
                $finalLeft = $slot.Left + ($slot.Width - $finalW) / 2
                $finalTop = $slot.Top + ($slot.Height - $finalH) / 2
            }

            $lastErr = $null
            $done = $false
            for ($attempt = 1; $attempt -le 3 -and -not $done; $attempt++) {
                try {
                    $shape = $ws.Shapes.AddPicture($tempFile, $false, $true, [single]$finalLeft, [single]$finalTop, [single]$finalW, [single]$finalH)
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

# "전일 내용 불러오기" 기능용: 오늘 시트를 제외한 가장 최근 날짜 시트의 값을 읽어서
# Invoke-Submit이 받는 payload와 같은 모양으로 되돌려준다 (역파싱). 셀에 저장된 문구가
# 표준 문구(예: "일일 안전점검 결과 이상 무")와 정확히 같으면 체크 ok=true로, 다르면
# 그 문구를 note로 보고 ok=false로 되돌린다.
function Get-PrevDayInfo {
    $today = Get-Date
    $todayLabel = Get-KoreanDateLabel $today
    $sorted = Get-DateSheetsSorted
    $prevCandidates = $sorted | Where-Object { $_.Name.Trim() -ne $todayLabel.Trim() }
    if (-not $prevCandidates -or $prevCandidates.Count -eq 0) {
        return [PSCustomObject]@{ ok = $false; error = "불러올 이전 시트가 없습니다." }
    }
    $ws = $prevCandidates[0].Sheet
    $sheetName = $prevCandidates[0].Name

    $trades = @()
    for ($row = 8; $row -le 15; $row++) {
        $name = $ws.Range("A$row").Value2
        if ([string]::IsNullOrWhiteSpace([string]$name)) { continue }
        $count = $ws.Range("C$row").Value2
        $work = $ws.Range("F$row").Value2
        $trades += [PSCustomObject]@{
            name = ([string]$name).Trim()
            count = if ($null -ne $count) { [double]$count } else { 0 }
            work = if ($work) { ([string]$work).Trim() } else { "" }
        }
    }

    # B열이 "... 이상 무" / "... 이상 유"로 끝나면 새 형식(코멘트는 I열에 별도로 있음)이고,
    # 그 접미사 없이 그냥 자유 텍스트만 있으면 이 기능을 넣기 전(구 형식 - 체크 해제시 B열
    # 전체가 코멘트였음) 시트이므로 그 텍스트를 그대로 note로 본다.
    function Parse-CheckRow($ws, [string]$bAddr, [string]$iAddr) {
        $bText = [string]$ws.Range($bAddr).Value2
        $iText = [string]$ws.Range($iAddr).Value2
        $t = if ($bText) { $bText.Trim() } else { "" }
        if ($t -match '이상\s*무\s*$') { return [PSCustomObject]@{ ok = $true; note = "" } }
        if ($t -match '이상\s*유\s*$') { return [PSCustomObject]@{ ok = $false; note = if ($iText) { $iText.Trim() } else { "" } } }
        return [PSCustomObject]@{ ok = ($t -eq ""); note = $t }
    }

    $safety = Parse-CheckRow $ws "B17" "I17"
    $hazmat = Parse-CheckRow $ws "B18" "I18"
    $heat   = Parse-CheckRow $ws "B20" "I20"
    $patrol = Parse-CheckRow $ws "B22" "I22"

    $b19 = [string]$ws.Range("B19").Value2
    $slText = if ($b19) { $b19.Trim() } else { "" }
    $slMatch = [regex]::Match($slText, '^S/L\s*\((\d+)대\)\s*점검 결과 이상\s*(무|유)$')
    if ($slMatch.Success -and $slMatch.Groups[2].Value -eq "무") {
        $sl = [PSCustomObject]@{ ok = $true; note = ""; count = [int]$slMatch.Groups[1].Value }
    } elseif ($slMatch.Success) {
        $i19 = [string]$ws.Range("I19").Value2
        $sl = [PSCustomObject]@{ ok = $false; note = if ($i19) { $i19.Trim() } else { "" }; count = [int]$slMatch.Groups[1].Value }
    } else {
        $sl = [PSCustomObject]@{ ok = ($slText -eq ""); note = $slText; count = 5 }
    }

    $b21 = [string]$ws.Range("B21").Value2
    $i21 = [string]$ws.Range("I21").Value2
    $equipText = if ($b21) { $b21.Trim() } else { "" }
    $equipMatch = [regex]::Match($equipText, '^장비\s*\((.+)\)\s*점검 결과 이상\s*(무|유)$')
    $equipItems = @()
    $equipOk = $true
    $equipNote = ""
    if ($equipMatch.Success) {
        foreach ($part in ($equipMatch.Groups[1].Value -split ',\s*')) {
            $pm = [regex]::Match($part.Trim(), '^(.+?)\s+(\d+)대$')
            if ($pm.Success) {
                $equipItems += [PSCustomObject]@{ name = $pm.Groups[1].Value.Trim(); qty = [int]$pm.Groups[2].Value }
            }
        }
        $equipOk = $equipMatch.Groups[2].Value -eq "무"
        if (-not $equipOk) { $equipNote = if ($i21) { $i21.Trim() } else { "" } }
    } elseif ($equipText) {
        $equipOk = $false
        $equipNote = $equipText
    }

    $eduDaily = [string]$ws.Range("C23").Value2
    $eduItems = @()
    for ($row = 24; $row -le 29; $row++) {
        $trade = $ws.Range("B$row").Value2
        if ([string]::IsNullOrWhiteSpace([string]$trade)) { continue }
        $text = $ws.Range("C$row").Value2
        $eduItems += [PSCustomObject]@{ trade = ([string]$trade).Trim(); text = if ($text) { ([string]$text).Trim() } else { "" } }
    }

    $progressTodayRaw = [string]$ws.Range("B30").Value2
    $progressPlanRaw = [string]$ws.Range("B31").Value2
    $progressToday = if ($progressTodayRaw) { ($progressTodayRaw -replace '^금일 진행사항:\s*', '').Trim() } else { "" }
    $progressPlan = if ($progressPlanRaw) { ($progressPlanRaw -replace '^금주 및 내주 예정사항\s*:\s*', '').Trim() } else { "" }

    $prevWorkHours = $ws.Range("I6").Value2
    $excludeAm = $ws.Range("W16").Value2
    $excludePm = $ws.Range("Y16").Value2

    return [PSCustomObject]@{
        ok = $true
        sheetName = $sheetName
        trades = $trades
        checks = [PSCustomObject]@{
            safety = $safety
            hazmat = $hazmat
            sl = $sl
            heat = $heat
            patrol = $patrol
            equip = [PSCustomObject]@{ ok = $equipOk; note = $equipNote; items = $equipItems }
        }
        edu = [PSCustomObject]@{ daily = if ($eduDaily) { $eduDaily.Trim() } else { "" }; items = $eduItems }
        progressToday = $progressToday
        progressPlan = $progressPlan
        prevWorkHours = if ($null -ne $prevWorkHours) { [double]$prevWorkHours } else { 10 }
        excludeAm = if ($null -ne $excludeAm) { [double]$excludeAm } else { 0 }
        excludePm = if ($null -ne $excludePm) { [double]$excludePm } else { 0 }
    }
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
    # 전일 무재해시간은 폼에서 직접 받은 값을 그대로 스냅샷으로 저장한다 (예전에는
    # 전일 시트 C5를 실시간으로 참조하는 수식을 넣었는데, 그 전일 시트가 나중에 삭제/보관되면
    # 참조가 끊겨 =#REF! 로 깨지고 그 뒤 모든 날짜의 무재해시간이 연쇄적으로 깨지는 문제가 있었다.
    # 값으로 저장하면 그런 연쇄 손상이 없고, 격주 근무 등으로 전일 시트를 다른 사람이 작성해서
    # 자동 이월 값이 실제와 다를 때도 사용자가 직접 고쳐 넣을 수 있다.)
    Set-CellValue $ws "T20" ([double]$payload.prevAccidentFreeHours)

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

    # 일일 위험성평가 - B열에는 "{항목} 이상 무/유"를 항상 명시적으로 적고, I열("위험성평가에
    # 관한 조치")에는 유(이상 있음)일 때 사용자가 입력한 조치/코멘트를 적는다.
    # I16은 그 조치란의 표 header 칸이라 절대 덮어쓰지 않는다.
    $c = $payload.checks
    Write-CheckRow $ws "B17" "I17" "일일 안전점검 결과" $c.safety
    Write-CheckRow $ws "B18" "I18" "위험물 저장소 점검 결과" $c.hazmat
    Write-CheckRow $ws "B19" "I19" "S/L ($([int]$c.sl.count)대) 점검 결과" $c.sl
    Write-CheckRow $ws "B20" "I20" "온열질환 자율 점검 결과" $c.heat
    Write-CheckRow $ws "B22" "I22" "작업장 순회점검 결과" $c.patrol

    # 장비는 여러 대를 사용자가 실제로 쓰던 형식대로 "(장비명 1대, 장비명 1대, ...)" 한 칸에
    # 콤마로 묶어서 B21에 적고, 이상 무일 때는 I21에 "장비 Checklist 확인" 고정 문구를,
    # 이상 유일 때는 사용자가 입력한 조치/코멘트를 적는다. 장비를 하나도 안 넣었으면 비워둔다.
    $equipItems = @($c.equip.items) | Where-Object { $_ -and $_.name }
    if ($equipItems.Count -gt 0) {
        $parts = $equipItems | ForEach-Object { "$($_.name) $([int]$_.qty)대" }
        $equipSuffix = if ($c.equip.ok) { "이상 무" } else { "이상 유" }
        Set-CellValue $ws "B21" "장비 ($($parts -join ', ')) 점검 결과 $equipSuffix"
        Set-CellValue $ws "I21" (if ($c.equip.ok) { "장비 Checklist 확인" } else { [string]$c.equip.note })
    } else {
        Set-CellValue $ws "B21" $null
        Set-CellValue $ws "I21" $null
    }
    Set-CellValue $ws "S21" $null
    Set-CellValue $ws "Z21" $null

    # 안전관계기록 - 공종별 안전지시사항은 엑셀 물리 슬롯이 6개(24~29행)뿐이라
    # 폼에서 온 개수만큼(최대 6개) 라벨(B열)과 내용(C열)을 함께 채우고, 남는 칸은 비운다.
    Set-CellValue $ws "C23" ([string]$payload.edu.daily)
    $eduItems = @($payload.edu.items) | Where-Object { $_ -and $_.trade }
    $eduRowNumbers = 24..29
    for ($i = 0; $i -lt $eduRowNumbers.Count; $i++) {
        $row = $eduRowNumbers[$i]
        if ($i -lt $eduItems.Count) {
            Set-CellValue $ws "B$row" ([string]$eduItems[$i].trade)
            Set-CellValue $ws "C$row" ([string]$eduItems[$i].text)
        } else {
            Set-CellValue $ws "B$row" $null
            Set-CellValue $ws "C$row" $null
        }
    }

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
            elseif ($request.HttpMethod -eq "GET" -and $path -eq "/api/prev-day") {
                try {
                    Open-MasterForOp
                    try { Send-Json $response (Get-PrevDayInfo) }
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
