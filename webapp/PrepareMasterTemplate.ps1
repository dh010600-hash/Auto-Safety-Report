# 마스터 워크북 안에 "TEMPLATE" 이라는 숨김 시트를 하나 만들어 둔다.
# (별도 파일로 SaveAs 하지 않고, 같은 파일 안에서 시트를 복제 -> 정리 -> 숨김 처리 후
#  마스터 파일 자체를 그대로 Save() 한다. 이후 서버는 매일 이 TEMPLATE 시트를
#  같은 워크북 안에서 복제만 하면 되므로 SaveAs가 전혀 필요 없다.)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$masterPath = Join-Path $root "안전일지(교도소 신축시설 현장) 남동현.xlsx"

$logPath = Join-Path $PSScriptRoot "PrepareMasterTemplate.log"
function Log($msg) {
    "$(Get-Date -Format 'HH:mm:ss')  $msg" | Out-File -FilePath $logPath -Append -Encoding utf8
}
Remove-Item $logPath -ErrorAction SilentlyContinue
Log "시작"

if (-not (Test-Path $masterPath)) { throw "마스터 파일을 찾을 수 없습니다: $masterPath" }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
Log "Excel 인스턴스 생성 완료"

try {
    $masterWb = $excel.Workbooks.Open($masterPath)   # 읽기/쓰기로 오픈
    Log "마스터 워크북 열기 완료 (읽기/쓰기)"

    # 이미 TEMPLATE 시트가 있으면 건너뜀 (재실행 대비)
    $already = $null
    foreach ($ws in $masterWb.Worksheets) {
        if ($ws.Name -eq "TEMPLATE") { $already = $ws }
    }
    if ($already) {
        Log "TEMPLATE 시트가 이미 존재함 - 종료"
        Write-Output "OK: TEMPLATE 시트가 이미 마스터 워크북에 있습니다."
    }
    else {
        $srcSheet = $masterWb.Worksheets.Item(1)
        $srcSheet.Select()
        Log "원본 시트 선택 완료: $($srcSheet.Name)"

        # 같은 워크북의 맨 끝에 복제
        $lastIndex = $masterWb.Worksheets.Count
        $srcSheet.Copy([System.Type]::Missing, $masterWb.Worksheets.Item($lastIndex))
        [void]$excel.GetType().InvokeMember("CutCopyMode", [System.Reflection.BindingFlags]::SetProperty, $null, $excel, @($false))
        $newWs = $masterWb.Worksheets.Item($masterWb.Worksheets.Count)
        Log "같은 워크북 내 시트 복제 완료: $($newWs.Name)"

        # 사진대지 영역(행 30 이상)의 이미지만 삭제, 상단 로고(행 1,4)는 유지
        for ($i = $newWs.Shapes.Count; $i -ge 1; $i--) {
            $sh = $newWs.Shapes.Item($i)
            if ($sh.TopLeftCell.Row -ge 30) { $sh.Delete() }
        }
        Log "사진 도형 삭제 완료"

        $newWs.Range("C4").Value2 = $null
        foreach ($r in 8..15) {
            $newWs.Range("C$r").Value2 = $null
            $newWs.Range("F$r").Value2 = $null
        }
        $newWs.Range("W16").Value2 = 0
        $newWs.Range("Y16").Value2 = 0
        $newWs.Range("T20").Value2 = $null
        $newWs.Range("I16").Value2 = $null
        $newWs.Range("B21").Value2 = $null
        $newWs.Range("S21").Value2 = $null
        $newWs.Range("C23").Value2 = $null
        foreach ($r in 24..29) { $newWs.Range("C$r").Value2 = $null }
        $newWs.Range("B30").Value2 = $null
        $newWs.Range("B31").Value2 = $null
        Log "가변 셀 비우기 완료"

        $newWs.Name = "TEMPLATE"
        $newWs.Visible = 2   # xlSheetVeryHidden - 엑셀 UI(시트 우클릭)에서도 숨겨짐
        Log "이름 변경 및 숨김 처리 완료"

        Log "Save 시작"
        $masterWb.Save()
        Log "Save 완료"

        Write-Output "OK: TEMPLATE 시트 생성 및 저장 완료"
    }

    $masterWb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Log "종료"
}
