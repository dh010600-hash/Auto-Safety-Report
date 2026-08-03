# 평택시 팽성읍 기준 당일 날씨 조회 (Open-Meteo, API 키 불필요)
# 좌표는 최초 지오코딩 결과를 고정값으로 사용 (매번 지오코딩 API를 부르지 않기 위함)

$script:PaengseongLat = 36.96531
$script:PaengseongLon = 127.05814

function Get-KoreanWeatherText {
    param(
        [int]$WeatherCode,
        [double]$TMax,
        [double]$TMin
    )

    $desc = switch ($WeatherCode) {
        0 { "맑음" }
        1 { "대체로 맑음" }
        2 { "구름 조금" }
        3 { "흐림" }
        45 { "안개" }
        48 { "안개(서리)" }
        51 { "약한 이슬비" }
        53 { "이슬비" }
        55 { "강한 이슬비" }
        56 { "약한 착빙성 이슬비" }
        57 { "착빙성 이슬비" }
        61 { "약한 비" }
        63 { "비" }
        65 { "강한 비" }
        66 { "약한 착빙성 비" }
        67 { "착빙성 비" }
        71 { "약한 눈" }
        73 { "눈" }
        75 { "강한 눈" }
        77 { "진눈깨비" }
        80 { "약한 소나기" }
        81 { "소나기" }
        82 { "강한 소나기" }
        85 { "약한 소낙눈" }
        86 { "소낙눈" }
        95 { "뇌우" }
        96 { "뇌우(우박 동반)" }
        99 { "강한 뇌우(우박 동반)" }
        default { "흐림" }
    }

    # 기상청 폭염특보 기준(일 최고기온): 35도 이상 경보, 33도 이상 주의보
    $suffix = ""
    if ($TMax -ge 35) { $suffix = " (폭염경보)" }
    elseif ($TMax -ge 33) { $suffix = " (폭염주의보)" }

    return "$desc$suffix"
}

function Get-TodayWeather {
    $uri = "https://api.open-meteo.com/v1/forecast?latitude=$($script:PaengseongLat)&longitude=$($script:PaengseongLon)&daily=weathercode,temperature_2m_max,temperature_2m_min&timezone=Asia%2FSeoul"
    $w = Invoke-RestMethod -Uri $uri -TimeoutSec 15
    $code = [int]$w.daily.weathercode[0]
    $tmax = [double]$w.daily.temperature_2m_max[0]
    $tmin = [double]$w.daily.temperature_2m_min[0]
    $text = Get-KoreanWeatherText -WeatherCode $code -TMax $tmax -TMin $tmin
    return [PSCustomObject]@{
        Text = $text
        TMax = $tmax
        TMin = $tmin
        Code = $code
    }
}
