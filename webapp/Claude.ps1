# Anthropic Claude API 연동 (선택 기능)
# - webapp/anthropic_api_key.txt 에 키가 있을 때만 동작하고, 없거나 호출이 실패하면
#   호출한 쪽(server.ps1)이 알아서 기존 로컬 규칙 기반 방식으로 대체한다.

function Get-AnthropicApiKey {
    $keyPath = Join-Path $PSScriptRoot "anthropic_api_key.txt"
    if (-not (Test-Path $keyPath)) { return $null }
    $lines = Get-Content -Path $keyPath -Encoding UTF8 | Where-Object { $_ -and -not $_.TrimStart().StartsWith("#") }
    $key = $lines | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($key)) { return $null }
    return $key.Trim()
}

$script:AnthropicApiKey = Get-AnthropicApiKey
$script:ClaudeModel = "claude-haiku-4-5-20251001"

# system/user 프롬프트를 보내고 응답 텍스트(문자열)를 그대로 돌려준다.
# 키가 없거나 호출이 실패하면 예외를 던진다 - 호출부에서 try/catch로 로컬 방식 대체 처리.
function Invoke-ClaudeMessages([string]$systemPrompt, [string]$userText, [int]$maxTokens = 1024) {
    if (-not $script:AnthropicApiKey) {
        throw "Anthropic API 키가 설정되어 있지 않습니다 (webapp/anthropic_api_key.txt)."
    }
    $body = @{
        model = $script:ClaudeModel
        max_tokens = $maxTokens
        system = $systemPrompt
        messages = @(@{ role = "user"; content = $userText })
    } | ConvertTo-Json -Depth 10 -Compress
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)

    # Windows PowerShell 5.1의 Invoke-RestMethod는 응답의 charset=utf-8을 무시하고 한글을
    # 깨뜨리는 경우가 있어(실제 확인함), HttpWebRequest로 직접 호출하고 응답 스트림을
    # UTF-8로 명시해서 읽는다.
    $req = [System.Net.HttpWebRequest]::Create("https://api.anthropic.com/v1/messages")
    $req.Method = "POST"
    $req.ContentType = "application/json"
    $req.Timeout = 30000
    $req.Headers.Add("x-api-key", $script:AnthropicApiKey)
    $req.Headers.Add("anthropic-version", "2023-06-01")
    $req.ContentLength = $bodyBytes.Length
    $reqStream = $req.GetRequestStream()
    $reqStream.Write($bodyBytes, 0, $bodyBytes.Length)
    $reqStream.Close()

    try {
        $webResp = $req.GetResponse()
    } catch [System.Net.WebException] {
        $errResp = $_.Exception.Response
        if ($errResp) {
            $reader = New-Object System.IO.StreamReader($errResp.GetResponseStream(), [System.Text.Encoding]::UTF8)
            $errText = $reader.ReadToEnd()
            $reader.Close()
            throw "Claude API 오류: $errText"
        }
        throw
    }
    $reader = New-Object System.IO.StreamReader($webResp.GetResponseStream(), [System.Text.Encoding]::UTF8)
    $respText = $reader.ReadToEnd()
    $reader.Close()
    $webResp.Close()

    $respObj = $respText | ConvertFrom-Json
    return $respObj.content[0].text
}

# 응답이 ```json ... ``` 코드블록으로 감싸져 오는 경우가 있어 벗겨내고 JSON으로 파싱한다.
function ConvertFrom-ClaudeJson([string]$text) {
    $t = $text.Trim()
    $t = $t -replace '^```(?:json)?\s*', ''
    $t = $t -replace '\s*```$', ''
    return $t | ConvertFrom-Json
}

# 카카오톡 붙여넣기 텍스트를 분석해서 공종별 인원/작업내용/장비 구조로 변환.
# 반환 형태는 index.html의 parseKakaoText()가 반환하는 것과 동일하게 맞춰서,
# 클라이언트의 applyParsedTrades()를 그대로 재사용할 수 있게 한다.
function Get-KakaoParseFromAI([string]$text) {
    $systemPrompt = @"
너는 건설 현장 카카오톡 단체방에 올라오는 공종별 예정작업사항 메시지를 구조화된 JSON으로 변환하는 도우미다.
반드시 JSON 객체 하나만 출력하고 다른 설명은 절대 쓰지 마라.

스키마:
{"trades": [{"label": "공종명", "count": 숫자, "work": "작업내용 요약(콤마로 구분)", "equipMentions": [{"name":"장비명","qty":숫자}]}]}

규칙:
- 인식되는 팀/공정은 아래 표준 공종명 중 가장 가까운 것으로 label을 맞춰라 (표준 공종명과 무관한 새 팀/회사명이면 원문 그대로 사용):
  직  원 / 직  영, 전기/ESS, 기계설비, 수  장, 도  장, 토  목, 펜  스, 주방기구
- "포비엠"은 수  장으로, "대한ENG"/"FCU"는 기계설비로, "아이디스"는 주방기구로 매핑해라.
- "직원"과 "직영" 메시지는 둘 다 "직  원 / 직  영" 하나로 합쳐서 인원을 더해라.
- "총인원"/"총 인원" 같은 합계 줄은 다른 항목과 중복되므로 인원 집계에 포함하지 마라.
- 장비는 "OO 1대" 형태로 언급된 것만 equipMentions에 넣어라 (숫자 없이 장비명만 언급된 경우 1대로 간주).
- 장비 항목이 "없음"이면(장비명이 아니라 "장비 없다"는 뜻) equipMentions에 아무것도 넣지 마라. "없음"을 장비명으로 착각해서 1대로 넣지 마라.
- "T/L"이라는 장비는 equipMentions에서 제외해라 (언급되어도 넣지 않는다).
- work가 "없음"/"작업없음"이면 빈 문자열로 둬라.
"@
    $raw = Invoke-ClaudeMessages $systemPrompt $text 1500
    $parsed = ConvertFrom-ClaudeJson $raw
    return $parsed
}

# 공종별 작업내용(+ 날씨)을 보고 안전관계기록(일일 교육 + 공종별 안전지시사항)을 생성.
function Get-SafetySuggestionsFromAI($trades, [string]$weather) {
    $systemPrompt = @"
너는 건설 현장 안전관리자를 돕는 도우미다. 아래로 주어지는 공종별 오늘 작업내용을 보고,
산업안전보건기준에 관한 규칙 등 한국 산업안전보건법령에 근거한 구체적이고 간결한 한국어
안전지시사항을 공종별로 하나씩 작성해라.

반드시 JSON 객체 하나만 출력하고 다른 설명은 절대 쓰지 마라.
스키마: {"daily": "현장 전체 대상 일일 안전교육 한 줄", "items": [{"trade":"공종명","text":"그 공종 안전지시사항"}]}

규칙:
- items는 입력으로 받은 공종 목록과 순서/이름이 정확히 같아야 한다.
- 작업내용이 비어있는 공종은 text를 "금일 작업 X"로 써라.
- 날씨에 "폭염주의보"가 있으면 매시간 10분, "폭염경보"가 있으면 매시간 15분 물·그늘·휴식 수칙을
  주로 외부(토목/펜스 등) 작업 공종의 안전지시사항에 포함시켜라.
- 문구는 실제 현장에서 쓰는 간결한 문장으로 작성하고, 공종별로 서로 다른 항목을 나열하듯 콤마로 이어써라.
"@
    $userPayload = @{ weather = $weather; trades = $trades } | ConvertTo-Json -Depth 10 -Compress
    $raw = Invoke-ClaudeMessages $systemPrompt $userPayload 1500
    $parsed = ConvertFrom-ClaudeJson $raw
    return $parsed
}
