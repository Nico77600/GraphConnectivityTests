<#
.SYNOPSIS
    Graph Connectivity Tests - HTTP layer and trace of the exchanges (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Every HTTP request of the tool (Microsoft Graph, Entra ID, Autodiscover, EWS, the Exchange Online
    administration API) goes through Invoke-GctHttp:
      - Send-GctHttpRequest sends it. It is the only function that touches the network for HTTP; the
        tests and the documentation tool replace it with the simulator.
      - The exchange is added to the trace of the run: the request sent and the response received
        as text (start line, headers, body), SOAP and XML indented, JSON indented, HTML summarised.
    Add-GctStep attaches the exchanges recorded since the previous check to that check: the report
    shows, for every check, what the client sent and what it received. DNS queries are added to the
    same trace (Add-GctTraceEntry, method DNS).

    Secrets never reach the trace: access, refresh and ID tokens, client assertions, device and
    authorization codes are replaced by their length; for Basic only the user name is shown; a
    Negotiate or NTLM token only by its length. Cookies are masked, except the affinity cookies of
    Exchange (they name a back-end server, they authenticate nothing).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

# Exchanges of the current run (List of objects), $null outside a run. Stage: current stage.
$script:GctTrace = $null
$script:GctTraceStage = $null
$script:GctTraceLimit = 30000
# Cookie container of the HTTP client of the run (affinity cookies sent back): only to describe the requests in the trace.
$script:GctCookies = $null
$script:GctTraceSecrets = @('access_token', 'refresh_token', 'id_token', 'device_code', 'code', 'client_secret', 'code_verifier', 'assertion', 'client_assertion', 'password')
# Cookies shown in the trace: server affinity of Exchange (on-premises and Exchange Online).
$script:GctAffinityCookies = @('X-BackEndOverrideCookie', 'X-BackEndCookie', 'X-BackEndCookie2', 'exchangecookie', 'X-OWA-CANARY-Affinity')
$script:GctReasons = @{
    200 = 'OK'; 301 = 'Moved Permanently'; 302 = 'Found'; 303 = 'See Other'; 307 = 'Temporary Redirect'; 308 = 'Permanent Redirect'
    400 = 'Bad Request'; 401 = 'Unauthorized'; 403 = 'Forbidden'; 404 = 'Not Found'; 429 = 'Too Many Requests'; 440 = 'Login Timeout'
    500 = 'Internal Server Error'; 502 = 'Bad Gateway'; 503 = 'Service Unavailable'; 504 = 'Gateway Timeout'
}
# Response headers that tell which Exchange servers answered (shown in the details of the checks).
$script:GctServerHeaders = @('X-FEServer', 'X-BEServer', 'X-CalculatedBETarget', 'X-DiagInfo', 'X-TargetBEServer', 'X-BackEndHttpStatus', 'request-id', 'client-request-id', 'X-MS-Diagnostics', 'X-ProxyBackendServerStatus', 'x-ms-ags-diagnostic')

function Send-GctHttpRequest {
    <# Sends one request, never follows a redirect. The only function of the tool that touches the network. #>
    param([Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient, [Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request)

    $response = $HttpClient.SendAsync($Request).GetAwaiter().GetResult()
    try {
        $bytes = $response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
        $headers = @{}
        $lines = [Collections.Generic.List[string]]::new()
        foreach ($set in @($response.Headers.NonValidated, $response.Content.Headers.NonValidated)) {
            foreach ($header in $set) {
                $values = @($header.Value)
                $headers[$header.Key] = $values -join ','
                foreach ($value in $values) { [void]$lines.Add("$($header.Key): $value") }
            }
        }
        [pscustomobject]@{
            StatusCode  = [int]$response.StatusCode
            Reason      = $response.ReasonPhrase
            Headers     = $headers
            HeaderLines = @($lines)
            Challenges  = @(foreach ($challenge in $response.Headers.WwwAuthenticate) { $challenge.ToString() })
            Body        = [byte[]]$bytes
        }
    }
    finally {
        $response.Dispose()
    }
}

function Get-GctNetworkErrorText {
    <# Readable reason of a request that got no answer (network or TLS error), or $null. #>
    param([Parameter(Mandatory = $true)][Exception]$Exception)

    $chain = [Collections.Generic.List[Exception]]::new()
    for ($e = $Exception; $e; $e = $e.InnerException) { $chain.Add($e) }
    $outer = $chain | Where-Object {
        $_ -is [Net.Http.HttpRequestException] -or $_ -is [IO.IOException] -or $_ -is [Net.Sockets.SocketException] -or
        $_ -is [Security.Authentication.AuthenticationException] -or $_ -is [OperationCanceledException]
    } | Select-Object -First 1
    if (-not $outer) { return $null }
    $text = ($outer.Message -replace '[,.]?\s*see inner exception\.?\s*$', '').Trim()
    if ($outer -is [OperationCanceledException]) { $text = 'No answer within the timeout (Test.HttpTimeoutSeconds).' }
    $inner = $chain[$chain.Count - 1]
    if ($inner -ne $outer -and $inner -isnot [OperationCanceledException] -and $inner -isnot [TimeoutException] -and
        -not [string]::IsNullOrWhiteSpace($inner.Message) -and -not $text.Contains($inner.Message.Trim())) {
        $text = "$($text.TrimEnd('.')): $($inner.Message.Trim())"
    }
    return $text
}

function Invoke-GctHttp {
    <#
        Sends a request (Send-GctHttpRequest) and adds the exchange to the trace. Collapse: requests
        repeated with the same answer (token polling, waiting for a message) are counted on the first one.
    #>
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient,
        [Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request,
        [string]$Note,
        [string]$Collapse,
        [string]$Label,
        [string]$Operation
    )

    $body = if ($Request.Content) { [byte[]]$Request.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult() } else { [byte[]]@() }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    try {
        $response = Send-GctHttpRequest -HttpClient $HttpClient -Request $Request
    }
    catch {
        $failure = $_.Exception
        while ($failure.InnerException) { $failure = $failure.InnerException }
        Add-GctTraceExchange -Request $Request -RequestBody $body -Failure $failure.Message -DurationMs $clock.ElapsedMilliseconds -Note $Note -Label $Label -Operation $Operation
        $text = Get-GctNetworkErrorText -Exception $_.Exception
        if ($text) { throw [Net.Http.HttpRequestException]::new($text, $_.Exception) }
        throw
    }
    Add-GctTraceExchange -Request $Request -RequestBody $body -Response $response -DurationMs $clock.ElapsedMilliseconds -Note $Note -Collapse $Collapse -Label $Label -Operation $Operation
    return $response
}

function Protect-GctCookieList {
    <# Cookie header value: the affinity cookies of Exchange shown, the others replaced by <hidden>. #>
    param([AllowEmptyString()][string]$Value, [switch]$SetCookie)

    $show = { param([string]$Name) $script:GctAffinityCookies -contains $Name.Trim() }
    if ($SetCookie) {
        $m = [regex]::Match($Value, '^\s*([^=;]+)=([^;]*)(.*)$')
        if (-not $m.Success) { return $Value }
        if (& $show $m.Groups[1].Value) { return $Value }
        return "$($m.Groups[1].Value)=<hidden, $($m.Groups[2].Value.Length) characters>$($m.Groups[3].Value)"
    }
    return [regex]::Replace($Value, '([^=;\s]+)=([^;]*)', {
            param($m)
            if (& $show $m.Groups[1].Value) { $m.Value } else { "$($m.Groups[1].Value)=<hidden>" }
        })
}

function Protect-GctHeaderValue {
    <# Header value as written to the trace: secrets replaced by their length, Windows tokens decoded, affinity cookies shown. #>
    param([Parameter(Mandatory = $true)][string]$Name, [AllowEmptyString()][string]$Value)

    if ($Name -ieq 'Authorization' -or $Name -ieq 'WWW-Authenticate') {
        $parts = $Value.Trim().Split(' ', 2)
        if ($parts.Count -lt 2 -or [string]::IsNullOrWhiteSpace($parts[1])) { return $Value }
        $scheme = $parts[0]
        if ($scheme -iin 'Negotiate', 'NTLM', 'Kerberos') {
            return "$scheme <Windows token: $($parts[1].Trim().Length) characters, never written>"
        }
        if ($Name -ieq 'WWW-Authenticate') { return $Value }
        if ($scheme -ieq 'Basic') {
            $decoded = $null
            try { $decoded = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($parts[1].Trim())) } catch { $decoded = $null }
            if ($decoded -and $decoded.Contains(':')) { return "Basic <user $($decoded.Split(':', 2)[0]), password never written>" }
            return "Basic <credentials: $($parts[1].Length) characters, never written>"
        }
        return "$scheme <access token: $($parts[1].Length) characters, never written>"
    }
    if ($Name -ieq 'Set-Cookie') { return Protect-GctCookieList -Value $Value -SetCookie }
    if ($Name -ieq 'Cookie') { return Protect-GctCookieList -Value $Value }
    return $Value
}

function Limit-GctTraceText {
    param([AllowEmptyString()][string]$Text, [int]$Limit = $script:GctTraceLimit)
    if ($Text.Length -le $Limit) { return $Text }
    return $Text.Substring(0, $Limit) + "`n... (cut: $($Text.Length) characters in total)"
}

function Format-GctXml {
    <# XML (SOAP envelope, Autodiscover) indented; the text unchanged when it is not XML. #>
    param([Parameter(Mandatory = $true)][string]$Text)

    try {
        $doc = [Xml.XmlDocument]::new()
        $doc.PreserveWhitespace = $false
        $doc.LoadXml($Text)
        $builder = [Text.StringBuilder]::new()
        $settings = [Xml.XmlWriterSettings]::new()
        $settings.Indent = $true
        $settings.IndentChars = '  '
        $settings.OmitXmlDeclaration = $true
        $writer = [Xml.XmlWriter]::Create($builder, $settings)
        try { $doc.Save($writer) } finally { $writer.Dispose() }
        return $builder.ToString()
    }
    catch {
        return $Text
    }
}

function Format-GctFormBody {
    <# application/x-www-form-urlencoded body, one field per line, URL-decoded, secrets masked. #>
    param([AllowEmptyString()][string]$Text)

    $lines = foreach ($pair in $Text.Split('&')) {
        if (-not $pair) { continue }
        $name, $value = $pair.Split('=', 2)
        $name = [Uri]::UnescapeDataString($name.Replace('+', ' '))
        $value = if ($null -ne $value) { [Uri]::UnescapeDataString($value.Replace('+', ' ')) } else { '' }
        if ($name -in $script:GctTraceSecrets) { $value = "<$($value.Length) characters, never written>" }
        "$name=$value"
    }
    return "# form fields, one per line, URL-decoded`n" + (@($lines) -join "`n")
}

function Format-GctJsonBody {
    <# JSON body indented, secrets masked. Throws if the text is not JSON. #>
    param([Parameter(Mandatory = $true)][string]$Text)

    $value = $Text | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    if ($value -is [System.Collections.IDictionary]) {
        foreach ($key in @($value.Keys)) {
            if ($key -in $script:GctTraceSecrets -and $null -ne $value[$key]) { $value[$key] = "<$($key): $(([string]$value[$key]).Length) characters, never written>" }
        }
    }
    return (ConvertTo-Json -InputObject $value -Depth 8)
}

function Format-GctHtmlSummary {
    <# An HTML page summarised: title, AD FS error (MSIS), Entra ID error (AADSTS), or the beginning of the visible text. #>
    param([AllowEmptyString()][string]$Text, [int]$Length)

    $lines = [Collections.Generic.List[string]]::new()
    [void]$lines.Add("# HTML page, $Length bytes: summary")
    $title = [regex]::Match($Text, '<title[^>]*>(.*?)</title>', 'IgnoreCase, Singleline')
    if ($title.Success) { [void]$lines.Add('Title: ' + [Net.WebUtility]::HtmlDecode($title.Groups[1].Value).Trim()) }
    $pageError = [regex]::Match($Text, '(MSIS\d{4}|AADSTS\d+)[^<"\\]{0,240}')
    if ($pageError.Success) { [void]$lines.Add('Error: ' + [Net.WebUtility]::HtmlDecode($pageError.Value).Trim()) }
    if ($Text -match 'passwordInput|userNameInput|loginForm') { [void]$lines.Add('Sign-in form: user name and password fields') }
    if (-not $pageError.Success -and $Text -notmatch 'passwordInput|userNameInput|loginForm') {
        $visible = [regex]::Replace($Text, '<(script|style)[^>]*>.*?</\1>', ' ', 'IgnoreCase, Singleline')
        $visible = [Net.WebUtility]::HtmlDecode([regex]::Replace($visible, '<[^>]+>', ' '))
        $visible = [regex]::Replace($visible, '\s+', ' ').Trim()
        if ($visible) { [void]$lines.Add('Text: ' + $(if ($visible.Length -gt 600) { $visible.Substring(0, 600) + '...' } else { $visible })) }
    }
    return $lines -join "`n"
}

function Format-GctTraceBody {
    <# Body of a request or a response as text: SOAP and XML indented, form and JSON readable, HTML summarised. #>
    param([AllowNull()][byte[]]$Body, [string]$ContentType)

    if ($null -eq $Body -or $Body.Length -eq 0) { return '' }
    $type = ([string]$ContentType).ToLowerInvariant()
    $text = [Text.Encoding]::UTF8.GetString($Body).TrimStart([char]0xFEFF)
    if ($type -like '*x-www-form-urlencoded*') { return Format-GctFormBody $text }
    if ($type -like '*json*' -or $text.TrimStart().StartsWith('{')) {
        try { return Format-GctJsonBody $text } catch { }
    }
    if ($type -like '*html*' -or $text -match '^\s*<(!doctype|html)') { return Format-GctHtmlSummary -Text $text -Length $Body.Length }
    if ($type -like '*xml*' -or $text.TrimStart().StartsWith('<')) {
        $kind = if ($text -match 'soap:Envelope|s:Envelope|Envelope xmlns') { 'SOAP' } else { 'XML' }
        return "# $kind, $($Body.Length) bytes, indented`n" + (Format-GctXml $text)
    }
    return Limit-GctTraceText $text 4000
}

function Get-GctTraceRequestText {
    <# The request as sent: start line, headers (secrets masked), blank line, body. #>
    param([Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request, [byte[]]$Body, [string]$CookieHeader)

    $uri = $Request.RequestUri
    $lines = [Collections.Generic.List[string]]::new()
    [void]$lines.Add("$($Request.Method.Method) $($uri.PathAndQuery) HTTP/1.1")
    [void]$lines.Add("Host: $($uri.Authority)")
    $contentType = $null
    foreach ($set in @($Request.Headers.NonValidated, $(if ($Request.Content) { $Request.Content.Headers.NonValidated }))) {
        if ($null -eq $set) { continue }
        foreach ($header in $set) {
            if ($header.Key -ieq 'Content-Length') { continue }
            $value = $header.Value.ToString()
            if ($header.Key -ieq 'Content-Type') { $contentType = $value }
            [void]$lines.Add("$($header.Key): $(Protect-GctHeaderValue -Name $header.Key -Value $value)")
        }
    }
    # Cookies added by the cookie container of the client (affinity): not in the request object itself.
    if ($CookieHeader) { [void]$lines.Add("Cookie: $(Protect-GctHeaderValue -Name 'Cookie' -Value $CookieHeader)") }
    if ($Request.Content) { [void]$lines.Add("Content-Length: $(if ($Body) { $Body.Length } else { 0 })") }
    $text = $lines -join "`n"
    $bodyText = Format-GctTraceBody -Body $Body -ContentType $contentType
    if (-not $bodyText -and $uri.Query.Length -gt 1 -and ($uri.Query.Split('&').Count -ge 3)) {
        $bodyText = (Format-GctFormBody $uri.Query.TrimStart('?')).Replace('# form fields, one per line', '# query parameters, one per line')
    }
    if ($bodyText) { $text += "`n`n" + $bodyText }
    return Limit-GctTraceText $text
}

function Get-GctTraceResponseText {
    <# The response as received: status line, headers (secrets masked, Windows tokens decoded), blank line, body. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Response)

    $code = [int]$Response.StatusCode
    $reason = [string](Get-GctField $Response 'Reason')
    if (-not $reason) { $reason = if ($script:GctReasons.ContainsKey($code)) { $script:GctReasons[$code] } else { '' } }
    $lines = [Collections.Generic.List[string]]::new()
    [void]$lines.Add("HTTP/1.1 $code $reason".TrimEnd())
    $headerLines = @(Get-GctField $Response 'HeaderLines' | Where-Object { $_ })
    if (-not $headerLines.Count) {
        # Response without the raw header lines (simulator): its headers, one line per challenge.
        $headerLines = @(foreach ($c in @(Get-GctField $Response 'Challenges')) { "WWW-Authenticate: $c" })
        $headers = Get-GctField $Response 'Headers'
        if ($headers) { $headerLines += @(foreach ($key in ($headers.Keys | Sort-Object)) { if ($key -ine 'WWW-Authenticate') { "$($key): $($headers[$key])" } }) }
    }
    $contentType = $null
    foreach ($line in $headerLines) {
        $name, $value = ([string]$line).Split(':', 2)
        $value = ([string]$value).Trim()
        if ($name -ieq 'Content-Type') { $contentType = $value }
        [void]$lines.Add("$($name): $(Protect-GctHeaderValue -Name $name -Value $value)")
    }
    $text = $lines -join "`n"
    $bodyText = Format-GctTraceBody -Body ([byte[]](Get-GctField $Response 'Body')) -ContentType $contentType
    if ($bodyText) { $text += "`n`n" + $bodyText }
    return Limit-GctTraceText $text
}

function Add-GctTraceEntry {
    <# Adds one exchange to the trace of the run (nothing outside a run). #>
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$Request,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Response,
        [AllowNull()][object]$StatusCode,
        [string]$Reason,
        [long]$DurationMs,
        [string]$Note,
        [string]$Collapse,
        [string]$Label,
        [string]$Operation,
        [string]$Server
    )

    if ($null -eq $script:GctTrace) { return }
    if ($Collapse -and $script:GctTrace.Count) {
        $last = $script:GctTrace[$script:GctTrace.Count - 1]
        if ($last.Collapse -eq $Collapse -and $null -eq $last.Step -and $last.StatusCode -eq $StatusCode -and $last.ResponseKey -eq [string]$Response.GetHashCode()) {
            $last.Repeated++
            $last.Note = "Sent $($last.Repeated) times with the same answer (waiting): only the first one is shown."
            return
        }
    }
    $entry = [pscustomobject]@{
        Sequence     = $script:GctTrace.Count + 1
        TimestampUtc = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        Stage        = $script:GctTraceStage
        Step         = $null
        StepName     = $null
        Method       = $Method
        Url          = $Url
        Operation    = $Operation
        Label        = $Label
        StatusCode   = $StatusCode
        Reason       = $Reason
        Server       = $Server
        DurationMs   = [int]$DurationMs
        Repeated     = 1
        Note         = $Note
        Collapse     = $Collapse
        ResponseKey  = if ($Collapse) { [string]$Response.GetHashCode() } else { $null }
        Request      = $Request
        Response     = $Response
    }
    $script:GctTrace.Add($entry)
    $status = if ($null -ne $StatusCode) { "$StatusCode $Reason".TrimEnd() } else { 'no response' }
    $what = if ($Operation) { " [$Operation]" } else { '' }
    Write-GctLog 'INFO' ("HTTP #{0} {1} {2}{3} -> {4} ({5} ms)" -f $entry.Sequence, $Method, $Url, $what, $status, $entry.DurationMs)
}

function Get-GctResponseServer {
    <# The Exchange servers named by a response: front end, back end (X-FEServer, X-BEServer, X-CalculatedBETarget, X-DiagInfo). #>
    param([AllowNull()][pscustomobject]$Response)

    $headers = Get-GctField $Response 'Headers'
    if (-not $headers) { return $null }
    $get = { param([string]$Name) foreach ($k in $headers.Keys) { if ($k -ieq $Name) { return [string]$headers[$k] } }; $null }
    $fe = & $get 'X-FEServer'
    $be = & $get 'X-BEServer'
    if (-not $be) { $be = & $get 'X-CalculatedBETarget' }
    $diag = & $get 'X-DiagInfo'
    $ags = & $get 'x-ms-ags-diagnostic'
    if ($ags) {
        # Microsoft Graph: {"ServerInfo":{"DataCenter":"France Central","Slice":"E","Ring":"3","ScaleUnit":"001","RoleInstance":"PA1PEPF..."}}
        try {
            $info = ($ags | ConvertFrom-Json -ErrorAction Stop).ServerInfo
            return ("Graph {0} ({1}, ring {2})" -f $info.RoleInstance, $info.DataCenter, $info.Ring)
        }
        catch { return "Graph $ags" }
    }
    $parts = @()
    if ($fe) { $parts += "FE $fe" }
    if ($be) { $parts += "BE $be" }
    elseif ($diag) { $parts += "server $diag" }
    return ($parts -join ' > ')
}

function Add-GctTraceExchange {
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request,
        [byte[]]$RequestBody,
        [pscustomobject]$Response,
        [string]$Failure,
        [long]$DurationMs,
        [string]$Note,
        [string]$Collapse,
        [string]$Label,
        [string]$Operation
    )

    if ($null -eq $script:GctTrace) { return }
    $cookie = $null
    if ($script:GctCookies) {
        $cookie = $script:GctCookies.GetCookieHeader($Request.RequestUri)
        if ([string]::IsNullOrEmpty($cookie)) { $cookie = $null }
    }
    $requestText = Get-GctTraceRequestText -Request $Request -Body $RequestBody -CookieHeader $cookie
    # What makes this request different from its neighbours: the credentials sent, affinity, impersonation.
    $hints = [Collections.Generic.List[string]]::new()
    if ($Label) { $hints.Add($Label) }
    $operation = $Operation
    $isService = $Request.RequestUri.AbsolutePath -match '/EWS/|/autodiscover/|/adminapi/' -or $Request.RequestUri.Host -match '^(graph\.microsoft\.|graph\.microsoft\.us|dod-graph\.microsoft\.us|microsoftgraph\.chinacloudapi\.cn)'
    if ($isService) {
        if (-not $operation -and $RequestBody -and $RequestBody.Length) {
            $m = [regex]::Match([Text.Encoding]::UTF8.GetString($RequestBody), '<soap:Body>\s*<(?:\w+:)?(\w+?)(?:Request)?[\s>/]')
            if ($m.Success) { $operation = $m.Groups[1].Value }
        }
        $auth = $Request.Headers.Authorization
        if (-not $auth) { if ($Label -notmatch 'Windows|NTLM|Negotiate') { $hints.Add('no credentials') } }
        elseif (-not $auth.Parameter) { $hints.Add("empty $($auth.Scheme) header") }
        elseif ($auth.Scheme -ieq 'Basic') { $hints.Add('user name and password (Basic)') }
        else { $hints.Add('access token') }
        $values = $null
        if ($Request.Headers.TryGetValues('X-AnchorMailbox', [ref]$values)) { $hints.Add("anchor $(@($values) -join '')") }
        if ($cookie -and $cookie -match 'X-BackEnd') { $hints.Add('affinity cookie') }
    }
    if ($RequestBody -and $RequestBody.Length -and $Request.Content -and "$($Request.Content.Headers.ContentType)" -like '*form-urlencoded*') {
        $grant = [regex]::Match([Text.Encoding]::UTF8.GetString($RequestBody), '(?:^|&)grant_type=([^&]*)')
        if ($grant.Success) { $hints.Add('grant ' + ([Uri]::UnescapeDataString($grant.Groups[1].Value) -replace '^urn:ietf:params:oauth:grant-type:', '')) }
    }
    $label = $hints -join ' + '
    if ($Response) {
        $code = [int]$Response.StatusCode
        $reason = [string](Get-GctField $Response 'Reason')
        if (-not $reason -and $script:GctReasons.ContainsKey($code)) { $reason = $script:GctReasons[$code] }
        Add-GctTraceEntry -Method $Request.Method.Method -Url $Request.RequestUri.AbsoluteUri -Request $requestText -Response (Get-GctTraceResponseText -Response $Response) `
            -StatusCode $code -Reason $reason -DurationMs $DurationMs -Note $Note -Collapse $Collapse -Label $label -Operation $operation -Server (Get-GctResponseServer -Response $Response)
    }
    else {
        Add-GctTraceEntry -Method $Request.Method.Method -Url $Request.RequestUri.AbsoluteUri -Request $requestText -Response "No response: $Failure" `
            -StatusCode $null -DurationMs $DurationMs -Note $Note -Label $label -Operation $operation
    }
}

function Complete-GctTraceStep {
    <# Attaches the exchanges recorded since the previous check to this check. Returns their numbers. #>
    param([Parameter(Mandatory = $true)][int]$Step, [Parameter(Mandatory = $true)][string]$Name)

    if ($null -eq $script:GctTrace) { return @() }
    $ids = foreach ($entry in $script:GctTrace) {
        if ($null -eq $entry.Step) { $entry.Step = $Step; $entry.StepName = $Name; $entry.Sequence }
    }
    return @($ids)
}

# SIG # Begin signature block
# MIImKwYJKoZIhvcNAQcCoIImHDCCJhgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDMT+t8QWxu71WL
# sm4Oj5UyHeIECqgWpDndpH51U5xt7KCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYEjCCGA4CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCBgI6tfzIo
# /VGHfBrtfbrWpD356gjDxCbFFuUOYYCr+DANBgkqhkiG9w0BAQEFAASCAgACBTWs
# AItmcYsZvSqMwUVZkkkGAr68qqTw69BhJH8GUNDG2n593XuJIkcYSwwM/kRSYBj1
# tG6dZn4AKFTgGmBjh/8O08wYanbRo2Tmh0Y6WzY+AisRn7OqJWwq0gV7bbBiK6IM
# m2RJQ8kWrTcxEV8lN2Fk3Ni4EVbln3a23wEqWzgIRR4IQ879gjxXAZLc7dX/vH0g
# NpBqGI24RNr8pcCnNHs/CAA/KJxXTc2y/fukJiAgMwPXLgOVk1+WswgcOw79T3PQ
# tHsy7Va9nTrqAxekTxwmtXkiCAcHmBE6FGVw1iVnq3dd9SEFdlM8v6zhZL4k36Wa
# dcwa0yHyjoqkRtimRFxxQjJBV17qm4LKMqleLwBZpMsQtJ8SpttzRXNErBpXZYKI
# MxXC8/XhWBmAlfbzJcmK4Ygh6qUqqaJ0oJNVX2Q8Aerew0RWVLfysv+ILNbvgd+p
# SOSokCLix81eetg0uuAL6VpFWjgCPpyGlJCv60josZdaYBNoCTUHXhEdSvheP2JK
# 16YYQn79epMRgbU9Lp3dqwO7vsN9WIjWl7hrxzBk3IlczA8aHqSlT+1vSZz8yHls
# 6e5F4Ym8A8z+ZNRH6JlPfPNto7e/8Dx0c2D3I2YdbSDc1CDQgSw3ktP4mklZCyJ4
# KJsl+jR68xATI4ZlPcuTIYa4Z2qigBgytx/NqaGCFM8wghTLBgorBgEEAYI3AwMB
# MYIUuzCCFLcGCSqGSIb3DQEHAqCCFKgwghSkAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIFHloDx4kU2Vo4aSTY13lyeYBE7ANrGBHicYewPwHDpvAghx55NuRtHd
# +BgPMjAyNjExMTAxMDIzNTVaMAMCAQECEB1E7PQUNvxKq07xXUbAdrCgghGnMIIE
# 8zCCAtugAwIBAgIQVEP4N71a7hgBhHK4XNqO4zANBgkqhkiG9w0BAQsFADByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyMB4XDTI2MDYxMjE3MjI0NloXDTM2MDYwOTE3MjI0NVow
# bjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9u
# MREwDwYDVQQKDAhTU0wgQ29ycDEqMCgGA1UEAwwhU1NMLmNvbSBUaW1lc3RhbXBp
# bmcgVW5pdCAyMDI2IEUxMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE6CWLuGYT
# LkiSOPMLNskOBthXsIhRSA46feOQwvSLcDaNh6vxUzBKEx1KdFBRFs7OHrnH/K0V
# A6MbDQIFpUg0PaOCAVIwggFOMB8GA1UdIwQYMBaAFHCaZcFDgKP86bFkiW7xMJwJ
# Icl8ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3Ns
# LmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVIyLmNlcjBRBgNVHSAESjBIMDwG
# DCsGAQQBgqkwAQMGATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29t
# L3JlcG9zaXRvcnkwCAYGZ4EMAQQCMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIG
# A1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmxzLnNzbC5jb20vU1NMLmNvbS10aW1l
# U3RhbXBpbmctSS1SMi5jcmwwHQYDVR0OBBYEFCesA1a6KGq6FzuDPUB/eagk7MM8
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAT3ZuzvwDnsIZTXBn
# +hjmuBXgzB12fbYxT9uYe6ZvhvpbZWKcyDQgacETkPw8k6yeqnYCYyFZ063lMPEq
# AhAD2nMG614Zh7PtHW9DWqfsgivxbPYBeN+t7UxgkYT3expbWMdPEqBtWVyHOvLc
# KvMxgJWfJjANwVkwVDfky/j7cf28+iimN8PKc7S/u8/LenRRnV2W2YfncRnkz9Cf
# s5nRDzaZguiHrrp2GxwgnNQSH4GBjHOl3PzLsnVbGoZKi4ZEjgXups4mGGOK7+GB
# 2HAlQn+OS944cHqlkH0gElt2WezoCNRpELph9qaxuBnjSEqXWWfgE6yGBwO0uxKN
# qpWVTHmAwr6SvBPofxubXIuEn9ivR4CLOuAxUSrJ41ZBg/qqG0oJjF6GkRKbxq7c
# yWSkEZYJhmKCg2Hg5fXHagOxiwbmn5JPZ1vhWCrbk3v/dTx2gvTK5+erC5H9/rTc
# ljUI776yd2rfhjfkpvyAsbtXs4D7nrnw8i9wor5+fdtRSJBVxqzLmykFftoL0Up2
# Q++tyKCUdribdcnrB4HKS3s+t3AjdcbA06DCL8GifQi9fL0yRBRNlA4o6BP+Sr75
# 5MaCNjURYZDvDyxJHBBT7cqjZIwsiR8s+6860noFbBBtdge2Iw7CPC/1Q31I27hV
# olS7B0yao3v8X2MtWy3msZlXTcEwggXdMIIDxaADAgECAgh7LJvTFoAymTANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8GA1UEAwwo
# U1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAeFw0xNjAy
# MTIxNzM5MzlaFw00MTAyMTIxNzM5MzlaMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0
# aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3Jp
# dHkgUlNBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+Q/doyt9y9Aq
# /uxnhabnLhu6d+Hj9a+k7PpKXZHEV0drGHdrdvL9k+Q9D8IWngtmw1aUnheDhc5W
# 7/IW/QBi9SIJVOhlF05BueBPRpeqG8i4bmJeabFf2yoCfvxsyvNB2O3Q6Pw/YUjt
# sAMUHRAOSxngu07shmX/NvNeZwILnYZVYf16OO3+4hkAt2+hUGJ1dDyg+sglkrRu
# eiLH+B6h47LdkTGrKx0E/6VKBDfphaQzK/3i1lU0fBmkSmjHsqjTt8qhk4jrwZe8
# jPkd2SKEJHTHBD1qqSmTzOu4W+H+XyWqNFjIwSNUnRuYEcM4nH49hmylD0CGfAL0
# XAJPKMuucZ8POsgz/hElNer8usVgPdl8GNWyqdN1eANyIso6wx/vLOUuqfqeLLZR
# Rv2vA9bqYGjqhRY2a4XpHsCz3cQk3IAqgUFtlD7I4MmBQQCeXr9/xQiYohgsQkCz
# +W84J0tOgPQ9gUfgiHzqHM61dVxRLhwrfxpyKOcAtdF0xtfkn60Hk7ZTNTX8N+TD
# 9l0WviFz3pIK+KBjaryWkmo++LxlVZve9Q2JJgT8JRqmJWnLwm3KfOJZX5es6+8u
# yLzXG1k8K8zyGciTaydjGc/86Sb4ynGbf5P+NGeETpnr/LN4CTNwumamdu0bc+sa
# pQ3EIhMglFYKTixsTrH9z5wJuqIz7YcCAwEAAaNjMGEwHQYDVR0OBBYEFN0ECQei
# 9Xp9UlMSkpXuOIAlDaZZMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU3QQJ
# B6L1en1SUxKSle44gCUNplkwDgYDVR0PAQH/BAQDAgGGMA0GCSqGSIb3DQEBCwUA
# A4ICAQAgGBGUKfsmnRweHnBh8ZVyk3EkrWiTWI4yrxuzcAP8JSt0hZA9eGr0uYul
# lzu1GJG7Hqf5QFuR+VWZrx4R0Fwdp2bjsZQHDDI5puobsHnYHZxwROOK3cT5lR+K
# OEM/AYWlR6c9RrK85SJo93uc2Cw+CiHILTOsv8WBmTF0wXVxxb6x8CNF9J1r/Blj
# naO8BMYYCyW7U4kPs4BQ3kXuRH+rlHhkmNP2KN2H2HBldPsOuRPrpw9hqTKWzN67
# 7WNMGLupQPegVG4giHF1GOp6tDRy4CMnd1y2kOqGJUCr7zMPy5+CvqIg+/a1LRrm
# woWxdA/7yGUCpFIBR91JIsG/2OtrrH7e7GMzFbcjCI/GD41BWt2OxbmP5UU/eNu6
# 0htAsf5xTT/ggaK6XrTsFeCT3QgffuFVmQsh3pOeCvvmo0m9NjD+53eyoHWXtS2B
# iBdlIPfakACfyVLMMso1fPU9D9gr1/UmbMkGNJYW6nBZGjJ5eQu2iH8PUkg9v2zY
# okQu0U63cljTiROV/kSr+NeLG26cvCygW9VqAK9fN+HV+hALmJyG5yaPzvDsbopX
# C4DjTrLAoGNhkLpVaDd0araS25+hhiK2ZScO7LafQmDkZ8K12kELxNOLYRu8+h+R
# K9dEB166KazZxenvU0ha64DxKFghzbAGVfsnP1OQcKkEHlcnuTCCBsswggSzoAMC
# AQICEFKdu8K4iYf1g4DgsOuiBjYwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9T
# U0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSBSU0EwHhcNMjYwNjEyMTY0MjE1WhcNNDEwMjEyMTczOTM5
# WjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0
# b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFt
# cGluZyBJc3N1aW5nIFJTQSBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAny+r5nNNkLwxRps2B3laUGNU9y7vWiJ5013yf0iXz65bQdwa9bjj8l/S
# keXqHGqfSh2Foae+80E8mDEf93oIwjsy2CX2cXmKxqtF8HJGYB4f6yRhKED50wkb
# OysL02IXf+/EZna44jffhhVz4OU/k52xQ/PnIg/uFhc2Wk+Sc/CHqAr+lYVkPz7L
# y5RqB9f1RKGr9JC37W1LgVdHC50dNIdsa3QXPhltYZubR3npE4Bbfi0joUW5n5O0
# ayweN45R/YVLxA0YGlp1OHVfVnsqXq8bhnw1Lc4kjJ6XOgXlOlw/a8ze57TvbduZ
# l3ehewTmvO0Aky/nqOWDS9IPPxdxdG6vprBbSuHlxNANBAbHTiFgkfZmf0dZYlT/
# AnZIoFxAq0HILPolf4OgWgET76SiWSlm2XYRIXwVjpeqG0plvXwRIQBzSiEJj7Qe
# QFnCvHx4gQMINvN3+yG/1rxhEbajGve0tTP3lX4PPOHVAE0QJp4ZzPI1YYJkhVRr
# iVVxVRU4Fsg4BPwP6I+YLmHb1tZ4RLwOFNLh/gJYDgPlChCoxOogFxXgwPZZrDyE
# RYB0aeOexW0OKacN4RRu9cqF2KKu/Ab5AL7eYDLVGUPrEEyYbIMApVKqpebc0/jD
# QniC+3rkauRAjgVc8m1ip/CQgG1Dw4Z9ulVRAY9VR6Za26LsTbUCAwEAAaOCAVEw
# ggFNMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKS
# le44gCUNplkwTQYIKwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2Vy
# dC5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY2VyMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybHMuc3NsLmNvbS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNy
# bDAdBgNVHQ4EFgQUcJplwUOAo/zpsWSJbvEwnAkhyXwwDgYDVR0PAQH/BAQDAgGG
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlFFvNMemA6iBiBYo56IKQjWhWuQaYKuo4qpf0
# KJNERAt1gz+HsA2SL3xA/V+EA8SRPNP9SvX0is+5mDsppYvpKyYGqYP36YFW0mbF
# 1F0ZtXxumCUUO/ZV5mBNKDf5Q4PY7Qb0eWAepUc0rNJr/Mvn5ZiXfJPxDb2xmcg8
# aRD652/nr+Ork2UBP0E43iJdI4LF7Vs29R3SGRLYZOxM15DI4QebsKRl0KDlf4Om
# YGCKc7umTxibkZtN1C1KWgnf9jadU/JwnySaYEBgx6/v+FVC6SGONziY7TcvI9MV
# +qYxHUPlOm6AaQP2dS4bjWXnm86ldCbUkxtOACumSfT97077JNiVa7PZ1IBkd/2P
# E1AC8gs7U+/Yk1BN9XRPIi15k/uA1xF9oeee5KGjyiGKkqIT7W8QkgUT5zeCinFK
# kcKafZMcROSVYpXtmtQqGGR2g4uPYqJ0K1hms6gJXrHt/34p8M9PrV8oNWe0Zl1E
# rDMtXndUg3qhJL4pA3ZUlo7w5lM5ml+A7ijr/3Gk2v4QRoDeE2ShmTXdHgtqo8L/
# fT119mYcETycE2NYGRXsBl3N0rzDsHRr8P3TrprqhRrX04iS0P7eQ+lDxgS3yhGn
# 3OzWGhDt9Gf2x1qe0/qdhFAVOkXxzMZz4Uat6xsAfzX0+/m8GEKW5Ojw/HY+kLvV
# wzSPOzGCAlcwggJTAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjM1NVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIJC0
# EziURse2QRpP2rycPorq6XZOAiNKjzZyT3N3gIPeMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIESDBGAiEA/a44ppnY264KxCuTKNQEVj4BPEY5j3S/Ju1WdEBvr1MCIQCf
# tThaMBOC7e3CcgJ+W2EBq30A7htG3L5gfb6R8fgFqQ==
# SIG # End signature block
