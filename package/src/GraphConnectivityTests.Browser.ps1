<#
.SYNOPSIS
    Graph Connectivity Tests - the sign-in window (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    The OAuth sign-in in a window, like a mail app: Microsoft Edge (or Google Chrome) opens the
    authorization page of Entra ID in an app window with a temporary profile - no account of the
    workstation, no cookie, no extension - and the user signs in there: password, MFA, Conditional
    Access. The tool drives the window through the DevTools protocol of the browser, on the loopback
    interface and a port chosen by the browser, and catches the redirect that carries the
    authorization code before the browser follows it (the native-client page of Entra ID, or
    http://localhost for Microsoft Graph Command Line Tools). The window then shows a completion page
    and closes; the profile folder is deleted.

    During the sign-in only the browser talks to Entra ID: those pages are not in the trace of the
    run. The tool records what it opened and the redirect it caught (code masked), then exchanges the
    code itself (traced).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

# Page shown in the window once the code is caught, before it closes.
$script:GctSignInDonePage = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Graph Connectivity Tests - sign-in complete</title><style>' +
    'body{font:15px "Segoe UI",sans-serif;background:#f7f4ef;color:#242424;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}' +
    'div{background:#fff;border:1px solid #dedede;border-top:4px solid #b11f4b;border-radius:12px;padding:26px 30px;max-width:360px}' +
    'h1{font-size:18px;margin:0 0 8px}p{margin:0;color:#5c5c5c;line-height:1.5}</style></head><body><div><h1>Sign-in complete</h1>' +
    '<p>Graph Connectivity Tests received the authorization code and goes on with the test. This window closes by itself.</p></div></body></html>'

# Text of the page, kept when the window shows an Entra ID error (reported if the window is closed).
$script:GctSignInErrorProbe = '(() => { const t = (document.body && document.body.innerText) || ""; const m = t.match(/(AADSTS\d+|MSIS\d{4})[^\n]{0,240}/); return m ? m[0] : ""; })()'

function Find-GctBrowser {
    <# Microsoft Edge, else Google Chrome (both speak the DevTools protocol): name and path, or $null. -All: every one installed. #>
    param([switch]$All)
    if (-not $IsWindows) { return $null }
    $found = [Collections.Generic.List[object]]::new()
    $candidates = @(
        @{ Name = 'Microsoft Edge'; Exe = 'msedge.exe'; Paths = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe", "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe") }
        @{ Name = 'Google Chrome'; Exe = 'chrome.exe'; Paths = @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe", "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe") }
    )
    foreach ($candidate in $candidates) {
        $paths = [Collections.Generic.List[string]]::new()
        foreach ($hive in 'HKLM:', 'HKCU:') {
            $item = Get-ItemProperty -LiteralPath "$hive\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$($candidate.Exe)" -ErrorAction SilentlyContinue
            $value = [string](Get-GctField $item '(default)')
            if ($value) { $paths.Add($value.Trim('"')) }
        }
        foreach ($path in $candidate.Paths) { $paths.Add($path) }
        $path = $paths | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
        if ($path) {
            $browser = [pscustomobject]@{ Name = $candidate.Name; Path = $path }
            if (-not $All) { return $browser }
            $found.Add($browser)
        }
    }
    if ($All) { return $found.ToArray() }
    return $null
}

function Test-GctBrowserPolicyBlock {
    <#
        An organisation policy forbids the DevTools protocol in this browser (RemoteDebuggingAllowed = 0,
        machine or user): the sign-in window cannot catch the answer of the server with it.
    #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Browser)

    $key = if ($Browser.Name -eq 'Google Chrome') { 'SOFTWARE\Policies\Google\Chrome' } else { 'SOFTWARE\Policies\Microsoft\Edge' }
    foreach ($hive in 'HKLM:', 'HKCU:') {
        $value = Get-GctField (Get-ItemProperty -LiteralPath "$hive\$key" -ErrorAction SilentlyContinue) 'RemoteDebuggingAllowed'
        if ($null -ne $value -and "$value" -eq '0') { return $true }
    }
    return $false
}

function Test-GctDesktopSession {
    <# A window can be shown: Windows and an interactive desktop - not a service, a task as SYSTEM or an SSH session (session 0). #>
    if (-not $IsWindows -or -not [Environment]::UserInteractive) { return $false }
    try { return (Get-Process -Id $PID -ErrorAction Stop).SessionId -ne 0 } catch { return $false }
}

function Test-GctRedirectMatch {
    <#
        The URL is the redirect URI with the answer of the server (?code=..., ?error=...). Compared
        without case and final '/': Entra ID answers com.apple.preferences://oauth-redirect/?code=...
        to com.apple.Preferences://oauth-redirect.
    #>
    param([AllowEmptyString()][AllowNull()][string]$Url, [Parameter(Mandatory = $true)][string]$RedirectUri)

    if ([string]::IsNullOrEmpty($Url)) { return $false }
    $prefix = $RedirectUri.TrimEnd('/')
    if (-not $Url.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return $false }
    $rest = $Url.Substring($prefix.Length)
    return ($rest.Length -eq 0 -or $rest.Substring(0, 1) -in '/', '?', '#')
}

function Send-GctDevTools {
    <# Sends one DevTools command (no wait). Returns its id. #>
    param([Parameter(Mandatory = $true)][hashtable]$Session, [Parameter(Mandatory = $true)][string]$Method, [hashtable]$Params = @{}, [string]$SessionId)

    $Session.NextId++
    $message = [ordered]@{ id = $Session.NextId; method = $Method; params = $Params }
    if ($SessionId) { $message.sessionId = $SessionId }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($message | ConvertTo-Json -Depth 12 -Compress))
    [void]$Session.Socket.SendAsync([ArraySegment[byte]]::new($bytes), [Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
    return $Session.NextId
}

function Receive-GctDevTools {
    <# Next message of the browser (hashtable), or $null after WaitMs without one. Throws when the browser has gone. #>
    param([Parameter(Mandatory = $true)][hashtable]$Session, [int]$WaitMs = 250)

    if ($Session.Queue.Count) { return $Session.Queue.Dequeue() }
    while ($true) {
        if (-not $Session.Pending) {
            $Session.Pending = $Session.Socket.ReceiveAsync([ArraySegment[byte]]::new($Session.Buffer), [Threading.CancellationToken]::None)
        }
        try { if (-not $Session.Pending.Wait($WaitMs)) { return $null } }
        catch { throw [InvalidOperationException]::new('The sign-in window was closed before the sign-in was completed.') }
        $result = $Session.Pending.Result
        $Session.Pending = $null
        if ($result.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close) { throw [InvalidOperationException]::new('The sign-in window was closed before the sign-in was completed.') }
        $Session.Stream.Write($Session.Buffer, 0, $result.Count)
        if ($result.EndOfMessage) {
            $text = [Text.Encoding]::UTF8.GetString($Session.Stream.ToArray())
            $Session.Stream.SetLength(0)
            return ($text | ConvertFrom-Json -AsHashtable -Depth 64)
        }
    }
}

function Invoke-GctDevTools {
    <# Sends a DevTools command and waits for its answer (events received meanwhile are kept in order). #>
    param([Parameter(Mandatory = $true)][hashtable]$Session, [Parameter(Mandatory = $true)][string]$Method, [hashtable]$Params = @{}, [string]$SessionId, [int]$TimeoutSeconds = 15)

    $id = Send-GctDevTools -Session $Session -Method $Method -Params $Params -SessionId $SessionId
    $until = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
    $held = [Collections.Generic.List[object]]::new()
    try {
        while ([DateTimeOffset]::UtcNow -lt $until) {
            Invoke-GctUiPump
            $message = Receive-GctDevTools -Session $Session -WaitMs 200
            if ($null -eq $message) { continue }
            if ($message['id'] -eq $id) {
                if ($message['error']) { throw "The browser refused $($Method): $($message['error']['message'])" }
                return $message['result']
            }
            $held.Add($message)
        }
        throw "The browser did not answer $Method within $TimeoutSeconds seconds."
    }
    finally {
        foreach ($m in $held) { $Session.Queue.Enqueue($m) }
    }
}

function Get-GctPausedRedirect {
    <# The redirect URL with the code, when a paused request or response is the answer of the server; $null otherwise. #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Params, [Parameter(Mandatory = $true)][string]$RedirectUri)

    $status = $Params['responseStatusCode']
    if ($null -eq $status) {
        $url = [string]$Params['request']['url']
        if (Test-GctRedirectMatch -Url $url -RedirectUri $RedirectUri) { return $url }
        return $null
    }
    if ([int]$status -in 301, 302, 303, 307, 308) {
        $location = @($Params['responseHeaders'] | Where-Object { $_ -and [string]$_['name'] -ieq 'location' } | ForEach-Object { [string]$_['value'] }) | Select-Object -First 1
        if (Test-GctRedirectMatch -Url $location -RedirectUri $RedirectUri) { return $location }
    }
    return $null
}

function Invoke-GctBrowserAuthorization {
    <#
        Opens Url (an authorization request, response_type=code) in the sign-in window and returns the
        redirect URL caught - RedirectUri with code and state, or with error. Throws when the window is
        closed, on timeout or cancel. The browser is always closed and its temporary profile deleted.
        When the window cannot even start (the browser closes at once, or does not open its DevTools
        endpoint - an organisation policy can forbid it: RemoteDebuggingAllowed), it throws a
        NotSupportedException: the caller can then use the device code.
    #>
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Browser,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$RedirectUri,
        [int]$TimeoutSeconds = 600,
        [int]$StartTimeoutSeconds = 20
    )

    # Profiles left by a run that was killed (the browser holds them while it runs): deleted when old.
    Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'GraphConnectivityTests-signin-*' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTimeUtc -lt [DateTime]::UtcNow.AddHours(-1) } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
    $profileDir = Join-Path ([IO.Path]::GetTempPath()) ('GraphConnectivityTests-signin-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
    [void][IO.Directory]::CreateDirectory($profileDir)
    # The window starts on a page of its own, found again among the pages the browser opens (an organisation
    # policy can add a new-tab or start page): the others are closed.
    $startHtml = "<!DOCTYPE html><html><head><meta charset=""utf-8""><title>Graph Connectivity Tests - sign-in</title></head><body style=""font:15px 'Segoe UI',sans-serif;color:#5c5c5c;padding:24px"">Opening the sign-in page...<!-- $([guid]::NewGuid().ToString('N')) --></body></html>"
    $startPage = 'data:text/html;base64,' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($startHtml))
    # App window (no address bar), new profile: no workstation account, no single sign-on, no extension.
    $arguments = @(
        "--user-data-dir=`"$profileDir`"", '--remote-debugging-port=0', '--no-first-run', '--no-default-browser-check',
        '--disable-sync', '--disable-extensions', '--window-size=520,760', "--app=$startPage"
    )
    $session = @{
        Socket = $null; NextId = 0; Buffer = [byte[]]::new(65536); Stream = [IO.MemoryStream]::new()
        Pending = $null; Queue = [Collections.Generic.Queue[object]]::new()
    }
    $process = $null
    try {
        # Start: anything that fails here means the window cannot be used on this computer (NotSupportedException).
        try {
            $process = Start-Process -FilePath $Browser.Path -ArgumentList $arguments -PassThru -ErrorAction Stop
            # The browser writes the port it listens on (loopback only) and the path of its DevTools endpoint.
            $portFile = Join-Path $profileDir 'DevToolsActivePort'
            $until = [DateTimeOffset]::UtcNow.AddSeconds($StartTimeoutSeconds)
            $lines = @()
            while ($lines.Count -lt 2) {
                if ($process.HasExited -and -not (Test-Path -LiteralPath $portFile)) { throw "$($Browser.Name) closed at once (exit code $($process.ExitCode))." }
                if ([DateTimeOffset]::UtcNow -gt $until) { throw "$($Browser.Name) did not open its DevTools endpoint within $StartTimeoutSeconds seconds: an organisation policy can forbid it (RemoteDebuggingAllowed)." }
                Wait-GctSeconds 0.2
                if (Test-Path -LiteralPath $portFile) { $lines = @(Get-Content -LiteralPath $portFile -ErrorAction SilentlyContinue | Where-Object { $_ }) }
            }
            $session.Socket = [Net.WebSockets.ClientWebSocket]::new()
            [void]$session.Socket.ConnectAsync([Uri]"ws://127.0.0.1:$($lines[0].Trim())$($lines[1].Trim())", [Threading.CancellationToken]::None).GetAwaiter().GetResult()

            $target = $null
            $until = [DateTimeOffset]::UtcNow.AddSeconds(15)
            while (-not $target) {
                $pages = @((Invoke-GctDevTools -Session $session -Method 'Target.getTargets')['targetInfos'] | Where-Object { $_['type'] -eq 'page' })
                $target = $pages | Where-Object { [string]$_['url'] -eq $startPage } | Select-Object -First 1
                if ($target) { break }
                if ([DateTimeOffset]::UtcNow -gt $until) { throw "$($Browser.Name) did not show the sign-in window (pages: $(@($pages | ForEach-Object { [string]$_['url'] -replace '^(data:[^,]{0,20}).*', '$1...' }) -join ', '))." }
                Wait-GctSeconds 0.2
            }
            foreach ($other in @($pages | Where-Object { $_['targetId'] -ne $target['targetId'] })) {
                [void](Send-GctDevTools -Session $session -Method 'Target.closeTarget' -Params @{ targetId = [string]$other['targetId'] })
            }
            $sessionId = [string](Invoke-GctDevTools -Session $session -Method 'Target.attachToTarget' -Params @{ targetId = $target['targetId']; flatten = $true })['sessionId']
        }
        catch {
            if ($_.Exception.Message -eq 'Cancelled by the operator.') { throw }
            throw [NotSupportedException]::new("The sign-in window could not start: $($_.Exception.Message)", $_.Exception)
        }
        # The answer comes back either as a redirect (302 to the redirect URI) or as a navigation
        # (Entra ID: native-client page or com.apple.Preferences://): documents paused at the response, the
        # redirect URI paused before it is requested, and every request seen.
        $patterns = [Collections.Generic.List[object]]::new()
        $patterns.Add(@{ urlPattern = '*'; resourceType = 'Document'; requestStage = 'Response' })
        if ($RedirectUri -match '^https?://') { $patterns.Add(@{ urlPattern = "$RedirectUri*"; requestStage = 'Request' }) }
        [void](Invoke-GctDevTools -Session $session -Method 'Fetch.enable' -Params @{ patterns = @($patterns) } -SessionId $sessionId)
        [void](Invoke-GctDevTools -Session $session -Method 'Network.enable' -SessionId $sessionId)
        [void](Invoke-GctDevTools -Session $session -Method 'Page.enable' -SessionId $sessionId)
        # Not awaited: the answer of Page.navigate waits for the paused response handled below.
        [void](Send-GctDevTools -Session $session -Method 'Page.navigate' -Params @{ url = $Url } -SessionId $sessionId)

        $donePage = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($script:GctSignInDonePage))
        $fulfil = { param([string]$RequestId) try { [void](Send-GctDevTools -Session $session -Method 'Fetch.fulfillRequest' -SessionId $sessionId -Params @{ requestId = $RequestId; responseCode = 200; responseHeaders = @(@{ name = 'Content-Type'; value = 'text/html; charset=utf-8' }); body = $donePage }) } catch { } }
        $deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
        $captured = $null
        $lastError = $null
        $probes = @{}
        $showDone = $false
        try {
            while (-not $captured) {
                if ([DateTimeOffset]::UtcNow -gt $deadline) { throw "No sign-in in the window within $(Format-GctDuration $TimeoutSeconds) (Test.OAuthPollTimeoutSeconds)." }
                Invoke-GctUiPump
                Assert-GctNotCancelled
                $message = Receive-GctDevTools -Session $session -WaitMs 250
                if ($null -eq $message) { continue }
                if ($null -ne $message['id']) {
                    if ($probes.ContainsKey([int]$message['id'])) {
                        $text = [string]$message['result']['result']['value']
                        if ($text) { $lastError = $text }
                    }
                    continue
                }
                $params = $message['params']
                switch ([string]$message['method']) {
                    'Fetch.requestPaused' {
                        $hit = Get-GctPausedRedirect -Params $params -RedirectUri $RedirectUri
                        if ($hit) { $captured = $hit; & $fulfil ([string]$params['requestId']) }
                        else { [void](Send-GctDevTools -Session $session -Method 'Fetch.continueRequest' -Params @{ requestId = [string]$params['requestId'] } -SessionId $sessionId) }
                    }
                    'Network.requestWillBeSent' {
                        $url = [string]$params['request']['url']
                        if (Test-GctRedirectMatch -Url $url -RedirectUri $RedirectUri) { $captured = $url; $showDone = $true }
                    }
                    'Page.loadEventFired' {
                        $id = Send-GctDevTools -Session $session -Method 'Runtime.evaluate' -Params @{ expression = $script:GctSignInErrorProbe; returnByValue = $true } -SessionId $sessionId
                        $probes[[int]$id] = $true
                    }
                    'Target.detachedFromTarget' {
                        if ([string]$params['sessionId'] -eq $sessionId) { throw [InvalidOperationException]::new('The sign-in window was closed before the sign-in was completed.') }
                    }
                    'Inspector.detached' {
                        throw [InvalidOperationException]::new('The sign-in window was closed before the sign-in was completed.')
                    }
                }
            }
        }
        catch {
            $why = $_.Exception.Message
            if ($lastError) { $why += " The window showed: $lastError" }
            throw $why
        }

        # A navigation the browser cannot follow (custom scheme): show the completion page instead.
        # The code is caught: from here the window is only cosmetic, a failure of the browser is ignored.
        try {
            if ($showDone) { [void](Send-GctDevTools -Session $session -Method 'Page.navigate' -Params @{ url = "data:text/html;base64,$donePage" } -SessionId $sessionId) }
            # Let the completion page show; requests still paused are answered so the window does not hang.
            $until = [DateTimeOffset]::UtcNow.AddSeconds(1.5)
            while ([DateTimeOffset]::UtcNow -lt $until) {
                Invoke-GctUiPump
                $message = Receive-GctDevTools -Session $session -WaitMs 150
                if ($message -and [string]$message['method'] -eq 'Fetch.requestPaused') {
                    $requestId = [string]$message['params']['requestId']
                    if (Get-GctPausedRedirect -Params $message['params'] -RedirectUri $RedirectUri) { & $fulfil $requestId }
                    else { [void](Send-GctDevTools -Session $session -Method 'Fetch.continueRequest' -Params @{ requestId = $requestId } -SessionId $sessionId) }
                }
            }
        }
        catch { Write-GctLog 'INFO' "Sign-in window: $($_.Exception.Message) (after the code was received)." }
        return $captured
    }
    finally {
        if ($session.Socket -and $session.Socket.State -eq [Net.WebSockets.WebSocketState]::Open) {
            try { [void](Send-GctDevTools -Session $session -Method 'Browser.close') } catch { }
        }
        if ($process) {
            $until = [DateTimeOffset]::UtcNow.AddSeconds(5)
            while (-not $process.HasExited -and [DateTimeOffset]::UtcNow -lt $until) { Start-Sleep -Milliseconds 100 }
            if (-not $process.HasExited) { try { Stop-Process -Id $process.Id -Force -ErrorAction Stop } catch { } }
        }
        if ($session.Socket) { $session.Socket.Dispose() }
        $session.Stream.Dispose()
        # The profile holds the cookies of the sign-in: it is deleted (the browser may hold it for a moment).
        for ($attempt = 0; $attempt -lt 20 -and (Test-Path -LiteralPath $profileDir); $attempt++) {
            try { Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction Stop } catch { Start-Sleep -Milliseconds 250 }
        }
        if (Test-Path -LiteralPath $profileDir) { Write-GctLog 'WARN' "The temporary profile of the sign-in window could not be deleted: $profileDir" }
    }
}
