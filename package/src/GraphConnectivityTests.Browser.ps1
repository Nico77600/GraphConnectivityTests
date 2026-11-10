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

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAj1O8gXsivU6s5
# HpD2TJ+ksx9pa4/Ok34BF5zP32J6kKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYEDCCGAwCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBmZzCrEEUa
# TXsIFVXMHsZrifE//cBBGGsZIUElvqAf7TANBgkqhkiG9w0BAQEFAASCAgBXNEHY
# W8AB5VWrC7KIbV/qutFHkdAbO61BKWyNkMW+vr/zAmAiVVskWqx2MkHa8q+cTn4s
# AZLoFgcU+eFAVZo2RcjL7usxblFUWGX02za6CRR1Dznop/7SHtmtJD/QVzLslP/W
# qFHuXMhSWvzYbcFl7L9OEnIn5z0LcR3ghJmZnxQbSIc+t1OQdFYt35MlYzOU1V2/
# RI9GSVamLt9uke67WReMuiOCd9DElgLgyPwA1FmPIb1Ek7t6xslNws0QlrFMODBn
# sgwxlGXU3kw/x9bvleQuMqKRlCg9RWxGzFqgC7fO9kofauA95WFkem6wlVWp8FDx
# Y/rw6isn8tUsMIi6FFKiA/iBcIThdwd1b1QJKi4R7MW5iMVlQogv3gJpKVjAf/2b
# pQjD+cnWm1KtAZU7ExYgJgMZvzge/TeZGaRaBHkHIIH3MKJj2kWjuFgSqfk5uZ7A
# AKOsgBnvXfvWSqlsCvkl3OXgoJqsbewQ0TT9W2Ll/b+SEtIsCCQImB4+5VrdFHhg
# HOAdleIcCWJovZYwjaU7/gw2z+7BWjrdJspQg55y2xAcorxa2xA3IiMOIdW2oqkm
# mTTatKipBx7bthVkrYh46YLvMqeKoE7n2Z0LDCvZqREIRbuRwmZCfCyzeixa22BZ
# 2GeYGT0RQXaqlU+Cj1HRnnpXpTOA1i8REduKSKGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIJOOzr/LAWI4Y1hVdKCHNhsbxmlwc1JpdFIoCvtK7f0MAghOpDVwcVHI
# IRgPMjAyNjExMTAxMDIzMDdaMAMCAQECEGH44U50+SHsL/MJt3J3SRSgghGnMIIE
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
# wzSPOzGCAlUwggJRAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjMwN1owKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIOhH
# p04yRHr+P/ULNGFGgoC9dusKimAfQiglJXsDuHmNMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiAsHDZDuuUr5tVpIEcriHPQuQ+AGVqvD5XKF85OEOVbkwIgJf6v
# x2d4Ar4aeikhbztT1hd2WkgwyJX/Z63eB9DQ8b0=
# SIG # End signature block
