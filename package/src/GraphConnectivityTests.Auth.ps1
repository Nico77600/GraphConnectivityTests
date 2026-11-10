<#
.SYNOPSIS
    Graph Connectivity Tests - sign-in and tokens (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    The user signs in once, like a mail app, and the tool keeps the tokens in memory for the run:

      User         Microsoft Office (default, no application to register, no consent), Microsoft Graph
                   Command Line Tools, or your own multi-tenant application (Identity.ClientId). Sign-in
                   window (Edge or Chrome, temporary profile: password, MFA, Conditional Access) with the
                   authorization code and PKCE, or a device code. The refresh token gives the other
                   tokens of the run without a new prompt: EWS on-premises with hybrid modern
                   authentication (the on-premises URL is a service principal name of Office 365 Exchange
                   Online, same client).
      Application  your application with a certificate (client credentials): for automation, the same
                   checks on /users/{source} instead of /me.
      Admin        the optional administrator checks (Test.AdminChecks) read the hybrid configuration
                   through the Exchange Online administration API with the client of Connect-ExchangeOnline:
                   one more sign-in, of an Exchange administrator.

    Access, refresh and ID tokens never leave the memory: not in the console, the log, the report.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

function Get-GctAuthority {
    <# Tenant segment of the sign-in: Identity.TenantId, else the tenant found for the source domain, else organizations. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    if ([string]$Context.Config.TenantId) { return [string]$Context.Config.TenantId }
    if ($Context.SourceTenantId) { return [string]$Context.SourceTenantId }
    return 'organizations'
}

function Get-GctOAuthEndpoints {
    param([Parameter(Mandatory = $true)][hashtable]$Context, [string]$Tenant)
    if (-not $Tenant) { $Tenant = Get-GctAuthority -Context $Context }
    $root = "https://$($Context.Cloud.Login)/$([Uri]::EscapeDataString($Tenant))/oauth2/v2.0"
    [pscustomobject]@{ Authorize = "$root/authorize"; Token = "$root/token"; DeviceCode = "$root/devicecode"; Tenant = $Tenant }
}

function Get-GctScope {
    <#
        Scope of a token request. Resource: https://graph.microsoft.com, https://<on-premises host>, https://outlook.office365.com.
        Explicit clients (Graph Command Line Tools, Identity.Scopes) ask for named permissions of Graph; the others
        <resource>/.default (what the client is pre-authorized or consented for). offline_access gives the refresh token.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][pscustomobject]$Client, [Parameter(Mandatory = $true)][string]$Resource, [switch]$NoRefresh)

    $scopes = [Collections.Generic.List[string]]::new()
    $named = @($Context.Config.Scopes | Where-Object { $_ })
    $isGraph = $Resource -match [regex]::Escape($Context.Cloud.Graph)
    if ($isGraph -and ($named.Count -or $Client.Explicit)) {
        if (-not $named.Count) { $named = @($script:DefaultScopes) + @(if ($Context.Config.AllowWrite) { 'MailboxSettings.ReadWrite' }) }
        foreach ($s in $named) { $scopes.Add($(if ($s -match '^https?://') { $s } else { "$Resource/$s" })) }
    }
    elseif ($Resource -notmatch [regex]::Escape($Context.Cloud.Graph) -and $Resource -notmatch 'outlook\.office' -and $Client.Key -eq 'MicrosoftOffice') {
        # EWS on-premises with hybrid modern authentication: the delegated permission of Exchange, like Outlook.
        $scopes.Add("$Resource/EWS.AccessAsUser.All")
    }
    else { $scopes.Add("$Resource/.default") }
    if (-not $NoRefresh) { foreach ($s in 'offline_access', 'openid', 'profile') { $scopes.Add($s) } }
    return ($scopes -join ' ')
}

function Get-GctSignInMode {
    <#
        How this run signs in a user: 'Window' (authorization code in the sign-in window) or 'DeviceCode'.
        Test.SignIn Auto: the window when this session can show one, the device code otherwise (Reason says why).
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Configuration)

    $wanted = if ([string]$Configuration.SignIn) { [string]$Configuration.SignIn } else { 'Auto' }
    if ($wanted -eq 'DeviceCode') { return [pscustomobject]@{ Mode = 'DeviceCode'; Browser = $null; Reason = $null } }
    $why = $null
    $browser = $null
    if (-not (Test-GctDesktopSession)) { $why = 'no interactive desktop in this session (service, scheduled task, SSH)' }
    else {
        $browsers = @(Find-GctBrowser -All | Where-Object { $_ })
        $browser = $browsers | Where-Object { -not (Test-GctBrowserPolicyBlock -Browser $_) } | Select-Object -First 1
        if (-not $browsers) { $why = 'neither Microsoft Edge nor Google Chrome is installed' }
        elseif (-not $browser) { $why = "an organisation policy forbids the DevTools protocol in $(@($browsers.Name) -join ' and ') (RemoteDebuggingAllowed = 0)" }
    }
    if ($why -and $wanted -eq 'Window') { throw "The sign-in window cannot be opened: $why. Run with -SignIn DeviceCode." }
    if ($why) { return [pscustomobject]@{ Mode = 'DeviceCode'; Browser = $null; Reason = $why } }
    [pscustomobject]@{ Mode = 'Window'; Browser = $browser; Reason = $null }
}

function Get-GctOAuthError {
    <# 'error - description. Cause: hint' of an OAuth error answer. #>
    param([AllowNull()][object]$Json, [string]$Resource)
    $code = [string](Get-GctField $Json 'error')
    $first = ([string](Get-GctField $Json 'error_description') -split "`r?`n")[0]
    $hint = Get-GctEntraErrorHint -Description $first -Resource $Resource
    $text = "$code - $first".Trim(' ', '-')
    if ($hint) { $text += " Cause: $hint" }
    return $text
}

function ConvertFrom-GctTokenAnswer {
    param([Parameter(Mandatory = $true)][pscustomobject]$Answer, [Parameter(Mandatory = $true)][string]$Flow, [string]$Resource)
    $access = [string](Get-GctField $Answer.Json 'access_token')
    if (-not $access) { throw "Entra ID returned a token answer without access_token." }
    [pscustomobject]@{
        AccessToken  = $access
        RefreshToken = [string](Get-GctField $Answer.Json 'refresh_token')
        Scope        = [string](Get-GctField $Answer.Json 'scope')
        Flow         = $Flow
        Resource     = $Resource
        Claims       = Get-GctTokenClaims -AccessToken $access
    }
}

function Invoke-GctWindowSignIn {
    <# Sign-in in the window (authorization code with PKCE), then the code exchanged for the tokens, traced. #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][pscustomobject]$Browser,
        [Parameter(Mandatory = $true)][pscustomobject]$Client,
        [Parameter(Mandatory = $true)][string]$Scope,
        [Parameter(Mandatory = $true)][string]$Resource,
        [string]$Hint
    )

    $ep = Get-GctOAuthEndpoints -Context $Context
    $verifier = ConvertTo-GctBase64Url ([Security.Cryptography.RandomNumberGenerator]::GetBytes(32))
    $challenge = ConvertTo-GctBase64Url ([Security.Cryptography.SHA256]::HashData([Text.Encoding]::ASCII.GetBytes($verifier)))
    $state = [guid]::NewGuid().ToString('N')
    $query = [ordered]@{
        response_type = 'code'; client_id = $Client.Id; redirect_uri = $Client.RedirectUri; scope = $Scope; state = $state
        code_challenge = $challenge; code_challenge_method = 'S256'; prompt = 'select_account'
    }
    if ($Hint) { $query.login_hint = $Hint }
    $url = $ep.Authorize + '?' + (@($query.Keys | ForEach-Object { '{0}={1}' -f $_, [Uri]::EscapeDataString([string]$query[$_]) }) -join '&')
    $timeout = [int]$Context.Config.OAuthPollTimeoutSeconds
    Write-GctItem Info ('Sign-in window ({0}, temporary profile): sign in as {1} - password, then MFA if asked. The window closes by itself.' -f $Browser.Name, $(if ($Hint) { $Hint } else { 'the account of the test' })) -Icon Key
    Write-GctItem Info ('Waiting for the sign-in in the window (up to {0}).' -f (Format-GctDuration $timeout)) -Icon Clock
    $opened = "GET $url`n`nOpened in the sign-in window: $($Browser.Name) in app mode, temporary profile deleted afterwards.`nThe sign-in pages (password, MFA) are exchanged by the browser and are not recorded."
    $clock = [Diagnostics.Stopwatch]::StartNew()
    try {
        $redirect = Invoke-GctBrowserAuthorization -Browser $Browser -Url $url -RedirectUri $Client.RedirectUri -TimeoutSeconds $timeout
    }
    catch {
        Add-GctTraceEntry -Method 'GET' -Url $url -Request $opened -Response "No authorization code: $($_.Exception.Message)" -StatusCode $null -DurationMs $clock.ElapsedMilliseconds -Label 'sign-in window'
        if ($_.Exception -is [NotSupportedException]) { throw }
        $text = $_.Exception.Message
        $hint = Get-GctEntraErrorHint -Description $text -Resource $Resource
        throw "The sign-in in the window did not complete: $text$(if ($hint) { " Cause: $hint" })"
    }
    $masked = [regex]::Replace($redirect, '([?&#]code=)([^&#]+)', { param($m) $m.Groups[1].Value + "<$($m.Groups[2].Value.Length) characters, never written>" })
    Add-GctTraceEntry -Method 'GET' -Url $url -Request $opened -Response "Redirect caught by the tool (not followed by the browser):`n$masked" -StatusCode 302 -Reason 'Redirect caught' -DurationMs $clock.ElapsedMilliseconds -Label 'sign-in window'
    $fields = Get-GctUrlFields -Url $redirect
    if ($fields['error']) { throw ('Entra ID answered the sign-in with the error {0}' -f (Get-GctOAuthError -Json $fields -Resource $Resource)) }
    if ([string]$fields['state'] -ne $state) { throw 'Entra ID returned an authorization code with another state: it does not answer this sign-in.' }
    $code = [string]$fields['code']
    if (-not $code) { throw "Entra ID redirected to $($Client.RedirectUri) without an authorization code." }
    $answer = Invoke-GctFormPost -HttpClient $Context.HttpClient -Uri $ep.Token -UserAgent ([string]$Context.Config.UserAgent) -Fields ([ordered]@{
            grant_type = 'authorization_code'; client_id = $Client.Id; code = $code; redirect_uri = $Client.RedirectUri; code_verifier = $verifier; scope = $Scope })
    if ($answer.StatusCode -ne 200) { throw ('Entra ID refused the authorization code (HTTP {0}): {1}' -f $answer.StatusCode, (Get-GctOAuthError -Json $answer.Json -Resource $Resource)) }
    return ConvertFrom-GctTokenAnswer -Answer $answer -Flow 'authorization code with PKCE (sign-in window)' -Resource $Resource
}

function Invoke-GctDeviceCodeSignIn {
    <# Device-code flow (RFC 8628), every request traced; the code is shown in the console and the window. #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][pscustomobject]$Client,
        [Parameter(Mandatory = $true)][string]$Scope,
        [Parameter(Mandatory = $true)][string]$Resource
    )

    $ep = Get-GctOAuthEndpoints -Context $Context
    $agent = [string]$Context.Config.UserAgent
    $requested = Invoke-GctFormPost -HttpClient $Context.HttpClient -Uri $ep.DeviceCode -UserAgent $agent -Fields ([ordered]@{ client_id = $Client.Id; scope = $Scope })
    if ($requested.StatusCode -ne 200) { throw ('Entra ID refused the device-code request (HTTP {0}): {1}' -f $requested.StatusCode, (Get-GctOAuthError -Json $requested.Json -Resource $Resource)) }
    $device = $requested.Json
    $deviceCode = [string](Get-GctField $device 'device_code')
    if (-not $deviceCode) { throw 'Entra ID did not return a device code.' }
    $expiresIn = 900
    [void][int]::TryParse([string](Get-GctField $device 'expires_in'), [ref]$expiresIn)
    $interval = 5
    $serverInterval = 0
    if ([int]::TryParse([string](Get-GctField $device 'interval'), [ref]$serverInterval) -and $serverInterval -gt 0) { $interval = [Math]::Max($serverInterval, 5) }
    $page = [string](Get-GctField $device 'verification_uri')
    $userCode = [string](Get-GctField $device 'user_code')
    Write-GctItem Info ("Code: {0}  -  page: {1}" -f $userCode, $page) -Icon Key
    Write-GctItem Info 'Sign in with the account of the test: if the browser is already signed in with another account, open the page in a private window.' -Icon Key
    if ($script:Ui -and $script:Ui.DeviceCode) { & $script:Ui.DeviceCode $userCode $page }
    $limit = [Math]::Min($expiresIn, [int]$Context.Config.OAuthPollTimeoutSeconds)
    Write-GctItem Info ('Waiting for the sign-in in the browser (up to {0}).' -f (Format-GctDuration $limit)) -Icon Clock
    try { Start-Process -FilePath $page -ErrorAction Stop | Out-Null } catch { Write-GctItem Info "No browser could be opened here: open $page on any device and enter the code." -Icon Key }
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($limit)
    while ([DateTimeOffset]::UtcNow -lt $deadline) {
        Wait-GctSeconds $interval
        $answer = Invoke-GctFormPost -HttpClient $Context.HttpClient -Uri $ep.Token -UserAgent $agent -Collapse 'TokenPoll' -Fields ([ordered]@{
                grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; client_id = $Client.Id; device_code = $deviceCode })
        if ($answer.StatusCode -eq 200) { return ConvertFrom-GctTokenAnswer -Answer $answer -Flow 'device code' -Resource $Resource }
        $code = [string](Get-GctField $answer.Json 'error')
        if ($code -eq 'authorization_pending') { continue }
        if ($code -eq 'slow_down') { $interval += 5; continue }
        if ($code -in 'access_denied', 'authorization_declined') { throw 'The sign-in was denied in the browser.' }
        if ($code -eq 'expired_token') { throw 'The device code expired before the sign-in was completed.' }
        throw ('Entra ID token request failed: {0}' -f (Get-GctOAuthError -Json $answer.Json -Resource $Resource))
    }
    throw 'No sign-in before the timeout (Test.OAuthPollTimeoutSeconds).'
}

function Invoke-GctUserSignIn {
    <#
        Signs a user in for one resource: the window when possible (falls back to the device code when the
        window cannot start or the redirect URI is refused), else the device code. Returns the token object.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][pscustomobject]$Client,
        [Parameter(Mandatory = $true)][string]$Resource,
        [string]$Hint
    )

    $scope = Get-GctScope -Context $Context -Client $Client -Resource $Resource
    $mode = Get-GctSignInMode -Configuration $Context.Config
    if ($mode.Mode -eq 'DeviceCode' -and $mode.Reason) { Write-GctItem Info "No sign-in window here ($($mode.Reason)): sign-in with a device code, on any device." -Icon Key }
    if ($mode.Mode -eq 'Window') {
        try {
            $token = Invoke-GctWindowSignIn -Context $Context -Browser $mode.Browser -Client $Client -Scope $scope -Resource $Resource -Hint $Hint
            $token | Add-Member -NotePropertyName Browser -NotePropertyValue $mode.Browser.Name
            return $token
        }
        catch [NotSupportedException] {
            if ([string]$Context.Config.SignIn -eq 'Window') { throw "$($_.Exception.Message) Run with -SignIn DeviceCode." }
            Write-GctItem Info "$($_.Exception.Message) Sign-in with a device code instead." -Icon Key
        }
        catch {
            if ($_.Exception.Message -notmatch 'AADSTS50011|AADSTS500113' -or [string]$Context.Config.SignIn -eq 'Window') { throw }
            Write-GctItem Warn "The client $($Client.Name) does not accept the redirect URI $($Client.RedirectUri): sign-in with a device code instead (no redirect URI)."
        }
    }
    return Invoke-GctDeviceCodeSignIn -Context $Context -Client $Client -Scope $scope -Resource $Resource
}

function Get-GctTokenFromRefresh {
    <# Another token of the run for another resource, from the refresh token of the sign-in (no prompt). #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][pscustomobject]$Client,
        [Parameter(Mandatory = $true)][string]$RefreshToken,
        [Parameter(Mandatory = $true)][string]$Resource
    )
    $ep = Get-GctOAuthEndpoints -Context $Context
    $scope = Get-GctScope -Context $Context -Client $Client -Resource $Resource
    $answer = Invoke-GctFormPost -HttpClient $Context.HttpClient -Uri $ep.Token -UserAgent ([string]$Context.Config.UserAgent) -Fields ([ordered]@{
            grant_type = 'refresh_token'; client_id = $Client.Id; refresh_token = $RefreshToken; scope = $scope })
    if ($answer.StatusCode -ne 200) { throw ('Entra ID refused a token for {0} (HTTP {1}): {2}' -f $Resource, $answer.StatusCode, (Get-GctOAuthError -Json $answer.Json -Resource $Resource)) }
    return ConvertFrom-GctTokenAnswer -Answer $answer -Flow 'refresh token of the sign-in' -Resource $Resource
}

function Get-GctApplicationCertificate {
    <# Certificate of the application with its private key: Cert:\CurrentUser\My, then Cert:\LocalMachine\My. #>
    param([Parameter(Mandatory = $true)][string]$Thumbprint)
    foreach ($store in 'Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') {
        $cert = Get-Item -LiteralPath (Join-Path $store $Thumbprint) -ErrorAction SilentlyContinue
        if ($cert) {
            if (-not $cert.HasPrivateKey) { throw "The certificate $Thumbprint ($store) has no private key: the application cannot sign its assertion." }
            return [pscustomobject]@{ Certificate = $cert; Store = $store }
        }
    }
    throw "Certificate $Thumbprint not found in Cert:\CurrentUser\My nor Cert:\LocalMachine\My (Identity.CertificateThumbprint)."
}

function New-GctClientAssertion {
    <# JWT signed with the certificate of the application (RS256, x5t), audience the token endpoint, valid 10 minutes. #>
    param([Parameter(Mandatory = $true)][Security.Cryptography.X509Certificates.X509Certificate2]$Certificate, [Parameter(Mandatory = $true)][string]$ClientId, [Parameter(Mandatory = $true)][string]$Audience)

    $now = [DateTimeOffset]::UtcNow
    $header = [ordered]@{ alg = 'RS256'; typ = 'JWT'; x5t = ConvertTo-GctBase64Url $Certificate.GetCertHash() }
    $payload = [ordered]@{ aud = $Audience; iss = $ClientId; sub = $ClientId; jti = [guid]::NewGuid().ToString(); nbf = $now.AddSeconds(-30).ToUnixTimeSeconds(); exp = $now.AddMinutes(10).ToUnixTimeSeconds() }
    $data = '{0}.{1}' -f (ConvertTo-GctBase64Url ([Text.Encoding]::UTF8.GetBytes(($header | ConvertTo-Json -Compress)))), (ConvertTo-GctBase64Url ([Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Compress))))
    $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if (-not $rsa) { throw "The certificate $($Certificate.Thumbprint) has no RSA private key usable here." }
    try { $signature = $rsa.SignData([Text.Encoding]::ASCII.GetBytes($data), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1) }
    finally { $rsa.Dispose() }
    return '{0}.{1}' -f $data, (ConvertTo-GctBase64Url $signature)
}

function Invoke-GctApplicationSignIn {
    <# Client credentials with the certificate of the application, scope <resource>/.default. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Resource)

    $cfg = $Context.Config
    $ep = Get-GctOAuthEndpoints -Context $Context -Tenant ([string]$cfg.TenantId)
    $found = Get-GctApplicationCertificate -Thumbprint ([string]$cfg.CertificateThumbprint)
    $answer = Invoke-GctFormPost -HttpClient $Context.HttpClient -Uri $ep.Token -UserAgent ([string]$cfg.UserAgent) -Fields ([ordered]@{
            grant_type = 'client_credentials'; client_id = [string]$cfg.ClientId; scope = "$Resource/.default"
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion = New-GctClientAssertion -Certificate $found.Certificate -ClientId ([string]$cfg.ClientId) -Audience $ep.Token })
    if ($answer.StatusCode -ne 200) { throw ('Entra ID refused the application sign-in (HTTP {0}): {1}' -f $answer.StatusCode, (Get-GctOAuthError -Json $answer.Json -Resource $Resource)) }
    $token = ConvertFrom-GctTokenAnswer -Answer $answer -Flow 'client credentials (certificate)' -Resource $Resource
    $token | Add-Member -NotePropertyName Certificate -NotePropertyValue ("{0} ({1}, {2}, valid until {3:yyyy-MM-dd})" -f $found.Certificate.Thumbprint, $found.Store, $found.Certificate.Subject, $found.Certificate.NotAfter)
    return $token
}

function Get-GctTokenSummary {
    <# What a token says, for the report: audience, permissions, user, client, tenant, expiry. Never the token itself. #>
    param([AllowNull()][hashtable]$Claims, [DateTimeOffset]$Now = [DateTimeOffset]::UtcNow)
    # Same keys when the token is not readable: the callers read .TenantId and .User (strict mode).
    if ($null -eq $Claims) { return [ordered]@{ Readable = $false; Audience = $null; Scope = $null; Roles = $null; User = $null; ObjectId = $null; ClientId = $null; AppName = $null; TenantId = $null; ExpiresUtc = $null; MinutesLeft = $null } }
    $exp = 0L
    $expires = if ([long]::TryParse([string]$Claims['exp'], [ref]$exp)) { [DateTimeOffset]::FromUnixTimeSeconds($exp) } else { $null }
    [ordered]@{
        Audience    = @($Claims['aud']) -join ', '
        Scope       = [string]$Claims['scp']
        Roles       = @($Claims['roles'] | Where-Object { $_ }) -join ', '
        User        = $(foreach ($n in 'upn', 'unique_name', 'preferred_username') { if ($Claims[$n]) { [string]$Claims[$n]; break } })
        ObjectId    = [string]$Claims['oid']
        ClientId    = $(if ($Claims['appid']) { [string]$Claims['appid'] } else { [string]$Claims['azp'] })
        AppName     = [string]$Claims['app_displayname']
        TenantId    = [string]$Claims['tid']
        ExpiresUtc  = if ($expires) { $expires.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { $null }
        MinutesLeft = if ($expires) { [int][Math]::Floor(($expires - $Now).TotalMinutes) } else { $null }
    }
}

function Test-GctGraphToken {
    <#
        The Graph token can read free/busy: audience Microsoft Graph, a calendar permission (scp for a user,
        roles for an application), tenant of the source, not expired. Returns Status, Message, Details.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [AllowNull()][hashtable]$Claims)

    $d = Get-GctTokenSummary -Claims $Claims
    if ($null -eq $Claims) { return [pscustomobject]@{ Status = 'Warning'; Message = 'The access token is not a readable JWT: its claims were not checked.'; Details = $d } }
    $app = [string]$Context.Config.Context -eq 'Application'
    $granted = @(if ($app) { @($Claims['roles']) } else { ([string]$Claims['scp']) -split '\s+' }) | Where-Object { $_ }
    $issues = [Collections.Generic.List[string]]::new()
    if ($d.Audience -notmatch [regex]::Escape($Context.Cloud.Graph) -and $d.Audience -ne $script:GraphAppId) { $issues.Add("the audience is $($d.Audience), not Microsoft Graph") }
    $calendar = @($granted | Where-Object { $_ -in $script:CalendarScopes })
    if (-not $calendar.Count) {
        $issues.Add("no calendar permission in the token ($(if ($app) { 'roles' } else { 'scp' }): $(if ($granted) { $granted -join ' ' } else { 'none' })): getSchedule will be refused (403). $(if ($app) { 'Grant the application permission Calendars.Read (or Calendars.ReadBasic.All) with admin consent.' } else { 'Use the client MicrosoftOffice, or consent Calendars.Read to the client.' })")
    }
    if ($Context.SourceTenantId -and $d.TenantId -and $d.TenantId -ine $Context.SourceTenantId) { $issues.Add("the token comes from the tenant $($d.TenantId), not from the tenant of the source mailbox $($Context.SourceTenantId): sign in with the source account") }
    if (-not $app -and $d.User -and $d.User -ine [string]$Context.Config.SourceMailbox) {
        $Context.SignedInUser = $d.User
        $note = " Signed in as $($d.User): the UPN can differ from the SMTP address; if it is another person, /me is that person, not $($Context.Config.SourceMailbox)."
    }
    else { $note = '' }
    if ($null -ne $d.MinutesLeft -and $d.MinutesLeft -lt 0) { return [pscustomobject]@{ Status = 'Failed'; Message = "The token expired at $($d.ExpiresUtc)."; Details = $d } }
    if ($issues.Count) { return [pscustomobject]@{ Status = 'Warning'; Message = ('Token received, but ' + ($issues -join '; ') + '.' + $note); Details = $d } }
    return [pscustomobject]@{ Status = 'Passed'; Message = "Microsoft Graph token with $($calendar -join ', ') (valid for $($d.MinutesLeft) min).$note"; Details = $d }
}

function Invoke-GctStageSignIn {
    <# Sign-in of the source user (or of the application) for Microsoft Graph, and the claims of the token. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'SignIn'
    $cfg = $Context.Config
    $graph = "https://$($Context.Cloud.Graph)"
    if ($Context.Tokens['Graph']) {
        Add-GctStep $Context $stage 'Access token' Passed 'Access token supplied by the caller: sign-in skipped.' ([ordered]@{ Source = 'Caller' })
    }
    elseif ([string]$cfg.Context -eq 'Application') {
        $token = Invoke-GctApplicationSignIn -Context $Context -Resource $graph
        $Context.Tokens.Graph = $token
        $Context.SignIn = 'ClientCredentials'
        Add-GctStep $Context $stage 'Application sign-in' Passed "Token received for the application $($cfg.ClientId) (client credentials, certificate $($token.Certificate)): the checks use /users/$($cfg.SourceMailbox) instead of /me." ([ordered]@{
                Flow = $token.Flow; ClientId = $cfg.ClientId; Tenant = $cfg.TenantId; Scope = "$graph/.default" })
    }
    else {
        $client = $Context.Client
        $token = Invoke-GctUserSignIn -Context $Context -Client $client -Resource $graph -Hint ([string]$cfg.SourceMailbox)
        $Context.Tokens.Graph = $token
        $Context.SignIn = $token.Flow
        $why = if ($client.Key -eq 'MicrosoftOffice') { 'a public client of Microsoft present in every tenant, pre-authorized for Microsoft Graph: no application to register, no consent' } elseif ($client.Known) { 'a public client of Microsoft' } else { 'your application' }
        Add-GctStep $Context $stage 'User sign-in' Passed "Token received for $($cfg.SourceMailbox) with $($client.Name) ($why), $($token.Flow)." ([ordered]@{
                Flow = $token.Flow; Browser = Get-GctField $token 'Browser'; ClientId = $client.Id; Client = $client.Name; RedirectUri = $client.RedirectUri
                Authority = (Get-GctOAuthEndpoints -Context $Context).Authorize; ScopeAsked = Get-GctScope -Context $Context -Client $client -Resource $graph; RefreshToken = [bool]$token.RefreshToken })
    }
    $check = Test-GctGraphToken -Context $Context -Claims $Context.Tokens.Graph.Claims
    $Context.Token = $check.Details
    if ($check.Details.TenantId -and -not $Context.SourceTenantId) { $Context.SourceTenantId = $check.Details.TenantId }
    Add-GctStep $Context $stage 'Token claims' $check.Status $check.Message $check.Details
    if ($check.Status -eq 'Failed') { $Context.Stop = $true }
}

function Get-GctOnPremisesToken {
    <#
        Token for EWS on-premises with hybrid modern authentication: resource https://<host> (a service principal
        name of Office 365 Exchange Online), scope EWS.AccessAsUser.All, from the refresh token of the sign-in.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$HostName)

    $key = "EWS:$HostName"
    if ($Context.Tokens[$key]) { return $Context.Tokens[$key] }
    $graph = $Context.Tokens['Graph']
    if (-not $graph -or -not $graph.RefreshToken) { throw 'No refresh token from the sign-in: the token for EWS on-premises cannot be asked without a new sign-in.' }
    $token = Get-GctTokenFromRefresh -Context $Context -Client $Context.Client -RefreshToken $graph.RefreshToken -Resource "https://$HostName"
    $Context.Tokens[$key] = $token
    return $token
}

function Get-GctAdminToken {
    <# Token of an Exchange administrator for the Exchange Online administration API (one more sign-in, client of Connect-ExchangeOnline). #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    if ($Context.Tokens['Admin']) { return $Context.Tokens['Admin'] }
    $resource = "https://$($Context.Cloud.Outlook)"
    if ([string]$Context.Config.Context -eq 'Application') {
        $token = Invoke-GctApplicationSignIn -Context $Context -Resource $resource
    }
    else {
        $client = Resolve-GctClient -ClientId ([string]$Context.Config.AdminClientId)
        Write-GctItem Info "Administrator checks: sign in as an Exchange administrator (View-Only Organization Management is enough) with $($client.Name), the client of Connect-ExchangeOnline." -Icon Key
        $token = Invoke-GctUserSignIn -Context $Context -Client $client -Resource $resource -Hint $null
    }
    $Context.Tokens.Admin = $token
    return $token
}

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAAufE4zILm1CSP
# cEmcxCbDa/2i2wsXnoLbto17ETkRYKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYETCCGA0CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCWCyTuNMSs
# tv/qvzog6LeHF8zYRzQnH0QnAkK2gFQ9OzANBgkqhkiG9w0BAQEFAASCAgCVsAC/
# cFYazFV4aLV9cTlV687Obv0RtlYvqg0vYaVCNwcq3WmAv3NK6ydonoZyrCv9tQm5
# D1P+gBTGCADYFDz8huWsDfUHVd1CgTYUWrK1B74jRy7EbztvPCIRrUyfb9oiJXFS
# 1GrW5b/rkf56Cj2c5krkETfiHBQKpXw4oosmOZaf7SIiPfpNhrwR+BKvNLUTAPml
# TWbNOnrkphLhMrdFErRI+JGP/c69kysrQk68EBhHkVces9iEeGFzfLHA8UU+nvMN
# InN4oNKBPzHnETgEYd/mdAiwU0I9pDjY1W1afTEzACVUtnFeaqLgS527vKOZlhhk
# AOeuQhtijqRReJwbCQyX0PTb3A6RtZuvtfq3+3qUkUngM0mXkzO3hh4PHNH8V/Na
# E5Gy2LLugEdD3vFFgZzpDVSR5LdGxnoe5Mfr2B0gUpRInLTpnCipkDgtMk8aWuU6
# 34fXfxV2YhOi++UEuzebefcKFuvTRFGkz8C1zO/7zkYV4+yMtGkL4OYXIOFJncgf
# vkyae0Ga0I+wS+dHa1bcsDba9t24k5JwNVaUAn7AacSnWVFHqnjjbiH6rQhaonGW
# Ev/GYTtLfNDE1AXC9fTE6UUFA4oIJcSd079dOglUFDEUExzq9wyYrw2IIajATWsm
# qpOEvTozfW8ywwZR3TrjEcrDcbXedwU3AEUvL6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIDdn+AUaGZgkg+FjG5w/2kxelip+cqHmmUTFvIXWyxdUAghr68KtJN3i
# /BgPMjAyNjExMTAxMDIzMDRaMAMCAQECEF1Iqgtu4yMMQ1RldgGRsdygghGnMIIE
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
# wzSPOzGCAlYwggJSAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjMwNFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEID7f
# ltYirAQ40vHz/Xo547Wb0UWQQtkaW761L/DOVrnWMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEAnWJ8V3FWG39Bsv0aGHV2WGUcvejUddtVrdAeA5F7HSMCIAXx
# 9sy+sM7HKpY1qGwA0wT3TEcy6wYfghGxG0ivBFBb
# SIG # End signature block
