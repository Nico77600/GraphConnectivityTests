<#
.SYNOPSIS
    Graph Connectivity Tests - helpers and OAuth building blocks (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Shared with Web Services Client for Exchange (same author, same behaviour): object helpers, waits
    that keep the window responsive, traced GET and form POST, TLS certificate of a server,
    WWW-Authenticate challenges, authorization server named by Exchange, the Entra ID errors met most
    often, token claims, Entra ID tenant and user realm.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

function Get-GctField {
    <# Property of an object or key of a dictionary; $null when absent (safe under strict mode). #>
    param([AllowNull()][object]$Object, [Parameter(Mandatory = $true)][string]$Name)

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function ConvertTo-GctBase64Url {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)
    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function ConvertFrom-GctBase64Url {
    param([Parameter(Mandatory = $true)][string]$Text)
    $value = $Text.Replace('-', '+').Replace('_', '/')
    switch ($value.Length % 4) { 2 { $value += '==' } 3 { $value += '=' } }
    return [Convert]::FromBase64String($value)
}

function Invoke-GctUiPump {
    <# Keeps the GUI responsive during long waits (no effect on the command line). #>
    if ($script:Ui -and $script:Ui.Pump) { & $script:Ui.Pump }
}

function Assert-GctNotCancelled {
    if ($script:Ui -and $script:Ui.Cancel) { throw 'Cancelled by the operator.' }
}

function Wait-GctSeconds {
    param([Parameter(Mandatory = $true)][double]$Seconds)
    $until = [DateTimeOffset]::UtcNow.AddSeconds($Seconds)
    while ([DateTimeOffset]::UtcNow -lt $until) {
        Invoke-GctUiPump
        Assert-GctNotCancelled
        Start-Sleep -Milliseconds 200
    }
}

function Invoke-GctWebRequest {
    <#
        GET without credentials and without following redirects (Autodiscover, AD FS pages), traced.
        Returns StatusCode, Location and Content (text).
    #>
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient,
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$UserAgent,
        [hashtable]$Headers
    )

    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Get, $Uri)
    try {
        if ($UserAgent) { $null = $request.Headers.TryAddWithoutValidation('User-Agent', $UserAgent) }
        if ($Headers) { foreach ($name in $Headers.Keys) { $null = $request.Headers.TryAddWithoutValidation([string]$name, [string]$Headers[$name]) } }
        $response = Invoke-GctHttp -HttpClient $HttpClient -Request $request
        [pscustomobject]@{
            StatusCode = $response.StatusCode
            Location   = if ($response.Headers -and $response.Headers.ContainsKey('Location')) { [string]$response.Headers['Location'] } else { $null }
            Content    = if ($response.Body) { [Text.Encoding]::UTF8.GetString([byte[]]$response.Body) } else { '' }
        }
    }
    finally {
        $request.Dispose()
    }
}

function Invoke-GctFormPost {
    <# POST of form fields to AD FS (device code, token), traced. Returns StatusCode and the JSON answer (or $null). #>
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$HttpClient,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Fields,
        [string]$UserAgent,
        [string]$Collapse
    )

    $pairs = [Collections.Generic.List[Collections.Generic.KeyValuePair[string, string]]]::new()
    foreach ($key in $Fields.Keys) { $pairs.Add([Collections.Generic.KeyValuePair[string, string]]::new([string]$key, [string]$Fields[$key])) }
    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Post, $Uri)
    try {
        $request.Content = [Net.Http.FormUrlEncodedContent]::new($pairs)
        $null = $request.Headers.TryAddWithoutValidation('User-Agent', $(if ($UserAgent) { $UserAgent } else { $script:GctUserAgent }))
        $null = $request.Headers.TryAddWithoutValidation('Accept', 'application/json')
        $response = Invoke-GctHttp -HttpClient $HttpClient -Request $request -Collapse $Collapse
        $json = $null
        if ($response.Body -and $response.Body.Length) {
            try { $json = [Text.Encoding]::UTF8.GetString([byte[]]$response.Body) | ConvertFrom-Json -ErrorAction Stop } catch { $json = $null }
        }
        [pscustomobject]@{ StatusCode = $response.StatusCode; Json = $json }
    }
    finally {
        $request.Dispose()
    }
}

function Get-GctTlsCertificate {
    <#
        Server certificate of HostName:Port with the default Windows validation. Reachable = $false
        when no direct TCP connection is possible (proxy, firewall); Valid = $false when the TLS
        handshake fails: Interrupted = $true when the connection was closed before the server sent a
        certificate (firewall, NSG, proxy or VPN client on the path), $false when the certificate was
        received and rejected.
    #>
    param([Parameter(Mandatory = $true)][string]$HostName, [int]$Port = 443, [int]$TimeoutSeconds = 10)

    $result = [ordered]@{ HostName = $HostName; Port = $Port; Reachable = $false; Valid = $false; Interrupted = $false; Subject = $null; Issuer = $null; NotAfterUtc = $null; DaysLeft = $null; Protocol = $null; Error = $null }
    $tcp = [Net.Sockets.TcpClient]::new()
    try {
        try {
            if (-not $tcp.ConnectAsync($HostName, $Port).Wait([TimeSpan]::FromSeconds($TimeoutSeconds))) {
                $result.Error = "no TCP connection within $TimeoutSeconds s"
                return [pscustomobject]$result
            }
        }
        catch {
            $inner = $_.Exception; while ($inner.InnerException) { $inner = $inner.InnerException }
            $result.Error = $inner.Message
            return [pscustomobject]$result
        }
        $result.Reachable = $true
        $tcp.ReceiveTimeout = $TimeoutSeconds * 1000
        $tcp.SendTimeout = $TimeoutSeconds * 1000
        # Same decision as Windows (no policy error), and a trace of whether a certificate arrived at all.
        $received = @{ Certificate = $false; Errors = $null }
        $validate = [Net.Security.RemoteCertificateValidationCallback]{
            param($sender, $certificate, $chain, $errors)
            if ($certificate) { $received.Certificate = $true }
            if ($errors -ne [Net.Security.SslPolicyErrors]::None) {
                $status = @($chain.ChainStatus | ForEach-Object { [string]$_.Status } | Where-Object { $_ -ne 'NoError' } | Select-Object -Unique)
                $received.Errors = "The remote certificate is invalid: $errors$(if ($status) { " ($($status -join ', '))" })"
            }
            return $errors -eq [Net.Security.SslPolicyErrors]::None
        }.GetNewClosure()
        $ssl = [Net.Security.SslStream]::new($tcp.GetStream(), $false, $validate)
        try {
            try {
                $ssl.AuthenticateAsClient($HostName)
            }
            catch {
                $inner = $_.Exception; while ($inner.InnerException) { $inner = $inner.InnerException }
                $result.Error = if ($received.Errors) { $received.Errors } else { $inner.Message }
                $result.Interrupted = -not $received.Certificate
                return [pscustomobject]$result
            }
            $certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($ssl.RemoteCertificate)
            $notAfter = $certificate.NotAfter.ToUniversalTime()
            $result.Valid = $true
            $result.Subject = $certificate.Subject
            $result.Issuer = $certificate.Issuer
            $result.NotAfterUtc = $notAfter.ToString('yyyy-MM-ddTHH:mm:ssZ')
            $result.DaysLeft = [int][Math]::Floor(($notAfter - [DateTime]::UtcNow).TotalDays)
            $result.Protocol = [string]$ssl.SslProtocol
            return [pscustomobject]$result
        }
        finally {
            $ssl.Dispose()
        }
    }
    finally {
        $tcp.Dispose()
    }
}

function Get-GctChallengeInfo {
    <# Schemes of the WWW-Authenticate challenges, and the parameters of the Bearer challenge. #>
    param([AllowEmptyCollection()][string[]]$Challenges)

    $schemes = [Collections.Generic.List[string]]::new()
    $parameters = @{}
    $bearer = $false
    foreach ($challenge in @($Challenges)) {
        if ([string]::IsNullOrWhiteSpace($challenge)) { continue }
        $scheme = ($challenge.Trim() -split '\s+', 2)[0].TrimEnd(',')
        if (-not $schemes.Contains($scheme)) { [void]$schemes.Add($scheme) }
        if ($scheme -ieq 'Bearer') {
            $bearer = $true
            foreach ($m in [regex]::Matches($challenge, '([A-Za-z_]+)\s*=\s*"([^"]*)"')) { $parameters[$m.Groups[1].Value.ToLowerInvariant()] = $m.Groups[2].Value }
        }
    }
    [pscustomobject]@{
        Schemes          = @($schemes)
        Bearer           = $bearer
        AuthorizationUri = $parameters['authorization_uri']
        IssuerKind       = $parameters['issuer_kind']
        # Entra ID (hybrid modern authentication): <token service ID>@<tenant ID>, comma separated.
        TrustedIssuers   = @(([string]$parameters['trusted_issuers']).Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        Error            = $parameters['error']
    }
}

function Get-GctAuthorityInfo {
    <#
        What an authorization URL given by Exchange points to: AD FS (https://<host>/adfs/oauth2/authorize,
        AdfsRoot) or Entra ID (login.microsoftonline.com, login.windows.net..., Tenant: the path segment:
        common, organizations, a tenant ID or a domain). Kind None when there is no URL, Other otherwise.
    #>
    param([AllowEmptyString()][string]$Uri)

    $info = [ordered]@{ Kind = 'None'; Uri = $Uri; Host = $null; AdfsRoot = $null; Tenant = $null; Name = 'no authorization server' }
    $parsed = $null
    if (-not $Uri -or -not [Uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsed)) { return [pscustomobject]$info }
    $info.Host = $parsed.Host
    $adfs = [regex]::Match($Uri, '^(https://[^/?#]+/adfs)/oauth2/authorize', 'IgnoreCase')
    if ($adfs.Success) {
        $info.Kind = 'ADFS'; $info.AdfsRoot = $adfs.Groups[1].Value; $info.Name = "AD FS ($($parsed.Host))"
    }
    elseif ($script:EntraHosts -contains $parsed.Host.ToLowerInvariant()) {
        $info.Kind = 'EntraID'
        $info.Tenant = ($parsed.AbsolutePath.Trim('/') -split '/')[0]
        $info.Name = "Entra ID ($($parsed.Host))"
    }
    else {
        $info.Kind = 'Other'; $info.Name = $parsed.Host
    }
    [pscustomobject]$info
}

function Get-GctOverallStatus {
    <# Failed > Blocked > Warning > Passed. Skipped steps never decide the status alone. #>
    param([AllowEmptyCollection()][object[]]$Steps)

    $statuses = @($Steps | ForEach-Object { $_.Status })
    if ($statuses -contains 'Failed') { return 'Failed' }
    if ($statuses -contains 'Blocked') { return 'Blocked' }
    if ($statuses -contains 'Warning') { return 'Warning' }
    if (@($statuses | Where-Object { $_ -eq 'Passed' }).Count) { return 'Passed' }
    return 'Failed'
}

function Get-GctEntraErrorHint {
    <# What to check for the Entra ID errors (AADSTS codes) a sign-in meets most often. #>
    param([AllowEmptyString()][string]$Description, [string]$Resource)

    $code = [regex]::Match($Description, 'AADSTS(\d+)').Groups[1].Value
    switch ($code) {
        '500011' {
            if ($Resource -match 'graph\.microsoft|outlook\.office') { return "the resource $Resource is not found in this tenant: check the tenant (Identity.TenantId, the domain of the source mailbox)." }
            return "the URL $Resource is not a service principal name of Office 365 Exchange Online ($($script:ExchangeAppId)) in this tenant: hybrid modern authentication is not configured for this on-premises URL (Hybrid Configuration Wizard, or servicePrincipalNames of Office 365 Exchange Online)."
        }
        '65001' { return 'the user or an administrator has not consented to this client: grant the consent in Entra ID (Enterprise applications), or use the default client (Identity.ClientId MicrosoftOffice), which needs no consent.' }
        '50011' { return 'the redirect URI is not one of the client: set Identity.RedirectUri to a redirect URI of the application (Mobile and desktop applications), or sign in with -SignIn DeviceCode.' }
        '50079' { return 'the user must register for multi-factor authentication first.' }
        '50126' { return 'wrong user name or password.' }
        '50053' { return 'the account is locked (too many sign-in attempts) or the sign-in was blocked as suspicious.' }
        '50057' { return 'the account is disabled in Entra ID.' }
        '530003' { return 'Conditional Access requires a compliant or managed device: run the test from such a device, or ask for an exclusion for the test account.' }
        '90094' { return 'an administrator must consent to this client (users cannot consent in this tenant): grant admin consent to the application in Entra ID (Enterprise applications).' }
        '53003' { return 'a Conditional Access policy blocked the sign-in: read the sign-in log of the user in Entra ID (Conditional Access tab).' }
        '50105' { return 'the user is not assigned to the application (assignment required).' }
        '700016' { return 'this client ID does not exist in the tenant (Identity.ClientId); a single-tenant application cannot sign in users of another tenant.' }
        '7000218' { return 'the client is not a public client: the device-code flow needs a public client (Allow public client flows).' }
        { $_ -in '50020', '50034', '90072' } { return 'the account used in the browser is not a user of this tenant.' }
        '50076' { return 'multi-factor authentication is required: complete it in the browser.' }
        '7000215' { return 'the client secret is wrong or expired (Certificates & secrets of the application).' }
        '700027' { return 'the certificate is not one of the application (Certificates & secrets), or the assertion is invalid.' }
        '1002012' { return 'the scope of an application must end with /.default.' }
        '90002' { return 'the tenant does not exist (Identity.TenantId or the domain of the source mailbox).' }
        default { return $null }
    }
}

function Get-GctUrlFields {
    <# Parameters of the query (and fragment) of a URL, decoded. #>
    param([Parameter(Mandatory = $true)][string]$Url)

    $fields = @{}
    $start = $Url.IndexOfAny([char[]]'?#')
    if ($start -lt 0) { return $fields }
    foreach ($pair in $Url.Substring($start + 1).Split([char[]]'&#', [StringSplitOptions]::RemoveEmptyEntries)) {
        $parts = $pair.Split('=', 2)
        $fields[[Uri]::UnescapeDataString($parts[0])] = if ($parts.Count -gt 1) { [Uri]::UnescapeDataString($parts[1].Replace('+', ' ')) } else { '' }
    }
    return $fields
}

function Get-GctTokenClaims {
    <# Payload of a JWT access token as a hashtable, or $null if the token is not a readable JWT. The signature is not checked. #>
    param([Parameter(Mandatory = $true)][string]$AccessToken)

    $parts = $AccessToken.Split('.')
    if ($parts.Count -lt 2) { return $null }
    try {
        $json = [Text.Encoding]::UTF8.GetString((ConvertFrom-GctBase64Url $parts[1]))
        $claims = $json | ConvertFrom-Json -AsHashtable
        if ($claims -isnot [hashtable]) { return $null }
        return $claims
    }
    catch {
        return $null
    }
}

function Resolve-GctEntraTenant {
    <#
        Tenant ID from the OpenID configuration Entra ID publishes for a tenant ID or a domain (traced,
        no sign-in). Returns TenantId (empty when the tenant is not found), Issuer, Url and Error.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Name)

    $url = "https://$($Context.Cloud.Login)/$([Uri]::EscapeDataString($Name))/v2.0/.well-known/openid-configuration"
    $response = Invoke-GctWebRequest -HttpClient $Context.HttpClient -Uri $url -UserAgent $script:GctUserAgent -Headers @{ Accept = 'application/json' }
    $json = $null
    try { $json = $response.Content | ConvertFrom-Json -ErrorAction Stop } catch { $json = $null }
    $issuer = [string](Get-GctField $json 'issuer')
    $tenantId = [regex]::Match($issuer, '[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}').Value
    $problem = if ($response.StatusCode -eq 200 -and $tenantId) { $null }
    elseif ($json -and (Get-GctField $json 'error_description')) { ([string](Get-GctField $json 'error_description') -split "`r?`n")[0] }
    else { "HTTP $($response.StatusCode)" }
    [pscustomobject]@{
        TenantId     = if ($problem) { $null } else { $tenantId }
        Issuer       = $issuer
        Region       = [string](Get-GctField $json 'tenant_region_scope')
        SubRegion    = [string](Get-GctField $json 'tenant_region_sub_scope')
        CloudInstance = [string](Get-GctField $json 'cloud_instance_name')
        Url          = $url
        Error        = $problem
    }
}

function Get-GctUserRealm {
    <#
        How Entra ID signs in a user, before any sign-in (traced): Managed (password hash
        synchronization, pass-through authentication), Federated (AD FS or another identity
        provider: AuthUrl) or Unknown (the domain is not a domain of a tenant).
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$User)

    $url = "https://$($Context.Cloud.Login)/common/userrealm/$([Uri]::EscapeDataString($User))?api-version=2.0"
    $response = Invoke-GctWebRequest -HttpClient $Context.HttpClient -Uri $url -UserAgent $script:GctUserAgent -Headers @{ Accept = 'application/json' }
    $json = $null
    try { $json = $response.Content | ConvertFrom-Json -ErrorAction Stop } catch { $json = $null }
    [pscustomobject]@{
        HttpStatus          = $response.StatusCode
        NameSpaceType       = [string](Get-GctField $json 'NameSpaceType')
        DomainName          = [string](Get-GctField $json 'DomainName')
        FederationBrandName = [string](Get-GctField $json 'FederationBrandName')
        AuthUrl             = [string](Get-GctField $json 'AuthURL')
        ConsumerDomain      = [bool](Get-GctField $json 'ConsumerDomain')
    }
}

