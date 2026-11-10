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


# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBnOrIxu4tezdL9
# fsceMKETbX/FzFXCDIDqZxj1eSWp2KCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCACI3M0fhj8
# e86pra88uyOjyajv1gKqdo2as8aEu1bO7TANBgkqhkiG9w0BAQEFAASCAgCAP1oO
# 7L7p6PAs7N7LasRyVQBIqsfTadrfiZqHQT033neZyAhqDYwBBqZ2GxeROHMwzkTV
# fh9FzEBrUtYRrfYfJ9mDzVymCJ/gWHdFYTkTQWsCtWsSv27irbAtSZIbhC0szr4C
# 50Ym0CGvK9HAHCryf1lhFnQK6S1PgsZAqB5Tl75rvn7rPEaNeduqUoyv5oJrHCZD
# f4ye77l0zFSoS0kzX/ITGywBsUwmNGy509doW6MBlD99qLLS64xeuCn0M0EKABV3
# ZZCPpRPCIKDl6W+sEo9KGhgP1aU/gfJwrn0HMpp5P/cLjgqPGSbvMrddX7b+EIIJ
# MiWSk+TweN/Hb0O1tanrqVR7iYooZp3hv3FlGeKkSmDMPmUTwcCZLxdLKtS+4m4F
# lacSRPAQiA0LDwVGHhgS0AQ0uiwLU5lKkpIe6OyXx6FYq0VE2cKA3BWOVSRPRewx
# vkyXLHXUTSZuKtGRo70Qc61GL0+j6Dt20FZR+uKwzBR5DkAfTwVarGSIHP/rk3EY
# TCOuTVFGPdd61N80ipfm+07ex6B4FE13qlP23gQCovmRHIPZdMa7EbWHnTOnIjSa
# t8DMV2j4b+JmOYIgF6LN9zzBLiDxA9Mcmgsh6Ns2ljUEe7LQD+BzHgy7Ji44LjUp
# WSar8FG6vIIDdyBmAJJNZaRk+l4KfE2s7/94P6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIOJK+0EOL6gMWqNf/PFXTiWLyFBOqVvIczuRUfoU7nslAgh51LtZXc6x
# OxgPMjAyNjExMTAxMDI0MDNaMAMCAQECEAHNvNKpdYfa4Z/pxyZUcJOgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjQwM1owKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIExj
# HK2F6ATiibuGD3DBRVNpwxlF+cSiLCAAykNAvHNRMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEA7XHbWo6mBB/fKLa4eCTOVsLX12kQpwjLXPb4Siuk1sMCIBso
# 6w9HCU+KpF5QKCWbUmj0FjSR2tsHqceCxOHvW6vb
# SIG # End signature block
