<#
.SYNOPSIS
    Graph Connectivity Tests - Microsoft Graph requests and the free/busy answers (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Invoke-GctGraph sends one request with the token of the run: client-request-id and
    return-client-request-id, $select of the properties needed only, Prefer headers, throttling (HTTP
    429, 503, 504 with Retry-After, Test.MaxRetries attempts). The server that answered comes from
    x-ms-ags-diagnostic. It never throws for an HTTP error: the caller reads HttpStatus, ErrorCode
    and ErrorMessage.

    The free/busy answers of Microsoft Graph (getSchedule) and of EWS (GetUserAvailability) are read
    into the same object - one per target mailbox: availability view, calendar items, working hours,
    the detail the source is allowed to see, and for an error the cause and what to check
    (Get-GctAvailabilityHint, the errors of the Availability service of Exchange).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>

$script:GctMaxWaitSeconds = 120
# PowerShell 7.5+: dates of the JSON answers kept as text (7.4 turns ISO strings ending in Z into local DateTime values).
$script:GctJsonDateKind = (Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')

function Get-GctHeader {
    <# Value of a response header (case-insensitive), or $null. #>
    param([AllowNull()][pscustomobject]$Response, [Parameter(Mandatory = $true)][string]$Name)
    $headers = Get-GctField $Response 'Headers'
    if (-not $headers) { return $null }
    foreach ($k in $headers.Keys) { if ($k -ieq $Name) { return [string]$headers[$k] } }
    return $null
}

function ConvertTo-GctODataText {
    <# A text inside an OData string literal ('' for '), URL-encoded. #>
    param([AllowEmptyString()][string]$Text)
    return [Uri]::EscapeDataString($Text.Replace("'", "''"))
}

function Get-GctMePath {
    <# /me for a user, /users/{source} for an application. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    if ([string]$Context.Config.Context -eq 'Application') { return '/users/' + [Uri]::EscapeDataString([string]$Context.Config.SourceMailbox) }
    return '/me'
}

function Invoke-GctGraph {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Method = 'GET',
        [object]$Body,
        [hashtable]$Headers,
        [string]$Operation,
        [string]$Token,
        [string]$Version = 'v1.0'
    )

    Invoke-GctUiPump
    Assert-GctNotCancelled
    $uri = if ($Path -match '^https://') { $Path } else { "https://$($Context.Cloud.Graph)/$Version$Path" }
    if (-not $Operation) { $Operation = "$Method " + (([Uri]$uri).AbsolutePath -replace '^/(v1\.0|beta)', '' -replace '/users/[^/]+', '/users/{id}') }
    $json = if ($null -ne $Body) { ConvertTo-Json -InputObject $Body -Depth 10 -Compress } else { $null }
    $bearer = if ($Token) { $Token } elseif ($Context.Tokens['Graph']) { $Context.Tokens['Graph'].AccessToken } else { $null }
    $retries = [int]$Context.Config.MaxRetries
    $attempt = 0
    while ($true) {
        $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::new($Method), $uri)
        try {
            $add = { param([string]$Name, [string]$Value) $null = $request.Headers.TryAddWithoutValidation($Name, $Value) }
            & $add 'Accept' 'application/json'
            & $add 'User-Agent' ([string]$Context.Config.UserAgent)
            & $add 'client-request-id' ([guid]::NewGuid().ToString())
            & $add 'return-client-request-id' 'true'
            if ($Headers) { foreach ($k in $Headers.Keys) { & $add ([string]$k) ([string]$Headers[$k]) } }
            if ($bearer) { $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $bearer) }
            if ($null -ne $json) { $request.Content = [Net.Http.StringContent]::new($json, [Text.Encoding]::UTF8, 'application/json') }
            $response = Invoke-GctHttp -HttpClient $Context.HttpClient -Request $request -Operation $Operation
        }
        finally { $request.Dispose() }
        if ($response.StatusCode -notin 429, 503, 504 -or $attempt -ge $retries) { break }
        $after = 0
        $wait = if ([int]::TryParse([string](Get-GctHeader $response 'Retry-After'), [ref]$after) -and $after -gt 0) { $after } else { 5 }
        if ($wait -gt $script:GctMaxWaitSeconds) { Write-GctItem Warn "Microsoft Graph throttles the requests and asks to wait $wait s: the request is not sent again."; break }
        $attempt++
        $Context.Throttled++
        Write-GctItem Warn "Microsoft Graph throttles the requests (HTTP $($response.StatusCode)): new attempt $attempt/$retries in $wait s, as asked."
        Wait-GctSeconds $wait
    }
    $data = $null
    $text = if ($response.Body -and $response.Body.Length) { [Text.Encoding]::UTF8.GetString([byte[]]$response.Body) } else { '' }
    if ($text.TrimStart().StartsWith('{')) { try { $data = if ($script:GctJsonDateKind) { $text | ConvertFrom-Json -AsHashtable -Depth 32 -DateKind String -ErrorAction Stop } else { $text | ConvertFrom-Json -AsHashtable -Depth 32 -ErrorAction Stop } } catch { $data = $null } }
    $failure = if ($data -and $data['error']) { $data['error'] } else { $null }
    $server = Get-GctResponseServer -Response $response
    [pscustomobject]@{
        Operation    = $Operation
        Response     = $response
        HttpStatus   = [int]$response.StatusCode
        Json         = $data
        ErrorCode    = if ($failure) { [string]$failure['code'] } else { $null }
        ErrorMessage = if ($failure) { [string]$failure['message'] } else { $null }
        Server       = $server
        RequestId    = Get-GctHeader $response 'request-id'
        Retries      = $attempt
    }
}

function Test-GctGraphOk {
    param([Parameter(Mandatory = $true)][pscustomobject]$Answer)
    return $Answer.HttpStatus -ge 200 -and $Answer.HttpStatus -lt 300
}

function Get-GctGraphDetails {
    <# Details common to every Graph check: operation, HTTP status, error, server, request ID. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Answer, [System.Collections.IDictionary]$More)
    $d = [ordered]@{ Operation = $Answer.Operation; HttpStatus = $Answer.HttpStatus }
    if ($Answer.ErrorCode) { $d.ErrorCode = $Answer.ErrorCode; $d.ErrorMessage = $Answer.ErrorMessage }
    if ($Answer.Server) { $d.Servers = $Answer.Server }
    if ($Answer.RequestId) { $d.RequestId = $Answer.RequestId }
    if ($Answer.Retries) { $d.Retries = $Answer.Retries }
    if ($More) { foreach ($k in $More.Keys) { $d[$k] = $More[$k] } }
    return $d
}

function Get-GctGraphFailureText {
    <# Why a Graph request failed, with what to check. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][pscustomobject]$Answer, [string]$Permission)

    $app = [string]$Context.Config.Context -eq 'Application'
    $what = if ($Answer.ErrorCode) { "$($Answer.ErrorCode): $($Answer.ErrorMessage)" } else { 'no error body' }
    $hint = switch ($Answer.HttpStatus) {
        401 { 'Microsoft Graph refused the token (audience, expiry, tenant).' }
        403 {
            if ($app) { "the application permission$(if ($Permission) { " $Permission" }) with admin consent, and the scope of RBAC for Applications or of an application access policy." }
            elseif ($Permission) { "the token has no $Permission (scp of the token): the client $($Context.Client.Name) is not allowed it. Run with -ClientId GraphCommandLineTools (consent) or your application." }
            else { 'the delegated permission of the token, consent, or a Conditional Access policy.' }
        }
        404 {
            switch -Regex ([string]$Answer.ErrorCode) {
                'MailboxNotEnabledForRESTAPI|MailboxNotSupportedForRESTAPI' { 'the mailbox is not in Exchange Online (on-premises, or no Exchange license): Microsoft Graph cannot reach it.' }
                'ErrorInvalidUser|ResourceNotFound|Request_ResourceNotFound' { 'no such user or object in the tenant of the token.' }
                default { 'not found.' }
            }
        }
        429 { 'Microsoft Graph throttles the requests: wait and try again.' }
        default { $null }
    }
    $id = if ($Answer.RequestId) { " Request ID $($Answer.RequestId)." } else { '' }
    return "HTTP $($Answer.HttpStatus) ($what)$(if ($hint) { ": $hint" } else { '.' })$id"
}

#region Free/busy answers ---------------------------------------------------------------------------

function Get-GctAvailabilityHint {
    <#
        The cause of an error of the Availability service (getSchedule error.responseCode / GetUserAvailability
        ResponseCode, with its message) and what to check, for the path of the request.
    #>
    param([AllowEmptyString()][string]$Code, [AllowEmptyString()][string]$Message, [string]$Path)

    $m = [string]$Message
    $remote = $Path -in 'CloudToOnPremises', 'OnPremisesToCloud', 'CrossOrganisation'
    $inner = $null
    if ($m -match '\(401\)|401 Unauthorized|Unauthorized') { $inner = 'the remote Exchange refused the token of the requesting organisation (HTTP 401): OAuth trust - on-premises Get-AuthServer for the tenant (EvoSts or ACS), Get-PartnerApplication "Exchange Online" enabled, OAuth certificate (Get-AuthConfig) published to Entra ID, Test-OAuthConnectivity; with DAuth the federation trust (Get-FederationTrust, Test-FederationTrust).' }
    elseif ($m -match '\(403\)|Forbidden') { $inner = 'the remote endpoint answered HTTP 403: a reverse proxy, a firewall or an EWS policy (Get-OrganizationConfig EwsApplicationAccessPolicy, Get-CASMailbox EwsEnabled) blocks the request.' }
    elseif ($m -match '\(404\)|Not Found') { $inner = 'the remote URL does not exist (HTTP 404): TargetSharingEpr / DiscoveryEndpoint / TargetAutodiscoverEpr point to a wrong address.' }
    elseif ($m -match '\(503\)|\(502\)|\(504\)|Service Unavailable|Bad Gateway') { $inner = 'the remote Exchange or its load balancer is not available (HTTP 5xx).' }
    elseif ($m -match 'could not be processed in time') { $inner = 'Exchange Online sent the request but got no answer in time: the endpoint can be reachable yet slow to answer (on-premises Exchange validating the token or calling Entra ID and Exchange Online through an outbound proxy), or the IntraOrganizationConnector / organization relationship points to another URL (TargetSharingEpr, DiscoveryEndpoint, Hybrid Agent). Run the test again; on-premises, look in the EWS and HttpProxy logs for the request from Exchange Online.' }
    elseif ($m -match 'Unable to connect|actively refused|timed out|timeout|could not be resolved|remote name|No such host|connection attempt failed') { $inner = 'Exchange Online could not reach the remote endpoint from the Internet: public DNS name, firewall and reverse proxy must let the Microsoft 365 IP addresses in on TCP 443 (/EWS and /Autodiscover in pass-through).' }
    elseif ($m -match 'trust relationship|SSL|TLS|certificate|secure channel') { $inner = 'the TLS connection to the remote endpoint failed: certificate of a public authority, complete chain, name of the URL, TLS 1.2.' }
    elseif ($m -match 'delegation token|Federation Gateway|federated') { $inner = 'the federation token (DAuth) could not be obtained: federation trust and federated organization identifier of both organisations (Get-FederatedOrganizationIdentifier, Test-FederationTrust).' }
    elseif ($m -match 'autodiscover') { $inner = 'Autodiscover of the remote organisation failed: TargetAutodiscoverEpr (organization relationship) or DiscoveryEndpoint (IntraOrganizationConnector), and the publication of /Autodiscover.' }
    switch -Regex ([string]$Code) {
        'ErrorMailRecipientNotFound|ErrorInvalidSmtpAddress' {
            return "Exchange does not know this address: no mailbox, mail user or contact with it in the source organisation$(if ($remote) { ', and no organization relationship, IntraOrganizationConnector or availability address space for its domain' }). In hybrid, the on-premises mailbox must be a mail user in Exchange Online (directory synchronization)."
        }
        'ErrorAvailabilityConfigNotFound|ErrorNoOrganizationRelationship|ErrorOrganizationNotFederated' {
            return "Exchange Online has no free/busy configuration for this domain: create an organization relationship (or an IntraOrganizationConnector in hybrid, or an availability address space) with the domain in DomainNames / TargetAddressDomains."
        }
        'ErrorNoFreeBusyAccess' { return 'The target refuses to share its free/busy with the source: calendar permission Default / Anonymous / of the source set to None, or FreeBusyAccessEnabled / FreeBusyAccessLevel None in the organization relationship of the target organisation.' }
        'ErrorProxyRequestProcessingFailed|ErrorProxyRequestNotAllowed|ErrorProxyCallFailed|ErrorProxyServiceDiscoveryFailed|ErrorServiceDiscoveryFailed|ErrorAutoDiscoverFailed|ErrorConnectionFailed|ErrorTimeoutExpired' {
            $base = "The request was sent on to the other side ($Code) and failed there"
            if ($inner) { return "$($base): $inner" }
            return "$base. Read the full message: it gives the remote URL and the error; then check the IntraOrganizationConnector (Get-IntraOrganizationConnector) or the organization relationship (Get-OrganizationRelationship) and the publication of /EWS and /Autodiscover."
        }
        'ErrorFreeBusyGenerationFailed' { return 'Exchange found the mailbox but could not build its free/busy (calendar of the target not readable, mailbox being moved, or a remote error): read the message.' }
        'ErrorFreeBusyDLLimitReached|ErrorDistributionListMemberNotExist' { return 'The address is a group: free/busy is given for at most 100 members, and only for a group that Exchange can expand.' }
        'ErrorServerBusy|ErrorTooManyObjectsOpened' { return 'Exchange throttled the request: wait and try again.' }
        default { if ($inner) { return $inner } return $null }
    }
}

function Get-GctDetailLevel {
    <# The detail of the calendar the source may see, deduced from the items received. #>
    param([AllowEmptyCollection()][object[]]$Events, [AllowEmptyString()][string]$View)
    $items = @($Events | Where-Object { $_ })
    if (-not $items.Count) {
        if ($View -match '[1-4]') { return 'Free/busy time only (busy periods without items)' }
        return 'Free/busy time (no item in the period)'
    }
    $withSubject = @($items | Where-Object { $_.Subject -or $_.Location })
    if ($withSubject.Count) { return 'Limited details or more (subject and location)' }
    return 'Free/busy time only (no subject)'
}

function ConvertFrom-GctGraphDate {
    <# dateTimeTimeZone of Graph (UTC asked with Prefer outlook.timezone) -> UTC text. #>
    param([AllowNull()][object]$Value)
    if (-not $Value) { return $null }
    $t = [string]$Value['dateTime']
    if (-not $t) { return $null }
    $d = [datetime]::Parse($t, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal)
    return $d.ToString('yyyy-MM-ddTHH:mm:ssZ')
}

function ConvertFrom-GctSchedule {
    <# One scheduleInformation of getSchedule -> the free/busy object of the tool. #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Info, [string]$Path, [int]$IntervalMinutes, [datetime]$StartUtc)

    $mailbox = [string]$Info['scheduleId']
    $err = $Info['error']
    $events = @(foreach ($i in @($Info['scheduleItems'])) {
            if (-not $i) { continue }
            $s = ConvertFrom-GctGraphDate $i['start']; $e = ConvertFrom-GctGraphDate $i['end']
            [pscustomobject]@{
                Mailbox = $mailbox; StartUtc = $s; EndUtc = $e
                Start = if ($s) { ([datetime]$s).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { $null }
                End = if ($e) { ([datetime]$e).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { $null }
                BusyType = switch ([string]$i['status']) { 'oof' { 'OOF' } 'workingElsewhere' { 'WorkingElsewhere' } default { (Get-Culture).TextInfo.ToTitleCase([string]$i['status']) } }
                Subject = [string]$i['subject']; Location = [string]$i['location']; IsPrivate = [bool]$i['isPrivate']; IsMeeting = $null
            }
        })
    $wh = $Info['workingHours']
    $days = @(if ($wh) { @($wh['daysOfWeek']) | ForEach-Object { ([string]$_).ToLowerInvariant() } })
    $toMinutes = { param([string]$T) if (-not $T) { return $null }; $p = $T.Split(':'); [int]$p[0] * 60 + [int]$p[1] }
    $zone = if ($wh -and $wh['timeZone']) { [string]$wh['timeZone']['name'] } else { $null }
    $code = if ($err) { [string]$err['responseCode'] } else { $null }
    $message = if ($err) { [string]$err['message'] } else { $null }
    $view = [string]$Info['availabilityView']
    [pscustomobject]@{
        Mailbox         = $mailbox
        Protocol        = 'Microsoft Graph getSchedule'
        Path            = $Path
        Status          = if ($err) { 'Failed' } elseif (-not $view) { 'Warning' } else { 'Passed' }
        ResponseCode    = $code
        Message         = $message
        Hint            = if ($err) { Get-GctAvailabilityHint -Code $code -Message $message -Path $Path } else { $null }
        View            = $view
        IntervalMinutes = $IntervalMinutes
        StartUtc        = $StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
        DetailLevel     = if ($err) { $null } else { Get-GctDetailLevel -Events $events -View $view }
        WorkingHours    = if ($wh) { '{0} {1}-{2} ({3})' -f ($days -join ' '), ([string]$wh['startTime']).Substring(0, 5), ([string]$wh['endTime']).Substring(0, 5), $zone } else { $null }
        WorkDays        = $days
        WorkStart       = if ($wh) { & $toMinutes ([string]$wh['startTime']) } else { $null }
        WorkEnd         = if ($wh) { & $toMinutes ([string]$wh['endTime']) } else { $null }
        WorkTimeZone    = $zone
        WorkOffsetMinutes = Get-GctZoneOffset -Name $zone -AtUtc $StartUtc
        Events          = $events
    }
}

#endregion

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDl2CRakhr2A0Hf
# w1x4GWerNYNG/bVjYQJ28bSG7mwBDqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDSLA3o8Zth
# 9WFNk7GlXN52MpAZ2ULxlswoEWJqWh+NqDANBgkqhkiG9w0BAQEFAASCAgCVTf7W
# yu6+tgngt4uH2HsP9Q+e9JdQCnOpenINVjuc4ksi2NpH5+5YZ+NK8NfK+VPPECWh
# ZVkUqjkY12w76GHPMx0UFMqEHMwx88vdBWraIFykeK9siaZyubQ9qR1kU4+hiKIv
# 1UWAXAL9FofuqorhobqFJjOfy3Jihor4S7jk6FPDeS/Kde68yDsTXTf+j+A9B438
# pkcSSVg3wzvYSbEv+0MFu9VLJKvABmSQzhXUWfojTVrd4N1m4CF0BSVbDbOU6maF
# bi++iWiDWULNCWq2gnwkrgLHYCf5rgUT1hT27OHfUYrtW4wwvxoUKjV6Dz/ghJ6s
# RKQvXu9ZuCThnmMF8wprxn7F7cKFFMjUa3DqIUaiKA6rMxNG6fCw/XCkrFPaZcuC
# bLPI926kpoXSTdWfUs3MRIgoKkH3VfIfb8nWfLXEDOX2ve8nL2n6S5LtXkBiTV32
# hSkD4GZ+t2LIVFLokoauL7vyVzlnP6DIBJftxHwPkcJprxEz78pVT8oijhOAWsRV
# hnDoJ2b+ycvH9/SwTQ6iPzgKUcwUqf+h9/vjbyX/0Fl5GMQTXl4nALqREdpw1adu
# TP+1v14btlIisT79tLmbxvtvDhLVNpfh43AeNKK0KJcvrGw7bNOZuzghSqoo/GKz
# eqa+l0ZBBN0DjXCghOrPCHV9YvBMDewOYeDYmaGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIG5GfB4MKUzE7Rr/S9z8lqwivYSSr1bcyJBOfyG4tIQBAghSl7Ib4b3j
# pxgPMjAyNjExMTAxODI2MTdaMAMCAQECEA+BcZTY3EDn2YI5jWo1052gghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDE4MjYxN1owKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIC6r
# keXixPM1lt+uBtyWEMa61UfOQbtEgPtCVQEcp/8BMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiAA9F9+ibVcOCWVaBnCllgdeSZCBBwaoheppIDJKfoKaQIhAKIq
# uD4d4ZkhORDTtuoiAseu/Fbf62acuspzQOOn3Dim
# SIG # End signature block
