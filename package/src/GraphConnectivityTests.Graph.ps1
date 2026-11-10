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
    Version : 1.0.0
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
