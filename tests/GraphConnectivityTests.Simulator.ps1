<#
.SYNOPSIS
    Graph Connectivity Tests - simulated Entra ID, Microsoft Graph, Exchange Online and Exchange on-premises, for the tests.

.DESCRIPTION
    Dot-source this file. Get-SimHttpResponse answers every HTTP request of the tool like the real servers,
    Get-SimDnsAnswer every DNS query, Get-SimCertificate every TLS check. The simulated world:

      contoso.com    tenant 7d4e2a91-..., hybrid: alice, bob, room1 (room), erin (no license, no mailbox) in
                     Exchange Online, carol on-premises (mail.contoso.com, Exchange Server SE) - a mail user in
                     Exchange Online; bob.smith@contoso.com is a secondary address of bob; nobody@ does not exist.
      fabrikam.com   another organisation (tenant b6f1c0d2-...), federated sign-in, dave in Exchange Online,
                     an organization relationship with contoso (free/busy only, MailTips).
      gmail.com      a consumer domain: the tenant of the Microsoft accounts (9cd80435-...), ConsumerDomain.

    Entra ID: OpenID configuration per domain, user realm, device code (pending once, then the token),
    authorization code with PKCE (checked), refresh token (the on-premises URL must be a service principal
    name), client credentials with a client assertion. Tokens are unsigned JWTs with real claims.
    Exchange Online: Autodiscover v2 (a 302 to the on-premises Autodiscover for carol), the administration API
    (InvokeCommand). Exchange on-premises: anonymous 401 with Bearer (trusted_issuers) and Negotiate, EWS
    GetUserAvailability with a Bearer token or a Windows session. Microsoft Graph: users, getSchedule,
    getMailTips, delta, mailboxSettings, licenseDetails, the Inbox probe, userPurpose, domains.

      New-SimState          what the simulated organisations do (every switch of the tests)
      New-SimToken          unsigned JWT with the claims of Entra ID
      Get-SimHttpResponse   answer to one HTTP request of the tool
      Install-SimNetwork    replaces the network inside the loaded module (Pester)

    Every secret the simulator hands out (tokens, codes, assertions) is recorded in State.Secrets, so that the
    tests can prove that none of them reaches the trace or a report file.

    Nothing here is used by the tool itself, and it is not part of the package.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:SimContoso = '7d4e2a91-3c5b-4f6e-8a1d-2b9c0e5f4a37'
$script:SimFabrikam = 'b6f1c0d2-4e8a-4c3b-9a71-5e2d8f0c6a14'
$script:SimConsumer = '9cd80435-793b-4f48-844b-6b3f37d1c1f3'
$script:SimTenants = @{ 'contoso.com' = $script:SimContoso; 'contoso.onmicrosoft.com' = $script:SimContoso; 'fabrikam.com' = $script:SimFabrikam; 'gmail.com' = $script:SimConsumer; 'outlook.com' = $script:SimConsumer }
$script:SimOfficeClient = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
$script:SimAdminClient = 'fb78d390-0c51-40cd-8e17-fdbfab77341b'
$script:SimNs = 'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"'
$script:SimReasons = @{ 200 = 'OK'; 201 = 'Created'; 204 = 'No Content'; 302 = 'Found'; 400 = 'Bad Request'; 401 = 'Unauthorized'; 403 = 'Forbidden'; 404 = 'Not Found'; 429 = 'Too Many Requests'; 500 = 'Internal Server Error'; 503 = 'Service Unavailable' }

function ConvertTo-SimBase64Url([byte[]]$Bytes) { [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_') }

function New-SimToken {
    <# Unsigned JWT with the claims of an Entra ID access token (v1.0): aud, iss, tid, upn, scp or roles, exp. #>
    param([string]$Audience = 'https://graph.microsoft.com', [string]$Scope, [string]$Upn, [string]$AppId = $script:SimOfficeClient, [string]$TenantId = $script:SimContoso, [string[]]$Roles, [int]$ExpiresIn = 3600, [string]$ObjectId, [string]$Name)
    $encode = { param($o) ConvertTo-SimBase64Url ([Text.Encoding]::UTF8.GetBytes(($o | ConvertTo-Json -Compress -Depth 4))) }
    $now = [DateTimeOffset]::UtcNow
    $payload = [ordered]@{ aud = $Audience; iss = "https://sts.windows.net/$TenantId/"; iat = $now.AddSeconds(-30).ToUnixTimeSeconds(); nbf = $now.AddSeconds(-30).ToUnixTimeSeconds(); exp = $now.AddSeconds($ExpiresIn).ToUnixTimeSeconds(); appid = $AppId; appidacr = '0'; tid = $TenantId; ver = '1.0' }
    if ($Upn) { $payload.upn = $Upn; $payload.unique_name = $Upn; $payload.name = $(if ($Name) { $Name } else { $Upn }); $payload.oid = $(if ($ObjectId) { $ObjectId } else { [guid]::NewGuid().ToString() }) }
    if ($Scope) { $payload.scp = $Scope }
    if ($Roles) { $payload.roles = @($Roles); $payload.app_displayname = 'Contoso free/busy checker'; $payload.appidacr = '2'; $payload.oid = [guid]::NewGuid().ToString() }
    '{0}.{1}.{2}' -f (& $encode ([ordered]@{ typ = 'JWT'; alg = 'RS256'; kid = 'SIM-KEY-1' })), (& $encode $payload), (ConvertTo-SimBase64Url ([Security.Cryptography.RandomNumberGenerator]::GetBytes(64)))
}

function Get-SimClaims([string]$Token) {
    if (-not $Token) { return $null }
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { return $null }
    $p = $parts[1].Replace('-', '+').Replace('_', '/'); switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
    try { [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json -AsHashtable } catch { $null }
}

function New-SimUser {
    param([string]$Address, [string]$Name, [string]$Tenant = $script:SimContoso, [ValidateSet('EXO', 'OnPrem', 'None')][string]$Location = 'EXO', [string]$Purpose = 'user', [bool]$Licensed = $true, [bool]$Synced = $true, [bool]$Enabled = $true, [string[]]$Aliases = @())
    @{
        Id = [guid]::NewGuid().ToString(); Address = $Address; Upn = $Address; DisplayName = $Name; Tenant = $Tenant; Location = $Location; Purpose = $Purpose
        Licensed = $Licensed; Synced = $Synced; Enabled = $Enabled; Aliases = @($Aliases); ImmutableId = if ($Synced) { [Convert]::ToBase64String([guid]::NewGuid().ToByteArray()) } else { $null }
    }
}

function New-SimState {
    <# Default: a healthy hybrid organisation contoso.com, a partner fabrikam.com, a consumer domain. #>
    $users = [ordered]@{}
    foreach ($u in @(
            (New-SimUser 'alice@contoso.com' 'Alice Martin')
            (New-SimUser 'bob@contoso.com' 'Bob Smith' -Aliases 'bob.smith@contoso.com')
            (New-SimUser 'room1@contoso.com' 'Meeting Room 1' -Purpose 'room' -Licensed $false -Synced $false -Enabled $false)
            (New-SimUser 'carol@contoso.com' 'Carol Dupont' -Location 'OnPrem' -Licensed $false)
            (New-SimUser 'erin@contoso.com' 'Erin Leroy' -Location 'None' -Licensed $false)
            (New-SimUser 'admin@contoso.com' 'Exchange Administrator' -Synced $false)
            (New-SimUser 'dave@fabrikam.com' 'Dave Jones' -Tenant $script:SimFabrikam)
        )) { $users[$u.Address] = $u }
    @{
        Calls              = [Collections.Generic.List[object]]::new()
        Secrets            = [Collections.Generic.List[string]]::new()
        ValidTokens        = [Collections.Generic.List[string]]::new()
        RefreshTokens      = @{}
        DeviceCodes        = @{}
        AuthCodes          = @{}
        Waits              = [Collections.Generic.List[double]]::new()
        Patches            = [Collections.Generic.List[object]]::new()
        Users              = $users
        SignInUser         = 'alice@contoso.com'
        AdminUser          = 'admin@contoso.com'
        GraphScopes        = 'Calendars.ReadWrite Calendars.ReadWrite.Shared Contacts.ReadWrite Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite People.Read User.Read User.Read.All Domain.Read.All'
        AppRoles           = @('Calendars.Read', 'User.Read.All', 'MailboxSettings.Read', 'Mail.ReadBasic.All', 'Domain.Read.All')
        PendingPolls       = 1
        DeviceDenied       = $false
        DesktopSession     = $false
        OnPremSpns         = @('mail.contoso.com')
        OnPremBearer       = $true
        OnPremTrustedTenant = $script:SimContoso
        OnPremProxy        = $false
        WindowsAccepted    = $true
        OnPremBasic        = $false
        OnPremPassword     = 'Sim-Pa55word'
        OnPremAccounts     = @{ 'contoso\carol' = 'carol@contoso.com'; 'carol@contoso.com' = 'carol@contoso.com' }
        CloudToOnPrem      = 'Success'
        OnPremToCloud      = 'Success'
        CrossOrg           = 'Success'
        AutoReply          = @{ status = 'disabled'; externalAudience = 'all'; internalReplyMessage = '<html><body>Original internal text</body></html>'; externalReplyMessage = '<html><body>Original external text</body></html>'
                                scheduledStartDateTime = @{ dateTime = '2026-11-01T08:00:00.0000000'; timeZone = 'UTC' }; scheduledEndDateTime = @{ dateTime = '2026-11-02T08:00:00.0000000'; timeZone = 'UTC' } }
        AutoReplyForbidden = $false
        BobOutOfOffice     = $true
        ThrottleOnce       = $null
        CertificateDays    = 365
        TlsFailure         = $null
        DnsDown            = $false
        Dns                = @{
            'contoso.com|MX'                          = @('0 contoso-com.mail.protection.outlook.com')
            'contoso.com|TXT'                         = @('MS=ms83421977', 'v=spf1 include:spf.protection.outlook.com -all')
            'autodiscover.contoso.com|CNAME'          = @('mail.contoso.com')
            'mail.contoso.com|A'                      = @('203.0.113.25')
            'mta-sts.contoso.com|A'                   = @('203.0.113.30')
            'selector1._domainkey.contoso.com|CNAME'  = @('selector1-contoso-com._domainkey.contoso.onmicrosoft.com')
            'selector2._domainkey.contoso.com|CNAME'  = @('selector2-contoso-com._domainkey.contoso.onmicrosoft.com')
            '_dmarc.contoso.com|TXT'                  = @('v=DMARC1; p=quarantine; rua=mailto:dmarc@contoso.com')
            '_mta-sts.contoso.com|TXT'                = @('v=STSv1; id=20261101T000000')
            '_smtp._tls.contoso.com|TXT'              = @('v=TLSRPTv1; rua=mailto:tlsrpt@contoso.com')
            'enterpriseregistration.contoso.com|CNAME' = @('enterpriseregistration.windows.net')
            'fabrikam.com|MX'                         = @('20 mx2.fabrikam.com', '10 mx1.fabrikam.com')
            'fabrikam.com|TXT'                        = @('v=spf1 ip4:198.51.100.0/24 include:_spf.fabrikam.com ~all')
            '_spf.fabrikam.com|TXT'                   = @('v=spf1 include:spf.protection.outlook.com -all')
            'autodiscover.fabrikam.com|CNAME'         = @('autodiscover.outlook.com')
        }
        AdminForbidden     = $false
        AdminCommands      = @{
            'Get-IntraOrganizationConnector' = @(@{ Name = 'HybridIOC - 7d4e2a91'; Enabled = $true; TargetAddressDomains = @('contoso.com'); DiscoveryEndpoint = 'https://mail.contoso.com/autodiscover/autodiscover.svc'; TargetSharingEpr = $null })
            'Get-OrganizationRelationship'   = @(
                @{ Name = 'O365 to On-premises - 7d4e2a91'; Enabled = $true; DomainNames = @('contoso.com'); FreeBusyAccessEnabled = $true; FreeBusyAccessLevel = 'LimitedDetails'; TargetAutodiscoverEpr = 'https://autodiscover.contoso.com/autodiscover/autodiscover.svc/WSSecurity'; TargetSharingEpr = $null; TargetApplicationUri = 'FYDIBOHF25SPDLT.contoso.com'; MailTipsAccessEnabled = $true }
                @{ Name = 'Fabrikam'; Enabled = $true; DomainNames = @('fabrikam.com'); FreeBusyAccessEnabled = $true; FreeBusyAccessLevel = 'AvailabilityOnly'; TargetAutodiscoverEpr = 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity'; TargetSharingEpr = $null; TargetApplicationUri = 'outlook.com'; MailTipsAccessEnabled = $true }
            )
            'Get-AvailabilityAddressSpace'   = @()
        }
    }
}

function Find-SimUser([hashtable]$State, [string]$Key) {
    if (-not $Key) { return $null }
    $k = $Key.ToLowerInvariant()
    foreach ($u in $State.Users.Values) { if ($u.Address -eq $k -or $u.Id -eq $k -or $u.Upn -eq $k -or $u.Aliases -contains $k) { return $u } }
    return $null
}

function Get-SimTenantOf([string]$Name) {
    $n = ([string]$Name).ToLowerInvariant()
    if ($n -match '@') { $n = $n.Split('@')[1] }
    if ($script:SimTenants.ContainsKey($n)) { return $script:SimTenants[$n] }
    if ($n -in $script:SimTenants.Values) { return $n }
    return $null
}

function New-SimResponse {
    <# Same shape as Send-GctHttpRequest: StatusCode, Reason, Headers (hashtable), HeaderLines, Challenges, Body (bytes). #>
    param([int]$Code, [string]$Body = '', [hashtable]$Headers = @{}, [string[]]$Challenges = @(), [string]$ContentType = 'application/json; charset=utf-8')
    $h = @{} + $Headers
    if ($Body -and -not $h.ContainsKey('Content-Type')) { $h['Content-Type'] = $ContentType }
    $h['Date'] = [DateTime]::UtcNow.ToString('r')
    $lines = [Collections.Generic.List[string]]::new()
    foreach ($k in ($h.Keys | Sort-Object)) { $lines.Add("$($k): $($h[$k])") }
    if ($Challenges) { $h['WWW-Authenticate'] = $Challenges -join ','; foreach ($c in $Challenges) { $lines.Add("WWW-Authenticate: $c") } }
    [pscustomobject]@{ StatusCode = $Code; Reason = $script:SimReasons[$Code]; Headers = $h; HeaderLines = @($lines); Challenges = @($Challenges); Body = [Text.Encoding]::UTF8.GetBytes($Body) }
}

function New-SimJson([int]$Code, $Object, [hashtable]$Headers = @{}) {
    $text = if ($null -eq $Object) { '' } else { ConvertTo-Json -InputObject $Object -Depth 12 -Compress }
    New-SimResponse $Code $text $Headers
}

function Get-SimFormFields([string]$Body) {
    $fields = @{}
    foreach ($pair in $Body.Split('&')) { if (-not $pair) { continue }; $k, $v = $pair.Split('=', 2); $fields[[Uri]::UnescapeDataString($k.Replace('+', ' '))] = [Uri]::UnescapeDataString(([string]$v).Replace('+', ' ')) }
    $fields
}

#region Free/busy -----------------------------------------------------------------------------------

function Get-SimSchedule {
    <#
        A plausible week for a mailbox, from StartUtc: meetings, a tentative one, a day out of office, an afternoon
        elsewhere - varied by mailbox. Events in UTC and the merged view (0 free, 1 tentative, 2 busy, 3 away,
        4 elsewhere), one digit per interval, like Exchange returns it. Bob has a meeting whose subject is a
        spreadsheet formula (CSV injection check).
    #>
    param([string]$Mailbox, [datetime]$StartUtc, [int]$Days = 7, [int]$IntervalMinutes = 30)

    $seed = [int]$Mailbox.ToLowerInvariant()[0] % 3
    $plan = switch ($seed) {
        0 { @(@(0, 9, 0, 60, 'Busy', 'Weekly review', 'Room 1'), @(0, 14, 0, 90, 'Tentative', 'Design sync', 'Teams'), @(1, 8, 0, 600, 'OOF', 'Out of office', ''), @(2, 11, 0, 60, 'Busy', 'Customer call', 'Teams'), @(3, 13, 0, 240, 'WorkingElsewhere', 'Working from home', ''), @(4, 10, 0, 120, 'Busy', 'Change board', 'Room 2')) }
        1 { @(@(0, 10, 0, 60, 'Busy', 'Project kick-off', 'Room 3'), @(0, 15, 0, 60, 'Busy', 'Exchange migration', 'Teams'), @(1, 9, 30, 90, 'Busy', 'Training', 'Room 1'), @(2, 14, 0, 120, 'Tentative', 'Workshop', 'Room 2'), @(4, 13, 0, 300, 'OOF', 'Day off afternoon', '')) }
        default { @(@(0, 8, 30, 30, 'Busy', 'Stand-up', 'Teams'), @(0, 12, 0, 60, 'Busy', 'Lunch with team', ''), @(1, 8, 30, 30, 'Busy', 'Stand-up', 'Teams'), @(2, 10, 0, 180, 'WorkingElsewhere', 'Customer site', 'Lyon'), @(3, 15, 0, 60, 'Tentative', 'Optional demo', 'Teams')) }
    }
    if ($Mailbox -ieq 'bob@contoso.com') { $plan = @($plan) + , @(0, 16, 0, 30, 'Busy', '=HYPERLINK("https://evil.example","Agenda")', '@Room 9') }
    $localStart = $StartUtc.ToLocalTime().Date
    $events = foreach ($p in $plan) {
        if ($p[0] -ge $Days) { continue }
        $s = $localStart.AddDays($p[0]).AddHours($p[1]).AddMinutes($p[2]).ToUniversalTime()
        [pscustomobject]@{ Start = $s; End = $s.AddMinutes($p[3]); BusyType = $p[4]; Subject = $p[5]; Location = $p[6] }
    }
    $code = @{ Free = 0; Tentative = 1; Busy = 2; OOF = 3; WorkingElsewhere = 4 }
    $rank = @{ 3 = 5; 2 = 4; 1 = 3; 4 = 2; 0 = 1 }
    $slots = [int]($Days * 1440 / $IntervalMinutes)
    $view = [Text.StringBuilder]::new()
    for ($i = 0; $i -lt $slots; $i++) {
        $t0 = $StartUtc.AddMinutes($i * $IntervalMinutes); $t1 = $t0.AddMinutes($IntervalMinutes); $best = 0
        foreach ($e in $events) { if ($e.Start -lt $t1 -and $e.End -gt $t0) { $v = $code[$e.BusyType]; if ($rank[$v] -gt $rank[$best]) { $best = $v } } }
        [void]$view.Append($best)
    }
    [pscustomobject]@{ Events = @($events); View = $view.ToString() }
}

function Get-SimFreeBusyOutcome {
    <#
        What the Availability service answers for one address. From 'Cloud' (Exchange Online asks: getSchedule) or
        'OnPrem' (Exchange on-premises asks: GetUserAvailability). Returns Code, Message and Detail (Full: subject
        and location; FreeBusy: busy periods only).
    #>
    param([hashtable]$State, [hashtable]$Source, [string]$Address, [ValidateSet('Cloud', 'OnPrem')][string]$From)

    $u = Find-SimUser $State $Address
    $domainTenant = Get-SimTenantOf $Address
    $ok = { param([string]$Detail = 'Full') [pscustomobject]@{ Code = 'Success'; Message = $null; Detail = $Detail } }
    $fail = { param([string]$Code, [string]$Message) [pscustomobject]@{ Code = $Code; Message = $Message; Detail = $null } }
    $proxy = {
        param([string]$Mode, [string]$Url)
        switch ($Mode) {
            'Unauthorized' { & $fail 'ErrorProxyRequestProcessingFailed' "Proxy web request failed., inner exception: The remote server returned an error: (401) Unauthorized., Url: $Url" }
            'Unreachable' { & $fail 'ErrorProxyRequestProcessingFailed' "Proxy web request failed., inner exception: Unable to connect to the remote server, Url: $Url" }
            default { & $ok 'Full' }
        }
    }
    if ($u -and $u.Tenant -eq $Source.Tenant) {
        if ($From -eq 'Cloud') {
            if ($u.Location -eq 'EXO') { return & $ok 'Full' }
            if ($u.Location -eq 'OnPrem') { return & $proxy $State.CloudToOnPrem 'https://mail.contoso.com/ews/exchange.asmx' }
            return & $fail 'ErrorMailRecipientNotFound' 'Mailbox does not exist.'
        }
        if ($u.Location -eq 'OnPrem') { return & $ok 'Full' }
        if ($u.Location -eq 'EXO') { return & $proxy $State.OnPremToCloud 'https://outlook.office365.com/EWS/Exchange.asmx' }
        return & $fail 'ErrorMailRecipientNotFound' 'Mailbox does not exist.'
    }
    if ($u) {
        if ($State.CrossOrg -eq 'Success') { return & $ok 'FreeBusy' }
        return & $fail $State.CrossOrg "Unable to get the free/busy of $Address from the remote organization."
    }
    if ($domainTenant -eq $Source.Tenant) { return & $fail 'ErrorMailRecipientNotFound' 'Mailbox does not exist.' }
    return & $fail 'ErrorAvailabilityConfigNotFound' "Unable to find the Availability configuration for the domain $(($Address -split '@')[-1])."
}

#endregion

#region Exchange on-premises (EWS) ------------------------------------------------------------------

function Get-SimEwsAvailability([hashtable]$State, [hashtable]$Source, [string]$Soap) {
    $addresses = @([regex]::Matches($Soap, '<t:Address>([^<]*)</t:Address>') | ForEach-Object { [Net.WebUtility]::HtmlDecode($_.Groups[1].Value) })
    $start = [datetime]::SpecifyKind([datetime][regex]::Match($Soap, '<t:StartTime>([^<]*)<').Groups[1].Value, 'Utc')
    $end = [datetime]::SpecifyKind([datetime][regex]::Match($Soap, '<t:EndTime>([^<]*)<').Groups[1].Value, 'Utc')
    $interval = [int][regex]::Match($Soap, '<t:MergedFreeBusyIntervalInMinutes>(\d+)<').Groups[1].Value
    $f = 'yyyy-MM-ddTHH:mm:ss'
    $responses = foreach ($a in $addresses) {
        $o = Get-SimFreeBusyOutcome -State $State -Source $Source -Address $a -From OnPrem
        if ($o.Code -ne 'Success') { "<m:FreeBusyResponse><m:ResponseMessage ResponseClass=""Error""><m:MessageText>$([Security.SecurityElement]::Escape($o.Message))</m:MessageText><m:ResponseCode>$($o.Code)</m:ResponseCode><m:DescriptiveLinkKey>0</m:DescriptiveLinkKey></m:ResponseMessage></m:FreeBusyResponse>"; continue }
        $plan = Get-SimSchedule -Mailbox $a -StartUtc $start -Days ([int]($end - $start).TotalDays) -IntervalMinutes $interval
        $items = ($plan.Events | ForEach-Object { "<t:CalendarEvent><t:StartTime>$($_.Start.ToString($f))</t:StartTime><t:EndTime>$($_.End.ToString($f))</t:EndTime><t:BusyType>$($_.BusyType)</t:BusyType><t:CalendarEventDetails><t:Subject>$([Security.SecurityElement]::Escape($_.Subject))</t:Subject><t:Location>$([Security.SecurityElement]::Escape($_.Location))</t:Location><t:IsMeeting>true</t:IsMeeting><t:IsRecurring>false</t:IsRecurring><t:IsPrivate>false</t:IsPrivate></t:CalendarEventDetails></t:CalendarEvent>" }) -join ''
        "<m:FreeBusyResponse><m:ResponseMessage ResponseClass=""Success""><m:ResponseCode>NoError</m:ResponseCode></m:ResponseMessage><m:FreeBusyView><t:FreeBusyViewType xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">DetailedMerged</t:FreeBusyViewType><t:MergedFreeBusy>$($plan.View)</t:MergedFreeBusy><t:CalendarEventArray>$items</t:CalendarEventArray><t:WorkingHours><t:TimeZone><t:Bias>-60</t:Bias><t:StandardTime><t:Bias>0</t:Bias><t:Time>03:00:00</t:Time><t:DayOrder>5</t:DayOrder><t:Month>10</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:StandardTime><t:DaylightTime><t:Bias>-60</t:Bias><t:Time>02:00:00</t:Time><t:DayOrder>5</t:DayOrder><t:Month>3</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:DaylightTime></t:TimeZone><t:WorkingPeriodArray><t:WorkingPeriod><t:DayOfWeek>Monday Tuesday Wednesday Thursday Friday</t:DayOfWeek><t:StartTimeInMinutes>480</t:StartTimeInMinutes><t:EndTimeInMinutes>1020</t:EndTimeInMinutes></t:WorkingPeriod></t:WorkingPeriodArray></t:WorkingHours></m:FreeBusyView></m:FreeBusyResponse>"
    }
    "<?xml version=""1.0"" encoding=""utf-8""?><s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/""><s:Header><h:ServerVersionInfo MajorVersion=""15"" MinorVersion=""2"" MajorBuildNumber=""2562"" MinorBuildNumber=""17"" Version=""V2017_07_11"" xmlns:h=""http://schemas.microsoft.com/exchange/services/2006/types"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/types""/></s:Header><s:Body><GetUserAvailabilityResponse xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"" $($script:SimNs)><m:FreeBusyResponseArray>$($responses -join '')</m:FreeBusyResponseArray></GetUserAvailabilityResponse></s:Body></s:Envelope>"
}

function Get-SimOnPremResponse([hashtable]$State, [Net.Http.HttpRequestMessage]$Request, [string]$Body) {
    $path = $Request.RequestUri.AbsolutePath.ToLowerInvariant()
    $auth = $Request.Headers.Authorization
    $headers = @{ 'X-FEServer' = 'EXCH01'; 'request-id' = [guid]::NewGuid().ToString(); 'X-Powered-By' = 'ASP.NET'; 'Server' = 'Microsoft-IIS/10.0' }
    $bearer = "Bearer client_id=""00000002-0000-0ff1-ce00-000000000000"", trusted_issuers=""00000001-0000-0000-c000-000000000000@$($State.OnPremTrustedTenant)"", token_types=""app_asserted_user_v1 service_asserted_app_v1"", authorization_uri=""https://login.windows.net/common/oauth2/authorize"""
    $schemes = @(if ($State.OnPremBearer) { $bearer }) + @('Negotiate', 'NTLM') + @(if ($State.OnPremBasic) { 'Basic realm="mail.contoso.com"' })
    if ($path -notin '/ews/exchange.asmx', '/autodiscover/autodiscover.svc') { return New-SimResponse 404 '<html><head><title>404 - File or directory not found.</title></head><body>Server Error</body></html>' $headers -ContentType 'text/html' }
    if ($State.OnPremProxy) { return New-SimResponse 302 '' @{ Location = 'https://sts.contoso.com/adfs/ls/?wa=wsignin1.0&wtrealm=https%3a%2f%2fmail.contoso.com%2f'; Server = 'BigIP' } }
    $isAvailability = $Body -match 'GetUserAvailabilityRequest'
    # Credential mode, NTLM / Negotiate: the handler answers the challenge with the account the tool put in the options
    # (the simulator sees no handshake). Accepted when the account is known and its scheme offered.
    $typed = $null
    $null = ([Collections.Generic.IDictionary[string, object]]$Request.Options).TryGetValue('GctWindowsCredential', [ref]$typed)
    $typedUser = $null
    if ($typed) {
        $scheme, $account = ([string]$typed).Split(' ', 2)
        if ($scheme -in 'Negotiate', 'NTLM') { $typedUser = $State.OnPremAccounts[$account.ToLowerInvariant()] }
    }
    $windowsSession = -not $auth -and $isAvailability -and $(if ($typed) { [bool]$typedUser } else { $State.WindowsAccepted })
    if (-not $auth -and -not $windowsSession) { return New-SimResponse 401 '' $headers $schemes }
    $caller = $null
    if ($auth -and $auth.Scheme -eq 'Basic') {
        $pair = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($auth.Parameter)).Split(':', 2)
        $mapped = $State.OnPremAccounts[$pair[0].ToLowerInvariant()]
        if (-not $State.OnPremBasic -or -not $mapped -or $pair[1] -cne $State.OnPremPassword) { return New-SimResponse 401 '' $headers $schemes }
        $caller = Find-SimUser $State $mapped
    }
    elseif ($auth) {
        $claims = Get-SimClaims $auth.Parameter
        $valid = $auth.Scheme -eq 'Bearer' -and $State.ValidTokens.Contains($auth.Parameter) -and $claims -and $claims['aud'] -eq "https://$($Request.RequestUri.Host)" -and $claims['tid'] -eq $State.OnPremTrustedTenant
        if (-not $valid) { $headers['x-ms-diagnostics'] = '2000003;reason="The audience claim value is invalid for current resource.";error_category="invalid_resource"'; return New-SimResponse 401 '' $headers $schemes }
        $caller = Find-SimUser $State ([string]$claims['upn'])
    }
    elseif ($typedUser) { $caller = Find-SimUser $State $typedUser }
    else { $caller = Find-SimUser $State $State.SignInUser }
    if (-not $isAvailability) { return New-SimResponse 400 '' $headers }
    return New-SimResponse 200 (Get-SimEwsAvailability -State $State -Source $caller -Soap $Body) $headers -ContentType 'text/xml; charset=utf-8'
}

#endregion

#region Entra ID ------------------------------------------------------------------------------------

function New-SimTokenAnswer {
    <# Tokens for a scope: the resource decides the audience and the permissions; offline_access gives a refresh token. #>
    param([hashtable]$State, [string]$Scope, [string]$ClientId, [string]$User)

    $items = @($Scope -split '\s+' | Where-Object { $_ })
    $resourceScope = $items | Where-Object { $_ -notin 'offline_access', 'openid', 'profile', 'email' } | Select-Object -First 1
    $resource = ([Uri]$resourceScope).GetLeftPart([UriPartial]::Authority)
    $hostName = ([Uri]$resource).Host
    $u = Find-SimUser $State $User
    if ($hostName -eq 'graph.microsoft.com') { $aud = 'https://graph.microsoft.com'; $scp = $State.GraphScopes }
    elseif ($hostName -eq 'outlook.office365.com') { $aud = 'https://outlook.office365.com'; $scp = 'Exchange.Manage EWS.AccessAsUser.All' }
    elseif ($hostName -in $State.OnPremSpns) { $aud = $resource; $scp = 'EWS.AccessAsUser.All' }
    else {
        return New-SimJson 400 ([ordered]@{ error = 'invalid_resource'; error_description = "AADSTS500011: The resource principal named $resource was not found in the tenant named Contoso. This can happen if the application has not been installed by the administrator of the tenant or consented to by any user in the tenant. Trace ID: $([guid]::NewGuid()) Correlation ID: $([guid]::NewGuid()) Timestamp: $([DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ssZ'))"; error_codes = @(500011); trace_id = [guid]::NewGuid().ToString(); correlation_id = [guid]::NewGuid().ToString() })
    }
    $access = New-SimToken -Audience $aud -Scope $scp -Upn $u.Upn -AppId $ClientId -TenantId $u.Tenant -ObjectId $u.Id -Name $u.DisplayName
    $State.ValidTokens.Add($access); $State.Secrets.Add($access)
    $answer = [ordered]@{ token_type = 'Bearer'; scope = $scp; expires_in = 4523; ext_expires_in = 4523; access_token = $access }
    if ($items -contains 'offline_access') {
        $refresh = '0.ARoAkSpOfVs8bk-KHSucDl9KN9YOWdOzUgJBrv-tItKasBwaAA.' + (ConvertTo-SimBase64Url ([Security.Cryptography.RandomNumberGenerator]::GetBytes(96)))
        $State.RefreshTokens[$refresh] = @{ User = $u.Upn; ClientId = $ClientId }
        $State.Secrets.Add($refresh)
        $answer.refresh_token = $refresh
    }
    if ($items -contains 'openid') {
        $id = New-SimToken -Audience $ClientId -Upn $u.Upn -AppId $ClientId -TenantId $u.Tenant -ObjectId $u.Id
        $State.Secrets.Add($id)
        $answer.id_token = $id
    }
    New-SimJson 200 $answer @{ 'x-ms-request-id' = [guid]::NewGuid().ToString(); 'Cache-Control' = 'no-store, no-cache' }
}

function Get-SimEntraResponse([hashtable]$State, [Net.Http.HttpRequestMessage]$Request, [string]$Body) {
    $path = [Uri]::UnescapeDataString($Request.RequestUri.AbsolutePath)
    $entraError = {
        param([string]$Code, [string]$Description, [int]$Number)
        New-SimJson 400 ([ordered]@{ error = $Code; error_description = "$Description Trace ID: $([guid]::NewGuid()) Correlation ID: $([guid]::NewGuid()) Timestamp: $([DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ssZ'))"; error_codes = @($Number); timestamp = [DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ssZ'); trace_id = [guid]::NewGuid().ToString(); correlation_id = [guid]::NewGuid().ToString() }) @{ 'x-ms-request-id' = [guid]::NewGuid().ToString() }
    }
    if ($path -match '^/([^/]+)/v2\.0/\.well-known/openid-configuration$') {
        $tid = Get-SimTenantOf $Matches[1]
        if (-not $tid) { return & $entraError 'invalid_tenant' "AADSTS90002: Tenant '$($Matches[1])' not found. Check to make sure you have the correct tenant ID and are signing into the correct cloud." 90002 }
        $region = switch ($tid) { $script:SimContoso { 'EU' } $script:SimFabrikam { 'NA' } default { $null } }
        $config = [ordered]@{
            token_endpoint = "https://login.microsoftonline.com/$tid/oauth2/v2.0/token"; authorization_endpoint = "https://login.microsoftonline.com/$tid/oauth2/v2.0/authorize"
            device_authorization_endpoint = "https://login.microsoftonline.com/$tid/oauth2/v2.0/devicecode"; issuer = "https://login.microsoftonline.com/$tid/v2.0"
            cloud_instance_name = 'microsoftonline.com'; msgraph_host = 'graph.microsoft.com'
        }
        if ($region) { $config.tenant_region_scope = $region }
        return New-SimJson 200 $config
    }
    if ($path -match '^/common/userrealm/(.+)$') {
        $address = $Matches[1]; $domain = ($address -split '@')[-1].ToLowerInvariant()
        $realm = switch ($domain) {
            'contoso.com' { [ordered]@{ ver = '2.0'; NameSpaceType = 'Managed'; Login = $address; DomainName = 'contoso.com'; FederationBrandName = 'Contoso'; TenantBrandingInfo = $null; cloud_instance_name = 'microsoftonline.com' } }
            'fabrikam.com' { [ordered]@{ ver = '2.0'; NameSpaceType = 'Federated'; Login = $address; DomainName = 'fabrikam.com'; FederationBrandName = 'Fabrikam'; AuthURL = "https://sts.fabrikam.com/adfs/ls/?username=$([Uri]::EscapeDataString($address))&wa=wsignin1.0&wtrealm=urn%3afederation%3aMicrosoftOnline"; cloud_instance_name = 'microsoftonline.com' } }
            { $_ -in 'gmail.com', 'outlook.com' } { [ordered]@{ ver = '2.0'; NameSpaceType = 'Unknown'; Login = $address; DomainName = $domain; ConsumerDomain = $true; cloud_instance_name = 'microsoftonline.com' } }
            default { [ordered]@{ ver = '2.0'; NameSpaceType = 'Unknown'; Login = $address; DomainName = $domain } }
        }
        return New-SimJson 200 $realm
    }
    $fields = Get-SimFormFields $Body
    if ($path -match '^/([^/]+)/oauth2/v2\.0/devicecode$') {
        $code = 'SIM-DC-' + [guid]::NewGuid().ToString('N') + (ConvertTo-SimBase64Url ([Security.Cryptography.RandomNumberGenerator]::GetBytes(24)))
        $State.DeviceCodes[$code] = @{ Scope = $fields.scope; ClientId = $fields.client_id; Polls = 0 }
        $State.Secrets.Add($code)
        return New-SimJson 200 ([ordered]@{ user_code = 'GQ7XK2PLM'; device_code = $code; verification_uri = 'https://microsoft.com/devicelogin'; expires_in = 900; interval = 5; message = 'To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code GQ7XK2PLM to authenticate.' })
    }
    if ($path -match '^/([^/]+)/oauth2/v2\.0/token$') {
        $userFor = { param([string]$ClientId) if ($ClientId -eq $script:SimAdminClient) { $State.AdminUser } else { $State.SignInUser } }
        switch ($fields.grant_type) {
            'urn:ietf:params:oauth:grant-type:device_code' {
                $entry = $State.DeviceCodes[[string]$fields.device_code]
                if (-not $entry) { return & $entraError 'invalid_grant' 'AADSTS70000: The provided value for the device_code is not valid.' 70000 }
                if ($State.DeviceDenied) { return & $entraError 'authorization_declined' 'AADSTS70000: The user declined the sign-in.' 70000 }
                if ($entry.Polls -lt $State.PendingPolls) { $entry.Polls++; return & $entraError 'authorization_pending' 'AADSTS70016: OAuth 2.0 device flow error. Authorization is pending. Continue polling.' 70016 }
                return New-SimTokenAnswer -State $State -Scope $entry.Scope -ClientId $entry.ClientId -User (& $userFor $entry.ClientId)
            }
            'refresh_token' {
                $entry = $State.RefreshTokens[[string]$fields.refresh_token]
                if (-not $entry) { return & $entraError 'invalid_grant' 'AADSTS9002313: Invalid request. Request is malformed or invalid.' 9002313 }
                return New-SimTokenAnswer -State $State -Scope $fields.scope -ClientId $fields.client_id -User $entry.User
            }
            'authorization_code' {
                $entry = $State.AuthCodes[[string]$fields.code]
                if (-not $entry) { return & $entraError 'invalid_grant' 'AADSTS70008: The provided authorization code or refresh token has expired due to inactivity.' 70008 }
                $hash = ConvertTo-SimBase64Url ([Security.Cryptography.SHA256]::HashData([Text.Encoding]::ASCII.GetBytes([string]$fields.code_verifier)))
                if ($hash -ne $entry.Challenge) { return & $entraError 'invalid_grant' 'AADSTS501481: The Code_Verifier does not match the code_challenge supplied in the authorization request.' 501481 }
                if ($fields.redirect_uri -ne $entry.RedirectUri) { return & $entraError 'invalid_grant' 'AADSTS50011: The redirect URI specified in the request does not match the redirect URIs configured for the application.' 50011 }
                return New-SimTokenAnswer -State $State -Scope $fields.scope -ClientId $fields.client_id -User (& $userFor $fields.client_id)
            }
            'client_credentials' {
                $assertion = [string]$fields.client_assertion
                if ($assertion) { $State.Secrets.Add($assertion) }
                $header = if ($assertion) { Get-SimClaims ('x.' + $assertion.Split('.')[0]) } else { $null }
                if (-not $header -or -not $header['x5t'] -or $fields.client_assertion_type -ne 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer') { return & $entraError 'invalid_client' 'AADSTS700027: Client assertion failed signature validation.' 700027 }
                $scope = [string]$fields.scope
                $aud = ([Uri]$scope).GetLeftPart([UriPartial]::Authority)
                $tid = Get-SimTenantOf ($path -split '/')[1]
                $roles = if ($aud -eq 'https://outlook.office365.com') { @('Exchange.ManageAsApp') } else { $State.AppRoles }
                $access = New-SimToken -Audience $aud -Roles $roles -AppId $fields.client_id -TenantId $tid
                $State.ValidTokens.Add($access); $State.Secrets.Add($access)
                return New-SimJson 200 ([ordered]@{ token_type = 'Bearer'; expires_in = 3599; ext_expires_in = 3599; access_token = $access })
            }
        }
        return & $entraError 'unsupported_grant_type' 'AADSTS70003: The app requested an unsupported grant type.' 70003
    }
    return New-SimResponse 404 '' -ContentType 'text/html'
}

#endregion

#region Microsoft Graph -----------------------------------------------------------------------------

function ConvertTo-SimGraphUser([hashtable]$U) {
    [ordered]@{
        id = $U.Id; displayName = $U.DisplayName; userPrincipalName = $U.Upn; mail = $U.Address; userType = 'Member'; accountEnabled = $U.Enabled
        onPremisesSyncEnabled = $(if ($U.Synced) { $true } else { $null }); onPremisesImmutableId = $U.ImmutableId; usageLocation = 'FR'; createdDateTime = '2024-03-11T09:12:44Z'
        proxyAddresses = @("SMTP:$($U.Address)") + @($U.Aliases | ForEach-Object { "smtp:$_" })
        assignedPlans = @(if ($U.Licensed) { [ordered]@{ assignedDateTime = '2024-03-11T09:20:00Z'; capabilityStatus = 'Enabled'; service = 'exchange'; servicePlanId = 'efb87545-963c-4e0d-99df-69c6916d9eb0' }; [ordered]@{ assignedDateTime = '2024-03-11T09:20:00Z'; capabilityStatus = 'Enabled'; service = 'TeamspaceAPI'; servicePlanId = '57ff2da0-773e-42df-b2af-ffb7a2317929' } })
    }
}

function Get-SimGraphResponse {
    param([hashtable]$State, [Net.Http.HttpRequestMessage]$Request, [string]$Body)

    $uri = $Request.RequestUri
    $method = $Request.Method.Method
    $path = [Uri]::UnescapeDataString($uri.AbsolutePath) -replace '^/v1\.0', ''
    $query = [Uri]::UnescapeDataString($uri.Query)
    $auth = $Request.Headers.Authorization
    $headers = @{ 'request-id' = [guid]::NewGuid().ToString(); 'x-ms-ags-diagnostic' = '{"ServerInfo":{"DataCenter":"France Central","Slice":"E","Ring":"3","ScaleUnit":"001","RoleInstance":"PA1PEPF0000SIM1"}}'; 'Strict-Transport-Security' = 'max-age=31536000' }
    $cid = $null; if ($Request.Headers.TryGetValues('client-request-id', [ref]$cid)) { $headers['client-request-id'] = @($cid)[0] }
    $send = { param([int]$Code, $Object) $h = $headers.Clone(); if ($null -eq $Object) { New-SimResponse $Code '' $h } else { $h['Content-Type'] = 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false;charset=utf-8'; New-SimResponse $Code (ConvertTo-Json -InputObject $Object -Depth 12 -Compress) $h } }
    $fail = { param([int]$Code, [string]$Err, [string]$Message) & $send $Code ([ordered]@{ error = [ordered]@{ code = $Err; message = $Message; innerError = [ordered]@{ date = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss'); 'request-id' = $headers['request-id']; 'client-request-id' = $headers['client-request-id'] } } }) }
    $noMailbox = { & $fail 404 'MailboxNotEnabledForRESTAPI' 'The mailbox is either inactive, soft-deleted, or is hosted on-premise.' }
    $denied = { & $fail 403 'ErrorAccessDenied' 'Access is denied. Check credentials and try again.' }

    if (-not $auth -or $auth.Scheme -ne 'Bearer' -or -not $auth.Parameter) { return New-SimResponse 401 '{"error":{"code":"InvalidAuthenticationToken","message":"Access token is empty."}}' $headers.Clone() @('Bearer realm="", authorization_uri="https://login.microsoftonline.com/common/oauth2/authorize", client_id="00000003-0000-0000-c000-000000000000"') }
    $claims = Get-SimClaims $auth.Parameter
    if (-not $State.ValidTokens.Contains($auth.Parameter) -or -not $claims) { return & $fail 401 'InvalidAuthenticationToken' 'IDX14100: JWT is not well formed, there are no dots (.).' }
    if ($claims['aud'] -notin 'https://graph.microsoft.com', 'https://graph.microsoft.com/', '00000003-0000-0000-c000-000000000000') { return & $fail 401 'InvalidAuthenticationToken' 'Access token validation failure. Invalid audience.' }
    if ([long]$claims['exp'] -lt [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()) { return & $fail 401 'InvalidAuthenticationToken' 'Lifetime validation failed, the token is expired.' }
    $isApp = -not $claims['scp']
    $granted = if ($isApp) { @($claims['roles']) } else { @(([string]$claims['scp']) -split ' ') }
    $tenant = [string]$claims['tid']
    $caller = Find-SimUser $State ([string]$claims['upn'])
    $has = { param([string]$Pattern) [bool](@($granted | Where-Object { $_ -match $Pattern }).Count) }
    $data = if ($Body) { $Body | ConvertFrom-Json -AsHashtable } else { @{} }

    if ($State.ThrottleOnce -and $path -match [regex]::Escape($State.ThrottleOnce)) {
        $State.ThrottleOnce = $null
        $h = $headers.Clone(); $h['Retry-After'] = '2'
        return New-SimResponse 429 '{"error":{"code":"ApplicationThrottled","message":"Application is over its MailboxConcurrency limit."}}' $h
    }

    # The mailbox a /me or /users/{id} path names.
    $owner = $null
    if ($path -match '^/me(/|$)') {
        if ($isApp) { return & $fail 400 'BadRequest' '/me request is only valid with delegated authentication flow.' }
        $owner = $caller; $rest = $path.Substring(3)
    }
    elseif ($path -match '^/users/([^/]+)(/.*)?$') {
        $owner = Find-SimUser $State $Matches[1]; $rest = [string]$Matches[2]
        if (-not $owner -or $owner.Tenant -ne $tenant) { return & $fail 404 'Request_ResourceNotFound' "Resource '$($Matches[1])' does not exist or one of its queried reference-property objects are not present." }
    }
    else { $rest = $null }

    if ($null -ne $rest) {
        $mailboxRoute = $rest -match '^/(calendar|getMailTips|mailFolders|mailboxSettings)'
        if ($mailboxRoute -and $owner.Location -ne 'EXO') { return & $noMailbox }
        switch -Regex ("$method $rest") {
            '^GET $' { return & $send 200 (ConvertTo-SimGraphUser $owner) }
            '^POST /calendar/getSchedule$' {
                if (-not (& $has '^Calendars\.')) { return & $denied }
                $fbStart = [datetime]::SpecifyKind([datetime]::Parse([string]$data.startTime.dateTime, [Globalization.CultureInfo]::InvariantCulture), 'Utc')
                $fbEnd = [datetime]::SpecifyKind([datetime]::Parse([string]$data.endTime.dateTime, [Globalization.CultureInfo]::InvariantCulture), 'Utc')
                $interval = [int]$data.availabilityViewInterval
                $status = @{ Busy = 'busy'; Tentative = 'tentative'; OOF = 'oof'; WorkingElsewhere = 'workingElsewhere'; Free = 'free' }
                $values = foreach ($a in @($data.schedules)) {
                    $o = Get-SimFreeBusyOutcome -State $State -Source $owner -Address $a -From Cloud
                    if ($o.Code -ne 'Success') { [ordered]@{ scheduleId = $a; availabilityView = ''; error = [ordered]@{ message = $o.Message; responseCode = $o.Code } }; continue }
                    $plan = Get-SimSchedule -Mailbox $a -StartUtc $fbStart -Days ([Math]::Max(1, [int]($fbEnd - $fbStart).TotalDays)) -IntervalMinutes $interval
                    $items = @($plan.Events | ForEach-Object {
                            $i = [ordered]@{ isPrivate = $false; status = $status[$_.BusyType]; start = [ordered]@{ dateTime = $_.Start.ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' }; end = [ordered]@{ dateTime = $_.End.ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' } }
                            if ($o.Detail -eq 'Full') { $i.subject = $_.Subject; $i.location = $_.Location; $i.isMeeting = $true; $i.isRecurring = $false; $i.isException = $false; $i.isReminderSet = $true }
                            $i
                        })
                    [ordered]@{ scheduleId = $a; availabilityView = $plan.View; scheduleItems = $items; workingHours = [ordered]@{ daysOfWeek = @('monday', 'tuesday', 'wednesday', 'thursday', 'friday'); startTime = '08:00:00.0000000'; endTime = '17:00:00.0000000'; timeZone = [ordered]@{ name = 'Romance Standard Time' } } }
                }
                return & $send 200 ([ordered]@{ '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.scheduleInformation)'; value = @($values) })
            }
            '^POST /getMailTips$' {
                $values = foreach ($a in @($data.EmailAddresses)) {
                    $u = Find-SimUser $State $a
                    $domainTenant = Get-SimTenantOf $a
                    if (-not $u -and $domainTenant -eq $owner.Tenant) { [ordered]@{ emailAddress = [ordered]@{ name = ''; address = $a }; error = [ordered]@{ message = "The recipient $a was not found."; code = 'ErrorInvalidRecipients' } }; continue }
                    $scope = if ($u -and $u.Tenant -eq $owner.Tenant) { 'internal' } elseif ($u -and $State.CrossOrg -eq 'Success') { 'externalPartner' } else { 'external' }
                    $oof = if ($u -and $u.Address -eq 'bob@contoso.com' -and $State.BobOutOfOffice) { '<html><body>I am out of the office until Monday.</body></html>' } else { '' }
                    [ordered]@{ emailAddress = [ordered]@{ name = $(if ($u) { $u.DisplayName } else { '' }); address = $a }; automaticReplies = [ordered]@{ message = $oof; messageLanguage = [ordered]@{ locale = 'en-US'; displayName = 'English (United States)' } }; mailboxFull = $false; maxMessageSize = 37748736; recipientScope = $scope; deliveryRestricted = $false; isModerated = $false; totalMemberCount = 1; externalMemberCount = 0 }
                }
                return & $send 200 ([ordered]@{ '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.mailTips)'; value = @($values) })
            }
            '^GET /mailFolders/delta$' {
                if (-not (& $has '^Mail\.')) { return & $denied }
                $base = "https://graph.microsoft.com/v1.0$($path)"
                if ($query -match 'deltatoken=') { return & $send 200 ([ordered]@{ value = @(); '@odata.deltaLink' = "$base`?`$deltatoken=SIM-DELTA-F2" }) }
                if ($query -match 'skiptoken=') { return & $send 200 ([ordered]@{ value = @(foreach ($n in 'Archive', 'Conversation History', 'Projects') { [ordered]@{ id = "F-$n"; displayName = $n; parentFolderId = 'F-ROOT' } }); '@odata.deltaLink' = "$base`?`$deltatoken=SIM-DELTA-F1" }) }
                return & $send 200 ([ordered]@{ value = @(foreach ($n in 'Inbox', 'Drafts', 'Sent Items', 'Deleted Items', 'Junk Email') { [ordered]@{ id = "F-$n"; displayName = $n; parentFolderId = 'F-ROOT' } }); '@odata.nextLink' = "$base`?`$skiptoken=SIM-SKIP-1" })
            }
            '^GET /mailFolders/inbox/messages/delta$' {
                if (-not (& $has '^Mail\.')) { return & $denied }
                $base = "https://graph.microsoft.com/v1.0$($path)"
                if ($query -match 'deltatoken=') { return & $send 200 ([ordered]@{ value = @([ordered]@{ id = 'M-4'; subject = 'New message since the first synchronization'; receivedDateTime = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ') }); '@odata.deltaLink' = "$base`?`$deltatoken=SIM-DELTA-M2" }) }
                return & $send 200 ([ordered]@{ value = @(1..3 | ForEach-Object { [ordered]@{ id = "M-$_"; subject = "Message $_"; receivedDateTime = [DateTime]::UtcNow.AddDays(-$_).ToString('yyyy-MM-ddTHH:mm:ssZ') } }); '@odata.deltaLink' = "$base`?`$deltatoken=SIM-DELTA-M1" })
            }
            '^GET /mailboxSettings/automaticRepliesSetting$' {
                if ($State.AutoReplyForbidden) { return & $denied }
                $o = [ordered]@{ '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#users(''alice'')/mailboxSettings/automaticRepliesSetting' }
                foreach ($k in $State.AutoReply.Keys) { $o[$k] = $State.AutoReply[$k] }
                return & $send 200 $o
            }
            '^PATCH /mailboxSettings$' {
                if (-not (& $has '^MailboxSettings\.ReadWrite$')) { return & $denied }
                $State.Patches.Add($data)
                $set = $data.automaticRepliesSetting
                foreach ($k in $set.Keys) { $State.AutoReply[$k] = $set[$k] }
                return & $send 200 ([ordered]@{ automaticRepliesSetting = $State.AutoReply })
            }
            '^GET /licenseDetails$' {
                $plans = @(if ($owner.Licensed) { [ordered]@{ id = 'L-1'; skuId = '6fd2c87f-b296-42f0-b197-1e91e994b900'; skuPartNumber = 'ENTERPRISEPACK'; servicePlans = @([ordered]@{ servicePlanName = 'EXCHANGE_S_ENTERPRISE'; provisioningStatus = 'Success'; appliesTo = 'User' }, [ordered]@{ servicePlanName = 'TEAMS1'; provisioningStatus = 'Success'; appliesTo = 'User' }) } })
                return & $send 200 ([ordered]@{ value = $plans })
            }
            '^GET /mailFolders/inbox$' {
                if ($isApp -or ($caller -and $caller.Id -eq $owner.Id)) { return & $send 200 ([ordered]@{ id = 'AAMkAGI2TG93AAA=' }) }
                return & $denied
            }
            '^GET /mailboxSettings/userPurpose$' { return & $send 200 ([ordered]@{ '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#users(''x'')/mailboxSettings/userPurpose'; value = $owner.Purpose }) }
        }
        return & $fail 400 'BadRequest' "Not simulated: $method $path"
    }

    if ($method -eq 'GET' -and $path -eq '/users') {
        $inTenant = @($State.Users.Values | Where-Object { $_.Tenant -eq $tenant })
        if ($query -match 'proxyAddresses/any') {
            $consistency = $null
            if (-not $Request.Headers.TryGetValues('ConsistencyLevel', [ref]$consistency) -or @($consistency)[0] -ne 'eventual') { return & $fail 400 'Request_UnsupportedQuery' "Unsupported or invalid query filter clause specified for property 'proxyAddresses' of resource 'User'." }
            $a = [regex]::Match($query, "p eq 'smtp:((?:[^']|'')*)'").Groups[1].Value.Replace("''", "'")
            $found = @($inTenant | Where-Object { $_.Aliases -contains $a.ToLowerInvariant() -or $_.Address -eq $a.ToLowerInvariant() })
            return & $send 200 ([ordered]@{ '@odata.count' = $found.Count; value = @($found | ForEach-Object { ConvertTo-SimGraphUser $_ }) })
        }
        $a = [regex]::Match($query, "mail eq '((?:[^']|'')*)'").Groups[1].Value.Replace("''", "'").ToLowerInvariant()
        $found = @($inTenant | Where-Object { $_.Address -eq $a -or $_.Upn -eq $a })
        return & $send 200 ([ordered]@{ value = @($found | ForEach-Object { ConvertTo-SimGraphUser $_ }) })
    }
    if ($method -eq 'GET' -and $path -match '^/domains/([^/]+)(/serviceConfigurationRecords)?$') {
        $domain = $Matches[1].ToLowerInvariant()
        if ($domain -ne 'contoso.com' -or $tenant -ne $script:SimContoso) { return & $fail 404 'Request_ResourceNotFound' "Resource '$domain' does not exist or one of its queried reference-property objects are not present." }
        if (-not $Matches[2]) { return & $send 200 ([ordered]@{ id = 'contoso.com'; authenticationType = 'Managed'; isAdminManaged = $true; isDefault = $true; isInitial = $false; isRoot = $true; isVerified = $true; supportedServices = @('Email', 'OfficeCommunicationsOnline', 'Intune') }) }
        return & $send 200 ([ordered]@{ value = @(
                    [ordered]@{ id = 'R-1'; isOptional = $false; label = 'contoso.com'; recordType = 'Mx'; supportedService = 'Email'; ttl = 3600; mailExchange = 'contoso-com.mail.protection.outlook.com'; preference = 0 }
                    [ordered]@{ id = 'R-2'; isOptional = $false; label = 'contoso.com'; recordType = 'Txt'; supportedService = 'Email'; ttl = 3600; text = 'v=spf1 include:spf.protection.outlook.com -all' }
                    [ordered]@{ id = 'R-3'; isOptional = $false; label = 'autodiscover.contoso.com'; recordType = 'CName'; supportedService = 'Email'; ttl = 3600; canonicalName = 'autodiscover.outlook.com' }
                    [ordered]@{ id = 'R-4'; isOptional = $false; label = 'enterpriseregistration.contoso.com'; recordType = 'CName'; supportedService = 'Intune'; ttl = 3600; canonicalName = 'enterpriseregistration.windows.net' }
                    [ordered]@{ id = 'R-5'; isOptional = $false; label = 'enterpriseenrollment.contoso.com'; recordType = 'CName'; supportedService = 'Intune'; ttl = 3600; canonicalName = 'enterpriseenrollment-s.manage.microsoft.com' }
                    [ordered]@{ id = 'R-6'; isOptional = $true; label = '_sipfederationtls._tcp.contoso.com'; recordType = 'Srv'; supportedService = 'OfficeCommunicationsOnline'; ttl = 3600; nameTarget = 'sipfed.online.lync.com'; port = 5061; priority = 100; weight = 1; protocol = '_tcp'; service = '_sipfederationtls' }
                ) })
    }
    return & $fail 400 'BadRequest' "Not simulated: $method $path"
}

#endregion

#region Exchange Online -----------------------------------------------------------------------------

function Get-SimOutlookResponse([hashtable]$State, [Net.Http.HttpRequestMessage]$Request, [string]$Body) {
    $path = [Uri]::UnescapeDataString($Request.RequestUri.AbsolutePath)
    $headers = @{ 'request-id' = [guid]::NewGuid().ToString(); 'X-CalculatedBETarget' = 'PA4P190MB1234.EURP190.PROD.OUTLOOK.COM'; 'X-FEServer' = 'PA7P264CA0123' }
    if ($path -match '^/autodiscover/autodiscover\.json/v1\.0/(.+)$') {
        $address = $Matches[1]
        $u = Find-SimUser $State $address
        if ($u -and $u.Location -eq 'OnPrem') { return New-SimResponse 302 '' (@{ Location = "https://autodiscover.contoso.com/autodiscover/autodiscover.json/v1.0/$([Uri]::EscapeDataString($address))?Protocol=EWS" } + $headers) }
        # Exchange Online answers for any address, known or not.
        return New-SimJson 200 ([ordered]@{ Protocol = 'EWS'; Url = 'https://outlook.office365.com/EWS/Exchange.asmx' }) $headers
    }
    if ($path -match '^/adminapi/beta/([^/]+)/InvokeCommand$') {
        $auth = $Request.Headers.Authorization
        $claims = if ($auth) { Get-SimClaims $auth.Parameter } else { $null }
        if (-not $claims -or -not $State.ValidTokens.Contains($auth.Parameter) -or $claims['aud'] -ne 'https://outlook.office365.com' -or $claims['tid'] -ne $Matches[1]) { return New-SimJson 401 ([ordered]@{ error = [ordered]@{ code = 'Unauthorized'; message = 'Access denied: the token is not valid for this tenant.' } }) $headers }
        if ($State.AdminForbidden) { return New-SimJson 403 ([ordered]@{ error = [ordered]@{ code = 'Forbidden'; message = "The user isn't assigned to any management roles." } }) $headers }
        $name = (ConvertFrom-Json $Body -AsHashtable).CmdletInput.CmdletName
        if (-not $State.AdminCommands.ContainsKey($name)) { return New-SimJson 400 ([ordered]@{ error = [ordered]@{ code = 'BadRequest'; message = "The term '$name' is not recognized." } }) $headers }
        return New-SimJson 200 ([ordered]@{ '@odata.context' = "https://outlook.office365.com/adminapi/beta/$($Matches[1])/`$metadata#Collection(Exchange.GenericHashTable)"; 'adminapi.warnings@odata.type' = '#Collection(String)'; '@adminapi.warnings' = @(); value = @($State.AdminCommands[$name]) }) $headers
    }
    return New-SimResponse 404 '' $headers
}

#endregion

function Get-SimHttpResponse {
    <# Answer to one HTTP request of the tool. Every call is recorded in State.Calls. #>
    param([Parameter(Mandatory = $true)][Net.Http.HttpRequestMessage]$Request, [Parameter(Mandatory = $true)][hashtable]$State)

    $uri = $Request.RequestUri
    $body = if ($Request.Content) { [Text.Encoding]::UTF8.GetString($Request.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()) } else { '' }
    $auth = $Request.Headers.Authorization
    $State.Calls.Add([pscustomobject]@{ Method = $Request.Method.Method; Url = $uri.AbsoluteUri; Host = $uri.Host; Path = [Uri]::UnescapeDataString($uri.AbsolutePath); Scheme = if ($auth) { $auth.Scheme } else { $null }; Body = $body })
    switch ($uri.Host.ToLowerInvariant()) {
        'login.microsoftonline.com' { return Get-SimEntraResponse -State $State -Request $Request -Body $body }
        'graph.microsoft.com' { return Get-SimGraphResponse -State $State -Request $Request -Body $body }
        'outlook.office365.com' { return Get-SimOutlookResponse -State $State -Request $Request -Body $body }
        'autodiscover.contoso.com' {
            if ($uri.AbsolutePath -match '^/autodiscover/autodiscover\.json/v1\.0/') { return New-SimJson 200 ([ordered]@{ Protocol = 'EWS'; Url = 'https://mail.contoso.com/EWS/Exchange.asmx' }) @{ 'X-FEServer' = 'EXCH01'; 'request-id' = [guid]::NewGuid().ToString() } }
            return New-SimResponse 404 ''
        }
        'mail.contoso.com' { return Get-SimOnPremResponse -State $State -Request $Request -Body $body }
        'mta-sts.contoso.com' {
            if ($uri.AbsolutePath -eq '/.well-known/mta-sts.txt') { return New-SimResponse 200 "version: STSv1`nmode: enforce`nmx: *.mail.protection.outlook.com`nmax_age: 604800`n" -ContentType 'text/plain' }
            return New-SimResponse 404 ''
        }
    }
    throw [Net.Http.HttpRequestException]::new("No such host is known. ($($uri.Host):443)", [Net.Sockets.SocketException]::new(11001))
}

function Get-SimDnsAnswer {
    <# Like Resolve-DnsName -DnsOnly: the CNAME chain, then the records of the type asked; nothing for NXDOMAIN. #>
    param([hashtable]$State, [string]$Name, [string]$Type)
    if ($State.DnsDown) { throw 'This operation returned because the timeout period expired' }
    $n = $Name.ToLowerInvariant().TrimEnd('.')
    $records = [Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt 5 -and $Type -ne 'CNAME' -and $State.Dns.ContainsKey("$n|CNAME"); $i++) {
        $target = @($State.Dns["$n|CNAME"])[0]
        $records.Add([pscustomobject]@{ Type = 'CNAME'; Name = $n; Data = $target })
        $n = $target
    }
    foreach ($d in @($State.Dns["$n|$Type"] | Where-Object { $_ })) { $records.Add([pscustomobject]@{ Type = $Type; Name = $n; Data = $d }) }
    return @($records)
}

function Get-SimCertificate([string]$HostName, [int]$Port, [hashtable]$State) {
    $State.Calls.Add([pscustomobject]@{ Method = 'TLS'; Url = "tls://$($HostName):$Port"; Host = $HostName; Path = ''; Scheme = $null; Body = '' })
    $result = [ordered]@{ HostName = $HostName; Port = $Port; Reachable = $true; Valid = $true; Interrupted = $false; Subject = "CN=$HostName, O=Contoso Ltd, C=FR"; Issuer = 'CN=DigiCert Global G2 TLS RSA SHA256 2020 CA1, O=DigiCert Inc, C=US'; NotAfterUtc = [DateTime]::UtcNow.AddDays($State.CertificateDays).ToString('yyyy-MM-ddTHH:mm:ssZ'); DaysLeft = $State.CertificateDays; Protocol = 'Tls13'; Error = $null }
    switch ($State.TlsFailure) {
        'Unreachable' { $result.Reachable = $false; $result.Valid = $false; $result.Subject = $null; $result.Error = 'A connection attempt failed because the connected party did not properly respond after a period of time' }
        'Untrusted' { $result.Valid = $false; $result.Subject = $null; $result.Error = 'The remote certificate is invalid: RemoteCertificateChainErrors (UntrustedRoot)' }
    }
    [pscustomobject]$result
}

function Install-SimNetwork {
    <#
        Replaces the network inside the loaded module: Send-GctHttpRequest answers with Get-SimHttpResponse,
        Invoke-GctDnsQuery with Get-SimDnsAnswer, the TLS check with Get-SimCertificate; the waits are not waited,
        no page is opened, the sign-in window returns an authorization code, the certificate of the application
        is a self-signed one in memory. Remove-Module (or Import-Module -Force) restores them.
    #>
    param([Parameter(Mandatory = $true)][psmoduleinfo]$Module, [Parameter(Mandatory = $true)][hashtable]$State)

    $rsa = [Security.Cryptography.RSA]::Create(2048)
    $certificate = [Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=Contoso free-busy checker', $rsa, 'SHA256', [Security.Cryptography.RSASignaturePadding]::Pkcs1).CreateSelfSigned([DateTimeOffset]::UtcNow.AddDays(-1), [DateTimeOffset]::UtcNow.AddDays(365))
    $hooks = @{ Http = ${function:Get-SimHttpResponse}; Dns = ${function:Get-SimDnsAnswer}; Tls = ${function:Get-SimCertificate}; Certificate = $certificate }
    . $Module {
        param($State, $Hooks)
        $script:SimState = $State; $script:SimHooks = $Hooks
        function script:Send-GctHttpRequest { param($HttpClient, $Request) & $script:SimHooks.Http -Request $Request -State $script:SimState }
        function script:Invoke-GctDnsQuery { param([string]$Name, [string]$Type) & $script:SimHooks.Dns -State $script:SimState -Name $Name -Type $Type }
        function script:Get-GctTlsCertificate { param([string]$HostName, [int]$Port = 443, [int]$TimeoutSeconds) & $script:SimHooks.Tls $HostName $Port $script:SimState }
        function script:Wait-GctSeconds { param([double]$Seconds) $script:SimState.Waits.Add($Seconds) }
        function script:Start-Process { }
        function script:Test-GctDesktopSession { [bool]$script:SimState.DesktopSession }
        function script:Find-GctBrowser { param([switch]$All) $b = [pscustomobject]@{ Name = 'Microsoft Edge'; Path = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' }; if ($All) { @($b) } else { $b } }
        function script:Test-GctBrowserPolicyBlock { param($Browser) $false }
        function script:Invoke-GctBrowserAuthorization {
            param($Browser, [string]$Url, [string]$RedirectUri, [int]$TimeoutSeconds, [int]$StartTimeoutSeconds)
            $q = @{}; foreach ($p in $Url.Split('?', 2)[1].Split('&')) { $k, $v = $p.Split('=', 2); $q[$k] = [Uri]::UnescapeDataString($v) }
            $code = '1.ARoAkSpOfVs8bk-SIMCODE.' + [Convert]::ToBase64String([Security.Cryptography.RandomNumberGenerator]::GetBytes(48)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
            $script:SimState.AuthCodes[$code] = @{ Challenge = $q.code_challenge; RedirectUri = $q.redirect_uri; Scope = $q.scope }
            $script:SimState.Secrets.Add($code)
            "$RedirectUri`?code=$code&state=$($q.state)&session_state=$([guid]::NewGuid())"
        }
        function script:Get-GctApplicationCertificate { param([string]$Thumbprint) [pscustomobject]@{ Certificate = $script:SimHooks.Certificate; Store = 'Cert:\CurrentUser\My' } }
    } $State $hooks
}
