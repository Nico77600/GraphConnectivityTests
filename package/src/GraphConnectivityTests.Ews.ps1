<#
.SYNOPSIS
    Graph Connectivity Tests - EWS GetUserAvailability for a source mailbox on-premises (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Microsoft Graph reaches only the mailboxes of Exchange Online. When the source mailbox is on-premises
    (direction on-premises to Microsoft 365), the request starts where Outlook would send it: the Availability service of Exchange on-premises
    (GetUserAvailability), which then asks Exchange Online through its IntraOrganizationConnector or its
    organization relationship.

    Authentication (Identity.OnPremisesAuthentication):
      OAuth       hybrid modern authentication (default): the token of the same sign-in, for the on-premises
                  URL, from the refresh token - no new prompt.
      Windows     the current Windows account (Negotiate: Kerberos, else NTLM), from a computer of the domain.
      Credential  an account and its password typed by the operator (-OnPremisesCredential, Get-Credential, or
                  the fields of the window), from any
                  computer: NTLM, Negotiate or Basic (Identity.OnPremisesCredentialScheme; Auto = the best scheme
                  EWS offers, Negotiate > NTLM > Basic). The password stays in memory, never in a file, the log
                  or the trace (Basic: the user name only).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:EwsNs = @{
    soap = 'http://schemas.xmlsoap.org/soap/envelope/'
    t    = 'http://schemas.microsoft.com/exchange/services/2006/types'
    m    = 'http://schemas.microsoft.com/exchange/services/2006/messages'
}

function ConvertTo-GctXmlText {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ($null -eq $Text) { return '' }
    return [Security.SecurityElement]::Escape($Text)
}

function New-GctAvailabilityEnvelope {
    <# GetUserAvailability for several mailboxes, times in UTC (zone with a bias of 0), DetailedMerged view. #>
    param([Parameter(Mandatory = $true)][string[]]$Mailboxes, [Parameter(Mandatory = $true)][datetime]$StartUtc, [Parameter(Mandatory = $true)][datetime]$EndUtc, [int]$IntervalMinutes = 30)

    $zone = '<t:TimeZone><t:Bias>0</t:Bias><t:StandardTime><t:Bias>0</t:Bias><t:Time>00:00:00</t:Time><t:DayOrder>0</t:DayOrder><t:Month>0</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:StandardTime>' +
        '<t:DaylightTime><t:Bias>0</t:Bias><t:Time>00:00:00</t:Time><t:DayOrder>0</t:DayOrder><t:Month>0</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek></t:DaylightTime></t:TimeZone>'
    $data = ($Mailboxes | ForEach-Object { '<t:MailboxData><t:Email><t:Address>{0}</t:Address><t:RoutingType>SMTP</t:RoutingType></t:Email><t:AttendeeType>Required</t:AttendeeType><t:ExcludeConflicts>false</t:ExcludeConflicts></t:MailboxData>' -f (ConvertTo-GctXmlText $_) }) -join ''
    $f = 'yyyy-MM-ddTHH:mm:ss'
    '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '<soap:Header><t:RequestServerVersion Version="Exchange2016"/></soap:Header><soap:Body><m:GetUserAvailabilityRequest>' + $zone + '<m:MailboxDataArray>' + $data + '</m:MailboxDataArray><t:FreeBusyViewOptions><t:TimeWindow>' +
    ('<t:StartTime>{0}</t:StartTime><t:EndTime>{1}</t:EndTime></t:TimeWindow><t:MergedFreeBusyIntervalInMinutes>{2}</t:MergedFreeBusyIntervalInMinutes><t:RequestedView>DetailedMerged</t:RequestedView>' -f $StartUtc.ToString($f), $EndUtc.ToString($f), $IntervalMinutes) +
    '</t:FreeBusyViewOptions></m:GetUserAvailabilityRequest></soap:Body></soap:Envelope>'
}

function Get-GctNodeText {
    param([AllowNull()][Xml.XmlNode]$Node, [Parameter(Mandatory = $true)][string]$XPath, [Parameter(Mandatory = $true)][Xml.XmlNamespaceManager]$Ns)
    if (-not $Node) { return $null }
    $n = $Node.SelectSingleNode($XPath, $Ns)
    if ($n) { return $n.InnerText }
    return $null
}

function Get-GctOnPremisesCredential {
    <# The account of the Credential mode: the one given (-OnPremisesCredential, the window), else asked now (Get-Credential), once per run. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    if ($Context.OnPremisesCredential) { return $Context.OnPremisesCredential }
    if ($script:Ui) { throw 'On-premises authentication Credential: type the user name and the password of the on-premises account in the window.' }
    $user = if ([string]$Context.Config.OnPremisesUserName) { [string]$Context.Config.OnPremisesUserName } else { [string]$Context.Config.SourceMailbox }
    $credential = $null
    try { $credential = Get-Credential -UserName $user -Message "Account of the source mailbox on-premises, for EWS ($([string]$Context.Config.OnPremisesCredentialScheme)): DOMAIN\user or user@domain" }
    catch { throw "On-premises authentication Credential: no account given and none can be asked here ($($_.Exception.Message)). Pass -OnPremisesCredential (Get-Credential)." }
    if (-not $credential) { throw 'Cancelled by the operator.' }
    $Context.OnPremisesCredential = $credential
    return $credential
}

function Select-GctCredentialScheme {
    <# The scheme of the Credential mode: the one asked, or with Auto the best one EWS offers (Negotiate > NTLM > Basic); $null when none. #>
    param([AllowEmptyString()][string]$Requested, [AllowEmptyCollection()][string[]]$Offered)
    if ($Requested -and $Requested -ne 'Auto') { return $Requested }
    foreach ($scheme in 'Negotiate', 'NTLM', 'Basic') { if (@($Offered) -icontains $scheme) { return $scheme } }
    return $null
}

function Invoke-GctEwsAvailability {
    <#
        GetUserAvailability on the on-premises EWS of the source, with X-AnchorMailbox of the source. Returns the
        HTTP answer, the XML read, the SOAP fault, and the authentication used.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$EwsUrl, [Parameter(Mandatory = $true)][string]$Envelope)

    $cfg = $Context.Config
    $mode = [string]$cfg.OnPremisesAuthentication
    $uri = [Uri]$EwsUrl
    $client = $Context.HttpClient
    $handler = $null
    $how = $null
    $basic = $null
    $label = $null
    $windowsScheme = $null
    $offered = @()
    if ($mode -eq 'Credential') {
        $credential = Get-GctOnPremisesCredential -Context $Context
        $network = $credential.GetNetworkCredential()
        $account = if ($network.Domain) { "$($network.Domain)\$($network.UserName)" } else { $network.UserName }
        # The schemes EWS offers: those the on-premises stage saw, else one request without credentials now.
        if ($Context.OnPremisesSchemes.ContainsKey($uri.Host)) { $offered = @($Context.OnPremisesSchemes[$uri.Host]) }
        elseif ([string]$cfg.OnPremisesCredentialScheme -eq 'Auto') {
            try { $offered = @((Test-GctAnonymousEndpoint -Context $Context -Url $EwsUrl -Method 'POST').Schemes); $Context.OnPremisesSchemes[$uri.Host] = $offered } catch { $offered = @() }
        }
        $scheme = Select-GctCredentialScheme -Requested ([string]$cfg.OnPremisesCredentialScheme) -Offered $offered
        if (-not $scheme) { throw "EWS on-premises offers no scheme for a user name and a password (offered: $(if ($offered.Count) { $offered -join ', ' } else { 'none seen' })): set -OnPremisesCredentialScheme NTLM, Negotiate or Basic (Identity.OnPremisesCredentialScheme), or use OAuth." }
        $notOffered = if ($offered.Count -and @($offered) -inotcontains $scheme) { "; EWS does not offer $scheme (offered: $($offered -join ', '))" } else { '' }
        if ($scheme -eq 'Basic') {
            if ($uri.Scheme -ne 'https') { throw "Basic over $($uri.Scheme)://: the password would travel in clear text. Use the https:// URL of EWS." }
            # Sent at once (no challenge first); the trace shows the user name only.
            $basic = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$($credential.UserName):$($network.Password)"))
            $how = "Basic, account $($credential.UserName) (password typed, sent in each request)$notOffered"
        }
        else {
            $cache = [Net.CredentialCache]::new()
            $cache.Add([Uri]$uri.GetLeftPart([UriPartial]::Authority), $scheme, $network)
            $handler = [Net.Http.HttpClientHandler]::new()
            $handler.Credentials = $cache
            $handler.UseDefaultCredentials = $false
            $handler.AllowAutoRedirect = $false
            $client = [Net.Http.HttpClient]::new($handler)
            $client.Timeout = [TimeSpan]::FromSeconds([int]$cfg.HttpTimeoutSeconds)
            $windowsScheme = "$scheme $account"
            $label = "$scheme, account $account"
            $how = "$scheme, account $account (password typed$(if ($scheme -eq 'Negotiate') { ', Kerberos else NTLM' }))$notOffered"
        }
    }
    elseif ($mode -eq 'Windows') {
        $handler = [Net.Http.HttpClientHandler]::new()
        $handler.UseDefaultCredentials = $true
        $handler.AllowAutoRedirect = $false
        $client = [Net.Http.HttpClient]::new($handler)
        $client.Timeout = [TimeSpan]::FromSeconds([int]$cfg.HttpTimeoutSeconds)
        $label = 'Windows session'
        $how = "Windows (current account $([Environment]::UserDomainName)\$([Environment]::UserName), Negotiate)"
    }
    else {
        $token = Get-GctOnPremisesToken -Context $Context -HostName $uri.Host
        $how = "OAuth, hybrid modern authentication (token for https://$($uri.Host), $($token.Flow))"
    }
    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Post, $EwsUrl)
    try {
        $null = $request.Headers.TryAddWithoutValidation('User-Agent', [string]$cfg.UserAgent)
        $null = $request.Headers.TryAddWithoutValidation('X-AnchorMailbox', [string]$cfg.SourceMailbox)
        $null = $request.Headers.TryAddWithoutValidation('client-request-id', [guid]::NewGuid().ToString())
        $null = $request.Headers.TryAddWithoutValidation('return-client-request-id', 'true')
        if ($basic) { $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Basic', $basic) }
        elseif ($mode -eq 'OAuth') { $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $token.AccessToken) }
        # The account the handler answers the NTLM / Negotiate challenge with (read by the simulator of the tests).
        if ($windowsScheme) { ([Collections.Generic.IDictionary[string, object]]$request.Options).Add('GctWindowsCredential', $windowsScheme) }
        $request.Content = [Net.Http.StringContent]::new($Envelope, [Text.Encoding]::UTF8, 'text/xml')
        $response = Invoke-GctHttp -HttpClient $client -Request $request -Operation 'EWS GetUserAvailability' -Label $label
    }
    finally {
        $request.Dispose()
        if ($handler) { $client.Dispose(); $handler.Dispose() }
        $basic = $null
    }
    $xml = $null; $ns = $null; $fault = $null
    $text = if ($response.Body) { [Text.Encoding]::UTF8.GetString([byte[]]$response.Body).TrimStart([char]0xFEFF) } else { '' }
    if ($text.TrimStart().StartsWith('<')) {
        try {
            $xml = [Xml.XmlDocument]::new(); $xml.LoadXml($text)
            $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
            foreach ($k in $script:EwsNs.Keys) { $ns.AddNamespace($k, $script:EwsNs[$k]) }
            $f = $xml.SelectSingleNode('//soap:Fault', $ns)
            if ($f) {
                $code = $f.SelectSingleNode('detail/*[local-name()="ResponseCode"]')
                $fault = [pscustomobject]@{ Code = if ($code) { $code.InnerText } else { [string]$f.SelectSingleNode('faultcode').InnerText }; Text = [string]$f.SelectSingleNode('faultstring').InnerText }
            }
        }
        catch { $xml = $null }
    }
    $version = if ($xml) { $sv = $xml.SelectSingleNode('//t:ServerVersionInfo', $ns); if ($sv) { '{0}.{1}.{2}.{3}' -f $sv.GetAttribute('MajorVersion'), $sv.GetAttribute('MinorVersion'), $sv.GetAttribute('MajorBuildNumber'), $sv.GetAttribute('MinorBuildNumber') } }
    [pscustomobject]@{
        HttpStatus = [int]$response.StatusCode; Response = $response; Xml = $xml; Ns = $ns; Fault = $fault; Authentication = $how; Mode = $mode; Offered = @($offered)
        Challenges = @($response.Challenges); Server = Get-GctResponseServer -Response $response; Version = $version
        ErrorHeader = Get-GctHeader $response 'X-MS-Diagnostics'
    }
}

function Get-GctZoneOffset {
    <# UTC offset in minutes of a time zone name of Exchange or Windows at an instant ($null when unknown). #>
    param([AllowEmptyString()][string]$Name, [datetime]$AtUtc = [datetime]::UtcNow)
    if (-not $Name) { return $null }
    try { return [int]([TimeZoneInfo]::FindSystemTimeZoneById($Name).GetUtcOffset([datetime]::SpecifyKind($AtUtc, 'Utc')).TotalMinutes) } catch { return $null }
}

function ConvertFrom-GctEwsAvailability {
    <# One free/busy object per mailbox of a GetUserAvailability answer, like ConvertFrom-GctSchedule. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Answer, [Parameter(Mandatory = $true)][string[]]$Mailboxes, [hashtable]$Paths = @{}, [int]$IntervalMinutes = 30, [datetime]$StartUtc)

    $ns = $Answer.Ns
    $responses = @($Answer.Xml.SelectNodes('//m:FreeBusyResponseArray/m:FreeBusyResponse', $ns))
    for ($i = 0; $i -lt $Mailboxes.Count; $i++) {
        $mailbox = $Mailboxes[$i]
        $r = if ($i -lt $responses.Count) { $responses[$i] } else { $null }
        $message = if ($r) { $r.SelectSingleNode('m:ResponseMessage', $ns) } else { $null }
        $view = if ($r) { $r.SelectSingleNode('m:FreeBusyView', $ns) } else { $null }
        $class = if ($message) { $message.GetAttribute('ResponseClass') } else { 'Error' }
        $code = if ($message) { Get-GctNodeText $message 'm:ResponseCode' $ns } else { 'NoResponse' }
        $text = if ($message) { Get-GctNodeText $message 'm:MessageText' $ns } else { 'No FreeBusyResponse for this mailbox.' }
        $events = @(if ($view) {
                foreach ($e in $view.SelectNodes('t:CalendarEventArray/t:CalendarEvent', $ns)) {
                    $s = Get-GctNodeText $e 't:StartTime' $ns; $en = Get-GctNodeText $e 't:EndTime' $ns
                    $su = if ($s) { [datetime]::SpecifyKind([datetime]$s, 'Utc') } else { $null }; $eu = if ($en) { [datetime]::SpecifyKind([datetime]$en, 'Utc') } else { $null }
                    [pscustomobject]@{
                        Mailbox = $mailbox; StartUtc = if ($su) { $su.ToString('yyyy-MM-ddTHH:mm:ssZ') }; EndUtc = if ($eu) { $eu.ToString('yyyy-MM-ddTHH:mm:ssZ') }
                        Start = if ($su) { $su.ToLocalTime().ToString('yyyy-MM-dd HH:mm') }; End = if ($eu) { $eu.ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
                        BusyType = Get-GctNodeText $e 't:BusyType' $ns
                        Subject = Get-GctNodeText $e 't:CalendarEventDetails/t:Subject' $ns; Location = Get-GctNodeText $e 't:CalendarEventDetails/t:Location' $ns
                        IsPrivate = (Get-GctNodeText $e 't:CalendarEventDetails/t:IsPrivate' $ns) -eq 'true'; IsMeeting = (Get-GctNodeText $e 't:CalendarEventDetails/t:IsMeeting' $ns) -eq 'true'
                    }
                }
            })
        $period = if ($view) { $view.SelectSingleNode('t:WorkingHours/t:WorkingPeriodArray/t:WorkingPeriod', $ns) } else { $null }
        $days = @(if ($period) { ([string](Get-GctNodeText $period 't:DayOfWeek' $ns)).ToLowerInvariant().Split(' ', [StringSplitOptions]::RemoveEmptyEntries) })
        $from = if ($period) { [int](Get-GctNodeText $period 't:StartTimeInMinutes' $ns) } else { $null }
        $to = if ($period) { [int](Get-GctNodeText $period 't:EndTimeInMinutes' $ns) } else { $null }
        $bias = if ($view) { Get-GctNodeText $view 't:WorkingHours/t:TimeZone/t:Bias' $ns } else { $null }
        $merged = if ($view) { [string](Get-GctNodeText $view 't:MergedFreeBusy' $ns) } else { '' }
        $path = [string]$Paths[$mailbox.ToLowerInvariant()]
        $failed = $class -ne 'Success'
        [pscustomobject]@{
            Mailbox         = $mailbox
            Protocol        = 'EWS GetUserAvailability (on-premises)'
            Path            = $path
            Status          = if ($failed) { 'Failed' } elseif (-not $merged) { 'Warning' } else { 'Passed' }
            ResponseCode    = $code
            Message         = $text
            Hint            = if ($failed) { Get-GctAvailabilityHint -Code $code -Message $text -Path $path } else { $null }
            View            = $merged
            IntervalMinutes = $IntervalMinutes
            StartUtc        = $StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
            DetailLevel     = if ($failed) { $null } else { Get-GctDetailLevel -Events $events -View $merged }
            WorkingHours    = if ($period) { '{0} {1:00}:{2:00}-{3:00}:{4:00} (bias {5} min)' -f ($days -join ' '), [Math]::Floor($from / 60), ($from % 60), [Math]::Floor($to / 60), ($to % 60), $bias } else { $null }
            WorkDays        = $days
            WorkStart       = $from
            WorkEnd         = $to
            WorkTimeZone    = if ($null -ne $bias) { "bias $bias" } else { $null }
            WorkOffsetMinutes = if ($null -ne $bias -and "$bias" -ne '') { - [int]$bias } else { $null }
            Events          = $events
        }
    }
}
