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

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBQXgbtWBOQCIoa
# XeEn8JN9VWfExr7R+sTuvf8fGNtwZqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCNHqaRsWIL
# Ok4LhVOkRAhBiwm4I378SImgOpf9/Nwu7jANBgkqhkiG9w0BAQEFAASCAgA77nto
# MNfPF5ZqAOr49wPrk+tIzlrbTey4BjIhyZliOgqwrcSB8ACg7jJ6EPrnU/wqC9U6
# DFbCSLEJVwJXO4qaug8iQHeBDnPt2WWX4KfOnYtNiv2E/C6bPtvvYOFwdYr/uJBS
# FmP9JayKgcTsCrmw4LowlGxNHCVUJeM+dhusYUALuTDuEx4W1yXaOiSdUH6zBqEi
# CnOrEjoOtVhmkxE7/i50T7+vCUucwFjj4+LlD/IlIiMf3gtpU2PgXQtAw2XHQdVW
# KBu/PCsS3Bl7ILnySGyuz7N9szFOsq0fQNKkYzJv/UGyi5nedZHAR+PLBvMaCFn/
# MIDpeQyGzMliwlThShq4DEmXHi6V2FmrBvvpf6KGrXOaMzWNMQ8aO5hNNKDuWUSh
# W7lgQDbMM0eBHHkGP6DHk7n16Q2K0lt0s9e9tJYV4vmLkejiWZyjlZZ/tfnNTzQn
# LzKM+c8IMdKCqXr483JPSxhhvG4/6axKD2FAt7EBT0Rkk3+Yzo3YFSRWLNaICQYq
# rjc01XjO+BNMdcHQ7IZpZq6ybiiqOBdDIehtGeq178qOOg00G1ddFPMOMG+rUjN4
# JMQtc13QutHsjXSWDdn7PPs7XVSdV/vEfg8jurxlKg1geqeb7EO7wrFcEVJ2GMFg
# 9XQ/9CMRXoTXpA3xsVuWM+OdNYFXsrf6E18bMqGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIGKW0VIa81RApCOwixFvG3Ye+cPkL0mRHP8inpdxvR6gAgg2+q3Fbdl7
# cxgPMjAyNjExMTAxMDIzMzFaMAMCAQECEDueiwlGKije10QHjYsbBE2gghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjMzMVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIGno
# 2M6xrlxneDyUq+/FCd2hES+f5fc31IwX97S6cAReMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEA0+qqz2LBLkZHhSbvIcqVh1MTHq+DB2Qv6ahimaYME0oCIHmW
# bE75Ok+gZa2DaA1sCZFqZM1EHo0aTME//NjijVUX
# SIG # End signature block
