<#
.SYNOPSIS
    Graph Connectivity Tests - where the mailboxes are, and the on-premises endpoint (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Before any sign-in:

      Discovery   for the source and each target mailbox: the Entra ID tenant of its domain (OpenID
                  configuration), how the domain signs in (managed, federated), and where the mailbox
                  is - Autodiscover v2 of Exchange Online, anonymous, then the Autodiscover of the
                  domain. From that, the path of the free/busy request: same organisation, Microsoft 365
                  to on-premises, on-premises to Microsoft 365, another organisation.
      OnPremises  the endpoint Exchange Online calls for an on-premises mailbox (the Availability
                  service, Target.OnPremisesHost or the host Autodiscover gives): DNS, TLS certificate,
                  and what EWS and Autodiscover answer without credentials - Exchange Online signs in
                  there with OAuth (IntraOrganizationConnector): the Bearer challenge must be offered
                  and must trust the tenant (trusted_issuers); a reverse proxy with pre-authentication
                  stops it.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>

#region DNS -----------------------------------------------------------------------------------------

# The DNS servers of the run (Test.DnsServer), set for each run (Run.ps1); empty: the DNS servers of Windows.
$script:GctDnsServer = @()

function Invoke-GctDnsQuery {
    <#
        One DNS query with the resolver of Windows. The only function of the tool that sends DNS queries; the tests
        replace it. Returns the records (Type, Name, Data) - an empty list for NXDOMAIN or no record - or throws
        when no DNS server answers.
    #>
    param([Parameter(Mandatory = $true)][string]$Name, [Parameter(Mandatory = $true)][ValidateSet('A', 'AAAA', 'CNAME', 'MX', 'TXT', 'SRV', 'NS')][string]$Type)

    if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) { throw 'Resolve-DnsName (module DnsClient of Windows) is not available here: the DNS checks need Windows.' }
    # Test.DnsServer (-DnsServer): the DNS servers asked instead of those of Windows.
    $server = @{}
    if (@($script:GctDnsServer).Count) { $server.Server = [string[]]@($script:GctDnsServer) }
    try { $answers = @(Resolve-DnsName -Name $Name -Type $Type -DnsOnly -QuickTimeout -ErrorAction Stop @server) }
    catch {
        if ($_.Exception.Message -match 'does not exist|DNS name does not exist|No records|9003|9501|NXDOMAIN|n.existe pas') { return @() }
        throw
    }
    $records = foreach ($a in $answers) {
        $section = [string](Get-GctField $a 'Section')
        if ($section -and $section -ne 'Answer') { continue }
        $t = [string](Get-GctField $a 'Type')
        $data = switch ($t) {
            'A' { [string](Get-GctField $a 'IPAddress') }
            'AAAA' { [string](Get-GctField $a 'IPAddress') }
            'CNAME' { [string](Get-GctField $a 'NameHost') }
            'MX' { '{0} {1}' -f (Get-GctField $a 'Preference'), (Get-GctField $a 'NameExchange') }
            'TXT' { (@(Get-GctField $a 'Strings') -join '') }
            'SRV' { '{0} {1} {2} {3}' -f (Get-GctField $a 'Priority'), (Get-GctField $a 'Weight'), (Get-GctField $a 'Port'), (Get-GctField $a 'NameTarget') }
            'NS' { [string](Get-GctField $a 'NameHost') }
            default { $null }
        }
        if ($null -ne $data) { [pscustomobject]@{ Type = $t; Name = [string](Get-GctField $a 'Name'); Data = $data } }
    }
    return @($records)
}

function Resolve-GctDns {
    <# A DNS query, added to the trace of the run (method DNS). Returns Records (of the type asked), Chain (CNAMEs) and Error. #>
    param([Parameter(Mandatory = $true)][string]$Name, [Parameter(Mandatory = $true)][string]$Type)

    $clock = [Diagnostics.Stopwatch]::StartNew()
    $err = $null
    $all = @()
    try { $all = @(Invoke-GctDnsQuery -Name $Name -Type $Type) } catch { $err = $_.Exception.Message }
    $records = @($all | Where-Object { $_.Type -eq $Type })
    $chain = @($all | Where-Object { $_.Type -eq 'CNAME' -and $Type -ne 'CNAME' })
    $text = if ($err) { "No answer: $err" } elseif (-not $all.Count) { 'No record (NXDOMAIN or empty answer).' } else { (@($all | ForEach-Object { "$($_.Name)`t$($_.Type)`t$($_.Data)" }) -join "`n") }
    Add-GctTraceEntry -Method 'DNS' -Url "dns:$Name/$Type" -Request "$Type $Name ($(Get-GctDnsServerText))" -Response $text -StatusCode $(if ($err) { $null } else { 0 }) -Reason $(if ($err) { '' } elseif ($records.Count) { "$($records.Count) record(s)" } else { 'no record' }) -DurationMs $clock.ElapsedMilliseconds -Operation "DNS $Type"
    [pscustomobject]@{ Name = $Name; Type = $Type; Records = $records; Chain = $chain; Error = $err }
}

function Get-GctDnsServerText {
    <# The DNS servers asked, for the trace and the messages: Test.DnsServer, else those of Windows. #>
    if (@($script:GctDnsServer).Count) { return "DNS server $(@($script:GctDnsServer) -join ', ')" }
    return 'resolver of Windows'
}

function Test-GctDnsAnswer {
    <#
        Whether a DNS server answers at all: one query of Test.DnsProbeName (a name that always exists). An empty
        answer or NXDOMAIN is still an answer; no answer (timeout, no server, no network) is not. Returns Answers,
        Server (text), Probe and Error.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    $probe = if ([string]$Context.Config.DnsProbeName) { [string]$Context.Config.DnsProbeName } else { 'microsoft.com' }
    $q = Resolve-GctDns -Name $probe -Type 'A'
    [pscustomobject]@{ Answers = -not $q.Error; Server = Get-GctDnsServerText; Probe = $probe; Error = $q.Error }
}

#endregion

#region Where is the mailbox -----------------------------------------------------------------------

function Test-GctExchangeOnlineHost {
    param([AllowEmptyString()][string]$HostName)
    return ([string]$HostName).ToLowerInvariant() -in @('outlook.office365.com', 'outlook.office.com', 'outlook.office365.us', 'outlook-dod.office365.us', 'partner.outlook.cn', 'outlook.live.com')
}

function Get-GctAutodiscoverV2 {
    <#
        Autodiscover v2 (JSON), anonymous, like Outlook: first Exchange Online
        (it answers for its mailboxes and sends the others to their organisation), then the Autodiscover of the
        domain. Redirects are followed (5 at most) and traced. Returns EwsUrl, Source, Hops and Error.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Address)

    $domain = Get-GctAddressDomain $Address
    $escaped = [Uri]::EscapeDataString($Address)
    $starts = @(
        @{ Url = "https://$($Context.Cloud.Outlook)/autodiscover/autodiscover.json/v1.0/$($escaped)?Protocol=EWS"; Source = 'Exchange Online Autodiscover v2' }
        @{ Url = "https://autodiscover.$domain/autodiscover/autodiscover.json/v1.0/$($escaped)?Protocol=EWS"; Source = "Autodiscover v2 of $domain" }
    )
    $errors = [Collections.Generic.List[string]]::new()
    foreach ($start in $starts) {
        $url = $start.Url
        $hops = [Collections.Generic.List[string]]::new()
        for ($i = 0; $i -lt 6; $i++) {
            $hops.Add($url)
            try { $r = Invoke-GctWebRequest -HttpClient $Context.HttpClient -Uri $url -UserAgent ([string]$Context.Config.UserAgent) -Headers @{ Accept = 'application/json' } }
            catch { $errors.Add("$($start.Source): $($_.Exception.Message)"); break }
            if ($r.StatusCode -in 301, 302, 307, 308 -and $r.Location) {
                $next = $null
                if (-not [Uri]::TryCreate([Uri]$url, $r.Location, [ref]$next) -or $next.Scheme -ne 'https') { $errors.Add("$($start.Source): redirect to $($r.Location) refused (not HTTPS)"); break }
                $url = $next.AbsoluteUri
                continue
            }
            if ($r.StatusCode -eq 200) {
                $json = $null
                try { $json = $r.Content | ConvertFrom-Json -ErrorAction Stop } catch { $json = $null }
                $ews = [string](Get-GctField $json 'Url')
                if ($ews -match '^https://') { return [pscustomobject]@{ EwsUrl = $ews; Source = $start.Source; Hops = @($hops); Error = $null } }
                $errors.Add("$($start.Source): HTTP 200 without an EWS URL")
                break
            }
            $errors.Add("$($start.Source): HTTP $($r.StatusCode)")
            break
        }
    }
    [pscustomobject]@{ EwsUrl = $null; Source = $null; Hops = @(); Error = ($errors -join '; ') }
}

function Get-GctMailboxLocation {
    <#
        Where one mailbox is, without sign-in: tenant of its domain, how the domain signs in, EWS URL from Autodiscover.
        Location: ExchangeOnline, OnPremises or Unknown.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Address, [Parameter(Mandatory = $true)][string]$Role)

    $domain = Get-GctAddressDomain $Address
    if (-not $Context.Tenants.ContainsKey($domain)) {
        $Context.Tenants[$domain] = try { Resolve-GctEntraTenant -Context $Context -Name $domain } catch { [pscustomobject]@{ TenantId = $null; Error = $_.Exception.Message; Region = $null; SubRegion = $null; CloudInstance = $null; Issuer = $null; Url = $null } }
    }
    $tenant = $Context.Tenants[$domain]
    $realm = try { Get-GctUserRealm -Context $Context -User $Address } catch { $null }
    # A consumer domain (Outlook.com, Gmail...) answers with the tenant of the Microsoft accounts: not an organisation.
    $consumer = [bool](($realm -and $realm.ConsumerDomain) -or $tenant.TenantId -eq $script:ConsumerTenantId)
    $tenantId = if ($consumer) { $null } else { $tenant.TenantId }
    $auto = Get-GctAutodiscoverV2 -Context $Context -Address $Address
    $ewsHost = if ($auto.EwsUrl) { ([Uri]$auto.EwsUrl).Host } else { $null }
    # Exchange Online answers Autodiscover v2 for any address (even unknown): without a tenant it says nothing;
    # with one, only the directory and the free/busy confirm the mailbox. A redirect to another host is on-premises.
    $location = if (-not $auto.EwsUrl) { 'Unknown' } elseif (-not (Test-GctExchangeOnlineHost $ewsHost)) { 'OnPremises' } elseif (-not $tenantId) { 'External' } else { 'ExchangeOnline' }
    [pscustomobject]@{
        Address       = $Address
        Role          = $Role
        Domain        = $domain
        TenantId      = $tenantId
        Consumer      = $consumer
        TenantRegion  = (@($tenant.Region, $tenant.SubRegion) | Where-Object { $_ }) -join ' / '
        TenantError   = $tenant.Error
        Realm         = if ($realm) { $realm.NameSpaceType } else { $null }
        FederationUrl = if ($realm) { $realm.AuthUrl } else { $null }
        Location      = $location
        EwsUrl        = $auto.EwsUrl
        EwsHost       = $ewsHost
        Autodiscover  = if ($auto.EwsUrl) { $auto.Source } else { $auto.Error }
        Hops          = @($auto.Hops)
        Relation      = $null
    }
}

function Get-GctLocationText {
    param([Parameter(Mandatory = $true)][pscustomobject]$Mailbox)
    switch ($Mailbox.Location) {
        'ExchangeOnline' { return 'Exchange Online' }
        'OnPremises' { return "Exchange on-premises ($($Mailbox.EwsHost))" }
        'External' { if ($Mailbox.Consumer) { return 'a consumer address (not a Microsoft 365 organisation)' } return 'outside Microsoft 365 (no tenant for the domain)' }
        default { return 'not found by Autodiscover' }
    }
}

function Get-GctFreeBusyPath {
    <#
        The way the free/busy request goes from the source to one target, and what it needs. Path:
        SameOrganisation, CloudToOnPremises, OnPremisesToCloud, OnPremises, CrossOrganisation, Unknown.
    #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Source, [Parameter(Mandatory = $true)][pscustomobject]$Target)

    $sameTenant = $Source.TenantId -and $Target.TenantId -and $Source.TenantId -ieq $Target.TenantId
    $s = $Source.Location; $t = $Target.Location
    if (-not $sameTenant -and $Target.TenantId -and $Source.TenantId) {
        return [pscustomobject]@{ Path = 'CrossOrganisation'; Text = "another organisation (tenant $($Target.TenantId)): the source organisation needs an organization relationship (or an OAuth IntraOrganizationConnector) for $($Target.Domain), and the target organisation must share its free/busy with it" }
    }
    if ($Target.Consumer) {
        return [pscustomobject]@{ Path = 'CrossOrganisation'; Text = "a consumer address ($($Target.Domain)): no organisation shares free/busy for it, expect ErrorMailRecipientNotFound or ErrorAvailabilityConfigNotFound" }
    }
    if (-not $Target.TenantId -and $t -ne 'Unknown' -and -not $sameTenant) {
        return [pscustomobject]@{ Path = 'CrossOrganisation'; Text = "an organisation without Microsoft 365 ($($Target.Domain)): the source organisation needs an availability address space or an organization relationship for $($Target.Domain)" }
    }
    if ($s -eq 'ExchangeOnline' -and $t -eq 'ExchangeOnline') { return [pscustomobject]@{ Path = 'SameOrganisation'; Text = 'both mailboxes in Exchange Online, same tenant: Exchange Online answers alone, the calendar permission of the target decides the detail' } }
    if ($s -eq 'ExchangeOnline' -and $t -eq 'OnPremises') { return [pscustomobject]@{ Path = 'CloudToOnPremises'; Text = "Microsoft 365 to on-premises (hybrid): Exchange Online calls $($Target.EwsHost) from the Internet - OAuth through the IntraOrganizationConnector, else the organization relationship (DAuth)" } }
    if ($s -eq 'OnPremises' -and $t -eq 'ExchangeOnline') { return [pscustomobject]@{ Path = 'OnPremisesToCloud'; Text = "on-premises to Microsoft 365 (hybrid): Exchange on-premises ($($Source.EwsHost)) calls Exchange Online - OAuth through its IntraOrganizationConnector, else its organization relationship" } }
    if ($s -eq 'OnPremises' -and $t -eq 'OnPremises') { return [pscustomobject]@{ Path = 'OnPremises'; Text = 'both mailboxes on-premises: Exchange on-premises answers alone (no Microsoft 365 in the path)' } }
    return [pscustomobject]@{ Path = 'Unknown'; Text = 'Autodiscover did not say where one of the mailboxes is: the free/busy request shows what Exchange Online knows of it' }
}

function Invoke-GctStageDiscovery {
    <# Source and targets: tenant, realm, Autodiscover, then the path of each free/busy request. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Discovery'
    $cfg = $Context.Config
    $addresses = [Collections.Generic.List[object]]::new()
    $addresses.Add(@{ Address = [string]$cfg.SourceMailbox; Role = 'Source' })
    foreach ($t in @($cfg.TargetMailboxes)) { if ($t -ine [string]$cfg.SourceMailbox) { $addresses.Add(@{ Address = [string]$t; Role = 'Target' }) } }
    foreach ($a in $addresses) {
        $m = Get-GctMailboxLocation -Context $Context -Address $a.Address -Role $a.Role
        $Context.Mailboxes[$a.Address.ToLowerInvariant()] = $m
        if ($a.Role -eq 'Source') { $Context.Source = $m; if ($m.TenantId -and -not $cfg.TenantId) { $Context.SourceTenantId = $m.TenantId } }
        $details = [ordered]@{ Role = $m.Role; Domain = $m.Domain; TenantId = $m.TenantId; TenantRegion = $m.TenantRegion; SignIn = $m.Realm; FederationUrl = $m.FederationUrl; Location = $m.Location; EwsUrl = $m.EwsUrl; FoundBy = $m.Autodiscover; Redirects = ($m.Hops -join ' > ') }
        $where = Get-GctLocationText -Mailbox $m
        $tenantText = if ($m.TenantId) { "tenant $($m.TenantId)$(if ($m.Realm -eq 'Federated') { ', federated sign-in' })" } elseif ($m.Consumer) { 'consumer domain (Microsoft accounts)' } else { 'no Microsoft 365 tenant for this domain' }
        if ($a.Role -eq 'Source' -and $m.Location -eq 'External') {
            Add-GctStep $Context $stage "$($a.Role) $($a.Address)" Failed "$($tenantText): the source must be a mailbox of a Microsoft 365 organisation or of Exchange on-premises." $details
            $Context.Stop = $true
            return
        }
        if ($m.Location -eq 'Unknown') {
            $status = if ($a.Role -eq 'Source' -and -not $m.TenantId) { 'Failed' } else { 'Warning' }
            Add-GctStep $Context $stage "$($a.Role) $($a.Address)" $status "$tenantText; Autodiscover did not find the mailbox ($($m.Autodiscover))." $details
            if ($status -eq 'Failed') { $Context.Stop = $true; return }
        }
        elseif ($m.Location -eq 'External') { Add-GctStep $Context $stage "$($a.Role) $($a.Address)" Warning "$(if ($m.Consumer) { 'A consumer address (Microsoft accounts domain)' } else { 'No Microsoft 365 tenant for this domain' }): only an availability address space or an organization relationship of the source organisation can reach its free/busy." $details }
        else { Add-GctStep $Context $stage "$($a.Role) $($a.Address)" Passed "$where, $tenantText." $details }
    }
    if (-not @($cfg.TargetMailboxes).Count) { return }
    foreach ($t in @($cfg.TargetMailboxes)) {
        $m = $Context.Mailboxes[([string]$t).ToLowerInvariant()]
        if (-not $m -or $m.Role -eq 'Source') { continue }
        $path = Get-GctFreeBusyPath -Source $Context.Source -Target $m
        $m.Relation = $path.Path
        Add-GctStep $Context $stage "Path to $t" Passed "Path of the request: $($path.Text)." ([ordered]@{ Path = $path.Path; Source = Get-GctLocationText $Context.Source; Target = Get-GctLocationText $m })
    }
}

#endregion

#region On-premises endpoint -----------------------------------------------------------------------

function Get-GctOnPremisesHosts {
    <# Hosts to check: Target.OnPremisesHost, else the EWS host Autodiscover gave for an on-premises mailbox. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    if ([string]$Context.Config.OnPremisesHost) { return @([string]$Context.Config.OnPremisesHost) }
    return @($Context.Mailboxes.Values | Where-Object { $_.Location -eq 'OnPremises' -and $_.EwsHost } | ForEach-Object { $_.EwsHost } | Select-Object -Unique)
}

function Test-GctAnonymousEndpoint {
    <#
        What an Exchange endpoint answers without credentials: HTTP status, the WWW-Authenticate schemes, the Bearer
        challenge (trusted_issuers: the tenants on-premises Exchange accepts tokens from), and whether Exchange itself
        answered (X-FEServer, request-id) or a reverse proxy (sign-in page).
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Url, [string]$Method = 'POST')

    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::new($Method), $Url)
    try {
        $null = $request.Headers.TryAddWithoutValidation('User-Agent', [string]$Context.Config.UserAgent)
        # An empty Bearer header, as Exchange Online and Outlook send it: Exchange offers its OAuth challenge
        # (trusted_issuers, authorization_uri) only to a request that announces Bearer.
        $null = $request.Headers.TryAddWithoutValidation('Authorization', 'Bearer')
        if ($Method -eq 'POST') { $request.Content = [Net.Http.StringContent]::new('', [Text.Encoding]::UTF8, 'text/xml') }
        $response = Invoke-GctHttp -HttpClient $Context.HttpClient -Request $request -Label 'no credentials'
    }
    finally { $request.Dispose() }
    $challenge = Get-GctChallengeInfo -Challenges @($response.Challenges)
    $headers = $response.Headers
    $exchange = $false
    foreach ($k in $headers.Keys) { if ($k -in 'X-FEServer', 'X-OWA-Version', 'request-id', 'X-CalculatedBETarget') { $exchange = $true } }
    $body = if ($response.Body) { [Text.Encoding]::UTF8.GetString([byte[]]$response.Body) } else { '' }
    [pscustomobject]@{
        Url            = $Url
        HttpStatus     = [int]$response.StatusCode
        Schemes        = @($challenge.Schemes)
        Bearer         = $challenge.Bearer
        TrustedIssuers = @($challenge.TrustedIssuers)
        AuthorizationUri = $challenge.AuthorizationUri
        Exchange       = $exchange
        Server         = Get-GctResponseServer -Response $response
        Location       = if ($headers.ContainsKey('Location')) { [string]$headers['Location'] } else { $null }
        SignInPage     = $body -match '<form|login|logon|sign in' -and $response.StatusCode -eq 200
    }
}

function Invoke-GctStageOnPremises {
    <# The on-premises endpoint Exchange Online calls for free/busy: DNS, certificate, EWS and Autodiscover without credentials. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'OnPremises'
    $hosts = @(Get-GctOnPremisesHosts -Context $Context)
    if (-not $hosts.Count) {
        Add-GctStep $Context $stage 'On-premises endpoint' Skipped 'No on-premises mailbox and no Target.OnPremisesHost: nothing to check on-premises.'
        return
    }
    foreach ($h in $hosts) {
        # The tenant that must be trusted: the one of the on-premises mailboxes behind this host (hybrid), else the source.
        $tenants = @($Context.Mailboxes.Values | Where-Object { $_.EwsHost -eq $h -and $_.TenantId } | ForEach-Object { $_.TenantId } | Select-Object -Unique)
        if (-not $tenants.Count -and $Context.SourceTenantId) { $tenants = @($Context.SourceTenantId) }
        $dns = Resolve-GctDns -Name $h -Type 'A'
        $ips = @($dns.Records | ForEach-Object { $_.Data })
        if ($dns.Error -or -not $ips.Count) {
            Add-GctStep $Context $stage "DNS $h" Failed "$h does not resolve$(if ($dns.Error) { " ($($dns.Error))" }): Exchange Online cannot reach it either. Publish the name in the public DNS." ([ordered]@{ Host = $h; Error = $dns.Error })
            continue
        }
        Add-GctStep $Context $stage "DNS $h" Passed "$h resolves to $($ips -join ', ')$(if ($dns.Chain.Count) { " (through $(@($dns.Chain | ForEach-Object { $_.Data }) -join ' > '))" }). Exchange Online resolves it in the public DNS: the address must be the published one." ([ordered]@{ Host = $h; Addresses = $ips -join ', '; Cname = @($dns.Chain | ForEach-Object { $_.Data }) -join ' > ' })

        $tls = Get-GctTlsCertificate -HostName $h -TimeoutSeconds ([Math]::Min(15, [int]$Context.Config.HttpTimeoutSeconds))
        $d = [ordered]@{ Host = $h; Subject = $tls.Subject; Issuer = $tls.Issuer; NotAfterUtc = $tls.NotAfterUtc; DaysLeft = $tls.DaysLeft; Protocol = $tls.Protocol; Error = $tls.Error }
        if (-not $tls.Reachable) {
            Add-GctStep $Context $stage "HTTPS $h" Failed "No TCP connection to $($h):443 from here ($($tls.Error)). Exchange Online must reach it from the Internet (firewall, reverse proxy, Microsoft 365 IP addresses)." $d
            continue
        }
        if (-not $tls.Valid) {
            $why = if ($tls.Interrupted) { 'the connection was closed before the certificate (firewall, TLS inspection, TLS 1.2 not enabled)' } else { 'the certificate is not trusted' }
            Add-GctStep $Context $stage "Certificate $h" Failed "TLS failed: $why - $($tls.Error). Exchange Online needs a certificate of a public authority with the name $h." $d
            continue
        }
        $warn = [int]$Context.Config.CertificateWarningDays
        $status = if ($tls.DaysLeft -lt $warn) { 'Warning' } else { 'Passed' }
        Add-GctStep $Context $stage "Certificate $h" $status "Trusted certificate ($($tls.Subject), $($tls.Protocol)), valid until $($tls.NotAfterUtc) ($($tls.DaysLeft) days$(if ($status -eq 'Warning') { ": renew it, Exchange Online stops at the expiry" }))." $d

        foreach ($endpoint in @(
                @{ Name = 'EWS'; Url = "https://$h/EWS/Exchange.asmx"; Method = 'POST'; Role = 'the Availability service (GetUserAvailability)' }
                @{ Name = 'Autodiscover'; Url = "https://$h/autodiscover/autodiscover.svc"; Method = 'POST'; Role = 'the DiscoveryEndpoint of the IntraOrganizationConnector' })) {
            $name = "$($endpoint.Name) $h"
            try { $r = Test-GctAnonymousEndpoint -Context $Context -Url $endpoint.Url -Method $endpoint.Method }
            catch { Add-GctStep $Context $stage $name Failed "No answer from $($endpoint.Url): $($_.Exception.Message)" ([ordered]@{ Url = $endpoint.Url }); continue }
            if ($endpoint.Name -eq 'EWS' -and $r.HttpStatus -eq 401 -and $Context.ContainsKey('OnPremisesSchemes')) { $Context.OnPremisesSchemes[$h] = @($r.Schemes) }
            $d = [ordered]@{ Url = $r.Url; HttpStatus = $r.HttpStatus; Schemes = $r.Schemes -join ', '; Bearer = $r.Bearer; TrustedIssuers = $r.TrustedIssuers -join ', '; AuthorizationUri = $r.AuthorizationUri; Exchange = $r.Exchange; Servers = $r.Server; Location = $r.Location }
            if ($r.HttpStatus -in 301, 302, 303, 307 -or $r.SignInPage) {
                Add-GctStep $Context $stage $name Failed "$($endpoint.Url) answers HTTP $($r.HttpStatus)$(if ($r.Location) { " to $($r.Location)" }) instead of an Exchange challenge: a reverse proxy with pre-authentication (or a redirect) is in front of $($endpoint.Role). Exchange Online cannot pre-authenticate: publish /EWS and /Autodiscover in pass-through." $d
                continue
            }
            if ($r.HttpStatus -ne 401) {
                $st = if ($r.HttpStatus -eq 404) { 'Failed' } else { 'Warning' }
                Add-GctStep $Context $stage $name $st "$($endpoint.Url) answers HTTP $($r.HttpStatus) without credentials (expected: 401 with the schemes Exchange accepts)$(if ($r.HttpStatus -eq 404) { ': the virtual directory is not published under this name' })." $d
                continue
            }
            if (-not $r.Bearer) {
                Add-GctStep $Context $stage $name Warning "$($endpoint.Role) does not offer OAuth (schemes: $($r.Schemes -join ', ')): Exchange Online cannot use the IntraOrganizationConnector here and falls back to the organization relationship (DAuth) if there is one. Check OAuthAuthentication on the virtual directory, the Hybrid Configuration Wizard, and that the reverse proxy keeps the WWW-Authenticate: Bearer header." $d
                continue
            }
            $issuerTenants = @($r.TrustedIssuers | ForEach-Object { ($_ -split '@', 2)[-1] } | Where-Object { $_ -match '^[0-9a-fA-F-]{36}$' })
            $missing = @($tenants | Where-Object { $_ -notin $issuerTenants })
            if ($issuerTenants.Count -and $missing.Count) {
                Add-GctStep $Context $stage $name Warning "OAuth offered, but the tenant $($missing -join ', ') is not among the issuers on-premises Exchange trusts ($($issuerTenants -join ', ')): its tokens will be refused (401). Run the Hybrid Configuration Wizard (or New-AuthServer for the tenant)." $d
                continue
            }
            $trust = if ($issuerTenants.Count) { " and trusts the tenant $($issuerTenants -join ', ')" } else { '' }
            Add-GctStep $Context $stage $name Passed "Exchange answers 401 with OAuth (Bearer)${trust}: $($endpoint.Role) is ready for the tokens of Exchange Online. Schemes: $($r.Schemes -join ', ')." $d
        }
    }
}

#endregion

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAETt1677PfFCNr
# lxNSty+yP+JBPJLvone/EXvckGjVpqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA1SWdcny59
# igz4Y43AOsmveV6Hdbe77g4oZI+aRgu2bTANBgkqhkiG9w0BAQEFAASCAgCGl+QD
# h/O1BnpS86BFs6rWvWUzgvHMCMDequdGyMDV2mm3PMc4GW5G4XdI8E/iyxSKs441
# lNmnd6CiivJ3EoZPK0C11UE2M7a9xS8cg+CVUndcteFLcj0qo6MEW/hQAbBJacA0
# xQko+oHKcisuIv3tNz3DKSJcNPxVrWIOC666Vyoh7wzYSM2RMsrHXJlswg2W41dW
# fqxAv04Z7EfUFDlgl63tycJU+4Ky7DrPkBL5wCGCKgyNNTeBxGEmeapKM7GLJPpZ
# V8sk6O3v0Y02QRXWbF1aAV03YESU1f3t3Z2Z0fqTABcKjVV1U9wlZzBykc2taEsO
# whmAmFM8Zmw9ETrKekdg5vIlQPwGsbB96aiWRuFjXKTf4vdFlDSxLz/qNMdvKpos
# 7rG/LHfOcIO+2NADOJCSVpb5AMaFiDtagZ0l1uFD7iiXrqOyr7Zculfl31R05av6
# rA0Stq43pYMEYXep9Gm5aCeGFUS6joCaL3b+pYL+R556h/yCHNwm+OHPYVyXftzu
# EfeIh2GBHE/s8rdzYfLBR1bo1CZQFuKWoyb5bmdn1InOc1EfqXnTFQoQoXsvJ0Nn
# mOky8O+SKrkIhEvle/mlS7cEr32ldZSZup0ETw/o8H/cEivLKvmcOxB91TLx5yMX
# jytGLcdBmpUhkTVR7xZRiw4VLVxq6u0fw77HvKGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIBlY3dg1h0fC7VJiLETNcdMRlQdcPhVgpH48V5yFHCLwAggWItnOeup1
# XBgPMjAyNjExMTAxODI2MDlaMAMCAQECEHNV6zHqOdqNREpo/NTaKeagghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDE4MjYwOVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIGtF
# 64sJj0u9Yn5bSxGx5X3225zBWaISXeyxnfPDlsOaMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiB9d6qgVQ/SOmiRwzCpy2yrxUrTILzI2PvRy5jY1XU2aAIhAOi2
# c0yG6ByKGUojPIbsVL9xhpah/EgGUSS7PAnUr3ZF
# SIG # End signature block
