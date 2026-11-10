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
    Version : 1.0.0
#>

#region DNS -----------------------------------------------------------------------------------------

function Invoke-GctDnsQuery {
    <#
        One DNS query with the resolver of Windows. The only function of the tool that sends DNS queries; the tests
        replace it. Returns the records (Type, Name, Data) - an empty list for NXDOMAIN or no record - or throws
        when no DNS server answers.
    #>
    param([Parameter(Mandatory = $true)][string]$Name, [Parameter(Mandatory = $true)][ValidateSet('A', 'AAAA', 'CNAME', 'MX', 'TXT', 'SRV', 'NS')][string]$Type)

    if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) { throw 'Resolve-DnsName (module DnsClient of Windows) is not available here: the DNS checks need Windows.' }
    try { $answers = @(Resolve-DnsName -Name $Name -Type $Type -DnsOnly -QuickTimeout -ErrorAction Stop) }
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
    Add-GctTraceEntry -Method 'DNS' -Url "dns:$Name/$Type" -Request "$Type $Name (resolver of Windows)" -Response $text -StatusCode $(if ($err) { $null } else { 0 }) -Reason $(if ($err) { '' } elseif ($records.Count) { "$($records.Count) record(s)" } else { 'no record' }) -DurationMs $clock.ElapsedMilliseconds -Operation "DNS $Type"
    [pscustomobject]@{ Name = $Name; Type = $Type; Records = $records; Chain = $chain; Error = $err }
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
