<#
.SYNOPSIS
    Graph Connectivity Tests - the hybrid and sharing configuration of Exchange Online, read by an administrator (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Optional (Test.AdminChecks, -AdminChecks): what Exchange Online is configured with to find the free/busy
    of each domain of the targets. A user cannot see it; an Exchange administrator
    can, read-only (View-Only Organization Management is enough). No module to install: the cmdlets
    Get-IntraOrganizationConnector, Get-OrganizationRelationship and Get-AvailabilityAddressSpace are called
    through the administration API Connect-ExchangeOnline uses (InvokeCommand of
    https://outlook.office365.com/adminapi/beta/<tenant>), with its client (Identity.AdminClientId).

    That API is the one of the ExchangeOnlineManagement module: it is not documented as a public API and
    can change. A failure here never fails the test: the step says how to read the same with
    Connect-ExchangeOnline.

    The configuration of Exchange on-premises (its IntraOrganizationConnector to Exchange Online, Get-AuthServer,
    Test-OAuthConnectivity) is out of reach from the Internet: the steps name the commands to run there.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

function Invoke-GctExoCommand {
    <# One read-only Exchange Online cmdlet through the administration API. Returns Answer (Graph-like) and Rows. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Name, [hashtable]$Parameters = @{})

    $token = $Context.Tokens.Admin
    $tenant = if ($token.Claims -and $token.Claims['tid']) { [string]$token.Claims['tid'] } else { Get-GctAuthority -Context $Context }
    $url = "https://$($Context.Cloud.Outlook)/adminapi/beta/$tenant/InvokeCommand"
    $headers = @{ 'X-ResponseFormat' = 'json'; 'X-ClientApplication' = 'ExoManagementModule'; 'X-CmdletName' = $Name; Prefer = 'odata.maxpagesize=1000' }
    $upn = if ($token.Claims) { [string]$token.Claims['upn'] } else { $null }
    if ($upn) { $headers['X-AnchorMailbox'] = "UPN:$upn" }
    elseif ([string]$Context.Config.TenantId -match '\.') { $headers['X-AnchorMailbox'] = "APP:SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}@$($Context.Config.TenantId)" }
    $body = @{ CmdletInput = @{ CmdletName = $Name; Parameters = $Parameters } }
    $answer = Invoke-GctGraph -Context $Context -Method POST -Path $url -Body $body -Headers $headers -Token $token.AccessToken -Operation "InvokeCommand $Name"
    # @() around the if: a single row must stay an array (not a hashtable counted by its keys), no row an empty array.
    $rows = @(if ((Test-GctGraphOk $answer) -and $answer.Json) { $answer.Json['value'] | Where-Object { $_ } })
    [pscustomobject]@{ Answer = $answer; Rows = $rows }
}

function Get-GctList {
    <# A multi-valued property of a cmdlet answer as strings (array, or text with commas). #>
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return @() }
    $items = if ($Value -is [string]) { $Value -split '[,;]\s*' } else { @($Value) }
    return @($items | ForEach-Object { if ($_ -is [System.Collections.IDictionary]) { [string]($_['Domain'] ?? $_['Address'] ?? $_['Name'] ?? ($_.Values | Select-Object -First 1)) } else { [string]$_ } } | Where-Object { $_ } | ForEach-Object { $_.Trim() })
}

function Test-GctDomainMatch {
    <# The domain is in the list (exact, or a wildcard *.contoso.com). #>
    param([string[]]$List, [Parameter(Mandatory = $true)][string]$Domain)
    foreach ($d in @($List)) { if ($d -ieq $Domain -or ($d.StartsWith('*.') -and $Domain.EndsWith($d.Substring(1), [StringComparison]::OrdinalIgnoreCase))) { return $true } }
    return $false
}

function Invoke-GctStageHybrid {
    <# The configuration Exchange Online uses for each domain of the targets, and what to check on-premises. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Hybrid'
    $cfg = $Context.Config
    $manual = 'Connect-ExchangeOnline, then Get-IntraOrganizationConnector | fl; Get-OrganizationRelationship | fl; Get-AvailabilityAddressSpace | fl'
    try { $token = Get-GctAdminToken -Context $Context }
    catch {
        Add-GctStep $Context $stage 'Administrator sign-in' Warning "No administrator token: $($_.Exception.Message) Read the configuration yourself: $manual." ([ordered]@{ Client = $cfg.AdminClientId })
        return
    }
    $summary = Get-GctTokenSummary -Claims $token.Claims
    Add-GctStep $Context $stage 'Administrator sign-in' Passed "Token of $($summary.User) for the Exchange Online administration API ($($token.Flow))." $summary

    $data = @{}
    foreach ($cmd in 'Get-IntraOrganizationConnector', 'Get-OrganizationRelationship', 'Get-AvailabilityAddressSpace') {
        $r = Invoke-GctExoCommand -Context $Context -Name $cmd
        if (-not (Test-GctGraphOk $r.Answer)) {
            $why = switch ($r.Answer.HttpStatus) { 401 { 'the token was refused: the account is not an Exchange administrator, or the API did not accept the client' } 403 { 'the account has no role that allows this cmdlet (View-Only Organization Management)' } default { Get-GctGraphFailureText -Context $Context -Answer $r.Answer } }
            Add-GctStep $Context $stage $cmd Warning "Not read: $why. Read it yourself: $manual." (Get-GctGraphDetails -Answer $r.Answer)
            return
        }
        $data[$cmd] = $r.Rows
        $names = @($r.Rows | ForEach-Object { [string]$_['Name'] }) -join ', '
        Add-GctStep $Context $stage $cmd Passed "$($r.Rows.Count) object(s)$(if ($names) { ": $names" })." (Get-GctGraphDetails -Answer $r.Answer -More ([ordered]@{ Objects = $names }))
    }
    $Context.Hybrid = [ordered]@{
        IntraOrganizationConnectors = @($data['Get-IntraOrganizationConnector'] | ForEach-Object { [ordered]@{ Name = $_['Name']; Enabled = $_['Enabled']; TargetAddressDomains = (Get-GctList $_['TargetAddressDomains']) -join ', '; DiscoveryEndpoint = [string]$_['DiscoveryEndpoint']; TargetSharingEpr = [string]$_['TargetSharingEpr'] } })
        OrganizationRelationships   = @($data['Get-OrganizationRelationship'] | ForEach-Object { [ordered]@{ Name = $_['Name']; Enabled = $_['Enabled']; DomainNames = (Get-GctList $_['DomainNames']) -join ', '; FreeBusyAccessEnabled = $_['FreeBusyAccessEnabled']; FreeBusyAccessLevel = [string]$_['FreeBusyAccessLevel']; TargetAutodiscoverEpr = [string]$_['TargetAutodiscoverEpr']; TargetSharingEpr = [string]$_['TargetSharingEpr']; TargetApplicationUri = [string]$_['TargetApplicationUri']; MailTipsAccessEnabled = $_['MailTipsAccessEnabled'] } })
        AvailabilityAddressSpaces   = @($data['Get-AvailabilityAddressSpace'] | ForEach-Object { [ordered]@{ ForestName = [string]$_['ForestName']; AccessMethod = [string]$_['AccessMethod']; ProxyUrl = [string]$_['ProxyUrl']; TargetAutodiscoverEpr = [string]$_['TargetAutodiscoverEpr'] } })
    }

    # The domains Exchange Online must find outside itself: on-premises targets and targets of other organisations.
    $domains = @($cfg.TargetMailboxes | ForEach-Object { $Context.Mailboxes[$_.ToLowerInvariant()] } | Where-Object { $_ -and $_.Relation -in 'CloudToOnPremises', 'CrossOrganisation', 'Unknown' } | ForEach-Object { $_.Domain } | Select-Object -Unique)
    if ($Context.SourceProtocol -eq 'EWS') {
        Add-GctStep $Context $stage 'On-premises configuration' Passed "The source is on-premises: the request to Exchange Online is configured on-premises. There, run: Get-IntraOrganizationConnector | fl TargetAddressDomains,DiscoveryEndpoint,Enabled; Get-AuthServer | fl Name,Type,Enabled,IsDefaultAuthorizationEndpoint; Test-OAuthConnectivity -Service EWS -TargetUri https://$($Context.Cloud.Outlook)/ews/exchange.asmx -Mailbox $($cfg.SourceMailbox) -Verbose | fl."
    }
    if (-not $domains.Count) {
        Add-GctStep $Context $stage 'Domains' Passed 'Every target is a mailbox of Exchange Online in the tenant: no hybrid or sharing configuration is involved.'
        return
    }
    $onPremHosts = @(Get-GctOnPremisesHosts -Context $Context)
    foreach ($domain in $domains) {
        $ioc = @($data['Get-IntraOrganizationConnector'] | Where-Object { Test-GctDomainMatch -List (Get-GctList $_['TargetAddressDomains']) -Domain $domain })
        $rel = @($data['Get-OrganizationRelationship'] | Where-Object { Test-GctDomainMatch -List (Get-GctList $_['DomainNames']) -Domain $domain })
        $aas = @($data['Get-AvailabilityAddressSpace'] | Where-Object { [string]$_['ForestName'] -ieq $domain })
        $d = [ordered]@{ Domain = $domain; IntraOrganizationConnector = @($ioc | ForEach-Object { $_['Name'] }) -join ', '; OrganizationRelationship = @($rel | ForEach-Object { $_['Name'] }) -join ', '; AvailabilityAddressSpace = @($aas | ForEach-Object { $_['AccessMethod'] }) -join ', ' }
        $iocOn = @($ioc | Where-Object { $_['Enabled'] -ne $false })
        $relOn = @($rel | Where-Object { $_['Enabled'] -ne $false })
        if ($iocOn.Count) {
            $c = $iocOn[0]
            $endpoint = [string]$c['DiscoveryEndpoint']
            $d.DiscoveryEndpoint = $endpoint
            $endpointHost = try { ([Uri]$endpoint).Host } catch { $null }
            $note = if ($endpointHost -and $onPremHosts.Count -and $endpointHost -notin $onPremHosts) { " Its DiscoveryEndpoint host ($endpointHost) is not the host checked ($($onPremHosts -join ', ')): run the test with -OnPremisesHost $endpointHost." } else { '' }
            $status = if ($note) { 'Warning' } else { 'Passed' }
            Add-GctStep $Context $stage "Configuration for $domain" $status "OAuth: the IntraOrganizationConnector $($c['Name']) sends the requests for $domain to $endpoint (Autodiscover, then EWS, with an OAuth token of the tenant). It wins over any organization relationship.$note" $d
            continue
        }
        if ($ioc.Count -and -not $iocOn.Count) { $d.Disabled = "IntraOrganizationConnector $($ioc[0]['Name']) disabled" }
        if ($relOn.Count) {
            $r = $relOn[0]
            $d.FreeBusyAccessEnabled = $r['FreeBusyAccessEnabled']; $d.FreeBusyAccessLevel = $r['FreeBusyAccessLevel']; $d.TargetAutodiscoverEpr = $r['TargetAutodiscoverEpr']; $d.TargetSharingEpr = $r['TargetSharingEpr']
            if ($r['FreeBusyAccessEnabled'] -eq $false) {
                Add-GctStep $Context $stage "Configuration for $domain" Failed "The organization relationship $($r['Name']) covers $domain but FreeBusyAccessEnabled is False: Exchange Online does not ask it for free/busy (Set-OrganizationRelationship -FreeBusyAccessEnabled `$true -FreeBusyAccessLevel LimitedDetails)." $d
            }
            else {
                $ep = if ([string]$r['TargetSharingEpr']) { [string]$r['TargetSharingEpr'] } else { [string]$r['TargetAutodiscoverEpr'] }
                Add-GctStep $Context $stage "Configuration for $domain" Passed "Organization relationship $($r['Name']) (DAuth, federation): free/busy asked at $ep, level $($r['FreeBusyAccessLevel']) - the other organisation must have a relationship for this organisation too.$(if ($d['Disabled']) { " ($($d['Disabled']).)" })" $d
            }
            continue
        }
        if ($aas.Count) {
            Add-GctStep $Context $stage "Configuration for $domain" Passed "Availability address space for $domain ($($aas[0]['AccessMethod'])$(if ($aas[0]['ProxyUrl']) { ", $($aas[0]['ProxyUrl'])" }))." $d
            continue
        }
        Add-GctStep $Context $stage "Configuration for $domain" Failed "Exchange Online has no IntraOrganizationConnector, organization relationship or availability address space for $domain$(if ($d['Disabled']) { " ($($d['Disabled']))" }): it cannot ask anybody for this free/busy (ErrorAvailabilityConfigNotFound). In hybrid run the Hybrid Configuration Wizard; across organisations create an organization relationship (New-OrganizationRelationship -DomainNames $domain -FreeBusyAccessEnabled `$true)." $d
    }
}
