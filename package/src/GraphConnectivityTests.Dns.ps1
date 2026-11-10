<#
.SYNOPSIS
    Graph Connectivity Tests - DNS records of a domain for Exchange Online (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Dns            without sign-in, with the resolver of Windows: MX, SPF, Autodiscover, DKIM (selector1, selector2), DMARC,
                   MTA-STS (record and policy file), TLS-RPT, and the device registration records
    DomainRecords  signed in: the domain in the tenant (verified, services) and the records Microsoft 365
                   expects for it (GET /domains/{domain}/serviceConfigurationRecords), each compared with
                   what the DNS publishes

    The domain is Target.Domain (-Domain), else the domain of the source mailbox.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

function Get-GctTestDomain {
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    if ([string]$Context.Config.Domain) { return ([string]$Context.Config.Domain).ToLowerInvariant() }
    return Get-GctAddressDomain ([string]$Context.Config.SourceMailbox)
}

function Get-GctSpfInclusion {
    <#
        Whether an SPF record includes Exchange Online, directly or through its includes (followed 3 levels deep,
        traced), and how many DNS lookups it costs (include, a, mx, ptr, exists, redirect: 10 at most, RFC 7208).
    #>
    param([Parameter(Mandatory = $true)][string]$Record, [int]$Depth = 0, [hashtable]$Seen = @{})

    $eop = 'spf\.protection\.outlook\.com|spf\.protection\.office365\.us|spf\.protection\.partner\.outlook\.cn'
    $lookups = ([regex]::Matches($Record, '(?:^|\s)[+~?-]?(include:|a\b|a:|mx\b|mx:|ptr|exists:|redirect=)')).Count
    if ($Record -match "include:($eop)") { return [pscustomobject]@{ Found = $true; Via = $null; Lookups = $lookups } }
    $found = $false; $via = $null
    foreach ($m in [regex]::Matches($Record, '(?:include:|redirect=)([^\s]+)')) {
        $name = $m.Groups[1].Value
        if ($Depth -ge 3 -or $Seen.ContainsKey($name)) { continue }
        $Seen[$name] = $true
        $q = Resolve-GctDns -Name $name -Type 'TXT'
        $inner = @($q.Records | Where-Object { $_.Data -match '^v=spf1' } | ForEach-Object { $_.Data }) | Select-Object -First 1
        if (-not $inner) { continue }
        $r = Get-GctSpfInclusion -Record $inner -Depth ($Depth + 1) -Seen $Seen
        $lookups += $r.Lookups
        if ($r.Found -and -not $found) { $found = $true; $via = if ($r.Via) { "$name > $($r.Via)" } else { $name } }
    }
    [pscustomobject]@{ Found = $found; Via = $via; Lookups = $lookups }
}

function Add-GctDnsRow {
    param([Parameter(Mandatory = $true)][hashtable]$Context, [string]$Record, [string]$Name, [string]$Type, [string]$Found, [string]$Expected, [string]$Status, [string]$Source = 'DNS')
    $Context.DnsRecords.Add([pscustomobject]@{ Record = $Record; Name = $Name; Type = $Type; Found = $Found; Expected = $Expected; Status = $Status; Source = $Source })
}

function Invoke-GctStageDns {
    <# The public DNS records of the domain for Exchange Online, read with the resolver of Windows. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Dns'
    $domain = Get-GctTestDomain -Context $Context
    $protection = $Context.Cloud.Protection
    $Context.Domain = $domain

    # MX
    $mx = Resolve-GctDns -Name $domain -Type 'MX'
    $hosts = @($mx.Records | ForEach-Object { $_.Data } | Sort-Object { [int](($_ -split ' ')[0]) })
    $d = [ordered]@{ Name = $domain; Records = $hosts -join ' | ' }
    if ($mx.Error) { Add-GctStep $Context $stage 'MX' Failed "No DNS answer for $domain ($($mx.Error))." $d; Add-GctDnsRow $Context 'MX' $domain 'MX' '' "*.$protection" 'Failed'; return }
    if (-not $hosts.Count) { Add-GctStep $Context $stage 'MX' Failed "$domain has no MX record: no mail can be delivered to it." $d; Add-GctDnsRow $Context 'MX' $domain 'MX' '' "*.$protection" 'Failed' }
    else {
        $toEop = @($hosts | Where-Object { $_ -match [regex]::Escape($protection) + '\.?$' })
        $first = ($hosts[0] -split ' ', 2)[1]
        if ($toEop.Count -and $hosts[0] -match [regex]::Escape($protection)) { Add-GctStep $Context $stage 'MX' Passed "Mail goes to Exchange Online Protection: $first." $d; $st = 'Passed' }
        elseif ($toEop.Count) { Add-GctStep $Context $stage 'MX' Warning "An MX points to Exchange Online Protection, but not the one of lowest preference ($first): mail goes there first." $d; $st = 'Warning' }
        else { Add-GctStep $Context $stage 'MX' Passed "Mail goes to $first (not to Exchange Online Protection: a third-party gateway or on-premises; normal in a hybrid with centralized mail transport or a filtering service)." $d; $st = 'Passed' }
        Add-GctDnsRow $Context 'MX' $domain 'MX' ($hosts -join ' | ') "*.$protection (lowest preference)" $st
    }

    # SPF
    $txt = Resolve-GctDns -Name $domain -Type 'TXT'
    $spf = @($txt.Records | Where-Object { $_.Data -match '^v=spf1' } | ForEach-Object { $_.Data })
    $d = [ordered]@{ Name = $domain; Spf = $spf -join ' | ' }
    if (-not $spf.Count) { Add-GctStep $Context $stage 'SPF' Warning "$domain has no SPF record: receivers cannot tell which servers may send for it. Publish v=spf1 include:spf.protection.outlook.com -all (with your other senders)." $d; $st = 'Warning' }
    elseif ($spf.Count -gt 1) { Add-GctStep $Context $stage 'SPF' Failed "$($spf.Count) SPF records: SPF fails with PermError when there is more than one. Merge them into one." $d; $st = 'Failed' }
    else {
        $spfInfo = Get-GctSpfInclusion -Record $spf[0]
        $d.Lookups = $spfInfo.Lookups
        if ($spfInfo.Via) { $d.IncludedThrough = $spfInfo.Via }
        if (-not $spfInfo.Found) { Add-GctStep $Context $stage 'SPF' Warning "The SPF record does not include Exchange Online, directly or through its includes ($($spf[0])): mail sent from Exchange Online fails SPF." $d; $st = 'Warning' }
        else {
            $st = if ($spfInfo.Lookups -gt 10) { 'Warning' } else { 'Passed' }
            $via = if ($spfInfo.Via) { " through $($spfInfo.Via)" } else { '' }
            Add-GctStep $Context $stage 'SPF' $st "SPF includes Exchange Online$via, ending with $(([regex]::Match($spf[0], '[~?+-]all')).Value); $($spfInfo.Lookups) DNS lookup(s) counted$(if ($spfInfo.Lookups -gt 10) { ': more than 10, receivers answer PermError' })." $d
        }
    }
    Add-GctDnsRow $Context 'SPF' $domain 'TXT' ($spf -join ' | ') 'v=spf1 include:spf.protection.outlook.com ...' $st

    # Autodiscover
    $ad = Resolve-GctDns -Name "autodiscover.$domain" -Type 'CNAME'
    $target = @($ad.Records | ForEach-Object { $_.Data.TrimEnd('.') }) | Select-Object -First 1
    $d = [ordered]@{ Name = "autodiscover.$domain"; Cname = $target }
    if ($target -match 'autodiscover\.outlook\.com$') { Add-GctStep $Context $stage 'Autodiscover' Passed "autodiscover.$domain points to $target (Exchange Online)." $d; $st = 'Passed' }
    elseif ($target) { Add-GctStep $Context $stage 'Autodiscover' Passed "autodiscover.$domain points to ${target}: Exchange on-premises answers Autodiscover (normal in hybrid while mailboxes are on-premises)." $d; $st = 'Passed' }
    else {
        $a = Resolve-GctDns -Name "autodiscover.$domain" -Type 'A'
        if ($a.Records.Count) { Add-GctStep $Context $stage 'Autodiscover' Passed "autodiscover.$domain is an A record ($(@($a.Records | ForEach-Object { $_.Data }) -join ', ')): an on-premises Autodiscover." $d; $st = 'Passed'; $target = @($a.Records | ForEach-Object { $_.Data }) -join ', ' }
        else { Add-GctStep $Context $stage 'Autodiscover' Warning "No autodiscover.$domain record: Outlook finds Exchange Online only through other ways (it may fail). Publish CNAME autodiscover.$domain -> autodiscover.outlook.com." $d; $st = 'Warning' }
    }
    Add-GctDnsRow $Context 'Autodiscover' "autodiscover.$domain" 'CNAME' $target 'autodiscover.outlook.com' $st

    # DKIM
    $found = 0
    foreach ($sel in 'selector1', 'selector2') {
        $k = Resolve-GctDns -Name "$sel._domainkey.$domain" -Type 'CNAME'
        $v = @($k.Records | ForEach-Object { $_.Data.TrimEnd('.') }) | Select-Object -First 1
        if (-not $v) { $t = Resolve-GctDns -Name "$sel._domainkey.$domain" -Type 'TXT'; $v = @($t.Records | ForEach-Object { $_.Data }) | Select-Object -First 1 }
        if ($v) { $found++ }
        Add-GctDnsRow $Context "DKIM $sel" "$sel._domainkey.$domain" 'CNAME' $v "$sel-<domain>._domainkey.<tenant>.onmicrosoft.com" $(if ($v) { 'Passed' } else { 'Warning' })
    }
    $d = [ordered]@{ Selectors = "$found of 2" }
    if ($found -eq 2) { Add-GctStep $Context $stage 'DKIM' Passed 'selector1 and selector2 are published: DKIM signing can be enabled (Get-DkimSigningConfig in Exchange Online tells whether it is).' $d }
    else { Add-GctStep $Context $stage 'DKIM' Warning "$found of the 2 DKIM selectors of Exchange Online are published: DKIM signing of $domain cannot be enabled (New-DkimSigningConfig, then the CNAME records it gives)." $d }

    # DMARC
    $dm = Resolve-GctDns -Name "_dmarc.$domain" -Type 'TXT'
    $dmarc = @($dm.Records | Where-Object { $_.Data -match '^v=DMARC1' } | ForEach-Object { $_.Data }) | Select-Object -First 1
    # "$dmarc", not [string]$dmarc: a cast $null reaches the .NET method as null (no DMARC record) and Match throws.
    $policy = ([regex]::Match("$dmarc", 'p=(\w+)')).Groups[1].Value
    $d = [ordered]@{ Name = "_dmarc.$domain"; Record = $dmarc; Policy = $policy }
    if (-not $dmarc) { Add-GctStep $Context $stage 'DMARC' Warning "No DMARC record: publish _dmarc.$domain TXT v=DMARC1; p=none; rua=mailto:... then move to quarantine or reject." $d; $st = 'Warning' }
    elseif ($policy -eq 'none') { Add-GctStep $Context $stage 'DMARC' Passed "DMARC published with p=none (monitoring only)." $d; $st = 'Passed' }
    else { Add-GctStep $Context $stage 'DMARC' Passed "DMARC published with p=$policy." $d; $st = 'Passed' }
    Add-GctDnsRow $Context 'DMARC' "_dmarc.$domain" 'TXT' $dmarc 'v=DMARC1; p=...' $st

    # MTA-STS and TLS-RPT (optional)
    $sts = Resolve-GctDns -Name "_mta-sts.$domain" -Type 'TXT'
    $stsRecord = @($sts.Records | Where-Object { $_.Data -match '^v=STSv1' } | ForEach-Object { $_.Data }) | Select-Object -First 1
    $rpt = Resolve-GctDns -Name "_smtp._tls.$domain" -Type 'TXT'
    $rptRecord = @($rpt.Records | Where-Object { $_.Data -match '^v=TLSRPTv1' } | ForEach-Object { $_.Data }) | Select-Object -First 1
    $d = [ordered]@{ MtaSts = $stsRecord; TlsRpt = $rptRecord }
    if ($stsRecord) {
        $policyUrl = "https://mta-sts.$domain/.well-known/mta-sts.txt"
        $file = try { Invoke-GctWebRequest -HttpClient $Context.HttpClient -Uri $policyUrl -UserAgent ([string]$Context.Config.UserAgent) } catch { $null }
        $mode = if ($file -and $file.StatusCode -eq 200) { ([regex]::Match($file.Content, 'mode:\s*(\w+)')).Groups[1].Value } else { $null }
        $d.PolicyUrl = $policyUrl; $d.Mode = $mode
        $mxList = if ($file -and $file.StatusCode -eq 200) { @([regex]::Matches($file.Content, 'mx:\s*(\S+)') | ForEach-Object { $_.Groups[1].Value }) } else { @() }
        $d.PolicyMx = $mxList -join ', '
        if (-not $mode) { Add-GctStep $Context $stage 'MTA-STS' Failed "_mta-sts.$domain is published but the policy $policyUrl cannot be read$(if ($file) { " (HTTP $($file.StatusCode))" }): senders that enforce MTA-STS may refuse to deliver." $d; $st = 'Failed' }
        else { Add-GctStep $Context $stage 'MTA-STS' Passed "MTA-STS policy in mode $mode (MX allowed: $($mxList -join ', '))$(if ($rptRecord) { '; TLS-RPT published' } else { '; no TLS-RPT record (reports of TLS failures)' })." $d; $st = 'Passed' }
    }
    else { Add-GctStep $Context $stage 'MTA-STS' Passed "No MTA-STS record (optional: it makes senders require TLS to the MX)$(if ($rptRecord) { '; TLS-RPT published' })." $d; $st = 'Passed' }
    Add-GctDnsRow $Context 'MTA-STS' "_mta-sts.$domain" 'TXT' $stsRecord 'v=STSv1; id=... (optional)' $st
    Add-GctDnsRow $Context 'TLS-RPT' "_smtp._tls.$domain" 'TXT' $rptRecord 'v=TLSRPTv1; rua=... (optional)' $(if ($rptRecord) { 'Passed' } else { 'Skipped' })

    # Device registration and enrollment (Entra ID, Intune)
    $reg = @(foreach ($n in 'enterpriseregistration', 'enterpriseenrollment') {
            $c = Resolve-GctDns -Name "$n.$domain" -Type 'CNAME'
            $v = @($c.Records | ForEach-Object { $_.Data.TrimEnd('.') }) | Select-Object -First 1
            Add-GctDnsRow $Context $n "$n.$domain" 'CNAME' $v $(if ($n -eq 'enterpriseregistration') { 'enterpriseregistration.windows.net' } else { 'enterpriseenrollment-s.manage.microsoft.com' }) $(if ($v) { 'Passed' } else { 'Skipped' })
            if ($v) { "$n -> $v" }
        })
    Add-GctStep $Context $stage 'Device registration' Passed $(if ($reg.Count) { "Published: $($reg -join '; ')." } else { 'enterpriseregistration and enterpriseenrollment are not published (optional: device registration and Intune enrollment with the domain of the users).' }) ([ordered]@{ Records = $reg -join '; ' })
}

function Invoke-GctStageDomainRecords {
    <# The domain in the tenant and the records Microsoft 365 expects for it, compared with the DNS. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'DomainRecords'
    $domain = Get-GctTestDomain -Context $Context
    $Context.Domain = $domain
    $answer = Invoke-GctGraph -Context $Context -Path "/domains/$([Uri]::EscapeDataString($domain))" -Operation 'GET domains/{domain}'
    if (-not (Test-GctGraphOk $answer)) {
        $status = if ($answer.HttpStatus -eq 404) { 'Failed' } else { 'Warning' }
        $text = if ($answer.HttpStatus -eq 404) { "$domain is not a domain of this tenant." } else { "The domain was not read: $(Get-GctGraphFailureText -Context $Context -Answer $answer -Permission 'Domain.Read.All')" }
        Add-GctStep $Context $stage "Domain $domain" $status $text (Get-GctGraphDetails -Answer $answer)
        return
    }
    $dom = $answer.Json
    $services = @($dom['supportedServices']) -join ', '
    $d = Get-GctGraphDetails -Answer $answer -More ([ordered]@{ Verified = $dom['isVerified']; Default = $dom['isDefault']; Authentication = $dom['authenticationType']; Services = $services })
    if (-not $dom['isVerified']) { Add-GctStep $Context $stage "Domain $domain" Failed "$domain is in the tenant but not verified: Exchange Online does not accept mail for it. Publish the verification TXT record and verify it." $d; return }
    Add-GctStep $Context $stage "Domain $domain" Passed "Verified, $($dom['authenticationType']) sign-in, services: $(if ($services) { $services } else { 'none' })$(if ($services -notmatch 'Email') { ' - Email is not among the services of the domain' })." $d

    $records = Invoke-GctGraph -Context $Context -Path "/domains/$([Uri]::EscapeDataString($domain))/serviceConfigurationRecords" -Operation 'GET domains/{domain}/serviceConfigurationRecords'
    if (-not (Test-GctGraphOk $records)) { Add-GctStep $Context $stage 'Expected records' Warning "Not read: $(Get-GctGraphFailureText -Context $Context -Answer $records -Permission 'Domain.Read.All')" (Get-GctGraphDetails -Answer $records); return }
    $ok = 0; $bad = [Collections.Generic.List[string]]::new()
    foreach ($r in @($records.Json['value'])) {
        $type = [string]$r['recordType']; $label = ([string]$r['label']).TrimEnd('.')
        $service = [string]$r['supportedService']
        $expected = switch ($type) { 'Mx' { [string]$r['mailExchange'] } 'CName' { [string]$r['canonicalName'] } 'Txt' { [string]$r['text'] } 'Srv' { '{0} {1} {2} {3}' -f $r['priority'], $r['weight'], $r['port'], $r['nameTarget'] } default { '' } }
        $dnsType = switch ($type) { 'Mx' { 'MX' } 'CName' { 'CNAME' } 'Txt' { 'TXT' } 'Srv' { 'SRV' } default { 'TXT' } }
        $q = Resolve-GctDns -Name $label -Type $dnsType
        $values = @($q.Records | ForEach-Object { if ($dnsType -eq 'MX') { (($_.Data -split ' ', 2)[1]).TrimEnd('.') } else { $_.Data.TrimEnd('.') } })
        $match = switch ($dnsType) {
            'TXT' { @($values | Where-Object { $_ -eq $expected -or ($expected -match '^v=spf1' -and $_ -match 'include:spf\.protection') }).Count -gt 0 }
            default { @($values | Where-Object { $_ -ieq $expected.TrimEnd('.') }).Count -gt 0 }
        }
        $optional = [bool]$r['isOptional']
        $st = if ($match) { 'Passed' } elseif ($optional) { 'Skipped' } else { 'Warning' }
        Add-GctDnsRow $Context "$service $type" $label $dnsType ($values -join ' | ') $expected $st 'Microsoft 365'
        if ($match) { $ok++ } elseif (-not $optional) { $bad.Add("$service $dnsType $label") }
    }
    $total = @($records.Json['value']).Count
    $d = Get-GctGraphDetails -Answer $records -More ([ordered]@{ Expected = $total; Matching = $ok; Missing = $bad -join '; ' })
    if ($bad.Count) { Add-GctStep $Context $stage 'Expected records' Warning "$ok of $total records Microsoft 365 expects are published as expected; not as expected: $($bad -join '; ') (normal for the MX and Autodiscover while mail or mailboxes stay on-premises)." $d }
    else { Add-GctStep $Context $stage 'Expected records' Passed "The $ok record(s) Microsoft 365 expects for $domain are published as expected (optional ones aside)." $d }
}
