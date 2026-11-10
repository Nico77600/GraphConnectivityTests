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

# SIG # Begin signature block
# MIImKwYJKoZIhvcNAQcCoIImHDCCJhgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCLYeC2ydZKOUyi
# 1KlGA/PsJQO2GxOOG6CTYco5qhkNs6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYEjCCGA4CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCbvbnkk33u
# MgqgipsTz9RF56l1wyA2u6FqjxHyd5YKPzANBgkqhkiG9w0BAQEFAASCAgBd17cb
# 256s+VBiY33qyvkL62HNjHe/C75WnVhU+mcu7ifMhkzaRjpEN2yFMtkR+Fk/iosI
# bzufU92NgHKcO7LkB3/MKpAkCpnn6dT8cwT3K6ssIrVdGhsX8ClswlrAZ8UOjgBl
# nDrkEBmaGKZXRFsTjJLzFL2VM2VRfRw7NOAgLcuQDN1AiAhs0bo9Eg1ORShXp535
# rC/+WqdiWlocMjC1YEy9jtAQE/Lh1bQ6zSTxJOUSCTT1w/CF1Ffq632cQqxk2SnM
# ke82YFsDZPND6MzkzqvhSV/mnY0UzDMGOgBxROLNGVMy8mT7X6gqqoPRbHWmKHU5
# 7qaFCugLSgqHQDPO3nHHl8SKcU1gSfPcMTyu9j6W9Z8+rQO9ddQU6jZgT7cfHH2t
# 8c0G9CJUjTv4CcboOjA9mFNMnfMhRJyNRxUDD3Jwhi8xEUUIUi0sgAsc/bKJc6V9
# /MAglKMgvCEcJZ3B88YoV3TLKrNNCDPTiolZ8Ts7/MJrpnLprIziZRO+5GyJb5z5
# p5hDLJnR3liP+1FqZyTnfRUAPsVluKD0u+LiaNU3ZxQwtqLQbcMIHqwwZv2vry/Z
# yR7n1hb4q7qSUJHvTQ3p5JPaNcTzEopvaLkW/ZP/cO2tWEteJgQwfLDy9p+uW4Qm
# MikUEe8kLidHLgZ+4pg+VPqw4/2E7vDYBWbJKKGCFM8wghTLBgorBgEEAYI3AwMB
# MYIUuzCCFLcGCSqGSIb3DQEHAqCCFKgwghSkAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIKAh3Z6aou/4HvbDOgAz+DKTAAo/dd2BAJCxa1BVCZ7TAggztqRSsvT3
# 8xgPMjAyNjExMTAxMDIzMjZaMAMCAQECEDFMk1uK4+UyTMhPKnGwQSmgghGnMIIE
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
# wzSPOzGCAlcwggJTAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjMyNlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIIrF
# b8r4MdTUS6m+Acxwl75Y46z+lMpozWWE01Itf5cRMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIESDBGAiEAmFHZNgwRcashnpphLrEH1QBBFwtdL6Wf7+5PXV6iO+wCIQCD
# lwl3CNAF7eY16+HgcFsA2hmUv60qVYcxV1Q1umVMOw==
# SIG # End signature block
