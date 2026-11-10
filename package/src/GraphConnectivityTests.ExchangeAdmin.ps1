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

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAF+vjbL50AqJgs
# wtE0cOKy+1meNMm7huXjetDktJCcWqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYEDCCGAwCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCwpGIpoGYS
# FATBXW0F8ezV00R1U5ZYR4i93fJyvdEiujANBgkqhkiG9w0BAQEFAASCAgAh2sGF
# 4jUsjLKSygYeQigV//zHvJDvql0eP/okjIvlK5BZRzysSKxLwWdUJ0Du6wUsZ2a3
# XZO7jlPmAkOJrjRWr2K6t1t94xCmsX1MfLYRHhehg8jiHzbl267LcrHjxJZVaRSd
# 0X3qjXy1oCeBm1AnKE6QNZKHWaqQE4vjZX6DDXkDgYz5JoHfcdin9bKTqTIgFZJf
# TqEkVAUAvkreB8LURC98LiS36CzxdiGBAAd3LY5BmLF1gml0EJ3bL9RMUcuOUC4N
# ozjq2euSCOocnt08eABGTvBTX2A1FZmJC5cM9edY9CsFeLMR+Pcb1ebH4Iekls5u
# 1GoPKsGo0y6TcSGnLuGBpdw5LfeQeNjG1Ic3aUIQUlBZPMT/jp5eZjr2G9Kb3doq
# hehOzCHNR62mfhCyj9ogu+OVFMixUPotcNjh+Z2xKYvcME7GpsOKfYMeoTf5V1D6
# tAvG4Q5bhED+0SznoYUoZbqIcX3vQSg5nIXwpV+lyBZTgaRztDm83OfXgaJ1eyNL
# 6TRM2cZ46k6DpjndfD+rYz2tJFJud1R5oO7fGvwtiKvQkKvMxQviz0Id8BSwtPZP
# OI3ORjNvHtfTOO16RCfs4nbH/zqkq+ysz/mc2rLMkSDuywwXHnHNJEBJVyAGaIX/
# ZTlAbIK5y1hMLwJmZXzpmIkZOY5IKUahwz+LPaGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEILuPouM3O014oRosQKsp7jycwOdVYFs6sUrtR+UvpyYXAggFm7jzkrzI
# chgPMjAyNjExMTAxMDIzMzZaMAMCAQECEBlHP5D8dvaiSf/Qttxoqs+gghGnMIIE
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
# wzSPOzGCAlUwggJRAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjMzNlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIFgx
# XCp+Ek6thISqam8t1ycUE6fDnHLkVXvExF/g4nFmMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiBNKCvj5VO0NaL1PvJ2jSrz8dQ9YpBgiHdkxzTcZemI0AIgW27K
# 32dHHy7Ncu4FfMYUI36koNCTpoT1lMXWWxAzzIw=
# SIG # End signature block
