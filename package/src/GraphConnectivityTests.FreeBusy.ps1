<#
.SYNOPSIS
    Graph Connectivity Tests - the free/busy test (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Outlook signs in the source mailbox and asks the Availability service for the free/busy of a target
    mailbox; Exchange then finds the target - in the same organisation, on-premises through the hybrid configuration, or in another
    organisation through an organization relationship - and says what went wrong. This test does the
    same in the user context with Microsoft Graph, which replaces EWS in Exchange Online:

      Source    the signed-in user and its mailbox: Exchange Online (Microsoft Graph), or on-premises
                (Graph cannot reach it: the request goes to EWS on-premises, like Outlook)
      Target    each target in the directory of the source, and its MailTips (how Exchange Online sees
                the recipient: internal, external partner, external)
      FreeBusy  getSchedule (or GetUserAvailability on-premises) for all the targets at once: the
                availability view, the calendar items and the detail the source may see, the working
                hours, and for each error the cause and what to check
      Hybrid    optional (Test.AdminChecks): what Exchange Online is configured with for the domains of
                the targets - IntraOrganizationConnector, organization relationships, availability
                address spaces - read by an Exchange administrator through the administration API

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

function Invoke-GctStageSource {
    <# The signed-in user (or the source for an application), then its mailbox: reachable through Graph, else on-premises. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Source'
    $cfg = $Context.Config
    $me = Get-GctMePath -Context $Context
    $select = 'id,displayName,userPrincipalName,mail,onPremisesSyncEnabled,usageLocation,accountEnabled'
    $user = Invoke-GctGraph -Context $Context -Path "$($me)?`$select=$select" -Operation "GET $me"
    if (Test-GctGraphOk $user) {
        $u = $user.Json
        $Context.SourceUser = $u
        $d = Get-GctGraphDetails -Answer $user -More ([ordered]@{ DisplayName = $u['displayName']; UserPrincipalName = $u['userPrincipalName']; Mail = $u['mail']; SyncedFromAD = $u['onPremisesSyncEnabled']; UsageLocation = $u['usageLocation']; ObjectId = $u['id'] })
        $mail = [string]$u['mail']
        if ($mail -and $mail -ine [string]$cfg.SourceMailbox -and [string]$u['userPrincipalName'] -ine [string]$cfg.SourceMailbox) {
            Add-GctStep $Context $stage 'Signed-in user' Warning "Signed in as $($u['displayName']) ($($u['userPrincipalName']), mail $mail), not as $($cfg.SourceMailbox): the free/busy is asked by this account. Sign in with the source account, or set it as the source." $d
        }
        else { Add-GctStep $Context $stage 'Signed-in user' Passed "$($u['displayName']) ($($u['userPrincipalName']))$(if ($u['onPremisesSyncEnabled']) { ', synchronized from Active Directory' })." $d }
    }
    else {
        Add-GctStep $Context $stage 'Signed-in user' Warning "The user object was not read: $(Get-GctGraphFailureText -Context $Context -Answer $user -Permission 'User.Read')" (Get-GctGraphDetails -Answer $user)
    }

    # The mailbox: getSchedule of the source itself needs only a calendar permission and gives the working hours and the time zone.
    $start = [datetime]::UtcNow.Date
    $body = @{ schedules = @([string]$cfg.SourceMailbox); startTime = @{ dateTime = $start.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }; endTime = @{ dateTime = $start.AddDays(1).ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }; availabilityViewInterval = 60 }
    $answer = Invoke-GctGraph -Context $Context -Method POST -Path "$me/calendar/getSchedule" -Body $body -Headers @{ Prefer = 'outlook.timezone="UTC"' } -Operation 'POST calendar/getSchedule (source)'
    $info = if ((Test-GctGraphOk $answer) -and $answer.Json['value']) { @($answer.Json['value'])[0] } else { $null }
    if ($info -and -not $info['error']) {
        $own = ConvertFrom-GctSchedule -Info $info -Path 'Self' -IntervalMinutes 60 -StartUtc $start
        $Context.SourceProtocol = 'Graph'
        $Context.SourceSchedule = $own
        $text = if ($Context.Source -and $Context.Source.Location -eq 'OnPremises') { ' Autodiscover had named an on-premises server: the mailbox may have been moved recently (Autodiscover cache).' } else { '' }
        Add-GctStep $Context $stage 'Source mailbox' Passed "Mailbox in Exchange Online, reached through Microsoft Graph; working hours $($own.WorkingHours).$text" (Get-GctGraphDetails -Answer $answer -More ([ordered]@{ WorkingHours = $own.WorkingHours; TimeZone = $own.WorkTimeZone }))
        return
    }
    $code = if ($info -and $info['error']) { [string]$info['error']['responseCode'] } else { [string]$answer.ErrorCode }
    $onPremises = $code -match 'MailboxNotEnabledForRESTAPI|MailboxNotSupportedForRESTAPI|ErrorMailRecipientNotFound|ProxyRequest|RecipientNotFound' -or ($Context.Source -and $Context.Source.Location -eq 'OnPremises' -and $answer.HttpStatus -in 400, 404)
    $ewsUrl = if ($Context.Source -and $Context.Source.Location -eq 'OnPremises') { $Context.Source.EwsUrl } elseif ([string]$cfg.OnPremisesHost) { "https://$($cfg.OnPremisesHost)/EWS/Exchange.asmx" } else { $null }
    if ($onPremises -and $ewsUrl -and [string]$cfg.Context -eq 'User') {
        $Context.SourceProtocol = 'EWS'
        $Context.SourceEwsUrl = $ewsUrl
        Add-GctStep $Context $stage 'Source mailbox' Passed "The mailbox is not in Exchange Online ($code): it is on-premises. Microsoft Graph cannot reach it; the free/busy request goes to the Availability service on-premises, $ewsUrl, like Outlook (direction on-premises to Microsoft 365)." (Get-GctGraphDetails -Answer $answer -More ([ordered]@{ EwsUrl = $ewsUrl; OnPremisesAuthentication = $cfg.OnPremisesAuthentication }))
        return
    }
    $why = if ($info -and $info['error']) { "$code - $($info['error']['message'])" } else { Get-GctGraphFailureText -Context $Context -Answer $answer -Permission 'Calendars.Read' }
    $more = if ($onPremises) { ' The mailbox seems to be on-premises, but no on-premises EWS URL is known: give Target.OnPremisesHost (-OnPremisesHost).' } else { '' }
    Add-GctStep $Context $stage 'Source mailbox' Failed "The source mailbox cannot be used: $why.$more" (Get-GctGraphDetails -Answer $answer)
    $Context.Stop = $true
}

function Get-GctTargetsInTenant {
    <# The targets whose domain belongs to the tenant of the token (their directory object can be read). #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    $tid = if ($Context.Token) { [string]$Context.Token.TenantId } else { $Context.SourceTenantId }
    foreach ($t in @($Context.Config.TargetMailboxes)) {
        $m = $Context.Mailboxes[([string]$t).ToLowerInvariant()]
        if ($m -and $m.Location -eq 'External') { continue }
        if (-not $m -or -not $m.TenantId -or -not $tid -or $m.TenantId -ieq $tid) { $t }
    }
}

function Find-GctDirectoryUser {
    <# The user of the tenant with this address as mail or user principal name. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Address, [string]$Select = 'id,displayName,mail,userPrincipalName,userType,accountEnabled,onPremisesSyncEnabled,proxyAddresses')
    $a = ConvertTo-GctODataText $Address
    $answer = Invoke-GctGraph -Context $Context -Path "/users?`$filter=mail eq '$a' or userPrincipalName eq '$a'&`$select=$Select" -Operation 'GET users?$filter=mail'
    $user = if ((Test-GctGraphOk $answer) -and @($answer.Json['value']).Count) { @($answer.Json['value'])[0] } else { $null }
    if (-not $user -and (Test-GctGraphOk $answer)) {
        # A secondary address: proxyAddresses needs an advanced query (ConsistencyLevel eventual, $count).
        $more = Invoke-GctGraph -Context $Context -Path "/users?`$filter=proxyAddresses/any(p:p eq 'smtp:$a')&`$count=true&`$select=$Select" -Headers @{ ConsistencyLevel = 'eventual' } -Operation 'GET users?$filter=proxyAddresses'
        if ((Test-GctGraphOk $more) -and @($more.Json['value']).Count) { $user = @($more.Json['value'])[0]; $answer = $more }
    }
    [pscustomobject]@{ Answer = $answer; User = $user }
}

function Invoke-GctStageTarget {
    <# The targets in the directory of the source tenant, and their MailTips. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Target'
    if (-not @($Context.Config.TargetMailboxes).Count) { Add-GctStep $Context $stage 'Target mailboxes' Skipped 'No target mailbox given (Target.TargetMailboxes, -TargetMailbox).'; return }
    $cfg = $Context.Config
    $inTenant = @(Get-GctTargetsInTenant -Context $Context)
    foreach ($t in @($cfg.TargetMailboxes)) {
        $m = $Context.Mailboxes[([string]$t).ToLowerInvariant()]
        if ($t -notin $inTenant) {
            Add-GctStep $Context $stage "Directory $t" Passed "$t is $(if ($m.TenantId) { "in another organisation (tenant $($m.TenantId))" } else { Get-GctLocationText $m }): it is not in the directory of the source; Exchange Online can only find it through an organization relationship or an availability address space." ([ordered]@{ TenantId = $m.TenantId; Path = $m.Relation })
            continue
        }
        $found = Find-GctDirectoryUser -Context $Context -Address $t
        $a = $found.Answer
        if (-not (Test-GctGraphOk $a)) {
            Add-GctStep $Context $stage "Directory $t" Warning "Not read: $(Get-GctGraphFailureText -Context $Context -Answer $a -Permission 'User.ReadBasic.All')" (Get-GctGraphDetails -Answer $a)
            continue
        }
        $u = $found.User
        if (-not $u) {
            Add-GctStep $Context $stage "Directory $t" Warning "No user of the tenant has $t as mail, user principal name or secondary address: Exchange Online may still know it as a contact or a group; otherwise expect ErrorMailRecipientNotFound." (Get-GctGraphDetails -Answer $a)
            continue
        }
        $Context.TargetUsers[$t.ToLowerInvariant()] = $u
        $d = Get-GctGraphDetails -Answer $a -More ([ordered]@{ DisplayName = $u['displayName']; Mail = $u['mail']; UserPrincipalName = $u['userPrincipalName']; UserType = $u['userType']; AccountEnabled = $u['accountEnabled']; SyncedFromAD = $u['onPremisesSyncEnabled']; ObjectId = $u['id'] })
        $where = if ($m -and $m.Location -eq 'OnPremises') { ' Its mailbox is on-premises: in Exchange Online it must be a mail user (RemoteMailbox on-premises) with this address.' } else { '' }
        Add-GctStep $Context $stage "Directory $t" Passed "$($u['displayName']) ($($u['userType'])$(if ($u['onPremisesSyncEnabled']) { ', synchronized from Active Directory' }))$(if ($u['accountEnabled'] -eq $false) { ', account disabled (normal for a shared mailbox or a room)' }).$where" $d
    }

    # MailTips: how Exchange Online sees each target from the source.
    $me = Get-GctMePath -Context $Context
    $body = @{ EmailAddresses = @($cfg.TargetMailboxes); MailTipsOptions = 'automaticReplies, mailboxFullStatus, recipientScope, maxMessageSize, deliveryRestriction, externalMemberCount, totalMemberCount' }
    $tips = Invoke-GctGraph -Context $Context -Method POST -Path "$me/getMailTips" -Body $body -Operation 'POST getMailTips'
    if (-not (Test-GctGraphOk $tips)) {
        Add-GctStep $Context $stage 'MailTips' Warning "MailTips not read: $(Get-GctGraphFailureText -Context $Context -Answer $tips -Permission 'Mail.Read')" (Get-GctGraphDetails -Answer $tips)
        return
    }
    foreach ($tip in @($tips.Json['value'])) {
        $address = [string]$tip['emailAddress']['address']
        $m = $Context.Mailboxes[$address.ToLowerInvariant()]
        $path = if ($m) { $m.Relation } else { $null }
        $scope = [string]$tip['recipientScope']
        # A MailTip in error has only emailAddress and error: no automaticReplies to index.
        $replies = if ($tip['automaticReplies']) { [string]$tip['automaticReplies']['message'] } else { '' }
        $d = [ordered]@{ RecipientScope = $scope; AutomaticReplies = $replies; MailboxFull = $tip['mailboxFull']; MaxMessageSize = $tip['maxMessageSize']; DeliveryRestricted = $tip['deliveryRestricted']; TotalMembers = $tip['totalMemberCount'] }
        if ($tip['error']) {
            Add-GctStep $Context $stage "MailTips $address" Warning "Exchange Online gave no MailTips: $($tip['error']['code']) - $($tip['error']['message'])" $d
            continue
        }
        $oof = if ($replies) { ' Automatic replies are on.' } else { '' }
        $status = 'Passed'
        # Anchored: 'externalPartner' must not also match the 'external' case.
        $text = switch -Regex ($scope) {
            '^internal$' { "Exchange Online sees $address as internal (a recipient of the organisation$(if ($path -eq 'CloudToOnPremises') { ': the mail user of the hybrid' }))." }
            '^externalPartner$' { "Exchange Online sees $address as an external partner: an organization relationship (MailTips access) exists for its domain." }
            '^(external|externalNonPartner)$' {
                if ($path -in 'CloudToOnPremises', 'SameOrganisation') { $status = 'Warning'; "Exchange Online sees $address as external although it should be in the organisation: its domain is not an accepted domain, or the mail user / mailbox does not exist in Exchange Online." }
                else { "Exchange Online sees $address as external, not as a partner: no organization relationship shares MailTips with its domain (MailTips and free/busy are separate settings of the relationship)." }
            }
            default { "Recipient scope: $(if ($scope) { $scope } else { 'none' })." }
        }
        Add-GctStep $Context $stage "MailTips $address" $status "$text$oof" $d
    }
}

function Get-GctFreeBusyWindow {
    param([Parameter(Mandatory = $true)][hashtable]$Context)
    $start = (Get-Date).Date.ToUniversalTime()
    [pscustomobject]@{ StartUtc = $start; EndUtc = $start.AddDays([int]$Context.Config.FreeBusyDays); IntervalMinutes = [int]$Context.Config.FreeBusyIntervalMinutes; Days = [int]$Context.Config.FreeBusyDays }
}

function Invoke-GctStageFreeBusy {
    <# The free/busy of the targets, asked as the source: getSchedule (Exchange Online) or GetUserAvailability (on-premises). #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'FreeBusy'
    if (-not @($Context.Config.TargetMailboxes).Count) { Add-GctStep $Context $stage 'Free/busy' Skipped 'No target mailbox given (Target.TargetMailboxes, -TargetMailbox): the own free/busy is in the Availability stage.'; return }
    $cfg = $Context.Config
    $targets = @($cfg.TargetMailboxes)
    $w = Get-GctFreeBusyWindow -Context $Context
    $Context.FreeBusyWindow = [ordered]@{ StartUtc = $w.StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'); EndUtc = $w.EndUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'); IntervalMinutes = $w.IntervalMinutes; Days = $w.Days }
    $paths = @{}
    foreach ($t in $targets) { $m = $Context.Mailboxes[$t.ToLowerInvariant()]; $paths[$t.ToLowerInvariant()] = if ($m) { $m.Relation } else { $null } }
    $results = @()
    if ($Context.SourceProtocol -eq 'EWS') {
        $envelope = New-GctAvailabilityEnvelope -Mailboxes $targets -StartUtc $w.StartUtc -EndUtc $w.EndUtc -IntervalMinutes $w.IntervalMinutes
        try { $answer = Invoke-GctEwsAvailability -Context $Context -EwsUrl $Context.SourceEwsUrl -Envelope $envelope }
        catch {
            Add-GctStep $Context $stage 'GetUserAvailability' Failed "No answer from EWS on-premises ($($Context.SourceEwsUrl)): $($_.Exception.Message)" ([ordered]@{ EwsUrl = $Context.SourceEwsUrl })
            return
        }
        $d = [ordered]@{ EwsUrl = $Context.SourceEwsUrl; HttpStatus = $answer.HttpStatus; Authentication = $answer.Authentication; Servers = $answer.Server; ExchangeVersion = $answer.Version; Diagnostics = $answer.ErrorHeader; Challenges = $answer.Challenges -join ' | ' }
        if ($answer.HttpStatus -ne 200 -or -not $answer.Xml -or $answer.Fault) {
            $hint = switch ($answer.HttpStatus) {
                401 {
                    switch ([string]$cfg.OnPremisesAuthentication) {
                        'OAuth' { "Exchange on-premises refused the token of hybrid modern authentication$(if ($answer.ErrorHeader) { " ($($answer.ErrorHeader))" }): EvoSts authorization server (Get-AuthServer), OAuthAuthentication on the EWS virtual directory, the URL among the service principal names of Office 365 Exchange Online. Or try -OnPremisesAuthentication Credential (NTLM or Basic with the account of the source) or Windows from a computer of the domain." }
                        'Credential' { "Exchange on-premises refused the account ($($answer.Authentication)): password, account locked or expired, user name format (DOMAIN\user or the UPN); the scheme must be enabled on the EWS virtual directory (Get-WebServicesVirtualDirectory: BasicAuthentication, WindowsAuthentication) and kept by the reverse proxy$(if (@($answer.Offered).Count) { ". EWS offers: $(@($answer.Offered) -join ', ')" })." }
                        default { 'Exchange on-premises refused the Windows account: run from a computer of the domain with the source account, or use Credential (an account typed) or OAuth.' }
                    }
                }
                403 { 'EWS is blocked for this user or application (Get-CASMailbox EwsEnabled, EwsApplicationAccessPolicy) or by a reverse proxy.' }
                default { if ($answer.Fault) { "SOAP fault $($answer.Fault.Code): $($answer.Fault.Text)" } else { 'no EWS answer.' } }
            }
            Add-GctStep $Context $stage 'GetUserAvailability' Failed "EWS on-premises answered HTTP $($answer.HttpStatus): $hint" $d
            return
        }
        $basicNote = if ($answer.Authentication -like 'Basic*') { ' Basic works here, but the password goes in every request: prefer hybrid modern authentication (OAuth) for Outlook.' } else { '' }
        Add-GctStep $Context $stage 'GetUserAvailability' Passed "EWS on-premises answered (Exchange $($answer.Version), $($answer.Authentication)) for $($targets.Count) mailbox(es).$basicNote" $d
        $results = @(ConvertFrom-GctEwsAvailability -Answer $answer -Mailboxes $targets -Paths $paths -IntervalMinutes $w.IntervalMinutes -StartUtc $w.StartUtc)
    }
    else {
        $me = Get-GctMePath -Context $Context
        $body = @{
            schedules = $targets
            startTime = @{ dateTime = $w.StartUtc.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }
            endTime   = @{ dateTime = $w.EndUtc.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }
            availabilityViewInterval = $w.IntervalMinutes
        }
        $answer = Invoke-GctGraph -Context $Context -Method POST -Path "$me/calendar/getSchedule" -Body $body -Headers @{ Prefer = 'outlook.timezone="UTC"' } -Operation 'POST calendar/getSchedule'
        if (-not (Test-GctGraphOk $answer)) {
            Add-GctStep $Context $stage 'getSchedule' Failed "Free/busy not asked: $(Get-GctGraphFailureText -Context $Context -Answer $answer -Permission 'Calendars.Read')" (Get-GctGraphDetails -Answer $answer)
            return
        }
        Add-GctStep $Context $stage 'getSchedule' Passed "Microsoft Graph answered for $($targets.Count) mailbox(es), $($w.Days) day(s) from $($w.StartUtc.ToLocalTime().ToString('yyyy-MM-dd')), slots of $($w.IntervalMinutes) min." (Get-GctGraphDetails -Answer $answer)
        $byAddress = @{}
        foreach ($info in @($answer.Json['value'])) { $byAddress[([string]$info['scheduleId']).ToLowerInvariant()] = $info }
        $results = @(foreach ($t in $targets) {
                $info = $byAddress[$t.ToLowerInvariant()]
                if ($info) { ConvertFrom-GctSchedule -Info $info -Path $paths[$t.ToLowerInvariant()] -IntervalMinutes $w.IntervalMinutes -StartUtc $w.StartUtc }
                else {
                    [pscustomobject]@{ Mailbox = $t; Protocol = 'Microsoft Graph getSchedule'; Path = $paths[$t.ToLowerInvariant()]; Status = 'Failed'; ResponseCode = 'NoResponse'; Message = 'getSchedule returned no scheduleInformation for this address.'; Hint = $null
                        View = ''; IntervalMinutes = $w.IntervalMinutes; StartUtc = $w.StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'); DetailLevel = $null; WorkingHours = $null; WorkDays = @(); WorkStart = $null; WorkEnd = $null; WorkTimeZone = $null; WorkOffsetMinutes = $null; Events = @() }
                }
            })
    }
    $Context.FreeBusy = $results
    foreach ($r in $results) {
        $d = [ordered]@{ Protocol = $r.Protocol; Path = $r.Path; ResponseCode = $r.ResponseCode; DetailLevel = $r.DetailLevel; WorkingHours = $r.WorkingHours; Items = @($r.Events).Count; AvailabilityView = $(if ($r.View.Length -gt 96) { $r.View.Substring(0, 96) + '...' } else { $r.View }) }
        if ($r.Status -eq 'Failed') {
            $d.Message = $r.Message
            Add-GctStep $Context $stage "Free/busy of $($r.Mailbox)" Failed "$($r.ResponseCode): $(([string]$r.Message -split "`r?`n")[0])$(if ($r.Hint) { " Cause: $($r.Hint)" })" $d
            continue
        }
        $busy = @($r.Events | Where-Object { $_.BusyType -notin 'Free' }).Count
        $slots = @($r.View.ToCharArray() | Where-Object { $_ -in '1', '2', '3', '4' }).Count
        $text = "$($r.DetailLevel): $busy calendar item(s), $slots busy slot(s) of $($r.IntervalMinutes) min over $($w.Days) day(s)$(if ($r.WorkingHours) { "; working hours $($r.WorkingHours)" })."
        if ($r.Status -eq 'Warning') { Add-GctStep $Context $stage "Free/busy of $($r.Mailbox)" Warning "No error but no availability view: Exchange answered without free/busy data." $d }
        else { Add-GctStep $Context $stage "Free/busy of $($r.Mailbox)" Passed $text $d }
    }
    if (-not $script:Quiet) { Write-GctFreeBusyGrid -Results $results -StartUtc $w.StartUtc -IntervalMinutes $w.IntervalMinutes -Days ([Math]::Min($w.Days, 7)) }
}

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAK20OxL6YBcSpT
# s6WCkDwUQHlb2sjw4biYYyF9MqERkqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAC4WfKWt95
# Y7W374RkZKJD+MadLOgHhPWKWYnEHDEu4TANBgkqhkiG9w0BAQEFAASCAgCOADH8
# 525+RBjS67lBCzCCgKZRHC+GQ/kyzcERanFCnUMkZt71usdqSeBamu3EUbR+iNcg
# TzceOfijGZWUtxIBOtkfSBG+Ioft2MuOWbnd3CCk3NQtQeVWJJU46XU+AnaQLf3C
# bhu+wjRhISnpQXflq95SKzvTo8EQVppf7O3FTnPzXaPBsY/p9XpYc4fgFQNf4PC/
# AazTYMZ1KIC4yFemd6oyhvil+KG75idddqEdAY55lRnTiL8X+v8Pjifll95Dic3+
# JfCgn2HC36tGfMv6r0W8a3vfujDWA9ouhcLUcYpstdl+7+qDA74SdYfuxRYf6nF/
# mmF1x+6G/XUsgEdNkgH81DOr15aEolEqidTqqhCVrUZwmLWstxixS30tr6vBlJ66
# +hVGIbBJb752zvggZhnCbuPZm96A5zpGNSapcxyJQ785QUzxjB6AZ7VOjOaKor5A
# 6BL7leN6HmH0eCqnS9RVe2QepWxA6AM9rYJoq8l6oGW1lnQF1GSsQxld+ev0udTE
# Rz6/h4E83gkxyKc2YlzpJfNm4RqbZh6a1nOe662+D5aWcWCJ9nB02We464gupnSz
# 6KvEsdVBHvVqMUFSQDHglsuQI9V4Ski6RXlLUgz6RCnRGseNEuCvPScfAcj83xsQ
# yw08kasJlhh5GEXaevfwwZSHYHSVPQF+/5wZSaGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIJ39XvQFUvQfFJnKpUgKc8SbiK+gQSGhxeZ65uwtrhr/AggbTc/9Tf2q
# WRgPMjAyNjExMTAxMDIzMzlaMAMCAQECEH+BS5sp0OiluDfX8gsJeRigghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjMzOVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIDtb
# a/e6vrW4C11NxTc6J4E3AKq7hv3BCGfbpAnm+9zYMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiAiG/tJWKX1suzaoP47DodT4SzSaIJczqLdk6tXM0eKjAIhAKtn
# 80DZC5yYcMfxSTiUIT3JvTHiCTWasvlha9IqM0aG
# SIG # End signature block
