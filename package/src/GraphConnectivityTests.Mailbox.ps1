<#
.SYNOPSIS
    Graph Connectivity Tests - mailbox services and mailbox provisioning (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    MailboxServices: the services a mail client relies on - in EWS SyncFolderHierarchy, SyncFolderItems,
    push or streaming notifications, GetUserAvailability, Get/SetUserOofSettings. Their Microsoft Graph
    equivalents:

      Sync          mailFolders/delta and messages/delta: a first synchronization, then the deltaLink
                    sent again must return only the changes
      Notifications change notifications (subscriptions) need a public HTTPS endpoint that Microsoft
                    Graph calls: not testable from a workstation, the step says why (delta replaces it)
      Availability  getSchedule of the mailbox itself, and its MailTips
      AutoReply     mailboxSettings/automaticRepliesSetting read; with -AllowWrite, scheduled in the
                    future, read back and restored - only when the automatic replies are disabled

    Provisioning: like the Mailbox Provisioning test, for each address: the user in Entra ID, its
    Exchange licenses, synchronization from Active Directory, and whether Exchange Online has a mailbox
    for it - without reading anything of the mailbox: Microsoft Graph answers 403 (a mailbox the signed-in
    user cannot open) or 404 MailboxNotEnabledForRESTAPI (no mailbox in Exchange Online).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:GctOofTestText = '[Graph Connectivity Tests] Test of the automatic replies, restored at once by the tool.'

function Invoke-GctDeltaRound {
    <# A first synchronization (pages followed up to MaxPages), then the deltaLink sent again. Returns counts and links. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context, [Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Name, [int]$MaxPages = 20)

    $url = $Path
    $count = 0; $pages = 0; $delta = $null; $last = $null
    while ($url -and $pages -lt $MaxPages) {
        $last = Invoke-GctGraph -Context $Context -Path $url -Headers @{ Prefer = 'odata.maxpagesize=50' } -Operation "GET $Name delta (page $($pages + 1))"
        if (-not (Test-GctGraphOk $last)) { return [pscustomobject]@{ Ok = $false; Answer = $last; Count = $count; Pages = $pages; DeltaLink = $null; Changes = $null } }
        $pages++
        $count += @($last.Json['value']).Count
        $delta = [string]$last.Json['@odata.deltaLink']
        $url = [string]$last.Json['@odata.nextLink']
        if ($delta) { break }
    }
    $changes = $null; $again = $null
    if ($delta) {
        $again = Invoke-GctGraph -Context $Context -Path $delta -Headers @{ Prefer = 'odata.maxpagesize=50' } -Operation "GET $Name delta (deltaLink)"
        if (Test-GctGraphOk $again) { $changes = @($again.Json['value']).Count }
    }
    [pscustomobject]@{ Ok = $true; Answer = $last; Again = $again; Count = $count; Pages = $pages; DeltaLink = [bool]$delta; Changes = $changes; Truncated = [bool]($url -and -not $delta) }
}

function Invoke-GctStageSync {
    <# Folder and message synchronization (delta), and why change notifications are not tested here. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Sync'
    $me = Get-GctMePath -Context $Context
    if ($Context.SourceProtocol -eq 'EWS') { Add-GctStep $Context $stage 'Synchronization' Skipped 'The source mailbox is on-premises: Microsoft Graph cannot synchronize it.'; return }
    $folders = Invoke-GctDeltaRound -Context $Context -Path "$me/mailFolders/delta?`$select=displayName,parentFolderId" -Name 'mailFolders'
    if (-not $folders.Ok) { Add-GctStep $Context $stage 'Folder synchronization' Failed "mailFolders/delta refused: $(Get-GctGraphFailureText -Context $Context -Answer $folders.Answer -Permission 'Mail.ReadBasic')" (Get-GctGraphDetails -Answer $folders.Answer) }
    elseif (-not $folders.DeltaLink) { Add-GctStep $Context $stage 'Folder synchronization' Warning "$($folders.Count) folder(s) in $($folders.Pages) page(s) without reaching the deltaLink (more than $($folders.Pages) pages)." ([ordered]@{ Folders = $folders.Count; Pages = $folders.Pages }) }
    else {
        $status = if ($null -eq $folders.Changes) { 'Warning' } else { 'Passed' }
        Add-GctStep $Context $stage 'Folder synchronization' $status "First synchronization: $($folders.Count) folder(s) in $($folders.Pages) page(s); the deltaLink sent again returns $(if ($null -eq $folders.Changes) { 'an error' } else { "$($folders.Changes) change(s)" }), like SyncFolderHierarchy with a sync state." ([ordered]@{ Folders = $folders.Count; Pages = $folders.Pages; ChangesAfter = $folders.Changes })
    }
    $since = [datetime]::UtcNow.AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $messages = Invoke-GctDeltaRound -Context $Context -Path "$me/mailFolders/inbox/messages/delta?`$select=subject,receivedDateTime&`$filter=receivedDateTime+ge+$since" -Name 'messages' -MaxPages 10
    if (-not $messages.Ok) { Add-GctStep $Context $stage 'Message synchronization' Failed "messages/delta refused: $(Get-GctGraphFailureText -Context $Context -Answer $messages.Answer -Permission 'Mail.ReadBasic')" (Get-GctGraphDetails -Answer $messages.Answer) }
    else {
        $status = if ($messages.DeltaLink -and $null -ne $messages.Changes) { 'Passed' } else { 'Warning' }
        $tail = if ($messages.DeltaLink) { "the deltaLink sent again returns $($messages.Changes) change(s)" } else { "the deltaLink was not reached in $($messages.Pages) pages" }
        Add-GctStep $Context $stage 'Message synchronization' $status "Inbox, last 7 days: $($messages.Count) message(s) (headers only) in $($messages.Pages) page(s); $tail, like SyncFolderItems." ([ordered]@{ Messages = $messages.Count; Pages = $messages.Pages; ChangesAfter = $messages.Changes; Since = $since })
    }
    Add-GctStep $Context $stage 'Change notifications' Skipped 'Not testable from a workstation: Microsoft Graph sends change notifications (subscriptions) to a public HTTPS endpoint of the application, which must answer its validation. The delta synchronization above is what a client without such an endpoint uses.'
}

function Invoke-GctStageAvailability {
    <# The free/busy of the mailbox itself and its MailTips. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Availability'
    $cfg = $Context.Config
    $me = Get-GctMePath -Context $Context
    if ($Context.SourceProtocol -eq 'EWS') { Add-GctStep $Context $stage 'Availability' Skipped 'The source mailbox is on-premises: run the FreeBusy test (EWS on-premises).'; return }
    $w = Get-GctFreeBusyWindow -Context $Context
    $body = @{ schedules = @([string]$cfg.SourceMailbox); startTime = @{ dateTime = $w.StartUtc.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }; endTime = @{ dateTime = $w.EndUtc.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }; availabilityViewInterval = $w.IntervalMinutes }
    $answer = Invoke-GctGraph -Context $Context -Method POST -Path "$me/calendar/getSchedule" -Body $body -Headers @{ Prefer = 'outlook.timezone="UTC"' } -Operation 'POST calendar/getSchedule (own)'
    $info = if ((Test-GctGraphOk $answer) -and $answer.Json['value']) { @($answer.Json['value'])[0] } else { $null }
    if (-not $info) { Add-GctStep $Context $stage 'Own free/busy' Failed "getSchedule refused: $(Get-GctGraphFailureText -Context $Context -Answer $answer -Permission 'Calendars.Read')" (Get-GctGraphDetails -Answer $answer) }
    else {
        $r = ConvertFrom-GctSchedule -Info $info -Path 'Self' -IntervalMinutes $w.IntervalMinutes -StartUtc $w.StartUtc
        if (-not @($Context.FreeBusy | Where-Object { $_.Mailbox -ieq $r.Mailbox }).Count) { $Context.FreeBusy = @($Context.FreeBusy) + $r }
        if (-not $Context.FreeBusyWindow) { $Context.FreeBusyWindow = [ordered]@{ StartUtc = $w.StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'); EndUtc = $w.EndUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'); IntervalMinutes = $w.IntervalMinutes; Days = $w.Days } }
        if ($r.Status -eq 'Failed') { Add-GctStep $Context $stage 'Own free/busy' Failed "$($r.ResponseCode): $($r.Message)" (Get-GctGraphDetails -Answer $answer) }
        else { Add-GctStep $Context $stage 'Own free/busy' Passed "$(@($r.Events).Count) calendar item(s) over $($w.Days) day(s); working hours $($r.WorkingHours), like GetUserAvailability." (Get-GctGraphDetails -Answer $answer -More ([ordered]@{ WorkingHours = $r.WorkingHours; Items = @($r.Events).Count })) }
    }
    $tips = Invoke-GctGraph -Context $Context -Method POST -Path "$me/getMailTips" -Body @{ EmailAddresses = @([string]$cfg.SourceMailbox); MailTipsOptions = 'automaticReplies, mailboxFullStatus, maxMessageSize, recipientScope' } -Operation 'POST getMailTips (own)'
    if (-not (Test-GctGraphOk $tips)) { Add-GctStep $Context $stage 'Own MailTips' Warning "MailTips not read: $(Get-GctGraphFailureText -Context $Context -Answer $tips -Permission 'Mail.Read')" (Get-GctGraphDetails -Answer $tips); return }
    $t = @($tips.Json['value'])[0]
    $Context.OwnMailTips = $t
    $oof = [string]$t['automaticReplies']['message']
    Add-GctStep $Context $stage 'Own MailTips' Passed "Mailbox $(if ($t['mailboxFull']) { 'FULL' } else { 'not full' }), maximum message size $([Math]::Round([double]$t['maxMessageSize'] / 1MB, 1)) MB, automatic replies $(if ($oof) { 'on' } else { 'off' })." (Get-GctGraphDetails -Answer $tips -More ([ordered]@{ MailboxFull = $t['mailboxFull']; MaxMessageSize = $t['maxMessageSize']; AutomaticReply = $(if ($oof) { [regex]::Replace($oof, '<[^>]+>', ' ').Trim() }) }))
}

function Invoke-GctStageAutoReply {
    <# Automatic replies: read; with -AllowWrite (User context) scheduled in the future, read back, restored. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'AutoReply'
    $cfg = $Context.Config
    $me = Get-GctMePath -Context $Context
    if ($Context.SourceProtocol -eq 'EWS') { Add-GctStep $Context $stage 'Automatic replies' Skipped 'The source mailbox is on-premises: Microsoft Graph cannot reach it.'; return }
    $read = Invoke-GctGraph -Context $Context -Path "$me/mailboxSettings/automaticRepliesSetting" -Operation 'GET mailboxSettings/automaticRepliesSetting'
    if (-not (Test-GctGraphOk $read)) {
        $from = if ($Context.OwnMailTips) { " MailTips say they are $(if ([string]$Context.OwnMailTips['automaticReplies']['message']) { 'on' } else { 'off' })." } else { '' }
        $status = if ($read.HttpStatus -eq 403) { 'Warning' } else { 'Failed' }
        Add-GctStep $Context $stage 'Read the automatic replies' $status "Not read: $(Get-GctGraphFailureText -Context $Context -Answer $read -Permission 'MailboxSettings.Read')$from" (Get-GctGraphDetails -Answer $read)
        return
    }
    $current = $read.Json
    $state = [string]$current['status']
    Add-GctStep $Context $stage 'Read the automatic replies' Passed "Automatic replies $state (external audience: $($current['externalAudience'])), like GetUserOofSettings." (Get-GctGraphDetails -Answer $read -More ([ordered]@{ Status = $state; ExternalAudience = $current['externalAudience']; ScheduledStart = (ConvertFrom-GctGraphDate $current['scheduledStartDateTime']); ScheduledEnd = (ConvertFrom-GctGraphDate $current['scheduledEndDateTime']) }))
    if (-not $cfg.AllowWrite) { Add-GctStep $Context $stage 'Change the automatic replies' Skipped 'Not changed: -AllowWrite was not given (the test would schedule the automatic replies in the future, read them back and restore them).'; return }
    if ([string]$cfg.Context -eq 'Application') { Add-GctStep $Context $stage 'Change the automatic replies' Skipped 'Changed only in the User context.'; return }
    if ($state -ne 'disabled') { Add-GctStep $Context $stage 'Change the automatic replies' Skipped "Not changed: the automatic replies are $state - the tool changes them only when they are disabled, so that a real absence is never touched."; return }

    $start = [datetime]::UtcNow.Date.AddDays(60); $end = $start.AddDays(1)
    $patch = @{ automaticRepliesSetting = @{ status = 'scheduled'; scheduledStartDateTime = @{ dateTime = $start.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }; scheduledEndDateTime = @{ dateTime = $end.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }; internalReplyMessage = $script:GctOofTestText; externalReplyMessage = $script:GctOofTestText } }
    $set = Invoke-GctGraph -Context $Context -Method PATCH -Path "$me/mailboxSettings" -Body $patch -Operation 'PATCH mailboxSettings (schedule)'
    $Context.Actions.Add([pscustomobject]@{ TimeUtc = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'); Action = 'Schedule automatic replies'; Target = [string]$cfg.SourceMailbox; Result = "HTTP $($set.HttpStatus)"; Detail = "$($start.ToString('yyyy-MM-dd')) to $($end.ToString('yyyy-MM-dd')) UTC" })
    if (-not (Test-GctGraphOk $set)) { Add-GctStep $Context $stage 'Change the automatic replies' Failed "Not changed: $(Get-GctGraphFailureText -Context $Context -Answer $set -Permission 'MailboxSettings.ReadWrite')" (Get-GctGraphDetails -Answer $set); return }
    $back = Invoke-GctGraph -Context $Context -Path "$me/mailboxSettings/automaticRepliesSetting" -Operation 'GET automaticRepliesSetting (read back)'
    $ok = (Test-GctGraphOk $back) -and [string]$back.Json['status'] -eq 'scheduled'
    $restore = Invoke-GctGraph -Context $Context -Method PATCH -Path "$me/mailboxSettings" -Body @{ automaticRepliesSetting = @{ status = 'disabled'; internalReplyMessage = [string]$current['internalReplyMessage']; externalReplyMessage = [string]$current['externalReplyMessage'] } } -Operation 'PATCH mailboxSettings (restore)'
    $Context.Actions.Add([pscustomobject]@{ TimeUtc = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'); Action = 'Restore automatic replies'; Target = [string]$cfg.SourceMailbox; Result = "HTTP $($restore.HttpStatus)"; Detail = 'status disabled, original messages' })
    $restored = Test-GctGraphOk $restore
    $status = if ($ok -and $restored) { 'Passed' } elseif (-not $restored) { 'Failed' } else { 'Warning' }
    $text = if (-not $restored) { "Scheduled, but NOT restored: $(Get-GctGraphFailureText -Context $Context -Answer $restore) Disable the automatic replies in Outlook." }
    elseif ($ok) { 'Scheduled in 60 days, read back as scheduled, then restored (disabled, original messages), like SetUserOofSettings.' }
    else { 'Changed and restored, but the read-back did not show the scheduled state.' }
    Add-GctStep $Context $stage 'Change the automatic replies' $status $text (Get-GctGraphDetails -Answer $restore -More ([ordered]@{ ReadBack = $(if (Test-GctGraphOk $back) { [string]$back.Json['status'] }) }))
}

function Get-GctExchangePlans {
    <# Exchange service plans of a user (assignedPlans of the user object), and whether one is enabled. #>
    param([AllowNull()][object]$User)
    $plans = @($User['assignedPlans'] | Where-Object { $_ -and [string]$_['service'] -match '^exchange$|^exchangeonline' })
    [pscustomobject]@{ Plans = $plans; Enabled = @($plans | Where-Object { [string]$_['capabilityStatus'] -eq 'Enabled' }).Count -gt 0 }
}

function Invoke-GctStageProvisioning {
    <# For each address: user, licenses, synchronization, mailbox in Exchange Online or not, and where Autodiscover sends it. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'Provisioning'
    $cfg = $Context.Config
    $addresses = if (@($cfg.TargetMailboxes).Count) { @($cfg.TargetMailboxes) } else { @([string]$cfg.SourceMailbox) }
    $tid = if ($Context.Token) { [string]$Context.Token.TenantId } else { $null }
    foreach ($address in $addresses) {
        $m = $Context.Mailboxes[$address.ToLowerInvariant()]
        if ($m -and $m.TenantId -and $tid -and $m.TenantId -ine $tid) {
            Add-GctStep $Context $stage "Provisioning of $address" Skipped "$address belongs to another tenant ($($m.TenantId)): its directory is not readable from here. Autodiscover: $(Get-GctLocationText $m)." ([ordered]@{ TenantId = $m.TenantId; Location = $m.Location })
            continue
        }
        $found = Find-GctDirectoryUser -Context $Context -Address $address -Select 'id,displayName,mail,userPrincipalName,userType,accountEnabled,onPremisesSyncEnabled,onPremisesImmutableId,proxyAddresses,assignedPlans,usageLocation,createdDateTime'
        if (-not (Test-GctGraphOk $found.Answer)) { Add-GctStep $Context $stage "Provisioning of $address" Failed "Directory not read: $(Get-GctGraphFailureText -Context $Context -Answer $found.Answer -Permission 'User.Read.All')" (Get-GctGraphDetails -Answer $found.Answer); continue }
        $u = $found.User
        if (-not $u) {
            $where = if ($m) { " Autodiscover: $(Get-GctLocationText $m)." } else { '' }
            Add-GctStep $Context $stage "Provisioning of $address" Failed "No user of the tenant has this address (mail, user principal name or proxy address): no mailbox of the tenant can answer for it. A contact or a group may still use it.$where" (Get-GctGraphDetails -Answer $found.Answer)
            continue
        }
        $id = [string]$u['id']
        $exchange = Get-GctExchangePlans -User $u
        $licenses = Invoke-GctGraph -Context $Context -Path "/users/$id/licenseDetails?`$select=skuPartNumber,servicePlans" -Operation 'GET users/{id}/licenseDetails'
        $skus = @(if (Test-GctGraphOk $licenses) { @($licenses.Json['value']) | ForEach-Object { [string]$_['skuPartNumber'] } })
        $exoPlans = @(if (Test-GctGraphOk $licenses) { @($licenses.Json['value']) | ForEach-Object { @($_['servicePlans']) } | Where-Object { [string]$_['servicePlanName'] -match '^EXCHANGE' } | ForEach-Object { '{0} ({1})' -f $_['servicePlanName'], $_['provisioningStatus'] } })
        # Whether Exchange Online has a mailbox: the Inbox asked without reading it - 200 (own), 403 (someone else's), 404 MailboxNotEnabledForRESTAPI (none).
        $probe = Invoke-GctGraph -Context $Context -Path "/users/$id/mailFolders/inbox?`$select=id" -Operation 'GET users/{id}/mailFolders/inbox (exists?)'
        $hasMailbox = switch ($probe.HttpStatus) { 200 { $true } 403 { $true } 404 { if ($probe.ErrorCode -match 'MailboxNotEnabledForRESTAPI|MailboxNotSupportedForRESTAPI') { $false } else { $null } } default { $null } }
        $purpose = Invoke-GctGraph -Context $Context -Path "/users/$id/mailboxSettings/userPurpose" -Operation 'GET users/{id}/mailboxSettings/userPurpose'
        $type = if (Test-GctGraphOk $purpose) { [string]$purpose.Json['value'] } else { $null }
        $d = [ordered]@{
            DisplayName = $u['displayName']; UserPrincipalName = $u['userPrincipalName']; Mail = $u['mail']; UserType = $u['userType']; AccountEnabled = $u['accountEnabled']
            SyncedFromAD = $u['onPremisesSyncEnabled']; ImmutableId = $u['onPremisesImmutableId']; UsageLocation = $u['usageLocation']; Created = $u['createdDateTime']
            Licenses = $skus -join ', '; ExchangePlans = $exoPlans -join ', '; ExchangePlanEnabled = $exchange.Enabled
            MailboxInExchangeOnline = $hasMailbox; MailboxProbe = "HTTP $($probe.HttpStatus) $($probe.ErrorCode)".Trim(); MailboxType = $type
            Autodiscover = if ($m) { Get-GctLocationText $m } else { $null }; ProxyAddresses = (@($u['proxyAddresses']) -join ', ')
        }
        $Context.Provisioning.Add([pscustomobject]$d)
        $who = "$($u['displayName'])$(if ($u['onPremisesSyncEnabled']) { ', synchronized from Active Directory' })"
        $typeText = if ($type) { " ($type mailbox)" } else { '' }
        if ($hasMailbox -eq $true) {
            $warn = if (-not $exchange.Enabled -and $type -notin 'shared', 'room', 'equipment') { ' No Exchange plan is enabled for this user: the mailbox will be disabled after the grace period.' } else { '' }
            Add-GctStep $Context $stage "Provisioning of $address" $(if ($warn) { 'Warning' } else { 'Passed' }) "$who has a mailbox in Exchange Online$typeText$(if ($skus) { "; licenses $($skus -join ', ')" }).$warn" $d
        }
        elseif ($hasMailbox -eq $false) {
            if ($m -and $m.Location -eq 'OnPremises') { Add-GctStep $Context $stage "Provisioning of $address" Passed "$who has no mailbox in Exchange Online: its mailbox is on-premises ($($m.EwsHost)), Exchange Online knows it as a mail user." $d }
            else {
                $why = if (-not $exchange.Enabled) { 'no Exchange license is assigned (or the plan is disabled)' } else { 'the Exchange plan is enabled but the mailbox is not (yet) provisioned - wait, or check the provisioning errors of the user (Get-User / Get-Recipient in Exchange Online)' }
                Add-GctStep $Context $stage "Provisioning of $address" Failed "$who has no mailbox in Exchange Online and none was found on-premises: $why." $d
            }
        }
        else { Add-GctStep $Context $stage "Provisioning of $address" Warning "$who exists; whether it has a mailbox could not be told (HTTP $($probe.HttpStatus) $($probe.ErrorCode))." $d }
    }
}

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBeiyOZYzz6Ne7H
# u5+2hfB+NrTUDg0/PJbwTWVg/81In6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCqOdAuV3VO
# lVu40vhxbLW/3VKSCUlUr4DyC67WeSglgDANBgkqhkiG9w0BAQEFAASCAgBTkzdK
# fx/aMCsAxN8DA5hR/tHtZX3XNe2vlD667H+NoWApLTndMTdbIjkwFsKHY5VfDFIR
# a6ZDNrN1itQ96Si3eVlY/D+sxfKOk5NwrTV4HnAV4Z4ZMPHT09kb6uqTI2rsMG8N
# J3PxBTBSPO/gq9Cy4Udlg4DeD7vZIU0RdIWQ6wt7LWPAqivRchcBKrYQxuEGXyLd
# mAYySLCe8ydMs+dvoJ3HXYKliWnkbJfxp1TA8gK1m9I/N9t+jZeHY6MW8XjENtDa
# 2W67JNQX6yGyMzbIZ/nzMuHvkc9eflesaw0WIH22NhAT4uBMrR2YKvJkokVvs5Lb
# 2STjPov0XNECavwVIqtP9wzuMy3EgkmL9J1jJjmp++f5eUNzE1hJgOGb7pT4UptU
# YNEVEG32Us5A4JNcx/NteRfhWeBDOY4sG0BxALZY+NeAtq99hYOnZ69cEB9TxuJH
# ymDGAnuKSroWE/KxrxWfeNxD5dR2C52VIvBJqkyIkgeNccJekGt88bJFrXjOayrs
# Xyb94UTYT+MZdxGHjeHYFwIt+Zxtd3a+aCz+vtGZyHOSr7WO3Jtzl/Qli9tCYLpj
# kgDMXpQ8b7fYihyarLRm2HB+0CqVrQ0rQ9ezJkWmPnLJ3yaOr+AmuVhKecdkEI5o
# yHsb1akA6rzPi4qq9AtfG8Jbo7WgHD8dObq4VaGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEICHH6zbK8dSQpiSmta+aSD61ZKyhSU3tO2x4TWwuZ5FiAggIAXLXkt+k
# qhgPMjAyNjExMTAxMDIzNTlaMAMCAQECEDMnk/BtVlRb6cjOD4l1aiqgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjM1OVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIA1k
# 56xCkm/Z+pRJ9JUkZjurpFvhcp8rOK7K6l7hjrvRMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiAdDEqY7Qg4zHfCB4UQkE+1j6OTtXqcLns8HzXGv6/kMgIhAP4y
# K88oJoYW+CbMJQbIqrWKVEIy7gYUVJxfksKkCGoz
# SIG # End signature block
