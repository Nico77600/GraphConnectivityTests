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
