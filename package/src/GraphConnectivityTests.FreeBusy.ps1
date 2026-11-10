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
