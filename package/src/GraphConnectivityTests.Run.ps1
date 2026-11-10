<#
.SYNOPSIS
    Graph Connectivity Tests - steps and the orchestration of a run (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Invoke-GctConnectivityTest runs the stages of a test in order with one shared context (tokens, HTTP
    client, where each mailbox is, free/busy, records...). Every check adds one step with a status:
    Passed, Warning, Blocked, Failed or Skipped. The discovery of the source, the sign-in and the source
    mailbox stop the test when they fail (the next stages are Skipped); the other stages fail alone.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>

function Add-GctStep {
    <# Records one check: status, message, details and the time spent since the previous step. #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Context,
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][ValidateSet('Passed', 'Warning', 'Blocked', 'Failed', 'Skipped')][string]$Status,
        [Parameter(Mandatory = $true)][string]$Message,
        [System.Collections.IDictionary]$Details
    )

    # Messages are built from parts that often end with a period: one period only (an ellipsis is kept).
    $Message = [regex]::Replace($Message, '(?<!\.)\.\.(?!\.)', '.')
    $elapsed = $Context.Clock.Elapsed.TotalMilliseconds
    $Context.Clock.Restart()
    $copy = [ordered]@{}
    if ($Details) { foreach ($key in $Details.Keys) { if ($null -ne $Details[$key] -and [string]$Details[$key] -ne '') { $copy[$key] = $Details[$key] } } }
    $trace = Complete-GctTraceStep -Step ($Context.Steps.Count + 1) -Name $Name
    $Context.Steps.Add([pscustomobject]@{
            Stage        = $Stage
            Name         = $Name
            Status       = $Status
            Message      = $Message
            Details      = $copy
            Trace        = @($trace)
            DurationMs   = [int][Math]::Round($elapsed)
            TimestampUtc = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        })
    $item = @{ Passed = 'Ok'; Warning = 'Warn'; Blocked = 'Block'; Failed = 'Fail'; Skipped = 'Skip' }[$Status]
    # The window receives the check itself (its details and HTTP exchanges for the diagnostics), not a line of text.
    $toWindow = $script:Ui -and $script:Ui['StepAdded']
    if ($toWindow) { & $script:Ui['StepAdded'] $Context.Steps[$Context.Steps.Count - 1] @($Context.Trace | Where-Object { $_.Sequence -in @($trace) }) }
    Write-GctItem $item ('{0}: {1}' -f $Name, $Message) -NoUi:([bool]$toWindow)
}

function Invoke-GctConnectivityTest {
    <#
    .SYNOPSIS
        Runs one test and returns the result (steps, mailboxes, free/busy, records, actions, trace).
    .PARAMETER Configuration
        Settings hashtable (Import-GctConfiguration). Missing keys take their default value.
    .PARAMETER TestType
        FreeBusy, MailboxServices, Provisioning, DomainDns, MessageHeaders or All. Default: the TestType of the configuration.
    .PARAMETER AccessToken
        A Microsoft Graph token already obtained (integration only): the sign-in is skipped. Never written anywhere.
    .PARAMETER OnPremisesCredential
        On-premises authentication Credential: the account EWS on-premises is called with (NTLM, Negotiate or Basic).
        Without it, the account is asked (Get-Credential) when a source mailbox on-premises needs it. Kept in memory only.
    .PARAMETER Quiet
        No console output (log and window still receive the lines).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [string]$TestType, [string]$AccessToken, [pscredential]$OnPremisesCredential, [switch]$Quiet)

    $cfg = Get-GctDefaultConfiguration
    foreach ($key in $Configuration.Keys) { $cfg[$key] = $Configuration[$key] }
    if ($TestType) { $cfg.TestType = $TestType }
    $cfg.TargetMailboxes = @($cfg.TargetMailboxes | Where-Object { $_ } | ForEach-Object { ([string]$_).Trim() } | Select-Object -Unique)
    $validation = Test-GctConfiguration -Configuration $cfg
    if (-not $validation.IsValid) { throw ("Invalid configuration:`n - " + ($validation.Problems -join "`n - ")) }
    $stages = @(Get-GctScenarioStages -TestType $cfg.TestType -Configuration $cfg)
    $scenario = $script:Scenarios | Where-Object Name -eq $cfg.TestType

    $previous = @{ Quiet = $script:Quiet; Trace = $script:GctTrace; Cookies = $script:GctCookies; DnsServer = $script:GctDnsServer }
    $script:Quiet = [bool]$Quiet
    $script:GctDnsServer = @($cfg.DnsServer | Where-Object { $_ } | ForEach-Object { ([string]$_).Trim() })
    $started = [DateTimeOffset]::UtcNow
    $context = @{
        Config         = $cfg
        Cloud          = $script:Clouds[[string]$cfg.Cloud]
        Client         = Resolve-GctClient -ClientId ([string]$cfg.ClientId) -RedirectUri ([string]$cfg.RedirectUri)
        Tokens         = @{}
        SourceTenantId = if ($cfg.TenantId -match '^[0-9a-fA-F-]{36}$') { [string]$cfg.TenantId } else { $null }
        Tenants        = @{}
        Mailboxes      = @{}
        Source         = $null
        SourceUser     = $null
        SourceProtocol = $null
        SourceEwsUrl   = $null
        SourceSchedule = $null
        TargetUsers    = @{}
        Token          = $null
        SignIn         = $null
        SignedInUser   = $null
        FreeBusy       = @()
        FreeBusyWindow = $null
        OwnMailTips    = $null
        Hybrid         = $null
        Provisioning   = [Collections.Generic.List[object]]::new()
        DnsRecords     = [Collections.Generic.List[object]]::new()
        Domain         = $null
        Actions        = [Collections.Generic.List[object]]::new()
        Steps          = [Collections.Generic.List[object]]::new()
        Trace          = [Collections.Generic.List[object]]::new()
        Clock          = [Diagnostics.Stopwatch]::StartNew()
        Throttled      = 0
        Stop           = $false
        HttpClient     = $null
        OnPremisesCredential = $OnPremisesCredential
        OnPremisesSchemes    = @{}
        Headers        = $null
        HeaderRoute    = @()
    }
    if ($AccessToken) { $context.Tokens.Graph = [pscustomobject]@{ AccessToken = $AccessToken; RefreshToken = $null; Flow = 'supplied'; Claims = Get-GctTokenClaims -AccessToken $AccessToken; Scope = $null; Resource = $null } }
    $handler = [Net.Http.HttpClientHandler]::new()
    $handler.AllowAutoRedirect = $false
    $handler.UseDefaultCredentials = $false
    $handler.CookieContainer = [Net.CookieContainer]::new()
    $handler.AutomaticDecompression = [Net.DecompressionMethods]::GZip -bor [Net.DecompressionMethods]::Deflate
    $script:GctTrace = $context.Trace
    $script:GctCookies = $handler.CookieContainer
    $context.HttpClient = [Net.Http.HttpClient]::new($handler)
    $context.HttpClient.Timeout = [TimeSpan]::FromSeconds([int]$cfg.HttpTimeoutSeconds)

    # A failure of these stages stops the test; the others fail alone.
    $blocking = @('Discovery', 'SignIn', 'Source', 'HeaderRead')
    try {
        for ($i = 0; $i -lt $stages.Count; $i++) {
            $stage = $stages[$i]
            $title = Get-GctStageTitle -Stage $stage -Configuration $cfg
            $script:GctTraceStage = $stage
            Write-GctStep ($i + 1) $stages.Count $title -Icon $script:StageInfo[$stage].Icon
            if ($context.Stop) { Add-GctStep $context $stage $title Skipped 'Not run: an earlier step failed.'; continue }
            $before = $context.Steps.Count
            try {
                switch ($stage) {
                    'Discovery' { Invoke-GctStageDiscovery -Context $context }
                    'OnPremises' { Invoke-GctStageOnPremises -Context $context }
                    'Dns' { Invoke-GctStageDns -Context $context }
                    'SignIn' { Invoke-GctStageSignIn -Context $context }
                    'Source' { Invoke-GctStageSource -Context $context }
                    'Target' { Invoke-GctStageTarget -Context $context }
                    'FreeBusy' { Invoke-GctStageFreeBusy -Context $context }
                    'Hybrid' { Invoke-GctStageHybrid -Context $context }
                    'Sync' { Invoke-GctStageSync -Context $context }
                    'Availability' { Invoke-GctStageAvailability -Context $context }
                    'AutoReply' { Invoke-GctStageAutoReply -Context $context }
                    'Provisioning' { Invoke-GctStageProvisioning -Context $context }
                    'DomainRecords' { Invoke-GctStageDomainRecords -Context $context }
                    'HeaderRead' { Invoke-GctStageHeaderRead -Context $context }
                    'HeaderDns' { Invoke-GctStageHeaderDns -Context $context }
                    'HeaderRoute' { Invoke-GctStageHeaderRoute -Context $context }
                }
            }
            catch {
                Add-GctStep $context $stage $title Failed $_.Exception.Message
                if ($stage -in $blocking -or $_.Exception.Message -eq 'Cancelled by the operator.') { $context.Stop = $true }
            }
            if ($context.Steps.Count -eq $before) { Add-GctStep $context $stage $title Passed 'Nothing to check.' }
        }
    }
    finally {
        $context.HttpClient.Dispose()
        $handler.Dispose()
        $script:Quiet = $previous.Quiet
        $script:GctTrace = $previous.Trace
        $script:GctCookies = $previous.Cookies
        $script:GctDnsServer = $previous.DnsServer
        $script:GctTraceStage = $null
    }

    $steps = @($context.Steps)
    $counts = [ordered]@{}
    foreach ($s in 'Passed', 'Warning', 'Blocked', 'Failed', 'Skipped') { $counts[$s] = @($steps | Where-Object Status -eq $s).Count }
    $firstProblem = $steps | Where-Object { $_.Status -eq 'Failed' } | Select-Object -First 1
    if (-not $firstProblem) { $firstProblem = $steps | Where-Object { $_.Status -eq 'Blocked' } | Select-Object -First 1 }
    $completed = [DateTimeOffset]::UtcNow
    $titles = [ordered]@{}
    foreach ($s in $stages) { $titles[$s] = Get-GctStageTitle -Stage $s -Configuration $cfg }
    [pscustomobject]@{
        Tool            = 'Graph Connectivity Tests'
        Version         = $script:ToolVersion
        Status          = Get-GctOverallStatus -Steps $steps
        TestType        = $cfg.TestType
        Scenario        = $scenario.DisplayName
        StartedUtc      = $started.ToString('yyyy-MM-ddTHH:mm:ssZ')
        CompletedUtc    = $completed.ToString('yyyy-MM-ddTHH:mm:ssZ')
        DurationSeconds = [Math]::Round(($completed - $started).TotalSeconds, 1)
        Cloud           = $context.Cloud.Name
        SourceMailbox   = [string]$cfg.SourceMailbox
        TargetMailboxes = @($cfg.TargetMailboxes)
        Context         = [string]$cfg.Context
        ClientId        = $context.Client.Id
        ClientName      = $context.Client.Name
        SignIn          = $context.SignIn
        SignedInUser    = if ($context.Token) { $context.Token.User } else { $null }
        TenantId        = $context.SourceTenantId
        Token           = $context.Token
        SourceProtocol  = $context.SourceProtocol
        SourceEwsUrl    = $context.SourceEwsUrl
        OnPremisesHost  = [string]$cfg.OnPremisesHost
        Domain          = $context.Domain
        AllowWrite      = [bool]$cfg.AllowWrite
        AdminChecks     = [bool]$cfg.AdminChecks
        StageTitles     = $titles
        Mailboxes       = @($context.Mailboxes.Values | Sort-Object { if ($_.Role -eq 'Source') { 0 } else { 1 } }, Address | Select-Object Address, Role, Domain, TenantId, Consumer, TenantRegion, Realm, Location, EwsUrl, EwsHost, Relation, Autodiscover)
        FreeBusy        = @($context.FreeBusy | ForEach-Object { $_ | Select-Object Mailbox, Protocol, Path, Status, ResponseCode, Message, Hint, View, IntervalMinutes, StartUtc, DetailLevel, WorkingHours, WorkDays, WorkStart, WorkEnd, WorkTimeZone, WorkOffsetMinutes, @{ n = 'Items'; e = { @($_.Events).Count } } })
        FreeBusyWindow  = $context.FreeBusyWindow
        FreeBusyEvents  = @($context.FreeBusy | ForEach-Object { @($_.Events) })
        Hybrid          = $context.Hybrid
        Provisioning    = @($context.Provisioning)
        DnsRecords      = @($context.DnsRecords)
        Actions         = @($context.Actions)
        Headers         = $context.Headers
        HeaderSummary   = Get-GctHeaderSummary -Analysis $context.Headers
        Route           = @($context.HeaderRoute)
        Throttled       = $context.Throttled
        Counts          = $counts
        Steps           = $steps
        Trace           = @($context.Trace | Select-Object -Property * -ExcludeProperty Collapse, ResponseKey)
        Error           = if ($firstProblem) { "$($firstProblem.Name): $($firstProblem.Message)" } else { $null }
    }
}
