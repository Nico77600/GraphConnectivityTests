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

# SIG # Begin signature block
# MIImKwYJKoZIhvcNAQcCoIImHDCCJhgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCp8FHWE3vixiTo
# 7WobEmrZej25pcBPhcTNSzt1PxUvL6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAlECpIvO5E
# c2mb/bkrc1pBsnwC3wkP/vBbqWAIvUGWtjANBgkqhkiG9w0BAQEFAASCAgBUxLsK
# RPGky1t/EkznO6PDuX56aUnSya1oUEjBV0vnPOwFCbKX4UwBjQGjXdWW5m+U6fQ3
# LTT+LdDMoiQ5GMPnhQngtNHXr/u6NLjFmew9Jy0W7pqn/r9C4hiW0g1W06FMn78o
# W+wAGhOGG0Igp/3PS6WFKwbBIB+wfBvsOIYjg80fCIemF/9h0MtP4wt049Cp5GaZ
# 8KRu2IhWOVAEYPwydEEOSYqDezFV5itsArNRjRCwIZRWUcAqqo+FzyXW6Rp2ssTy
# hydh6JzZbfArCIWsfn1LhJ2QgY9wHWhOsgvCfbwETZA103vvFjbrXPjJsvkZ3xz5
# hitZRF2/1JAyccpBOOquZ7eHp/qk6KHxqiL8jmN8ibl4m3jBwoegT9DFfjBB1K9R
# j2ckkYTUup69KHpV8EJ7/SN19/VZqHe9FxXwn4fY2atBv4cN7wuRN9lpk0qaT5ON
# hVGJ0BKdFGMNUjlungwgFcK4grPlNnOrqLgxFOozSxJGgkwS4iXb5is4iHs7uajh
# VZPxbQB0rVhAZjJl0/a6tfVqZJltRSKMeSsX0+OXXSHRGu4mPOReaPuE+AEV4nFH
# y4a3GtGB7eUkoqP0CjUMttetMeeGRhyZGNuUi4q67XDtqBJTFE626OiT5PkvAy8L
# XuT/88vzbqCLCs3tphfEqhHgY7lAftEvNDSvjqGCFM8wghTLBgorBgEEAYI3AwMB
# MYIUuzCCFLcGCSqGSIb3DQEHAqCCFKgwghSkAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIG0yeWuvWa+iWnHakfZOWfg5XABDa3Z1zGJ+1bVX407JAgg9BjPyQnnx
# cxgPMjAyNjExMTAxODI3MDJaMAMCAQECEDvW6ve51F3mxzP+22y2X8mgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDE4MjcwMlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEINzS
# Xy8K6fszzEVROMTjEg1KtRvvnhgf6wo41NrNuharMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIESDBGAiEAtkMY3fMpVpHqsvbrRMSVx9y+xn9pAUXLW7uRpIwUC6gCIQDw
# jg7X5oNPh7RGp87nMU5qtB4bIUkLbIKZHWBJ25EG/Q==
# SIG # End signature block
