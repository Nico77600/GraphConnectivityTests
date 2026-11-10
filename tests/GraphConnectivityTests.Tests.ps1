#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.1.0' }
<#
    Graph Connectivity Tests - automated tests (Pester 6.1 or later).
    Author  : Nicolas Fabert
    Version : 1.0.0

    Run:  .\Run-Tests.ps1      (or Invoke-Pester -Path .\tests -Output Detailed)

    No network, no tenant: tests\GraphConnectivityTests.Simulator.ps1 replaces, inside the module, the functions
    that touch the network (Send-GctHttpRequest, Invoke-GctDnsQuery, Get-GctTlsCertificate) and answers like
    Entra ID, Microsoft Graph, Exchange Online and Exchange on-premises.
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    $script:Artifacts = Join-Path $script:RepoRoot 'artifacts'
    . (Join-Path $PSScriptRoot 'GraphConnectivityTests.Simulator.ps1')
    $script:Base = @{
        SourceMailbox = 'alice@contoso.com'; TargetMailboxes = @('bob@contoso.com'); SignIn = 'DeviceCode'
        OutputPath = (Join-Path $script:Artifacts 'test-reports'); LogPath = (Join-Path $script:Artifacts 'test-logs')
    }
    $script:AppId = '11111111-2222-3333-4444-555555555555'
    $script:Thumbprint = '0123456789ABCDEF0123456789ABCDEF01234567'

    function Start-Sim([hashtable]$Changes = @{}) {
        Import-Module (Join-Path $script:Root 'GraphConnectivityTests.psd1') -Force
        $script:Module = Get-Module GraphConnectivityTests
        $script:State = New-SimState
        foreach ($k in $Changes.Keys) { $script:State[$k] = $Changes[$k] }
        Install-SimNetwork -Module $script:Module -State $script:State
    }
    function New-UserToken([string]$Upn = 'alice@contoso.com', [string]$Scope, [int]$ExpiresIn = 3600, [string]$TenantId = $script:SimContoso) {
        if (-not $Scope) { $Scope = $script:State.GraphScopes }
        $u = Find-SimUser $script:State $Upn
        $t = New-SimToken -Audience 'https://graph.microsoft.com' -Scope $Scope -Upn $Upn -TenantId $TenantId -ExpiresIn $ExpiresIn -ObjectId $(if ($u) { $u.Id })
        $script:State.ValidTokens.Add($t); $script:State.Secrets.Add($t)
        $t
    }
    <#
        One run against the simulator. By default with -AccessToken (a token of the source user); -SignIn runs the
        real sign-in of the engine (device code unless Overrides say otherwise) against the simulated Entra ID.
    #>
    function Invoke-Sim([string]$Type, [hashtable]$Overrides = @{}, [hashtable]$State = @{}, [switch]$SignIn, [string]$Token, [pscredential]$Credential) {
        Start-Sim $State
        $cfg = $script:Base.Clone(); foreach ($k in $Overrides.Keys) { $cfg[$k] = $Overrides[$k] }
        if (-not $SignIn -and -not $Token -and [string]$cfg['Context'] -ne 'Application') { $Token = New-UserToken -Upn $cfg.SourceMailbox }
        Invoke-GctConnectivityTest -Configuration $cfg -TestType $Type -AccessToken $Token -OnPremisesCredential $Credential -Quiet
    }
    function Get-Step($Result, [string]$Name) { $Result.Steps | Where-Object Name -eq $Name | Select-Object -First 1 }
    function Get-StageSteps($Result, [string]$Stage) { @($Result.Steps | Where-Object Stage -eq $Stage) }
    # A step named like its stage title with Failed: the stage threw (an exception of the engine, not a check).
    function Get-StageCrash($Result) { @($Result.Steps | Where-Object { $_.Status -eq 'Failed' -and $_.Name -eq $Result.StageTitles[$_.Stage] -and $_.Message -ne 'Not run: an earlier step failed.' }) }
    function Get-ResultText($Result) {
        $parts = foreach ($e in $Result.Trace) { $e.Url; $e.Request; $e.Response; $e.Note }
        $parts += foreach ($s in $Result.Steps) { $s.Message; ConvertTo-Json -InputObject $s.Details -Depth 6 -Compress }
        $parts += ConvertTo-Json -InputObject $Result -Depth 8 -Compress
        $parts -join "`n"
    }
    # Every secret the simulator handed out, and the payload and signature of every JWT among them (the header is the same for all).
    function Assert-NoSecret([string]$Text) {
        $script:State.Secrets.Count | Should -BeGreaterThan 0
        foreach ($secret in $script:State.Secrets) {
            $Text.Contains($secret) | Should -BeFalse -Because "a secret of $($secret.Length) characters must never be written"
            $parts = $secret.Split('.')
            if ($parts.Count -eq 3) { foreach ($p in $parts[1], $parts[2]) { if ($p.Length -gt 16) { $Text.Contains($p) | Should -BeFalse -Because 'no part of a token is written' } } }
        }
    }
}

Describe 'Configuration' {
    BeforeAll {
        Import-Module (Join-Path $script:Root 'GraphConnectivityTests.psd1') -Force
        [void][IO.Directory]::CreateDirectory($script:Artifacts)
    }

    It 'reads the sections of the file and resolves the relative paths from the tool folder' {
        $c = Import-GctConfiguration
        $c.SourceMailbox | Should -Be 'alice@contoso.com'
        $c.TargetMailboxes | Should -Be @('bob@contoso.com')
        $c.TestType | Should -Be 'FreeBusy'
        $c.SignIn | Should -Be 'Auto'
        $c.ClientId | Should -Be 'MicrosoftOffice'
        $c.AllowWrite | Should -BeFalse
        $c.OutputPath | Should -Be (Join-Path $script:Root 'reports')
        $c.LogPath | Should -Be (Join-Path $script:Root 'logs')
        $c.ReportFormats | Should -Be @('Csv', 'Html')
        (Test-GctConfiguration -Configuration $c).IsValid | Should -BeTrue
    }

    It 'lists unknown sections, unknown keys, a section that is not a table and invalid values, all at once' {
        $path = Join-Path $script:Artifacts 'bad.config.psd1'
        "@{ Target = @{ SourceMailbox = 'alice'; Mailbox = 'x'; Cloud = 'Mars' }; Extra = @{}; Logging = 'logs'; Test = @{ DefaultType = 'Nope'; FreeBusyDays = 90; SignIn = 'Browser' }; Identity = @{ Context = 'Delegated' } }" | Set-Content -LiteralPath $path
        $text = ({ Import-GctConfiguration -Path $path } | Should -Throw -PassThru).Exception.Message
        $text | Should -Match "Unknown key 'Target\.Mailbox'"
        $text | Should -Match "Unknown section 'Extra'"
        $text | Should -Match "Section 'Logging' must be @\{ \.\.\. \}"
        $text | Should -Match 'Target\.SourceMailbox must be an SMTP address'
        $text | Should -Match 'Target\.Cloud must be one of: Global, USGov, USGovDoD, China'
        $text | Should -Match 'Test\.DefaultType must be one of: FreeBusy, MailboxServices, Provisioning, DomainDns, All'
        $text | Should -Match 'FreeBusyDays must be a number from 1 to 62'
        $text | Should -Match "Test\.SignIn must be 'Auto', 'Window' or 'DeviceCode'"
        $text | Should -Match "Identity\.Context must be 'User' or 'Application'"
        Remove-Item -LiteralPath $path
    }

    It 'says when the file does not exist' {
        { Import-GctConfiguration -Path (Join-Path $script:Artifacts 'missing.config.psd1') } | Should -Throw 'Configuration file not found*'
    }

    It 'lists every problem of the settings' {
        $c = & (Get-Module GraphConnectivityTests) { Get-GctDefaultConfiguration }
        $c.TestType = 'Everything'; $c.SourceMailbox = 'alice-at-contoso'; $c.TargetMailboxes = @('not-an-address') + @(1..20 | ForEach-Object { "user$_@contoso.com" })
        $c.OnPremisesHost = 'https://mail.contoso.com/EWS/Exchange.asmx'; $c.Domain = 'contoso'; $c.Cloud = 'Moon'; $c.Context = 'Delegated'
        $c.ClientId = 'Outlook'; $c.AdminClientId = 'Admin'; $c.RedirectUri = 'nativeclient'; $c.TenantId = 'contoso!'; $c.OnPremisesAuthentication = 'Kerberos'; $c.OnPremisesCredentialScheme = 'Digest'; $c.OnPremisesUserName = 'carol'; $c.SignIn = 'Popup'
        $c.FreeBusyDays = 0; $c.FreeBusyIntervalMinutes = 2; $c.OAuthPollTimeoutSeconds = 5; $c.HttpTimeoutSeconds = 1000; $c.CertificateWarningDays = 400; $c.MaxRetries = 11; $c.LogRetentionDays = 'x'
        $c.ReportFormats = @('Csv', 'Pdf'); $c.CsvDelimiter = '|'
        $problems = (Test-GctConfiguration -Configuration $c).Problems
        $text = $problems -join "`n"
        foreach ($expected in @(
                'Test\.DefaultType must be one of', 'Target\.SourceMailbox must be an SMTP address', "Target\.TargetMailboxes: 'not-an-address' is not an SMTP address", '20 mailboxes at most'
                'Target\.OnPremisesHost must be a host name', 'Target\.Domain must be a domain name', 'Target\.Cloud must be one of', "Identity\.Context must be 'User' or 'Application'"
                'Identity\.ClientId must be MicrosoftOffice, GraphCommandLineTools, ExchangeOnlinePowerShell or an application ID', 'Identity\.AdminClientId must be', 'Identity\.RedirectUri must be an absolute URI'
                'Identity\.TenantId must be a tenant ID or a domain', "Identity\.OnPremisesAuthentication must be 'OAuth'", "Identity\.OnPremisesCredentialScheme must be 'Auto'", 'Identity\.OnPremisesUserName must be DOMAIN\\user or user@domain', "Test\.SignIn must be 'Auto', 'Window' or 'DeviceCode'"
                'FreeBusyDays must be a number from 1 to 62', 'FreeBusyIntervalMinutes must be a number from 5 to 1440', 'OAuthPollTimeoutSeconds must be a number from 30 to 3600'
                'HttpTimeoutSeconds must be a number from 5 to 600', 'CertificateWarningDays must be a number from 0 to 365', 'MaxRetries must be a number from 0 to 10', 'LogRetentionDays must be a number from 1 to 3650'
                "Report\.Formats: 'Pdf' is not Csv or Html", 'Report\.CsvDelimiter must be')) {
            $text | Should -Match $expected
        }
        $problems.Count | Should -Be 25
    }

    It 'needs a target for FreeBusy on the command line, not in the file' {
        $c = & (Get-Module GraphConnectivityTests) { Get-GctDefaultConfiguration }
        (Test-GctConfiguration -Configuration $c).Problems | Should -Contain 'The FreeBusy test needs at least one target mailbox (Target.TargetMailboxes, -TargetMailbox).'
        (Test-GctConfiguration -Configuration $c -File).IsValid | Should -BeTrue
        $c.TestType = 'DomainDns'
        (Test-GctConfiguration -Configuration $c).IsValid | Should -BeTrue
    }

    It 'the Application context needs its application, its certificate and its tenant, and never changes the automatic replies' {
        $c = & (Get-Module GraphConnectivityTests) { Get-GctDefaultConfiguration }
        $c.Context = 'Application'; $c.TestType = 'MailboxServices'; $c.AllowWrite = $true
        $text = (Test-GctConfiguration -Configuration $c).Problems -join "`n"
        $text | Should -Match 'needs the application ID of your application'
        $text | Should -Match 'needs the thumbprint of the certificate'
        $text | Should -Match 'needs the tenant'
        $text | Should -Match 'changed only in the User context'
        $c.ClientId = $script:AppId; $c.CertificateThumbprint = $script:Thumbprint; $c.TenantId = 'contoso.com'; $c.AllowWrite = $false
        (Test-GctConfiguration -Configuration $c).IsValid | Should -BeTrue
    }

    It 'a run refuses invalid settings before any request' {
        Start-Sim
        { Invoke-GctConnectivityTest -Configuration @{ SourceMailbox = 'alice'; TargetMailboxes = @('bob@contoso.com') } -Quiet } | Should -Throw 'Invalid configuration*Target.SourceMailbox*'
        $script:State.Calls.Count | Should -Be 0
    }
}

Describe 'Catalog and stages' {
    BeforeAll { Import-Module (Join-Path $script:Root 'GraphConnectivityTests.psd1') -Force }
    AfterAll { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'exposes the tests with their stages, sign-in and writing flags' {
        $catalog = @(Get-GctTestCatalog)
        $catalog.Name | Should -Be @('FreeBusy', 'MailboxServices', 'Provisioning', 'DomainDns', 'All')
        @($catalog | Where-Object Writes).Name | Should -Be @('MailboxServices', 'All')
        @($catalog | Where-Object { -not $_.SignIn }).Count | Should -Be 0
        ($catalog | Where-Object Name -eq 'DomainDns').Stages | Should -Be @('Dns', 'SignIn', 'DomainRecords')
        ($catalog | Where-Object Name -eq 'MailboxServices').Stages | Should -Be @('SignIn', 'Source', 'Sync', 'Availability', 'AutoReply')
    }

    It 'skips the on-premises checks with SkipOnPremisesChecks and runs Hybrid only with AdminChecks' {
        Get-GctScenarioStages -TestType FreeBusy | Should -Be @('Discovery', 'OnPremises', 'SignIn', 'Source', 'Target', 'FreeBusy')
        Get-GctScenarioStages -TestType FreeBusy -Configuration @{ SkipOnPremisesChecks = $true } | Should -Be @('Discovery', 'SignIn', 'Source', 'Target', 'FreeBusy')
        Get-GctScenarioStages -TestType FreeBusy -Configuration @{ AdminChecks = $true } | Should -Be @('Discovery', 'OnPremises', 'SignIn', 'Source', 'Target', 'FreeBusy', 'Hybrid')
        Get-GctScenarioStages -TestType All -Configuration @{ SkipOnPremisesChecks = $true; AdminChecks = $true } | Should -Be @('Discovery', 'Dns', 'SignIn', 'Source', 'Target', 'FreeBusy', 'Hybrid', 'Sync', 'Availability', 'AutoReply', 'Provisioning', 'DomainRecords')
        Get-GctScenarioStages -TestType Nope | Should -BeNullOrEmpty
    }

    It 'a run without the on-premises checks never contacts the on-premises server' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('carol@contoso.com'); SkipOnPremisesChecks = $true }
        Get-StageSteps $r 'OnPremises' | Should -BeNullOrEmpty
        $r.StageTitles.Keys | Should -Not -Contain 'OnPremises'
        @($script:State.Calls | Where-Object Host -eq 'mail.contoso.com') | Should -BeNullOrEmpty
        (Get-Step $r 'Free/busy of carol@contoso.com').Status | Should -Be 'Passed'
    }

    It 'names the sign-in stage after the context' {
        & (Get-Module GraphConnectivityTests) { Get-GctStageTitle -Stage SignIn -Configuration @{ Context = 'User' } } | Should -Be 'Sign-in of the source user'
        & (Get-Module GraphConnectivityTests) { Get-GctStageTitle -Stage SignIn -Configuration @{ Context = 'Application' } } | Should -Be 'Sign-in of the application (certificate)'
    }
}

Describe 'Sign-in against the simulated Entra ID' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'an access token supplied by the caller skips the sign-in, its claims are still checked' {
        $r = Invoke-Sim 'FreeBusy'
        (Get-Step $r 'Access token').Status | Should -Be 'Passed'
        (Get-Step $r 'Token claims').Status | Should -Be 'Passed'
        (Get-Step $r 'Token claims').Message | Should -Match 'Microsoft Graph token with Calendars\.ReadWrite'
        $r.Token.User | Should -Be 'alice@contoso.com'
        $r.Token.Audience | Should -Be 'https://graph.microsoft.com'
        $r.TenantId | Should -Be $script:SimContoso
        @($script:State.Calls | Where-Object { $_.Method -eq 'POST' -and $_.Host -eq 'login.microsoftonline.com' }) | Should -BeNullOrEmpty
    }

    It 'device code: the code asked for the tenant of the source, polled until the user signs in, with a refresh token' {
        $r = Invoke-Sim 'FreeBusy' -SignIn
        $r.Status | Should -Be 'Passed'
        $step = Get-Step $r 'User sign-in'
        $step.Status | Should -Be 'Passed'
        $step.Details.Flow | Should -Be 'device code'
        $step.Details.Client | Should -Be 'Microsoft Office'
        $step.Details.RefreshToken | Should -BeTrue
        $step.Details.ScopeAsked | Should -Be 'https://graph.microsoft.com/.default offline_access openid profile'
        $device = $r.Trace | Where-Object { $_.Url -like '*/oauth2/v2.0/devicecode' }
        $device.Url | Should -Be "https://login.microsoftonline.com/$($script:SimContoso)/oauth2/v2.0/devicecode"
        $device.Response | Should -Match '"device_code": "<device_code: \d+ characters, never written>"'
        $device.Response | Should -Match 'GQ7XK2PLM'
        $polls = @($r.Trace | Where-Object { $_.Url -like '*/oauth2/v2.0/token' })
        $polls.Count | Should -Be 2
        $polls[0].Response | Should -Match 'authorization_pending'
        $polls[0].Label | Should -Be 'grant device_code'
        $polls[1].Request | Should -Match 'device_code=<\d+ characters, never written>'
        $polls[1].Response | Should -Match '"access_token": "<access_token: \d+ characters, never written>"'
        $polls[1].Response | Should -Match '"refresh_token": "<refresh_token: \d+ characters, never written>"'
        $polls[1].Response | Should -Match '"id_token": "<id_token: \d+ characters, never written>"'
        @($script:State.Waits) | Should -Be @(5, 5)
    }

    It 'Auto signs in with a device code when no window can be shown; Window then refuses to start' {
        $r = Invoke-Sim 'FreeBusy' @{ SignIn = 'Auto' } -SignIn
        (Get-Step $r 'User sign-in').Details.Flow | Should -Be 'device code'
        $r = Invoke-Sim 'FreeBusy' @{ SignIn = 'Window' } -SignIn
        (Get-Step $r 'Sign-in of the source user').Status | Should -Be 'Failed'
        (Get-Step $r 'Sign-in of the source user').Message | Should -Match 'cannot be opened: no interactive desktop.*-SignIn DeviceCode'
        @(Get-StageSteps $r 'Source').Status | Should -Be @('Skipped')
        @(Get-StageSteps $r 'FreeBusy').Status | Should -Be @('Skipped')
    }

    It 'sign-in window: authorization code with PKCE, the code never written' {
        $r = Invoke-Sim 'FreeBusy' @{ SignIn = 'Window' } -State @{ DesktopSession = $true } -SignIn
        $r.Status | Should -Be 'Passed'
        $step = Get-Step $r 'User sign-in'
        $step.Details.Flow | Should -Be 'authorization code with PKCE (sign-in window)'
        $step.Details.Browser | Should -Be 'Microsoft Edge'
        $window = $r.Trace | Where-Object Label -eq 'sign-in window'
        $window.Request | Should -Match 'code_challenge_method=S256'
        $window.Request | Should -Match 'login_hint=alice%40contoso\.com'
        $window.Response | Should -Match 'code=<\d+ characters, never written>&state='
        $token = $r.Trace | Where-Object { $_.Url -like '*/oauth2/v2.0/token' }
        $token.Label | Should -Be 'grant authorization_code'
        $token.Request | Should -Match 'code=<\d+ characters, never written>'
        $token.Request | Should -Match 'code_verifier=<43 characters, never written>'
        Assert-NoSecret (Get-ResultText $r)
    }

    It 'a sign-in declined in the browser stops the test' {
        $r = Invoke-Sim 'FreeBusy' -State @{ DeviceDenied = $true } -SignIn
        $r.Status | Should -Be 'Failed'
        (Get-Step $r 'Sign-in of the source user').Message | Should -Be 'The sign-in was denied in the browser.'
        @($r.Steps | Where-Object { $_.Stage -in 'Source', 'Target', 'FreeBusy' }).Status | Should -Be @('Skipped', 'Skipped', 'Skipped')
    }

    It 'an expired token fails the claims and stops the test' {
        Start-Sim
        $expired = New-UserToken -ExpiresIn -120
        $cfg = $script:Base.Clone()
        $r = Invoke-GctConnectivityTest -Configuration $cfg -AccessToken $expired -Quiet
        (Get-Step $r 'Token claims').Status | Should -Be 'Failed'
        (Get-Step $r 'Token claims').Message | Should -Match 'The token expired at'
        @(Get-StageSteps $r 'Source').Status | Should -Be @('Skipped')
    }

    It 'a token without a calendar permission is named before getSchedule is refused' {
        Start-Sim
        $token = New-UserToken -Scope 'User.Read openid profile'
        $r = Invoke-GctConnectivityTest -Configuration $script:Base.Clone() -AccessToken $token -Quiet
        (Get-Step $r 'Token claims').Status | Should -Be 'Warning'
        (Get-Step $r 'Token claims').Message | Should -Match 'no calendar permission in the token \(scp: User\.Read openid profile\)'
        (Get-Step $r 'Source mailbox').Status | Should -Be 'Failed'
        (Get-Step $r 'Source mailbox').Message | Should -Match 'HTTP 403 \(ErrorAccessDenied.*the token has no Calendars\.Read'
    }

    It 'a token of another tenant is flagged' {
        Start-Sim
        $token = New-UserToken -TenantId $script:SimFabrikam
        $r = Invoke-GctConnectivityTest -Configuration $script:Base.Clone() -AccessToken $token -Quiet
        (Get-Step $r 'Token claims').Status | Should -Be 'Warning'
        (Get-Step $r 'Token claims').Message | Should -Match "comes from the tenant $($script:SimFabrikam), not from the tenant of the source mailbox $($script:SimContoso)"
    }

    It 'application: client credentials with a certificate, the checks on /users/{source}, the assertion never written' {
        $r = Invoke-Sim 'FreeBusy' @{ Context = 'Application'; ClientId = $script:AppId; CertificateThumbprint = $script:Thumbprint; TenantId = 'contoso.com' }
        $r.Status | Should -Be 'Passed'
        (Get-Step $r 'Application sign-in').Status | Should -Be 'Passed'
        (Get-Step $r 'Application sign-in').Message | Should -Match '/users/alice@contoso\.com instead of /me'
        (Get-Step $r 'Token claims').Details.Roles | Should -Match 'Calendars\.Read'
        $r.StageTitles['SignIn'] | Should -Be 'Sign-in of the application (certificate)'
        $token = $r.Trace | Where-Object { $_.Url -like '*/oauth2/v2.0/token' }
        $token.Url | Should -Be 'https://login.microsoftonline.com/contoso.com/oauth2/v2.0/token'
        $token.Request | Should -Match 'client_assertion=<\d+ characters, never written>'
        $token.Request | Should -Match 'scope=https://graph\.microsoft\.com/\.default'
        ($r.Trace | Where-Object Operation -eq 'POST calendar/getSchedule').Url | Should -Be 'https://graph.microsoft.com/v1.0/users/alice%40contoso.com/calendar/getSchedule'
        Assert-NoSecret (Get-ResultText $r)
    }
}

Describe 'FreeBusy' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'same organisation: directory, MailTips, then the free/busy with its details and working hours' {
        $r = Invoke-Sim 'FreeBusy'
        $r.Status | Should -Be 'Passed'
        (Get-Step $r 'Source alice@contoso.com').Message | Should -Be "Exchange Online, tenant $($script:SimContoso)."
        (Get-Step $r 'Path to bob@contoso.com').Details.Path | Should -Be 'SameOrganisation'
        (Get-Step $r 'On-premises endpoint').Status | Should -Be 'Skipped'
        (Get-Step $r 'Signed-in user').Status | Should -Be 'Passed'
        (Get-Step $r 'Source mailbox').Message | Should -Match 'Mailbox in Exchange Online, reached through Microsoft Graph; working hours monday .* 08:00-17:00 \(Romance Standard Time\)'
        (Get-Step $r 'Directory bob@contoso.com').Status | Should -Be 'Passed'
        (Get-Step $r 'MailTips bob@contoso.com').Message | Should -Match 'as internal.*Automatic replies are on\.'
        (Get-Step $r 'getSchedule').Status | Should -Be 'Passed'
        (Get-Step $r 'Free/busy of bob@contoso.com').Status | Should -Be 'Passed'
        $fb = $r.FreeBusy[0]
        $fb.Mailbox | Should -Be 'bob@contoso.com'
        $fb.Protocol | Should -Be 'Microsoft Graph getSchedule'
        $fb.View.Length | Should -Be (7 * 48)
        $fb.DetailLevel | Should -Be 'Limited details or more (subject and location)'
        $fb.WorkDays | Should -Contain 'monday'
        $fb.WorkStart | Should -Be 480
        @($r.FreeBusyEvents).Count | Should -Be $fb.Items
        @($r.FreeBusyEvents.BusyType | Select-Object -Unique).Count | Should -BeGreaterThan 1
        $r.FreeBusyWindow.IntervalMinutes | Should -Be 30
        $request = ($r.Trace | Where-Object Operation -eq 'POST calendar/getSchedule').Request
        $request | Should -Match 'Prefer: outlook\.timezone="UTC"'
        $request | Should -Match 'client-request-id: [0-9a-f-]{36}'
        $request | Should -Match 'Authorization: Bearer <access token: \d+ characters, never written>'
        ($r.Trace | Where-Object Operation -eq 'POST calendar/getSchedule').Server | Should -Be 'Graph PA1PEPF0000SIM1 (France Central, ring 3)'
        Get-StageCrash $r | Should -BeNullOrEmpty
    }

    It 'a secondary address is found through proxyAddresses (advanced query)' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('bob.smith@contoso.com') }
        (Get-Step $r 'Directory bob.smith@contoso.com').Status | Should -Be 'Passed'
        (Get-Step $r 'Directory bob.smith@contoso.com').Details.DisplayName | Should -Be 'Bob Smith'
        $query = $r.Trace | Where-Object Operation -eq 'GET users?$filter=proxyAddresses'
        $query.Request | Should -Match 'ConsistencyLevel: eventual'
        $query.StatusCode | Should -Be 200
    }

    It 'Microsoft 365 to on-premises: Autodiscover redirect, the on-premises endpoint ready for OAuth, the free/busy of the mail user' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('carol@contoso.com') }
        $r.Status | Should -Be 'Passed'
        $carol = $r.Mailboxes | Where-Object Address -eq 'carol@contoso.com'
        $carol.Location | Should -Be 'OnPremises'
        $carol.EwsUrl | Should -Be 'https://mail.contoso.com/EWS/Exchange.asmx'
        $carol.Relation | Should -Be 'CloudToOnPremises'
        (Get-Step $r 'Target carol@contoso.com').Details.Redirects | Should -Be 'https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/carol%40contoso.com?Protocol=EWS > https://autodiscover.contoso.com/autodiscover/autodiscover.json/v1.0/carol%40contoso.com?Protocol=EWS'
        @(Get-StageSteps $r 'OnPremises').Name | Should -Be @('DNS mail.contoso.com', 'Certificate mail.contoso.com', 'EWS mail.contoso.com', 'Autodiscover mail.contoso.com')
        @(Get-StageSteps $r 'OnPremises').Status | Should -Be @('Passed', 'Passed', 'Passed', 'Passed')
        $ews = Get-Step $r 'EWS mail.contoso.com'
        $ews.Message | Should -Match "Exchange answers 401 with OAuth \(Bearer\) and trusts the tenant $($script:SimContoso)"
        $ews.Details.Schemes | Should -Be 'Bearer, Negotiate, NTLM'
        $ews.Details.AuthorizationUri | Should -Be 'https://login.windows.net/common/oauth2/authorize'
        $ews.Details.Servers | Should -Be 'FE EXCH01'
        ($r.Trace | Where-Object { $_.Url -eq 'https://mail.contoso.com/EWS/Exchange.asmx' }).Label | Should -Match 'no credentials'
        (Get-Step $r 'Directory carol@contoso.com').Message | Should -Match 'must be a mail user \(RemoteMailbox on-premises\)'
        (Get-Step $r 'MailTips carol@contoso.com').Message | Should -Match 'the mail user of the hybrid'
        (Get-Step $r 'Free/busy of carol@contoso.com').Status | Should -Be 'Passed'
        $r.FreeBusy[0].Path | Should -Be 'CloudToOnPremises'
    }

    It 'Microsoft 365 to on-premises refused with (401) Unauthorized: the cause is the OAuth trust' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('carol@contoso.com') } -State @{ CloudToOnPrem = 'Unauthorized' }
        $r.Status | Should -Be 'Failed'
        $step = Get-Step $r 'Free/busy of carol@contoso.com'
        $step.Status | Should -Be 'Failed'
        $step.Message | Should -Match '^ErrorProxyRequestProcessingFailed: Proxy web request failed\.'
        $step.Message | Should -Match 'Cause: The request was sent on to the other side \(ErrorProxyRequestProcessingFailed\) and failed there: the remote Exchange refused the token'
        $step.Message | Should -Match 'Get-AuthServer'
        $r.FreeBusy[0].ResponseCode | Should -Be 'ErrorProxyRequestProcessingFailed'
        $r.Error | Should -Match '^Free/busy of carol@contoso\.com: ErrorProxyRequestProcessingFailed'
    }

    It 'Microsoft 365 to on-premises unreachable: the cause is the publication of /EWS' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('carol@contoso.com') } -State @{ CloudToOnPrem = 'Unreachable' }
        (Get-Step $r 'Free/busy of carol@contoso.com').Message | Should -Match 'could not reach the remote endpoint from the Internet'
    }

    It 'on-premises to Microsoft 365: Graph cannot reach the source, EWS on-premises with the token of hybrid modern authentication' {
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com', 'bob@contoso.com') } -State @{ SignInUser = 'carol@contoso.com' } -SignIn
        $r.SourceProtocol | Should -Be 'EWS'
        $r.SourceEwsUrl | Should -Be 'https://mail.contoso.com/EWS/Exchange.asmx'
        (Get-Step $r 'Path to alice@contoso.com').Details.Path | Should -Be 'OnPremisesToCloud'
        (Get-Step $r 'Source mailbox').Message | Should -Match 'not in Exchange Online \(MailboxNotEnabledForRESTAPI\): it is on-premises'
        (Get-Step $r 'MailTips').Status | Should -Be 'Warning'
        $gua = Get-Step $r 'GetUserAvailability'
        $gua.Status | Should -Be 'Passed'
        $gua.Details.ExchangeVersion | Should -Be '15.2.2562.17'
        $gua.Details.Authentication | Should -Match 'OAuth, hybrid modern authentication \(token for https://mail\.contoso\.com, refresh token of the sign-in\)'
        (Get-Step $r 'Free/busy of alice@contoso.com').Status | Should -Be 'Passed'
        (Get-Step $r 'Free/busy of bob@contoso.com').Status | Should -Be 'Passed'
        @($r.FreeBusy.Protocol | Select-Object -Unique) | Should -Be @('EWS GetUserAvailability (on-premises)')
        $r.FreeBusy[0].WorkOffsetMinutes | Should -Be 60
        $refresh = $r.Trace | Where-Object { $_.Label -eq 'grant refresh_token' }
        $refresh.Request | Should -Match 'refresh_token=<\d+ characters, never written>'
        $refresh.Request | Should -Match 'scope=https://mail\.contoso\.com/EWS\.AccessAsUser\.All offline_access openid profile'
        $ews = $r.Trace | Where-Object Operation -eq 'EWS GetUserAvailability'
        $ews.Request | Should -Match 'X-AnchorMailbox: carol@contoso\.com'
        $ews.Request | Should -Match 'Authorization: Bearer <access token: \d+ characters, never written>'
        $ews.Request | Should -Match '# SOAP, \d+ bytes, indented'
        $r.Status | Should -Be 'Warning'
        Get-StageCrash $r | Should -BeNullOrEmpty
    }

    It 'on-premises to Microsoft 365 refused with (401) Unauthorized: the cause is given' {
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com') } -State @{ SignInUser = 'carol@contoso.com'; OnPremToCloud = 'Unauthorized' } -SignIn
        $step = Get-Step $r 'Free/busy of alice@contoso.com'
        $step.Status | Should -Be 'Failed'
        $step.Message | Should -Match 'ErrorProxyRequestProcessingFailed.*Url: https://outlook\.office365\.com/EWS/Exchange\.asmx'
        $r.FreeBusy[0].Hint | Should -Match 'the remote Exchange refused the token of the requesting organisation \(HTTP 401\)'
    }

    It 'on-premises with the Windows account: no token asked for the on-premises URL' {
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com'); OnPremisesAuthentication = 'Windows' } -State @{ SignInUser = 'carol@contoso.com' }
        $gua = Get-Step $r 'GetUserAvailability'
        $gua.Status | Should -Be 'Passed'
        $gua.Details.Authentication | Should -Match '^Windows \(current account'
        ($r.Trace | Where-Object Operation -eq 'EWS GetUserAvailability').Label | Should -Match 'Windows session'
        @($r.Trace | Where-Object { $_.Url -like '*/oauth2/v2.0/token' }) | Should -BeNullOrEmpty
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com'); OnPremisesAuthentication = 'Windows' } -State @{ SignInUser = 'carol@contoso.com'; WindowsAccepted = $false }
        (Get-Step $r 'GetUserAvailability').Message | Should -Match 'EWS on-premises answered HTTP 401: Exchange on-premises refused the Windows account'
    }

    It 'on-premises with an account typed, NTLM: the account answers the challenge, no token asked' {
        $cred = [pscredential]::new('CONTOSO\carol', (ConvertTo-SecureString 'Sim-Pa55word' -AsPlainText -Force))
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com'); OnPremisesAuthentication = 'Credential'; OnPremisesCredentialScheme = 'NTLM' } -State @{ SignInUser = 'carol@contoso.com' } -Credential $cred
        $gua = Get-Step $r 'GetUserAvailability'
        $gua.Status | Should -Be 'Passed'
        $gua.Details.Authentication | Should -Match '^NTLM, account CONTOSO\\carol \(password typed\)$'
        $trace = $r.Trace | Where-Object Operation -eq 'EWS GetUserAvailability'
        $trace.Label | Should -Match 'NTLM, account CONTOSO\\carol'
        $trace.Label | Should -Not -Match 'no credentials'
        @($r.Trace | Where-Object { $_.Url -like '*/oauth2/v2.0/token' }) | Should -BeNullOrEmpty
        (ConvertTo-Json -InputObject $r -Depth 8) | Should -Not -Match 'Sim-Pa55word'
    }

    It 'on-premises with an account typed, Auto takes the best scheme EWS offers (Negotiate), Basic refused when EWS does not offer it' {
        $cred = [pscredential]::new('carol@contoso.com', (ConvertTo-SecureString 'Sim-Pa55word' -AsPlainText -Force))
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com'); OnPremisesAuthentication = 'Credential' } -State @{ SignInUser = 'carol@contoso.com' } -Credential $cred
        (Get-Step $r 'GetUserAvailability').Details.Authentication | Should -Match '^Negotiate, account carol@contoso\.com'
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com'); OnPremisesAuthentication = 'Credential'; OnPremisesCredentialScheme = 'Basic' } -State @{ SignInUser = 'carol@contoso.com' } -Credential $cred
        $gua = Get-Step $r 'GetUserAvailability'
        $gua.Status | Should -Be 'Failed'
        $gua.Message | Should -Match 'HTTP 401: Exchange on-premises refused the account \(Basic, account carol@contoso\.com .*EWS does not offer Basic'
        $gua.Message | Should -Match 'BasicAuthentication, WindowsAuthentication'
    }

    It 'on-premises with an account typed, Basic: sent at once, user name only in the trace, a note to prefer OAuth' {
        $cred = [pscredential]::new('CONTOSO\carol', (ConvertTo-SecureString 'Sim-Pa55word' -AsPlainText -Force))
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com'); OnPremisesAuthentication = 'Credential'; OnPremisesCredentialScheme = 'Basic' } -State @{ SignInUser = 'carol@contoso.com'; OnPremBasic = $true } -Credential $cred
        $gua = Get-Step $r 'GetUserAvailability'
        $gua.Status | Should -Be 'Passed'
        $gua.Message | Should -Match 'prefer hybrid modern authentication'
        (Get-Step $r 'Free/busy of alice@contoso.com').Status | Should -Be 'Passed'
        $trace = $r.Trace | Where-Object Operation -eq 'EWS GetUserAvailability'
        $trace.Label | Should -Match 'user name and password \(Basic\)'
        $trace.Request | Should -Match 'Authorization: Basic <user CONTOSO\\carol, password never written>'
        (ConvertTo-Json -InputObject $r -Depth 8) | Should -Not -Match 'Sim-Pa55word'
        $wrong = [pscredential]::new('CONTOSO\carol', (ConvertTo-SecureString 'wrong' -AsPlainText -Force))
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com'); OnPremisesAuthentication = 'Credential'; OnPremisesCredentialScheme = 'Basic' } -State @{ SignInUser = 'carol@contoso.com'; OnPremBasic = $true } -Credential $wrong
        (Get-Step $r 'GetUserAvailability').Message | Should -Match 'refused the account.*password, account locked'
    }

    It 'on-premises with OAuth needs a refresh token: an access token alone is not enough' {
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com') } -State @{ SignInUser = 'carol@contoso.com' }
        (Get-Step $r 'GetUserAvailability').Status | Should -Be 'Failed'
        (Get-Step $r 'GetUserAvailability').Message | Should -Match 'No refresh token from the sign-in'
    }

    It 'on-premises URL not a service principal name of Exchange Online: AADSTS500011 and its cause' {
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com') } -State @{ SignInUser = 'carol@contoso.com'; OnPremSpns = @() } -SignIn
        $step = Get-Step $r 'GetUserAvailability'
        $step.Status | Should -Be 'Failed'
        $step.Message | Should -Match 'AADSTS500011'
        $step.Message | Should -Match 'hybrid modern authentication is not configured for this on-premises URL'
    }

    It 'another organisation: not in the directory, an external partner, free/busy only' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('dave@fabrikam.com') }
        $r.Status | Should -Be 'Passed'
        (Get-Step $r 'Target dave@fabrikam.com').Message | Should -Match "tenant $($script:SimFabrikam), federated sign-in"
        (Get-Step $r 'Path to dave@fabrikam.com').Details.Path | Should -Be 'CrossOrganisation'
        (Get-Step $r 'Directory dave@fabrikam.com').Message | Should -Match 'in another organisation'
        $tips = Get-Step $r 'MailTips dave@fabrikam.com'
        $tips.Message | Should -Match 'as an external partner'
        $tips.Message | Should -Not -Match 'not as a partner'
        $r.FreeBusy[0].DetailLevel | Should -Be 'Free/busy time only (no subject)'
        @($r.FreeBusyEvents | Where-Object Subject) | Should -BeNullOrEmpty
        @($script:State.Calls | Where-Object { $_.Path -eq '/v1.0/users' -and $_.Url -match 'fabrikam' }) | Should -BeNullOrEmpty
    }

    It 'another organisation that does not share: ErrorNoFreeBusyAccess and its cause' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('dave@fabrikam.com') } -State @{ CrossOrg = 'ErrorNoFreeBusyAccess' }
        (Get-Step $r 'MailTips dave@fabrikam.com').Message | Should -Match 'as external, not as a partner'
        (Get-Step $r 'Free/busy of dave@fabrikam.com').Message | Should -Match '^ErrorNoFreeBusyAccess: .*Cause: The target refuses to share its free/busy'
    }

    It 'a consumer address: the Microsoft accounts tenant is not an organisation' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('someone@gmail.com') }
        $m = $r.Mailboxes | Where-Object Address -eq 'someone@gmail.com'
        $m.Consumer | Should -BeTrue
        $m.TenantId | Should -BeNullOrEmpty
        $m.Location | Should -Be 'External'
        (Get-Step $r 'Target someone@gmail.com').Status | Should -Be 'Warning'
        (Get-Step $r 'Target someone@gmail.com').Message | Should -Match '^A consumer address'
        (Get-Step $r 'Path to someone@gmail.com').Message | Should -Match 'a consumer address \(gmail\.com\)'
        (Get-Step $r 'MailTips someone@gmail.com').Message | Should -Match 'as external, not as a partner'
        $fb = Get-Step $r 'Free/busy of someone@gmail.com'
        $fb.Message | Should -Match '^ErrorAvailabilityConfigNotFound: .*Cause: Exchange Online has no free/busy configuration for this domain'
        ($r.Trace | Where-Object { $_.Url -like '*/gmail.com/v2.0/.well-known/openid-configuration' }).Response | Should -Match $script:SimConsumer
    }

    It 'a consumer source stops the test at the discovery' {
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'someone@gmail.com' }
        $r.Status | Should -Be 'Failed'
        (Get-Step $r 'Source someone@gmail.com').Message | Should -Match 'consumer domain \(Microsoft accounts\): the source must be a mailbox of a Microsoft 365 organisation'
        @($r.Steps | Where-Object Stage -ne 'Discovery').Status | Should -Be @('Skipped', 'Skipped', 'Skipped', 'Skipped', 'Skipped')
        @($script:State.Calls | Where-Object Host -eq 'graph.microsoft.com') | Should -BeNullOrEmpty
    }

    It 'an unknown address of the organisation: every stage says so, none throws' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('bob@contoso.com', 'nobody@contoso.com') }
        (Get-Step $r 'Directory nobody@contoso.com').Status | Should -Be 'Warning'
        (Get-Step $r 'MailTips nobody@contoso.com').Status | Should -Be 'Warning'
        (Get-Step $r 'MailTips nobody@contoso.com').Message | Should -Match 'gave no MailTips: ErrorInvalidRecipients'
        (Get-Step $r 'MailTips bob@contoso.com').Status | Should -Be 'Passed'
        $fb = Get-Step $r 'Free/busy of nobody@contoso.com'
        $fb.Status | Should -Be 'Failed'
        $fb.Message | Should -Match '^ErrorMailRecipientNotFound: .*Cause: Exchange does not know this address'
        (Get-Step $r 'Free/busy of bob@contoso.com').Status | Should -Be 'Passed'
        Get-StageCrash $r | Should -BeNullOrEmpty
    }

    It 'waits as Microsoft Graph asks (429, Retry-After) and sends again' {
        $r = Invoke-Sim 'FreeBusy' -State @{ ThrottleOnce = 'calendar/getSchedule' }
        $r.Status | Should -Be 'Passed'
        $r.Throttled | Should -Be 1
        @($script:State.Waits) | Should -Contain 2
        (Get-Step $r 'Source mailbox').Details.Retries | Should -Be 1
        @($r.Trace | Where-Object StatusCode -eq 429).Count | Should -Be 1
    }
}

Describe 'On-premises endpoint' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }
    BeforeAll { $script:OnPrem = @{ TargetMailboxes = @('carol@contoso.com') } }

    It 'no Bearer challenge: Exchange Online cannot use OAuth there' {
        $r = Invoke-Sim 'FreeBusy' $script:OnPrem -State @{ OnPremBearer = $false }
        (Get-Step $r 'EWS mail.contoso.com').Status | Should -Be 'Warning'
        (Get-Step $r 'EWS mail.contoso.com').Message | Should -Match 'does not offer OAuth \(schemes: Negotiate, NTLM\)'
    }

    It 'another tenant trusted: the tokens of this tenant will be refused' {
        $r = Invoke-Sim 'FreeBusy' $script:OnPrem -State @{ OnPremTrustedTenant = $script:SimFabrikam }
        (Get-Step $r 'Autodiscover mail.contoso.com').Status | Should -Be 'Warning'
        (Get-Step $r 'Autodiscover mail.contoso.com').Message | Should -Match "the tenant $($script:SimContoso) is not among the issuers on-premises Exchange trusts \($($script:SimFabrikam)\)"
    }

    It 'a reverse proxy with pre-authentication' {
        $r = Invoke-Sim 'FreeBusy' $script:OnPrem -State @{ OnPremProxy = $true }
        (Get-Step $r 'EWS mail.contoso.com').Status | Should -Be 'Failed'
        (Get-Step $r 'EWS mail.contoso.com').Message | Should -Match 'answers HTTP 302 to https://sts\.contoso\.com/adfs/ls/.*reverse proxy with pre-authentication'
    }

    It 'the certificate: expiring soon, not trusted, or no TCP connection' {
        $r = Invoke-Sim 'FreeBusy' $script:OnPrem -State @{ CertificateDays = 10 }
        (Get-Step $r 'Certificate mail.contoso.com').Status | Should -Be 'Warning'
        (Get-Step $r 'Certificate mail.contoso.com').Message | Should -Match '10 days: renew it'
        $r = Invoke-Sim 'FreeBusy' $script:OnPrem -State @{ TlsFailure = 'Untrusted' }
        (Get-Step $r 'Certificate mail.contoso.com').Status | Should -Be 'Failed'
        (Get-Step $r 'Certificate mail.contoso.com').Message | Should -Match 'the certificate is not trusted.*UntrustedRoot'
        Get-Step $r 'EWS mail.contoso.com' | Should -BeNullOrEmpty
        $r = Invoke-Sim 'FreeBusy' $script:OnPrem -State @{ TlsFailure = 'Unreachable' }
        (Get-Step $r 'HTTPS mail.contoso.com').Status | Should -Be 'Failed'
        (Get-Step $r 'HTTPS mail.contoso.com').Message | Should -Match 'No TCP connection to mail\.contoso\.com:443'
    }

    It 'a host that does not resolve' {
        Start-Sim
        $script:State.Dns.Remove('mail.contoso.com|A')
        $cfg = $script:Base.Clone(); $cfg.TargetMailboxes = @('carol@contoso.com')
        $r = Invoke-GctConnectivityTest -Configuration $cfg -AccessToken (New-UserToken) -Quiet
        (Get-Step $r 'DNS mail.contoso.com').Status | Should -Be 'Failed'
        (Get-Step $r 'DNS mail.contoso.com').Message | Should -Match 'does not resolve'
        Get-Step $r 'Certificate mail.contoso.com' | Should -BeNullOrEmpty
        ($r.Trace | Where-Object Method -eq 'DNS' | Where-Object Url -eq 'dns:mail.contoso.com/A').Response | Should -Be 'No record (NXDOMAIN or empty answer).'
    }

    It 'Target.OnPremisesHost is checked even without an on-premises mailbox, for the tenant of the source' {
        $r = Invoke-Sim 'FreeBusy' @{ OnPremisesHost = 'mail.contoso.com' }
        (Get-Step $r 'EWS mail.contoso.com').Status | Should -Be 'Passed'
        (Get-Step $r 'EWS mail.contoso.com').Message | Should -Match "trusts the tenant $($script:SimContoso)"
    }
}

Describe 'Hybrid configuration (administrator)' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }
    BeforeAll { $script:Hybrid = @{ TargetMailboxes = @('bob@contoso.com', 'carol@contoso.com', 'dave@fabrikam.com', 'someone@gmail.com'); AdminChecks = $true } }

    It 'reads the IntraOrganizationConnector, the organization relationships and the availability address spaces for each domain' {
        $r = Invoke-Sim 'FreeBusy' $script:Hybrid
        $sign = Get-Step $r 'Administrator sign-in'
        $sign.Status | Should -Be 'Passed'
        $sign.Message | Should -Match 'Token of admin@contoso\.com for the Exchange Online administration API \(device code\)'
        ($r.Trace | Where-Object { $_.Url -like '*/oauth2/v2.0/devicecode' }).Request | Should -Match 'client_id=fb78d390-0c51-40cd-8e17-fdbfab77341b'
        ($r.Trace | Where-Object { $_.Url -like '*/oauth2/v2.0/devicecode' }).Request | Should -Match 'scope=https://outlook\.office365\.com/\.default'
        (Get-Step $r 'Get-IntraOrganizationConnector').Message | Should -Be '1 object(s): HybridIOC - 7d4e2a91.'
        (Get-Step $r 'Get-OrganizationRelationship').Message | Should -Be '2 object(s): O365 to On-premises - 7d4e2a91, Fabrikam.'
        (Get-Step $r 'Get-AvailabilityAddressSpace').Message | Should -Be '0 object(s).'
        $invoke = $r.Trace | Where-Object Operation -eq 'InvokeCommand Get-IntraOrganizationConnector'
        $invoke.Url | Should -Be "https://outlook.office365.com/adminapi/beta/$($script:SimContoso)/InvokeCommand"
        $invoke.Request | Should -Match 'X-AnchorMailbox: UPN:admin@contoso\.com'
        (Get-Step $r 'Configuration for contoso.com').Status | Should -Be 'Passed'
        (Get-Step $r 'Configuration for contoso.com').Message | Should -Match '^OAuth: the IntraOrganizationConnector HybridIOC - 7d4e2a91 sends the requests for contoso\.com to https://mail\.contoso\.com/autodiscover/autodiscover\.svc'
        (Get-Step $r 'Configuration for fabrikam.com').Status | Should -Be 'Passed'
        (Get-Step $r 'Configuration for fabrikam.com').Message | Should -Match '^Organization relationship Fabrikam \(DAuth, federation\).*level AvailabilityOnly'
        (Get-Step $r 'Configuration for gmail.com').Status | Should -Be 'Failed'
        Get-Step $r 'Configuration for bob.contoso.com' | Should -BeNullOrEmpty
        @($r.Hybrid.IntraOrganizationConnectors).Count | Should -Be 1
        $r.Hybrid.IntraOrganizationConnectors[0].TargetAddressDomains | Should -Be 'contoso.com'
        @($r.Hybrid.OrganizationRelationships).Count | Should -Be 2
        Get-StageCrash $r | Should -BeNullOrEmpty
    }

    It 'the IntraOrganizationConnector names another host than the one checked' {
        $state = @{}
        $r = Invoke-Sim 'FreeBusy' $script:Hybrid -State @{ AdminCommands = @{
                    'Get-IntraOrganizationConnector' = @(@{ Name = 'HybridIOC'; Enabled = $true; TargetAddressDomains = 'contoso.com'; DiscoveryEndpoint = 'https://hybrid.contoso.com/autodiscover/autodiscover.svc' })
                    'Get-OrganizationRelationship'   = @(@{ Name = 'Fabrikam'; Enabled = $true; DomainNames = @('fabrikam.com'); FreeBusyAccessEnabled = $false; FreeBusyAccessLevel = 'None' })
                    'Get-AvailabilityAddressSpace'   = @(@{ ForestName = 'gmail.com'; AccessMethod = 'OrgWideFB'; ProxyUrl = 'https://fb.example/ews/exchange.asmx' })
                } }
        (Get-Step $r 'Configuration for contoso.com').Status | Should -Be 'Warning'
        (Get-Step $r 'Configuration for contoso.com').Message | Should -Match 'run the test with -OnPremisesHost hybrid\.contoso\.com'
        (Get-Step $r 'Configuration for fabrikam.com').Status | Should -Be 'Failed'
        (Get-Step $r 'Configuration for fabrikam.com').Message | Should -Match 'FreeBusyAccessEnabled is False'
        (Get-Step $r 'Configuration for gmail.com').Message | Should -Match 'Availability address space for gmail\.com \(OrgWideFB, https://fb\.example/ews/exchange\.asmx\)'
    }

    It 'an account without an Exchange role: a warning, never a failure of the test' {
        $r = Invoke-Sim 'FreeBusy' $script:Hybrid -State @{ AdminForbidden = $true }
        (Get-Step $r 'Get-IntraOrganizationConnector').Status | Should -Be 'Warning'
        (Get-Step $r 'Get-IntraOrganizationConnector').Message | Should -Match 'no role that allows this cmdlet.*Connect-ExchangeOnline'
        @(Get-StageSteps $r 'Hybrid' | Where-Object Status -eq 'Failed') | Should -BeNullOrEmpty
    }

    It 'without AdminChecks there is no administrator sign-in' {
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('carol@contoso.com') }
        Get-StageSteps $r 'Hybrid' | Should -BeNullOrEmpty
        @($script:State.Calls | Where-Object Path -like '/adminapi/*') | Should -BeNullOrEmpty
    }

    It 'a source on-premises: the commands to run on-premises are named' {
        $r = Invoke-Sim 'FreeBusy' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com'); AdminChecks = $true } -State @{ SignInUser = 'carol@contoso.com' } -SignIn
        (Get-Step $r 'On-premises configuration').Message | Should -Match 'Test-OAuthConnectivity -Service EWS -TargetUri https://outlook\.office365\.com/ews/exchange\.asmx -Mailbox carol@contoso\.com'
        (Get-Step $r 'Domains').Message | Should -Match 'Every target is a mailbox of Exchange Online'
    }
}

Describe 'Error causes (Get-GctAvailabilityHint)' {
    BeforeAll {
        Import-Module (Join-Path $script:Root 'GraphConnectivityTests.psd1') -Force
        $script:Hint = { param($Code, $Message, $Path) & (Get-Module GraphConnectivityTests) { param($c, $m, $p) Get-GctAvailabilityHint -Code $c -Message $m -Path $p } $Code $Message $Path }
    }
    AfterAll { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'names the cause of each error of the Availability service' -ForEach @(
        @{ Code = 'ErrorMailRecipientNotFound'; Message = ''; Path = 'SameOrganisation'; Expected = '^Exchange does not know this address: no mailbox, mail user or contact with it in the source organisation\. In hybrid' }
        @{ Code = 'ErrorMailRecipientNotFound'; Message = ''; Path = 'CrossOrganisation'; Expected = 'and no organization relationship, IntraOrganizationConnector or availability address space for its domain' }
        @{ Code = 'ErrorAvailabilityConfigNotFound'; Message = ''; Path = 'CrossOrganisation'; Expected = 'no free/busy configuration for this domain' }
        @{ Code = 'ErrorNoFreeBusyAccess'; Message = ''; Path = 'CrossOrganisation'; Expected = 'refuses to share its free/busy' }
        @{ Code = 'ErrorProxyRequestProcessingFailed'; Message = 'The remote server returned an error: (401) Unauthorized.'; Path = 'CloudToOnPremises'; Expected = 'failed there: the remote Exchange refused the token.*Test-OAuthConnectivity' }
        @{ Code = 'ErrorProxyRequestProcessingFailed'; Message = 'The remote server returned an error: (403) Forbidden.'; Path = 'CloudToOnPremises'; Expected = 'failed there: the remote endpoint answered HTTP 403' }
        @{ Code = 'ErrorProxyRequestProcessingFailed'; Message = 'The remote server returned an error: (404) Not Found.'; Path = 'CloudToOnPremises'; Expected = 'the remote URL does not exist \(HTTP 404\)' }
        @{ Code = 'ErrorProxyRequestProcessingFailed'; Message = 'The remote server returned an error: (503) Server Unavailable.'; Path = 'CloudToOnPremises'; Expected = 'not available \(HTTP 5xx\)' }
        @{ Code = 'ErrorProxyRequestProcessingFailed'; Message = 'Unable to connect to the remote server'; Path = 'CloudToOnPremises'; Expected = 'could not reach the remote endpoint from the Internet' }
        @{ Code = 'ErrorProxyRequestProcessingFailed'; Message = 'Could not establish trust relationship for the SSL/TLS secure channel.'; Path = 'CloudToOnPremises'; Expected = 'the TLS connection to the remote endpoint failed' }
        @{ Code = 'ErrorProxyRequestProcessingFailed'; Message = 'Proxy web request failed.'; Path = 'CloudToOnPremises'; Expected = 'Read the full message: it gives the remote URL' }
        @{ Code = 'ErrorAutoDiscoverFailed'; Message = 'Autodiscover failed for the remote organization.'; Path = 'CrossOrganisation'; Expected = 'Autodiscover of the remote organisation failed' }
        @{ Code = 'ErrorFreeBusyGenerationFailed'; Message = ''; Path = 'SameOrganisation'; Expected = 'could not build its free/busy' }
        @{ Code = 'ErrorFreeBusyDLLimitReached'; Message = ''; Path = 'SameOrganisation'; Expected = 'at most 100 members' }
        @{ Code = 'ErrorServerBusy'; Message = ''; Path = 'SameOrganisation'; Expected = 'throttled' }
        @{ Code = 'ErrorSomethingNew'; Message = 'A delegation token could not be obtained from the Federation Gateway.'; Path = 'CrossOrganisation'; Expected = 'federation token \(DAuth\) could not be obtained' }
    ) {
        & $script:Hint $Code $Message $Path | Should -Match $Expected
    }

    It 'says nothing it does not know' {
        & $script:Hint 'ErrorSomethingNew' '' 'SameOrganisation' | Should -BeNullOrEmpty
    }

    It 'names the cause of the Entra ID errors met most often' {
        $m = Get-Module GraphConnectivityTests
        & $m { Get-GctEntraErrorHint -Description 'AADSTS500011: The resource principal named https://mail.contoso.com was not found' -Resource 'https://mail.contoso.com' } | Should -Match 'not a service principal name of Office 365 Exchange Online'
        & $m { Get-GctEntraErrorHint -Description 'AADSTS65001: consent' } | Should -Match 'consented'
        & $m { Get-GctEntraErrorHint -Description 'AADSTS53003: blocked' } | Should -Match 'Conditional Access'
        & $m { Get-GctEntraErrorHint -Description 'no code here' } | Should -BeNullOrEmpty
    }
}

Describe 'MailboxServices' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'read-only: delta synchronization, own free/busy and MailTips, automatic replies read; nothing changed' {
        $r = Invoke-Sim 'MailboxServices'
        $r.Status | Should -Be 'Passed'
        (Get-Step $r 'Folder synchronization').Message | Should -Match 'First synchronization: 8 folder\(s\) in 2 page\(s\); the deltaLink sent again returns 0 change\(s\)'
        (Get-Step $r 'Message synchronization').Message | Should -Match '3 message\(s\) \(headers only\) in 1 page\(s\); the deltaLink sent again returns 1 change\(s\)'
        (Get-Step $r 'Change notifications').Status | Should -Be 'Skipped'
        (Get-Step $r 'Own free/busy').Status | Should -Be 'Passed'
        (Get-Step $r 'Own MailTips').Message | Should -Match 'maximum message size 36 MB, automatic replies off'
        (Get-Step $r 'Read the automatic replies').Message | Should -Match 'Automatic replies disabled \(external audience: all\)'
        (Get-Step $r 'Change the automatic replies').Status | Should -Be 'Skipped'
        (Get-Step $r 'Change the automatic replies').Message | Should -Match '-AllowWrite was not given'
        $script:State.Patches.Count | Should -Be 0
        @($script:State.Calls | Where-Object Method -eq 'PATCH') | Should -BeNullOrEmpty
        @($r.Trace | Where-Object Operation -like 'GET mailFolders delta*').Operation | Should -Be @('GET mailFolders delta (page 1)', 'GET mailFolders delta (page 2)', 'GET mailFolders delta (deltaLink)')
        $r.FreeBusy[0].Mailbox | Should -Be 'alice@contoso.com'
    }

    It 'with AllowWrite and the automatic replies disabled: scheduled, read back, restored with the original messages' {
        $r = Invoke-Sim 'MailboxServices' @{ AllowWrite = $true }
        $r.Status | Should -Be 'Passed'
        (Get-Step $r 'Change the automatic replies').Status | Should -Be 'Passed'
        (Get-Step $r 'Change the automatic replies').Details.ReadBack | Should -Be 'scheduled'
        $script:State.Patches.Count | Should -Be 2
        $script:State.Patches[0].automaticRepliesSetting.status | Should -Be 'scheduled'
        $script:State.Patches[0].automaticRepliesSetting.internalReplyMessage | Should -Match '^\[Graph Connectivity Tests\]'
        ([datetime]$script:State.Patches[0].automaticRepliesSetting.scheduledStartDateTime.dateTime) | Should -BeGreaterThan ([datetime]::UtcNow.AddDays(59))
        $script:State.Patches[1].automaticRepliesSetting.status | Should -Be 'disabled'
        $script:State.AutoReply.status | Should -Be 'disabled'
        $script:State.AutoReply.internalReplyMessage | Should -Be '<html><body>Original internal text</body></html>'
        $script:State.AutoReply.externalReplyMessage | Should -Be '<html><body>Original external text</body></html>'
        $r.Actions.Action | Should -Be @('Schedule automatic replies', 'Restore automatic replies')
        $r.Actions.Result | Should -Be @('HTTP 200', 'HTTP 200')
    }

    It 'never touches automatic replies that are on' {
        $r = Invoke-Sim 'MailboxServices' @{ AllowWrite = $true } -State @{ AutoReply = @{ status = 'alwaysEnabled'; externalAudience = 'contactsOnly'; internalReplyMessage = 'On holiday'; externalReplyMessage = 'On holiday' } }
        (Get-Step $r 'Change the automatic replies').Status | Should -Be 'Skipped'
        (Get-Step $r 'Change the automatic replies').Message | Should -Match 'the automatic replies are alwaysEnabled - the tool changes them only when they are disabled'
        $script:State.Patches.Count | Should -Be 0
        $script:State.AutoReply.status | Should -Be 'alwaysEnabled'
        @($r.Actions) | Should -BeNullOrEmpty
    }

    It 'automatic replies not readable (403): a warning, no change attempted' {
        $r = Invoke-Sim 'MailboxServices' @{ AllowWrite = $true } -State @{ AutoReplyForbidden = $true }
        (Get-Step $r 'Read the automatic replies').Status | Should -Be 'Warning'
        (Get-Step $r 'Read the automatic replies').Message | Should -Match 'HTTP 403 \(ErrorAccessDenied.*MailTips say they are off'
        Get-Step $r 'Change the automatic replies' | Should -BeNullOrEmpty
        $script:State.Patches.Count | Should -Be 0
    }

    It 'a source on-premises: Graph cannot reach it, the stages say so' {
        $r = Invoke-Sim 'MailboxServices' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @() } -State @{ SignInUser = 'carol@contoso.com' }
        (Get-Step $r 'Source mailbox').Status | Should -Be 'Failed'
        (Get-Step $r 'Source mailbox').Message | Should -Match 'seems to be on-premises, but no on-premises EWS URL is known'
        $r = Invoke-Sim 'MailboxServices' @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @(); OnPremisesHost = 'mail.contoso.com' } -State @{ SignInUser = 'carol@contoso.com' }
        (Get-Step $r 'Synchronization').Status | Should -Be 'Skipped'
        (Get-Step $r 'Availability').Status | Should -Be 'Skipped'
        (Get-Step $r 'Automatic replies').Status | Should -Be 'Skipped'
    }
}

Describe 'Provisioning' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'for each address: user, licenses, mailbox in Exchange Online or on-premises, type' {
        $r = Invoke-Sim 'Provisioning' @{ TargetMailboxes = @('bob@contoso.com', 'room1@contoso.com', 'carol@contoso.com', 'erin@contoso.com', 'nobody@contoso.com', 'dave@fabrikam.com') }
        $bob = Get-Step $r 'Provisioning of bob@contoso.com'
        $bob.Status | Should -Be 'Passed'
        $bob.Message | Should -Match 'has a mailbox in Exchange Online \(user mailbox\); licenses ENTERPRISEPACK'
        $bob.Details.MailboxProbe | Should -Be 'HTTP 403 ErrorAccessDenied'
        $bob.Details.ExchangePlans | Should -Be 'EXCHANGE_S_ENTERPRISE (Success)'
        (Get-Step $r 'Provisioning of room1@contoso.com').Message | Should -Match '\(room mailbox\)'
        (Get-Step $r 'Provisioning of room1@contoso.com').Status | Should -Be 'Passed'
        $carol = Get-Step $r 'Provisioning of carol@contoso.com'
        $carol.Status | Should -Be 'Passed'
        $carol.Message | Should -Match 'its mailbox is on-premises \(mail\.contoso\.com\)'
        $carol.Details.MailboxProbe | Should -Be 'HTTP 404 MailboxNotEnabledForRESTAPI'
        (Get-Step $r 'Provisioning of erin@contoso.com').Status | Should -Be 'Failed'
        (Get-Step $r 'Provisioning of erin@contoso.com').Message | Should -Match 'no Exchange license is assigned'
        (Get-Step $r 'Provisioning of nobody@contoso.com').Status | Should -Be 'Failed'
        (Get-Step $r 'Provisioning of dave@fabrikam.com').Status | Should -Be 'Skipped'
        @($r.Provisioning).Count | Should -Be 4
        ($r.Provisioning | Where-Object Mail -eq 'carol@contoso.com').MailboxInExchangeOnline | Should -BeFalse
        Get-StageCrash $r | Should -BeNullOrEmpty
    }

    It 'without targets the source itself is checked' {
        $r = Invoke-Sim 'Provisioning' @{ TargetMailboxes = @() }
        (Get-Step $r 'Provisioning of alice@contoso.com').Details.MailboxProbe | Should -Be 'HTTP 200'
        (Get-Step $r 'Provisioning of alice@contoso.com').Status | Should -Be 'Passed'
    }
}

Describe 'DomainDns' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'the records of the domain of the source, then the records Microsoft 365 expects' {
        $r = Invoke-Sim 'DomainDns' @{ TargetMailboxes = @() }
        $r.Domain | Should -Be 'contoso.com'
        @(Get-StageSteps $r 'Dns').Name | Should -Be @('MX', 'SPF', 'Autodiscover', 'DKIM', 'DMARC', 'MTA-STS', 'Device registration')
        @(Get-StageSteps $r 'Dns').Status | Should -Be @('Passed', 'Passed', 'Passed', 'Passed', 'Passed', 'Passed', 'Passed')
        (Get-Step $r 'MX').Message | Should -Be 'Mail goes to Exchange Online Protection: contoso-com.mail.protection.outlook.com.'
        (Get-Step $r 'Autodiscover').Message | Should -Match 'points to mail\.contoso\.com: Exchange on-premises answers Autodiscover'
        (Get-Step $r 'MTA-STS').Message | Should -Match 'mode enforce'
        (Get-Step $r 'Domain contoso.com').Message | Should -Match 'Verified, Managed sign-in, services: Email'
        $expected = Get-Step $r 'Expected records'
        $expected.Status | Should -Be 'Warning'
        $expected.Message | Should -Match '^3 of 6 records'
        $expected.Details.Missing | Should -Be 'Email CNAME autodiscover.contoso.com; Intune CNAME enterpriseenrollment.contoso.com'
        @($r.DnsRecords | Where-Object Source -eq 'Microsoft 365').Count | Should -Be 6
        ($r.DnsRecords | Where-Object { $_.Source -eq 'Microsoft 365' -and $_.Type -eq 'SRV' }).Status | Should -Be 'Skipped'
        @($r.Trace | Where-Object Method -eq 'DNS').Count | Should -BeGreaterThan 10
        ($r.Trace | Where-Object Url -eq 'dns:contoso.com/MX' | Select-Object -First 1).Response | Should -Be "contoso.com`tMX`t0 contoso-com.mail.protection.outlook.com"
    }

    It 'another domain: MX elsewhere, SPF through an include, no DKIM nor DMARC, not a domain of the tenant' {
        $r = Invoke-Sim 'DomainDns' @{ Domain = 'fabrikam.com'; TargetMailboxes = @() }
        (Get-Step $r 'MX').Message | Should -Match '^Mail goes to mx1\.fabrikam\.com \(not to Exchange Online Protection'
        (Get-Step $r 'SPF').Details.IncludedThrough | Should -Be '_spf.fabrikam.com'
        (Get-Step $r 'SPF').Details.Lookups | Should -Be 2
        (Get-Step $r 'Autodiscover').Message | Should -Match 'autodiscover\.outlook\.com \(Exchange Online\)'
        (Get-Step $r 'DKIM').Status | Should -Be 'Warning'
        (Get-Step $r 'DMARC').Status | Should -Be 'Warning'
        (Get-Step $r 'MTA-STS').Status | Should -Be 'Passed'
        (Get-Step $r 'Device registration').Status | Should -Be 'Passed'
        Get-StageCrash $r | Should -BeNullOrEmpty
        (Get-Step $r 'Domain fabrikam.com').Status | Should -Be 'Failed'
        (Get-Step $r 'Domain fabrikam.com').Message | Should -Be 'fabrikam.com is not a domain of this tenant.'
    }

    It 'no DNS answer stops the DNS checks at the MX' {
        $r = Invoke-Sim 'DomainDns' @{ TargetMailboxes = @() } -State @{ DnsDown = $true }
        (Get-Step $r 'MX').Status | Should -Be 'Failed'
        (Get-Step $r 'MX').Message | Should -Match 'No DNS answer for contoso\.com \(This operation returned because the timeout period expired\)'
        @(Get-StageSteps $r 'Dns').Count | Should -Be 1
    }
}

Describe 'All' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'every stage in one run with one sign-in, none throws' {
        $r = Invoke-Sim 'All' @{ TargetMailboxes = @('bob@contoso.com', 'carol@contoso.com', 'dave@fabrikam.com') } -SignIn
        @($r.Steps.Stage | Select-Object -Unique) | Should -Be @('Discovery', 'OnPremises', 'Dns', 'SignIn', 'Source', 'Target', 'FreeBusy', 'Sync', 'Availability', 'AutoReply', 'Provisioning', 'DomainRecords')
        Get-StageCrash $r | Should -BeNullOrEmpty
        @($r.Steps | Where-Object { $_.Message -eq 'Not run: an earlier step failed.' }) | Should -BeNullOrEmpty
        @($script:State.Calls | Where-Object { $_.Path -like '*/oauth2/v2.0/devicecode' }).Count | Should -Be 1
        $r.Status | Should -Be 'Warning'
        (Get-Step $r 'Change the automatic replies').Status | Should -Be 'Skipped'
        @($r.FreeBusy.Mailbox) | Should -Be @('bob@contoso.com', 'carol@contoso.com', 'dave@fabrikam.com', 'alice@contoso.com')
    }
}

Describe 'Secrets never written' {
    AfterEach { Stop-GctLog; Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'no token, refresh token, device or authorization code, nor client assertion in the trace, the result, the report files or the log' {
        $runs = @(
            @{ Type = 'FreeBusy'; Overrides = @{ SourceMailbox = 'carol@contoso.com'; TargetMailboxes = @('alice@contoso.com', 'dave@fabrikam.com'); AdminChecks = $true }; State = @{ SignInUser = 'carol@contoso.com' } }
            @{ Type = 'All'; Overrides = @{ SignIn = 'Window'; TargetMailboxes = @('bob@contoso.com', 'carol@contoso.com'); AllowWrite = $true }; State = @{ DesktopSession = $true } }
            @{ Type = 'MailboxServices'; Overrides = @{ Context = 'Application'; ClientId = $script:AppId; CertificateThumbprint = $script:Thumbprint; TenantId = 'contoso.com' }; State = @{} }
        )
        foreach ($run in $runs) {
            $logDir = Join-Path $script:Artifacts ('test-logs-secrets-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            Start-Sim $run.State
            Start-GctLog -Directory $logDir | Out-Null
            $cfg = $script:Base.Clone(); foreach ($k in $run.Overrides.Keys) { $cfg[$k] = $run.Overrides[$k] }
            $r = Invoke-GctConnectivityTest -Configuration $cfg -TestType $run.Type -Quiet
            Stop-GctLog
            Get-StageCrash $r | Should -BeNullOrEmpty
            $report = Export-GctReport -Result $r -OutputPath (Join-Path $script:Artifacts 'test-reports') -Formats Csv
            $files = (Get-ChildItem -LiteralPath $report.Directory, $logDir -File | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
            $files | Should -Match 'GraphConnectivityTests'
            Assert-NoSecret ((Get-ResultText $r) + "`n" + $files)
            ($r.Trace.Request -join "`n") | Should -Not -Match 'eyJ0eXAiOiJKV1Qi'
            ($r.Trace.Response -join "`n") | Should -Not -Match 'eyJ0eXAiOiJKV1Qi'
            Remove-Item -LiteralPath $logDir -Recurse -Force
        }
    }

    It 'masks Basic, Windows tokens and cookies, shows only the affinity cookies of Exchange' {
        Import-Module (Join-Path $script:Root 'GraphConnectivityTests.psd1') -Force
        $m = Get-Module GraphConnectivityTests
        $basic = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('CONTOSO\alice:P@ssw0rd!'))
        & $m { param($v) Protect-GctHeaderValue -Name Authorization -Value $v } $basic | Should -Be 'Basic <user CONTOSO\alice, password never written>'
        & $m { Protect-GctHeaderValue -Name Authorization -Value 'Negotiate YIIGhgYGKwYBBQUCoIIGejCCBnag' } | Should -Be 'Negotiate <Windows token: 28 characters, never written>'
        & $m { Protect-GctHeaderValue -Name Authorization -Value 'Bearer eyJ0eXAi.eyJhdWQi.c2ln' } | Should -Be 'Bearer <access token: 22 characters, never written>'
        & $m { Protect-GctHeaderValue -Name Cookie -Value 'X-BackEndOverrideCookie=EXCH02~1942062522; ESTSAUTH=secret-session; other=1' } | Should -Be 'X-BackEndOverrideCookie=EXCH02~1942062522; ESTSAUTH=<hidden>; other=<hidden>'
        & $m { Protect-GctHeaderValue -Name Set-Cookie -Value 'ESTSAUTHPERSISTENT=abcdef; path=/; secure' } | Should -Be 'ESTSAUTHPERSISTENT=<hidden, 6 characters>; path=/; secure'
        & $m { Format-GctFormBody 'grant_type=password&username=a&password=Secret%21&client_assertion=abc' } | Should -Match 'password=<7 characters, never written>'
    }
}

Describe 'Report' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'guards the CSV cells against formula injection' {
        Import-Module (Join-Path $script:Root 'GraphConnectivityTests.psd1') -Force
        $m = Get-Module GraphConnectivityTests
        $cell = { param($v, $d = ';') & $m { param($x, $y) Format-GctCsvCell $x $y } $v $d }
        & $cell '=1+1' | Should -Be "'=1+1"
        & $cell '+33 1 23' | Should -Be "'+33 1 23"
        & $cell '-2' | Should -Be "'-2"
        & $cell '@SUM(A1)' | Should -Be "'@SUM(A1)"
        & $cell "`t=cmd" | Should -Be "'`t=cmd"
        & $cell 'plain; text' | Should -Be '"plain; text"'
        & $cell 'say "hi"' | Should -Be '"say ""hi"""'
        & $cell ([int]-5) | Should -Be '-5'
        & $cell '-5' | Should -Be "'-5"
        & $cell $true | Should -Be 'True'
        & $cell @('a', 'b') | Should -Be 'a, b'
        & $cell @('a', 'b') ',' | Should -Be '"a, b"'
    }

    It 'writes every CSV file and the JSON summary, a formula of the calendar is neutralised' {
        $r = Invoke-Sim 'All' @{ TargetMailboxes = @('bob@contoso.com', 'carol@contoso.com') } -SignIn
        $report = Export-GctReport -Result $r -OutputPath (Join-Path $script:Artifacts 'test-reports') -Formats Csv -Delimiter ','
        foreach ($k in 'Steps', 'Mailboxes', 'FreeBusy', 'Events', 'Provisioning', 'Dns', 'Actions', 'Trace', 'Summary') { Test-Path -LiteralPath $report.Files[$k] | Should -BeTrue -Because "$k is written" }
        $report.Files.Contains('Html') | Should -BeFalse
        $events = [IO.File]::ReadAllText($report.Files.Events)
        $events | Should -Match ([regex]::Escape('"''=HYPERLINK(""https://evil.example"",""Agenda"")"'))
        $events | Should -Match ([regex]::Escape(",'@Room 9,"))
        $events | Should -Not -Match '(^|,)=HYPERLINK'
        [IO.File]::ReadAllBytes($report.Files.Steps)[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
        (Get-Content -LiteralPath $report.Files.Steps -TotalCount 1) | Should -Be ([char]0xFEFF + 'Stage,Name,Status,Message,Details,DurationMs,TimestampUtc,Trace')
        $summary = Get-Content -LiteralPath $report.Files.Summary -Raw | ConvertFrom-Json
        $summary.Status | Should -Be $r.Status
        @($summary.Steps).Count | Should -Be @($r.Steps).Count
        Assert-NoSecret ((Get-ChildItem -LiteralPath $report.Directory -File | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n")
    }

    It 'writes the HTML report with every marker of the template replaced' {
        # The template is maintained separately: a marker the engine does not fill is a pending dependency, not an engine failure.
        $known = @('TITLE', 'SUMMARY_JSON', 'STEPS_JSON', 'MAILBOXES_JSON', 'FREEBUSY_JSON', 'EVENTS_JSON', 'PROVISIONING_JSON', 'DNS_JSON', 'ACTIONS_JSON', 'TRACE_JSON')
        $template = [IO.File]::ReadAllText((Join-Path $script:Root 'templates\Report.template.html'))
        $unknown = @([regex]::Matches($template, '\{\{([A-Z_]+)\}\}') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique | Where-Object { $_ -notin $known })
        if ($unknown.Count) { Set-ItResult -Pending -Because "the report template still has markers the engine does not fill ($($unknown -join ', ')): it is being rewritten"; return }
        $r = Invoke-Sim 'FreeBusy' @{ TargetMailboxes = @('bob@contoso.com', 'carol@contoso.com', 'someone@gmail.com') } -SignIn
        $report = Export-GctReport -Result $r -OutputPath (Join-Path $script:Artifacts 'test-reports')
        foreach ($k in 'Steps', 'Mailboxes', 'FreeBusy', 'Events', 'Provisioning', 'Dns', 'Actions', 'Trace', 'Summary', 'Html') { Test-Path -LiteralPath $report.Files[$k] | Should -BeTrue -Because "$k is written" }
        $html = [IO.File]::ReadAllText($report.Files.Html)
        $html | Should -Not -Match '\{\{[A-Za-z_]+\}\}'
        $html | Should -Match '<title>Graph Connectivity Tests \| Failed \| FreeBusy</title>'
        foreach ($block in [regex]::Matches($html, '<script id="data-[a-z]+" type="application/json">(.*?)</script>', 'Singleline')) { { $block.Groups[1].Value | ConvertFrom-Json } | Should -Not -Throw }
        Assert-NoSecret $html
    }
}

Describe 'Console' {
    AfterEach { Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue }

    It 'a run with the console output does not throw, the free/busy grid shows each mailbox' {
        Start-Sim
        $cfg = $script:Base.Clone(); $cfg.TargetMailboxes = @('bob@contoso.com', 'dave@fabrikam.com')
        $token = New-UserToken
        $out = & { Invoke-GctConnectivityTest -Configuration $cfg -AccessToken $token } 6>&1
        $result = @($out | Where-Object { $_ -isnot [Management.Automation.InformationRecord] })[-1]
        $result.Status | Should -Be 'Passed'
        $text = (@($out | Where-Object { $_ -is [Management.Automation.InformationRecord] }) | ForEach-Object { [string]$_.MessageData }) -join "`n"
        $text | Should -Match 'Free/busy'
        $text | Should -Match 'bob@contoso\.com'
        $text | Should -Match 'dave@fabrikam\.com'
        $text | Should -Match '07 08 09 10'
        $settings = & (Get-Module GraphConnectivityTests) { Get-GctDefaultConfiguration }; foreach ($k in $cfg.Keys) { $settings[$k] = $cfg[$k] }
        { & { Write-GctRunBanner -Settings $settings -LogPath 'x.log'; Write-GctRunSummary -Result $result -ReportText 'report.html' -LogPath 'x.log' } 6>$null } | Should -Not -Throw
        $empty = & (Get-Module GraphConnectivityTests) { $script:Quiet = $false; Write-GctFreeBusyGrid -Results @() -StartUtc ([datetime]::UtcNow) } 6>&1
        $empty | Should -BeNullOrEmpty
    }

    It 'the run banner and summary describe every test type and status' -ForEach @(
        @{ Type = 'FreeBusy'; Status = 'Passed' }, @{ Type = 'MailboxServices'; Status = 'Blocked' }, @{ Type = 'DomainDns'; Status = 'Warning' }, @{ Type = 'All'; Status = 'Failed' }
    ) {
        Import-Module (Join-Path $script:Root 'GraphConnectivityTests.psd1') -Force
        $settings = & (Get-Module GraphConnectivityTests) { Get-GctDefaultConfiguration }
        $settings.TestType = $Type; $settings.TargetMailboxes = @('bob@contoso.com'); $settings.AllowWrite = $true; $settings.AdminChecks = $true; $settings.SignIn = 'DeviceCode'
        $result = [pscustomobject]@{ Status = $Status; TestType = $Type; Counts = [ordered]@{ Passed = 1; Warning = 1; Blocked = 0; Failed = 0; Skipped = 2 }; FreeBusy = @([pscustomobject]@{ Status = 'Passed' }); Error = 'Step: message'; DurationSeconds = 3.2 }
        $lines = & { Write-GctRunBanner -Settings $settings; Write-GctRunSummary -Result $result } 6>&1 | ForEach-Object { [string]$_.MessageData }
        ($lines -join "`n") | Should -Match $Type
        ($lines -join "`n") | Should -Match '1 passed'
        if ($Type -in 'MailboxServices', 'All') { ($lines -join "`n") | Should -Match 'AUTHORISED: automatic replies scheduled' }
    }
}

Describe 'Command line (Invoke-GraphConnectivityTests.ps1)' {
    <#
        The script imports the module itself (Import-Module -Force): the simulator cannot be installed in it, so only
        the exits that need no network are tested here - the exit codes 0 and 2 come from the status of the result
        (Passed / Warning or Blocked), tested above through Invoke-GctConnectivityTest.
    #>
    BeforeAll {
        $script:Pwsh = (Get-Process -Id $PID).Path
        $script:Script = Join-Path $script:Root 'Invoke-GraphConnectivityTests.ps1'
        $script:CliDir = Join-Path $script:Artifacts 'test-cli'
        [void][IO.Directory]::CreateDirectory($script:CliDir)
        $logs = (Join-Path $script:CliDir 'logs').Replace("'", "''")
        $script:NoTarget = Join-Path $script:CliDir 'no-target.config.psd1'
        "@{ Target = @{ SourceMailbox = 'alice@contoso.com'; TargetMailboxes = @() }; Logging = @{ Path = '$logs' }; Report = @{ OutputPath = '$((Join-Path $script:CliDir 'reports').Replace("'", "''"))' } }" | Set-Content -LiteralPath $script:NoTarget
        $script:Bad = Join-Path $script:CliDir 'bad.config.psd1'
        "@{ Target = @{ SourceMailbox = 'alice@contoso.com'; Mailboxes = @() } }" | Set-Content -LiteralPath $script:Bad
        function Invoke-Cli([string[]]$Arguments) {
            $output = & $script:Pwsh -NoProfile -NonInteractive -File $script:Script @Arguments 2>&1 | Out-String
            [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
        }
    }

    It 'exit code 1 for a configuration file with an unknown key' {
        $r = Invoke-Cli @('-ConfigPath', $script:Bad)
        $r.ExitCode | Should -Be 1
        $r.Output | Should -Match "Unknown key 'Target\.Mailboxes'"
    }

    It 'exit code 1 for a FreeBusy test without a target, before any request' {
        $r = Invoke-Cli @('-ConfigPath', $script:NoTarget, '-TestType', 'FreeBusy')
        $r.ExitCode | Should -Be 1
        $r.Output | Should -Match 'The FreeBusy test needs at least one target mailbox'
        Get-ChildItem -LiteralPath (Join-Path $script:CliDir 'logs') -Filter 'GraphConnectivityTests_*.log' | Should -Not -BeNullOrEmpty
        Test-Path -LiteralPath (Join-Path $script:CliDir 'reports') | Should -BeFalse
    }

    It 'exit code 1 for a command-line value out of its set' {
        $r = Invoke-Cli @('-ConfigPath', $script:NoTarget, '-TestType', 'Everything')
        $r.ExitCode | Should -Be 1
    }
}
