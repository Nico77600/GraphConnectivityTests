<#
.SYNOPSIS
    Graph Connectivity Tests - configuration, clouds, clients and tests (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    The configuration file has sections (Target, Identity, Test, Report, Logging), like the other
    tools of the author. It is flattened into one settings hashtable used by the command line, the
    window and the tests; unknown sections or keys and invalid values are all reported at once.
    Secrets are never part of it: the sign-in is interactive, an application signs in with a
    certificate of the certificate store.

    No application to register: by default the user signs in with Microsoft Office, a public client
    of Microsoft present in every tenant and pre-authorized for Microsoft Graph (no consent). Another
    client can be named (Identity.ClientId): Microsoft Graph Command Line Tools, or your own
    multi-tenant application.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>

# Section.Key of the configuration file -> key of the settings hashtable.
$script:ConfigSchema = [ordered]@{
    Target   = [ordered]@{
        SourceMailbox = 'SourceMailbox'; TargetMailboxes = 'TargetMailboxes'; OnPremisesHost = 'OnPremisesHost'; SkipOnPremisesChecks = 'SkipOnPremisesChecks'
        Domain = 'Domain'; Cloud = 'Cloud'; HeaderFile = 'HeaderFile'; HeaderText = 'HeaderText'
    }
    Identity = [ordered]@{
        Context = 'Context'; ClientId = 'ClientId'; RedirectUri = 'RedirectUri'; Scopes = 'Scopes'; TenantId = 'TenantId'; CertificateThumbprint = 'CertificateThumbprint'
        OnPremisesAuthentication = 'OnPremisesAuthentication'; OnPremisesCredentialScheme = 'OnPremisesCredentialScheme'; OnPremisesUserName = 'OnPremisesUserName'
        AdminClientId = 'AdminClientId'
    }
    Test     = [ordered]@{
        DefaultType = 'TestType'; SignIn = 'SignIn'; AllowWrite = 'AllowWrite'; AdminChecks = 'AdminChecks'; FreeBusyDays = 'FreeBusyDays'
        FreeBusyIntervalMinutes = 'FreeBusyIntervalMinutes'; OAuthPollTimeoutSeconds = 'OAuthPollTimeoutSeconds'; HttpTimeoutSeconds = 'HttpTimeoutSeconds'
        CertificateWarningDays = 'CertificateWarningDays'; MaxRetries = 'MaxRetries'; UserAgent = 'UserAgent'; SkipDnsCrossCheck = 'SkipDnsCrossCheck'
        DnsServer = 'DnsServer'; DnsProbeName = 'DnsProbeName'
    }
    Report   = [ordered]@{ OutputPath = 'OutputPath'; FilePrefix = 'ReportPrefix'; Formats = 'ReportFormats'; CsvDelimiter = 'CsvDelimiter' }
    Logging  = [ordered]@{ Path = 'LogPath'; RetentionDays = 'LogRetentionDays' }
}

# Microsoft clouds: sign-in, Microsoft Graph, Exchange Online (EWS, Autodiscover, administration API).
$script:Clouds = [ordered]@{
    Global   = @{ Name = 'Microsoft 365 (worldwide)'; Login = 'login.microsoftonline.com'; Graph = 'graph.microsoft.com'; Outlook = 'outlook.office365.com'; Protection = 'mail.protection.outlook.com' }
    USGov    = @{ Name = 'Microsoft 365 GCC High'; Login = 'login.microsoftonline.us'; Graph = 'graph.microsoft.us'; Outlook = 'outlook.office365.us'; Protection = 'mail.protection.office365.us' }
    USGovDoD = @{ Name = 'Microsoft 365 DoD'; Login = 'login.microsoftonline.us'; Graph = 'dod-graph.microsoft.us'; Outlook = 'outlook-dod.office365.us'; Protection = 'mail.protection.office365.us' }
    China    = @{ Name = 'Microsoft 365 operated by 21Vianet'; Login = 'login.partner.microsoftonline.cn'; Graph = 'microsoftgraph.chinacloudapi.cn'; Outlook = 'partner.outlook.cn'; Protection = 'mail.protection.partner.outlook.cn' }
}
$script:EntraHosts = @('login.microsoftonline.com', 'login.windows.net', 'login.microsoft.com', 'sts.windows.net', 'login.microsoftonline.us', 'login.partner.microsoftonline.cn', 'login.chinacloudapi.cn')
$script:ExchangeAppId = '00000002-0000-0ff1-ce00-000000000000'
$script:EvoStsId = '00000001-0000-0000-c000-000000000000'
$script:GraphAppId = '00000003-0000-0000-c000-000000000000'
# Tenant Entra ID gives for the consumer domains (Outlook.com, Gmail...): the Microsoft accounts, not an organisation.
$script:ConsumerTenantId = '9cd80435-793b-4f48-844b-6b3f37d1c1f3'

# Public clients a user can sign in with, without registering an application.
#   MicrosoftOffice        first-party client of Microsoft, in every tenant, pre-authorized for Microsoft Graph and
#                          Exchange (calendars, mail, directory read): no consent, the scope is <resource>/.default.
#   GraphCommandLineTools  the client of Connect-MgGraph: explicit scopes, consent of the user or of an administrator.
#   ExchangeOnlinePowerShell  the client of Connect-ExchangeOnline (Identity.AdminClientId): only for the optional
#                          administrator checks (Test.AdminChecks), Exchange Online administration API.
$script:KnownClients = [ordered]@{
    MicrosoftOffice          = @{ Id = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'; Name = 'Microsoft Office'; Redirect = 'https://login.microsoftonline.com/common/oauth2/nativeclient'; Explicit = $false }
    GraphCommandLineTools    = @{ Id = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; Name = 'Microsoft Graph Command Line Tools'; Redirect = 'http://localhost'; Explicit = $true }
    ExchangeOnlinePowerShell = @{ Id = 'fb78d390-0c51-40cd-8e17-fdbfab77341b'; Name = 'Microsoft Exchange REST API Based Powershell'; Redirect = 'https://login.microsoftonline.com/organizations/oauth2/nativeclient'; Explicit = $false }
}
# Delegated permissions asked when the client needs explicit scopes (GraphCommandLineTools, or Identity.Scopes).
$script:DefaultScopes = @('User.Read', 'User.ReadBasic.All', 'Calendars.Read', 'Calendars.Read.Shared', 'MailboxSettings.Read', 'Mail.ReadBasic', 'Mail.Read', 'Domain.Read.All')
# Permissions that let getSchedule read free/busy (delegated scp or application roles).
$script:CalendarScopes = @('Calendars.ReadBasic', 'Calendars.Read', 'Calendars.ReadWrite', 'Calendars.Read.Shared', 'Calendars.ReadWrite.Shared', 'Calendars.ReadBasic.All', 'Calendars.Read.All', 'Calendars.ReadWrite.All')

# Tests and the stages they run, in order. A stage that fails stops the test: the next stages are Skipped.
# Writes: the test changes something (and needs AllowWrite for that part).
$script:Scenarios = @(
    @{ Name = 'FreeBusy'; DisplayName = 'Free/busy of other mailboxes'; Stages = @('Discovery', 'OnPremises', 'SignIn', 'Source', 'Target', 'FreeBusy', 'Hybrid')
       Description = 'Can the source mailbox see the free/busy of the target mailboxes? In both directions of a hybrid organisation and across organisations: where each mailbox is (Autodiscover, Entra ID), the on-premises endpoint (certificate, OAuth offered to Exchange Online), the sign-in of the source user, MailTips, then the free/busy itself - Microsoft Graph getSchedule from Exchange Online, EWS GetUserAvailability from on-premises - with the cause of each error. Optional: the hybrid configuration read by an Exchange administrator.' }
    @{ Name = 'MailboxServices'; DisplayName = 'Mailbox services (sync, availability, automatic replies)'; Stages = @('SignIn', 'Source', 'Sync', 'Availability', 'AutoReply'); Writes = $true
       Description = 'The services a mail client relies on, through Microsoft Graph: folder and message synchronization (delta), the own free/busy (getSchedule), MailTips, and the automatic replies - read, and with -AllowWrite set, read back and restored.' }
    @{ Name = 'Provisioning'; DisplayName = 'Mailbox provisioning'; Stages = @('Discovery', 'SignIn', 'Provisioning')
       Description = 'For each address (the target mailboxes, else the source): does a mailbox exist, where (Exchange Online or on-premises), of which type (user, shared, room, equipment), with which Exchange license and synchronization from Active Directory.' }
    @{ Name = 'DomainDns'; DisplayName = 'Domain DNS records'; Stages = @('Dns', 'SignIn', 'DomainRecords')
       Description = 'The DNS records of a domain for Exchange Online (MX, SPF, Autodiscover, DKIM, DMARC, MTA-STS, TLS-RPT), then, signed in, the records Microsoft 365 expects for this domain (domains/serviceConfigurationRecords) compared with the published ones.' }
    @{ Name = 'MessageHeaders'; DisplayName = 'Message headers (route, authentication, anti-spam)'; Stages = @('HeaderRead', 'HeaderDns', 'HeaderRoute')
       Description = 'How a message travelled and what Microsoft 365 decided, from its headers (an .eml or .txt file, or text pasted) - nothing signs in. The chronology from the sender to the mailbox: each hop with its time, the wait and who held the message (sender, gateway, Exchange on-premises, Microsoft 365), SPF, DKIM, DMARC and composite authentication, the spam verdict, the folder and the hybrid attribution, each finding with what to do. Then, optional, the DNS of today (SPF for the sending IP, DMARC, DKIM keys, MX), so that a finding already handled since is marked fixed.' }
    @{ Name = 'All'; DisplayName = 'All the tests'; Stages = @('Discovery', 'OnPremises', 'Dns', 'SignIn', 'Source', 'Target', 'FreeBusy', 'Hybrid', 'Sync', 'Availability', 'AutoReply', 'Provisioning', 'DomainRecords'); Writes = $true
       Description = 'Every test in one run, with one sign-in: free/busy, mailbox services (automatic replies only read unless -AllowWrite), provisioning and DNS.' }
)

$script:StageInfo = [ordered]@{
    Discovery     = @{ Title = 'Where are the mailboxes (Entra ID, Autodiscover)'; Icon = 'Search' }
    OnPremises    = @{ Title = 'On-premises endpoint (certificate, OAuth)'; Icon = 'Server' }
    Dns           = @{ Title = 'DNS records of the domain'; Icon = 'Search' }
    SignIn        = @{ Title = 'Sign-in'; Icon = 'Key' }
    Source        = @{ Title = 'Source mailbox'; Icon = 'People' }
    Target        = @{ Title = 'Target mailboxes (directory, MailTips)'; Icon = 'Target' }
    FreeBusy      = @{ Title = 'Free/busy'; Icon = 'Clock' }
    Hybrid        = @{ Title = 'Hybrid configuration (administrator)'; Icon = 'Settings' }
    Sync          = @{ Title = 'Synchronization (delta)'; Icon = 'Folder' }
    Availability  = @{ Title = 'Availability and MailTips of the mailbox'; Icon = 'Clock' }
    AutoReply     = @{ Title = 'Automatic replies'; Icon = 'Mail' }
    Provisioning  = @{ Title = 'Mailbox provisioning'; Icon = 'People' }
    DomainRecords = @{ Title = 'Records expected by Microsoft 365'; Icon = 'Shield' }
    HeaderRead    = @{ Title = 'Message headers'; Icon = 'Mail' }
    HeaderDns     = @{ Title = 'Is it still the case? (DNS of today)'; Icon = 'Search' }
    HeaderRoute   = @{ Title = 'Route of the message (chronology)'; Icon = 'Clock' }
}

function Get-GctTestCatalog {
    <# The tests: name, description, stages, whether they sign in and whether they can change something. #>
    [CmdletBinding()]
    param()

    foreach ($s in $script:Scenarios) {
        [pscustomobject]@{
            Name        = $s.Name
            DisplayName = $s.DisplayName
            Description = $s.Description
            Stages      = @($s.Stages)
            SignIn      = $s.Stages -contains 'SignIn'
            Writes      = [bool]($s.ContainsKey('Writes') -and $s.Writes)
        }
    }
}

function Get-GctScenarioStages {
    <#
        Stages a test runs with these settings: the on-premises checks only when not skipped, the
        administrator checks only when asked (Test.AdminChecks), the DNS of today unless skipped (Test.SkipDnsCrossCheck).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$TestType, [hashtable]$Configuration = @{})

    $scenario = $script:Scenarios | Where-Object { $_.Name -eq $TestType } | Select-Object -First 1
    if (-not $scenario) { return }
    foreach ($stage in $scenario.Stages) {
        if ($stage -eq 'OnPremises' -and $Configuration['SkipOnPremisesChecks']) { continue }
        if ($stage -eq 'Hybrid' -and -not $Configuration['AdminChecks']) { continue }
        if ($stage -eq 'HeaderDns' -and $Configuration['SkipDnsCrossCheck']) { continue }
        $stage
    }
}

function Get-GctStageTitle {
    param([Parameter(Mandatory = $true)][string]$Stage, [hashtable]$Configuration = @{})
    if ($Stage -eq 'SignIn') {
        if ([string]$Configuration['Context'] -eq 'Application') { return 'Sign-in of the application (certificate)' }
        return 'Sign-in of the source user'
    }
    return $script:StageInfo[$Stage].Title
}

function Get-GctDefaultConfiguration {
    @{
        SourceMailbox            = 'alice@contoso.com'
        TargetMailboxes          = @()
        OnPremisesHost           = ''
        SkipOnPremisesChecks     = $false
        Domain                   = ''
        Cloud                    = 'Global'
        HeaderFile               = ''
        HeaderText               = ''
        Context                  = 'User'
        ClientId                 = 'MicrosoftOffice'
        RedirectUri              = ''
        Scopes                   = @()
        TenantId                 = ''
        CertificateThumbprint    = ''
        OnPremisesAuthentication = 'OAuth'
        OnPremisesCredentialScheme = 'Auto'
        OnPremisesUserName       = ''
        AdminClientId            = 'ExchangeOnlinePowerShell'
        TestType                 = 'FreeBusy'
        SignIn                   = 'Auto'
        AllowWrite               = $false
        AdminChecks              = $false
        FreeBusyDays             = 7
        FreeBusyIntervalMinutes  = 30
        OAuthPollTimeoutSeconds  = 600
        HttpTimeoutSeconds       = 60
        CertificateWarningDays   = 30
        MaxRetries               = 2
        UserAgent                = 'GraphConnectivityTests/1.0'
        SkipDnsCrossCheck        = $false
        DnsServer                = @()
        DnsProbeName             = 'microsoft.com'
        OutputPath               = Join-Path $script:ToolRoot 'reports'
        ReportPrefix             = 'GraphConnectivityTests'
        ReportFormats            = @('Csv', 'Html')
        CsvDelimiter             = ';'
        LogPath                  = Join-Path $script:ToolRoot 'logs'
        LogRetentionDays         = 14
    }
}

function Resolve-GctPath {
    param([AllowEmptyString()][string]$Path, [Parameter(Mandatory = $true)][string]$Root)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ([IO.Path]::IsPathRooted($Path)) { return $Path }
    return [IO.Path]::GetFullPath((Join-Path $Root $Path))
}

function Resolve-GctClient {
    <#
        The client of a sign-in: a known client by name (MicrosoftOffice, GraphCommandLineTools,
        ExchangeOnlinePowerShell) or an application ID. Returns Id, Name, RedirectUri and whether it
        needs explicit scopes.
    #>
    param([Parameter(Mandatory = $true)][string]$ClientId, [string]$RedirectUri)

    $known = $null
    foreach ($key in $script:KnownClients.Keys) {
        $c = $script:KnownClients[$key]
        if ($ClientId -ieq $key -or $ClientId -ieq $c.Id) { $known = $c; $known.Key = $key; break }
    }
    if ($known) {
        return [pscustomobject]@{ Id = $known.Id; Name = $known.Name; Key = $known.Key; RedirectUri = $(if ($RedirectUri) { $RedirectUri } else { $known.Redirect }); Explicit = [bool]$known.Explicit; Known = $true }
    }
    [pscustomobject]@{ Id = $ClientId; Name = "application $ClientId"; Key = 'Custom'; RedirectUri = $(if ($RedirectUri) { $RedirectUri } else { 'https://login.microsoftonline.com/common/oauth2/nativeclient' }); Explicit = $false; Known = $false }
}

function Get-GctAddressDomain {
    param([AllowEmptyString()][string]$Address)
    if ($Address -match '@([^@\s]+)$') { return $Matches[1].ToLowerInvariant() }
    return $null
}

function Test-GctConfiguration {
    <# Checks the settings; every problem is listed. Returns IsValid and Problems. #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [switch]$File)

    $c = $Configuration
    $problems = [Collections.Generic.List[string]]::new()
    $address = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    $guid = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    $names = @($script:Scenarios | ForEach-Object { $_.Name })
    if ([string]$c.TestType -notin $names) { $problems.Add("Test.DefaultType must be one of: $($names -join ', ').") }
    if ([string]$c.TestType -ne 'MessageHeaders' -and [string]$c.SourceMailbox -notmatch $address) { $problems.Add("Target.SourceMailbox must be an SMTP address (the mailbox that signs in and asks for the free/busy): '$($c.SourceMailbox)'.") }
    foreach ($t in @($c.TargetMailboxes | Where-Object { $_ })) { if ([string]$t -notmatch $address) { $problems.Add("Target.TargetMailboxes: '$t' is not an SMTP address.") } }
    if (@($c.TargetMailboxes | Where-Object { $_ }).Count -gt 20) { $problems.Add('Target.TargetMailboxes: 20 mailboxes at most in one run (limit of getSchedule).') }
    if (-not $File -and [string]$c.TestType -eq 'FreeBusy' -and -not @($c.TargetMailboxes | Where-Object { $_ }).Count) { $problems.Add('The FreeBusy test needs at least one target mailbox (Target.TargetMailboxes, -TargetMailbox).') }
    if ([string]$c.OnPremisesHost -and [string]$c.OnPremisesHost -notmatch '^[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$') {
        $problems.Add("Target.OnPremisesHost must be a host name (FQDN of the Availability service, for example mail.contoso.com, without https:// nor /EWS/Exchange.asmx): '$($c.OnPremisesHost)'.")
    }
    if ([string]$c.Domain -and [string]$c.Domain -notmatch '^[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$') { $problems.Add("Target.Domain must be a domain name: '$($c.Domain)'.") }
    if ([string]$c.HeaderText -and [string]$c.HeaderFile) { $problems.Add('Target.HeaderFile and Target.HeaderText: give the headers one way only (a file, or the text).') }
    if ([string]$c.HeaderFile -and -not [string]$c.HeaderText -and -not (Test-Path -LiteralPath ([string]$c.HeaderFile) -PathType Leaf)) { $problems.Add("Target.HeaderFile: file not found: '$($c.HeaderFile)'.") }
    if (([string]$c.HeaderText).Length -gt 2MB) { $problems.Add('Target.HeaderText: more than 2 MB of text - paste the headers only, not the whole message.') }
    if (-not $File -and [string]$c.TestType -eq 'MessageHeaders' -and -not ([string]$c.HeaderText).Trim() -and -not [string]$c.HeaderFile) { $problems.Add('The MessageHeaders test needs the headers of a message: a file (Target.HeaderFile, -HeaderFile, .eml or .txt) or the text (Target.HeaderText, -HeaderText).') }
    $dnsName = '^[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+\.?$'
    foreach ($s in @($c.DnsServer | Where-Object { $_ })) {
        $ip = $null
        if (-not [Net.IPAddress]::TryParse([string]$s, [ref]$ip) -and [string]$s -notmatch $dnsName) { $problems.Add("Test.DnsServer: '$s' is not an IP address nor a host name (leave it empty for the DNS servers of Windows).") }
    }
    if ([string]$c.DnsProbeName -notmatch $dnsName) { $problems.Add("Test.DnsProbeName must be a DNS name that always answers (default microsoft.com): '$($c.DnsProbeName)'.") }
    if ([string]$c.Cloud -notin @($script:Clouds.Keys)) { $problems.Add("Target.Cloud must be one of: $(@($script:Clouds.Keys) -join ', ').") }
    if ([string]$c.Context -notin 'User', 'Application') { $problems.Add("Identity.Context must be 'User' or 'Application'.") }
    if ([string]$c.ClientId -notmatch $guid -and [string]$c.ClientId -notin @($script:KnownClients.Keys)) { $problems.Add("Identity.ClientId must be $(@($script:KnownClients.Keys) -join ', ') or an application ID (GUID): '$($c.ClientId)'.") }
    if ([string]$c.AdminClientId -notmatch $guid -and [string]$c.AdminClientId -notin @($script:KnownClients.Keys)) { $problems.Add("Identity.AdminClientId must be ExchangeOnlinePowerShell or an application ID (GUID).") }
    if ([string]$c.RedirectUri -and [string]$c.RedirectUri -notmatch '^(https?|ms-appx-web|msal[0-9a-f-]+)://') { $problems.Add("Identity.RedirectUri must be an absolute URI: '$($c.RedirectUri)'.") }
    if ([string]$c.Context -eq 'Application') {
        if ([string]$c.ClientId -notmatch $guid) { $problems.Add('The Application context needs the application ID of your application (Identity.ClientId, a GUID).') }
        if ([string]$c.CertificateThumbprint -notmatch '^[0-9A-Fa-f]{40}$') { $problems.Add('The Application context needs the thumbprint of the certificate of the application (Identity.CertificateThumbprint, 40 hexadecimal characters).') }
        if (-not [string]$c.TenantId) { $problems.Add('The Application context needs the tenant (Identity.TenantId: tenant ID or domain).') }
        if ([string]$c.TestType -eq 'MailboxServices' -and $c.AllowWrite) { $problems.Add('The automatic replies are changed only in the User context (delegated MailboxSettings.ReadWrite).') }
    }
    if ([string]$c.TenantId -and [string]$c.TenantId -notmatch $guid -and [string]$c.TenantId -notmatch '^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') { $problems.Add("Identity.TenantId must be a tenant ID or a domain: '$($c.TenantId)'.") }
    if ([string]$c.OnPremisesAuthentication -notin 'OAuth', 'Windows', 'Credential') { $problems.Add("Identity.OnPremisesAuthentication must be 'OAuth' (hybrid modern authentication), 'Windows' (current Windows account) or 'Credential' (an account and its password typed).") }
    if ([string]$c.OnPremisesCredentialScheme -notin 'Auto', 'Negotiate', 'NTLM', 'Basic') { $problems.Add("Identity.OnPremisesCredentialScheme must be 'Auto', 'Negotiate', 'NTLM' or 'Basic'.") }
    if ([string]$c.OnPremisesUserName -and [string]$c.OnPremisesUserName -notmatch '^([^\\@\s]+\\[^\\@\s]+|[^@\s\\]+@[^@\s]+\.[^@\s]+)$') { $problems.Add("Identity.OnPremisesUserName must be DOMAIN\user or user@domain: '$($c.OnPremisesUserName)'.") }
    if ([string]$c.SignIn -notin 'Auto', 'Window', 'DeviceCode') { $problems.Add("Test.SignIn must be 'Auto', 'Window' or 'DeviceCode'.") }
    foreach ($pair in @(@('FreeBusyDays', 1, 62), @('FreeBusyIntervalMinutes', 5, 1440), @('OAuthPollTimeoutSeconds', 30, 3600), @('HttpTimeoutSeconds', 5, 600), @('CertificateWarningDays', 0, 365), @('MaxRetries', 0, 10), @('LogRetentionDays', 1, 3650))) {
        $v = 0
        if (-not [int]::TryParse([string]$c[$pair[0]], [ref]$v) -or $v -lt $pair[1] -or $v -gt $pair[2]) { $problems.Add("$($pair[0]) must be a number from $($pair[1]) to $($pair[2]).") }
    }
    foreach ($f in @($c.ReportFormats)) { if ($f -notin 'Csv', 'Html') { $problems.Add("Report.Formats: '$f' is not Csv or Html.") } }
    if ([string]$c.CsvDelimiter -notin ';', ',', "`t") { $problems.Add('Report.CsvDelimiter must be ; , or a tab.') }
    [pscustomobject]@{ IsValid = $problems.Count -eq 0; Problems = @($problems) }
}

function Import-GctConfiguration {
    <#
    .SYNOPSIS
        Reads the configuration file (sections Target, Identity, Test, Report, Logging) into one settings hashtable.
    #>
    [CmdletBinding()]
    param([string]$Path = (Join-Path $script:ToolRoot 'config\GraphConnectivityTests.config.psd1'), [string]$Root = $script:ToolRoot)

    if (-not (Test-Path -LiteralPath $Path)) { throw "Configuration file not found: $Path" }
    $data = Import-PowerShellDataFile -LiteralPath $Path
    $settings = Get-GctDefaultConfiguration
    $problems = [Collections.Generic.List[string]]::new()
    foreach ($section in $data.Keys) {
        if (-not $script:ConfigSchema.Contains($section)) { $problems.Add("Unknown section '$section'."); continue }
        $values = $data[$section]
        if ($values -isnot [hashtable]) { $problems.Add("Section '$section' must be @{ ... }."); continue }
        foreach ($key in $values.Keys) {
            if (-not $script:ConfigSchema[$section].Contains($key)) { $problems.Add("Unknown key '$section.$key'."); continue }
            $settings[$script:ConfigSchema[$section][$key]] = $values[$key]
        }
    }
    $settings.TargetMailboxes = @($settings.TargetMailboxes | Where-Object { $_ } | ForEach-Object { ([string]$_).Trim() })
    $settings.Scopes = @($settings.Scopes | Where-Object { $_ })
    $settings.ReportFormats = @($settings.ReportFormats)
    $settings.OutputPath = Resolve-GctPath -Path ([string]$settings.OutputPath) -Root $Root
    $settings.LogPath = Resolve-GctPath -Path ([string]$settings.LogPath) -Root $Root
    $settings.HeaderFile = Resolve-GctPath -Path ([string]$settings.HeaderFile) -Root $Root
    $check = Test-GctConfiguration -Configuration $settings -File
    foreach ($p in $check.Problems) { $problems.Add($p) }
    if ($problems.Count) { throw ("Configuration $Path`n - " + ($problems -join "`n - ")) }
    return $settings
}

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDATnkuk5tZAkLO
# 0Z5OsLVNWM5YqRQXxazfs6fFOXHl9aCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA5xjhJgz8Y
# NJZud6emObMZq7uiSnyfckaFCZSj1eAGMTANBgkqhkiG9w0BAQEFAASCAgCacPy2
# x+y8nXasFzc+BYoFKqGsj8gUQ46N/ZaPvtG3ps2IqIW48sMRIRvff2Z2MGDA0Xsm
# Se/zENnu8csDxA2xFmCrlKWRagmfzehDXDsnalPzpOhYWpdXcIrn7feSNyHlrKQm
# 0jx9jOJLByggxs88eeVTkl8UcP/+w5g0eE9Mhtv/WCVZ8C+TH8pIBNl8PNuNiCFE
# ruTD2q0UP+n997vYO0PrGmN4Op3EOy5dhXMXj8ZuuN9kPh3iidShdzWploBWResJ
# P8cC4auGDAShOaWrvWi6PkUIgiWb6LpAbMkI6A3lE+JcjYgotJDthm8Moz050DOA
# W4lliXobu7FCRog4UkkmclUWoyse4gqp3SNurmMyEMyqjuXcLsR0lDNKbgfhjBeq
# L1lD2d62szWT0WI6T/GWwEI0l4D1bQ346UCRUFjZdxh6kdF/paeHaF2Wwa/8624J
# 2gO4W8pw3IN+IkpOquDoKIZTrZa528wSabDLTgyB0aIQNFTNhXliVikUUsNdJcwh
# xS91O2qC7C3SgQxS28QCTVrkLgg+K7FgS9gDCUCnGbJVticDT8BHpvFeA6IHkx2k
# OJzPE+8XLrwQgZS1uJTNngDMawloRAxw/swygPbNbWSKeVcNy9T43/a1urfbHmxd
# E819Y4pVTJoPCJg8KsCVpAPdHMEpiuZdUjXbgKGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIOKkXkWqsHr+VnzbqFGdmHXXw5ZdbnaoUbXqJv7LxC83AghpyygLmdJB
# uhgPMjAyNjExMTAxODI1NDRaMAMCAQECEG/IKai8OK18fUXwUdZCwLSgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDE4MjU0NFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIBXg
# W76mQwaEqLQPHMQmXamqsrPxSICLrZUsI3R051DXMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiBwyxrgav0iwpKo1obO8Ahh0vIyQsoeg7WKWpeYsCPdlQIhAKiV
# Sxh8zhnksWUiaDaCTLs+ri6ezX55WCcbRY2JzwHo
# SIG # End signature block
