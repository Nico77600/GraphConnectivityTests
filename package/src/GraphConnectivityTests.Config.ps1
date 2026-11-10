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
