<#
.SYNOPSIS
    Graph Connectivity Tests - Microsoft Graph connectivity tests in the user context: free/busy, mailbox services, provisioning, DNS, message headers.

.DESCRIPTION
    Runs one test and tells exactly which step works and which one does not, with the request sent and the
    response received for each check:

      FreeBusy         can the source mailbox see the free/busy of the target mailboxes? Where each mailbox is
                       (Entra ID, Autodiscover), the on-premises endpoint (DNS, certificate, OAuth offered to
                       Exchange Online), the sign-in of the source user, the targets in the directory and their
                       MailTips, then the free/busy - Microsoft Graph getSchedule from Exchange Online, EWS
                       GetUserAvailability from a source on-premises - with the cause of each error.
                       -AdminChecks adds the hybrid configuration read by an Exchange administrator.
      MailboxServices  synchronization (delta), availability, MailTips and automatic replies of the source
                       (scheduled, read back and restored with -AllowWrite)
      Provisioning     does each address have a mailbox, where, of which type, with which license
      DomainDns        MX, SPF, Autodiscover, DKIM, DMARC, MTA-STS, TLS-RPT of a domain, and the records
                       Microsoft 365 expects for it
      MessageHeaders   the headers of a message (an .eml or .txt file, or text): the chronology from the sender to
                       the mailbox with the wait at each hop, SPF, DKIM, DMARC, the spam verdict, the hybrid
                       attribution and gateways, each finding with what to do; then the DNS of today, so that a
                       finding already handled since is marked fixed. Nothing signs in.
      All              every test, one sign-in (not MessageHeaders)

    No application to register: the source user signs in with Microsoft Office, a public client of Microsoft
    pre-authorized for Microsoft Graph in every tenant (sign-in window or device code, MFA and Conditional
    Access included). -ClientId names another client, or your multi-tenant application; -Context Application
    runs with your application and a certificate.

    Writes CSV, JSON and HTML report files in a new folder, and a daily log file. Everything is set in
    config\GraphConnectivityTests.config.psd1; the parameters below override it. Tokens stay in memory and
    are never written anywhere.

.PARAMETER TestType
    FreeBusy, MailboxServices, Provisioning, DomainDns, MessageHeaders or All. Default: Test.DefaultType.
    Giving -HeaderFile or -HeaderText alone selects MessageHeaders.
.PARAMETER SourceMailbox
    The mailbox that signs in and asks for the free/busy (Target.SourceMailbox).
.PARAMETER TargetMailbox
    The mailboxes whose free/busy is asked, 20 at most (Target.TargetMailboxes).
.PARAMETER OnPremisesHost
    FQDN of the on-premises Availability service, when Autodiscover does not give it (Target.OnPremisesHost).
.PARAMETER SkipOnPremisesChecks
    Skips the checks of the on-premises endpoint (Target.SkipOnPremisesChecks).
.PARAMETER Domain
    DomainDns: the domain to check (default: the domain of the source).
.PARAMETER HeaderFile
    MessageHeaders: an .eml file, or a .txt file with the headers of the message (Target.HeaderFile). Only the header
    block is read; the body and the attachments are not.
.PARAMETER HeaderText
    MessageHeaders: the headers of the message as text (Target.HeaderText), for example (Get-Clipboard -Raw).
.PARAMETER SkipDnsCrossCheck
    MessageHeaders: does not compare the header with the DNS of today (Test.SkipDnsCrossCheck): no DNS query at all.
.PARAMETER DnsServer
    The DNS servers asked (IP addresses or names, Test.DnsServer) instead of those of Windows - DomainDns and the
    DNS of today of MessageHeaders. When no DNS server answers, MessageHeaders reads the header offline.
.PARAMETER Cloud
    Global, USGov, USGovDoD or China (Target.Cloud).
.PARAMETER Context
    User (the source user signs in) or Application (your application with a certificate) (Identity.Context).
.PARAMETER ClientId
    MicrosoftOffice (default), GraphCommandLineTools or an application ID (Identity.ClientId).
.PARAMETER RedirectUri
    Redirect URI of your application (Identity.RedirectUri).
.PARAMETER TenantId
    Tenant ID or domain (Identity.TenantId). Needed for the Application context.
.PARAMETER CertificateThumbprint
    Application context: certificate of the application (Identity.CertificateThumbprint).
.PARAMETER OnPremisesAuthentication
    A source mailbox on-premises: OAuth (hybrid modern authentication), Windows (the current Windows account) or
    Credential (an account and its password: NTLM, Negotiate or Basic, from any computer) (Identity.OnPremisesAuthentication).
.PARAMETER OnPremisesCredentialScheme
    Credential: Auto (the best scheme EWS offers: Negotiate, NTLM, then Basic), Negotiate, NTLM or Basic (Identity.OnPremisesCredentialScheme).
.PARAMETER OnPremisesCredential
    Credential: the on-premises account (Get-Credential). Passing it selects Credential. Without it, the account is
    asked when the source mailbox turns out to be on-premises. The password is never written anywhere.
.PARAMETER SignIn
    Auto, Window or DeviceCode (Test.SignIn).
.PARAMETER AllowWrite
    MailboxServices: schedules the automatic replies in the future, reads them back and restores them.
.PARAMETER AdminChecks
    FreeBusy: also reads the hybrid configuration of Exchange Online as an Exchange administrator.
.PARAMETER FreeBusyDays
    Number of days from today (Test.FreeBusyDays).
.PARAMETER Gui
    Opens the window: same tests, same report, progress shown live.
.PARAMETER OutputPath
    Overrides Report.OutputPath.
.PARAMETER NoReport
    No report file (console and log only).
.PARAMETER AccessToken
    A Microsoft Graph token already obtained (integration only): the sign-in is skipped. It can stay in the PowerShell history.
.PARAMETER ConfigPath
    Configuration file. Default: config\GraphConnectivityTests.config.psd1.

.EXAMPLE
    .\Invoke-GraphConnectivityTests.ps1 -SourceMailbox alice@contoso.com -TargetMailbox bob@contoso.com
    Free/busy of Bob seen by Alice: Alice signs in (Microsoft Office, no application to register).

.EXAMPLE
    .\Invoke-GraphConnectivityTests.ps1 -SourceMailbox alice@contoso.com -TargetMailbox onprem.user@contoso.com -OnPremisesHost mail.contoso.com -AdminChecks
    Hybrid, Microsoft 365 to on-premises: the on-premises endpoint, the free/busy, and the IntraOrganizationConnector read by an administrator.

.EXAMPLE
    .\Invoke-GraphConnectivityTests.ps1 -SourceMailbox onprem.user@contoso.com -TargetMailbox alice@contoso.com
    Hybrid, on-premises to Microsoft 365: the request goes to EWS on-premises with the token of hybrid modern authentication.

.EXAMPLE
    .\Invoke-GraphConnectivityTests.ps1 -SourceMailbox onprem.user@contoso.com -TargetMailbox alice@contoso.com -OnPremisesCredential (Get-Credential CONTOSO\onprem.user) -OnPremisesCredentialScheme Basic
    The same without hybrid modern authentication: EWS on-premises called with the account typed, in Basic (NTLM with -OnPremisesCredentialScheme NTLM).

.EXAMPLE
    .\Invoke-GraphConnectivityTests.ps1 -TestType DomainDns -Domain contoso.com -SourceMailbox alice@contoso.com
    The DNS records of contoso.com, then the records Microsoft 365 expects for it.

.EXAMPLE
    .\Invoke-GraphConnectivityTests.ps1 -HeaderFile .\message.eml
    The route of the message, its authentication and anti-spam verdict, then the DNS of today (SPF for the sending IP,
    DMARC, DKIM keys, MX). Nothing signs in.

.EXAMPLE
    .\Invoke-GraphConnectivityTests.ps1 -HeaderText (Get-Clipboard -Raw) -SkipDnsCrossCheck
    The headers copied from Outlook (File > Properties > Internet headers), read without any DNS query.

.EXAMPLE
    .\Invoke-GraphConnectivityTests.ps1 -SourceMailbox room-checker@contoso.com -TargetMailbox room1@contoso.com -Context Application -ClientId 11111111-2222-3333-4444-555555555555 -CertificateThumbprint 0123456789ABCDEF0123456789ABCDEF01234567 -TenantId contoso.com
    The same free/busy test as an application (Calendars.Read), for a scheduled task.

.EXAMPLE
    .\Invoke-GraphConnectivityTests.ps1 -Gui

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
    Exit codes : 0 = Passed, 1 = Failed, 2 = finished with warnings or blocked.
    Documentation : docs\GraphConnectivityTests-UserGuide.html, docs\GraphConnectivityTests-Guide.html
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [ValidateSet('FreeBusy', 'MailboxServices', 'Provisioning', 'DomainDns', 'MessageHeaders', 'All')][string]$TestType,
    [string]$SourceMailbox,
    [string[]]$TargetMailbox,
    [string]$OnPremisesHost,
    [switch]$SkipOnPremisesChecks,
    [string]$Domain,
    [string]$HeaderFile,
    [string]$HeaderText,
    [switch]$SkipDnsCrossCheck,
    [string[]]$DnsServer,
    [ValidateSet('Global', 'USGov', 'USGovDoD', 'China')][string]$Cloud,
    [ValidateSet('User', 'Application')][string]$Context,
    [string]$ClientId,
    [string]$RedirectUri,
    [string]$TenantId,
    [string]$CertificateThumbprint,
    [ValidateSet('OAuth', 'Windows', 'Credential')][string]$OnPremisesAuthentication,
    [ValidateSet('Auto', 'Negotiate', 'NTLM', 'Basic')][string]$OnPremisesCredentialScheme,
    [pscredential]$OnPremisesCredential,
    [ValidateSet('Auto', 'Window', 'DeviceCode')][string]$SignIn,
    [switch]$AllowWrite,
    [switch]$AdminChecks,
    [int]$FreeBusyDays,
    [switch]$Gui,
    [string]$OutputPath,
    [switch]$NoReport,
    [string]$AccessToken,
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\GraphConnectivityTests.config.psd1')
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$exitCode = 1
$moduleLoaded = $false

try {
    Import-Module (Join-Path $PSScriptRoot 'GraphConnectivityTests.psd1') -Force
    $moduleLoaded = $true

    $settings = Import-GctConfiguration -Path $ConfigPath -Root $PSScriptRoot
    foreach ($name in 'TestType', 'SourceMailbox', 'OnPremisesHost', 'Domain', 'Cloud', 'Context', 'ClientId', 'RedirectUri', 'TenantId', 'CertificateThumbprint', 'OnPremisesAuthentication', 'OnPremisesCredentialScheme', 'SignIn') {
        if ($PSBoundParameters.ContainsKey($name)) { $settings[$name] = [string]$PSBoundParameters[$name] }
    }
    if ($PSBoundParameters.ContainsKey('TargetMailbox')) { $settings.TargetMailboxes = @($TargetMailbox | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ }) }
    if ($PSBoundParameters.ContainsKey('HeaderFile')) { $settings.HeaderFile = [IO.Path]::GetFullPath($HeaderFile, (Get-Location).Path); $settings.HeaderText = '' }
    if ($PSBoundParameters.ContainsKey('HeaderText')) { $settings.HeaderText = $HeaderText; if (-not $PSBoundParameters.ContainsKey('HeaderFile')) { $settings.HeaderFile = '' } }
    if (($PSBoundParameters.ContainsKey('HeaderFile') -or $PSBoundParameters.ContainsKey('HeaderText')) -and -not $PSBoundParameters.ContainsKey('TestType')) { $settings.TestType = 'MessageHeaders' }
    foreach ($switch in 'SkipOnPremisesChecks', 'AllowWrite', 'AdminChecks', 'SkipDnsCrossCheck') { if ($PSBoundParameters.ContainsKey($switch)) { $settings[$switch] = [bool]$PSBoundParameters[$switch] } }
    if ($PSBoundParameters.ContainsKey('DnsServer')) { $settings.DnsServer = @($DnsServer | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ }) }
    if ($PSBoundParameters.ContainsKey('FreeBusyDays')) { $settings.FreeBusyDays = $FreeBusyDays }
    if ($OnPremisesCredential -and -not $PSBoundParameters.ContainsKey('OnPremisesAuthentication')) { $settings.OnPremisesAuthentication = 'Credential' }
    if ($OutputPath) { $settings.OutputPath = [IO.Path]::GetFullPath($OutputPath, (Get-Location).Path) }

    $logPath = Start-GctLog -Directory $settings.LogPath -RetentionDays $settings.LogRetentionDays
    if ($Gui) {
        Write-GctLog 'STEP' '=== Graph Connectivity Tests - window opened ==='
        Show-GctTestGui -Configuration $settings
        $exitCode = 0
    }
    else {
        $check = Test-GctConfiguration -Configuration $settings
        if (-not $check.IsValid) { throw ("Invalid value:`n - " + ($check.Problems -join "`n - ")) }
        Write-GctRunBanner -Settings $settings -LogPath $logPath -NoReport:$NoReport
        if ($AccessToken) { Write-GctItem Warn 'An access token was passed on the command line: it can stay in the PowerShell history.' }
        $result = Invoke-GctConnectivityTest -Configuration $settings -TestType $settings.TestType -AccessToken $AccessToken -OnPremisesCredential $OnPremisesCredential
        $reportText = 'none (-NoReport)'
        if (-not $NoReport) {
            $report = Export-GctReport -Result $result -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Formats $settings.ReportFormats -Delimiter $settings.CsvDelimiter
            $reportText = if ($report.Files.Contains('Html')) { $report.Files.Html } else { $report.Directory }
        }
        Write-GctRunSummary -Result $result -ReportText $reportText -LogPath $logPath
        $exitCode = switch ($result.Status) { 'Passed' { 0 } 'Failed' { 1 } default { 2 } }
    }
}
catch {
    if ($moduleLoaded) { Write-GctItem Fail $_.Exception.Message; Write-Host '' }
    else { Write-Host "Graph Connectivity Tests: $($_.Exception.Message)" -ForegroundColor Red }
    $exitCode = 1
}
finally {
    if ($moduleLoaded) { Stop-GctLog }
    [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}
exit $exitCode
