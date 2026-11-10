<#
.SYNOPSIS
    Graph Connectivity Tests - PowerShell module.

.DESCRIPTION
    Loads the parts of the tool, in the order of an execution:

        src\GraphConnectivityTests.Console.ps1       console output and log file
        src\GraphConnectivityTests.Config.ps1        configuration, clouds, clients and tests
        src\GraphConnectivityTests.Http.ps1          HTTP requests and their trace (request sent, response received)
        src\GraphConnectivityTests.OAuth.ps1         helpers and OAuth building blocks shared with Web Services Client for Exchange
        src\GraphConnectivityTests.Browser.ps1       sign-in window (Edge or Chrome, DevTools protocol)
        src\GraphConnectivityTests.Auth.ps1          sign-in and tokens: user (no application to register), application, administrator
        src\GraphConnectivityTests.Discovery.ps1     where the mailboxes are (Entra ID, Autodiscover), DNS, the on-premises endpoint
        src\GraphConnectivityTests.Graph.ps1         Microsoft Graph requests, free/busy answers and the causes of their errors
        src\GraphConnectivityTests.Ews.ps1           EWS GetUserAvailability for a source mailbox on-premises
        src\GraphConnectivityTests.FreeBusy.ps1      the free/busy test: source, targets and MailTips, free/busy
        src\GraphConnectivityTests.ExchangeAdmin.ps1 optional: hybrid and sharing configuration of Exchange Online (administrator)
        src\GraphConnectivityTests.Mailbox.ps1       mailbox services (sync, availability, automatic replies) and provisioning
        src\GraphConnectivityTests.Dns.ps1           DNS records of a domain, and the records Microsoft 365 expects
        src\GraphConnectivityTests.Headers.ps1       message headers: route, authentication, anti-spam, and the DNS of today
        src\GraphConnectivityTests.Run.ps1           steps and Invoke-GctConnectivityTest
        src\GraphConnectivityTests.Report.ps1        CSV, JSON and HTML report
        src\GraphConnectivityTests.Gui.ps1           WPF window (Fluent theme of Windows 11)

    Access, refresh and ID tokens stay in memory: they are never written to the console, the log or the report.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
    History : see CHANGELOG.md
#>
#Requires -Version 7.4
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http

$script:ToolVersion = '1.1.0'
$script:ToolRoot = $PSScriptRoot
$script:GctUserAgent = 'GraphConnectivityTests/1.0'
$script:LogWriter = $null
$script:LogPath = $null
$script:Quiet = $false
# Window hooks, set only while the window runs a test: Sink (progress lines), Pump (keeps the window responsive), Cancel, DeviceCode.
$script:Ui = $null

foreach ($part in 'Console', 'Config', 'Http', 'OAuth', 'Browser', 'Auth', 'Discovery', 'Graph', 'Ews', 'FreeBusy', 'ExchangeAdmin', 'Mailbox', 'Dns', 'Headers', 'Run', 'Report', 'Gui') {
    . (Join-Path $PSScriptRoot "src\GraphConnectivityTests.$part.ps1")
}

Export-ModuleMember -Function @(
    'Import-GctConfiguration', 'Test-GctConfiguration', 'Get-GctTestCatalog', 'Get-GctScenarioStages'
    'Invoke-GctConnectivityTest', 'Export-GctReport', 'Show-GctTestGui'
    'Start-GctLog', 'Stop-GctLog', 'Write-GctLog', 'Write-GctBanner', 'Write-GctStep', 'Write-GctItem', 'Write-GctSummary', 'Write-GctRunBanner', 'Write-GctRunSummary', 'Format-GctDuration'
)
