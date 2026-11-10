#
#  Graph Connectivity Tests - module manifest
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : see ModuleVersion
#
#  Loaded by Invoke-GraphConnectivityTests.ps1 (Import-Module by path).
#
@{
    RootModule        = 'GraphConnectivityTests.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = '5f7c2a9e-3d41-4b8a-9c6e-2a1f0b7d4e93'
    Author            = 'Nicolas Fabert'
    Copyright         = '(c) 2026 Nicolas Fabert. MIT License.'
    Description       = 'Graph Connectivity Tests: Microsoft Graph connectivity tests in the user context (no application to register) - can a mailbox see the free/busy of others, in both directions of a hybrid organisation and across organisations, with the cause of each Availability error; plus mailbox services (delta sync, availability, automatic replies), mailbox provisioning, the DNS records of a domain, and the analysis of message headers (route, authentication, anti-spam, compared with the DNS of today) - every request and response traced, CSV, JSON and HTML reports, a console and a window.'
    PowerShellVersion = '7.4'

    FunctionsToExport = @(
        'Import-GctConfiguration', 'Test-GctConfiguration', 'Get-GctTestCatalog', 'Get-GctScenarioStages'
        'Invoke-GctConnectivityTest', 'Export-GctReport', 'Show-GctTestGui'
        'Start-GctLog', 'Stop-GctLog', 'Write-GctLog', 'Write-GctBanner', 'Write-GctStep', 'Write-GctItem', 'Write-GctSummary', 'Write-GctRunBanner', 'Write-GctRunSummary', 'Format-GctDuration'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('Exchange', 'ExchangeOnline', 'MicrosoftGraph', 'FreeBusy', 'Hybrid', 'OAuth', 'EntraID', 'Autodiscover', 'DNS', 'Diagnostic', 'MessageHeaders', 'SPF', 'DKIM', 'DMARC')
            LicenseUri = 'https://opensource.org/licenses/MIT'
        }
    }
}
