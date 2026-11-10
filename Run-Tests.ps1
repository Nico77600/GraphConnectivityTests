#Requires -Version 7.4
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Import-Module Pester -MinimumVersion 6.1.0
$configuration = New-PesterConfiguration
$configuration.Run.Path = Join-Path $PSScriptRoot 'tests\GraphConnectivityTests.Tests.ps1'
$configuration.Run.PassThru = $true
$configuration.Output.Verbosity = 'Normal'
$configuration.TestDrive.Enabled = $false
$configuration.TestRegistry.Enabled = $false
$configuration.TestResult.Enabled = $true
$configuration.TestResult.OutputPath = Join-Path $PSScriptRoot 'artifacts\pester-results.xml'
[void][IO.Directory]::CreateDirectory((Join-Path $PSScriptRoot 'artifacts'))
$result = Invoke-Pester -Configuration $configuration
if ($result.FailedCount -gt 0 -or $result.Result -ne 'Passed') {
    throw "Tests failed: $($result.FailedCount)."
}
