<#
.SYNOPSIS
    Graph Connectivity Tests - report files (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    One folder per execution, <FilePrefix>_<Test>_<yyyyMMdd-HHmmss>, with:
      <prefix>-Steps.csv         every check: stage, status, message, details, duration
      <prefix>-Mailboxes.csv     where each mailbox is: tenant, sign-in, Exchange Online or on-premises, path
      <prefix>-FreeBusy.csv      one line per target: status, error and its cause, detail, working hours
      <prefix>-Events.csv        the calendar items returned, one line per item
      <prefix>-Provisioning.csv  user, licenses, mailbox in Exchange Online or not
      <prefix>-Dns.csv           DNS records found and expected
      <prefix>-Actions.csv       what the run changed (automatic replies scheduled and restored)
      <prefix>-Route.csv         message headers: one line per step of the chronology (sender, hops, mailbox) and its findings
      <prefix>-Trace.csv         every HTTP request sent and the response received, and the DNS queries (secrets masked)
      <prefix>-Summary.json      the whole result, for scripts
      <prefix>.html              self-contained report (templates\Report.template.html)
    CSV files: UTF-8 with BOM, configurable delimiter, cells starting with = + - @ prefixed with an
    apostrophe (no formula injection in Excel). Tokens are never part of the result.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>

$script:ReportColumns = [ordered]@{
    Steps        = @('Stage', 'Name', 'Status', 'Message', 'Details', 'DurationMs', 'TimestampUtc', 'Trace')
    Mailboxes    = @('Address', 'Role', 'Domain', 'TenantId', 'TenantRegion', 'Realm', 'Location', 'EwsUrl', 'Relation', 'Autodiscover')
    FreeBusy     = @('Mailbox', 'Status', 'ResponseCode', 'Message', 'Hint', 'Protocol', 'Path', 'DetailLevel', 'WorkingHours', 'Items', 'IntervalMinutes', 'StartUtc', 'View')
    Events       = @('Mailbox', 'Start', 'End', 'BusyType', 'Subject', 'Location', 'IsPrivate', 'StartUtc', 'EndUtc')
    Provisioning = @('Mail', 'DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled', 'SyncedFromAD', 'MailboxInExchangeOnline', 'MailboxType', 'Licenses', 'ExchangePlans', 'ExchangePlanEnabled', 'Autodiscover', 'MailboxProbe', 'UsageLocation', 'Created', 'ProxyAddresses')
    Dns          = @('Record', 'Name', 'Type', 'Found', 'Expected', 'Status', 'Source')
    Actions      = @('TimeUtc', 'Action', 'Target', 'Result', 'Detail')
    Route        = @('Step', 'Kind', 'TimeUtc', 'Wait', 'SegmentName', 'Server', 'From', 'Ip', 'Tls', 'LongestWait', 'EntersMicrosoft365', 'Status', 'FindingsText')
    Trace        = @('Sequence', 'TimestampUtc', 'Stage', 'Step', 'StepName', 'Method', 'Url', 'Operation', 'Label', 'StatusCode', 'Reason', 'Server', 'DurationMs', 'Repeated', 'Note', 'Request', 'Response')
}

function Format-GctDetails {
    <# Details of a step on one line: Key=Value; Key=Value (empty values left out). #>
    param([AllowNull()][System.Collections.IDictionary]$Details)

    if ($null -eq $Details) { return '' }
    $parts = foreach ($key in $Details.Keys) {
        $value = $Details[$key]
        if ($null -eq $value -or ([string]$value) -eq '') { continue }
        if ($value -is [array]) { $value = $value -join ', ' }
        '{0}={1}' -f $key, $value
    }
    return (@($parts) -join '; ')
}

function Format-GctCsvCell {
    param([AllowNull()][object]$Value, [Parameter(Mandatory = $true)][string]$Delimiter)

    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Collections.IDictionary]) { $text = Format-GctDetails $Value }
    elseif ($Value -is [array]) { $text = $Value -join ', ' }
    elseif ($Value -is [bool]) { $text = if ($Value) { 'True' } else { 'False' } }
    elseif ($Value -is [string]) {
        $text = $Value
        if ($text -match '^[=+\-@\t\r]') { $text = "'" + $text }
    }
    else { $text = [string]$Value }
    if ($text.Contains($Delimiter) -or $text.Contains('"') -or $text -match '[\r\n]') { $text = '"' + $text.Replace('"', '""') + '"' }
    return $text
}

function Write-GctCsv {
    param([AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory = $true)][string[]]$Columns, [Parameter(Mandatory = $true)][string]$Path, [string]$Delimiter = ';')

    $builder = [Text.StringBuilder]::new()
    [void]$builder.AppendLine((@($Columns | ForEach-Object { Format-GctCsvCell $_ $Delimiter }) -join $Delimiter))
    foreach ($row in @($Rows)) {
        [void]$builder.AppendLine((@($Columns | ForEach-Object { Format-GctCsvCell (Get-GctField $row $_) $Delimiter }) -join $Delimiter))
    }
    [IO.File]::WriteAllText($Path, $builder.ToString(), [Text.UTF8Encoding]::new($true))
}

function ConvertTo-GctEmbeddedJson {
    <# JSON safe inside a <script type="application/json"> block. #>
    param([AllowNull()][object]$Value)
    $json = ConvertTo-Json -InputObject $Value -Depth 8 -Compress
    if ([string]::IsNullOrEmpty($json)) { $json = 'null' }
    return $json.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
}

function New-GctRunFolder {
    param([Parameter(Mandatory = $true)][string]$OutputPath, [Parameter(Mandatory = $true)][string]$Prefix, [string]$TestType = 'Run')
    $base = Join-Path $OutputPath ('{0}_{1}_{2}' -f $Prefix, $TestType, (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $path = $base
    $n = 2
    while (Test-Path -LiteralPath $path) { $path = "$base-$n"; $n++ }
    [void][IO.Directory]::CreateDirectory($path)
    return $path
}

function Export-GctReport {
    <#
    .SYNOPSIS
        Writes the CSV, JSON and HTML files of one result in a new folder under OutputPath.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$Prefix = 'GraphConnectivityTests',
        [ValidateSet('Csv', 'Html')][string[]]$Formats = @('Csv', 'Html'),
        [ValidateSet(';', ',', "`t")][string]$Delimiter = ';'
    )

    $testType = [string](Get-GctField $Result 'TestType')
    $runPath = New-GctRunFolder -OutputPath $OutputPath -Prefix $Prefix -TestType $(if ($testType) { $testType } else { 'Run' })
    $data = [ordered]@{
        Steps        = @(Get-GctField $Result 'Steps')
        Mailboxes    = @(Get-GctField $Result 'Mailboxes')
        FreeBusy     = @(Get-GctField $Result 'FreeBusy')
        Events       = @(Get-GctField $Result 'FreeBusyEvents')
        Provisioning = @(Get-GctField $Result 'Provisioning')
        Dns          = @(Get-GctField $Result 'DnsRecords')
        Actions      = @(Get-GctField $Result 'Actions')
        Route        = @(Get-GctField $Result 'Route')
        Trace        = @(Get-GctField $Result 'Trace')
    }
    $files = [ordered]@{}
    if ($Formats -contains 'Csv') {
        foreach ($name in $data.Keys) {
            $files[$name] = Join-Path $runPath "$Prefix-$name.csv"
            Write-GctCsv -Rows $data[$name] -Columns $script:ReportColumns[$name] -Path $files[$name] -Delimiter $Delimiter
        }
    }
    $files.Summary = Join-Path $runPath "$Prefix-Summary.json"
    [IO.File]::WriteAllText($files.Summary, (ConvertTo-Json -InputObject $Result -Depth 8), [Text.UTF8Encoding]::new($false))

    if ($Formats -contains 'Html') {
        $summary = [ordered]@{}
        foreach ($key in 'Tool', 'Version', 'Status', 'TestType', 'Scenario', 'StartedUtc', 'CompletedUtc', 'DurationSeconds', 'Cloud', 'SourceMailbox', 'TargetMailboxes', 'Context', 'ClientId', 'ClientName',
            'SignIn', 'SignedInUser', 'TenantId', 'Token', 'SourceProtocol', 'SourceEwsUrl', 'OnPremisesHost', 'Domain', 'AllowWrite', 'AdminChecks', 'StageTitles', 'FreeBusyWindow', 'Hybrid', 'HeaderSummary', 'Throttled', 'Counts', 'Error') {
            $summary[$key] = Get-GctField $Result $key
        }
        $template = [IO.File]::ReadAllText((Join-Path $script:ToolRoot 'templates\Report.template.html'))
        $status = if ($summary.Status) { $summary.Status } else { 'Unknown' }
        $summary.TargetMailboxes = @(Get-GctField $Result 'TargetMailboxes')
        # Every marker filled in one pass over the template: the data inserted (which can contain {{...}}) is never read again.
        $values = @{ TITLE = [Net.WebUtility]::HtmlEncode("Graph Connectivity Tests | $status | $testType"); SUMMARY_JSON = ConvertTo-GctEmbeddedJson $summary }
        foreach ($name in $data.Keys) { $values["$($name.ToUpperInvariant())_JSON"] = ConvertTo-GctEmbeddedJson @($data[$name]) }
        $markers = @([regex]::Matches($template, '\{\{([A-Z_]+)\}\}') | ForEach-Object { $_.Groups[1].Value })
        $unknown = @($markers | Where-Object { -not $values.ContainsKey($_) })
        if ($unknown.Count) { throw "Report template marker not replaced: {{$($unknown[0])}}" }
        $missing = @($values.Keys | Where-Object { $_ -notin $markers })
        if ($missing.Count) { throw "Report template marker missing: {{$($missing[0])}}" }
        $html = [regex]::Replace($template, '\{\{([A-Z_]+)\}\}', [Text.RegularExpressions.MatchEvaluator]{ param($m) $values[$m.Groups[1].Value] })
        $files.Html = Join-Path $runPath "$Prefix.html"
        [IO.File]::WriteAllText($files.Html, $html, [Text.UTF8Encoding]::new($true))
    }
    [pscustomobject]@{ Directory = $runPath; Files = $files }
}
