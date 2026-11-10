<#
.SYNOPSIS
    Graph Connectivity Tests - console output and log file (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    Same rules as the other tools of the author (Web Services Client for Exchange, Exchange Log Report):
      - ANSI colours are disabled when the output is redirected or NO_COLOR is set;
        GCT_FORCE_COLOR=1 forces them.
      - Icons: emoji in Windows Terminal / VS Code, symbols of the classic console fonts elsewhere.
        GCT_ICONS = Emoji | Symbols | Ascii forces a style.
      - Every line shown is also written to the daily log file, without colours or icons.
      - During a GUI run, the same lines are sent to the progress box of the window.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>

$script:C = @{ Reset = ''; Bold = ''; Dim = ''; Accent = ''; AccentBg = ''; Green = ''; Yellow = ''; Red = ''; White = '' }
if ($env:GCT_FORCE_COLOR -eq '1' -or (-not [Console]::IsOutputRedirected -and -not $env:NO_COLOR)) {
    $e = [char]27
    $script:C = @{
        Reset = "$e[0m"; Bold = "$e[1m"; Dim = "$e[90m"; White = "$e[97m"
        Accent = "$e[38;2;214;62;115m"; AccentBg = "$e[48;2;177;31;75m$e[97m"
        Green = "$e[38;2;80;200;120m"; Yellow = "$e[38;2;240;200;90m"; Red = "$e[38;2;240;90;90m"
    }
}
$script:IconStyle = if ($env:GCT_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:GCT_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }

function Get-GctIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' {
            return @{
                Logo = & $u 0x1F4C5; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F; Fail = & $u 0x274C; Info = & $u 0x1F539
                Skip = & $u 0x23E9; Block = & $u 0x26D4; Key = & $u 0x1F511; Server = & $u 0x1F5A5; Shield = & $u 0x1F512
                Folder = & $u 0x1F4C1; Mail = & $u 0x1F4E8; People = & $u 0x1F465; File = & $u 0x1F4C4; Log = & $u 0x1F4DD
                Report = & $u 0x1F4CA; Done = & $u 0x1F389; Target = & $u 0x1F3AF; Search = & $u 0x1F50E; Clock = & $u 0x23F3
                Settings = & $u 0x1F527
            }
        }
        'Symbols' {
            return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022
                Skip = & $u 0x00BB; Block = & $u 0x25A0; Key = & $u 0x00A7; Server = & $u 0x2261; Shield = & $u 0x25CA
                Folder = & $u 0x2302; Mail = '@'; People = & $u 0x2192; File = & $u 0x25AC; Log = & $u 0x00B6
                Report = & $u 0x2261; Done = & $u 0x221A; Target = & $u 0x25D9; Search = & $u 0x25BA; Clock = & $u 0x25CB
                Settings = & $u 0x263C
            }
        }
        default {
            return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Block = '#'; Key = 'k'; Server = '='
                Shield = 'o'; Folder = '>'; Mail = '@'; People = '&'; File = '-'; Log = '='; Report = '='; Done = '*'
                Target = 'o'; Search = '?'; Clock = '~'; Settings = '%'
            }
        }
    }
}

function Get-GctFrameSet {
    <# Rounded corners in modern terminals (emoji style), square corners elsewhere (present in every console font). #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    if ($Style -eq 'Ascii') {
        return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' }
    }
    if ($Style -eq 'Symbols') {
        return @{ TopLeft = [char]0x250C; TopRight = [char]0x2510; BottomLeft = [char]0x2514; BottomRight = [char]0x2518; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
    }
    return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
}

$script:Icons = Get-GctIconSet $script:IconStyle
$script:Frame = Get-GctFrameSet $script:IconStyle
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }

function Get-GctIcon { param([Parameter(Mandatory = $true)][string]$Name) return $script:Icons[$Name] + $script:IconPad }

function Format-GctDuration {
    param([Parameter(Mandatory = $true)][double]$Seconds)

    $inv = [Globalization.CultureInfo]::InvariantCulture
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    if ($t.TotalHours -ge 1) { return [string]::Format($inv, '{0} h {1:00} min', [int][Math]::Floor($t.TotalHours), $t.Minutes) }
    if ($t.TotalMinutes -ge 1) { return [string]::Format($inv, '{0} min {1:00} s', $t.Minutes, $t.Seconds) }
    return [string]::Format($inv, '{0:0.0} s', $t.TotalSeconds)
}

function Send-GctUi {
    <# Forwards a console line to the GUI progress box while a GUI run is in progress. #>
    param([string]$Status, [string]$Text)
    if ($script:Ui -and $script:Ui.Sink) { & $script:Ui.Sink $Status $Text }
}

function Start-GctLog {
    <# Opens (or continues) today's log file and deletes the log files older than the retention. #>
    param([Parameter(Mandatory = $true)][string]$Directory, [int]$RetentionDays = 14)

    Stop-GctLog
    [void][IO.Directory]::CreateDirectory($Directory)
    $script:LogPath = Join-Path $Directory ('GraphConnectivityTests_{0:yyyyMMdd}.log' -f (Get-Date))
    $stream = [IO.FileStream]::new($script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $script:LogWriter = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    $script:LogWriter.AutoFlush = $true
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Directory -Filter 'GraphConnectivityTests_*.log' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt $limit | Remove-Item -Force -ErrorAction SilentlyContinue
    return $script:LogPath
}

function Stop-GctLog {
    if ($script:LogWriter) { $script:LogWriter.Dispose(); $script:LogWriter = $null }
}

function Write-GctLog {
    <# One line in the log file only. The log never contains colours, icons or tokens. #>
    param(
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP')][string]$Level = 'INFO',
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Message
    )
    if ($script:LogWriter) { $script:LogWriter.WriteLine(('{0:yyyy-MM-ddTHH:mm:ss.fffzzz} [{1,-5}] {2}' -f (Get-Date), $Level, $Message)) }
}

function Write-GctBanner {
    <# Title card at the start of an execution, followed by the context rows (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Subtitle,
        [System.Collections.Specialized.OrderedDictionary]$Details
    )

    Write-GctLog 'STEP' "=== $Title v$($script:ToolVersion) ==="
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $v = $Details[$key]
            Write-GctLog 'INFO' ('{0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
        }
    }
    if ($script:Quiet) { return }
    $C = $script:C; $F = $script:Frame; $width = 74
    $right = "v$($script:ToolVersion) $([char]0x00B7) Nicolas Fabert"
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $left = "  $($script:Icons.Logo)  $Title"
    $gap = [Math]::Max(1, $width - ($left.Length - $script:Icons.Logo.Length + $iconWidth) - $right.Length - 2)
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.TopLeft, [string]::new($F.Horizontal, $width), $F.TopRight, $C.Reset)
    Write-Host ('  {0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Bold, $left, $C.Reset, [string]::new(' ', $gap), $C.Dim, $right, '  ', ($C.Accent + $F.Vertical), $C.Reset)
    if ($Subtitle) {
        $sub = "     $Subtitle"
        if ($sub.Length -gt $width - 2) { $sub = $sub.Substring(0, $width - 5) + '...' }
        Write-Host ('  {0}{1}{2}{3}{4}{5}{0}{6}{2}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Dim, $sub.PadRight($width), $C.Reset, $F.Vertical)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            $icon, $text = if ($value -is [array]) { (Get-GctIcon $value[0]), $value[1] } else { '   ', $value }
            Write-Host ('     {0}{1}{2,-11}{3} {4}' -f $icon, $C.Dim, $key, $C.Reset, $text)
        }
    }
}

function Write-GctStep {
    <# Step header with a coloured number pill and an icon:  ─ 3/5 ─ 🔑  Sign-in #>
    param(
        [Parameter(Mandatory = $true)][int]$Number,
        [Parameter(Mandatory = $true)][int]$Total,
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Icon = 'Info'
    )

    Write-GctLog 'STEP' "[$Number/$Total] $Title"
    Send-GctUi 'Step' "[$Number/$Total] $Title"
    if ($script:Quiet) { return }
    $C = $script:C
    Write-Host ''
    Write-Host ('  {0} {1}/{2} {3} {4}{5}{6}{3}' -f $C.AccentBg, $Number, $Total, $C.Reset, (Get-GctIcon $Icon), $C.Bold, $Title)
}

function Write-GctItem {
    <# One indented result line with a status icon, also written to the log and to the GUI. #>
    param(
        [ValidateSet('Ok', 'Warn', 'Fail', 'Info', 'Skip', 'Block')][string]$Status = 'Info',
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [string]$Icon,
        [switch]$NoUi
    )

    $level = @{ Ok = 'OK'; Warn = 'WARN'; Fail = 'ERROR'; Info = 'INFO'; Skip = 'INFO'; Block = 'WARN' }[$Status]
    Write-GctLog $level $Text
    if (-not $NoUi) { Send-GctUi $Status $Text }
    if ($script:Quiet) { return }
    $color = @{ Ok = $script:C.Green; Warn = $script:C.Yellow; Fail = $script:C.Red; Info = ''; Skip = $script:C.Dim; Block = $script:C.Yellow }[$Status]
    $symbol = Get-GctIcon $(if ($Icon) { $Icon } else { $Status })
    $textColor = if ($Status -in 'Warn', 'Fail', 'Skip', 'Block') { $color } else { '' }
    Write-Host ('      {0}{1}{2}{3}{4}{2}' -f $color, $symbol, $script:C.Reset, $textColor, $Text)
}

function Write-GctSummary {
    <# Final summary card (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][System.Collections.Specialized.OrderedDictionary]$Values,
        [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok'
    )

    foreach ($key in $Values.Keys) {
        $v = $Values[$key]
        Write-GctLog 'INFO' ('Summary - {0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
    }
    if ($script:Quiet) { return }
    $C = $script:C; $F = $script:Frame; $width = 74
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $icon = $script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $iconWidth))
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}{0}{5}{6}{7}' -f $color, $F.TopLeft, $F.Horizontal, $C.Bold, $head, ($C.Reset + $color), ([string]::new($F.Horizontal, $rest) + $F.TopRight), $C.Reset)
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $rowIcon, $text = if ($value -is [array]) { (Get-GctIcon $value[0]), $value[1] } else { '   ', $value }
        Write-Host ('    {0}{1}{2,-10}{3} {4}' -f $rowIcon, $C.Dim, $key, $C.Reset, $text)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $color, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    Write-Host ''
}

function Get-GctSignInText {
    <# One line that says who signs in and how (banner, window footer). #>
    param([Parameter(Mandatory = $true)][hashtable]$Configuration)

    $dot = [char]0x00B7
    if ([string]$Configuration.Context -eq 'Application') { return "Application $($Configuration.ClientId) $dot certificate $($Configuration.CertificateThumbprint) $dot tenant $($Configuration.TenantId)" }
    $client = Resolve-GctClient -ClientId ([string]$Configuration.ClientId) -RedirectUri ([string]$Configuration.RedirectUri)
    $what = if ($client.Key -eq 'MicrosoftOffice') { 'no application to register' } elseif ($client.Explicit) { 'consent of the permissions' } else { 'your application' }
    return "User $($Configuration.SourceMailbox) $dot $($client.Name) ($what)"
}

function Write-GctRunBanner {
    <# Title card of a command-line run: test, source, targets, sign-in, writes, report and log. #>
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [string]$LogPath, [switch]$NoReport)

    $scenario = Get-GctTestCatalog | Where-Object Name -eq $Settings.TestType
    $dot = [char]0x00B7
    $banner = [ordered]@{}
    $banner['Test'] = @('Target', "$($scenario.Name) $dot $($scenario.DisplayName)")
    if ($Settings.TestType -eq 'MessageHeaders') {
        $banner['Headers'] = @('Mail', $(if (([string]$Settings.HeaderText).Trim()) { "text pasted ($(([string]$Settings.HeaderText).Length) characters)" } else { [string]$Settings.HeaderFile }))
        $banner['DNS'] = @('Search', $(if ($Settings.SkipDnsCrossCheck) { 'not compared (-SkipDnsCrossCheck): no DNS query' } else { $(if (@($Settings.DnsServer | Where-Object { $_ }).Count) { "compared with the DNS of today (DNS server $(@($Settings.DnsServer) -join ', ')); offline when no DNS answers" } else { 'compared with the DNS of today (resolver of Windows); offline when no DNS answers' }) }))
        $banner['Sign-in'] = @('Key', 'none: the header is read on this computer')
        $banner['Report'] = @('Report', $(if ($NoReport) { 'none (-NoReport)' } else { $Settings.OutputPath }))
        if ($LogPath) { $banner['Log'] = @('Log', $LogPath) }
        Write-GctBanner -Title 'Graph Connectivity Tests' -Subtitle "Message headers $dot route $dot authentication $dot anti-spam" -Details $banner
        return
    }
    $banner['Source'] = @('People', [string]$Settings.SourceMailbox)
    if (@($Settings.TargetMailboxes).Count) { $banner['Targets'] = @('People', (@($Settings.TargetMailboxes) -join ', ')) }
    if ($Settings.TestType -in 'DomainDns', 'All') { $banner['Domain'] = @('Search', $(if ($Settings.Domain) { [string]$Settings.Domain } else { "$(Get-GctAddressDomain ([string]$Settings.SourceMailbox)) (domain of the source)" })) }
    if ([string]$Settings.OnPremisesHost) { $banner['On-premises'] = @('Server', [string]$Settings.OnPremisesHost) }
    $signIn = Get-GctSignInText -Configuration $Settings
    if ([string]$Settings.Context -ne 'Application') {
        $how = try { if ((Get-GctSignInMode -Configuration $Settings).Mode -eq 'Window') { 'sign-in window' } else { 'device code' } } catch { 'sign-in window not available here' }
        $signIn = "$signIn $dot $how"
    }
    $banner['Sign-in'] = @('Key', $signIn)
    $banner['Cloud'] = @('Server', $script:Clouds[[string]$Settings.Cloud].Name)
    if ($Settings.AdminChecks) { $banner['Admin'] = @('Settings', "hybrid configuration read by an administrator ($([string]$Settings.AdminClientId))") }
    if ($scenario.Writes) { $banner['Writes'] = @('Shield', $(if ($Settings.AllowWrite) { 'AUTHORISED: automatic replies scheduled, read back and restored (only when disabled)' } else { 'none (-AllowWrite not given): read-only' })) }
    $banner['Report'] = @('Report', $(if ($NoReport) { 'none (-NoReport)' } else { $Settings.OutputPath }))
    if ($LogPath) { $banner['Log'] = @('Log', $LogPath) }
    Write-GctBanner -Title 'Graph Connectivity Tests' -Subtitle "Microsoft Graph $dot user context $dot free/busy, mailbox, provisioning, DNS" -Details $banner
}

function Write-GctRunSummary {
    <# Final card of a command-line run: status, counts, first issue, report, log and what to do next. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result, [string]$ReportText = 'none (-NoReport)', [string]$LogPath)

    $dot = [char]0x00B7
    $n = $Result.Counts
    $values = [ordered]@{}
    $values['Status'] = @($(switch ($Result.Status) { 'Passed' { 'Ok' } 'Failed' { 'Fail' } 'Blocked' { 'Block' } default { 'Warn' } }), "$($Result.Status) $dot $($Result.TestType)")
    $values['Checks'] = @('Target', ('{0} passed {5} {1} warning(s) {5} {2} blocked {5} {3} failed {5} {4} skipped' -f $n.Passed, $n.Warning, $n.Blocked, $n.Failed, $n.Skipped, $dot))
    $fb = @($Result.FreeBusy)
    if ($fb.Count) { $values['Free/busy'] = @('Clock', ('{0} of {1} mailbox(es) answered' -f @($fb | Where-Object Status -ne 'Failed').Count, $fb.Count)) }
    $hs = Get-GctField $Result 'HeaderSummary'
    if ($hs) {
        $values['Route'] = @('Clock', "$($hs.Hops) hop(s) $dot delivered in $($hs.Transit)$(if ($hs.LongestWait) { " $dot longest wait $($hs.LongestWait)" })")
        $values['Auth'] = @('Shield', "SPF $(if ($hs.Spf) { $hs.Spf } else { '-' }) $dot DKIM $(if ($hs.Dkim) { $hs.Dkim } else { '-' }) $dot DMARC $(if ($hs.Dmarc) { $hs.Dmarc } else { '-' })$(if ($hs.Compauth) { " $dot compauth $($hs.Compauth)" })")
        if ($hs.DnsCrossCheck) { $values['DNS today'] = @('Search', "$($hs.Fixed) finding(s) fixed since the message was received") }
    }
    if ($Result.Error) { $values['First issue'] = @('Fail', $Result.Error) }
    $values['Duration'] = @('Clock', (Format-GctDuration $Result.DurationSeconds))
    $values['Report'] = @('Report', $ReportText)
    if ($LogPath) { $values['Log'] = @('Log', $LogPath) }
    $values['Next'] = @('Info', $(switch ($Result.Status) {
                'Passed' { 'Nothing to do. The report shows every request and response.' }
                'Blocked' { 'A writing step needs -AllowWrite (test mailbox only).' }
                'Warning' { 'Read the warnings in the report: each one gives the cause and what to check.' }
                default { 'Open the report: each failed step gives the error of Exchange, its cause and what to check.' }
            }))
    $title = switch ($Result.Status) { 'Passed' { 'Test passed' } 'Warning' { 'Test finished with warnings' } 'Blocked' { 'Test blocked' } default { 'Test failed' } }
    $card = switch ($Result.Status) { 'Passed' { 'Ok' } 'Failed' { 'Fail' } default { 'Warn' } }
    Write-GctSummary -Title $title -Values $values -Status $card
}

function Write-GctFreeBusyGrid {
    <#
        Free/busy as a small map in the console: one line per day and mailbox, one cell per hour from 07:00
        to 20:00 (local time of this computer), the strongest status of the hour. Also in the log, as text.
        Free .  Tentative ~  Busy #  Away O  Working elsewhere w  No data ?
    #>
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Results, [Parameter(Mandatory = $true)][datetime]$StartUtc, [int]$IntervalMinutes = 30, [int]$Days = 7, [int]$FirstHour = 7, [int]$LastHour = 20)

    $C = $script:C
    $ascii = $script:IconStyle -eq 'Ascii'
    $glyph = if ($ascii) { @{ 0 = ' . '; 1 = ' ~ '; 2 = '###'; 3 = 'OOO'; 4 = ' w '; 5 = ' ? ' } }
             else { @{ 0 = [string]::new([char]0x00B7, 1).PadLeft(2).PadRight(3); 1 = [string]::new([char]0x2592, 3); 2 = [string]::new([char]0x2588, 3); 3 = [string]::new([char]0x2593, 3); 4 = [string]::new([char]0x2591, 3); 5 = ' ? ' } }
    $color = @{ 0 = $C.Dim; 1 = $C.Yellow; 2 = "$([char]27)[38;2;71;158;245m"; 3 = "$([char]27)[38;2;180;140;230m"; 4 = "$([char]27)[38;2;90;200;200m"; 5 = $C.Dim }
    if (-not $C.Reset) { $color = @{ 0 = ''; 1 = ''; 2 = ''; 3 = ''; 4 = ''; 5 = '' } }
    $rank = @{ 3 = 6; 2 = 5; 1 = 4; 4 = 3; 0 = 2; 5 = 1 }
    $codes = @{ free = 0; tentative = 1; busy = 2; oof = 3; workingelsewhere = 4 }
    $startLocal = $StartUtc.ToLocalTime()
    $ruler = (($FirstHour..($LastHour - 1)) | ForEach-Object { '{0:00} ' -f $_ }) -join ''
    $shown = $false
    foreach ($r in $Results) {
        $view = [string]$r.View
        if (-not $view) { continue }
        $items = @($r.Events | Where-Object { $_ -and $_.StartUtc } | ForEach-Object { [pscustomobject]@{ S = ([datetime]$_.StartUtc).ToUniversalTime(); E = ([datetime]$_.EndUtc).ToUniversalTime(); C = $codes[([string]$_.BusyType).ToLowerInvariant()] } } | Where-Object { $null -ne $_.C })
        if (-not $shown -and -not $script:Quiet) { Write-Host ('      {0}{1,-12}{2}{3}' -f $C.Dim, 'Free/busy', $ruler, $C.Reset) }
        $shown = $true
        if (-not $script:Quiet) { Write-Host ('      {0}{1}{2}' -f $C.Bold, $r.Mailbox, $C.Reset) }
        Write-GctLog 'INFO' "Free/busy map of $($r.Mailbox) (hours $FirstHour-$LastHour, . free ~ tentative # busy O away w elsewhere)"
        for ($d = 0; $d -lt $Days; $d++) {
            $day = $startLocal.AddDays($d)
            $cells = [Text.StringBuilder]::new(); $plain = [Text.StringBuilder]::new()
            for ($h = $FirstHour; $h -lt $LastHour; $h++) {
                $from = $day.Date.AddHours($h).ToUniversalTime(); $to = $from.AddHours(1)
                $best = 5
                for ($t = $from; $t -lt $to; $t = $t.AddMinutes($IntervalMinutes)) {
                    $i = [int][Math]::Floor(($t - $StartUtc).TotalMinutes / $IntervalMinutes)
                    if ($i -ge 0 -and $i -lt $view.Length) { $v = [int][string]$view[$i]; if ($v -le 4 -and ($rank[$v] -gt $rank[$best] -or $best -eq 5)) { $best = $v } }
                }
                foreach ($e in $items) { if ($e.S -lt $to -and $e.E -gt $from -and ($rank[$e.C] -gt $rank[$best] -or $best -eq 5)) { $best = $e.C } }
                [void]$cells.Append($color[$best] + $glyph[$best] + $C.Reset)
                [void]$plain.Append(@{ 0 = '.'; 1 = '~'; 2 = '#'; 3 = 'O'; 4 = 'w'; 5 = '?' }[$best])
            }
            $label = $day.ToString('ddd dd MMM', [Globalization.CultureInfo]::GetCultureInfo('en-GB'))
            if (-not $script:Quiet) { Write-Host ('      {0}{1,-12}{2}{3}' -f $C.Dim, $label, $C.Reset, $cells.ToString()) }
            Write-GctLog 'INFO' ('  {0}  {1}' -f $label, $plain.ToString())
        }
    }
    if ($shown -and -not $script:Quiet) {
        $legend = if ($ascii) { '. free  ~ tentative  # busy  O away  w elsewhere' } else { "$($glyph[0].Trim()) free  $([char]0x2592) tentative  $([char]0x2588) busy  $([char]0x2593) away  $([char]0x2591) elsewhere" }
        Write-Host ('      {0}{1,-12}{2}{3}' -f $C.Dim, '', $legend, $C.Reset)
    }
}
