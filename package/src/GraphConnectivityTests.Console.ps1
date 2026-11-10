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

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCByTqiMuHsRlw8Y
# q3xDm0vmPzQbPecuwIwZV3ybY/yrq6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCABW2lh/+9Z
# I+HzPlEVAudA2/aOBsmjuoncF7nVZNRL6zANBgkqhkiG9w0BAQEFAASCAgCw3wD5
# ZyLHh9aW5NaGXYBObliXnuZPxFW0FnUv4BhUVKt3fjgDU96HnGSbaz1gjT5YGnUe
# zz+m6hOYOLWR4pWsQ67/Zy3koz3QSgPTsHstUPIZnL0RLL7ngJkpEXxtNcw5aCd8
# 3jD+SlpMC9YjRoWBrlgvkx15ewnVDlPd2zbXRJgLiqOrfFe+FbC6h9N94BiTe6c6
# qWHjEN1cElb52bHt+j3TiX1f0RA621ED7QrrpAezE4rU1cCZQ0LN7hQG45pxaepd
# nBl0BUToKsRb0O1W3QOBr8G2wxnOQo7GMnxpmboShgOTrf0VpMl87/oxticAupY7
# N2qnIt98vlflzX/3L2FXBH2Mn8bsrhOaykeThDeFkyv7zNWbD6bdmrw3DZHU55Sw
# IroHKyKzi3yQKMlYiU3Qa80LinP4arIKeOlavpHxWzH67In2ua0A95opbP0d/2Zb
# KhkVyGMbKm6xWqxjJQod5aZYJLfbOnzZ0mysw9c5U4vXu/QIZWKmC+WPPbhdNR+v
# frvKUCC2QLOQFh0ZHiYmypEPWPQKxyyC2PeHs8LD+kCOjg4HGFz+SM3ePvuo2BcI
# nspy7CPMCsgw9gtGtxK94E4XdJ+rJmxT42FI2XLjgPfV59dJ8PmLLLX6dXqx89Ia
# lrq1kXi/KLXceFK/ATMfHzlF5HCqmQXgLCHvu6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIJNNvTFbvxw26fsn7mFkEr+/8yvyOjUqlMg3LWecHFW2Aghp11RxGCKI
# NxgPMjAyNjExMTAxODI1NTJaMAMCAQECEG+TVGgn9a+n7WjH/z3w5WigghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDE4MjU1MlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIPqS
# eTGNf6lZjYstYBEz5n2f+8LEIUUoaCSrwIh3LFA7MIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEAslLIiPWwuJicXxFRCZ28VWlfT0dNpih8bQB6TEF2cDACIAM0
# xUCGQbd8bF9Qs2CrjOJmMbfq0DProbCxGT8FIsag
# SIG # End signature block
