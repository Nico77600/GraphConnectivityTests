#Requires -Version 7.4

<#
.SYNOPSIS
    Renders the screenshots of the guides (package\docs\images\gct-*.png) from the real tool, against the
    simulated Entra ID, Microsoft Graph and Exchange.

.DESCRIPTION
    No tenant is needed and the images always match the current code: the console, the window and the report
    are produced by the module itself, with tests\GraphConnectivityTests.Simulator.ps1 standing in for the
    network (Install-SimNetwork replaces only the network calls, the browser and the waits).

    The run: a FreeBusy test with -AdminChecks in the hybrid organisation contoso.com of the simulator.
    alice@contoso.com (Exchange Online) asks the free/busy of bob@contoso.com, the room room1@contoso.com,
    carol@contoso.com (on-premises, behind mail.contoso.com) and dave@fabrikam.com (another organisation).
    Exchange on-premises refuses the token of Exchange Online (HTTP 401): carol fails with its cause; the
    others answer. Only the fictitious names of the simulator appear (contoso.com, fabrikam.com, their
    simulated tenant IDs); the paths shown are C:\Tools\GraphConnectivityTests\...

        gct-report-overview.png  header, tiles and overview of the HTML report
        gct-report-freebusy.png  the free/busy view (one row per target, the error and its cause)
        gct-report-trace.png     the HTTP trace, with the getSchedule exchange open
        gct-console.png          the console of the same test (ANSI colours and emoji, as in Windows Terminal)
        gct-gui.png              the window after the test, light theme
        gct-gui-dark.png         the window after the test, dark theme
        gct-report-route.png     the Route panel of a MessageHeaders report (sample tests\samples\headers\gateway-in-front.txt)
        gct-gui-headers.png      the window after a MessageHeaders test of the same sample, light theme
        gct-gui-headers-dark.png the same, dark theme

    The HTML pages are captured with Microsoft Edge (headless). The window is opened in a separate STA
    PowerShell (New-GctTestForm -Theme Light|Dark, the hook of Show-GctTestGui for the documentation), on
    screen and topmost, and captured with CopyFromScreen: the session must be unlocked during that part.
    The reports and the work pages are kept in artifacts\doc-images (ignored by Git).
    Run tools\Build-Documentation.ps1 afterwards: the guides embed the images.

.PARAMETER OutputFolder
    Default: package\docs\images next to the tools folder.

.PARAMETER SkipWindow
    Does not render the two window images (no desktop session, or a locked screen).

.EXAMPLE
    .\tools\New-DocumentationImages.ps1; .\tools\Build-Documentation.ps1

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0  (from Web Services Client for Exchange 1.0.0)
    Part of : Graph Connectivity Tests (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$SkipWindow
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'package\docs\images' }
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path $root 'artifacts\doc-images'
if (Test-Path $work) { Remove-Item $work -Recurse -Force }
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$shown = 'C:\Tools\GraphConnectivityTests'
$stamp = Get-Date -Format 'yyyyMMdd'

#region Edge helpers (same as Web Services Client for Exchange) ---------------------------------------
function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height, [string]$Fragment) {
    $url = 'file:///' + ($Html -replace '\\', '/') + $(if ($Fragment) { "#$Fragment" })
    if (Test-Path $Png) { Remove-Item $Png -Force }
    # Start-Process, not &: an Edge helper process can keep the output pipe open after the capture.
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,$Height", '--force-device-scale-factor=1', '--virtual-time-budget=3000', "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
}

function Get-PageHeight([string]$Html, [int]$Width, [string]$Fragment) {
    # The page writes its height in body[data-h]; Edge returns the DOM with --dump-dom.
    $url = 'file:///' + ($Html -replace '\\', '/') + $(if ($Fragment) { "#$Fragment" })
    $dom = Join-Path $work ('dom-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html')
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,2000", '--virtual-time-budget=3000', '--dump-dom', "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $dom
    if (-not $proc.WaitForExit(45000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    # Edge helper processes inherit the output handle: read in shared mode, retry until written.
    $m = $null
    for ($i = 0; $i -lt 20 -and -not ($m -and $m.Success); $i++) {
        $stream = [IO.File]::Open($dom, 'Open', 'Read', 'ReadWrite')
        try { $text = [IO.StreamReader]::new($stream).ReadToEnd() } finally { $stream.Dispose() }
        $m = [regex]::Match($text, 'data-h="(\d+)"')
        if (-not $m.Success) { Start-Sleep -Milliseconds 250 }
    }
    if (-not $m.Success) { throw "Height not measured: $Html" }
    return [int]$m.Groups[1].Value
}

function Save-Page([string]$Html, [string]$Name, [int]$Width, [int]$MaxHeight = 0, [string]$Fragment) {
    $height = Get-PageHeight $Html $Width $Fragment
    if ($MaxHeight -gt 0 -and $height -gt $MaxHeight) { $height = $MaxHeight }
    Save-Screenshot $Html (Join-Path $OutputFolder "$Name.png") $Width $height $Fragment
    Write-Host ("  {0,-26} {1} x {2}" -f "$Name.png", $Width, $height)
}

function Assert-Anonymous([string]$Text, [string]$What) {
    # No local path, user name or computer name of this computer in an image.
    foreach ($leak in @($root, $work, $env:USERPROFILE, $env:USERNAME, $env:COMPUTERNAME, [IO.Path]::GetTempPath().TrimEnd('\')) | Where-Object { $_ -and $_.Length -ge 3 }) {
        if ($Text.IndexOf($leak, [StringComparison]::OrdinalIgnoreCase) -ge 0) { throw "$What contains '$leak': not anonymous." }
    }
}
#endregion

#region Simulated organisation --------------------------------------------------------------------
# The console theme is chosen when the module loads: colours and emoji, as in Windows Terminal.
$env:GCT_FORCE_COLOR = '1'
$env:GCT_ICONS = 'Emoji'
$culture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
Remove-Module GraphConnectivityTests -ErrorAction SilentlyContinue
Import-Module (Join-Path $root 'package\GraphConnectivityTests.psd1') -Force
$module = Get-Module GraphConnectivityTests
. (Join-Path $root 'tests\GraphConnectivityTests.Simulator.ps1')

# Bob has a meeting whose subject is a spreadsheet formula (the CSV injection test of the simulator): an
# ordinary subject in the images. The simulator itself is not changed.
${function:Get-SimScheduleOriginal} = ${function:Get-SimSchedule}
function Get-SimSchedule {
    param([string]$Mailbox, [datetime]$StartUtc, [int]$Days = 7, [int]$IntervalMinutes = 30)
    $s = Get-SimScheduleOriginal -Mailbox $Mailbox -StartUtc $StartUtc -Days $Days -IntervalMinutes $IntervalMinutes
    foreach ($e in $s.Events) { if ($e.Subject -like '=*') { $e.Subject = 'Agenda review'; $e.Location = 'Room 2' } }
    return $s
}

function New-DocState {
    # The hybrid organisation of the simulator; Exchange on-premises refuses the token of Exchange Online
    # (carol fails, with its cause); a desktop session, so the sign-in window is used (simulated).
    $s = New-SimState
    $s.CloudToOnPrem = 'Unauthorized'
    $s.DesktopSession = $true
    return $s
}

$settings = Import-GctConfiguration
$settings.TestType = 'FreeBusy'
$settings.SourceMailbox = 'alice@contoso.com'
$settings.TargetMailboxes = @('bob@contoso.com', 'room1@contoso.com', 'carol@contoso.com', 'dave@fabrikam.com')
$settings.AdminChecks = $true
$settings.SignIn = 'Auto'
$settings.OutputPath = Join-Path $work 'reports'
$settings.LogPath = Join-Path $work 'logs'
#endregion

try {
    #region Report images ---------------------------------------------------------------------------
    Write-Host 'Report (simulated run)...'
    Install-SimNetwork -Module $module -State (New-DocState)
    $result = Invoke-GctConnectivityTest -Configuration $settings -TestType FreeBusy -Quiet
    $report = Export-GctReport -Result $result -OutputPath $settings.OutputPath
    Write-Host ("  {0}, {1} checks, free/busy {2} of {3}" -f $result.Status, @($result.Steps).Count, @($result.FreeBusy | Where-Object Status -ne 'Failed').Count, @($result.FreeBusy).Count)
    $reportHtml = [IO.File]::ReadAllText($report.Files.Html).Replace($settings.OutputPath, "$shown\reports").Replace($settings.OutputPath.Replace('\', '\\'), "$shown\reports".Replace('\', '\\'))
    Assert-Anonymous $reportHtml 'The HTML report'

    function Save-ReportView([string]$Name, [string]$Css, [string]$Script, [string]$Tab, [int]$MaxHeight = 0) {
        # A copy of the report with a style and a script; the height is the bottom of the page content.
        $measure = "document.body.setAttribute('data-h', Math.ceil(document.querySelector('.wrap').getBoundingClientRect().bottom + 16));"
        $inject = "<style>$Css</style><script>window.addEventListener('load', () => { $Script; setTimeout(() => { $measure }, 150); });</script></body>"
        $page = Join-Path $work "$Name.html"
        [IO.File]::WriteAllText($page, $reportHtml.Replace('</body>', $inject), [Text.UTF8Encoding]::new($false))
        Save-Page $page $Name 1400 $MaxHeight "tab=$Tab&theme=light"
    }
    Save-ReportView 'gct-report-overview' 'footer { display:none !important; }' '' 'overview' 1700
    Save-ReportView 'gct-report-freebusy' 'header.hero, footer { display:none !important; } body { padding-top:8px; }' '' 'freebusy' 1000
    $openTrace = "const x = Array.from(document.querySelectorAll('#panel-trace details.exchange')).find(d => /getSchedule/.test(d.querySelector('summary').textContent) && !/source/.test(d.querySelector('summary').textContent)); if (x) { x.open = true; const all = Array.from(document.querySelectorAll('#panel-trace details.exchange')); const i = all.indexOf(x); all.slice(0, Math.max(0, i - 2)).forEach(d => d.hidden = true); }"
    Save-ReportView 'gct-report-trace' 'header.hero, footer { display:none !important; } body { padding-top:8px; }' $openTrace 'trace' 1150

    # MessageHeaders: a sample header, the DNS of today from the simulated DNS.
    Write-Host 'Report of a message header analysis (sample, simulated DNS)...'
    Install-SimNetwork -Module $module -State (New-DocState)
    $headerSettings = $settings.Clone()
    $headerSettings.TestType = 'MessageHeaders'
    $headerSettings.HeaderText = [IO.File]::ReadAllText((Join-Path $root 'tests\samples\headers\gateway-in-front.txt'))
    $headerResult = Invoke-GctConnectivityTest -Configuration $headerSettings -TestType MessageHeaders -Quiet
    $headerReport = Export-GctReport -Result $headerResult -OutputPath $settings.OutputPath
    $reportHtml = [IO.File]::ReadAllText($headerReport.Files.Html).Replace($settings.OutputPath, "$shown\reports").Replace($settings.OutputPath.Replace('\', '\\'), "$shown\reports".Replace('\', '\\'))
    Assert-Anonymous $reportHtml 'The HTML report of the header analysis'
    Save-ReportView 'gct-report-route' 'footer { display:none !important; }' '' 'route' 2200
    #endregion

    #region Console image ---------------------------------------------------------------------------
    Write-Host 'Console (simulated run)...'
    Install-SimNetwork -Module $module -State (New-DocState)
    $consoleSettings = $settings.Clone()
    $consoleSettings.AdminChecks = $false
    $consoleSettings.FreeBusyDays = 5
    $consoleSettings.TargetMailboxes = @('bob@contoso.com', 'carol@contoso.com', 'dave@fabrikam.com')
    $displayed = $consoleSettings.Clone(); $displayed.OutputPath = "$shown\reports"
    $log = "$shown\logs\GraphConnectivityTests_$stamp.log"
    $records = & {
        Write-GctRunBanner -Settings $displayed -LogPath $log
        $script:DocResult = Invoke-GctConnectivityTest -Configuration $consoleSettings -TestType FreeBusy
        $script:DocReport = Export-GctReport -Result $script:DocResult -OutputPath (Join-Path $work 'console-reports')
        Write-GctRunSummary -Result $script:DocResult -ReportText "$shown\reports\$(Split-Path $script:DocReport.Directory -Leaf)\GraphConnectivityTests.html" -LogPath $log
    } 6>&1 | Where-Object { $_ -is [Management.Automation.InformationRecord] }

    function ConvertFrom-Ansi([string]$Line) {
        # SGR codes of the console theme: 0 reset, 1 bold, 90 dim, 97 white, 38;2;r;g;b and 48;2;r;g;b.
        $out = [Text.StringBuilder]::new()
        $fg = $null; $bg = $null; $bold = $false
        foreach ($part in [regex]::Split($Line, '(\x1b\[[0-9;]*m)')) {
            if ($part -match '^\x1b\[([0-9;]*)m$') {
                $codes = @($Matches[1].Split(';') | ForEach-Object { if ($_ -eq '') { 0 } else { [int]$_ } })
                for ($i = 0; $i -lt $codes.Count; $i++) {
                    switch ($codes[$i]) {
                        0 { $fg = $null; $bg = $null; $bold = $false }
                        1 { $bold = $true }
                        22 { $bold = $false }
                        39 { $fg = $null }
                        49 { $bg = $null }
                        90 { $fg = '#8a8a8a' }
                        97 { $fg = '#ffffff' }
                        38 { $fg = 'rgb({0},{1},{2})' -f $codes[$i + 2], $codes[$i + 3], $codes[$i + 4]; $i += 4 }
                        48 { $bg = 'rgb({0},{1},{2})' -f $codes[$i + 2], $codes[$i + 3], $codes[$i + 4]; $i += 4 }
                    }
                }
                continue
            }
            if ($part -eq '') { continue }
            $style = @()
            if ($fg) { $style += "color:$fg" }
            if ($bg) { $style += "background:$bg" }
            if ($bold) { $style += 'font-weight:700' }
            $text = [Net.WebUtility]::HtmlEncode($part)
            [void]$out.Append($(if ($style) { "<span style=""$($style -join ';')"">$text</span>" } else { $text }))
        }
        return $out.ToString()
    }

    $lines = foreach ($r in $records) {
        $data = $r.MessageData
        $text = if ($data -is [Management.Automation.HostInformationMessage]) { [string]$data.Message } else { [string]$data }
        $text.Replace($settings.OutputPath, "$shown\reports").Replace($settings.LogPath, "$shown\logs").Replace($work, $shown)
    }
    Assert-Anonymous ($lines -join "`n") 'The console output'
    $command = ".\Invoke-GraphConnectivityTests.ps1 -SourceMailbox alice@contoso.com -TargetMailbox $($consoleSettings.TargetMailboxes -join ',') -FreeBusyDays 5"
    $body = ($lines | ForEach-Object { ConvertFrom-Ansi $_ }) -join "`n"
    $console = @"
<!doctype html><html><head><meta charset="utf-8"><style>
body { margin:0; background:#ffffff; font-family:"Segoe UI", sans-serif; }
.win { width:1498px; margin:0; border-radius:10px; overflow:hidden; background:#0c0c0c; border:1px solid #2b2b2b; }
.bar { display:flex; align-items:center; gap:10px; height:38px; padding:0 14px; background:#202020; color:#d0d0d0; font-size:12.5px; }
.tab { padding:6px 14px; background:#0c0c0c; border-radius:8px 8px 0 0; margin-top:8px; }
pre { margin:0; padding:14px 18px 18px; color:#cccccc; font:13.5px/1.42 "Cascadia Mono", Consolas, "Segoe UI Emoji", monospace; white-space:pre-wrap; word-break:break-word; }
.prompt { color:#cccccc; }
</style></head><body><div class="win"><div class="bar"><span class="tab">PowerShell 7.5</span></div>
<pre><span class="prompt">PS $shown&gt; $([Net.WebUtility]::HtmlEncode($command))</span>
$body</pre></div>
<script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.win').getBoundingClientRect().height) + 2);</script></body></html>
"@
    $page = Join-Path $work 'gct-console.html'
    [IO.File]::WriteAllText($page, $console, [Text.UTF8Encoding]::new($false))
    Save-Page $page 'gct-console' 1500
    #endregion
}
finally {
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Remove-Module GraphConnectivityTests -Force -ErrorAction SilentlyContinue
    Remove-Item Env:\GCT_FORCE_COLOR, Env:\GCT_ICONS -ErrorAction SilentlyContinue
    [Threading.Thread]::CurrentThread.CurrentCulture = $culture
}

#region Window images -----------------------------------------------------------------------------
# A separate STA PowerShell: WPF needs it, and the window must not inherit the console settings above.
$windowScript = @'
param([string]$Root, [string]$Theme, [string]$Png, [string]$Work, [string]$Shown, [string]$Focus = 'Free/busy of carol@*', [string]$Test = 'FreeBusy', [string]$Sample)
$ErrorActionPreference = 'Stop'
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
Import-Module (Join-Path $Root 'package\GraphConnectivityTests.psd1') -Force
$module = Get-Module GraphConnectivityTests
. (Join-Path $Root 'tests\GraphConnectivityTests.Simulator.ps1')
$state = New-SimState
$state.CloudToOnPrem = 'Unauthorized'
$state.DesktopSession = $true
Install-SimNetwork -Module $module -State $state
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class GctDocWin { [StructLayout(LayoutKind.Sequential)] public struct R { public int L, T, Ri, B; }
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr h, int a, ref int v, int s);
 [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
 [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h, int a, out R r, int s); }
"@
& $module {
    param($Theme, $Png, $Work, $Shown, $Focus, $Test, $Sample)
    $reports = Join-Path $Work "gui-$Theme\reports"
    $cfg = Import-GctConfiguration
    $cfg.TestType = 'FreeBusy'
    $cfg.SourceMailbox = 'alice@contoso.com'
    $cfg.TargetMailboxes = @('bob@contoso.com', 'room1@contoso.com', 'carol@contoso.com', 'dave@fabrikam.com')
    $cfg.AdminChecks = $true
    $cfg.SignIn = 'Auto'
    $cfg.OutputPath = $reports
    $cfg.LogPath = Join-Path $Work "gui-$Theme\logs"
    if ($Test -eq 'MessageHeaders') { $cfg.TestType = 'MessageHeaders'; $cfg.HeaderText = [IO.File]::ReadAllText($Sample); $cfg.HeaderFile = '' }
    $f = New-GctTestForm -Configuration $cfg -Theme $Theme
    $form = $f.Form
    $form.WindowStartupLocation = 'Manual'
    $form.Left = 24; $form.Top = 16
    $form.Topmost = $true
    # Dark theme: the title bar dark too, as Windows draws it when the applications are dark.
    $form.Show()
    $hwnd = [Windows.Interop.WindowInteropHelper]::new($form).Handle
    if ($Theme -eq 'Dark') {
        # DWMWA_USE_IMMERSIVE_DARK_MODE, then SWP_FRAMECHANGED | NOMOVE | NOSIZE | NOZORDER to redraw the frame.
        $on = 1; [void][GctDocWin]::DwmSetWindowAttribute($hwnd, 20, [ref]$on, 4)
        [void][GctDocWin]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, 0, 0, 0, 0x27)
    }
    Invoke-GctGuiRun
    # Anonymised paths in the progress and the footer.
    $walk = {
        param($node)
        if ($node -is [Windows.Controls.TextBlock] -and $node.Text -and $node.Text.Contains($reports)) { $node.Text = $node.Text.Replace($reports, "$Shown\reports") }
        foreach ($child in [Windows.LogicalTreeHelper]::GetChildren($node)) { if ($child -is [Windows.DependencyObject]) { & $walk $child } }
    }
    & $walk $f.Controls.Timeline
    $f.Controls.Footer.Text = $f.Controls.Footer.Text.Replace($reports, "$Shown\reports")
    # The timeline from the failed mailbox, its diagnostics open (details and HTTP exchanges).
    for ($k = 0; $k -lt 10; $k++) { Invoke-GctGuiPump }
    $failed = $f.Gui.Checks | Where-Object { $_.Step.Name -like $Focus } | Select-Object -First 1
    if ($failed) {
        Set-GctGuiCheckOpen -Check $failed -Open $true
        for ($k = 0; $k -lt 10; $k++) { Invoke-GctGuiPump }
        $y = $failed.Root.TranslatePoint([Windows.Point]::new(0, 0), $f.Controls.Timeline).Y
        $f.Controls.LogScroll.ScrollToVerticalOffset([Math]::Max(0, $y - 40))
    }
    for ($k = 0; $k -lt 20; $k++) { Invoke-GctGuiPump; Start-Sleep -Milliseconds 40 }
    # Windows can minimise a window that opened in the background (the frame is then 219 x 39): restore it first.
    $form.WindowState = 'Normal'
    [void]$form.Activate()
    for ($k = 0; $k -lt 10; $k++) { Invoke-GctGuiPump; Start-Sleep -Milliseconds 30 }
    [void][GctDocWin]::SetForegroundWindow($hwnd)
    Start-Sleep -Milliseconds 600
    for ($k = 0; $k -lt 5; $k++) { Invoke-GctGuiPump }
    $r = New-Object GctDocWin+R
    [void][GctDocWin]::DwmGetWindowAttribute($hwnd, 9, [ref]$r, 16)   # DWMWA_EXTENDED_FRAME_BOUNDS: no shadow
    $bmp = [Drawing.Bitmap]::new($r.Ri - $r.L, $r.B - $r.T)
    $gr = [Drawing.Graphics]::FromImage($bmp)
    $gr.CopyFromScreen($r.L, $r.T, 0, 0, $bmp.Size)
    # 8 pixels off each edge: the rounded corners and the frame would show the desktop behind the window.
    $inner = $bmp.Clone([Drawing.Rectangle]::new(8, 8, $bmp.Width - 16, $bmp.Height - 16), $bmp.PixelFormat)
    $inner.Save($Png, [Drawing.Imaging.ImageFormat]::Png); $inner.Dispose(); $gr.Dispose(); $bmp.Dispose()
    $form.Topmost = $false
    $form.Close()
    "$($r.Ri - $r.L) x $($r.B - $r.T)"
} $Theme $Png $Work $Shown $Focus $Test $Sample
'@

if (-not $SkipWindow) {
    Write-Host 'Window (simulated run, on screen)...'
    $scriptPath = Join-Path $work 'window.ps1'
    [IO.File]::WriteAllText($scriptPath, $windowScript, [Text.UTF8Encoding]::new($true))
    $pwsh = (Get-Process -Id $PID).Path
    $sample = Join-Path $root 'tests\samples\headers\gateway-in-front.txt'
    foreach ($shot in @(@{ Theme = 'Light'; Name = 'gct-gui' }, @{ Theme = 'Dark'; Name = 'gct-gui-dark' },
            @{ Theme = 'Light'; Name = 'gct-gui-headers'; Test = 'MessageHeaders'; Focus = '4. * Microsoft 365 entry*' }, @{ Theme = 'Dark'; Name = 'gct-gui-headers-dark'; Test = 'MessageHeaders'; Focus = '4. * Microsoft 365 entry*' })) {
        $extra = if ($shot['Test']) { @('-Test', $shot.Test, '-Focus', "`"$($shot.Focus)`"", '-Sample', "`"$sample`"") } else { @() }
        $png = Join-Path $OutputFolder "$($shot.Name).png"
        if (Test-Path $png) { Remove-Item $png -Force }
        $out = Join-Path $work "window-$($shot.Theme).txt"
        # A window minimised by Windows gives a tiny image (219 x 39): up to three attempts.
        for ($attempt = 1; $attempt -le 3; $attempt++) {
        $proc = Start-Process -FilePath $pwsh -ArgumentList (@('-NoProfile', '-STA', '-File', "`"$scriptPath`"", '-Root', "`"$root`"", '-Theme', $shot.Theme, '-Png', "`"$png`"", '-Work', "`"$work`"", '-Shown', "`"$shown`"") + $extra) -PassThru -WindowStyle Hidden -RedirectStandardOutput $out -RedirectStandardError "$out.err"
        if (-not $proc.WaitForExit(180000)) { Stop-Process -Id $proc.Id -Force; throw "The window ($($shot.Theme)) did not close in 3 minutes." }
        if (-not (Test-Path $png)) { throw "Window image not written ($($shot.Theme)): $(Get-Content "$out.err" -Raw)" }
        if ((Get-Content $out -Raw).Trim() -match '^(\d+) x' -and [int]$Matches[1] -ge 1000) { break }
        Remove-Item $png -Force
        if ($attempt -eq 3) { throw "Window image too small ($($shot.Theme)): the window was minimised; run again without using the computer meanwhile." }
        }
        Write-Host ("  {0,-26} {1} ({2})" -f "$($shot.Name).png", (Get-Content $out -Raw).Trim(), $shot.Theme)
    }
}
#endregion

Write-Host "Reports and work pages: $work"
