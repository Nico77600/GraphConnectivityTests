<#
.SYNOPSIS
    Graph Connectivity Tests - message header analysis (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    The MessageHeaders test reads the header block of a message (an .eml or .txt file, or text pasted) and tells
    how the message travelled and what Microsoft 365 decided about it. Nothing signs in and no Microsoft Graph
    permission is used: the header is read on this computer.

        HeaderRead   the fields of the header (folded lines joined, encoded words decoded), the message, and the
                     commands to follow it in Exchange Online and on-premises
        HeaderDns    optional (Test.SkipDnsCrossCheck, -SkipDnsCrossCheck): the DNS of today for what the header
                     shows - SPF of the MAIL FROM domain evaluated for the original sending IP (RFC 7208), DMARC of
                     the From domain, the DKIM keys of the signatures, the MX of the recipient domain - so that a
                     finding already handled since is marked "fixed since"
        HeaderRoute  the chronology, one check per step: the sender itself (Date field, client, first host), each
                     Received hop (time, wait, who held the message), the mailbox (folder, spam verdict). Each
                     finding sits on the step it is about: authentication and hybrid attribution where the
                     message enters Microsoft 365, the Enhanced Filtering tip on a gateway, the spam verdict and
                     the folder on the mailbox.

    The parser and the findings are the same as the ones of the web page of Graph Connectivity Tests
    (wwwroot/js/headers.js): the code tables are one shared file (data\header-codes.json) and the tests of
    both use the same sample headers.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>

$script:HeaderCodes = $null
$script:HeaderMaxBytes = 2MB
try { [Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance) } catch { Write-Verbose 'Code pages already available.' }

function Get-GctHeaderCodes {
    <# The code tables (data\header-codes.json), read once. #>
    if (-not $script:HeaderCodes) {
        $path = Join-Path $script:ToolRoot 'data\header-codes.json'
        if (-not (Test-Path -LiteralPath $path)) { throw "Code tables of the message header analysis not found: $path" }
        $script:HeaderCodes = [IO.File]::ReadAllText($path) | ConvertFrom-Json -AsHashtable
    }
    return $script:HeaderCodes
}

#region Reading the header block -------------------------------------------------------------------

function ConvertFrom-GctHeaderBlock {
    <# The fields of a header (RFC 5322): folded lines joined, in their order (newest Received first). The body of an .eml is not read. #>
    param([AllowEmptyString()][AllowNull()][string]$Text)

    $fields = [Collections.Generic.List[object]]::new()
    $t = ([string]$Text) -replace "`r`n?", "`n"
    if ($t.Length -and $t[0] -eq [char]0xFEFF) { $t = $t.Substring(1) }
    $started = $false
    foreach ($line in $t.Split("`n")) {
        $blank = -not $line.Trim()
        if (-not $started -and $blank) { continue }
        if ($started -and $blank) { break }
        $m = [regex]::Match($line, '^([!-9;-~]+):[ \t]?(.*)$')
        if ($m.Success -and $line -notmatch '^[ \t]') {
            $fields.Add([ordered]@{ Name = $m.Groups[1].Value; Value = $m.Groups[2].Value })
            $started = $true
        }
        elseif ($fields.Count) {
            # A folded line, or a line whose indentation was lost in a copy: it continues the previous field.
            $last = $fields[$fields.Count - 1]
            $last.Value = $last.Value + ' ' + ($line -replace '^[ \t]+', '')
        }
    }
    foreach ($f in $fields) { $f.Value = $f.Value -replace '[ \t]+$', '' }
    return , @($fields)
}

function Get-GctTextFromBytes {
    param([byte[]]$Bytes, [string]$Charset)
    try { return [Text.Encoding]::GetEncoding($Charset).GetString($Bytes) } catch { return [Text.Encoding]::UTF8.GetString($Bytes) }
}

function ConvertFrom-GctEncodedWord {
    <# Encoded words of RFC 2047 (=?charset?B?...?= and =?charset?Q?...?=) decoded; the rest left as it is. #>
    param([AllowEmptyString()][AllowNull()][string]$Value)

    $text = [string]$Value
    if (-not $text.Contains('=?')) { return $text }
    # White space between two encoded words is not displayed (RFC 2047, 6.2).
    $joined = [regex]::Replace($text, '(=\?[^?\s]+\?[BbQq]\?[^?\s]*\?=)\s+(?==\?[^?\s]+\?[BbQq]\?)', '$1')
    return [regex]::Replace($joined, '=\?([^?\s]+?)(?:\*[^?]*)?\?([BbQq])\?([^?\s]*)\?=', [Text.RegularExpressions.MatchEvaluator] {
            param($m)
            try {
                $data = $m.Groups[3].Value
                if ($m.Groups[2].Value -ieq 'B') {
                    $b64 = $data -replace '[^A-Za-z0-9+/=]', ''
                    $b64 = $b64.TrimEnd('=')
                    while ($b64.Length % 4) { $b64 += '=' }
                    $bytes = [Convert]::FromBase64String($b64)
                }
                else {
                    $list = [Collections.Generic.List[byte]]::new()
                    for ($i = 0; $i -lt $data.Length; $i++) {
                        $ch = $data[$i]
                        if ($ch -eq '_') { $list.Add(32) }
                        elseif ($ch -eq '=' -and $i + 2 -lt $data.Length -and $data.Substring($i + 1, 2) -match '^[0-9A-Fa-f]{2}$') { $list.Add([Convert]::ToByte($data.Substring($i + 1, 2), 16)); $i += 2 }
                        else { $list.Add([byte]([int][char]$ch -band 0xFF)) }
                    }
                    $bytes = $list.ToArray()
                }
                return Get-GctTextFromBytes -Bytes $bytes -Charset $m.Groups[1].Value
            }
            catch { return $m.Value }
        })
}

$script:MailMonths = @{ jan = 1; feb = 2; mar = 3; apr = 4; may = 5; jun = 6; jul = 7; aug = 8; sep = 9; oct = 10; nov = 11; dec = 12 }
$script:MailZones = @{ UT = 0; UTC = 0; GMT = 0; Z = 0; EST = -300; EDT = -240; CST = -360; CDT = -300; MST = -420; MDT = -360; PST = -480; PDT = -420 }

function ConvertFrom-GctMailDate {
    <# A date of RFC 5322 (and the usual variants) as milliseconds since 1970 (UTC), or $null. #>
    param([AllowEmptyString()][AllowNull()][string]$Text)

    $s = [regex]::Replace([string]$Text, '\([^)]*\)', ' ')
    $m = [regex]::Match($s, '(\d{1,2})[\s-]+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s-]+(\d{2,4})\s+(\d{1,2}):(\d{2})(?::(\d{2})(?:[.,](\d{1,7}))?)?\s*([+-]\d{4}|[A-Za-z]{1,3}\b)?', 'IgnoreCase')
    if ($m.Success) {
        $year = [int]$m.Groups[3].Value
        if ($year -lt 100) { $year += $(if ($year -lt 50) { 2000 } else { 1900 }) }
        $ms = if ($m.Groups[7].Success) { [Math]::Floor([double]('0.' + $m.Groups[7].Value) * 1000 + 0.5) } else { 0 }
        $offset = 0
        $zone = $m.Groups[8].Value
        if ($zone -match '^[+-]\d{4}$') {
            $sign = if ($zone[0] -eq '-') { -1 } else { 1 }
            $offset = $sign * ([int]$zone.Substring(1, 2) * 60 + [int]$zone.Substring(3, 2))
        }
        elseif ($zone -and $script:MailZones.ContainsKey($zone.ToUpperInvariant())) { $offset = $script:MailZones[$zone.ToUpperInvariant()] }
        try {
            $seconds = if ($m.Groups[6].Success) { [int]$m.Groups[6].Value } else { 0 }
            $t = [DateTimeOffset]::new($year, $script:MailMonths[$m.Groups[2].Value.ToLowerInvariant()], [int]$m.Groups[1].Value, 0, 0, 0, [TimeSpan]::Zero)
            $t = $t.AddHours([int]$m.Groups[4].Value).AddMinutes([int]$m.Groups[5].Value - $offset).AddSeconds($seconds).AddMilliseconds($ms)
            return $t.ToUnixTimeMilliseconds()
        }
        catch { return $null }
    }
    $dto = [DateTimeOffset]::MinValue
    if ($s.Trim() -and [DateTimeOffset]::TryParse($s.Trim(), [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)) { return $dto.ToUnixTimeMilliseconds() }
    return $null
}

#endregion

#region Received ------------------------------------------------------------------------------------

function Test-GctIPv4 { param([AllowEmptyString()][AllowNull()][string]$Text) return [string]$Text -match '^(?:25[0-5]|2[0-4]\d|1?\d?\d)(?:\.(?:25[0-5]|2[0-4]\d|1?\d?\d)){3}$' }
function Test-GctIPv6 { param([AllowEmptyString()][AllowNull()][string]$Text) $t = [string]$Text; return ($t -match '^[0-9a-f:]+(?:%\w+)?$' -and $t.Contains(':') -and $t.Split(':').Count -ge 3) }
function Test-GctIpText { param([AllowEmptyString()][AllowNull()][string]$Text) return ((Test-GctIPv4 $Text) -or (Test-GctIPv6 $Text)) }

function Split-GctHeaderComment {
    <# Comments (parentheses, nested) taken out of a field: the text with a mark for each, and the comments. #>
    param([AllowEmptyString()][string]$Value)
    $out = [Text.StringBuilder]::new()
    $comments = [Collections.Generic.List[string]]::new()
    $current = [Text.StringBuilder]::new()
    $depth = 0
    foreach ($ch in $Value.ToCharArray()) {
        if ($ch -eq [char]'(') {
            if ($depth -eq 0) { [void]$current.Clear() } else { [void]$current.Append($ch) }
            $depth++
        }
        elseif ($ch -eq [char]')' -and $depth -gt 0) {
            $depth--
            if ($depth -eq 0) { [void]$out.Append(" $([char]1)$($comments.Count)$([char]1) "); $comments.Add($current.ToString().Trim()) }
            else { [void]$current.Append($ch) }
        }
        elseif ($depth -gt 0) { [void]$current.Append($ch) }
        else { [void]$out.Append($ch) }
    }
    if ($depth -gt 0) { $comments.Add($current.ToString().Trim()) }
    return @{ Text = $out.ToString(); Comments = @($comments) }
}

function Restore-GctHeaderComment {
    param([AllowEmptyString()][string]$Text, [string[]]$Comments)
    $r = [regex]::Replace($Text, "$([char]1)(\d+)$([char]1)", [Text.RegularExpressions.MatchEvaluator] { param($m) '(' + $Comments[[int]$m.Groups[1].Value] + ')' })
    return ($r -replace '\s+', ' ').Trim()
}
function Get-GctCommentsIn {
    param([AllowEmptyString()][string]$Text, [string[]]$Comments)
    return @([regex]::Matches($Text, "$([char]1)(\d+)$([char]1)") | ForEach-Object { $Comments[[int]$_.Groups[1].Value] })
}
function Remove-GctCommentMark {
    param([AllowEmptyString()][AllowNull()][string]$Text)
    return (([string]$Text -replace "$([char]1)\d+$([char]1)", ' ') -replace '\s+', ' ').Trim()
}

function Get-GctIpInText {
    param([AllowEmptyString()][AllowNull()][string]$Text)
    if (-not $Text) { return $null }
    $b = [regex]::Match($Text, '\[(?:IPv6:)?([0-9A-Fa-f:.]+)(?:%\w+)?\]')
    if ($b.Success -and (Test-GctIpText $b.Groups[1].Value)) { return $b.Groups[1].Value }
    foreach ($token in [regex]::Split($Text, '[\s,;()\[\]]+')) {
        $t = ($token -replace '^IPv6:', '') -replace '%\w+$', ''
        if (Test-GctIpText $t) { return $t }
    }
    return $null
}

function ConvertFrom-GctReceived {
    <# One Received field: from, by, with, id, for, via, the date, the IP of the sending server, TLS. #>
    param([AllowEmptyString()][string]$Value)

    $split = Split-GctHeaderComment -Value $Value
    $text = $split.Text
    $comments = [string[]]@($split.Comments)
    $semi = $text.LastIndexOf(';')
    $main = if ($semi -ge 0) { $text.Substring(0, $semi) } else { $text }
    $dateText = if ($semi -ge 0) { $text.Substring($semi + 1).Trim() } else { '' }
    $clauses = @{}
    $marks = @([regex]::Matches($main, '(?:^|\s)(from|by|via|with|id|for)\s+', 'IgnoreCase'))
    for ($i = 0; $i -lt $marks.Count; $i++) {
        $end = if ($i + 1 -lt $marks.Count) { $marks[$i + 1].Index } else { $main.Length }
        $start = $marks[$i].Index + $marks[$i].Length
        $key = $marks[$i].Groups[1].Value.ToLowerInvariant()
        if (-not $clauses.ContainsKey($key)) { $clauses[$key] = $main.Substring($start, [Math]::Max(0, $end - $start)) }
    }
    $head = { param($s) if ($s) { (Remove-GctCommentMark $s).Split(' ')[0] } else { $null } }
    $fromRaw = [string]$clauses['from']
    $fromComments = @(Get-GctCommentsIn -Text $fromRaw -Comments $comments)
    $byRaw = [string]$clauses['by']
    $withRaw = [string]$clauses['with']
    $fromHost = & $head $fromRaw
    $byHost = & $head $byRaw
    # "from helo (rdns [ip])": the name the server gave (HELO) and, in the comment, its reverse DNS and IP.
    $rdns = $null
    foreach ($c in $fromComments) {
        $name = @($c -split '[\s\[\]]+' | Where-Object { $_ -match '^[a-z0-9-]+(\.[a-z0-9-]+)+$' -and -not (Test-GctIpText $_) }) | Select-Object -First 1
        if ($name) { $rdns = $name.TrimEnd('.'); break }
    }
    $bare = if ($fromHost) { $fromHost -replace '^\[|\]$', '' } else { $null }
    $ip = Get-GctIpInText ($fromComments -join ' ')
    if (-not $ip) { $ip = if (Test-GctIpText $bare) { $bare } else { Get-GctIpInText $fromRaw } }
    $allText = Restore-GctHeaderComment -Text $text -Comments $comments
    $withValue = if ((Remove-GctCommentMark $withRaw) -match '^Microsoft SMTP Server\b') { 'Microsoft SMTP Server' } else { & $head $withRaw }
    $id = & $head ([string]$clauses['id'])
    $tlsMatch = [regex]::Match(($comments -join ' '), 'version=(TLS[\w.]+)|(TLSv?1[._]?[0-3])|using (TLS\S*)', 'IgnoreCase')
    $tls = if ($tlsMatch.Success) {
        $v = @($tlsMatch.Groups[1].Value, $tlsMatch.Groups[2].Value, $tlsMatch.Groups[3].Value | Where-Object { $_ })[0]
        ([regex]'_').Replace($v, '.', 1)
    }
    elseif ([string]$withValue -match '^(ESMTPS|ESMTPSA|SMTPS|UTF8SMTPS|HTTPS|LMTPS)$') { 'yes' }
    else { $null }
    $for = $null
    if ($clauses.ContainsKey('for')) { $fm = [regex]::Match((Remove-GctCommentMark $clauses['for']), '<?([^<>\s;]+@[^<>\s;]+)>?'); if ($fm.Success) { $for = $fm.Groups[1].Value } }
    $exchangeId = if ([string]$id -match '^15\.\d+\.\d+\.\d+$') { $id } else { $mm = [regex]::Match($allText, 'mapi id (15\.\d+\.\d+\.\d+)', 'IgnoreCase'); if ($mm.Success) { $mm.Groups[1].Value } else { $null } }
    [ordered]@{
        Raw = $allText
        FromHost = if ($fromHost) { $fromHost.TrimEnd('.') } else { $null }
        Rdns = $rdns
        Ip = $ip
        ByHost = if ($byHost) { $byHost.TrimEnd('.') } else { $null }
        With = $withValue
        WithDetail = if ($withRaw) { Restore-GctHeaderComment -Text $withRaw -Comments $comments } else { $null }
        Id = $id
        Via = if ($clauses.ContainsKey('via')) { Remove-GctCommentMark $clauses['via'] } else { $null }
        For = $for
        DateText = $dateText
        Time = ConvertFrom-GctMailDate $dateText
        Tls = $tls
        ExchangeId = $exchangeId
    }
}

#endregion

#region Addresses, domains and structured fields --------------------------------------------------

function ConvertFrom-GctAddress {
    <# "Name <a@b>" or "a@b": the display name, the address and its domain. #>
    param([AllowEmptyString()][AllowNull()][string]$Value)
    $v = (ConvertFrom-GctEncodedWord $Value).Trim()
    if (-not $v) { return $null }
    $angle = [regex]::Match($v, '<([^<>]*)>\s*$')
    if (-not $angle.Success) { $angle = [regex]::Match($v, '<([^<>]*)>') }
    $address = if ($angle.Success) { $angle.Groups[1].Value } else { $am = [regex]::Match($v, '[^\s<>,;"]+@[^\s<>,;"]+'); if ($am.Success) { $am.Value } else { '' } }
    $address = $address.Trim()
    $name = if ($angle.Success) { $v.Substring(0, $angle.Index).Trim().Trim('"') } else { '' }
    if (-not $address -and -not $name) { return $null }
    $at = $address.LastIndexOf('@')
    [ordered]@{ Name = $(if ($name) { $name } else { $null }); Address = $(if ($address) { $address } else { $null }); Domain = $(if ($at -gt 0) { $address.Substring($at + 1).ToLowerInvariant().TrimEnd('.') } else { $null }); Text = $v }
}

function ConvertFrom-GctAddressList {
    param([AllowEmptyString()][AllowNull()][string]$Value)
    if (-not $Value) { return , @() }
    $parts = [Collections.Generic.List[string]]::new()
    $depth = 0; $quote = $false; $cur = [Text.StringBuilder]::new()
    foreach ($ch in (ConvertFrom-GctEncodedWord $Value).ToCharArray()) {
        if ($ch -eq [char]'"') { $quote = -not $quote }
        if (-not $quote -and $ch -eq [char]'<') { $depth++ }
        if (-not $quote -and $ch -eq [char]'>') { $depth-- }
        if (-not $quote -and $depth -eq 0 -and $ch -eq [char]',') { $parts.Add($cur.ToString()); [void]$cur.Clear() } else { [void]$cur.Append($ch) }
    }
    $parts.Add($cur.ToString())
    return , @($parts | ForEach-Object { ConvertFrom-GctAddress $_ } | Where-Object { $_ -and $_.Address })
}

function Get-GctDomainOf {
    param([AllowEmptyString()][AllowNull()][string]$Value)
    if (-not $Value) { return $null }
    $v = $Value.Trim() -replace '^<|>$', ''
    $at = $v.LastIndexOf('@')
    $d = $(if ($at -ge 0) { $v.Substring($at + 1) } else { $v }).ToLowerInvariant().TrimEnd('.')
    if ($d -cmatch '^[a-z0-9-]+(\.[a-z0-9-]+)+$') { return $d }
    return $null
}

function Get-GctOrgDomain {
    <# The organisational domain (RFC 7489): the registered domain, with a short list of multi-label suffixes. #>
    param([AllowEmptyString()][AllowNull()][string]$Domain, [hashtable]$Codes)
    if (-not $Domain) { return $null }
    $labels = @($Domain.ToLowerInvariant().Split('.') | Where-Object { $_ })
    if ($labels.Count -le 2) { return $labels -join '.' }
    $lastTwo = ($labels[-2..-1]) -join '.'
    $suffixes = @(if ($Codes) { $Codes['multiLabelSuffixes'] })
    if ($suffixes -contains $lastTwo) { return ($labels[-3..-1]) -join '.' }
    return $lastTwo
}

function ConvertFrom-GctColonList {
    <# "KEY:VALUE;KEY:VALUE" (X-Forefront-Antispam-Report, X-Microsoft-Antispam, Mailbox-Delivery), in order. #>
    param([AllowEmptyString()][AllowNull()][string]$Value)
    $out = [ordered]@{}
    foreach ($part in ([string]$Value).Split(';')) {
        $i = $part.IndexOf(':')
        if ($i -le 0) { continue }
        $k = $part.Substring(0, $i).Trim()
        if ($k -and -not $out.Contains($k)) { $out[$k] = $part.Substring($i + 1).Trim() }
    }
    return $out
}

function ConvertFrom-GctTagList {
    <# "tag=value; tag=value" (DKIM-Signature, ARC-Seal, DMARC): tags in lower case, spaces removed from the values. #>
    param([AllowEmptyString()][AllowNull()][string]$Value)
    $out = [ordered]@{}
    foreach ($part in ([string]$Value).Split(';')) {
        $i = $part.IndexOf('=')
        if ($i -le 0) { continue }
        $k = $part.Substring(0, $i).Trim().ToLowerInvariant()
        if ($k -and -not $out.Contains($k)) { $out[$k] = ($part.Substring($i + 1) -replace '\s+', '').Trim() }
    }
    return $out
}

function Split-GctOutside {
    param([AllowEmptyString()][string]$Value, [char]$Separator)
    $parts = [Collections.Generic.List[string]]::new()
    $depth = 0; $quote = $false; $cur = [Text.StringBuilder]::new()
    foreach ($ch in $Value.ToCharArray()) {
        if ($ch -eq [char]'"') { $quote = -not $quote }
        if (-not $quote -and $ch -eq [char]'(') { $depth++ }
        if (-not $quote -and $ch -eq [char]')' -and $depth -gt 0) { $depth-- }
        if (-not $quote -and $depth -eq 0 -and $ch -eq $Separator) { $parts.Add($cur.ToString()); [void]$cur.Clear() } else { [void]$cur.Append($ch) }
    }
    $parts.Add($cur.ToString())
    return , @($parts | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function ConvertFrom-GctAuthResults {
    <# Authentication-Results (RFC 8601): the server that wrote it and each method with its result and properties. #>
    param([AllowEmptyString()][AllowNull()][string]$Value)
    $parts = Split-GctOutside -Value ([string]$Value) -Separator ';'
    $authserv = $null
    $results = [Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $split = Split-GctHeaderComment -Value $parts[$i]
        $t = Remove-GctCommentMark $split.Text
        $m = [regex]::Match($t, '^([a-z][\w.-]*)\s*=\s*([\w.-]+)(.*)$', 'IgnoreCase')
        if (-not $m.Success) { if ($i -eq 0) { $authserv = @($t.Split(' '))[0]; if (-not $authserv) { $authserv = $null } }; continue }
        $props = [ordered]@{}
        foreach ($pm in [regex]::Matches($m.Groups[3].Value, '([a-z][\w.-]*)\s*=\s*("[^"]*"|[^\s;]+)', 'IgnoreCase')) { $props[$pm.Groups[1].Value.ToLowerInvariant()] = $pm.Groups[2].Value.Trim('"') }
        $results.Add([ordered]@{ Method = $m.Groups[1].Value.ToLowerInvariant(); Result = $m.Groups[2].Value.ToLowerInvariant(); Props = $props; Comment = $(if ($split.Comments.Count) { $split.Comments -join '; ' } else { $null }) })
    }
    return [ordered]@{ Authserv = $authserv; Results = @($results) }
}

#endregion

#region Servers, route and the whole analysis -----------------------------------------------------

$script:HeaderSegments = [ordered]@{
    origin         = 'Sender'
    sender         = 'Sender side'
    senderService  = 'Sending service'
    senderTenant   = 'Sender tenant (Microsoft 365)'
    gateway        = 'Third-party gateway'
    onPremises     = 'Exchange on-premises'
    entry          = 'Microsoft 365 entry (EOP)'
    exchangeOnline = 'Exchange Online'
    beyond         = 'After Microsoft 365'
    delivery       = 'Mailbox'
}

function Test-GctMicrosoftHost { param([AllowNull()][string]$HostName) return ([bool]$HostName -and $HostName -match '(\.prod\.outlook\.com|\.prod\.exchangelabs\.com|\.outlook\.office365\.com|\.office365\.com|\.protection\.outlook\.com|\.outlook\.com|\.hotmail\.com|\.mx\.microsoft)$') }
function Test-GctMicrosoftOutbound { param([AllowNull()][string]$HostName) return ([bool]$HostName -and $HostName -match '\.outbound\.protection\.outlook\.com$') }

function Find-GctHeaderProduct {
    param([object[]]$List, [AllowNull()][string]$HostName)
    if (-not $HostName) { return $null }
    foreach ($g in $List) { if ($g['hosts'] -and $HostName -match $g['hosts']) { return $g['name'] } }
    return $null
}

function Get-GctExchangeVersion {
    param([AllowNull()][string]$Id)
    if (-not $Id) { return $null }
    $n = $Id -replace '\.0+(\d)', '.$1'
    if ($n -match '^15\.0\.') { return "Exchange Server 2013 ($Id)" }
    if ($n -match '^15\.1\.') { return "Exchange Server 2016 ($Id)" }
    if ($n -match '^15\.2\.') { return "Exchange Server 2019 or Subscription Edition ($Id)" }
    return $null
}

function Format-GctHeaderDuration {
    <# 36 s, 2 min 16 s, 3 h 05 min, 2 days 4 h. #>
    param([AllowNull()][object]$Milliseconds)
    if ($null -eq $Milliseconds) { return '-' }
    $s = [long][Math]::Floor([double]$Milliseconds / 1000 + 0.5)
    if ($s -lt 60) { return "$s s" }
    if ($s -lt 3600) { return '{0} min {1:00} s' -f [Math]::Floor($s / 60), ($s % 60) }
    $h = [long][Math]::Floor($s / 3600)
    if ($h -lt 48) { return '{0} h {1:00} min' -f $h, [Math]::Floor(($s % 3600) / 60) }
    return '{0} days {1} h' -f [Math]::Floor($h / 24), ($h % 24)
}

function Format-GctHeaderTime {
    param([AllowNull()][object]$Milliseconds, [switch]$Full)
    if ($null -eq $Milliseconds) { return '-' }
    $t = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$Milliseconds)
    if ($Full) { return $t.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) + ' UTC' }
    return $t.ToString('HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
}

function Set-GctHopSegment {
    <# Who each server is: sender side, gateway, Exchange on-premises, Microsoft 365 entry, Exchange Online... Returns the entry index. #>
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Hops, [Parameter(Mandatory = $true)][hashtable]$Codes)
    $entry = -1
    for ($i = 0; $i -lt $Hops.Count; $i++) {
        $h = $Hops[$i]
        $fromOutside = -not (Test-GctMicrosoftHost $h.FromHost) -and -not (Test-GctMicrosoftHost $h.Rdns)
        $fromOtherTenant = (Test-GctMicrosoftOutbound $h.FromHost) -or (Test-GctMicrosoftOutbound $h.Rdns)
        if ((Test-GctMicrosoftHost $h.ByHost) -and ($fromOutside -or $fromOtherTenant)) { $entry = $i; break }
    }
    $lastMicrosoft = -1
    for ($i = 0; $i -lt $Hops.Count; $i++) { if (Test-GctMicrosoftHost $Hops[$i].ByHost) { $lastMicrosoft = $i } }
    for ($i = 0; $i -lt $Hops.Count; $i++) {
        $h = $Hops[$i]
        $gateway = Find-GctHeaderProduct -List $Codes['gateways'] -HostName $h.ByHost
        $service = Find-GctHeaderProduct -List $Codes['senders'] -HostName $h.ByHost
        $h.Product = $null
        $h.ExchangeVersion = if (-not (Test-GctMicrosoftHost $h.ByHost)) { Get-GctExchangeVersion $h.ExchangeId } else { $null }
        if (Test-GctMicrosoftHost $h.ByHost) {
            $h.Segment = if ($entry -ge 0 -and $i -lt $entry) { 'senderTenant' } elseif ($i -eq $entry) { 'entry' } else { 'exchangeOnline' }
        }
        elseif ($gateway) { $h.Segment = 'gateway'; $h.Product = $gateway }
        elseif ($h.ExchangeVersion) { $h.Segment = 'onPremises' }
        elseif ($service) { $h.Segment = 'senderService'; $h.Product = $service }
        elseif ($lastMicrosoft -ge 0 -and $i -gt $lastMicrosoft) { $h.Segment = 'beyond' }
        else { $h.Segment = 'sender' }
        $h.IsEntry = $i -eq $entry
    }
    return $entry
}

function Get-GctHeaderAnalysis {
    <#
    .SYNOPSIS
        The whole analysis of a header block: fields, summary, the sender, the route, the mailbox, authentication,
        anti-spam, Exchange, gateways, findings (each on its step), the DNS request for the cross-check and commands.
    #>
    param([AllowEmptyString()][AllowNull()][string]$Text, [hashtable]$Codes = (Get-GctHeaderCodes))

    $raw = ConvertFrom-GctHeaderBlock -Text $Text
    $fields = [Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $raw.Count; $i++) { $fields.Add([ordered]@{ Name = $raw[$i].Name; Value = $raw[$i].Value; Index = $i; Decoded = ConvertFrom-GctEncodedWord $raw[$i].Value }) }
    $all = { param([string]$Name) @($fields | Where-Object { $_.Name -ieq $Name }) }
    $first = { param([string]$Name) $f = @(& $all $Name); if ($f.Count) { $f[0].Value } else { $null } }
    $firstDecoded = { param([string]$Name) $f = @(& $all $Name); if ($f.Count) { $f[0].Decoded } else { $null } }
    $nz = { param($a, $b) if ($a) { $a } else { $b } }

    # Summary.
    $from = ConvertFrom-GctAddress (& $first 'From')
    $client = & $nz (& $firstDecoded 'X-Mailer') (& $firstDecoded 'User-Agent')
    $originating = ([string](& $first 'X-Originating-IP') -replace '[\[\]]', '').Trim()
    $summary = [ordered]@{
        Subject = & $firstDecoded 'Subject'
        From = $from
        Sender = ConvertFrom-GctAddress (& $first 'Sender')
        ReplyTo = ConvertFrom-GctAddressList (& $first 'Reply-To')
        ReturnPath = ConvertFrom-GctAddress (& $first 'Return-Path')
        To = ConvertFrom-GctAddressList (& $first 'To')
        Cc = ConvertFrom-GctAddressList (& $first 'Cc')
        Date = & $first 'Date'
        DateTime = ConvertFrom-GctMailDate (& $first 'Date')
        MessageId = $(if (& $first 'Message-ID') { ([string](& $first 'Message-ID')).Trim() } else { $null })
        NetworkMessageId = & $nz (& $first 'X-MS-Exchange-Organization-Network-Message-Id') (& $first 'X-MS-Exchange-CrossTenant-Network-Message-Id')
        TenantId = & $first 'X-MS-Exchange-CrossTenant-Id'
        Client = $client
        OriginatingIp = $(if ($originating) { $originating } else { $null })
        Directionality = & $first 'X-MS-Exchange-Organization-MessageDirectionality'
    }

    # Route: Received fields, oldest first.
    $receivedFields = @(& $all 'Received')
    [array]::Reverse($receivedFields)
    $hops = [Collections.Generic.List[object]]::new()
    foreach ($f in $receivedFields) {
        $h = ConvertFrom-GctReceived -Value $f.Value
        $h.FieldIndex = $f.Index
        $h.Index = $hops.Count
        foreach ($k in 'Segment', 'Product', 'IsEntry', 'ExchangeVersion', 'DelayMs', 'Skew', 'HeldBy', 'HeldSegment') { $h[$k] = $null }
        $hops.Add($h)
    }
    $entry = Set-GctHopSegment -Hops @($hops) -Codes $Codes
    $previous = $summary.DateTime
    $skew = $false
    for ($i = 0; $i -lt $hops.Count; $i++) {
        $h = $hops[$i]
        $h.Skew = $false
        $base = if ($i -eq 0) { $summary.DateTime } else { $previous }
        if ($null -ne $h.Time -and $null -ne $base) {
            $d = [long]$h.Time - [long]$base
            if ($d -lt -1000) { $h.Skew = $true; $skew = $true; $h.DelayMs = 0 } else { $h.DelayMs = [Math]::Max(0, $d) }
            $h.HeldBy = if ($i -eq 0) { 'the sender (client to first server)' } elseif ($hops[$i - 1].ByHost) { $hops[$i - 1].ByHost } else { 'previous server' }
            $h.HeldSegment = if ($i -eq 0) { 'sender' } else { $hops[$i - 1].Segment }
        }
        if ($null -ne $h.Time) { $previous = $h.Time }
    }
    $timed = @($hops | Where-Object { $null -ne $_.Time })
    $firstTime = if ($timed.Count) { $timed[0].Time } else { $null }
    $lastTime = if ($timed.Count) { $timed[-1].Time } else { $null }
    $biggest = $null
    foreach ($h in $hops) { if ($null -ne $h.DelayMs -and (-not $biggest -or $h.DelayMs -gt $biggest.DelayMs)) { $biggest = $h } }
    $route = [ordered]@{
        Hops = $hops.Count
        EntryIndex = $entry
        FirstTime = $firstTime
        LastTime = $lastTime
        TransitMs = $(if ($null -ne $firstTime -and $null -ne $lastTime) { $lastTime - $firstTime } else { $null })
        EndToEndMs = $(if ($null -ne $summary.DateTime -and $null -ne $lastTime) { $lastTime - $summary.DateTime } else { $null })
        Biggest = $(if ($biggest -and $biggest.DelayMs -gt 0) { [ordered]@{ Index = $biggest.Index; DelayMs = $biggest.DelayMs; HeldBy = $biggest.HeldBy; Segment = $biggest.HeldSegment } } else { $null })
        Skew = $skew
    }

    # Authentication: the topmost Authentication-Results is the verdict of the receiving organisation.
    $arFields = @(& $all 'Authentication-Results')
    $ar = if ($arFields.Count) { ConvertFrom-GctAuthResults $arFields[0].Value } else { $null }
    $pick = { param([string]$Method) @(if ($ar) { $ar.Results | Where-Object { $_.Method -eq $Method } }) }
    $spfR = @(& $pick 'spf') | Select-Object -First 1
    $dmarcR = @(& $pick 'dmarc') | Select-Object -First 1
    $compR = @(& $pick 'compauth') | Select-Object -First 1
    $rspfRaw = & $first 'Received-SPF'
    $rspf = $null
    if ($rspfRaw) {
        $rspf = ConvertFrom-GctTagList ((($rspfRaw.Trim() -replace '^\w+\s*', '') -replace '^\([^)]*\)\s*', ''))
        $rspf.Insert(0, 'result', ([regex]::Split($rspfRaw.Trim(), '[\s(;]')[0]).ToLowerInvariant())
    }
    $signatures = @(& $all 'DKIM-Signature' | ForEach-Object { $t = ConvertFrom-GctTagList $_.Value; [ordered]@{ D = ([string]$t['d']).ToLowerInvariant(); S = ([string]$t['s']).ToLowerInvariant(); A = $t['a']; I = $t['i'] } })
    $forefront = ConvertFrom-GctColonList (& $first 'X-Forefront-Antispam-Report')
    $senderIp = $null
    if ($spfR -and $spfR.Comment) { $sm = [regex]::Match($spfR.Comment, 'sender IP is ([0-9A-Fa-f:.]+)', 'IgnoreCase'); if ($sm.Success) { $senderIp = $sm.Groups[1].Value } }
    if (-not $senderIp -and $rspf -and $rspf['client-ip']) { $senderIp = $rspf['client-ip'] }
    if (-not $senderIp -and (Test-GctIpText $forefront['CIP']) -and $forefront['CIP'] -ne '255.255.255.255') { $senderIp = $forefront['CIP'] }
    if (-not $senderIp -and $entry -ge 0) { $senderIp = $hops[$entry].Ip }
    $mailFrom = if ($spfR -and $spfR.Props['smtp.mailfrom']) { $spfR.Props['smtp.mailfrom'] } elseif ($rspf -and $rspf['envelope-from']) { $rspf['envelope-from'] } elseif ($summary.ReturnPath) { $summary.ReturnPath.Address } else { $null }
    $mailFromDomain = Get-GctDomainOf $mailFrom
    $fromDomain = Get-GctDomainOf $(if ($dmarcR) { $dmarcR.Props['header.from'] })
    if (-not $fromDomain -and $from) { $fromDomain = $from.Domain }
    $dkimResults = @(& $pick 'dkim' | ForEach-Object {
            $d = if ($_.Props['header.d']) { $_.Props['header.d'] } elseif ($_.Props['header.i']) { $_.Props['header.i'] -replace '^.*@', '' } else { '' }
            [ordered]@{ Result = $_.Result; D = ([string]$d).ToLowerInvariant(); S = ([string]$_.Props['header.s']).ToLowerInvariant(); Comment = $_.Comment }
        })
    $orgFrom = Get-GctOrgDomain $fromDomain $Codes
    $align = { param($d) [ordered]@{ Relaxed = [bool]($d -and $fromDomain -and (Get-GctOrgDomain $d $Codes) -eq $orgFrom); Strict = [bool]($d -and $d -eq $fromDomain) } }
    $spfPass = ($spfR -and $spfR.Result -eq 'pass') -or ($rspf -and $rspf['result'] -eq 'pass')
    $spfAligned = & $align $mailFromDomain
    $dkimPassing = @($dkimResults | Where-Object { $_.Result -eq 'pass' -and $_.D -and $_.D -ne 'none' })
    $dkimAligned = [bool](@($dkimPassing | Where-Object { (& $align $_.D).Relaxed }).Count)
    $dkimAlignedStrict = [bool](@($dkimPassing | Where-Object { (& $align $_.D).Strict }).Count)
    $arc = @(& $all 'ARC-Seal' | ForEach-Object { $t = ConvertFrom-GctTagList $_.Value; [ordered]@{ I = [int]$(if ($t['i']) { $t['i'] } else { 0 }); Cv = ([string]$t['cv']).ToLowerInvariant(); D = ([string]$t['d']).ToLowerInvariant(); S = $t['s']; Authserv = $null; Results = $null } } | Sort-Object { $_.I })
    foreach ($f in @(& $all 'ARC-Authentication-Results')) {
        $im = [regex]::Match($f.Value, 'i\s*=\s*(\d+)', 'IgnoreCase')
        $seal = if ($im.Success) { $arc | Where-Object { $_.I -eq [int]$im.Groups[1].Value } | Select-Object -First 1 }
        if ($seal) {
            $parsed = ConvertFrom-GctAuthResults ($f.Value -replace '^\s*i\s*=\s*\d+\s*;?', '')
            $seal.Authserv = $parsed.Authserv
            $seal.Results = (@($parsed.Results | ForEach-Object { "$($_.Method)=$($_.Result)" }) -join ' ')
        }
    }
    $auth = [ordered]@{
        Present = [bool]$ar
        Authserv = $(if ($ar) { $ar.Authserv })
        Results = @(if ($ar) { $ar.Results })
        Others = @($arFields | Select-Object -Skip 1 | ForEach-Object { $_.Decoded })
        Spf = $(if ($spfR) { [ordered]@{ Result = $spfR.Result; Comment = $spfR.Comment } } elseif ($rspf) { [ordered]@{ Result = $rspf['result']; Comment = 'from Received-SPF' } } else { $null })
        SenderIp = $senderIp
        MailFrom = $mailFrom
        MailFromDomain = $mailFromDomain
        ReceivedSpf = $rspf
        Dkim = $dkimResults
        Signatures = $signatures
        Dmarc = $(if ($dmarcR) { [ordered]@{ Result = $dmarcR.Result; Action = ([string]$dmarcR.Props['action']).ToLowerInvariant(); HeaderFrom = $dmarcR.Props['header.from'] } } else { $null })
        Compauth = $(if ($compR) { [ordered]@{ Result = $compR.Result; Reason = $compR.Props['reason'] } } else { $null })
        Arc = $arc
        Alignment = [ordered]@{
            FromDomain = $fromDomain; OrgFromDomain = $orgFrom; MailFromDomain = $mailFromDomain; SpfPass = [bool]$spfPass
            SpfAligned = $spfAligned; DkimAligned = $dkimAligned; DkimAlignedStrict = $dkimAlignedStrict
            RelaxedPass = [bool](($spfPass -and $spfAligned.Relaxed) -or $dkimAligned)
            StrictPass = [bool](($spfPass -and $spfAligned.Strict) -or $dkimAlignedStrict)
        }
    }

    # Anti-spam and Exchange.
    $untrusted = ConvertFrom-GctColonList (& $first 'X-Forefront-Antispam-Report-Untrusted')
    $msAntispam = ConvertFrom-GctColonList (& $first 'X-Microsoft-Antispam')
    $msAntispamUntrusted = ConvertFrom-GctColonList (& $first 'X-Microsoft-Antispam-Untrusted')
    $delivery = ConvertFrom-GctColonList (& $first 'X-Microsoft-Antispam-Mailbox-Delivery')
    $exchange = [ordered]@{
        Scl = & $first 'X-MS-Exchange-Organization-SCL'
        AuthAs = & $first 'X-MS-Exchange-Organization-AuthAs'
        AuthSource = & $first 'X-MS-Exchange-Organization-AuthSource'
        AuthMechanism = & $first 'X-MS-Exchange-Organization-AuthMechanism'
        CrossTenantAuthAs = & $first 'X-MS-Exchange-CrossTenant-AuthAs'
        CrossTenantAuthSource = & $first 'X-MS-Exchange-CrossTenant-AuthSource'
        FromEntity = & $first 'X-MS-Exchange-CrossTenant-FromEntityHeader'
        CrossTenantId = $summary.TenantId
        OriginalArrival = & $first 'X-MS-Exchange-CrossTenant-OriginalArrivalTime'
        MailboxType = & $first 'X-MS-Exchange-CrossTenant-MailboxType'
        Directionality = $summary.Directionality
        SkipListed = & $first 'X-MS-Exchange-SkipListedInternetSender'
        ExternalOriginal = & $first 'X-MS-Exchange-ExternalOriginalInternetSender'
        OriginalClientIp = & $first 'X-MS-Exchange-Organization-OriginalClientIPAddress'
        OriginalServerIp = & $first 'X-MS-Exchange-Organization-OriginalServerIPAddress'
    }
    $antispam = [ordered]@{
        Report = $(if ($forefront.Count) { $forefront } else { $null })
        Untrusted = $(if ($untrusted.Count) { $untrusted } else { $null })
        Microsoft = $(if ($msAntispam.Count) { $msAntispam } else { $null })
        MicrosoftUntrusted = $(if ($msAntispamUntrusted.Count) { $msAntispamUntrusted } else { $null })
        Delivery = $(if ($delivery.Count) { $delivery } else { $null })
        Scl = $(if ($forefront.Contains('SCL')) { $forefront['SCL'] } else { $exchange.Scl })
        Bcl = $(if ($msAntispam.Contains('BCL')) { $msAntispam['BCL'] } else { $null })
    }

    # Third-party gateways and sending services.
    $gateways = [Collections.Generic.List[object]]::new()
    foreach ($g in $Codes['gateways']) {
        $byHeader = if ($g['headers']) { $fields | Where-Object { $_.Name -match $g['headers'] } | Select-Object -First 1 } else { $null }
        $hop = $null
        if ($g['hosts']) { $hop = $hops | Where-Object { @($_.ByHost, $_.FromHost, $_.Rdns) | Where-Object { $_ -and $_ -match $g['hosts'] } } | Select-Object -First 1 }
        if ($byHeader -or $hop) {
            $before = if ($hop) { ($entry -lt 0) -or ($hop.Index -lt $entry) -or ($hop.Index -eq $entry -and [string]$(if ($hop.FromHost) { $hop.FromHost } elseif ($hop.Rdns) { $hop.Rdns } else { '' }) -match $g['hosts']) } else { $null }
            $gateways.Add([ordered]@{ Name = $g['name']; Header = $(if ($byHeader) { $byHeader.Name } else { $null }); Hop = $(if ($hop) { $hop.Index } else { $null }); BeforeEntry = $before })
        }
    }
    $services = @($hops | Where-Object { $_.Segment -eq 'senderService' } | ForEach-Object { $_.Product } | Select-Object -Unique)
    $recipientDomain = Get-GctDomainOf $(if ($summary.To.Count) { $summary.To[0].Address })
    if (-not $recipientDomain) { $withFor = @($hops | Where-Object { $_.For }); if ($withFor.Count) { $recipientDomain = Get-GctDomainOf $withFor[-1].For } }

    # The first step of the chronology: the sender itself; the last: the mailbox.
    $firstHop = if ($hops.Count) { $hops[0] } else { $null }
    $origin = [ordered]@{
        Time = $summary.DateTime; DateText = $summary.Date; From = $from; Client = $client
        ClientIp = $(if ($summary.OriginatingIp) { $summary.OriginatingIp } elseif ($exchange.OriginalClientIp) { $exchange.OriginalClientIp } else { $null })
        Host = $(if ($firstHop) { $firstHop.FromHost }); Rdns = $(if ($firstHop) { $firstHop.Rdns }); Ip = $(if ($firstHop) { $firstHop.Ip }); FirstServer = $(if ($firstHop) { $firstHop.ByHost })
    }
    $delivered = [ordered]@{
        Time = $lastTime
        Folder = $(if ($antispam.Delivery) { $antispam.Delivery['dest'] } else { $null })
        Recipients = @($summary.To | ForEach-Object { $_.Address })
        Mailbox = $(if ($hops.Count) { $hops[-1].ByHost } else { $null })
    }
    $analysis = [ordered]@{
        Fields = @($fields); Summary = $summary; Origin = $origin; Hops = @($hops); Delivered = $delivered; Route = $route; Auth = $auth
        Antispam = $antispam; Exchange = $exchange; Gateways = @($gateways); Services = $services; RecipientDomain = $recipientDomain
        Findings = @(); CheckRequest = $null; Commands = @(); DnsToday = $null; DnsOffline = $false
    }
    $analysis.Findings = @(Get-GctHeaderFinding -Analysis $analysis -Codes $Codes)
    Set-GctFindingNode -Analysis $analysis
    $analysis.CheckRequest = Get-GctHeaderCheckRequest -Analysis $analysis -Codes $Codes
    $analysis.Commands = @(Get-GctHeaderCommand -Analysis $analysis)
    return $analysis
}

#endregion

#region Findings -----------------------------------------------------------------------------------

function ConvertTo-GctFindingStatus {
    param([AllowNull()][string]$Level)
    switch ($Level) { 'bad' { 'Failed' } 'warn' { 'Warning' } 'good' { 'Passed' } default { 'Info' } }
}

function Get-GctCodeRange {
    param([object[]]$Table, [AllowNull()][object]$Value)
    $n = 0.0
    if ($null -eq $Value -or -not [double]::TryParse([string]$Value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $null }
    foreach ($r in $Table) { if ($n -ge $r['min'] -and $n -le $r['max']) { return $r } }
    return $null
}

function Get-GctDelayTip {
    param([AllowNull()][string]$Segment)
    switch ($Segment) {
        'onPremises' { 'The on-premises Exchange held the message: look at its queues (Get-Queue) and at the message tracking log of the server (command below), and at the Send connector to Exchange Online (TLS, smart host, certificate).' }
        'gateway' { 'The gateway held the message: look at the message log of the gateway (queues, sandboxing of attachments, greylisting).' }
        { $_ -in 'entry', 'exchangeOnline' } { 'Microsoft 365 held the message: check the message trace (command below), Safe Attachments dynamic delivery and service health; a delay at the entry can also be throttling of the sending IP.' }
        'senderTenant' { 'The Microsoft 365 organisation of the sender held the message: its administrator can check its message trace.' }
        'beyond' { 'The servers after Microsoft 365 held the message (on-premises mailbox, archiving, journaling): look at their logs.' }
        default { 'The sending side held the message (client Outbox, sending server, greylisting at the first receiver, or retries after a refusal): ask the sender for the logs of its server.' }
    }
}

function Get-GctHeaderFinding {
    <# What the header tells, one finding each: status (Passed, Warning, Failed, Info), title, text, what to do, and what DNS can check today. #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Analysis, [Parameter(Mandatory = $true)][hashtable]$Codes)

    $list = [Collections.Generic.List[object]]::new()
    $add = { param([System.Collections.IDictionary]$F) $o = [ordered]@{ Id = $F.Id; Area = $F.Area; Status = $F.Status; Title = $F.Title; Text = $F.Text; Tip = $F['Tip']; Check = $F['Check']; Node = $null; Evidence = @(); Verdict = $null }; $list.Add($o) }
    $a = $Analysis
    $route = $a.Route; $hops = $a.Hops; $auth = $a.Auth; $antispam = $a.Antispam; $exchange = $a.Exchange; $summary = $a.Summary
    $seg = { param($k) if ($script:HeaderSegments.Contains([string]$k)) { $script:HeaderSegments[[string]$k] } else { [string]$k } }
    $dur = { param($ms) Format-GctHeaderDuration $ms }

    # Route.
    if (-not $hops.Count) {
        & $add @{ Id = 'route-none'; Area = 'Route'; Status = 'Warning'; Title = 'No Received field'; Text = 'The text has no Received field: the route of the message cannot be read.'; Tip = 'Paste all the headers of the message as the recipient sees them: in Outlook, open the message, File > Properties > Internet headers; in Outlook on the web, ... > View > View message details.' }
    }
    else {
        $b = $route.Biggest
        if ($b -and $b.DelayMs -ge 300000) { & $add @{ Id = 'route-delay'; Area = 'Route'; Status = 'Warning'; Title = "Longest wait: $(& $dur $b.DelayMs) at $($b.HeldBy)"; Text = "The message waited $(& $dur $b.DelayMs) before hop $($b.Index + 1), held by $($b.HeldBy) ($(& $seg $b.Segment)). The whole route took $(& $dur $route.TransitMs)."; Tip = Get-GctDelayTip $b.Segment } }
        elseif ($b -and $b.DelayMs -ge 60000) { & $add @{ Id = 'route-delay'; Area = 'Route'; Status = 'Info'; Title = "Longest wait: $(& $dur $b.DelayMs) at $($b.HeldBy)"; Text = "A wait of more than a minute before hop $($b.Index + 1), held by $($b.HeldBy) ($(& $seg $b.Segment)). The whole route took $(& $dur $route.TransitMs)."; Tip = Get-GctDelayTip $b.Segment } }
        else { & $add @{ Id = 'route-delay'; Area = 'Route'; Status = 'Passed'; Title = "Delivered in $(& $dur $(if ($null -ne $route.TransitMs) { $route.TransitMs } else { 0 }))"; Text = "$($hops.Count) hop(s), no wait of a minute or more$(if ($b) { "; the longest was $(& $dur $b.DelayMs) at $($b.HeldBy)" })." } }
        if ($null -ne $route.EndToEndMs -and $route.EndToEndMs -gt 900000 -and -not ($b -and $b.Index -eq 0)) { & $add @{ Id = 'route-latency'; Area = 'Route'; Status = 'Info'; Title = "Written $(& $dur $route.EndToEndMs) before its delivery"; Text = 'The Date field (set by the mail client when the message is sent) is far before the first Received: the message waited in the Outbox, was sent later, or the clock of the client is wrong.' } }
        if ($route.Skew) { & $add @{ Id = 'route-skew'; Area = 'Route'; Status = 'Info'; Title = 'Clocks of the servers disagree'; Text = 'At least one hop is dated before the previous one: the clocks of two servers differ. The delays around it are shown as 0; trust the times written by the receiving organisation.' } }
        $entryHop = if ($route.EntryIndex -ge 0) { $hops[$route.EntryIndex] } else { $null }
        if ($entryHop -and -not $entryHop.Tls -and [string]$entryHop.With -match '^E?SMTP$') {
            $who = if ($entryHop.FromHost) { $entryHop.FromHost } elseif ($entryHop.Ip) { $entryHop.Ip } else { 'The sending server' }
            & $add @{ Id = 'route-tls'; Area = 'Route'; Status = 'Info'; Title = 'Entry into Microsoft 365 without TLS'; Text = "$who sent the message to Microsoft 365 with $($entryHop.With), without TLS: the message travelled unencrypted on the Internet." }
        }
    }

    # Authentication.
    if (-not $auth.Present -and -not $auth.ReceivedSpf) { & $add @{ Id = 'auth-none'; Area = 'Authentication'; Status = 'Info'; Title = 'No authentication result'; Text = 'No Authentication-Results field: the message was read before its arrival in the receiving organisation (a copy from Sent Items), or it stayed inside one organisation.' } }
    $gatewayFront = [bool](@($a.Gateways | Where-Object { $_.BeforeEntry }).Count)
    if ($auth.Spf) {
        $r = [string]$auth.Spf.Result
        $code = if ($Codes['auth']['spf'].ContainsKey($r)) { $Codes['auth']['spf'][$r] } else { @{ level = 'warn'; text = "SPF result $r." } }
        $tip = if ($r -eq 'pass') { $null } elseif ($gatewayFront) { 'The IP checked is the one of the gateway in front of Microsoft 365, not the one of the sender: use Enhanced Filtering for Connectors on the inbound connector of the gateway, so that Microsoft 365 checks the original sender.' } else { "Add the sending service or IP to the SPF record of $(if ($auth.MailFromDomain) { $auth.MailFromDomain } else { 'the MAIL FROM domain' }), or send with a MAIL FROM whose domain allows it." }
        & $add @{ Id = 'spf'; Area = 'Authentication'; Status = ConvertTo-GctFindingStatus $code['level']; Title = "SPF $r$(if ($auth.MailFromDomain) { " for $($auth.MailFromDomain)" })"
            Text = "$($code['text'])$(if ($auth.SenderIp) { " Sending IP: $($auth.SenderIp)." })$(if ($auth.MailFrom) { " MAIL FROM: $($auth.MailFrom)." })"; Tip = $tip
            Check = $(if ($auth.MailFromDomain -and $auth.SenderIp) { [ordered]@{ Kind = 'spf'; Domain = $auth.MailFromDomain; Ip = $auth.SenderIp; Result = $r } }) }
    }
    if ($auth.Present -or $auth.Signatures.Count) {
        $dk = @($auth.Dkim)
        $passing = @($dk | Where-Object { $_.Result -eq 'pass' })
        $failing = @($dk | Where-Object { $_.Result -in 'fail', 'permerror' })
        $fromDom = $auth.Alignment.FromDomain
        $keys = @($auth.Signatures | Where-Object { $_.D -and $_.S } | ForEach-Object { [ordered]@{ Domain = $_.D; Selector = $_.S } })
        if ($failing.Count) {
            $r = $failing[0]
            $tip = if ($gatewayFront) { 'A gateway in front of Microsoft 365 that changes the message (subject tag, disclaimer, rewritten links) breaks the DKIM signature: let the gateway add an ARC seal and trust it in Microsoft 365, or use Enhanced Filtering for Connectors.' } else { 'Check that the key of the selector is published in DNS, and that no server on the way changes the message (disclaimer, footer, rewritten links).' }
            & $add @{ Id = 'dkim'; Area = 'Authentication'; Status = 'Failed'; Title = "DKIM $($r.Result)$(if ($r.D) { " for $($r.D)" })"; Text = "$($Codes['auth']['dkim'][$r.Result]['text'])$(if ($r.Comment) { " Microsoft 365 said: $($r.Comment)." })"; Tip = $tip
                Check = $(if ($keys.Count) { [ordered]@{ Kind = 'dkim'; Keys = $keys; Result = $r.Result } }) }
        }
        elseif ($passing.Count) {
            & $add @{ Id = 'dkim'; Area = 'Authentication'; Status = $(if ($auth.Alignment.DkimAligned) { 'Passed' } else { 'Info' }); Title = "DKIM pass for $(@($passing | ForEach-Object { $_.D }) -join ', ')"
                Text = "The DKIM signature was verified$(if ($auth.Alignment.DkimAligned) { ', for a domain aligned with the From address' } else { ", but for a domain that is not aligned with the From address ($fromDom)" })."
                Check = $(if ($keys.Count) { [ordered]@{ Kind = 'dkim'; Keys = $keys; Result = 'pass' } }) }
        }
        elseif ($dk.Count -or $auth.Present) {
            $ms = @(if ($fromDom) { [ordered]@{ Domain = $fromDom; Selector = 'selector1' }; [ordered]@{ Domain = $fromDom; Selector = 'selector2' } })
            $check = if ($keys.Count) { [ordered]@{ Kind = 'dkim'; Keys = $keys; Result = 'none' } } elseif ($ms.Count) { [ordered]@{ Kind = 'dkim'; Keys = $ms; Result = 'none'; ExchangeOnline = $true } } else { $null }
            & $add @{ Id = 'dkim'; Area = 'Authentication'; Status = 'Warning'; Title = 'Not signed with DKIM'; Text = "$($Codes['auth']['dkim']['none']['text'])$(if ($fromDom) { " DMARC then relies on SPF alone for $fromDom." })"
                Tip = "Sign the mail of $(if ($fromDom) { $fromDom } else { 'the domain' }) with DKIM: in Exchange Online, publish the CNAME records selector1._domainkey and selector2._domainkey and enable signing (Get-DkimSigningConfig, Set-DkimSigningConfig -Enabled `$true)."; Check = $check }
        }
    }
    if ($auth.Dmarc) {
        $d = $auth.Dmarc
        $c = if ($Codes['auth']['dmarc'].ContainsKey([string]$d.Result)) { $Codes['auth']['dmarc'][[string]$d.Result] } else { @{ level = 'warn'; text = "DMARC result $($d.Result)." } }
        $actionText = if ($d.Action) { if ($Codes['auth']['dmarcAction'].ContainsKey($d.Action)) { $Codes['auth']['dmarcAction'][$d.Action] } else { $d.Action } } else { $null }
        $al = $auth.Alignment
        $tip = $null
        if ($d.Result -eq 'fail') {
            $passingDkim = @($auth.Dkim | Where-Object { $_.Result -eq 'pass' })
            if ($al.SpfPass -and -not $al.SpfAligned.Relaxed) { $tip = "SPF passed for $($al.MailFromDomain), but that domain is not aligned with the From domain $($al.FromDomain). Send with a MAIL FROM in $($al.OrgFromDomain) (custom return-path of the sending service), or sign with DKIM d=$($al.OrgFromDomain)." }
            elseif ($passingDkim.Count -and -not $al.DkimAligned) { $tip = "The message is signed with DKIM for $(@($passingDkim | ForEach-Object { $_.D }) -join ', '), not for $($al.FromDomain): configure the sending service to sign with d=$($al.OrgFromDomain)." }
            else { $tip = "Neither SPF nor DKIM passed for $($al.FromDomain): fix SPF (the sending IP) or DKIM (the signature) for the From domain." }
        }
        elseif ($d.Result -in 'bestguesspass', 'none') { $tip = "Publish a DMARC record for $(if ($al.OrgFromDomain) { $al.OrgFromDomain } else { 'the domain' }) (_dmarc TXT, start with p=none and reports, then p=quarantine or p=reject)." }
        & $add @{ Id = 'dmarc'; Area = 'Authentication'; Status = ConvertTo-GctFindingStatus $c['level']; Title = "DMARC $($d.Result)$(if ($d.HeaderFrom) { " for $($d.HeaderFrom)" })"; Text = "$($c['text'])$(if ($actionText) { " Action: $actionText." })"; Tip = $tip
            Check = $(if ($al.FromDomain) { [ordered]@{ Kind = 'dmarc'; Domain = $al.FromDomain; OrgDomain = $al.OrgFromDomain; Result = $d.Result; Action = $d.Action } }) }
        if ($d.Result -eq 'fail' -and $al.SpfPass -and -not $al.SpfAligned.Relaxed) {
            & $add @{ Id = 'alignment'; Area = 'Authentication'; Status = 'Warning'; Title = 'SPF passed, but not for the From domain'; Text = "SPF checked $($al.MailFromDomain) (MAIL FROM), the reader sees $($al.FromDomain) (From). DMARC needs one of SPF or DKIM to pass for the From domain or a subdomain of $($al.OrgFromDomain) (relaxed alignment)."
                Check = [ordered]@{ Kind = 'dmarc'; Domain = $al.FromDomain; OrgDomain = $al.OrgFromDomain; Result = $d.Result; Alignment = $true } }
        }
    }
    if ($auth.Compauth) {
        $cr = $auth.Compauth
        $reason = $null
        if ($cr.Reason) { foreach ($r in $Codes['compauth']['reasons']) { if ($cr.Reason -match $r['match']) { $reason = $r; break } } }
        $status = switch ($cr.Result) { 'fail' { 'Failed' } 'softpass' { 'Warning' } 'pass' { 'Passed' } default { 'Info' } }
        $resultText = if ($Codes['compauth']['results'].ContainsKey([string]$cr.Result)) { $Codes['compauth']['results'][[string]$cr.Result] } else { '' }
        & $add @{ Id = 'compauth'; Area = 'Authentication'; Status = $status; Title = "Composite authentication $($cr.Result)$(if ($cr.Reason) { " (reason $($cr.Reason))" })"; Text = "$resultText$(if ($reason) { " $($reason['text'])" })"
            Tip = $(if ($cr.Result -eq 'fail') { 'Composite authentication combines SPF, DKIM, DMARC and the history of the sender: fix the failing check above. Do not allow-list the sender to work around it: an allow entry lets real spoofing in too.' })
            Check = $(if ($cr.Result -eq 'fail' -and $auth.Alignment.FromDomain) { [ordered]@{ Kind = 'dmarc'; Domain = $auth.Alignment.FromDomain; OrgDomain = $auth.Alignment.OrgFromDomain; Result = $(if ($auth.Dmarc) { $auth.Dmarc.Result } else { 'fail' }); Compauth = $true } }) }
    }
    if ($auth.Arc.Count) {
        $broken = $auth.Arc | Where-Object { $_.Cv -eq 'fail' } | Select-Object -First 1
        & $add @{ Id = 'arc'; Area = 'Authentication'; Status = $(if ($broken) { 'Warning' } else { 'Info' }); Title = "ARC chain of $($auth.Arc.Count) seal(s)"
            Text = $(if ($broken) { "The ARC chain is broken at seal $($broken.I) ($($broken.D)): a server changed the message after sealing it." } else { "Sealed by $(@($auth.Arc | ForEach-Object { "$($_.D) (i=$($_.I), cv=$($_.Cv))" }) -join ', '). ARC keeps the authentication results of the first servers through forwarders and gateways." })
            Tip = $(if ($broken) { 'Microsoft 365 can trust the seal of a known intermediary: add it to the trusted ARC sealers (Set-ArcConfig).' }) }
    }

    # Anti-spam.
    $rep = $antispam.Report
    if ($rep) {
        if ($rep['SFV']) {
            $s = $Codes['sfv'][$rep['SFV']]
            $tip = if ($rep['SFV'] -in 'SKA', 'SFE', 'SKN') { 'An allow entry skips spam filtering for every message that matches it, spoofed ones included. Prefer an entry of the Tenant Allow/Block List, limited in time, or fix the authentication of the sender.' } elseif ($rep['SFV'] -in 'SKB', 'BLK') { 'Check the blocked senders and domains of the anti-spam policy or of the mailbox if the sender is legitimate.' } else { $null }
            & $add @{ Id = 'sfv'; Area = 'Anti-spam'; Status = $(if ($s) { ConvertTo-GctFindingStatus $s['level'] } else { 'Info' }); Title = "Spam filtering verdict: $($rep['SFV'])"; Text = $(if ($s) { $s['text'] } else { 'Value not documented.' }); Tip = $tip }
        }
        $sclRow = Get-GctCodeRange -Table $Codes['scl'] -Value $antispam.Scl
        if ($sclRow -and [double]$antispam.Scl -ge 5) { & $add @{ Id = 'scl'; Area = 'Anti-spam'; Status = ConvertTo-GctFindingStatus $sclRow['level']; Title = "Spam confidence level $($antispam.Scl)"; Text = $sclRow['text']; Tip = 'If the message is legitimate, submit it to Microsoft (Submissions page of the Defender portal) and fix its authentication; do not lower the policy for everybody.' } }
        if ($rep['CAT'] -and $rep['CAT'] -ne 'NONE') {
            $catName = $Codes['cat'][$rep['CAT']]
            & $add @{ Id = 'cat'; Area = 'Anti-spam'; Status = $(if ($rep['CAT'] -in 'SPM', 'BULK', 'OSPM') { 'Warning' } else { 'Failed' }); Title = "Protection applied: $(if ($catName) { $catName } else { $rep['CAT'] })"
                Text = "The category $($rep['CAT'])$(if ($catName) { " ($catName)" }) decided the action on the message.$(if ($rep['SFTY'] -and $Codes['sfty'][$rep['SFTY']]) { " $($Codes['sfty'][$rep['SFTY']])" })"
                Tip = $(if ($rep['CAT'] -eq 'SPOOF') { 'Spoof intelligence found the From address not authenticated: fix SPF, DKIM and DMARC of the sender domain, or allow the spoofed sender pair in the Tenant Allow/Block List if the sending infrastructure is legitimate.' }) }
        }
        if ($rep['IPV'] -eq 'CAL') { & $add @{ Id = 'ipv'; Area = 'Anti-spam'; Status = 'Warning'; Title = 'IP in the IP Allow List'; Text = $Codes['ipv']['CAL']; Tip = 'The IP Allow List skips spam filtering for any message from the IP: keep it for your own servers only.' } }
        if ($rep['DIR']) {
            $dirText = if ($Codes['dir'][$rep['DIR']]) { $Codes['dir'][$rep['DIR']] } else { $rep['DIR'] }
            & $add @{ Id = 'dir'; Area = 'Anti-spam'; Status = 'Info'; Title = "Direction: $($rep['DIR'])"; Text = "$dirText$(if ($rep['CIP']) { "; connecting IP $($rep['CIP'])" })$(if ($rep['CTRY']) { " ($($rep['CTRY']))" })$(if ($rep['H']) { ", HELO $($rep['H'])" })." }
        }
    }
    $bclRow = Get-GctCodeRange -Table $Codes['bcl'] -Value $antispam.Bcl
    if ($bclRow -and [double]$antispam.Bcl -ge 4) {
        $threshold = if ($Codes['bclThresholdDefault']) { $Codes['bclThresholdDefault'] } else { 7 }
        & $add @{ Id = 'bcl'; Area = 'Anti-spam'; Status = $(if ([double]$antispam.Bcl -ge $threshold) { 'Warning' } else { 'Info' }); Title = "Bulk complaint level $($antispam.Bcl)"; Text = "$($bclRow['text']) The default threshold of the anti-spam policy is $($threshold): at or above it, the bulk action applies." }
    }
    if ($antispam.Untrusted -or $antispam.MicrosoftUntrusted) { & $add @{ Id = 'untrusted'; Area = 'Anti-spam'; Status = 'Info'; Title = 'Verdict of the sending organisation'; Text = 'The fields ending in -Untrusted were written by the Microsoft 365 organisation that sent the message (its outbound filtering), not by the receiving organisation. Read the fields without -Untrusted for the verdict of the recipient.' } }
    $dl = $antispam.Delivery
    if ($dl -and $dl['dest']) {
        $folder = if ($Codes['mailboxDelivery']['dest'][$dl['dest']]) { $Codes['mailboxDelivery']['dest'][$dl['dest']] } else { $dl['dest'] }
        $ofr = if ($dl['OFR']) { $Codes['mailboxDelivery']['ofr'][$dl['OFR']] } else { $null }
        $junk = $dl['dest'] -eq 'J'
        $tip = $null
        if ($junk -and $dl['OFR'] -eq 'CustomRules') { $tip = 'An inbox rule of the mailbox moved the message: check the rules (Get-InboxRule -Mailbox <address>).' }
        elseif ($junk -and $dl['OFR'] -eq 'ExclusiveSettings') { $tip = 'The Blocked Senders list of the mailbox matched: check Get-MailboxJunkEmailConfiguration -Identity <address> (BlockedSendersAndDomains).' }
        elseif ($junk -and $null -ne $antispam.Scl -and [double]$antispam.Scl -lt 5) { $tip = 'The service did not find spam (low SCL): the junk email settings of the mailbox moved it. Check the Blocked Senders and the junk filter of the mailbox.' }
        elseif ($junk) { $tip = 'The spam verdict of the service sent it to the Junk Email folder: see the anti-spam and authentication findings.' }
        & $add @{ Id = 'delivery'; Area = 'Delivery'; Status = $(if ($junk) { 'Warning' } else { 'Passed' }); Title = "Delivered to: $folder"
            Text = "$(if ($junk) { 'The message went to the Junk Email folder of the mailbox.' } else { "The mailbox filed the message in: $folder." })$(if ($ofr) { " $($ofr['text'])" } elseif ($dl['OFR']) { " Reason: $($dl['OFR'])." })"; Tip = $tip }
    }

    # Hybrid and Exchange.
    $onPremBeforeEntry = $route.EntryIndex -gt 0 -and [bool](@($hops | Select-Object -First $route.EntryIndex | Where-Object { $_.Segment -eq 'onPremises' }).Count)
    $fe = [string]$exchange.FromEntity
    $authAsText = if ($exchange.CrossTenantAuthAs) { ", AuthAs $($exchange.CrossTenantAuthAs)" } else { '' }
    if ($fe -eq 'HybridOnPrem') { & $add @{ Id = 'hybrid'; Area = 'Hybrid'; Status = 'Passed'; Title = 'Recognised as the on-premises organisation'; Text = "Exchange Online attributed the message to the on-premises organisation of the hybrid (FromEntityHeader HybridOnPrem$authAsText): the inbound connector from on-premises matched, and the message is treated as internal." } }
    elseif ($onPremBeforeEntry -and ($fe -eq 'Internet' -or ([string]$exchange.CrossTenantAuthAs).ToLowerInvariant() -eq 'anonymous')) {
        & $add @{ Id = 'hybrid'; Area = 'Hybrid'; Status = 'Warning'; Title = 'On-premises message not recognised'; Text = "The message came from Exchange on-premises but Exchange Online received it as Internet mail ($(if ($fe) { "FromEntityHeader $fe" })$authAsText): no inbound connector of type OnPremises matched."
            Tip = 'Check the inbound connector from on-premises in Exchange Online: ConnectorType OnPremises, TlsSenderCertificateName equal to the certificate name the on-premises Send connector presents (or SenderIPAddresses), RequireTls. Once it matches, the messages are attributed (HybridOnPrem, AuthAs Internal) and no longer filtered as anonymous Internet mail.' }
    }
    elseif ($fe) {
        $entity = if ($Codes['exchange']['fromEntity'][$fe]) { $Codes['exchange']['fromEntity'][$fe] } else { $fe }
        $direction = if ($exchange.Directionality) { ", $(if ($Codes['exchange']['directionality'][$exchange.Directionality]) { $Codes['exchange']['directionality'][$exchange.Directionality] } else { $exchange.Directionality })" } else { '' }
        & $add @{ Id = 'hybrid'; Area = 'Hybrid'; Status = 'Info'; Title = "Attributed to: $entity"; Text = "FromEntityHeader $fe$authAsText$direction." }
    }
    if ($exchange.AuthAs -and -not $fe) {
        $c = $Codes['exchange']['authAs'][$exchange.AuthAs]
        & $add @{ Id = 'authas'; Area = 'Hybrid'; Status = 'Info'; Title = "Authenticated as: $($exchange.AuthAs)"; Text = "$(if ($c) { $c['text'] })$(if ($exchange.AuthSource) { " Server: $($exchange.AuthSource)." })" }
    }

    # Gateways.
    $front = @($a.Gateways | Where-Object { $_.BeforeEntry })
    $other = @($a.Gateways | Where-Object { -not $_.BeforeEntry })
    if ($front.Count) {
        $names = @($front | ForEach-Object { $_.Name }) -join ', '
        $mx = if ($a.RecipientDomain) { [ordered]@{ Kind = 'mx'; Domain = $a.RecipientDomain; Gateway = $front[0].Name } } else { $null }
        if ($exchange.SkipListed) { & $add @{ Id = 'gateway'; Area = 'Gateway'; Status = 'Passed'; Title = "Enhanced Filtering skips $names"; Text = "The gateway is in front of Microsoft 365 and Enhanced Filtering for Connectors is on: Microsoft 365 skipped its IP and checked the original sender ($($exchange.SkipListed))."; Check = $mx } }
        else {
            & $add @{ Id = 'gateway'; Area = 'Gateway'; Status = $(if ($auth.Spf -and $auth.Spf.Result -ne 'pass') { 'Warning' } else { 'Info' }); Title = "$names in front of Microsoft 365"
                Text = "The message reached Microsoft 365 through $names. Without Enhanced Filtering for Connectors, Microsoft 365 checks SPF and IP reputation on the IP of the gateway, not of the sender."
                Tip = 'On the inbound connector of the gateway, enable Enhanced Filtering for Connectors (skip listing of the gateway IPs), or point the MX directly to Exchange Online Protection.'; Check = $mx }
        }
    }
    if ($other.Count) { & $add @{ Id = 'gateway-other'; Area = 'Gateway'; Status = 'Info'; Title = "Also seen: $(@($other | ForEach-Object { $_.Name }) -join ', ')"; Text = "Fields or servers of $(@($other | ForEach-Object { $_.Name }) -join ', ') are in the message ($(@($other | ForEach-Object { if ($_.Header) { $_.Header } else { 'route' } }) -join ', ')): usually the filtering service of the sender, or a service after Microsoft 365." } }
    if ($exchange.ExternalOriginal) { & $add @{ Id = 'original-sender'; Area = 'Gateway'; Status = 'Info'; Title = 'Original sender of the Internet'; Text = "Microsoft 365 recorded the original sender before the skipped servers: $($exchange.ExternalOriginal)." } }

    # Sender identity.
    $fromDomain = if ($summary.From) { $summary.From.Domain } else { $null }
    $replyOther = @($summary.ReplyTo | Where-Object { $_.Domain -and $fromDomain -and (Get-GctOrgDomain $_.Domain $Codes) -ne (Get-GctOrgDomain $fromDomain $Codes) })
    if ($replyOther.Count) { & $add @{ Id = 'reply-to'; Area = 'Sender'; Status = 'Info'; Title = 'Replies go to another domain'; Text = "Reply-To is $(@($replyOther | ForEach-Object { $_.Address }) -join ', '), while the From domain is $fromDomain. Normal for some services (ticketing, newsletters); a frequent sign of phishing otherwise." } }
    if ($summary.Sender -and $summary.Sender.Address -and $summary.From -and $summary.From.Address -and $summary.Sender.Address.ToLowerInvariant() -ne $summary.From.Address.ToLowerInvariant()) {
        & $add @{ Id = 'sender'; Area = 'Sender'; Status = 'Info'; Title = 'Sent on behalf of'; Text = "The Sender field ($($summary.Sender.Address)) differs from From ($($summary.From.Address)): the message was sent on behalf of the From address (delegate, Send on Behalf, or a service)." }
    }
    return @($list)
}

$script:HeaderEvidence = @{
    'spf' = @('Authentication-Results', 'Received-SPF'); 'dkim' = @('Authentication-Results', 'DKIM-Signature'); 'dmarc' = @('Authentication-Results')
    'alignment' = @('Authentication-Results', 'Return-Path', 'From'); 'compauth' = @('Authentication-Results'); 'arc' = @('ARC-Seal', 'ARC-Authentication-Results')
    'sfv' = @('X-Forefront-Antispam-Report'); 'scl' = @('X-Forefront-Antispam-Report', 'X-MS-Exchange-Organization-SCL'); 'cat' = @('X-Forefront-Antispam-Report')
    'ipv' = @('X-Forefront-Antispam-Report'); 'dir' = @('X-Forefront-Antispam-Report'); 'bcl' = @('X-Microsoft-Antispam')
    'untrusted' = @('X-Forefront-Antispam-Report-Untrusted', 'X-Microsoft-Antispam-Untrusted'); 'delivery' = @('X-Microsoft-Antispam-Mailbox-Delivery')
    'hybrid' = @('X-MS-Exchange-CrossTenant-FromEntityHeader', 'X-MS-Exchange-CrossTenant-AuthAs', 'X-MS-Exchange-Organization-AuthAs')
    'authas' = @('X-MS-Exchange-Organization-AuthAs', 'X-MS-Exchange-Organization-AuthSource'); 'gateway' = @('X-MS-Exchange-SkipListedInternetSender')
    'original-sender' = @('X-MS-Exchange-ExternalOriginalInternetSender'); 'reply-to' = @('Reply-To', 'From'); 'sender' = @('Sender', 'From'); 'route-latency' = @('Date')
}

function Set-GctFindingNode {
    <#
        Puts each finding on the step of the chronology it is about: 'origin' (the sender), a hop index, or 'delivery'
        (the mailbox). Authentication and attribution are decided where the message enters Microsoft 365; the spam
        verdict and the folder show at the delivery.
    #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Analysis)
    $a = $Analysis
    $entry = if ($a.Route.EntryIndex -ge 0) { [int]$a.Route.EntryIndex } else { $null }
    $lastSenderTenant = -1
    $firstSkew = -1
    for ($i = 0; $i -lt $a.Hops.Count; $i++) {
        if ($a.Hops[$i].Segment -eq 'senderTenant') { $lastSenderTenant = $i }
        if ($a.Hops[$i].Skew -and $firstSkew -lt 0) { $firstSkew = $i }
    }
    $orDelivery = { param($n) if ($null -ne $n) { $n } else { 'delivery' } }
    foreach ($f in $a.Findings) {
        $node = switch ($f.Id) {
            { $_ -in 'spf', 'dkim', 'dmarc', 'alignment', 'compauth', 'arc', 'auth-none', 'hybrid', 'authas', 'original-sender', 'dir', 'ipv', 'route-tls' } { & $orDelivery $entry; break }
            { $_ -in 'gateway', 'gateway-other' } {
                $g = if ($_ -eq 'gateway') { $a.Gateways | Where-Object { $_.BeforeEntry } | Select-Object -First 1 } else { $a.Gateways | Where-Object { -not $_.BeforeEntry } | Select-Object -First 1 }
                if ($g -and $null -ne $g.Hop) { [int]$g.Hop } else { & $orDelivery $entry }
                break
            }
            'untrusted' { if ($lastSenderTenant -ge 0) { $lastSenderTenant } else { 'origin' }; break }
            'route-delay' { if ($a.Route.Biggest) { [int]$a.Route.Biggest.Index } else { 'delivery' }; break }
            'route-skew' { if ($firstSkew -ge 0) { $firstSkew } else { 'delivery' }; break }
            { $_ -in 'route-latency', 'reply-to', 'sender', 'route-none' } { 'origin'; break }
            default { 'delivery' }
        }
        $f.Node = $node
        $names = [Collections.Generic.List[string]]::new()
        if ($script:HeaderEvidence.ContainsKey($f.Id)) { foreach ($n in $script:HeaderEvidence[$f.Id]) { $names.Add($n) } }
        if ($f.Id -like 'gateway*') { foreach ($g in $a.Gateways) { if ($g.Header) { $names.Add($g.Header) } } }
        $f.Evidence = @($names)
    }
}

function Get-GctHeaderNodeFinding {
    <# The findings of one step of the chronology ('origin', a hop index or 'delivery'). #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Analysis, [Parameter(Mandatory = $true)][object]$Node)
    return , @($Analysis.Findings | Where-Object { ($Node -is [string] -and $_.Node -is [string] -and $_.Node -eq $Node) -or ($Node -isnot [string] -and $_.Node -isnot [string] -and [int]$_.Node -eq [int]$Node) })
}

function Get-GctHeaderCheckRequest {
    <# What the DNS cross-check asks: the MAIL FROM domain and the sending IP, the From domain, DKIM keys, the recipient domain. #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Analysis, [hashtable]$Codes)
    $req = [ordered]@{ MailFromDomain = $null; Ip = $null; FromDomain = $null; FromOrgDomain = $null; RecipientDomain = $null; Dkim = [Collections.Generic.List[object]]::new() }
    foreach ($f in $Analysis.Findings) {
        $c = $f.Check
        if (-not $c) { continue }
        switch ($c.Kind) {
            'spf' { $req.MailFromDomain = $c.Domain; $req.Ip = $c.Ip }
            'dmarc' { $req.FromDomain = $c.Domain; $req.FromOrgDomain = $c.OrgDomain }
            'dkim' { foreach ($k in $c.Keys) { if (-not @($req.Dkim | Where-Object { $_.Domain -eq $k.Domain -and $_.Selector -eq $k.Selector }).Count -and $req.Dkim.Count -lt 4) { $req.Dkim.Add($k) } } }
            'mx' { $req.RecipientDomain = $c.Domain }
        }
    }
    if (-not $req.FromDomain -and $Analysis.Auth.Alignment.FromDomain) { $req.FromDomain = $Analysis.Auth.Alignment.FromDomain; $req.FromOrgDomain = Get-GctOrgDomain $Analysis.Auth.Alignment.FromDomain $Codes }
    $req.Dkim = @($req.Dkim)
    if ($req.MailFromDomain -or $req.FromDomain -or $req.RecipientDomain -or $req.Dkim.Count) { return $req }
    return $null
}

function Get-GctHeaderCommand {
    <# The commands to follow the message: message trace of Exchange Online, message tracking log on-premises, Defender. #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Analysis)
    $id = $Analysis.Summary.MessageId
    $when = if ($null -ne $Analysis.Route.LastTime) { $Analysis.Route.LastTime } else { $Analysis.Summary.DateTime }
    $day = { param([int]$Days) $base = if ($null -ne $when) { [DateTimeOffset]::FromUnixTimeMilliseconds([long]$when) } else { [DateTimeOffset]::UtcNow }; $base.AddDays($Days).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) }
    $out = [Collections.Generic.List[object]]::new()
    if ($id) {
        $q = $id.Replace("'", "''")
        $out.Add([ordered]@{ Title = 'Exchange Online: message trace'; Text = "Get-MessageTraceV2 -MessageId '$q' -StartDate '$(& $day -1)' -EndDate '$(& $day 1)' |`n    Format-Table Received, SenderAddress, RecipientAddress, Status, MessageTraceId -AutoSize"; Note = 'Exchange Online PowerShell (Connect-ExchangeOnline), messages of the last 90 days. Then Get-MessageTraceDetailV2 -MessageTraceId <id> -RecipientAddress <address> for the events.' })
        $out.Add([ordered]@{ Title = 'Exchange on-premises: message tracking log'; Text = "Get-ExchangeServer | Where-Object { `$_.IsHubTransportServer -or `$_.IsMailboxServer } |`n    Get-MessageTrackingLog -MessageId '$q' -Start '$(& $day -1)' -End '$(& $day 1)' -ResultSize Unlimited |`n    Sort-Object Timestamp | Format-Table Timestamp, ServerHostname, EventId, Source, ConnectorId, RecipientStatus -AutoSize"; Note = 'Exchange Management Shell on-premises, when a hop of the route is on-premises.' })
    }
    if ($Analysis.Summary.NetworkMessageId) { $out.Add([ordered]@{ Title = 'Microsoft Defender: the same message'; Text = $Analysis.Summary.NetworkMessageId; Note = 'Network Message ID: search it in Explorer or Advanced hunting (EmailEvents | where NetworkMessageId == "...") of the Defender portal.' }) }
    return @($out)
}

#endregion

#region SPF evaluation (RFC 7208) ------------------------------------------------------------------

function Test-GctIpInNetwork {
    param([Parameter(Mandatory = $true)][Net.IPAddress]$Ip, [Parameter(Mandatory = $true)][Net.IPAddress]$Network, [int]$Prefix)
    if ($Ip.AddressFamily -ne $Network.AddressFamily) { return $false }
    $a = $Ip.GetAddressBytes(); $b = $Network.GetAddressBytes()
    $full = [Math]::Floor($Prefix / 8)
    for ($i = 0; $i -lt $full; $i++) { if ($a[$i] -ne $b[$i]) { return $false } }
    $bits = $Prefix % 8
    if ($bits -eq 0) { return $true }
    $mask = (0xFF -shl (8 - $bits)) -band 0xFF
    return (($a[$full] -band $mask) -eq ($b[$full] -band $mask))
}

function Expand-GctSpfMacro {
    <# Macros of RFC 7208, section 7 (%{s} %{l} %{o} %{d} %{i} %{p} %{v} %{h}, digits, r, delimiters). #>
    param([Parameter(Mandatory = $true)][hashtable]$State, [Parameter(Mandatory = $true)][string]$Spec, [Parameter(Mandatory = $true)][string]$Domain)
    if (-not $Spec.Contains('%')) { return $Spec.TrimEnd('.') }
    $sb = [Text.StringBuilder]::new()
    for ($i = 0; $i -lt $Spec.Length; $i++) {
        $c = $Spec[$i]
        if ($c -ne [char]'%') { [void]$sb.Append($c); continue }
        $i++
        if ($i -ge $Spec.Length) { throw "SPFPERM|Bad macro in '$Spec'." }
        switch -CaseSensitive ([string]$Spec[$i]) {
            '%' { [void]$sb.Append('%') }
            '_' { [void]$sb.Append(' ') }
            '-' { [void]$sb.Append('%20') }
            '{' {
                $end = $Spec.IndexOf('}', $i)
                if ($end -lt 0) { throw "SPFPERM|Bad macro in '$Spec'." }
                $m = [regex]::Match($Spec.Substring($i + 1, $end - $i - 1), '^([slodipvhSLODIPVH])(\d*)(r?)([.\-+,/_=]*)$')
                if (-not $m.Success) { throw "SPFPERM|Bad macro in '$Spec'." }
                $mailFrom = $State.Sender
                $v4 = $State.Ip.AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork
                $value = switch ($m.Groups[1].Value.ToLowerInvariant()) {
                    's' { $mailFrom }
                    'l' { $mailFrom.Substring(0, $mailFrom.LastIndexOf('@')) }
                    'o' { $mailFrom.Substring($mailFrom.LastIndexOf('@') + 1) }
                    'd' { $Domain }
                    'i' { if ($v4) { $State.Ip.ToString() } else { (([Convert]::ToHexString($State.Ip.GetAddressBytes())).ToLowerInvariant().ToCharArray()) -join '.' } }
                    'p' { 'unknown' }
                    'v' { if ($v4) { 'in-addr' } else { 'ip6' } }
                    'h' { if ($State.Helo) { $State.Helo } else { $Domain } }
                }
                $delims = if ($m.Groups[4].Value) { $m.Groups[4].Value.ToCharArray() } else { @([char]'.') }
                $parts = [Collections.Generic.List[string]]::new()
                foreach ($p in $value.Split($delims)) { $parts.Add($p) }
                if ($m.Groups[3].Value) { $parts.Reverse() }
                if ($m.Groups[2].Value) {
                    $n = [int]$m.Groups[2].Value
                    if ($n -eq 0) { throw "SPFPERM|Bad macro in '$Spec'." }
                    if ($n -lt $parts.Count) { $parts = [Collections.Generic.List[string]]@($parts | Select-Object -Last $n) }
                }
                $text = $parts -join '.'
                [void]$sb.Append($(if ([char]::IsUpper($m.Groups[1].Value[0])) { [Uri]::EscapeDataString($text) } else { $text }))
                $i = $end
            }
            default { throw "SPFPERM|Bad macro in '$Spec'." }
        }
    }
    $expanded = $sb.ToString().TrimEnd('.')
    while ($expanded.Length -gt 253 -and $expanded.Contains('.')) { $expanded = $expanded.Substring($expanded.IndexOf('.') + 1) }
    return $expanded
}

function Invoke-GctSpfQuery {
    <# One DNS query of the SPF evaluation (traced by Resolve-GctDns): temperror when DNS does not answer. #>
    param([hashtable]$State, [string]$Name, [string]$Type)
    $q = Resolve-GctDns -Name $Name -Type $Type
    if ($q.Error) { throw "SPFTEMP|DNS did not answer for $Name ($Type): $($q.Error)" }
    return @($q.Records | ForEach-Object { [string]$_.Data })
}

function Add-GctSpfLookup {
    param([hashtable]$State)
    $State.Lookups++
    if ($State.Lookups -gt 10) { throw 'SPFPERM|More than 10 DNS lookups: receivers give permerror.' }
}
function Add-GctSpfVoid {
    param([hashtable]$State)
    $State.Voids++
    if ($State.Voids -gt 2) { throw 'SPFPERM|More than 2 lookups without answer (void lookups).' }
}

function Test-GctSpfAddress {
    param([hashtable]$State, [string]$HostName, [int]$Cidr4, [int]$Cidr6, [bool]$CountVoid)
    $v4 = $State.Ip.AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork
    $records = @(Invoke-GctSpfQuery -State $State -Name $HostName -Type $(if ($v4) { 'A' } else { 'AAAA' }))
    if (-not $records.Count) { if ($CountVoid) { Add-GctSpfVoid $State }; return $false }
    foreach ($r in $records) {
        $addr = $null
        if ([Net.IPAddress]::TryParse($r, [ref]$addr) -and (Test-GctIpInNetwork -Ip $State.Ip -Network $addr -Prefix $(if ($v4) { $Cidr4 } else { $Cidr6 }))) { return $true }
    }
    return $false
}

function Split-GctSpfCidr {
    param([AllowEmptyString()][string]$Arg)
    $c4 = 32; $c6 = 128; $s = $Arg
    $dbl = $s.IndexOf('//')
    if ($dbl -ge 0) {
        if (-not [int]::TryParse($s.Substring($dbl + 2), [ref]$c6) -or $c6 -lt 0 -or $c6 -gt 128) { throw "SPFPERM|Bad IPv6 prefix in '$Arg'." }
        $s = $s.Substring(0, $dbl)
    }
    $one = $s.LastIndexOf('/')
    if ($one -ge 0) {
        if (-not [int]::TryParse($s.Substring($one + 1), [ref]$c4) -or $c4 -lt 0 -or $c4 -gt 32) { throw "SPFPERM|Bad IPv4 prefix in '$Arg'." }
        $s = $s.Substring(0, $one)
    }
    return @{ Domain = $s; Cidr4 = $c4; Cidr6 = $c6 }
}

function Invoke-GctSpfCheckHost {
    param([hashtable]$State, [string]$Domain, [int]$Depth)
    if ($Depth -gt 12) { throw 'SPFPERM|The includes and redirects go too deep (a loop?).' }
    $records = @(Invoke-GctSpfQuery -State $State -Name $Domain -Type 'TXT' | Where-Object { $_ -ieq 'v=spf1' -or $_ -like 'v=spf1 *' })
    if (-not $records.Count) { return @{ Result = 'none'; Mechanism = $null } }
    if ($records.Count -gt 1) { throw "SPFPERM|$Domain has $($records.Count) SPF records: only one is allowed." }
    $record = $records[0]
    if ($Depth -eq 0) { $State.Record = $record }
    $redirect = $null
    foreach ($term in @($record.Split(' ', [StringSplitOptions]::RemoveEmptyEntries) | Select-Object -Skip 1)) {
        $mod = [regex]::Match($term, '^([a-zA-Z][a-zA-Z0-9_.-]*)=(.*)$')
        if ($mod.Success) {
            if ($mod.Groups[1].Value -ieq 'redirect') {
                if ($null -ne $redirect) { throw "SPFPERM|$($Domain): two redirect= modifiers." }
                $redirect = $mod.Groups[2].Value
            }
            continue
        }
        $qualifier = if ('+-~?'.Contains($term[0])) { $term[0] } else { [char]'+' }
        $body = if ('+-~?'.Contains($term[0])) { $term.Substring(1) } else { $term }
        $cut = $body.IndexOfAny([char[]]@(':', '/'))
        $mech = $(if ($cut -lt 0) { $body } else { $body.Substring(0, $cut) }).ToLowerInvariant()
        $rest = if ($cut -lt 0) { '' } else { $body.Substring($cut) }
        $arg = if ($rest.StartsWith(':')) { $rest.Substring(1) } else { $rest }
        $matched = $null
        switch ($mech) {
            'all' { if ($rest) { throw "SPFPERM|$($Domain): bad term '$term'." }; $matched = $term }
            'include' {
                if (-not $arg) { throw "SPFPERM|$($Domain): include without a domain." }
                Add-GctSpfLookup $State
                $target = Expand-GctSpfMacro -State $State -Spec $arg -Domain $Domain
                $r = Invoke-GctSpfCheckHost -State $State -Domain $target -Depth ($Depth + 1)
                if ($r.Result -eq 'pass') { $matched = "include:$target" + $(if ($r.Mechanism) { " > $($r.Mechanism)" } else { '' }) }
                elseif ($r.Result -eq 'none') { throw "SPFPERM|include:$target has no SPF record." }
            }
            { $_ -in 'a', 'mx' } {
                Add-GctSpfLookup $State
                $parts = Split-GctSpfCidr $arg
                $target = if ($parts.Domain) { Expand-GctSpfMacro -State $State -Spec $parts.Domain -Domain $Domain } else { $Domain }
                if ($mech -eq 'a') { if (Test-GctSpfAddress -State $State -HostName $target -Cidr4 $parts.Cidr4 -Cidr6 $parts.Cidr6 -CountVoid $true) { $matched = $term } }
                else {
                    $mx = @(Invoke-GctSpfQuery -State $State -Name $target -Type 'MX')
                    if (-not $mx.Count) { Add-GctSpfVoid $State }
                    if ($mx.Count -gt 10) { throw "SPFPERM|$target has more than 10 MX records." }
                    foreach ($h in $mx) {
                        $hostName = ($h.Split(' ', 2)[-1]).TrimEnd('.')
                        if (Test-GctSpfAddress -State $State -HostName $hostName -Cidr4 $parts.Cidr4 -Cidr6 $parts.Cidr6 -CountVoid $false) { $matched = $term; break }
                    }
                }
            }
            'ptr' { Add-GctSpfLookup $State; $State.Notes.Add('ptr is not evaluated (deprecated by RFC 7208): it is taken as not matching.') }
            { $_ -in 'ip4', 'ip6' } {
                $slash = $arg.IndexOf('/')
                $addrText = if ($slash -lt 0) { $arg } else { $arg.Substring(0, $slash) }
                $net = $null
                $family = if ($mech -eq 'ip4') { [Net.Sockets.AddressFamily]::InterNetwork } else { [Net.Sockets.AddressFamily]::InterNetworkV6 }
                if (-not [Net.IPAddress]::TryParse($addrText, [ref]$net) -or $net.AddressFamily -ne $family) { throw "SPFPERM|$($Domain): bad address in '$term'." }
                $max = if ($mech -eq 'ip4') { 32 } else { 128 }
                $prefix = $max
                if ($slash -ge 0 -and (-not [int]::TryParse($arg.Substring($slash + 1), [ref]$prefix) -or $prefix -lt 0 -or $prefix -gt $max)) { throw "SPFPERM|$($Domain): bad prefix in '$term'." }
                if (Test-GctIpInNetwork -Ip $State.Ip -Network $net -Prefix $prefix) { $matched = $term }
            }
            'exists' {
                if (-not $arg) { throw "SPFPERM|$($Domain): exists without a domain." }
                Add-GctSpfLookup $State
                $target = Expand-GctSpfMacro -State $State -Spec $arg -Domain $Domain
                if (@(Invoke-GctSpfQuery -State $State -Name $target -Type 'A').Count) { $matched = "exists:$target" } else { Add-GctSpfVoid $State }
            }
            default { throw "SPFPERM|$($Domain): unknown mechanism '$term'." }
        }
        if ($matched) {
            $result = switch ([string]$qualifier) { '-' { 'fail' } '~' { 'softfail' } '?' { 'neutral' } default { 'pass' } }
            return @{ Result = $result; Mechanism = $matched }
        }
    }
    if ($null -ne $redirect) {
        Add-GctSpfLookup $State
        $target = Expand-GctSpfMacro -State $State -Spec $redirect -Domain $Domain
        $r = Invoke-GctSpfCheckHost -State $State -Domain $target -Depth ($Depth + 1)
        if ($r.Result -eq 'none') { throw "SPFPERM|redirect=$target has no SPF record." }
        return @{ Result = $r.Result; Mechanism = "redirect=$target" + $(if ($r.Mechanism) { " > $($r.Mechanism)" } else { '' }) }
    }
    return @{ Result = 'neutral'; Mechanism = $null }
}

function Test-GctSpf {
    <#
    .SYNOPSIS
        SPF of a domain for an IP address, as a receiver evaluates it today (RFC 7208): ip4, ip6, a, mx, include,
        exists, redirect, the four qualifiers, macros, 10 DNS lookups and 2 void lookups at most. ptr is not
        evaluated (deprecated): it never matches, and a note says so. Returns Result, Mechanism, Record, Lookups,
        Error and Notes.
    #>
    param([Parameter(Mandatory = $true)][string]$Ip, [Parameter(Mandatory = $true)][string]$Domain, [string]$MailFrom, [string]$Helo)
    $address = $null
    if (-not [Net.IPAddress]::TryParse($Ip, [ref]$address)) { throw "Not an IP address: $Ip" }
    if ($address.IsIPv4MappedToIPv6) { $address = $address.MapToIPv4() }
    $state = @{ Ip = $address; Sender = $(if ($MailFrom -and $MailFrom.Contains('@')) { $MailFrom } else { "postmaster@$Domain" }); Helo = $Helo; Lookups = 0; Voids = 0; Record = $null; Notes = [Collections.Generic.List[string]]::new() }
    $result = 'neutral'; $mechanism = $null; $err = $null
    try { $r = Invoke-GctSpfCheckHost -State $state -Domain $Domain -Depth 0; $result = $r.Result; $mechanism = $r.Mechanism }
    catch {
        $msg = $_.Exception.Message
        if ($msg.StartsWith('SPFPERM|')) { $result = 'permerror'; $err = $msg.Substring(8) }
        elseif ($msg.StartsWith('SPFTEMP|')) { $result = 'temperror'; $err = $msg.Substring(8) }
        else { throw }
    }
    [pscustomobject]@{ Result = $result; Mechanism = $mechanism; Record = $state.Record; Lookups = $state.Lookups; Error = $err; Notes = @($state.Notes | Select-Object -Unique) }
}

#endregion

#region DNS of today and the verdicts --------------------------------------------------------------

function ConvertFrom-GctDnsTag {
    param([AllowEmptyString()][string]$Record)
    $out = @{}
    foreach ($t in $Record.Split(';')) {
        $p = $t.Split('=', 2)
        if ($p.Count -eq 2 -and $p[0].Trim() -and -not $out.ContainsKey($p[0].Trim().ToLowerInvariant())) { $out[$p[0].Trim().ToLowerInvariant()] = $p[1].Trim() }
    }
    return $out
}

function Get-GctDkimKeyBits {
    param([string]$KeyType, [string]$Base64)
    if ($KeyType -eq 'ed25519') { return 256 }
    try {
        $length = [Convert]::FromBase64String($Base64).Length
        $estimate = ($length - 38) * 8
        return (@(512, 1024, 2048, 3072, 4096) | Sort-Object { [Math]::Abs($_ - $estimate) } | Select-Object -First 1)
    }
    catch { return $null }
}

function Invoke-GctStageHeaderRead {
    <# The header block: its fields, the message, the commands to follow it. Without a field, the test stops. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'HeaderRead'
    $cfg = $Context.Config
    $source = 'text pasted'
    $text = [string]$cfg.HeaderText
    if (-not $text.Trim()) {
        $path = [string]$cfg.HeaderFile
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Header file not found: $path" }
        $item = Get-Item -LiteralPath $path
        if ($item.Length -gt $script:HeaderMaxBytes) {
            # A big .eml: the header block is at the start, read only its beginning.
            $stream = [IO.File]::OpenRead($item.FullName)
            try { $buffer = [byte[]]::new($script:HeaderMaxBytes); $n = $stream.Read($buffer, 0, $buffer.Length); $text = [Text.Encoding]::UTF8.GetString($buffer, 0, $n) } finally { $stream.Dispose() }
        }
        else { $text = [IO.File]::ReadAllText($item.FullName) }
        $source = $item.FullName
    }
    $a = Get-GctHeaderAnalysis -Text $text
    $a.Source = $source
    $Context.Headers = $a
    if (-not $a.Fields.Count) {
        $Context.Stop = $true
        Add-GctStep $Context $stage 'Message headers' Failed 'No header field found: the text must start with lines like "Received: ..." or "From: ...". In Outlook: open the message, File > Properties > Internet headers; in Outlook on the web: ... > View > View message details.' ([ordered]@{ Source = $source })
        return
    }
    $s = $a.Summary
    $to = @($s.To | ForEach-Object { $_.Address }) -join ', '
    $message = "$($a.Fields.Count) field(s), $($a.Hops.Count) Received hop(s)$(if ($s.Subject) { ": $($s.Subject)" })$(if ($s.From) { ", from $($s.From.Address)" })$(if ($to) { " to $to" })."
    Add-GctStep $Context $stage 'Message headers' Passed $message ([ordered]@{
            Source = $source; Subject = $s.Subject; From = $(if ($s.From) { $s.From.Text }); To = $to; Date = $s.Date; 'Message-ID' = $s.MessageId
            'Network Message ID' = $s.NetworkMessageId; 'Tenant (recipient)' = $s.TenantId; 'Mail client' = $s.Client; Direction = $s.Directionality
        })
    if ($a.Commands.Count) {
        $d = [ordered]@{}
        foreach ($c in $a.Commands) { $d[$c.Title] = $c.Text }
        Add-GctStep $Context $stage 'Follow the message' Passed "Commands to find the message in the logs ($(@($a.Commands | ForEach-Object { $_.Title }) -join '; ')): open the details to copy them." $d
    }
}

function Invoke-GctStageHeaderDns {
    <# The DNS of today for what the header shows: SPF for the sending IP, DMARC, the DKIM keys, the MX of the recipient. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'HeaderDns'
    $a = $Context.Headers
    $req = $a.CheckRequest
    if (-not $req) { Add-GctStep $Context $stage 'DNS of today' Skipped 'Nothing in the header to compare with the DNS (no domain of the sender, no DKIM signature).'; return }
    # No DNS server answering (offline, closed network, Test.DnsServer unreachable): the header is read offline.
    $dns = Test-GctDnsAnswer -Context $Context
    if (-not $dns.Answers) {
        $a.DnsOffline = $true
        Add-GctStep $Context $stage 'DNS of today' Skipped "No answer from the $($dns.Server) for $($dns.Probe): $($dns.Error) The header is read offline: every finding stays as the header shows it, without the comparison with the DNS of today. Set Test.DnsServer (-DnsServer) to a DNS server that answers, or Test.SkipDnsCrossCheck (-SkipDnsCrossCheck) not to try." ([ordered]@{ DnsServer = $dns.Server; Probe = $dns.Probe; Error = $dns.Error })
        return
    }
    $today = [ordered]@{ Spf = $null; Dmarc = $null; Dkim = [Collections.Generic.List[object]]::new(); Mx = $null }

    if ($req.MailFromDomain -and $req.Ip) {
        $spf = Test-GctSpf -Ip $req.Ip -Domain $req.MailFromDomain
        # The DNS of today qualifies the findings of the message: it is never a failure of the message itself.
        $status = if ($spf.Result -eq 'pass') { 'Passed' } else { 'Warning' }
        $text = switch ($spf.Result) {
            'pass' { "Today the SPF record of $($req.MailFromDomain) allows $($req.Ip) ($($spf.Mechanism))." }
            'none' { "$($req.MailFromDomain) has no SPF record today." }
            'permerror' { "The SPF record of $($req.MailFromDomain) is broken today: $($spf.Error)" }
            'temperror' { "DNS did not answer the SPF check today: $($spf.Error)" }
            default { "Today the SPF record of $($req.MailFromDomain) gives $($spf.Result) for $($req.Ip)$(if ($spf.Mechanism) { " ($($spf.Mechanism))" } else { ' (no mechanism matched)' })." }
        }
        if ($spf.Notes.Count) { $text += ' ' + ($spf.Notes -join ' ') }
        Add-GctStep $Context $stage 'SPF today' $status $text ([ordered]@{ Domain = $req.MailFromDomain; SendingIp = $req.Ip; Result = $spf.Result; Mechanism = $spf.Mechanism; Record = $spf.Record; DnsLookups = "$($spf.Lookups) of 10" })
        $today.Spf = [ordered]@{ Domain = $req.MailFromDomain; Ip = $req.Ip; Result = $spf.Result; Mechanism = $spf.Mechanism; Record = $spf.Record; Lookups = $spf.Lookups; Error = $spf.Error }
    }

    if ($req.FromDomain) {
        $org = if ($req.FromOrgDomain) { $req.FromOrgDomain } else { $req.FromDomain }
        $found = $null; $record = $null; $problem = $null
        foreach ($d in @(@($req.FromDomain, $org) | Select-Object -Unique)) {
            $q = Resolve-GctDns -Name "_dmarc.$d" -Type 'TXT'
            $records = @($q.Records | ForEach-Object { [string]$_.Data } | Where-Object { $_ -like 'v=DMARC1*' })
            if ($records.Count -gt 1) { $problem = "_dmarc.$d has $($records.Count) DMARC records today: receivers ignore them all."; $found = $d; break }
            if ($records.Count -eq 1) { $found = $d; $record = $records[0]; break }
            if ($q.Error) { $problem = "DNS did not answer for _dmarc.$d today."; break }
        }
        if ($problem) {
            Add-GctStep $Context $stage 'DMARC today' Warning $problem ([ordered]@{ Domain = $req.FromDomain })
            $today.Dmarc = [ordered]@{ Domain = $(if ($found) { $found } else { $req.FromDomain }); Found = $false; Error = $problem }
        }
        elseif (-not $record) {
            Add-GctStep $Context $stage 'DMARC today' Warning "No DMARC record today for $($req.FromDomain)$(if ($org -ne $req.FromDomain) { " or $org" })." ([ordered]@{ Domain = $req.FromDomain })
            $today.Dmarc = [ordered]@{ Domain = $req.FromDomain; Found = $false }
        }
        else {
            $tags = ConvertFrom-GctDnsTag $record
            $p = if ($tags['p']) { $tags['p'].ToLowerInvariant() } else { 'none' }
            $sp = if ($tags['sp']) { $tags['sp'].ToLowerInvariant() } else { $null }
            $applied = if ($found -ne $req.FromDomain -and $sp) { $sp } else { $p }
            $pct = if ($tags['pct']) { $tags['pct'] } else { '100' }
            $what = switch ($applied) { 'reject' { 'mail that fails is refused' } 'quarantine' { 'mail that fails goes to junk' } default { 'monitoring only, mail that fails is still delivered' } }
            $today.Dmarc = [ordered]@{ Domain = $found; Found = $true; Record = $record; P = $p; Sp = $sp; Pct = $pct; Adkim = $(if ($tags['adkim']) { $tags['adkim'].ToLowerInvariant() } else { 'r' }); Aspf = $(if ($tags['aspf']) { $tags['aspf'].ToLowerInvariant() } else { 'r' }); Rua = $tags['rua'] }
            $message = "DMARC of $found today: p=$p$(if ($sp) { ", sp=$sp" })$(if ($pct -ne '100') { ", pct=$pct" }) - $what$(if ($found -ne $req.FromDomain) { " (record of the organisational domain, for $($req.FromDomain))" })."
            Add-GctStep $Context $stage 'DMARC today' $(if ($applied -eq 'none') { 'Warning' } else { 'Passed' }) $message ([ordered]@{ Domain = $found; Record = $record; Policy = $applied; Alignment = "SPF $($today.Dmarc.Aspf), DKIM $($today.Dmarc.Adkim)" })
        }
    }

    foreach ($k in $req.Dkim) {
        $name = "$($k.Selector)._domainkey.$($k.Domain)"
        $q = Resolve-GctDns -Name $name -Type 'TXT'
        $record = @($q.Records | ForEach-Object { [string]$_.Data } | Where-Object { $_ -match 'p=' }) | Select-Object -First 1
        $keyType = $null; $bits = $null
        if ($q.Error) { $state = 'unknown' }
        elseif (-not $record) { $state = 'missing' }
        else {
            $tags = ConvertFrom-GctDnsTag $record
            $p = ([string]$tags['p']) -replace '\s', ''
            $keyType = if ($tags['k']) { $tags['k'].ToLowerInvariant() } else { 'rsa' }
            if (-not $p) { $state = 'revoked' } else { $state = 'present'; $bits = Get-GctDkimKeyBits -KeyType $keyType -Base64 $p }
        }
        $status = if ($state -eq 'present') { 'Passed' } else { 'Warning' }
        $message = switch ($state) {
            'present' { "The key $name is published today ($keyType$(if ($bits) { ", about $bits bits" }))." }
            'revoked' { "The key $name is revoked today (empty p=): signatures with this selector fail." }
            'missing' { "No key at $name today: signatures with this selector fail." }
            default { "DNS did not answer for $name today." }
        }
        Add-GctStep $Context $stage "DKIM key $($k.Selector) of $($k.Domain)" $status $message ([ordered]@{ Name = $name; State = $state })
        $today.Dkim.Add([ordered]@{ Domain = $k.Domain; Selector = $k.Selector; State = $state; KeyType = $keyType; KeyBits = $bits })
    }

    if ($req.RecipientDomain) {
        $q = Resolve-GctDns -Name $req.RecipientDomain -Type 'MX'
        $hosts = @($q.Records | ForEach-Object { [string]$_.Data } | Sort-Object { [int](($_ -split ' ')[0]) } | ForEach-Object { ($_.Split(' ', 2)[-1]).TrimEnd('.') })
        $eop = $hosts.Count -gt 0 -and -not @($hosts | Where-Object { $_ -notmatch '\.mail\.protection\.outlook\.com$|\.mx\.microsoft$' }).Count
        $message = if (-not $hosts.Count) { "No MX record today for $($req.RecipientDomain)." } elseif ($eop) { "The MX of $($req.RecipientDomain) points to Exchange Online Protection today ($($hosts[0]))." } else { "The MX of $($req.RecipientDomain) points to $($hosts[0]) today, not to Exchange Online Protection." }
        Add-GctStep $Context $stage 'MX today' $(if ($hosts.Count) { 'Passed' } else { 'Warning' }) $message ([ordered]@{ Records = $hosts -join ' | '; ExchangeOnlineProtection = $(if ($eop) { 'Yes' } else { 'No' }) })
        $today.Mx = [ordered]@{ Domain = $req.RecipientDomain; Records = $hosts; Eop = $eop }
    }
    $today.Dkim = @($today.Dkim)
    $a.DnsToday = $today
    Set-GctHeaderVerdict -Analysis $a -Today $today
}

function Get-GctHeaderVerdictOf {
    <# The verdict of one finding against the DNS of today: still, fixed, changed or unknown (or $null). #>
    param([System.Collections.IDictionary]$Finding, [System.Collections.IDictionary]$Today, [System.Collections.IDictionary]$Analysis)
    $c = $Finding.Check
    if (-not $c) { return $null }
    $v = { param([string]$State, [string]$Text) [ordered]@{ State = $State; Text = $Text } }
    switch ($c.Kind) {
        'spf' {
            $s = $Today['Spf']
            if (-not $s) { return $null }
            $now = $s['Result']; $was = $c.Result
            $how = if ($s['Mechanism']) { " ($($s['Mechanism']))" } else { '' }
            if ($now -eq 'temperror') { return & $v 'unknown' "DNS did not answer for $($c.Domain) today: try again later." }
            if ($was -eq 'pass') { if ($now -eq 'pass') { return & $v 'still' "Still allowed today$how." } else { return & $v 'changed' "Changed since: the SPF record of $($c.Domain) gives $now for $($c.Ip) today. Mail from this server would no longer pass." } }
            if ($now -eq 'pass') { return & $v 'fixed' "Fixed since: the SPF record of $($c.Domain) allows $($c.Ip) today$how. This tip is already handled." }
            if ($now -eq $was) { return & $v 'still' "Still the case: the SPF record of $($c.Domain) still gives $now for $($c.Ip)." }
            return & $v 'changed' "Changed since: $was in the message, $now today for $($c.Ip)$(if ($s['Record']) { " (record: $($s['Record']))" })."
        }
        'dkim' {
            $keys = @($Today['Dkim'] | Where-Object { $k = $_; @($c.Keys | Where-Object { $_.Domain -eq $k.Domain -and $_.Selector -eq $k.Selector }).Count })
            if (-not $keys.Count) { return $null }
            $present = @($keys | Where-Object { $_.State -eq 'present' })
            $list = @($keys | ForEach-Object { "$($_.Selector)._domainkey.$($_.Domain): $($_.State)" }) -join '; '
            if ($c.Result -eq 'pass') { if ($present.Count -eq $keys.Count) { return & $v 'still' "The keys are still published ($list)." } else { return & $v 'changed' "Changed since: $list. The key was removed or revoked (key rotation): new messages must use another selector." } }
            if ($c['ExchangeOnline']) { if ($present.Count) { return & $v 'changed' "Changed since: the Exchange Online selectors are published today ($list). Check that signing is enabled: Get-DkimSigningConfig -Identity $($c.Keys[0].Domain)." } else { return & $v 'still' "Still the case: no Exchange Online selector published ($list)." } }
            if ($c.Result -eq 'none') { if ($present.Count) { return & $v 'changed' "A key is published ($list)." } else { return & $v 'still' "Still the case: $list." } }
            if (-not $present.Count) { return & $v 'still' "Still the case: no usable key in DNS ($list). Publish the key of the selector." }
            return & $v 'changed' "The key is published today ($list). If the failure was a missing key, it is fixed; if the message was changed on the way (gateway, disclaimer, list), DNS does not fix it."
        }
        'dmarc' {
            $d = $Today['Dmarc']
            if (-not $d) { return $null }
            $policy = if ($d['Found']) { "p=$($d['P'])$(if ($d['Sp'] -and $d['Domain'] -ne $c.Domain) { " (sp=$($d['Sp']) for subdomains)" })$(if ($d['Pct'] -and $d['Pct'] -ne '100') { ", pct=$($d['Pct'])" })" } else { 'no record' }
            $spfFixed = $Today['Spf'] -and $Today['Spf'].Result -eq 'pass' -and $Analysis.Auth.Alignment.SpfAligned.Relaxed
            if ($c.Result -in 'bestguesspass', 'none') { if ($d['Found']) { return & $v 'fixed' "Fixed since: a DMARC record is published today for $($d['Domain']) ($policy)." } else { return & $v 'still' "Still the case: no DMARC record for $($c.Domain) or $($c.OrgDomain)." } }
            if ($c.Result -eq 'pass') {
                if (-not $d['Found']) { return & $v 'changed' "Changed since: no DMARC record today for $($c.Domain)." }
                return & $v 'still' "Still published: $policy.$(if ($d['P'] -eq 'none') { ' Monitoring only: mail that fails DMARC is still delivered.' })"
            }
            if ($c['Alignment'] -or $c.Result -eq 'fail' -or $c['Compauth']) {
                if ($spfFixed) { return & $v 'fixed' "Probably fixed since: SPF passes today for the aligned domain, and DMARC ($policy) would pass through SPF." }
                if (-not $d['Found']) { return & $v 'changed' "Changed since: no DMARC record today for $($c.Domain): DMARC no longer applies (and gives no protection)." }
                $changedPolicy = if ($c['Action'] -and $d['P'] -and $c['Action'] -ne $d['P'] -and -not ([string]$c['Action']).EndsWith($d['P'])) { " The policy changed: $($c['Action']) at the time, p=$($d['P']) today." } else { '' }
                return & $v 'still' "Still the case: DNS today does not solve the alignment ($policy, aspf=$(if ($d['Aspf'] -ne 's') { 'r' } else { 's' }), adkim=$(if ($d['Adkim'] -ne 's') { 'r' } else { 's' })).$changedPolicy"
            }
            return $null
        }
        'mx' {
            $m = $Today['Mx']
            if (-not $m) { return $null }
            if (-not @($m['Records']).Count) { return & $v 'unknown' "No MX record found today for $($c.Domain)." }
            if ($m['Eop']) {
                $label = if ($Finding.Status -eq 'Passed') { 'Changed since' } else { 'Fixed since' }
                return & $v $(if ($Finding.Status -eq 'Passed') { 'changed' } else { 'fixed' }) "$($label): the MX of $($c.Domain) points to Exchange Online Protection today ($(@($m['Records'])[0])): the gateway is no longer in front for incoming mail."
            }
            return & $v 'still' "Still the case: the MX of $($c.Domain) points to $(@($m['Records'])[0]) today, not to Exchange Online Protection."
        }
    }
    return $null
}

function Set-GctHeaderVerdict {
    <# Gives each finding its verdict from the DNS of today. #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Analysis, [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Today)
    foreach ($f in $Analysis.Findings) { $f.Verdict = Get-GctHeaderVerdictOf -Finding $f -Today $Today -Analysis $Analysis }
}

#endregion

#region The chronology ------------------------------------------------------------------------------

function Get-GctHeaderNode {
    <# The steps of the chronology: the sender, each Received hop, the mailbox - each with its facts and its findings. #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Analysis, [hashtable]$Codes = (Get-GctHeaderCodes))
    $a = $Analysis
    $nodes = [Collections.Generic.List[object]]::new()
    $o = $a.Origin
    $hostText = (@($o.Host, $(if ($o.Rdns -and $o.Rdns -ne $o.Host) { $o.Rdns })) | Where-Object { $_ }) -join ' / '
    $nodes.Add([ordered]@{
            Key = 'origin'; Time = $o.Time; DelayMs = $null; Skew = $false; Segment = 'origin'; Chip = 'Sender'; Biggest = $false; IsEntry = $false
            Title = $(if ($o.From -and $o.From.Address) { $o.From.Address } elseif ($o.Host) { $o.Host } else { 'The sender' })
            Sub = (@($o.Client, $(if ($o.Host) { "from $($o.Host)$(if ($o.Ip) { " ($($o.Ip))" })" })) | Where-Object { $_ }) -join ', '
            Facts = [ordered]@{ 'Written (Date field)' = $o.DateText; From = $(if ($o.From) { $o.From.Text }); 'Mail client' = $o.Client; 'Client IP (X-Originating-IP)' = $o.ClientIp; 'Sending host' = $hostText; 'Sending IP' = $o.Ip; 'First server' = $o.FirstServer }
            Raw = $null; Findings = Get-GctHeaderNodeFinding -Analysis $a -Node 'origin'
        })
    foreach ($h in $a.Hops) {
        $chip = if ($h.Product) { $h.Product } else { $script:HeaderSegments[[string]$h.Segment] }
        $nodes.Add([ordered]@{
                Key = "h$($h.Index)"; Time = $h.Time; DelayMs = $h.DelayMs; Skew = [bool]$h.Skew; Segment = $h.Segment; Chip = $chip
                Biggest = [bool]($a.Route.Biggest -and $a.Route.Biggest.Index -eq $h.Index); IsEntry = [bool]$h.IsEntry
                Title = $(if ($h.ByHost) { $h.ByHost } else { '(unknown server)' })
                Sub = "from $(if ((@($h.FromHost, $h.Ip) | Where-Object { $_ }) ) { (@($h.FromHost, $h.Ip) | Where-Object { $_ }) -join ' ' } else { '(not given)' })$(if ($h.Tls) { if ($h.Tls -eq 'yes') { ', TLS' } else { ", $($h.Tls)" } } elseif ($h.With) { ", $($h.With)" })"
                Facts = [ordered]@{
                    'Received from' = (@($h.FromHost, $(if ($h.Rdns -and $h.Rdns -ne $h.FromHost) { "reverse DNS $($h.Rdns)" })) | Where-Object { $_ }) -join ' / '
                    'Sending IP' = $h.Ip; Protocol = $h.With; TLS = $h.Tls; Server = $h.ExchangeVersion; Via = $h.Via; 'Time (UTC)' = Format-GctHeaderTime $h.Time -Full
                    'Waited before this hop' = $(if ($null -ne $h.DelayMs) { "$(Format-GctHeaderDuration $h.DelayMs), held by $($h.HeldBy)" })
                }
                Raw = $h; Findings = Get-GctHeaderNodeFinding -Analysis $a -Node $h.Index
            })
    }
    $d = $a.Delivered
    $folder = if ($d.Folder) { if ($Codes['mailboxDelivery']['dest'][$d.Folder]) { $Codes['mailboxDelivery']['dest'][$d.Folder] } else { $d.Folder } } else { $null }
    $nodes.Add([ordered]@{
            Key = 'delivery'; Time = $d.Time; DelayMs = $null; Skew = $false; Segment = 'delivery'; Chip = 'Mailbox'; Biggest = $false; IsEntry = $false
            Title = $(if ($folder) { "Delivered to: $folder" } else { 'Delivered' })
            Sub = (@(($d.Recipients -join ', '), $(if ($d.Mailbox) { "on $($d.Mailbox)" })) | Where-Object { $_ }) -join ' '
            Facts = [ordered]@{ Recipients = ($d.Recipients -join ', '); Folder = $folder; 'Spam confidence level' = $a.Antispam.Scl; 'Spam filtering verdict' = $(if ($a.Antispam.Report) { $a.Antispam.Report['SFV'] }); 'Bulk complaint level' = $a.Antispam.Bcl; 'Mailbox server' = $d.Mailbox }
            Raw = $null; Findings = Get-GctHeaderNodeFinding -Analysis $a -Node 'delivery'
        })
    return , @($nodes)
}

function Get-GctHeaderNodeStatus {
    param([object[]]$Findings)
    if (@($Findings | Where-Object { $_.Status -eq 'Failed' }).Count) { return 'Failed' }
    if (@($Findings | Where-Object { $_.Status -eq 'Warning' }).Count) { return 'Warning' }
    return 'Passed'
}

function Invoke-GctStageHeaderRoute {
    <# One check per step of the chronology, with the findings about that step in its details. #>
    param([Parameter(Mandatory = $true)][hashtable]$Context)

    $stage = 'HeaderRoute'
    $a = $Context.Headers
    $nodes = Get-GctHeaderNode -Analysis $a
    $rows = [Collections.Generic.List[object]]::new()
    $number = 0
    foreach ($n in $nodes) {
        $number++
        $status = Get-GctHeaderNodeStatus -Findings $n.Findings
        $wait = if ($n.Skew) { 'clock skew' } elseif ($null -ne $n.DelayMs) { '+' + (Format-GctHeaderDuration $n.DelayMs) } else { '' }
        $name = '{0}. {1} {2} - {3}' -f $number, (Format-GctHeaderTime $n.Time), $n.Chip, $n.Title
        $issues = @($n.Findings | Where-Object { $_.Status -in 'Failed', 'Warning' } | ForEach-Object { $_.Title + $(if ($_.Verdict -and $_.Verdict.State -eq 'fixed') { ' (fixed since)' } else { '' }) })
        $marks = @(if ($n.Biggest) { 'longest wait' }; if ($n.IsEntry) { 'enters Microsoft 365' }) -join ', '
        $message = (@($(if ($wait) { "$wait." }), "$($n.Sub)$(if ($n.Sub) { '.' })", $(if ($marks) { "$($marks.Substring(0,1).ToUpperInvariant())$($marks.Substring(1))." }), $(if ($issues.Count) { "To look at: $($issues -join '; ')." } elseif ($n.Findings.Count) { "$($n.Findings.Count) note(s), nothing to fix." })) | Where-Object { $_ -and $_ -ne '.' }) -join ' '
        if (-not $message) { $message = 'Nothing to note.' }
        $details = [ordered]@{}
        foreach ($k in $n.Facts.Keys) { if ($n.Facts[$k]) { $details[$k] = [string]$n.Facts[$k] } }
        foreach ($f in $n.Findings) {
            $value = $f.Text
            if ($f.Tip) { $value += " What to do: $($f.Tip)" }
            if ($f.Verdict) { $value += " DNS today: $($f.Verdict.Text)" }
            $details["[$($f.Status)] $($f.Title)"] = $value
        }
        # Diagnostics: the raw values the step is read from.
        if ($n.Raw) { $details['Received field'] = $n.Raw.Raw }
        $seen = @{}
        foreach ($f in $n.Findings) {
            foreach ($ev in $f.Evidence) {
                foreach ($field in @($a.Fields | Where-Object { $_.Name -ieq $ev })) {
                    $key = "Field $($field.Name)$(if ($seen.ContainsKey($field.Name)) { " ($($field.Index + 1))" })"
                    if ($seen.ContainsKey("$($field.Index)")) { continue }
                    $seen["$($field.Index)"] = $true; $seen[$field.Name] = $true
                    $details[$key] = $field.Value
                }
            }
        }
        Add-GctStep $Context $stage $name $status $message $details
        $rows.Add([pscustomobject]@{
                Step = $number; Kind = $(if ($n.Key -eq 'origin') { 'Sender' } elseif ($n.Key -eq 'delivery') { 'Mailbox' } else { 'Hop' }); TimeUtc = Format-GctHeaderTime $n.Time -Full
                WaitMs = $n.DelayMs; Wait = $wait; Segment = $n.Segment; SegmentName = $n.Chip; Server = $n.Title; From = $n.Sub
                Ip = $(if ($n.Raw) { $n.Raw.Ip } elseif ($n.Key -eq 'origin') { $a.Origin.Ip }); Tls = $(if ($n.Raw) { $n.Raw.Tls }); LongestWait = [bool]$n.Biggest; EntersMicrosoft365 = [bool]$n.IsEntry
                Status = $status; FindingsText = (@($n.Findings | ForEach-Object { "[$($_.Status)] $($_.Title)$(if ($_.Verdict) { " (DNS today: $($_.Verdict.State))" })" }) -join '; ')
                Findings = @($n.Findings | ForEach-Object { [pscustomobject]@{ Status = $_.Status; Area = $_.Area; Title = $_.Title; Text = $_.Text; Tip = $_.Tip; Verdict = $(if ($_.Verdict) { $_.Verdict.State }); VerdictText = $(if ($_.Verdict) { $_.Verdict.Text }) } })
            })
    }
    $Context.HeaderRoute = @($rows)
    if (-not $script:Ui) { Write-GctHeaderChronology -Rows @($rows) }
}

function Get-GctHeaderSummary {
    <# The message and its route in a few values (console, window, HTML report); $null when no header was read. #>
    param([AllowNull()][System.Collections.IDictionary]$Analysis)
    if (-not $Analysis -or -not $Analysis.Fields.Count) { return $null }
    $s = $Analysis.Summary
    $au = $Analysis.Auth
    $b = $Analysis.Route.Biggest
    [ordered]@{
        Source = $Analysis['Source']; Subject = $s.Subject; From = $(if ($s.From) { $s.From.Text }); To = @($s.To | ForEach-Object { $_.Address }); Date = $s.Date
        MessageId = $s.MessageId; NetworkMessageId = $s.NetworkMessageId; TenantId = $s.TenantId; Fields = $Analysis.Fields.Count; Hops = $Analysis.Hops.Count
        TransitMs = $Analysis.Route.TransitMs; Transit = Format-GctHeaderDuration $Analysis.Route.TransitMs
        LongestWait = $(if ($b) { "$(Format-GctHeaderDuration $b.DelayMs) at $($b.HeldBy)" } else { $null })
        Spf = $(if ($au.Spf) { $au.Spf.Result }); Dkim = $(if ($au.Dkim.Count) { $au.Dkim[0].Result } elseif ($au.Present) { 'none' }); Dmarc = $(if ($au.Dmarc) { $au.Dmarc.Result })
        Compauth = $(if ($au.Compauth) { "$($au.Compauth.Result)$(if ($au.Compauth.Reason) { " ($($au.Compauth.Reason))" })" })
        Scl = $Analysis.Antispam.Scl; Sfv = $(if ($Analysis.Antispam.Report) { $Analysis.Antispam.Report['SFV'] }); Folder = $Analysis.Delivered.Folder
        DnsCrossCheck = [bool]$Analysis.DnsToday; DnsOffline = [bool]$Analysis.DnsOffline
        Fixed = @($Analysis.Findings | Where-Object { $_.Verdict -and $_.Verdict.State -eq 'fixed' }).Count
        Commands = @($Analysis.Commands)
    }
}

function Write-GctHeaderChronology {
    <# The chronology in the console, one line per step: time, wait, who, server - the longest wait marked. Also in the log. #>
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows)
    $C = $script:C
    $color = @{ Passed = $C.Green; Warning = $C.Yellow; Failed = $C.Red }
    Write-GctLog 'INFO' 'Chronology of the message (UTC): time, wait, who held it, server'
    foreach ($r in $Rows) {
        $line = '{0,-8} {1,-12} {2,-30} {3}' -f ($r.TimeUtc -replace '^\S+ ', '' -replace ' UTC$', ''), $r.Wait, $r.SegmentName, $r.Server
        if ($r.LongestWait) { $line += '  <- longest wait' }
        Write-GctLog 'INFO' $line
        if (-not $script:Quiet) { Write-Host ('      {0}{1}{2}' -f $color[[string]$r.Status], $line, $C.Reset) }
    }
}

#endregion

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAmKZn288jUgCgo
# lIuWvmrmFMU1ysCrkrsusLuNHEtDnaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBlPJtsP4E9
# E81b215e7KuRRL9vdossWv/tNhK1uStVfTANBgkqhkiG9w0BAQEFAASCAgA+y3gm
# s4d94InOyqDxLeOOfj8mO/HTX9oV4MG75iSeauDY4Uy6lGIu4EtkKf4LFe3AwYdw
# 8pmalBvr/4LDbFjNZNf8IIyeG3bauK/RMAURIn7lW+62bp0xtHmyB7CL3zIHl+QC
# J219a0l3nykFm3waGOk5fpN95AsPpHuwggyJj0f+83b0aE81kWKX3UlX2jwfdPgX
# 16aEjMM8A1Go6X+RMw5osQLG2UYHON5Clp6bFe6eaiU15x+beOePqMRr7g/YtT1r
# kAtlIRCpBvqnXjz032royNlinng9KJa6rqXKjigvnwpW65bPVqvjUoky8aFfU7r3
# OGG/DC4EMgnAcVacGx0BqJnIU4/wHoj10H0VG8SjFHPjxVknq4uMHXLx3vSawMSC
# fMHyzV0pzwWiss0aKryOoHg3moxU3zbUImflp0ORjdeByMzoEWJuoh2r1uCl1+By
# zBRX553VAjvobjO4y+YbSI4D/k4onEYL8zy+Wukh0RKD0a4wnKvyICYRgJ3kxW6z
# pKwepwY1LQQrMEdHR9EJ0bsoCqr5uacTbU1diynS1UbS7xRliTWv8X4Z+Qk3jTVZ
# bTjGvlPqYIZg5qtt99SQL+fPZe2tF2gFliZhbad55v+mbsv7D3kKmVTNZfq4w6Ys
# OL95dvZKClhX9yhTRUkJe8uQAguBgwz2YCnoaKGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIB0Oer03Yv3kto9+qtgtSYNeGTKULM0lDn8HopUODPAIAghzbV4UKeqq
# kBgPMjAyNjExMTAxODI2NDBaMAMCAQECEHe9eN3QYDghXFwaGg98CLqgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDE4MjY0MFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIF69
# I6Rx+tsrTsqkhKKeXiuMvz5tz57wyLxoFV5tAzuzMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiAbRuJQx38K3EI78xMlXSfNkrp1mTmmYoHGk2Vau+3nygIhAPrm
# 1RQx9DN2H1MuCtAqXAQABUcv6j9QBXttxAaO8/7R
# SIG # End signature block
