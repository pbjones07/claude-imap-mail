# ImapMail calendar commands (CalDAV). Dot-sourced by ImapMail.psm1.
# Uses the same app password as mail. Every change saves a backup of the original
# event to $DataDir\calendar-backups so it can be put back with Restore-CalendarEvent.

$script:CalPresets = @{
    'iCloud'   = 'https://caldav.icloud.com/'
    'Fastmail' = 'https://caldav.fastmail.com/'
    'Yahoo'    = 'https://caldav.calendar.yahoo.com/'
    'Zoho'     = 'https://calendar.zoho.com/caldav/'
}
$script:TzMap = @{
    'America/Los_Angeles' = 'Pacific Standard Time'; 'America/Denver' = 'Mountain Standard Time'
    'America/Phoenix' = 'US Mountain Standard Time'; 'America/Chicago' = 'Central Standard Time'
    'America/New_York' = 'Eastern Standard Time'; 'America/Anchorage' = 'Alaskan Standard Time'
    'Pacific/Honolulu' = 'Hawaiian Standard Time'; 'Europe/London' = 'GMT Standard Time'
    'Europe/Paris' = 'Romance Standard Time'; 'Europe/Berlin' = 'W. Europe Standard Time'
    'Europe/Madrid' = 'Romance Standard Time'; 'Europe/Rome' = 'W. Europe Standard Time'
    'Asia/Tokyo' = 'Tokyo Standard Time'; 'Asia/Kolkata' = 'India Standard Time'
    'Australia/Sydney' = 'AUS Eastern Standard Time'; 'UTC' = 'UTC'; 'Etc/UTC' = 'UTC'
}

# ---------------------------------------------------------------- internals

function Get-IcsTimeZone([string]$id) {
    if (-not $id) { return [TimeZoneInfo]::Local }
    $id = $id.Trim('"')
    try { return [TimeZoneInfo]::FindSystemTimeZoneById($id) } catch { }
    if ($script:TzMap.ContainsKey($id)) { try { return [TimeZoneInfo]::FindSystemTimeZoneById($script:TzMap[$id]) } catch { } }
    [TimeZoneInfo]::Local
}

function Get-CalDavSettings {
    $cfg = Get-MailConfig; $cred = Get-MailCredential
    $url = if ($cfg.PSObject.Properties['CalDavUrl'] -and $cfg.CalDavUrl) { $cfg.CalDavUrl }
           elseif ($script:CalPresets.ContainsKey([string]$cfg.Provider)) { $script:CalPresets[[string]$cfg.Provider] }
           else { throw 'No calendar server is configured for this account. Run Initialize-ImapMail and enter a CalDAV URL.' }
    $user = if ($cfg.PSObject.Properties['CalDavUser'] -and $cfg.CalDavUser) { $cfg.CalDavUser } else { $cfg.Address }
    $pair = "${user}:$($cred.GetNetworkCredential().Password)"
    [pscustomobject]@{ Url = $url; Auth = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair)) }
}

function Invoke-Dav {
    param([string]$Url, [string]$Method, [string]$Body, [string]$Depth, [string]$Auth, [hashtable]$Headers,
          [string]$ContentType = 'application/xml; charset=utf-8')
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $req = [Net.HttpWebRequest]::Create($Url)
    $req.Method = $Method; $req.Headers['Authorization'] = $Auth
    $req.Timeout = 30000; $req.ReadWriteTimeout = 30000
    if ($Depth) { $req.Headers['Depth'] = $Depth }
    if ($Headers) { foreach ($k in $Headers.Keys) { $req.Headers[$k] = $Headers[$k] } }
    if ($Body) {
        $req.ContentType = $ContentType
        $b = [Text.Encoding]::UTF8.GetBytes($Body); $req.ContentLength = $b.Length
        $s = $req.GetRequestStream(); $s.Write($b, 0, $b.Length); $s.Close()
    }
    try { $r = $req.GetResponse() } catch [Net.WebException] { $r = $_.Exception.Response; if (-not $r) { throw } }
    try {
        $text = (New-Object IO.StreamReader($r.GetResponseStream(), [Text.Encoding]::UTF8)).ReadToEnd()
        [pscustomobject]@{ Code = [int]$r.StatusCode; Body = $text; ETag = $r.Headers['ETag'] }
    } finally { $r.Close() }
}

function Get-CalDavHome($s) {
    $a = Invoke-Dav -Url $s.Url -Method PROPFIND -Depth 0 -Auth $s.Auth -Body '<d:propfind xmlns:d="DAV:"><d:prop><d:current-user-principal/></d:prop></d:propfind>'
    if ($a.Code -eq 401) { throw 'Calendar sign-in failed (401). For iCloud, the calendar username must be your full Apple ID email address.' }
    $p = [regex]::Match($a.Body, '(?s)current-user-principal[^>]*>\s*<(?:\w+:)?href[^>]*>([^<]+)<').Groups[1].Value.Trim()
    if (-not $p) { throw "Could not find a calendar account at $($s.Url) (HTTP $($a.Code))." }
    $pu = [uri]::new([uri]$s.Url, $p).AbsoluteUri
    $h = Invoke-Dav -Url $pu -Method PROPFIND -Depth 0 -Auth $s.Auth -Body '<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:prop><c:calendar-home-set/></d:prop></d:propfind>'
    $hm = [regex]::Match($h.Body, '(?s)calendar-home-set[^>]*>\s*<(?:\w+:)?href[^>]*>([^<]+)<').Groups[1].Value.Trim()
    if (-not $hm) { throw 'Could not find the calendar home for this account.' }
    [uri]::new([uri]$pu, $hm).AbsoluteUri
}

function Get-CalendarCollections($s) {
    $calHome = Get-CalDavHome $s
    $l = Invoke-Dav -Url $calHome -Method PROPFIND -Depth 1 -Auth $s.Auth -Body '<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:prop><d:displayname/><d:resourcetype/><c:supported-calendar-component-set/></d:prop></d:propfind>'
    [xml]$x = $l.Body
    foreach ($r in $x.GetElementsByTagName('response', 'DAV:')) {
        if ($r.GetElementsByTagName('calendar', 'urn:ietf:params:xml:ns:caldav').Count -eq 0) { continue }
        $href = ($r.GetElementsByTagName('href', 'DAV:')[0]).InnerText.Trim()
        $name = ($r.GetElementsByTagName('displayname', 'DAV:') | Select-Object -First 1).InnerText
        $comps = @($r.GetElementsByTagName('comp', 'urn:ietf:params:xml:ns:caldav') | ForEach-Object { $_.GetAttribute('name') })
        [pscustomobject]@{ Name = $name; Components = ($comps -join ','); Url = [uri]::new([uri]$calHome, $href).AbsoluteUri }
    }
}

function Unescape-Ics([string]$v) { if ($null -eq $v) { return $null }; $v -replace '\\[nN]', "`n" -replace '\\,', ',' -replace '\\;', ';' -replace '\\\\', '\' }
function Escape-Ics([string]$v) { if ($null -eq $v) { return '' }; ($v -replace '\\', '\\' -replace ';', '\;' -replace ',', '\,') -replace "\r?\n", '\n' }

function Get-IcsProps([string]$block) {
    $props = @{}
    foreach ($line in (($block -replace "\r?\n[ \t]", '') -split "\r?\n")) {
        if ($line -match '^([A-Za-z-]+)((?:;(?:"[^"]*"|[^:"])*)?):(.*)$') {
            $k = $matches[1].ToUpper()
            if (-not $props.ContainsKey($k)) { $props[$k] = [pscustomobject]@{ Params = $matches[2]; Value = $matches[3] } }
        }
    }
    $props
}

function ConvertFrom-IcsDate($prop) {
    # Returns @{ Date = local DateTime; AllDay = bool }
    if (-not $prop) { return $null }
    $v = $prop.Value.Trim()
    if ($prop.Params -match 'VALUE=DATE(?!-)' -or $v -match '^\d{8}$') {
        return @{ Date = [datetime]::ParseExact($v.Substring(0, 8), 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture); AllDay = $true }
    }
    if ($v -match '^(\d{8}T\d{6})Z$') {
        $u = [datetime]::ParseExact($matches[1], "yyyyMMdd'T'HHmmss", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
        return @{ Date = $u.ToLocalTime(); AllDay = $false }
    }
    $d = [datetime]::ParseExact($v.Substring(0, 15), "yyyyMMdd'T'HHmmss", [Globalization.CultureInfo]::InvariantCulture)
    if ($prop.Params -match 'TZID=("[^"]*"|[^;:]+)') {
        $tz = Get-IcsTimeZone $matches[1]
        return @{ Date = [TimeZoneInfo]::ConvertTime([datetime]::SpecifyKind($d, 'Unspecified'), $tz, [TimeZoneInfo]::Local); AllDay = $false }
    }
    @{ Date = $d; AllDay = $false }
}

function Get-CalendarResource($s, [string]$Href) {
    $g = Invoke-Dav -Url $Href -Method GET -Auth $s.Auth
    if ($g.Code -ne 200) { throw "Could not read event at $Href (HTTP $($g.Code))." }
    $g
}

function Save-CalendarBackup([string]$Href, [string]$Ics, [string]$Uid) {
    $dir = Join-Path $script:DataDir 'calendar-backups'; New-Item -ItemType Directory -Force $dir | Out-Null
    $base = Join-Path $dir ("{0}_{1}" -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), (($Uid -replace '[^\w-]', '_') -replace '^(.{40}).*$', '$1'))
    [IO.File]::WriteAllText("$base.ics", $Ics, (New-Object Text.UTF8Encoding($false)))
    Set-Content "$base.url" $Href -Encoding UTF8
    "$base.ics"
}

function Format-IcsLine([string]$line) {
    if ($line.Length -le 74) { return $line }
    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add($line.Substring(0, 74)); $i = 74
    while ($i -lt $line.Length) { $n = [Math]::Min(73, $line.Length - $i); $parts.Add(' ' + $line.Substring($i, $n)); $i += $n }
    $parts -join "`r`n"
}

# ---------------------------------------------------------------- public

function Get-Calendar {
    <# Lists calendars (and reminder lists) on the account. #>
    [CmdletBinding()] param()
    Get-CalendarCollections (Get-CalDavSettings)
}

function Get-CalendarEvent {
    <# Lists events in a date range, with recurring events expanded into individual occurrences. #>
    [CmdletBinding()]
    param([datetime]$Start = (Get-Date).Date, [int]$Days = 7, [datetime]$End, [string]$Calendar = '*', [string]$Search)
    if (-not $PSBoundParameters.ContainsKey('End')) { $End = $Start.AddDays($Days) }
    $s = Get-CalDavSettings
    $fmt = "yyyyMMdd'T'HHmmss'Z'"
    $su = $Start.ToUniversalTime().ToString($fmt); $eu = $End.ToUniversalTime().ToString($fmt)
    $q = "<c:calendar-query xmlns:d=`"DAV:`" xmlns:c=`"urn:ietf:params:xml:ns:caldav`"><d:prop><d:getetag/><c:calendar-data><c:expand start=`"$su`" end=`"$eu`"/></c:calendar-data></d:prop><c:filter><c:comp-filter name=`"VCALENDAR`"><c:comp-filter name=`"VEVENT`"><c:time-range start=`"$su`" end=`"$eu`"/></c:comp-filter></c:comp-filter></c:filter></c:calendar-query>"
    $out = foreach ($cal in (Get-CalendarCollections $s | Where-Object { $_.Components -match 'VEVENT' -and $_.Name -like $Calendar })) {
        $rep = Invoke-Dav -Url $cal.Url -Method REPORT -Depth 1 -Auth $s.Auth -Body $q
        foreach ($resp in [regex]::Matches($rep.Body, '(?s)<(?:\w+:)?response[\s>].*?</(?:\w+:)?response>')) {
            $href = [regex]::Match($resp.Value, '<(?:\w+:)?href[^>]*>([^<]+)<').Groups[1].Value.Trim()
            $data = [regex]::Match($resp.Value, '(?s)<(?:\w+:)?calendar-data[^>]*>(.*?)</(?:\w+:)?calendar-data>').Groups[1].Value
            if (-not $data) { continue }
            $ics = [Net.WebUtility]::HtmlDecode($data)
            foreach ($ev in [regex]::Matches($ics, '(?s)BEGIN:VEVENT(.*?)END:VEVENT')) {
                $p = Get-IcsProps $ev.Groups[1].Value
                $st = ConvertFrom-IcsDate $p['DTSTART']; if (-not $st) { continue }
                $en = ConvertFrom-IcsDate $p['DTEND']
                $endDate = if ($en) { $en.Date } elseif ($st.AllDay) { $st.Date.AddDays(1) } else { $st.Date }
                if ($st.AllDay) { $endDate = $endDate.AddDays(-1) }   # show the last day, not the exclusive end
                $rid = ConvertFrom-IcsDate $p['RECURRENCE-ID']
                $title = Unescape-Ics ($p['SUMMARY'].Value)
                if ($Search -and $title -notlike "*$Search*") { continue }
                [pscustomobject][ordered]@{
                    Calendar = $cal.Name; Title = $title
                    Start = $st.Date; End = $endDate; AllDay = $st.AllDay
                    Location = Unescape-Ics ($p['LOCATION'].Value)
                    Notes = Unescape-Ics ($p['DESCRIPTION'].Value)
                    Recurring = [bool]$rid
                    RecurrenceId = if ($rid) { $rid.Date } else { $null }
                    Uid = $p['UID'].Value
                    Href = [uri]::new([uri]$cal.Url, $href).AbsoluteUri
                }
            }
        }
    }
    $out | Sort-Object Start
}

function New-CalendarEvent {
    <# Creates an event. Times are local; all-day events use -AllDay with -Start (and optional -End as the last day). #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][datetime]$Start,
        [datetime]$End, [int]$DurationMinutes = 60, [switch]$AllDay,
        [string]$Location, [string]$Notes, [string]$Calendar
    )
    $s = Get-CalDavSettings
    $cals = @(Get-CalendarCollections $s | Where-Object { $_.Components -match 'VEVENT' })
    $cal = if ($Calendar) { $cals | Where-Object { $_.Name -like $Calendar } | Select-Object -First 1 }
           else { ($cals | Where-Object { $_.Name -in 'Calendar', 'Home', 'Personal', 'Default' } | Select-Object -First 1), $cals[0] | Where-Object { $_ } | Select-Object -First 1 }
    if (-not $cal) { throw "Calendar '$Calendar' not found. Calendars: $(($cals.Name) -join ', ')" }
    $uid = [guid]::NewGuid().ToString().ToUpper()
    $fmt = "yyyyMMdd'T'HHmmss'Z'"
    $lines = @('BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//ImapMail//PowerShell//EN', 'CALSCALE:GREGORIAN', 'BEGIN:VEVENT',
               "UID:$uid", "DTSTAMP:$((Get-Date).ToUniversalTime().ToString($fmt))")
    if ($AllDay) {
        $last = if ($PSBoundParameters.ContainsKey('End')) { $End.Date } else { $Start.Date }
        $lines += "DTSTART;VALUE=DATE:$($Start.ToString('yyyyMMdd'))"
        $lines += "DTEND;VALUE=DATE:$($last.AddDays(1).ToString('yyyyMMdd'))"
    } else {
        if (-not $PSBoundParameters.ContainsKey('End')) { $End = $Start.AddMinutes($DurationMinutes) }
        if ($End -le $Start) { throw 'End must be after Start.' }
        $lines += "DTSTART:$($Start.ToUniversalTime().ToString($fmt))"
        $lines += "DTEND:$($End.ToUniversalTime().ToString($fmt))"
    }
    $lines += "SUMMARY:$(Escape-Ics $Title)"
    if ($Location) { $lines += "LOCATION:$(Escape-Ics $Location)" }
    if ($Notes) { $lines += "DESCRIPTION:$(Escape-Ics $Notes)" }
    $lines += 'END:VEVENT', 'END:VCALENDAR'
    $ics = (($lines | ForEach-Object { Format-IcsLine $_ }) -join "`r`n") + "`r`n"
    $url = $cal.Url.TrimEnd('/') + "/$uid.ics"
    $r = Invoke-Dav -Url $url -Method PUT -Auth $s.Auth -Body $ics -ContentType 'text/calendar; charset=utf-8' -Headers @{ 'If-None-Match' = '*' }
    if ($r.Code -notin 200, 201, 204) { throw "Could not create event (HTTP $($r.Code)): $($r.Body)" }
    [pscustomobject][ordered]@{ Created = $true; Calendar = $cal.Name; Title = $Title; Start = $Start; End = if ($AllDay) { $last } else { $End }; AllDay = [bool]$AllDay; Uid = $uid; Href = $url }
}

function Remove-CalendarEvent {
    <#
      Removes an event. Pipe in results from Get-CalendarEvent, or pass -Href (and -RecurrenceId for one occurrence).
      For a recurring event, removes only that occurrence unless -WholeSeries is given.
      The original is backed up first; undo with Restore-CalendarEvent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)][string]$Href,
        [Parameter(ValueFromPipelineByPropertyName)][Nullable[datetime]]$RecurrenceId,
        [Parameter(ValueFromPipelineByPropertyName)][string]$Title,
        [Parameter(ValueFromPipelineByPropertyName)][Nullable[datetime]]$Start,
        [switch]$WholeSeries
    )
    begin { $s = Get-CalDavSettings }
    process {
        $res = Get-CalendarResource $s $Href
        $ics = $res.Body
        $blocks = [regex]::Matches($ics, '(?s)BEGIN:VEVENT.*?END:VEVENT\r?\n?')
        $master = $blocks | Where-Object { $_.Value -notmatch '(?m)^RECURRENCE-ID' } | Select-Object -First 1
        $mp = if ($master) { Get-IcsProps $master.Value } else { Get-IcsProps $blocks[0].Value }
        $uid = $mp['UID'].Value
        $isSeries = $master -and $master.Value -match '(?m)^RRULE'
        $backup = Save-CalendarBackup $Href $ics $uid

        if (-not $isSeries -or $WholeSeries) {
            $r = Invoke-Dav -Url $Href -Method DELETE -Auth $s.Auth -Headers @{ 'If-Match' = $res.ETag }
            if ($r.Code -notin 200, 204) { throw "Delete failed (HTTP $($r.Code)): $($r.Body)" }
            [pscustomobject][ordered]@{ Removed = $(if ($isSeries) { 'whole series' } else { 'event' }); Title = Unescape-Ics $mp['SUMMARY'].Value; Backup = $backup }
            return
        }
        if (-not $RecurrenceId) { throw "'$(Unescape-Ics $mp['SUMMARY'].Value)' repeats. Pass the occurrence (pipe it from Get-CalendarEvent) or use -WholeSeries." }

        # Build an EXDATE in the same form as the series' DTSTART.
        $ds = $mp['DTSTART']; $occ = [datetime]$RecurrenceId
        if ($ds.Params -match 'VALUE=DATE(?!-)' -or $ds.Value -match '^\d{8}$') {
            # Expanded all-day occurrences carry a UTC RECURRENCE-ID; the local Start date is the real day.
            $day = if ($Start) { ([datetime]$Start).Date } else { $occ.ToUniversalTime().Date }
            $ex = "EXDATE;VALUE=DATE:$($day.ToString('yyyyMMdd'))"
            $occ = $day
        } elseif ($ds.Value -match 'Z$') {
            $ex = "EXDATE:$($occ.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'"))"
        } elseif ($ds.Params -match 'TZID=("[^"]*"|[^;:]+)') {
            $tzid = $matches[1]; $tz = Get-IcsTimeZone $tzid
            $wall = [TimeZoneInfo]::ConvertTime([datetime]::SpecifyKind($occ, 'Local'), $tz)
            $ex = "EXDATE;TZID=${tzid}:$($wall.ToString("yyyyMMdd'T'HHmmss"))"
        } else {
            $ex = "EXDATE:$($occ.ToString("yyyyMMdd'T'HHmmss"))"
        }
        $newMaster = [regex]::new('(?m)^(DTSTART[^\r\n]*\r?\n)').Replace($master.Value, "`$1$ex`r`n", 1)
        $newIcs = $ics.Replace($master.Value, $newMaster)
        # Drop any edited copy of that same occurrence.
        foreach ($b in $blocks) {
            if ($b.Value -notmatch '(?m)^RECURRENCE-ID') { continue }
            $rid = ConvertFrom-IcsDate ((Get-IcsProps $b.Value)['RECURRENCE-ID'])
            if ($rid -and [Math]::Abs(($rid.Date - $occ).TotalMinutes) -lt 1) { $newIcs = $newIcs.Replace($b.Value, '') }
        }
        $r = Invoke-Dav -Url $Href -Method PUT -Auth $s.Auth -Body $newIcs -ContentType 'text/calendar; charset=utf-8' -Headers @{ 'If-Match' = $res.ETag }
        if ($r.Code -notin 200, 201, 204) { throw "Update failed (HTTP $($r.Code)): $($r.Body)" }
        [pscustomobject][ordered]@{ Removed = 'one occurrence'; Title = Unescape-Ics $mp['SUMMARY'].Value; Occurrence = $occ; Backup = $backup }
    }
}

function Restore-CalendarEvent {
    <# Puts an event back exactly as it was before a change, from a backup made by Remove-CalendarEvent. #>
    [CmdletBinding()] param([Parameter(Mandatory)][string]$Path)
    $s = Get-CalDavSettings
    $url = (Get-Content ([IO.Path]::ChangeExtension($Path, '.url')) -Raw).Trim()
    $r = Invoke-Dav -Url $url -Method PUT -Auth $s.Auth -Body ([IO.File]::ReadAllText($Path)) -ContentType 'text/calendar; charset=utf-8'
    if ($r.Code -notin 200, 201, 204) { throw "Restore failed (HTTP $($r.Code)): $($r.Body)" }
    "Restored event from $Path"
}

function Get-CalendarBackup {
    <# Lists calendar backups, newest first. #>
    $dir = Join-Path $script:DataDir 'calendar-backups'
    if (Test-Path $dir) { Get-ChildItem $dir -Filter *.ics | Sort-Object LastWriteTime -Descending | Select-Object Name, LastWriteTime, FullName }
}
