# ImapMail - read, search, organize and draft email over IMAP/SMTP from PowerShell.
# Windows PowerShell 5.1+. MIT License.
#
# Your settings live in $env:IMAPMAIL_HOME (default: %APPDATA%\ImapMail):
#   config.json     - account and server settings
#   credential.xml  - your app password, encrypted with Windows DPAPI (only your
#                     Windows account on this PC can decrypt it)
#   logs\           - CSV logs of every bulk move to Trash, so they can be undone
#   bin\            - compiled helper library (built automatically from ImapMail.cs)

$script:ModuleDir = $PSScriptRoot
$script:DataDir = if ($env:IMAPMAIL_HOME) { $env:IMAPMAIL_HOME } else { Join-Path $env:APPDATA 'ImapMail' }

function Initialize-ImapMailLibrary {
    if ('ImapMail.ImapClient' -as [type]) { return }
    $cs = Join-Path $script:ModuleDir 'ImapMail.cs'
    $sha = [Security.Cryptography.SHA256]::Create()
    $hash = -join ($sha.ComputeHash([IO.File]::ReadAllBytes($cs))[0..5] | ForEach-Object { $_.ToString('x2') })
    $bin = Join-Path $script:DataDir 'bin'
    $dll = Join-Path $bin "ImapMail-$hash.dll"
    if (-not (Test-Path $dll)) {
        New-Item -ItemType Directory -Force $bin | Out-Null
        Add-Type -TypeDefinition ([IO.File]::ReadAllText($cs)) -Language CSharp -OutputAssembly $dll -OutputType Library -ReferencedAssemblies System.dll
    }
    Add-Type -Path $dll
}
Initialize-ImapMailLibrary

$script:Presets = [ordered]@{
    'iCloud'   = @{ ImapHost = 'imap.mail.me.com';    ImapPort = 993; SmtpHost = 'smtp.mail.me.com';    SmtpPort = 587; SmtpTls = 'StartTls'; Help = 'Create an app-specific password at account.apple.com > Sign-In and Security > App-Specific Passwords. Use your @icloud.com / @me.com address.' }
    'Gmail'    = @{ ImapHost = 'imap.gmail.com';      ImapPort = 993; SmtpHost = 'smtp.gmail.com';      SmtpPort = 465; SmtpTls = 'Implicit'; Help = 'Requires 2-Step Verification. Create an app password at myaccount.google.com/apppasswords.' }
    'Fastmail' = @{ ImapHost = 'imap.fastmail.com';   ImapPort = 993; SmtpHost = 'smtp.fastmail.com';   SmtpPort = 465; SmtpTls = 'Implicit'; Help = 'Create an app password in Settings > Privacy & Security > Manage app passwords.' }
    'Yahoo'    = @{ ImapHost = 'imap.mail.yahoo.com'; ImapPort = 993; SmtpHost = 'smtp.mail.yahoo.com'; SmtpPort = 465; SmtpTls = 'Implicit'; Help = 'Generate an app password in Account Info > Account Security.' }
    'AOL'      = @{ ImapHost = 'imap.aol.com';        ImapPort = 993; SmtpHost = 'smtp.aol.com';        SmtpPort = 465; SmtpTls = 'Implicit'; Help = 'Generate an app password in Account Security.' }
    'Zoho'     = @{ ImapHost = 'imap.zoho.com';       ImapPort = 993; SmtpHost = 'smtp.zoho.com';       SmtpPort = 465; SmtpTls = 'Implicit'; Help = 'Enable IMAP access and create an app-specific password in your Zoho account security settings.' }
}

# ---------------------------------------------------------------- setup

function Initialize-ImapMail {
    <# Interactive setup. Run in a PowerShell window; prompts for provider, address and app password. #>
    [CmdletBinding()] param()
    Write-Host ''
    Write-Host '  ImapMail setup' -ForegroundColor Cyan
    Write-Host "  Settings are stored in $script:DataDir. Your password is encrypted with your"
    Write-Host '  Windows account (DPAPI); only you, signed in to this PC, can decrypt it.'
    Write-Host ''
    $names = @($script:Presets.Keys) + 'Other (enter servers manually)'
    for ($i = 0; $i -lt $names.Count; $i++) { Write-Host ("  {0}. {1}" -f ($i + 1), $names[$i]) }
    do { $n = Read-Host '  Choose your email provider (number)' } until ($n -match '^\d+$' -and [int]$n -ge 1 -and [int]$n -le $names.Count)
    $choice = $names[[int]$n - 1]
    if ($script:Presets.Contains($choice)) {
        $s = $script:Presets[$choice]; $provider = $choice
        Write-Host "  $($s.Help)" -ForegroundColor Yellow
    } else {
        $provider = 'Custom'
        $s = @{
            ImapHost = Read-Host '  IMAP server (e.g. imap.example.com)'
            ImapPort = [int](Read-Host '  IMAP port (usually 993)')
            SmtpHost = Read-Host '  SMTP server (e.g. smtp.example.com)'
            SmtpPort = [int](Read-Host '  SMTP port (465 or 587)')
        }
        $s.SmtpTls = if ($s.SmtpPort -eq 465) { 'Implicit' } else { 'StartTls' }
    }
    $addr = (Read-Host '  Your email address').Trim()
    $user = (Read-Host "  Sign-in username [press Enter for $addr]").Trim(); if (-not $user) { $user = $addr }
    $name = (Read-Host '  Name to show on emails you send').Trim()
    $pw = Read-Host '  App password' -AsSecureString
    # Calendar (CalDAV) uses the same app password.
    $calUrl = $null; $calUser = $null
    $calPresets = @{ 'iCloud' = 'https://caldav.icloud.com/'; 'Fastmail' = 'https://caldav.fastmail.com/'; 'Yahoo' = 'https://caldav.calendar.yahoo.com/'; 'Zoho' = 'https://calendar.zoho.com/caldav/' }
    if ($calPresets.ContainsKey($provider)) {
        $calUrl = $calPresets[$provider]
        $calUser = if ($provider -eq 'iCloud') { (Read-Host "  Apple ID email for Calendar [press Enter for $addr]").Trim() } else { '' }
        if (-not $calUser) { $calUser = $addr }
    } elseif ($provider -eq 'Custom') {
        $calUrl = (Read-Host '  CalDAV calendar URL (optional, press Enter to skip)').Trim()
        if ($calUrl) { $calUser = $user }
    }
    $cred = New-Object System.Management.Automation.PSCredential($user, $pw)

    Write-Host ''
    Write-Host "  Testing sign-in to $($s.ImapHost)..." -NoNewline
    try {
        $imap = New-Object ImapMail.ImapClient($s.ImapHost, [int]$s.ImapPort)
        $imap.Login($cred.UserName, $cred.GetNetworkCredential().Password)
        $imap.Dispose()
        Write-Host ' OK' -ForegroundColor Green
    } catch {
        Write-Host ' FAILED' -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        Write-Host '  Check your address and app password, then run setup again.'
        return
    }
    New-Item -ItemType Directory -Force $script:DataDir | Out-Null
    $cred | Export-Clixml (Join-Path $script:DataDir 'credential.xml')
    [pscustomobject][ordered]@{
        Provider = $provider; Address = $addr; DisplayName = $name
        ImapHost = $s.ImapHost; ImapPort = [int]$s.ImapPort
        SmtpHost = $s.SmtpHost; SmtpPort = [int]$s.SmtpPort; SmtpTls = $s.SmtpTls
        CalDavUrl = $calUrl; CalDavUser = $calUser
    } | ConvertTo-Json | Set-Content (Join-Path $script:DataDir 'config.json') -Encoding UTF8
    if ($calUrl) {
        Write-Host '  Testing calendar sign-in...' -NoNewline
        try { [void](Get-Calendar); Write-Host ' OK' -ForegroundColor Green }
        catch { Write-Host ' failed' -ForegroundColor Yellow; Write-Host "  Mail works, but calendar access didn't: $($_.Exception.Message)" -ForegroundColor Yellow }
    }
    Write-Host '  All set. You can close this window.' -ForegroundColor Green
}

function Get-MailConfig {
    $p = Join-Path $script:DataDir 'config.json'
    if (-not (Test-Path $p)) { throw "ImapMail is not set up yet. Run Initialize-ImapMail in a PowerShell window." }
    Get-Content $p -Raw | ConvertFrom-Json
}

function Get-MailCredential {
    $p = Join-Path $script:DataDir 'credential.xml'
    if (-not (Test-Path $p)) { throw "ImapMail is not set up yet. Run Initialize-ImapMail in a PowerShell window." }
    Import-Clixml $p
}

# ---------------------------------------------------------------- internals

function Open-MailImap {
    $cfg = Get-MailConfig; $c = Get-MailCredential
    $imap = New-Object ImapMail.ImapClient($cfg.ImapHost, [int]$cfg.ImapPort)
    $imap.Login($c.UserName, $c.GetNetworkCredential().Password)
    $imap
}

function Use-MailImap([scriptblock]$Script) {
    $imap = Open-MailImap
    try { & $Script $imap } finally { $imap.Dispose() }
}

function Format-ImapDate([datetime]$d) { $d.ToString('dd-MMM-yyyy', [Globalization.CultureInfo]::InvariantCulture) }
function QS([string]$s) { [ImapMail.ImapClient]::Q($s) }

function Get-AddressList([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return @() }
    $t = [regex]::Replace($s, '"[^"]*"', '')
    $a = @([regex]::Matches($t, '<([^<>\s]+@[^<>\s]+)>') | ForEach-Object { $_.Groups[1].Value })
    $t2 = [regex]::Replace($t, '<[^>]*>', '')
    $a += @([regex]::Matches($t2, "[\w.+'-]+@[\w-]+(\.[\w-]+)+") | ForEach-Object { $_.Value })
    @($a | Select-Object -Unique)
}

function ConvertTo-MailObject($p, $Uid, $Folder, [int]$MaxChars) {
    $body = $p.BestText()
    if ($MaxChars -gt 0 -and $body.Length -gt $MaxChars) { $body = $body.Substring(0, $MaxChars) + "`n...[truncated; $($body.Length) chars total]" }
    [pscustomobject][ordered]@{
        Uid = $Uid; Folder = $Folder
        Date = $p.Header('Date'); From = $p.Header('From'); To = $p.Header('To'); Cc = $p.Header('Cc')
        ReplyTo = $p.Header('Reply-To'); Subject = $p.Header('Subject'); MessageId = $p.Header('Message-ID')
        Attachments = (@($p.Attachments | ForEach-Object { "$($_.Name) ($([math]::Ceiling($_.Size / 1KB)) KB)" }) -join '; ')
        Body = $body
    }
}

function Split-Batch([long[]]$Items, [int]$Size = 200) {
    for ($i = 0; $i -lt $Items.Count; $i += $Size) { , ($Items[$i..([Math]::Min($i + $Size, $Items.Count) - 1)]) }
}

# ---------------------------------------------------------------- reading

function Get-MailFolder {
    <# Lists mailbox folders. -Counts adds total and unread counts. #>
    [CmdletBinding()] param([switch]$Counts)
    Use-MailImap {
        param($imap)
        foreach ($f in $imap.ListFolders()) {
            if ($Counts -and $f.Attributes -notmatch 'Noselect') { $imap.FillStatus($f) }
            $o = [ordered]@{ Name = $f.Name; Attributes = $f.Attributes }
            if ($Counts) { $o.Messages = $f.Messages; $o.Unread = $f.Unread }
            [pscustomobject]$o
        }
    }
}

function Find-Mail {
    <# Searches a folder and returns message summaries, newest first. #>
    [CmdletBinding()]
    param(
        [string]$Folder = 'INBOX',
        [datetime]$Since, [datetime]$Before,
        [string]$From, [string]$To, [string]$Subject, [string]$Text,
        [switch]$Unread, [switch]$Flagged,
        [string]$RawCriteria,
        [int]$Max = 50
    )
    $c = @()
    if ($PSBoundParameters.ContainsKey('Since'))  { $c += "SINCE $(Format-ImapDate $Since)" }
    if ($PSBoundParameters.ContainsKey('Before')) { $c += "BEFORE $(Format-ImapDate $Before)" }
    if ($From)    { $c += "FROM $(QS $From)" }
    if ($To)      { $c += "TO $(QS $To)" }
    if ($Subject) { $c += "SUBJECT $(QS $Subject)" }
    if ($Text)    { $c += "TEXT $(QS $Text)" }
    if ($Unread)  { $c += 'UNSEEN' }
    if ($Flagged) { $c += 'FLAGGED' }
    if ($RawCriteria) { $c += $RawCriteria }
    if ($c.Count -eq 0) { $c = @('ALL') }
    $criteria = $c -join ' '
    if ($criteria -match '[^\x00-\x7F]') { $criteria = "CHARSET UTF-8 $criteria" }
    Use-MailImap {
        param($imap)
        $imap.Select($Folder, $true)
        $u = $imap.SearchUids($criteria)
        if ($u.Count -gt $Max) { $u = $u.GetRange($u.Count - $Max, $Max) }
        $imap.FetchSummaries($u) | Select-Object Uid, Folder, Date, From, To, Subject, Unread, Flagged, Answered, SizeKB
    }
}

function Read-Mail {
    <# Returns full messages (headers, plain-text body, attachment list). Does not mark them read. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][long[]]$Uid, [string]$Folder = 'INBOX', [int]$MaxChars = 15000)
    Use-MailImap {
        param($imap)
        $imap.Select($Folder, $true)
        foreach ($id in $Uid) {
            $raw = $imap.FetchRaw($id)
            if (-not $raw) { Write-Warning "UID $id not found in $Folder"; continue }
            ConvertTo-MailObject ([ImapMail.Mime]::Parse($raw)) $id $Folder $MaxChars
        }
    }
}

function Save-MailAttachment {
    [CmdletBinding()]
    param([Parameter(Mandatory)][long]$Uid, [string]$Folder = 'INBOX', [string]$Destination = (Join-Path $env:USERPROFILE 'Downloads'), [string]$Name)
    Use-MailImap {
        param($imap)
        $imap.Select($Folder, $true)
        $raw = $imap.FetchRaw($Uid)
        if (-not $raw) { throw "UID $Uid not found in $Folder" }
        foreach ($a in ([ImapMail.Mime]::Parse($raw)).Attachments) {
            if ($Name -and $a.Name -notlike $Name) { continue }
            $fn = ($a.Name -replace '[\\/:*?"<>|]', '_')
            $path = Join-Path $Destination $fn
            $i = 1; while (Test-Path $path) { $path = Join-Path $Destination ("{0} ({1}){2}" -f [IO.Path]::GetFileNameWithoutExtension($fn), $i++, [IO.Path]::GetExtension($fn)) }
            [IO.File]::WriteAllBytes($path, $a.Data)
            Get-Item $path
        }
    }
}

# ---------------------------------------------------------------- organizing

function Set-MailFlag {
    [CmdletBinding()]
    param([Parameter(Mandatory)][long[]]$Uid, [string]$Folder = 'INBOX', [switch]$Read, [switch]$Unread, [switch]$Flag, [switch]$Unflag)
    Use-MailImap {
        param($imap)
        $imap.Select($Folder, $false)
        foreach ($b in (Split-Batch $Uid)) {
            if ($Read)   { $imap.Store($b, '+FLAGS.SILENT (\Seen)') }
            if ($Unread) { $imap.Store($b, '-FLAGS.SILENT (\Seen)') }
            if ($Flag)   { $imap.Store($b, '+FLAGS.SILENT (\Flagged)') }
            if ($Unflag) { $imap.Store($b, '-FLAGS.SILENT (\Flagged)') }
        }
        "Updated $($Uid.Count) message(s) in $Folder"
    }
}

function Move-Mail {
    [CmdletBinding()]
    param([Parameter(Mandatory)][long[]]$Uid, [string]$Folder = 'INBOX', [Parameter(Mandatory)][string]$Destination)
    Use-MailImap {
        param($imap)
        $imap.Select($Folder, $false)
        foreach ($b in (Split-Batch $Uid)) { $imap.Move($b, $Destination) }
        "Moved $($Uid.Count) message(s) from $Folder to $Destination"
    }
}

function Remove-Mail {
    <# Moves messages to the provider's Trash folder (recoverable) and logs them to a CSV first. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][long[]]$Uid, [string]$Folder = 'INBOX', [string]$Label = 'removed')
    Use-MailImap {
        param($imap)
        $trash = $imap.FindSpecialFolder('\Trash', 'Trash')
        $imap.Select($Folder, $false)
        $logDir = Join-Path $script:DataDir 'logs'; New-Item -ItemType Directory -Force $logDir | Out-Null
        $log = Join-Path $logDir ("{0}_{1}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'), ($Label -replace '[^\w-]', '_'))
        $imap.FetchSummaries([Collections.Generic.List[long]]$Uid) | Select-Object Uid, Folder, Date, From, Subject, MessageId |
            Export-Csv $log -NoTypeInformation -Encoding UTF8
        foreach ($b in (Split-Batch $Uid)) { $imap.Move($b, $trash) }
        [pscustomobject]@{ Moved = $Uid.Count; From = $Folder; To = $trash; Log = $log }
    }
}

# ---------------------------------------------------------------- writing

function New-MailDraft {
    <# Saves a message to Drafts. With -ReplyToUid it threads the reply and quotes the original. Nothing is sent. #>
    [CmdletBinding()]
    param(
        [string]$To, [string]$Cc, [string]$Bcc, [string]$Subject,
        [Parameter(Mandatory)][string]$Body,
        [long]$ReplyToUid, [string]$ReplyFolder = 'INBOX', [switch]$ReplyAll, [switch]$NoQuote
    )
    $cfg = Get-MailConfig
    $from = if ($cfg.DisplayName) { "$($cfg.DisplayName) <$($cfg.Address)>" } else { $cfg.Address }
    Use-MailImap {
        param($imap)
        $inReplyTo = $null; $refs = $null
        if ($ReplyToUid) {
            $imap.Select($ReplyFolder, $true)
            $raw = $imap.FetchRaw($ReplyToUid)
            if (-not $raw) { throw "UID $ReplyToUid not found in $ReplyFolder" }
            $o = [ImapMail.Mime]::Parse($raw)
            if (-not $To) { $To = if ($o.Header('Reply-To')) { $o.Header('Reply-To') } else { $o.Header('From') } }
            if ($ReplyAll -and -not $Cc) {
                $me = @($cfg.Address.ToLower())
                $toAddrs = @(Get-AddressList $To | ForEach-Object { $_.ToLower() })
                $others = @((Get-AddressList $o.Header('To')) + (Get-AddressList $o.Header('Cc'))) |
                    Where-Object { $me -notcontains $_.ToLower() -and $toAddrs -notcontains $_.ToLower() } | Select-Object -Unique
                if ($others) { $Cc = $others -join ', ' }
            }
            if (-not $Subject) {
                $s = $o.Header('Subject'); if (-not $s) { $s = '' }
                $Subject = if ($s -match '^\s*re:') { $s } else { "Re: $s" }
            }
            $inReplyTo = $o.Header('Message-ID')
            $refs = (@($o.Header('References'), $inReplyTo) | Where-Object { $_ }) -join ' '
            if (-not $NoQuote) {
                $q = ($o.BestText() -split "`n" | ForEach-Object { "> $_" }) -join "`n"
                $Body = "$Body`n`nOn $($o.Header('Date')), $($o.Header('From')) wrote:`n$q"
            }
        }
        if (-not $To) { throw 'A recipient (-To) is required.' }
        $mid = $null
        $bytes = [ImapMail.Mime]::BuildMessage($from, $To, $Cc, $Bcc, $Subject, $Body, $inReplyTo, $refs, [ref]$mid)
        $drafts = $imap.FindSpecialFolder('\Drafts', 'Drafts')
        $newUid = $imap.Append($drafts, '\Draft \Seen', $bytes)
        [pscustomobject][ordered]@{ DraftUid = $newUid; DraftsFolder = $drafts; To = $To; Cc = $Cc; Subject = $Subject; MessageId = $mid }
    }
}

function Send-MailDraft {
    <# Sends a saved draft over SMTP, files a copy in Sent, and deletes the draft. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][long]$Uid, [switch]$KeepDraft, [switch]$NoSentCopy)
    $cfg = Get-MailConfig; $cred = Get-MailCredential
    Use-MailImap {
        param($imap)
        $drafts = $imap.FindSpecialFolder('\Drafts', 'Drafts')
        $sent = $imap.FindSpecialFolder('\Sent', 'Sent')
        $imap.Select($drafts, $false)
        $raw = $imap.FetchRaw($Uid)
        if (-not $raw) { throw "Draft UID $Uid not found in $drafts" }
        $p = [ImapMail.Mime]::Parse($raw)
        $rcpts = @((Get-AddressList $p.Header('To')) + (Get-AddressList $p.Header('Cc')) + (Get-AddressList $p.Header('Bcc'))) | Select-Object -Unique
        if (-not $rcpts) { throw 'Draft has no recipients.' }
        $wire = [ImapMail.Mime]::RemoveHeader($raw, 'Bcc')
        [ImapMail.SmtpSender]::Send($cfg.SmtpHost, [int]$cfg.SmtpPort, ($cfg.SmtpTls -eq 'Implicit'), $cred.UserName, $cred.GetNetworkCredential().Password, $cfg.Address, [string[]]$rcpts, $wire)
        # Gmail files sent mail automatically; other providers need the copy appended.
        if (-not $NoSentCopy -and $cfg.Provider -ne 'Gmail') {
            try { [void]$imap.Append($sent, '\Seen', $raw) } catch { Write-Warning "Sent, but could not save a copy to ${sent}: $_" }
        }
        if (-not $KeepDraft) { $imap.Expunge([long[]]@($Uid)) }
        [pscustomobject][ordered]@{ Sent = $true; To = $p.Header('To'); Cc = $p.Header('Cc'); Subject = $p.Header('Subject'); Recipients = ($rcpts -join ', ') }
    }
}

function Remove-MailDraft {
    [CmdletBinding()] param([Parameter(Mandatory)][long]$Uid)
    Use-MailImap {
        param($imap)
        $drafts = $imap.FindSpecialFolder('\Drafts', 'Drafts')
        $imap.Select($drafts, $false)
        $imap.Expunge([long[]]@($Uid))
        "Deleted draft $Uid"
    }
}

. (Join-Path $PSScriptRoot 'ImapMail.Calendar.ps1')

Export-ModuleMember -Function Initialize-ImapMail, Get-MailConfig, Get-MailFolder, Find-Mail, Read-Mail, Save-MailAttachment,
    Set-MailFlag, Move-Mail, Remove-Mail, New-MailDraft, Send-MailDraft, Remove-MailDraft,
    Get-Calendar, Get-CalendarEvent, New-CalendarEvent, Remove-CalendarEvent, Restore-CalendarEvent, Get-CalendarBackup
