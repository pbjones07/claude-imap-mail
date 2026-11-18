# ImapMail for Claude

Let Claude read, summarize, clean up and draft replies in **any IMAP email account** (iCloud, Gmail, Fastmail, Yahoo, AOL, Zoho or a custom server), with no third-party service in between. Claude drives a small PowerShell module on your Windows PC that talks directly to your mail server.

**What it can do**
- Summarize today's mail, or search by sender, subject, date or text
- Read full messages and save attachments
- Mark messages read or flagged, move them between folders, and bulk-clean promotions and notifications (to Trash, with an undo log)
- Write replies as **drafts** in your Drafts folder, and send them only after you approve

**Safety by design**
- Your app password is encrypted with Windows DPAPI under your Windows account. Claude never sees it, and it never leaves your PC.
- Claude saves drafts. It never sends anything unless you approve that specific message.
- Cleanup moves mail to Trash (recoverable) and logs every move.

## Requirements
- Windows 10 or 11 with Windows PowerShell 5.1 (built in)
- An email account with IMAP enabled and an **app password**
- Claude with access to a PowerShell shell on your PC (Claude Code on Windows, or a Claude app linked to your computer)

## Install

### 1. Install the Claude plugin
In Claude Code:
```
/plugin marketplace add pbjones07/claude-imap-mail
/plugin install imap-mail@imap-mail
```
Or copy `plugins/imap-mail/skills/imap-mail/SKILL.md` into your own skills.

### 2. Install the PowerShell module and connect your account
Download or clone this repo, then run:
```powershell
powershell -ExecutionPolicy Bypass -File .\plugins\imap-mail\scripts\Install.ps1
```
Choose your provider, enter your address, and paste your app password into the prompt. Setup tests the sign-in before saving.

| Provider | Where to get an app password |
|---|---|
| iCloud | account.apple.com, then Sign-In and Security, then App-Specific Passwords |
| Gmail | myaccount.google.com/apppasswords (needs 2-Step Verification) |
| Fastmail | Settings, then Privacy & Security, then Manage app passwords |
| Yahoo / AOL | Account Security, then Generate app password |
| Other | Choose "Other" and enter your IMAP/SMTP servers |

Outlook.com and Microsoft 365 accounts no longer allow app-password IMAP sign-in, so they aren't supported.

### 3. Ask Claude
> "Summarize today's emails"
> "Clear out the promotional mail in my inbox"
> "Draft a reply to the email from Sam saying Thursday works"

## Using the module yourself
```powershell
Import-Module ImapMail
Find-Mail -Since (Get-Date).Date
Read-Mail -Uid 1234
New-MailDraft -ReplyToUid 1234 -Body "Thursday works for me."
```
Run `Get-Command -Module ImapMail` to list every command.

## Where things are stored
`%APPDATA%\ImapMail\` holds `config.json` (servers and address), `credential.xml` (encrypted password), `logs\` (CSV records of cleanup moves) and `bin\` (the helper library, compiled automatically from `ImapMail.cs` on first use). To disconnect, delete that folder and revoke the app password with your provider.

## License
MIT
