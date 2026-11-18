# ImapMail for Claude

**Email and calendar for Claude on Windows, with nothing to install.** No Node.js, no Python, no Docker, and no password sitting in a config file.

Let Claude read, summarize, clean up and draft replies in **any IMAP email account** (iCloud, Gmail, Fastmail, Yahoo, AOL, Zoho or a custom server). It runs on the PowerShell that already comes with Windows and talks straight to your mail server, with no third-party service in between.

## Why this one?

Most email integrations for Claude are MCP servers that need a developer toolchain installed and keep your password in an environment variable or `.env` file. ImapMail takes a different approach:

| | ImapMail | Typical IMAP MCP server |
|---|---|---|
| Runtime | Windows PowerShell 5.1, already on every Windows 10/11 PC | Node.js, Python/uv or Docker |
| Install | Run one script | Install the runtime, then a package, then edit a JSON config |
| Password storage | Encrypted with Windows DPAPI, tied to your Windows account | Often plain text in env vars or a `.env` file |
| Setup | Guided prompt that tests your sign-in before saving | Edit config by hand |
| Sending | Drafts first; sends only what you approve | Varies; many send directly |
| Bulk cleanup | Built in, moves mail to Trash with a CSV undo log | Usually per-message only |
| How Claude uses it | A skill that drives a local PowerShell shell, so it works from Claude Code or a Claude app linked to your PC | Must be registered as an MCP server in each client |

If you're on macOS or Linux, or want email tools inside clients other than Claude, an MCP server is probably a better fit. ImapMail is for Windows users who want it to just work.

**What it can do**
- Summarize today's mail, or search by sender, subject, date or text
- Read full messages and save attachments
- Mark messages read or flagged, move them between folders, and bulk-clean promotions and notifications (to Trash, with an undo log)
- Write replies as **drafts** in your Drafts folder, and send them only after you approve
- **Calendar:** see what's coming up, add events, and remove an event or a single occurrence of a repeating one (iCloud, Fastmail, Yahoo, Zoho or any CalDAV server; every change is backed up first)

**Safety by design**
- Your app password is encrypted with Windows DPAPI under your Windows account. Claude never sees it, and it never leaves your PC.
- Claude saves drafts. It never sends anything unless you approve that specific message.
- Cleanup moves mail to Trash (recoverable) and logs every move.
- Calendar changes are confirmed with you first, and the original event is saved so it can be restored exactly.

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

Outlook.com and Microsoft 365 accounts no longer allow app-password IMAP sign-in, so they aren't supported. Gmail mail works, but Gmail **calendar** doesn't (Google requires OAuth for calendar access).

### 3. Ask Claude
> "Summarize today's emails"
> "Clear out the promotional mail in my inbox"
> "Draft a reply to the email from Sam saying Thursday works"
> "What's on my calendar this week?"
> "Take work off my calendar for the day after Thanksgiving"

## Using the module yourself
```powershell
Import-Module ImapMail
Find-Mail -Since (Get-Date).Date
Read-Mail -Uid 1234
New-MailDraft -ReplyToUid 1234 -Body "Thursday works for me."
```
Run `Get-Command -Module ImapMail` to list every command.

## Where things are stored
`%APPDATA%\ImapMail\` holds `config.json` (servers and address), `credential.xml` (encrypted password), `logs\` (CSV records of cleanup moves), `calendar-backups\` (originals of changed events) and `bin\` (the helper library, compiled automatically from `ImapMail.cs` on first use). To disconnect, delete that folder and revoke the app password with your provider.

## License
MIT
