---
name: imap-mail
description: Reads, searches, summarizes, cleans up and drafts replies in the user's email (iCloud, Gmail, Fastmail, Yahoo or any IMAP account) through the ImapMail PowerShell module on their Windows PC. Use when the user asks to check, summarize, search, sort, clean out, or reply to their email or inbox and no dedicated email connector covers that account.
---

# Email over IMAP with the ImapMail PowerShell module

The ImapMail module talks to the user's mail server over IMAP (reading and organizing) and SMTP (sending) from their Windows PC. Run it with whatever PowerShell tool reaches that PC: a local shell in Claude Code, or a remote-device PowerShell tool.

## Before anything else

Start every command with:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
Import-Module ImapMail -Force
```

Each tool call may be a fresh process, so import every time.

- **Module not found:** it isn't installed. Tell the user to run `Install.ps1` from this plugin's `scripts` folder. If you can, launch it in a visible window for them: `Start-Process powershell.exe -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','"<path>\Install.ps1"'`. It copies the module and walks them through setup.
- **"Not set up yet," or sign-in fails:** launch setup in a visible window: `Start-Process powershell.exe -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-NoExit','-Command','Import-Module ImapMail; Initialize-ImapMail'`. The user picks a provider and types their app password there.
- **Never ask the user to paste a password into the chat,** never print `credential.xml`, and never pass a password on a command line.

## Commands

| Command | Use |
|---|---|
| `Get-MailFolder [-Counts]` | List folders, with total and unread counts |
| `Find-Mail [-Folder INBOX] [-Since d] [-Before d] [-From s] [-To s] [-Subject s] [-Text s] [-Unread] [-Flagged] [-RawCriteria s] [-Max 50]` | Search. Returns Uid, Date, From, To, Subject, Unread, Flagged, Answered, SizeKB, newest first |
| `Read-Mail -Uid <long[]> [-Folder] [-MaxChars 15000]` | Full message: headers, plain-text body, attachment list. Doesn't mark it read |
| `Save-MailAttachment -Uid n [-Folder] [-Destination] [-Name pattern]` | Save attachments (default: Downloads) |
| `Set-MailFlag -Uid <long[]> [-Folder] -Read\|-Unread\|-Flag\|-Unflag` | Change flags |
| `Move-Mail -Uid <long[]> [-Folder] -Destination name` | Move between folders |
| `Remove-Mail -Uid <long[]> [-Folder] [-Label text]` | Move to Trash (recoverable) and write a CSV log |
| `New-MailDraft -Body s [-To] [-Cc] [-Bcc] [-Subject] [-ReplyToUid n [-ReplyFolder] [-ReplyAll] [-NoQuote]]` | Save a draft. As a reply it fills in To, the Re: subject, threading headers and the quoted original |
| `Send-MailDraft -Uid <draftUid>` | Send a saved draft, file a copy in Sent, delete the draft |
| `Remove-MailDraft -Uid n` | Delete a draft |
| `Get-MailConfig` | Show the account and server settings (no password) |

UIDs belong to one folder. After a message moves, search for it again in the new folder.

## Rules

- **Never send without explicit approval of that specific message.** Write a draft with `New-MailDraft`, show the user the recipients, subject and body, and call `Send-MailDraft` only after they say to send it. Mention they can also review the draft in their own mail app.
- **Clean up with `Remove-Mail`, never by deleting.** It moves mail to Trash, where most providers keep it for about 30 days, and writes a log in `%APPDATA%\ImapMail\logs`.
- **Bulk cleanup:** fetch summaries for the whole folder and group by sender address. Classify by sender and subject: promotional, notifications, receipts and orders, bills and statements, security codes and alerts, real people. Show the categories with counts and examples, then confirm the scope with the user before moving hundreds of messages. Spot-check the selection for receipts, statements, tickets or account notices that got caught, and pull those out unless the user included that category.
- **Summaries:** group into what needs attention, security, money and everything else. Point out anything that needs action (deadlines, expiring services, payments due). Marketing bodies are mostly links, so read them with a small `-MaxChars` and describe them in your own words.
- **Point out likely phishing:** an unknown sender, an unexpected attachment (.docx, .html, .zip), a gibberish or name-only subject, or urgency about passwords or payments.
- **"Today":** IMAP `SINCE` uses the server's date. Filter results by the user's local date.
- **Gmail:** labels appear as folders, and `[Gmail]/All Mail` holds everything. Removing from INBOX in Gmail only archives unless the mail is moved to `[Gmail]/Trash`, which `Remove-Mail` does.
- Keep message contents in the conversation. Don't save them to files unless the user asks.
