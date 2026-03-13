# outlook-mcp

## What this is

A Node.js MCP server that talks to Microsoft Outlook on Mac via AppleScript. Single file: `index.js`, uses `@modelcontextprotocol/sdk`. No additional dependencies.

## Prerequisite: Legacy Outlook

**New Outlook for Mac completely breaks AppleScript access to Exchange data.** `every exchange account` returns 0, `every mail folder` only returns empty local folders.

To fix: **Help > Revert to Legacy Outlook**. Microsoft even has "AppleScript" as a reason option when they ask why you reverted. Once reverted, all 7 tools work.

If you see "No Exchange accounts found" errors from the read tools, this is almost certainly the cause.

## Tools

| Tool | What it does |
|------|-------------|
| `list_folders` | Lists all mail folders across Exchange, IMAP, and local accounts with message counts |
| `read_emails` | Returns ID, **folder**, subject, sender, date, preview. Default: Inbox, max 50 |
| `get_email` | Full content of a single email by ID. Accepts optional `folder` for fast lookup |
| `search_emails` | Case-insensitive search by subject or sender. Returns ID, **folder**, sender, subject, date |
| `create_draft` | Creates a new compose window. Body supports markdown. Signature preserved |
| `create_reply_draft` | Threaded reply to an email by ID. Accepts optional `folder` for fast lookup. Body supports markdown |
| `move_email` | Moves an email to a folder in the same account (e.g. Archive). Accepts optional `folder` for fast lookup |

### Folder pass-through pattern

`read_emails` and `search_emails` return a `Folder` field with each email. Pass this value as the `folder` parameter to `get_email`, `create_reply_draft`, or `move_email` to skip the full folder scan. If `folder` is omitted, the tool falls back to scanning all folders (slow but works).

## Architecture

```
outlook-mcp/
  index.js          — MCP server, all tools in one file
  package.json
  SPEC.md           — This file
```

Everything is in `index.js`. No `lib/` split — the file is small enough to stay monolithic.

**The tool definitions in `index.js` are the source of truth.** That's what Claude sees via MCP. If a capability isn't registered as a tool there, Claude won't know about it regardless of what this spec says.

## Exchange accounts

The server has access to Exchange accounts configured in Legacy Outlook. Account names are read from the SQLite database at runtime — no account details are hardcoded.

## Key implementation details

### Folder lookup order

All read tools search folders in this order:
1. Exchange account folders (`every mail folder of acct` for each Exchange account)
2. IMAP account folders
3. Local folders (`every mail folder`)

This matters because `every mail folder` returns both local AND Exchange folders, but local folders come first and are usually empty. The old code hit the empty local Inbox before the Exchange one.

### Folder deduplication in list_folders

`list_folders` collects folder IDs from Exchange/IMAP accounts, then only shows local folders whose IDs weren't already listed. This prevents the same folder appearing twice (once under the account name, once under "Local/").

### Sender property access

Outlook's AppleScript bridge doesn't handle chained property access on sender records. `name of sender of msg` silently returns empty. The fix is two-step:
```applescript
set sRec to sender of msg
set msgSender to name of sRec
```

### Markdown-to-HTML in drafts

Both `create_draft` and `create_reply_draft` convert the body from markdown to HTML before sending to Outlook. Outlook renders HTML in the `content` property. Plain newlines get eaten otherwise.

Supported markdown:
- **Bold** (`**text**` or `__text__`)
- *Italic* (`*text*` or `_text_`)
- Headings (`#`, `##`, `###`)
- Bullet lists (`-` or `*`)
- Numbered lists (`1.`, `2.`)
- Inline code (`` `code` ``)
- Links (`[text](url)`)
- Horizontal rules (`---`)
- Line breaks (single `\n` becomes `<br>`, double `\n\n` becomes paragraph break)

### Error handling

- **30 second timeout** on all `execSync` calls
- **-1728 retry**: if AppleScript returns error -1728 (object not found / not loaded yet), waits 2 seconds and retries once
- **No Exchange accounts**: returns a helpful message about reverting to Legacy Outlook

### Message lookup by ID (get_email, create_reply_draft, move_email)

These tools find a message by numeric ID. If a `folder` parameter is provided (from `read_emails`/`search_emails` output), it searches only that folder — fast. Without `folder`, it falls back to a linear scan of all folders across all accounts — slow but guaranteed to find the message. The `findMessageScript()` helper in `index.js` generates the AppleScript for both paths.

## Known limitations

- **New Outlook is unsupported.** Must use Legacy Outlook for any Exchange access.
- **Folder name collisions.** If two Exchange accounts both have an "Inbox", `read_emails` hits the first one found (first account). No way to specify which account's Inbox without using the full `account/folder` path (not currently implemented).
- **Message ID search is slow without folder hint.** `get_email`, `create_reply_draft`, and `move_email` scan every folder if no `folder` parameter is provided. Always pass `folder` from `read_emails`/`search_emails` output.
- **Search is client-side.** `search_emails` iterates every message in the folder and shells out to `tr` for case-insensitive comparison. Slow on large folders.
- **No attachment support.** Can't read or send attachments.
- **No read/unread filtering.** `read_emails` returns all messages, no way to filter by read status.

## Operations without dedicated tools

These can be done via `osascript` in bash using message IDs from `read_emails` or `search_emails`. No MCP tool needed.

### Mark as read/unread
```applescript
tell application "Microsoft Outlook"
    -- find targetMsg by ID first (same pattern as move_email)
    set read status of targetMsg to true
    -- or false to mark unread
end tell
```

## Future improvements (parked)

### Account-scoped folder access
Add an optional `account` parameter to `read_emails` and `search_emails` so you can target a specific Exchange account's folders.

### Microsoft Graph API (Option 4)
Use OAuth2 + Graph API for reading. Most reliable, supports shared mailboxes, cross-platform. Requires Azure app registration. Consider this when deploying to another machine or needing shared mailbox access.

### Direct .emlx reading (Option 3)
Read Apple Mail's local `.emlx` files directly. No AppleScript flakiness. Requires Full Disk Access. Consider this if Legacy Outlook's AppleScript bridge proves unreliable.
