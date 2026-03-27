# mcp-outlook-mac

## What this is

A Node.js MCP server that talks to Microsoft Outlook on Mac via AppleScript and SQLite. Single file: `index.js`, uses `@modelcontextprotocol/sdk` and `marked`.

## Prerequisite: Legacy Outlook

**New Outlook for Mac breaks AppleScript access to Exchange data.** `every exchange account` returns 0, `every mail folder` only returns empty local folders.

To fix: **Help > Revert to Legacy Outlook**. Once reverted, all tools work.

## Tools

| Tool | What it does |
|------|-------------|
| `list_folders` | Lists all folders with message counts |
| `search_emails` | Search or list emails by subject/sender. Supports date range, unread filter, sort, account filter |
| `get_email` | Full email content by ID. Strips signatures and quoted replies. Shows read/flagged status |
| `compose` | Draft new email, reply, or forward. Body is markdown |
| `move_email` | Move email to a folder (e.g. Archive) |
| `archive_emails` | Batch archive — moves multiple emails to Archive |
| `download_attachment` | Save attachment(s) to disk |
| `search_body` | Full-text body search using FTS5 index with BM25 ranking |
| `index_now` | Trigger immediate FTS5 index pass |

## Architecture

```
mcp-outlook-mac/
  index.js          — MCP server, all tools in one file
  package.json
  SPEC.md           — This file
  ~/.mcp-outlook-mac/   — FTS5 body search index (created at runtime)
```

Everything is in `index.js`. The tool definitions in `index.js` are the source of truth.

## Key implementation details

### SQLite read path

`list_folders`, `search_emails` read directly from Outlook's SQLite database (`~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite`). This is much faster than AppleScript iteration. `get_email` uses AppleScript for full message content since the SQLite DB only stores previews.

### FTS5 body search

`search_body` uses a local FTS5 index at `~/.mcp-outlook-mac/body-index.db`. The index builds incrementally:
- Every `get_email` call indexes that email's body text
- `index_now` triggers a batch pass using message previews from SQLite
- BM25 relevance ranking with `>>>` / `<<<` snippet markers

### Signature and quote stripping

`get_email` runs `cleanBody()` which strips:
- Standard `-- ` signatures
- Mobile signatures (Sent from my iPhone/iPad/Galaxy, Get Outlook for iOS/Android)
- "On ... wrote:" quoted reply chains
- Outlook-style "From: ... Sent: ..." blocks
- Trailing `>` quoted lines
- Forwarded message delimiters

### Compose with RTF injection

The `compose` tool uses `textutil` to convert markdown→HTML→RTF, then injects the RTF directly into the message via `read POSIX file ... as «class RTF »`. This avoids clipboard hijacking.

### Search features

`search_emails` supports:
- `query` — subject/sender search (omit for recent emails)
- `folder` — restrict to a folder (default: all Inboxes)
- `account` — restrict to an account
- `unread_only` — only unread emails
- `after` / `before` — date range (YYYY-MM-DD)
- `sort` — `desc` (default) or `asc`

Results show `●` for unread, `⚑` for flagged, `📎` for attachments.

### Sender property access

Outlook's AppleScript bridge doesn't handle chained property access on sender records. `name of sender of msg` silently returns empty. The fix:
```applescript
set sRec to sender of msg
set msgSender to name of sRec
```

### Error handling

- **30s timeout** on AppleScript, 15s on SQLite
- **-1728 auto-retry**: waits 2s and retries once on "object not found"
- **Diagnostic messages** for common error codes (-1728, -1712, -600, -10810, ETIMEDOUT)
- **⏱ timing** on every response

### Message lookup by ID

`get_email`, `compose` (reply/forward), `move_email` use `message id N` for direct lookup — no folder scanning needed.

## Known limitations

- **New Outlook is unsupported.** Must use Legacy Outlook.
- **FTS index uses previews for batch indexing.** Full body text is only indexed when `get_email` is called. Run `index_now` for broader coverage, but snippets will be from preview text only.
- **SQLite column names may vary.** `Message_IsRead` and `Message_IsFlagged` are assumed — if these don't exist in your Outlook version, unread/flagged indicators will be missing from search results.

## Operations without dedicated tools

### Mark as read/unread
```applescript
tell application "Microsoft Outlook"
    set read status of (message id 12345) to true
end tell
```
