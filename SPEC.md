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
| `search_emails` | **The one search tool.** Searches subject, sender, recipients (To/CC) and body across ALL folders incl. Sent, by default. Trigram substring + fuzzy matching, exact matches ranked first. Multi-word queries decomposed and ranked by parts matched. Just pass a query. |
| `get_email` | Full email content by ID. Strips signatures and quoted replies by default; pass `include_quoted: true` to keep them. Shows read/flagged status |
| `compose` | Draft new email, reply, or forward. Body is markdown |
| `move_email` | Move email to a folder (e.g. Archive) |
| `archive_emails` | Batch archive — moves multiple emails to Archive |
| `download_attachment` | Save attachment(s) to disk |
| `index_now` | Build/catch up the search index immediately |

There is **one** search tool (`search_emails`). The old `search_body` is gone — body content is now searched as part of the unified search.

## CLI

`index.js` doubles as a CLI (via the `mcp-outlook` bin). Same logic as the MCP tools, no server needed:

```
mcp-outlook search 'tamm@'                       # everything to/from/mentioning tamm@, ranked
mcp-outlook search 'Project X with person A'     # multi-term, ranked by parts matched
mcp-outlook search 'invoice' --folder Sent       # narrow to a folder
mcp-outlook index                                # build/catch up the index
mcp-outlook folders                              # list folders
mcp-outlook get 12345 [--quoted]                 # full email by id
mcp-outlook serve                                # run as MCP stdio server (the default when launched by a client)
```

With no subcommand: launched by an MCP client (no TTY) → starts the stdio server; run bare in a terminal → prints usage.

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

`list_folders` reads directly from Outlook's SQLite database (`~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite`). `search_emails` reads from the local FTS index (below), which is synced from that same Outlook DB. `get_email` uses AppleScript for full message content since the SQLite DB only stores previews. All sqlite invocations use `execFileSync` (no shell) — query text is passed as a literal arg, so arbitrary user input cannot inject shell commands.

### Unified search index (FTS5 trigram)

`search_emails` queries a local FTS5 index at `~/.mcp-outlook-mac/body-index.db` (override with `MCP_OUTLOOK_FTS_DB`). One `bodies` table holds every searchable field — subject, sender name, sender address, To, CC, body — plus stored metadata (folder, account, timestamp, read flag, attachment flag) so a single query renders and filters results with **no round-trip to Outlook**. An external-content FTS5 virtual table (`tokenize='trigram'`) indexes the six text columns.

- **Trigram tokenizer** gives substring-anywhere matching for queries ≥3 chars (e.g. `tamm@` matches mid-string), and natural fuzzy matching via shared trigrams.
- **Ranking** (in `handleSearch`): `ORDER BY exact_hit DESC, terms_matched DESC, best_bm ASC, ts DESC`.
  - `exact_hit` — does the full query appear as a literal substring (LIKE) in any column? Exact substrings always rank above fuzzy near-matches.
  - `terms_matched` — for multi-word queries, how many distinct terms a row matched (per-term `UNION ALL` + `GROUP BY`). More parts matched → higher.
  - `best_bm` — BM25 with column weights `[8,10,10,9,9,2]` (subject/sender/recipients weighted above body, which is preview-only in bulk).
  - The `per_term` CTE is `MATERIALIZED` — required so BM25 (an FTS5 auxiliary function) is evaluated in the same query as its MATCH.
- **Fuzzy expansion** (`expandTerms`): each term also searches its alphanumeric core (leading/trailing punctuation stripped), so `tamm@` additionally surfaces `@tamm` — ranked below exact hits.
- **Short queries** (all tokens ≤2 chars, below the trigram floor) fall back to a pure LIKE-substring path.
- Index builds incrementally: `index_now` / background refresh batch-syncs all fields using `Message_Preview` for body; every `get_email` upgrades that row's body to full text. BM25 snippets use `>>>` / `<<<` markers.
- **Schema versioned** via `PRAGMA user_version` (`FTS_SCHEMA_VERSION`). On a version bump, `ensureFtsDb` drops and rebuilds the tables and clears `indexed` to force a full re-sync — a one-time pass on first run after upgrade.

### Signature and quote stripping

`get_email` runs `cleanBody()` by default which strips:
- Standard `-- ` signatures
- Mobile signatures (Sent from my iPhone/iPad/Galaxy, Get Outlook for iOS/Android)
- "On ... wrote:" quoted reply chains
- Outlook-style "From: ... Sent: ..." blocks
- Trailing `>` quoted lines
- Forwarded message delimiters

Pass `include_quoted: true` to skip all stripping and return the full raw body.

### Compose with RTF injection

The `compose` tool uses `textutil` to convert markdown→HTML→RTF, then injects the RTF directly into the message via `read POSIX file ... as «class RTF »`. This avoids clipboard hijacking.

### Reply and reply-all recipient handling

`compose` with `mode: "reply"` delegates recipient population to Outlook's native `reply to` AppleScript command. When `reply_all: true`, the native `reply to all true` parameter is used — recipients are populated by Outlook itself, never reconstructed from the SQLite database.

The `to` and `cc` arguments are **optional overrides** with per-field semantics:

- Omit both → native reply / reply-all population is used as-is
- Provide `to` only → `to` is wiped and rebuilt from the arg; native `cc` is preserved
- Provide `cc` only → `cc` is wiped and rebuilt from the arg; native `to` is preserved
- Provide both → both fields wiped and rebuilt

This means a caller can add an extra CC to a reply-all without having to re-specify the native To list, or redirect the `to` while keeping the auto-populated CC chain.

`to` and `cc` accept a string, a comma/semicolon-separated string, a JSON array, or a JSON-stringified array (defensive: some MCP clients stringify array-typed args on the wire).

### Search parameters

`search_emails` parameters (all optional):
- `query` — matched against subject, sender name + address, To, CC and body. Substrings, fuzzy near-matches, and multi-word queries all handled automatically. Omit to list newest emails.
- `folder` — restrict to a named folder. **Omit to search all folders** (including Sent) — almost always correct.
- `account` — restrict to one account. Omit for all accounts.
- `unread_only` — only unread emails.
- `after` / `before` — date bounds (YYYY-MM-DD). Omit unless a date range is specifically needed.
- `limit` — default 500, hard cap 1000. Omit — the default is sufficient.

Results show `●` for unread and `📎` for attachments, with a BM25 snippet under each hit.

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
- **Body search is preview-text for un-opened mail, full text once opened.** Bulk indexing uses `Message_Preview` (full body is only reachable via AppleScript, one message at a time). Subject, sender and recipients are indexed completely and exactly. Opening an email via `get_email` upgrades its body to full text.
- **No BCC, no attachment-filename search.** Outlook's local SQLite doesn't store BCC or attachment names (only a has-attachment flag).
- **First run after a schema bump rebuilds the index** (one-time, well under a minute for ~35k emails; ~100 MB on disk with body trigram-indexed).
- **SQLite column names may vary.** `Message_ReadFlag` / `Message_HasAttachment` are assumed — if absent in your Outlook version, those indicators will be missing.

## Operations without dedicated tools

### Mark as read/unread
```applescript
tell application "Microsoft Outlook"
    set read status of (message id 12345) to true
end tell
```
