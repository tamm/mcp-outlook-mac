# mcp-outlook-mac

An MCP server that gives Claude (or any MCP client) read/write access to Microsoft Outlook on macOS. Search, read, reply, move, archive, download attachments — via AppleScript and direct SQLite reads.

## Requirements

- macOS with **Legacy Outlook for Mac**. New Outlook broke the AppleScript bridge for Exchange data — you must revert (`Help → Revert to Legacy Outlook`).
- Node.js 20+ (uses the built-in `node --test` runner).
- Outlook opened at least once so its SQLite profile exists at `~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite`.

## Install

```
git clone https://github.com/tamm/mcp-outlook-mac.git
cd mcp-outlook-mac
npm install
```

Wire the pre-commit hook (runs the test suite on every commit):

```
git config core.hooksPath .githooks
```

## Run the tests

```
npm test
```

That's the one canonical entry point — don't invoke `node --test` directly. It runs everything in `test/*.test.js` using Node's built-in test runner (no extra dependencies).

## Use it as an MCP server

Point your MCP client at `node /absolute/path/to/mcp-outlook-mac/index.js`. Example `.mcp.json` entry:

```json
{
  "mcpServers": {
    "mcp-outlook-mac": {
      "command": "node",
      "args": ["/absolute/path/to/mcp-outlook-mac/index.js"]
    }
  }
}
```

First invocation triggers a background FTS5 index build at `~/.mcp-outlook-mac/body-index.db`. Call `index_now` if you want to force a full catch-up pass.

## Tools

See `SPEC.md` for the full list and semantics. Quick summary:

| Tool | Purpose |
|------|---------|
| `list_folders` | All mail folders with counts |
| `search_emails` | List/filter inbox emails (biases toward full view — narrow only when needed) |
| `search_body` | Full-text body search (FTS5, BM25) |
| `get_email` | Full content by ID (strips signatures and quotes by default) |
| `compose` | Draft new / reply / forward — body is markdown |
| `move_email` | Move to any folder |
| `archive_emails` | Batch archive (move to Archive) |
| `download_attachment` | Save attachment(s) to disk |
| `index_now` | Force an immediate FTS index pass |

## Layout

```
index.js        — MCP server, all tools (source of truth)
SPEC.md         — tool semantics and implementation notes
test/           — node --test suite
.githooks/      — pre-commit runs `npm test`
```

## Contributing

1. Make changes in `index.js`.
2. Update/add tests under `test/`.
3. `npm test` must be green.
4. Keep `SPEC.md` in sync with tool changes.
5. Commit — the pre-commit hook re-runs tests.

## Known limitations

- New Outlook for Mac is unsupported. Legacy only.
- FTS index uses message previews for batch passes; full body only indexed after `get_email` reads it. Snippets from bulk-indexed rows come from the preview text.
