import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { execSync } from "child_process";
import { marked } from "marked";
import fs from "fs";
import os from "os";
import path from "path";

const server = new Server(
  { name: "mcp-outlook-mac", version: "2.0.0" },
  { capabilities: { tools: {} } }
);

// --- Response helpers ---

function ok(text) {
  return { content: [{ type: "text", text }] };
}

function err(text) {
  return { content: [{ type: "text", text: `Error: ${text}` }], isError: true };
}

// --- AppleScript execution ---

function diagnoseAppleScriptError(error) {
  const msg = error.message || "";
  if (msg.includes("ETIMEDOUT") || error.killed)
    return "AppleScript timed out (>30s). Legacy Outlook may be unresponsive — try quitting and reopening it.";
  if (msg.includes("-1728"))
    return "Outlook object not found (-1728). The email may have been deleted or moved. Try search_emails to get fresh IDs.";
  if (msg.includes("-1712"))
    return "Outlook is busy with a dialog or modal window (-1712). Dismiss any open dialogs in Outlook and retry.";
  if (msg.includes("not running") || msg.includes("-600"))
    return "Microsoft Outlook is not running. Open Legacy Outlook (Help > Revert to Legacy Outlook) and retry.";
  if (msg.includes("-10810") || msg.includes("launch"))
    return "Could not launch Outlook. Check that Legacy Outlook is installed and not blocked by macOS.";
  const execMatch = msg.match(/execution error: (.+?) \(-?\d+\)/);
  if (execMatch) return `Outlook AppleScript error: ${execMatch[1]}`;
  return `AppleScript error: ${msg.slice(0, 300)}`;
}

function runAppleScriptHeredoc(script, retried = false) {
  try {
    const result = execSync(`osascript << 'APPLESCRIPT_EOF'\n${script}\nAPPLESCRIPT_EOF`, {
      encoding: "utf-8",
      maxBuffer: 10 * 1024 * 1024,
      shell: "/bin/bash",
      timeout: 30000,
    });
    return result.trim();
  } catch (error) {
    if (!retried && error.message && error.message.includes("-1728")) {
      execSync("sleep 2", { shell: "/bin/bash" });
      return runAppleScriptHeredoc(script, true);
    }
    throw new Error(diagnoseAppleScriptError(error));
  }
}

// --- SQLite (Outlook database) ---

const DB_PATH = path.join(
  os.homedir(),
  "Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite"
);

function runSqlite(query) {
  const flat = query.replace(/\s+/g, " ").trim();
  try {
    const result = execSync(`sqlite3 -json ${JSON.stringify(DB_PATH)} ${JSON.stringify(flat)}`, {
      encoding: "utf-8",
      maxBuffer: 10 * 1024 * 1024,
      timeout: 15000,
      shell: "/bin/bash",
    });
    const trimmed = result.trim();
    if (!trimmed) return [];
    return JSON.parse(trimmed);
  } catch (error) {
    const msg = error.message || "";
    if (msg.includes("database is locked"))
      throw new Error("Outlook database is locked — Outlook may be syncing. Wait a moment and retry.");
    if (msg.includes("no such table") || msg.includes("no such column"))
      throw new Error(`Outlook database schema mismatch: ${msg.slice(0, 200)}`);
    if (msg.includes("unable to open database") || msg.includes("not a database"))
      throw new Error("Cannot open Outlook database — ensure Legacy Outlook is installed and has been opened at least once.");
    if (msg.includes("ETIMEDOUT") || error.killed)
      throw new Error("SQLite query timed out (>15s). The database may be very large or locked.");
    throw new Error(`Outlook database error: ${msg.slice(0, 300)}`);
  }
}

// --- Text cleaning ---

function stripSignature(text) {
  if (!text) return "";
  const sigIdx = text.search(/\n-- \n/);
  if (sigIdx !== -1) text = text.slice(0, sigIdx);
  const mobilePatterns = [
    /\n\s*Sent from my iPhone\s*$/i,
    /\n\s*Sent from my iPad\s*$/i,
    /\n\s*Sent from my Galaxy\s*$/i,
    /\n\s*Get Outlook for iOS\s*$/i,
    /\n\s*Get Outlook for Android\s*$/i,
    /\n\s*Sent from Mail for Windows\s*$/i,
  ];
  for (const pat of mobilePatterns) text = text.replace(pat, "");
  return text.trim();
}

function stripQuotedReplies(text) {
  if (!text) return "";
  const onWroteIdx = text.search(/\nOn .+wrote:\s*\n/i);
  if (onWroteIdx !== -1) return text.slice(0, onWroteIdx).trim();
  const outlookIdx = text.search(/\nFrom: .+\nSent: /i);
  if (outlookIdx !== -1) return text.slice(0, outlookIdx).trim();
  const fwdIdx = text.search(/\n-{5,}\s*Forwarded message\s*-{5,}/i);
  if (fwdIdx !== -1) return text.slice(0, fwdIdx).trim();
  const lines = text.split("\n");
  let lastContentLine = lines.length - 1;
  while (lastContentLine >= 0 && /^\s*>/.test(lines[lastContentLine])) lastContentLine--;
  if (lastContentLine < lines.length - 1) {
    while (lastContentLine >= 0 && lines[lastContentLine].trim() === "") lastContentLine--;
    return lines.slice(0, lastContentLine + 1).join("\n").trim();
  }
  return text.trim();
}

function cleanBody(text) {
  return stripQuotedReplies(stripSignature(text));
}

// --- Markdown to HTML ---

function markdownToHtml(text) {
  const html = marked.parse(text, { async: false });
  return `<div style="font-family: Aptos, Calibri, sans-serif; font-size: 12pt;">${html}</div>`;
}

// --- Body injection: RTF preferred, clipboard paste fallback ---

function setRtfBody(htmlBody, msgVar = "newMsg") {
  const ts = Date.now();
  const rtfPath = `/tmp/mcp-outlook-mac-body-${ts}.rtf`;
  const tmpHtml = `/tmp/mcp-outlook-mac-body-${ts}.html`;
  fs.writeFileSync(tmpHtml, htmlBody, "utf-8");
  try {
    const rtf = execSync(
      `textutil -inputencoding UTF-8 -format html -convert rtf -stdout ${JSON.stringify(tmpHtml)}`,
      { timeout: 10000 }
    );
    fs.writeFileSync(rtfPath, rtf);
  } finally {
    try { fs.unlinkSync(tmpHtml); } catch {}
  }
  return {
    tmpPath: rtfPath,
    snippet: `set content of ${msgVar} to (read POSIX file "${rtfPath}" as «class RTF »)`,
  };
}

// Fallback: copy rich text to clipboard via textutil | pbcopy, then paste with System Events
function pasteViaClipboard(htmlBody) {
  const tmpHtml = `/tmp/mcp-outlook-mac-paste-${Date.now()}.html`;
  fs.writeFileSync(tmpHtml, htmlBody, "utf-8");
  try {
    execSync(
      `textutil -inputencoding UTF-8 -format html -convert rtf -stdout ${JSON.stringify(tmpHtml)} | pbcopy`,
      { shell: "/bin/bash", timeout: 10000 }
    );
  } finally {
    try { fs.unlinkSync(tmpHtml); } catch {}
  }
}

function pasteIntoFrontWindow() {
  return `
tell application "System Events"
    tell process "Microsoft Outlook"
        set frontWin to front window
        set bodyField to missing value
        try
            set bodyField to first text area of frontWin
        end try
        if bodyField is missing value then
            try
                set bodyField to first scroll area of frontWin
            end try
        end if
        if bodyField is not missing value then
            click bodyField
        end if
        keystroke "v" using command down
    end tell
end tell`;
}

// --- Message helpers ---

function lookupEmailLocation(emailId) {
  try {
    const rows = runSqlite(
      `SELECT f.Folder_Name, COALESCE(ae.Account_Name, am.Account_Name) AS AccountName
       FROM Mail m
       JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
       LEFT JOIN AccountsExchange ae ON ae.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
       LEFT JOIN AccountsMail am ON am.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
       WHERE m.Record_RecordID = ${emailId} LIMIT 1`
    );
    if (rows && rows.length > 0) {
      return { folderName: rows[0].Folder_Name, accountName: rows[0].AccountName };
    }
  } catch {}
  return null;
}

function findMessageScript(emailId) {
  return `
    set targetMsg to missing value
    try
        set targetMsg to message id ${emailId}
    end try`;
}

function uniquePath(dir, emailId, filename) {
  const ext = path.extname(filename);
  const base = path.basename(filename, ext);
  const prefix = `${emailId}_${base}`;
  let candidate = path.join(dir, `${prefix}${ext}`);
  if (!fs.existsSync(candidate)) return candidate;
  const trailingNum = prefix.match(/_(\d+)$/);
  let stem = prefix;
  let n = 1;
  if (trailingNum) {
    stem = prefix.slice(0, -trailingNum[0].length);
    n = parseInt(trailingNum[1], 10) + 1;
  }
  while (true) {
    candidate = path.join(dir, `${stem}_${n}${ext}`);
    if (!fs.existsSync(candidate)) return candidate;
    n++;
  }
}

function escapeForAppleScript(str) {
  if (!str) return "";
  return str.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\n/g, "\\n").replace(/\r/g, "\\r");
}

function extractEmail(raw) {
  if (raw == null) return "";
  let s = String(raw).trim();
  if (!s) return "";
  // "Display Name <email@example.com>" → email@example.com
  const angleMatch = s.match(/<([^>]+)>/);
  if (angleMatch) return angleMatch[1].trim();
  // Strip leading/trailing JSON-array/quote junk from malformed upstream input
  s = s.replace(/^[\[\s"']+/, "").replace(/[\]\s"']+$/, "");
  return s;
}

function parseRecipients(input) {
  if (!input) return [];
  if (Array.isArray(input)) return input.map(extractEmail).filter(Boolean);
  // Defensive: some clients stringify arrays. Try JSON-parse if it looks like one.
  if (typeof input === "string") {
    const trimmed = input.trim();
    if (trimmed.startsWith("[") && trimmed.endsWith("]")) {
      try {
        const parsed = JSON.parse(trimmed);
        if (Array.isArray(parsed)) return parsed.map(extractEmail).filter(Boolean);
      } catch {}
    }
    return trimmed.split(/[,;]/).map(extractEmail).filter(Boolean);
  }
  return [];
}

function recipientLines(addresses, type, msgVar) {
  return addresses.map(addr =>
    `make new ${type} recipient at ${msgVar} with properties {email address:{address:"${escapeForAppleScript(addr)}"}}`
  ).join("\n    ");
}

// --- Resolve folder IDs for queries ---

function resolveFolder(folderName) {
  const rows = runSqlite(
    `SELECT Record_RecordID FROM Folders WHERE Folder_Name = ${JSON.stringify(folderName)}
     ORDER BY (SELECT COUNT(*) FROM Mail m WHERE m.Record_FolderID = Record_RecordID) DESC LIMIT 1`
  );
  if (!rows || rows.length === 0) return null;
  return rows[0].Record_RecordID;
}

function resolveInboxIds() {
  const rows = runSqlite(`SELECT Record_RecordID FROM Folders WHERE Folder_Name = 'Inbox'`);
  if (!rows || rows.length === 0) return null;
  return rows.map(r => r.Record_RecordID);
}

// --- FTS5 body search index ---

const FTS_DIR = path.join(os.homedir(), ".mcp-outlook-mac");
const FTS_DB = path.join(FTS_DIR, "body-index.db");
const INDEX_BATCH_SIZE = 10000;
const INDEX_REFRESH_INTERVAL_MS = 60 * 1000;
let indexRefreshRunning = false;
let indexRefreshTimer = null;

function runFts(query, options = {}) {
  const q = query.replace(/\s+/g, " ").trim();
  try {
    return execSync(
      `sqlite3 ${JSON.stringify(FTS_DB)} ${JSON.stringify(q)}`,
      {
        encoding: "utf-8",
        maxBuffer: 10 * 1024 * 1024,
        timeout: options.timeout || 10000,
      }
    ).trim();
  } catch (error) {
    throw new Error(`FTS error: ${(error.message || "").slice(0, 200)}`);
  }
}

function parseFts(query) {
  const q = query.replace(/\s+/g, " ").trim();
  try {
    const raw = execSync(
      `sqlite3 -json ${JSON.stringify(FTS_DB)} ${JSON.stringify(q)}`,
      { encoding: "utf-8", maxBuffer: 10 * 1024 * 1024, timeout: 10000 }
    ).trim();
    if (!raw) return [];
    return JSON.parse(raw);
  } catch { return []; }
}

function ensureFtsDb() {
  if (!fs.existsSync(FTS_DIR)) fs.mkdirSync(FTS_DIR, { recursive: true });
  runFts(`CREATE TABLE IF NOT EXISTS indexed (id INTEGER PRIMARY KEY, ts INTEGER)`);
  runFts(`CREATE TABLE IF NOT EXISTS bodies (id INTEGER PRIMARY KEY, subject TEXT, sender TEXT, body TEXT)`);
  runFts(`CREATE VIRTUAL TABLE IF NOT EXISTS fts USING fts5(subject, sender, body, content='bodies', content_rowid='id')`);
  runFts(`CREATE TRIGGER IF NOT EXISTS bodies_ai AFTER INSERT ON bodies BEGIN INSERT INTO fts(rowid, subject, sender, body) VALUES (new.id, new.subject, new.sender, new.body); END`);
  runFts(`CREATE TRIGGER IF NOT EXISTS bodies_ad AFTER DELETE ON bodies BEGIN INSERT INTO fts(fts, rowid, subject, sender, body) VALUES ('delete', old.id, old.subject, old.sender, old.body); END`);
}

function isIndexed(emailId) {
  return parseFts(`SELECT 1 FROM indexed WHERE id = ${Number(emailId)} LIMIT 1`).length > 0;
}

function indexEmail(emailId, subject, sender, bodyText) {
  const id = Number(emailId);
  if (isIndexed(id)) return;
  const body = (bodyText || "").slice(0, 50000).replace(/'/g, "''");
  const subj = (subject || "").replace(/'/g, "''");
  const from = (sender || "").replace(/'/g, "''");
  try {
    runFts(`INSERT OR REPLACE INTO bodies (id, subject, sender, body) VALUES (${id}, '${subj}', '${from}', '${body}')`);
    runFts(`INSERT OR REPLACE INTO indexed (id, ts) VALUES (${id}, ${Math.floor(Date.now() / 1000)})`);
  } catch {}
}

function syncFtsIndex(options = {}) {
  ensureFtsDb();

  const batchSize = Math.max(1, Number(options.batchSize) || INDEX_BATCH_SIZE);
  const maxBatches = Math.max(1, Number(options.maxBatches) || Number.MAX_SAFE_INTEGER);
  const escapedDbPath = DB_PATH.replace(/'/g, "''");
  let batches = 0;
  let inserted = 0;
  let lastBatchSize = 0;

  while (batches < maxBatches) {
    const batchInserted = Number(runFts(`
      ATTACH DATABASE '${escapedDbPath}' AS outlook;
      DROP TABLE IF EXISTS temp.batch_rows;
      CREATE TEMP TABLE batch_rows AS
      SELECT
        m.Record_RecordID AS id,
        COALESCE(m.Message_NormalizedSubject, '') AS subject,
        COALESCE(m.Message_SenderList, '') AS sender,
        COALESCE(m.Message_Preview, '') AS body
      FROM outlook.Mail m
      LEFT JOIN indexed i ON i.id = m.Record_RecordID
      LEFT JOIN bodies b ON b.id = m.Record_RecordID
      WHERE i.id IS NULL OR b.id IS NULL
      ORDER BY m.Record_RecordID
      LIMIT ${batchSize};
      BEGIN;
      INSERT OR IGNORE INTO bodies (id, subject, sender, body)
      SELECT id, subject, sender, body FROM batch_rows;
      INSERT OR IGNORE INTO indexed (id, ts)
      SELECT id, strftime('%s', 'now') FROM batch_rows;
      COMMIT;
      SELECT COUNT(*) FROM batch_rows;
      DROP TABLE temp.batch_rows;
      DETACH DATABASE outlook;
    `, { timeout: 60000 }));

    lastBatchSize = Number.isFinite(batchInserted) ? batchInserted : 0;
    inserted += lastBatchSize;
    batches++;

    if (lastBatchSize < batchSize) break;
  }

  return {
    batches,
    indexed: inserted,
    complete: lastBatchSize < batchSize,
    stats: ftsIndexStats(),
  };
}

function triggerIndexRefresh() {
  if (indexRefreshRunning) return;
  indexRefreshRunning = true;

  setImmediate(() => {
    try {
      syncFtsIndex({ maxBatches: 1 });
    } catch {}
    indexRefreshRunning = false;
  });
}

function startIndexRefreshLoop() {
  if (indexRefreshTimer) return;
  indexRefreshTimer = setInterval(() => {
    triggerIndexRefresh();
  }, INDEX_REFRESH_INTERVAL_MS);
  if (typeof indexRefreshTimer.unref === "function") indexRefreshTimer.unref();
}

function ftsIndexStats() {
  try {
    const indexed = parseFts(`SELECT COUNT(*) as n FROM indexed`);
    const total = runSqlite(`SELECT COUNT(*) as n FROM Mail`);
    return {
      indexed: indexed[0]?.n || 0,
      total: total[0]?.n || 0,
    };
  } catch {
    return { indexed: 0, total: 0 };
  }
}

function ftsStatusLine(stats) {
  const { indexed, total } = stats;
  if (!total) return "Index: empty";
  const pct = Math.round((indexed / total) * 100);
  if (pct >= 100) return `Index: complete (${indexed.toLocaleString()} emails)`;
  return `Index: ${pct}% — ${indexed.toLocaleString()}/${total.toLocaleString()} indexed`;
}

// --- Tool definitions ---

const TOOLS = [
  {
    name: "list_folders",
    description: "List all mail folders with message counts.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "search_emails",
    description: "List inbox emails matching optional filters. Omit all filters to return the full inbox (newest first). Bias toward unbounded scope — only narrow when the task actually requires it. Good for triage, where missing older unread items is worse than a larger result set.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Optional substring match on subject or sender. Omit to match all emails." },
        folder: { type: "string", description: "Optional folder name. Omit to search all Inboxes." },
        account: { type: "string", description: "Optional account name. Omit to search all accounts." },
        limit: { type: "number", description: "Optional max results (hard cap 1000). Omit to return up to 500 — prefer omitting for triage so older items are not missed." },
        unread_only: { type: "boolean", description: "Only unread emails." },
        sort: { type: "string", enum: ["desc", "asc"], description: "Sort by date (default: desc)." },
        after: { type: "string", description: "Optional date lower bound (YYYY-MM-DD). Only use when you specifically need to bound by date. Omit to search regardless of date." },
        before: { type: "string", description: "Optional date upper bound (YYYY-MM-DD). Only use when you specifically need to bound by date. Omit to search regardless of date." },
      },
    },
  },
  {
    name: "get_email",
    description: "Get full email content by ID. Strips signatures and quoted replies by default.",
    inputSchema: {
      type: "object",
      properties: {
        email_id: { type: "number", description: "Email ID from search_emails." },
        include_quoted: { type: "boolean", description: "Keep quoted replies and signatures (default false)." },
      },
      required: ["email_id"],
    },
  },
  {
    name: "compose",
    description: "Draft new email, reply, or forward. Body is markdown.",
    inputSchema: {
      type: "object",
      properties: {
        mode: { type: "string", enum: ["new", "reply", "forward"], description: "Compose mode." },
        to: { oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }], description: "Recipient(s). String, comma-separated string, or array." },
        subject: { type: "string", description: "Subject (new)." },
        body: { type: "string", description: "Markdown body." },
        cc: { oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }], description: "CC recipient(s). String, comma-separated string, or array." },
        email_id: { type: "number", description: "Email ID (reply/forward)." },
        reply_all: { type: "boolean", description: "Reply all." },
      },
      required: ["mode", "body"],
    },
  },
  {
    name: "move_email",
    description: "Move email to a folder (e.g. Archive). Use list_folders for names.",
    inputSchema: {
      type: "object",
      properties: {
        email_id: { type: "number", description: "Email ID." },
        destination_folder: { type: "string", description: "Target folder name." },
      },
      required: ["email_id", "destination_folder"],
    },
  },
  {
    name: "archive_emails",
    description: "Archive emails (move to Archive folder). Takes multiple IDs.",
    inputSchema: {
      type: "object",
      properties: {
        email_ids: {
          type: "array",
          items: { type: "number" },
          description: "Array of email IDs to archive.",
        },
      },
      required: ["email_ids"],
    },
  },
  {
    name: "download_attachment",
    description: "Download attachment(s) from an email. Use get_email to see filenames.",
    inputSchema: {
      type: "object",
      properties: {
        email_id: { type: "number", description: "Email ID." },
        attachment_name: { type: "string", description: "Filename (omit for all)." },
        destination: { type: "string", description: "Save directory (default: /tmp/outlook-attachments/)." },
      },
      required: ["email_id"],
    },
  },
  {
    name: "search_body",
    description: "Full-text body search across all indexed emails (FTS5, BM25-ranked). Index refreshes in the background. Bias toward unbounded scope — only narrow limit when the task actually requires it.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Search terms." },
        limit: { type: "number", description: "Optional max results (hard cap 1000). Omit to return up to 500." },
      },
      required: ["query"],
    },
  },
  {
    name: "index_now",
    description: "Build or catch up the full FTS5 index immediately.",
    inputSchema: { type: "object", properties: {} },
  },
];

// --- Tool routing ---

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  const t0 = Date.now();
  try {
    let result;
    switch (name) {
      case "list_folders": result = handleListFolders(args); break;
      case "search_emails": result = handleSearchEmails(args); break;
      case "get_email": result = handleGetEmail(args); break;
      case "compose": result = handleCompose(args); break;
      case "move_email": result = handleMoveEmail(args); break;
      case "archive_emails": result = handleArchiveEmails(args); break;
      case "download_attachment": result = handleDownloadAttachment(args); break;
      case "search_body": result = handleSearchBody(args); break;
      case "index_now": result = handleIndexNow(args); break;
      default: return err(`Unknown tool: ${name}`);
    }
    const ms = Date.now() - t0;
    if (result.content?.[0]?.text) result.content[0].text += `\n\n⏱ ${ms}ms`;
    return result;
  } catch (error) {
    return err(error.message);
  }
});

// --- Handler: list_folders ---

function handleListFolders() {
  const rows = runSqlite(`
    SELECT f.Folder_Name,
           COALESCE(ae.Account_Name, am.Account_Name, 'Local') AS AccountName,
           (SELECT COUNT(*) FROM Mail m WHERE m.Record_FolderID = f.Record_RecordID) AS MsgCount
    FROM Folders f
    LEFT JOIN AccountsExchange ae ON ae.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
    LEFT JOIN AccountsMail am ON am.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
    WHERE f.Folder_Name NOT LIKE 'Placeholder%' AND f.Folder_Name != ''
    ORDER BY AccountName, f.Folder_Name
  `);
  if (!rows || rows.length === 0) return ok("No folders found.");
  const lines = rows.map(r => `${r.AccountName}/${r.Folder_Name} (${r.MsgCount} messages)`);
  return ok(lines.join("\n"));
}

// --- Handler: search_emails ---

function handleSearchEmails(args) {
  const query = args?.query || null;
  const folderName = args?.folder || null;
  const accountName = args?.account || null;
  const limit = Math.min(args?.limit || 500, 1000);
  const sortDir = args?.sort === "asc" ? "ASC" : "DESC";

  // Resolve folder filter
  let folderClause;
  let folderLabel;
  if (folderName) {
    const folderId = resolveFolder(folderName);
    if (!folderId) return err(`Folder '${folderName}' not found.`);
    folderClause = `m.Record_FolderID = ${folderId}`;
    folderLabel = folderName;
  } else {
    const inboxIds = resolveInboxIds();
    if (!inboxIds) return err("No Inbox folders found.");
    folderClause = `m.Record_FolderID IN (${inboxIds.join(",")})`;
    folderLabel = "All Inboxes";
  }

  // Account filter
  let accountClause = "";
  if (accountName) {
    accountClause = `AND (ae.Account_Name = ${JSON.stringify(accountName)} OR am.Account_Name = ${JSON.stringify(accountName)})`;
  }

  // Search filter
  let searchClause = "";
  if (query) {
    const escaped = query.replace(/'/g, "''");
    searchClause = `AND (m.Message_NormalizedSubject LIKE '%${escaped}%' OR m.Message_SenderList LIKE '%${escaped}%')`;
  }

  // Unread filter
  const unreadClause = args?.unread_only ? "AND m.Message_ReadFlag = 0" : "";

  // Date filters (Outlook stores Message_TimeReceived as Unix timestamp)
  let dateClause = "";
  if (args?.after) {
    const ts = Math.floor(new Date(args.after).getTime() / 1000);
    if (!isNaN(ts)) dateClause += ` AND m.Message_TimeReceived >= ${ts}`;
  }
  if (args?.before) {
    const ts = Math.floor(new Date(args.before + "T23:59:59").getTime() / 1000);
    if (!isNaN(ts)) dateClause += ` AND m.Message_TimeReceived <= ${ts}`;
  }

  const emails = runSqlite(`
    SELECT m.Record_RecordID, m.Message_NormalizedSubject, m.Message_SenderList,
           m.Message_TimeReceived, m.Message_HasAttachment,
           m.Message_ReadFlag, m.Record_FlagStatus,
           f.Folder_Name,
           COALESCE(ae.Account_Name, am.Account_Name, 'Local') AS AccountName
    FROM Mail m
    JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
    LEFT JOIN AccountsExchange ae ON ae.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
    LEFT JOIN AccountsMail am ON am.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
    WHERE ${folderClause}
      ${accountClause}
      ${searchClause}
      ${unreadClause}
      ${dateClause}
    ORDER BY m.Message_TimeReceived ${sortDir}
    LIMIT ${limit}
  `);

  if (!emails || emails.length === 0) {
    return ok(query ? `No emails matching "${query}" in ${folderLabel}.` : `No emails in ${folderLabel}.`);
  }

  const lines = emails.map((msg) => {
    const date = msg.Message_TimeReceived
      ? new Date(msg.Message_TimeReceived * 1000).toISOString().slice(0, 10)
      : "";
    const from = msg.Message_SenderList || "";
    const account = msg.AccountName ? `${msg.AccountName}/` : "";
    const unread = msg.Message_ReadFlag ? " " : "●";
    const flagged = msg.Record_FlagStatus ? "⚑ " : "";
    const attach = msg.Message_HasAttachment ? " 📎" : "";
    return `ID:${msg.Record_RecordID} | ${unread} ${flagged}${date} | ${from} | ${msg.Message_NormalizedSubject || "(no subject)"}${attach} [${account}${msg.Folder_Name}]`;
  });
  return ok(lines.join("\n"));
}

// --- Handler: get_email ---

function handleGetEmail(args) {
  const emailId = args?.email_id;
  if (!emailId) return err("email_id is required.");

  const script = `
tell application "Microsoft Outlook"
    set sourceAcct to missing value
${findMessageScript(emailId)}

    if targetMsg is missing value then
        return "NOT_FOUND"
    end if

    set msgSubject to subject of targetMsg
    set msgSender to ""
    set msgSenderEmail to ""
    try
        set sRec to sender of targetMsg
        set msgSender to name of sRec
        set msgSenderEmail to address of sRec
    end try
    set msgDate to time received of targetMsg as string
    set msgContent to ""
    try
        set msgContent to plain text content of targetMsg
    end try
    set isRead to is read of targetMsg
    set flagStatus to todo flag of targetMsg

    set toList to ""
    try
        set toRecips to to recipients of targetMsg
        repeat with r in toRecips
            set toList to toList & (email address of r as string) & ", "
        end repeat
    end try

    set ccList to ""
    try
        set ccRecips to cc recipients of targetMsg
        repeat with r in ccRecips
            set ccList to ccList & (email address of r as string) & ", "
        end repeat
    end try

    set attInfo to ""
    try
        set attList to attachments of targetMsg
        set attCount to count of attList
        if attCount > 0 then
            set attInfo to "\\nAttachments (" & attCount & "):"
            repeat with att in attList
                set attName to name of att
                set attSize to file size of att
                set attInfo to attInfo & "\\n  - " & attName & " (" & attSize & " bytes)"
            end repeat
        end if
    end try

    set readFlag to "false"
    if isRead then set readFlag to "true"
    set flaggedFlag to "false"
    if flagStatus is not not flagged then set flaggedFlag to "true"

    set output to "Subject: " & msgSubject & "\\n"
    set output to output & "From: " & msgSender & " <" & msgSenderEmail & ">\\n"
    set output to output & "To: " & toList & "\\n"
    if ccList is not "" then
        set output to output & "CC: " & ccList & "\\n"
    end if
    set output to output & "Date: " & msgDate & "\\n"
    set output to output & "Read: " & readFlag & " | Flagged: " & flaggedFlag & "\\n"
    if attInfo is not "" then
        set output to output & attInfo & "\\n"
    end if
    set output to output & "\\n" & msgContent

    return output
end tell`;

  const result = runAppleScriptHeredoc(script);
  if (result === "NOT_FOUND") return err(`Email ${emailId} not found.`);

  // Supplement CC from DB if AppleScript returned none (Outlook UI/AS bug can hide CC)
  let resultWithCc = result;
  if (!/^CC: .+/m.test(result)) {
    try {
      const rows = runSqlite(
        `SELECT Message_CCRecipientAddressList FROM Mail WHERE Record_RecordID = ${Number(emailId)} LIMIT 1`
      );
      const dbCc = rows?.[0]?.Message_CCRecipientAddressList;
      if (dbCc && dbCc.trim()) {
        resultWithCc = result.replace(/^(To: .*)$/m, `$1\nCC: ${dbCc}`);
      }
    } catch {}
  }

  // Split off headers from body for cleaning
  const bodyMarker = resultWithCc.indexOf("\n\n");
  let output;
  if (bodyMarker !== -1) {
    const headers = resultWithCc.slice(0, bodyMarker);
    const rawBody = resultWithCc.slice(bodyMarker + 2);
    output = headers + "\n\n" + (args?.include_quoted ? rawBody.trim() : cleanBody(rawBody));
  } else {
    output = resultWithCc;
  }

  // Index for FTS (fire-and-forget)
  try {
    const subjectMatch = resultWithCc.match(/^Subject: (.*)$/m);
    const senderMatch = resultWithCc.match(/^From: (.*)$/m);
    const bodyText = bodyMarker !== -1 ? resultWithCc.slice(bodyMarker + 2) : "";
    indexEmail(emailId, subjectMatch?.[1] || "", senderMatch?.[1] || "", bodyText);
  } catch {}

  return ok(output);
}

// --- Handler: compose ---

function handleCompose(args) {
  const mode = args?.mode;
  const body = args?.body || "";
  const htmlBody = markdownToHtml(body);

  if (mode === "new") {
    return composeNew(args, body, htmlBody);
  }

  if (mode === "reply" || mode === "forward") {
    return composeReplyOrForward(args, mode, body, htmlBody);
  }

  return err(`Invalid mode: ${mode}. Use "new", "reply", or "forward".`);
}

function composeNew(args, body, htmlBody) {
  const subject = args?.subject || "";
  const toAddrs = parseRecipients(args?.to);
  const ccAddrs = parseRecipients(args?.cc);

  // Try RTF injection first
  if (body.trim()) {
    let tmpPath = null;
    try {
      const rtf = setRtfBody(htmlBody, "newMsg");
      tmpPath = rtf.tmpPath;
      let script = `tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {subject:"${escapeForAppleScript(subject)}"}`;
      if (toAddrs.length) script += `\n    ${recipientLines(toAddrs, "to", "newMsg")}`;
      if (ccAddrs.length) script += `\n    ${recipientLines(ccAddrs, "cc", "newMsg")}`;
      script += `\n    ${rtf.snippet}`;
      script += `\n    open newMsg\n    activate\nend tell`;
      runAppleScriptHeredoc(script);
      return ok(`Draft created: ${subject}`);
    } catch (rtfErr) {
      // RTF injection failed — fall back to HTML content property
      console.error(`RTF injection failed, falling back to HTML: ${rtfErr.message}`);
    } finally {
      if (tmpPath) try { fs.unlinkSync(tmpPath); } catch {}
    }
  }

  // Fallback: set HTML via content property directly
  const escapedBody = htmlBody.replace(/"/g, '\\"');
  let script = `tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {subject:"${escapeForAppleScript(subject)}", content:"${escapedBody}"}`;
  if (toAddrs.length) script += `\n    ${recipientLines(toAddrs, "to", "newMsg")}`;
  if (ccAddrs.length) script += `\n    ${recipientLines(ccAddrs, "cc", "newMsg")}`;
  script += `\n    open newMsg\n    activate\nend tell`;
  runAppleScriptHeredoc(script);
  return ok(`Draft created: ${subject}`);
}

// Build an AppleScript snippet that wipes and rebuilds recipient lists.
// Only the fields explicitly provided are touched; the other is left as-is
// (e.g. reply-all's native CC is preserved when the caller overrides only `to`).
function setRecipientsScript(toAddrs, ccAddrs, msgVar) {
  const hasTo = Array.isArray(toAddrs);
  const hasCc = Array.isArray(ccAddrs);
  const to = hasTo ? toAddrs.filter(Boolean) : [];
  const cc = hasCc ? ccAddrs.filter(Boolean) : [];
  const lines = [];
  if (hasTo) {
    lines.push(`delete every to recipient of ${msgVar}`);
    for (const addr of to) {
      lines.push(`make new to recipient at ${msgVar} with properties {email address:{address:"${escapeForAppleScript(addr)}"}}`);
    }
  }
  if (hasCc) {
    lines.push(`delete every cc recipient of ${msgVar}`);
    for (const addr of cc) {
      lines.push(`make new cc recipient at ${msgVar} with properties {email address:{address:"${escapeForAppleScript(addr)}"}}`);
    }
  }
  return lines.join("\n    ");
}

function composeReplyOrForward(args, mode, body, htmlBody) {
  const emailId = args?.email_id;
  if (!emailId) return err("email_id required for reply/forward.");
  const replyAll = args?.reply_all || false;

  // Verify the message exists
  const findScript = `
tell application "Microsoft Outlook"
${findMessageScript(emailId)}
    if targetMsg is missing value then return "NOT_FOUND"
    return id of targetMsg
end tell`;
  const findResult = runAppleScriptHeredoc(findScript);
  if (findResult === "NOT_FOUND") return err(`Email ${emailId} not found.`);

  // Use Outlook's native reply-all — it populates recipients correctly on its own.
  const action = mode === "reply"
    ? (replyAll ? "reply to targetMsg reply to all true" : "reply to targetMsg")
    : "forward targetMsg";

  // Only override recipients when the caller explicitly passes `to`/`cc`.
  // Provided field wipes and rebuilds that field only; the other is left untouched
  // so native reply/reply-all population is preserved for the unspecified side.
  let recipientOverride = null;
  if (mode === "reply") {
    const hasTo = args?.to != null;
    const hasCc = args?.cc != null;
    if (hasTo || hasCc) {
      recipientOverride = {
        to: hasTo ? parseRecipients(args.to) : null,
        cc: hasCc ? parseRecipients(args.cc) : null,
      };
    }
  }

  // Step 1: open the reply/forward window and capture the new message's id.
  // Targeting by id avoids the `message of front window` race entirely —
  // Outlook has a number of visible=false phantom windows (Contact Suggestions
  // autocomplete popups, stale inspectors) that compete for "front window".
  const openScript = `
tell application "Microsoft Outlook"
    activate
${findMessageScript(emailId)}
    if targetMsg is missing value then return "NOT_FOUND"
    set composeMsg to ${action} with opening window
    return (id of composeMsg) as string
end tell`;
  const composeIdRaw = runAppleScriptHeredoc(openScript);
  if (composeIdRaw === "NOT_FOUND") return err(`Email ${emailId} not found.`);
  const composeId = Number(composeIdRaw);
  if (!Number.isFinite(composeId) || composeId <= 0) {
    return err(`Failed to open ${mode} window for email ${emailId}: got ${composeIdRaw}`);
  }

  // Step 2: inject body + recipient overrides by targeting the draft by id.
  const recipientFix = recipientOverride
    ? `\n    ${setRecipientsScript(recipientOverride.to, recipientOverride.cc, "composeMsg")}`
    : "";

  if (body.trim()) {
    let tmpPath = null;
    try {
      const rtf = setRtfBody(htmlBody, "composeMsg");
      tmpPath = rtf.tmpPath;
      const injectScript = `
tell application "Microsoft Outlook"
    set composeMsg to message id ${composeId}
    ${rtf.snippet}${recipientFix}
end tell`;
      const result = runAppleScriptHeredoc(injectScript);
      if (result && result.startsWith("Error:")) return err(result.slice(7));
      return ok(`${mode === "reply" ? "Reply" : "Forward"} draft created for email ${emailId}.`);
    } catch (rtfErr) {
      console.error(`[compose] RTF injection failed for draft ${composeId}, falling back to clipboard paste: ${rtfErr.message}`);
      pasteViaClipboard(htmlBody);
      const fallbackScript = `
tell application "Microsoft Outlook"
    set composeMsg to message id ${composeId}${recipientFix}
end tell
${pasteIntoFrontWindow()}`;
      const result = runAppleScriptHeredoc(fallbackScript);
      if (result && result.startsWith("Error:")) return err(result.slice(7));
      return ok(`${mode === "reply" ? "Reply" : "Forward"} draft created for email ${emailId}.`);
    } finally {
      if (tmpPath) try { fs.unlinkSync(tmpPath); } catch {}
    }
  }

  // No body: just apply recipient overrides if any.
  if (recipientOverride) {
    const fixScript = `
tell application "Microsoft Outlook"
    set composeMsg to message id ${composeId}${recipientFix}
end tell`;
    const result = runAppleScriptHeredoc(fixScript);
    if (result && result.startsWith("Error:")) return err(result.slice(7));
  }
  return ok(`${mode === "reply" ? "Reply" : "Forward"} draft created for email ${emailId}.`);
}

// --- Handler: move_email ---

function handleMoveEmail(args) {
  const emailId = args?.email_id;
  const destFolder = args?.destination_folder;
  if (!emailId || !destFolder) return err("email_id and destination_folder required.");

  const escapedDest = escapeForAppleScript(destFolder);
  const loc = lookupEmailLocation(emailId);

  let destFolderScript;
  if (loc && loc.accountName) {
    const escapedAcct = escapeForAppleScript(loc.accountName);
    destFolderScript = `
    set destF to missing value
    try
        set destF to folder "${escapedDest}" of exchange account "${escapedAcct}"
    end try
    if destF is missing value then
        repeat with f in every mail folder whose name is "${escapedDest}"
            set destF to f
            exit repeat
        end repeat
    end if`;
  } else {
    destFolderScript = `
    set destF to missing value
    repeat with f in every mail folder whose name is "${escapedDest}"
        set destF to f
        exit repeat
    end repeat`;
  }

  const script = `
tell application "Microsoft Outlook"
${findMessageScript(emailId)}
    if targetMsg is missing value then
        return "Error: Email ${emailId} not found."
    end if
${destFolderScript}
    if destF is missing value then
        return "Error: Folder '${escapedDest}' not found."
    end if
    move targetMsg to destF
    return "Moved email ${emailId} to ${escapedDest}."
end tell`;

  const result = runAppleScriptHeredoc(script);
  if (result.startsWith("Error:")) return err(result.slice(7));
  return ok(result);
}

// --- Handler: archive_emails ---

function coerceEmailIds(raw) {
  if (raw == null) return [];
  let ids = raw;
  if (typeof ids === "string") {
    try { ids = JSON.parse(ids); } catch { ids = [ids]; }
  }
  if (!Array.isArray(ids)) ids = [ids];
  return ids.map(Number).filter(n => !isNaN(n) && n > 0);
}

function handleArchiveEmails(args) {
  const emailIds = coerceEmailIds(args?.email_ids);
  if (emailIds.length === 0) return err("email_ids required.");

  // Group emails by account for efficient batch moves
  const byAccount = {};
  const notFound = [];
  for (const id of emailIds) {
    const loc = lookupEmailLocation(Number(id));
    if (!loc) { notFound.push(id); continue; }
    const key = loc.accountName || "unknown";
    if (!byAccount[key]) byAccount[key] = [];
    byAccount[key].push(Number(id));
  }

  let archived = 0;
  const errors = [];

  for (const [accountName, ids] of Object.entries(byAccount)) {
    const escapedAcct = escapeForAppleScript(accountName);
    const script = `
tell application "Microsoft Outlook"
    set destF to missing value
    try
        set destF to folder "Archive" of exchange account "${escapedAcct}"
    end try
    if destF is missing value then
        repeat with f in every mail folder whose name is "Archive"
            set destF to f
            exit repeat
        end repeat
    end if
    if destF is missing value then
        return "Error: Archive folder not found."
    end if
    set moveCount to 0
    ${ids.map(id => `
    try
        set msg to message id ${id}
        move msg to destF
        set moveCount to moveCount + 1
    end try`).join("")}
    return moveCount as string
end tell`;

    try {
      const count = parseInt(runAppleScriptHeredoc(script), 10) || 0;
      archived += count;
    } catch (e) {
      errors.push(`${accountName}: ${e.message}`);
    }
  }

  const parts = [`Archived ${archived} of ${emailIds.length} emails.`];
  if (notFound.length > 0) parts.push(`Not found: ${notFound.join(", ")}`);
  if (errors.length > 0) parts.push(`Errors: ${errors.join("; ")}`);
  return ok(parts.join("\n"));
}

// --- Handler: download_attachment ---

function handleDownloadAttachment(args) {
  const emailId = args?.email_id;
  if (!emailId) return err("email_id is required.");

  const attachmentName = args?.attachment_name || null;
  const escapedAttName = attachmentName ? escapeForAppleScript(attachmentName) : "";
  const saveDir = args?.destination || "/tmp/outlook-attachments";

  fs.mkdirSync(saveDir, { recursive: true });
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "outlook-att-"));

  const script = attachmentName
    ? `
tell application "Microsoft Outlook"
${findMessageScript(emailId)}
    if targetMsg is missing value then return "Error: Email ${emailId} not found."
    set attList to attachments of targetMsg
    set found to false
    repeat with att in attList
        if name of att is "${escapedAttName}" then
            set savePath to "${tmpDir}/" & name of att
            save att in POSIX file savePath
            set found to true
            return (name of att) & "\\t" & (file size of att)
        end if
    end repeat
    if not found then
        set available to ""
        repeat with att in attList
            set available to available & "\\n  - " & name of att
        end repeat
        return "Error: Attachment \\"${escapedAttName}\\" not found. Available:" & available
    end if
end tell`
    : `
tell application "Microsoft Outlook"
${findMessageScript(emailId)}
    if targetMsg is missing value then return "Error: Email ${emailId} not found."
    set attList to attachments of targetMsg
    set attCount to count of attList
    if attCount is 0 then return "Error: Email has no attachments."
    set output to ""
    repeat with att in attList
        set savePath to "${tmpDir}/" & name of att
        save att in POSIX file savePath
        if output is not "" then set output to output & "\\n"
        set output to output & (name of att) & "\\t" & (file size of att)
    end repeat
    return output
end tell`;

  const result = runAppleScriptHeredoc(script);
  if (result.startsWith("Error:")) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    return err(result.slice(7));
  }

  const lines = result.split("\n").filter(Boolean);
  const output = [];
  for (const line of lines) {
    const [filename, sizeStr] = line.split("\t");
    const tmpFile = path.join(tmpDir, filename);
    const destFile = uniquePath(saveDir, emailId, filename);
    fs.renameSync(tmpFile, destFile);
    output.push(`Saved: ${destFile} (${sizeStr} bytes)`);
  }
  fs.rmSync(tmpDir, { recursive: true, force: true });
  return ok(output.join("\n"));
}

// --- Handler: search_body ---

function handleSearchBody(args) {
  const query = args?.query;
  if (!query) return err("query is required.");
  const limit = Math.min(args?.limit || 500, 1000);
  triggerIndexRefresh();

  const escapedQuery = query.replace(/'/g, "''").replace(/"/g, '""');

  // Try quoted match first, then unquoted
  let results = parseFts(`
    SELECT b.id, b.subject, b.sender, snippet(fts, 2, '>>>', '<<<', '...', 40) as snippet, rank
    FROM fts JOIN bodies b ON b.id = fts.rowid
    WHERE fts MATCH '"${escapedQuery}"'
    ORDER BY rank LIMIT ${limit}
  `);

  if (!results || results.length === 0) {
    results = parseFts(`
      SELECT b.id, b.subject, b.sender, snippet(fts, 2, '>>>', '<<<', '...', 40) as snippet, rank
      FROM fts JOIN bodies b ON b.id = fts.rowid
      WHERE fts MATCH '${escapedQuery}'
      ORDER BY rank LIMIT ${limit}
    `);
  }

  if (!results || results.length === 0) {
    const stats = ftsIndexStats();
    return ok(`No results for "${query}". ${ftsStatusLine(stats)}`);
  }

  // Get dates from Outlook DB
  const ids = results.map(r => r.id).join(",");
  let dateMap = {};
  try {
    const dates = runSqlite(`SELECT Record_RecordID as id, Message_TimeReceived FROM Mail WHERE Record_RecordID IN (${ids})`);
    for (const d of dates) dateMap[d.id] = d.Message_TimeReceived;
  } catch {}

  const lines = results.map(r => {
    const date = dateMap[r.id]
      ? new Date(dateMap[r.id] * 1000).toISOString().slice(0, 10)
      : "";
    const snippet = (r.snippet || "").replace(/\n/g, " ").trim();
    return `ID:${r.id} | ${date} | ${r.sender || ""} | ${r.subject || "(no subject)"}\n  ${snippet}`;
  });

  const stats = ftsIndexStats();
  lines.push(`\n${ftsStatusLine(stats)}`);
  return ok(lines.join("\n"));
}

// --- Handler: index_now ---

function handleIndexNow() {
  const result = syncFtsIndex();
  return ok(`Index sync complete after ${result.batches} batch(es). ${ftsStatusLine(result.stats)}`);
}

// --- Start server ---

async function main() {
  ensureFtsDb();
  startIndexRefreshLoop();
  triggerIndexRefresh();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Outlook MCP server running on stdio");
}

const isDirectRun = process.argv[1] && import.meta.url.endsWith(process.argv[1].replace(/^\//, ""));
if (isDirectRun) {
  main().catch(console.error);
}

export {
  extractEmail,
  parseRecipients,
  recipientLines,
  escapeForAppleScript,
  stripSignature,
  stripQuotedReplies,
  cleanBody,
  markdownToHtml,
  coerceEmailIds,
  setRecipientsScript,
  diagnoseAppleScriptError,
  uniquePath,
  ftsStatusLine,
  ok,
  err,
};
