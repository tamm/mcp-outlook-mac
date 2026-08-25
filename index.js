#!/usr/bin/env node
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { execSync, execFileSync } from "child_process";
import { realpathSync } from "fs";
import { fileURLToPath } from "url";
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
    // execFileSync passes args directly to sqlite3 with no shell — no quoting/injection risk.
    const result = execFileSync("sqlite3", ["-json", DB_PATH, flat], {
      encoding: "utf-8",
      maxBuffer: 64 * 1024 * 1024,
      timeout: 15000,
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

// Two things fight us on the way to an Outlook draft:
//   1. textutil's HTML importer silently DROPS <br>, so marked's `breaks: true`
//      alone still delivers a one-per-line list glued into one paragraph.
//   2. Outlook does not render the default paragraph spacing textutil emits the
//      way textutil's own txt conversion suggests, so ANY reliance on default
//      <p> margins puts a stray gap in the middle of a run.
// So state every paragraph explicitly and lean on no defaults at all: each source
// line becomes its own margin:0 paragraph, and a markdown paragraph break becomes
// a real empty paragraph. That is what an email body looks like as plain text,
// which is what a mail body is.
const SPACER = '<p style="margin:0">&nbsp;</p>';

function emailParagraphs(html) {
  const out = html.replace(/<p>([\s\S]*?)<\/p>/g, (whole, inner) =>
    inner
      .split(/<br\s*\/?>\s*/)
      .map(line => `<p style="margin:0">${line}</p>`)
      .join("") + SPACER
  );
  // the trailing spacer after the final paragraph is just dead space
  return out.endsWith(SPACER + "\n") ? out.slice(0, -(SPACER.length + 1)) + "\n"
       : out.endsWith(SPACER) ? out.slice(0, -SPACER.length)
       : out;
}

function markdownToHtml(text) {
  const html = emailParagraphs(marked.parse(text, { async: false, breaks: true }));
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

// AppleScript sometimes yields empty `to recipients` / `cc recipients` (always on
// Sent Items). Backfill those header lines from the Mail DB address-list columns.
function fillRecipientsFromDb(text, dbTo, dbCc) {
  // Only ever look at, and touch, the FIRST To:/CC: lines — the real header block.
  // A quoted reply below carries its own "To:"/"CC:" lines, and matching those made
  // the fallback conclude the header was already populated and skip the fill.
  const lines = text.split("\n");
  const to = lines.findIndex(l => l.startsWith("To:"));
  if (to === -1) return text;
  if (dbTo && dbTo.trim() && !lines[to].slice(3).trim()) {
    lines[to] = `To: ${dbTo.trim()}`;
  }
  if (dbCc && dbCc.trim()) {
    const cc = lines[to + 1]?.startsWith("CC:") ? to + 1 : -1;
    if (cc === -1) lines.splice(to + 1, 0, `CC: ${dbCc.trim()}`);
    else if (!lines[cc].slice(3).trim()) lines[cc] = `CC: ${dbCc.trim()}`;
  }
  return lines.join("\n");
}


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

// --- Query helpers ---

// Split a multi-word query into individual terms (strips punctuation, dedupes, ignores short words)
function queryTerms(query) {
  return [...new Set(
    query.toLowerCase().split(/[\s,;|&()"']+/).filter(t => t.length > 2)
  )];
}

// For each term, also search its alphanumeric core (leading/trailing punctuation stripped).
// This is what makes 'tamm@' fuzzy-match '@tamm': the literal 'tamm@' matches only exact substrings,
// while the core 'tamm' surfaces near-matches like '@tamm'. Exact substrings still rank higher
// (via exact_hit), so cores only add lower-ranked fuzzy candidates. Cores ≤2 chars are dropped.
function expandTerms(terms) {
  const out = [];
  for (const t of terms) {
    out.push(t);
    const core = t.replace(/^[^a-z0-9]+/i, "").replace(/[^a-z0-9]+$/i, "");
    if (core && core !== t && core.length > 2) out.push(core);
  }
  return [...new Set(out)];
}

// --- FTS5 body search index ---

// FTS index location. Override with MCP_OUTLOOK_FTS_DB (used by tests + isolated CLI runs).
const FTS_DB = process.env.MCP_OUTLOOK_FTS_DB || path.join(os.homedir(), ".mcp-outlook-mac", "body-index.db");
const FTS_DIR = path.dirname(FTS_DB);
const INDEX_BATCH_SIZE = 10000;
const INDEX_REFRESH_INTERVAL_MS = 60 * 1000;
let indexRefreshRunning = false;
let indexRefreshTimer = null;

function runFts(query, options = {}) {
  const q = query.replace(/\s+/g, " ").trim();
  try {
    // execFileSync: no shell, query passed as a literal arg — safe for arbitrary user text.
    return execFileSync("sqlite3", [FTS_DB, q], {
      encoding: "utf-8",
      maxBuffer: 64 * 1024 * 1024,
      timeout: options.timeout || 10000,
    }).trim();
  } catch (error) {
    throw new Error(`FTS error: ${(error.message || "").slice(0, 200)}`);
  }
}

function parseFts(query) {
  const q = query.replace(/\s+/g, " ").trim();
  try {
    const raw = execFileSync("sqlite3", ["-json", FTS_DB, q], {
      encoding: "utf-8", maxBuffer: 64 * 1024 * 1024, timeout: 10000,
    }).trim();
    if (!raw) return [];
    return JSON.parse(raw);
  } catch { return []; }
}

// Bump this whenever the FTS/bodies schema changes — forces a one-time rebuild.
const FTS_SCHEMA_VERSION = 2;

// All FTS-indexed text columns, in order — used for bm25 weights and trigger column lists.
const FTS_COLUMNS = ["subject", "sender_name", "sender_addr", "to_addr", "cc_addr", "body"];

let ftsReady = false;

function ensureFtsDb() {
  if (ftsReady) return;
  if (!fs.existsSync(FTS_DIR)) fs.mkdirSync(FTS_DIR, { recursive: true });

  const current = Number(runFts(`PRAGMA user_version`)) || 0;
  if (current < FTS_SCHEMA_VERSION) {
    migrateFtsDb();
  } else {
    createFtsSchema();
  }
  ftsReady = true;
}

function createFtsSchema() {
  // indexed: tracks which Outlook message IDs have been synced (and at what ts).
  runFts(`CREATE TABLE IF NOT EXISTS indexed (id INTEGER PRIMARY KEY, ts INTEGER)`);
  // bodies: the searchable content + stored metadata for rendering/filtering without an Outlook round-trip.
  runFts(`CREATE TABLE IF NOT EXISTS bodies (
    id INTEGER PRIMARY KEY,
    subject TEXT, sender_name TEXT, sender_addr TEXT,
    to_addr TEXT, cc_addr TEXT, body TEXT,
    folder TEXT, account TEXT, ts INTEGER, read_flag INTEGER, has_attach INTEGER
  )`);
  // External-content trigram FTS over every searchable text column.
  runFts(`CREATE VIRTUAL TABLE IF NOT EXISTS fts USING fts5(
    ${FTS_COLUMNS.join(", ")},
    content='bodies', content_rowid='id', tokenize='trigram'
  )`);
  const cols = FTS_COLUMNS.join(", ");
  const newCols = FTS_COLUMNS.map(c => `new.${c}`).join(", ");
  const oldCols = FTS_COLUMNS.map(c => `old.${c}`).join(", ");
  runFts(`CREATE TRIGGER IF NOT EXISTS bodies_ai AFTER INSERT ON bodies BEGIN
    INSERT INTO fts(rowid, ${cols}) VALUES (new.id, ${newCols}); END`);
  runFts(`CREATE TRIGGER IF NOT EXISTS bodies_ad AFTER DELETE ON bodies BEGIN
    INSERT INTO fts(fts, rowid, ${cols}) VALUES ('delete', old.id, ${oldCols}); END`);
  // External-content FTS5 silently desyncs on a plain UPDATE without this trigger.
  runFts(`CREATE TRIGGER IF NOT EXISTS bodies_au AFTER UPDATE ON bodies BEGIN
    INSERT INTO fts(fts, rowid, ${cols}) VALUES ('delete', old.id, ${oldCols});
    INSERT INTO fts(rowid, ${cols}) VALUES (new.id, ${newCols}); END`);
}

// Drop everything and rebuild cleanly. user_version is set LAST so a crash mid-migration re-runs.
function migrateFtsDb() {
  runFts(`DROP TRIGGER IF EXISTS bodies_ai`);
  runFts(`DROP TRIGGER IF EXISTS bodies_ad`);
  runFts(`DROP TRIGGER IF EXISTS bodies_au`);
  runFts(`DROP TABLE IF EXISTS fts`);   // virtual table before its content table
  runFts(`DROP TABLE IF EXISTS bodies`);
  runFts(`CREATE TABLE IF NOT EXISTS indexed (id INTEGER PRIMARY KEY, ts INTEGER)`);
  runFts(`DELETE FROM indexed`);        // force a full re-sync into the new schema
  createFtsSchema();
  runFts(`PRAGMA user_version = ${FTS_SCHEMA_VERSION}`);
}

function isIndexed(emailId) {
  return parseFts(`SELECT 1 FROM indexed WHERE id = ${Number(emailId)} LIMIT 1`).length > 0;
}

// Single-quote escape for embedding in an FTS SQL string literal.
function sqlLit(s) {
  return (s || "").replace(/'/g, "''");
}

// Called from get_email: upgrade an existing row's body to FULL text (bulk sync only has the preview).
// Safe to UPDATE because bodies_au keeps the FTS index in sync. Creates the row if it doesn't exist yet.
function indexEmail(emailId, subject, sender, bodyText) {
  const id = Number(emailId);
  if (!Number.isFinite(id)) return;
  const body = sqlLit((bodyText || "").slice(0, 50000));
  const subj = sqlLit(subject);
  const from = sqlLit(sender);
  try {
    const exists = parseFts(`SELECT 1 FROM bodies WHERE id = ${id} LIMIT 1`).length > 0;
    if (exists) {
      // Upgrade body to full text; refresh subject/sender if we have them.
      runFts(`UPDATE bodies SET body = '${body}'
              ${subject ? `, subject = '${subj}'` : ""}
              ${sender ? `, sender_name = '${from}'` : ""}
              WHERE id = ${id}`);
    } else {
      runFts(`INSERT OR REPLACE INTO bodies (id, subject, sender_name, body) VALUES (${id}, '${subj}', '${from}', '${body}')`);
    }
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
        COALESCE(m.Message_SenderList, '') AS sender_name,
        COALESCE(m.Message_SenderAddressList, '') AS sender_addr,
        COALESCE(m.Message_ToRecipientAddressList, '') AS to_addr,
        COALESCE(m.Message_CCRecipientAddressList, '') AS cc_addr,
        COALESCE(m.Message_Preview, '') AS body,
        COALESCE(f.Folder_Name, '') AS folder,
        COALESCE(ae.Account_Name, am.Account_Name, 'Local') AS account,
        m.Message_TimeReceived AS ts,
        COALESCE(m.Message_ReadFlag, 0) AS read_flag,
        COALESCE(m.Message_HasAttachment, 0) AS has_attach
      FROM outlook.Mail m
      LEFT JOIN outlook.Folders f ON f.Record_RecordID = m.Record_FolderID
      LEFT JOIN outlook.AccountsExchange ae ON ae.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
      LEFT JOIN outlook.AccountsMail am ON am.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
      LEFT JOIN indexed i ON i.id = m.Record_RecordID
      LEFT JOIN bodies b ON b.id = m.Record_RecordID
      WHERE i.id IS NULL OR b.id IS NULL
      ORDER BY m.Record_RecordID
      LIMIT ${batchSize};
      BEGIN;
      INSERT OR IGNORE INTO bodies (id, subject, sender_name, sender_addr, to_addr, cc_addr, body, folder, account, ts, read_flag, has_attach)
      SELECT id, subject, sender_name, sender_addr, to_addr, cc_addr, body, folder, account, ts, read_flag, has_attach FROM batch_rows;
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
  // Indexed count comes from the local FTS db (always available).
  let indexed = 0;
  try { indexed = parseFts(`SELECT COUNT(*) as n FROM bodies`)[0]?.n || 0; } catch {}
  // Total comes from Outlook — best effort; 0 if Outlook is unreachable.
  let total = 0;
  try { total = runSqlite(`SELECT COUNT(*) as n FROM Mail`)[0]?.n || 0; } catch {}
  return { indexed, total };
}

function ftsStatusLine(stats) {
  const { indexed, total } = stats;
  if (!indexed && !total) return "Index: empty";
  if (!total) return `Index: ${indexed.toLocaleString()} emails indexed`;
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
    description: "Search email. This is the ONLY search tool — it searches EVERYTHING for each message: subject, sender name, sender address, To, CC, and body — across ALL folders including Sent, and all accounts, by default. Just pass a natural query; partial words, substrings and fuzzy matches are found and ranked automatically, with exact matches first. e.g. query 'tamm@' returns every email to/from/mentioning that address (and surfaces '@tamm'-style near-matches, ranked lower). Multi-word queries are decomposed and ranked by how many parts match, so you never get a bare zero when only part matches. Do NOT pass folder/limit/date filters unless you specifically must narrow — the defaults are right for almost every task.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "What to find — matched against subject, sender, To, CC and body. Substrings/fuzzy/multi-word all handled. Omit to list newest emails." },
        folder: { type: "string", description: "Restrict to one folder. Accepts a bare name, a shorthand ('Sent' matches 'Sent Items'), or the 'Account/Folder' string list_folders prints. OMIT to search all folders — almost always correct." },
        account: { type: "string", description: "Restrict to one account. OMIT to search all accounts." },
        limit: { type: "number", description: "Max results (default 500, cap 1000). OMIT — the default is right for most tasks." },
        unread_only: { type: "boolean", description: "Only unread emails. OMIT unless you specifically want unread-only." },
        after: { type: "string", description: "Date lower bound (YYYY-MM-DD). OMIT unless a date range is specifically required." },
        before: { type: "string", description: "Date upper bound (YYYY-MM-DD). OMIT unless a date range is specifically required." },
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
    name: "index_now",
    description: "Build or catch up the search index immediately (subject/sender/recipients/preview for all mail; full body for opened emails). search_emails refreshes the index in the background, so you rarely need this — use it after a large mailbox change or if search reports the index is incomplete.",
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
      case "search_emails": result = handleSearch(args); break;
      case "get_email": result = handleGetEmail(args); break;
      case "compose": result = handleCompose(args); break;
      case "move_email": result = handleMoveEmail(args); break;
      case "archive_emails": result = handleArchiveEmails(args); break;
      case "download_attachment": result = handleDownloadAttachment(args); break;
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

// --- Unified search (one tool, every field, trigram fuzzy + exact ranking) ---

// bm25 column weights, in FTS_COLUMNS order: subject, sender_name, sender_addr, to_addr, cc_addr, body.
// Sender/recipients/subject are weighted above body (which is preview-only in bulk).
const BM25_WEIGHTS = [8, 10, 10, 9, 9, 2];
const MAX_QUERY_TERMS = 6; // cap the per-term UNION to keep the query bounded

// Turn a raw term into an FTS5 string token: double internal quotes, wrap in quotes,
// then escape for the surrounding SQL string literal. Disables FTS operators (OR/AND/*/parens).
function ftsToken(term) {
  const ftsQuoted = `"${term.replace(/"/g, '""')}"`;
  return sqlLit(ftsQuoted);
}

// LIKE pattern for exact-substring detection, escaped for a SQL string literal.
// Escapes LIKE wildcards so a literal % or _ in the query stays literal (ESCAPE '\').
function likePattern(query) {
  const esc = query.replace(/[%_\\]/g, c => "\\" + c);
  return sqlLit(`%${esc}%`);
}

// Build WHERE fragments for the optional filters, operating on stored bodies columns.
function buildFilterClause(args) {
  const parts = [];
  if (args?.folder) {
    // Match tolerantly: list_folders prints "Account/Folder", the index stores the
    // bare folder name, and callers reasonably say "Sent" for "Sent Items". Take the
    // last path segment and prefix-match case-insensitively.
    const name = String(args.folder).split("/").pop().trim().toLowerCase();
    const esc = sqlLit(name.replace(/[%_\\]/g, c => "\\" + c));
    parts.push(`LOWER(b.folder) LIKE '${esc}%' ESCAPE '\\'`);
  }
  if (args?.account) parts.push(`b.account = '${sqlLit(args.account)}'`);
  if (args?.unread_only) parts.push(`b.read_flag = 0`);
  if (args?.after) {
    // Parse as start of the LOCAL day (no Z) so date bounds match what the user sees.
    const ts = Math.floor(new Date(args.after + "T00:00:00").getTime() / 1000);
    if (!isNaN(ts)) parts.push(`b.ts >= ${ts}`);
  }
  if (args?.before) {
    // Parse as end of the LOCAL day so `before` is inclusive of that whole day.
    const ts = Math.floor(new Date(args.before + "T23:59:59").getTime() / 1000);
    if (!isNaN(ts)) parts.push(`b.ts <= ${ts}`);
  }
  return parts.length ? "AND " + parts.join(" AND ") : "";
}

const TEXT_COLS = ["subject", "sender_name", "sender_addr", "to_addr", "cc_addr", "body"];

// exact_hit: 1 if the full query appears as a literal substring in any column.
function exactHitExpr(query) {
  const pat = likePattern(query);
  return "(" + TEXT_COLS.map(c => `b.${c} LIKE '${pat}' ESCAPE '\\'`).join(" OR ") + ")";
}

function rowSelectCols() {
  // Pull everything needed to render + rank, no Outlook round-trip.
  return `b.id, b.subject, b.sender_name, b.sender_addr, b.folder, b.account,
          b.ts, b.read_flag, b.has_attach`;
}

// Render a Unix epoch (seconds) as a YYYY-MM-DD string in the machine's LOCAL
// timezone. toISOString() would force UTC and shift the date across midnight.
function localDate(epochSeconds) {
  const d = new Date(epochSeconds * 1000);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function formatSearchRows(rows) {
  return rows.map(r => {
    const date = r.ts ? localDate(r.ts) : "";
    const from = r.sender_name || r.sender_addr || "";
    const account = r.account ? `${r.account}/` : "";
    const unread = r.read_flag ? " " : "●";
    const attach = r.has_attach ? " 📎" : "";
    const snippet = (r.snippet || "").replace(/\s+/g, " ").trim();
    const tail = snippet ? `\n  ${snippet}` : "";
    return `ID:${r.id} | ${unread} ${date} | ${from} | ${r.subject || "(no subject)"}${attach} [${account}${r.folder || ""}]${tail}`;
  });
}

function handleSearch(args) {
  ensureFtsDb();
  const query = (args?.query || "").trim();
  const limit = Math.min(Math.max(Number(args?.limit) || 500, 1), 1000);
  const filterClause = buildFilterClause(args);

  // No query → list newest emails (optionally filtered).
  if (!query) {
    const rows = parseFts(`
      SELECT ${rowSelectCols()}, '' AS snippet FROM bodies b
      WHERE 1=1 ${filterClause}
      ORDER BY b.ts DESC LIMIT ${limit}`);
    if (!rows.length) return ok(`No emails. ${ftsStatusLine(ftsIndexStats())}`);
    return ok(formatSearchRows(rows).join("\n") + `\n\n${ftsStatusLine(ftsIndexStats())}`);
  }

  const exact = exactHitExpr(query);
  const terms = expandTerms(queryTerms(query)).slice(0, MAX_QUERY_TERMS);

  let rows;
  if (terms.length === 0) {
    // Query too short for trigram (all tokens ≤2 chars) — LIKE-only path.
    rows = parseFts(`
      SELECT ${rowSelectCols()}, ${exact} AS exact_hit, '' AS snippet
      FROM bodies b
      WHERE ${exact} ${filterClause}
      ORDER BY b.ts DESC
      LIMIT ${limit}`);
  } else {
    // Per-term UNION ALL → GROUP BY: terms_matched + best bm25, then tier by exact_hit.
    const perTerm = terms.map(t =>
      `SELECT rowid AS id, bm25(fts, ${BM25_WEIGHTS.join(", ")}) AS bm FROM fts WHERE fts MATCH '${ftsToken(t)}'`
    ).join("\n      UNION ALL\n      ");

    // Snippet via a correlated subquery so it NEVER filters the result set (a JOIN on one term
    // would drop rows that matched only the other terms). Matches the body column on the first term.
    const snipSub = `(SELECT snippet(fts, 5, '>>>', '<<<', '...', 12) FROM fts
                      WHERE fts.rowid = b.id AND fts MATCH '${ftsToken(terms[0])}')`;

    // MATERIALIZED is required: bm25() is an FTS5 auxiliary function that must be evaluated in the
    // same query as its MATCH. Without it, SQLite inlines per_term into the aggregating query and
    // errors with "unable to use function bm25 in the requested context".
    rows = parseFts(`
      WITH per_term AS MATERIALIZED (
        ${perTerm}
      ),
      agg AS (
        SELECT id, COUNT(*) AS terms_matched, MIN(bm) AS best_bm
        FROM per_term GROUP BY id
      )
      SELECT ${rowSelectCols()},
             ${exact} AS exact_hit,
             a.terms_matched,
             COALESCE(${snipSub}, '') AS snippet
      FROM agg a
      JOIN bodies b ON b.id = a.id
      WHERE 1=1 ${filterClause}
      ORDER BY exact_hit DESC, a.terms_matched DESC, a.best_bm ASC, b.ts DESC
      LIMIT ${limit}`);
  }

  if (!rows || rows.length === 0) {
    return ok(`No results for "${query}". ${ftsStatusLine(ftsIndexStats())}`);
  }

  const allFuzzy = rows.every(r => Number(r.exact_hit) === 0);
  const lines = formatSearchRows(rows);
  if (allFuzzy) lines.push(`\n(No exact match for "${query}" — showing best fuzzy/partial matches, ranked.)`);
  lines.push(`\n${ftsStatusLine(ftsIndexStats())}`);
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

  // Supplement To/CC from the DB when AppleScript returned none.
  // Outlook's AppleScript recipient lists come back empty for some messages —
  // reliably so on Sent Items — while the DB columns are always populated.
  let resultWithCc = result;
  const headerEnd = result.indexOf("\n\n");
  const headerBlock = headerEnd === -1 ? result : result.slice(0, headerEnd);
  if (!/^To: .+/m.test(headerBlock) || !/^CC: .+/m.test(headerBlock)) {
    try {
      const rows = runSqlite(
        `SELECT Message_ToRecipientAddressList, Message_CCRecipientAddressList
         FROM Mail WHERE Record_RecordID = ${Number(emailId)} LIMIT 1`
      );
      resultWithCc = fillRecipientsFromDb(
        result,
        rows?.[0]?.Message_ToRecipientAddressList,
        rows?.[0]?.Message_CCRecipientAddressList
      );
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
  {
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
      // Outlook refuses «class RTF » on message objects here, so fall back to the
      // HTML content property — the same thing composeNew falls back to, and the
      // reason `new` has always worked. Prepending rather than replacing keeps the
      // quoted original thread intact.
      //
      // This used to paste via the clipboard into the FRONT window, which put the
      // body into whatever window happened to be frontmost — or nowhere — while
      // still reporting success.
      console.error(`[compose] RTF injection failed for draft ${composeId}, using HTML content property: ${rtfErr.message}`);
      const escapedHtml = escapeForAppleScript(htmlBody);
      const fallbackScript = `
tell application "Microsoft Outlook"
    set composeMsg to message id ${composeId}
    set content of composeMsg to ("${escapedHtml}" & content of composeMsg)${recipientFix}
end tell`;
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

// --- Handler: index_now ---

function handleIndexNow() {
  const result = syncFtsIndex();
  return ok(`Index sync complete after ${result.batches} batch(es). ${ftsStatusLine(result.stats)}`);
}

// --- Entrypoints: MCP server (default) + CLI ---

async function main() {
  ensureFtsDb();
  startIndexRefreshLoop();
  triggerIndexRefresh();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Outlook MCP server running on stdio");
}

const CLI_USAGE = `mcp-outlook — Outlook for Mac search & mail CLI

Usage:
  mcp-outlook search <query> [--folder F] [--account A] [--after YYYY-MM-DD] [--before YYYY-MM-DD] [--limit N] [--unread]
  mcp-outlook index                 Build/catch up the full search index
  mcp-outlook folders               List folders with message counts
  mcp-outlook get <id> [--quoted]   Show full email content by ID
  mcp-outlook serve                 Run as an MCP stdio server (default when launched by a client)

Search hits subject, sender, recipients (To/CC) and body across ALL folders incl. Sent.
Substrings and fuzzy matches are found and ranked automatically (exact matches first).`;

// Minimal arg parser: positional values + --flag / --flag value.
function parseCliArgs(argv) {
  const positional = [];
  const flags = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith("--")) {
      const key = a.slice(2);
      const next = argv[i + 1];
      if (next === undefined || next.startsWith("--")) {
        flags[key] = true;            // boolean flag
      } else {
        flags[key] = next; i++;       // flag with value
      }
    } else {
      positional.push(a);
    }
  }
  return { positional, flags };
}

function printResult(result) {
  const text = result?.content?.[0]?.text ?? "";
  process.stdout.write(text + "\n");
}

function cliMain(argv) {
  const command = argv[0];
  const { positional, flags } = parseCliArgs(argv.slice(1));
  ensureFtsDb();

  switch (command) {
    case "search": {
      // Catch up one batch synchronously so a freshly-built index isn't empty.
      try { syncFtsIndex({ maxBatches: 1 }); } catch {}
      const query = positional.join(" ");
      printResult(handleSearch({
        query,
        folder: flags.folder,
        account: flags.account,
        after: flags.after,
        before: flags.before,
        unread_only: flags.unread === true,
        limit: flags.limit ? Number(flags.limit) : undefined,
      }));
      break;
    }
    case "index":
      printResult(handleIndexNow());
      break;
    case "folders":
      printResult(handleListFolders());
      break;
    case "get": {
      const id = Number(positional[0]);
      if (!Number.isFinite(id)) { process.stderr.write("get: numeric email id required\n"); process.exit(2); }
      printResult(handleGetEmail({ email_id: id, include_quoted: flags.quoted === true }));
      break;
    }
    default:
      process.stderr.write(CLI_USAGE + "\n");
      process.exit(command ? 2 : 0);
  }
  process.exit(0);
}

// realpathSync resolves symlinks so a globally-linked bin still matches the real file.
let isDirectRun = false;
try {
  isDirectRun = process.argv[1] &&
    realpathSync(process.argv[1]) === realpathSync(fileURLToPath(import.meta.url));
} catch { isDirectRun = false; }

if (isDirectRun) {
  const cmd = process.argv[2];
  const KNOWN = new Set(["search", "index", "folders", "get", "serve"]);
  if (cmd === "serve" || (!cmd && !process.stdout.isTTY)) {
    // Explicit serve, or launched by an MCP client (stdio, not a terminal).
    main().catch(console.error);
  } else if (cmd && KNOWN.has(cmd)) {
    cliMain(process.argv.slice(2));
  } else {
    // Bare TTY invocation or unknown subcommand → usage.
    process.stderr.write(CLI_USAGE + "\n");
    process.exit(cmd ? 2 : 0);
  }
}

export {
  extractEmail,
  parseRecipients,
  recipientLines,
  escapeForAppleScript,
  stripSignature,
  stripQuotedReplies,
  cleanBody,
  fillRecipientsFromDb,
  emailParagraphs,
  markdownToHtml,
  coerceEmailIds,
  setRecipientsScript,
  diagnoseAppleScriptError,
  uniquePath,
  ftsStatusLine,
  ok,
  err,
  // search internals (for tests)
  queryTerms,
  expandTerms,
  ftsToken,
  likePattern,
  sqlLit,
  ensureFtsDb,
  createFtsSchema,
  handleSearch,
  parseCliArgs,
  FTS_COLUMNS,
  BM25_WEIGHTS,
};
