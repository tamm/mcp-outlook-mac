/**
 * E2E test harness for mcp-outlook-mac.
 *
 * Spawns `node index.js` as a child process and speaks JSON-RPC 2.0 over
 * stdio — exactly how a real MCP client would consume it.  Never trusts the
 * tool's own `ok` response as proof of success; every test that mutates state
 * verifies the outcome with an independent AppleScript or SQLite probe.
 *
 * Usage:
 *   import { startServer, stopServer, callTool, probeAppleScript, probeSqlite } from "./harness.js";
 *
 *   // In your test file:
 *   let server;
 *   before(async () => { server = await startServer(); });
 *   after(async () => { await stopServer(server); });
 */

import { spawn, execSync } from "node:child_process";
import { createInterface } from "node:readline";
import path from "node:path";
import { fileURLToPath } from "node:url";
import os from "node:os";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SERVER_ENTRY = path.resolve(__dirname, "../index.js");

// Outlook SQLite DB path — mirrors DB_PATH in index.js
const DB_PATH = path.join(
  os.homedir(),
  "Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite"
);

/**
 * Start the MCP server process.  Returns a handle object used by all other
 * harness functions.
 *
 * @param {{ timeoutMs?: number }} [options]
 * @returns {Promise<ServerHandle>}
 */
export async function startServer(options = {}) {
  const timeoutMs = options.timeoutMs ?? 30_000;

  const child = spawn(process.execPath, [SERVER_ENTRY], {
    stdio: ["pipe", "pipe", "pipe"],
    env: { ...process.env },
  });

  let nextId = 1;
  const pending = new Map(); // id → { resolve, reject }
  let readyResolve;
  const readyPromise = new Promise((res) => (readyResolve = res));

  const rl = createInterface({ input: child.stdout, crlfDelay: Infinity });

  rl.on("line", (line) => {
    if (!line.trim()) return;
    let msg;
    try {
      msg = JSON.parse(line);
    } catch {
      return; // non-JSON output (e.g. debug logs) — ignore
    }

    // MCP notifications (no id) — currently just mark server as ready
    if (msg.id === undefined) {
      readyResolve(true);
      return;
    }

    const handler = pending.get(msg.id);
    if (!handler) return;
    pending.delete(msg.id);

    if (msg.error) {
      handler.reject(new Error(`JSON-RPC error ${msg.error.code}: ${msg.error.message}`));
    } else {
      handler.resolve(msg.result);
    }
  });

  // Stderr is the server's diagnostic channel — capture for test output on failure
  const stderrLines = [];
  child.stderr.on("data", (chunk) => {
    const text = chunk.toString();
    stderrLines.push(text);
    // Server writes "Outlook MCP server running on stdio" once ready
    if (text.includes("Outlook MCP server running on stdio")) {
      readyResolve(true);
    }
  });

  child.on("error", (err) => {
    // Reject all pending requests on spawn error
    for (const { reject } of pending.values()) reject(err);
    pending.clear();
  });

  child.on("exit", (code) => {
    for (const { reject } of pending.values()) {
      reject(new Error(`Server exited with code ${code}`));
    }
    pending.clear();
  });

  // Wait for server to signal readiness, with a timeout
  await Promise.race([
    readyPromise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`Server did not start within ${timeoutMs}ms`)), timeoutMs)
    ),
  ]);

  // Send the MCP initialize handshake
  await sendRequest(child, pending, nextId++, "initialize", {
    protocolVersion: "2024-11-05",
    capabilities: {},
    clientInfo: { name: "e2e-harness", version: "1.0.0" },
  });

  // Send initialized notification
  sendNotification(child, "notifications/initialized", {});

  const handle = {
    child,
    pending,
    stderrLines,
    get nextId() { return nextId; },
    set nextId(v) { nextId = v; },
  };

  return handle;
}

/**
 * Gracefully stop the server process.
 *
 * @param {ServerHandle} handle
 */
export async function stopServer(handle) {
  handle.child.stdin.end();
  await new Promise((res) => handle.child.on("exit", res));
}

/**
 * Call an MCP tool and return the first text content block.
 *
 * @param {ServerHandle} handle
 * @param {string} toolName
 * @param {object} args
 * @param {{ timeoutMs?: number }} [options]
 * @returns {Promise<string>}
 */
export async function callTool(handle, toolName, args = {}, options = {}) {
  const timeoutMs = options.timeoutMs ?? 30_000;
  const id = handle.nextId++;

  const result = await Promise.race([
    sendRequest(handle.child, handle.pending, id, "tools/call", {
      name: toolName,
      arguments: args,
    }),
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`Tool '${toolName}' timed out after ${timeoutMs}ms`)), timeoutMs)
    ),
  ]);

  const content = result?.content;
  if (!Array.isArray(content) || content.length === 0) {
    throw new Error(`Tool '${toolName}' returned no content`);
  }

  return content[0]?.text ?? "";
}

/**
 * Run an AppleScript via osascript and return the trimmed stdout.
 * This is the independent verification path — never reuse the tool's own
 * AppleScript to confirm what the tool just did.
 *
 * @param {string} script
 * @param {{ timeoutMs?: number }} [options]
 * @returns {string}
 */
export function probeAppleScript(script, options = {}) {
  const timeoutMs = options.timeoutMs ?? 30_000;
  return execSync(`osascript << 'APPLESCRIPT_EOF'\n${script}\nAPPLESCRIPT_EOF`, {
    encoding: "utf-8",
    timeout: timeoutMs,
    shell: "/bin/bash",
  }).trim();
}

/**
 * Run a SQLite query against Outlook's database and return parsed JSON rows.
 *
 * @param {string} query
 * @returns {object[]}
 */
export function probeSqlite(query) {
  const flat = query.replace(/\s+/g, " ").trim();
  const raw = execSync(`sqlite3 -json ${JSON.stringify(DB_PATH)} ${JSON.stringify(flat)}`, {
    encoding: "utf-8",
    timeout: 15_000,
    shell: "/bin/bash",
  }).trim();
  if (!raw) return [];
  return JSON.parse(raw);
}

/**
 * Poll until `predicate()` returns a truthy value, or throw after timeout.
 *
 * Useful for waiting for a self-sent fixture email to land in the Outlook DB
 * before running assertions against it.
 *
 * @param {() => any} predicate  Called repeatedly; first truthy return value is returned.
 * @param {{ intervalMs?: number, timeoutMs?: number, label?: string }} [options]
 * @returns {Promise<any>}
 */
export async function pollUntil(predicate, options = {}) {
  const intervalMs = options.intervalMs ?? 2_000;
  const timeoutMs = options.timeoutMs ?? 120_000;
  const label = options.label ?? "condition";
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    const val = predicate();
    if (val) return val;
    await sleep(intervalMs);
  }
  throw new Error(`Timed out waiting for: ${label}`);
}

/**
 * Seed a fixture email by sending a self-email with a marker subject.
 * Returns the marker subject string used so tests can query for it.
 *
 * @param {string} [tag]  Optional extra tag appended to the marker subject.
 * @returns {{ subject: string, isoTag: string }}
 */
export function fixtureSubject(tag = "") {
  const isoTag = new Date().toISOString().replace(/[:.]/g, "-");
  const subject = `[mcp-e2e ${isoTag}]${tag ? " " + tag : ""}`;
  return { subject, isoTag };
}

/**
 * Poll the Outlook SQLite DB until an email with the given subject marker
 * appears.  Returns the matching row (with Record_RecordID at minimum).
 *
 * @param {string} subject   Exact subject string to look for.
 * @param {{ timeoutMs?: number }} [options]
 * @returns {Promise<object>}
 */
export async function waitForFixture(subject, options = {}) {
  const timeoutMs = options.timeoutMs ?? 120_000;
  const escaped = subject.replace(/'/g, "''");

  return pollUntil(
    () => {
      const rows = probeSqlite(
        `SELECT Record_RecordID, Message_NormalizedSubject, Message_TimeReceived
         FROM Mail
         WHERE Message_NormalizedSubject = '${escaped}'
         ORDER BY Message_TimeReceived DESC
         LIMIT 1`
      );
      return rows.length > 0 ? rows[0] : null;
    },
    { timeoutMs, label: `fixture email "${subject}"` }
  );
}

// --- Internal helpers ---

function sendRequest(child, pending, id, method, params) {
  return new Promise((resolve, reject) => {
    pending.set(id, { resolve, reject });
    const msg = JSON.stringify({ jsonrpc: "2.0", id, method, params });
    child.stdin.write(msg + "\n");
  });
}

function sendNotification(child, method, params) {
  const msg = JSON.stringify({ jsonrpc: "2.0", method, params });
  child.stdin.write(msg + "\n");
}

function sleep(ms) {
  return new Promise((res) => setTimeout(res, ms));
}
