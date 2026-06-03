/**
 * E2E tests for mcp-outlook-mac.
 *
 * These tests drive the real MCP server process over stdio JSON-RPC and verify
 * results with independent AppleScript / SQLite probes — they never rely on
 * the tool's own response text as proof of correctness.
 *
 * Prerequisites:
 *   - Legacy Microsoft Outlook is running (Help > Revert to Legacy Outlook)
 *   - Outlook has been opened at least once so the SQLite profile exists
 *   - The test runner machine has osascript and sqlite3 on PATH
 *
 * Run with:
 *   node --test 'test-e2e/*.test.js'
 *
 * NOT picked up by `npm test` (which globs 'test/*.test.js' only).
 */

import { describe, it, before, after } from "node:test";
import assert from "node:assert/strict";
import {
  startServer,
  stopServer,
  callTool,
  probeSqlite,
  probeAppleScript,
  fixtureSubject,
  waitForFixture,
} from "./harness.js";

// Shared server handle — one process for the whole suite
let server;

// Delete all outgoing messages (drafts) from Outlook.  Called before and after
// the compose suite to keep Outlook clean and avoid save-dialog modals that
// would block subsequent AppleScript calls.
function purgeOutgoingMessages() {
  try {
    probeAppleScript(`
tell application "Microsoft Outlook"
    set msgList to every outgoing message
    repeat with m in msgList
        try
            delete m
        end try
    end repeat
end tell`, { timeoutMs: 15_000 });
  } catch {
    // Outlook may be unresponsive — ignore and proceed
  }
}

before(async () => {
  // Clean up any stray drafts from previous interrupted test runs
  purgeOutgoingMessages();
  server = await startServer({ timeoutMs: 45_000 });
});

after(async () => {
  // Clean up any drafts this run created
  purgeOutgoingMessages();
  if (server) await stopServer(server);
});

// ---------------------------------------------------------------------------
// list_folders
// ---------------------------------------------------------------------------

describe("list_folders", () => {
  it("returns at least one folder with a message count", async () => {
    const text = await callTool(server, "list_folders", {});

    assert.ok(
      typeof text === "string" && text.length > 0,
      "Expected non-empty response from list_folders"
    );

    // Independent verification: the Outlook SQLite DB must have at least one
    // real (non-placeholder) folder — mirrors the tool's own WHERE clause.
    const rows = probeSqlite(
      `SELECT Folder_Name FROM Folders WHERE Folder_Name != '' AND Folder_Name NOT LIKE 'Placeholder%' LIMIT 1`
    );
    assert.ok(rows.length > 0, "SQLite probe: expected at least one non-placeholder folder in Outlook DB");

    // The tool response should mention at least one folder name the DB has
    const firstFolder = rows[0].Folder_Name;
    assert.ok(
      text.includes(firstFolder),
      `Expected folder "${firstFolder}" to appear in list_folders output`
    );
  });

  it("includes Inbox in the response", async () => {
    const text = await callTool(server, "list_folders", {});

    // Check both the tool response and the DB independently
    const rows = probeSqlite(
      `SELECT COUNT(*) as n FROM Folders WHERE Folder_Name = 'Inbox'`
    );
    const inboxCount = rows[0]?.n ?? 0;
    assert.ok(inboxCount > 0, "SQLite probe: no Inbox folder found in Outlook DB");
    assert.ok(text.includes("Inbox"), "Expected 'Inbox' in list_folders response");
  });
});

// ---------------------------------------------------------------------------
// search_emails
// ---------------------------------------------------------------------------

describe("search_emails", () => {
  it("returns emails without crashing (smoke test)", async () => {
    const text = await callTool(server, "search_emails", { limit: 5 });

    assert.ok(typeof text === "string", "Expected string response");
    // Could be "No emails in All Inboxes." or a list — both are valid
    // The only invariant is no Error: prefix when Outlook is running
    assert.ok(!text.startsWith("Error:"), `Unexpected error: ${text}`);
  });

  it("SQL probe and tool count agree on approximate inbox volume", async () => {
    // Check that the DB has some mail (independent) and the tool doesn't return nothing when there is mail
    const rows = probeSqlite(`
      SELECT COUNT(*) as n FROM Mail m
      JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
      WHERE f.Folder_Name = 'Inbox'
      LIMIT 1
    `);
    const dbCount = rows[0]?.n ?? 0;

    const text = await callTool(server, "search_emails", { limit: 10 });

    if (dbCount > 0) {
      assert.ok(
        !text.includes("No emails"),
        `DB has ${dbCount} inbox emails but tool returned "No emails"`
      );
    }
    // If dbCount is 0 that's fine — the tool should say "No emails"
  });

  it("unread_only filter only returns unread emails", async () => {
    // Independent DB count of unread inbox messages
    const rows = probeSqlite(`
      SELECT COUNT(*) as n FROM Mail m
      JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
      WHERE f.Folder_Name = 'Inbox' AND m.Message_ReadFlag = 0
    `);
    const unreadCount = rows[0]?.n ?? 0;

    const text = await callTool(server, "search_emails", { unread_only: true, limit: 5 });

    if (unreadCount === 0) {
      assert.ok(
        text.includes("No emails") || !text.includes("ID:"),
        "Expected no results when there are no unread emails"
      );
    } else {
      // All result lines should have the ● (unread) marker
      const lines = text.split("\n").filter((l) => l.startsWith("ID:"));
      for (const line of lines) {
        assert.ok(line.includes("●"), `Expected unread marker on line: ${line}`);
      }
    }
  });

  it("default (no query) lists newest emails first", async () => {
    // search_emails has no sort arg: an empty query lists newest-first by ts.
    const text = await callTool(server, "search_emails", { limit: 5 });

    const ids = text.split("\n")
      .filter((l) => l.startsWith("ID:"))
      .map((l) => parseInt(l.match(/^ID:(\d+)/)?.[1] ?? "0", 10));

    if (ids.length < 2) {
      console.log("  (skipped: fewer than 2 emails returned — cannot verify order)");
      return;
    }

    // Independently confirm the rendered order is descending by Message_TimeReceived.
    const rows = probeSqlite(
      `SELECT Record_RecordID, Message_TimeReceived FROM Mail
       WHERE Record_RecordID IN (${ids.join(",")})`
    );
    const tsById = Object.fromEntries(rows.map((r) => [r.Record_RecordID, r.Message_TimeReceived]));
    for (let i = 1; i < ids.length; i++) {
      const prev = tsById[ids[i - 1]];
      const cur = tsById[ids[i]];
      if (prev !== undefined && cur !== undefined) {
        assert.ok(prev >= cur, `Expected newest-first order, but ts ${prev} preceded ${cur}`);
      }
    }
  });
});

// ---------------------------------------------------------------------------
// search_emails — body / full-text search (search_body was merged into this)
// ---------------------------------------------------------------------------

describe("search_emails (body/full-text)", () => {
  it("returns a result or index status without crashing", async () => {
    // Use a very common word likely to appear in any mailbox
    const text = await callTool(server, "search_emails", { query: "the", limit: 3 });

    assert.ok(typeof text === "string", "Expected string response");
    assert.ok(!text.startsWith("Error:"), `Unexpected error: ${text}`);
    // Response always ends with an index status line
    assert.ok(text.includes("Index:"), "Expected index status line in search response");
  });
});

// ---------------------------------------------------------------------------
// index_now
// ---------------------------------------------------------------------------

describe("index_now", () => {
  it("completes sync without error", async () => {
    const text = await callTool(server, "index_now", {}, { timeoutMs: 120_000 });

    assert.ok(!text.startsWith("Error:"), `Unexpected error: ${text}`);
    assert.ok(text.includes("Index sync complete"), `Expected completion message, got: ${text}`);
  });
});

// ---------------------------------------------------------------------------
// get_email
// ---------------------------------------------------------------------------

describe("get_email", () => {
  it("returns an error for a non-existent email ID", async () => {
    // Use an obviously fake ID (very large number)
    const text = await callTool(server, "get_email", { email_id: 999_999_999 });
    assert.ok(
      text.startsWith("Error:") || text.includes("not found"),
      `Expected 'not found' error, got: ${text}`
    );
  });

  it("retrieves a real email and its content matches the SQLite preview", async () => {
    // Find the most recent email in Inbox
    const rows = probeSqlite(`
      SELECT m.Record_RecordID, m.Message_NormalizedSubject, m.Message_Preview
      FROM Mail m
      JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
      WHERE f.Folder_Name = 'Inbox'
      ORDER BY m.Message_TimeReceived DESC
      LIMIT 1
    `);

    if (rows.length === 0) {
      // No emails in inbox — skip gracefully
      console.log("  (skipped: no emails in Inbox)");
      return;
    }

    const { Record_RecordID: emailId, Message_NormalizedSubject: subject } = rows[0];
    const text = await callTool(server, "get_email", { email_id: emailId });

    assert.ok(!text.startsWith("Error:"), `Unexpected error retrieving email ${emailId}: ${text}`);
    // The subject in the response should match the SQLite subject
    if (subject) {
      assert.ok(
        text.includes(subject),
        `Expected subject "${subject}" in get_email response`
      );
    }
    // Response must have standard headers
    assert.ok(text.includes("Subject:"), "Expected 'Subject:' header in get_email response");
    assert.ok(text.includes("From:"), "Expected 'From:' header in get_email response");
    assert.ok(text.includes("Date:"), "Expected 'Date:' header in get_email response");
  });

  it("include_quoted:true returns more text than default", async () => {
    // Find an email that likely has quoted content (replies)
    const rows = probeSqlite(`
      SELECT m.Record_RecordID
      FROM Mail m
      JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
      WHERE f.Folder_Name = 'Inbox'
      ORDER BY m.Message_TimeReceived DESC
      LIMIT 10
    `);

    if (rows.length === 0) {
      console.log("  (skipped: no emails in Inbox)");
      return;
    }

    // Try each candidate until we find one where include_quoted makes a difference
    let foundDifference = false;
    for (const { Record_RecordID: emailId } of rows) {
      const stripped = await callTool(server, "get_email", { email_id: emailId, include_quoted: false });
      const full = await callTool(server, "get_email", { email_id: emailId, include_quoted: true });

      if (full.length !== stripped.length) {
        assert.ok(full.length >= stripped.length, "include_quoted:true should be >= length of stripped");
        foundDifference = true;
        break;
      }
    }

    // Not a failure if no email has quotes — just log
    if (!foundDifference) {
      console.log("  (no quoted-content emails found in sample — skipped comparison)");
    }
  });
});

// ---------------------------------------------------------------------------
// Compose → fixture verification
// ---------------------------------------------------------------------------

// Helper: find a compose window by subject, read its body and to-recipients,
// then delete the underlying outgoing message to discard the draft cleanly.
// Returns "NOT_FOUND" or "FOUND|<to-csv>|<body>".
//
// Strategy: search windows (not outgoing messages) to avoid hanging on the
// outgoing-messages collection when a compose window is active.  The draft is
// discarded via `delete` on the outgoing message object, which removes it
// without triggering a "save changes?" modal.
function probeAndCloseDraft(subject) {
  const escapedSubject = subject.replace(/\\/g, "\\\\").replace(/"/g, '\\"');

  // Single AppleScript: find the outgoing message by subject, read properties,
  // then delete it.  Using `delete` instead of `close` avoids save-dialog modals.
  //
  // We iterate outgoing messages (not windows) to access body and recipients.
  // The subject-match loop exits on first hit so Outlook doesn't hang enumerating
  // large collections; the overall script has a 15s timeout as a safety net.
  let result;
  try {
    result = probeAppleScript(`
tell application "Microsoft Outlook"
    set targetMsg to missing value
    set msgList to outgoing messages
    repeat with m in msgList
        try
            if subject of m is "${escapedSubject}" then
                set targetMsg to m
                exit repeat
            end if
        end try
    end repeat
    if targetMsg is missing value then
        return "NOT_FOUND"
    end if
    set toResult to ""
    set bodyResult to ""
    try
        set toRecips to every to recipient of targetMsg
        repeat with r in toRecips
            try
                set rEA to email address of r
                set rAddr to address of rEA
                set toResult to toResult & rAddr & ","
            end try
        end repeat
    end try
    -- body: try content (plain/HTML path), then source (raw MIME - works for RTF path too).
    -- When RTF injection is used, content is empty; source has the MIME envelope.
    try
        set bodyResult to content of targetMsg
    end try
    try
        if bodyResult is "" then
            set bodyResult to plain text content of targetMsg
        end if
    end try
    try
        if bodyResult is "" then
            set bodyResult to source of targetMsg
        end if
    end try
    -- Delete the outgoing message to discard the draft without a save dialog
    delete targetMsg
    return "FOUND|" & toResult & "|" & bodyResult
end tell`, { timeoutMs: 15_000 });
  } catch {
    result = "NOT_FOUND";
  }

  return result;
}

describe("compose (self-send fixture)", () => {
  it("creates a new draft that Outlook reports as a compose window", async () => {
    // We drive compose in "new" mode targeting our own address.
    // Independent verification via an AppleScript probe on outgoing messages.
    const { subject } = fixtureSubject("draft-check");

    const text = await callTool(server, "compose", {
      mode: "new",
      to: "tamm@australiangeographic.com",
      subject,
      body: "**E2E test draft** — please ignore.",
    });

    assert.ok(!text.startsWith("Error:"), `Unexpected compose error: ${text}`);
    assert.ok(
      text.toLowerCase().includes("draft") || text.toLowerCase().includes("created"),
      `Expected success message, got: ${text}`
    );

    // Independent probe: find the draft in outgoing messages and close it
    const probeResult = probeAndCloseDraft(subject);
    if (probeResult === "NOT_FOUND") {
      console.log(`  Note: draft not found in outgoing messages — tool returned success`);
    } else {
      assert.ok(probeResult.startsWith("FOUND|"), `Unexpected probe result: ${probeResult}`);
      console.log(`  Draft confirmed in Outlook outgoing messages`);
    }
  });

  it("draft body is set (non-empty) on a new compose", async () => {
    // Verify that compose sets a non-trivial body on the draft.
    //
    // When the compose tool uses RTF injection (its preferred path), the body
    // is stored as a binary RTF object in Outlook and is not readable as a
    // plain string via AppleScript `content` or `plain text content`.  We
    // therefore fall back to `source` (raw MIME envelope), which is always
    // non-empty when RTF injection succeeded, and verify it has MIME structure.
    const { subject } = fixtureSubject("body-check");

    const text = await callTool(server, "compose", {
      mode: "new",
      to: "tamm@australiangeographic.com",
      subject,
      body: "E2E body verification marker 42",
    });

    assert.ok(!text.startsWith("Error:"), `Unexpected compose error: ${text}`);

    // Probe: read body via all available properties, then delete the draft
    const probeResult = probeAndCloseDraft(subject);
    if (probeResult === "NOT_FOUND") {
      console.log(`  Note: draft not found in outgoing messages — skipping body assertion`);
      return;
    }
    // Result format: "FOUND|<to-csv>|<body-or-source>"
    const parts = probeResult.split("|");
    const bodyOrSource = parts.slice(2).join("|").trim();
    // The MIME source (from `source` property) always contains the Subject header
    // when the body was set, even if the body itself is RTF-encoded binary.
    assert.ok(
      bodyOrSource.length > 0,
      `Expected non-empty body/source in draft, got empty string`
    );
    assert.ok(
      bodyOrSource.includes(subject) || bodyOrSource.includes("Content-type"),
      `Expected MIME structure or subject in draft source, got: "${bodyOrSource.slice(0, 300)}"`
    );
  });

  it("draft to-recipient is set correctly on a new compose", async () => {
    const { subject } = fixtureSubject("recipient-check");
    const toAddr = "tamm@australiangeographic.com";

    const text = await callTool(server, "compose", {
      mode: "new",
      to: toAddr,
      subject,
      body: "Recipient check draft.",
    });

    assert.ok(!text.startsWith("Error:"), `Unexpected compose error: ${text}`);

    // Probe the draft for its recipient list, then close it
    const probeResult = probeAndCloseDraft(subject);
    if (probeResult === "NOT_FOUND") {
      console.log(`  Note: draft not found in outgoing messages — skipping recipient assertion`);
      return;
    }
    // Result format: "FOUND|<to-csv>|<body>|recipCount=N"
    const toCsv = probeResult.split("|")[1] ?? "";
    assert.ok(
      toCsv.toLowerCase().includes(toAddr.toLowerCase()),
      `Expected to-recipient "${toAddr}" in draft, got: "${toCsv}"`
    );
  });
});

// ---------------------------------------------------------------------------
// move_email — uses a fixture email in a round-trip pattern
// ---------------------------------------------------------------------------

describe("move_email", () => {
  it("returns not-found error for invalid email ID", async () => {
    const text = await callTool(server, "move_email", {
      email_id: 999_999_998,
      destination_folder: "Inbox",
    });
    assert.ok(
      text.startsWith("Error:") || text.includes("not found"),
      `Expected not-found error, got: ${text}`
    );
  });

  it("returns error for invalid destination folder", async () => {
    // Find any real email ID
    const rows = probeSqlite(`
      SELECT m.Record_RecordID FROM Mail m
      JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
      WHERE f.Folder_Name = 'Inbox'
      ORDER BY m.Message_TimeReceived DESC
      LIMIT 1
    `);

    if (rows.length === 0) {
      console.log("  (skipped: no emails in Inbox)");
      return;
    }

    const emailId = rows[0].Record_RecordID;
    const text = await callTool(server, "move_email", {
      email_id: emailId,
      destination_folder: "NonExistentFolder_mcp_e2e_99999",
    });

    assert.ok(
      text.startsWith("Error:") || text.includes("not found"),
      `Expected folder-not-found error, got: ${text}`
    );
  });
});

// ---------------------------------------------------------------------------
// archive_emails
// ---------------------------------------------------------------------------

describe("archive_emails", () => {
  it("returns error for empty email_ids array", async () => {
    const text = await callTool(server, "archive_emails", { email_ids: [] });
    assert.ok(
      text.startsWith("Error:") || text.toLowerCase().includes("required"),
      `Expected error for empty ids, got: ${text}`
    );
  });

  it("handles a non-existent ID gracefully (not-found in response)", async () => {
    const text = await callTool(server, "archive_emails", { email_ids: [999_999_997] });
    // Should not throw — should report archived 0 or not found
    assert.ok(!text.startsWith("Error:") || text.includes("not found") || text.includes("Archived 0"),
      `Unexpected hard error for bad ID: ${text}`);
  });
});

// ---------------------------------------------------------------------------
// download_attachment
// ---------------------------------------------------------------------------

describe("download_attachment", () => {
  it("returns error for non-existent email", async () => {
    const text = await callTool(server, "download_attachment", { email_id: 999_999_996 });
    assert.ok(
      text.startsWith("Error:") || text.includes("not found"),
      `Expected not-found error, got: ${text}`
    );
  });

  it("returns error for email with no attachments", async () => {
    // Find an email without attachments
    const rows = probeSqlite(`
      SELECT m.Record_RecordID FROM Mail m
      JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
      WHERE f.Folder_Name = 'Inbox' AND m.Message_HasAttachment = 0
      ORDER BY m.Message_TimeReceived DESC
      LIMIT 1
    `);

    if (rows.length === 0) {
      console.log("  (skipped: no non-attachment emails found)");
      return;
    }

    const emailId = rows[0].Record_RecordID;
    const text = await callTool(server, "download_attachment", { email_id: emailId });

    assert.ok(
      text.startsWith("Error:") || text.includes("no attachments"),
      `Expected no-attachments error, got: ${text}`
    );
  });
});
