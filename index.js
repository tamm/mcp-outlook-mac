import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { execSync } from "child_process";
import fs from "fs";
import os from "os";
import path from "path";

const server = new Server(
  {
    name: "outlook-mcp",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

const LEGACY_OUTLOOK_HINT =
  "If Exchange accounts show 0 results, you may be running New Outlook. " +
  "Revert to Legacy Outlook via Help > Revert to Legacy Outlook for full AppleScript support.";

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function diagnoseAppleScriptError(error) {
  const msg = error.message || "";
  if (msg.includes("ETIMEDOUT") || error.killed) {
    return "AppleScript timed out (>30s). Legacy Outlook may be unresponsive — try quitting and reopening it. Share this with the developer if it persists.";
  }
  if (msg.includes("-1728")) {
    return "Outlook object not found (-1728). The email may have been deleted or moved. Try read_emails to get fresh IDs.";
  }
  if (msg.includes("-1712")) {
    return "Outlook is busy with a dialog or modal window (-1712). Dismiss any open dialogs in Outlook and retry.";
  }
  if (msg.includes("not running") || msg.includes("-600")) {
    return "Microsoft Outlook is not running. Open Legacy Outlook (Help > Revert to Legacy Outlook) and retry.";
  }
  if (msg.includes("-10810") || msg.includes("launch")) {
    return "Could not launch Outlook. Check that Legacy Outlook is installed and not blocked by macOS.";
  }
  // Extract the core osascript error if present
  const execMatch = msg.match(/execution error: (.+?) \(-?\d+\)/);
  if (execMatch) {
    return `Outlook AppleScript error: ${execMatch[1]}. Share this with the developer if unexpected.`;
  }
  return `AppleScript error: ${msg.slice(0, 300)}. Share this with the developer to diagnose.`;
}

function runAppleScript(script) {
  try {
    const result = execSync(`osascript -e '${script.replace(/'/g, "'\"'\"'")}'`, {
      encoding: "utf-8",
      maxBuffer: 10 * 1024 * 1024,
      timeout: 30000,
    });
    return result.trim();
  } catch (error) {
    throw new Error(diagnoseAppleScriptError(error));
  }
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
      return runAppleScriptHeredocRetry(script);
    }
    throw new Error(diagnoseAppleScriptError(error));
  }
}

function runAppleScriptHeredocRetry(script) {
  execSync("sleep 2", { shell: "/bin/bash" });
  return runAppleScriptHeredoc(script, true);
}

// Looks up the account name and folder name for an email ID from SQLite.
function lookupEmailLocation(emailId) {
  try {
    const rows = runSqlite(
      `SELECT f.Folder_Name, COALESCE(ae.Account_Name, am.Account_Name) AS AccountName FROM Mail m JOIN Folders f ON f.Record_RecordID = m.Record_FolderID LEFT JOIN AccountsExchange ae ON ae.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF) LEFT JOIN AccountsMail am ON am.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF) WHERE m.Record_RecordID = ${emailId} LIMIT 1`
    );
    if (rows && rows.length > 0) {
      return { folderName: rows[0].Folder_Name, accountName: rows[0].AccountName };
    }
  } catch (e) {
    // ignore, caller will fall back
  }
  return null;
}

// Generates AppleScript to find a message by ID.
// Uses direct `message id N` reference — no iteration needed.
function findMessageScript(emailId) {
  return `
    set targetMsg to missing value
    try
        set targetMsg to message id ${emailId}
    end try`;
}

function markdownToHtml(text) {
  // Escape HTML entities first
  let html = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  // Bold: **text** or __text__
  html = html.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  html = html.replace(/__(.+?)__/g, "<strong>$1</strong>");

  // Italic: *text* or _text_
  html = html.replace(/\*(.+?)\*/g, "<em>$1</em>");
  html = html.replace(/(?<!\w)_(.+?)_(?!\w)/g, "<em>$1</em>");

  // Inline code: `code`
  html = html.replace(/`(.+?)`/g, '<code style="background:#f0f0f0;padding:2px 4px;border-radius:3px;">$1</code>');

  // Links: [text](url)
  html = html.replace(/\[(.+?)\]\((.+?)\)/g, '<a href="$2">$1</a>');

  // Headings: # ## ###
  html = html.replace(/^### (.+)$/gm, "<h3>$1</h3>");
  html = html.replace(/^## (.+)$/gm, "<h2>$1</h2>");
  html = html.replace(/^# (.+)$/gm, "<h1>$1</h1>");

  // Horizontal rules: --- or ***
  html = html.replace(/^[\-\*]{3,}$/gm, "<hr>");

  // Unordered lists: lines starting with - or *
  html = html.replace(/^[\-\*] (.+)$/gm, "<li>$1</li>");

  // Ordered lists: lines starting with 1. 2. etc
  html = html.replace(/^\d+\. (.+)$/gm, "<li>$1</li>");

  // Wrap consecutive <li> lines in <ul>, consuming newlines between them
  html = html.replace(/((?:<li>.*<\/li>(?:\n|$))+)/g, (match) => {
    const items = match.trim().replace(/\n/g, "");
    return "<ul>" + items + "</ul>";
  });

  // Paragraphs: double newlines become paragraph breaks
  html = html.replace(/\n\n+/g, "</p><p>");

  // Single newlines become <br>
  html = html.replace(/\n/g, "<br>");

  // Wrap in paragraph tags
  html = "<p>" + html + "</p>";

  // Clean up empty paragraphs and stray tags
  html = html.replace(/<p><\/p>/g, "");
  html = html.replace(/<p>(<h[1-3]>)/g, "$1");
  html = html.replace(/(<\/h[1-3]>)<\/p>/g, "$1");
  html = html.replace(/<p>(<ul>)/g, "$1");
  html = html.replace(/(<\/ul>)<\/p>/g, "$1");
  html = html.replace(/<p><hr><\/p>/g, "<hr>");

  return html;
}

// Build a unique file path: {dir}/{emailId}_{name}, appending _N if it already exists.
// If the base already ends in _N, increment that number instead of double-suffixing.
function uniquePath(dir, emailId, filename) {
  const ext = path.extname(filename);
  const base = path.basename(filename, ext);
  const prefix = `${emailId}_${base}`;
  let candidate = path.join(dir, `${prefix}${ext}`);
  if (!fs.existsSync(candidate)) return candidate;

  // Check if prefix already ends with _N
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

const DB_PATH = path.join(
  os.homedir(),
  "Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite"
);

function runSqlite(query) {
  try {
    const flat = query.replace(/\s+/g, " ").trim();
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
    if (msg.includes("database is locked")) {
      throw new Error("Outlook database is locked — Outlook may be syncing or another process has it open. Wait a moment and retry. Share this with the developer if it persists.");
    }
    if (msg.includes("no such table") || msg.includes("no such column")) {
      throw new Error(`Outlook database schema mismatch — the database structure may have changed. Share this with the developer: ${msg.slice(0, 200)}`);
    }
    if (msg.includes("unable to open database") || msg.includes("not a database")) {
      throw new Error("Cannot open Outlook database — the file may be missing or corrupt. Ensure Legacy Outlook is installed and has been opened at least once. Share this with the developer if unexpected.");
    }
    if (msg.includes("ETIMEDOUT") || error.killed) {
      throw new Error("SQLite query timed out (>15s). The database may be very large or locked. Share this with the developer if it persists.");
    }
    throw new Error(`Outlook database error: ${msg.slice(0, 300)}. Share this with the developer to diagnose.`);
  }
}

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: "list_folders",
        description: "List all mail folders in Outlook. Requires Legacy Outlook for Exchange access (Help > Revert to Legacy Outlook).",
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "read_emails",
        description: "Read emails from Outlook. Defaults to all Inboxes across all accounts. Returns ID, folder, subject, sender, date, and preview. Emails with attachments are marked with 📎.",
        inputSchema: {
          type: "object",
          properties: {
            folder: {
              type: "string",
              description: "Folder name to read from. Defaults to all Inboxes across all accounts. Use list_folders to see available folders.",
            },
            limit: {
              type: "number",
              description: "Maximum number of emails to return (default: 10, max: 50)",
            },
          },
        },
      },
      {
        name: "get_email",
        description: "Get the full content of a specific email by its ID. Lists any attachments with filenames and sizes. Use download_attachment to save them to disk.",
        inputSchema: {
          type: "object",
          properties: {
            email_id: {
              type: "number",
              description: "The ID of the email to retrieve",
            },
          },
          required: ["email_id"],
        },
      },
      {
        name: "create_draft",
        description: "Create a new draft email in Outlook. Opens the draft in a compose window. Body supports markdown formatting (bold, italic, lists, headings, links).",
        inputSchema: {
          type: "object",
          properties: {
            to: {
              type: "string",
              description: "Recipient email address",
            },
            subject: {
              type: "string",
              description: "Email subject",
            },
            body: {
              type: "string",
              description: "Email body content",
            },
            cc: {
              type: "string",
              description: "CC recipient email address (optional)",
            },
          },
          required: ["subject", "body"],
        },
      },
      {
        name: "create_reply_draft",
        description: "Create a threaded reply draft to an existing email. Opens reply in compose window with quoted content. Body MUST be markdown only — no HTML or plain text.",
        inputSchema: {
          type: "object",
          properties: {
            email_id: {
              type: "number",
              description: "The ID of the email to reply to",
            },
            body: {
              type: "string",
              description: "Reply body content in markdown format. MUST be markdown only — no raw HTML.",
            },
            reply_all: {
              type: "boolean",
              description: "Whether to reply to all recipients (default: false)",
            },
          },
          required: ["email_id", "body"],
        },
      },
      {
        name: "search_emails",
        description: "Search emails by subject or sender across all Inboxes by default. Pass a folder name to restrict. Returns ID, folder, sender, subject, date.",
        inputSchema: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "Search query (searches subject and sender)",
            },
            folder: {
              type: "string",
              description: "Folder to search in. Defaults to all Inboxes across all accounts.",
            },
            limit: {
              type: "number",
              description: "Maximum results (default: 20)",
            },
          },
          required: ["query"],
        },
      },
      {
        name: "download_attachment",
        description: "Download an attachment from an email and return the file path. Files are saved as {email_id}_{filename} to avoid collisions. Use get_email first to see attachment names.",
        inputSchema: {
          type: "object",
          properties: {
            email_id: {
              type: "number",
              description: "The ID of the email containing the attachment",
            },
            attachment_name: {
              type: "string",
              description: "Exact filename of the attachment (from get_email output). If omitted, downloads all attachments.",
            },
            destination: {
              type: "string",
              description: "Directory to save attachments to (default: /tmp/outlook-attachments/)",
            },
          },
          required: ["email_id"],
        },
      },
      {
        name: "move_email",
        description: "Move an email to a different folder (e.g. Archive). Use list_folders to see available destination folder names.",
        inputSchema: {
          type: "object",
          properties: {
            email_id: {
              type: "number",
              description: "The ID of the email to move",
            },
            destination_folder: {
              type: "string",
              description: "Name of the destination folder (e.g. 'Archive', 'Junk Email')",
            },
          },
          required: ["email_id", "destination_folder"],
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case "list_folders": {
        const rows = runSqlite(`
          SELECT
            f.Folder_Name,
            f.Record_AccountUID,
            COALESCE(ae.Account_Name, am.Account_Name, 'Local') AS AccountName,
            (SELECT COUNT(*) FROM Mail m WHERE m.Record_FolderID = f.Record_RecordID) AS MsgCount
          FROM Folders f
          LEFT JOIN AccountsExchange ae ON ae.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
          LEFT JOIN AccountsMail am ON am.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
          WHERE f.Folder_Name NOT LIKE 'Placeholder%'
            AND f.Folder_Name != ''
          ORDER BY AccountName, f.Folder_Name
        `);

        if (!rows || rows.length === 0) {
          return { content: [{ type: "text", text: "No folders found" }] };
        }

        const output = rows
          .map((r) => `${r.AccountName}/${r.Folder_Name} (${r.MsgCount} messages)`)
          .join("\n");

        return { content: [{ type: "text", text: output }] };
      }

      case "read_emails": {
        const folder = args?.folder || null;
        const limit = Math.min(args?.limit || 10, 50);

        let folderClause;
        let folderLabel;
        if (folder) {
          const folderRows = runSqlite(
            `SELECT Record_RecordID FROM Folders WHERE Folder_Name = ${JSON.stringify(folder)} ORDER BY (SELECT COUNT(*) FROM Mail m WHERE m.Record_FolderID = Record_RecordID) DESC LIMIT 1`
          );
          if (!folderRows || folderRows.length === 0) {
            return {
              content: [{ type: "text", text: `Error: Folder '${folder}' not found` }],
              isError: true,
            };
          }
          folderClause = `Record_FolderID = ${folderRows[0].Record_RecordID}`;
          folderLabel = folder;
        } else {
          const inboxRows = runSqlite(`SELECT Record_RecordID FROM Folders WHERE Folder_Name = 'Inbox'`);
          if (!inboxRows || inboxRows.length === 0) {
            return { content: [{ type: "text", text: "No Inbox folders found" }], isError: true };
          }
          const ids = inboxRows.map(r => r.Record_RecordID).join(",");
          folderClause = `Record_FolderID IN (${ids})`;
          folderLabel = "All Inboxes";
        }

        const emails = runSqlite(`
          SELECT m.Record_RecordID, m.Message_NormalizedSubject, m.Message_SenderList,
                 m.Message_TimeReceived, m.Message_Preview, m.Message_HasAttachment,
                 f.Folder_Name,
                 COALESCE(ae.Account_Name, am.Account_Name, 'Local') AS AccountName
          FROM Mail m
          JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
          LEFT JOIN AccountsExchange ae ON ae.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
          LEFT JOIN AccountsMail am ON am.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
          WHERE ${folderClause}
          ORDER BY m.Message_TimeReceived DESC
          LIMIT ${limit}
        `);

        if (!emails || emails.length === 0) {
          return { content: [{ type: "text", text: `No emails in ${folderLabel}` }] };
        }

        const output = emails.map((msg) => {
          const date = msg.Message_TimeReceived
            ? new Date(msg.Message_TimeReceived * 1000).toString()
            : "";
          const preview = msg.Message_Preview
            ? String(msg.Message_Preview).slice(0, 100) + (String(msg.Message_Preview).length > 100 ? "..." : "")
            : "";
          const account = msg.AccountName ? `${msg.AccountName}/` : "";
          const attach = msg.Message_HasAttachment ? " 📎" : "";
          return `---\nID: ${msg.Record_RecordID}\nFolder: ${account}${msg.Folder_Name}\nFrom: ${msg.Message_SenderList || ""}\nSubject: ${msg.Message_NormalizedSubject || ""}${attach}\nDate: ${date}\nPreview: ${preview}`;
        }).join("\n");

        return { content: [{ type: "text", text: output }] };
      }

      case "get_email": {
        const emailId = args?.email_id;
        if (!emailId) {
          return {
            content: [{ type: "text", text: "Error: email_id is required" }],
            isError: true,
          };
        }

        const script = `
tell application "Microsoft Outlook"
    set sourceAcct to missing value
${findMessageScript(emailId)}

    if targetMsg is missing value then
        return "Error: Email with ID ${emailId} not found"
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

    set output to "Subject: " & msgSubject & "\\n"
    set output to output & "From: " & msgSender & " <" & msgSenderEmail & ">\\n"
    set output to output & "To: " & toList & "\\n"
    if ccList is not "" then
        set output to output & "CC: " & ccList & "\\n"
    end if
    set output to output & "Date: " & msgDate & "\\n"
    if attInfo is not "" then
        set output to output & attInfo & "\\n"
    end if
    set output to output & "\\n--- Content ---\\n" & msgContent

    return output
end tell`;
        const result = runAppleScriptHeredoc(script);
        return {
          content: [{ type: "text", text: result }],
        };
      }

      case "create_draft": {
        const subject = args?.subject || "";
        const body = args?.body || "";
        const to = args?.to || "";
        const cc = args?.cc || "";

        const escapedSubject = subject.replace(/"/g, '\\"');
        const htmlBody = markdownToHtml(body);
        const escapedBody = htmlBody.replace(/"/g, '\\"');

        let script = `
tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {subject:"${escapedSubject}", content:"${escapedBody}"}`;

        if (to) {
          script += `
    make new to recipient at newMsg with properties {email address:{address:"${to}"}}`;
        }
        if (cc) {
          script += `
    make new cc recipient at newMsg with properties {email address:{address:"${cc}"}}`;
        }

        script += `
    open newMsg
    return "Draft created successfully with subject: ${escapedSubject}"
end tell`;

        const result = runAppleScriptHeredoc(script);
        return {
          content: [{ type: "text", text: result }],
        };
      }

      case "create_reply_draft": {
        const emailId = args?.email_id;
        const body = args?.body || "";
        const replyAll = args?.reply_all || false;

        if (!emailId) {
          return {
            content: [{ type: "text", text: "Error: email_id is required" }],
            isError: true,
          };
        }

        const htmlBody = markdownToHtml(body);
        // Wrap in a styled container so it matches Outlook's default font
        const styledBody = `<div style="font-family: Aptos, Calibri, sans-serif; font-size: 12pt;">${htmlBody}</div>`;
        const replyAllFlag = replyAll ? " with reply to all" : "";

        // Step 1: Find the email (no UI)
        const findScript = `
tell application "Microsoft Outlook"
    set sourceAcct to missing value
${findMessageScript(emailId)}
    if targetMsg is missing value then
        return "Error: Email with ID ${emailId} not found"
    end if
    return id of targetMsg
end tell`;

        const findResult = runAppleScriptHeredoc(findScript);
        if (findResult.startsWith("Error:")) {
          return { content: [{ type: "text", text: findResult }], isError: true };
        }

        // Step 2: Copy rich text to clipboard
        // Write HTML to temp file to avoid shell interpolation of $ in dollar amounts
        const tmpHtml = path.join(os.tmpdir(), `outlook-reply-${emailId}.html`);
        fs.writeFileSync(tmpHtml, styledBody, "utf-8");
        try {
          execSync(`textutil -inputencoding UTF-8 -format html -convert rtf -stdout ${JSON.stringify(tmpHtml)} | pbcopy`, {
            shell: "/bin/bash",
            timeout: 10000,
          });
        } finally {
          fs.unlinkSync(tmpHtml);
        }

        // Step 3: Activate Outlook, open reply, click body, paste
        const replyScript = `
tell application "Microsoft Outlook"
    activate
    set sourceAcct to missing value
${findMessageScript(emailId)}
    reply to targetMsg${replyAllFlag}
end tell
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

        const replyResult = runAppleScriptHeredoc(replyScript);
        if (replyResult && replyResult.startsWith("Error:")) {
          return { content: [{ type: "text", text: replyResult }], isError: true };
        }

        return {
          content: [{ type: "text", text: `Reply draft created for email ID ${emailId}` }],
        };
      }

      case "search_emails": {
        const query = args?.query || "";
        const folder = args?.folder || null;
        const limit = Math.min(args?.limit || 20, 50);

        if (!query) {
          return {
            content: [{ type: "text", text: "Error: query is required" }],
            isError: true,
          };
        }

        let folderClause;
        let folderLabel;
        if (folder) {
          const folderRows = runSqlite(
            `SELECT Record_RecordID FROM Folders WHERE Folder_Name = ${JSON.stringify(folder)} ORDER BY (SELECT COUNT(*) FROM Mail m WHERE m.Record_FolderID = Record_RecordID) DESC LIMIT 1`
          );
          if (!folderRows || folderRows.length === 0) {
            return {
              content: [{ type: "text", text: `Error: Folder '${folder}' not found` }],
              isError: true,
            };
          }
          folderClause = `m.Record_FolderID = ${folderRows[0].Record_RecordID}`;
          folderLabel = folder;
        } else {
          const inboxRows = runSqlite(`SELECT Record_RecordID FROM Folders WHERE Folder_Name = 'Inbox'`);
          if (!inboxRows || inboxRows.length === 0) {
            return { content: [{ type: "text", text: "No Inbox folders found" }], isError: true };
          }
          const ids = inboxRows.map(r => r.Record_RecordID).join(",");
          folderClause = `m.Record_FolderID IN (${ids})`;
          folderLabel = "All Inboxes";
        }

        const likeQuery = query.replace(/'/g, "''");

        const emails = runSqlite(`
          SELECT m.Record_RecordID, m.Message_NormalizedSubject, m.Message_SenderList,
                 m.Message_TimeReceived, f.Folder_Name,
                 COALESCE(ae.Account_Name, am.Account_Name, 'Local') AS AccountName
          FROM Mail m
          JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
          LEFT JOIN AccountsExchange ae ON ae.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
          LEFT JOIN AccountsMail am ON am.Record_RecordID = (f.Record_AccountUID & 0xFFFFFFFF)
          WHERE ${folderClause}
            AND (m.Message_NormalizedSubject LIKE '%${likeQuery}%' OR m.Message_SenderList LIKE '%${likeQuery}%')
          ORDER BY m.Message_TimeReceived DESC
          LIMIT ${limit}
        `);

        if (!emails || emails.length === 0) {
          return {
            content: [{ type: "text", text: `No emails found matching "${query}" in ${folderLabel}` }],
          };
        }

        const output = emails.map((msg) => {
          const date = msg.Message_TimeReceived
            ? new Date(msg.Message_TimeReceived * 1000).toString()
            : "";
          const account = msg.AccountName ? `${msg.AccountName}/` : "";
          return `ID: ${msg.Record_RecordID} | Folder: ${account}${msg.Folder_Name} | From: ${msg.Message_SenderList || ""} | Subject: ${msg.Message_NormalizedSubject || ""} | Date: ${date}`;
        }).join("\n---\n");

        return { content: [{ type: "text", text: output + "\n---\n" }] };
      }

      case "download_attachment": {
        const emailId = args?.email_id;
        if (!emailId) {
          return {
            content: [{ type: "text", text: "Error: email_id is required" }],
            isError: true,
          };
        }

        const attachmentName = args?.attachment_name || null;
        const escapedAttName = attachmentName ? attachmentName.replace(/"/g, '\\"') : "";
        const saveDir = args?.destination || "/tmp/outlook-attachments";

        // Ensure save directory exists
        fs.mkdirSync(saveDir, { recursive: true });

        // Use a per-call temp dir so AppleScript saves don't collide
        const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "outlook-att-"));

        const script = attachmentName
          ? `
tell application "Microsoft Outlook"
${findMessageScript(emailId)}
    if targetMsg is missing value then
        return "Error: Email with ID ${emailId} not found"
    end if

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
    if targetMsg is missing value then
        return "Error: Email with ID ${emailId} not found"
    end if

    set attList to attachments of targetMsg
    set attCount to count of attList
    if attCount is 0 then
        return "Error: Email has no attachments"
    end if

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
          // Clean up temp dir
          fs.rmSync(tmpDir, { recursive: true, force: true });
          return { content: [{ type: "text", text: result }], isError: true };
        }

        // Move files from temp dir to destination with unique names
        const lines = result.split("\n").filter(Boolean);
        const output = [];
        for (const line of lines) {
          const [filename, sizeStr] = line.split("\t");
          const tmpFile = path.join(tmpDir, filename);
          const destFile = uniquePath(saveDir, emailId, filename);
          fs.renameSync(tmpFile, destFile);
          output.push(`Saved: ${destFile} (${sizeStr} bytes)`);
        }

        // Clean up temp dir
        fs.rmSync(tmpDir, { recursive: true, force: true });

        return {
          content: [{ type: "text", text: output.join("\n") }],
        };
      }

      case "move_email": {
        const emailId = args?.email_id;
        const destFolder = args?.destination_folder;

        if (!emailId || !destFolder) {
          return {
            content: [{ type: "text", text: "Error: email_id and destination_folder are required" }],
            isError: true,
          };
        }

        const escapedDest = destFolder.replace(/"/g, '\\"');

        // Look up the account so we can scope the destination folder to the same account
        const loc = lookupEmailLocation(emailId);
        let destFolderScript;
        if (loc && loc.accountName) {
          const escapedAcct = loc.accountName.replace(/"/g, '\\"');
          destFolderScript = `
    set destF to missing value
    try
        set destF to folder "${escapedDest}" of exchange account "${escapedAcct}"
    end try
    if destF is missing value then
        -- Fall back to searching all folders
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
        return "Error: Email with ID ${emailId} not found"
    end if
${destFolderScript}

    if destF is missing value then
        return "Error: Folder '${escapedDest}' not found"
    end if

    move targetMsg to destF
    return "Moved email ID ${emailId} to ${escapedDest}"
end tell`;

        const result = runAppleScriptHeredoc(script);
        return {
          content: [{ type: "text", text: result }],
        };
      }

      default:
        return {
          content: [{ type: "text", text: `Unknown tool: ${name}` }],
          isError: true,
        };
    }
  } catch (error) {
    return {
      content: [{ type: "text", text: `Error: ${error.message}` }],
      isError: true,
    };
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Outlook MCP server running on stdio");
}

main().catch(console.error);
