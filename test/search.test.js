import { describe, it, before } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";

// Isolate the FTS index to a temp DB BEFORE importing index.js (FTS_DB is read at module load).
const TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), "mcp-outlook-search-"));
const FIXTURE_DB = path.join(TMP_DIR, "fixture.db");
process.env.MCP_OUTLOOK_FTS_DB = FIXTURE_DB;

const mod = await import("../index.js");
const {
  queryTerms, ftsToken, likePattern, sqlLit,
  ensureFtsDb, handleSearch, parseCliArgs, FTS_COLUMNS,
} = mod;

// Seed the fixture index directly via the bodies table (triggers populate FTS).
// Mirrors what syncFtsIndex would produce, without needing Outlook.
function seed() {
  ensureFtsDb();
  const rows = [
    // id, subject, sender_name, sender_addr, to_addr, cc_addr, body, folder, account, ts, read, attach
    [1, "Q3 budget call", "Tamm Sjodin", "tamm@example.com", "finance@racat.com", "", "Let us lock the Q3 numbers.", "Inbox", "Work", 1700000000, 1, 0],
    [2, "Re: lunch", "Bob Jones", "bob@other.com", "someone@x.com", "", "see you at the @tamm cafe near the park", "Inbox", "Work", 1700000100, 1, 0],
    [3, "Project X kickoff", "Alice Smith", "alice@x.com", "tamm@example.com", "team@x.com", "Project X with person A is going well, lots to do", "Sent Items", "Work", 1700000200, 0, 1],
    [4, "Newsletter", "News Bot", "news@foo.com", "list@foo.com", "", "nothing relevant here at all", "Inbox", "Work", 1700000300, 1, 0],
    [5, "person A sync", "Carol", "carol@y.com", "tamm@example.com", "", "Quick note about person A and the roadmap", "Archive", "Work", 1700000400, 1, 0],
  ];
  const values = rows.map(r => {
    const cells = r.map((c, i) => (i >= 1 && i <= 8) ? `'${sqlLit(String(c))}'` : c);
    return `(${cells.join(",")})`;
  }).join(",");
  execFileSync("sqlite3", [FIXTURE_DB,
    `INSERT INTO bodies (id, subject, sender_name, sender_addr, to_addr, cc_addr, body, folder, account, ts, read_flag, has_attach) VALUES ${values};`
  ]);
}

function idsFrom(result) {
  const text = result.content[0].text;
  return [...text.matchAll(/ID:(\d+)/g)].map(m => Number(m[1]));
}

before(() => { seed(); });

describe("queryTerms", () => {
  it("splits, lowercases, dedupes, drops ≤2-char tokens", () => {
    assert.deepEqual(queryTerms("Project X with person A"), ["project", "with", "person"]);
  });
  it("handles punctuation separators", () => {
    assert.deepEqual(queryTerms("budget, Q3; call"), ["budget", "call"]);
  });
  it("returns empty for all-short input", () => {
    assert.deepEqual(queryTerms("a X to"), []);
  });
});

describe("ftsToken (FTS5 string-token escaping)", () => {
  it("wraps in double quotes", () => {
    assert.equal(ftsToken("tamm@"), `"tamm@"`);
  });
  it("doubles internal double-quotes then quotes", () => {
    assert.equal(ftsToken('a"b'), `"a""b"`);
  });
  it("single quotes get doubled for SQL literal layer", () => {
    assert.equal(ftsToken("o'brien"), `"o''brien"`);
  });
});

describe("likePattern", () => {
  it("wraps in % and escapes LIKE wildcards", () => {
    assert.equal(likePattern("50%_off"), "%50\\%\\_off%");
  });
});

describe("parseCliArgs", () => {
  it("separates positionals from flags with values", () => {
    const { positional, flags } = parseCliArgs(["tamm@", "--folder", "Sent", "--limit", "20"]);
    assert.deepEqual(positional, ["tamm@"]);
    assert.equal(flags.folder, "Sent");
    assert.equal(flags.limit, "20");
  });
  it("treats trailing --flag as boolean", () => {
    const { flags } = parseCliArgs(["x", "--unread"]);
    assert.equal(flags.unread, true);
  });
});

describe("handleSearch end-to-end (fixture index)", () => {
  it("finds 'tamm@' across sender, recipient and body — exact matches ranked above fuzzy '@tamm'", () => {
    const res = handleSearch({ query: "tamm@" });
    const ids = idsFrom(res);
    // 1 (sender), 3 & 5 (recipient) are exact substring hits; 2 (@tamm body) is fuzzy.
    assert.ok(ids.includes(1) && ids.includes(3) && ids.includes(5), `expected exact hits, got ${ids}`);
    assert.ok(ids.includes(2), `expected fuzzy @tamm hit, got ${ids}`);
    // The fuzzy-only row (2) must rank below at least one exact row.
    assert.ok(ids.indexOf(2) > 0, `fuzzy row should not rank first: ${ids}`);
  });

  it("searches Sent and Archive by default (not inbox-only)", () => {
    const ids = idsFrom(handleSearch({ query: "tamm@" }));
    assert.ok(ids.includes(3), "Sent folder email must be found");   // id 3 = Sent
    assert.ok(ids.includes(5), "Archive folder email must be found"); // id 5 = Archive
  });

  it("multi-term query ranks by number of terms matched", () => {
    const ids = idsFrom(handleSearch({ query: "Project X with person A" }));
    // id 3 matches project+person; id 5 matches person only. 3 should rank above 5.
    assert.ok(ids.includes(3) && ids.includes(5), `got ${ids}`);
    assert.ok(ids.indexOf(3) < ids.indexOf(5), `more-terms row should rank higher: ${ids}`);
  });

  it("never returns a bare zero when only part of a multi-term query matches", () => {
    const ids = idsFrom(handleSearch({ query: "budget rocketship aardvark" }));
    assert.ok(ids.includes(1), `partial match on 'budget' should surface id 1: ${ids}`);
  });

  it("folder filter narrows scope", () => {
    const ids = idsFrom(handleSearch({ query: "tamm@", folder: "Sent" }));
    assert.deepEqual(ids, [3], `folder=Sent should match "Sent Items", got ${ids}`);
  });

  it("folder filter accepts the Account/Folder string list_folders prints", () => {
    const ids = idsFrom(handleSearch({ query: "tamm@", folder: "Work/Sent Items" }));
    assert.deepEqual(ids, [3], `composite folder id should work, got ${ids}`);
  });

  it("folder filter is case-insensitive", () => {
    const ids = idsFrom(handleSearch({ query: "tamm@", folder: "archive" }));
    assert.deepEqual(ids, [5], `lowercase folder should match Archive, got ${ids}`);
  });

  it("folder filter does not silently match everything on junk input", () => {
    const ids = idsFrom(handleSearch({ query: "tamm@", folder: "Nope" }));
    assert.deepEqual(ids, [], `unknown folder should return nothing, got ${ids}`);
  });

  it("matches body substrings, not just metadata", () => {
    const ids = idsFrom(handleSearch({ query: "roadmap" }));
    assert.deepEqual(ids, [5], `body-only term should find id 5: ${ids}`);
  });

  it("empty query lists newest first", () => {
    const ids = idsFrom(handleSearch({ query: "" }));
    assert.equal(ids[0], 5, `newest (ts) should be first: ${ids}`); // id 5 has highest ts
  });

  it("FTS_COLUMNS covers all six searchable fields", () => {
    assert.deepEqual(FTS_COLUMNS, ["subject", "sender_name", "sender_addr", "to_addr", "cc_addr", "body"]);
  });
});
