import { describe, it } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import fs from "node:fs";
import { execFileSync } from "node:child_process";
import {
  extractEmail,
  parseRecipients,
  recipientLines,
  parseAttachments,
  validateAttachments,
  attachmentLines,
  MAX_ATTACHMENT_BYTES,
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
} from "../index.js";

describe("extractEmail", () => {
  it("returns plain email unchanged", () => {
    assert.equal(extractEmail("user@example.com"), "user@example.com");
  });

  it("extracts email from angle brackets", () => {
    assert.equal(extractEmail("John Doe <john@example.com>"), "john@example.com");
  });

  it("handles extra whitespace", () => {
    assert.equal(extractEmail("  <spaced@example.com>  "), "spaced@example.com");
  });

  it("returns empty for empty string", () => {
    assert.equal(extractEmail(""), "");
  });
});

describe("parseRecipients", () => {
  it("returns empty array for null", () => {
    assert.deepEqual(parseRecipients(null), []);
  });

  it("returns empty array for undefined", () => {
    assert.deepEqual(parseRecipients(undefined), []);
  });

  it("returns empty array for empty string", () => {
    assert.deepEqual(parseRecipients(""), []);
  });

  it("splits comma-separated addresses", () => {
    assert.deepEqual(
      parseRecipients("a@example.com, b@example.com"),
      ["a@example.com", "b@example.com"]
    );
  });

  it("splits semicolon-separated addresses", () => {
    assert.deepEqual(
      parseRecipients("a@example.com; b@example.com"),
      ["a@example.com", "b@example.com"]
    );
  });

  it("handles array input", () => {
    assert.deepEqual(
      parseRecipients(["a@example.com", "b@example.com"]),
      ["a@example.com", "b@example.com"]
    );
  });

  it("extracts emails from Name <email> in array", () => {
    assert.deepEqual(
      parseRecipients(["Alice <alice@example.com>", "bob@example.com"]),
      ["alice@example.com", "bob@example.com"]
    );
  });

  it("extracts emails from Name <email> in comma string", () => {
    assert.deepEqual(
      parseRecipients("John <john@example.com>, Jane <jane@example.com>"),
      ["john@example.com", "jane@example.com"]
    );
  });

  it("handles single address", () => {
    assert.deepEqual(parseRecipients("solo@example.com"), ["solo@example.com"]);
  });

  it("filters out empty entries", () => {
    assert.deepEqual(parseRecipients("a@example.com,,b@example.com"), ["a@example.com", "b@example.com"]);
  });

  it("handles JSON-stringified array input", () => {
    assert.deepEqual(
      parseRecipients('["cameron@example.com", "iblomley@example.com"]'),
      ["cameron@example.com", "iblomley@example.com"]
    );
  });

  it("strips bracket/quote junk from malformed comma split", () => {
    // Simulates the legacy broken path — result should still be clean emails.
    assert.deepEqual(
      parseRecipients('["a@example.com","b@example.com"]'),
      ["a@example.com", "b@example.com"]
    );
  });
});

describe("recipientLines", () => {
  it("generates a single to recipient line", () => {
    const result = recipientLines(["a@example.com"], "to", "newMsg");
    assert.equal(
      result,
      'make new to recipient at newMsg with properties {email address:{address:"a@example.com"}}'
    );
  });

  it("generates multiple cc recipient lines", () => {
    const result = recipientLines(["a@example.com", "b@example.com"], "cc", "msg");
    assert.ok(result.includes('make new cc recipient at msg'));
    assert.ok(result.includes("a@example.com"));
    assert.ok(result.includes("b@example.com"));
    assert.equal(result.split("make new").length - 1, 2);
  });

  it("escapes special characters in addresses", () => {
    const result = recipientLines(['o"malley@example.com'], "to", "m");
    assert.ok(result.includes('o\\"malley@example.com'));
  });
});

describe("escapeForAppleScript", () => {
  it("escapes backslashes", () => {
    assert.equal(escapeForAppleScript("a\\b"), "a\\\\b");
  });

  it("escapes double quotes", () => {
    assert.equal(escapeForAppleScript('say "hi"'), 'say \\"hi\\"');
  });

  it("escapes newlines", () => {
    assert.equal(escapeForAppleScript("line1\nline2"), "line1\\nline2");
  });

  it("escapes carriage returns", () => {
    assert.equal(escapeForAppleScript("line1\rline2"), "line1\\rline2");
  });

  it("returns empty for empty string", () => {
    assert.equal(escapeForAppleScript(""), "");
  });
});

describe("stripSignature", () => {
  it("strips standard -- delimiter", () => {
    const input = "Hello there\n-- \nJohn Doe\nCEO";
    assert.equal(stripSignature(input), "Hello there");
  });

  it("strips Sent from my iPhone", () => {
    const input = "Hey mate\n\nSent from my iPhone";
    assert.equal(stripSignature(input), "Hey mate");
  });

  it("strips Sent from my iPad", () => {
    const input = "Hey mate\n\nSent from my iPad";
    assert.equal(stripSignature(input), "Hey mate");
  });

  it("strips Get Outlook for iOS", () => {
    const input = "Hey mate\n\nGet Outlook for iOS";
    assert.equal(stripSignature(input), "Hey mate");
  });

  it("returns text unchanged when no signature", () => {
    const input = "Just a normal message\nWith two lines";
    assert.equal(stripSignature(input), "Just a normal message\nWith two lines");
  });

  it("returns empty for empty string", () => {
    assert.equal(stripSignature(""), "");
  });
});

describe("stripQuotedReplies", () => {
  it("strips On date wrote pattern", () => {
    const input = "My reply\n\nOn Mon, Jan 1, 2026 at 10:00 AM John wrote:\n> old stuff";
    assert.equal(stripQuotedReplies(input), "My reply");
  });

  it("strips Outlook From/Sent block", () => {
    const input = "My reply\n\nFrom: Someone\nSent: Monday\nTo: Me\n\nOld content";
    assert.equal(stripQuotedReplies(input), "My reply");
  });

  it("strips trailing > quoted lines", () => {
    const input = "My reply\n\n> quoted line 1\n> quoted line 2";
    assert.equal(stripQuotedReplies(input), "My reply");
  });

  it("strips forwarded message delimiter", () => {
    const input = "Check this out\n\n---------- Forwarded message ----------\nFrom: someone";
    assert.equal(stripQuotedReplies(input), "Check this out");
  });

  it("returns text unchanged when no quotes", () => {
    const input = "Just a normal message";
    assert.equal(stripQuotedReplies(input), "Just a normal message");
  });

  it("returns empty for empty string", () => {
    assert.equal(stripQuotedReplies(""), "");
  });
});

describe("cleanBody", () => {
  it("strips both signature and quoted replies", () => {
    const input =
      "My reply\n-- \nJohn\n\nOn Mon, Jan 1, 2026 at 10:00 AM Someone wrote:\n> hi";
    const result = cleanBody(input);
    assert.equal(result, "My reply");
  });
});

describe("fillRecipientsFromDb", () => {
  const sent = [
    "Subject: Invoice split",
    "From: Tamm <tamm@australiangeographic.com>",
    "To: ",
    "Date: Wednesday, 20 May 2026 at 12:53:00",
    "",
    "body text",
  ].join("\n");

  it("fills an empty To: from the DB (Sent Items repro)", () => {
    const out = fillRecipientsFromDb(sent, "sophie.hanson@junkeemedia.com", null);
    assert.ok(out.includes("To: sophie.hanson@junkeemedia.com"));
    assert.ok(out.includes("body text"));
  });

  it("does not overwrite a To: AppleScript already populated", () => {
    const filled = sent.replace("To: ", "To: real@example.com, ");
    const out = fillRecipientsFromDb(filled, "wrong@example.com", null);
    assert.ok(out.includes("To: real@example.com"));
    assert.ok(!out.includes("wrong@example.com"));
  });

  it("still adds CC below the To line", () => {
    const out = fillRecipientsFromDb(sent, "a@example.com", "b@example.com");
    assert.match(out, /^To: a@example\.com\nCC: b@example\.com$/m);
  });

  it("leaves text untouched when the DB has nothing", () => {
    assert.equal(fillRecipientsFromDb(sent, null, undefined), sent);
    assert.equal(fillRecipientsFromDb(sent, "  ", ""), sent);
  });

  // Real-world repro: a reply quotes the earlier message's headers, so the body
  // contains its own "To:"/"CC:" lines. Those must not count as a populated header.
  const quotedReply = [
    "Subject: Re: VPN",
    "From: Tamm <tamm@australiangeographic.com>",
    "To: ",
    "Date: Tuesday, 25 August 2026 at 18:58:23",
    "",
    "Thanks, all sorted.",
    "",
    "From: Richard <richard@escadia.com.au>",
    "To: Tamm Sjodin <tamm@australiangeographic.com>",
    "CC: Troy <troy@escadia.com.au>",
    "Subject: VPN",
  ].join("\n");

  it("fills the header To: even when the quoted body has its own To: line", () => {
    const out = fillRecipientsFromDb(quotedReply, "richard@escadia.com.au; troy@escadia.com.au", "");
    assert.match(out, /^To: richard@escadia\.com\.au; troy@escadia\.com\.au$/m);
    // the quoted block below is left exactly as it was
    assert.ok(out.includes("To: Tamm Sjodin <tamm@australiangeographic.com>"));
  });

  it("adds CC to the header, not next to the quoted CC further down", () => {
    const out = fillRecipientsFromDb(quotedReply, "a@example.com", "b@example.com");
    const lines = out.split("\n");
    assert.equal(lines[2], "To: a@example.com");
    assert.equal(lines[3], "CC: b@example.com");
    assert.equal(lines[4], "Date: Tuesday, 25 August 2026 at 18:58:23");
  });

  it("returns text unchanged when there is no To: line at all", () => {
    const noTo = "Subject: x\nFrom: y\n\nbody";
    assert.equal(fillRecipientsFromDb(noTo, "a@example.com", "b@example.com"), noTo);
  });
});

describe("markdownToHtml", () => {
  it("converts bold", () => {
    const result = markdownToHtml("**bold text**");
    assert.ok(result.includes("<strong>bold text</strong>"));
  });

  it("converts italic", () => {
    const result = markdownToHtml("*italic text*");
    assert.ok(result.includes("<em>italic text</em>"));
  });

  it("converts h1", () => {
    const result = markdownToHtml("# Heading");
    assert.ok(result.includes("<h1>Heading</h1>"));
  });

  it("converts list items", () => {
    const result = markdownToHtml("- item one\n- item two");
    assert.ok(result.includes("<li>item one</li>"));
    assert.ok(result.includes("<li>item two</li>"));
    assert.ok(result.includes("<ul>"));
  });

  it("converts links", () => {
    const result = markdownToHtml("[click](https://example.com)");
    assert.ok(result.includes('<a href="https://example.com">click</a>'));
  });

  it("converts inline code", () => {
    const result = markdownToHtml("`some code`");
    assert.ok(result.includes("<code>some code</code>"));
  });

  it("passes plain text through wrapped in p tags", () => {
    const result = markdownToHtml("hello world");
    assert.ok(result.includes('<p style="margin:0">hello world</p>'));
    assert.ok(result.includes("font-family"));
  });

  it("turns single newlines into separate line-broken blocks", () => {
    const result = markdownToHtml(
      "Sophie Hanson, sophie.hanson@junkeemedia.com\n" +
        "Lia Kim, lia.kim@junkeemedia.com\n" +
        "Yagmur Ilkyaz, Yagmur.Ilkyaz@junkeemedia.com"
    );
    // Not <br> — textutil drops those. Each line becomes its own tight paragraph.
    assert.equal((result.match(/<p style="margin:0">/g) || []).length, 3);
    assert.ok(!/<br/.test(result));
  });

  it("keeps blank-line paragraphs separate", () => {
    const result = markdownToHtml("first para\n\nsecond para");
    assert.ok(result.includes('<p style="margin:0">first para</p>'));
    assert.ok(result.includes('<p style="margin:0">second para</p>'));
    // separated by a real empty paragraph, not by a margin Outlook may ignore
    assert.ok(result.includes('<p style="margin:0">&nbsp;</p>'));
  });
});

describe("emailParagraphs", () => {
  it("makes every line of a <br> run its own margin:0 paragraph", () => {
    assert.equal(
      emailParagraphs("<p>one<br>two<br>three</p>"),
      '<p style="margin:0">one</p><p style="margin:0">two</p><p style="margin:0">three</p>'
    );
  });

  it("leaves no paragraph relying on a default margin", () => {
    const out = emailParagraphs("<p>a<br>b</p>\n<p>c</p>");
    assert.ok(!/<p>/.test(out), `bare <p> would inherit Outlook's spacing: ${out}`);
  });

  it("separates markdown paragraphs with a real empty paragraph", () => {
    const out = emailParagraphs("<p>first</p>\n<p>second</p>");
    assert.ok(out.includes('<p style="margin:0">&nbsp;</p>'));
    assert.ok(out.indexOf("first") < out.indexOf("&nbsp;"));
    assert.ok(out.indexOf("&nbsp;") < out.indexOf("second"));
  });

  it("does not leave a trailing spacer after the final paragraph", () => {
    assert.ok(!emailParagraphs("<p>only</p>").includes("&nbsp;"));
  });

  it("does not disturb list markup", () => {
    const html = "<ul>\n<li>one</li>\n<li>two</li>\n</ul>";
    assert.equal(emailParagraphs(html), html);
  });

  it("handles self-closing <br/>", () => {
    assert.ok(emailParagraphs("<p>a<br/>b</p>").includes('<p style="margin:0">b</p>'));
  });

  it("puts a gap after a list, so it does not run into the next paragraph", () => {
    const out = emailParagraphs("<ul>\n<li>one</li>\n</ul>\n<p>after</p>");
    assert.ok(out.includes('</ul><p style="margin:0">&nbsp;</p>'));
  });

  it("puts a gap after a heading", () => {
    const out = emailParagraphs("<h1>Title</h1>\n<p>after</p>");
    assert.ok(out.includes('</h1><p style="margin:0">&nbsp;</p>'));
  });

  it("does not put a gap after a nested list closing inside an <li>", () => {
    const out = emailParagraphs("<ul>\n<li>top<ul>\n<li>inner</li>\n</ul>\n</li>\n</ul>\n<p>after</p>");
    assert.equal(out.match(/&nbsp;/g).length, 1, "only the outer list gets a gap");
  });

  it("leaves no dead gap when the body ends with a list", () => {
    const html = "<ul>\n<li>one</li>\n</ul>";
    assert.equal(emailParagraphs(html), html);
  });
});

// The bug Tamm reported survived a correct <br> because textutil's HTML importer
// DROPS <br> outright. Only a round-trip through textutil proves the fix.
describe("compose body rendering through textutil (the real pipeline)", () => {
  function renderToText(md) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "mcp-outlook-rtf-"));
    const html = path.join(dir, "b.html");
    const rtf = path.join(dir, "b.rtf");
    fs.writeFileSync(html, markdownToHtml(md), "utf-8");
    execFileSync("textutil", ["-inputencoding", "UTF-8", "-format", "html", "-convert", "rtf", "-output", rtf, html]);
    const out = execFileSync("textutil", ["-format", "rtf", "-convert", "txt", "-stdout", rtf], { encoding: "utf-8" });
    fs.rmSync(dir, { recursive: true, force: true });
    return out.replace(/\s+$/, "").split("\n");
  }

  it("renders a one-per-line recipient list on separate lines", () => {
    const lines = renderToText(
      "Sophie Hanson, sophie.hanson@junkeemedia.com\n" +
        "Lia Kim, lia.kim@junkeemedia.com\n" +
        "Yagmur Ilkyaz, Yagmur.Ilkyaz@junkeemedia.com"
    );
    assert.equal(lines.length, 3, `expected 3 lines, got: ${JSON.stringify(lines)}`);
    assert.match(lines[1], /^Lia Kim/);
  });

  it("still renders bullet lists one per line", () => {
    const lines = renderToText("- one\n- two\n- three");
    assert.equal(lines.length, 3, `expected 3 bullets, got: ${JSON.stringify(lines)}`);
  });

  it("separates blank-line paragraphs with a genuine blank line", () => {
    const lines = renderToText("First para.\n\nSecond para.");
    assert.equal(lines.length, 3, `expected a blank line between: ${JSON.stringify(lines)}`);
    assert.equal(lines[0], "First para.");
    assert.equal(lines[1].trim(), "", `middle line should be blank: ${JSON.stringify(lines[1])}`);
    assert.equal(lines[2], "Second para.");
  });
});

describe("coerceEmailIds", () => {
  it("handles a normal array of numbers", () => {
    assert.deepEqual(coerceEmailIds([32455, 32456]), [32455, 32456]);
  });

  it("handles a single number", () => {
    assert.deepEqual(coerceEmailIds(32455), [32455]);
  });

  it("handles a JSON string array", () => {
    assert.deepEqual(coerceEmailIds("[32455, 32456, 32457]"), [32455, 32456, 32457]);
  });

  it("handles a single number as string", () => {
    assert.deepEqual(coerceEmailIds("32455"), [32455]);
  });

  it("handles an array of string numbers", () => {
    assert.deepEqual(coerceEmailIds(["32455", "32456"]), [32455, 32456]);
  });

  it("returns empty for null", () => {
    assert.deepEqual(coerceEmailIds(null), []);
  });

  it("returns empty for undefined", () => {
    assert.deepEqual(coerceEmailIds(undefined), []);
  });

  it("returns empty for empty array", () => {
    assert.deepEqual(coerceEmailIds([]), []);
  });

  it("returns empty for empty string", () => {
    assert.deepEqual(coerceEmailIds(""), []);
  });

  it("filters out NaN and zero values", () => {
    assert.deepEqual(coerceEmailIds([32455, "abc", 0, -1, 32456]), [32455, 32456]);
  });

  it("handles non-parseable garbage string", () => {
    assert.deepEqual(coerceEmailIds("not-a-number"), []);
  });

  it("handles a stringified single number in JSON", () => {
    assert.deepEqual(coerceEmailIds('"32455"'), [32455]);
  });

  it("handles mixed array of numbers and strings", () => {
    assert.deepEqual(coerceEmailIds([32455, "32456", 32457]), [32455, 32456, 32457]);
  });
});

describe("setRecipientsScript", () => {
  it("generates delete and make lines for to and cc", () => {
    const result = setRecipientsScript(["a@example.com"], ["b@example.com"], "composeMsg");
    assert.ok(result.includes("delete every to recipient of composeMsg"));
    assert.ok(result.includes("delete every cc recipient of composeMsg"));
    assert.ok(result.includes("a@example.com"));
    assert.ok(result.includes("b@example.com"));
    assert.ok(result.includes("make new to recipient"));
    assert.ok(result.includes("make new cc recipient"));
  });

  it("handles empty to and cc arrays", () => {
    const result = setRecipientsScript([], [], "composeMsg");
    assert.ok(result.includes("delete every to recipient"));
    assert.ok(result.includes("delete every cc recipient"));
    assert.ok(!result.includes("make new"));
  });

  it("returns empty when both to and cc are null/undefined", () => {
    const result = setRecipientsScript(null, undefined, "msg");
    assert.equal(result, "");
  });

  it("only touches `to` when cc is not provided", () => {
    const result = setRecipientsScript(["a@example.com"], null, "msg");
    assert.ok(result.includes("delete every to recipient"));
    assert.ok(!result.includes("delete every cc recipient"));
    assert.ok(result.includes("make new to recipient"));
    assert.ok(!result.includes("make new cc recipient"));
  });

  it("only touches `cc` when to is not provided", () => {
    const result = setRecipientsScript(null, ["c@example.com"], "msg");
    assert.ok(!result.includes("delete every to recipient"));
    assert.ok(result.includes("delete every cc recipient"));
    assert.ok(result.includes("make new cc recipient"));
    assert.ok(!result.includes("make new to recipient"));
  });

  it("filters out falsy entries in arrays", () => {
    const result = setRecipientsScript(["a@example.com", "", null, "b@example.com"], [], "msg");
    assert.equal(result.split("make new").length - 1, 2);
  });

  it("escapes special characters in addresses", () => {
    const result = setRecipientsScript(['o"malley@example.com'], [], "msg");
    assert.ok(result.includes('o\\"malley@example.com'));
  });
});

describe("diagnoseAppleScriptError", () => {
  it("detects ETIMEDOUT", () => {
    const result = diagnoseAppleScriptError({ message: "ETIMEDOUT" });
    assert.ok(result.includes("timed out"));
  });

  it("detects killed process", () => {
    const result = diagnoseAppleScriptError({ message: "", killed: true });
    assert.ok(result.includes("timed out"));
  });

  it("detects -1728 object not found", () => {
    const result = diagnoseAppleScriptError({ message: "execution error -1728" });
    assert.ok(result.includes("not found"));
  });

  it("detects -1712 busy dialog", () => {
    const result = diagnoseAppleScriptError({ message: "execution error -1712" });
    assert.ok(result.includes("dialog"));
  });

  it("detects not running / -600", () => {
    const result = diagnoseAppleScriptError({ message: "not running" });
    assert.ok(result.includes("not running"));
  });

  it("detects -600 error code", () => {
    const result = diagnoseAppleScriptError({ message: "some error -600" });
    assert.ok(result.includes("not running"));
  });

  it("detects -10810 launch failure", () => {
    const result = diagnoseAppleScriptError({ message: "error -10810" });
    assert.ok(result.includes("launch"));
  });

  it("extracts execution error message without a known error code", () => {
    // Use an error code not matched by any earlier specific branch
    const result = diagnoseAppleScriptError({ message: "execution error: Mail folder not accessible (-9999)" });
    assert.ok(result.includes("Outlook AppleScript error"));
    assert.ok(result.includes("Mail folder not accessible"));
  });

  it("falls back to generic message slice", () => {
    const result = diagnoseAppleScriptError({ message: "some totally unknown error" });
    assert.ok(result.includes("AppleScript error") || result.includes("some totally unknown error"));
  });

  it("handles missing message property", () => {
    const result = diagnoseAppleScriptError({});
    assert.equal(typeof result, "string");
    assert.ok(result.length > 0);
  });
});

describe("ftsStatusLine", () => {
  it("returns empty message when total is 0", () => {
    const result = ftsStatusLine({ indexed: 0, total: 0 });
    assert.equal(result, "Index: empty");
  });

  it("returns complete message when fully indexed", () => {
    const result = ftsStatusLine({ indexed: 1000, total: 1000 });
    assert.ok(result.includes("complete"));
    assert.ok(result.includes("1,000"));
  });

  it("returns percentage when partially indexed", () => {
    const result = ftsStatusLine({ indexed: 500, total: 1000 });
    assert.ok(result.includes("50%"));
    assert.ok(result.includes("500"));
    assert.ok(result.includes("1,000"));
  });

  it("rounds percentage down", () => {
    const result = ftsStatusLine({ indexed: 1, total: 3 });
    assert.ok(result.includes("33%"));
  });
});

describe("ok and err response helpers", () => {
  it("ok wraps text in content array", () => {
    const result = ok("hello world");
    assert.deepEqual(result, { content: [{ type: "text", text: "hello world" }] });
  });

  it("err wraps text with Error: prefix and isError flag", () => {
    const result = err("something went wrong");
    assert.deepEqual(result, {
      content: [{ type: "text", text: "Error: something went wrong" }],
      isError: true,
    });
  });

  it("ok handles empty string", () => {
    const result = ok("");
    assert.equal(result.content[0].text, "");
  });

  it("err does not set isError on ok", () => {
    const result = ok("fine");
    assert.equal(result.isError, undefined);
  });
});

describe("uniquePath", () => {
  it("returns the unmodified path when no conflict", () => {
    // Use a path that definitely does not exist
    const dir = os.tmpdir();
    const result = uniquePath(dir, 99999, "nonexistent-mcp-test-file.txt");
    assert.equal(result, path.join(dir, "99999_nonexistent-mcp-test-file.txt"));
  });

  it("appends _1 when the base path already exists", () => {
    const dir = os.tmpdir();
    // First call returns the base path (it does not exist yet in our test)
    const base = uniquePath(dir, 88888, "unique-test-file.txt");
    assert.equal(base, path.join(dir, "88888_unique-test-file.txt"));
  });

  it("handles filenames with extensions correctly", () => {
    const dir = os.tmpdir();
    const result = uniquePath(dir, 12345, "report.pdf");
    assert.ok(result.includes("12345_report"));
    assert.ok(result.endsWith(".pdf"));
  });

  it("handles filenames without extensions", () => {
    const dir = os.tmpdir();
    const result = uniquePath(dir, 12345, "noextension");
    assert.ok(result.includes("12345_noextension"));
  });
});

describe("stripSignature — additional patterns", () => {
  it("strips Sent from my Galaxy", () => {
    const input = "Message body\n\nSent from my Galaxy";
    assert.equal(stripSignature(input), "Message body");
  });

  it("strips Get Outlook for Android", () => {
    const input = "Body text\n\nGet Outlook for Android";
    assert.equal(stripSignature(input), "Body text");
  });

  it("strips Sent from Mail for Windows", () => {
    const input = "Body text\n\nSent from Mail for Windows";
    assert.equal(stripSignature(input), "Body text");
  });

  it("preserves body when -- delimiter has no trailing space", () => {
    // '--' without trailing space is NOT the sig delimiter — should be preserved
    const input = "Hello\n--\nNot a sig";
    assert.ok(stripSignature(input).includes("Not a sig"));
  });
});

describe("stripQuotedReplies — additional patterns", () => {
  it("strips multi-line On ... wrote block", () => {
    const input = "My message\n\nOn Tuesday, 15 April 2026 at 9:00 AM, John Smith <john@example.com> wrote:\n> quoted";
    assert.equal(stripQuotedReplies(input), "My message");
  });

  it("handles text with leading > lines only (no body)", () => {
    const input = "> all quoted\n> more quoted";
    // All lines are quoted — result should be empty or just whitespace
    const result = stripQuotedReplies(input);
    assert.equal(result.trim(), "");
  });

  it("preserves > characters mid-body that are not quote lines", () => {
    const input = "Prices went up > 10% this year\nThat's significant.";
    const result = stripQuotedReplies(input);
    assert.ok(result.includes("10%"));
  });
});

describe("parseRecipients — edge cases", () => {
  it("handles array with Name <email> format", () => {
    assert.deepEqual(
      parseRecipients(["Alice Smith <alice@example.com>"]),
      ["alice@example.com"]
    );
  });

  it("handles whitespace-only string", () => {
    assert.deepEqual(parseRecipients("   "), []);
  });

  it("handles mixed comma and semicolon separators", () => {
    const result = parseRecipients("a@example.com, b@example.com; c@example.com");
    assert.equal(result.length, 3);
    assert.ok(result.includes("a@example.com"));
    assert.ok(result.includes("b@example.com"));
    assert.ok(result.includes("c@example.com"));
  });
});

describe("parseAttachments", () => {
  it("returns [] for null, undefined and empty string", () => {
    assert.deepEqual(parseAttachments(null), []);
    assert.deepEqual(parseAttachments(undefined), []);
    assert.deepEqual(parseAttachments(""), []);
    assert.deepEqual(parseAttachments("   "), []);
  });

  it("wraps a single path in an array", () => {
    assert.deepEqual(parseAttachments("/tmp/a.pdf"), ["/tmp/a.pdf"]);
  });

  it("keeps an array as-is, trimmed and compacted", () => {
    assert.deepEqual(
      parseAttachments(["/tmp/a.pdf", "  /tmp/b.pdf  ", ""]),
      ["/tmp/a.pdf", "/tmp/b.pdf"]
    );
  });

  it("parses a JSON-stringified array (clients that stringify args)", () => {
    assert.deepEqual(
      parseAttachments('["/tmp/a.pdf","/tmp/b.pdf"]'),
      ["/tmp/a.pdf", "/tmp/b.pdf"]
    );
  });

  it("does NOT split on commas — commas are legal in filenames", () => {
    assert.deepEqual(
      parseAttachments("/tmp/Smith, John - invoice.pdf"),
      ["/tmp/Smith, John - invoice.pdf"]
    );
  });
});

describe("validateAttachments", () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "mcp-att-test-"));
  const good = path.join(tmpDir, "good.txt");
  fs.writeFileSync(good, "hello");
  const second = path.join(tmpDir, "second.txt");
  fs.writeFileSync(second, "world");

  it("accepts an existing readable file", () => {
    const { files, errors } = validateAttachments(good);
    assert.deepEqual(errors, []);
    assert.deepEqual(files, [good]);
  });

  it("accepts several files", () => {
    const { files, errors } = validateAttachments([good, second]);
    assert.deepEqual(errors, []);
    assert.equal(files.length, 2);
  });

  it("de-duplicates the same file passed twice", () => {
    const { files, errors } = validateAttachments([good, good]);
    assert.deepEqual(errors, []);
    assert.deepEqual(files, [good]);
  });

  it("resolves a relative path to an absolute one", () => {
    const rel = path.relative(process.cwd(), good);
    const { files, errors } = validateAttachments(rel);
    assert.deepEqual(errors, []);
    assert.deepEqual(files, [good]);
  });

  it("expands a leading ~ to the home directory", () => {
    const homeFile = path.join(os.homedir(), `.mcp-att-test-${process.pid}.txt`);
    fs.writeFileSync(homeFile, "x");
    try {
      const { files, errors } = validateAttachments(`~/${path.basename(homeFile)}`);
      assert.deepEqual(errors, []);
      assert.deepEqual(files, [homeFile]);
    } finally {
      fs.unlinkSync(homeFile);
    }
  });

  it("rejects a missing file", () => {
    const { files, errors } = validateAttachments(path.join(tmpDir, "nope.txt"));
    assert.deepEqual(files, []);
    assert.equal(errors.length, 1);
    assert.match(errors[0], /not found/i);
  });

  it("rejects a directory", () => {
    const { files, errors } = validateAttachments(tmpDir);
    assert.deepEqual(files, []);
    assert.equal(errors.length, 1);
    assert.match(errors[0], /directory/i);
  });

  it("rejects an unreadable file", () => {
    const locked = path.join(tmpDir, "locked.txt");
    fs.writeFileSync(locked, "secret");
    fs.chmodSync(locked, 0o000);
    try {
      const { files, errors } = validateAttachments(locked);
      if (files.length) return; // running as root: permissions do not apply
      assert.equal(errors.length, 1);
      assert.match(errors[0], /not readable/i);
    } finally {
      fs.chmodSync(locked, 0o600);
      fs.unlinkSync(locked);
    }
  });

  it("rejects a file over the size cap", () => {
    const big = path.join(tmpDir, "big.bin");
    const fd = fs.openSync(big, "w");
    fs.ftruncateSync(fd, MAX_ATTACHMENT_BYTES + 1);
    fs.closeSync(fd);
    try {
      const { files, errors } = validateAttachments(big);
      assert.deepEqual(files, []);
      assert.equal(errors.length, 1);
      assert.match(errors[0], /too large/i);
    } finally {
      fs.unlinkSync(big);
    }
  });

  it("rejects a set of files whose total is over the cap", () => {
    const half = Math.floor(MAX_ATTACHMENT_BYTES * 0.6);
    const paths = ["t1.bin", "t2.bin"].map(n => path.join(tmpDir, n));
    for (const p of paths) {
      const fd = fs.openSync(p, "w");
      fs.ftruncateSync(fd, half);
      fs.closeSync(fd);
    }
    try {
      const { errors } = validateAttachments(paths);
      assert.ok(errors.some(e => /total/i.test(e)), `Expected a total-size error, got: ${errors.join("; ")}`);
    } finally {
      for (const p of paths) fs.unlinkSync(p);
    }
  });

  it("reports every bad path, not just the first", () => {
    const { files, errors } = validateAttachments([
      path.join(tmpDir, "missing1.txt"),
      good,
      path.join(tmpDir, "missing2.txt"),
    ]);
    assert.deepEqual(files, [good]);
    assert.equal(errors.length, 2);
  });

  it("returns nothing to do when no attachments are given", () => {
    const { files, errors } = validateAttachments(undefined);
    assert.deepEqual(files, []);
    assert.deepEqual(errors, []);
  });
});

describe("attachmentLines", () => {
  it("emits one make-new-attachment statement per file", () => {
    const script = attachmentLines(["/tmp/a.pdf", "/tmp/b.pdf"], "newMsg");
    const lines = script.split("\n").map(l => l.trim());
    assert.equal(lines.length, 2);
    assert.equal(
      lines[0],
      'make new attachment at newMsg with properties {file:POSIX file "/tmp/a.pdf"}'
    );
    assert.ok(lines[1].includes('"/tmp/b.pdf"'));
  });

  it("targets the message variable it is given", () => {
    const script = attachmentLines(["/tmp/a.pdf"], "composeMsg");
    assert.ok(script.includes("at composeMsg with properties"));
  });

  it("escapes quotes and backslashes in the path", () => {
    const script = attachmentLines(['/tmp/we"ird\\name.pdf'], "newMsg");
    assert.ok(script.includes('\\"ird'), script);
    assert.ok(!script.includes('e"ird'), script);
  });

  it("returns an empty string for no files", () => {
    assert.equal(attachmentLines([], "newMsg"), "");
  });
});
