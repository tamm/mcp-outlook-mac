import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  extractEmail,
  parseRecipients,
  recipientLines,
  escapeForAppleScript,
  stripSignature,
  stripQuotedReplies,
  cleanBody,
  markdownToHtml,
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
    assert.ok(result.includes("<p>hello world</p>"));
    assert.ok(result.includes("font-family"));
  });
});
