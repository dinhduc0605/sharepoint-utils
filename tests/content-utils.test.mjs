import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const source = await readFile(new URL("../content-utils.js", import.meta.url), "utf8");
const context = { encodeURIComponent: value => globalThis.encodeURIComponent(value), Intl };
vm.runInNewContext(source, context, { filename: "content-utils.js" });
const utils = context.SharePointUtils;

test("normalizes punctuation and repeated whitespace in search terms", () => {
    assert.equal(utils.normalizeSearchInput("  Q1_report-2024  final  "), "Q1 report 2024 final");
});

test("matches search terms using Unicode normalization", () => {
    assert.equal(utils.matchesAllSearchTerms("ＡＢＣ レポート", ["abc", "レポート"]), true);
    assert.equal(utils.matchesAllSearchTerms("ABC notes", ["report"]), false);
});

test("escapes SharePoint OData values", () => {
    assert.equal(utils.encodeODataString("O'Reilly"), "O''Reilly");
    assert.equal(utils.encodeSharePointPath("/sites/foo/Shared Documents"), "%2Fsites%2Ffoo%2FShared%20Documents");
});

test("maps checkout enum values correctly", () => {
    assert.equal(utils.isFileCheckedOut(0), true);
    assert.equal(utils.isFileCheckedOut(1), true);
    assert.equal(utils.isFileCheckedOut(2), false);
    assert.equal(utils.isFileCheckedOut("invalid"), false);
});

test("categorizes files and chooses desktop protocols", () => {
    assert.equal(utils.getFileTypeCategory("plan.xlsx"), "excel");
    assert.equal(utils.getFileTypeCategory("brief.pdf"), "pdf");
    assert.equal(utils.getOfficeDesktopProtocol("brief.docx"), "ms-word:ofe|u|");
    assert.equal(utils.getOfficeDesktopProtocol("image.png"), "");
});

test("escapes HTML and regular expression text", () => {
    assert.equal(utils.escapeHtml("<a href='x'>&"), "&lt;a href=&#39;x&#39;&gt;&amp;");
    const regex = new RegExp(utils.escapeRegex("Q1.2024"), "i");
    assert.equal(regex.test("Q1.2024"), true);
    assert.equal(regex.test("Q1x2024"), false);
});

test("neutralizes KQL syntax in search terms", () => {
    assert.equal(utils.toKqlSearchTerm("Title:secret"), "Titlesecret");
    assert.equal(utils.toKqlSearchTerm("report*"), "report");
    assert.equal(utils.toKqlSearchTerm("OR"), "or");
    assert.equal(utils.toKqlSearchTerm("NOT"), "not");
    assert.equal(utils.toKqlSearchTerm("Oracle"), "Oracle");
    assert.equal(utils.toKqlSearchTerm("\"()\\"), "");
});
