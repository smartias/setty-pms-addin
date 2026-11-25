// "Save as accolade" in the Outlook pane: the quote trimmed from the email, the
// first guess at who earned it, and typed names matched to the staff list.
//
//   node --test saveAccolade.test.mjs
import { strict as assert } from "node:assert";
import test from "node:test";
import { readFileSync } from "node:fs";

const src = readFileSync(new URL("./taskpane.js", import.meta.url), "utf8").replace(/\r\n/g, "\n");
function fn(name) {
  const i = src.indexOf("function " + name + "(");
  assert.ok(i > 0, name + " not found in taskpane.js");
  return src.slice(i, src.indexOf("\n}", i) + 2);
}
const names = ["accoladeQuoteFrom", "accoladeGuessPeople", "accoladeResolvePeople"];
const code = names.map(fn).join("\n") + "\nexport { " + names.join(", ") + " };";
const { accoladeQuoteFrom, accoladeGuessPeople, accoladeResolvePeople } =
  await import("data:text/javascript;base64," + Buffer.from(code, "utf8").toString("base64"));

const staff = [
  { name: "Ana Diaz", email: "Ana@setty.com" },
  { name: "Bo Chen", email: "bo@setty.com" },
  { name: "Old Timer", email: "old@setty.com", inactive: true },
];

test("the quote is the newest message only, whitespace collapsed", () => {
  const body = "Hi Ana,\r\n\r\nThe drawings were  excellent.\r\nThanks!\r\n\r\nFrom: Ana Diaz\r\nSent: Monday\r\nSubject: drawings\r\nolder text";
  assert.equal(accoladeQuoteFrom(body), "Hi Ana, The drawings were excellent. Thanks!");
  assert.equal(accoladeQuoteFrom("Great work.\nOn Mon, Nov 23, 2026 Ana wrote:\n> earlier"), "Great work.");
  assert.equal(accoladeQuoteFrom("Nice job\n-----Original Message-----\nold"), "Nice job");
  assert.equal(accoladeQuoteFrom("> quoted only line\nreal line"), "real line");
  assert.equal(accoladeQuoteFrom(""), "");
  assert.equal(accoladeQuoteFrom("x".repeat(1500)).length, 1198);
});

test("the guess is staff on the To line, not the sender, then Cc, then you", () => {
  const to = [{ emailAddress: "client@acme.com" }, { emailAddress: "ANA@setty.com" }, { emailAddress: "old@setty.com" }];
  assert.deepEqual(accoladeGuessPeople(to, [], "client@acme.com", staff, { name: "Me" }), ["Ana Diaz"]);
  assert.deepEqual(accoladeGuessPeople([{ emailAddress: "bo@setty.com" }], [], "bo@setty.com", staff, { name: "Me" }), ["Me"]);
  assert.deepEqual(accoladeGuessPeople([{ emailAddress: "x@acme.com" }], [{ emailAddress: "bo@setty.com" }], "c@acme.com", staff, {}), ["Bo Chen"]);
  assert.deepEqual(accoladeGuessPeople(null, null, "", undefined, {}), []);
});

test("typed names match the staff list; others are kept without an email", () => {
  assert.deepEqual(accoladeResolvePeople(" ana diaz, Pat Client; Ana Diaz", staff),
    [{ name: "Ana Diaz", email: "ana@setty.com" }, { name: "Pat Client", email: "" }]);
  assert.deepEqual(accoladeResolvePeople("", staff), []);
});
