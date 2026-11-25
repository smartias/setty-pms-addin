// "Hi <first name>," in the add-in's assignment emails. The helper in taskpane.js
// mirrors greetingFirstName in setty-pms (greetingName.js); this slices the real
// code out of taskpane.js and checks it, so a drift shows up here.
//
//   node --test greetingName.test.mjs
import { strict as assert } from "node:assert";
import test from "node:test";
import { readFileSync } from "node:fs";

const src = readFileSync(new URL("./taskpane.js", import.meta.url), "utf8").replace(/\r\n/g, "\n");

// The marked helper block.
const a = src.indexOf("// ── GREETING NAME (begin)"), b = src.indexOf("// ── GREETING NAME (end)");
assert.ok(a > 0 && b > a, "greeting-name markers must exist in taskpane.js");
const block = src.slice(a, b);

// A top-level function, cut from its declaration to the first "}" at line start.
function fn(name) {
  const i = src.indexOf("function " + name + "(");
  assert.ok(i > 0, name + " not found in taskpane.js");
  const j = src.indexOf("\n}", i);
  return src.slice(i, j + 2);
}

const { greetingFirstName, resolveAssignee, assigneeRecipients } = new Function(
  block + "\n" + fn("resolveAssignee") + "\n" + fn("assigneeRecipients") +
  "\nreturn { greetingFirstName, resolveAssignee, assigneeRecipients };")();

// Same contact shapes the PMS tests use (setty-pms greetingName.test.mjs).
const CASES = [
  ["Rama, Monica", "Monica"],
  ["Bermudez , James", "James"],
  ["Millard Jensen , Gene", "Gene"],
  ["Schenne , John A.", "John"],
  ["Monica Rama", "Monica"],
  ["John Smith, PE", "John"],
  ["John Smith, P.E.", "John"],
  ["Kurt Bedore, Jr.", "Kurt"],
  ["Jane Doe, PhD", "Jane"],
  ["Smith, John, PE", "John"],
  ["Dr. Ann Lee", "Ann"],
  ["Rama,", "Rama"],
  ["Rama, Monica <monica@example.com>", "Monica"],
  ["monica@example.com", ""],
  ["", ""],
  [null, ""],
  [undefined, ""],
];

test("greets contact shapes by the given name", () => {
  for (const [name, want] of CASES) assert.equal(greetingFirstName(name), want, JSON.stringify(name));
});

const project = {
  teamMembers: [
    { id: "t1", name: "Sara Arias", email: "sara@setty.com", role: "Electrical Engineer" },
    { id: "t2", name: "Arias, Sara", email: "sara2@setty.com" },
    { id: "t3", email: "role@setty.com", role: "Mechanical Engineer" },   // no name on file
  ],
  subconsultants: [
    { id: "s1", firm: "Acme Engineering", contact: "Rama, Monica", email: "m@acme.com" },
    { id: "s2", firm: "Acme Engineering", email: "info@acme.com" },        // firm only, no contact
  ],
};

test("a staff assignee is greeted by first name, however the name is stored", () => {
  assert.equal(resolveAssignee("staff:t1", project).first, "Sara");
  assert.equal(resolveAssignee("staff:t2", project).first, "Sara");
  const recips = assigneeRecipients(["t1", "t2"], project);
  assert.deepEqual(recips.map(r => r.first), ["Sara", "Sara"]);
});

test("a subconsultant contact is greeted by first name", () => {
  assert.equal(resolveAssignee("sub:s1", project).first, "Monica");
});

test("a role or firm name is not a person: no first name, so the email says Hi,", () => {
  // name still falls back to the role/firm for the 'Assigned to …' status line…
  assert.equal(resolveAssignee("staff:t3", project).name, "Mechanical Engineer");
  assert.equal(resolveAssignee("sub:s2", project).name, "Acme Engineering");
  // …but is never used to greet.
  assert.equal(resolveAssignee("staff:t3", project).first, "");
  assert.equal(resolveAssignee("sub:s2", project).first, "");
  assert.equal(assigneeRecipients(["t3"], project)[0].first, "");
});

test("both assignment emails greet assignee.first, not the full name", () => {
  for (const name of ["buildRfiAssignmentEmailHtml", "buildSubAssignmentEmailHtml"]) {
    const body = fn(name);
    assert.match(body, /<p>Hi\$\{assignee\?\.first \? " " \+ esc\(assignee\.first\) : ""\},<\/p>/, name + " greets by first name");
    assert.doesNotMatch(body, /Hi \$\{esc\(assignee\?\.name/, name + " no longer greets by full name");
  }
});

test("the 'Complete <name>'s contact' button label uses the first name", () => {
  assert.match(src, /const who = greetingFirstName\(enrichRec\.name\);/);
});
