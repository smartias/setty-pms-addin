// Staff photos in the Outlook pane: which directory photo a project team row
// (a copy with no photoPath of its own) resolves to. Same cases as
// staffPhotos.test.mjs in setty-pms, run against the copy in taskpane.js.
//
//   node --test staffPhotos.test.mjs
import { strict as assert } from "node:assert";
import test from "node:test";
import { readFileSync } from "node:fs";

const src = readFileSync(new URL("./taskpane.js", import.meta.url), "utf8").replace(/\r\n/g, "\n");
function fn(name) {
  const i = src.indexOf("function " + name + "(");
  assert.ok(i > 0, name + " not found in taskpane.js");
  return src.slice(i, src.indexOf("\n}", i) + 2);
}
const code = ["buildStaffPhotoIndex", "staffPhotoPathFor", "staffInitials"].map(fn).join("\n") +
  "\nlet _staffPhotoIndex = { byEmail: new Map(), byName: new Map() };\nexport { buildStaffPhotoIndex, staffPhotoPathFor, staffInitials };";
const { buildStaffPhotoIndex, staffPhotoPathFor, staffInitials } =
  await import("data:text/javascript;base64," + Buffer.from(code, "utf8").toString("base64"));

const index = buildStaffPhotoIndex([
  { id: "a1", name: "Jane Smith", email: "JSmith@setty.com", photoPath: "a1/x.jpg" },
  { id: "b2", name: "Sam Lee", email: "", photoPath: "b2/y.jpg" },
  { id: "c3", name: "No Photo", email: "np@setty.com" },
  null,
]);

test("a team row matches the directory by email, then by name", () => {
  assert.equal(staffPhotoPathFor({ name: "Anyone", photoPath: "z/z.jpg" }, index), "z/z.jpg");
  assert.equal(staffPhotoPathFor({ name: "J. Smith", email: " jsmith@SETTY.com " }, index), "a1/x.jpg");
  assert.equal(staffPhotoPathFor({ name: "  jane smith" }, index), "a1/x.jpg");
  assert.equal(staffPhotoPathFor({ name: "Sam Lee", email: "slee@setty.com" }, index), "b2/y.jpg");
});
test("a same-named person with a different email gets no photo", () => {
  assert.equal(staffPhotoPathFor({ name: "Jane Smith", email: "jane.smith@other.com" }, index), "");
});
test("no photo, no person", () => {
  assert.equal(staffPhotoPathFor({ name: "No Photo", email: "np@setty.com" }, index), "");
  assert.equal(staffPhotoPathFor(null, index), "");
  assert.equal(staffPhotoPathFor({ name: "Jane Smith" }), "");
});
test("initials", () => {
  assert.equal(staffInitials("  jane  van Smith "), "JV");
  assert.equal(staffInitials(""), "");
});
