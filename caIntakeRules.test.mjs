// The CA shared-rules block in taskpane.js mirrors SettyPMS.html (numbering,
// review periods, stamp → status) and the Tracker's record naming for portal
// items. This slices that block out of the file and checks it, so a drift
// shows up here before it shows up as a duplicate RFI number.
//
//   node --test caIntakeRules.test.mjs
import { strict as assert } from "node:assert";
import test from "node:test";
import { readFileSync } from "node:fs";

const src = readFileSync(new URL("./taskpane.js", import.meta.url), "utf8");
const a = src.indexOf("// ── CA SHARED RULES (begin)"), b = src.indexOf("// ── CA SHARED RULES (end)");
assert.ok(a > 0 && b > a, "shared-rules markers must exist in taskpane.js");
const rules = new Function(src.slice(a, b) + "\nreturn { caNumberValue, caNextNumber, caBumpHighWater, caDueDays, subStatusForStamp, portalItemNumber, parseEmailList, caDistributionFor, addWorkdaysISO, subDueDefault };")();

test("next number is highest-ever + 1, never count + 1", () => {
  const rfis = [{ number: "RFI-001" }, { number: "RFI-003" }];           // RFI-002 was deleted
  assert.equal(rules.caNextNumber(rfis, "rfi"), "RFI-004");
  assert.equal(rules.caNextNumber([], "rfi"), "RFI-001");
  assert.equal(rules.caNextNumber(rfis, "rfi", 7), "RFI-008");          // high-water mark remembers a voided RFI-007
  assert.equal(rules.caNextNumber([{ number: "SUB-012" }], "submittal"), "SUB-013");
  // A Procore number in the log pushes the hand numbers past it, as the PMS does.
  assert.equal(rules.caNextNumber([{ number: "RFI-2218" }], "rfi"), "RFI-2219");
});

test("number value reads the trailing digits the way the PMS does", () => {
  assert.equal(rules.caNumberValue("RFI-091"), 91);
  assert.equal(rules.caNumberValue("RFI-121.1"), 121);
  assert.equal(rules.caNumberValue(""), null);
});

test("high-water mark only ever rises", () => {
  const p = { caNumberHighWater: { rfi: 9 } };
  assert.deepEqual(rules.caBumpHighWater(p, "rfi", "RFI-004"), { rfi: 9 });
  assert.deepEqual(rules.caBumpHighWater(p, "rfi", "RFI-012"), { rfi: 12 });
  assert.deepEqual(rules.caBumpHighWater({}, "submittal", "SUB-003"), { submittal: 3 });
});

test("review period comes from the project, with the firm defaults", () => {
  assert.equal(rules.caDueDays({ rfiDueDays: 7, submittalDueDays: 14 }, "rfi"), 7);
  assert.equal(rules.caDueDays({ rfiDueDays: 7, submittalDueDays: 14 }, "submittal"), 14);
  assert.equal(rules.caDueDays({}, "rfi"), 5);
  assert.equal(rules.caDueDays(null, "submittal"), 10);
});

test("a stamp that sends the submittal back sets Resubmittal Required", () => {
  assert.equal(rules.subStatusForStamp("Approved"), "Returned");
  assert.equal(rules.subStatusForStamp("Approved as Noted"), "Returned");
  assert.equal(rules.subStatusForStamp("Revise and Resubmit"), "Resubmittal Required");
  assert.equal(rules.subStatusForStamp("Rejected"), "Resubmittal Required");
});

test("portal items are numbered the way the Tracker's import names them", () => {
  assert.equal(rules.portalItemNumber("rfi", { extNum: "2218" }), "RFI-2218");
  assert.equal(rules.portalItemNumber("rfi", { extNum: "91" }), "RFI-091");
  assert.equal(rules.portalItemNumber("rfi", { extNum: "121.1" }), "RFI-121.1");
  assert.equal(rules.portalItemNumber("submittal", { extNum: "233400-001", rev: 1 }), "233400-001-1");
  assert.equal(rules.portalItemNumber("submittal", { extNum: "233400-001" }), "233400-001-0");
  assert.equal(rules.portalItemNumber("submittal", { extNum: "12" }), "SUB-012");
});

test("email lists parse the way the PMS distribution editor writes them", () => {
  assert.deepEqual(rules.parseEmailList("A@gc.com; b@owner.org , A@gc.com"), ["a@gc.com", "b@owner.org"]);
  assert.deepEqual(rules.parseEmailList(""), []);
});

test("distribution: contractor first in To, project list in Cc, no address twice", () => {
  const project = { caDistribution: { rfi: { to: "pm@owner.org", cc: "sub@mech.com; gc@gc.com" } } };
  assert.deepEqual(rules.caDistributionFor(project, "rfi", ["gc@gc.com"]), { to: ["gc@gc.com", "pm@owner.org"], cc: ["sub@mech.com"] });
  assert.deepEqual(rules.caDistributionFor({}, "submittal", ["gc@gc.com"]), { to: ["gc@gc.com"], cc: [] });
});

test("sub due date: 2 working days before the contractor's, never before tomorrow", () => {
  assert.equal(rules.addWorkdaysISO("2026-11-06", -2), "2026-11-04");       // Fri → Wed
  assert.equal(rules.addWorkdaysISO("2026-11-09", -2), "2026-11-05");       // Mon → Thu (skips the weekend)
  assert.equal(rules.subDueDefault("2026-11-12", "2026-11-03"), "2026-11-10");
  assert.equal(rules.subDueDefault("2026-11-04", "2026-11-03"), "2026-11-04"); // floor: next working day
  assert.equal(rules.subDueDefault("", "2026-11-06"), "2026-11-09");          // no contractor date → next working day (Mon)
});
