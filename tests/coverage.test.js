const test = require('node:test');
const assert = require('assert');

// Mock Apps Script environment if needed, though pure logic module shouldn't need much.
// We'll load the logic from Coverage.gs (we need to make sure we exported it for tests).
// Since we are in Google Apps script, Coverage.gs was loaded as global scope. We will just eval it.
const fs = require('fs');
const code = fs.readFileSync('Coverage.gs', 'utf8');
eval(code);

test('Coverage Logic', async (t) => {

  await t.test('parseShiftCell handles valid times', () => {
    assert.deepStrictEqual(parseShiftCell("15:00"), { type: 'shift', startMin: 15 * 60 });
    assert.deepStrictEqual(parseShiftCell("0:00"), { type: 'shift', startMin: 0 });
    assert.deepStrictEqual(parseShiftCell("22:00"), { type: 'shift', startMin: 22 * 60 });
  });

  await t.test('parseShiftCell handles OFF/VL/LOA/SL/blank', () => {
    assert.deepStrictEqual(parseShiftCell("OFF"), { type: 'off' });
    assert.deepStrictEqual(parseShiftCell("VL"), { type: 'vl' });
    assert.deepStrictEqual(parseShiftCell("LOA"), { type: 'loa' });
    assert.deepStrictEqual(parseShiftCell("SL"), { type: 'sl' });
    assert.deepStrictEqual(parseShiftCell(""), { type: 'blank' });
  });

  await t.test('coverageImpactOfRequest handles normal shifts', () => {
    const req = { vlDate: '2026-09-21', site: 'CEB' };
    const ctx = {
      shiftHours: 8,
      intervalMinutes: 60,
      minCoverage: 2,
      schedule: {
        people: [
          { email: 'a@b.com', site: 'CEB', days: { '2026-09-21': { type: 'shift', startMin: 15 * 60 } } }
        ]
      }
    };
    // Expected impact: intervals 15, 16, 17, 18, 19, 20, 21, 22 on 2026-09-21
    const impact = coverageImpactOfRequest(req, ctx);
    assert.strictEqual(impact.site, 'CEB');
    assert.strictEqual(impact.intervalsAffected.length, 8);
    assert.strictEqual(impact.intervalsAffected[0].intervalStart, '3:00 PM');
  });

  await t.test('coverageImpactOfRequest handles midnight crossing shifts (rollover)', () => {
    const req = { vlDate: '2026-09-21', site: 'CEB' };
    const ctx = {
      shiftHours: 8,
      intervalMinutes: 60,
      minCoverage: 2,
      schedule: {
        people: [
          { email: 'a@b.com', site: 'CEB', days: { '2026-09-21': { type: 'shift', startMin: 22 * 60 } } }
        ]
      }
    };
    // Expected impact: intervals 22, 23 on 2026-09-21, and 0, 1, 2, 3, 4, 5 on 2026-09-22
    const impact = coverageImpactOfRequest(req, ctx);
    assert.strictEqual(impact.intervalsAffected.length, 8);
    const starts = impact.intervalsAffected.map(i => i.date + " " + i.intervalStart);
    assert.ok(starts.includes('2026-09-21 10:00 PM'));
    assert.ok(starts.includes('2026-09-22 12:00 AM'));
    assert.ok(starts.includes('2026-09-22 5:00 AM'));
  });

});
