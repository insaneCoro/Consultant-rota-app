// Offline regression tests for storage safety and the leave-amendment path.
// Usage: TZ=Asia/Riyadh node apps-script/tests/run.js
'use strict';
const assert = require('assert');
const { createEnv, seed } = require('./harness');

const tests = [];
const test = (name, fn) => tests.push({ name, fn });
const SA = 'RotaAssignments';

function liveRows(env) {
  const sh = env.__db.getSheetByName(SA);
  const [h, ...rows] = sh.data;
  return rows.filter(r => r.some(v => v !== '' && v !== undefined)).map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function assignment(versionId, weekNo, place, id) {
  return { VersionID: versionId, WeekNo: weekNo, WeekStart: '2026-07-05', WeekEnd: '2026-07-11', TeamID: 'T1',
    AssignmentPlace: place, ConsultantID: id, Source: 'Auto', Notes: '', PeriodID: 'P1' };
}
function baseEnv() {
  const env = createEnv();
  seed(env, 'Teams', [{ TeamID: 'T1', TeamName: 'Medicine', Active: 'Yes', TeamType: 'SubTeam', SubTeamConfig: 'TCU,EMCU', SubTeamFrom: '' }]);
  seed(env, 'Consultants', ['A', 'B', 'C', 'D'].map(x => ({ ConsultantID: x, Name: x, TeamID: 'T1', CanOffDuty: 'Yes',
    ConsultantType: 'Regular', CurrentStatus: 'Active', ActiveFrom: '2026-01-01', SubTeamAssignment: 'Any' })));
  ['Settings', 'EligibilityHistory', 'LeaveRequests', 'ConflictLog', 'OverrideLog', 'AuditLog', 'AdminUsers'].forEach(n => seed(env, n, []));
  seed(env, 'RotaPeriods', [{ PeriodID: 'P1', PeriodName: '2026/27', Status: 'Active', StartDate: '2026-07-05', EndDate: '2027-07-03', Weeks: 52 }]);
  seed(env, 'RotaVersions', [
    { VersionID: 'V1', AmendmentNo: 1, Status: 'Archived', PeriodID: 'P1', StartDate: '2026-07-05', EndDate: '2027-07-03', Weeks: 52 },
    { VersionID: 'V2', AmendmentNo: 2, Status: 'Official', PeriodID: 'P1', StartDate: '2026-07-05', EndDate: '2027-07-03', Weeks: 52 }]);
  return env;
}

test('U1: removing a version deletes only its rows, even when interleaved with blank spacer rows', () => {
  const env = baseEnv();
  seed(env, SA, [assignment('V1', 1, 'TCU', 'A'), assignment('V2', 1, 'TCU', 'B'), assignment('V1', 2, 'OPD', 'C'),
    {}, assignment('V2', 2, 'EMCU', 'D'), assignment('V1', 3, 'Off-duty', 'D')]);
  const removed = env.removeVersionAssignmentsFromLive_('V1');
  assert.strictEqual(removed, 3);
  const rows = liveRows(env);
  assert.deepStrictEqual(rows.map(r => r.VersionID + r.ConsultantID), ['V2B', 'V2D']);
});

test('U1: a Sheets failure part-way never removes rows of other versions', () => {
  const env = baseEnv();
  seed(env, SA, [assignment('V1', 1, 'TCU', 'A'), assignment('V2', 1, 'TCU', 'B'), assignment('V1', 2, 'OPD', 'C'),
    assignment('V2', 2, 'EMCU', 'D'), assignment('V1', 3, 'OPD', 'D')]);
  env.__db.writeBudget = 1; // first block delete succeeds, second fails
  assert.throws(() => env.removeVersionAssignmentsFromLive_('V1'), /Simulated/);
  const v2 = liveRows(env).filter(r => r.VersionID === 'V2');
  assert.strictEqual(v2.length, 2, 'official rows must survive');
});

test('U2: an unreadable archive file stops archiving and keeps the stored archive link', () => {
  const env = baseEnv();
  env.__files.OLDARCH = { getId: () => 'OLDARCH' };
  seed(env, 'RotaPeriods', [{ PeriodID: 'P1', PeriodName: '2026/27', Status: 'Active', StartDate: '2026-07-05',
    EndDate: '2027-07-03', Weeks: 52, ArchiveSpreadsheetID: 'OLDARCH', ArchiveSpreadsheetURL: 'u' }]);
  seed(env, SA, [assignment('V1', 1, 'TCU', 'A'), assignment('V2', 1, 'TCU', 'B')]);
  env.__openFails = true;
  assert.throws(() => env.archiveVersionAssignments_('V1'), /could not be opened/);
  const period = env.readObjects_('RotaPeriods')[0];
  assert.strictEqual(period.ArchiveSpreadsheetID, 'OLDARCH');
  assert.strictEqual(liveRows(env).length, 2, 'live rows untouched');
  assert.strictEqual(Object.keys(env.__files).filter(k => k.startsWith('ARCH')).length, 0, 'no replacement archive created');
});

test('U2: a period without an archive still gets one created and linked', () => {
  const env = baseEnv();
  seed(env, SA, [assignment('V1', 1, 'TCU', 'A'), assignment('V2', 1, 'TCU', 'B')]);
  const r = env.archiveVersionAssignments_('V1');
  assert.strictEqual(r.archived, true);
  assert.ok(env.readObjects_('RotaPeriods')[0].ArchiveSpreadsheetID.startsWith('ARCH'));
  assert.deepStrictEqual(liveRows(env).map(x => x.VersionID), ['V2']);
});

test('Restore: a partially moved version is rebuilt from its complete archive copy', () => {
  const env = baseEnv();
  seed(env, SA, [assignment('V1', 1, 'TCU', 'A'), assignment('V1', 2, 'OPD', 'C'), assignment('V2', 1, 'TCU', 'B'), assignment('V1', 3, 'OPD', 'D')]);
  // Simulate: archive written, then deletion interrupted after the bottom block.
  const archive = env.getOrCreatePeriodArchive_('P1');
  const version = env.readObjects_('RotaVersions')[0];
  env.writeArchiveSheet_(archive, env.assignmentArchiveSheetName_(version), liveRows(env).filter(r => r.VersionID === 'V1'));
  env.__db.writeBudget = 1;
  assert.throws(() => env.removeVersionAssignmentsFromLive_('V1'));
  env.__db.writeBudget = Infinity;
  assert.strictEqual(liveRows(env).filter(r => r.VersionID === 'V1').length, 2, 'partial state reproduced');
  const restored = env.restoreVersionAssignmentsToLive_('V1');
  assert.strictEqual(restored.length, 3);
  const v1 = liveRows(env).filter(r => r.VersionID === 'V1').map(r => r.WeekNo).sort();
  assert.deepStrictEqual(v1, [1, 2, 3], 'complete, without duplicates');
  assert.strictEqual(liveRows(env).filter(r => r.VersionID === 'V2').length, 1);
});

test('Migration: archiving history keeps Official/Draft rows exactly and removes only archived versions', () => {
  const env = baseEnv();
  seed(env, SA, [assignment('V1', 1, 'TCU', 'A'), assignment('V2', 1, 'TCU', 'B'), assignment('V1', 2, 'OPD', 'C'), assignment('V2', 2, 'OPD', 'D')]);
  const r = env.migrateAssignmentStorage_();
  assert.strictEqual(r.archivedVersions, 1);
  assert.deepStrictEqual(liveRows(env).map(x => x.VersionID + x.ConsultantID), ['V2B', 'V2D']);
});

test('U3: Clear Generated Rota is refused unless ALLOW_ROTA_RESET=YES', () => {
  const env = baseEnv();
  env.requireAdminPin_ = () => ({ isMainAdmin: true, name: 'Admin' });
  env.authenticateAdmin_ = env.requireAdminPin_;
  env.getAdminData = () => ({});
  seed(env, SA, [assignment('V2', 1, 'TCU', 'B')]);
  assert.throws(() => env.resetRotaData('t', 'RESET ROTA'), /disabled on this deployment/);
  assert.strictEqual(liveRows(env).length, 1);
  assert.strictEqual(env.readObjects_('RotaVersions').length, 2);
  env.PropertiesService.getScriptProperties().setProperty('ALLOW_ROTA_RESET', 'YES');
  env.resetRotaData('t', 'RESET ROTA');
  assert.strictEqual(liveRows(env).length, 0, 'staging can still reset when explicitly enabled');
});

test('End-to-end: approved leave amendment archives the old official, preserves past weeks, recalculates future', () => {
  const env = baseEnv();
  seed(env, 'RotaVersions', [{ VersionID: 'V2', AmendmentNo: 1, Status: 'Official', PeriodID: 'P1', StartDate: '2026-07-05', EndDate: '2027-07-03', Weeks: 52 }]);
  const base = env.buildRotaRows_({ versionId: 'V2', startDate: env.toDate_('2026-07-05'), weeks: 52, fromWeekNo: 1, previousRows: [], teamIds: ['T1'] });
  seed(env, SA, base.rows.map(r => Object.assign(r, { PeriodID: 'P1' })));
  seed(env, 'LeaveRequests', [{ LeaveID: 'L1', ConsultantID: 'A', TeamID: 'T1', StartDate: '2027-02-07', EndDate: '2027-02-20', Status: 'Approved' }]);
  const result = env.createAmendmentFromDate_('2027-02-07', 'Leave amendment', 'test', null, ['T1'], '2027-02-20');
  assert.strictEqual(result.amendmentCreated, true);
  const versions = env.readObjects_('RotaVersions');
  assert.strictEqual(versions.find(v => v.VersionID === 'V2').Status, 'Archived');
  const newId = result.version.VersionID;
  const live = liveRows(env);
  assert.ok(live.every(r => r.VersionID === newId), 'live sheet holds only the new official version');
  const cutoff = env.getWeekNoForDate_(env.toDate_('2026-07-05'), env.toDate_('2027-02-07'), 52);
  const sig = rows => rows.filter(r => Number(r.WeekNo) < cutoff).map(r => [r.WeekNo, r.AssignmentPlace, r.ConsultantID].join('|')).sort().join(';');
  assert.strictEqual(sig(live), sig(base.rows), 'weeks before the effective week are unchanged');
  const leaveWeeks = live.filter(r => r.ConsultantID === 'A' && Number(r.WeekNo) >= cutoff && Number(r.WeekNo) <= cutoff + 1);
  assert.ok(leaveWeeks.every(r => r.AssignmentPlace === 'Leave'), 'consultant on leave is not rostered');
  for (let w = cutoff; w <= 52; w++) ['TCU', 'EMCU'].forEach(role =>
    assert.ok(live.some(r => Number(r.WeekNo) === w && r.AssignmentPlace === role), `week ${w} ${role} covered`));
  const archived = env.readArchivedAssignments_(env.readObjects_('RotaVersions').find(v => v.VersionID === 'V2'));
  assert.strictEqual(archived.length, base.rows.length, 'old official fully archived');
});

// ---- What-if leave check ----
function whatIfEnv() {
  const env = baseEnv();
  env.requireLeaveAdmin_ = () => ({ isMainAdmin: true });
  seed(env, 'RotaVersions', [{ VersionID: 'V2', AmendmentNo: 1, Status: 'Official', PeriodID: 'P1', StartDate: '2026-07-05', EndDate: '2027-07-03', Weeks: 52 }]);
  const base = env.buildRotaRows_({ versionId: 'V2', startDate: env.toDate_('2026-07-05'), weeks: 52, fromWeekNo: 1, previousRows: [], teamIds: ['T1'] });
  seed(env, SA, base.rows.map(r => Object.assign(r, { PeriodID: 'P1' })));
  return env;
}
const snapshot = env => JSON.stringify(Object.keys(env.__db.sheets).sort().map(k => [k, env.__db.sheets[k].data]));

test('What-if: one consultant away -> 3 available, posts covered, nothing written', () => {
  const env = whatIfEnv();
  const before = snapshot(env);
  const r = env.previewLeaveWhatIf('t', { ConsultantID: 'A', StartDate: '2027-02-07', EndDate: '2027-02-20' });
  assert.strictEqual(snapshot(env), before, 'no sheet may change');
  assert.strictEqual(r.whatIf, true);
  const away = r.outlook.weeks.filter(w => w.onLeave.includes('A'));
  assert.strictEqual(away.length, 2);
  assert.ok(away.every(w => w.available === 3 && w.inpatient.length === 2 && !w.uncovered.length));
  assert.ok(r.outlook.lines.some(l => /3 consultants available/.test(l.text)));
  assert.ok(r.outlook.lines.some(l => /A returns in week/.test(l.text)));
  assert.ok(r.preview.length === r.outlook.weeks.length && r.preview.length >= 3, 'preview runs through the return week');
});

test('What-if: second overlapping leave -> only 2 left is called out', () => {
  const env = whatIfEnv();
  seed(env, 'LeaveRequests', [{ LeaveID: 'L1', ConsultantID: 'B', TeamID: 'T1', StartDate: '2027-02-07', EndDate: '2027-02-20', Status: 'Approved' }]);
  const r = env.previewLeaveWhatIf('t', { ConsultantID: 'A', StartDate: '2027-02-07', EndDate: '2027-02-13' });
  const line = r.outlook.lines.find(l => l.tone === 'bad');
  assert.ok(line && /only 2 consultant/.test(line.text), 'two-left warning present');
  assert.ok(r.outlook.lines.some(l => /Back-to-back inpatient/.test(l.text)));
});

test('What-if: dates in already-started weeks or outside the rota change nothing and say so', () => {
  const env = whatIfEnv();
  const past = env.previewLeaveWhatIf('t', { ConsultantID: 'A', StartDate: '2026-07-06', EndDate: '2026-07-08' });
  assert.strictEqual(past.outlook, null); assert.match(past.summary, /already started/);
  const later = env.previewLeaveWhatIf('t', { ConsultantID: 'A', StartDate: '2027-09-01', EndDate: '2027-09-05' });
  assert.strictEqual(later.outlook, null); assert.match(later.summary, /after the current rota/);
  assert.throws(() => env.previewLeaveWhatIf('t', { ConsultantID: 'A', StartDate: '2027-02-10', EndDate: '2027-02-01' }), /before start/);
  assert.throws(() => env.previewLeaveWhatIf('t', { ConsultantID: 'ZZ', StartDate: '2027-02-01', EndDate: '2027-02-02' }), /Select a consultant/);
});

test('What-if result equals what a real approval then publishes', () => {
  const env = whatIfEnv();
  const w = env.previewLeaveWhatIf('t', { ConsultantID: 'C', StartDate: '2027-03-07', EndDate: '2027-03-20' });
  seed(env, 'LeaveRequests', [{ LeaveID: 'L9', ConsultantID: 'C', TeamID: 'T1', StartDate: '2027-03-07', EndDate: '2027-03-20', Status: 'Approved' }]);
  env.createAmendmentFromDate_('2027-03-07', 'Leave amendment', 'test', null, ['T1'], '2027-03-20');
  const live = liveRows(env);
  w.preview.forEach(p => ['TCU', 'EMCU', 'OPD', 'Off-duty', 'Leave'].forEach(place => {
    const real = live.filter(r => Number(r.WeekNo) === p.weekNo && r.AssignmentPlace === place).map(r => r.ConsultantID).sort().join(',');
    assert.strictEqual((p[place] || []).map(x => x.id).sort().join(','), real, `week ${p.weekNo} ${place}`);
  }));
});

test('Leave preview for a real request now carries the same outlook', () => {
  const env = whatIfEnv();
  seed(env, 'LeaveRequests', [{ LeaveID: 'L5', ConsultantID: 'D', TeamID: 'T1', StartDate: '2027-04-04', EndDate: '2027-04-10', Status: 'Pending' }]);
  const r = env.previewLeaveImpact('t', 'L5');
  assert.ok(r.outlook && r.outlook.weeks.some(w => w.onLeave.includes('D')));
  assert.ok(Array.isArray(r.preview) && r.preview.length > 0, 'existing preview field kept');
});

test('API: previewLeaveWhatIf is reachable through doPost and rejects a missing session', () => {
  const env = baseEnv();
  const out = JSON.parse(env.doPost({ postData: { contents: JSON.stringify({ fn: 'previewLeaveWhatIf', args: ['', { ConsultantID: 'A', StartDate: '2027-02-07', EndDate: '2027-02-08' }] }) } }));
  assert.strictEqual(out.ok, false); assert.match(out.error, /sign in again/i);
});

let failed = 0;
tests.forEach(t => {
  try { t.fn(); console.log('PASS ' + t.name); }
  catch (e) { failed++; console.log('FAIL ' + t.name + '\n     ' + (e && e.stack || e).split('\n').slice(0, 3).join('\n     ')); }
});
console.log(`\n${tests.length - failed}/${tests.length} passed`);
process.exit(failed ? 1 : 0);
