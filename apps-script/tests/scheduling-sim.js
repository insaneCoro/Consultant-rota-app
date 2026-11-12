// Generates full 52-week sub-team rotas offline and reports fairness and
// TCU/EMCU alternation metrics for several staffing scenarios.
// Usage: TZ=Asia/Riyadh node apps-script/tests/scheduling-sim.js
'use strict';
const { createEnv, seed } = require('./harness');

const START = '2026-07-05'; // a Sunday

function scenario(name, consultants, leaves, opts) {
  return Object.assign({ name, consultants, leaves: leaves || [] }, opts || {});
}
function cons(id, extra) {
  return Object.assign({ ConsultantID: id, Name: id, TeamID: 'T1', CanOffDuty: 'Yes', ConsultantType: 'Regular',
    CurrentStatus: 'Active', ActiveFrom: '2026-01-01', InactiveFrom: '', SubTeamAssignment: 'Any', SubTeamAssignmentFrom: '' }, extra || {});
}
function leave(id, s, e) { return { LeaveID: 'L' + id + s, ConsultantID: id, TeamID: 'T1', StartDate: s, EndDate: e, Status: 'Approved' }; }

const SCENARIOS = [
  scenario('4 consultants, no leave', ['A', 'B', 'C', 'D'].map(x => cons(x))),
  scenario('5 consultants, no leave', ['A', 'B', 'C', 'D', 'E'].map(x => cons(x))),
  scenario('4 consultants, regular leave', ['A', 'B', 'C', 'D'].map(x => cons(x)), [
    leave('A', '2026-08-02', '2026-08-15'), leave('B', '2026-10-04', '2026-10-17'), leave('C', '2026-12-06', '2026-12-26'),
    leave('D', '2027-02-07', '2027-02-13'), leave('A', '2027-03-14', '2027-03-27'), leave('B', '2027-05-02', '2027-05-08')]),
  scenario('5 consultants, one cannot take Off', ['A', 'B', 'C', 'D'].map(x => cons(x)).concat([cons('E', { CanOffDuty: 'No' })]), [
    leave('A', '2026-09-06', '2026-09-19'), leave('E', '2026-11-01', '2026-11-14')]),
  scenario('5 consultants, new joiner + one leaves', ['A', 'B', 'C', 'D'].map(x => cons(x)).concat([
    cons('E', { ActiveFrom: '2026-10-04' })]).map(c => (c.ConsultantID === 'B' ? Object.assign(c, { InactiveFrom: '2027-01-03' }) : c))),
  scenario('5 consultants, one fixed to TCU', ['A', 'B', 'C', 'D'].map(x => cons(x)).concat([
    cons('E', { SubTeamAssignment: 'SubTeam1' })]))
];

function run(sc) {
  const env = createEnv();
  seed(env, 'Teams', [{ TeamID: 'T1', TeamName: 'Medicine', Active: 'Yes', TeamType: 'SubTeam', SubTeamConfig: 'TCU,EMCU', SubTeamFrom: '' }]);
  seed(env, 'Settings', []);
  seed(env, 'EligibilityHistory', []);
  const result = env.buildRotaRows_({ versionId: 'SIM', startDate: env.toDate_(START), weeks: 52, fromWeekNo: 1,
    previousRows: [], consultantsOverride: sc.consultants, leavesOverride: sc.leaves, teamIds: ['T1'] });
  const rows = result.rows;
  const byWeek = {};
  rows.forEach(r => { (byWeek[r.WeekNo] || (byWeek[r.WeekNo] = [])).push(r); });
  const ip = {}, eligible = {}, last = {}, lastWeek = {};
  let serviceRepeats = 0, serviceChances = 0, consecutive = 0, uncovered = 0;
  for (let w = 1; w <= 52; w++) {
    const wk = byWeek[w] || [];
    ['TCU', 'EMCU'].forEach(role => { if (!wk.some(r => r.AssignmentPlace === role)) uncovered++; });
    wk.forEach(r => {
      const id = r.ConsultantID;
      if (['TCU', 'EMCU', 'OPD', 'Off-duty'].includes(r.AssignmentPlace)) eligible[id] = (eligible[id] || 0) + 1;
      if (r.AssignmentPlace === 'TCU' || r.AssignmentPlace === 'EMCU') {
        ip[id] = (ip[id] || 0) + 1;
        if (last[id]) {
          const fixed = sc.consultants.find(c => c.ConsultantID === id).SubTeamAssignment !== 'Any';
          if (!fixed) { serviceChances++; if (last[id] === r.AssignmentPlace) serviceRepeats++; }
        }
        if (lastWeek[id] === w - 1) consecutive++;
        last[id] = r.AssignmentPlace; lastWeek[id] = w;
      }
    });
  }
  // Inpatient share relative to each consultant's eligible (working) weeks.
  const shares = Object.keys(eligible).map(id => (ip[id] || 0) / eligible[id]);
  const spread = Math.max(...shares) - Math.min(...shares);
  return { ip, eligible, serviceRepeats, serviceChances, consecutive, uncovered, spread, conflicts: result.conflicts.length };
}

if (require.main === module) {
  SCENARIOS.forEach(sc => {
    const m = run(sc);
    console.log(`\n${sc.name}`);
    console.log('  inpatient weeks / eligible weeks: ' + Object.keys(m.eligible).sort().map(id => `${id} ${m.ip[id] || 0}/${m.eligible[id]}`).join(', '));
    console.log(`  inpatient-share spread: ${(m.spread * 100).toFixed(1)} pts | same-service repeats: ${m.serviceRepeats}/${m.serviceChances}` +
      ` | consecutive inpatient: ${m.consecutive} | uncovered posts: ${m.uncovered}`);
  });
}
module.exports = { SCENARIOS, run };
