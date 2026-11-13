/**
 * CTU ROTA - Google Apps Script backend
 * Mobile-first web app + Google Sheets database.
 *
 * Deploy as a web app from a Google Sheet-bound Apps Script project.
 */

const APP = {
  timezone: 'Asia/Riyadh',
  title: 'CTU ROTA',
  version: '3.12.0-rc1'
};

const SHEETS = {
  Settings: 'Settings',
  AdminUsers: 'AdminUsers',
  Teams: 'Teams',
  Consultants: 'Consultants',
  LeaveRequests: 'LeaveRequests',
  RotaPeriods: 'RotaPeriods',
  RotaVersions: 'RotaVersions',
  RotaAssignments: 'RotaAssignments',
  ConflictLog: 'ConflictLog',
  OverrideLog: 'OverrideLog',
  AuditLog: 'AuditLog',
  EligibilityHistory: 'EligibilityHistory'
};

const HEADERS = {
  Settings: ['Key', 'Value'],
  AdminUsers: ['UserID', 'Username', 'Name', 'PIN', 'PINHash', 'PINSalt', 'AuthVersion', 'Role', 'TeamID', 'Active', 'CreatedAt', 'UpdatedAt', 'LastLoginAt', 'Notes'],
  Teams: ['TeamID', 'TeamName', 'ColorHex', 'DisplayOrder', 'Active', 'CreatedAt', 'UpdatedAt', 'TeamType', 'SubTeamConfig', 'SubTeamFrom', 'OneConsultantPriority', 'SubTeamUntil', 'SubTeamResumeFrom'],
  Consultants: ['ConsultantID', 'Name', 'TeamID', 'CanOffDuty', 'ConsultantType', 'ColorHex', 'CurrentStatus', 'ActiveFrom', 'InactiveFrom', 'Notes', 'CreatedAt', 'UpdatedAt', 'SubTeamAssignment', 'SubTeamAssignmentFrom'],
  LeaveRequests: ['LeaveID', 'ConsultantID', 'TeamID', 'StartDate', 'EndDate', 'Status', 'SubmittedAt', 'UserNote', 'AdminDecisionDate', 'AdminNote', 'ImpactStatus', 'ImpactSummary', 'SupersedesLeaveID', 'SupersededByLeaveID'],
  RotaPeriods: ['PeriodID', 'PeriodName', 'Status', 'StartDate', 'EndDate', 'Weeks', 'CreatedAt', 'CreatedBy', 'ActivatedAt', 'ArchivedAt', 'ArchiveSpreadsheetID', 'ArchiveSpreadsheetURL', 'Notes'],
  RotaVersions: ['VersionID', 'AmendmentNo', 'Status', 'Type', 'CreatedAt', 'CreatedBy', 'StartDate', 'EndDate', 'Weeks', 'EffectiveFrom', 'Reason', 'BasedOnVersionID', 'ExportDocId', 'ExportDocUrl', 'PeriodID', 'AssignmentStorage', 'ArchiveSheetName', 'AssignmentCount'],
  RotaAssignments: ['VersionID', 'WeekNo', 'WeekStart', 'WeekEnd', 'TeamID', 'AssignmentPlace', 'ConsultantID', 'Source', 'Notes', 'PeriodID'],
  ConflictLog: ['ConflictID', 'VersionID', 'WeekNo', 'WeekStart', 'TeamID', 'Level', 'Rule', 'Message', 'CreatedAt'],
  OverrideLog: ['OverrideID', 'VersionID', 'RequestType', 'RequestID', 'RuleBroken', 'Reason', 'CreatedAt'],
  AuditLog: ['AuditID', 'Timestamp', 'ActorName', 'ActorRole', 'ActorTeamID', 'Action', 'TargetType', 'TargetID', 'Details'],
  EligibilityHistory: ['HistoryID', 'ConsultantID', 'TeamID', 'EffectiveFrom', 'CanOffDuty', 'ConsultantType', 'SubTeamAssignment', 'RecordedAt']
};

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'manifest') {
    return ContentService.createTextOutput(JSON.stringify(WEB_APP_MANIFEST_)).setMimeType(ContentService.MimeType.JSON);
  }
  // Apps Script HTML files are referenced by their internal name.
  // Correct name is usually 'Index'. If the file was accidentally named
  // 'Index.html', this fallback prevents the common deployment error.
  try {
    return renderHtmlFile_('Index');
  } catch (err) {
    try {
      return renderHtmlFile_('Index.html');
    } catch (err2) {
      throw new Error('No HTML file named Index was found. In Apps Script, create an HTML file named exactly Index, then paste Index.html content into it.');
    }
  }
}

// JSON API for the externally-hosted frontend (GitHub Pages), used instead of google.script.run
// so the browser never has to depend on Apps Script's cross-origin sandboxed iframe bridge.
// The client sends { fn, args } as a text/plain body (not application/json) specifically to
// avoid a CORS preflight request, which Apps Script web apps cannot answer (no doOptions support).
function doPost(e) {
  return handleApiRequest_(e);
}

// Per-request counters only; never persist or disclose PINs, tokens, names, or arguments.
let API_PERF_CONTEXT_ = null;
// Sheet objects and validated headers are memoized for ONE API invocation only.
// Mutations are never cached across API calls or separate Apps Script executions.
let REQUEST_SHEET_MEMO_ = null;
let ADMIN_READ_MEMO_ = null;
function withAdminReadMemo_(fn) {
  const previous = ADMIN_READ_MEMO_;
  ADMIN_READ_MEMO_ = Object.create(null);
  try { return fn(); }
  finally { ADMIN_READ_MEMO_ = previous; }
}
function perfSheetRead_(rowCount) {
  if (!API_PERF_CONTEXT_) return;
  API_PERF_CONTEXT_.sheetReads++;
  API_PERF_CONTEXT_.rowsRead += Math.max(0, Number(rowCount || 0));
}
function perfSheetWrite_(rowCount) {
  if (!API_PERF_CONTEXT_) return;
  API_PERF_CONTEXT_.sheetWrites++;
  API_PERF_CONTEXT_.rowsWritten += Math.max(0, Number(rowCount || 0));
}
function handleApiRequest_(e) {
  const startedAt = Date.now();
  let route = 'invalid';
  let success = false;
  const context = { sheetReads: 0, rowsRead: 0, sheetWrites: 0, rowsWritten: 0, schemaChecks: 0, rowFinds: 0, cacheInvalidations: 0 };
  // Do not reuse a sheet reference between user requests or database copies.
  REQUEST_SHEET_MEMO_ = Object.create(null);
  API_PERF_CONTEXT_ = context;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = String(body.fn || '');
    const args = Array.isArray(body.args) ? body.args : [];
    if (!Object.prototype.hasOwnProperty.call(API_WHITELIST_, fn)) {
      throw new Error('Unknown API function: ' + fn);
    }
    route = fn; // Server-defined allowlisted name only, never user input.
    const result = API_WHITELIST_[fn].apply(null, args);
    success = true;
    return jsonOutput_({ ok: true, result: result });
  } catch (err) {
    return jsonOutput_({ ok: false, error: (err && err.message) || String(err) });
  } finally {
    // Accessible via Apps Script > Executions/Logs. No additional Sheets writes.
    try { Logger.log(JSON.stringify({ event: 'rota_api_perf', version: APP.version,
      operation: route, ok: success, durationMs: Date.now() - startedAt,
      reads: context.sheetReads, readRows: context.rowsRead,
      writes: context.sheetWrites, writtenRows: context.rowsWritten,
      schemaChecks: context.schemaChecks, rowFinds: context.rowFinds,
      cacheInvalidations: context.cacheInvalidations })); } catch (_) {}
    API_PERF_CONTEXT_ = null;
    REQUEST_SHEET_MEMO_ = null;
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Only functions the client actually calls are reachable through the API - this is a deliberate
// whitelist, not a generic "call any function by name" bridge. Admin-only functions still enforce
// their own requireAdminPin_ check inside the function itself, exactly as before.
const API_WHITELIST_ = {
  getBootstrap: getBootstrap,
  getStartupPayload: getStartupPayload,
  loginAdmin: loginAdmin,
  logoutAdminSession: logoutAdminSession,
  getAdminSession: getAdminSession,
  getAdminData: getAdminData,
  getAdminLeaveData: getAdminLeaveData,
  getAssignmentStorageSummary: getAssignmentStorageSummary,
  getAuditLog: getAuditLog,
  getRotaWeekList: getRotaWeekList,
  getRotaWeekTeamView: getRotaWeekTeamView,
  getLatestRotaView: getLatestRotaView,
  getLatestRotaWindow: getLatestRotaWindow,
  setupDatabaseAdmin: setupDatabaseAdmin,
  archiveHistoricalAssignments: archiveHistoricalAssignments,
  updateAdminPin: updateAdminPin,
  saveAdminUser: saveAdminUser,
  saveTeam: saveTeam,
  saveSubTeamCoverageChoice: saveSubTeamCoverageChoice,
  saveConsultant: saveConsultant,
  startNewRota: startNewRota,
  createDraftRota: createDraftRota,
  rearrangeTeamRota: rearrangeTeamRota,
  publishDraftRota: publishDraftRota,
  submitLeaveRequest: submitLeaveRequest,
  checkLeaveAvailability: checkLeaveAvailability,
  getLeaveRequestsForConsultant: getLeaveRequestsForConsultant,
  updateLeaveRequest: updateLeaveRequest,
  previewLeaveImpact: previewLeaveImpact,
  previewLeaveWhatIf: previewLeaveWhatIf,
  approveLeave: approveLeave,
  amendApprovedLeave: amendApprovedLeave,
  cancelApprovedLeave: cancelApprovedLeave,
  rejectLeave: rejectLeave,
  previewInactivity: previewInactivity,
  confirmInactivity: confirmInactivity,
  voidRotaVersion: voidRotaVersion,
  restoreRotaVersion: restoreRotaVersion,
  deleteDraftRota: deleteDraftRota,
  resetRotaData: resetRotaData,
  publishManualEditedWeek: publishManualEditedWeek,
  exportLatestRota: exportLatestRota
};

// Best-effort "Add to Home Screen" support for Android/Chrome. iOS Safari does not read this
// file directly (it relies on the apple-mobile-web-app meta tags + apple-touch-icon instead).
const WEB_APP_MANIFEST_ = {
  name: APP.title,
  short_name: 'CTU ROTA',
  start_url: '.',
  scope: '.',
  display: 'standalone',
  background_color: '#01163a',
  theme_color: '#01163a',
  icons: []
};

function renderHtmlFile_(name) {
  // HtmlService.addMetaTag() only supports the 'viewport' name - any other name throws.
  // Other meta tags (theme-color, apple-mobile-web-app-capable, etc.) are set directly
  // as plain <meta> elements in Index.html's <head> instead.
  return HtmlService.createTemplateFromFile(name)
    .evaluate()
    .setTitle(APP.title)
    // No maximum-scale here on purpose - locking pinch-zoom is an accessibility issue on both iOS and Android.
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupDatabase() {
  const cache = CacheService.getScriptCache();
  // Versioned cache key: schema/auth changes must be checked after every release upgrade.
  const cacheKey = 'dbReady_v3_10_reactivate';
  if (cache.get(cacheKey) === '1') {
    return { ok: true, message: 'Database sheets are ready.', sheets: Object.values(SHEETS), url: getDb_().getUrl() };
  }
  const ss = getDb_();
  Object.keys(SHEETS).forEach(key => ensureSheet_(SHEETS[key], HEADERS[key]));
  seedSettings_();
  ensureAdminAuthMigration_();
  ensureRotaPeriodMigration_();
  cache.put(cacheKey, '1', 3600);
  return { ok: true, message: 'Database sheets are ready.', sheets: Object.values(SHEETS), url: ss.getUrl() };
}

const PUBLIC_STARTUP_CACHE_BASE_ = 'publicStartupPayload_v3_12';
const PUBLIC_STARTUP_CACHE_SECONDS_ = 21600;
const PUBLIC_STARTUP_CACHE_CHUNK_SIZE_ = 78000;

// Backward-compatible endpoint for older cached frontends. From v1.9 it returns the
// complete public startup payload, including the current eight-week rota window, so an
// older v1.8 page also avoids making a second request after startup.
function getBootstrap() {
  return getStartupPayload();
}

function getStartupPayload() {
  const cached = getCachedJsonChunks_(PUBLIC_STARTUP_CACHE_BASE_);
  if (cached) return cached;
  return buildAndCachePublicStartupPayload_();
}

function buildAndCachePublicStartupPayload_() {
  const payload = buildPublicStartupPayload_();
  putCachedJsonChunks_(PUBLIC_STARTUP_CACHE_BASE_, payload, PUBLIC_STARTUP_CACHE_SECONDS_);
  return payload;
}

function buildPublicStartupPayload_() {
  // Public startup is deliberately read-only. It does not call setupDatabase(), add
  // columns, seed settings, run migrations, or repair sheets. Schema work remains under
  // Setup -> Safe Repair / Add Missing Columns.
  const periods = readObjectsNoRepair_(SHEETS.RotaPeriods);
  const versions = readObjectsNoRepair_(SHEETS.RotaVersions);
  const teams = sortTeams_(readObjectsNoRepair_(SHEETS.Teams));
  const consultants = readObjectsNoRepair_(SHEETS.Consultants);
  const assignments = readObjectsNoRepair_(SHEETS.RotaAssignments);
  const latest = getLatestVersionFromData_(periods, versions);
  const latestRota = latest
    ? buildRotaWindowFromData_(latest, assignments, teams, consultants, 8)
    : { meta: null, weeks: [], viewMode: 'window', window: null, summary: 'No official rota generated yet.' };
  return {
    app: { title: APP.title, version: APP.version, timezone: APP.timezone, subTeamFeatureVersion: 9 },
    serverTime: now_(),
    teams: teams,
    consultants: consultants,
    leaveRequests: [],
    latestVersion: latest,
    latestRota: latestRota
  };
}

function readObjectsNoRepair_(sheetName) {
  const sh = getDb_().getSheetByName(sheetName);
  if (!sh) throw new Error(`Database sheet ${sheetName} is missing. Main admin must run Setup -> Safe Repair / Add Missing Columns.`);
  const values = sh.getDataRange().getValues();
  perfSheetRead_(values.length > 0 ? values.length - 1 : 0);
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1)
    .filter(row => row.some(v => v !== '' && v !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = normalizeCell_(row[i]));
      return obj;
    });
}

function getLatestVersionFromData_(periods, versions) {
  const active = (periods || [])
    .filter(p => p.Status === 'Active')
    .sort((a, b) => String(b.ActivatedAt || b.CreatedAt || '').localeCompare(String(a.ActivatedAt || a.CreatedAt || '')))[0] || null;
  let official = (versions || []).filter(v => v.Status === 'Official');
  if (active) official = official.filter(v => String(v.PeriodID || '') === String(active.PeriodID));
  if (!official.length) return null;
  official.sort((a, b) => {
    const amendmentDiff = Number(b.AmendmentNo || 0) - Number(a.AmendmentNo || 0);
    return amendmentDiff || String(b.CreatedAt || '').localeCompare(String(a.CreatedAt || ''));
  });
  const version = official[0];
  const period = (periods || []).find(p => String(p.PeriodID) === String(version.PeriodID || '')) || null;
  return Object.assign({}, version, {
    PeriodName: period ? period.PeriodName : '',
    PeriodStatus: period ? period.Status : '',
    PeriodStartDate: period ? period.StartDate : version.StartDate,
    PeriodEndDate: period ? period.EndDate : version.EndDate
  });
}

function buildRotaWindowFromData_(latest, assignments, teams, consultants, windowWeeks) {
  const totalWeeks = Number(latest.Weeks || 52);
  const start = toDate_(latest.StartDate);
  const todayIso = Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd');
  const currentWeek = getWeekNoForDate_(start, toDate_(todayIso), totalWeeks);
  let startWeek;
  let summary;
  if (currentWeek < 1) {
    startWeek = 1;
    summary = 'The rota has not started yet. Showing the first upcoming weeks.';
  } else if (currentWeek > totalWeeks) {
    startWeek = Math.max(1, totalWeeks - 3);
    summary = 'This rota period has ended. Showing the final weeks.';
  } else {
    startWeek = currentWeek;
    summary = 'Showing the current rota week and upcoming weeks. Use View Full 52-Week Rota for the complete rota.';
  }
  const endWeek = Math.min(totalWeeks, startWeek + Math.max(1, Number(windowWeeks || 8)) - 1);
  const rows = (assignments || []).filter(r =>
    String(r.VersionID) === String(latest.VersionID) &&
    Number(r.WeekNo) >= startWeek && Number(r.WeekNo) <= endWeek
  );
  const view = buildRotaViewFromData_(latest, rows, teams, consultants);
  view.viewMode = 'window';
  view.window = { startWeek: startWeek, endWeek: endWeek, totalWeeks: totalWeeks, currentWeek: currentWeek };
  view.summary = summary;
  return view;
}

// Public rota display only. This does not remove, modify, or rewrite assignment records.
// A consultant with a future InactiveFrom remains visible until their last eligible
// week. A date within a week makes that whole weekly assignment ineligible,
// consistent with isActiveForWeek_ in the generator.
function hidePublicInactiveAssignment_(assignment, consultantLookup) {
  if (String(assignment.AssignmentPlace || '') === 'Inactive') return true;
  const consultant = consultantLookup[String(assignment.ConsultantID || '')];
  if (!consultant) return false; // Preserve independent external/manual cover.
  const inactiveFrom = String(consultant.InactiveFrom || '').slice(0, 10);
  const weekEnd = String(assignment.WeekEnd || assignment.WeekStart || '').slice(0, 10);
  if (inactiveFrom) return !!weekEnd && inactiveFrom <= weekEnd;
  // Legacy records may have CurrentStatus=Inactive but no date.
  return String(consultant.CurrentStatus || '').trim().toLowerCase() === 'inactive';
}

function buildRotaViewFromData_(latest, rows, teams, consultants) {
  const teamById = {};
  const teamOrder = {};
  (teams || []).forEach((t, idx) => { teamById[String(t.TeamID)] = t; teamOrder[String(t.TeamID)] = idx + 1; });
  const consultantById = {};
  (consultants || []).forEach(c => consultantById[String(c.ConsultantID)] = c);
  const weekMap = {};
  (rows || []).forEach(r => {
    const weekKey = String(r.WeekNo);
    if (!weekMap[weekKey]) {
      weekMap[weekKey] = { weekNo: Number(r.WeekNo), period: `${nice_(r.WeekStart)} – ${nice_(r.WeekEnd)}`, periodExport: `${dateOnlyNice_(r.WeekStart)} – ${dateOnlyNice_(r.WeekEnd)}`, teams: {} };
    }
    const team = teamById[String(r.TeamID)] || { TeamName: r.TeamID, ColorHex: '#0f172a' };
    if (!weekMap[weekKey].teams[r.TeamID]) {
      weekMap[weekKey].teams[r.TeamID] = { teamId: r.TeamID, teamName: team.TeamName, teamColor: normalizeHexColor_(team.ColorHex || '#0f172a'), subTeamNames: subTeamNamesForView_(team, r.WeekStart, [r]),
        scheduledStandardFrom: (!isSubTeamActiveForWeek_(team, r.WeekStart) ? String(team.SubTeamUntil || '').slice(0,10) : ''), scheduledSubTeamsFrom: String(team.SubTeamResumeFrom || '').slice(0,10), weekStart: String(r.WeekStart || '').slice(0,10), Team: [], OPD: [], 'Off-duty': [], Reserve: [], Leave: [], Inactive: [] };
    }
    if (hidePublicInactiveAssignment_(r, consultantById)) {
      if (String(r.AssignmentPlace) !== 'Inactive') weekMap[weekKey].teams[r.TeamID].hasHiddenInactiveAssignments = true;
      return;
    }
    const person = personFromAssignment_(r, consultantById);
    if (!weekMap[weekKey].teams[r.TeamID].subTeamNames.length) weekMap[weekKey].teams[r.TeamID].subTeamNames = subTeamNamesForView_(team, r.WeekStart, [r]);
    weekMap[weekKey].teams[r.TeamID].subTeamNames.forEach(role => { if (!weekMap[weekKey].teams[r.TeamID][role]) weekMap[weekKey].teams[r.TeamID][role] = []; });
    if (!weekMap[weekKey].teams[r.TeamID][r.AssignmentPlace]) weekMap[weekKey].teams[r.TeamID][r.AssignmentPlace] = [];
    weekMap[weekKey].teams[r.TeamID][r.AssignmentPlace].push(person);
  });
  const weeks = Object.values(weekMap)
    .sort((a, b) => a.weekNo - b.weekNo)
    .map(w => Object.assign({}, w, { teams: Object.values(w.teams).sort((a, b) => {
      const orderDiff = (teamOrder[String(a.teamId)] || 999999) - (teamOrder[String(b.teamId)] || 999999);
      return orderDiff || String(a.teamName || '').localeCompare(String(b.teamName || ''));
    }) }));
  return { meta: latest, weeks: weeks };
}

function getCachedJsonChunks_(baseKey) {
  const cache = CacheService.getScriptCache();
  const metaRaw = cache.get(baseKey + '_meta');
  if (!metaRaw) return null;
  try {
    const meta = JSON.parse(metaRaw);
    const todayKey = Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd');
    if (String(meta.dateKey || '') !== todayKey) {
      invalidatePublicStartupCache_();
      return null;
    }
    const count = Math.max(0, Number(meta.count || 0));
    if (!count) return null;
    const keys = [];
    for (let i = 0; i < count; i++) keys.push(baseKey + '_part_' + i);
    const values = cache.getAll(keys);
    const json = keys.map(k => values[k] || '').join('');
    if (!json) return null;
    return JSON.parse(json);
  } catch (err) {
    invalidatePublicStartupCache_();
    return null;
  }
}

function putCachedJsonChunks_(baseKey, value, seconds) {
  const cache = CacheService.getScriptCache();
  const json = JSON.stringify(value);
  const chunks = [];
  for (let i = 0; i < json.length; i += PUBLIC_STARTUP_CACHE_CHUNK_SIZE_) chunks.push(json.slice(i, i + PUBLIC_STARTUP_CACHE_CHUNK_SIZE_));
  invalidatePublicStartupCache_();
  const entries = {};
  chunks.forEach((chunk, idx) => entries[baseKey + '_part_' + idx] = chunk);
  if (chunks.length) cache.putAll(entries, seconds || PUBLIC_STARTUP_CACHE_SECONDS_);
  cache.put(baseKey + '_meta', JSON.stringify({ count: chunks.length, createdAt: now_(), dateKey: Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd') }), seconds || PUBLIC_STARTUP_CACHE_SECONDS_);
}

function invalidatePublicStartupCache_() {
  if (API_PERF_CONTEXT_) API_PERF_CONTEXT_.cacheInvalidations++;
  // The metadata key is the commit marker. Removing it makes ALL existing parts
  // unreachable to readers; stale chunks expire by their normal TTL. No need to
  // fetch the old metadata and delete potentially many part keys on each write.
  CacheService.getScriptCache().remove(PUBLIC_STARTUP_CACHE_BASE_ + '_meta');
}

function warmPublicStartupCache_() {
  // Administrative writes already invalidate the public snapshot. Rebuilding it here
  // makes leave approval and manual edits block on five extra Sheet reads plus
  // serialization. Let the next public read refresh the snapshot once, on demand.
  invalidatePublicStartupCache_();
  return null;
}

function sheetAffectsPublicStartup_(sheetName) {
  return [SHEETS.Teams, SHEETS.Consultants, SHEETS.RotaPeriods, SHEETS.RotaVersions, SHEETS.RotaAssignments].indexOf(sheetName) !== -1;
}

function getAdminData(adminPin, mode) {
  return withAdminReadMemo_(function() {
  const session = requireAnyAdmin_(adminPin);
  const latest = getLatestVersion_();
  const allTeams = getTeams_();
  const allConsultants = getConsultants_();
  const includeLeave = String(mode || '') !== 'dashboard';
  const allLeave = includeLeave ? readObjects_(SHEETS.LeaveRequests) : [];
  const allowedTeamIds = session.isMainAdmin ? null : [String(session.teamId || '')];
  const periods = session.isMainAdmin ? getRotaPeriods_() : [];
  const versions = session.isMainAdmin ? getRotaVersions_(periods) : [];
  const latestConflicts = getLatestConflicts_(latest, allTeams);
  return {
    adminSession: session,
    settings: getPublicSettings_(),
    teams: allowedTeamIds ? allTeams.filter(t => allowedTeamIds.indexOf(String(t.TeamID)) !== -1) : allTeams,
    allTeams: allTeams,
    consultants: allowedTeamIds ? allConsultants.filter(c => allowedTeamIds.indexOf(String(c.TeamID)) !== -1) : allConsultants,
    leaveRequests: allowedTeamIds ? allLeave.filter(l => allowedTeamIds.indexOf(String(l.TeamID)) !== -1) : allLeave,
    leaveRequestsIncluded: includeLeave,
    adminUsers: session.isMainAdmin ? getPublicAdminUsers_() : [],
    latestVersion: latest,
    rotaPeriods: periods,
    rotaVersions: versions,
    latestRota: latest ? { meta: latest, weeks: [], notLoaded: true } : { meta: null, weeks: [] },
    assignmentStorage: null,
    latestConflicts: session.isMainAdmin ? latestConflicts : latestConflicts.filter(c => allowedTeamIds.indexOf(String(c.TeamID)) !== -1)
  };
  });
}

function getAdminLeaveData(adminPin) {
  return withAdminReadMemo_(function() {
  const session = requireAnyAdmin_(adminPin);
  const allTeams = getTeams_();
  const allConsultants = getConsultants_();
  const allLeave = readObjects_(SHEETS.LeaveRequests);
  const allowedTeamIds = session.isMainAdmin ? null : [String(session.teamId || '')];
  return {
    adminSession: session,
    teams: allowedTeamIds ? allTeams.filter(t => allowedTeamIds.indexOf(String(t.TeamID)) !== -1) : allTeams,
    allTeams: allTeams,
    consultants: allowedTeamIds ? allConsultants.filter(c => allowedTeamIds.indexOf(String(c.TeamID)) !== -1) : allConsultants,
    leaveRequests: allowedTeamIds ? allLeave.filter(l => allowedTeamIds.indexOf(String(l.TeamID)) !== -1) : allLeave,
    leaveRequestsIncluded: true
  };
  });
}

function getAssignmentStorageSummary(adminPin) {
  requireAdminPin_(adminPin);
  return getAssignmentStorageSummaryCached_();
}

function getRotaWeekList(adminPin) {
  requireAdminPin_(adminPin);
  const latest = getLatestVersion_();
  if (!latest) return { meta: null, weeks: [] };
  const start = toDate_(latest.StartDate);
  const weeks = Number(latest.Weeks || 52);
  const list = [];
  for (let i = 1; i <= weeks; i++) {
    const ws = addDays_(start, (i - 1) * 7);
    const we = addDays_(ws, 6);
    list.push({ weekNo: i, period: `${nice_(ws)} – ${nice_(we)}`, periodExport: `${dateOnlyNice_(ws)} – ${dateOnlyNice_(we)}` });
  }
  return { meta: latest, weeks: list };
}

function getRotaWeekView(adminPin, weekNo) {
  requireAdminPin_(adminPin);
  return getRotaWeekView_(Number(weekNo || 0));
}

function getRotaWeekTeamView(adminPin, weekNo, teamId) {
  requireAdminPin_(adminPin);
  return getRotaWeekTeamView_(Number(weekNo || 0), String(teamId || '').trim());
}


function publishManualEditedWeek(adminPin, data) {
  const session = requireAdminPin_(adminPin);
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let version = null;
  let latest = null;
  let madeOfficial = false;
  try {
    latest = getLatestVersion_();
    if (!latest) throw new Error('No current official rota exists to edit.');
    const weekNo = Number(data.weekNo || 0);
    const weeks = Number(latest.Weeks || 52);
    if (!weekNo || weekNo < 1 || weekNo > weeks) throw new Error('Select a valid week to edit.');
    const edits = Array.isArray(data.edits) ? data.edits : [];
    if (!edits.length) throw new Error('No manual edits were submitted.');
    const start = toDate_(latest.StartDate), weekStart = addDays_(start, (weekNo - 1) * 7), weekEnd = addDays_(weekStart, 6);
    const weekStartIso = iso_(weekStart), weekEndIso = iso_(weekEnd);
    // The current week may require emergency correction, but closed weeks must
    // never be retrospectively overwritten through routine Manual Edit.
    if (weekStart < startOfWeekSunday_(new Date()))
      throw new Error('Completed rota weeks are read-only. Use the audit/restore process for historical records.');
    const reason = String(data.reason || '').trim() || `Manual week ${weekNo} rota edit with subsequent rota recalculation`;
    version = createRotaVersion_({ PeriodID: latest.PeriodID, AmendmentNo: nextAmendmentNoForPeriod_(latest.PeriodID), Status: 'Draft', Type: 'Manual rota edit + recalculation', StartDate: latest.StartDate, EndDate: latest.EndDate, Weeks: weeks, EffectiveFrom: weekStartIso, Reason: reason, BasedOnVersionID: latest.VersionID });
    const editMap = {};
    edits.forEach(e => { const teamId = String(e.teamId || '').trim(); if (teamId) editMap[teamId] = e; });
    const originalRows = getVersionAssignments_(latest.VersionID), replacePlaces = { 'Team': true, 'OPD': true, 'Off-duty': true, 'TCU': true, 'EMCU': true, 'Reserve': true }, fixedRows = [];
    getTeams_().filter(t => isSubTeamConfigured_(t)).forEach(t => subTeamNames_(t).forEach(role => replacePlaces[role] = true));
    originalRows.forEach(r => {
      const rowWeek = Number(r.WeekNo);
      // Copy ALL weeks for teams not being edited; otherwise this manual edit
      // would silently regenerate their published schedules.
      const isEditedTeam = Object.prototype.hasOwnProperty.call(editMap, String(r.TeamID));
      const shouldKeep = !isEditedTeam || rowWeek < weekNo ||
        (rowWeek === weekNo && !replacePlaces[String(r.AssignmentPlace)]);
      if (shouldKeep) fixedRows.push(Object.assign({}, r, { VersionID: version.VersionID, PeriodID: latest.PeriodID }));
    });
    Object.keys(editMap).forEach(teamId => {
      const e = editMap[teamId];
      const teamConfig = teamMap_()[teamId];
      // Never silently treat sub-team text as a legacy Team edit. Older UI
      // builds may submit fields before the new structure becomes effective.
      if (isSubTeamConfigured_(teamConfig) && !isSubTeamActiveForWeek_(teamConfig,weekStartIso) &&
          (e.subTeamFieldsSubmitted || parseManualNames_(e.tcuText).length || parseManualNames_(e.emcuText).length || parseManualNames_(e.reserveText).length)) {
        throw new Error('Sub-teams are configured but not active in Week '+weekNo+'. Activate the structure for this Sunday in Manual Edit, then reload the week before publishing. No rota was changed.');
      }
      if (isSubTeamActiveForWeek_(teamConfig,weekStartIso)) {
        const roles = subTeamNames_(teamMap_()[teamId]);
        appendManualNameRows_(fixedRows, version.VersionID, weekNo, weekStartIso, weekEndIso, teamId, roles[0], e.tcuText);
        appendManualNameRows_(fixedRows, version.VersionID, weekNo, weekStartIso, weekEndIso, teamId, roles[1], e.emcuText);
      } else appendManualNameRows_(fixedRows, version.VersionID, weekNo, weekStartIso, weekEndIso, teamId, 'Team', e.teamText);
      // Old frontend submissions can still contain reserveText; migrate the
      // submitted consultant(s) into OPD rather than recreating Reserve rows.
      const opdNames = parseManualNames_(e.opdText).concat(parseManualNames_(e.reserveText));
      const uniqueOpdNames = Array.from(new Set(opdNames.map(x=>normalizeName_(x)))).map(key=>opdNames.find(x=>normalizeName_(x)===key));
      appendManualNameRows_(fixedRows, version.VersionID, weekNo, weekStartIso, weekEndIso, teamId, 'OPD', uniqueOpdNames.join('\n'));
      appendManualNameRows_(fixedRows, version.VersionID, weekNo, weekStartIso, weekEndIso, teamId, 'Off-duty', e.offText);
    });
    let finalRows = fixedRows;
    let conflicts = validateManualEditedWeek_(fixedRows, weekNo, editMap);
    if (weekNo < weeks) {
      const generated = buildRotaRows_({ versionId: version.VersionID, startDate: start, weeks: weeks, fromWeekNo: weekNo + 1, previousRows: fixedRows, manualSeedWeeks: 0, manualAssignments: [], leavesOverride: getApprovedLeaves_(), teamIds: Object.keys(editMap) });
      finalRows = generated.rows;
      conflicts = conflicts.concat(generated.conflicts);
    }
    if (!finalRows.length) throw new Error('Safety stop: the manual edit generated no assignments. The live rota was not changed.');
    assertRowsBeforeEffectiveWeekUnchanged_(originalRows, finalRows, weekNo);
    // No manual edit of one parent team may mutate the remaining teams.
    const untouched = originalRows.filter(r => !Object.prototype.hasOwnProperty.call(editMap, String(r.TeamID)));
    const untouchedAfter = finalRows.filter(r => !Object.prototype.hasOwnProperty.call(editMap, String(r.TeamID)));
    if (assignmentSignatureAll_(untouched) !== assignmentSignatureAll_(untouchedAfter)) {
      throw new Error('Safety stop: manual edit unexpectedly changed another team. No amendment was published.');
    }
    appendRows_(SHEETS.RotaAssignments, finalRows);
    appendConflicts_(version.VersionID, conflicts);
    appendOverride_({ VersionID: version.VersionID, requestType: 'Manual edit', requestId: `Week ${weekNo}`, ruleBroken: 'Admin manual rota amendment', reason: reason });
    updateRow_(SHEETS.RotaVersions, 'VersionID', version.VersionID, { Status: 'Official' });
    madeOfficial = true;
    archiveVersion_(latest.VersionID);
    version.Status = 'Official';
    appendAudit_(session, 'MANUAL_ROTA_PUBLISHED', 'RotaVersion', version.VersionID, { amendmentNo: version.AmendmentNo, weekNo: weekNo, reason: reason, conflictCount: conflicts.length });
    warmPublicStartupCache_();
    return { ok: true, message: `Manual rota edit published as Amendment ${version.AmendmentNo}. Week ${weekNo} was kept as edited, and subsequent weeks were recalculated.`, version: version, conflicts: conflicts };
  } catch (err) {
    if (version && version.VersionID) {
      if (madeOfficial && latest) {
        restoreOfficialAfterRollback_(latest);
        try { updateRow_(SHEETS.RotaVersions, 'VersionID', version.VersionID, { Status: 'Draft' }); } catch (_) {}
      }
      try { deleteRowsWhere_(SHEETS.RotaAssignments, r => r.VersionID === version.VersionID); } catch (_) {}
      try { deleteRowsWhere_(SHEETS.ConflictLog, r => r.VersionID === version.VersionID); } catch (_) {}
      try { deleteRowsWhere_(SHEETS.OverrideLog, r => r.VersionID === version.VersionID); } catch (_) {}
      try { deleteRowsWhere_(SHEETS.RotaVersions, r => r.VersionID === version.VersionID); } catch (_) {}
    }
    throw err;
  } finally { lock.releaseLock(); }
}

function setupDatabaseAdmin(adminToken) {
  // Safe Repair always performs a fresh header check even if this request has
  // already read a sheet; never trust request-local schema memo for repairs.
  if (REQUEST_SHEET_MEMO_) REQUEST_SHEET_MEMO_ = Object.create(null);
  const session = requireAdminPin_(adminToken);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const cache = CacheService.getScriptCache();
    ['dbReady', 'dbReady_v1_6', 'dbReady_v1_7', 'dbReady_v1_8', 'dbReady_v1_9', 'dbReady_v2_3_subteams', 'dbReady_v3_10_reactivate'].forEach(k => cache.remove(k));
    const result = setupDatabase();
    const archiveSummary = assignmentStorageSummary_();
    warmPublicStartupCache_();
    appendAudit_(session, 'SAFE_REPAIR_RUN', 'Database', 'Schema', { appVersion: APP.version, assignmentArchive: archiveSummary });
    return Object.assign({}, result, {
      message: `${result.message} Authentication migration and assignment storage check completed.`,
      assignmentArchive: archiveSummary
    });
  } finally {
    lock.releaseLock();
  }
}

function updateAdminPin(adminToken, newPin) {
  const session = requireAdminPin_(adminToken);
  setupDatabase();
  const pin = String(newPin || '').trim();
  if (!/^\d{4,8}$/.test(pin)) throw new Error('Admin PIN must be 4 to 8 digits.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const users = getAdminUsersRaw_();
    const user = users.find(u => String(u.UserID) === String(session.userId));
    if (!user) throw new Error('Main admin account was not found. Run Safe Repair.');
    const salt = makePinSalt_();
    const nextAuthVersion = Number(user.AuthVersion || 1) + 1;
    updateRow_(SHEETS.AdminUsers, 'UserID', user.UserID, {
      PIN: '', PINHash: hashPin_(pin, salt), PINSalt: salt,
      AuthVersion: nextAuthVersion, UpdatedAt: now_()
    });
    appendAudit_(session, 'ADMIN_PIN_CHANGED', 'AdminUser', user.UserID, 'Main admin PIN changed. PIN value is not stored in the audit log. Existing sessions were invalidated.');
    return { ok: true, message: 'Admin PIN updated. Please sign in again with the new PIN.', sessionInvalidated: true };
  } finally {
    lock.releaseLock();
  }
}

function saveAdminUser(adminToken, data) {
  const session = requireAdminPin_(adminToken);
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const input = data || {};
    const name = String(input.Name || '').trim();
    const username = normalizeUsername_(input.Username);
    if (!name) throw new Error('Sub-admin name is required.');
    if (!isValidUsername_(username)) throw new Error('Username must be 3 to 40 characters using letters, numbers, dot, underscore, or hyphen.');
    const role = String(input.Role || 'TeamAdmin').trim() || 'TeamAdmin';
    if (role !== 'TeamAdmin') throw new Error('Only TeamAdmin sub-admin users can be created from this screen.');
    const teamId = String(input.TeamID || '').trim();
    if (!teamId) throw new Error('Assign the sub-admin to a team.');
    const team = getTeams_().find(t => String(t.TeamID) === teamId);
    if (!team) throw new Error('Selected team was not found.');
    const active = normalizeYesNo_(input.Active, 'Yes');
    const userId = String(input.UserID || '').trim() || makeId_('ADM');
    const users = getAdminUsersRaw_();
    const existing = users.find(u => String(u.UserID) === userId) || null;
    if (users.some(u => String(u.UserID) !== userId && normalizeUsername_(u.Username) === username)) {
      throw new Error('This username is already used. Choose a unique username.');
    }
    const pin = String(input.PIN || '').trim();
    if (!existing && !/^\d{4,8}$/.test(pin)) throw new Error('A 4 to 8 digit PIN is required for a new sub-admin.');
    if (pin && !/^\d{4,8}$/.test(pin)) throw new Error('Sub-admin PIN must be 4 to 8 digits.');
    const patch = {
      UserID: userId, Username: username, Name: name, Role: role, TeamID: teamId,
      Active: active, UpdatedAt: now_(), Notes: String(input.Notes || '').trim(), PIN: ''
    };
    let authChanged = !existing || normalizeUsername_(existing.Username) !== username || normalizeYesNo_(existing.Active, 'Yes') !== active;
    if (pin) {
      const salt = makePinSalt_();
      patch.PINHash = hashPin_(pin, salt);
      patch.PINSalt = salt;
      authChanged = true;
    }
    patch.AuthVersion = authChanged ? Number((existing && existing.AuthVersion) || 0) + 1 : Number(existing.AuthVersion || 1);
    if (existing) updateRow_(SHEETS.AdminUsers, 'UserID', userId, patch);
    else {
      patch.CreatedAt = now_();
      appendObject_(SHEETS.AdminUsers, patch);
    }
    appendAudit_(session, existing ? 'SUBADMIN_UPDATED' : 'SUBADMIN_CREATED', 'AdminUser', userId, { username: username, name: name, teamId: teamId, active: active, pinChanged: !!pin });
    return { ok: true, message: existing ? 'Sub-admin updated. A blank PIN field keeps the current PIN.' : 'Sub-admin created.' };
  } finally {
    lock.releaseLock();
  }
}

// A team may be converted or renamed before its first sub-team period is generated.
// Once ANY rota version (including a draft or an archived one) overlaps the
// effective date, the role names become part of the assignment history.
// Do not reinterpret those historical role strings by changing the configuration.
function subTeamStructureLocked_(team) {
  if (!team || String(team.TeamType) !== 'SubTeam') return false;
  const effective = String(team.SubTeamFrom || '').slice(0,10);
  const names = subTeamNames_(team);
  // The structure becomes permanently locked after its roles have appeared
  // in any stored version. Cache affirmative results to avoid repeatedly
  // scanning archive spreadsheets when administrators open Team Settings.
  const cache=typeof CacheService==='undefined'?null:CacheService.getScriptCache();
  const cacheKey='lockedSubTeam_v3_3_'+String(team.TeamID).slice(0,55)+'_'+effective+'_'+names.join('_').slice(0,90);
  if(cache && cache.get(cacheKey)==='yes')return true;
  // A Standard-rota version covering this date is NOT evidence the sub-team
  // structure was used. Lock only after a version really contains these roles.
  const locked=readObjects_(SHEETS.RotaVersions).some(v => {
    const end = String(v.EndDate || '').slice(0,10);
    if (!end || (effective && end < effective)) return false;
    return getVersionAssignments_(v.VersionID).some(r =>
      String(r.TeamID) === String(team.TeamID) &&
      (!effective || String(r.WeekStart || '').slice(0,10) >= effective) &&
      names.includes(String(r.AssignmentPlace)));
  });
  if(locked && cache)try{cache.put(cacheKey,'yes',3600);}catch(_){}
  return locked;
}

function saveTeam(adminPin, team) {
  const session = requireAdminPin_(adminPin);
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const now = now_();
    const name = String(team.TeamName || '').trim();
    if (!name) throw new Error('Team name is required.');
    const colorHex = normalizeHexColor_(team.ColorHex || '#0f172a');
    const displayOrder = normalizeDisplayOrder_(team.DisplayOrder, team.TeamID ? '' : nextTeamDisplayOrder_());
    const isUpdate = !!team.TeamID;
    const existingTeam = isUpdate ? getTeams_().find(t => String(t.TeamID) === String(team.TeamID)) : null;
    if (isUpdate && !existingTeam) throw new Error('This team no longer exists. Refresh the team list.');
    const teamType = team.TeamType === 'SubTeam' ? 'SubTeam' : (team.TeamType === 'Standard' ? 'Standard' : (existingTeam && existingTeam.TeamType) || 'Standard');
    const hasSubmittedConfig = Object.prototype.hasOwnProperty.call(team, 'SubTeamConfig');
    // An explicit blank is invalid; only omitted values inherit the saved configuration.
    let subConfig = teamType === 'SubTeam' ? String(hasSubmittedConfig ? team.SubTeamConfig : ((existingTeam && existingTeam.SubTeamConfig) || '')) : '';
    const subNames = subConfig.split(',').map(x => x.trim()).filter(Boolean);
    if (teamType === 'SubTeam' && (subNames.length !== 2 || new Set(subNames.map(x=>x.toLowerCase())).size !== 2 || subNames.some(x => ['team','opd','off-duty','reserve','leave','inactive'].includes(x.toLowerCase())) || subNames.some(x=>x.length > 60))) {
      throw new Error('Enter two different sub-team names (up to 60 characters each), excluding reserved assignment names.');
    }
    subConfig = teamType === 'SubTeam' ? subNames.join(',') : '';
    const latestForActivation = getLatestVersion_();
    const previousType = String((existingTeam && existingTeam.TeamType) || 'Standard');
    const previousConfig = String((existingTeam && existingTeam.SubTeamConfig) || '');
    const structureChanged = !!existingTeam && previousType === 'SubTeam' &&
      (teamType !== 'SubTeam' || previousConfig !== subConfig);
    if (structureChanged && subTeamStructureLocked_(existingTeam)) {
      throw new Error('This sub-team structure already has rota versions in its effective period. Names and team type are locked to protect historical assignments. Existing consultants, colours and display order can still be updated.');
    }
    // Default to next rota period. Main Admin can explicitly start on a Sunday
    // within the active period to prepare a selected-team Draft amendment.
    const savedFrom = (previousType === 'SubTeam' && existingTeam.SubTeamFrom) || '';
    const proposedFrom = String(team.SubTeamFrom || '').trim();
    let subFrom = teamType === 'SubTeam'
      ? (proposedFrom || savedFrom || (latestForActivation ? iso_(addDays_(toDate_(latestForActivation.EndDate), 1)) : '')) : '';
    const activationChanged = teamType === 'SubTeam' && previousType === 'SubTeam' &&
      !!proposedFrom && proposedFrom !== String(savedFrom);
    if (subFrom) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(subFrom) || iso_(startOfWeekSunday_(toDate_(subFrom))) !== subFrom) {
        throw new Error('The sub-team activation date must be a Sunday (YYYY-MM-DD).');
      }
      // Past completed weeks are immutable, including via a structure change.
      const todayWeekStart = iso_(startOfWeekSunday_(new Date()));
      if ((previousType !== 'SubTeam' || activationChanged) && subFrom < todayWeekStart) {
        throw new Error('Activation must start this Sunday or a future Sunday; past weeks cannot be converted.');
      }
    }
    if (activationChanged && subTeamStructureLocked_(existingTeam)) {
      throw new Error('This sub-team structure is already used in a rota. Activation cannot be moved; regenerate a new draft instead.');
    }
    // Only allow an explicit current-period activation for an existing team;
    // never interpret a metadata-only edit as an activation change.
    if (proposedFrom && (previousType !== 'SubTeam' || activationChanged) && latestForActivation && subFrom <= String(latestForActivation.EndDate).slice(0,10)) {
      if (!existingTeam) throw new Error('Current-period activation is intended for an existing team.');
      if (subFrom < String(latestForActivation.StartDate).slice(0,10)) throw new Error('Activation precedes the current rota period.');
    }
    let priority = String(team.OneConsultantPriority != null ? team.OneConsultantPriority : (existingTeam && existingTeam.OneConsultantPriority) || '').trim();
    if (teamType !== 'SubTeam') priority = '';
    // A draft already generated for the next period still contains the OLD team roles.
    // Do not silently change its meaning: require deleting/recreating that draft first.
    if (teamType === 'SubTeam' && (previousType !== 'SubTeam' || activationChanged) && subFrom &&
      readObjects_(SHEETS.RotaVersions).some(v => v.Status === 'Draft' && String(v.EndDate || '').slice(0,10) >= subFrom)) {
      throw new Error('A draft rota already exists for the next period. Delete that draft first, configure the existing team sub-teams, then generate a new draft. The published rota will remain unchanged.');
    }
    // Scheduled deactivation changes the future scheduling mode, not the team identity.
    // Keep the original names so historical versions and sub-team fairness remain interpretable.
    const priorUntil = String((existingTeam && existingTeam.SubTeamUntil) || '').slice(0,10);
    const stopSubmitted = Object.prototype.hasOwnProperty.call(team, 'SubTeamUntil');
    const subUntil = teamType === 'SubTeam'
      ? String(stopSubmitted ? team.SubTeamUntil : priorUntil).trim().slice(0,10) : '';
    if (previousType === 'SubTeam' && teamType !== 'SubTeam') {
      throw new Error('Schedule a sub-team deactivation Sunday instead of overwriting the team type. This preserves older published rota assignments.');
    }
    if (subUntil) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(subUntil) ||
          iso_(startOfWeekSunday_(toDate_(subUntil))) !== subUntil) {
        throw new Error('Deactivation date must be a Sunday (YYYY-MM-DD).');
      }
      if (subFrom && subUntil <= subFrom) throw new Error('Deactivation must be after the sub-team activation date.');
      const thisSunday = iso_(startOfWeekSunday_(new Date()));
      if (subUntil < thisSunday && subUntil !== priorUntil) throw new Error('Deactivation cannot change completed rota weeks.');
      if (priorUntil && priorUntil < thisSunday && priorUntil !== subUntil)
        throw new Error('A historical deactivation date cannot be changed.');
      // A draft generated under a different configuration must not be reinterpreted.
      if (subUntil !== priorUntil && readObjects_(SHEETS.RotaVersions).some(v =>
          v.Status === 'Draft' && String(v.EndDate || '').slice(0,10) >= (subUntil || priorUntil)))
        throw new Error('Delete the existing draft before scheduling sub-team deactivation, then create a fresh selected-team draft.');
    }
    if (!subUntil && priorUntil && readObjects_(SHEETS.RotaVersions).some(v =>
        v.Status === 'Draft' && String(v.EndDate || '').slice(0,10) >= priorUntil)) {
      throw new Error('Delete the existing draft before cancelling the sub-team deactivation date.');
    }
    // Optional single reactivation: keep the original start/stop dates and restore
    // the same two named services from a chosen future Sunday. This is deliberately
    // a second interval, never an overwrite of the historical activation date.
    const priorResume = String((existingTeam && existingTeam.SubTeamResumeFrom) || '').slice(0,10);
    const resumeSubmitted = Object.prototype.hasOwnProperty.call(team, 'SubTeamResumeFrom');
    const subResume = teamType === 'SubTeam'
      ? String(resumeSubmitted ? team.SubTeamResumeFrom : priorResume).trim() : '';
    if (subResume) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(subResume) ||
          iso_(startOfWeekSunday_(toDate_(subResume))) !== subResume) {
        throw new Error('Reactivation must be a Sunday (YYYY-MM-DD), or choose a rota week that starts Sunday.');
      }
      if (!subUntil) throw new Error('Schedule a deactivation Sunday before scheduling reactivation.');
      if (subResume <= subUntil) throw new Error('Reactivation must be after the deactivation Sunday.');
      const thisSunday = iso_(startOfWeekSunday_(new Date()));
      if (subResume < thisSunday && subResume !== priorResume)
        throw new Error('Reactivation cannot begin in a completed/past week.');
    }
    if (priorResume && priorResume < iso_(startOfWeekSunday_(new Date())) && subResume !== priorResume)
      throw new Error('A historical reactivation date cannot be changed or removed.');
    if (subResume !== priorResume && readObjects_(SHEETS.RotaVersions).some(v =>
        v.Status === 'Draft' && String(v.EndDate || '').slice(0,10) >= (subResume && priorResume
          ? (subResume < priorResume ? subResume : priorResume) : (subResume || priorResume)))) {
      throw new Error('Delete the existing draft before changing the reactivation date, then create a new selected-team draft.');
    }
    // Renaming sub-teams before activation should not leave an invalid saved default.
    if (priority && !subNames.includes(priority)) {
      if (team.OneConsultantPriority != null) throw new Error('Single-consultant coverage priority must match one of the two sub-team names.');
      priority = '';
    }
    const teamFields = { TeamType:teamType, SubTeamConfig:subConfig, SubTeamFrom:subFrom, OneConsultantPriority:priority, SubTeamUntil:subUntil, SubTeamResumeFrom:subResume };
    const teamId = isUpdate ? String(team.TeamID) : makeId_('T');
    if (isUpdate) updateRow_(SHEETS.Teams, 'TeamID', teamId, Object.assign({ TeamName: name, ColorHex: colorHex, DisplayOrder: displayOrder, Active: normalizeYesNo_(team.Active, 'Yes'), UpdatedAt: now }, teamFields));
    else appendObject_(SHEETS.Teams, Object.assign({ TeamID: teamId, TeamName: name, ColorHex: colorHex, DisplayOrder: displayOrder, Active: normalizeYesNo_(team.Active, 'Yes'), CreatedAt: now, UpdatedAt: now }, teamFields));
    appendAudit_(session, isUpdate ? 'TEAM_UPDATED' : 'TEAM_CREATED', 'Team', teamId, { name: name, displayOrder: displayOrder, active: normalizeYesNo_(team.Active, 'Yes'), teamType:teamType, subTeams:subNames, effectiveFrom:subFrom, inactiveSubTeamsFrom:subUntil, reactivatedSubTeamsFrom:subResume });
    warmPublicStartupCache_();
    return { ok: true, teamId:teamId, teamType:teamType, subTeamNames:subNames, effectiveFrom:subFrom, subTeamsUntil:subUntil, subTeamsResumeFrom:subResume, message: (isUpdate ? 'Existing team updated' : 'Team added') + (teamType === 'SubTeam' ? '. Two sub-teams saved; takes effect ' + (subFrom || 'with the first rota period') + (subUntil ? '; returns to the Standard single-team rota on ' + subUntil : '') + (subResume ? '; reactivates the two inpatient sub-teams on ' + subResume : '') + '. Create and publish a selected-team draft to update the active rota; the existing published rota is unchanged.' : '.') };
  } finally { lock.releaseLock(); }
}

// Main Admin selects which service remains covered for a specific one-consultant week.
// The decision is stored separately so choices may differ from one week to another.
function saveSubTeamCoverageChoice(adminToken, input) {
  const session = requireAdminPin_(adminToken);
  setupDatabase();
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const teamId = String((input && input.teamId) || '').trim();
    const team = getTeams_().find(t => String(t.TeamID) === teamId);
    if (!isSubTeamConfigured_(team)) throw new Error('Choose an existing sub-team parent team.');
    const week = iso_(input && input.weekStart);
    if (iso_(startOfWeekSunday_(week)) !== week) throw new Error('Coverage week must begin on Sunday.');
    const choice = String((input && input.subTeam) || '').trim();
    if (!subTeamNames_(team).includes(choice)) throw new Error('Choose one of this parent team\'s sub-teams.');
    updateSetting_('SubTeamCoverage|' + teamId + '|' + week, choice);
    appendAudit_(session, 'SUBTEAM_COVERAGE_CHOICE', 'Team', teamId, { weekStart:week, subTeam:choice });
    return { ok:true, message:'Coverage selection saved for '+week+'. Future rota generation/amendments will apply it; already-published rotas are not changed.' };
  } finally { lock.releaseLock(); }
}

function saveConsultant(adminPin, c) {
  requireAdminPin_(adminPin);
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const now = now_();
    const name = String(c.Name || '').trim();
    const teamId = String(c.TeamID || '').trim();
    if (!name) throw new Error('Consultant name is required.');
    if (!teamId) throw new Error('Assigned team is required.');
    const colorHex = normalizeHexColor_(c.ColorHex || '#2563eb');
    const newTeam = getTeams_().find(t=>String(t.TeamID)===teamId);
    if (!newTeam) throw new Error('Assigned parent team was not found.');
    const oldSubTeamSetting = c.ConsultantID ? ((findConsultant_(c.ConsultantID)||{}).SubTeamAssignment || 'Any') : 'Any';
    const requestedSubTeam = normalizeSubTeamAssignment_(Object.prototype.hasOwnProperty.call(c,'SubTeamAssignment')?c.SubTeamAssignment:oldSubTeamSetting);
    const subTeamSetting = isSubTeamConfigured_(newTeam)?requestedSubTeam:'Any';
    const oldSubTeam=(c.ConsultantID && findConsultant_(c.ConsultantID))||{};
    const nextUnstartedSunday=iso_(addDays_(startOfWeekSunday_(new Date()),7));
    const subTeamSettingFrom=subTeamSetting==='Any'?'':
      ((subTeamSetting===normalizeSubTeamAssignment_(oldSubTeam.SubTeamAssignment) && oldSubTeam.SubTeamAssignmentFrom)
      ?oldSubTeam.SubTeamAssignmentFrom:(c.ConsultantID?nextUnstartedSunday:(c.ActiveFrom||iso_(new Date()))));
    const currentStatus = String(c.CurrentStatus || 'Active').trim() || 'Active';
    const inactiveFrom = c.InactiveFrom ? iso_(toDate_(c.InactiveFrom)) : '';
    if (currentStatus === 'Inactive' && !inactiveFrom) throw new Error('Inactive status needs an inactive-from date.');

    if (c.ConsultantID) {
      const existing = findConsultant_(c.ConsultantID) || {};
      const activeFrom = c.ActiveFrom ? iso_(toDate_(c.ActiveFrom)) : (existing.ActiveFrom || iso_(new Date()));
      const oldInactiveFrom = existing.InactiveFrom || '';
      const oldActiveFrom = existing.ActiveFrom || '';
      const oldTeamId = existing.TeamID || teamId;
      const oldConsultantType = normalizeConsultantType_(existing.ConsultantType || 'Regular');
      const consultantType = 'Regular';
      updateRow_(SHEETS.Consultants, 'ConsultantID', c.ConsultantID, {
        Name: name,
        TeamID: teamId,
        CanOffDuty: normalizeYesNo_(c.CanOffDuty, 'Yes'),
        SubTeamAssignment: subTeamSetting,
        SubTeamAssignmentFrom: subTeamSettingFrom,
        ConsultantType: consultantType,
        ColorHex: colorHex,
        CurrentStatus: currentStatus,
        ActiveFrom: activeFrom,
        // Save the date whenever admin enters one. The rota engine uses InactiveFrom
        // as the effective date, even if CurrentStatus stays Active until that date.
        InactiveFrom: inactiveFrom,
        Notes: c.Notes || '',
        UpdatedAt: now
      });

      let msg = 'Consultant settings updated.';
      // Consolidate all staffing changes into ONE selected-team draft. A team
      // transfer recalculates both parent teams together; never publish two
      // separate amendments against inconsistent base versions.
      const modifiedDates = [], affectedTeams = {};
      if (inactiveFrom !== oldInactiveFrom) modifiedDates.push(inactiveFrom || nextUnstartedSunday);
      if (currentStatus !== String(existing.CurrentStatus||'Active')) modifiedDates.push(nextUnstartedSunday);
      if (activeFrom !== oldActiveFrom) modifiedDates.push(activeFrom);
      if (oldTeamId !== teamId) modifiedDates.push(activeFrom);
      if (consultantType !== oldConsultantType || normalizeYesNo_(c.CanOffDuty,'Yes') !== String(existing.CanOffDuty||'Yes') || subTeamSetting !== normalizeSubTeamAssignment_(existing.SubTeamAssignment))
        modifiedDates.push(iso_(new Date()));
      if (modifiedDates.length) {
        // Capture the old entitlement before the date-effective change; the
        // new record becomes active no earlier than the next unstarted Sunday.
        let eligibilityRows=[];
        const settingChanged = consultantType!==oldConsultantType ||
          normalizeYesNo_(c.CanOffDuty,'Yes')!==String(existing.CanOffDuty||'Yes') ||
          subTeamSetting!==normalizeSubTeamAssignment_(existing.SubTeamAssignment) || oldTeamId!==teamId;
        try {
          if(settingChanged && typeof SpreadsheetApp!=='undefined')
            eligibilityRows=recordEligibilityChange_(existing,Object.assign({},existing,{
              TeamID:teamId,ConsultantType:consultantType,CanOffDuty:normalizeYesNo_(c.CanOffDuty,'Yes'),
              SubTeamAssignment:subTeamSetting}),nextUnstartedSunday);
        } catch(err){updateRow_(SHEETS.Consultants,'ConsultantID',c.ConsultantID,existing);throw err;}
        affectedTeams[String(teamId)] = true;
        if (oldTeamId !== teamId) affectedTeams[String(oldTeamId)] = true;
        try {
          const earliest = modifiedDates.sort()[0];
          const amend = createStaffingDraftFromDateIfInsideRota_(earliest,Object.keys(affectedTeams),'Consultant staffing draft',
            name+' staffing updated ('+Object.keys(affectedTeams).join(', ')+')');
          if (amend.message) msg += ' ' + amend.message;
        } catch (err) {
          updateRow_(SHEETS.Consultants,'ConsultantID',c.ConsultantID,existing);
          if(eligibilityRows.length)deleteRowsWhere_(SHEETS.EligibilityHistory,r=>eligibilityRows.includes(String(r.HistoryID)));
          throw new Error('Consultant change reversed; no official rota changed: '+((err&&err.message)||String(err)));
        }
      }
      appendAudit_(authenticateAdmin_(adminPin), 'CONSULTANT_UPDATED', 'Consultant', c.ConsultantID, { name: name, teamId: teamId, consultantType: consultantType, canOffDuty: normalizeYesNo_(c.CanOffDuty, 'Yes'), subTeamAssignment:subTeamSetting, status: currentStatus, activeFrom: activeFrom, inactiveFrom: inactiveFrom });
      // Public cache was invalidated by sheet writes; rebuild on next read.

      return { ok: true, message: msg };
    }

    const activeFrom = c.ActiveFrom ? iso_(toDate_(c.ActiveFrom)) : iso_(new Date());
    const consultantType = normalizeConsultantType_(c.ConsultantType || 'Regular');
    const newId = makeId_('C');
    appendObject_(SHEETS.Consultants, {
      ConsultantID: newId,
      Name: name,
      TeamID: teamId,
      CanOffDuty: normalizeYesNo_(c.CanOffDuty, 'Yes'),
      SubTeamAssignment: subTeamSetting,
      SubTeamAssignmentFrom: subTeamSettingFrom,
      ConsultantType: consultantType,
      ColorHex: colorHex,
      CurrentStatus: currentStatus,
      ActiveFrom: activeFrom,
      InactiveFrom: inactiveFrom,
      Notes: c.Notes || '',
      CreatedAt: now,
      UpdatedAt: now
    });

    let msg = inactiveFrom ? 'Consultant added with inactive-from date.' : 'Consultant added.';
    try {
      const amend = createStaffingDraftFromDateIfInsideRota_(activeFrom,[teamId],'Consultant added draft',
        name+' added to '+getTeamName_(teamId)+' from '+activeFrom);
      if (amend.message) msg += ' ' + amend.message;
    } catch (err) {
      deleteRowsWhere_(SHEETS.Consultants,row=>String(row.ConsultantID)===String(newId));
      throw new Error('Consultant addition reversed; no official rota changed: '+((err&&err.message)||String(err)));
    }
    appendAudit_(authenticateAdmin_(adminPin), 'CONSULTANT_CREATED', 'Consultant', newId, { name: name, teamId: teamId, consultantType: consultantType, canOffDuty: normalizeYesNo_(c.CanOffDuty, 'Yes'), subTeamAssignment:subTeamSetting, status: currentStatus, activeFrom: activeFrom, inactiveFrom: inactiveFrom });
    // Public cache was invalidated by sheet writes; rebuild on next read.

    return { ok: true, message: msg };
  } finally {
    lock.releaseLock();
  }
}

function startNewRota(adminPin, params) {
  // Backward-compatible API name kept for old cached frontends.
  // From v1.2 onward, direct generation creates a Draft only. The official rota
  // changes only after admin reviews the draft and clicks Publish Draft.
  return createDraftRota(adminPin, params);
}

// Admin-controlled future-only reshuffle of ONE team within the current official period.
// This always creates a reviewable Draft; it NEVER publishes the official rota.
function rearrangeTeamRota(adminPin, params) {
  requireAdminPin_(adminPin);
  const latest = getLatestVersion_();
  if (!latest) throw new Error('A published rota is required before rearranging a team.');
  const teamId = String((params && params.teamId) || '').trim();
  if (!teamId || teamId === 'ALL' || !getTeams_().some(t => t.Active === 'Yes' && String(t.TeamID) === teamId)) {
    throw new Error('Choose one active team. Rearranging all teams is not supported.');
  }
  const weekNo = Number(params && params.weekNo);
  const weeks = Number(latest.Weeks || 52);
  if (!Number.isInteger(weekNo) || weekNo < 1 || weekNo > weeks) throw new Error('Choose a valid week in the published rota.');
  const start = toDate_(latest.StartDate);
  const chosenSunday = addDays_(start, (weekNo - 1) * 7);
  const currentSunday = startOfWeekSunday_(new Date());
  if (chosenSunday < currentSunday) throw new Error('Completed/past weeks are protected. Choose the current or a future week.');
  if (chosenSunday > toDate_(latest.EndDate)) throw new Error('Selected week is outside the active rota period.');
  return createDraftRota(adminPin, {
    teamId: teamId, startDate: iso_(chosenSunday), weeks: weeks,
    manualSeedWeeks: 0, manualAssignments: []
  });
}

function createDraftRota(adminPin, params) {
  requireAdminPin_(adminPin);
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const existingDraft = getRotaVersions_().find(v => v.Status === 'Draft');
    if (existingDraft) {
      const periodText = existingDraft.PeriodName ? ` in ${existingDraft.PeriodName}` : '';
      throw new Error(`There is already a draft rota: Amendment ${existingDraft.AmendmentNo}${periodText}. Publish or delete that draft before creating another one.`);
    }

    const selectedTeamId = String((params && params.teamId) || 'ALL').trim() || 'ALL';
    const start = startOfWeekSunday_(toDate_(params.startDate));
    const requestedWeeks = Math.max(1, Number(params.weeks || 52));
    const manualSeedWeeks = Math.min(4, Math.max(0, Number(params.manualSeedWeeks || 0)));
    const manualAssignments = params.manualAssignments || [];
    const latest = getLatestVersion_();
    const allActiveTeams = getTeams_().filter(t => t.Active === 'Yes');

    if (selectedTeamId !== 'ALL' && !allActiveTeams.some(t => t.TeamID === selectedTeamId)) {
      throw new Error('Selected team is not active or was not found.');
    }

    // Generating ALL teams creates a completely separate Draft rota period. The currently
    // active period and its latest official amendment remain untouched until Publish Draft.
    if (selectedTeamId === 'ALL' || !latest) {
      const periodEnd = iso_(addDays_(start, (requestedWeeks * 7) - 1));
      const period = createRotaPeriod_({
        PeriodName: String((params && params.periodName) || '').trim() || defaultPeriodName_(iso_(start), periodEnd),
        Status: 'Draft',
        StartDate: iso_(start),
        EndDate: periodEnd,
        Weeks: requestedWeeks,
        Notes: 'New rota period draft. It becomes active only after Publish Draft.'
      });
      const teamIds = selectedTeamId === 'ALL' ? null : [selectedTeamId];
      const version = createRotaVersion_({
        PeriodID: period.PeriodID,
        AmendmentNo: 1,
        Status: 'Draft',
        Type: selectedTeamId === 'ALL' ? 'Draft new rota period' : 'Draft new single-team rota period',
        StartDate: period.StartDate,
        EndDate: period.EndDate,
        Weeks: requestedWeeks,
        EffectiveFrom: period.StartDate,
        Reason: manualSeedWeeks > 0 ? `New period draft generated with ${manualSeedWeeks} manual seed week(s)` : 'New rota period draft generated',
        BasedOnVersionID: latest ? latest.VersionID : ''
      });

      let result;
      try {
        result = buildRotaRows_({
          versionId: version.VersionID,
          startDate: start,
          weeks: requestedWeeks,
          fromWeekNo: 1,
          previousRows: [],
          manualSeedWeeks: manualSeedWeeks,
          manualAssignments: manualAssignments,
          leavesOverride: getApprovedLeaves_(),
          teamIds: teamIds,
          // Retain the Off-duty cooldown from the preceding official rota
          // for the first two weeks of a new scheduling period.
          offCooldownWeeks: latest ? offDutyCarryOverWeeks_(getVersionAssignments_(latest.VersionID), start) : {},
          inpatientCooldownWeeks: latest ? inpatientCarryOverWeeks_(getVersionAssignments_(latest.VersionID), start) : {},
          inpatientCarryOverGapWeeks: latest ? inpatientCarryOverGapWeeks_(getVersionAssignments_(latest.VersionID), start) : {},
          roleCarryOver: latest ? previousPeriodRoleCarryOver_(getVersionAssignments_(latest.VersionID), start) : {}
        });
      } catch (err) {
        deleteRowsWhere_(SHEETS.RotaVersions, row => row.VersionID === version.VersionID);
        deleteRowsWhere_(SHEETS.RotaPeriods, row => row.PeriodID === period.PeriodID);
        throw err;
      }

      if (!result.rows.length) {
        deleteRowsWhere_(SHEETS.RotaVersions, row => row.VersionID === version.VersionID);
        deleteRowsWhere_(SHEETS.RotaPeriods, row => row.PeriodID === period.PeriodID);
        throw new Error('Safety stop: the new-period draft generated no rota assignments. Nothing was published.');
      }
      appendRows_(SHEETS.RotaAssignments, result.rows);
      appendConflicts_(version.VersionID, result.conflicts);
      appendAudit_(authenticateAdmin_(adminPin), 'ROTA_DRAFT_CREATED', 'RotaVersion', version.VersionID, { periodId: period.PeriodID, periodName: period.PeriodName, amendmentNo: 1, teamId: selectedTeamId, weeks: requestedWeeks, conflictCount: result.conflicts.length });
      return {
        ok: true,
        message: `${period.PeriodName} created as a Draft period with Amendment 1. The current official rota remains live until you publish this draft. Conflicts/warnings: ${result.conflicts.length}`,
        period: period,
        version: enrichVersionWithPeriod_(version),
        conflicts: result.conflicts
      };
    }

    // Selected-team generation is a Draft amendment inside the currently active rota period.
    // Other teams are copied exactly from the latest official amendment.
    const latestStart = toDate_(latest.StartDate);
    const latestEnd = toDate_(latest.EndDate);
    if (start < latestStart || start > latestEnd) {
      throw new Error('For selected-team generation, the start date must be inside the current active rota period.');
    }
    if (start < startOfWeekSunday_(new Date()))
      throw new Error('Completed/past weeks are protected. Select this Sunday or a future Sunday for the team draft.');
    const fromWeekNo = Math.floor((start.getTime() - latestStart.getTime()) / (7 * 24 * 60 * 60 * 1000)) + 1;
    const weeks = Number(latest.Weeks || requestedWeeks || 52);

    const version = createRotaVersion_({
      PeriodID: latest.PeriodID,
      AmendmentNo: nextAmendmentNoForPeriod_(latest.PeriodID),
      Status: 'Draft',
      Type: 'Draft selected team amendment',
      StartDate: latest.StartDate,
      EndDate: latest.EndDate,
      Weeks: weeks,
      EffectiveFrom: iso_(start),
      Reason: `Draft generated for ${getTeamName_(selectedTeamId)} from Week ${fromWeekNo}; other teams unchanged`,
      BasedOnVersionID: latest.VersionID
    });

    const previous = getVersionAssignments_(latest.VersionID);
    const fixedRows = previous
      .filter(r => !(r.TeamID === selectedTeamId && Number(r.WeekNo) >= fromWeekNo))
      .map(r => Object.assign({}, r, { VersionID: version.VersionID, PeriodID: latest.PeriodID }));

    let result;
    try {
      result = buildRotaRows_({
        versionId: version.VersionID,
        startDate: latestStart,
        weeks: weeks,
        fromWeekNo: fromWeekNo,
        previousRows: fixedRows,
        manualSeedWeeks: fromWeekNo + manualSeedWeeks - 1,
        manualAssignments: manualAssignments,
        leavesOverride: getApprovedLeaves_(),
        teamIds: [selectedTeamId]
      });
    } catch (err) {
      deleteRowsWhere_(SHEETS.RotaVersions, row => row.VersionID === version.VersionID);
      throw err;
    }

    if (!result.rows.length) {
      deleteRowsWhere_(SHEETS.RotaVersions, row => row.VersionID === version.VersionID);
      throw new Error('Safety stop: the selected-team draft generated no rota assignments. Nothing was published.');
    }
    // Guard the published baseline before persisting a new draft: prior weeks
    // and assignments belonging to other teams must remain byte-for-byte equivalent.
    try {
      assertRowsBeforeEffectiveWeekUnchanged_(previous, result.rows, fromWeekNo);
      const oldOther=previous.filter(r=>String(r.TeamID)!==selectedTeamId);
      const newOther=result.rows.filter(r=>String(r.TeamID)!==selectedTeamId);
      if (assignmentSignatureAll_(oldOther)!==assignmentSignatureAll_(newOther))
        throw new Error('Safety stop: rearrangement modified another team.');
    } catch(e) {
      deleteRowsWhere_(SHEETS.RotaVersions, row => row.VersionID === version.VersionID);
      throw e;
    }
    appendRows_(SHEETS.RotaAssignments, result.rows);
    appendConflicts_(version.VersionID, result.conflicts);
    appendAudit_(authenticateAdmin_(adminPin), 'ROTA_DRAFT_CREATED', 'RotaVersion', version.VersionID, { periodId: latest.PeriodID, amendmentNo: version.AmendmentNo, teamId: selectedTeamId, fromWeekNo: fromWeekNo, conflictCount: result.conflicts.length });
    return {
      ok: true,
      message: `${getTeamName_(selectedTeamId)} draft Amendment ${version.AmendmentNo} created inside ${latest.PeriodName || 'the active rota period'} from Week ${fromWeekNo}. Users still see the current official amendment until you publish. Conflicts/warnings: ${result.conflicts.length}`,
      version: enrichVersionWithPeriod_(version),
      conflicts: result.conflicts
    };
  } finally {
    lock.releaseLock();
  }
}

function publishDraftRota(adminPin, versionId) {
  const session = requireAdminPin_(adminPin);
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const draft = getRotaVersions_().find(v => v.VersionID === versionId);
    if (!draft) throw new Error('Draft rota version was not found.');
    if (draft.Status !== 'Draft') throw new Error('Only draft rota versions can be published.');
    if (!getVersionAssignments_(draft.VersionID).length) throw new Error('Safety stop: this draft has no rota assignments, so it cannot be published.');
    const latest = getLatestVersion_(), basedOn = String(draft.BasedOnVersionID || '').trim(), draftPeriod = getPeriodById_(draft.PeriodID);
    if (!draftPeriod) throw new Error('The draft rota period was not found. Run Safe Repair / Add Missing Columns, then try again.');
    const activePeriod = getActivePeriod_();
    const isNewPeriod = draftPeriod.Status === 'Draft' && (!activePeriod || String(draftPeriod.PeriodID) !== String(activePeriod.PeriodID));
    if (basedOn) {
      if (!latest || latest.VersionID !== basedOn) throw new Error('This draft is outdated because the official rota changed after it was created. Delete this draft and create a new one from the latest official rota.');
    } else if (latest) throw new Error('This draft was created when no official rota existed, but an official rota exists now. Delete this draft and create a new one from the latest official rota.');
    if (!isNewPeriod && activePeriod && String(draft.PeriodID) !== String(activePeriod.PeriodID)) throw new Error('This amendment draft does not belong to the active rota period. Delete it and create a new draft from the active period.');

    // Publish the validated draft first while the document lock is held. If a later archive/period
    // step fails, the catch block reverts it to Draft and restores the old official state.
    const oldOfficials = getRotaVersions_().filter(v => v.Status === 'Official');
    const oldPeriods = getRotaPeriods_().map(p => ({ PeriodID: p.PeriodID, Status: p.Status, ActivatedAt: p.ActivatedAt, ArchivedAt: p.ArchivedAt }));
    try {
      updateRow_(SHEETS.RotaVersions, 'VersionID', draft.VersionID, { Status: 'Official', Reason: `${draft.Reason || ''}${draft.Reason ? ' | ' : ''}PUBLISHED: Admin published this draft (${now_()})` });
      if (latest) getRotaVersions_().filter(v => v.Status === 'Official' && v.VersionID !== draft.VersionID && String(v.PeriodID || '') === String(latest.PeriodID || '')).forEach(v => archiveVersion_(v.VersionID));
      if (isNewPeriod) {
        if (activePeriod) archivePeriod_(activePeriod.PeriodID);
        activatePeriod_(draftPeriod.PeriodID);
        updateSetting_('CurrentExportDocId', ''); updateSetting_('CurrentExportDocUrl', '');
      }
    } catch (publishErr) {
      try { updateRow_(SHEETS.RotaVersions, 'VersionID', draft.VersionID, { Status: 'Draft', Reason: draft.Reason || '' }); } catch (_) {}
      oldOfficials.forEach(v => restoreOfficialAfterRollback_(v));
      oldPeriods.forEach(p => { try { updateRow_(SHEETS.RotaPeriods, 'PeriodID', p.PeriodID, { Status: p.Status, ActivatedAt: p.ActivatedAt, ArchivedAt: p.ArchivedAt }); } catch (_) {} });
      throw publishErr;
    }
    const latestAfter = getLatestVersion_();
    appendAudit_(session, 'ROTA_DRAFT_PUBLISHED', 'RotaVersion', draft.VersionID, { periodId: draft.PeriodID, periodName: draftPeriod.PeriodName, amendmentNo: draft.AmendmentNo, newPeriod: isNewPeriod });
    warmPublicStartupCache_();
    return { ok: true, message: isNewPeriod ? `${draftPeriod.PeriodName} is now the active rota period. The previous period was archived, and Amendment 1 is live.` : `Draft Amendment ${draft.AmendmentNo} published inside ${draftPeriod.PeriodName}. Users will now see it as the latest official amendment.`, latestVersion: latestAfter };
  } finally { lock.releaseLock(); }
}

function consultantUnavailableForLeave_(c) {
  if (!c) return true;
  if (String(c.CurrentStatus || '').toLowerCase() === 'inactive') return true;
  const inactiveFrom = String(c.InactiveFrom || '').slice(0,10);
  return !!inactiveFrom && inactiveFrom <= iso_(new Date());
}

function submitLeaveRequest(data) {
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const consultantId = String(data.ConsultantID || '').trim();
    const consultant = getConsultants_().find(c => c.ConsultantID === consultantId);
    if (!consultant) throw new Error('Please select a valid consultant.');
    if (consultantUnavailableForLeave_(consultant)) throw new Error('Inactive consultants cannot submit or edit leave requests. Contact the administrator if reactivated.');
    const start = toDate_(data.StartDate);
    const end = toDate_(data.EndDate);
    if (end < start) throw new Error('Leave end date cannot be before start date.');
    const candidate = {
      LeaveID: makeId_('L'), ConsultantID: consultant.ConsultantID, TeamID: consultant.TeamID,
      StartDate: iso_(start), EndDate: iso_(end), Status: 'Pending', SubmittedAt: now_(),
      UserNote: data.Note || '', AdminDecisionDate: '', AdminNote: '', ImpactStatus: '', ImpactSummary: ''
    };
    const precheck = analyzeLeaveCandidate_(candidate, { userMode: true });
    candidate.ImpactStatus = precheck.level || '';
    candidate.ImpactSummary = precheck.summary || '';
    appendObject_(SHEETS.LeaveRequests, candidate);
    appendAudit_({ name: consultant.Name, role: 'Consultant', teamId: consultant.TeamID }, 'LEAVE_SUBMITTED', 'LeaveRequest', candidate.LeaveID, { startDate: candidate.StartDate, endDate: candidate.EndDate });
    return { ok: true, message: 'Leave request submitted and pending admin approval.', leave: candidate, precheck: precheck };
  } finally {
    lock.releaseLock();
  }
}

function checkLeaveAvailability(data) {
  setupDatabase();
  const consultantId = String(data.ConsultantID || '').trim();
  const consultant = getConsultants_().find(c => c.ConsultantID === consultantId);
  if (!consultant) throw new Error('Please select your name.');
  if (consultantUnavailableForLeave_(consultant)) throw new Error('Inactive consultants cannot submit or edit leave requests. Contact the administrator if reactivated.');
  const start = toDate_(data.StartDate);
  const end = toDate_(data.EndDate);
  if (end < start) throw new Error('Leave end date cannot be before start date.');
  const candidate = {
    LeaveID: data.LeaveID || 'PRECHECK',
    ConsultantID: consultant.ConsultantID,
    TeamID: consultant.TeamID,
    StartDate: iso_(start),
    EndDate: iso_(end),
    Status: 'Pending',
    SubmittedAt: '',
    UserNote: data.Note || '',
    AdminDecisionDate: '',
    AdminNote: '',
    ImpactStatus: '',
    ImpactSummary: ''
  };
  return analyzeLeaveCandidate_(candidate, { userMode: true });
}

function getLeaveRequestsForConsultant(consultantId) {
  setupDatabase();
  const raw = String(consultantId || '').trim();
  if (!raw) return [];
  const rawLower = raw.toLowerCase();
  const consultant = getConsultants_().find(c =>
    String(c.ConsultantID || '').trim() === raw ||
    String(c.Name || '').trim().toLowerCase() === rawLower
  );
  if (!consultant) throw new Error('Please select a valid consultant.');
  const id = String(consultant.ConsultantID || '').trim();
  return readObjects_(SHEETS.LeaveRequests)
    .filter(l => {
      const rowConsultantId = String(l.ConsultantID || '').trim();
      const rowConsultantName = String(l.ConsultantName || '').trim().toLowerCase();
      const status = normalizeStatus_(l.Status);
      const sameConsultant = rowConsultantId === id || (!!rowConsultantName && rowConsultantName === String(consultant.Name || '').trim().toLowerCase());
      return sameConsultant && (status === 'pending' || status === 'rejected' || status === 'approved');
    })
    .sort((a, b) => String(b.SubmittedAt || '').localeCompare(String(a.SubmittedAt || '')))
    .map(l => ({
      LeaveID: l.LeaveID,
      ConsultantID: id,
      TeamID: l.TeamID || consultant.TeamID || '',
      StartDate: l.StartDate,
      EndDate: l.EndDate,
      Status: statusDisplay_(l.Status),
      SubmittedAt: l.SubmittedAt,
      UserNote: l.UserNote || l.Note || '',
      AdminNote: l.AdminNote || '',
      ImpactStatus: l.ImpactStatus || '',
      ImpactSummary: l.ImpactSummary || ''
    }));
}

function updateLeaveRequest(data) {
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const leaveId = String(data.LeaveID || '').trim();
    const consultantId = String(data.ConsultantID || '').trim();
    if (!leaveId) throw new Error('Select a pending/rejected leave request to edit.');
    const consultant = getConsultants_().find(c => c.ConsultantID === consultantId);
    if (!consultant) throw new Error('Please select a valid consultant.');
    if (consultantUnavailableForLeave_(consultant)) throw new Error('Inactive consultants cannot submit or edit leave requests. Contact the administrator if reactivated.');
    const leave = readObjects_(SHEETS.LeaveRequests).find(l => l.LeaveID === leaveId);
    if (!leave) throw new Error('Leave request was not found.');
    if (leave.ConsultantID !== consultant.ConsultantID) throw new Error('This leave request does not belong to the selected consultant.');
    if (!['pending','rejected'].includes(normalizeStatus_(leave.Status))) throw new Error('Only pending or rejected leave requests can be edited. Approved leave is archived in Admin and cannot be edited by users.');
    const start = toDate_(data.StartDate), end = toDate_(data.EndDate);
    if (end < start) throw new Error('Leave end date cannot be before start date.');
    const candidate = { LeaveID: leave.LeaveID, ConsultantID: consultant.ConsultantID, TeamID: consultant.TeamID, StartDate: iso_(start), EndDate: iso_(end), Status: 'Pending', SubmittedAt: now_(), UserNote: data.Note || '', AdminDecisionDate: '', AdminNote: '', ImpactStatus: '', ImpactSummary: '' };
    const precheck = analyzeLeaveCandidate_(candidate, { userMode: true });
    updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, { TeamID: consultant.TeamID, StartDate: candidate.StartDate, EndDate: candidate.EndDate, Status: 'Pending', SubmittedAt: candidate.SubmittedAt, UserNote: candidate.UserNote, AdminDecisionDate: '', AdminNote: '', ImpactStatus: precheck.level || '', ImpactSummary: precheck.summary || '' });
    appendAudit_({ name: consultant.Name, role: 'Consultant', teamId: consultant.TeamID }, 'LEAVE_UPDATED', 'LeaveRequest', leaveId, { startDate: candidate.StartDate, endDate: candidate.EndDate });
    return { ok: true, message: 'Leave request updated and returned to pending admin review.', precheck: precheck };
  } finally { lock.releaseLock(); }
}


function analyzeLeaveCandidate_(leave, options) {
  // User-facing pre-check is intentionally overlap-only.
  // It does not simulate or expose how the rota will look after approval.
  // Final rota impact remains an admin-only approval function.
  options = options || {};
  const latest = getLatestVersion_();
  const overlaps = getLeaveOverlapStats_(leave);
  let status = 'Safe';
  let level = 'Safe';
  if (overlaps.approvedCount > 0 || overlaps.pendingCount > 0) {
    level = 'Warning';
    status = 'Overlap found';
  }

  let rotaScope = 'Inside current rota';
  if (!latest) {
    rotaScope = 'No current rota yet';
  } else {
    const currentStart = toDate_(latest.StartDate);
    const currentEnd = toDate_(latest.EndDate);
    const leaveStart = toDate_(leave.StartDate);
    const leaveEnd = toDate_(leave.EndDate);
    if (leaveEnd < currentStart || leaveStart > currentEnd) rotaScope = 'Outside current rota period';
    else if (leaveStart < currentStart || leaveEnd > currentEnd) rotaScope = 'Partially inside current rota period';
  }

  const summary = summarizeUserLeaveCheck_(level, overlaps, [], rotaScope);
  return {
    status: status,
    level: level,
    summary: summary,
    overlaps: overlaps,
    rotaScope: rotaScope,
    hiddenDetails: true
  };
}

function getLeaveOverlapStats_(leave) {
  const leaveStart = toDate_(leave.StartDate);
  const leaveEnd = toDate_(leave.EndDate);
  const rows = readObjects_(SHEETS.LeaveRequests).filter(l => {
    if (l.LeaveID && leave.LeaveID && l.LeaveID === leave.LeaveID) return false;
    if (l.TeamID !== leave.TeamID) return false;
    // Only compare against different consultants from the same team.
    if (l.ConsultantID === leave.ConsultantID) return false;
    const status = normalizeStatus_(l.Status);
    if (status === 'rejected' || status === 'cancelled') return false;
    return datesOverlap_(toDate_(l.StartDate), toDate_(l.EndDate), leaveStart, leaveEnd);
  });
  return {
    approvedCount: rows.filter(l => normalizeStatus_(l.Status) === 'approved').length,
    pendingCount: rows.filter(l => normalizeStatus_(l.Status) === 'pending').length,
    totalCount: rows.length,
    overlaps: rows.map(l => ({
      status: l.Status,
      startDate: l.StartDate,
      endDate: l.EndDate
    })),
    privacyNote: 'User view shows overlap counts and dates only. Other consultants names are hidden; admin sees full details during approval.'
  };
}

function summarizeUserLeaveCheck_(level, overlaps, conflicts, fallback) {
  const overlapParts = [];
  if (overlaps.approvedCount) overlapParts.push(`${overlaps.approvedCount} approved overlapping leave request(s)`);
  if (overlaps.pendingCount) overlapParts.push(`${overlaps.pendingCount} pending overlapping leave request(s)`);
  if (!overlapParts.length) {
    return `No overlapping leave request was found from another consultant in the same team. Final approval remains with admin. Scope: ${fallback}.`;
  }
  return `Overlap found with ${overlapParts.join(' and ')} from another consultant in the same team. This may make approval less likely, but final approval remains with admin. Scope: ${fallback}.`;
}

function previewLeaveImpact(adminPin, leaveId) {
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const leave = readObjects_(SHEETS.LeaveRequests).find(l => l.LeaveID === leaveId);
    if (!leave) throw new Error('Leave request was not found.');
    requireLeaveAdmin_(adminPin, leave);
    const latest = getLatestVersion_();
    if (!latest) return { status: 'No current rota', level: 'Warning', summary: 'There is no generated rota yet. This leave will be stored and applied when a rota is generated.', conflicts: [], preview: [] };
  
    const start = toDate_(latest.StartDate);
    const weeks = Number(latest.Weeks || 52);
    const leaveStart = toDate_(leave.StartDate);
    const leaveEnd = toDate_(leave.EndDate);
    const currentEnd = toDate_(latest.EndDate);
  
    if (leaveStart > currentEnd) {
      return { status: 'Outside current rota', level: 'Info', summary: 'This leave is outside the current rota period. Approval will store it for a future rota and will not amend the current rota.', conflicts: [], preview: [] };
    }
  
    const effectiveDate = leaveStart < start ? start : leaveStart;
    const tempLeaves = getApprovedLeaves_().concat([leave]);
    const simulation = simulateAmendment_(latest, effectiveDate, tempLeaves, null, [leave.TeamID]);
    const affectedWeekStart = getWeekNoForDate_(start, effectiveDate, weeks);
    const previewRows = simulation.rows.filter(r => Number(r.WeekNo) >= affectedWeekStart && Number(r.WeekNo) <= affectedWeekStart + 2 && r.TeamID === leave.TeamID);
    const conflicts = simulation.conflicts.filter(c => c.TeamID === leave.TeamID);
    const level = deriveImpactLevel_(conflicts, leave, latest);
    const summary = summarizeImpact_(level, conflicts, leave, latest);
    updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, { ImpactStatus: level, ImpactSummary: summary });
    return { status: level, level: level, summary: summary, conflicts: conflicts.slice(0, 20), preview: groupRowsForTeamPreview_(previewRows),
      outlook: leaveStaffingOutlook_(simulation, latest, leave, affectedWeekStart) };
  } finally {
    lock.releaseLock();
  }
}

// Read-only "what if" check: shows how the current official rota would change if a
// consultant took leave on the given dates. It creates no leave request, writes nothing
// to any sheet and publishes nothing. The recalculation starts where a real approval
// would start (never before the next unstarted week).
function previewLeaveWhatIf(adminToken, data) {
  return withAdminReadMemo_(function() {
    data = data || {};
    const consultant = getConsultants_().find(c => String(c.ConsultantID) === String(data.ConsultantID || '').trim());
    if (!consultant) throw new Error('Select a consultant.');
    requireLeaveAdmin_(adminToken, { TeamID: consultant.TeamID });
    const startText = String(data.StartDate || '').trim().slice(0, 10), endText = String(data.EndDate || '').trim().slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(startText) || !/^\d{4}-\d{2}-\d{2}$/.test(endText)) throw new Error('Enter a start date and an end date.');
    const leaveStart = toDate_(startText), leaveEnd = toDate_(endText);
    if (iso_(leaveStart) !== startText || iso_(leaveEnd) !== endText) throw new Error('The dates are not valid.');
    if (leaveEnd < leaveStart) throw new Error('Leave end date cannot be before start date.');
    const leave = { LeaveID: 'WHATIF', ConsultantID: consultant.ConsultantID, TeamID: consultant.TeamID,
      StartDate: startText, EndDate: endText, Status: 'Approved' };
    const latest = getLatestVersion_();
    const nothing = (level, summary) => ({ status: level, level: level, summary: summary, conflicts: [], preview: [], outlook: null, whatIf: true });
    if (!latest) return nothing('Info', 'There is no generated rota yet, so there is nothing to compare against.');
    const start = toDate_(latest.StartDate), weeks = Number(latest.Weeks || 52), currentEnd = toDate_(latest.EndDate);
    if (leaveStart > currentEnd) return nothing('Info', 'These dates are after the current rota period, so the current rota would not change.');
    const safeStart = addDays_(startOfWeekSunday_(new Date()), 7);
    const requested = startOfWeekSunday_(leaveStart < start ? start : leaveStart);
    const effective = safeStart > requested ? safeStart : requested;
    if (effective > currentEnd || leaveEnd < effective) return nothing('Info', 'These dates only touch weeks that have already started. Published weeks are not recalculated automatically; use Manual Edit for the current week.');
    const tempLeaves = getApprovedLeaves_().concat([leave]);
    const simulation = simulateAmendment_(latest, effective, tempLeaves, null, [String(consultant.TeamID)]);
    const fromWeek = Math.max(1, getWeekNoForDate_(start, effective, weeks));
    const outlook = leaveStaffingOutlook_(simulation, latest, leave, fromWeek);
    const conflicts = simulation.conflicts.filter(c => String(c.TeamID) === String(consultant.TeamID) &&
      Number(c.WeekNo) >= fromWeek && Number(c.WeekNo) <= outlook.toWeek);
    const level = deriveImpactLevel_(conflicts, leave, latest);
    const previewRows = simulation.rows.filter(r => String(r.TeamID) === String(consultant.TeamID) &&
      Number(r.WeekNo) >= fromWeek && Number(r.WeekNo) <= outlook.toWeek);
    return { status: level, level: level, whatIf: true,
      summary: 'What-if only: nothing was saved. ' + consultant.Name + ' on leave ' + startText + ' to ' + endText + '.',
      conflicts: conflicts.slice(0, 20), preview: groupRowsForTeamPreview_(previewRows), outlook: outlook };
  });
}

// Plain-language staffing picture for a simulated leave: for each week from the first
// recalculated week until the week after the leave ends, how many consultants remain,
// who works inpatient in consecutive weeks, and whether OPD or the Off-duty reserve is
// empty. Pure calculation over simulation rows; it never reads or writes a sheet itself.
function leaveStaffingOutlook_(simulation, latest, leave, fromWeekNo) {
  const teamId = String(leave.TeamID), totalWeeks = Number(latest.Weeks || 52), rotaStart = toDate_(latest.StartDate);
  const leaveEndWeek = Math.max(1, getWeekNoForDate_(rotaStart, toDate_(leave.EndDate), totalWeeks));
  const fromWeek = Math.max(1, Number(fromWeekNo || 1));
  // Show at most 12 weeks, always including the return week when it is inside the rota.
  const toWeek = Math.min(totalWeeks, Math.max(fromWeek, Math.min(leaveEndWeek + 1, fromWeek + 11)));
  const consultants = consultantMap_();
  const nameOf = id => { const c = consultants[String(id)]; return c ? c.Name : String(id); };
  const team = getTeams_().find(t => String(t.TeamID) === teamId);
  // Inpatient = the parent 'Team' post or a registered sub-team service. Any other
  // label (e.g. the retired 'Reserve') is neither inpatient nor OPD/Off-duty.
  const inpatientPlaces = ['Team', 'TCU', 'EMCU'].concat(isSubTeamConfigured_(team) ? subTeamNames_(team) : []);
  const byWeek = {};
  (simulation.rows || []).forEach(r => {
    if (String(r.TeamID) !== teamId) return;
    const w = Number(r.WeekNo);
    if (w < fromWeek - 1 || w > toWeek) return;
    if (!byWeek[w]) byWeek[w] = { weekNo: w, weekStart: String(r.WeekStart).slice(0, 10), weekEnd: String(r.WeekEnd).slice(0, 10), inpatient: [], opd: [], off: [], leave: [] };
    const place = String(r.AssignmentPlace || ''), id = String(r.ConsultantID || '');
    if (place === 'Leave') { byWeek[w].leave.push(id); return; }
    if (place === 'Inactive' || isManualCoverId_(id)) return;
    if (place === 'OPD') byWeek[w].opd.push(id);
    else if (place === 'Off-duty') byWeek[w].off.push(id);
    else if (inpatientPlaces.indexOf(place) !== -1) byWeek[w].inpatient.push(id);
  });
  const uncoveredByWeek = {};
  (simulation.conflicts || []).forEach(c => {
    if (String(c.TeamID) === teamId && String(c.Level) === 'Red') (uncoveredByWeek[Number(c.WeekNo)] = uncoveredByWeek[Number(c.WeekNo)] || []).push(String(c.Message || ''));
  });
  const weeksOut = [];
  for (let w = fromWeek; w <= toWeek; w++) {
    const cur = byWeek[w]; if (!cur) continue;
    const prev = byWeek[w - 1];
    const backToBack = prev ? cur.inpatient.filter(id => prev.inpatient.indexOf(id) !== -1) : [];
    weeksOut.push({ weekNo: w, weekStart: cur.weekStart, weekEnd: cur.weekEnd,
      available: cur.inpatient.length + cur.opd.length + cur.off.length,
      onLeave: cur.leave.map(nameOf), inpatient: cur.inpatient.map(nameOf), backToBack: backToBack.map(nameOf),
      opdCount: cur.opd.length, offDuty: cur.off.map(nameOf), uncovered: uncoveredByWeek[w] || [] });
  }
  const span = list => {
    // Collapses consecutive week numbers into "Week 25" / "Weeks 25-26" with dates.
    const out = []; let i = 0;
    while (i < list.length) {
      let j = i; while (j + 1 < list.length && list[j + 1].weekNo === list[j].weekNo + 1) j++;
      out.push((i === j ? 'Week ' + list[i].weekNo : 'Weeks ' + list[i].weekNo + '-' + list[j].weekNo) +
        ' (' + dateOnlyNice_(list[i].weekStart) + ' - ' + dateOnlyNice_(list[j].weekEnd) + ')');
      i = j + 1;
    }
    return out.join(', ');
  };
  const lines = [];
  const who = nameOf(leave.ConsultantID);
  // Only the weeks this consultant is away; other people's leave is context, not a consequence.
  const leaveWeeks = weeksOut.filter(x => x.onLeave.indexOf(who) !== -1);
  const uncovered = weeksOut.filter(x => x.uncovered.length);
  if (uncovered.length) lines.push({ tone: 'bad', text: span(uncovered) + ': an inpatient post would be UNCOVERED. ' + uncovered[0].uncovered[0] });
  const two = leaveWeeks.filter(x => x.available <= 2 && !x.uncovered.length);
  if (two.length) lines.push({ tone: 'bad', text: span(two) + ': only ' + Math.min.apply(null, two.map(x => x.available)) +
    ' consultant(s) left. Everyone available works inpatient every week, with no OPD and no Off-duty reserve.' });
  const three = leaveWeeks.filter(x => x.available === 3);
  if (three.length) lines.push({ tone: 'warn', text: span(three) + ': 3 consultants available.' });
  const b2bWeeks = {};
  weeksOut.forEach(x => x.backToBack.forEach(n => (b2bWeeks[n] = b2bWeeks[n] || []).push(x.weekNo)));
  const b2b = Object.keys(b2bWeeks).map(n => n + ' (week' + (b2bWeeks[n].length > 1 ? 's ' : ' ') + b2bWeeks[n].join(', ') + ')');
  if (b2b.length) lines.push({ tone: 'warn', text: 'Back-to-back inpatient weeks: ' + b2b.join(', ') + '.' });
  else if (weeksOut.length) lines.push({ tone: 'ok', text: 'No one works back-to-back inpatient weeks.' });
  const noOpd = leaveWeeks.filter(x => x.opdCount === 0 && x.available > 2);
  if (noOpd.length) lines.push({ tone: 'info', text: span(noOpd) + ': no one is assigned to OPD' +
    (noOpd.some(x => x.offDuty.length) ? '; the Off-duty consultant is the reserve (' + [...new Set([].concat.apply([], noOpd.map(x => x.offDuty)))].join(', ') + ').' : '.') });
  const back = weeksOut.find(x => x.weekNo > leaveEndWeek && x.onLeave.indexOf(who) === -1);
  if (back) {
    const role = back.inpatient.indexOf(who) !== -1 ? 'inpatient' : back.offDuty.indexOf(who) !== -1 ? 'Off-duty' : 'OPD';
    lines.push({ tone: 'info', text: who + ' returns in week ' + back.weekNo + ' (' + dateOnlyNice_(back.weekStart) + ') and is scheduled for ' + role + '.' });
  }
  return { fromWeek: fromWeek, toWeek: toWeek, weeks: weeksOut, lines: lines };
}

function approveLeave(adminPin, leaveId, overrideReason) {
  setupDatabase();
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const leave = readObjects_(SHEETS.LeaveRequests).find(l => l.LeaveID === leaveId);
    if (!leave) throw new Error('Leave request was not found.');
    const session = requireLeaveAdmin_(adminPin, leave);
    if (normalizeStatus_(leave.Status) === 'approved') return { ok: true, message: 'This leave is already approved.' };
    const latest = getLatestVersion_();
    const approvedPatch = { Status: 'Approved', AdminDecisionDate: now_(), AdminNote: overrideReason || leave.AdminNote || '' };
    if (!latest) {
      updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, Object.assign({}, approvedPatch, { ImpactStatus: 'Future', ImpactSummary: 'No current rota exists yet. This leave will be applied when a rota is generated.' }));
      appendAudit_(session, 'LEAVE_APPROVED', 'LeaveRequest', leaveId, { consultantId: leave.ConsultantID, teamId: leave.TeamID, amendmentCreated: false, reason: 'No current rota' });
      return { ok: true, message: 'Leave approved. No current rota exists yet, so no amendment was created.' };
    }
    if (toDate_(leave.StartDate) > toDate_(latest.EndDate)) {
      updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, Object.assign({}, approvedPatch, { ImpactStatus: 'Future', ImpactSummary: 'Leave is outside the current rota period.' }));
      appendAudit_(session, 'LEAVE_APPROVED', 'LeaveRequest', leaveId, { consultantId: leave.ConsultantID, teamId: leave.TeamID, amendmentCreated: false, reason: 'Outside current period' });
      return { ok: true, message: 'Leave approved as future leave. It is outside the current rota period, so no amendment was created.' };
    }
    updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, approvedPatch);
    let result;
    try {
      const consultant = findConsultant_(leave.ConsultantID);
      result = createAmendmentFromDate_(leave.StartDate, 'Leave amendment', `Approved leave for ${consultant ? consultant.Name : leave.ConsultantID}`, overrideReason ? { requestType: 'Leave', requestId: leaveId, ruleBroken: 'Admin override', reason: overrideReason } : null, [leave.TeamID], leave.EndDate);
    } catch (err) {
      updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, { Status: leave.Status, AdminDecisionDate: leave.AdminDecisionDate || '', AdminNote: leave.AdminNote || '', ImpactStatus: leave.ImpactStatus || '', ImpactSummary: leave.ImpactSummary || '' });
      throw new Error('Leave approval was rolled back because the rota amendment failed: ' + ((err && err.message) || String(err)));
    }
    const teamConflicts = (result.conflicts || []).filter(c => c.TeamID === leave.TeamID), level = deriveImpactLevel_(teamConflicts, leave, latest), summary = summarizeImpact_(level, teamConflicts, leave, latest);
    updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, { ImpactStatus: level, ImpactSummary: summary });
    appendAudit_(session, 'LEAVE_APPROVED', 'LeaveRequest', leaveId, { consultantId: leave.ConsultantID, teamId: leave.TeamID, amendmentId: result.version && result.version.VersionID, amendmentNo: result.version && result.version.AmendmentNo, overrideReason: overrideReason || '' });
    return { ok: true, message: `Leave approved. ${result.message}`, version: result.version, conflicts: result.conflicts };
  } finally { lock.releaseLock(); }
}


/**
 * Main/team admins may replace an already approved leave with corrected dates.
 * The old record is kept as Superseded (audit trail), the replacement receives
 * a new LeaveID and a link to the original. Earlier rota weeks are never edited.
 * This does not auto-reconcile OTHER approved requests for the same consultant.
 */
function amendApprovedLeave(adminToken, originalLeaveId, newStartISO, newEndISO, adminNote) {
  setupDatabase();
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const oldLeave = readObjects_(SHEETS.LeaveRequests).find(l => String(l.LeaveID) === String(originalLeaveId));
    if (!oldLeave) throw new Error('Original leave request was not found.');
    const session = requireLeaveAdmin_(adminToken, oldLeave);
    if (normalizeStatus_(oldLeave.Status) !== 'approved')
      throw new Error('Only currently approved leave can be amended. Refresh leave data.');
    const startText = String(newStartISO || '').trim(), endText = String(newEndISO || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(startText) || !/^\d{4}-\d{2}-\d{2}$/.test(endText))
      throw new Error('Use YYYY-MM-DD for both dates.');
    const start = toDate_(startText), end = toDate_(endText);
    if (iso_(start) !== startText || iso_(end) !== endText || end < start)
      throw new Error('The amended leave dates are not valid.');
    if (startText === String(oldLeave.StartDate) && endText === String(oldLeave.EndDate))
      throw new Error('The amended dates are identical to the approved dates.');
    const note = String(adminNote || '').trim().slice(0, 500);
    if (!note) throw new Error('Enter a short amendment reason for the audit history.');

    const latest = getLatestVersion_();
    if (latest && getRotaVersions_().some(v => v.Status === 'Draft'))
      throw new Error('An unpublished draft exists. Publish or delete it before amending approved leave.');
    const newId = makeId_('L');
    const replacement = {
      LeaveID: newId, ConsultantID: oldLeave.ConsultantID, TeamID: oldLeave.TeamID,
      StartDate: startText, EndDate: endText, Status: 'Approved',
      SubmittedAt: now_(), UserNote: oldLeave.UserNote || '',
      AdminDecisionDate: now_(), AdminNote: note,
      ImpactStatus: '', ImpactSummary: '',
      SupersedesLeaveID: oldLeave.LeaveID, SupersededByLeaveID: ''
    };
    const prior = {
      Status: oldLeave.Status, AdminDecisionDate: oldLeave.AdminDecisionDate || '',
      AdminNote: oldLeave.AdminNote || '', ImpactStatus: oldLeave.ImpactStatus || '',
      ImpactSummary: oldLeave.ImpactSummary || '', SupersededByLeaveID: oldLeave.SupersededByLeaveID || ''
    };
    let inserted = false, changedOld = false, result = null;
    try {
      appendObject_(SHEETS.LeaveRequests, replacement); inserted = true;
      updateRow_(SHEETS.LeaveRequests, 'LeaveID', oldLeave.LeaveID, {
        Status: 'Superseded', SupersededByLeaveID: newId,
        ImpactStatus: 'Superseded', ImpactSummary: 'Replaced by approved leave ' + newId,
        AdminDecisionDate: now_(), AdminNote: note
      });
      changedOld = true;
      if (latest) {
        const from = toDate_(oldLeave.StartDate) < start ? oldLeave.StartDate : startText;
        const until = toDate_(oldLeave.EndDate) > end ? oldLeave.EndDate : endText;
        result = createAmendmentFromDate_(from, 'Approved leave dates amended',
          'Amended approved leave', null, [String(oldLeave.TeamID)], until);
      }
      if (!result) result = { message: 'New dates recorded for future rota generation.', amendmentCreated: false, conflicts: [] };
    } catch (err) {
      const rollbackErrors = [];
      if (changedOld) {
        try { updateRow_(SHEETS.LeaveRequests, 'LeaveID', oldLeave.LeaveID, prior); }
        catch (restoreErr) { rollbackErrors.push('original record'); }
      }
      if (inserted) {
        try { deleteRowsWhere_(SHEETS.LeaveRequests, r => String(r.LeaveID) === newId); }
        catch (deleteErr) { rollbackErrors.push('replacement record'); }
      }
      if (rollbackErrors.length) throw new Error('Leave amendment failed and rollback was incomplete (' + rollbackErrors.join(', ') + '). Inspect the staging database before retrying.');
      throw new Error('Leave amendment was not published; original leave retained: ' + ((err && err.message) || String(err)));
    }
    // A successful rota publication is the commit boundary. A later metadata
    // write failure must NOT restore superseded dates while leaving the new
    // published rota in place. Report that as a reconciliation warning.
    let metadataWarning = false;
    try {
      updateRow_(SHEETS.LeaveRequests, 'LeaveID', newId, {
        ImpactStatus: result.amendmentCreated ? 'Amended' : 'Future',
        ImpactSummary: result.message
      });
    } catch (metadataErr) { metadataWarning = true; }
    appendAudit_(session, 'APPROVED_LEAVE_AMENDED', 'LeaveRequest', newId, {
      supersedesLeaveId: oldLeave.LeaveID, consultantId: oldLeave.ConsultantID,
      oldStart: oldLeave.StartDate, oldEnd: oldLeave.EndDate,
      newStart: startText, newEnd: endText, publishedVersion: result.version && result.version.VersionID,
      amendmentCreated: !!result.amendmentCreated
    });
    return { ok: true, message: 'Leave dates amended. Previous request preserved as Superseded. ' + result.message +
      (metadataWarning ? ' Warning: impact summary could not be written; review staging leave details.' : ''),
      leaveId: newId, supersedesLeaveId: oldLeave.LeaveID,
      version: result.version || null, conflicts: result.conflicts || [] };
  } finally { lock.releaseLock(); }
}


function cancelApprovedLeave(adminPin, leaveId, note) {
  const session = requireAdminPin_(adminPin);
  setupDatabase();
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const leave = readObjects_(SHEETS.LeaveRequests).find(l => l.LeaveID === leaveId);
    if (!leave) throw new Error('Leave request was not found.');
    const cancelPatch = { Status: 'Cancelled', AdminDecisionDate: now_(), AdminNote: note || leave.AdminNote || 'Cancelled by admin' };
    if (normalizeStatus_(leave.Status) !== 'approved') {
      updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, cancelPatch);
      appendAudit_(session, 'LEAVE_CANCELLED', 'LeaveRequest', leaveId, { consultantId: leave.ConsultantID, amendmentCreated: false });
      return { ok: true, message: 'Leave request cancelled. No rota amendment was needed because it was not approved.' };
    }
    const latest = getLatestVersion_();
    updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, Object.assign({}, cancelPatch, { ImpactStatus: 'Cancelled', ImpactSummary: 'Approved leave was cancelled. Consultant is available again from the affected week onward.' }));
    if (!latest) {
      appendAudit_(session, 'LEAVE_CANCELLED', 'LeaveRequest', leaveId, { consultantId: leave.ConsultantID, amendmentCreated: false, reason: 'No current rota' });
      return { ok: true, message: 'Approved leave cancelled. No current rota exists, so no amendment was created.' };
    }
    const leaveEnd = toDate_(leave.EndDate), currentStart = toDate_(latest.StartDate), currentEnd = toDate_(latest.EndDate);
    if (leaveEnd < currentStart || toDate_(leave.StartDate) > currentEnd) {
      appendAudit_(session, 'LEAVE_CANCELLED', 'LeaveRequest', leaveId, { consultantId: leave.ConsultantID, amendmentCreated: false, reason: 'Outside current period' });
      return { ok: true, message: 'Approved leave cancelled. The leave is outside the current rota period, so no amendment was created.' };
    }
    let result;
    try {
      const consultant = findConsultant_(leave.ConsultantID), effective = toDate_(leave.StartDate) < currentStart ? latest.StartDate : leave.StartDate;
      result = createAmendmentFromDate_(effective, 'Leave cancellation', `Cancelled leave for ${consultant ? consultant.Name : leave.ConsultantID}`, { requestType: 'Leave', requestId: leaveId, ruleBroken: 'Leave cancelled', reason: note || 'Approved leave cancelled' }, [leave.TeamID], leave.EndDate);
    } catch (err) {
      updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, { Status: leave.Status, AdminDecisionDate: leave.AdminDecisionDate || '', AdminNote: leave.AdminNote || '', ImpactStatus: leave.ImpactStatus || '', ImpactSummary: leave.ImpactSummary || '' });
      throw new Error('Leave cancellation was rolled back because the rota amendment failed: ' + ((err && err.message) || String(err)));
    }
    appendAudit_(session, 'LEAVE_CANCELLED', 'LeaveRequest', leaveId, { consultantId: leave.ConsultantID, amendmentId: result.version && result.version.VersionID, amendmentNo: result.version && result.version.AmendmentNo });
    return { ok: true, message: `Leave cancelled and consultant made available again. ${result.message}`, version: result.version, conflicts: result.conflicts };
  } finally { lock.releaseLock(); }
}

function rejectLeave(adminPin, leaveId, note) {
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const leave = readObjects_(SHEETS.LeaveRequests).find(l => l.LeaveID === leaveId);
    if (!leave) throw new Error('Leave request was not found.');
    const session = requireLeaveAdmin_(adminPin, leave);
    updateRow_(SHEETS.LeaveRequests, 'LeaveID', leaveId, { Status: 'Rejected', AdminDecisionDate: now_(), AdminNote: note || '' });
    appendAudit_(session, 'LEAVE_REJECTED', 'LeaveRequest', leaveId, { consultantId: leave.ConsultantID, teamId: leave.TeamID, note: note || '' });
    return { ok: true, message: 'Leave request rejected.' };
  } finally { lock.releaseLock(); }
}

function previewInactivity(adminPin, data) {
  requireAdminPin_(adminPin);
  setupDatabase();
  const consultant = findConsultant_(data.ConsultantID);
  if (!consultant) throw new Error('Consultant not found.');
  const latest = getLatestVersion_();
  if (!latest) return { level: 'Warning', summary: 'No current rota exists. The consultant can be marked inactive, and future rota generation will exclude them.', conflicts: [], preview: [] };

  const inactiveDate = toDate_(data.InactiveFrom);
  const tempConsultants = getConsultants_().map(c => {
    if (c.ConsultantID === consultant.ConsultantID) {
      const copy = Object.assign({}, c);
      copy.CurrentStatus = 'Inactive';
      copy.InactiveFrom = iso_(inactiveDate);
      return copy;
    }
    return c;
  });

  const start = toDate_(latest.StartDate);
  const weeks = Number(latest.Weeks || 52);
  const currentEnd = toDate_(latest.EndDate);
  if (inactiveDate > currentEnd) {
    return { level: 'Info', summary: 'Inactive date is outside the current rota. No amendment is needed now.', conflicts: [], preview: [] };
  }

  const effectiveDate = inactiveDate < start ? start : inactiveDate;
  const simulation = simulateAmendment_(latest, effectiveDate, getApprovedLeaves_(), tempConsultants);
  const affectedWeekStart = getWeekNoForDate_(start, effectiveDate, weeks);
  const previewRows = simulation.rows.filter(r => Number(r.WeekNo) >= affectedWeekStart && Number(r.WeekNo) <= affectedWeekStart + 2 && r.TeamID === consultant.TeamID);
  const conflicts = simulation.conflicts.filter(c => c.TeamID === consultant.TeamID);
  const level = conflicts.some(c => c.Level === 'Red') ? 'Conflict' : conflicts.some(c => c.Level === 'Yellow') ? 'Warning' : 'Safe';
  const summary = level === 'Safe'
    ? `${consultant.Name} can be removed from future assignments from ${iso_(inactiveDate)} without major conflict.`
    : `${consultant.Name} can be marked inactive, but the rota will need adjustment from ${iso_(inactiveDate)}. ${conflicts.length} warning/conflict item(s) found.`;
  return { level: level, summary: summary, conflicts: conflicts.slice(0, 20), preview: groupRowsForTeamPreview_(previewRows) };
}

function confirmInactivity(adminPin, data) {
  const session = requireAdminPin_(adminPin);
  setupDatabase();
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const consultant = findConsultant_(data.ConsultantID);
    if (!consultant) throw new Error('Consultant not found.');
    const inactiveDate = iso_(toDate_(data.InactiveFrom));
    const patch = { CurrentStatus: 'Inactive', InactiveFrom: inactiveDate, Notes: data.Notes || consultant.Notes || '', UpdatedAt: now_() };
    updateRow_(SHEETS.Consultants, 'ConsultantID', consultant.ConsultantID, patch);
    const latest = getLatestVersion_();
    if (!latest || toDate_(inactiveDate) > toDate_(latest.EndDate)) {
      appendAudit_(session, 'CONSULTANT_INACTIVATED', 'Consultant', consultant.ConsultantID, { inactiveFrom: inactiveDate, amendmentCreated: false });
      return { ok: true, message: `${consultant.Name} marked inactive from ${inactiveDate}. No current rota amendment was needed.` };
    }
    let result;
    try { result = createStaffingDraftFromDateIfInsideRota_(inactiveDate, [consultant.TeamID], 'Consultant inactivity draft', `${consultant.Name} inactive from ${inactiveDate}`); }
    catch (err) {
      updateRow_(SHEETS.Consultants, 'ConsultantID', consultant.ConsultantID, { CurrentStatus: consultant.CurrentStatus, InactiveFrom: consultant.InactiveFrom || '', Notes: consultant.Notes || '', UpdatedAt: consultant.UpdatedAt || now_() });
      throw new Error('Inactivity change was rolled back because the draft could not be created: ' + ((err && err.message) || String(err)));
    }
    appendAudit_(session, 'CONSULTANT_INACTIVATED', 'Consultant', consultant.ConsultantID, { inactiveFrom: inactiveDate, amendmentId: result.version && result.version.VersionID, amendmentNo: result.version && result.version.AmendmentNo });
    return { ok: true, message: `${consultant.Name} marked inactive. ${result.message}`, version: result.version, conflicts: result.conflicts };
  } finally { lock.releaseLock(); }
}

function getRotaVersions(adminPin) {
  requireAdminPin_(adminPin);
  setupDatabase();
  return getRotaVersions_();
}

function voidRotaVersion(adminToken, versionId, reason) {
  const session = requireAdminPin_(adminToken);
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let version = null;
  let restored = null;
  try {
    const versions = getRotaVersions_();
    version = versions.find(v => v.VersionID === versionId);
    if (!version) throw new Error('Rota version was not found.');
    if (version.Status === 'Voided') return { ok: true, message: 'This rota version is already voided.', latestVersion: getLatestVersion_() };
    if (version.Status === 'Draft') throw new Error('Draft versions should be deleted with Delete Draft, not voided.');

    if (version.Status === 'Official') {
      const candidates = versions
        .filter(v => v.VersionID !== versionId && v.Status === 'Archived' && String(v.PeriodID || '') === String(version.PeriodID || ''))
        .sort((a, b) => Number(b.AmendmentNo || 0) - Number(a.AmendmentNo || 0));
      restored = candidates.length ? candidates[0] : null;
      // Secure the current official assignments in archive before changing status.
      archiveVersionAssignments_(version.VersionID);
      if (restored) restoreVersionAssignmentsToLive_(restored.VersionID);
    }

    updateRow_(SHEETS.RotaVersions, 'VersionID', versionId, {
      Status: 'Voided',
      Reason: `${version.Reason || ''}${version.Reason ? ' | ' : ''}VOIDED: ${reason || 'No reason provided'} (${now_()})`
    });
    if (restored) updateRow_(SHEETS.RotaVersions, 'VersionID', restored.VersionID, { Status: 'Official', AssignmentStorage: 'Live' });

    if (version.Status === 'Official') {
      updateSetting_('CurrentExportDocId', '');
      updateSetting_('CurrentExportDocUrl', '');
    }
    const remainingPeriodVersions = getRotaVersions_().filter(v => String(v.PeriodID || '') === String(version.PeriodID || '') && v.Status !== 'Voided');
    const period = getPeriodById_(version.PeriodID);
    if (period && period.Status !== 'Active' && !remainingPeriodVersions.length) {
      updateRow_(SHEETS.RotaPeriods, 'PeriodID', period.PeriodID, { Status: 'Voided', ArchivedAt: now_() });
    }
    appendAudit_(session, 'ROTA_VERSION_VOIDED', 'RotaVersion', versionId, { amendmentNo: version.AmendmentNo, periodId: version.PeriodID, reason: reason || '', restoredVersionId: restored ? restored.VersionID : '' });
    warmPublicStartupCache_();
    return {
      ok: true,
      message: restored
        ? `Rota version voided. Amendment ${restored.AmendmentNo} remains official within ${restored.PeriodName || 'the same rota period'}.`
        : 'Rota version voided. There is no official amendment in the active rota period until you restore another version or publish a draft.',
      latestVersion: getLatestVersion_()
    };
  } catch (err) {
    // Best-effort rollback: the previously official amendment must remain readable and live.
    if (restored) {
      try { archiveVersionAssignments_(restored.VersionID); } catch (_) {}
      try { updateRow_(SHEETS.RotaVersions, 'VersionID', restored.VersionID, { Status: 'Archived' }); } catch (_) {}
    }
    if (version && version.Status === 'Official') restoreOfficialAfterRollback_(version);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function restoreRotaVersion(adminToken, versionId, reason) {
  const session = requireAdminPin_(adminToken);
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  const versionSnapshots = [];
  const periodSnapshots = [];
  try {
    const version = getRotaVersions_().find(v => v.VersionID === versionId);
    if (!version) throw new Error('Rota version was not found.');
    if (version.Status === 'Voided') throw new Error('Voided versions cannot be restored directly. Generate a new draft instead.');
    if (version.Status === 'Draft') throw new Error('Draft versions cannot be restored directly. Publish the draft, or delete it and create a new draft.');
    if (version.Status !== 'Archived') throw new Error('Only archived rota versions can be restored.');
    const period = getPeriodById_(version.PeriodID);
    if (!period) throw new Error('The rota period for this version was not found. Run Safe Repair / Add Missing Columns.');

    getRotaVersions_().filter(v => v.Status === 'Official').forEach(v => versionSnapshots.push(v));
    getRotaPeriods_().forEach(p => periodSnapshots.push({ PeriodID: p.PeriodID, Status: p.Status, ActivatedAt: p.ActivatedAt, ArchivedAt: p.ArchivedAt }));
    const activePeriod = getActivePeriod_();
    if (!activePeriod || String(activePeriod.PeriodID) !== String(period.PeriodID)) {
      if (activePeriod) {
        archiveOfficialVersions_(activePeriod.PeriodID);
        archivePeriod_(activePeriod.PeriodID);
      }
      activatePeriod_(period.PeriodID);
    }
    archiveOfficialVersions_(period.PeriodID);
    restoreVersionAssignmentsToLive_(versionId);
    updateRow_(SHEETS.RotaVersions, 'VersionID', versionId, {
      Status: 'Official', AssignmentStorage: 'Live',
      Reason: `${version.Reason || ''}${version.Reason ? ' | ' : ''}RESTORED: ${reason || 'Admin restored this version'} (${now_()})`
    });
    updateSetting_('CurrentExportDocId', '');
    updateSetting_('CurrentExportDocUrl', '');
    appendAudit_(session, 'ROTA_VERSION_RESTORED', 'RotaVersion', versionId, { amendmentNo: version.AmendmentNo, periodId: version.PeriodID, reason: reason || '' });
    warmPublicStartupCache_();
    return { ok: true, message: `Amendment ${version.AmendmentNo} restored, and ${period.PeriodName} is now the active rota period.`, latestVersion: getLatestVersion_() };
  } catch (err) {
    // Remove a partially restored target from live storage, then restore the prior official state.
    try {
      const target = getRotaVersions_().find(v => v.VersionID === versionId);
      if (target && target.Status === 'Official') archiveVersionAssignments_(target.VersionID);
      if (target) updateRow_(SHEETS.RotaVersions, 'VersionID', target.VersionID, { Status: 'Archived' });
    } catch (_) {}
    periodSnapshots.forEach(p => { try { updateRow_(SHEETS.RotaPeriods, 'PeriodID', p.PeriodID, { Status: p.Status, ActivatedAt: p.ActivatedAt, ArchivedAt: p.ArchivedAt }); } catch (_) {} });
    versionSnapshots.forEach(v => restoreOfficialAfterRollback_(v));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function deleteDraftRota(adminPin, versionId) {
  requireAdminPin_(adminPin);
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const version = getRotaVersions_().find(v => v.VersionID === versionId);
    if (!version) throw new Error('Rota version was not found.');
    if (version.Status !== 'Draft') throw new Error('Only draft rota versions can be permanently deleted. Official/archived versions should be voided instead.');
    const periodId = String(version.PeriodID || '');
    deleteRowsWhere_(SHEETS.RotaAssignments, row => row.VersionID === versionId);
    deleteRowsWhere_(SHEETS.ConflictLog, row => row.VersionID === versionId);
    deleteRowsWhere_(SHEETS.OverrideLog, row => row.VersionID === versionId);
    deleteRowsWhere_(SHEETS.RotaVersions, row => row.VersionID === versionId);
    deleteDraftPeriodIfEmpty_(periodId);
    appendAudit_(authenticateAdmin_(adminPin), 'ROTA_DRAFT_DELETED', 'RotaVersion', versionId, { periodId: periodId, amendmentNo: version.AmendmentNo });
    return { ok: true, message: 'Draft rota deleted permanently. A separate draft period was also removed when it no longer contained any versions.' };
  } finally {
    lock.releaseLock();
  }
}

function resetRotaData(adminPin, confirmText) {
  requireAdminPin_(adminPin);
  // Production guard: deleting every rota period/version/assignment and trashing the
  // archive files must be explicitly enabled per deployment (e.g. staging only) via
  // Apps Script -> Project Settings -> Script Properties: ALLOW_ROTA_RESET = YES.
  if (String(PropertiesService.getScriptProperties().getProperty('ALLOW_ROTA_RESET') || '').trim().toUpperCase() !== 'YES') {
    throw new Error('Clear Generated Rota is disabled on this deployment to protect rota history. Nothing was changed.');
  }
  setupDatabase();
  if (String(confirmText || '').trim() !== 'RESET ROTA') throw new Error('Type RESET ROTA to confirm.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const before = {
      periods: readObjects_(SHEETS.RotaPeriods).length,
      versions: readObjects_(SHEETS.RotaVersions).length,
      assignments: readObjects_(SHEETS.RotaAssignments).length,
      conflicts: readObjects_(SHEETS.ConflictLog).length,
      overrides: readObjects_(SHEETS.OverrideLog).length
    };

    // Archive spreadsheets created by this app are moved to Drive trash during a confirmed full rota reset.
    const trashedArchives = trashAssignmentArchives_();
    // Use physical row deletion, not clearContent().
    // This avoids old blank rows, stale filters, and UI/export reads that can appear as if the old rota still exists.
    deleteAllDataRows_(SHEETS.RotaAssignments);
    deleteAllDataRows_(SHEETS.RotaVersions);
    deleteAllDataRows_(SHEETS.RotaPeriods);
    deleteAllDataRows_(SHEETS.ConflictLog);
    deleteAllDataRows_(SHEETS.OverrideLog);
    // The old export doc no longer matches any existing version - start fresh next export.
    updateSetting_('CurrentExportDocId', '');
    updateSetting_('CurrentExportDocUrl', '');

    const after = {
      periods: readObjects_(SHEETS.RotaPeriods).length,
      versions: readObjects_(SHEETS.RotaVersions).length,
      assignments: readObjects_(SHEETS.RotaAssignments).length,
      conflicts: readObjects_(SHEETS.ConflictLog).length,
      overrides: readObjects_(SHEETS.OverrideLog).length
    };

    appendAudit_(authenticateAdmin_(adminPin), 'ROTA_DATA_RESET', 'RotaSystem', 'ALL', before);
    warmPublicStartupCache_();
    return {
      ok: true,
      message: `Generated rota reset complete. Removed ${before.periods} rota period(s), ${before.versions} rota version(s), ${before.assignments} live assignment row(s), ${before.conflicts} conflict row(s), and ${before.overrides} override row(s). ${trashedArchives} assignment archive file(s) were moved to Drive trash. Teams, consultants, and leave requests were kept.`,
      before: before,
      after: after,
      state: getAdminData(adminPin)
    };
  } finally {
    lock.releaseLock();
  }
}


function appendManualNameRows_(rows, versionId, weekNo, weekStartIso, weekEndIso, teamId, place, text) {
  parseManualNames_(text).forEach(name => {
    const consultant = findConsultantByNameInTeam_(name, teamId);
    if (consultant) {
      rows.push(row_(versionId, weekNo, weekStartIso, weekEndIso, teamId, place, consultant.ConsultantID, 'Manual admin edit', 'Manually edited and published by admin'));
    } else {
      rows.push(row_(versionId, weekNo, weekStartIso, weekEndIso, teamId, place, manualCoverId_(name), 'Manual cover', 'External/manual cover; not used for future auto-generation'));
    }
  });
}

function parseManualNames_(text) {
  return String(text || '')
    .split(/[\n,;]+/)
    .map(x => x.trim())
    .filter(x => x && x !== '-');
}

function findConsultantByNameInTeam_(name, teamId) {
  const normalized = normalizeName_(name);
  return getConsultants_().find(c => c.TeamID === teamId && normalizeName_(c.Name) === normalized) || null;
}

function normalizeName_(name) {
  return String(name || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function manualCoverId_(name) {
  return 'COVER::' + String(name || '').trim().replace(/[\r\n]+/g, ' ');
}

function isManualCoverId_(id) {
  return String(id || '').indexOf('COVER::') === 0;
}

function displayNameFromAssignment_(r, consultants) {
  if (isManualCoverId_(r.ConsultantID)) return String(r.ConsultantID).replace(/^COVER::/, '');
  const consultant = consultants[r.ConsultantID];
  return consultant ? consultant.Name : r.ConsultantID;
}

function personFromAssignment_(r, consultants) {
  if (isManualCoverId_(r.ConsultantID)) {
    return { id: r.ConsultantID, name: displayNameFromAssignment_(r, consultants), color: '#64748b', manualCover: true };
  }
  const consultant = consultants[r.ConsultantID] || { Name: r.ConsultantID, ColorHex: '#64748b' };
  return { id: r.ConsultantID, name: consultant.Name, color: normalizeHexColor_(consultant.ColorHex || '#64748b'), manualCover: false };
}

function validateManualEditedWeek_(rows, weekNo, editMap) {
  const conflicts = [];
  const teams = teamMap_();
  Object.keys(editMap).forEach(teamId => {
    const ctx = { weekNo: weekNo, weekStart: toDate_(rows.find(r => Number(r.WeekNo) === weekNo && r.TeamID === teamId)?.WeekStart || new Date()), team: teams[teamId] || { TeamID: teamId, TeamName: teamId } };
    const teamRows = rows.filter(r => Number(r.WeekNo) === weekNo && r.TeamID === teamId);
    const necessary = subTeamNamesForView_(teams[teamId], ctx.weekStart, teamRows);
    if (!necessary.length) necessary.push('Team');
    necessary.forEach(role => {
    if (!teamRows.some(r => r.AssignmentPlace === role)) {
      conflicts.push(conflict_(ctx, 'Red', 'MANUAL_NO_TEAM_COVER', `Manual edit left ${role} without inpatient coverage.`));
    }
    });
    const seen = {};
    teamRows.filter(r => (necessary.includes(r.AssignmentPlace) || ['Team','TCU','EMCU','OPD','Off-duty','Reserve'].includes(r.AssignmentPlace)) && !isManualCoverId_(r.ConsultantID)).forEach(r => {
      if (seen[r.ConsultantID]) throw new Error('Manual edit blocked: '+displayNameFromAssignment_(r, consultantMap_())+' is allocated more than once in the same week.');
      seen[r.ConsultantID] = true;
    });
    teamRows.filter(r => ['Leave','Inactive'].includes(r.AssignmentPlace)).forEach(r => {
      if (seen[r.ConsultantID]) throw new Error('Manual edit blocked: a consultant on leave or inactive is assigned a working duty.');
    });
    // Manual edits may intentionally override fair rotation, but warn when
    // they give a consultant either inpatient sub-team two weeks in a row.
    if (isSubTeamActiveForWeek_(teams[teamId],ctx.weekStart)) {
      const clinicalRoles = subTeamNames_(teams[teamId]);
      const consultantRecords=consultantMap_();
      teamRows.filter(r => clinicalRoles.indexOf(r.AssignmentPlace) !== -1 && !isManualCoverId_(r.ConsultantID)).forEach(r => {
        const consultant=consultantRecords[r.ConsultantID];
        if (consultant && !consultantAllowedInSubTeam_(consultant,r.AssignmentPlace,clinicalRoles,ctx.weekStart))
          throw new Error('Manual edit blocked: '+consultant.Name+' is configured for a different sub-team. Update Consultant Settings before assigning this service.');
        const previous = rows.find(p => String(p.TeamID)===String(teamId) && p.ConsultantID===r.ConsultantID &&
          Number(p.WeekNo)===weekNo-1 && (p.AssignmentPlace==='Team' || clinicalRoles.indexOf(p.AssignmentPlace)!==-1));
        if (previous) conflicts.push(conflict_(ctx,'Yellow','MANUAL_CONSECUTIVE_INPATIENT',
          displayNameFromAssignment_(r,consultantMap_())+' assigned inpatient in consecutive Weeks '+(weekNo-1)+' and '+weekNo+'.'));
        const twoWeeksAgo = rows.find(p => String(p.TeamID)===String(teamId) && p.ConsultantID===r.ConsultantID &&
          Number(p.WeekNo)===weekNo-2 && (p.AssignmentPlace==='Team' || clinicalRoles.indexOf(p.AssignmentPlace)!==-1));
        if (twoWeeksAgo && !previous) conflicts.push(conflict_(ctx,'Yellow','MANUAL_RAPID_INPATIENT_RETURN',
          displayNameFromAssignment_(r,consultantMap_())+' returns to inpatient duty after only one intervening week. Consider another available consultant.'));
      });
    }
    // Manual amendments also protect the minimum three-week Off-duty spacing.
    // Past weeks are read-only; never rewrite them to resolve a bad edit.
    if (isSubTeamActiveForWeek_(teams[teamId],ctx.weekStart)) {
      const offRows = teamRows.filter(r => r.AssignmentPlace === 'Off-duty' && !isManualCoverId_(r.ConsultantID));
      if (offRows.length > 1) throw new Error('Manual edit blocked: this sub-team rotation permits at most one Off-duty consultant per week.');
      offRows.forEach(r => {
        const person = consultantMap_()[String(r.ConsultantID)];
        if (!person || String(person.TeamID)!==String(teamId))
          throw new Error('Manual edit blocked: Off-duty consultant is not registered in the selected parent team.');
        if (String(person.CanOffDuty)!=='Yes')
          throw new Error('Manual edit blocked: '+person.Name+' is not eligible for Off-duty in Consultant Settings.');
        if (!isActiveForWeek_(person,ctx.weekStart,addDays_(ctx.weekStart,6)) ||
            isInactiveForWeek_(person,ctx.weekStart,addDays_(ctx.weekStart,6)))
          throw new Error('Manual edit blocked: '+person.Name+' is not active for the selected week.');
        const precedingLeave=rows.some(p=>String(p.TeamID)===String(teamId) && String(p.ConsultantID)===String(r.ConsultantID) &&
          Number(p.WeekNo)===weekNo-1 && p.AssignmentPlace==='Leave');
        if (precedingLeave) throw new Error('Manual edit blocked: first week back from approved leave cannot be Off-duty. Assign OPD or an inpatient service.');
        const previous = rows.find(p => String(p.TeamID) === String(teamId) && p.ConsultantID === r.ConsultantID &&
          p.AssignmentPlace === 'Off-duty' && Number(p.WeekNo) < weekNo && Number(p.WeekNo) >= weekNo-2);
        if (previous) throw new Error('Manual edit blocked: '+displayNameFromAssignment_(r, consultantMap_())+
          ' was already Off-duty in Week '+previous.WeekNo+'. Keep at least three weeks between Off-duty assignments.');
        const priorInpatient=rows.find(p => String(p.TeamID)===String(teamId) && p.ConsultantID===r.ConsultantID &&
          Number(p.WeekNo)===weekNo-1 && (p.AssignmentPlace==='Team' || subTeamNames_(teams[teamId]).indexOf(p.AssignmentPlace)!==-1));
        if (priorInpatient) conflicts.push(conflict_(ctx,'Yellow','MANUAL_INPATIENT_TO_OFF',
          displayNameFromAssignment_(r,consultantMap_())+' moves directly from inpatient to Off-duty. This is allowed manually but discouraged by the automatic scheduler.'));
      });
    }
  });
  return conflicts;
}

function subTeamExportText_(team, index) {
  const name = (team.subTeamNames || [])[index];
  return name ? name + ': ' + peopleText_(team[name]) : '';
}

function exportLatestRota() {
  setupDatabase();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const latest = getLatestVersion_();
    if (!latest) throw new Error('No official rota exists yet. Generate a rota first.');

    const view = getLatestRotaView_();
    const assignmentCount = getVersionAssignments_(latest.VersionID).length;
    if (!assignmentCount || !view.weeks || !view.weeks.length) {
      throw new Error(`Latest Amendment ${latest.AmendmentNo} has no rota assignments, so the export was stopped to avoid creating an empty file. Restore the previous rota version or regenerate the rota, then export again.`);
    }
    const title = `CTU ROTA - ${latest.PeriodName || (latest.StartDate + ' to ' + latest.EndDate)} - Amendment ${latest.AmendmentNo}`;

    // One Google Doc is reused (content cleared and rewritten, then renamed) across every
    // amendment in the current rota cycle, instead of creating a new file each time.
    // A brand-new full rota generation (or a rota reset) clears CurrentExportDocId, so the
    // next export starts a fresh file for that new cycle.
    const existingDocId = String(getSetting_('CurrentExportDocId', '') || '').trim();
    if (existingDocId) {
      try {
        const doc = DocumentApp.openById(existingDocId);
        doc.getBody().clear();
        doc.setName(title);
        populateRotaExportDoc_(doc, latest, view);
        doc.saveAndClose();
        const docUrl = doc.getUrl();
        const docxUrl = `https://docs.google.com/document/d/${existingDocId}/export?format=docx`;
        updateRow_(SHEETS.RotaVersions, 'VersionID', latest.VersionID, { ExportDocId: existingDocId, ExportDocUrl: docUrl });
        appendAudit_({ name: 'App user', role: 'User', teamId: '' }, 'ROTA_EXPORTED', 'RotaVersion', latest.VersionID, { reused: true, amendmentNo: latest.AmendmentNo });
        return { ok: true, docUrl: docUrl, docxUrl: docxUrl, title: title, amendmentNo: latest.AmendmentNo, versionId: latest.VersionID, reused: true, message: 'Export file updated in place for this rota.' };
      } catch (err) {
        // Tracked doc was deleted/trashed outside the app - fall through and create a new one.
      }
    }

    const doc = DocumentApp.create(title);
    populateRotaExportDoc_(doc, latest, view);
    doc.saveAndClose();
    const docId = doc.getId();
    const docUrl = doc.getUrl();
    const docxUrl = `https://docs.google.com/document/d/${docId}/export?format=docx`;
    updateSetting_('CurrentExportDocId', docId);
    updateSetting_('CurrentExportDocUrl', docUrl);
    updateRow_(SHEETS.RotaVersions, 'VersionID', latest.VersionID, { ExportDocId: docId, ExportDocUrl: docUrl });
    appendAudit_({ name: 'App user', role: 'User', teamId: '' }, 'ROTA_EXPORTED', 'RotaVersion', latest.VersionID, { reused: false, amendmentNo: latest.AmendmentNo });
    return { ok: true, docUrl: docUrl, docxUrl: docxUrl, title: title, amendmentNo: latest.AmendmentNo, versionId: latest.VersionID, reused: false, message: 'New export file created for this rota.' };
  } finally {
    lock.releaseLock();
  }
}

function populateRotaExportDoc_(doc, latest, view) {
  const body = doc.getBody();
  body.appendParagraph('Consultant Weekly Rota').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Rota period: ${latest.PeriodName || '-'}`);
  body.appendParagraph(`Version: Amendment No. ${latest.AmendmentNo}`);
  body.appendParagraph(`Period dates: ${latest.StartDate} to ${latest.EndDate}`);
  body.appendParagraph(`Generated export: ${now_()}`);
  body.appendParagraph(`Reason: ${latest.Reason || '-'}`);
  body.appendParagraph('');

  const hasLegacy = (view.weeks || []).some(w => (w.teams || []).some(t => (t.Reserve || []).length));
  const tableData = [['Week No.', 'Week Period', 'Team', 'Inpatient Team', 'Sub-team 1', 'Sub-team 2', 'OPD', 'Off-duty', 'Leave', 'Inactive']];
  if (hasLegacy) tableData[0].push('Previously published legacy duty');
  const rowTeams = [];
  view.weeks.forEach(week => {
    week.teams.forEach(t => {
      const dataRow = [
        String(week.weekNo),
        week.periodExport || datesOnlyPeriodFromNicePeriod_(week.period),
        t.teamName,
        peopleText_(t.Team),
        subTeamExportText_(t, 0),
        subTeamExportText_(t, 1),
        peopleText_(t.OPD),
        peopleText_(t['Off-duty']),
        peopleText_(t.Leave),
        peopleText_(t.Inactive)
      ];
      if (hasLegacy) dataRow.push(peopleText_(t.Reserve));
      tableData.push(dataRow);
      rowTeams.push(t);
    });
  });

  const table = body.appendTable(tableData);
  styleRotaExportFlatTable_(table, rowTeams);

  body.appendParagraph('');
  body.appendParagraph('Note: Warnings and conflicts are intentionally not printed in this export. Admin can review them inside the web app.').setItalic(true);
}

function peopleText_(people) {
  if (!people || !people.length) return '-';
  // Use one name per line in exported cells. This keeps long/multiple names readable in Word.
  return people.map(p => typeof p === 'string' ? p : p.name).join('\n');
}

function styleRotaExportFlatTable_(table, rowTeams) {
  const header = table.getRow(0);
  for (let c = 0; c < header.getNumCells(); c++) {
    const cell = header.getCell(c);
    cell.setBackgroundColor('#0f172a').editAsText().setBold(true).setForegroundColor('#ffffff').setFontSize(9);
    styleExportCell_(cell, true);
  }

  rowTeams.forEach((team, index) => {
    const row = table.getRow(index + 1);
    const teamColor = normalizeHexColor_(team.teamColor || '#0f172a');
    for (let c = 0; c < row.getNumCells(); c++) styleExportCell_(row.getCell(c), false);
    row.getCell(0).editAsText().setBold(true);
    row.getCell(2).setBackgroundColor(teamColor).editAsText().setBold(true).setForegroundColor(contrastTextColor_(teamColor)).setFontSize(9);
    stylePeopleCell_(row.getCell(3), team.Team);
    stylePeopleCell_(row.getCell(4), team[(team.subTeamNames||[])[0]] || []);
    stylePeopleCell_(row.getCell(5), team[(team.subTeamNames||[])[1]] || []);
    stylePeopleCell_(row.getCell(6), team.OPD);
    stylePeopleCell_(row.getCell(7), team['Off-duty']);
    stylePeopleCell_(row.getCell(8), team.Leave);
    stylePeopleCell_(row.getCell(9), team.Inactive);
    if (row.getNumCells() > 10) stylePeopleCell_(row.getCell(10), team.Reserve);
  });
}

function styleExportCell_(cell, isHeader) {
  try { cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER); } catch (err) {}
  try { cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(4).setPaddingRight(4); } catch (err) {}
  for (let i = 0; i < cell.getNumChildren(); i++) {
    const child = cell.getChild(i);
    if (child.getType && child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      child.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      child.asParagraph().editAsText().setFontSize(isHeader ? 9 : 8);
    }
  }
}

function stylePeopleCell_(cell, people) {
  if (!people || !people.length) return;
  if (people.length === 1 && typeof people[0] !== 'string') {
    const color = normalizeHexColor_(people[0].color || '#64748b');
    cell.setBackgroundColor(color).editAsText().setForegroundColor(contrastTextColor_(color)).setFontSize(8);
    return;
  }
  const text = cell.editAsText();
  let pos = 0;
  people.forEach((person, idx) => {
    const name = typeof person === 'string' ? person : person.name;
    const color = typeof person === 'string' ? '#334155' : normalizeHexColor_(person.color || '#334155');
    const start = pos;
    const end = pos + String(name).length - 1;
    if (end >= start) {
      text.setForegroundColor(start, end, color);
      text.setBold(start, end, true);
      text.setFontSize(start, end, 8);
    }
    pos += String(name).length + (idx < people.length - 1 ? 1 : 0);
  });
}

function getLatestRotaView() {
  return getLatestRotaView_();
}

function getLatestRotaWindow() {
  const startup = getStartupPayload();
  return startup.latestRota || { meta: null, weeks: [], viewMode: 'window' };
}


// Staffing changes are reviewable drafts, never auto-published. This differs
// deliberately from the approved-leave auto-amendment workflow.
function createStaffingDraftFromDateIfInsideRota_(effectiveDateISO, teamIds, type, reason) {
  const latest=getLatestVersion_();
  if (!latest) return {message:'No active rota exists; the consultant setting will be used in the next generated draft.'};
  const currentSunday=startOfWeekSunday_(new Date());
  const earliestFutureSunday=addDays_(currentSunday,7);
  const requestedSunday=startOfWeekSunday_(toDate_(effectiveDateISO));
  // Do not silently reallocate weeks already started, even if ActiveFrom is historical.
  const effective=requestedSunday<earliestFutureSunday?earliestFutureSunday:requestedSunday;
  if (effective>toDate_(latest.EndDate)) return {message:'Consultant saved. No current-period amendment: the change takes effect in a later period.'};
  return createStaffingDraftFromDate_(iso_(effective),teamIds,type,reason);
}

function staffingDraftTeamIds_(version) {
  if(!version || String(version.Type)!=='Staffing batch draft') return [];
  const m=String(version.Reason||'').match(/\[STAFFING_TEAMS:([^\]]+)\]/);
  return m?m[1].split(',').map(x=>x.trim()).filter(Boolean):[];
}
// Two or more consultants may be added or transferred before publishing. A
// controlled staffing batch is rebuilt from the SAME official baseline; manual
// and unrelated drafts are not silently overwritten.
function createStaffingDraftFromDate_(effectiveDateISO,teamIds,type,reason) {
  const latest=getLatestVersion_();
  if(!latest)throw new Error('No current official rota exists.');
  const oldDraft=getRotaVersions_().find(v=>v.Status==='Draft');
  if(oldDraft && (String(oldDraft.Type)!=='Staffing batch draft' ||
      String(oldDraft.BasedOnVersionID)!==String(latest.VersionID)))
    throw new Error('An unrelated or outdated draft exists. Publish or delete it before changing staffing.');
  const previousTeams=oldDraft?staffingDraftTeamIds_(oldDraft):[];
  if(oldDraft && !previousTeams.length)
    throw new Error('Cannot merge staffing changes: existing draft lacks the approved staffing-team marker.');
  const teams=[...new Set(previousTeams.concat((teamIds||[]).map(String)))].filter(Boolean);
  if(!teams.length)throw new Error('No affected team was supplied for the staffing draft.');
  const start=toDate_(latest.StartDate),weeks=Number(latest.Weeks||52);
  const date=oldDraft && toDate_(oldDraft.EffectiveFrom)<toDate_(effectiveDateISO)?oldDraft.EffectiveFrom:effectiveDateISO;
  const effectiveWeek=Math.max(1,getWeekNoForDate_(start,toDate_(date),weeks));
  const oldRows=getVersionAssignments_(latest.VersionID);
  const version=createRotaVersion_({PeriodID:latest.PeriodID,
    AmendmentNo:nextAmendmentNoForPeriod_(latest.PeriodID),
    Status:'Draft',Type:'Staffing batch draft',StartDate:latest.StartDate,EndDate:latest.EndDate,Weeks:weeks,
    EffectiveFrom:iso_(addDays_(start,(effectiveWeek-1)*7)),
    Reason:'[STAFFING_TEAMS:'+teams.join(',')+'] '+reason+(oldDraft?' | consolidated staffing additions':''),
    BasedOnVersionID:latest.VersionID});
  try{
    const fixed=oldRows.filter(r=>!(teams.includes(String(r.TeamID)) && Number(r.WeekNo)>=effectiveWeek))
      .map(r=>Object.assign({},r,{VersionID:version.VersionID,PeriodID:latest.PeriodID}));
    const result=buildRotaRows_({versionId:version.VersionID,startDate:start,weeks:weeks,fromWeekNo:effectiveWeek,
      previousRows:fixed,manualSeedWeeks:0,manualAssignments:[],leavesOverride:getApprovedLeaves_(),teamIds:teams});
    assertRowsBeforeEffectiveWeekUnchanged_(oldRows,result.rows,effectiveWeek);
    assertUnaffectedTeamsUnchanged_(oldRows,result.rows,teams);
    if(!result.rows.length)throw new Error('Staffing draft contains no assignments.');
    appendRows_(SHEETS.RotaAssignments,result.rows);
    appendConflicts_(version.VersionID,result.conflicts);
    // Keep the earlier draft as an immutable superseded record. This avoids a
    // destructive delete if a Sheets call fails midway through a batch update.
    if(oldDraft)updateRow_(SHEETS.RotaVersions,'VersionID',oldDraft.VersionID,{Status:'Superseded'});
    return {message:'Staffing draft '+version.AmendmentNo+' ready for '+teams.map(getTeamName_).join(' and ')+
      ' from Week '+effectiveWeek+'. Added/updated consultants are included. Earlier weeks and other teams are protected. Review before publishing.',
      version:version,conflicts:result.conflicts};
  }catch(err){
    try{deleteRowsWhere_(SHEETS.RotaAssignments,r=>r.VersionID===version.VersionID);}catch(_){}
    try{deleteRowsWhere_(SHEETS.ConflictLog,r=>r.VersionID===version.VersionID);}catch(_){}
    try{deleteRowsWhere_(SHEETS.RotaVersions,r=>r.VersionID===version.VersionID);}catch(_){}
    throw err;
  }
}

// Auto-published leave amendments are strictly scoped to the affected team and
// to weeks that have not yet started. Historic leave is recorded without
// backdating published work assignments.
function createAmendmentFromDate_(effectiveDateISO, type, reason, overrideLog, affectedTeamIds, affectedEndISO) {
  const latest=getLatestVersion_();
  if(!latest)throw new Error('No current official rota exists.');
  const start=toDate_(latest.StartDate),weeks=Number(latest.Weeks||52),end=toDate_(latest.EndDate);
  const safeStart=addDays_(startOfWeekSunday_(new Date()),7);
  const requested=startOfWeekSunday_(toDate_(effectiveDateISO));
  const effective=safeStart>requested?safeStart:requested;
  const teams=[...new Set((affectedTeamIds||[]).map(String))].filter(Boolean);
  if(!teams.length)throw new Error('A leave amendment must specify its affected parent team.');
  if(effective>end || (affectedEndISO && toDate_(affectedEndISO)<effective))
    return {message:'Leave recorded; no future rota week is affected. Historical assignments were preserved.',version:latest,conflicts:[],amendmentCreated:false};
  const effectiveWeek=Math.max(1,getWeekNoForDate_(start,effective,weeks));
  const original=getVersionAssignments_(latest.VersionID);
  const draft=getRotaVersions_().find(v=>v.Status==='Draft');
  if(draft)throw new Error('An unpublished draft already exists; publish or delete it before automatic leave amendment. Leave status will be rolled back.');
  const version=createRotaVersion_({PeriodID:latest.PeriodID,AmendmentNo:nextAmendmentNoForPeriod_(latest.PeriodID),
    Status:'Draft',Type:type,StartDate:latest.StartDate,EndDate:latest.EndDate,Weeks:weeks,
    EffectiveFrom:iso_(addDays_(start,(effectiveWeek-1)*7)),Reason:reason,BasedOnVersionID:latest.VersionID});
  let madeOfficial=false;
  try{
    // Preserve every unrelated team, in addition to all weeks before cutoff.
    const fixed=copyRowsForVersion_(original.filter(r=>teams.indexOf(String(r.TeamID))===-1 || Number(r.WeekNo)<effectiveWeek),version.VersionID);
    const result=buildRotaRows_({versionId:version.VersionID,startDate:start,weeks:weeks,fromWeekNo:effectiveWeek,
      previousRows:fixed,manualSeedWeeks:0,manualAssignments:[],leavesOverride:getApprovedLeaves_(),teamIds:teams});
    assertRowsBeforeEffectiveWeekUnchanged_(original,result.rows,effectiveWeek);
    assertUnaffectedTeamsUnchanged_(original,result.rows,teams);
    if(!result.rows.length)throw new Error('Safety stop: amendment has no assignments.');
    appendRows_(SHEETS.RotaAssignments,result.rows);
    appendConflicts_(version.VersionID,result.conflicts);
    if(overrideLog)appendOverride_(Object.assign({},overrideLog,{VersionID:version.VersionID}));
    updateRow_(SHEETS.RotaVersions,'VersionID',version.VersionID,{Status:'Official'});
    madeOfficial=true;
    archiveVersion_(latest.VersionID);
    version.Status='Official';warmPublicStartupCache_();
    return {message:'Published Amendment '+version.AmendmentNo+' for '+teams.map(getTeamName_).join(', ')+
      ' from Week '+effectiveWeek+'. Earlier weeks and other teams were preserved.',version:version,conflicts:result.conflicts,amendmentCreated:true};
  }catch(err){
    if(madeOfficial){restoreOfficialAfterRollback_(latest);try{updateRow_(SHEETS.RotaVersions,'VersionID',version.VersionID,{Status:'Draft'});}catch(_){}}
    for(const sheet of [SHEETS.RotaAssignments,SHEETS.ConflictLog,SHEETS.OverrideLog,SHEETS.RotaVersions])
      try{deleteRowsWhere_(sheet,r=>r.VersionID===version.VersionID);}catch(_){}
    throw err;
  }
}

function copyRowsForVersion_(rows, versionId) {
  const version = readObjects_(SHEETS.RotaVersions).find(v => String(v.VersionID) === String(versionId));
  const periodId = version ? String(version.PeriodID || '') : '';
  return (rows || []).map(r => Object.assign({}, r, { VersionID: versionId, PeriodID: periodId }));
}

function assignmentSignatureAll_(rows) {
  return (rows || []).map(r => [r.WeekNo, r.WeekStart, r.WeekEnd, r.TeamID,
    r.AssignmentPlace, r.ConsultantID, r.Source, r.Notes].map(x => String(x == null ? '' : x)).join('¦')).sort().join('\n');
}

function assignmentSignatureBeforeWeek_(rows, effectiveWeek) {
  return (rows || [])
    .filter(r => Number(r.WeekNo) < effectiveWeek)
    .map(r => [
      Number(r.WeekNo),
      String(r.WeekStart || ''),
      String(r.WeekEnd || ''),
      String(r.TeamID || ''),
      String(r.AssignmentPlace || ''),
      String(r.ConsultantID || ''),
      String(r.Source || ''),
      String(r.Notes || '')
    ].join('¦'))
    .sort()
    .join('\n');
}

function assertUnaffectedTeamsUnchanged_(oldRows,newRows,affectedTeamIds) {
  const changed={};(affectedTeamIds||[]).forEach(id=>changed[String(id)]=true);
  const filtered=rows=>(rows||[]).filter(r=>!changed[String(r.TeamID)]);
  if (assignmentSignatureAll_(filtered(oldRows)) !== assignmentSignatureAll_(filtered(newRows)))
    throw new Error('Safety stop: amendment changed a team unrelated to the request. Nothing was published.');
}

function assertRowsBeforeEffectiveWeekUnchanged_(oldRows, newRows, effectiveWeek) {
  const oldSig = assignmentSignatureBeforeWeek_(oldRows, effectiveWeek);
  const newSig = assignmentSignatureBeforeWeek_(newRows, effectiveWeek);
  if (oldSig !== newSig) {
    throw new Error('Safety stop: the amendment tried to change weeks before the affected leave week. No new rota was published.');
  }
}

function simulateAmendment_(latest, effectiveDate, leavesOverride, consultantsOverride, affectedTeamIds) {
  const start = toDate_(latest.StartDate);
  const weeks = Number(latest.Weeks || 52);
  const effectiveWeek = Math.max(1, getWeekNoForDate_(start, effectiveDate, weeks));
  const previousRows = getVersionAssignments_(latest.VersionID).filter(r => !affectedTeamIds || !affectedTeamIds.includes(String(r.TeamID)) || Number(r.WeekNo) < effectiveWeek);
  return buildRotaRows_({
    versionId: 'SIMULATION',
    startDate: start,
    weeks: weeks,
    fromWeekNo: effectiveWeek,
    previousRows: previousRows,
    manualSeedWeeks: 0,
    manualAssignments: [],
    leavesOverride: leavesOverride,
    consultantsOverride: consultantsOverride || null,
    teamIds: affectedTeamIds || null
  });
}

function buildRotaRows_(opts) {
  const allowedTeamIds = opts.teamIds && opts.teamIds.length ? opts.teamIds.map(String) : null;
  const teams = getTeams_().filter(t => t.Active === 'Yes' && (!allowedTeamIds || allowedTeamIds.indexOf(String(t.TeamID)) !== -1));
  const consultants = opts.consultantsOverride || getConsultants_();
  const leaves = opts.leavesOverride || getApprovedLeaves_();
  const manualMap = manualAssignmentMap_(opts.manualAssignments || []);
  const rows = (opts.previousRows || []).map(r => Object.assign({}, r, { VersionID: opts.versionId }));
  const conflicts = [];
  const state = seedStateFromRows_(rows, consultants);
  state.offCooldownWeeks = opts.offCooldownWeeks || {};
  state.inpatientCooldownWeeks = opts.inpatientCooldownWeeks || {};
  state.inpatientCarryOverGapWeeks = opts.inpatientCarryOverGapWeeks || {};
  state.roleCarryOver = opts.roleCarryOver || {};
  // Fetched once per rebuild instead of per team/week - this setting rarely changes mid-rebuild.
  const maxOpd = Math.max(1, Number(getSetting_('OPDMaxPerTeam', '9') || 9));

  for (let weekNo = Number(opts.fromWeekNo || 1); weekNo <= Number(opts.weeks); weekNo++) {
    const weekStart = addDays_(opts.startDate, (weekNo - 1) * 7);
    const weekEnd = addDays_(weekStart, 6);
    teams.forEach(team => {
      const teamConsultants=consultants.filter(c=>String(c.TeamID)===String(team.TeamID));
      const buildFn = isSubTeamActiveForWeek_(team, weekStart) ? buildSubTeamWeekRows_ : buildTeamWeekRows_;
      const result = buildFn({
        versionId: opts.versionId,
        team: team,
        consultants: teamConsultants,
        leaves: leaves,
        weekNo: weekNo,
        weekStart: weekStart,
        weekEnd: weekEnd,
        rotaStartDate: opts.startDate,
        totalWeeks: Number(opts.weeks),
        state: state,
        maxOpd: maxOpd,
        manual: weekNo <= Number(opts.manualSeedWeeks || 0) ? manualMap[`${weekNo}|${team.TeamID}`] : null
      });
      rows.push.apply(rows, result.rows);
      conflicts.push.apply(conflicts, result.conflicts);
      updateStateFromWeekRows_(state, result.rows);
    });
  }
  return { rows: rows, conflicts: conflicts };
}

function subTeamNames_(team) {
  return String((team && team.SubTeamConfig) || 'TCU,EMCU').split(',').map(x => x.trim()).filter(Boolean).slice(0,2);
}
// Stored by *position*, not by label, so renaming TCU/EMCU does not
// silently switch a consultant to the wrong service.
function normalizeSubTeamAssignment_(value) {
  const raw=String(value==null?'':value).trim();
  if (!raw || raw==='Any') return 'Any';
  if (raw==='SubTeam1' || raw==='SubTeam2') return raw;
  throw new Error('Choose Any, Sub-team 1 only, or Sub-team 2 only.');
}
function consultantAllowedInSubTeam_(consultant, role, roles, weekStart) {
  const effective=String((consultant && consultant.SubTeamAssignmentFrom)||'').slice(0,10);
  if (effective && weekStart && iso_(weekStart)<effective) return true; // Historical weeks retain original flexibility.
  const setting=normalizeSubTeamAssignment_(consultant && consultant.SubTeamAssignment);
  if (setting==='Any') return true;
  const wanted=setting==='SubTeam1'?roles[0]:roles[1];
  return role===wanted;
}
function isSubTeamConfigured_(team) { return !!team && team.TeamType === 'SubTeam'; }
function isSubTeamActiveForWeek_(team, date) {
  if (!isSubTeamConfigured_(team)) return false;
  const week = iso_(date);
  const activated = !team.SubTeamFrom || week >= String(team.SubTeamFrom).slice(0,10);
  const stop = String(team.SubTeamUntil || '').slice(0,10);
  const resume = String(team.SubTeamResumeFrom || '').slice(0,10);
  // Sub-team weeks: [initial activation, stop), then [resume, onwards).
  // The interval between stop and resume remains a one-inpatient parent team.
  return activated && (!stop || week < stop || (!!resume && week >= resume));
}
function subTeamNamesForView_(team, date, teamRows) {
  const registered = isSubTeamConfigured_(team) ? subTeamNames_(team) : [];
  // Show configured service headings from their activation week. If the published
  // version still contains only legacy Team assignments the UI labels the new
  // headings PENDING, retains the actual Team coverage, and never invents staff.
  if (isSubTeamActiveForWeek_(team, date)) return registered;
  // Older published sub-team assignments still use their own historical labels.
  if ((teamRows || []).some(r => registered.indexOf(String(r.AssignmentPlace)) !== -1)) return registered;
  // Historical legacy names are preserved for teams configured before custom labels.
  if ((teamRows || []).some(r => r.AssignmentPlace === 'TCU' || r.AssignmentPlace === 'EMCU')) return ['TCU','EMCU'];
  return [];
}
function isSubTeamRole_(role, teamId) {
  // Only registered sub-team names count as inpatient work, including archived weeks.
  const team = getTeams_().find(t => String(t.TeamID) === String(teamId));
  return isSubTeamConfigured_(team) && subTeamNames_(team).indexOf(String(role)) !== -1;
}
// v2.6: Treat both inpatient sub-teams as ONE consecutive-duty category.
// The previous greedy picker could protect an OPD consultant for Off-duty and
// then allocate a previous-week inpatient consultant despite a workable alternative.
function wasInpatientPreviousWeek_(state, consultantId, weekNo) {
  const id = String(consultantId);
  const previous = Number(state && state.lastTeamWeek && state.lastTeamWeek[id] || 0);
  if (previous > 0 && previous === Number(weekNo) - 1) return true;
  return Number(weekNo) <= Number(state && state.inpatientCooldownWeeks && state.inpatientCooldownWeeks[id] || 0);
}

function inpatientCarryOverWeeks_(priorRows, newPeriodStart) {
  const result = {};
  const start = startOfWeekSunday_(toDate_(newPeriodStart));
  (priorRows || []).forEach(r => {
    if (isManualCoverId_(r.ConsultantID)) return;
    const role = String(r.AssignmentPlace || '');
    if (role !== 'Team' && !isSubTeamRole_(role, r.TeamID)) return;
    const prior = startOfWeekSunday_(toDate_(r.WeekStart));
    const elapsed = Math.round((start.getTime() - prior.getTime()) / (7*24*60*60*1000));
    if (elapsed === 1) result[String(r.ConsultantID)] = 1;
  });
  return result;
}

// Preserve TWO weeks of inpatient recency across a new rota period.
// The one-week cooldown above remains intact for backward compatibility.
function inpatientCarryOverGapWeeks_(priorRows, newPeriodStart) {
  const result={}, start=startOfWeekSunday_(toDate_(newPeriodStart));
  (priorRows||[]).forEach(r=>{
    if (isManualCoverId_(r.ConsultantID)) return;
    const role=String(r.AssignmentPlace||'');
    if (role!=='Team' && !isSubTeamRole_(role,r.TeamID)) return;
    const prior=startOfWeekSunday_(toDate_(r.WeekStart));
    const elapsed=Math.round((start.getTime()-prior.getTime())/(7*24*60*60*1000));
    const id=String(r.ConsultantID);
    if (elapsed>=1 && elapsed<=2 && (!result[id] || elapsed<result[id])) result[id]=elapsed;
  });
  return result;
}

function previousPeriodRoleCarryOver_(priorRows, newPeriodStart) {
  const result={}, start=startOfWeekSunday_(toDate_(newPeriodStart));
  (priorRows||[]).forEach(r=>{
    if (isManualCoverId_(r.ConsultantID)) return;
    const prior=startOfWeekSunday_(toDate_(r.WeekStart));
    const elapsed=Math.round((start.getTime()-prior.getTime())/(7*24*60*60*1000));
    if (elapsed!==1) return;
    const role=String(r.AssignmentPlace||'');
    if (role==='Off-duty' || role==='OPD' || role==='Team' || isSubTeamRole_(role,r.TeamID)) {
      result[String(r.ConsultantID)] = (role==='Team' || isSubTeamRole_(role,r.TeamID)) ? 'Team' : role;
    }
  });
  return result;
}

function inpatientRecencyGap_(state, consultantId, weekNo) {
  const id=String(consultantId), recent=Number((state.lastTeamWeek||{})[id]||0);
  if (recent>0) return Number(weekNo)-recent;
  const gap=Number((state.inpatientCarryOverGapWeeks||{})[id]||0);
  return gap ? Number(weekNo)+gap-1 : Infinity;
}

function returnedFromLeaveThisWeek_(ctx, consultant) {
  const id=String(consultant.ConsultantID), week=Number(ctx.weekNo);
  if (week>1 && Number((ctx.state.lastLeaveWeek || {})[id] || 0)===week-1) return true;
  if (week<=1) {
    // First week in a new period: inspect the last calendar week in approved leave.
    const lastStart=addDays_(ctx.weekStart,-7),lastEnd=addDays_(ctx.weekStart,-1);
    return isOnLeave_(id,lastStart,lastEnd,ctx.leaves) && !isOnLeave_(id,ctx.weekStart,ctx.weekEnd,ctx.leaves);
  }
  return false;
}

function buildSubTeamWeekRows_(ctx) {
  const roles = subTeamNames_(ctx.team);
  if (roles.length !== 2) throw new Error('Sub-team configuration requires exactly two services.');
  const rows = [], conflicts = [], teamId = ctx.team.TeamID, week = Number(ctx.weekNo);
  const start = iso_(ctx.weekStart), end = iso_(ctx.weekEnd);
  const available = [], leave = [], inactive = [];
  ctx.consultants.forEach(c => {
    // A future hire must not appear as Leave just because leave was entered
    // before their joining date. Effective staffing comes first.
    if (isInactiveForWeek_(c,ctx.weekStart,ctx.weekEnd)) inactive.push(c);
    else if (!isActiveForWeek_(c,ctx.weekStart,ctx.weekEnd)) return;
    else if (isOnLeave_(c.ConsultantID,ctx.weekStart,ctx.weekEnd,ctx.leaves)) leave.push(c);
    else available.push(c);
  });
  // A roster change starts a new fairness comparison cohort. Approved leave
  // does NOT count as a roster change: its weeks simply do not earn workload
  // opportunities, whether the leave is in the past or still scheduled.
  beginFairnessCohortWeek_(ctx.state,teamId,
    available.concat(leave).map(c=>c.ConsultantID),week);
  const manual = ctx.manual || {};
  const requested = [String(manual.subTeam1ConsultantId || '').trim(), String(manual.subTeam2ConsultantId || '').trim()];
  const manualOff = String(manual.offConsultantId || '').trim();
  const availableById = {};
  available.forEach(c => { availableById[String(c.ConsultantID)] = c; });

  // Off-duty remains optional and has a hard three-week minimum interval.
  // NEVER take someone out of inpatient consideration merely to secure Off-duty.
  const shouldOff = (leave.length > 0 && available.length >= 3) || (leave.length === 0 && available.length >= 4);
  // A newly hired consultant must receive an actual working assignment in the
  // first week of eligibility, not start their team rota on Off-duty.
  const firstEligibleWeek = c => !!c.ActiveFrom &&
    iso_(startOfWeekSunday_(toDate_(c.ActiveFrom))) === start;
  const eligibleOff = c => String(c.CanOffDuty) === 'Yes' &&
    !returnedFromLeaveThisWeek_(ctx,c) && !firstEligibleWeek(c) &&
    isOffDutyIntervalSafe_(c,ctx.state,week) && ctx.state.prevOff[teamId] !== c.ConsultantID;
  if (manualOff && (!shouldOff || !availableById[manualOff] || !eligibleOff(availableById[manualOff]))) {
    throw new Error('Manual seed: selected Off-duty consultant is unavailable or violates the minimum 3-week Off-duty interval for '+ctx.team.TeamName+' Week '+week+'.');
  }
  requested.forEach((id,i) => {
    if (id && availableById[id] && !consultantAllowedInSubTeam_(availableById[id],roles[i],roles,start)) throw new Error('Manual seed: consultant is restricted to a different sub-team ('+roles[i]+').');
    if (id && !availableById[id]) throw new Error('Manual seed: consultant for '+roles[i]+' is unavailable in '+ctx.team.TeamName+' Week '+week+'.');
  });
  if (requested[0] && requested[0] === requested[1]) {
    throw new Error('Manual seed: consultant unavailable or already assigned; the same consultant cannot cover both sub-teams in Week '+week+'.');
  }
  if (manualOff && requested.indexOf(manualOff) !== -1) {
    throw new Error('Manual seed: a consultant cannot be both inpatient and Off-duty in Week '+week+'.');
  }

  const assignments = [];
  if (available.length === 1) {
    const single = available[0];
    const selected = roles.find((role,i) => requested[i] === String(single.ConsultantID)) || '';
    const priority = String(selected || getSetting_('SubTeamCoverage|' + teamId + '|' + start, '') || ctx.team.OneConsultantPriority || '').trim();
    if (roles.indexOf(priority) === -1) {
      throw new Error('One consultant available in '+ctx.team.TeamName+' Week '+week+'. Main admin must select which sub-team to cover before generating or approving the rota.');
    }
    if (!consultantAllowedInSubTeam_(single,priority,roles,start))
      throw new Error('Single-consultant coverage decision conflicts with the consultant\'s fixed sub-team. Change the coverage choice or Consultant Settings.');
    const omitted = roles.find(x => x !== priority);
    assignments.push({role:priority,c:single});
    conflicts.push(conflict_(ctx,'Red','SUBTEAM_UNCOVERED',omitted+' uncovered; main-admin coverage priority applied.'));
  } else if (available.length >= 2) {
    // v3.0: Evaluate all placements together using recent workload fairness and previous TWO weeks, including
    // the category of the immediately preceding assignment. A short return to
    // inpatient, an inpatient->Off transition, and skipping Off all have explicit
    // costs. Skipping Off is allowed when preferable to a poor duty sequence.
    // A small, bounded next-week forecast discourages choices that cause an
    // otherwise avoidable inpatient repeat during an upcoming approved leave.
    const options = [];
    const compareScores = (a,b) => {
      for (let i=0;i<a.score.length;i++) if (a.score[i] !== b.score[i]) return a.score[i]-b.score[i];
      return a.key.localeCompare(b.key);
    };
    // A manual TCU/EMCU assignment is a new rotation anchor: on the next
    // inpatient assignment, prefer the OTHER service for that consultant.
    // This is tracked across OPD, Off-duty, approved leave and parent-Team
    // weeks. Fixed-service restrictions can make a swap impossible.
    const repeatedService = (c,role) => {
      const id=String(c.ConsultantID);
      const prev=(ctx.state.lastSubTeamRole||{})[String(teamId)+'|'+id]||'';
      return prev===role && roles.some(other=>other!==role && consultantAllowedInSubTeam_(c,other,roles,start)) ? 1 : 0;
    };
    const lastDutyGap = c => inpatientRecencyGap_(ctx.state,c.ConsultantID,week);
    const lastRole = c => {
      const id = String(c.ConsultantID);
      if (Number((ctx.state.lastAssignmentWeek || {})[id] || 0) === week-1)
        return String((ctx.state.lastRole||{})[id] || '');
      if (week===1) return String((ctx.state.roleCarryOver||{})[id]||'');
      return '';
    };
    const nextWeekStart = addDays_(ctx.weekStart,7), nextWeekEnd=addDays_(ctx.weekEnd,7);
    const nextWeekAvailable = ctx.consultants.filter(c =>
      !isInactiveForWeek_(c,nextWeekStart,nextWeekEnd) &&
      !isOnLeave_(c.ConsultantID,nextWeekStart,nextWeekEnd,ctx.leaves) &&
      isActiveForWeek_(c,nextWeekStart,nextWeekEnd));
    const observed = Number((ctx.state.subTeamWeeksObserved||{})[String(teamId)]||0);
    const pastOff = Number((ctx.state.subTeamOffWeeks||{})[String(teamId)]||0);
    // Off-duty remains common, not automatically eliminated: aim to fill at
    // least ~75% of eligible 4-person weeks when this competes with the
    // undesirable inpatient->Off transition. For 5-person teams, a safe Off
    // placement usually exists and is chosen every week. This is a preference,
    // not a minimum guarantee when leave or CanOffDuty rules prevent it.
    const offRate = observed ? pastOff/observed : 1;
    const skipOffCost = 18 + 320*Math.max(0,0.75-offRate);
    available.forEach(c1 => available.forEach(c2 => {
      const id1=String(c1.ConsultantID), id2=String(c2.ConsultantID);
      if (id1===id2 || (requested[0] && requested[0]!==id1) || (requested[1] && requested[1]!==id2)) return;
      if (!consultantAllowedInSubTeam_(c1,roles[0],roles,start) || !consultantAllowedInSubTeam_(c2,roles[1],roles,start)) return;
      if (manualOff && (manualOff===id1 || manualOff===id2)) return;
      const remainder=available.filter(c => c.ConsultantID!==c1.ConsultantID && c.ConsultantID!==c2.ConsultantID);
      const offOptions = shouldOff ? remainder.filter(eligibleOff) : [];
      const possibleOffs = manualOff ? offOptions.filter(c=>String(c.ConsultantID)===manualOff) : [null].concat(offOptions);
      possibleOffs.forEach(off => {
        const repeated = Number(wasInpatientPreviousWeek_(ctx.state,id1,week)) + Number(wasInpatientPreviousWeek_(ctx.state,id2,week));
        const serviceRepeat = repeatedService(c1,roles[0]) + repeatedService(c2,roles[1]);
        // gap=2 means Inpatient -> OPD/Off -> Inpatient, usually undesirable.
        // Gap=3 is preferable; >=4 has no recency penalty.
        const quickReturn = [c1,c2].reduce((total,c)=>{
          const gap=lastDutyGap(c);
          return total + (gap===2 ? 55 : gap===3 ? 9 : 0);
        },0);
        const offToInpatient = [c1,c2].filter(c=>lastRole(c)==='Off-duty').length*13;
        // All returning consultants use the normal inpatient preference.
        const returningRegular = available.filter(c=>returnedFromLeaveThisWeek_(ctx,c));
        const returnedInpatient = returningRegular.filter(c=>
          String(c.ConsultantID)===id1 || String(c.ConsultantID)===id2).length;
        const desiredReturnCount = returningRegular.length ? 1 : 0;
        const returnMissed = Math.abs(returnedInpatient-desiredReturnCount);
        const offAfterInpatient = off && lastDutyGap(off)===1 ? 34 : 0;
        const offSoonAfterInpatient = off && lastDutyGap(off)===2 ? 5 : 0;
        // Off-duty may be skipped, but not indefinitely without a good reason.
        const offSkipped = shouldOff && !off ? skipOffCost : 0;
        // Protect next week's two inpatient posts where known leave/inactivity
        // would otherwise force repeating one of THIS week's inpatient doctors.
        const forecastFresh = nextWeekAvailable.filter(c =>
          String(c.ConsultantID)!==id1 && String(c.ConsultantID)!==id2).length;
        const forecastRepeat = Math.max(0,Math.min(2,nextWeekAvailable.length)-forecastFresh)*24;
        // Opportunity-adjusted deficits prevent a new hire or a returning doctor
        // from being scheduled repeatedly to 'catch up' with pre-eligibility weeks.
        const proposed = {c1:c1,c2:c2,off:off};
        const fairnessCost = subTeamFairnessCost_(available,proposed,ctx.state,roles,teamId);
        // Off-duty equity has its own strong weight: inpatient and sequence
        // scores should not repeatedly bypass the consultant most overdue for Off.
        const offEquityCost = subTeamOffEquityCost_(available,off,ctx.state,teamId);
        // Compare the last EIGHT eligible weeks, not just total historic counts.
        // This prevents a consultant (including a consultant with no Off-duty)
        // from repeatedly absorbing inpatient duty over a short interval.
        const recentFairness = subTeamRecentFairnessCost_(available,proposed,ctx.state,roles,teamId,week);
        // Do not force the Off-duty slot if it worsens the clinical-duty sequence.
        const sequenceCost = quickReturn+offToInpatient+offAfterInpatient+
          offSoonAfterInpatient+offSkipped+forecastRepeat;
        // Hard priority: avoid avoidable consecutive inpatient assignments,
        // then give returning staff inpatient preference.
        // Sequence, rolling equity, and long-term equity jointly select the rest.
        // Protect inpatient spacing and return-from-leave first. Among
        // equally safe choices, enforce per-consultant TCU<->EMCU alternation
        // before optimizing workload totals. Without this explicit priority,
        // repeated TCU/EMCU assignments can follow a Manual Edit indefinitely.
        const score = [repeated, returnMissed, serviceRepeat,
          sequenceCost+fairnessCost+recentFairness+offEquityCost, recentFairness, sequenceCost, fairnessCost, offEquityCost];
        options.push({ c1:c1, c2:c2, off:off, score:score,
          key:id1+'|'+id2+'|'+(off?String(off.ConsultantID):'~') });
      });
    }));
    options.sort(compareScores);
    if (!options.length && !requested[0] && !requested[1] && !manualOff) {
      // Hard subgroup restrictions can make one of the posts impossible.
      // Never silently assign a TCU-only doctor to EMCU; show uncovered service.
      const chosen={};
      roles.forEach(role=>{
        const eligible=available.filter(c=>!chosen[String(c.ConsultantID)] && consultantAllowedInSubTeam_(c,role,roles,start));
        eligible.sort((a,b)=>inpatientRecencyGap_(ctx.state,b.ConsultantID,week)-inpatientRecencyGap_(ctx.state,a.ConsultantID,week));
        if (eligible.length) {const c=eligible[0];chosen[String(c.ConsultantID)]=true;assignments.push({role:role,c:c});}
        else conflicts.push(conflict_(ctx,'Red','SUBTEAM_UNCOVERED',role+' uncovered: no available consultant permitted for this sub-team. Review consultant restrictions.'));
      });
    } else if (!options.length) {
      throw new Error('No valid inpatient/Off-duty combination for '+ctx.team.TeamName+' Week '+week+'. Review the manual seed, consultant sub-team restrictions, and staffing.');
    }
    if (options.length) {
    const best = options[0];
    assignments.push({role:roles[0],c:best.c1},{role:roles[1],c:best.c2});
    if (best.off) assignments.push({role:'Off-duty',c:best.off});
    else if (shouldOff && available.some(c => String(c.CanOffDuty)==='Yes')) {
      // Skipping Off is planned, not a missing or inactive assignment.
      // All remaining available consultants are assigned to OPD below.
    }
    [best.c1,best.c2].forEach(c => {
      if (wasInpatientPreviousWeek_(ctx.state,c.ConsultantID,week)) {
        const manualRepeat = requested.indexOf(String(c.ConsultantID)) !== -1;
        conflicts.push(conflict_(ctx,'Yellow',manualRepeat ? 'MANUAL_CONSECUTIVE_INPATIENT' : 'CONSECUTIVE_INPATIENT_UNAVOIDABLE',
          c.Name+' has consecutive inpatient weeks. '+(manualRepeat ? 'This was explicitly selected by the administrator.' : 'No feasible non-consecutive assignment was available under current staffing and sub-team restrictions.')));
      }
    });
    }
  } else {
    roles.forEach(role => conflicts.push(conflict_(ctx,'Red','SUBTEAM_UNCOVERED',role+' has no available consultant.')));
  }
  const assigned = {};
  assignments.forEach(x => { assigned[String(x.c.ConsultantID)] = true; });
  const remaining = available.filter(c => !assigned[String(c.ConsultantID)]);
  remaining.sort((a,b)=>Number((ctx.state.opdCounts||{})[a.ConsultantID]||0)-Number((ctx.state.opdCounts||{})[b.ConsultantID]||0)||String(a.ConsultantID).localeCompare(String(b.ConsultantID)));
  remaining.forEach(c => assignments.push({role:'OPD',c:c}));
  assignments.forEach(x=>rows.push(row_(ctx.versionId,week,start,end,teamId,x.role,x.c.ConsultantID,'Auto-SubTeam','')));
  leave.forEach(c=>rows.push(row_(ctx.versionId,week,start,end,teamId,'Leave',c.ConsultantID,'Approved leave','')));
  inactive.forEach(c=>rows.push(row_(ctx.versionId,week,start,end,teamId,'Inactive',c.ConsultantID,'Inactive status','')));
  return {rows:rows,conflicts:conflicts};
}

function buildTeamWeekRows_(ctx) {
  const rows = [];
  const conflicts = [];
  const wkStartIso = iso_(ctx.weekStart);
  const wkEndIso = iso_(ctx.weekEnd);
  const prevTeam = ctx.state.prevTeam[ctx.team.TeamID] || '';
  const prevOff = ctx.state.prevOff[ctx.team.TeamID] || '';

  const inactive = [];
  const leave = [];
  const available = [];
  ctx.consultants.forEach(c => {
    if (isInactiveForWeek_(c, ctx.weekStart, ctx.weekEnd)) inactive.push(c);
    else if (!isActiveForWeek_(c, ctx.weekStart, ctx.weekEnd)) return;
    else if (isOnLeave_(c.ConsultantID, ctx.weekStart, ctx.weekEnd, ctx.leaves)) leave.push(c);
    else available.push(c);
  });

  beginFairnessCohortWeek_(ctx.state,ctx.team.TeamID,
    available.concat(leave).map(c=>c.ConsultantID),ctx.weekNo);
  let teamConsultant = null;
  let offConsultant = null;
  let source = 'Auto';

  if (ctx.manual && (ctx.manual.teamConsultantId || ctx.manual.offConsultantId)) {
    source = 'Manual';
    teamConsultant = available.find(c => c.ConsultantID === ctx.manual.teamConsultantId) || null;
    offConsultant = available.find(c => c.ConsultantID === ctx.manual.offConsultantId) || null;
    if (!teamConsultant) {
      conflicts.push(conflict_(ctx, 'Red', 'MANUAL_TEAM_MISSING', 'Manual seed week has no valid inpatient Team consultant. Auto-selection will be used.'));
      const combo = chooseWeeklyCombination_(available, ctx.state, ctx.team.TeamID, ctx);
      teamConsultant = combo.teamConsultant;
      offConsultant = combo.offConsultant;
      conflicts.push.apply(conflicts, combo.conflicts);
      source = 'Manual+AutoFix';
    } else {
      if (offConsultant && offConsultant.ConsultantID === teamConsultant.ConsultantID) offConsultant = null;
      if (offConsultant && String(offConsultant.CanOffDuty) !== 'Yes') offConsultant = null;
      if (offConsultant && returnedFromLeaveThisWeek_(ctx,offConsultant)) offConsultant = null;
      if (offConsultant && prevOff && offConsultant.ConsultantID === prevOff) {
        conflicts.push(conflict_(ctx, 'Yellow', 'CONSECUTIVE_OFF_DUTY_BLOCKED', `${offConsultant.Name} was entered for Off-duty two weeks in a row. Off-duty was skipped because this is not allowed.`));
        offConsultant = null;
      }
      if (offConsultant && !isOffDutyIntervalSafe_(offConsultant, ctx.state, ctx.weekNo)) {
        conflicts.push(conflict_(ctx, 'Yellow', 'OFF_DUTY_TOO_SOON_BLOCKED', `${offConsultant.Name} was entered for Off-duty before the normal three-week interval had passed. Off-duty was skipped to protect the rotation cycle.`));
        offConsultant = null;
      }
      if (prevTeam && teamConsultant.ConsultantID === prevTeam && available.length > 1) {
        conflicts.push(conflict_(ctx, 'Red', 'MANUAL_CONSECUTIVE_TEAM', `${teamConsultant.Name} was entered for inpatient Team two weeks in a row although another consultant from the same team is available.`));
      }

    }
  } else {
    const combo = chooseWeeklyCombination_(available, ctx.state, ctx.team.TeamID, ctx);
    teamConsultant = combo.teamConsultant;
    offConsultant = combo.offConsultant;
    conflicts.push.apply(conflicts, combo.conflicts);
  }

  if (!teamConsultant) {
    conflicts.push(conflict_(ctx, 'Red', 'NO_TEAM_COVER', 'No available consultant can cover inpatient Team work this week.'));
  }

  const assigned = {};
  if (teamConsultant) {
    assigned[teamConsultant.ConsultantID] = true;
    rows.push(row_(ctx.versionId, ctx.weekNo, wkStartIso, wkEndIso, ctx.team.TeamID, 'Team', teamConsultant.ConsultantID, source, ''));
  }
  if (offConsultant) {
    assigned[offConsultant.ConsultantID] = true;
    rows.push(row_(ctx.versionId, ctx.weekNo, wkStartIso, wkEndIso, ctx.team.TeamID, 'Off-duty', offConsultant.ConsultantID, source, ''));
  }

  const opdConsultants = available.filter(c => !assigned[c.ConsultantID]);
  const maxOpd = ctx.maxOpd;
  if (opdConsultants.length > maxOpd) {
    conflicts.push(conflict_(ctx, 'Yellow', 'OPD_OVER_MAX', `OPD has ${opdConsultants.length} consultants this week. The configured maximum is ${maxOpd}.`));
  }
  opdConsultants.forEach(c => {
    rows.push(row_(ctx.versionId, ctx.weekNo, wkStartIso, wkEndIso, ctx.team.TeamID, 'OPD', c.ConsultantID, source, ''));
  });
  leave.forEach(c => rows.push(row_(ctx.versionId, ctx.weekNo, wkStartIso, wkEndIso, ctx.team.TeamID, 'Leave', c.ConsultantID, 'Approved leave', '')));
  inactive.forEach(c => rows.push(row_(ctx.versionId, ctx.weekNo, wkStartIso, wkEndIso, ctx.team.TeamID, 'Inactive', c.ConsultantID, 'Inactive status', `Inactive from ${c.InactiveFrom}`)));
  return { rows: rows, conflicts: conflicts };
}

function chooseWeeklyCombination_(available, state, teamId, ctxForConflict) {
  const conflicts=[];
  if(!available.length) return {teamConsultant:null,offConsultant:null,conflicts:conflicts};
  const prevTeam=state.prevTeam[teamId]||'';
  const prevOff=state.prevOff[teamId]||'';
  const weekNo=Number((ctxForConflict&&ctxForConflict.weekNo)||0);
  const maxOpd=(ctxForConflict&&Number(ctxForConflict.maxOpd))||
    Math.max(1,Number(getSetting_('OPDMaxPerTeam','9')||9));
  const candidates=[];
  available.forEach(teamCandidate=>{
    if(available.length>1 && teamCandidate.ConsultantID===prevTeam)return;
    const offs=[null];
    if(available.length>=3)available.forEach(c=>{
      if(c.ConsultantID===teamCandidate.ConsultantID || String(c.CanOffDuty)!=='Yes' ||
         (ctxForConflict && returnedFromLeaveThisWeek_(ctxForConflict,c)) ||
         c.ConsultantID===prevOff || !isOffDutyIntervalSafe_(c,state,weekNo))return;
      offs.push(c);
    });
    offs.forEach(offCandidate=>{
      const opd=available.length-1-(offCandidate?1:0);
      if(available.length>=2 && opd<1)return;
      if(opd>maxOpd)return;
      candidates.push({teamConsultant:teamCandidate,offConsultant:offCandidate,
        score:scoreWeeklyCombination_(available,teamCandidate,offCandidate,state,teamId,weekNo,ctxForConflict)});
    });
  });
  // Standard four-consultant teams: prefer the explicit 1 Team + 2 OPD + 1 Off
  // allocation whenever an eligible Off-duty consultant can satisfy the cooldown.
  if (available.length === 4 && candidates.some(c=>c.offConsultant)) {
    for (let i=candidates.length-1; i>=0; i--) {
      if (!candidates[i].offConsultant) candidates.splice(i,1);
    }
  }
  if(!candidates.length){
    if(available.length===1)return {teamConsultant:available[0],offConsultant:null,
      conflicts:[conflict_(ctxForConflict||{weekNo:weekNo,weekStart:new Date(),team:{TeamID:teamId,TeamName:teamId}},
        'Yellow','ONLY_ONE_AVAILABLE',available[0].Name+' is the only available consultant; inpatient coverage may repeat.') ]};
    return {teamConsultant:null,offConsultant:null,conflicts:conflicts};
  }
  candidates.sort((a,b)=>a.score-b.score||
    String(a.teamConsultant.ConsultantID).localeCompare(String(b.teamConsultant.ConsultantID))||
    String(a.offConsultant?a.offConsultant.ConsultantID:'').localeCompare(String(b.offConsultant?b.offConsultant.ConsultantID:'')));
  const chosen=candidates[0];
  if(available.length>=3 && !chosen.offConsultant &&
     available.some(c=>String(c.CanOffDuty)==='Yes'))
    conflicts.push(conflict_(ctxForConflict||{weekNo:weekNo,weekStart:new Date(),team:{TeamID:teamId,TeamName:teamId}},
      'Yellow','OFF_DUTY_SKIPPED_BY_RULES','Off-duty was skipped to respect rotation constraints.'));
  return {teamConsultant:chosen.teamConsultant,offConsultant:chosen.offConsultant,conflicts:conflicts};
}

function returnedFromLeaveThisWeekStandard_(state, consultant, weekNo) {
  return Number((state.lastLeaveWeek||{})[String(consultant.ConsultantID)]||0)===Number(weekNo)-1;
}

function scoreWeeklyCombination_(available, teamConsultant, offConsultant, state, teamId, weekNo, ctx) {
  let score = 0;
  const assignedRole = {};
  assignedRole[teamConsultant.ConsultantID] = 'Team';
  if (offConsultant) assignedRole[offConsultant.ConsultantID] = 'Off-duty';
  available.forEach(c => {
    if (!assignedRole[c.ConsultantID]) assignedRole[c.ConsultantID] = 'OPD';
  });

  const prevOff = state.prevOff[teamId] || '';
  const returning=available.filter(c=>ctx ? returnedFromLeaveThisWeek_(ctx,c) : returnedFromLeaveThisWeekStandard_(state,c,weekNo));
  if (returning.length && !returning.some(c=>String(c.ConsultantID)===String(teamConsultant.ConsultantID))) score+=5000;
  const offEligible = available.filter(c => String(c.CanOffDuty) === 'Yes' &&
    !(ctx ? returnedFromLeaveThisWeek_(ctx,c) : returnedFromLeaveThisWeekStandard_(state,c,weekNo)));
  const eligibleNotPrevOff = offEligible.filter(c => c.ConsultantID !== prevOff);
  const dueForOff = eligibleNotPrevOff.filter(c => {
    const lastOffWeek = Number(state.lastOffWeek && state.lastOffWeek[c.ConsultantID] || 0);
    if (lastOffWeek > 0) return Number(weekNo || 0) >= lastOffWeek + 3;
    return expectedNextRole_(c, state.lastRole[c.ConsultantID] || '') === 'Off-duty';
  });
  const hasSeededRoles = available.some(c => state.lastRole[c.ConsultantID]);

  // Off-duty frequency rule:
  // - If at least 3 consultants are eligible for Off-duty, keeping one Off-duty every week can still be fair.
  // - If fewer than 3 are eligible, forcing Off-duty every week makes the eligible consultants off too often.
  //   In that case, Off-duty is assigned only when someone is actually due for Off-duty in the cycle
  //   or in the first seeded week to start the rotation.
  if (available.length >= 3 && eligibleNotPrevOff.length) {
    const shouldHaveOffThisWeek = offEligible.length >= 3 || dueForOff.length > 0 || !hasSeededRoles;
    if (shouldHaveOffThisWeek && !offConsultant) score += 8000;
    if (!shouldHaveOffThisWeek && offConsultant) score += 8000;
    if (dueForOff.length && offConsultant && !dueForOff.some(c => c.ConsultantID === offConsultant.ConsultantID)) score += 2500;
  }

  available.forEach(c => {
    const role = assignedRole[c.ConsultantID];
    const expected = expectedNextRole_(c, state.lastRole[c.ConsultantID] || '');
    if (expected) {
      if (role === expected) score -= 80;
      else score += roleMismatchPenalty_(expected, role);
    }

    // Eligibility-adjusted fairness: joining halfway through the period must
    // not create a fictitious inpatient backlog. Restrict historical counts to
    // the weeks in which the consultant was actually available to this team.
    const key=String(teamId)+'|'+String(c.ConsultantID);
    const actualTeam=cohortMetric_(state,teamId,'standardActualTeam',key)+(role==='Team'?1:0);
    const expectedTeam=cohortMetric_(state,teamId,'standardExpectedTeam',key)+1/available.length;
    score += (actualTeam-expectedTeam)*(actualTeam-expectedTeam)*115;
    // A long uninterrupted OPD streak should not be the default just because
    // previous OPD was rewarded by the old role-cycle rule.
    if(role==='OPD') {
      const streak=Number((state.standardOpdStreak||{})[key]||0);
      score += Math.pow(Math.max(0,streak-2),2)*65;
    }
    if(String(c.CanOffDuty)==='Yes') {
      const possibleOff=offConsultant ? 1/offEligible.length : 0;
      const actualOff=cohortMetric_(state,teamId,'standardActualOff',key)+(role==='Off-duty'?1:0);
      const expectedOff=cohortMetric_(state,teamId,'standardExpectedOff',key)+possibleOff;
      score += (actualOff-expectedOff)*(actualOff-expectedOff)*35;
    }

    // A consultant who cannot take Off-duty must never be selected for it.
    if (String(c.CanOffDuty) !== 'Yes' && role === 'Off-duty') score += 99999;
  });

  // Prefer assigning Off-duty to someone whose cycle expects Off-duty.
  if (offConsultant) {
    const expected = expectedNextRole_(offConsultant, state.lastRole[offConsultant.ConsultantID] || '');
    if (expected === 'Off-duty') score -= 180;
  }

  return score;
}

function expectedNextRole_(consultant, lastRole) {
  if (!lastRole) return '';
  if (String(consultant.CanOffDuty) === 'Yes') {
    if (lastRole === 'Team') return 'OPD';
    if (lastRole === 'OPD') return 'Off-duty';
    if (lastRole === 'Off-duty') return 'Team';
  } else {
    // Do not reward repeating OPD indefinitely for consultants unable to take Off.
    // Their inpatient turn comes from eligibility-adjusted workload fairness.
    if (lastRole === 'Team') return 'OPD';
    if (lastRole === 'OPD') return '';
    if (lastRole === 'Off-duty') return 'OPD';
  }
  return '';
}

function roleMismatchPenalty_(expected, actual) {
  if (expected === actual) return 0;
  if (expected === 'OPD' && actual === 'Team') return 12;
  if (expected === 'OPD' && actual === 'Off-duty') return 50;
  if (expected === 'Team' && actual === 'OPD') return 25;
  if (expected === 'Team' && actual === 'Off-duty') return 55;
  if (expected === 'Off-duty' && actual === 'OPD') return 30;
  if (expected === 'Off-duty' && actual === 'Team') return 70;
  return 40;
}

// Effective-dated eligibility snapshots. This avoids rewriting historical
// Off-duty expectations when an administrator changes a current preference.
function eligibilityHistory_() {
  // Isolated VM tests have no SpreadsheetApp. Real Apps Script always reads the
  // persisted history; an inaccessible real sheet should fail safely.
  if(typeof SpreadsheetApp==='undefined') return [];
  return readObjects_(SHEETS.EligibilityHistory).sort((a,b)=>String(a.EffectiveFrom).localeCompare(String(b.EffectiveFrom)));
}
function consultantAtWeek_(state, consultant, weekStart) {
  if(!weekStart)return consultant;
  const id=String(consultant.ConsultantID), date=iso_(weekStart);
  const history=((state.eligibilityHistory||{})[id]||[]);
  let chosen=null;
  for(const h of history)if(String(h.EffectiveFrom).slice(0,10)<=date)chosen=h;
  if(!chosen && history.length)chosen=history[0];
  return chosen?Object.assign({},consultant,{CanOffDuty:chosen.CanOffDuty,
    ConsultantType:chosen.ConsultantType,SubTeamAssignment:chosen.SubTeamAssignment,
    SubTeamAssignmentFrom:chosen.EffectiveFrom}):consultant;
}
function recordEligibilityChange_(oldConsultant,newConsultant,effective) {
  const prior=readObjects_(SHEETS.EligibilityHistory).filter(r=>String(r.ConsultantID)===String(oldConsultant.ConsultantID));
  const ids=[];
  const add=(c,date)=>{const id=makeId_('EH');appendObject_(SHEETS.EligibilityHistory,{
    HistoryID:id,ConsultantID:c.ConsultantID,TeamID:c.TeamID,EffectiveFrom:date,
    CanOffDuty:normalizeYesNo_(c.CanOffDuty,'Yes'),ConsultantType:normalizeConsultantType_(c.ConsultantType||'Regular'),
    SubTeamAssignment:normalizeSubTeamAssignment_(c.SubTeamAssignment),RecordedAt:now_()});ids.push(id);};
  if(!prior.length){
    const initial=String(oldConsultant.ActiveFrom||'1900-01-01').slice(0,10);
    const oldRestrictionFrom=String(oldConsultant.SubTeamAssignmentFrom||'').slice(0,10);
    // Existing database versions already recorded when a fixed service began.
    // Retain the earlier unrestricted eligibility rather than retroactively
    // treating every pre-assignment week as fixed to today's service.
    if(oldRestrictionFrom && oldRestrictionFrom>initial &&
       normalizeSubTeamAssignment_(oldConsultant.SubTeamAssignment)!=='Any'){
      add(Object.assign({},oldConsultant,{SubTeamAssignment:'Any'}),initial);
      add(oldConsultant,oldRestrictionFrom);
    }else add(oldConsultant,initial);
  }
  add(newConsultant,effective);
  return ids;
}

// Staffing-cohort fairness: old workload is real and remains in the history,
// but its *imbalance* is not a debt to carry from a 3-person rota into a new
// 5-person or 4-person rota. Off-duty spacing and last clinical roles are NOT reset.
const FAIRNESS_COHORT_METRICS_ = [
  'standardActualTeam','standardExpectedTeam','standardActualOff','standardExpectedOff',
  'scopedTeamCounts','scopedExpectedTeam','scopedSubCounts','scopedExpectedSub',
  'scopedOffCounts','scopedExpectedOff'
];
function beginFairnessCohortWeek_(state, teamId, eligibleIds, weekNo) {
  const team=String(teamId), ids=[...new Set((eligibleIds||[]).map(String).filter(Boolean))].sort();
  const signature=ids.join('|');
  if (!state.fairnessCohorts) state.fairnessCohorts={};
  const previous=state.fairnessCohorts[team];
  if (previous && previous.signature===signature) return;
  const base={};
  const prefixes=ids.map(id=>team+'|'+id);
  FAIRNESS_COHORT_METRICS_.forEach(metric=>{
    const current=state[metric]||{}, snapshot={};
    Object.keys(current).forEach(key=>{
      if (prefixes.some(prefix=>key===prefix || key.startsWith(prefix+'|')))
        snapshot[key]=Number(current[key]||0);
    });
    base[metric]=snapshot;
  });
  state.fairnessCohorts[team]={signature,firstWeek:Number(weekNo)||0,base};
  // Recent fairness also belongs to the new cohort. Historical rows and duty
  // recency are preserved independently for clinical continuity.
  if (state.subTeamRecentWeeks) state.subTeamRecentWeeks[team]=[];
}
function cohortMetric_(state, teamId, metric, key) {
  const current=Number(((state[metric]||{})[key])||0);
  const baseline=Number((((((state.fairnessCohorts||{})[String(teamId)]||{}).base||{})[metric]||{})[key])||0);
  return current-baseline;
}

function seedStateFromRows_(rows, consultants) {
  const state = { standardActualTeam: {}, standardExpectedTeam: {}, standardOpdStreak: {}, standardActualOff: {}, standardExpectedOff: {}, teamCounts: {}, offCounts: {}, prevTeam: {}, prevOff: {}, lastRole: {}, lastSubTeamRole: {}, lastTeamWeek: {}, lastOffWeek: {}, lastLeaveWeek: {}, lastAssignmentWeek: {}, lastOffTeamWeek: {}, subTeamWeeksObserved: {}, subTeamOffWeeks: {}, subTeamRecentWeeks: {}, subCounts: {}, opdCounts: {}, subRoleMap: {}, expectedTeamCounts: {}, expectedSubCounts: {}, expectedOffCounts: {}, scopedTeamCounts: {}, scopedSubCounts: {}, scopedOffCounts: {}, scopedExpectedTeam: {}, scopedExpectedSub: {}, scopedExpectedOff: {}, consultantRoster: {}, fairnessCohorts: {} };
  state.eligibilityHistory = {};
  eligibilityHistory_().forEach(h=>{const id=String(h.ConsultantID);(state.eligibilityHistory[id]||(state.eligibilityHistory[id]=[])).push(h);});
  getTeams_().forEach(t => { if (isSubTeamConfigured_(t)) state.subRoleMap[t.TeamID] = subTeamNames_(t); });
  (consultants || getConsultants_()).forEach(c => { state.consultantRoster[String(c.ConsultantID)] = c; });
  updateStateFromWeekRows_(state, rows || []);
  return state;
}

// Historical workload is eligibility-adjusted: no catch-up for weeks before
// joining or during leave. All consultants use the same rotation calculation.
function subTeamWorkloadOpportunity_(available, assignedRoles, state, teamId, weekStart) {
  const expected = { team:{}, sub:{}, off:{} };
  const roster = {}, roleNames=((state.subRoleMap||{})[String(teamId)]||[]);
  available=available.map(c=>consultantAtWeek_(state,c,weekStart));
  available.forEach(c=>{roster[String(c.ConsultantID)]=c;});
  // A consultant unable to take Off-duty has a slightly lower inpatient target,
  // but their inpatient target is never based solely on sub-team restrictions.
  const weight = c => String(c.CanOffDuty)==='No' ? 0.90 : 1;
  const offEligible=available.filter(c=>String(c.CanOffDuty)==='Yes');
  let inpatientSlots=0;
  assignedRoles.forEach(a=>{
    const role=String(a.role);
    if(role==='Off-duty'){
      if(offEligible.length)offEligible.forEach(c=>{
        const id=String(c.ConsultantID);expected.off[id]=(expected.off[id]||0)+1/offEligible.length;
      });
      return;
    }
    if(role!=='Team' && roleNames.indexOf(role)===-1)return;
    inpatientSlots++;
    // Per-SERVICE expectations are conditional on the consultant's permissions.
    // They are deliberately separate from the pooled inpatient entitlement.
    if(role!=='Team'){
      const eligible=available.filter(c=>consultantAllowedInSubTeam_(c,role,roleNames,weekStart));
      const denominator=eligible.reduce((sum,c)=>sum+weight(c),0);
      if(denominator)eligible.forEach(c=>{
        const key=String(c.ConsultantID)+'|'+role;
        expected.sub[key]=(expected.sub[key]||0)+weight(c)/denominator;
      });
    }
  });
  // Two clinicians with equal eligibility should have approximately the same
  // *combined* inpatient workload even when one is restricted to TCU only.
  // Allocating total duty by each service's restricted pool would incorrectly
  // leave the fixed-service consultant underutilized and overwork colleagues.
  const ipEligible=available.filter(c=>roleNames.length===0 || roleNames.some(role=>consultantAllowedInSubTeam_(c,role,roleNames,weekStart)));
  const denom=ipEligible.reduce((sum,c)=>sum+weight(c),0);
  if(denom)ipEligible.forEach(c=>{
    expected.team[String(c.ConsultantID)]=inpatientSlots*weight(c)/denom;
  });
  return expected;
}

function addSubTeamOpportunity_(state, teamRows) {
  const teamId=String(teamRows[0].TeamID || ''), names=state.subRoleMap[teamId]||[];
  if (!names.length) return; // Standard teams keep their historical rules unchanged.
  const eligibleRoles=['Team','OPD','Reserve','Off-duty'].concat(names);
  const eligibleIds={};
  const workAssignments=[];
  teamRows.forEach(r=>{
    const id=String(r.ConsultantID || '');
    if (!id || isManualCoverId_(id) || eligibleRoles.indexOf(r.AssignmentPlace)===-1) return;
    eligibleIds[id] = true;
    if (r.AssignmentPlace==='Team' || names.indexOf(r.AssignmentPlace)!==-1 || r.AssignmentPlace==='Off-duty')
      workAssignments.push({role:r.AssignmentPlace,id:id});
  });
  const available=Object.keys(eligibleIds).map(id=>consultantAtWeek_(state,
    state.consultantRoster[id]||{ConsultantID:id,ConsultantType:'Regular',CanOffDuty:'Yes'},teamRows[0].WeekStart));
  const next=subTeamWorkloadOpportunity_(available,workAssignments,state,teamId,teamRows[0].WeekStart);
  if (!state.subTeamRecentWeeks) state.subTeamRecentWeeks={};
  if (!state.subTeamRecentWeeks[teamId]) state.subTeamRecentWeeks[teamId]=[];
  state.subTeamRecentWeeks[teamId].push({
    weekNo:Number(teamRows[0].WeekNo)||0,
    eligible:available.map(c=>({
      id:String(c.ConsultantID),
      expected:Number(next.team[String(c.ConsultantID)]||0),
      inpatient:Number(workAssignments.some(a=>a.id===String(c.ConsultantID) && a.role!=='Off-duty'))
    }))
  });
  // Only the last eight weeks can contribute to the rolling fairness term.
  if (state.subTeamRecentWeeks[teamId].length>8) state.subTeamRecentWeeks[teamId].shift();
  Object.keys(next.team).forEach(id=>{
    state.expectedTeamCounts[id]=(state.expectedTeamCounts[id]||0)+next.team[id];
    const k=teamId+'|'+id;state.scopedExpectedTeam[k]=(state.scopedExpectedTeam[k]||0)+next.team[id];
  });
  Object.keys(next.sub).forEach(id=>{
    state.expectedSubCounts[id]=(state.expectedSubCounts[id]||0)+next.sub[id];
    const k=teamId+'|'+id;state.scopedExpectedSub[k]=(state.scopedExpectedSub[k]||0)+next.sub[id];
  });
  Object.keys(next.off).forEach(id=>{
    state.expectedOffCounts[id]=(state.expectedOffCounts[id]||0)+next.off[id];
    const k=teamId+'|'+id;state.scopedExpectedOff[k]=(state.scopedExpectedOff[k]||0)+next.off[id];
  });
}

// Recent workload has a separate fairness term so a balanced 52-week total
// cannot disguise an unfair six- to eight-week stretch. Expected workload is
// eligibility-adjusted, excluding pre-joining and approved-leave weeks.
function subTeamRecentFairnessCost_(available, selected, state, roles, teamId, weekNo) {
  const recent=((state.subTeamRecentWeeks||{})[String(teamId)]||[])
    .filter(x=>Number(x.weekNo)>=Number(weekNo)-7 && Number(x.weekNo)<Number(weekNo));
  const assigned=[{role:roles[0],id:String(selected.c1.ConsultantID)},
    {role:roles[1],id:String(selected.c2.ConsultantID)}];
  if (selected.off) assigned.push({role:'Off-duty',id:String(selected.off.ConsultantID)});
  const expectation=subTeamWorkloadOpportunity_(available,assigned,state,teamId);
  let imbalance=0, noOffExcess=0;
  available.forEach(c=>{
    const id=String(c.ConsultantID), history=recent.flatMap(x=>x.eligible.filter(e=>e.id===id));
    const actual=history.reduce((sum,e)=>sum+e.inpatient,0) +
      Number(id===String(selected.c1.ConsultantID)||id===String(selected.c2.ConsultantID));
    const expected=history.reduce((sum,e)=>sum+e.expected,0) + Number(expectation.team[id]||0);
    const deviation=actual-expected;
    imbalance+=deviation*deviation;
    // Extra protection for someone who cannot take Off-duty: avoid a recent
    // inpatient surplus, which is especially burdensome without regular Off.
    if (String(c.CanOffDuty)==='No') noOffExcess+=Math.pow(Math.max(0,deviation),2);
  });
  return imbalance*19+noOffExcess*28;
}

// Off-duty fairness across people who CAN take Off, adjusted for actual
// eligible exposure. Never compare a newly joined doctor against an older
// consultant's raw number of Off weeks.
function subTeamOffEquityCost_(available, chosenOff, state, teamId) {
  if (!chosenOff) return 0; // Skipping Off is handled by the separate rate/sequence cost.
  const pool=available.filter(c=>String(c.CanOffDuty)==='Yes');
  if (pool.length<2) return 0;
  const chosen=String(chosenOff.ConsultantID);
  const prior=pool.map(c=>{
    const key=String(teamId)+'|'+String(c.ConsultantID);
    return {id:String(c.ConsultantID), actual:cohortMetric_(state,teamId,'scopedOffCounts',key),
      expected:cohortMetric_(state,teamId,'scopedExpectedOff',key)};
  });
  // Expected Off allocation increases by 1/pool.length for this Off slot.
  // Penalize unequal *deficits*, not unequal historical totals.
  const diffs=prior.map(p=>p.actual+(p.id===chosen?1:0)-p.expected-1/pool.length);
  const dispersion=diffs.reduce((sum,d)=>sum+d*d,0);
  const mostOwed=prior.reduce((m,p)=>Math.max(m,p.expected-p.actual),-Infinity);
  const chosenOwed=prior.find(p=>p.id===chosen).expected-prior.find(p=>p.id===chosen).actual;
  return dispersion*75 + Math.max(0,mostOwed-chosenOwed)*35 - 85;
}

function subTeamFairnessCost_(available, selected, state, roles, teamId) {
  const assignments=[{role:roles[0],id:String(selected.c1.ConsultantID)}, {role:roles[1],id:String(selected.c2.ConsultantID)}];
  if(selected.off) assignments.push({role:'Off-duty',id:String(selected.off.ConsultantID)});
  const target=subTeamWorkloadOpportunity_(available,assignments,state,teamId);
  let inpatient=0, services=0, off=0;
  available.forEach(c=>{
    const id=String(c.ConsultantID);
    const chosen=(id===String(selected.c1.ConsultantID)||id===String(selected.c2.ConsultantID)) ? 1 : 0;
    const scope=String(teamId)+'|'+id;
    const actual=cohortMetric_(state,teamId,'scopedTeamCounts',scope)+chosen;
    const expected=cohortMetric_(state,teamId,'scopedExpectedTeam',scope)+Number(target.team[id]||0);
    inpatient+=(actual-expected)*(actual-expected);
    roles.forEach((role,i)=>{
      const key=id+'|'+role, scoped=String(teamId)+'|'+key;
      const value=cohortMetric_(state,teamId,'scopedSubCounts',scoped)+(String((i===0?selected.c1:selected.c2).ConsultantID)===id?1:0);
      const aim=cohortMetric_(state,teamId,'scopedExpectedSub',scoped)+Number(target.sub[key]||0);
      services+=(value-aim)*(value-aim);
    });
    if(String(c.CanOffDuty)==='Yes'){
      const actualOff=cohortMetric_(state,teamId,'scopedOffCounts',scope)+(selected.off&&String(selected.off.ConsultantID)===id?1:0);
      const targetOff=cohortMetric_(state,teamId,'scopedExpectedOff',scope)+Number(target.off[id]||0);
      off+=(actualOff-targetOff)*(actualOff-targetOff);
    }
  });
  return inpatient*24+services*9+off*11;
}


function updateStateFromWeekRows_(state, rows) {
  if (!state.lastRole) state.lastRole = {};
  if (!state.lastSubTeamRole) state.lastSubTeamRole = {};
  if (!state.lastTeamWeek) state.lastTeamWeek = {};
  if (!state.lastOffWeek) state.lastOffWeek = {};
  if (!state.lastLeaveWeek) state.lastLeaveWeek = {};
  if (!state.lastAssignmentWeek) state.lastAssignmentWeek = {};
  if (!state.lastOffTeamWeek) state.lastOffTeamWeek = {};
  if (!state.subTeamWeeksObserved) state.subTeamWeeksObserved = {};
  if (!state.subTeamOffWeeks) state.subTeamOffWeeks = {};
  if (!state.expectedTeamCounts) state.expectedTeamCounts = {};
  if (!state.expectedSubCounts) state.expectedSubCounts = {};
  if (!state.expectedOffCounts) state.expectedOffCounts = {};
  if (!state.consultantRoster) state.consultantRoster = {};
  for (const k of ['standardActualTeam','standardExpectedTeam','standardOpdStreak','standardActualOff','standardExpectedOff'])
    if (!state[k]) state[k] = {};
  for (const k of ['scopedTeamCounts','scopedSubCounts','scopedOffCounts','scopedExpectedTeam','scopedExpectedSub','scopedExpectedOff'])
    if (!state[k]) state[k] = {};
  if (!rows || !rows.length) return;

  const groups = {};
  rows.forEach(r => {
    const key = `${Number(r.WeekNo) || 0}|${String(r.TeamID || '')}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });

  Object.keys(groups).sort((a, b) => {
    const aw = Number(a.split('|')[0]) || 0;
    const bw = Number(b.split('|')[0]) || 0;
    if (aw !== bw) return aw - bw;
    return String(a).localeCompare(String(b));
  }).forEach(key => {
    const group = groups[key];
    const historicalTeamId=String(group[0].TeamID||'');
    const inpatientRoles=(state.subRoleMap||{})[historicalTeamId]||[];
    const present=group.filter(r=>!isManualCoverId_(r.ConsultantID) &&
      ['Team','OPD','Reserve','Off-duty','Leave'].concat(inpatientRoles).includes(String(r.AssignmentPlace)))
      .map(r=>String(r.ConsultantID));
    beginFairnessCohortWeek_(state,historicalTeamId,present,group[0].WeekNo);
    addSubTeamOpportunity_(state, group);
    const teamId = String(group[0].TeamID || '');
    // Snapshot fair opportunities from the assignments that were actually
    // eligible in this historic week. Do not retroactively count weeks before
    // a consultant joined, or weeks on Leave/Inactive, against them.
    if (group.some(r=>r.AssignmentPlace==='Team')) {
      const working = group.filter(r=>['Team','OPD','Off-duty','Reserve'].includes(r.AssignmentPlace) && !isManualCoverId_(r.ConsultantID));
      const byId = {};
      working.forEach(r=>{ byId[String(r.ConsultantID)] = r; });
      const ids = Object.keys(byId);
      ids.forEach(id=>{
        const key=teamId+'|'+id, role=byId[id].AssignmentPlace;
        state.standardExpectedTeam[key]=Number(state.standardExpectedTeam[key]||0)+1/ids.length;
        if(role==='Team') state.standardActualTeam[key]=Number(state.standardActualTeam[key]||0)+1;
        state.standardOpdStreak[key]=(role==='OPD'||role==='Reserve') ? Number(state.standardOpdStreak[key]||0)+1 : 0;
      });
      const offIds=ids.filter(id=>{
        const c=state.consultantRoster[id]||{ConsultantID:id,CanOffDuty:'Yes'};
        const historic=consultantAtWeek_(state,c,group[0].WeekStart);
        return String(historic.CanOffDuty)==='Yes';
      });
      if (offIds.length && working.some(r=>r.AssignmentPlace==='Off-duty')) {
        offIds.forEach(id=>{
          const key=teamId+'|'+id;
          state.standardExpectedOff[key]=Number(state.standardExpectedOff[key]||0)+1/offIds.length;
        });
      }
      working.filter(r=>r.AssignmentPlace==='Off-duty').forEach(r=>{
        const key=teamId+'|'+String(r.ConsultantID);
        state.standardActualOff[key]=Number(state.standardActualOff[key]||0)+1;
      });
    }

    let weekTeam = '';
    let weekOff = '';

    group.forEach(r => {
      // External/manual cover names are display-only. They must not influence future auto-generation.
      if (isManualCoverId_(r.ConsultantID)) return;
      if (r.AssignmentPlace === 'Team' || ((state.subRoleMap && state.subRoleMap[teamId]) || []).indexOf(r.AssignmentPlace) !== -1) {
        if (r.AssignmentPlace !== 'Team') {
          // Preserve the last actual sub-team assignment, even when this week
          // was manually edited. This is NOT reset by intervening OPD, Off,
          // leave or a temporarily deactivated parent-team rota.
          state.lastSubTeamRole[teamId+'|'+String(r.ConsultantID)] = String(r.AssignmentPlace);
          if (!state.subCounts) state.subCounts = {};
          const key = r.ConsultantID + '|' + r.AssignmentPlace;
          state.subCounts[key] = Number(state.subCounts[key] || 0) + 1;
        }
        if((state.subRoleMap[teamId]||[]).length){
          const k=teamId+'|'+r.ConsultantID;
          state.scopedTeamCounts[k]=Number(state.scopedTeamCounts[k]||0)+1;
          if(r.AssignmentPlace!=='Team'){ const sr=k+'|'+r.AssignmentPlace;state.scopedSubCounts[sr]=Number(state.scopedSubCounts[sr]||0)+1; }
        }
        state.teamCounts[r.ConsultantID] = Number(state.teamCounts[r.ConsultantID] || 0) + 1;
        state.lastRole[r.ConsultantID] = 'Team';
        state.lastTeamWeek[r.ConsultantID] = Number(r.WeekNo) || 0;
        weekTeam = r.ConsultantID;
      }
      // Previous-version Reserve shifts are counted as OPD for future fairness,
      // while archived assignment records remain untouched.
      if (r.AssignmentPlace === 'OPD' || r.AssignmentPlace === 'Reserve') {
        state.opdCounts = state.opdCounts || {};
        state.opdCounts[r.ConsultantID] = (state.opdCounts[r.ConsultantID] || 0) + 1;
        state.lastRole[r.ConsultantID] = 'OPD';
      }
      if (r.AssignmentPlace === 'Off-duty') {
        if((state.subRoleMap[teamId]||[]).length){ const k=teamId+'|'+r.ConsultantID;state.scopedOffCounts[k]=Number(state.scopedOffCounts[k]||0)+1; }
        state.offCounts[r.ConsultantID] = Number(state.offCounts[r.ConsultantID] || 0) + 1;
        state.lastRole[r.ConsultantID] = 'Off-duty';
        state.lastOffWeek[r.ConsultantID] = Number(r.WeekNo) || 0;
        weekOff = r.ConsultantID;
        state.lastOffTeamWeek[teamId] = Number(r.WeekNo) || 0;
      }
      if (['Team','OPD','Reserve','Off-duty'].indexOf(r.AssignmentPlace) !== -1 ||
          ((state.subRoleMap && state.subRoleMap[teamId]) || []).indexOf(r.AssignmentPlace) !== -1) {
        state.lastAssignmentWeek[r.ConsultantID] = Number(r.WeekNo) || 0;
      }
      // Leave pauses workload accumulation, but record its week so a returning
      // consultant is not sent straight into Off-duty on the first week back.
      if (r.AssignmentPlace === 'Leave') state.lastLeaveWeek[r.ConsultantID] = Number(r.WeekNo) || 0;
      // Inactive also pauses the consultant's last working role.
    });

    if ((state.subRoleMap[teamId] || []).some(role => group.some(r => r.AssignmentPlace === role))) {
      state.subTeamWeeksObserved[teamId] = Number(state.subTeamWeeksObserved[teamId]||0)+1;
      if (weekOff) state.subTeamOffWeeks[teamId] = Number(state.subTeamOffWeeks[teamId]||0)+1;
    }

    // These are true previous-week markers. If Off-duty was skipped this week,
    // prevOff is cleared so the last Off-duty consultant is not blocked forever.
    state.prevTeam[teamId] = weekTeam;
    state.prevOff[teamId] = weekOff;
  });
}


function getRotaWeekView_(weekNo) {
  const latest = getLatestVersion_();
  if (!latest) return { meta: null, week: null };
  if (!weekNo || weekNo < 1 || weekNo > Number(latest.Weeks || 52)) throw new Error('Select a valid week.');
  const rows = getVersionAssignments_(latest.VersionID).filter(r => Number(r.WeekNo) === Number(weekNo));
  const teams = teamMap_();
  const consultants = consultantMap_();
  let weekObj = null;
  rows.forEach(r => {
    if (!weekObj) {
      weekObj = { weekNo: Number(r.WeekNo), period: `${nice_(r.WeekStart)} – ${nice_(r.WeekEnd)}`, periodExport: `${dateOnlyNice_(r.WeekStart)} – ${dateOnlyNice_(r.WeekEnd)}`, teams: {} };
    }
    const team = teams[r.TeamID] || { TeamName: r.TeamID, ColorHex: '#0f172a' };
    if (!weekObj.teams[r.TeamID]) {
      weekObj.teams[r.TeamID] = { teamId: r.TeamID, teamName: team.TeamName, teamColor: normalizeHexColor_(team.ColorHex || '#0f172a'), subTeamNames: subTeamNamesForView_(team, r.WeekStart, [r]),
        scheduledStandardFrom: (!isSubTeamActiveForWeek_(team, r.WeekStart) ? String(team.SubTeamUntil || '').slice(0,10) : ''), scheduledSubTeamsFrom: String(team.SubTeamResumeFrom || '').slice(0,10), weekStart: String(r.WeekStart || '').slice(0,10), Team: [], OPD: [], 'Off-duty': [], Reserve: [], Leave: [], Inactive: [] };
    }
    const person = personFromAssignment_(r, consultants);
    if (!weekObj.teams[r.TeamID].subTeamNames.length) weekObj.teams[r.TeamID].subTeamNames = subTeamNamesForView_(team, r.WeekStart, [r]);
    weekObj.teams[r.TeamID].subTeamNames.forEach(role => { if (!weekObj.teams[r.TeamID][role]) weekObj.teams[r.TeamID][role] = []; });
    if (!weekObj.teams[r.TeamID][r.AssignmentPlace]) weekObj.teams[r.TeamID][r.AssignmentPlace] = [];
    weekObj.teams[r.TeamID][r.AssignmentPlace].push(person);
  });
  if (!weekObj) {
    const start = toDate_(latest.StartDate);
    const ws = addDays_(start, (weekNo - 1) * 7);
    const we = addDays_(ws, 6);
    weekObj = { weekNo: weekNo, period: `${nice_(ws)} – ${nice_(we)}`, periodExport: `${dateOnlyNice_(ws)} – ${dateOnlyNice_(we)}`, teams: [] };
  } else {
    weekObj = Object.assign({}, weekObj, { teams: Object.values(weekObj.teams).sort((a, b) => a.teamName.localeCompare(b.teamName)) });
  }
  return { meta: latest, week: weekObj };
}


function getRotaWeekTeamView_(weekNo, teamId) {
  const full = getRotaWeekView_(weekNo);
  if (!full.week) return full;
  if (!teamId) throw new Error('Select a team to edit.');
  const teams = Array.isArray(full.week.teams) ? full.week.teams : [];
  const selectedTeam = teams.find(t => String(t.teamId || '') === String(teamId));
  if (!selectedTeam) {
    const team = teamMap_()[teamId] || { TeamName: teamId, ColorHex: '#0f172a' };
    full.week.teams = [{
      teamId: teamId,
      teamName: team.TeamName || teamId,
      teamColor: normalizeHexColor_(team.ColorHex || '#0f172a'),
      Team: [], OPD: [], 'Off-duty': [], Reserve: [], Leave: [], Inactive: [], subTeamNames: subTeamNamesForView_(team, addDays_(toDate_(full.meta.StartDate), (weekNo - 1)*7), [])
    }];
  } else {
    full.week.teams = [selectedTeam];
  }
  // The manual editor must always reveal configured services, even if this
  // published week is still Standard or the effective date is still future.
  // A separate flag prevents publication until activation is valid.
  const selected = full.week.teams[0];
  const config = teamMap_()[teamId];
  const date = addDays_(toDate_(full.meta.StartDate), (weekNo - 1) * 7);
  selected.subTeamConfigured = isSubTeamConfigured_(config);
  selected.subTeamActivated = isSubTeamActiveForWeek_(config, date);
  selected.subTeamEffectiveFrom = (config && config.SubTeamFrom) || '';
  selected.subTeamEffectiveUntil = (config && config.SubTeamUntil) || '';
  selected.subTeamResumeFrom = (config && config.SubTeamResumeFrom) || '';
  if (selected.subTeamActivated) {
    selected.subTeamNames = subTeamNames_(config);
    selected.subTeamNames.forEach(role => { if (!selected[role]) selected[role] = []; });
  } else if (selected.subTeamConfigured) {
    const stopped = !!selected.subTeamEffectiveUntil && iso_(date) >= String(selected.subTeamEffectiveUntil).slice(0,10) &&
      (!selected.subTeamResumeFrom || iso_(date) < String(selected.subTeamResumeFrom).slice(0,10));
    // Before activation show configured fields with the existing activation warning.
    // After deactivation show the Standard Team editor, preserving old published data.
    selected.subTeamNames = stopped ? [] : subTeamNames_(config);
    selected.subTeamNames.forEach(role => { if (!selected[role]) selected[role] = []; });
  }
  return Object.assign({}, full, {subTeamEditorVersion:3});
}

function getLatestRotaView_() {
  const latest = getLatestVersion_();
  if (!latest) return { meta: null, weeks: [] };
  const rows = getVersionAssignments_(latest.VersionID);
  const teams = teamMap_();
  const consultants = consultantMap_();
  const weekMap = {};
  rows.forEach(r => {
    const weekKey = String(r.WeekNo);
    if (!weekMap[weekKey]) {
      weekMap[weekKey] = { weekNo: Number(r.WeekNo), period: `${nice_(r.WeekStart)} – ${nice_(r.WeekEnd)}`, periodExport: `${dateOnlyNice_(r.WeekStart)} – ${dateOnlyNice_(r.WeekEnd)}`, teams: {} };
    }
    const team = teams[r.TeamID] || { TeamName: r.TeamID, ColorHex: '#0f172a' };
    const teamName = team.TeamName;
    if (!weekMap[weekKey].teams[r.TeamID]) {
      weekMap[weekKey].teams[r.TeamID] = { teamId: r.TeamID, teamName: teamName, teamColor: normalizeHexColor_(team.ColorHex || '#0f172a'), subTeamNames: subTeamNamesForView_(team, r.WeekStart, [r]),
        scheduledStandardFrom: (!isSubTeamActiveForWeek_(team, r.WeekStart) ? String(team.SubTeamUntil || '').slice(0,10) : ''), scheduledSubTeamsFrom: String(team.SubTeamResumeFrom || '').slice(0,10), weekStart: String(r.WeekStart || '').slice(0,10), Team: [], OPD: [], 'Off-duty': [], Reserve: [], Leave: [], Inactive: [] };
    }
    if (hidePublicInactiveAssignment_(r, consultants)) {
      if (String(r.AssignmentPlace) !== 'Inactive') weekMap[weekKey].teams[r.TeamID].hasHiddenInactiveAssignments = true;
      return;
    }
    const person = personFromAssignment_(r, consultants);
    if (!weekMap[weekKey].teams[r.TeamID].subTeamNames.length) weekMap[weekKey].teams[r.TeamID].subTeamNames = subTeamNamesForView_(team, r.WeekStart, [r]);
    weekMap[weekKey].teams[r.TeamID].subTeamNames.forEach(role => { if (!weekMap[weekKey].teams[r.TeamID][role]) weekMap[weekKey].teams[r.TeamID][role] = []; });
    if (!weekMap[weekKey].teams[r.TeamID][r.AssignmentPlace]) weekMap[weekKey].teams[r.TeamID][r.AssignmentPlace] = [];
    weekMap[weekKey].teams[r.TeamID][r.AssignmentPlace].push(person);
  });
  const weeks = Object.values(weekMap)
    .sort((a, b) => a.weekNo - b.weekNo)
    .map(w => Object.assign({}, w, { teams: sortRotaTeamViews_(Object.values(w.teams)) }));
  return { meta: latest, weeks: weeks };
}


function getLatestRotaWindow_(windowWeeks) {
  const latest = getLatestVersion_();
  if (!latest) return { meta: null, weeks: [], viewMode: 'window' };
  const totalWeeks = Number(latest.Weeks || 52);
  const start = toDate_(latest.StartDate);
  const todayIso = Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd');
  const today = toDate_(todayIso);
  let currentWeek = getWeekNoForDate_(start, today, totalWeeks);
  let startWeek;
  let summary;
  if (currentWeek < 1) {
    startWeek = 1;
    summary = 'The rota has not started yet. Showing the first upcoming weeks.';
  } else if (currentWeek > totalWeeks) {
    startWeek = Math.max(1, totalWeeks - 3);
    summary = 'This rota period has ended. Showing the final weeks.';
  } else {
    startWeek = currentWeek;
    summary = 'Showing the current rota week and upcoming weeks. Use View Full 52-Week Rota for the complete rota.';
  }
  const endWeek = Math.min(totalWeeks, startWeek + Math.max(1, Number(windowWeeks || 8)) - 1);
  const view = getLatestRotaViewForWeekRange_(latest, startWeek, endWeek);
  view.viewMode = 'window';
  view.window = { startWeek: startWeek, endWeek: endWeek, totalWeeks: totalWeeks, currentWeek: currentWeek };
  view.summary = summary;
  return view;
}

function getLatestRotaViewForWeekRange_(latest, startWeek, endWeek) {
  if (!latest) return { meta: null, weeks: [] };
  const rows = getVersionAssignments_(latest.VersionID).filter(r => Number(r.WeekNo) >= Number(startWeek) && Number(r.WeekNo) <= Number(endWeek));
  const teams = teamMap_();
  const consultants = consultantMap_();
  const weekMap = {};
  rows.forEach(r => {
    const weekKey = String(r.WeekNo);
    if (!weekMap[weekKey]) {
      weekMap[weekKey] = { weekNo: Number(r.WeekNo), period: `${nice_(r.WeekStart)} – ${nice_(r.WeekEnd)}`, periodExport: `${dateOnlyNice_(r.WeekStart)} – ${dateOnlyNice_(r.WeekEnd)}`, teams: {} };
    }
    const team = teams[r.TeamID] || { TeamName: r.TeamID, ColorHex: '#0f172a' };
    const teamName = team.TeamName;
    if (!weekMap[weekKey].teams[r.TeamID]) {
      weekMap[weekKey].teams[r.TeamID] = { teamId: r.TeamID, teamName: teamName, teamColor: normalizeHexColor_(team.ColorHex || '#0f172a'), subTeamNames: subTeamNamesForView_(team, r.WeekStart, [r]),
        scheduledStandardFrom: (!isSubTeamActiveForWeek_(team, r.WeekStart) ? String(team.SubTeamUntil || '').slice(0,10) : ''), scheduledSubTeamsFrom: String(team.SubTeamResumeFrom || '').slice(0,10), weekStart: String(r.WeekStart || '').slice(0,10), Team: [], OPD: [], 'Off-duty': [], Reserve: [], Leave: [], Inactive: [] };
    }
    if (hidePublicInactiveAssignment_(r, consultants)) {
      if (String(r.AssignmentPlace) !== 'Inactive') weekMap[weekKey].teams[r.TeamID].hasHiddenInactiveAssignments = true;
      return;
    }
    const person = personFromAssignment_(r, consultants);
    if (!weekMap[weekKey].teams[r.TeamID].subTeamNames.length) weekMap[weekKey].teams[r.TeamID].subTeamNames = subTeamNamesForView_(team, r.WeekStart, [r]);
    weekMap[weekKey].teams[r.TeamID].subTeamNames.forEach(role => { if (!weekMap[weekKey].teams[r.TeamID][role]) weekMap[weekKey].teams[r.TeamID][role] = []; });
    if (!weekMap[weekKey].teams[r.TeamID][r.AssignmentPlace]) weekMap[weekKey].teams[r.TeamID][r.AssignmentPlace] = [];
    weekMap[weekKey].teams[r.TeamID][r.AssignmentPlace].push(person);
  });
  const weeks = Object.values(weekMap)
    .sort((a, b) => a.weekNo - b.weekNo)
    .map(w => Object.assign({}, w, { teams: sortRotaTeamViews_(Object.values(w.teams)) }));
  return { meta: latest, weeks: weeks };
}

function groupRowsForTeamPreview_(rows) {
  const consultants = consultantMap_();
  const map = {};
  rows.forEach(r => {
    const key = `${r.WeekNo}|${r.WeekStart}`;
    if (!map[key]) map[key] = { weekNo: Number(r.WeekNo), period: `${nice_(r.WeekStart)} – ${nice_(r.WeekEnd)}`, Team: [], OPD: [], 'Off-duty': [], Leave: [], Inactive: [] };
    const person = personFromAssignment_(r, consultants);
    if (!map[key][r.AssignmentPlace]) map[key][r.AssignmentPlace] = [];
    map[key][r.AssignmentPlace].push(person);
  });
  return Object.values(map).sort((a, b) => a.weekNo - b.weekNo);
}

function deriveImpactLevel_(conflicts, leave, latest) {
  const currentEnd = latest ? toDate_(latest.EndDate) : null;
  if (currentEnd && toDate_(leave.StartDate) > currentEnd) return 'Future';
  if (conflicts.some(c => c.Level === 'Red')) return 'Conflict';
  if (conflicts.some(c => c.Level === 'Yellow')) return 'Warning';
  return 'Safe';
}

function summarizeImpact_(level, conflicts, leave, latest) {
  if (level === 'Safe') return 'Leave can be approved and the rota can be amended without major conflict.';
  if (level === 'Warning') return `Leave can probably be approved, but ${conflicts.length} warning item(s) were found. Off-duty may be skipped in affected weeks.`;
  if (level === 'Conflict') return `Leave creates ${conflicts.length} conflict/warning item(s). Admin can reject it or approve with override if clinically necessary.`;
  return 'Leave is outside the current rota period.';
}

function manualAssignmentMap_(manualAssignments) {
  const map = {};
  (manualAssignments || []).forEach(m => {
    map[`${m.weekNo}|${m.teamId}`] = {
      teamConsultantId: m.teamConsultantId || '',
      offConsultantId: m.offConsultantId || '',
      subTeam1ConsultantId: m.subTeam1ConsultantId || '',
      subTeam2ConsultantId: m.subTeam2ConsultantId || ''
    };
  });
  return map;
}

function appendConflicts_(versionId, conflicts) {
  if (!conflicts.length) return;
  const rows = conflicts.map(c => ({
    ConflictID: makeId_('X'),
    VersionID: versionId,
    WeekNo: c.WeekNo,
    WeekStart: c.WeekStart,
    TeamID: c.TeamID,
    Level: c.Level,
    Rule: c.Rule,
    Message: c.Message,
    CreatedAt: now_()
  }));
  appendRows_(SHEETS.ConflictLog, rows);
}

function appendOverride_(o) {
  appendObject_(SHEETS.OverrideLog, {
    OverrideID: makeId_('O'),
    VersionID: o.VersionID || '',
    RequestType: o.requestType || '',
    RequestID: o.requestId || '',
    RuleBroken: o.ruleBroken || '',
    Reason: o.reason || '',
    CreatedAt: now_()
  });
}

function row_(versionId, weekNo, weekStart, weekEnd, teamId, place, consultantId, source, notes) {
  return { VersionID: versionId, WeekNo: weekNo, WeekStart: weekStart, WeekEnd: weekEnd, TeamID: teamId, AssignmentPlace: place, ConsultantID: consultantId, Source: source, Notes: notes || '' };
}

function conflict_(ctx, level, rule, message) {
  return { WeekNo: ctx.weekNo, WeekStart: iso_(ctx.weekStart), TeamID: ctx.team.TeamID, Level: level, Rule: rule, Message: message };
}

function getApprovedLeaves_() {
  return readObjects_(SHEETS.LeaveRequests).filter(l => normalizeStatus_(l.Status) === 'approved');
}

function isOnLeave_(consultantId, weekStart, weekEnd, leaves) {
  return leaves.some(l => l.ConsultantID === consultantId && datesOverlap_(toDate_(l.StartDate), toDate_(l.EndDate), weekStart, weekEnd));
}

function isActiveForWeek_(c, weekStart, weekEnd) {
  const activeFrom = c.ActiveFrom ? toDate_(c.ActiveFrom) : new Date('1900-01-01');
  const inactiveFrom = c.InactiveFrom ? toDate_(c.InactiveFrom) : null;
  if (activeFrom > weekEnd) return false;
  if (inactiveFrom && inactiveFrom <= weekEnd) return false;
  return true;
}

function isInactiveForWeek_(c, weekStart, weekEnd) {
  const inactiveFrom = c.InactiveFrom ? toDate_(c.InactiveFrom) : null;
  return !!(inactiveFrom && inactiveFrom <= weekEnd);
}


function periodSignature_(startDate, endDate, weeks) {
  return [String(startDate || ''), String(endDate || ''), String(weeks || '')].join('|');
}

function defaultPeriodName_(startDate, endDate) {
  return `Rota ${String(startDate || '')} to ${String(endDate || '')}`;
}

function createRotaPeriod_(p) {
  const obj = {
    PeriodID: makeId_('P'),
    PeriodName: String(p.PeriodName || '').trim() || defaultPeriodName_(p.StartDate, p.EndDate),
    Status: p.Status || 'Draft',
    StartDate: p.StartDate,
    EndDate: p.EndDate,
    Weeks: p.Weeks,
    CreatedAt: now_(),
    CreatedBy: Session.getActiveUser().getEmail() || 'Admin',
    ActivatedAt: p.Status === 'Active' ? now_() : '',
    ArchivedAt: p.Status === 'Archived' ? now_() : '',
    Notes: p.Notes || ''
  };
  appendObject_(SHEETS.RotaPeriods, obj);
  return obj;
}

function getRotaPeriods_() {
  return readObjects_(SHEETS.RotaPeriods).sort((a, b) => {
    const rank = { Active: 0, Draft: 1, Archived: 2, Voided: 3 };
    const ar = Object.prototype.hasOwnProperty.call(rank, a.Status) ? rank[a.Status] : 9;
    const br = Object.prototype.hasOwnProperty.call(rank, b.Status) ? rank[b.Status] : 9;
    if (ar !== br) return ar - br;
    const ds = String(b.StartDate || '').localeCompare(String(a.StartDate || ''));
    if (ds !== 0) return ds;
    return String(b.CreatedAt || '').localeCompare(String(a.CreatedAt || ''));
  });
}

function getPeriodById_(periodId) {
  const id = String(periodId || '').trim();
  return id ? getRotaPeriods_().find(p => String(p.PeriodID) === id) || null : null;
}

function getActivePeriod_() {
  const active = getRotaPeriods_().filter(p => p.Status === 'Active');
  if (!active.length) return null;
  active.sort((a, b) => String(b.ActivatedAt || b.CreatedAt || '').localeCompare(String(a.ActivatedAt || a.CreatedAt || '')));
  return active[0];
}

function nextAmendmentNoForPeriod_(periodId) {
  const id = String(periodId || '').trim();
  const values = readObjects_(SHEETS.RotaVersions)
    .filter(v => !id || String(v.PeriodID || '') === id)
    .map(v => Number(v.AmendmentNo || 0))
    .filter(n => isFinite(n));
  return values.length ? Math.max.apply(null, values) + 1 : 1;
}

function activatePeriod_(periodId) {
  const id = String(periodId || '').trim();
  if (!id) throw new Error('Rota period ID is required.');
  getRotaPeriods_().forEach(p => {
    if (p.Status === 'Active' && String(p.PeriodID) !== id) {
      updateRow_(SHEETS.RotaPeriods, 'PeriodID', p.PeriodID, { Status: 'Archived', ArchivedAt: now_() });
    }
  });
  updateRow_(SHEETS.RotaPeriods, 'PeriodID', id, { Status: 'Active', ActivatedAt: now_(), ArchivedAt: '' });
}

function archivePeriod_(periodId) {
  const id = String(periodId || '').trim();
  if (!id) return;
  const period = getPeriodById_(id);
  if (period && period.Status !== 'Voided') {
    updateRow_(SHEETS.RotaPeriods, 'PeriodID', id, { Status: 'Archived', ArchivedAt: now_() });
  }
}

function deleteDraftPeriodIfEmpty_(periodId) {
  const id = String(periodId || '').trim();
  if (!id) return;
  const period = getPeriodById_(id);
  if (!period || period.Status !== 'Draft') return;
  const remaining = readObjects_(SHEETS.RotaVersions).some(v => String(v.PeriodID || '') === id);
  if (!remaining) deleteRowsWhere_(SHEETS.RotaPeriods, row => String(row.PeriodID || '') === id);
}

function enrichVersionWithPeriod_(version) {
  if (!version) return version;
  const period = getPeriodById_(version.PeriodID);
  return Object.assign({}, version, {
    PeriodName: period ? period.PeriodName : '',
    PeriodStatus: period ? period.Status : '',
    PeriodStartDate: period ? period.StartDate : version.StartDate,
    PeriodEndDate: period ? period.EndDate : version.EndDate
  });
}

function ensureRotaPeriodMigration_() {
  const versions = readObjects_(SHEETS.RotaVersions);
  if (!versions.length) return;

  const existingPeriods = readObjects_(SHEETS.RotaPeriods);
  const periodById = {};
  existingPeriods.forEach(p => periodById[String(p.PeriodID)] = p);

  const groups = {};
  versions.forEach(v => {
    const suppliedId = String(v.PeriodID || '').trim();
    const signature = periodSignature_(v.StartDate, v.EndDate, v.Weeks);
    const key = suppliedId ? `ID:${suppliedId}` : `SIG:${signature}`;
    if (!groups[key]) groups[key] = { key: key, periodId: suppliedId, signature: signature, versions: [] };
    groups[key].versions.push(v);
  });

  // The group containing the latest legacy Official version becomes Active. Other date-range
  // groups are preserved as Draft/Archived periods rather than being mixed into one history.
  const officialVersions = versions.filter(v => v.Status === 'Official').sort((a, b) => {
    const an = Number(a.AmendmentNo || 0), bn = Number(b.AmendmentNo || 0);
    if (bn !== an) return bn - an;
    return String(b.CreatedAt || '').localeCompare(String(a.CreatedAt || ''));
  });
  const latestOfficialId = officialVersions.length ? String(officialVersions[0].VersionID) : '';
  let activeGroupKey = '';
  Object.keys(groups).forEach(key => {
    if (groups[key].versions.some(v => String(v.VersionID) === latestOfficialId)) activeGroupKey = key;
  });

  Object.keys(groups).forEach(key => {
    const g = groups[key];
    const sample = g.versions[0];
    let period = g.periodId ? periodById[g.periodId] : null;
    if (!period) {
      const hasDraft = g.versions.some(v => v.Status === 'Draft');
      const status = key === activeGroupKey ? 'Active' : (hasDraft && !g.versions.some(v => v.Status === 'Official') ? 'Draft' : 'Archived');
      period = createRotaPeriod_({
        PeriodName: defaultPeriodName_(sample.StartDate, sample.EndDate),
        Status: status,
        StartDate: sample.StartDate,
        EndDate: sample.EndDate,
        Weeks: sample.Weeks,
        Notes: 'Created automatically during v1.6 rota-period migration.'
      });
      g.periodId = period.PeriodID;
      periodById[g.periodId] = period;
    }
    g.versions.forEach(v => {
      if (String(v.PeriodID || '') !== String(g.periodId)) {
        updateRow_(SHEETS.RotaVersions, 'VersionID', v.VersionID, { PeriodID: g.periodId });
      }
    });
  });

  // Ensure exactly the period containing the current Official version is marked Active.
  if (latestOfficialId) {
    const refreshed = readObjects_(SHEETS.RotaVersions);
    const latest = refreshed.find(v => String(v.VersionID) === latestOfficialId);
    if (latest && latest.PeriodID) activatePeriod_(latest.PeriodID);
  }

  const versionToPeriod = {};
  readObjects_(SHEETS.RotaVersions).forEach(v => versionToPeriod[String(v.VersionID)] = String(v.PeriodID || ''));
  backfillAssignmentPeriodIds_(versionToPeriod);
}

function backfillAssignmentPeriodIds_(versionToPeriod) {
  const sh = ensureSheet_(SHEETS.RotaAssignments, HEADERS.RotaAssignments);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0].map(String);
  const versionCol = headers.indexOf('VersionID');
  const periodCol = headers.indexOf('PeriodID');
  if (versionCol === -1 || periodCol === -1) return;
  let changed = false;
  for (let i = 1; i < values.length; i++) {
    if (!String(values[i][periodCol] || '').trim()) {
      const periodId = versionToPeriod[String(values[i][versionCol] || '')] || '';
      if (periodId) {
        values[i][periodCol] = periodId;
        changed = true;
      }
    }
  }
  if (changed) sh.getRange(1, 1, values.length, headers.length).setValues(values);
}

function createRotaVersion_(v) {
  const periodId = String(v.PeriodID || '').trim() || String((getActivePeriod_() || {}).PeriodID || '');
  const obj = {
    VersionID: makeId_('V'),
    AmendmentNo: v.AmendmentNo || nextAmendmentNoForPeriod_(periodId),
    Status: v.Status || 'Official',
    Type: v.Type || '',
    CreatedAt: now_(),
    CreatedBy: Session.getActiveUser().getEmail() || 'Admin',
    StartDate: v.StartDate,
    EndDate: v.EndDate,
    Weeks: v.Weeks,
    EffectiveFrom: v.EffectiveFrom || v.StartDate,
    Reason: v.Reason || '',
    BasedOnVersionID: v.BasedOnVersionID || '',
    ExportDocId: '',
    ExportDocUrl: '',
    PeriodID: periodId,
    AssignmentStorage: v.AssignmentStorage || 'Live',
    ArchiveSheetName: v.ArchiveSheetName || '',
    AssignmentCount: Number(v.AssignmentCount || 0)
  };
  appendObject_(SHEETS.RotaVersions, obj);
  return enrichVersionWithPeriod_(obj);
}

function archiveOfficialVersions_(periodId) {
  const id = String(periodId || '').trim();
  const versions = readObjects_(SHEETS.RotaVersions).filter(v => v.Status === 'Official' && (!id || String(v.PeriodID || '') === id));
  versions.forEach(v => archiveVersion_(v.VersionID));
}

function archiveVersion_(versionId) {
  archiveVersionAssignments_(versionId);
  updateRow_(SHEETS.RotaVersions, 'VersionID', versionId, { Status: 'Archived' });
}

function assignmentArchiveSheetName_(version) {
  const amend = String(Number(version.AmendmentNo || 0)).padStart(3, '0');
  const suffix = String(version.VersionID || '').replace(/[^A-Za-z0-9]/g, '').slice(-10) || 'version';
  return (`A${amend}_${suffix}`).slice(0, 90);
}

function getOrCreatePeriodArchive_(periodId) {
  const period = getPeriodById_(periodId);
  if (!period) throw new Error('Rota period was not found for assignment archiving.');
  let archiveId = String(period.ArchiveSpreadsheetID || '').trim();
  let archive;
  if (archiveId) {
    // Never replace an existing archive link: earlier archived amendments of this
    // period live only in that file. A transient Drive error or a permission change
    // must stop the operation (callers roll back) instead of silently re-linking.
    try { archive = SpreadsheetApp.openById(archiveId); }
    catch (err) {
      throw new Error('The assignment archive spreadsheet for rota period ' + (period.PeriodName || period.PeriodID) +
        ' could not be opened (ID ' + archiveId + '). No changes were kept. Check that the script owner still has Drive access to that file, then retry.');
    }
    return archive;
  }
  // No archive exists yet for this period: create it once and record its ID.
  {
    archive = SpreadsheetApp.create(`${APP.title} Archive - ${period.PeriodName || period.PeriodID}`.slice(0, 95));
    archiveId = archive.getId();
    try {
      const sourceFile = DriveApp.getFileById(getDb_().getId());
      const targetFile = DriveApp.getFileById(archiveId);
      const parents = sourceFile.getParents();
      if (parents.hasNext()) {
        const folder = parents.next();
        folder.addFile(targetFile);
        try { DriveApp.getRootFolder().removeFile(targetFile); } catch (_) {}
      }
    } catch (_) {}
    updateRow_(SHEETS.RotaPeriods, 'PeriodID', period.PeriodID, {
      ArchiveSpreadsheetID: archiveId,
      ArchiveSpreadsheetURL: archive.getUrl()
    });
  }
  return archive;
}

function writeArchiveSheet_(archive, sheetName, rows) {
  let sh = archive.getSheetByName(sheetName);
  if (!sh) sh = archive.insertSheet(sheetName);
  sh.clear();
  const headers = HEADERS.RotaAssignments;
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
  if (rows.length) {
    const values = rows.map(r => headers.map(h => r[h] === undefined ? '' : r[h]));
    sh.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  sh.autoResizeColumns(1, headers.length);
}

function archiveVersionAssignments_(versionId) {
  const version = readObjects_(SHEETS.RotaVersions).find(v => String(v.VersionID) === String(versionId));
  if (!version) throw new Error('Rota version was not found for assignment archiving.');
  const liveRows = readObjects_(SHEETS.RotaAssignments).filter(r => String(r.VersionID) === String(versionId));
  if (!liveRows.length) {
    if (version.ArchiveSheetName) return { archived: false, count: Number(version.AssignmentCount || 0), alreadyArchived: true };
    return { archived: false, count: 0, alreadyArchived: false };
  }
  const archive = getOrCreatePeriodArchive_(version.PeriodID);
  const sheetName = String(version.ArchiveSheetName || assignmentArchiveSheetName_(version));
  writeArchiveSheet_(archive, sheetName, liveRows);
  const removed = removeVersionAssignmentsFromLive_(version.VersionID);
  if (removed !== liveRows.length) throw new Error(`Assignment archive safety check failed: expected to move ${liveRows.length} rows but moved ${removed}.`);
  updateRow_(SHEETS.RotaVersions, 'VersionID', version.VersionID, {
    AssignmentStorage: 'Archive', ArchiveSheetName: sheetName, AssignmentCount: liveRows.length
  });
  return { archived: true, count: liveRows.length, sheetName: sheetName };
}

function readArchivedAssignments_(version) {
  const period = getPeriodById_(version.PeriodID);
  const archiveId = period && String(period.ArchiveSpreadsheetID || '').trim();
  const sheetName = String(version.ArchiveSheetName || '').trim();
  if (!archiveId || !sheetName) return [];
  let archive;
  try { archive = SpreadsheetApp.openById(archiveId); }
  catch (_) { throw new Error('The assignment archive file could not be opened. Check Drive access or run Safe Repair.'); }
  const sh = archive.getSheetByName(sheetName);
  if (!sh) throw new Error(`Archived assignments for Amendment ${version.AmendmentNo} were not found.`);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(v => v !== '' && v !== null)).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = normalizeCell_(row[i]));
    return obj;
  });
}

function restoreVersionAssignmentsToLive_(versionId) {
  const version = readObjects_(SHEETS.RotaVersions).find(v => String(v.VersionID) === String(versionId));
  if (!version) throw new Error('Rota version was not found for restoration.');
  const existing = readObjects_(SHEETS.RotaAssignments).filter(r => String(r.VersionID) === String(versionId));
  let rows = null;
  if (existing.length) {
    // An interrupted archive move can leave only part of a version in the live sheet
    // while the archive copy (written first) is complete. Prefer the complete copy.
    const lookup = Object.assign({}, version, { ArchiveSheetName: version.ArchiveSheetName || assignmentArchiveSheetName_(version) });
    try { rows = readArchivedAssignments_(lookup); } catch (_) { return existing; }
    if (rows.length <= existing.length) return existing;
    deleteRowsWhere_(SHEETS.RotaAssignments, r => String(r.VersionID) === String(versionId));
  }
  if (!rows) rows = readArchivedAssignments_(version);
  if (!rows.length) throw new Error('This archived amendment has no stored assignments and cannot be restored.');
  appendRows_(SHEETS.RotaAssignments, rows.map(r => Object.assign({}, r, { VersionID: version.VersionID, PeriodID: version.PeriodID })));
  updateRow_(SHEETS.RotaVersions, 'VersionID', version.VersionID, { AssignmentStorage: 'Live', AssignmentCount: rows.length });
  return rows;
}

function clearAssignmentStorageSummaryCache_() {
  CacheService.getScriptCache().remove('assignmentStorageSummary_v1_8');
}

function getAssignmentStorageSummaryCached_() {
  const cache = CacheService.getScriptCache();
  const key = 'assignmentStorageSummary_v1_8';
  const raw = cache.get(key);
  if (raw) { try { return JSON.parse(raw); } catch (_) {} }
  const summary = assignmentStorageSummary_();
  try { cache.put(key, JSON.stringify(summary), 300); } catch (_) {}
  return summary;
}

function assignmentStorageSummary_() {
  const versions = readObjects_(SHEETS.RotaVersions);
  const assignments = readObjects_(SHEETS.RotaAssignments);
  const keep = {};
  versions.filter(v => v.Status === 'Official' || v.Status === 'Draft').forEach(v => keep[String(v.VersionID)] = true);
  const liveByVersion = {};
  assignments.forEach(r => { const id = String(r.VersionID || ''); if (id) liveByVersion[id] = (liveByVersion[id] || 0) + 1; });
  const eligible = versions.filter(v => !keep[String(v.VersionID)] && (liveByVersion[String(v.VersionID)] || 0) > 0);
  return {
    liveRows: assignments.length,
    currentAndDraftRows: Object.keys(keep).reduce((n, id) => n + Number(liveByVersion[id] || 0), 0),
    historicalVersionsReadyToArchive: eligible.length,
    historicalRowsReadyToArchive: eligible.reduce((n, v) => n + Number(liveByVersion[String(v.VersionID)] || 0), 0)
  };
}

function migrateAssignmentStorage_() {
  const versions = readObjects_(SHEETS.RotaVersions);
  if (!versions.length) {
    updateSetting_('AssignmentArchiveMigrationV18', 'Done');
    return { status: 'Done', archivedVersions: 0, archivedRows: 0, summary: assignmentStorageSummary_() };
  }
  const assignments = readObjects_(SHEETS.RotaAssignments);
  const liveByVersion = {};
  assignments.forEach(r => {
    const id = String(r.VersionID || '');
    if (!id) return;
    if (!liveByVersion[id]) liveByVersion[id] = [];
    liveByVersion[id].push(r);
  });
  const keep = {};
  versions.filter(v => v.Status === 'Official' || v.Status === 'Draft').forEach(v => keep[String(v.VersionID)] = true);
  const eligible = versions.filter(v => !keep[String(v.VersionID)] && (liveByVersion[String(v.VersionID)] || []).length);

  // Write every archive copy first. The live sheet is not reduced until all archive writes succeed.
  const archivePlans = [];
  eligible.forEach(v => {
    const rows = liveByVersion[String(v.VersionID)] || [];
    const archive = getOrCreatePeriodArchive_(v.PeriodID);
    const sheetName = String(v.ArchiveSheetName || assignmentArchiveSheetName_(v));
    writeArchiveSheet_(archive, sheetName, rows);
    archivePlans.push({ version: v, rows: rows, sheetName: sheetName });
  });

  // Mark the versions as archived before reducing the live sheet. If the rewrite fails, duplicate
  // live rows remain, which is safer than losing access to an amendment.
  archivePlans.forEach(p => updateRow_(SHEETS.RotaVersions, 'VersionID', p.version.VersionID, {
    AssignmentStorage: 'Archive', ArchiveSheetName: p.sheetName, AssignmentCount: p.rows.length
  }));
  versions.filter(v => keep[String(v.VersionID)]).forEach(v => {
    const count = (liveByVersion[String(v.VersionID)] || []).length;
    if (count) updateRow_(SHEETS.RotaVersions, 'VersionID', v.VersionID, { AssignmentStorage: 'Live', AssignmentCount: count });
  });

  // Remove only the rows that were just archived; kept rows are never rewritten.
  const archivedIds = {};
  archivePlans.forEach(p => { archivedIds[String(p.version.VersionID)] = true; });
  if (archivePlans.length) deleteRowsWhere_(SHEETS.RotaAssignments, r => !!archivedIds[String(r.VersionID || '')]);
  const archivedRows = archivePlans.reduce((n, p) => n + p.rows.length, 0);
  updateSetting_('AssignmentArchiveMigrationV18', 'Done');
  return { status: archivePlans.length ? 'Migrated' : 'Done', archivedVersions: archivePlans.length, archivedRows: archivedRows, summary: assignmentStorageSummary_() };
}

function archiveHistoricalAssignments(adminToken, confirmText) {
  const session = requireAdminPin_(adminToken);
  if (String(confirmText || '').trim() !== 'ARCHIVE HISTORY') throw new Error('Type ARCHIVE HISTORY to confirm.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const before = assignmentStorageSummary_();
    const result = migrateAssignmentStorage_();
    appendAudit_(session, 'HISTORICAL_ASSIGNMENTS_ARCHIVED', 'RotaAssignments', 'History', { before: before, result: result });
    return {
      ok: true,
      message: `Historical assignment archive completed. ${result.archivedVersions} version(s) and ${result.archivedRows} row(s) were moved out of the live RotaAssignments sheet.`,
      result: result
    };
  } finally {
    lock.releaseLock();
  }
}

function trashAssignmentArchives_() {
  const ids = {};
  readObjects_(SHEETS.RotaPeriods).forEach(p => { if (p.ArchiveSpreadsheetID) ids[String(p.ArchiveSpreadsheetID)] = true; });
  let trashed = 0;
  Object.keys(ids).forEach(id => {
    try { DriveApp.getFileById(id).setTrashed(true); trashed++; } catch (_) {}
  });
  return trashed;
}

function getRotaVersions_(knownPeriods) {
  const periods = {};
  (knownPeriods || getRotaPeriods_()).forEach(p => periods[String(p.PeriodID)] = p);
  const versions = readObjects_(SHEETS.RotaVersions).map(v => {
    const p = periods[String(v.PeriodID || '')];
    return Object.assign({}, v, {
      PeriodName: p ? p.PeriodName : '',
      PeriodStatus: p ? p.Status : '',
      PeriodStartDate: p ? p.StartDate : v.StartDate,
      PeriodEndDate: p ? p.EndDate : v.EndDate
    });
  });
  versions.sort((a, b) => {
    const pd = String(b.PeriodStartDate || '').localeCompare(String(a.PeriodStartDate || ''));
    if (pd !== 0) return pd;
    const ab = Number(b.AmendmentNo || 0) - Number(a.AmendmentNo || 0);
    if (ab !== 0) return ab;
    return String(b.CreatedAt || '').localeCompare(String(a.CreatedAt || ''));
  });
  return versions;
}

function deleteAllDataRows_(sheetName) {
  if (ADMIN_READ_MEMO_) delete ADMIN_READ_MEMO_[sheetName];
  const sh = ensureSheet_(sheetName, HEADERS[sheetName]);
  const last = sh.getLastRow();
  if (last <= 1) return;
  // Sheets refuses to delete every non-frozen row in one call. If there is no spare row
  // below the data, add one first so at least one non-frozen row survives the deletion.
  if (sh.getMaxRows() <= last) {
    sh.insertRowAfter(sh.getMaxRows());
  }
  sh.deleteRows(2, last - 1);
  if (sheetAffectsPublicStartup_(sheetName)) invalidatePublicStartupCache_();
}


function removeVersionAssignmentsFromLive_(versionId) {
  // Deletes only this version's rows (bottom-up, in contiguous blocks). Rows of other
  // versions are never cleared or rewritten, so an interrupted call can leave some of
  // this (already archived) version behind, but can never empty the live sheet.
  const id = String(versionId || '');
  if (!id) return 0;
  return deleteRowsWhere_(SHEETS.RotaAssignments, r => String(r.VersionID || '') === id);
}

function restoreOfficialAfterRollback_(version) {
  if (!version || !version.VersionID) return;
  try { restoreVersionAssignmentsToLive_(version.VersionID); } catch (_) {}
  try { updateRow_(SHEETS.RotaVersions, 'VersionID', version.VersionID, { Status: 'Official', AssignmentStorage: 'Live' }); } catch (_) {}
}

function deleteRowsWhere_(sheetName, predicate) {
  if (ADMIN_READ_MEMO_) delete ADMIN_READ_MEMO_[sheetName];
  const sh = ensureSheet_(sheetName, HEADERS[sheetName]);
  const values = sh.getDataRange().getValues();
  perfSheetRead_(values.length > 0 ? values.length - 1 : 0);
  if (values.length < 2) return 0;
  const headers = values[0].map(String);
  const matching = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    headers.forEach((h, idx) => row[h] = normalizeCell_(values[i][idx]));
    if (predicate(row)) matching.push(i + 1);
  }
  if (!matching.length) return 0;
  if (matching.length === values.length - 1 && sh.getMaxRows() <= values.length) {
    sh.insertRowAfter(sh.getMaxRows());
  }
  // Delete bottom-up in contiguous groups to avoid one RPC per assignment.
  for (let end = matching.length - 1; end >= 0;) {
    let begin = end;
    while (begin > 0 && matching[begin - 1] + 1 === matching[begin]) begin--;
    sh.deleteRows(matching[begin], end - begin + 1);
    end = begin - 1;
  }
  perfSheetWrite_(matching.length);
  if (sheetName === SHEETS.RotaAssignments) clearAssignmentStorageSummaryCache_();
  if (sheetAffectsPublicStartup_(sheetName)) invalidatePublicStartupCache_();
  return matching.length;
}

function getLatestConflicts_(knownLatest, knownTeams) {
  const latest = knownLatest || getLatestVersion_();
  if (!latest) return [];
  const teams = knownTeams ? knownTeams.reduce((m,t) => { m[String(t.TeamID)] = t; return m; }, {}) : teamMap_();
  const order = knownTeams ? knownTeams.reduce((m,t,i) => { m[String(t.TeamID)] = i + 1; return m; }, {}) : teamSortOrderMap_();
  return readObjects_(SHEETS.ConflictLog)
    .filter(c => c.VersionID === latest.VersionID)
    .filter(c => !isAdministrativeNonConflict_(c))
    .sort((a, b) => {
      const wk = Number(a.WeekNo || 0) - Number(b.WeekNo || 0);
      if (wk !== 0) return wk;
      return (order[String(a.TeamID)] || 999999) - (order[String(b.TeamID)] || 999999);
    })
    .map(c => Object.assign({}, c, { TeamName: teams[c.TeamID] ? teams[c.TeamID].TeamName : c.TeamID }));
}

function isAdministrativeNonConflict_(c) {
  const rule = String(c.Rule || '').toUpperCase();
  const message = String(c.Message || '').toLowerCase();
  // Off-duty is optional by design. Skipping it because of staffing, settings, or manual seed cleanup
  // is not a rota conflict and should not be displayed on the Admin conflicts panel.
  if (rule.indexOf('OFF') !== -1) return true;
  if (message.indexOf('off-duty skipped') !== -1) return true;
  return false;
}

function normalizeStatus_(status) {
  return String(status == null ? '' : status).trim().toLowerCase();
}

function statusDisplay_(status) {
  const s = String(status == null ? '' : status).trim();
  if (!s) return '';
  const n = s.toLowerCase();
  if (n === 'pending') return 'Pending';
  if (n === 'approved') return 'Approved';
  if (n === 'rejected') return 'Rejected';
  if (n === 'cancelled') return 'Cancelled';
  if (n === 'superseded') return 'Superseded';
  return s;
}

function getLatestVersion_() {
  const activePeriod = getActivePeriod_();
  let versions = readObjects_(SHEETS.RotaVersions).filter(v => v.Status === 'Official');
  if (activePeriod) versions = versions.filter(v => String(v.PeriodID || '') === String(activePeriod.PeriodID));
  if (!versions.length) return null;
  versions.sort((a, b) => {
    const an = Number(a.AmendmentNo || 0), bn = Number(b.AmendmentNo || 0);
    if (bn !== an) return bn - an;
    return String(b.CreatedAt || '').localeCompare(String(a.CreatedAt || ''));
  });
  return enrichVersionWithPeriod_(versions[0]);
}

function getVersionAssignments_(versionId) {
  const live = readObjects_(SHEETS.RotaAssignments).filter(r => String(r.VersionID) === String(versionId));
  if (live.length) return live;
  const version = readObjects_(SHEETS.RotaVersions).find(v => String(v.VersionID) === String(versionId));
  return version ? readArchivedAssignments_(version) : [];
}

function getTeams_() {
  return sortTeams_(readObjects_(SHEETS.Teams));
}

function normalizeDisplayOrder_(value, fallback) {
  const raw = String(value == null || value === '' ? (fallback == null ? '' : fallback) : value).trim();
  if (!raw) return '';
  const n = Number(raw);
  if (!isFinite(n)) return '';
  return String(Math.max(0, Math.round(n)));
}

function teamDisplayOrderValue_(team) {
  const raw = String(team && team.DisplayOrder != null ? team.DisplayOrder : '').trim();
  const n = Number(raw);
  return raw && isFinite(n) ? n : 999999;
}

function teamSortCompare_(a, b) {
  const ao = teamDisplayOrderValue_(a);
  const bo = teamDisplayOrderValue_(b);
  if (ao !== bo) return ao - bo;
  return String((a && a.TeamName) || '').localeCompare(String((b && b.TeamName) || ''));
}

function sortTeams_(teams) {
  return (teams || []).slice().sort(teamSortCompare_);
}

function nextTeamDisplayOrder_() {
  const nums = readObjects_(SHEETS.Teams)
    .map(t => Number(String(t.DisplayOrder || '').trim()))
    .filter(n => isFinite(n));
  if (!nums.length) return '10';
  return String(Math.max.apply(null, nums) + 10);
}

function teamSortOrderMap_() {
  const map = {};
  getTeams_().forEach((t, idx) => map[String(t.TeamID)] = idx + 1);
  return map;
}

function sortRotaTeamViews_(teamViews) {
  const order = teamSortOrderMap_();
  return (teamViews || []).slice().sort((a, b) => {
    const ao = order[String(a.teamId || a.TeamID || '')] || 999999;
    const bo = order[String(b.teamId || b.TeamID || '')] || 999999;
    if (ao !== bo) return ao - bo;
    return String(a.teamName || a.TeamName || '').localeCompare(String(b.teamName || b.TeamName || ''));
  });
}

function getTeamName_(teamId) {
  const team = getTeams_().find(t => t.TeamID === teamId);
  return team ? team.TeamName : teamId;
}

function getConsultants_() {
  return readObjects_(SHEETS.Consultants);
}

function findConsultant_(id) {
  return getConsultants_().find(c => c.ConsultantID === id);
}

function consultantMap_() {
  const map = {};
  getConsultants_().forEach(c => map[c.ConsultantID] = c);
  return map;
}

function teamMap_() {
  const map = {};
  getTeams_().forEach(t => map[t.TeamID] = t);
  return map;
}

function getDb_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this Apps Script project from a Google Sheet so SpreadsheetApp.getActiveSpreadsheet() is available.');
  return ss;
}

function ensureSheet_(name, headers) {
  if (REQUEST_SHEET_MEMO_ && REQUEST_SHEET_MEMO_[name]) return REQUEST_SHEET_MEMO_[name];
  const ss = getDb_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (API_PERF_CONTEXT_) API_PERF_CONTEXT_.schemaChecks++;
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (API_PERF_CONTEXT_) API_PERF_CONTEXT_.sheetWrites++;
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    sh.autoResizeColumns(1, headers.length);
  } else {
    const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    perfSheetRead_(1);
    const missing = headers.filter(h => current.indexOf(h) === -1);
    if (missing.length) {
      sh.getRange(1, sh.getLastColumn() + 1, 1, missing.length).setValues([missing]);
      perfSheetWrite_(1);
      sh.autoResizeColumns(1, sh.getLastColumn());
    }
  }
  if (REQUEST_SHEET_MEMO_) REQUEST_SHEET_MEMO_[name] = sh;
  return sh;
}

function seedSettings_() {
  const existing = readObjects_(SHEETS.Settings).reduce((acc, r) => { acc[r.Key] = true; return acc; }, {});
  const defaults = [
    { Key: 'WeekStartDay', Value: 'Sunday' },
    { Key: 'DefaultWeeks', Value: '52' },
    { Key: 'MaxTeams', Value: '10' },
    { Key: 'OffDutyOptional', Value: 'Yes' },
    { Key: 'PreventConsecutiveTeam', Value: 'Yes' },
    { Key: 'PreventConsecutiveOffDuty', Value: 'Yes' },
    { Key: 'OPDMode', Value: 'Remaining available consultants' },
    { Key: 'OPDMinPerTeam', Value: '1' },
    { Key: 'OPDMaxPerTeam', Value: '9' },
    { Key: 'WeeklyLeaveEffect', Value: 'Any overlap labels the whole week as Leave' },
    { Key: 'Timezone', Value: APP.timezone },
    { Key: 'AdminPIN', Value: '1234' },
    { Key: 'AssignmentArchiveMigrationV18', Value: 'Pending' }
  ];
  defaults.forEach(d => { if (!existing[d.Key]) appendObject_(SHEETS.Settings, d); });
}


function getPublicSettings_() {
  return readObjects_(SHEETS.Settings).filter(r => !['AdminPIN'].includes(String(r.Key)));
}

function getSetting_(key, fallback) {
  const row = readObjects_(SHEETS.Settings).find(r => r.Key === key);
  return row && row.Value !== '' ? String(row.Value) : fallback;
}

function updateSetting_(key, value) {
  const existing = readObjects_(SHEETS.Settings).find(r => r.Key === key);
  if (existing) {
    updateRow_(SHEETS.Settings, 'Key', key, { Value: String(value) });
  } else {
    appendObject_(SHEETS.Settings, { Key: key, Value: String(value) });
  }
}

function getAdminUsersRaw_() {
  return readObjects_(SHEETS.AdminUsers);
}


function getPublicAdminUsers_() {
  return getAdminUsersRaw_()
    .filter(u => String(u.Role || '') !== 'MainAdmin')
    .map(u => ({
      UserID: u.UserID, Username: u.Username, Name: u.Name, Role: u.Role,
      TeamID: u.TeamID, Active: u.Active, CreatedAt: u.CreatedAt,
      UpdatedAt: u.UpdatedAt, LastLoginAt: u.LastLoginAt, Notes: u.Notes
    }));
}

function normalizeUsername_(value) {
  return String(value || '').trim().toLowerCase();
}

function isValidUsername_(value) {
  return /^[a-z0-9._-]{3,40}$/.test(normalizeUsername_(value));
}

function uniqueUsername_(base, used) {
  let seed = normalizeUsername_(base).replace(/[^a-z0-9._-]+/g, '').replace(/^[._-]+|[._-]+$/g, '');
  if (seed.length < 3) seed = 'adminuser';
  seed = seed.slice(0, 32);
  let candidate = seed, n = 2;
  while (used[candidate]) candidate = `${seed.slice(0, 34)}${n++}`;
  used[candidate] = true;
  return candidate;
}

function makePinSalt_() {
  return Utilities.getUuid().replace(/-/g, '');
}

function bytesToHex_(bytes) {
  return bytes.map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join('');
}

function hashPin_(pin, salt) {
  return bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, `${String(salt || '')}:${String(pin || '')}`, Utilities.Charset.UTF_8));
}

function verifyPin_(pin, user) {
  const hash = String((user && user.PINHash) || '');
  const salt = String((user && user.PINSalt) || '');
  return !!hash && !!salt && hashPin_(pin, salt) === hash;
}

function ensureAdminAuthMigration_() {
  ensureSheet_(SHEETS.AdminUsers, HEADERS.AdminUsers);
  const users = getAdminUsersRaw_();
  const used = {};
  users.forEach(u => { const n = normalizeUsername_(u.Username); if (n) used[n] = true; });
  let main = users.find(u => String(u.Role || '') === 'MainAdmin');
  if (!main) {
    const legacyPin = getSetting_('AdminPIN', '1234');
    const salt = makePinSalt_();
    main = {
      UserID: makeId_('ADM'), Username: uniqueUsername_('admin', used), Name: 'Main Admin',
      PIN: '', PINHash: hashPin_(legacyPin, salt), PINSalt: salt,
      AuthVersion: 1, Role: 'MainAdmin', TeamID: 'ALL', Active: 'Yes',
      CreatedAt: now_(), UpdatedAt: now_(), LastLoginAt: '', Notes: 'System main administrator account'
    };
    appendObject_(SHEETS.AdminUsers, main);
  }
  getAdminUsersRaw_().forEach(u => {
    const patch = {};
    let username = normalizeUsername_(u.Username);
    if (!isValidUsername_(username)) {
      username = uniqueUsername_(String(u.Role || '') === 'MainAdmin' ? 'admin' : (u.Name || u.UserID || 'teamadmin'), used);
      patch.Username = username;
    }
    if (!u.PINHash || !u.PINSalt) {
      const legacy = String(u.PIN || '').trim() || (String(u.Role || '') === 'MainAdmin' ? String(getSetting_('AdminPIN', '1234')) : '');
      if (legacy) {
        const salt = makePinSalt_();
        patch.PINHash = hashPin_(legacy, salt); patch.PINSalt = salt;
      }
    }
    if (String(u.PIN || '')) patch.PIN = '';
    if (!Number(u.AuthVersion || 0)) patch.AuthVersion = 1;
    if (Object.keys(patch).length) {
      patch.UpdatedAt = now_();
      updateRow_(SHEETS.AdminUsers, 'UserID', u.UserID, patch);
    }
  });
}

function publicAdminSession_(session) {
  return {
    userId: session.userId || '', username: session.username || '',
    name: session.name || 'Admin', role: session.role || 'MainAdmin',
    teamId: session.teamId || 'ALL', teamName: session.teamName || 'All teams',
    isMainAdmin: !!session.isMainAdmin, canApproveLeave: !!session.canApproveLeave,
    canManageSettings: !!session.canManageSettings, canManualEdit: !!session.canManualEdit,
    canGenerateRota: !!session.canGenerateRota
  };
}

function sessionCacheKey_(token) {
  return 'admSess_' + hashPin_(String(token || ''), 'session').slice(0, 40);
}

function sessionPropertyKey_(token) {
  return 'ADMIN_SESSION_' + hashPin_(String(token || ''), 'persistent-session').slice(0, 48);
}

function removeAdminSession_(token) {
  if (!token) return;
  CacheService.getScriptCache().remove(sessionCacheKey_(token));
  PropertiesService.getScriptProperties().deleteProperty(sessionPropertyKey_(token));
}

function cleanupExpiredAdminSessions_() {
  // Expiry is checked on every authenticated request. Garbage-collecting ALL
  // Script Properties on every login adds unnecessary synchronous I/O.
  const cache = CacheService.getScriptCache();
  const marker = 'adminSessionCleanup_v3_7';
  if (cache.get(marker) === '1') return;
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const now = Date.now();
  Object.keys(all).filter(k => k.indexOf('ADMIN_SESSION_') === 0).forEach(k => {
    try {
      const payload = JSON.parse(all[k]);
      if (!payload.expiresAtEpoch || Number(payload.expiresAtEpoch) <= now) props.deleteProperty(k);
    } catch (_) { props.deleteProperty(k); }
  });
  cache.put(marker, '1', 900);
}


function createAdminSession_(user) {
  cleanupExpiredAdminSessions_();
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const nowEpoch = Date.now();
  const payload = {
    userId: String(user.UserID || ''),
    authVersion: Number(user.AuthVersion || 1),
    issuedAtEpoch: nowEpoch,
    lastPersistedAtEpoch: nowEpoch,
    lastSeenAtEpoch: nowEpoch,
    expiresAtEpoch: nowEpoch + (21600 * 1000)
  };
  const json = JSON.stringify(payload);
  CacheService.getScriptCache().put(sessionCacheKey_(token), json, 21600);
  PropertiesService.getScriptProperties().setProperty(sessionPropertyKey_(token), json);
  return token;
}

function adminAttemptKey_(username) {
  const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(username));
  return 'loginAttempts_'+Utilities.base64EncodeWebSafe(digest).slice(0,28);
}
function checkLoginThrottle_(store,key) {
  const v=JSON.parse(store.getProperty(key)||'{}'),now=Date.now();
  if(v.blockedUntil && now<v.blockedUntil)
    throw new Error('Too many unsuccessful attempts. Try again in 15 minutes.');
}
function recordLoginFailure_(store,key) {
  const now=Date.now();let v;
  try{v=JSON.parse(store.getProperty(key)||'{}');}catch(_){v={};}
  const count=(now-Number(v.startedAt||0)<15*60*1000)?Number(v.count||0)+1:1;
  store.setProperty(key,JSON.stringify({count:count,startedAt:count===1?now:v.startedAt,
    blockedUntil:count>=5?now+15*60*1000:0}));
}

function loginAdmin(username, pin) {
  setupDatabase();
  const uname = normalizeUsername_(username);
  const suppliedPin = String(pin || '').trim();
  if (!uname || !suppliedPin) throw new Error('Enter username and PIN.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const attempts=PropertiesService.getScriptProperties(), attemptKey=adminAttemptKey_(uname);
    checkLoginThrottle_(attempts,attemptKey);
    const user = getAdminUsersRaw_().find(u => normalizeYesNo_(u.Active, 'Yes') === 'Yes' && normalizeUsername_(u.Username) === uname);
    if (!user || !verifyPin_(suppliedPin, user)) {
      recordLoginFailure_(attempts,attemptKey);
      throw new Error('Invalid username or PIN.');
    }
    attempts.deleteProperty(attemptKey);
    const session = sessionFromUser_(user);
    const token = createAdminSession_(user);
    updateRow_(SHEETS.AdminUsers, 'UserID', user.UserID, { LastLoginAt: now_() });
    appendAudit_(session, 'ADMIN_LOGIN', 'AdminUser', user.UserID, { username: uname, role: user.Role, teamId: user.TeamID || '' });
    return { ok: true, sessionToken: token, adminSession: session, expiresInMinutes: 360 };
  } finally {
    lock.releaseLock();
  }
}

function logoutAdminSession(token) {
  let session = null;
  try { session = authenticateAdmin_(token); } catch (_) {}
  removeAdminSession_(token);
  if (!session) return { ok: true, message: 'Signed out.' };
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    appendAudit_(session, 'ADMIN_LOGOUT', 'AdminUser', session.userId, { username: session.username });
    return { ok: true, message: 'Signed out.' };
  } finally {
    lock.releaseLock();
  }
}

function sessionFromUser_(user) {
  const role = String(user.Role || 'TeamAdmin').trim() || 'TeamAdmin';
  if (role === 'MainAdmin') {
    return publicAdminSession_({ userId: user.UserID, username: user.Username, name: user.Name || 'Main Admin', role: 'MainAdmin', teamId: 'ALL', teamName: 'All teams', isMainAdmin: true, canApproveLeave: true, canManageSettings: true, canManualEdit: true, canGenerateRota: true });
  }
  if (role !== 'TeamAdmin') throw new Error('This admin role is not supported.');
  const team = getTeams_().find(t => String(t.TeamID) === String(user.TeamID || ''));
  if (!team) throw new Error('Sub-admin team was not found. Contact the main admin.');
  return publicAdminSession_({ userId: user.UserID, username: user.Username, name: user.Name || 'Team Sub-admin', role: 'TeamAdmin', teamId: user.TeamID, teamName: team.TeamName || user.TeamID, isMainAdmin: false, canApproveLeave: true, canManageSettings: false, canManualEdit: false, canGenerateRota: false });
}

function authenticateAdmin_(sessionToken) {
  const token = String(sessionToken || '').trim();
  if (!token) throw new Error('Admin session expired. Please sign in again.');
  const cache = CacheService.getScriptCache();
  const properties = PropertiesService.getScriptProperties();
  const cacheKey = sessionCacheKey_(token);
  const propertyKey = sessionPropertyKey_(token);
  let raw = cache.get(cacheKey) || properties.getProperty(propertyKey);
  if (!raw) throw new Error('Admin session expired. Please sign in again.');
  let payload;
  try { payload = JSON.parse(raw); }
  catch (_) { removeAdminSession_(token); throw new Error('Admin session expired. Please sign in again.'); }
  if (!payload.expiresAtEpoch || Number(payload.expiresAtEpoch) <= Date.now()) {
    removeAdminSession_(token);
    throw new Error('Admin session expired. Please sign in again.');
  }
  const user = getAdminUsersRaw_().find(u => String(u.UserID) === String(payload.userId) && normalizeYesNo_(u.Active, 'Yes') === 'Yes');
  if (!user || Number(user.AuthVersion || 1) !== Number(payload.authVersion || 0)) {
    removeAdminSession_(token);
    throw new Error('Admin session expired. Please sign in again.');
  }
  const nowEpoch = Date.now();
  payload.lastSeenAtEpoch = nowEpoch;
  payload.expiresAtEpoch = nowEpoch + (21600 * 1000);
  // Cache is refreshed on every request. Persist the durable fallback only
  // once every three minutes to avoid a slow Properties write on every click.
  const persist=nowEpoch-Number(payload.lastPersistedAtEpoch||payload.issuedAtEpoch||0)>=180000;
  if(persist) payload.lastPersistedAtEpoch=nowEpoch;
  raw=JSON.stringify(payload);
  cache.put(cacheKey,raw,21600);
  if(persist) properties.setProperty(propertyKey,raw);
  return sessionFromUser_(user);
}

function requireAnyAdmin_(sessionToken) {
  return authenticateAdmin_(sessionToken);
}

// Existing frontend restores its session through this endpoint after a reload.
// This only verifies the token; it does not change administrator state.
function getAdminSession(sessionToken) {
  return requireAnyAdmin_(sessionToken);
}

// Name retained for compatibility with existing server functions; argument is now a session token.
function requireAdminPin_(sessionToken) {
  const session = authenticateAdmin_(sessionToken);
  if (!session.isMainAdmin) throw new Error('Main admin permission is required for this action.');
  return session;
}

function requireLeaveAdmin_(sessionToken, leave) {
  const session = authenticateAdmin_(sessionToken);
  if (session.isMainAdmin) return session;
  if (!session.canApproveLeave) throw new Error('You do not have permission to manage leave requests.');
  const leaveTeamId = String((leave && leave.TeamID) || '');
  if (!leaveTeamId || leaveTeamId !== String(session.teamId)) throw new Error('You can only manage leave requests for your assigned team.');
  return session;
}


function auditDetailsText_(details) {
  if (details === undefined || details === null) return '';
  let text;
  if (typeof details === 'string') text = details;
  else {
    try { text = JSON.stringify(details); }
    catch (err) { text = String(details); }
  }
  return String(text || '').slice(0, 1500);
}

function appendAudit_(session, action, targetType, targetId, details) {
  try {
    const actor = session || {};
    appendObject_(SHEETS.AuditLog, {
      AuditID: makeId_('AUD'),
      Timestamp: now_(),
      ActorName: actor.name || 'System',
      ActorRole: actor.role || (actor.isMainAdmin ? 'MainAdmin' : 'System'),
      ActorTeamID: actor.teamId || '',
      Action: String(action || ''),
      TargetType: String(targetType || ''),
      TargetID: String(targetId || ''),
      Details: auditDetailsText_(details)
    });
  } catch (err) {
    console.error('Audit log write failed: ' + ((err && err.message) || String(err)));
  }
}

function getAuditLog(adminPin, limit) {
  requireAdminPin_(adminPin);
  setupDatabase();
  const maxRows = Math.min(500, Math.max(1, Number(limit || 100)));
  return readObjects_(SHEETS.AuditLog)
    .sort((a, b) => String(b.Timestamp || '').localeCompare(String(a.Timestamp || '')))
    .slice(0, maxRows);
}

function readObjects_(sheetName) {
  if (ADMIN_READ_MEMO_ && Object.prototype.hasOwnProperty.call(ADMIN_READ_MEMO_, sheetName)) {
    return ADMIN_READ_MEMO_[sheetName];
  }
  const sh = ensureSheet_(sheetName, HEADERS[sheetName]);
  const values = sh.getDataRange().getValues();
  perfSheetRead_(values.length > 0 ? values.length - 1 : 0);
  let result = [];
  if (values.length >= 2) {
    const headers = values[0].map(String);
    result = values.slice(1)
      .filter(row => row.some(v => v !== '' && v !== null))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = normalizeCell_(row[i]));
        return obj;
      });
  }
  if (ADMIN_READ_MEMO_) ADMIN_READ_MEMO_[sheetName] = result;
  return result;
}

function appendObject_(sheetName, obj) {
  appendRows_(sheetName, [obj]);
}

function appendRows_(sheetName, objects) {
  if (!objects || !objects.length) return;
  if (ADMIN_READ_MEMO_) delete ADMIN_READ_MEMO_[sheetName];
  let prepared = objects;
  if (sheetName === SHEETS.RotaAssignments) {
    const versionIds = {};
    objects.forEach(o => { if (o && o.VersionID) versionIds[String(o.VersionID)] = true; });
    const periodMap = {};
    readObjects_(SHEETS.RotaVersions).forEach(v => {
      if (versionIds[String(v.VersionID)]) periodMap[String(v.VersionID)] = String(v.PeriodID || '');
    });
    prepared = objects.map(o => Object.assign({}, o, {
      PeriodID: String(o.PeriodID || periodMap[String(o.VersionID)] || '')
    }));
  }
  const sh = ensureSheet_(sheetName, HEADERS[sheetName]);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  perfSheetRead_(1);
  const rows = prepared.map(obj => headers.map(h => obj[h] === undefined ? '' : obj[h]));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  perfSheetWrite_(rows.length);
  if (sheetAffectsPublicStartup_(sheetName)) invalidatePublicStartupCache_();
  if (sheetName === SHEETS.RotaAssignments) {
    clearAssignmentStorageSummaryCache_();
    const counts = {};
    prepared.forEach(r => { const id = String(r.VersionID || ''); if (id) counts[id] = (counts[id] || 0) + 1; });
    Object.keys(counts).forEach(id => {
      try {
        const current = readObjects_(SHEETS.RotaVersions).find(v => String(v.VersionID) === id);
        updateRow_(SHEETS.RotaVersions, 'VersionID', id, { AssignmentStorage: 'Live', AssignmentCount: Number((current && current.AssignmentCount) || 0) + counts[id] });
      } catch (_) {}
    });
  }
}

function updateRow_(sheetName, idHeader, idValue, patch) {
  const sh = ensureSheet_(sheetName, HEADERS[sheetName]);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error(`No rows in ${sheetName}.`);
  const lastCol = sh.getLastColumn();
  let targetRow, headers, rowValues;
  if (lastRow > 150) {
    // For large sheets, Google Sheets searches the identifier column on the server
    // instead of transferring every column/row to Apps Script for one-row updates.
    headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    perfSheetRead_(1);
    const idCol = headers.indexOf(idHeader);
    if (idCol < 0) throw new Error(`${idHeader} not found in ${sheetName}.`);
    const match = sh.getRange(2, idCol + 1, lastRow - 1, 1)
      .createTextFinder(String(idValue)).useRegularExpression(false)
      .matchEntireCell(true).matchCase(true).findNext();
    if (API_PERF_CONTEXT_) API_PERF_CONTEXT_.rowFinds++;
    if (!match) throw new Error(`Row ${idValue} not found in ${sheetName}.`);
    targetRow = match.getRow();
    rowValues = sh.getRange(targetRow, 1, 1, headers.length).getValues()[0];
    perfSheetRead_(1);
    // TextFinder may encounter a string-representation mismatch for numeric keys.
    if (String(rowValues[idCol]) !== String(idValue)) throw new Error(`Row ${idValue} not found in ${sheetName}.`);
  } else {
    // On small sheets one bulk read is generally cheaper than multiple RPCs.
    const values = sh.getDataRange().getValues();
    perfSheetRead_(values.length > 0 ? values.length - 1 : 0);
    headers = values[0].map(String);
    const idCol = headers.indexOf(idHeader);
    if (idCol < 0) throw new Error(`${idHeader} not found in ${sheetName}.`);
    const idx = values.findIndex((row, index) => index > 0 && String(row[idCol]) === String(idValue));
    if (idx < 0) throw new Error(`Row ${idValue} not found in ${sheetName}.`);
    targetRow = idx + 1;
    rowValues = values[idx].slice();
  }
  Object.keys(patch).forEach(key => {
    const col = headers.indexOf(key);
    if (col !== -1) rowValues[col] = patch[key];
  });
  sh.getRange(targetRow, 1, 1, headers.length).setValues([rowValues]);
  perfSheetWrite_(1);
  // An authenticated admin read memo may exist within this request. A write
  // invalidates that sheet snapshot so follow-on reads see the latest data.
  if (ADMIN_READ_MEMO_) delete ADMIN_READ_MEMO_[sheetName];
  if (sheetAffectsPublicStartup_(sheetName)) invalidatePublicStartupCache_();
}

function normalizeHexColor_(color) {
  let value = String(color || '').trim();
  if (!value) return '#64748b';
  if (value[0] !== '#') value = '#' + value;
  if (/^#[0-9a-fA-F]{3}$/.test(value)) {
    value = '#' + value.slice(1).split('').map(ch => ch + ch).join('');
  }
  if (!/^#[0-9a-fA-F]{6}$/.test(value)) throw new Error(`Invalid color code: ${color}. Use a hex color like #2563eb.`);
  return value.toUpperCase();
}

function contrastTextColor_(hex) {
  const value = normalizeHexColor_(hex).replace('#', '');
  const r = parseInt(value.slice(0, 2), 16);
  const g = parseInt(value.slice(2, 4), 16);
  const b = parseInt(value.slice(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b);
  return luminance > 160 ? '#111827' : '#ffffff';
}

function makeId_(prefix) {
  return `${prefix}-${Utilities.getUuid().slice(0, 8).toUpperCase()}`;
}

function now_() {
  return Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd HH:mm:ss');
}

function toDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  if (!value) throw new Error('Date value is required.');
  const s = String(value).slice(0, 10);
  const parts = s.split('-').map(Number);
  if (parts.length === 3 && parts[0] && parts[1] && parts[2]) return new Date(parts[0], parts[1] - 1, parts[2]);
  const d = new Date(value);
  if (isNaN(d.getTime())) throw new Error(`Invalid date: ${value}`);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function iso_(dateValue) {
  return Utilities.formatDate(toDate_(dateValue), APP.timezone, 'yyyy-MM-dd');
}

function nice_(dateValue) {
  return Utilities.formatDate(toDate_(dateValue), APP.timezone, 'EEE dd MMM yyyy');
}

function dateOnlyNice_(dateValue) {
  return Utilities.formatDate(toDate_(dateValue), APP.timezone, 'dd MMM yyyy');
}

function datesOnlyPeriodFromNicePeriod_(period) {
  // Converts "Sun 26 Jul 2026 – Sat 01 Aug 2026" to "26 Jul 2026 – 01 Aug 2026" for export.
  return String(period || '').replace(/\b(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+/g, '');
}

function addDays_(dateValue, days) {
  const d = toDate_(dateValue);
  d.setDate(d.getDate() + Number(days));
  return d;
}

function startOfWeekSunday_(dateValue) {
  const d = toDate_(dateValue);
  d.setDate(d.getDate() - d.getDay());
  return d;
}

function datesOverlap_(aStart, aEnd, bStart, bEnd) {
  return toDate_(aStart) <= toDate_(bEnd) && toDate_(bStart) <= toDate_(aEnd);
}

function getWeekNoForDate_(rotaStart, dateValue, weeks) {
  const start = toDate_(rotaStart);
  const weekStart = startOfWeekSunday_(dateValue);
  const diffMs = weekStart.getTime() - start.getTime();
  const diffWeeks = Math.floor(diffMs / (7 * 24 * 60 * 60 * 1000));
  return diffWeeks + 1;
}


function offDutyCarryOverWeeks_(previousRows, nextRotaStart) {
  const limits = {};
  const start = startOfWeekSunday_(toDate_(nextRotaStart));
  (previousRows || []).forEach(r => {
    if (String(r.AssignmentPlace) !== 'Off-duty' || isManualCoverId_(r.ConsultantID)) return;
    const previousStart = startOfWeekSunday_(toDate_(r.WeekStart));
    const weeksSincePrevious = Math.round((start.getTime() - previousStart.getTime()) / (7*24*60*60*1000));
    if (weeksSincePrevious > 0 && weeksSincePrevious < 3) {
      const id = String(r.ConsultantID);
      limits[id] = Math.max(Number(limits[id] || 0), 3 - weeksSincePrevious);
    }
  });
  return limits;
}

function isOffDutyIntervalSafe_(consultant, state, weekNo) {
  // New rota periods must not restart the three-week Off-duty clock.
  if (Number(weekNo || 0) <= Number(state && state.offCooldownWeeks && state.offCooldownWeeks[consultant.ConsultantID] || 0)) return false;
  const lastOffWeek = Number(state && state.lastOffWeek && state.lastOffWeek[consultant.ConsultantID] || 0);
  if (!lastOffWeek) return true;
  return Number(weekNo || 0) - lastOffWeek >= 3;
}

// Historical ConsultantType remains in the sheet for read compatibility only.
// All old consultant types, including retired legacy labels, are now treated
// as regular rotation members. No special scheduling cadence exists.
function normalizeConsultantType_(value) { return 'Regular'; }

function normalizeCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return iso_(value);
  return value;
}

function normalizeYesNo_(value, defaultValue) {
  if (value === true) return 'Yes';
  if (value === false) return 'No';
  const v = String(value || defaultValue || 'Yes').trim().toLowerCase();
  return ['no', 'n', 'false', '0'].indexOf(v) !== -1 ? 'No' : 'Yes';
}
