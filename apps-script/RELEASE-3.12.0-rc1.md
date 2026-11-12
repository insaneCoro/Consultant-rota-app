# Backend 3.12.0-rc1 (based on live 3.11.0-rc22)

Backend-only release. No frontend change, no sheet/column change, no data migration.

## Changes

1. **Live assignments can no longer be emptied by archiving.** Archiving an old
   amendment (on every leave approval, Manual Edit, publish) and the one-time
   "Archive Existing History" now delete only the archived version's rows. The previous
   fallback cleared the whole `RotaAssignments` sheet and rewrote it; a failure
   in between could leave the live rota empty. The now-unused
   `rewriteSheetObjects_` helper was removed.
2. **Archive links are never silently replaced.** If a period's archive
   spreadsheet cannot be opened, the operation stops and rolls back instead of
   creating a new archive and overwriting `RotaPeriods.ArchiveSpreadsheetID`
   (which made all earlier archived amendments of that period unreachable).
3. **Interrupted archive moves are repaired on rollback.** If a version was only
   partly removed from the live sheet, restoring it now rebuilds it from the
   complete archive copy instead of accepting the partial rows.
4. **"Clear Generated Rota" is disabled unless explicitly enabled.** It works
   only when the Script Property `ALLOW_ROTA_RESET` is `YES` (intended for
   staging). Production: do not set it.
5. **TCU/EMCU alternation is a weighted preference (option C).** Inpatient
   coverage, avoiding consecutive inpatient weeks and the return-from-leave
   preference still come first. Alternation is now part of the fairness cost
   (`SUBTEAM_ALTERNATION_WEIGHT_ = 120`), so a clearly overdue consultant is
   no longer passed over only to keep alternation. Ties still prefer alternation.
6. Version `3.12.0-rc1`; public startup cache key renamed so the new version
   shows immediately.

## Effect on existing rota data

- Published, archived and completed weeks are not touched.
- Change 5 applies only when future weeks are next generated (leave approval,
  Manual Edit, staffing draft, rearrangement) for that team. Those future weeks
  may then differ slightly more from the current plan than they would under rc22.

## Test results (offline, local only)

`TZ=Asia/Riyadh node apps-script/tests/run.js` uses an in-memory Sheets/Drive mock:
8/8 pass on 3.12.0-rc1. On unchanged rc22, 4/8 fail (partial-failure data loss,
archive re-link, partial restore, unguarded reset).

`TZ=Asia/Riyadh node apps-script/tests/scheduling-sim.js` runs 52-week simulations
(rc22 → rc1; spread = max−min inpatient share of eligible weeks):

| Scenario | Spread | Same-service repeats | Consecutive inpatient | Uncovered |
|---|---|---|---|---|
| 4 consultants, no leave | 0.0 → 0.0 | 0 → 0 | 0 → 0 | 0 → 0 |
| 5 consultants, no leave | 1.9 → 1.9 | 0 → 0 | 0 → 0 | 0 → 0 |
| 4 consultants, regular leave | 4.1 → 1.0 | 0 → 2 | 10 → 7 | 0 → 0 |
| 5, one cannot take Off | 2.3 → 2.0 | 1 → 2 | 0 → 0 | 0 → 0 |
| 5, joiner + leaver | 2.6 → 2.6 | 0 → 0 | 0 → 0 | 0 → 0 |
| 5, one fixed to TCU | 5.8 → 5.8 | 17 → 17 | 0 → 0 | 0 → 0 |

Not yet staging-tested. Not deployed to production.

## Staging test (before production)

1. Make a copy of the production spreadsheet (File → Make a copy). This also copies the bound script.
2. **In the copy, before any write:** clear `ArchiveSpreadsheetID` and
   `ArchiveSpreadsheetURL` in `RotaPeriods`, and `CurrentExportDocId` /
   `CurrentExportDocUrl` in `Settings`. Otherwise staging writes into the
   production archive and export files. Archived amendments will then be
   unreadable in staging; that is expected.
3. In the copy's Apps Script editor, paste this `Code.gs`, then deploy it as its own
   web app (Execute as: Me, Access: Anyone).
4. Point a local copy of `index.html` at the staging `/exec` URL (`API_URL`), and test:
   - leave approval
   - amend approved leave
   - cancel leave
   - Manual Edit and publish
   - staffing draft and publish
   - void/restore
   - Archive Existing History
   - the Danger Zone button, which must refuse
5. Compare the staging rota's future weeks with production's for reasonableness.

## Production deployment (only after approval)

1. Back up the production spreadsheet (File → Make a copy, keep it untouched).
2. Apps Script → paste `Code.gs` → Save.
3. Deploy → Manage deployments → edit the existing web app → Version: **New
   version** → Deploy. The `/exec` URL stays the same; no frontend change needed.
4. Check that the site footer shows `Backend 3.12.0-rc1`.

## Rollback

Deploy → Manage deployments → edit → select the previous version (rc22) →
Deploy. No data conversion is involved, so rolling back is immediate and safe.
