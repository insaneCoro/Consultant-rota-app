# Release 3.12.0-rc1 (based on live 3.11.0-rc22)

Backend (`apps-script/Code.gs`) and website (`index.html`). No sheet or column
changes and no data migration. **Scheduling, fairness, alternation and Off-duty
logic are unchanged**: for the real database, rebuilding the published plan gives
an identical result.

## What is new

### 1. What-if Leave Check (Leave Approval tab)
Pick a consultant and dates and press **Check** to see how the rota would look
if that leave were approved. It is a preview only: no leave request is created,
nothing is written to any sheet, and nothing is published.

It shows a **staffing outlook** in plain language:
- which weeks drop to 3 consultants, or to only 2 (everyone inpatient, no OPD, no Off-duty reserve)
- who would work back-to-back inpatient weeks
- weeks where OPD is empty and the Off-duty consultant is the reserve
- any inpatient post that would be uncovered
- when the consultant returns and what they are scheduled for

Below that is the rota week by week, from the first changed week through the return week (at most 12 weeks).

The same outlook now also appears when you press **Preview** on a real pending
leave request.

Backend: new read-only API `previewLeaveWhatIf`; `previewLeaveImpact` returns an
extra `outlook` field (existing fields unchanged).

### 2. Storage safety (no visible change in daily use)
- Archiving an old amendment now deletes only that amendment's rows. The old
  fallback that cleared and rewrote the whole `RotaAssignments` sheet is removed,
  so a failure part-way can no longer empty the live rota.
- If a period's archive spreadsheet cannot be opened, the operation stops and
  rolls back. Previously a new archive was created and the stored archive ID was
  overwritten, cutting off all earlier archived amendments.
- If an archive move is interrupted, rollback rebuilds the version from its
  complete archive copy.
- **Clear Generated Rota** is refused unless the Script Property
  `ALLOW_ROTA_RESET` is `YES`. Leave it unset in production.

### 3. Faster: far fewer spreadsheet round-trips
Each read or write of the Google Sheet is a slow round-trip. The backend now:
- reads each sheet at most once per request between writes (a per-request cache
  that is cleared on every write and whenever the database lock is taken)
- skips the per-request column check while the setup step's "verified" flag is
  valid (Safe Repair still forces a full check)
- keeps that flag for 6 hours instead of 1
- no longer rewrites the rota period's activation time on every setup check

Sheet reads per action, measured offline on a copy of the real database:

| Action | rc22 | 3.12 |
|---|---|---|
| Admin dashboard data | 14 | 7 |
| Leave approval list | 8 | 4 |
| Manual Edit: load a week | 14 | 6 |
| Preview a leave | 23 | 10 |
| Approve a leave (publishes) | 49 | 24 |
| Amend approved leave | 51 | 25 |
| Login (normal) | 5 | 3 |
| Login (first after 6 h) | 25 | 20 |

Writes are unchanged. Real seconds saved can only be measured on Google itself
(Apps Script → Executions shows `durationMs` per call); each call also has a
fixed Apps Script overhead that this cannot remove.

Results are unchanged: the same 34-call sequence (login, previews, approve,
amend, cancel, draft, publish, Manual Edit, inactivity, restore, void) gives
identical API responses and an identical final database with the cache on and
off, and an identical final database compared with rc22.

## Compatibility

- New backend + old website: works as today (the extra data is ignored).
- New website + old backend: everything works except the What-if card, which
  says the backend needs updating.
- So deploy the **backend first**, then the website.

## Test results

Offline only so far (not staging-tested, not deployed):

- `TZ=Asia/Riyadh node apps-script/tests/run.js`: 18/18 pass, using an
  in-memory Sheets/Drive mock with synthetic data. Includes a check that the
  what-if preview equals what a real approval then publishes, and that what-if
  writes nothing.
- On unchanged rc22, the 4 storage-risk tests fail (as expected).
- Against an offline copy of the real database: the published plan is reproduced
  exactly (28 of 28 future weeks), and the website was driven in a real browser
  (desktop and phone size) through login, Leave Approval and What-if with no page errors.

## Staging test (recommended before production)

1. Copy the production spreadsheet (File → Make a copy). The bound script is copied too.
2. **In the copy, before any write:** clear `ArchiveSpreadsheetID` and
   `ArchiveSpreadsheetURL` in `RotaPeriods`, and `CurrentExportDocId` /
   `CurrentExportDocUrl` in `Settings`. Otherwise the copy writes into the
   production archive and export files. (Older archived amendments will then be
   unreadable in the copy; that is expected.)
3. In the copy's Apps Script editor, paste `Code.gs`, then Deploy → New deployment →
   Web app (Execute as: Me, Access: Anyone).
4. Open a local copy of `index.html` with `API_URL` set to the staging `/exec` URL.
5. Try:
   - What-if Leave Check
   - Preview on a pending leave
   - approve, amend and cancel a leave
   - Manual Edit and publish
   - the Danger Zone button, which must refuse

## Production deployment (only after approval)

1. Back up the production spreadsheet (File → Make a copy; keep it untouched).
2. Apps Script → paste `Code.gs` → Save → Deploy → Manage deployments → edit the
   existing web app → Version: **New version** → Deploy. The `/exec` URL is unchanged.
3. Check the site footer shows `Backend 3.12.0-rc1`.
4. Merge the website change to `main` (GitHub Pages publishes it).

## Rollback

- Backend: Deploy → Manage deployments → edit → select the previous version → Deploy.
- Website: revert the `index.html` commit on `main`.
- No data conversion is involved, so both are immediate.
