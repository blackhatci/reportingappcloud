# Handoff: moving the payroll-initiation work from cloud to local Claude Code

## Why this handoff exists

The cloud/web Claude Code session (Claude Code on the web) runs inside a sandboxed
container whose outbound traffic goes through a policy-enforcing proxy. That proxy
blocks `api.knack.com` (confirmed via repeated 403s at the CONNECT level — not a
credentials or connection-string problem, the request never reaches Knack). All
Knack reads/writes from the cloud session had to be routed through Zapier's
`knack_make_api_get_request` / Knack connector actions instead, since Zapier's own
servers (not the sandboxed container) make the actual call to Knack.

A local Claude Code session has no such proxy, so direct calls to `api.knack.com`
should just work — the same way the workbook's own VBA (`KnackAPI.bas`) already
does from the user's machine.

## What's already done (all committed on this branch)

Branch: `claude/youthful-newton-gmdvoi`, up to date with `origin`.

1. **New Pull modules** (`knack-integration/Pull*.bas`) bring additional Knack
   tables into Excel sheets: Monthly Targets, CarryOver, AFDelinquency, Locations,
   Paid Holidays, plus a fix to the pre-existing `PullWages.bas`
   (`PullWages_FIXED.bas`). Each follows a header-driven column→Knack-field
   mapping pattern that reads the sheet's own header row at runtime and preserves
   whatever column order already exists.
2. **New Upload modules** (`knack-integration/Upload*.bas` +
   `KnackUploadHelpers.bas`) replace three Zapier-webhook buttons (Upload CO
   table, Upload Payroll, Upload Next Months Targets) with direct Knack API calls
   from VBA. See `UPLOADS_SETUP.md` for the full field-mapping confirmation
   method and the list of intentionally-unmapped Payroll columns.
3. **A full "payroll initiation" run was executed end-to-end from the cloud
   session**, via Zapier (OneDrive + Excel Graph API + Knack), for the pay period
   ending **9/30/2026**:
   - Knack App ID / API key (already in `KnackAPI.bas`'s `KNACK_APP_ID`/
     `KNACK_API_KEY` constants, used operationally per explicit user
     instruction — not rotated yet; App ID is `67a3de417009da15fb223d49`, the
     API key is deliberately NOT repeated here — pull it from `KnackAPI.bas`
     directly, never from a committed doc):
   - Copied `Templates/AF Reporting Book.xlsm` → `Payroll Reports/AF Reporting
     Book Completed Run for Pay Period 09302026.xlsm`, then moved to
     `Payroll Reports/Archive/` once complete.
   - Set `VariablesSheet!A2` = 9/30/2026 (Current period **End** of Pay Period —
     see important correction below), `B2` = 9/15/2026 (Prior period End of Pay
     Period), `G2` = TRUE (fiscal quarter-end flag).
   - Pulled and wrote: Hours (274 records), Leads (266, **see open issue below**),
     Area Location (28), Employee Wage Codes (35 of 41 active-rate records — 6
     dropped for inactive Local Manager), AF Delinquency (22), Monthly Targets
     (24, in scope because this was a month-end run), and — in a follow-up fix —
     OT Carryover (28, filtered to the Prior period 9/15/2026, which had
     initially been skipped and then had to be backfilled).
   - Updated the `Monthly Summary` sheet's G/H/I reconciliation columns (label /
     Knack count / sheet-row-count formula) for all 7 tables now pulled.

4. **Important correction made mid-run**: `VariablesSheet!A2`/`B2` must hold the
   **End** of Pay Period date for Current/Prior, not the **Beginning** of Pay
   Period as originally assumed from the user's plain-English instructions. This
   was discovered empirically — with Beginning-of-period dates, the Hours pull
   returned zero records, because the Hours table's connected "End of Pay
   Period" field (`field_177`) and the Knack-side `Report_Day` date-range filter
   both only make sense against End-of-period dates. Confirmed against the
   workbook's own `EndOfMonth` formula (`=IF(DAY(A2)>27,...)`) agreeing with
   Knack's own EndOfMonth flag only when `A2` is the End date. **Use End of Pay
   Period for both A2 and B2 going forward.**

## Open issue to resolve first, now that direct API access should work

**Leads count discrepancy**: the cloud run (via Zapier) produced **266** matching
Leads records. The user later queried Knack directly (via Excel/VBA) and saw
**267**. This was never root-caused — the Zapier session was interrupted mid
re-investigation when the user asked to stop using Zapier for data pulls. A
re-pull of the raw candidate count during that investigation showed Knack's
`object_9` returning **268** total raw records matching the Knack-side filter
(up from an earlier 267/266 at different pull times), which suggests this
object's record set is actively changing over time (new leads being entered or
converted) rather than a filtering-logic bug — but this is NOT confirmed, just
the leading theory. **First task in the local session: re-run the Leads pull
with a timestamp-pinned snapshot, compare record-by-record against whatever the
user's own direct Knack query returns, and identify whether the diff is (a) a
timing/data-drift issue, (b) an off-by-one in the date-boundary filters, or (c)
a connected-period-matching edge case** (see exact filter logic below).

### Leads pull filter logic (from `PullLeadsRecords.bas`, already verified once)

- Knack object: `object_9`
- Knack-side filters: `field_470` (Marked for Delete) is not "Yes"; `field_54`
  (Date Converted) is after `09/15/2026`; `field_54` is before `10/01/2026`.
- VBA-side filter (connected-period match, can't be done server-side): keep only
  records where `field_178` ("Pay Period Lead Was Converted" connection)
  resolves via `ProcessConnRecords.GetConnValue` to exactly `09/30/2026`.
- 21-column header→field mapping lives in the workbook's "Knack Field Mappings"
  sheet, rows 8-9 (not hardcoded in VBA) — reproduced in the cloud session's
  scratch Python helper (not committed; see below) as:
  ```
  headers = ['Id','Location','Date Lead Entered','Early Transfer','Lead Sources','Lead Name','Lead Status','Days Outstanding','Date Converted','Pay Period Lead Was Converted','Billing Method','Bonus','ConvCode','Local Manager','Area','Billing Code','Searchable Pay Period Lead Converted','Transfer Bonus','LeadMonthCode','MgrCode','Delete']
  field_keys = ['field_52','field_220','field_59','field_289','field_58','field_55','field_61','field_186','field_54','field_178','field_57','field_180','field_60','field_223','field_226','field_361','field_372','field_405','field_407','field_866','field_470']
  ```

## Setting up the local session

1. `export KNACK_APP_ID="67a3de417009da15fb223d49"` and `export
   KNACK_API_KEY="<copy the value of KNACK_API_KEY from KnackAPI.bas>"` —
   deliberately not repeated here so it never ends up duplicated in a second
   committed file (consider rotating this key at some point; it's been sitting
   in plaintext in the VBA project for a while — not urgent, user's call).
2. Verify direct access actually works before relying on it:
   ```bash
   curl -s "https://api.knack.com/v1/objects/object_9/records?rows_per_page=1" \
     -H "X-Knack-Application-Id: $KNACK_APP_ID" \
     -H "X-Knack-REST-API-Key: $KNACK_API_KEY"
   ```
3. OneDrive access: no connector is set up locally yet. Two options discussed
   with the user — simplest is pointing Claude at the local OneDrive-synced
   folder on disk (ordinary file read/write, sync client pushes changes up
   automatically); alternative is adding the same Zapier OneDrive connector
   locally if sync-timing is a concern. Not yet decided/configured.
4. The archived run file to reconcile against:
   `Payroll Reports/Archive/AF Reporting Book Completed Run for Pay Period
   09302026.xlsm` (OneDrive item id, if still needed:
   `01ATC5E22ZMDWPTN6FV5AIEFDINCUVDDOE`, though a fresh local OneDrive session
   should resolve it by path instead).

## Reusable logic (not yet committed to the repo — recreate or port if useful)

During the cloud session, a Python helper (`knack_pull.py`, scratch-only, not in
git) replicated the VBA account-specific quirk that Knack connection fields
return as HTML-wrapped plain strings (e.g. `<span ...>Name</span>`) rather than
ID-bearing objects in this account — mirroring `ProcessConnRecords.bas`'s
`GetFieldValue`/`GetConnValue`/`CleanKnackText`/`NormalizeConnectionText`. If the
local session ends up doing pulls in Python/another language rather than driving
the VBA macros directly, that logic will need to be reproduced (or the VBA run
directly, which a local session might finally be able to do if it has an actual
Excel/VBA execution path — worth checking, since neither session so far has had
one).
