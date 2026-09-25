# Direct-to-Knack uploads (replacing the Zapier webhook buttons)

Replaces the three "Upload" buttons, which currently POST staging-sheet
rows to Zapier "catch" webhooks and let Zapier map fields into Knack, with
direct Knack API calls from VBA. Nothing about the existing buttons or
their old macros is touched or deleted — these are new modules you point
the buttons at once you're ready to cut over.

| Old button | Old macro | New macro | Knack object |
|---|---|---|---|
| Upload CO table | `macroUploadCORecords` (Button2) | `UploadCOTable` | CarryOver (`object_34`) |
| Upload Payroll | `macroUploadPayroll` (Button3) | `UploadPayrollReport` | Full Payroll Report Table (`object_44`) |
| Upload Next Months Targets | `UploadMOnthlyTargetRecords` (Button5) | `UploadNextMonthsTargets` | Employee Monthly Targets (`object_28`) |

## How the field mappings were confirmed

The old macros only POST raw JSON to Zapier — the actual column→Knack-field
mapping lives inside the Zap configuration, which isn't visible from the
workbook. Rather than guess, each mapping below was confirmed against a
**real, already-uploaded record** pulled live from Knack (via Zapier's own
"Find Record" action), cross-checked against each object's live field
schema. Where a sheet column had no matching Knack field, it's explicitly
skipped rather than guessed at — see each module's header comment for the
full list.

## 1. Import the modules

Import, in this order (helpers first, since the other three depend on it):
1. `KnackUploadHelpers.bas`
2. `UploadCarryOverRecords.bas`
3. `UploadMonthlyTargetsRecords.bas`
4. `UploadPayrollRecords.bas`

## 2. What each one does

### `UploadCOTable` (CarryOver)
Reads the `Upload` sheet (7 columns) and creates one `object_34` record per
row. Resolves `Local Manager` and `ReportingPeriod` to their Knack
connection IDs by searching for a matching name/date — this replaces what
Zapier was doing invisibly. Always **creates** new records; it never
searches for or updates an existing one (matching the old webhook's plain
POST behavior).

### `UploadPayrollReport` (Full Payroll Report Table)
Reads `Payroll_Sheet_ACTIVE` (89 columns) and creates one `object_44`
record per row. **20 columns have no matching Knack field and are not
uploaded** — the confirmation prompt and the completion summary both list
them explicitly:

> pt, mgrhrsw, hrswagetype, Not Used, Salary_Pay2, Bonus Paid, Part-Time,
> DG, DG+, QtrBonusQualify, PriorQtrRev, PriorQtrYearRev, L1, L2, L3,
> CO OT, Transfers, PR

If any of these are supposed to land somewhere in Knack, tell me which
field (on `object_44` or elsewhere) and I'll wire it in — I did not guess.
Always **creates** new records, same as CO.

### `UploadNextMonthsTargets` (Employee Monthly Targets)
This one works differently from the old macro on purpose. The old
`UploadMOnthlyTargetRecords` mutated the live `MonthlyTargets` sheet in
place (overwriting Target Month/Year, then restoring the original values
afterward) before posting it. This version **never touches the
MonthlyTargets sheet** — for every current row, it creates one brand-new
`object_28` record for next month:

- `Location`, `Local Manager`, and the monthly target inputs (Monthly New
  Member Sign Up Target, etc.) carry through unchanged.
- **`Prior Month Dues Tap` on the new record = `Current Month Dues Tap` on
  the current record** — this is the "current actuals become prior
  actuals" rollover you asked for.
- `Current Month Dues Tap`, `Actual New Member Sign Ups`, `Monthly Member
  Sign Up Bonus`, `Monthly Member Sign Up Excess`, and `New Members Sign
  Ups` are all set to blank/0 on the new record — they're per-period
  actuals that aren't known yet for the new month.
- `Target Month`/`Target Year` come from `VariablesSheet!E2` ("Next Target
  Period", already computed as `=EOMONTH(current period, 1)`), not from
  adding a month myself — that's the same authoritative cell the old macro
  used.
- `MBCode` and `Manager and MB Code` aren't sent — they're Knack
  concatenation fields Knack computes automatically from Target
  Month/Year.
- **Area is not set** on the new record — the `MonthlyTargets` sheet has
  no Area column to source it from, and I didn't want to guess a fragile
  two-hop lookup (Location → its current Area) without confirming it's
  wanted. Tell me if this should be filled in and I'll add it.

If you want a *different* rollover rule (e.g. carrying the target number
forward with an increase, or handling New Members Sign Ups differently),
tell me and I'll adjust `UploadOneTargetRow` in
`UploadMonthlyTargetsRecords.bas` — the whole transform lives in one place.

## 3. The confirmation summary

All three now show a popup after running with **records created**, **rows
skipped/failed**, and (up to 25) specific reasons for each skip/failure —
e.g. "Row 14 skipped: Could not find a Local Manager record matching
'J. Smith'." This replaces the old bare "Successful/Unsuccessful: X/Y"
count with enough detail to actually fix a bad row without guessing.

## 4. Wiring the buttons

Once you've tested each macro standalone (`Alt+F8`) and are ready to cut
over: right-click each button → **Assign Macro** → point it at the new
macro name instead of the old one. The old macros and their Zapier hooks
keep working untouched until you do this, so you can test side-by-side.

## Email notification on Actual New Member Sign Ups changes

Not part of these three uploads — recommended separately as a **Knack
Flow**: Trigger *Record Updated → Employee Monthly Targets*, Condition
*Actual New Member Sign Ups changed*, Action *Send Email*. This fires
regardless of what updates the field (this macro, a manual edit, anything
else), which a check built only into `UploadPayrollReport` or any other
single macro can't guarantee.
