# Additional Knack pulls

Callable pulls of Knack tables into `AF_Reporting_Book.xlsm`, all
using `PullHoursRecords.bas` as the template (page through Knack →
write mapped values into the existing sheet, preserving its current
column order).

Knack application: **Anytime Fitness Reporting App**
(`67a3de417009da15fb223d49`).

| Module | Knack table | Sub to run | Writes into | Scope |
|---|---|---|---|---|
| `PullMonthlyTargetsRecords.bas` | Employee Monthly Targets (`object_28`) | `PullMonthlyTargets` | `MonthlyTargets` | Current reporting period |
| `PullCarryOverRecords.bas` | CarryOver (`object_34`) | `PullCarryOver` | `OTCarryOver` | Current reporting period |
| `PullAFDelinquencyRecords.bas` | AFDelinquency (`object_12`) | `PullAFDelinquency` | `afdelinquency` | Current reporting period |
| `PullLocationsRecords.bas` | Locations (`object_22`) | `PullLocations` | `Area Location` | All active locations |
| `PullPaidHolidaysRecords.bas` | Paid Holidays (`object_36`) | `PullPaidHolidays` | `PaidHolidays` | All records |

## ⚠️ Rotate the Knack API key

While reviewing the workbook's `KnackAPI` VBA module, the Knack
Application ID and REST API Key were found hardcoded in plaintext:

```
KNACK_APP_ID  = 67a3de417009da15fb223d49
KNACK_API_KEY = <redacted - see the KnackAPI module in the .xlsm>
```

Anyone who opens the workbook (or unzips the `.xlsm`, which is just a
zip file) can read this key and make authenticated calls against your
Knack app. Recommend rotating it now in the Knack Builder:
**Settings → API & Code → regenerate REST API Key** — then update the
`KNACK_API_KEY` constant in the `KnackAPI` module with the new value.

## How these modules work

All five follow the same shape:

- Each reads the **existing header row** of its target sheet at
  runtime and maps every header it recognizes (by name) to a Knack
  field key. **Column order is whatever is already on the sheet —
  nothing is reordered, inserted, or removed.** No changes to the
  `Knack Field Mappings` sheet are needed; each module's field map
  lives in its own code (e.g. `MtgtFieldKeyMap`, `CoFieldKeyMap`,
  `DelFieldKeyMap`).
- Each only ever reads/writes/clears the columns it recognizes as
  Knack-backed. Anything on the sheet it doesn't recognize is left
  completely alone:
  - `MonthlyTargets`: `Next Target Period` and anything past it are
    untouched, since `macroUploadMonthlyTargets.bas` depends on them.
  - `afdelinquency`: `Employee Wage Code` is left unmapped —
    object_12 only has one wage-code field
    (`EmployeeWageCode`/`field_373`), already mapped to the
    `1 AFEmployeeWages_Code` column, so this avoids guessing which of
    the two columns should receive it.
- The three period-scoped pulls read the target reporting period from
  `VariablesSheet!A2` and narrow to it. `object_28` (Monthly Targets)
  and `object_34` (CarryOver) have no plain date field, only a
  connection to `AFReporting_Period`, so those two page in *all*
  records and filter to the matching period in VBA. `object_12`
  (AFDelinquency) also has no plain pay-period date field (only the
  "Pay Period when the Account was Closed" connection), so it's
  filtered the same way — the one difference is it *does* have a real
  `Mark for Delete` field, so that part is filtered on Knack's side
  first (same as `PullHoursRecords`/`PullLeadsRecords` do for their
  delete fields).
- `PullLocationsRecords` and `PullPaidHolidaysRecords` are reference
  tables, not tied to a reporting period, so neither reads
  `VariablesSheet`. Locations filters to `Active Location = Yes` on
  Knack's side; Paid Holidays pulls every record (it's a small table
  used elsewhere via Month/Year lookups).

## 1. Import the VBA modules

1. Open `AF_Reporting_Book.xlsm`.
2. `Alt+F11` to open the VBA editor.
3. `File → Import File…` and select each `.bas` file (in this
   folder) you want to add.
4. Save the workbook as macro-enabled (`.xlsm`).

## 2. Run them

`Alt+F8` → pick the Sub → `Run` (see table above for which Sub goes
with which module), or wire each to a button
(`Developer → Insert → Button`, assign the matching macro) the same
way `Button5_Click` triggers `UploadMOnthlyTargetRecords`.

To chain any of these into a larger "run everything" macro, just add
a line calling the Sub by name (e.g. `PullCarryOver`) — no other setup
needed once the module is imported.

## If you add or reorder columns on one of these sheets later

Nothing to change in the macro as long as the header text matches one
of the names in that module's field map — the column can move
anywhere. Adding a brand-new Knack-backed column just needs one more
`map.Add "Header Name", "field_xxx"` line in that module's `*FieldKeyMap`
function.

## Adding pulls for other tables

Copy whichever of these three modules is the closest fit (a
connected-only period field → base it on `PullMonthlyTargetsRecords`
or `PullCarryOverRecords`; a table with its own delete flag → base it
on `PullAFDelinquencyRecords`), rename its prefix/constants, swap in
the new object key and field map, and point its output-sheet constant
at that table's sheet.
