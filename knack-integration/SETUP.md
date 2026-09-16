# Additional Knack pull: Monthly Targets

Adds a new, callable pull of Knack `object_28` ("Employee Monthly
Targets") into `AF_Reporting_Book.xlsm`, using `PullHoursRecords.bas`
as the template (page through Knack → filter to the current
reporting period → write mapped values into the sheet).

Knack application: **Anytime Fitness Reporting App**
(`67a3de417009da15fb223d49`).

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

## What this module does differently from the other Pull modules

`MonthlyTargets` (VBA codename `Sheet4`) is the **live** sheet that
`macroUploadMonthlyTargets.bas` reads and writes directly — including
a `Next Target Period` column it fills in after every upload, and
temporary staging columns further to the right. Because of that, this
module deliberately does not behave exactly like `PullHoursRecords`
in one respect:

- It reads the sheet's **existing header row** at runtime and maps
  each header it recognizes (by name) to a Knack field key. **Column
  order is whatever is already on the sheet — nothing is
  reordered, inserted, or removed.**
- It only ever reads/writes/clears the columns it recognizes as
  Knack-backed (`ID` through `New Members Sign Ups`, 17 columns as of
  this writing). `Next Target Period` and anything to its right are
  never touched, so the upload workflow keeps working unchanged.
- Since `object_28` has no plain date field (only a connection to
  `AFReporting_Period`), it pages in *all* records and then filters
  them in VBA to the period in `VariablesSheet!A2` — the same
  connected-field technique `PullHoursRecords`/`PullLeadsRecords`
  already use for their date-range narrowing.

No changes to the `Knack Field Mappings` sheet are needed — the
field map lives in code (`MtgtFieldKeyMap`), keyed off the header
text already on `MonthlyTargets`.

## 1. Import the VBA module

1. Open `AF_Reporting_Book.xlsm`.
2. `Alt+F11` to open the VBA editor.
3. `File → Import File…` and select `PullMonthlyTargetsRecords.bas`
   (in this folder).
4. Save the workbook as macro-enabled (`.xlsm`).

## 2. Run it

`Alt+F8` → `PullMonthlyTargets` → `Run`. It:

1. Reads the target period from `VariablesSheet!A2`.
2. Pages through every `object_28` record via
   `KnackAPI.KnackGetRecords`.
3. Keeps only the records whose `AFReporting_Period` connection
   matches that date.
4. Writes them into `MonthlyTargets`, column-for-column matching the
   sheet's current header order, leaving `Next Target Period` (and
   anything past it) untouched.

Optionally wire a button to it (`Developer → Insert → Button`, assign
macro `PullMonthlyTargets`), the same way `Button5_Click` triggers
`UploadMOnthlyTargetRecords`.

## If you add or reorder columns on MonthlyTargets later

Nothing to change in the macro as long as the header text matches one
of the names in `MtgtFieldKeyMap` (in `PullMonthlyTargetsRecords.bas`)
— the column can move anywhere. Adding a brand-new Knack-backed
column just needs one more `map.Add "Header Name", "field_xxx"` line
there.

## Adding pulls for other tables

To pull a different object (e.g. Location Revenue / object_68,
Marketing Campaigns / object_60), copy
`PullMonthlyTargetsRecords.bas`, rename the `Mtgt` prefix/constants,
swap in the new object key and field map, and point
`MTGT_OUTPUT_SHEET` at that table's sheet.
