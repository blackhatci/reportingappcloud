# Additional Knack pull: Monthly Targets

Adds a new, read-only pull of Knack `object_28` ("Employee Monthly
Targets") into `AF_Reporting_Book.xlsm`, following the exact same
pattern already used by `PullLeadsRecords.bas` (Leads → `Leads_Pull`
sheet). It does **not** touch the existing `MonthlyTargets` sheet or
the Zapier upload macro (`UploadMOnthlyTargetRecords`) — those keep
working exactly as they do today. This is additive only.

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

## 1. Import the new VBA module

1. Open `AF_Reporting_Book.xlsm`.
2. `Alt+F11` to open the VBA editor.
3. `File → Import File…` and select `PullMonthlyTargets.bas` (in this
   folder).
4. Save the workbook as macro-enabled (`.xlsm`).

## 2. Add the field mapping row

On the **`Knack Field Mappings`** sheet, add a new header/field-key
pair in **rows 10–11**, columns A→Q (this follows the same layout as
the Hours mapping in rows 1–2 and the Leads mapping in rows 8–9):

| Col | Row 10 (header)                     | Row 11 (Knack field key) |
|-----|--------------------------------------|---------------------------|
| A   | ID                                    | field_363 |
| B   | Location                              | field_368 |
| C   | Local Manager                         | field_364 |
| D   | AFReporting_Period                    | field_367 |
| E   | Monthly New Member Sign Up Target     | field_366 |
| F   | Prior Month Dues Tap                  | field_365 |
| G   | Current Month Dues Tap                | field_390 |
| H   | Monthly Revenue Bonus                 | field_389 |
| I   | Monthly Revenue 500 Bonus             | field_392 |
| J   | Target Year                           | field_387 |
| K   | Target Month                          | field_385 |
| L   | MBCode                                | field_384 |
| M   | Actual New Member Sign Ups            | field_393 |
| N   | Monthly Member Sign Up Bonus          | field_394 |
| O   | Monthly Member Sign Up Excess         | field_395 |
| P   | Manager and MB Code                   | field_397 |
| Q   | New Members Sign Ups                  | field_478 |

A ready-to-paste copy of these two rows is in
`monthly_targets_mapping.csv` — paste it starting at `A10` on the
mapping sheet.

## 3. Create the output sheet and table

1. Add a new sheet named exactly **`MonthlyTargets_Pull`**.
2. In `A1` put the label `Target Month` and leave `B1` blank (or type
   a value like `09` to filter).
3. In `A2` put the label `Target Year` and leave `B2` blank (or type
   `2026` to filter). Leaving both blank pulls every record.
4. Starting at `A4`, add the same 17 header names as row 10 of the
   mapping sheet (in the same order), then select that single header
   row and **Insert → Table** to create an Excel Table named
   **`MonthlyTargetsTablePull`** (Table Design → Table Name).

## 4. Run it

`Alt+F8` → `PullMonthlyTargets` → `Run`. It pages through all matching
`object_28` records via `KnackAPI.GetRecordsByFilters` and writes them
into `MonthlyTargetsTablePull`, same as the existing Leads/Hours pulls.

Optionally wire a button to it (`Developer → Insert → Button`, assign
macro `PullMonthlyTargets`) the same way `Button5_Click` is wired to
`UploadMOnthlyTargetRecords`.

## Adding further pulls later

To pull another table (e.g. Location Revenue / object_68, Marketing
Campaigns / object_60, etc.), copy `PullMonthlyTargets.bas`, rename
the `Mtgt` prefix and constants for the new object, look up its field
keys, and repeat steps 2–3 with the next free mapping-row pair (rows
12–13, then 14–15, and so on).
