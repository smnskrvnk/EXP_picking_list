# Build history — CSS source addition

> Covers the build up to Step 6 (10.09.2026). Later changes — whole-pack rule,
> Packaging query, settings sheet — are recorded in `PICKING_LIST.md`, section 7.

Working document. Status as of 10.09.2026.

**Steps 3–6 are DONE and applied to `_shippingPlan_v04_PANEL_WORKING.xlsm`.**
See §5. For a full description of the finished system (not just this feature)
see `PICKING_LIST.md`.

---

## 1. Goal

Warehouse can generate a picking list from **two sources**:

- **PLAN** — shipment plan file, one sheet per Slovak day name, ready ~10:00
- **CSS** — long-term customer order file, `Ship Sched` sheet, ready ~07:00, covers ~16 working days ahead

CSS lets the warehouse start earlier and pick for days beyond tomorrow.

---

## 2. Current state — DONE and tested

### Module3_CSS (new module)

| Procedure | Purpose |
|---|---|
| `RefreshCSS` | Opens CSS file read-only, copies `Ship Sched` into local `CSS` sheet as values, then calls `BuildExportDateList` |
| `BuildExportDateList` | Fills the F5 date list — CSS FIX dates if the CSS sheet is loaded, else the next 15 working days; writes them to `PACK_SIZE` col Z, (re)creates `CSS_DATES`, sets F5 validation. Replaces the old `BuildCSSDateList`. |
| `NextWorkingDays` | Next N Mon–Fri dates starting from tomorrow |
| `SlovakDayName` | Date → `Pondelok`…`Piatok`, `""` for a weekend (`Weekday()` based, no locale) |
| `ExportDate` | PANEL F5 as a real `Date` — explicit `DD.MM.YYYY` split-parse, `0` if empty/invalid |
| `CSSHasData` | True if the local CSS sheet currently holds FIX date columns |
| `CSSDateLoaded` | True if a given `DD.MM.YYYY` is one of the loaded CSS columns |
| `CSSFixDateCols` | Finds FIX day columns — anchors on `Ship_Sch_Backlog` in row 21, walks right while row 20 holds a date |
| `FindHeaderCol` | Generic header lookup by name in a given row |
| `LoadItemsCSS` | Reads chosen date's rows into uniform item array — moved here from `Module4_CSS_Debug` in Step 6 |
| `LoadItemsPLAN` | Same for PLAN sheet — logic moved unchanged from old `Vytvor_zoznam` |
| `GetOrCreateSheet` | Private helper |

### Module4_CSS_Debug (new module)

`TestCSSDates`, `TestLoadCSS`, `DumpCSSHeaders` — all output to Immediate window (Ctrl+G).

### Module1

`Vytvor_zoznam` refactored. Reads a source selector, calls one loader or the other, loops the
returned array. In Step 3 only 4 lines differed in the 240-line preserved core (the
`deliveryNote` → `groupHeader` rename and `Next i` → `Next n`). Step 6 (4a/4b) then changed the
STOCK-scan loop to read from an in-memory `stockData` array and wrapped the sub in a
`CleanExit` / `CleanFail` handler — the allocation / sort / layout code below the scan is
still untouched.

### Constants / cells in use

| Item | Location |
|---|---|
| Plan file path | PACK_SIZE **X1** |
| CSS file path | PACK_SIZE **X2** |
| CSS date list (helper) | PACK_SIZE **Z1:Zn** |
| Defined name | `CSS_DATES` |
| Source selector | PANEL **B5** (`PANEL_SOURCE_CELL`) — renamed from `CSS_SOURCE_CELL` in Step 6 |
| Date selector | PANEL **F5** (`PANEL_DATE_CELL`) — renamed from `CSS_DATE_CELL` in Step 6 |
| CSS refresh stamp | PANEL **B18** |

### Item array format

Both loaders return the same structure:

```
items(n,1) = group header   ' delivery note (PLAN) or customer (CSS)
items(n,2) = Acme PN
items(n,3) = Customer PN
items(n,4) = quantity
```

### Verified

- PLAN mode output identical to pre-refactor version
- CSS mode: 27 items, 1390 pcs, two customers, totals match the source file
- No missing pack sizes on test date

---

## 3. CSS file structure — reference

Sheet `Ship Sched`:

| Row | Content |
|---|---|
| 1–19 | Summary block — **contains merged cells, do not copy** |
| 20 | Real dates (as serial numbers after values-only copy) |
| 21 | Table headers |
| 22+ | Data |

Columns:

| Col | Header |
|---|---|
| A | Customer |
| E | Unloading Point |
| F | Ship-To |
| J | Customer Item2 → Customer PN |
| K | Acme Item → Acme PN |
| Q | `RSS_Backlog` — start of RSS block (NOT used) |
| BB | `Ship_Sch_Backlog` — anchor for FIX block |
| BC:BR | FIX daily quantities — **this is what we use** |
| BS, BT | `SUM in new FIX 1/2` — stop scanning here |
| BW, BX | `Ship_Sch_Backlog_D-1`, `Ship_Sch_Backlog2` — old block, ignore |

Row-21 header names in the FIX block are auto-generated junk
(`Schip_Sched_FIX_Sum_Check135222222...`) and vary between file versions.
**Never match on them.** Anchor on `Ship_Sch_Backlog` and use row 20 dates.

Customers are contiguous — 18 customers, 18 stretches, no repeats. No sorting needed.

---

## 4. Hard-won lessons — do not regress

| Problem | Cause | Fix |
|---|---|---|
| `Dir()` fails on CSS path | Japanese folder name `<company-folder-with-Japanese-name>`; `Dir` is ANSI and mangles it | Removed the `Dir` check; `Workbooks.Open` is COM and handles Unicode |
| Error 1004 "merged cell" | Copy range included rows 1–19 summary block | Start copy at row 20 |
| Only 3 rows loaded instead of 27 | `Range.Copy` transfers **visible cells only**; someone left a filter on Customer | Direct `.Value = .Value` assignment, immune to filters |
| Blank error message | `On Error Resume Next` resets `Err` before `MsgBox` | Capture `Err.Number` / `Err.Description` into variables first |
| Zero FIX columns found | Values-only paste turns dates into serial numbers; `IsDate(46273)` is False | Accept numeric values in range 30000–80000 as dates |
| Nothing happens on test sub | `Debug.Print` needs Immediate window | Ctrl+G |

From Step 3–4 (PANEL rework + PLAN date check):

| Problem | Cause | Fix |
|---|---|---|
| `Štvrtok` imported as `Ĺ tvrtok` | VBE *File → Import* reads the `.bas` as the system ANSI codepage (1250), not UTF-8 | Save the `.bas` as **cp1250** before importing (Import then works); or paste into the code pane; or load via COM `CodeModule.AddFromString` |
| `ExportDate()` fragile off a Slovak box | relied on `IsDate`/`CDate` to parse `"DD.MM.YYYY"` — locale-dependent | explicit `Split(s, ".")` → `DateSerial(y, m, d)` |
| any Monday maps to *the* `Pondelok` sheet | PLAN file reuses 5 day-name sheets every week | Step 4: after the copy, compare `PLAN!C1` (*Vývoz:* date) to the chosen date; on mismatch warn, on "Nie" wipe PLAN |
| day name from a cell went stale | B5 used to hold the Slovak day name | derive it with `SlovakDayName(ExportDate())` — `Weekday()` based, no locale, no cell |

Also from earlier work — still applies:

- Slovak decimal separators: `Replace(rawPack, ".", ",")` before `IsNumeric`
- Delivery note read **before** the `planPN` empty check, so header-only rows carry forward
- Dynamic column detection by header name — positions vary

---

## 5. Steps — status

### Step 3 — PANEL rework — DONE (08.09.2026)

Decided: **option 2** — source first, then date. Single date selector, real dates not day names.

New layout:

| Row | B:D | F:G |
|---|---|---|
| 4 | label `Zdroj dát` | label `Deň exportu` |
| 5 | dropdown PLAN / CSS | dropdown of dates |
| 7 | `Aktualizácia skladu` | `Vytvor zoznam` (F7:G9) |
| 9 | `Aktualizácia zdroja` — dispatches on source | |
| 11 | Clear plan sheet | |
| 12 | Clear stock sheet | |
| 14–18 | Status block | |

Work:

- [x] Back up PANEL sheet, export Module1, Module2, and the PANEL sheet module
- [x] Rewrite `BuildPanel` with new layout
- [x] B5 becomes source dropdown (PLAN / CSS), validated list
- [x] F5 becomes date dropdown from `CSS_DATES`
- [x] Date list source: CSS header dates; fall back to next 15 computed working days if CSS not yet refreshed
- [x] Single refresh button (`Btn_RefreshSource`) dispatching to `RefreshPlan` / `RefreshCSS`
- [x] `CSS_SOURCE_CELL` → `B5`, `CSS_DATE_CELL` → `F5` (constants later renamed in Step 6)
- [x] Remove temporary B20 / F20
- [x] Guard in `Btn_GenerateList`: CSS empty / date not loaded / PLAN weekend → stop with a message

Also folded in from the review: shared `RefreshStock` helper (skips non-query tables),
`RefreshPlan` → `Public Function ... As Boolean` with a full error handler, `.Value` copy
instead of `.Copy`, `EnableEvents=False` on open.

### Step 4 — RefreshPlan date verification — DONE (08.09.2026)

- [x] Slovak day name derived from the chosen date via `SlovakDayName()` (`Weekday()` based)
- [x] After copy, `VerifyPlanDate` compares `PlanSheetDate(wsDest)` (`PLAN!C1`, *Vývoz:*) to the chosen date
- [x] Mismatch or missing C1 → Áno/Nie dialog; **Nie** wipes PLAN and stamps B16, so a
      following `Vytvor_zoznam` hits the "no items" guard and stops

### Step 5 — RefreshPlan hardening — DONE (08.09.2026)

- [x] `.Copy` → direct `.Value` assignment (kills the filter-truncation bug)
- [x] error handler with captured `Err.Number` / `Err.Description`
- [x] `Application.EnableEvents = False` while opening the plan file
- [ ] `%OneDriveCommercial%` in X1 instead of the literal Japanese path — *still open, optional*

### Step 6 — cleanup pass — DONE (10.09.2026)

Non-urgent tidy-ups from the review. Applied in rounds; each round test-compiled on a
throwaway copy (Slovak intact) and `Vytvor_zoznam` output verified **byte-identical** in
PLAN and CSS mode before applying.

- [x] **#8/#9/#10** dead lines in `Vytvor_zoznam` — two top `DoEvents`, `missingPacks = 0`,
      merged `wsOut` into `wsDest`
- [x] **#11** `Option Explicit` added to Module1 and Module4 (every var was already declared)
- [x] **#16** `LoadItemsCSS` moved from `Module4_CSS_Debug` to `Module3_CSS`, beside `LoadItemsPLAN`
- [x] rename `CSS_SOURCE_CELL` → `PANEL_SOURCE_CELL`, `CSS_DATE_CELL` → `PANEL_DATE_CELL`
- [x] **#17 (4b)** `Vytvor_zoznam` perf — STOCK columns `A:R` read once into `stockData`
      (was ~250k per-cell `wsStock.Cells(j, …)` reads); indices `C=3 G=7 H=8 R=18`
- [x] **4a** `Vytvor_zoznam` — `ScreenUpdating` off around the build and restored on
      **every** exit; real error handler (`CleanExit` / `CleanFail`) so a mid-run error
      restores the screen and shows a message instead of a frozen screen + half-built sheet

Still open from the review, deliberately not done: **#12** `BuildPanel` doesn't recreate the
B5/F5 validation on a full rebuild (mitigated — `SetDefaultExport` runs at the end and rebuilds F5);
**#17 (partial)** the per-item pack-size `VLookup` on `PACK_SIZE` B:C is **not** batched into a
dictionary — small next to the STOCK read, left as-is; **#19** inline `Cells(1,24)` for the plan
path vs named consts. (**#18** `UpdateDayDisplay` was removed entirely in Step 3, so its guard is moot.)

---

## 6. Open questions / watch list

- **Pack size coverage.** CSS carries all 18 customers; PACK_SIZE has ~492 rows built around
  the PLAN scope. Test date had no gaps, but customers vary by day. Watch for
  `Chýba PACK SIZE` messages on other dates.
- **Column balance.** Three-column split targets a third of rows per column. CSS produces
  fewer, chunkier sections than PLAN's many delivery notes — may look lopsided.
- **Red section headers.** The `Left(...,1) = "-"` rule colours anything starting with `-`,
  including the `"-- ... --"` wrapper. Pre-existing, applies to delivery notes too.
  Fix only if it bothers the warehouse.
- **Duplicate PNs across customers.** Decided: allocations stay **independent** per row,
  same as PLAN today. Two customers needing the same PN may be pointed at the same location.

---

## 7. Backup checklist before each step

- Export the module being changed (File → Export File) to a folder outside the workbook
- Keep the whole `.xlsm` copy too — VBA state can corrupt independently of the code
- Test PLAN mode after any shared-code change, before testing CSS
