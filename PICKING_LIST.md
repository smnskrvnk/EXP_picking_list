# Picking list generator — system reference

What the workbook does and how, end to end. `CSS_plan.md` in this same folder is the
build history (why each decision was made, lessons learned along the way); this doc is
the finished-state reference. Status as of 25.09.2026.

---

## 1. Purpose

One button turns the warehouse's shipment plan for a chosen day into a printable,
location-sorted picking list — allocated against live stock, split across three columns,
shortages flagged in red. Two independent data sources feed it:

- **PLAN** — an external "Shipment planning" workbook, one sheet per Slovak day name,
  refreshed daily. Covers roughly the next 1–2 working days.
- **CSS** — a separate long-range order file (`Ship Sched` sheet), refreshed less often,
  covering ~16 working days ahead. Lets the warehouse pick further out than PLAN allows.

Both sources normalise into the same internal shape before the allocation logic ever runs.

---

## 2. Workbook layout

| Sheet | Role |
|---|---|
| **PANEL** | Control panel — source/date pickers, buttons, status lines. The only sheet a warehouse user interacts with directly. |
| **STOCK** | Live inventory, refreshed from a QAD query (QAD query connection) via `ListObjects`/`QueryTables`. Read-only from the macro's side. |
| **PLAN** | Local landing sheet for whichever day's plan was last pulled from the external PLAN file. Overwritten each refresh. |
| **CSS** | Local landing sheet for the external CSS file's `Ship Sched` data, copied in as values (row 20 down). Overwritten each refresh. |
| **PACKAGING** | Pack sizes. A Power Query table (query `Packaging`) loaded from the `Packaging` sheet of the PLAN workbook. Columns: `Customer PN`, `Acme PN`, `Total content` (= pack size), `LoadedAt` (time the query ran — used to detect a failed refresh). |
| **NASTAVENIA** | Settings. Two labelled cells holding the file paths: the PLAN workbook (`PATH_PLAN`) and the CSS file (`PATH_CSS`). Paths are written with a `%USERPROFILE%` placeholder instead of the user's profile folder, so the same workbook runs on any PC or login. The cells are **named ranges**, so they survive column/row edits. |
| **PACK_SIZE** | Now holds only the date-dropdown helper list (column Z, feeding the named range `CSS_DATES`). It used to hold a hand-typed pack-size table (columns A–H) and the file paths (`X1`, `X2`); the paths moved to **NASTAVENIA** and the table is superseded by **PACKAGING**, so both are unused. Any cleanup: *Clear Contents* only — never delete columns or rows, and keep column Z. |
| **TVOJ ZOZNAM** | Generated output. Deleted and rebuilt from scratch on every "Vytvor zoznam" run — nothing here is meant to survive between runs. |

### VBA module map

| Module | Contains |
|---|---|
| `Module1_Vytvor_zoznam` | `Vytvor_zoznam` (the allocation/layout engine), `RefreshPlan` and its Step-4 date-verification helpers (`PlanSheetDate`, `PlausibleDate`, `CoerceDate`, `VerifyPlanDate`) |
| `Module2_PANEL` | Everything PANEL: button handlers, `BuildPanel`, `SetDefaultExport`, shared `RefreshStock`/`RefreshPackaging`/`SourceMode`/`Stamp` helpers |
| `Module3_CSS` | All source-cell constants, `RefreshCSS`, the shared export-date-list builder, both item loaders (`LoadItemsPLAN`, `LoadItemsCSS`), `LoadPackSizes`, the path helpers `ResolvePath` and `SyncPackagingPath`, and the CSS column-finding helpers |
| `Module4_CSS_Debug` | `TestLoadCSS`, `TestCSSDates`, `DumpCSSHeaders` — Immediate-window (Ctrl+G) diagnostics only, nothing here is called from the UI |
| `Tento_zošit` (ThisWorkbook) | `Workbook_Open` → `SetDefaultExport` |
| `Sheet2` (PANEL's code-behind) | `Worksheet_Change` — watches B5 only, rebuilds the F5 date list when the source changes |

---

## 3. The PANEL UI, as built today

```
┌─────────────────────────────────────────────────┐
│           Picking list - kontrolný panel          │  ← title, B2:G2
│                                                    │
│   Zdroj dát          │   Deň exportu               │  ← labels, row 4
│  [ PLAN ▾ ]          │  [ 23.09.2026 ▾ ]           │  ← B5 / F5 dropdowns
│                                                    │
│ [Aktualizácia skladu]│                              │
│                       │   [  Vytvor zoznam  ]       │  ← spans F7:G9
│ [Aktualizácia zdroja] │                              │
│                                                    │
│  Stav                                              │  ← label, row 14
│  Sklad: ...                                        │  ← B15
│  Zdroj: ...                                        │  ← B16
│  Zoznam: ...                                       │  ← B17
│  CSS: ...                                          │  ← B18
└─────────────────────────────────────────────────┘
```

Five shapes total: `BtnRefreshStock`, `BtnRefreshSource`, `BtnGenerateList` (the
prominent blue one), plus the B5/F5 selector cells (data-validation dropdowns, not
shapes). **There is no "clear PLAN" / "clear STOCK" button** — that pair (and the
`DangerButton` helper that drew them) was present through Step 3–5 and was deliberately
dropped afterward: no longer needed for the warehouse workflow. Rows 11–12 are now
plain spacer rows.

Running `BuildPanel` wipes and redraws the whole sheet from scratch, then calls
`SetDefaultExport` — so it always leaves B5/F5 in a valid state, not just a valid layout.

---

## 4. Both source modes, end to end

### Common front door: `Btn_GenerateList`

Before calling `Vytvor_zoznam`, it checks (in order):
1. F5 has a date at all.
2. If source = CSS: the local CSS sheet has FIX date columns (`CSSHasData`), and the
   chosen date is actually one of them (`CSSDateLoaded`).
3. If source = PLAN: the chosen date isn't a weekend (`SlovakDayName` returns `""` for
   Sat/Sun).

Any failure stops here with a message — `Vytvor_zoznam` never starts on a source it
can't use.

### PLAN mode

1. **`Btn_RefreshSource`** → `RefreshPlan()`.
2. The chosen date (F5) is converted to a Slovak day name via `SlovakDayName` —
   **not** read from a cell; B5/F5 hold the source and date, nothing holds a day name
   anymore.
3. Opens the external plan workbook read-only (path from `NASTAVENIA`, cell `PATH_PLAN`, resolved by `ResolvePath`), pulls the sheet matching
   that day name, copies it into the local `PLAN` sheet via direct `.Value` assignment
   (immune to any filter left active on the source — a lesson from early testing).
4. **Date verification (Step 4):** the plan file only has 5 day-named sheets, reused
   every week, so pulling "Streda" doesn't by itself guarantee it's *this* Wednesday's
   plan. `PlanSheetDate` reads the sheet's own `Vývoz:` date out of `C1`; if it doesn't
   match the chosen date, `VerifyPlanDate` pops an Áno/Nie dialog. **Nie wipes the PLAN
   sheet** so a following `Vytvor zoznam` finds nothing rather than silently using stale
   data.
5. `Btn_GenerateList` → `Vytvor_zoznam` → `LoadItemsPLAN` reads `PLAN` by header name
   (row 3): `Acme PN`, `Zakaznik PN`, `KS`, `Delivery note`. Only rows under a
   `1000`- or `2000`-prefixed delivery note are kept; the delivery note is read *before*
   the part-number check so header-only rows still carry the note forward correctly.

### CSS mode

1. **`Btn_RefreshSource`** → `RefreshCSS()`.
2. Opens the external CSS workbook read-only (path from `NASTAVENIA`, cell `PATH_CSS`), locates `Ship Sched`,
   anchors on the `Acme Item` header (row 21) to find the last data row, and copies
   row 20 downward into the local `CSS` sheet as values — starting at row 20
   deliberately skips the merged summary block in rows 1–19.
3. `BuildExportDateList` scans for the `Ship_Sch_Backlog` anchor column, walks right
   while row 20 holds a real date (or a values-only-pasted serial number in the
   30000–80000 range), and writes the found dates to `PACK_SIZE!Z1:Zn`. This rebuilds
   the `CSS_DATES` named range and the F5 validation list. **This same function is also
   the shared F5 list builder** — if CSS has no data loaded, it falls back to the next
   15 computed working days instead, so F5 always has *something* valid regardless of
   which source is selected.
4. `Btn_GenerateList` → `Vytvor_zoznam` → `LoadItemsCSS` matches the chosen date against
   the FIX columns, then reads `Customer` / `Acme Item` / `Customer Item2` for each
   row, carrying the customer name forward across blank cells (CSS data is grouped in
   contiguous customer blocks, not repeated per row).

### Shared core: `Vytvor_zoznam` (from "items loaded" onward, identical for both modes)

1. `RefreshStock` refreshes every QAD-backed `ListObject` on STOCK (silently skips any
   non-query table rather than erroring the whole run). Then `RefreshPackaging` refreshes
   the pack-size query (see *Pack sizes* below) and `LoadPackSizes` builds the lookup.
2. STOCK's relevant columns (`A:R`, covering the ones actually used — pack-quantity,
   location, item PN, pick-status) are read into an in-memory array **once**, not
   re-queried per item — this was the Step 6 performance fix.
3. For each item: look up its pack size (dictionary from the `PACKAGING` table), sum available quantity per
   storage location (excluding `LOC_EXCL_1`/`LOC_EXCL_2`/`LOC_EXCL_3`/`LOC_EXCL_4`/`LOC_EXCL_5` locations and
   anything already marked `PICKED`), then allocate — exact-location match first, else
   largest-location-first until the required quantity is covered. Any shortfall becomes
   a `"-N"` entry (N = packs still missing), rendered in red on the output sheet.

   **Whole-pack rule (added 24.09.2026, after a warehouse report).** STOCK holds
   separate rows (lots) per part number and location, and they used to be summed
   blindly — so a leftover of 15 pcs next to a full 16 was counted as 31 pickable
   pieces. Now, when `requiredQty >= packSize`, each STOCK row is floored to a whole
   multiple of the pack size *before* it is summed (`Int(qty / packSize) * packSize`).
   An incomplete pack (e.g. 15 of 16) stays in the warehouse and never reaches the
   list. When `requiredQty < packSize` the order is itself a partial-pack quantity, so
   raw quantities count as before. The shortfall logic itself is unchanged — it only
   ever sees usable stock now.

   | required | pack size | stock (lots) | list shows |
   |---|---|---|---|
   | 16 | 16 | 15 | `-1` (no location) |
   | 32 | 16 | 15 | `-2` (no location) |
   | 32 | 16 | 16 | location `16` + `-1` |
   | 32 | 16 | 16 + 15 | location `16` + `-1` |
   | 28 | 36 | 20 | location `20` + `-1` |
4. Blocks are laid out into three columns on `TVOJ ZOZNAM`, trying to keep a delivery
   note / customer header together with its first item rather than splitting across a
   column break.
5. Wrapped in a real error handler (`CleanExit`/`CleanFail`) — `ScreenUpdating` is always
   restored on exit, success or failure, instead of leaving a frozen, half-built sheet
   behind.

### Pack sizes (Packaging query)

Pack sizes are no longer typed in by hand. A Power Query query named `Packaging` reads
the `Packaging` sheet of the PLAN workbook and loads it into the `PACKAGING` sheet.

- **`RefreshPackaging`** (Module2) refreshes the query at the start of every
  `Vytvor_zoznam` run, right after `RefreshStock`. It does not trust the refresh to
  report its own failure: it reads the `LoadedAt` stamp afterwards, and if that is older
  than `PKG_MAX_AGE_MIN` (5 minutes) the refresh did not really happen. The user is then
  asked whether to continue with the last loaded sizes (*Áno*) or stop (*Nie*). This
  happens before the old list is deleted, so *Nie* leaves the previous list intact.
- **`SyncPackagingPath`** (Module3) runs just before the refresh. The query stores its
  own copy of the PLAN file path, so this reads it, compares it with `PATH_PLAN`, and
  rewrites it only if they differ. The path therefore lives in one place. (The first
  refresh on a new PC/profile may show a one-time Power Query permission prompt.)
- **`LoadPackSizes`** (Module3) turns the table into a dictionary once per run
  (part number, trimmed and upper-case → size). Rows with a size of 0 or less are
  ignored; if a part number appears more than once, the first valid row wins.
- A part number with no usable size (missing, or 0 in `Packaging`) is skipped with
  *"Chýba PACK SIZE …"*. Wrong or missing sizes are corrected at the source — the PLAN
  workbook — not in this tool.
- The `Packaging` sheet is not perfectly clean: a few part numbers appear on two rows
  with different sizes, and a few have a 0. The rules above decide what happens; the
  comparison helper `ComparePackSizes` (Module4) was used once to check the switch.

---

## 5. Constants / cells reference

| Item | Location | Constant |
|---|---|---|
| Source selector (PLAN / CSS) | PANEL **B5** | `PANEL_SOURCE_CELL` |
| Export date | PANEL **F5** | `PANEL_DATE_CELL` |
| Plan file path | NASTAVENIA, named cell `PATH_PLAN` | `PATH_PLAN_NAME` |
| CSS file path | NASTAVENIA, named cell `PATH_CSS` | `PATH_CSS_NAME` |
| Packaging query name | Power Query `Packaging` | `PACKAGING_QUERY` |
| Date-dropdown helper list | PACK_SIZE **Z1:Zn** | `CSS_LIST_COL`; named range `CSS_DATES` |
| Pack-size table | sheet **PACKAGING** (query `Packaging`) | `PACKAGING_SHEET` |
| Query run-time column | PACKAGING, column `LoadedAt` | `PKG_STAMP_HEADER` |
| Max age of query data | 5 minutes | `PKG_MAX_AGE_MIN` |
| Stock stamp | PANEL **B15** | — |
| Source stamp | PANEL **B16** | — |
| List stamp | PANEL **B17** | — |
| CSS refresh stamp | PANEL **B18** | — |

Item array shape both loaders return:

```
items(n,1) = group header   ' delivery note (PLAN) or customer (CSS)
items(n,2) = Acme PN
items(n,3) = Customer PN
items(n,4) = quantity
```

---

## 6. Known limitations (deliberate, not bugs)

- **Pack sizes are only as good as the `Packaging` sheet.** Duplicates resolve to the
  first valid row and zeros count as "no size" (see *Pack sizes*). If one part number
  genuinely has two pack sizes (e.g. per plant), a lookup by part number alone cannot
  pick the right one — that would need the plant in the key.
- **Column balance.** The three-column split targets an even third of *rows*. CSS
  produces fewer, chunkier customer blocks than PLAN's many delivery notes, so CSS-mode
  output can look less balanced across the three columns than PLAN-mode.
- **Duplicate part numbers across customers/delivery notes allocate independently.**
  Two rows needing the same part number can each be pointed at the same physical
  location — the allocator has no cross-row awareness. Decided behaviour, not a defect.
- **Red "shortage" colouring keys off a leading `-`**, which also happens to be the
  first character of every `"-- delivery note --"` group header. Cosmetic overlap,
  pre-existing, harmless.
- **File paths depend on the synced folder sitting directly under the user's profile.**
  Paths are stored as `%USERPROFILE%\<synced folder>\…`, so they work on any PC or login
  *provided* the synced folder has the same name and location relative to the profile
  (true on both the development and warehouse PCs). A `%OneDriveCommercial%` placeholder
  was investigated first (23.09.2026) and does not work here: that folder is not exposed
  through any OneDrive environment variable (only `OneDrive`/`OneDriveCommercial` exist,
  and neither matches). `USERPROFILE` is always defined, which is why it works.
  `ResolvePath` uses the built-in `Environ$` rather than `WScript.Shell` on purpose — a
  scripting-host object is a pattern security tools tend to flag in macros. Limitation:
  `Environ$` can garble a profile name with accented characters (none in use).

## 7. Open items

None currently. The two items tracked here (Clear-Plan/Clear-Stock removal intent;
the `RefreshCSS` `MsgBox` argument-order crash) were both resolved on 23.09.2026 —
the button removal was confirmed intentional, and the `MsgBox` fix has been applied
to the live workbook.

Also applied on 24.09.2026 (found while reviewing the code after the whole-pack change):
- `RefreshPlan`: the rejection branch now uses `ThisWorkbook.Worksheets("PANEL")`.
  The bare `Worksheets("PANEL")` resolved against the *active* workbook, which is
  not guaranteed to be this one right after the external plan file is closed.
- Three remaining English `MsgBox` texts translated to Slovak.
- `LoadItemsCSS` message no longer points at a `'Refresh CSS'` button that doesn't
  exist — it says `'Aktualizácia zdroja'`.
- B17 list stamp reads `Zoznam:` instead of `List:`.

Applied on 25.09.2026: pack sizes now come from the `Packaging` query instead of a
hand-typed table (see *Pack sizes*). A comparison run before the switch showed the old
table was stale or wrong in about a dozen places — new part-number revisions missing,
a few sizes that disagreed — which is why it was retired rather than kept as a fallback.

Also applied on 25.09.2026: the PLAN/CSS file paths moved from `PACK_SIZE!X1`/`X2` to
named cells on a new `NASTAVENIA` sheet, stored with `%USERPROFILE%` (see *Deploying*),
and the `Packaging` query's path is now kept in sync with `PATH_PLAN` automatically.
This came after the deployed copy failed on the development PC with "Sorry, we couldn't
find …" — a path pointing at the other PC's profile folder.

## 8. Code conventions

- **Comments are in Slovak** (translated across all modules on 24.09.2026), as is all
  user-facing text (`MsgBox`, labels, status lines).
- **Identifiers stay English** — procedure, variable and constant names are not
  translated. Sheet names (`STOCK`, `PLAN`, `CSS`, `PACK_SIZE`), cell references and
  established terms like `PACK SIZE` are also left as-is inside Slovak sentences.
- `Option Explicit` is on in every module. `Vytvor_zoznam` carries a header comment
  describing the allocation rules — update it if the rules change.

## 9. Deploying to another PC or account

File paths live in **one place**: the two named cells on the `NASTAVENIA` sheet
(`PATH_PLAN`, `PATH_CSS`), written with a `%USERPROFILE%` placeholder. `ResolvePath`
expands the placeholder at run time, and `SyncPackagingPath` keeps the Power Query's
own copy of the PLAN path in step with `PATH_PLAN`. So on a PC where the synced folder
has the same name, **nothing needs editing** after copying the workbook.

After copying the workbook to a new PC or login:
1. Enable macros and data connections.
2. Click *Aktualizácia zdroja* for each source, and generate one test list. The first
   Packaging refresh may show a one-time Power Query permission prompt — accept it.
3. If the synced folder is named or placed differently on that PC, edit the two cells
   on `NASTAVENIA` — nothing else.

**Rolling out a code change to the deployed copy:** deploy the new code *before*
clearing anything old. A copy still running the old code reads its paths from
`PACK_SIZE!X1`/`X2`, and clearing those first would break it.

**Symptom of a wrong path:** Excel's *"Sorry, we couldn't find `<path>`"* (run-time error
1004), or the Slovak *"Cesta k … súboru je prázdna"* if the `NASTAVENIA` cell is empty.
First check which PC and account the workbook is actually running on. This has already
happened once: the deployed (warehouse) copy was opened on the development PC, so its
paths pointed at a profile folder that didn't exist there.
