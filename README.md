# Warehouse picking-list generator (Excel / VBA)

An Excel macro workbook that turns a shipment plan into a printable, location-sorted
picking list. It allocates each required quantity against live stock, flags shortages,
and lays the result out in three columns for printing.

> **Status:** internal tool, in daily use. This repository contains the VBA source and
> documentation only — **not** the workbook itself (it holds live stock and customer
> data). Code comments and all user-facing text are in Slovak; identifiers are English.

## What it does

- Two interchangeable data sources, chosen on the control panel:
  - **PLAN** — a short-term shipment-planning workbook (one sheet per weekday).
  - **CSS** — a long-range order file covering roughly 16 working days ahead.
- Reads live stock from a QAD query table and allocates per storage location:
  exact-location match first, otherwise largest location first.
- **Whole-pack rule:** when an order needs at least one full pack, only whole packs
  count as pickable stock — an incomplete pack (e.g. 15 of 16 pcs) stays in the
  warehouse. Shortages are printed in red as `-N` (packs still missing).
- **Pack sizes are not typed in by hand.** A Power Query reads them from the
  `Packaging` sheet of the plan workbook. After every refresh the tool checks a load
  timestamp, so a silently failed refresh is caught and the user is asked before
  stale sizes are used.
- Verifies that the PLAN sheet it pulled is really dated for the chosen day, and
  discards it if the user rejects a stale plan.
- File paths are stored once, with a `%USERPROFILE%` placeholder, so the same workbook
  runs on any PC or Windows login.

Full description: [`docs/PICKING_LIST.md`](docs/PICKING_LIST.md).
Design decisions and lessons learned: [`docs/BUILD_HISTORY.md`](docs/BUILD_HISTORY.md).

## Layout

```
src/
  Module1_Vytvor_zoznam.bas   allocation / layout engine, PLAN refresh
  Module2_PANEL.bas           control panel UI, button handlers, pack-size refresh
  Module3_CSS.bas             CSS refresh, date list, item loaders, path helpers
  Module4_CSS_Debug.bas       Immediate-window diagnostics only
  Sheet2_PANEL.cls            PANEL sheet code-behind (Worksheet_Change)
  ThisWorkbook.cls            Workbook_Open
docs/
  PICKING_LIST.md             system reference
  BUILD_HISTORY.md            build log and design decisions
```

## Setting up a workbook

Requires Excel with macros enabled (`.xlsm`) and Power Query (Excel 2016 or later).

| Sheet | What it is |
|---|---|
| `PANEL` | control panel; run `BuildPanel` once to draw it |
| `STOCK` | live stock, loaded by a query table (a QAD query) |
| `PLAN`, `CSS` | local landing sheets, filled by the refresh buttons |
| `PACKAGING` | pack sizes, loaded by the `Packaging` Power Query (see below) |
| `NASTAVENIA` | settings: the two file paths (see below) |
| `PACK_SIZE` | only the date-dropdown helper list in column Z; keep this sheet |
| `TVOJ ZOZNAM` | the output; created on each run |

### `NASTAVENIA` — file paths

Two cells holding full paths, each given a **workbook-scope name** (*Formulas → Name
Manager*): `PATH_PLAN` (the plan workbook) and `PATH_CSS` (the CSS file). A path may
start with `%USERPROFILE%`, which the code replaces with the current user's profile
folder. This works when the synced folder has the same name and location relative to
the profile on every PC.

### `PACKAGING` — pack sizes (Power Query)

Create a query named **`Packaging`** (*Data → Get Data → From Excel Workbook*, pointing
at the plan workbook, sheet `Packaging`) with these steps:

1. Promote the first row to headers (once).
2. Keep three columns, in this order: `Customer PN`, `Acme PN`, `Total content`.
3. Trim and upper-case `Acme PN`; set `Total content` to Whole Number.
4. Remove rows where `Acme PN` is empty.
5. Add a custom column **`LoadedAt`** = `DateTime.LocalNow()`.
6. Load to a table on a new sheet named **`PACKAGING`**.

The code depends on: sheet name `PACKAGING`, query name `Packaging`, part number in
column 2, pack size in column 3, and the header `LoadedAt`. Before each refresh the
code rewrites the file path inside the query from `PATH_PLAN`, so the path only ever
needs changing in one place.

## Placeholder names

This is a sanitised copy: company- and site-specific values were replaced with
placeholders. **The code will not run against real data until you set your own values.**

| Placeholder | Where | What to set |
|---|---|---|
| `Acme PN` | header of the part-number column on the PLAN sheet and in the `Packaging` sheet | your column header text |
| `Acme Item` | header of the part-number column in the CSS file | your column header text |
| `"1000"`, `"2000"` | `LoadItemsPLAN` in `Module3_CSS.bas` | the delivery-note prefixes you want to pick for |
| `LOC_EXCL_1` … `LOC_EXCL_5` | `Vytvor_zoznam` in `Module1_Vytvor_zoznam.bas` | stock locations to exclude from picking (quarantine, scrap, etc.) |

Header strings are matched by name, so they must equal your data exactly.

## Importing the source into a workbook

The files in `src/` are stored as **UTF-8** so they display correctly on GitHub. The VBA
editor's *File → Import File* reads the system ANSI code page instead, so Slovak letters
would come out garbled. Before importing, open each `.bas` in Notepad and use
*Save As → Encoding: ANSI*, or paste the code into the editor instead of importing.
Sheet and workbook modules (`.cls`) must be pasted into the existing module rather than
imported.

## Not included

The workbook, the Power Query itself (rebuilt from the steps above), backups, source
data files, and any real paths or credentials. See `.gitignore`. No licence has been
chosen yet.
