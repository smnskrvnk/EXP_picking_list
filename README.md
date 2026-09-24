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
- Verifies that the PLAN sheet it pulled is really dated for the chosen day, and
  discards it if the user rejects a stale plan.

Full description: [`docs/PICKING_LIST.md`](docs/PICKING_LIST.md).
Design decisions and lessons learned: [`docs/BUILD_HISTORY.md`](docs/BUILD_HISTORY.md).

## Layout

```
src/
  Module1_Vytvor_zoznam.bas   allocation / layout engine, PLAN refresh
  Module2_PANEL.bas           control panel UI and button handlers
  Module3_CSS.bas             CSS refresh, date list, item loaders
  Module4_CSS_Debug.bas       Immediate-window diagnostics only
  Sheet2_PANEL.cls            PANEL sheet code-behind (Worksheet_Change)
  ThisWorkbook.cls            Workbook_Open
docs/
  PICKING_LIST.md             system reference
  BUILD_HISTORY.md            build log and design decisions
```

## Requirements

- Excel with macros enabled (`.xlsm`).
- Sheets: `PANEL`, `STOCK` (query-backed), `PLAN`, `CSS`, `PACK_SIZE`; the output sheet
  `TVOJ ZOZNAM` is created on each run.
- `PACK_SIZE`: item → pack size in columns B:C; file paths in `X1` (PLAN) and `X2` (CSS).
- Run `BuildPanel` once to draw the control panel.

## Importing the source into a workbook

The files in `src/` are stored as **UTF-8** so they display correctly on GitHub. The VBA
editor's *File → Import File* reads the system ANSI code page instead, so Slovak letters
would come out garbled. Before importing, open each `.bas` in Notepad and use
*Save As → Encoding: ANSI*, or paste the code into the editor instead of importing.
Sheet and workbook modules (`.cls`) must be pasted into the existing module rather than
imported.

## Placeholder names

This is a sanitised copy: company- and site-specific values were replaced with
placeholders. **The code will not run against real data until you set your own values.**

| Placeholder | Where | What to set |
|---|---|---|
| `Acme PN` | header of the part-number column on the PLAN sheet | your column header text |
| `Acme Item` | header of the part-number column in the CSS file | your column header text |
| `"1000"`, `"2000"` | `LoadItemsPLAN` in `Module3_CSS.bas` | the delivery-note prefixes you want to pick for |
| `LOC_EXCL_1` … `LOC_EXCL_5` | `Vytvor_zoznam` in `Module1_Vytvor_zoznam.bas` | stock locations to exclude from picking (quarantine, scrap, etc.) |

Header strings are matched by name, so they must equal your data exactly.

## Not included

## Not included

The workbook, backups, source data files, and any real paths or credentials. See
`.gitignore`. No licence has been chosen yet.
