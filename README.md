# Python-Enabled RAID Log

An automated RAID log built in Excel and powered by embedded Python.
Buttons inside the workbook trigger Python scripts that handle the
repetitive work of running a project log — colour-coding, flagging
stale items, building dashboards, generating reports — so project
managers can focus on the work, not the log.

> Built as part of a digital-products line for project managers.
> This repo holds the Python automation layer and project documentation.

---

## Why I Built This

Project managers live in RAID logs. Most of them are static Excel
files that require constant manual upkeep — conditional formatting,
sorting, status-chasing, weekly report rebuilds. That's hours of
work every week that doesn't add any project value.

This tool puts the routine work behind buttons. The PM clicks
**Flag Status** and every row colour-codes itself based on status
and due date. They click **Refresh Dashboard** and the project's
state rebuilds visually in seconds. The log stops being a chore
and starts being a tool.

It's also a deliberate exercise in building something sellable as
a solo developer — a single `.xlsm` file that ships with embedded
Python, distributed via Gumroad.

---

## How It Works

The architecture is three layers stitched together:

**Excel** — the interface. A pre-formatted workbook with a Home
tab full of buttons, a clean RAID Log structure, and supporting
tabs for outputs (Dashboard, Weekly Report, My Items).

**VBA** — the trigger layer. Each button on the Home tab calls a
small VBA macro that hands control over to Python via `xlwings`'
`RunPython` mechanism.

**Python** — the engine. The `raid_functions` module holds every
piece of automation logic. It reads the workbook in bulk (to avoid
slow per-cell round-trips), runs the logic in memory, and writes
results back with screen updating frozen for speed.

When packaged for distribution, the Python code is embedded
directly inside the `.xlsm` using `xlwings code embed`, so the
buyer downloads a single file with no external dependencies
beyond Python itself.

---

## Features

**Flag Status** — colour-codes the ID cell on every row based on
Status and Due Date. Cancelled items grey, complete green, overdue
red, due-within-7-days amber, on-track green. Idempotent — clears
its previous output before reapplying so the log stays accurate
as data changes.

**Mark Stale** — parses dated entries from the Actions Taken
column (users prefix updates with `08-May-26: chased vendor`)
and flags rows that haven't had a dated update in 14+ days.

**Refresh Dashboard** — rebuilds a Dashboard tab with stat cards
(overdue, due this week, open, complete), breakdowns by RAID
Category and Owner, and a criticality grouping.

**Generate Weekly Report** — writes a dated, formatted status
snapshot to the Weekly Report tab. Sections for overdue, due this
week, and in-progress items.

**Show My Items** — filters the log to a single owner (selected
from a dropdown on the Home tab) and writes the filtered view to
its own tab.

Future features under development: PDF export, Outlook email
draft generation, PowerPoint status deck generator, and a
visually upgraded dashboard using Plotly-rendered charts.

---

## Tech Stack

- **Python 3** — automation logic
- **xlwings** — Python ↔ Excel bridge
- **openpyxl, pandas** — data manipulation
- **VBA** — button triggers via `RunPython` macros
- **Excel** — the host environment and UI

---

## Workbook Structure

The workbook is built around a single combined RAID Log tab
covering Risks, Actions, Issues, and Dependencies. Headers in
row 1, data from row 2, no blank rows in the middle (so the
data-detection logic works cleanly).

The full structure:

- **Home** — control panel with buttons
- **Instructions** — short user guide
- **RAID Log** — the data (18 columns A–R)
- **Dashboard** — built by Refresh Dashboard
- **Weekly Report** — built by Generate Weekly Report
- **My Items** — built by Show My Items

---

## Engineering Notes

A few design decisions worth calling out:

**Idempotent functions.** Every automation can be re-run safely.
Each function clears its own previous output before reapplying,
which means the user can click the same button repeatedly without
accumulating stale state. This matters for a tool used over many
weeks.

**Bulk reads, single-call writes where possible.** Per-cell
Excel access is slow because each call crosses the Python/Excel
process boundary. The functions bulk-read the data block in one
call, do all logic in Python, and write back as efficiently as
the formatting operations allow.

**Screen updating frozen during writes.** Excel tries to repaint
the screen between every cell change, which is both slow and
visually jarring. Each function freezes screen updates at the
start and restores them at the end.

**Code embedded at packaging time.** The development version
uses a separate `.py` file. For distribution, `xlwings code embed`
pulls the Python into a hidden sheet inside the workbook so the
buyer ships and runs one file.

---

## Roadmap

This repo represents the first version of the tool. Planned
additions:

- One-click PDF report export
- Outlook email draft generation
- Auto-generated PowerPoint status deck from RAID data
- Plotly-powered visual dashboard upgrade
- Multi-project portfolio rollup
- Standalone executable distribution (no Python install required)
- Companion web-based dashboard via Streamlit or Dash

---

## About

Built by Gagan — PMP-certified project manager with five years
of project, reporting, and analytics experience, building the
tools I wished existed during my own PM work.

The full, packaged version of the tool will be available on
Gumroad. Watch this repo for release news.

---

## License

This codebase is shared for portfolio and educational purposes.
The compiled, sellable workbook is a separate licensed product.
See LICENSE.md for terms once published.