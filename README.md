# Farmer Income Reconciliation

Monthly and YTD gap analysis between **actual income from farmers** (paid for hours, days, or output) and **budgeted income**, broken down per agricultural school across the network.

Built for the finance/control function of *Adam Ve'Adama* — answers the recurring monthly question: *which schools are over- or under-performing against the income plan, and on what payment model?*

## What it does

For a target month (defaults to current month), the pipeline:

1. **Loads** the monthly invoices workbook (`מרכז חשבוניות חקלאים <month>.xlsx`) — one sheet per school.
2. **Identifies payment type** per row from column E:
   - **Hourly** (default) — paid per hour
   - **`יומי`** (daily) — paid per worker per day
   - **`תפוקה`** (output) — paid per unit produced (e.g. "דולב 130", "ק"ג")
3. **Computes per school × month**:
   - Hourly KPIs: rate per hour, hours per worker, workers per day
   - Income from non-hourly rows kept separate (so it doesn't distort hourly KPIs)
   - YTD aggregates — actuals computed item-by-item across all months; budget weighted by working days per month
4. **Compares to budget** (`תקציב.xlsx`): gap in absolute terms and as % of plan.
5. **Produces two Excel deliverables**:
   - **Per-school annual workbook** — one tab per month + YTD tab + non-hourly income table
   - **Network summary workbook** — `תצוגת על` overview tab listing all schools' KPIs side-by-side, plus a dedicated non-hourly-income tab

## Tech stack

- **Python 3.11**
- **openpyxl** — reads structured monthly Excel workbooks, writes styled multi-sheet output (with conditional formatting on gap cells: green for actual > budget, red for actual < budget, amber for piece-rate blocks)

## Project layout

```
analyze_monthly.py          Single-file pipeline (~1.5k LOC, fully documented)
נתונים מרוכזים/            Source data
  מרכז חשבוניות חקלאים *.xlsx   Monthly invoice workbooks (one per month, multiple school sheets each)
  תקציב.xlsx                     Annual budget per school × month
דוחות/                      Generated reports — one folder per run
  YYYY-MM-Month/
    <School>_YYYY.xlsx          Per-school annual workbook
    סיכום_רשת_YYYY_MM.xlsx      Network summary workbook
```

## Source format (per school sheet)

| Col | Header        | Notes                                                                 |
|-----|---------------|-----------------------------------------------------------------------|
| A   | תאריך          | Date — `DD-MMM`, datetime, or `YYYY-MM-DD HH:MM:SS`                  |
| B   | עובדים         | Workers — optional; usually present for daily, missing for output    |
| C   | שעות           | Hours — required for hourly rows only                                |
| D   | תשלום          | Payment in NIS                                                       |
| E   | אופן תשלום     | Optional — contains `תפוקה` or `יומי` to flag non-hourly rows        |

## Usage

```bash
python analyze_monthly.py              # current month
python analyze_monthly.py 2026 03      # specific year + month
```

Output lands in `דוחות/YYYY-MM-Month/`.

## Notes

- All UI, data and reports are in **Hebrew** (RTL).
- `BASE_DIR` at the top of the script is hardcoded — adjust before running elsewhere.
- Sheets named `פרטי חקלאים` or starting with `מרכז חשבוניות` are skipped (not school data).
- Network total row is weighted by working days per school (so each working day contributes equally regardless of which school it came from).

## Status

Production — runs monthly for finance/control at *Adam Ve'Adama*.

---

*Author: Dror Nadel · 2026*
