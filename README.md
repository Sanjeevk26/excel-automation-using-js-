# Excel SFT Dashboard Automation (Office.js / Script Lab)

This repo contains an **Office.js Excel Script (Script Lab)** that transforms a raw dataset into a formatted **Dashboard** with:
- KPI calculations (Revenue, Margin, Trends, YoY deltas)
- Margin health classification + conditional formatting
- A “Chart Data” table
- A combo chart (clustered columns for margins + line on secondary axis for total revenue)

It’s designed for the workflow: **Input workbook → run script → Final dashboard workbook**.

---

## What this script builds

### On the `Dashboard` sheet
**Rows 8–39** (product x quarter grid)
- **Total Revenue** (SUMIFS from Raw Data)
- **Weighted Avg Margin**
- **Rolling Trend (Quarter-over-Quarter delta)**
- **Year-over-Year Margin Delta**
- **Margin Health**: Strong / Moderate / At Risk
- Conditional formatting applied to Margin Health

**Rows 42–51**
- “Chart Data” table:
  - Quarter (2023 Q1 → 2024 Q4)
  - Margin by product
  - Total Revenue by quarter (computed via JS from Raw Data)

**Chart**
- Clustered columns for product margin series
- “Total Revenue” plotted as a line on the secondary axis

> Note: The script intentionally excludes **2023 Q1** from chart plotting (sets the first data row to `#N/A` for series columns).

---

## Workbook requirements (must match)

Your workbook must contain these worksheets:
- `Raw Data`
- `Dashboard`

### Raw Data required columns (header row)
The script expects these exact headers (case-insensitive trim helps, but spelling must match):
- `Year`
- `Quarter`
- `Revenue`

And the dashboard formulas expect Raw Data fields in these columns (as used by the formulas):
- **B** = Year
- **C** = Quarter
- **D** = Product
- **E** = Revenue
- **G** = Margin numerator / margin value used in weighted calc

If your Raw Data layout differs, adjust the formula ranges in the script (see “Customize” below).

---

## How to run (Script Lab)

1. Open the workbook in **Excel Desktop** (recommended) or Excel Web.
2. Install **Script Lab** add-in (Home → Add-ins → search “Script Lab”).
3. Open Script Lab → **Code**.
4. Paste the script from:  
   `script/excel-sft-dashboard.ts`  *(or copy from this README’s code section in the repo)*
5. Click **Run**.

The script will populate formulas, apply formatting, generate the chart, and refresh the dashboard.

---

