# Blinkit Sales Dashboard — Excel Data Analysis Project
## Overview

This project simulates a real-world retail analytics task: taking a messy grocery sales export and transforming it into a decision-ready dashboard, using only native Microsoft Excel tools (formulas, PivotTables, PivotCharts, conditional formatting, and slicers).

## Dataset

- **Source sheet:** `Blinkit Grocery Data`
- **Rows:** 8,523 item-level sales records
- **Columns:** Item Identifier, Item Fat Content, Item Type, Outlet Establishment Year, Outlet Identifier, Outlet Location Type, Outlet Size, Outlet Type, Item Visibility, Item Weight, Sales, Rating

## Data Cleaning

| Issue | Fix Applied |
|---|---|
| Inconsistent text casing in `Item Fat Content` (`"low fat"` vs `"Low Fat"`) | Standardized all entries to `"Low Fat"` using Find & Replace (match case) |
| 1,463 missing values in `Item Weight` | Imputed using the **average weight for that item's Item Type**, calculated via an array formula (`AVERAGE(IF(...))`) and applied with `VLOOKUP` |

No duplicate rows or other missing values were found elsewhere in the dataset.

## Analysis: PivotTables & Charts

Five PivotTables were built, each on its own worksheet, paired with a chart:

| Pivot Sheet | Chart Type | Insight |
|---|---|---|
| `Pivot - Outlet Type` | Clustered Bar | Total sales by outlet type (Grocery Store vs. Supermarket Type 1/2/3) |
| `Pivot - Fat Content` | Pie / Doughnut | Sales share between Low Fat and Regular products |
| `Pivot - Year Trend` | Line with Markers | Sales trend across outlet establishment years |
| `Pivot - Item Type` | Horizontal Bar | Total sales across all 16 item categories |
| `Pivot - Outlet Size` | Pie / Doughnut | Sales distribution across Small, Medium, and High outlet sizes |

## Dashboard

A single-page `Dashboard` sheet consolidates the analysis, styled to match Blinkit's brand palette (black background, yellow accents):

- **KPI cards:** Total Revenue, Average Sales per Item, Total Items Sold, Average Rating
- **5 linked charts** (listed above), all connected to live PivotTable data
- **4 interactive slicers:** Outlet Location Type, Outlet Size, Item Fat Content, Outlet Establishment Year — each connected to all 5 PivotTables for synchronized cross-filtering
- **Navigation buttons** linking to the raw data sheet and back

## Tools & Techniques Used

- Excel formulas: `VLOOKUP`, `SUMIFS`, array formulas (`AVERAGE(IF(...))` with Ctrl+Shift+Enter), `COUNTA`
- PivotTables & PivotCharts
- Slicers with multi-pivot report connections
- Conditional formatting
- Custom cell/shape styling and hyperlinked navigation shapes

## File

- `Blinkit.xlsx` — the complete workbook (raw data, 5 pivot sheets, and dashboard)

## How to Use

1. Open `Blinkit.xlsx` in Microsoft Excel (desktop recommended — some slicer/PivotTable features may not render fully in Excel Online or Google Sheets).
2. Start on the `Dashboard` sheet.
3. Use the slicers at the top to filter all charts simultaneously by location tier, outlet size, fat content, or establishment year.
4. Use the "View Raw Data" button to inspect the underlying cleaned dataset.

## Author

Animesh Deshmukh
