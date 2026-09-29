## Daily Sales Tracking
Interactive HTML dashboard for daily commercial performance tracking, paired with the Python pipeline that consolidates and updates the underlying sales data.

Overview
This project automates a recurring commercial reporting workflow:

Data consolidation: the Python script (daily_sales_stratos.py) reads daily sales source files, merges them into the master workbook, applies lookup tables (ADVP range, segment, booking window), and updates the pivot-driven summary reports (Matriz Comitê, Proj. Diária) via Excel COM automation.
Visualization: the standalone HTML dashboard (Dashboard Daily Sales (In).html) renders the consolidated data as an interactive report: daily/MTD revenue vs. forecast vs. budget, passenger volume, ticket average, and day-over-day trends.
Dashboard
Open Dashboard Daily Sales (In).html directly in any browser, no server or build step required. All data and rendering logic are self-contained in the file.

---
Highlights:

MTD revenue vs. forecast vs. budget, with variance indicators
Daily sales curve with day-over-day comparison
Passenger volume (PAX) and average ticket tracking
Responsive layout, print-friendly
Automation script
daily_sales_stratos.py runs the daily data pipeline in three stages:

Stage	What it does
1. Análises Diárias	Reads new daily source files, deduplicates by date, appends to the master workbook, and recalculates lookup-driven columns
2. Matriz Comitê	Updates the BRL/USD committee matrix from the refreshed daily data
3. Proj. Diária	Updates the daily sales projection sheet, cross-checking totals against the source for data integrity
Built with pandas, openpyxl, and pywin32 (Excel COM automation) for surgical, formula-safe updates to large shared workbooks, preserving existing formatting, formulas, and pivot tables instead of regenerating the file from scratch.

Network paths are read from environment variables (SALES_BASE_PATH) rather than hardcoded, so the script can run against any equivalent folder structure.

Tech stack
Python · pandas · openpyxl · pywin32 · HTML/CSS/JavaScript (SVG-based charts, no external libraries)

Note on data
The dataset shown in the dashboard is illustrative and does not represent real financial figures.
