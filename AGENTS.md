# AGENTS.md

## Project scope
This repository is a GPS vs MIS fleet reconciliation dashboard built in Python with Streamlit. The work here centers on operational data quality, trip comparison, KPI reporting, and dashboarding for vehicle, route, site, and shift performance.

## Core principles
- Prefer clean, auditable data processing over opaque transformations.
- Treat missing GPS/MIS records as operational signals, not as harmless blanks.
- Normalize vehicle IDs, route names, dates, and site labels before comparison.
- Keep business rules explicit and traceable.
- Surface exceptions and critical mismatches clearly in outputs.

## Standard operating workflow
1. Validate source files before computing KPIs.
2. Standardize identifiers and date formats.
3. Merge GPS and MIS data on business keys such as Vehicle + Date + Route + Shift.
4. Record missing, duplicate, and unmatched rows separately.
5. Compute difference metrics and apply classification logic.
6. Generate KPI summaries and exception reports.
7. Present results in dashboard-friendly tables and charts.

## Business rules
- Difference <= 2 KM: Match
- Difference 3–10 KM: Review
- Difference > 10 KM: Exception
- GPS Missing: Critical
- MIS Missing: Critical

## Data quality expectations
- Never silently drop invalid rows without logging the reason.
- Flag duplicates distinctly from valid trip records.
- Keep route and site names canonical to avoid false mismatches.
- Keep all KPI calculations reproducible from the cleaned master dataset.

## Preferred tooling
- Python for ETL and analytics
- pandas for transformations and aggregations
- Streamlit for dashboard UX and filters
- Plotly for KPI and comparison visualization
- Excel/CSV/PDF export for operational reporting

## Implementation expectations
- Use named functions for repeated transformation logic.
- Favor readable code over clever one-liners.
- Include robust handling for empty inputs, missing columns, and null values.
- Ensure dashboards expose business filters such as Date, Month, Site, Client, Vehicle, Route, and Shift.

## Deliverable expectations
Every analytics update should include:
- a summary of the data issue or business question
- the cleaning/standardization applied
- the reconciliation result or KPI outcome
- any critical exceptions or anomalies found
- recommended next action for operational review

## Use this repo for
- GPS/MIS comparison logic
- fleet reconciliation dashboards
- KPI monitoring and route variance analysis
- exception reporting and operational investigations

## Do not use this repo for
- unrelated product or web app work
- non-fleet datasets without explicit business-context alignment
- hidden or unexplained rule changes without documentation
