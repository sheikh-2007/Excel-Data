# 📊 Automated Customer Analytics & Sales Performance Dashboard (Excel VBA)

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Automation-blue?style=for-the-badge)
![BI](https://img.shields.io/badge/Analytics-Business_Intelligence-orange?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Production_Ready-brightgreen?style=for-the-badge)

An enterprise-grade, pixel-perfect executive dashboard built entirely programmatically in Microsoft Excel using Visual Basic for Applications (VBA). The macro extracts raw transactional customer and sales call records, constructs normalized PivotTables on a dedicated staging sheet, calculates high-level financial KPIs, and renders an interactive multi-visual analytical workspace equipped with dark-theme slicers.

---

## 📌 Executive Overview

Manual reporting in enterprise environments often leads to formula corruption, broken chart ranges, and slow refresh cycles. This project implements a **zero-manual-touch automation engine** that turns raw call logs into an executive-ready reporting suite with a single click.

* **Idempotent & Self-Healing:** Safely wipes previous staging and visualization layers before executing, preventing duplicate caches or name-scope collisions.
* **Pure Programmatic Modeling:** Direct object referencing (`Worksheets`, `ListObjects`, `PivotTables`, `Slicers`, `Shapes`) without relying on `.Select` or `.Activate`.
* **Zero Disruption to Raw Data:** Raw logs are converted to a structured Excel Table (`tbl_CustomerData`), with all aggregations decoupled on a background calculation sheet (`Pivot`).

---

## 🏗️ Architecture & Data Pipeline

```text
[ Raw Transaction Logs (Customer_Data) ]
                   │
                   ▼  (VBA Table Standardization)
          [ tbl_CustomerData ]
                   │
                   ▼  (Shared PivotCache Optimization)
           [ Staging Layer (Pivot Sheet) ]
    ┌──────────────┴────────────────────────────┐
    ▼                                           ▼
[ 5 Aggregated PivotTables ]           [ KPI Named Range Engine ]
- PT_Trend (Month × Year)              - KPI_TotalRevenue
- PT_Agent (Agent Leaderboard)         - KPI_TotalLeads
- PT_Region (Geographic Revenue)       - KPI_AvgRevLead
- PT_Product (Product/Service Mix)     - KPI_ConvertedLeads
- PT_Conversion (Conversion Split)     - KPI_ConversionRate
    │                                  - KPI_UniqueAgents
    └──────────────┬────────────────────────────┘
                   ▼
  [ Executive Workspace (Dashboard Sheet) ]
   ├── 6 Dynamic KPI Metric Cards (Programmatic Rounded Rectangles)
   ├── 3 Global Dark-Styled Slicers (Year, Customer Type, Product/Service)
   └── 5 Custom-Styled Synchronized Visuals (Line, Columns, Bar, Doughnut)
