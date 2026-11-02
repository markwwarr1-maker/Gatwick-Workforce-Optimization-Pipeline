# Real-Time Operational Command & Workforce Optimization Pipeline

## ✈️ Executive Project Overview
This production-grade data engineering project reconstructs a **Real-Time Operational Command Suite** designed to solve a complex workforce asset bottleneck at Gatwick Airport's Immigration Control. 

As an Operations Leader responsible for a **500+ rostered workforce**, I designed this automated data solution to bridge the gap between volatile stakeholder demands (unpredictable passenger inflows) and highly constrained resource variables (sickness reporting, sudden emergency leave, and rolling training demands). 

By migrating legacy workforce tracking from fragmented files into a structured, automated **ETL (Extract, Transform, Load)** pipeline, this tool replaces manual resource scheduling with high-velocity data modeling.

---

## 🛠️ The Tech Stack & Architecture
* **Data Engineering Engine:** Excel Power Query (M-Code Language)
* **Automation UI Layer:** Excel VBA Macros
* **Reporting Architecture:** Top-Down Geographic Pivot Matrices & Cross-Functional Data Modeling
* **Data Schema Protocol:** Relational Star-Schema Design Patterns (`fct` and `dim` layers)

[ Ingestion Gateway ] ───> [ ETL Cleaning Layers ] ───> [ Unified Reporting Layer ]
• Raw TAM (Kronos) Logs     • Automated .xlsx filters     • Dynamic Multi-Key Joins
• SharePoint Absence Logs    • Legacy Metric Dropping      • Fractional Asset Allocations
• Deduplication Protocols     • Interactive Slicers & VBA UI

---

## 📊 Core Data Engineering Steps (The ETL Pipeline)

### 1. Ingestion Gateway (Folder Consolidation)
* Engineered a dynamic folder ingestion connector pointing to the raw **TAM (Kronos)** database dump directory (`01_Raw_TAM_Exports`).
* Implemented a restrictive metadata extension filter (`.xlsx` gate) directly within the ingestion node to ensure system stability and protect the pipeline against background file debris.

### 2. Dimensionality Reduction & Data Minimization
* Identified and stripped away redundant legacy metrics (`Scheduled Hours`, `Job Covered`, `Transfer Job`) from the raw data streams during the ingestion stage.
* This proactive data minimization step optimized local memory footprints and reduced internal processing latency across the pipeline.

### 3. Chronological Consistency & Deduplication
* Handled the overlapping date duplications inherent to continuous fortnightly payroll downloads by building a composite key deduplication layer.
* Enforced structural data integrity by executing multi-column deduplication across `Staff ID` + `Shift Start Date` + `Shift Start Time`.

### 4. Cross-Functional Data Merging
* Executed an advanced **Left Outer Merge** linking the active operational roster with a separate live **SharePoint Leave & Absence Tracker** query.
* Conducted a highly precise **Multi-Key Join** executing simultaneous matches across both employee identity (`Staff_ID`) and the chronological day (`Shift_Start_Date` = `Log_Date`).

### 5. Fractional Asset Logic (Conditional Modeling)
* Authored clean, multi-clause conditional logic to evaluate real-time workforce availability variables:
  * **Sickness / Emergency Leave / Annual Leave:** Evaluated down to `0` (Resource Unavailable).
  * **Half-Day Annual Leave:** Modeled as a fractional asset value of `0.5` (Resource available for exactly 50% of the operational window).
  * **Active Rostered Duty:** Evaluated as `1` (Resource fully available).

---

## 🏢 Translating Data into Leadership Outcomes

This project explicitly highlights advanced workforce asset optimization and inclusive team management practices:

* **Top-Down Geographic Hierarchy:** Designed the presentation canvas to group rows by **Assigned Post** first and **Operational Role** second. This maps perfectly to tactical command structures, allowing a Duty Manager to evaluate terminal readiness at a glance.
* **Granular Grade Enforcement:** Successfully embedded strict public-sector organizational hierarchies, properly segregating tactical commanders (SEOs overseeing Early Support, Early Duty Officer, and Call rotations) from core managers (HEOs) and frontline personnel (EOs/AOs/AAs).
* **Automated Data Governance:** Implemented **Worksheet Protection Protocols** on the underlying fact data layer (`fct_Workforce_Availability`) to safeguard structural data integrity while keeping it visible to technical stakeholders for audit transparency.

---

## 💻 Technical Demonstration (The Automation Script)
To deliver a true single-click update interface for busy operations directors, the entire background Power Query pipeline is wrapped inside a customized VBA macro layer attached to a front-facing action control element:

```vba
Sub Refresh_Gatwick_Pipeline()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Refreshing Gatwick Operational Pipeline..."
    
    ThisWorkbook.RefreshAll
    DoEvents 
    
    Dim ws As Worksheet, pt As PivotTable
    Set ws = ThisWorkbook.Sheets("Dashboard")
    For Each pt In ws.PivotTables: pt.RefreshTable: Next pt
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Pipeline successfully updated!", vbInformation, "ETL Success"
End Sub
```