# Power BI Desktop Setup Guide — Meridian Precast Analytics

## 📋 Prerequisites

- **Power BI Desktop** (free download from [Microsoft](https://powerbi.microsoft.com/desktop))
- **Python 3.x** with `pandas` and `openpyxl` installed
- Source file: `Meridian Precast - Production Data.xlsx`

---

## 🚀 Quick Start

### Step 1: Generate the Data Package

```bash
pip install pandas openpyxl
python generate_powerbi_package.py
```

This creates a `powerbi_package/` folder with all required files.

### Step 2: Open Power BI Desktop

1. Launch **Power BI Desktop**
2. Go to **File → New** to start a fresh report

---

## 📥 Step 3: Import Data

### Option A: CSV Import (Recommended)

For each CSV file in `powerbi_package/`:

1. **Home → Get Data → Text/CSV**
2. Navigate to `powerbi_package/` folder
3. Import files in this order:
   - `dim_date.csv`
   - `dim_work_centres.csv`
   - `dim_products.csv`
   - `dim_jobs.csv`
   - `fact_labour_records.csv`
   - `violations_summary.csv`

### Option B: Power Query (Advanced)

1. **Home → Transform Data → Advanced Editor**
2. Open `power_query_m.pq` from the package
3. Copy the M script for each table
4. Update `FolderPath` variable to your actual folder path

---

## 🔗 Step 4: Create Relationships

Go to **Model View** (left sidebar, icon looks like interconnected boxes).

Create these relationships by dragging columns between tables:

| # | From Table | From Column | → | To Table | To Column | Cardinality |
|---|-----------|-------------|---|----------|-----------|-------------|
| 1 | fact_labour_records | job_number | → | dim_jobs | job_number | Many-to-One |
| 2 | fact_labour_records | product_code | → | dim_products | product_code | Many-to-One |
| 3 | fact_labour_records | work_centre | → | dim_work_centres | work_centre | Many-to-One |
| 4 | fact_labour_records | date | → | dim_date | date | Many-to-One |

### Star Schema Diagram

```
                    ┌──────────────┐
                    │   dim_date   │
                    │   (date)     │
                    └──────┬───────┘
                           │
┌──────────────┐   ┌───────┴────────┐   ┌──────────────────┐
│  dim_jobs    ├───┤  fact_labour   ├───┤ dim_work_centres │
│ (job_number) │   │   _records     │   │  (work_centre)   │
└──────────────┘   └───────┬────────┘   └──────────────────┘
                           │
                    ┌──────┴───────┐
                    │ dim_products │
                    │(product_code)│
                    └──────────────┘
```

---

## 📐 Step 5: Add DAX Measures

Open `dax_measures.dax` from the package. For each measure:

1. Click on the **fact_labour_records** table in the Fields pane
2. **Modeling → New Measure**
3. Paste the measure formula
4. Press **Enter** to confirm

### Key Measures to Add First

| Priority | Measure Name | Purpose |
|----------|-------------|---------|
| 1 | Total Records | Count of labour entries |
| 2 | Total Hours | Sum of hours worked |
| 3 | Total Labour Cost | Hours × Rate calculation |
| 4 | Total Units Scrapped | Scrap count |
| 5 | Scrap Rate | Scrapped / (Completed + Scrapped) |
| 6 | Missing Operator Count | Q1 data quality |
| 7 | Product Margin | Sell price − cost |
| 8 | Product Margin % | Margin as percentage |

---

## 📊 Step 6: Build Report Pages

### Page 1: Executive Summary

| Visual Type | Data Fields | Size/Position |
|-------------|-------------|---------------|
| **Card** | Total Records | Top row, 1st card |
| **Card** | Total Hours | Top row, 2nd card |
| **Card** | Total Labour Cost | Top row, 3rd card |
| **Card** | Total Rule Violations | Top row, 4th card |
| **Card** | Scrap Rate | Top row, 5th card |
| **Clustered Bar** | violations_summary[rule], violations_summary[count] | Middle left |
| **Table/Matrix** | violations_summary[question, rule, check, finding, severity] | Middle right |
| **Horizontal Bar** | operator_name (axis), Total Hours (value) | Bottom left |
| **Donut Chart** | dim_work_centres[name] (legend), Total Hours (value) | Bottom right |

**Slicer:** Add a slicer with `violations_summary[rule]` for interactive filtering.

### Page 2: Data Quality Audit

| Visual Type | Data Fields |
|-------------|-------------|
| **Table** | Q1: Records where flag_missing_operator = TRUE |
| **Table** | Q3: Overtime (operator_id, date, DailyHours > 8) |
| **Table** | Q4: Records where flag_has_scrap_no_reason = TRUE |
| **Table** | Q5: Records where flag_post_close_labour = TRUE |
| **Stacked Bar** | Overtime visualization (8h allowed + excess) |

### Page 3: Labour Cost Analysis

| Visual Type | Data Fields |
|-------------|-------------|
| **Card × 4** | Labour cost per work centre (use WC slicer) |
| **Clustered Bar** | Work centre name (axis), Labour Cost By WC (value) |
| **Donut** | Work centre hours distribution |
| **Matrix** | WC detail: name, records, hours, rate, cost, status |

**Format tip:** Use conditional formatting on the matrix to highlight WC-350's $0 rate in amber.

### Page 4: Product Profitability

| Visual Type | Data Fields |
|-------------|-------------|
| **Grouped Bar** | product_code (axis), cost_per_unit + sell_price_per_unit (values) |
| **Bar Chart** | product_code (axis), margin_dollars (value) — use conditional coloring |
| **Matrix** | Full product profitability table |
| **Card** | Loss Making Products count |

**Format tip:** Use rules-based conditional formatting: Red if margin < 0, Green if > 0.

### Page 5: Job Comparison (Q9)

| Visual Type | Data Fields |
|-------------|-------------|
| **Cards** | JOB-1006 Hours, JOB-1010 Hours, Hour Difference |
| **Grouped Bar** | stage (axis), hours filtered by JOB-1006 and JOB-1010 |
| **Radar/Rose** | Multi-metric comparison (hours, records, scrapped, days) |
| **Table** | Side-by-side job metrics |

---

## 🎨 Step 7: Apply Theme

Create a custom theme for the Power BI color palette. Save this as `meridian_theme.json`:

```json
{
    "name": "Meridian Precast",
    "dataColors": [
        "#118DFF", "#12B5CB", "#E66C37", "#6B8E23",
        "#D64045", "#7030A0", "#F2C811", "#004C6D"
    ],
    "background": "#F2F2F2",
    "foreground": "#252423",
    "tableAccent": "#118DFF",
    "visualStyles": {
        "*": {
            "*": {
                "general": [{
                    "responsive": true
                }]
            }
        }
    }
}
```

Apply via: **View → Themes → Browse for themes → select `meridian_theme.json`**

---

## ✅ Step 8: Add Slicers for Interactivity

Add these slicers to enable cross-filtering:

| Slicer Field | Type | Placement |
|--------------|------|-----------|
| dim_date[month_name] | Dropdown | Top of every page |
| dim_work_centres[name] | Buttons | Labour Cost page |
| dim_jobs[job_number] | Dropdown | Comparison page |
| dim_products[description] | Buttons | Products page |
| violations_summary[severity] | Buttons | Quality page |

---

## 📁 Package Contents Reference

```
powerbi_package/
├── fact_labour_records.csv    ← Main fact table (169 rows, enriched)
├── dim_jobs.csv               ← Jobs dimension (15 jobs)
├── dim_products.csv           ← Products with margins (5 products)
├── dim_work_centres.csv       ← Work centres with rates (4 centres)
├── dim_date.csv               ← Date dimension (Jun–Jul 2026)
├── violations_summary.csv     ← Rule violations (8 rules)
├── dax_measures.dax           ← 30+ ready-to-use DAX measures
├── power_query_m.pq           ← Power Query M scripts
└── relationships.json         ← Data model relationship definitions
```

---

## 💡 Tips & Best Practices

1. **Mark dim_date as a Date Table**: Right-click dim_date → Mark as date table → Select "date" column
2. **Sort month by number**: Click month_name column → Sort by Column → month
3. **Create a Measures table**: Modeling → New Table → `_Measures = { BLANK() }` — move all measures here for organization
4. **Use bookmarks**: Create bookmarks for "All Violations", "Failures Only", "Warnings Only" to add button navigation
5. **Enable Q&A**: Power BI's Q&A feature works well with this star schema — try asking "show hours by operator"
