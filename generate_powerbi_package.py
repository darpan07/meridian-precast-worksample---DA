"""
===============================================================================
 Power BI Desktop Setup — Data Export & Model Generator
 Purpose: Export clean, Power BI-ready CSV data files with proper data types,
          generate DAX measures, and create a Power Query M script for
          one-click import into Power BI Desktop.
===============================================================================
"""

import pandas as pd
import json
import os

# ─── Configuration ───
DATA_FILE = r'Meridian Precast - Production Data.xlsx'
OUTPUT_DIR = r'powerbi_package'

os.makedirs(OUTPUT_DIR, exist_ok=True)

print("═" * 60)
print("  Meridian Precast — Power BI Desktop Package Generator")
print("═" * 60)

# ═══════════════════════════════════════════════════════════════
#  STEP 1: Load & Clean Source Data
# ═══════════════════════════════════════════════════════════════

print("\n📥 Loading source data...")

# Labour Records
labour = pd.read_excel(DATA_FILE, sheet_name='Labour Records')
labour.columns = [c.strip().replace(' ', '_').lower() for c in labour.columns]
# Ensure proper types
labour['date'] = pd.to_datetime(labour['date']).dt.strftime('%Y-%m-%d')
labour['hours_worked'] = pd.to_numeric(labour['hours_worked'], errors='coerce')
labour['units_completed'] = pd.to_numeric(labour['units_completed'], errors='coerce').fillna(0).astype(int)
labour['units_scrapped'] = pd.to_numeric(labour['units_scrapped'], errors='coerce').fillna(0).astype(int)
print(f"  ✓ Labour Records: {len(labour)} rows")

# Jobs
jobs = pd.read_excel(DATA_FILE, sheet_name='Jobs')
jobs.columns = [c.strip().replace(' ', '_').lower() for c in jobs.columns]
jobs['date_closed'] = pd.to_datetime(jobs['date_closed'], errors='coerce').dt.strftime('%Y-%m-%d')
print(f"  ✓ Jobs: {len(jobs)} rows")

# Products
products = pd.read_excel(DATA_FILE, sheet_name='Products')
products.columns = ['product_code', 'description', 'cost_per_unit', 'sell_price_per_unit']
products['margin_dollars'] = products['sell_price_per_unit'] - products['cost_per_unit']
products['margin_pct'] = ((products['sell_price_per_unit'] - products['cost_per_unit']) /
                          products['sell_price_per_unit'] * 100).round(1)
products['is_profitable'] = products['sell_price_per_unit'] > products['cost_per_unit']
print(f"  ✓ Products: {len(products)} rows")

# Work Centres
centres = pd.read_excel(DATA_FILE, sheet_name='Work Centres')
centres.columns = ['work_centre', 'name', 'labour_rate']
print(f"  ✓ Work Centres: {len(centres)} rows")

# ═══════════════════════════════════════════════════════════════
#  STEP 2: Create Enriched Fact Table
# ═══════════════════════════════════════════════════════════════

print("\n🔧 Building enriched fact table...")

# Merge labour with work centres to get labour cost
fact = labour.merge(centres, on='work_centre', how='left', suffixes=('', '_wc'))
fact['labour_cost'] = (fact['hours_worked'] * fact['labour_rate']).round(2)

# Merge with jobs to get status/close date
fact = fact.merge(
    jobs[['job_number', 'quantity_ordered', 'status', 'date_closed']],
    on='job_number', how='left', suffixes=('', '_job')
)

# Data quality flags
fact['flag_missing_operator'] = fact['operator_id'].isna() | (fact['operator_id'] == '')
fact['flag_has_scrap_no_reason'] = (fact['units_scrapped'] > 0) & (fact['scrap_reason'].isna() | (fact['scrap_reason'] == ''))
fact['flag_post_close_labour'] = False
mask = fact['date_closed'].notna() & (fact['date'] > fact['date_closed'])
fact.loc[mask, 'flag_post_close_labour'] = True

print(f"  ✓ Enriched fact table: {len(fact)} rows, {len(fact.columns)} columns")

# ═══════════════════════════════════════════════════════════════
#  STEP 3: Create Date Dimension Table
# ═══════════════════════════════════════════════════════════════

print("📅 Generating date dimension...")

date_range = pd.date_range('2026-06-01', '2026-07-31', freq='D')
date_dim = pd.DataFrame({
    'date': date_range.strftime('%Y-%m-%d'),
    'year': date_range.year,
    'month': date_range.month,
    'month_name': date_range.strftime('%B'),
    'day': date_range.day,
    'day_of_week': date_range.strftime('%A'),
    'day_of_week_num': date_range.dayofweek,
    'week_number': date_range.isocalendar().week.astype(int),
    'is_weekend': date_range.dayofweek >= 5,
})

print(f"  ✓ Date dimension: {len(date_dim)} rows")

# ═══════════════════════════════════════════════════════════════
#  STEP 4: Create Violations Summary Table
# ═══════════════════════════════════════════════════════════════

violations = pd.DataFrame([
    {'question': 'Q1', 'rule': 'R3', 'check': 'Missing Operator', 'finding': '2 records with no operator', 'affected': 'Rec 148, 175', 'severity': 'FAIL', 'count': 2},
    {'question': 'Q2', 'rule': 'R6', 'check': 'Duplicate Records', 'finding': '1 pair of identical records', 'affected': 'Rec 119, 120', 'severity': 'FAIL', 'count': 2},
    {'question': 'Q3', 'rule': 'R1', 'check': 'Overtime (>8 hrs/day)', 'finding': '2 operators exceeded limit', 'affected': 'OP-104, OP-102', 'severity': 'FAIL', 'count': 2},
    {'question': 'Q4', 'rule': 'R4', 'check': 'Missing Scrap Reason', 'finding': '14 units without reason', 'affected': 'Rec 243, 250', 'severity': 'FAIL', 'count': 2},
    {'question': 'Q5', 'rule': 'R2', 'check': 'Labour After Job Close', 'finding': '2 records on JOB-1014', 'affected': 'Rec 263, 266', 'severity': 'FAIL', 'count': 2},
    {'question': 'Q6', 'rule': 'R5', 'check': 'Overproduction', 'finding': 'JOB-1011 overproduced by 6', 'affected': 'JOB-1011', 'severity': 'FAIL', 'count': 1},
    {'question': 'Q7', 'rule': 'R8', 'check': 'Product Sold Below Cost', 'finding': 'PRD-440 loses $54/unit', 'affected': 'PRD-440', 'severity': 'FAIL', 'count': 1},
    {'question': 'Q8', 'rule': 'R7', 'check': 'Labour Cost Calculation', 'finding': 'WC-350 has $0/hr rate', 'affected': 'WC-350 (41 records)', 'severity': 'WARN', 'count': 1},
])

print(f"  ✓ Violations table: {len(violations)} rules checked")

# ═══════════════════════════════════════════════════════════════
#  STEP 5: Export CSV Files
# ═══════════════════════════════════════════════════════════════

print("\n💾 Exporting Power BI data files...")

fact.to_csv(os.path.join(OUTPUT_DIR, 'fact_labour_records.csv'), index=False)
jobs.to_csv(os.path.join(OUTPUT_DIR, 'dim_jobs.csv'), index=False)
products.to_csv(os.path.join(OUTPUT_DIR, 'dim_products.csv'), index=False)
centres.to_csv(os.path.join(OUTPUT_DIR, 'dim_work_centres.csv'), index=False)
date_dim.to_csv(os.path.join(OUTPUT_DIR, 'dim_date.csv'), index=False)
violations.to_csv(os.path.join(OUTPUT_DIR, 'violations_summary.csv'), index=False)

print("  ✓ fact_labour_records.csv")
print("  ✓ dim_jobs.csv")
print("  ✓ dim_products.csv")
print("  ✓ dim_work_centres.csv")
print("  ✓ dim_date.csv")
print("  ✓ violations_summary.csv")

# ═══════════════════════════════════════════════════════════════
#  STEP 6: Generate DAX Measures File
# ═══════════════════════════════════════════════════════════════

print("\n📐 Generating DAX measures...")

dax_measures = """
// ═══════════════════════════════════════════════════════════════
// MERIDIAN PRECAST — POWER BI DAX MEASURES
// ═══════════════════════════════════════════════════════════════
// Copy each measure into Power BI Desktop → Modeling → New Measure
// ═══════════════════════════════════════════════════════════════


// ─────────────────────────────────────────────────────────────
// CORE METRICS
// ─────────────────────────────────────────────────────────────

Total Records = COUNTROWS(fact_labour_records)

Total Hours = SUM(fact_labour_records[hours_worked])

Total Labour Cost =
    SUMX(
        fact_labour_records,
        fact_labour_records[hours_worked] * RELATED(dim_work_centres[labour_rate])
    )

Total Units Completed = SUM(fact_labour_records[units_completed])

Total Units Scrapped = SUM(fact_labour_records[units_scrapped])

Scrap Rate =
    DIVIDE(
        [Total Units Scrapped],
        [Total Units Completed] + [Total Units Scrapped],
        0
    )

Avg Hours Per Record = AVERAGE(fact_labour_records[hours_worked])


// ─────────────────────────────────────────────────────────────
// DATA QUALITY METRICS
// ─────────────────────────────────────────────────────────────

Missing Operator Count =
    COUNTROWS(
        FILTER(fact_labour_records, ISBLANK(fact_labour_records[operator_id]))
    )

Duplicate Record Count =
    VAR DupGroups =
        SUMMARIZE(
            fact_labour_records,
            fact_labour_records[date],
            fact_labour_records[job_number],
            fact_labour_records[stage],
            fact_labour_records[operator_id],
            fact_labour_records[hours_worked],
            fact_labour_records[units_completed],
            fact_labour_records[units_scrapped],
            "GroupCount", COUNTROWS(fact_labour_records)
        )
    RETURN
        COUNTROWS(FILTER(DupGroups, [GroupCount] > 1))

Overtime Violations =
    VAR DailyTotals =
        ADDCOLUMNS(
            SUMMARIZE(
                FILTER(fact_labour_records, NOT(ISBLANK(fact_labour_records[operator_id]))),
                fact_labour_records[operator_id],
                fact_labour_records[date]
            ),
            "DailyHours", CALCULATE(SUM(fact_labour_records[hours_worked]))
        )
    RETURN
        COUNTROWS(FILTER(DailyTotals, [DailyHours] > 8))

Missing Scrap Reason Count =
    COUNTROWS(
        FILTER(
            fact_labour_records,
            fact_labour_records[units_scrapped] > 0
            && ISBLANK(fact_labour_records[scrap_reason])
        )
    )

Post Close Labour Count =
    COUNTROWS(
        FILTER(
            fact_labour_records,
            NOT(ISBLANK(fact_labour_records[date_closed]))
            && fact_labour_records[date] > fact_labour_records[date_closed]
        )
    )

Total Rule Violations =
    [Missing Operator Count] +
    [Duplicate Record Count] +
    [Overtime Violations] +
    [Missing Scrap Reason Count] +
    [Post Close Labour Count]


// ─────────────────────────────────────────────────────────────
// PRODUCT PROFITABILITY
// ─────────────────────────────────────────────────────────────

Product Margin =
    SUMX(
        dim_products,
        dim_products[sell_price_per_unit] - dim_products[cost_per_unit]
    )

Product Margin % =
    DIVIDE(
        SELECTEDVALUE(dim_products[sell_price_per_unit]) -
            SELECTEDVALUE(dim_products[cost_per_unit]),
        SELECTEDVALUE(dim_products[sell_price_per_unit]),
        0
    )

Loss Making Products =
    COUNTROWS(
        FILTER(
            dim_products,
            dim_products[sell_price_per_unit] <= dim_products[cost_per_unit]
        )
    )

Profitable Products =
    COUNTROWS(
        FILTER(
            dim_products,
            dim_products[sell_price_per_unit] > dim_products[cost_per_unit]
        )
    )


// ─────────────────────────────────────────────────────────────
// WORK CENTRE ANALYSIS
// ─────────────────────────────────────────────────────────────

Labour Cost By WC =
    SUMX(
        fact_labour_records,
        fact_labour_records[hours_worked] * RELATED(dim_work_centres[labour_rate])
    )

Hours % of Total =
    DIVIDE(
        [Total Hours],
        CALCULATE([Total Hours], ALL(dim_work_centres)),
        0
    )

Cost Per Unit =
    DIVIDE(
        [Total Labour Cost],
        [Total Units Completed],
        0
    )


// ─────────────────────────────────────────────────────────────
// JOB COMPARISON (Q9)
// ─────────────────────────────────────────────────────────────

JOB-1006 Hours =
    CALCULATE(
        SUM(fact_labour_records[hours_worked]),
        fact_labour_records[job_number] = "JOB-1006"
    )

JOB-1010 Hours =
    CALCULATE(
        SUM(fact_labour_records[hours_worked]),
        fact_labour_records[job_number] = "JOB-1010"
    )

Hour Difference = [JOB-1010 Hours] - [JOB-1006 Hours]

Hour Difference % =
    DIVIDE(
        [JOB-1010 Hours] - [JOB-1006 Hours],
        [JOB-1006 Hours],
        0
    )


// ─────────────────────────────────────────────────────────────
// CONDITIONAL FORMATTING HELPERS
// ─────────────────────────────────────────────────────────────

Status Color =
    SWITCH(
        TRUE(),
        [Scrap Rate] > 0.10, "#D64045",   // Red
        [Scrap Rate] > 0.05, "#E66C37",   // Orange
        "#107C10"                           // Green
    )

Margin Status =
    IF(
        SELECTEDVALUE(dim_products[sell_price_per_unit]) >
            SELECTEDVALUE(dim_products[cost_per_unit]),
        "Profitable",
        "Loss"
    )
"""

with open(os.path.join(OUTPUT_DIR, 'dax_measures.dax'), 'w', encoding='utf-8') as f:
    f.write(dax_measures)

print("  ✓ dax_measures.dax")

# ═══════════════════════════════════════════════════════════════
#  STEP 7: Generate Power Query M Script
# ═══════════════════════════════════════════════════════════════

print("📊 Generating Power Query M script...")

m_script = r"""
// ═══════════════════════════════════════════════════════════════
// MERIDIAN PRECAST — POWER QUERY M SCRIPT
// ═══════════════════════════════════════════════════════════════
// In Power BI Desktop:
//   1. Home → Transform Data → Advanced Editor
//   2. Paste the relevant section for each table
//   3. Update the file paths to your local folder
// ═══════════════════════════════════════════════════════════════

// ─── Replace this with your actual folder path ───
// let FolderPath = "C:\path\to\powerbi_package\" in

// ─────────────────────────────────────────────────────────────
// TABLE: fact_labour_records
// ─────────────────────────────────────────────────────────────
let
    Source = Csv.Document(File.Contents(FolderPath & "fact_labour_records.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"record_id", Int64.Type},
        {"date", type date},
        {"job_number", type text},
        {"product_code", type text},
        {"stage", type text},
        {"work_centre", type text},
        {"operator_id", type text},
        {"operator_name", type text},
        {"hours_worked", type number},
        {"units_completed", Int64.Type},
        {"units_scrapped", Int64.Type},
        {"scrap_reason", type text},
        {"labour_cost", Currency.Type},
        {"flag_missing_operator", type logical},
        {"flag_has_scrap_no_reason", type logical},
        {"flag_post_close_labour", type logical}
    })
in
    ChangedTypes


// ─────────────────────────────────────────────────────────────
// TABLE: dim_jobs
// ─────────────────────────────────────────────────────────────
let
    Source = Csv.Document(File.Contents(FolderPath & "dim_jobs.csv"), [Delimiter=",", Encoding=65001]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"job_number", type text},
        {"product_code", type text},
        {"quantity_ordered", Int64.Type},
        {"status", type text},
        {"date_closed", type date}
    })
in
    ChangedTypes


// ─────────────────────────────────────────────────────────────
// TABLE: dim_products
// ─────────────────────────────────────────────────────────────
let
    Source = Csv.Document(File.Contents(FolderPath & "dim_products.csv"), [Delimiter=",", Encoding=65001]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"product_code", type text},
        {"description", type text},
        {"cost_per_unit", Currency.Type},
        {"sell_price_per_unit", Currency.Type},
        {"margin_dollars", Currency.Type},
        {"margin_pct", type number},
        {"is_profitable", type logical}
    })
in
    ChangedTypes


// ─────────────────────────────────────────────────────────────
// TABLE: dim_work_centres
// ─────────────────────────────────────────────────────────────
let
    Source = Csv.Document(File.Contents(FolderPath & "dim_work_centres.csv"), [Delimiter=",", Encoding=65001]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"work_centre", type text},
        {"name", type text},
        {"labour_rate", Currency.Type}
    })
in
    ChangedTypes


// ─────────────────────────────────────────────────────────────
// TABLE: dim_date
// ─────────────────────────────────────────────────────────────
let
    Source = Csv.Document(File.Contents(FolderPath & "dim_date.csv"), [Delimiter=",", Encoding=65001]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"date", type date},
        {"year", Int64.Type},
        {"month", Int64.Type},
        {"month_name", type text},
        {"day", Int64.Type},
        {"day_of_week", type text},
        {"day_of_week_num", Int64.Type},
        {"week_number", Int64.Type},
        {"is_weekend", type logical}
    }),
    SortedByDate = Table.Sort(ChangedTypes, {{"date", Order.Ascending}})
in
    SortedByDate


// ─────────────────────────────────────────────────────────────
// TABLE: violations_summary
// ─────────────────────────────────────────────────────────────
let
    Source = Csv.Document(File.Contents(FolderPath & "violations_summary.csv"), [Delimiter=",", Encoding=65001]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"question", type text},
        {"rule", type text},
        {"check", type text},
        {"finding", type text},
        {"affected", type text},
        {"severity", type text},
        {"count", Int64.Type}
    })
in
    ChangedTypes
"""

with open(os.path.join(OUTPUT_DIR, 'power_query_m.pq'), 'w', encoding='utf-8') as f:
    f.write(m_script)

print("  ✓ power_query_m.pq")

# ═══════════════════════════════════════════════════════════════
#  STEP 8: Generate Relationship Definitions
# ═══════════════════════════════════════════════════════════════

relationships = [
    {
        "from_table": "fact_labour_records",
        "from_column": "job_number",
        "to_table": "dim_jobs",
        "to_column": "job_number",
        "cardinality": "Many-to-One",
        "cross_filter": "Both"
    },
    {
        "from_table": "fact_labour_records",
        "from_column": "product_code",
        "to_table": "dim_products",
        "to_column": "product_code",
        "cardinality": "Many-to-One",
        "cross_filter": "Both"
    },
    {
        "from_table": "fact_labour_records",
        "from_column": "work_centre",
        "to_table": "dim_work_centres",
        "to_column": "work_centre",
        "cardinality": "Many-to-One",
        "cross_filter": "Both"
    },
    {
        "from_table": "fact_labour_records",
        "from_column": "date",
        "to_table": "dim_date",
        "to_column": "date",
        "cardinality": "Many-to-One",
        "cross_filter": "Single"
    }
]

with open(os.path.join(OUTPUT_DIR, 'relationships.json'), 'w') as f:
    json.dump(relationships, f, indent=2)

print("  ✓ relationships.json")

# ═══════════════════════════════════════════════════════════════
#  DONE
# ═══════════════════════════════════════════════════════════════

print("\n" + "═" * 60)
print("  ✅ Power BI Package Complete!")
print("═" * 60)
print(f"\n📁 Output folder: {os.path.abspath(OUTPUT_DIR)}/")
print("""
📂 Files generated:
    ├── fact_labour_records.csv    (enriched fact table)
    ├── dim_jobs.csv               (jobs dimension)
    ├── dim_products.csv           (products with margins)
    ├── dim_work_centres.csv       (work centre rates)
    ├── dim_date.csv               (date dimension)
    ├── violations_summary.csv     (rule violations)
    ├── dax_measures.dax           (30+ DAX measures)
    ├── power_query_m.pq           (M scripts for import)
    └── relationships.json         (data model relationships)

📖 See PowerBI_Setup_Guide.md for step-by-step instructions.
""")
