# Meridian Precast — Data Analyst Work Sample

## 📁 Deliverable Files

| # | File | Tool | Purpose |
|---|------|------|---------|
| 1 | **`Work Sample Answers.md`** | — | Final answers to all 9 questions |
| 2 | **`analysis.py`** | Python | Automated analysis with charts & CSV exports |
| 3 | **`analysis.sql`** | SQL | All 9 questions as SQL queries (SQLite) |
| 4 | **`Meridian Precast - Analysis Workbook.xlsx`** | Excel | Enhanced workbook with Q1-Q9 tabs & charts |
| 5 | **`dashboard.html`** | HTML/JS | Interactive data quality dashboard |
| 6 | **`load_to_sqlite.py`** | Python | Helper: loads Excel → SQLite database |
| 7 | **`analysis_output/`** | — | Generated CSV files and chart images |
| 8 | **`powerbi_dashboard.html`** | HTML/JS | Power BI-themed interactive dashboard |
| 9 | **`generate_powerbi_package.py`** | Python | Generates Power BI Desktop data package |
| 10 | **`PowerBI_Setup_Guide.md`** | — | Step-by-step Power BI Desktop setup guide |
| 11 | **`powerbi_package/`** | — | CSV data, DAX measures, Power Query M scripts |

## 🚀 How to Run

### Python Analysis
```bash
pip install pandas openpyxl matplotlib seaborn
python analysis.py
```

### SQL Queries
```bash
python load_to_sqlite.py        # Creates meridian_precast.db
sqlite3 meridian_precast.db < analysis.sql
```

### Dashboard
Simply open `dashboard.html` in any modern web browser.

### Power BI Dashboard (Web Version)
Open `powerbi_dashboard.html` in any modern browser — fully interactive with slicers, cross-filtering, and 5 report pages.

### Power BI Desktop
```bash
python generate_powerbi_package.py   # Creates powerbi_package/ folder
```
Then follow `PowerBI_Setup_Guide.md` to import into Power BI Desktop.

### Excel Workbook
Open `Meridian Precast - Analysis Workbook.xlsx` in Excel.

## 📊 Summary of Findings

| Question | Rule | Issue Found | Details |
|----------|------|------------|---------|
| Q1 | R3 | ✅ Yes — 2 records | Record 148, 175 missing operator |
| Q2 | R6 | ✅ Yes — 1 pair | Records 119, 120 are duplicates |
| Q3 | R1 | ✅ Yes — 2 instances | OP-104 (11.92h), OP-102 (11.75h) |
| Q4 | R4 | ✅ Yes — 2 records | 14 units scrapped without reason |
| Q5 | R2 | ✅ Yes — 2 records | JOB-1014 has post-close labour |
| Q6 | R5 | ✅ Yes — 1 job | JOB-1011 overproduced by 6 units |
| Q7 | R8 | ✅ Yes — 1 product | PRD-440 sold at $54 loss/unit |
| Q8 | R7 | ⚠️ Concern | WC-350 Finishing at $0/hr |
| Q9 | — | 59.32h difference | OP-106 (trainee?) on JOB-1010 |
