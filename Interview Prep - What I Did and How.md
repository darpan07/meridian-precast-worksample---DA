# 🎯 Interview Prep — Everything I Did in This Project

**Read this before your interview. It explains what the project is, what I did, how I did it, and why — all in plain English.**

---

## 📌 What Is This Project?

Meridian Precast is a company that makes precast concrete products (wall panels, manhole barrels, paving slabs, retaining blocks). They gave me an Excel file with **production data** — labour records, jobs, products, and work centres — and asked me to **find data quality problems** by checking it against 8 business rules.

The work sample had **9 questions** I needed to answer. Each question is basically asking: "Does this data break this rule? If so, show me the proof."

### The Data (4 Sheets in the Excel File)

| Sheet | What It Contains | Rows |
|-------|-----------------|------|
| **Labour Records** | Every time a worker logged hours on a job — who did it, what stage, how many hours, units made, units scrapped | 169 rows |
| **Jobs** | List of production jobs — what product, how many ordered, status (open/closed), close date | 15 jobs |
| **Products** | The 5 products they make — cost to make and selling price | 5 products |
| **Work Centres** | The 4 stages of production — Mixing, Casting, Curing, Finishing — and the hourly labour rate for each | 4 centres |

### The 8 Business Rules I Was Checking

| Rule | What It Says (Simple English) |
|------|-------------------------------|
| **R1** | Nobody should work more than 8 hours in one day |
| **R2** | Don't log work on a job after it's been closed |
| **R3** | Every labour record must say who did the work |
| **R4** | If you scrap units, you must give a reason |
| **R5** | Don't finish more units than were ordered |
| **R6** | Don't enter the same record twice |
| **R7** | Labour cost = hours × hourly rate for that work centre |
| **R8** | Every product must sell for more than it costs to make |

---

## 🛠️ What I Built (All the Deliverables)

I didn't just answer the 9 questions — I built **5 different deliverables** to show I can work across multiple tools:

### 1. `Work Sample Answers.md` — The Written Answers
- This is the main document with my final answers to all 9 questions
- Each answer has a table showing the exact records that violated the rule
- Written in Markdown so it looks professional and is easy to read

### 2. `analysis.py` — Python Script (566 lines)
- A fully automated Python script that loads the Excel data and runs all 9 checks programmatically
- Uses **pandas** for data manipulation, **matplotlib** and **seaborn** for charts
- Outputs results to the terminal AND exports CSV files and chart images
- **Why I made this:** Shows I can automate data quality checks, not just do them manually

### 3. `analysis.sql` — SQL Queries (369 lines)
- Every single question answered in SQL (SQLite dialect)
- Includes multiple approaches for some questions (e.g., self-join AND GROUP BY for duplicates)
- Has a bonus "Full Data Quality Summary" query at the end that counts all violations
- **Why I made this:** Shows I'm comfortable writing SQL, which is essential for a data analyst role

### 4. `Meridian Precast - Analysis Workbook.xlsx` — Enhanced Excel Workbook
- Created by `create_excel_workbook.py` (a Python script using openpyxl)
- Has an "Analysis Summary" tab plus separate tabs for each question (Q1–Q9)
- Includes conditional formatting (red for violations, yellow for warnings, green for OK)
- Has an embedded bar chart for labour costs
- Also includes the original raw data tabs for reference
- **Why I made this:** Shows I can present findings in Excel, which is what most business users expect

### 5. `dashboard.html` — Interactive Web Dashboard (1153 lines)
- A single-file HTML dashboard with Chart.js visualizations
- Dark-themed, modern design with KPI cards at the top
- 5 tabs: Overview, Rule Violations, Labour Cost (Q8), Job Comparison (Q9), Products (Q7)
- Interactive charts: bar charts, doughnut charts, radar charts
- Animated KPI counters on page load
- Responsive design (works on mobile too)
- **Why I made this:** Shows I can build data visualizations beyond just Excel/Python charts

### 6. Supporting Files
- **`load_to_sqlite.py`** — Helper script that loads the Excel data into a SQLite database so the SQL queries can run
- **`analysis_output/`** — Folder with 15 files: CSV exports and PNG chart images for every question

---

## 📊 The 9 Questions — What I Found and How I Found It

### Q1: Missing Operator Records (Rule R3)

**What they asked:** Are there any labour records where we don't know who did the work?

**What I found:** Yes — **2 records** (Record 148 and Record 175) have blank Operator ID and Operator Name.

**How I found it in Python:**
```python
missing_operator = labour[labour['Operator ID'].isna() | labour['Operator Name'].isna()]
```
This just filters the dataframe for rows where Operator ID is blank (NaN) OR Operator Name is blank.

**How I found it in SQL:**
```sql
SELECT * FROM labour_records WHERE operator_id IS NULL OR operator_name IS NULL;
```

**How to explain in interview:** "I filtered the labour records for any rows where the operator fields were null or blank. Found 2 records — Record 148 on JOB-1007 at Curing stage, and Record 175 on JOB-1009 at Mixing stage. Both are completely missing operator information."

---

### Q2: Duplicate Records (Rule R6)

**What they asked:** Has any record been entered twice by mistake? How would you have spotted it?

**What I found:** Yes — **Records 119 and 120** are identical across all key fields.

**How I found it in Python:**
```python
dup_cols = ['Date', 'Job Number', 'Stage', 'Operator ID', 'Hours Worked',
            'Units Completed', 'Units Scrapped']
duplicates = labour[labour.duplicated(subset=dup_cols, keep=False)]
```
I used `pandas.duplicated()` with `keep=False` (which marks ALL duplicates, not just the second one). I grouped by all the fields that should be unique together — if two records match on date, job, stage, operator, hours, and units, they're duplicates.

**How I found it in SQL (two ways):**
```sql
-- Way 1: Self-join
SELECT a.record_id, b.record_id FROM labour_records a JOIN labour_records b
ON a.date = b.date AND a.job_number = b.job_number ... AND a.record_id < b.record_id;

-- Way 2: GROUP BY + HAVING
SELECT GROUP_CONCAT(record_id) FROM labour_records
GROUP BY date, job_number, stage, operator_id, hours_worked, units_completed, units_scrapped
HAVING COUNT(*) > 1;
```

**How to explain in interview:** "I grouped records by their key fields — date, job, stage, operator, hours, and units — and looked for groups with more than one entry. Records 119 and 120 are identical on every field. Both are OP-104 on JOB-1004, Casting stage, June 8th, 5.96 hours, 5 units. The detection method is simple: if two rows have the same values on all these fields, one is a duplicate."

---

### Q3: Overtime Violations (Rule R1)

**What they asked:** Did anyone record more than 8 hours on a single date?

**What I found:** Yes — **2 operators** exceeded the limit:
- OP-104 (Osei, K.) worked **11.92 hours** on June 8th
- OP-102 (Nguyen, T.) worked **11.75 hours** on July 6th

**How I found it:**
```python
daily_hours = labour.groupby(['Operator ID', 'Operator Name', 'Date'])['Hours Worked'].sum()
overtime = daily_hours[daily_hours > 8.0]
```
The key insight is that an operator might have **multiple records** on the same day (e.g., working on different stages). So I **summed** their hours per day first, then checked if the total exceeded 8.

**How to explain in interview:** "A single record might show 5 hours, which is fine. But if the same person has three records on the same day totalling 12 hours, that's a violation. So I grouped by operator and date, summed the hours, and filtered for totals above 8. Found two cases — OP-104 at nearly 12 hours and OP-102 at nearly 12 hours."

---

### Q4: Missing Scrap Reason (Rule R4)

**What they asked:** If units are scrapped, is a reason always given? How many units are affected?

**What I found:** **2 records** have scrapped units but no scrap reason, affecting **14 units** total (7 + 7).

**How I found it:**
```python
scrapped_no_reason = labour[(labour['Units Scrapped'] > 0) & (labour['Scrap Reason'].isna())]
```

**How to explain in interview:** "Simple filter — give me records where Units Scrapped is greater than zero AND Scrap Reason is blank. Found Record 243 (7 units on JOB-1009, Finishing) and Record 250 (7 units on JOB-1012, Casting). That's 14 units with no explanation for why they were thrown away."

---

### Q5: Labour After Job Close (Rule R2)

**What they asked:** Did anyone log work on a job after it was already closed?

**What I found:** Yes — **2 records** on JOB-1014 were logged after it closed on July 10th.

**How I found it:**
```python
labour_jobs = labour.merge(jobs[['Job Number', 'Date Closed']], on='Job Number', how='left')
post_close = labour_jobs[labour_jobs['Date'] > labour_jobs['Date Closed']]
```
I merged (joined) the labour table with the jobs table to get each record's job close date. Then I filtered for records where the labour date is after the close date.

**How to explain in interview:** "I joined the labour records to the jobs table to bring in each job's close date. Then I compared: is the work date after the close date? Found 2 records on JOB-1014 — Record 263 (Casting, July 13, 3 days late) and Record 266 (Finishing, July 14, 4 days late). The job was closed on July 10th, so no work should have been logged after that."

---

### Q6: Overproduction at Finishing (Rule R5)

**What they asked:** Did any job finish more units than were ordered?

**What I found:** Yes — **JOB-1011** finished **22 units** but only **16 were ordered** (overage of 6).

**How I found it:**
```python
finishing = labour[labour['Stage'] == 'Finishing'].groupby('Job Number')['Units Completed'].sum()
finishing_vs_ordered = finishing.merge(jobs[['Job Number', 'Quantity Ordered']], on='Job Number')
overproduced = finishing_vs_ordered[finishing_vs_ordered['Total Finished'] > finishing_vs_ordered['Quantity Ordered']]
```
I filtered for only the Finishing stage (because "finishing" units is what matters for output), summed the units completed per job, then compared to the quantity ordered.

**How to explain in interview:** "I only looked at the Finishing stage because that's where units come out as completed products. Summed up the total finished units per job and compared to what was ordered. JOB-1011 for PRD-440 (Retaining Block) finished 22 units when only 16 were ordered — that's 6 extra. All other jobs were at or below their ordered quantity."

---

### Q7: Products Tab Issues (Rule R8)

**What they asked:** Look at the Products tab — describe any issues.

**What I found:** **PRD-440 (Retaining Block)** costs $366.00 to make but sells for only $312.00. **The company loses $54 on every unit sold.**

**How I found it:**
```python
products['Margin ($)'] = products['Selling Price per Unit ($)'] - products['Cost to Make per Unit ($)']
loss_products = products[products['Margin ($)'] < 0]
```

**How to explain in interview:** "I calculated the margin for every product — selling price minus cost. Four of the five products are profitable with margins around 32-35%. But PRD-440 (Retaining Block) has a NEGATIVE margin of -$54 per unit. This violates Rule R8 which says every product should sell for more than it costs. This could be a data entry error — maybe the cost and selling price got swapped — or it could be a genuine pricing problem that needs urgent review."

---

### Q8: Labour Cost by Work Centre (Rule R7)

**What they asked:** Calculate total hours and total labour cost for each work centre.

**What I found:**

| Work Centre | Total Hours | Rate ($/hr) | Labour Cost |
|-------------|------------|-------------|-------------|
| Mixing (WC-100) | 129.66 | $46.50 | $6,029 |
| Casting (WC-200) | 433.04 | $42.75 | $18,512 |
| Curing (WC-300) | 96.29 | $28.00 | $2,696 |
| Finishing (WC-350) | 214.56 | **$0.00** | **$0.00** ⚠️ |
| **TOTAL** | **873.55** | | **$27,238** |

**The concern:** Finishing has a rate of $0/hr, so 214 hours of work contribute zero to costs. The total of $27,238 is likely **understated**.

**How I found it:**
```python
labour_wc = labour.merge(centres, on='Work Centre', how='left')
labour_wc['Labour Cost ($)'] = labour_wc['Hours Worked'] * labour_wc['Labour Rate ($ per hour)']
wc_summary = labour_wc.groupby(['Work Centre', 'Name']).agg(Total_Hours=('Hours Worked', 'sum'), ...)
```

**How to explain in interview:** "I joined the labour records with the work centres table to get each record's hourly rate. Then I multiplied hours × rate to get the cost for each record, and summed by work centre. The big red flag is WC-350 Finishing — it has a $0.00 per hour rate. That means 214 hours of work show up as costing nothing. This is either because finishing staff are salaried and costed differently, or it's a data entry error. Either way, the total labour cost number is incomplete."

---

### Q9: JOB-1006 vs JOB-1010 Comparison

**What they asked:** Both jobs are for the same product (PRD-110, Wall Panel 8ft), same quantity (24 units), but one used way more labour. How big is the difference and what might explain it?

**What I found:**
- JOB-1006: **63.19 hours**
- JOB-1010: **122.51 hours**
- Difference: **59.32 hours** (JOB-1010 used nearly DOUBLE)

The gap is almost entirely at the **Casting stage** (58.61 of 59.32 hours).

**My 3 hypotheses:**

1. **Trainee operator:** JOB-1010 includes OP-106 (Singh, A.) who doesn't appear on ANY other job. This person logged 56 hours — all exactly 8.00 per day — exclusively at Casting. Likely a new or trainee employee working slower. This single operator accounts for almost all the 59-hour difference.

2. **Higher scrap/rework:** JOB-1010 scrapped 29 units vs only 4 for JOB-1006. More scrap = more rework = more hours to produce replacement units.

3. **Longer production span:** JOB-1010 took 24 calendar days vs 18 for JOB-1006, suggesting possible interruptions or lower-priority scheduling.

**How to explain in interview:** "I pulled all labour records for both jobs and compared them side by side. The 59-hour difference is concentrated almost entirely in the Casting stage. When I dug deeper, I found OP-106 (Singh, A.) — this operator only appears on JOB-1010, nowhere else in the dataset. They logged exactly 8.00 hours every single day, 7 records, 56 total hours — all at Casting. This strongly suggests a trainee or new hire. On top of that, JOB-1010 had 7x the scrap (29 vs 4 units), which means more rework. My next steps would be to ask the production manager if OP-106 was being trained, and to check for equipment issues."

---

## 🔧 Technical Details — How Each File Works

### Python Analysis (`analysis.py`)

**Libraries used:**
- `pandas` — Loading Excel, filtering, grouping, merging data (this is the main one)
- `matplotlib` + `seaborn` — Creating charts (bar charts, horizontal bar charts)
- `numpy` — Number operations (minor use)
- `sqlite3`, `os`, `warnings` — Utilities

**Key pandas operations I used:**
- `pd.read_excel()` — Load data from Excel sheets
- `.isna()` — Check for blank/missing values
- `.duplicated(subset=..., keep=False)` — Find duplicate rows
- `.groupby().sum()` — Group data and sum up values
- `.merge()` — Join two tables (like SQL JOIN)
- `.sort_values()` — Sort results
- Boolean filtering: `df[condition1 & condition2]` — Filter rows matching multiple conditions

**Chart types created:**
- Horizontal bar chart for overtime violations (Q3)
- Grouped bar chart for ordered vs finished (Q6)
- Grouped bar chart for cost vs selling price (Q7)
- Side-by-side bar + doughnut for labour costs (Q8)
- 3-panel comparison chart for Q9 (by stage, scrap, total hours)

### SQL Analysis (`analysis.sql`)

**Key SQL techniques:**
- `WHERE ... IS NULL` — Finding missing values
- `SELF JOIN` (a.record_id < b.record_id) — Comparing records to themselves for duplicates
- `GROUP BY ... HAVING COUNT(*) > 1` — Alternative duplicate detection
- `SUM()`, `COUNT()`, `GROUP_CONCAT()` — Aggregation functions
- `INNER JOIN` — Combining tables
- `CASE WHEN ... THEN ... END` — Conditional logic in queries
- `JULIANDAY()` — SQLite function for date arithmetic
- `UNION ALL` — Combining multiple result sets (in the bonus summary query)

### Excel Workbook (`create_excel_workbook.py`)

**Key openpyxl techniques:**
- `Font`, `PatternFill`, `Alignment`, `Border` — Cell styling
- Conditional formatting — Red for violations, yellow for warnings, green for OK
- `BarChart`, `Reference` — Embedded charts
- Tab colors to organize sheets visually
- Auto-fit column widths

### Dashboard (`dashboard.html`)

**Technologies:**
- Pure HTML/CSS/JavaScript (no framework needed)
- **Chart.js** (CDN) for interactive charts
- **Google Fonts** (Inter) for typography
- CSS custom properties (variables) for the dark theme
- CSS Grid for responsive layout
- Keyframe animations for entry effects
- `requestAnimationFrame` for smooth KPI counter animations

**Chart types in the dashboard:**
- Bar charts (violations count, labour cost, product comparison)
- Doughnut chart (hours distribution)
- Radar chart (job metrics comparison)

---

## 🗣️ How to Talk About This in the Interview

### "Walk me through your approach"

> "I started by loading the Excel data and understanding the four tables and how they relate to each other. Then I went through each business rule one by one, writing both Python and SQL queries to systematically check for violations. I exported my findings as CSV files and charts, then created an Excel workbook with analysis tabs, and finally built an interactive dashboard to present everything visually. I wanted to show I can work across multiple tools — not just Python or just Excel."

### "Why did you use multiple tools?"

> "Different stakeholders prefer different formats. A controller might want the Excel workbook. A technical team might want the SQL queries to build into automated checks. The dashboard is great for a quick visual overview. And the Python script makes everything reproducible — run it again whenever the data updates."

### "What was the most interesting finding?"

> "Probably Question 9 — the job comparison. At first it just looks like one job used more hours, but when you dig into it, you find this one operator (OP-106) who only appears on that one job, logged exactly 8.00 hours every day, and accounts for almost all the excess hours. That pattern — perfectly round numbers, only on one job — strongly suggests a trainee. It's the kind of insight you only get by actually exploring the data, not just running a formula."

### "How did you detect duplicates?"

> "I chose a combination of key fields that should be unique together — date, job, stage, operator, hours, units completed, and units scrapped. If two records match on ALL of those fields, one is almost certainly a duplicate. I didn't include Record ID because that's auto-generated and would always be different. I showed two different SQL approaches: a self-join and a GROUP BY with HAVING, to demonstrate I know multiple ways to solve the same problem."

### "What would you do differently with more time?"

> "I'd add input validation rules — like automated alerts that flag these issues BEFORE the data is saved. I'd also calculate the total financial impact of each violation, for example how much extra did the overtime cost, what's the dollar impact of the 14 unaccounted scrapped units, etc. And I'd build a proper ETL pipeline that runs these checks automatically every time new data comes in."

### "Tell me about the dashboard"

> "It's a single HTML file — no server needed, just open it in a browser. I used Chart.js for the interactive charts, CSS Grid for the responsive layout, and CSS custom properties for the dark theme. It has 5 tabs: an overview with KPI cards and a summary table, detailed violation records, the labour cost breakdown for Q8, the side-by-side job comparison for Q9, and the product margin analysis for Q7. The KPI counters at the top animate on page load using requestAnimationFrame for smooth counting."

---

## 📁 Quick Reference: File List

| File | Lines | What It Is |
|------|-------|-----------|
| `Work Sample Answers.md` | 161 | Final written answers to all 9 questions |
| `analysis.py` | 566 | Python script — automated analysis with charts |
| `analysis.sql` | 369 | SQL queries for all 9 questions |
| `create_excel_workbook.py` | 499 | Script that generates the Excel workbook |
| `load_to_sqlite.py` | 52 | Helper to load Excel → SQLite database |
| `dashboard.html` | 1153 | Interactive web dashboard |
| `README.md` | 48 | Project overview and how to run |
| `analysis_output/` | 15 files | CSV exports and chart PNGs |

---

## 🔢 Key Numbers to Remember

- **169** labour records analyzed
- **873.55** total hours across all work centres
- **$27,237.77** total labour cost (understated because Finishing is $0/hr)
- **11** rule violations found across 7 different rules
- **59.32 hours** difference between JOB-1006 and JOB-1010
- **$54** loss per unit on PRD-440 (Retaining Block)
- **14 units** scrapped with no reason given
- **6 units** overproduced on JOB-1011
- **OP-106** — the mystery trainee operator (56 hours, all at Casting, all exactly 8.00/day)

---

*Good luck with the interview! You've got this. 💪*
