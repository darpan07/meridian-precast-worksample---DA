# Meridian Precast — Work Sample Answers

**Position:** Data Analyst — Precon  
**Candidate:** [Your Name]  
**Date:** September 2026  
**Tools Used:** Python (pandas, matplotlib), SQL (SQLite), Excel (openpyxl), HTML/CSS/Chart.js  

---

## Question 1 — Missing Operator Records (Rule R3)

**Yes, there are 2 labour records with no operator information.**

| Record ID | Date       | Job      | Stage  | Operator ID | Operator Name |
|-----------|-----------|----------|--------|-------------|---------------|
| 148       | 2026-06-12 | JOB-1007 | Curing | *(blank)*   | *(blank)*     |
| 175       | 2026-06-18 | JOB-1009 | Mixing | *(blank)*   | *(blank)*     |

Both the Operator ID and Operator Name fields are blank in these records, violating Rule R3.

---

## Question 2 — Duplicate Records (Rule R6)

**Yes, 1 pair of records has been entered twice by mistake.**

| Record ID | Date       | Job      | Stage   | Operator        | Hours | Units |
|-----------|-----------|----------|---------|-----------------|-------|-------|
| 119       | 2026-06-08 | JOB-1004 | Casting | OP-104 Osei, K. | 5.96  | 5     |
| 120       | 2026-06-08 | JOB-1004 | Casting | OP-104 Osei, K. | 5.96  | 5     |

**Detection method:** I grouped all records by (Date, Job Number, Stage, Operator ID, Hours Worked, Units Completed, Units Scrapped) and looked for groups with more than one entry. Records 119 and 120 are identical across all key fields.

---

## Question 3 — Overtime Violations (Rule R1)

**Yes, 2 operators exceeded the 8-hour limit on a single date.**

| Operator ID | Operator Name | Date       | Total Hours | Excess |
|-------------|--------------|-----------|-------------|--------|
| OP-104      | Osei, K.     | 2026-06-08 | 11.92       | +3.92  |
| OP-102      | Nguyen, T.   | 2026-07-06 | 11.75       | +3.75  |

*Note: This was found by summing each operator's hours across all their records on each date, then filtering where the total exceeded 8.00.*

---

## Question 4 — Missing Scrap Reason (Rule R4)

**Yes, 2 records have scrapped units but no scrap reason. A total of 14 units are affected.**

| Record ID | Date       | Job      | Stage    | Units Scrapped | Scrap Reason |
|-----------|-----------|----------|----------|---------------|-------------|
| 243       | 2026-07-06 | JOB-1009 | Finishing | 7             | *(blank)*   |
| 250       | 2026-07-07 | JOB-1012 | Casting  | 7             | *(blank)*   |

---

## Question 5 — Labour After Job Close (Rule R2)

**Yes, 2 records on JOB-1014 were dated after the job's close date.**

JOB-1014 was closed on **2026-07-10**, but the following records were logged after that:

| Record ID | Labour Date | Job      | Stage    | Job Closed | Days After |
|-----------|-----------|----------|----------|-----------|-----------|
| 263       | 2026-07-13 | JOB-1014 | Casting  | 2026-07-10 | +3 days   |
| 266       | 2026-07-14 | JOB-1014 | Finishing | 2026-07-10 | +4 days   |

---

## Question 6 — Overproduction at Finishing (Rule R5)

**Yes, JOB-1011 finished more units than ordered — by 6 units.**

| Job      | Product | Qty Ordered | Total Finished | Overage |
|----------|---------|-------------|----------------|---------|
| JOB-1011 | PRD-440 | 16          | 22             | **+6**  |

All other closed jobs with Finishing records finished at or below the quantity ordered.

---

## Question 7 — Products Tab Issues (Rule R8)

**PRD-440 (Retaining Block) is sold at a loss, violating Rule R8.**

| Product | Description      | Cost ($) | Sell ($) | Margin ($) | Margin % | Status |
|---------|-----------------|----------|----------|-----------|----------|--------|
| PRD-110 | Wall Panel 8ft  | 268.50   | 412.00   | +143.50   | +34.8%   | ✅     |
| PRD-120 | Wall Panel 10ft | 331.75   | 505.00   | +173.25   | +34.3%   | ✅     |
| PRD-310 | Manhole Barrel  | 214.60   | 318.00   | +103.40   | +32.5%   | ✅     |
| **PRD-440** | **Retaining Block** | **366.00** | **312.00** | **−54.00** | **−17.3%** | **❌** |
| PRD-450 | Paving Slab     | 248.75   | 372.00   | +123.25   | +33.1%   | ✅     |

The company **loses $54.00 on every Retaining Block sold**. This could be a data entry error (e.g., cost and selling price swapped) or a genuine pricing problem that needs urgent review.

---

## Question 8 — Labour Cost by Work Centre (Rule R7)

| Work Centre | Name      | Records | Total Hours | Rate ($/hr) | Labour Cost ($) |
|-------------|-----------|---------|-------------|-------------|-----------------|
| WC-100      | Mixing    | 28      | 129.66      | $46.50      | $6,029.19       |
| WC-200      | Casting   | 77      | 433.04      | $42.75      | $18,512.46      |
| WC-300      | Curing    | 23      | 96.29       | $28.00      | $2,696.12       |
| WC-350      | Finishing | 41      | 214.56      | **$0.00**   | **$0.00** ⚠️    |
| **TOTAL**   |           | **169** | **873.55**  |             | **$27,237.77**  |

**Concern:** WC-350 (Finishing) has a labour rate of **$0.00 per hour**. This means 214.56 hours of Finishing work contribute nothing to labour costs. The total of $27,237.77 is likely **understated**.

This could mean: (1) Finishing is done by salaried staff not costed hourly, (2) it's an automated process, or (3) the rate is a data entry error. In any case, if Finishing labour should carry a cost, then every job's cost is currently incomplete.

---

## Question 9 — JOB-1006 vs JOB-1010 Comparison

### How big is the difference?

JOB-1006 used **63.19 hours** and JOB-1010 used **122.51 hours**. The difference is **59.32 hours** — JOB-1010 used nearly double the labour despite being for the same product (PRD-110, Wall Panel 8ft) and the same quantity (24 units).

**Hours by stage:**

| Stage    | JOB-1006 | JOB-1010 | Difference |
|----------|----------|----------|-----------|
| Mixing   | 11.52    | 11.64    | 0.12      |
| Casting  | 27.79    | **86.40**| **58.61** |
| Curing   | 6.99     | 7.09     | 0.10      |
| Finishing| 16.89    | 17.38    | 0.49      |

The gap is almost entirely in **Casting** (58.61 of the 59.32-hour difference).

### My best explanations:

1. **Operator experience:** JOB-1010 includes OP-106 (Singh, A.) who does not appear on any other job in the dataset. This operator logged **56.00 hours** across 7 records — all exactly 8.00 hours per day — exclusively at the Casting stage. OP-106 is likely a **new or trainee operator** who works slower. This single operator accounts for virtually all of the 59-hour difference.

2. **Higher scrap / rework:** JOB-1010 scrapped **29 units** vs. only 4 for JOB-1006. More scrap means more rework — extra hours and material to produce replacement units.

3. **Longer production span:** JOB-1010 took **24 calendar days** vs. 18 for JOB-1006, suggesting possible production interruptions or lower-priority scheduling.

### What I would want to find out next:

- Ask the **Production Manager** whether OP-106 was being trained during JOB-1010.
- Check if there were **equipment issues or downtime** logged during JOB-1010's dates.
- Ask the Controller whether OP-106's **perfectly round 8.00-hour entries** represent actual hours or are estimates.
- Review whether JOB-1010 had any **material or quality issues** not captured in the scrap reason field.

---

## Appendix — Files Included

| File | Description |
|------|-------------|
| `analysis.py` | Python script — all 9 answers programmatically with charts |
| `analysis.sql` | SQL queries — all 9 questions in SQLite-compatible SQL |
| `load_to_sqlite.py` | Helper script to create SQLite database from Excel |
| `Meridian Precast - Analysis Workbook.xlsx` | Enhanced Excel with Q1-Q9 analysis tabs |
| `dashboard.html` | Interactive HTML dashboard with Chart.js visualizations |
| `analysis_output/` | Exported CSV files and chart images |
