-- =============================================================================
-- Meridian Precast Ltd. — Production Data Quality Analysis (SQL Version)
-- =============================================================================
-- Analyst:     [Your Name]
-- Date:        2026-09-01
-- Database:    SQLite (loaded from Excel via Python)
-- Purpose:     Answer all 9 questions using SQL queries
-- =============================================================================

-- NOTE: This SQL file is designed to run against a SQLite database created by
-- the companion script 'load_to_sqlite.py'. Each query corresponds to one
-- question from the Work Sample Brief.

-- =============================================================================
-- TABLE STRUCTURE (for reference)
-- =============================================================================
-- labour_records (
--   record_id INT, date TEXT, job_number TEXT, product_code TEXT, stage TEXT,
--   work_centre TEXT, operator_id TEXT, operator_name TEXT, hours_worked REAL,
--   units_completed INT, units_scrapped INT, scrap_reason TEXT, notes TEXT
-- )
-- jobs (job_number TEXT, product_code TEXT, quantity_ordered INT, status TEXT, date_closed TEXT)
-- products (product_code TEXT, description TEXT, cost_per_unit REAL, sell_price_per_unit REAL)
-- work_centres (work_centre TEXT, name TEXT, labour_rate REAL)


-- =============================================================================
-- QUESTION 1: Missing Operator Records (Rule R3)
-- "Are there any labour records that do not say who did the work?"
-- =============================================================================

SELECT 
    record_id,
    date,
    job_number,
    stage,
    operator_id,
    operator_name,
    'MISSING OPERATOR - Violates Rule R3' AS violation
FROM labour_records
WHERE operator_id IS NULL 
   OR operator_name IS NULL;

-- Expected: Record IDs 148 and 175


-- =============================================================================
-- QUESTION 2: Duplicate Records (Rule R6)
-- "Has any record been entered into the system twice by mistake?"
-- =============================================================================

-- Method: Find records with identical key fields using a self-join approach
SELECT 
    a.record_id AS record_id_1,
    b.record_id AS record_id_2,
    a.date,
    a.job_number,
    a.stage,
    a.operator_id,
    a.hours_worked,
    a.units_completed,
    a.units_scrapped,
    'DUPLICATE ENTRY - Violates Rule R6' AS violation
FROM labour_records a
JOIN labour_records b
  ON a.date = b.date
 AND a.job_number = b.job_number
 AND a.stage = b.stage
 AND a.operator_id = b.operator_id
 AND a.hours_worked = b.hours_worked
 AND a.units_completed = b.units_completed
 AND a.units_scrapped = b.units_scrapped
 AND a.record_id < b.record_id;

-- Expected: Records 119 and 120 (JOB-1004, OP-104, 2026-06-08)
-- Detection: GROUP BY all key fields and look for COUNT > 1

-- Alternative approach using GROUP BY + HAVING:
SELECT 
    date, job_number, stage, operator_id, hours_worked,
    units_completed, units_scrapped,
    GROUP_CONCAT(record_id) AS duplicate_record_ids,
    COUNT(*) AS occurrence_count
FROM labour_records
GROUP BY date, job_number, stage, operator_id, hours_worked,
         units_completed, units_scrapped
HAVING COUNT(*) > 1;


-- =============================================================================
-- QUESTION 3: Overtime Violations (Rule R1 — max 8 hours per day)
-- "Did anyone record more than 8 hours on one date?"
-- =============================================================================

-- Step 1: Calculate total hours per operator per day
-- Step 2: Filter where total exceeds 8 hours

SELECT 
    operator_id,
    operator_name,
    date,
    SUM(hours_worked) AS total_hours,
    COUNT(*) AS record_count,
    GROUP_CONCAT(record_id) AS contributing_records,
    'OVERTIME - Violates Rule R1 (>8 hrs)' AS violation
FROM labour_records
WHERE operator_id IS NOT NULL
GROUP BY operator_id, operator_name, date
HAVING SUM(hours_worked) > 8.0
ORDER BY total_hours DESC;

-- Expected:
--   OP-102 (Nguyen, T.) on 2026-07-06: 11.75 hours
--   OP-104 (Osei, K.)   on 2026-06-08: 11.92 hours


-- =============================================================================
-- QUESTION 4: Missing Scrap Reason (Rule R4)
-- "If any units are scrapped, a scrap reason must be recorded."
-- =============================================================================

SELECT 
    record_id,
    date,
    job_number,
    stage,
    operator_name,
    units_scrapped,
    scrap_reason,
    'MISSING SCRAP REASON - Violates Rule R4' AS violation
FROM labour_records
WHERE units_scrapped > 0
  AND (scrap_reason IS NULL OR TRIM(scrap_reason) = '');

-- Summary: total affected units
SELECT 
    COUNT(*) AS records_affected,
    SUM(units_scrapped) AS total_units_without_reason
FROM labour_records
WHERE units_scrapped > 0
  AND (scrap_reason IS NULL OR TRIM(scrap_reason) = '');

-- Expected: 2 records, 14 total units (Records 243 and 250)


-- =============================================================================
-- QUESTION 5: Labour After Job Close (Rule R2)
-- "No labour should be dated after a job's Date Closed."
-- =============================================================================

SELECT 
    lr.record_id,
    lr.date AS labour_date,
    lr.job_number,
    j.date_closed,
    CAST(JULIANDAY(lr.date) - JULIANDAY(j.date_closed) AS INTEGER) AS days_after_close,
    lr.stage,
    lr.operator_name,
    lr.hours_worked,
    'POST-CLOSE LABOUR - Violates Rule R2' AS violation
FROM labour_records lr
INNER JOIN jobs j ON lr.job_number = j.job_number
WHERE j.date_closed IS NOT NULL
  AND lr.date > j.date_closed
ORDER BY lr.job_number, lr.date;

-- Expected: 2 records for JOB-1014 (Records 263 and 266)


-- =============================================================================
-- QUESTION 6: Overproduction at Finishing (Rule R5)
-- "A job should not finish more units than were ordered."
-- =============================================================================

-- Sum completed units at Finishing stage, compare to quantity ordered
SELECT 
    lr.job_number,
    j.product_code,
    j.quantity_ordered,
    SUM(lr.units_completed) AS total_finished_units,
    SUM(lr.units_completed) - j.quantity_ordered AS overage,
    CASE 
        WHEN SUM(lr.units_completed) > j.quantity_ordered 
        THEN 'OVERPRODUCTION - Violates Rule R5'
        ELSE 'OK'
    END AS status
FROM labour_records lr
INNER JOIN jobs j ON lr.job_number = j.job_number
WHERE lr.stage = 'Finishing'
GROUP BY lr.job_number, j.product_code, j.quantity_ordered
ORDER BY overage DESC;

-- Expected: JOB-1011 overproduced by 6 units (finished 22, ordered 16)


-- =============================================================================
-- QUESTION 7: Products Tab Issues (Rule R8)
-- "Every product should sell for more than it costs to make."
-- =============================================================================

SELECT 
    product_code,
    description,
    cost_per_unit,
    sell_price_per_unit,
    (sell_price_per_unit - cost_per_unit) AS margin_dollars,
    ROUND((sell_price_per_unit - cost_per_unit) * 100.0 / sell_price_per_unit, 1) AS margin_pct,
    CASE 
        WHEN sell_price_per_unit <= cost_per_unit 
        THEN 'LOSS MAKER - Violates Rule R8'
        ELSE 'Profitable'
    END AS status
FROM products
ORDER BY margin_dollars;

-- Expected: PRD-440 (Retaining Block) costs $366, sells for $312 → -$54 loss per unit


-- =============================================================================
-- QUESTION 8: Labour Cost by Work Centre (Rule R7)
-- "Labour cost = Hours x Labour Rate for that work centre"
-- =============================================================================

SELECT 
    wc.work_centre,
    wc.name AS work_centre_name,
    COUNT(lr.record_id) AS total_records,
    ROUND(SUM(lr.hours_worked), 2) AS total_hours,
    wc.labour_rate AS rate_per_hour,
    ROUND(SUM(lr.hours_worked * wc.labour_rate), 2) AS total_labour_cost,
    CASE 
        WHEN wc.labour_rate = 0 
        THEN '⚠ $0 rate - see note below'
        ELSE ''
    END AS concern
FROM labour_records lr
INNER JOIN work_centres wc ON lr.work_centre = wc.work_centre
GROUP BY wc.work_centre, wc.name, wc.labour_rate
ORDER BY wc.work_centre;

-- Grand total
SELECT 
    'TOTAL' AS work_centre,
    '' AS work_centre_name,
    COUNT(*) AS total_records,
    ROUND(SUM(lr.hours_worked), 2) AS total_hours,
    NULL AS rate_per_hour,
    ROUND(SUM(lr.hours_worked * wc.labour_rate), 2) AS total_labour_cost,
    '' AS concern
FROM labour_records lr
INNER JOIN work_centres wc ON lr.work_centre = wc.work_centre;

-- NOTE: WC-350 (Finishing) has $0/hr rate → 214.56 hours contribute $0 to costs.
-- This is either intentional (salaried staff) or a data error. Either way,
-- total labour cost ($27,237.77) is understated if Finishing should be costed.


-- =============================================================================
-- QUESTION 9: JOB-1006 vs JOB-1010 Comparison
-- =============================================================================

-- Overall comparison
SELECT 
    job_number,
    COUNT(*) AS record_count,
    ROUND(SUM(hours_worked), 2) AS total_hours,
    SUM(units_completed) AS total_completed,
    SUM(units_scrapped) AS total_scrapped,
    MIN(date) AS start_date,
    MAX(date) AS end_date,
    CAST(JULIANDAY(MAX(date)) - JULIANDAY(MIN(date)) + 1 AS INTEGER) AS calendar_days
FROM labour_records
WHERE job_number IN ('JOB-1006', 'JOB-1010')
GROUP BY job_number;

-- Hours breakdown by stage
SELECT 
    job_number,
    stage,
    ROUND(SUM(hours_worked), 2) AS stage_hours,
    SUM(units_completed) AS units_completed,
    SUM(units_scrapped) AS units_scrapped,
    COUNT(*) AS records
FROM labour_records
WHERE job_number IN ('JOB-1006', 'JOB-1010')
GROUP BY job_number, stage
ORDER BY job_number, 
    CASE stage 
        WHEN 'Mixing' THEN 1 
        WHEN 'Casting' THEN 2 
        WHEN 'Curing' THEN 3 
        WHEN 'Finishing' THEN 4 
    END;

-- Operator breakdown
SELECT 
    job_number,
    operator_id,
    operator_name,
    ROUND(SUM(hours_worked), 2) AS hours,
    COUNT(*) AS records,
    SUM(units_scrapped) AS scrapped
FROM labour_records
WHERE job_number IN ('JOB-1006', 'JOB-1010')
GROUP BY job_number, operator_id, operator_name
ORDER BY job_number, hours DESC;

-- Key finding: OP-106 (Singh, A.) only appears on JOB-1010
-- with exactly 8.00 hours per record (7 records, 56 total hours).
-- This operator accounts for the majority of the 59.32-hour difference.

-- Investigate OP-106 specifically
SELECT 
    record_id,
    date,
    job_number,
    stage,
    hours_worked,
    units_completed,
    units_scrapped,
    scrap_reason
FROM labour_records
WHERE operator_id = 'OP-106'
ORDER BY date;


-- =============================================================================
-- BONUS: Full Data Quality Summary
-- =============================================================================

-- Count all violations by rule
SELECT 'R1 - Overtime (>8 hrs/day)' AS rule,
    (SELECT COUNT(*) FROM (
        SELECT operator_id, date, SUM(hours_worked) AS total
        FROM labour_records WHERE operator_id IS NOT NULL
        GROUP BY operator_id, date HAVING total > 8
    )) AS violations
UNION ALL
SELECT 'R2 - Post-close labour',
    (SELECT COUNT(*) FROM labour_records lr
     JOIN jobs j ON lr.job_number = j.job_number
     WHERE j.date_closed IS NOT NULL AND lr.date > j.date_closed)
UNION ALL
SELECT 'R3 - Missing operator',
    (SELECT COUNT(*) FROM labour_records WHERE operator_id IS NULL)
UNION ALL
SELECT 'R4 - Missing scrap reason',
    (SELECT COUNT(*) FROM labour_records 
     WHERE units_scrapped > 0 AND scrap_reason IS NULL)
UNION ALL
SELECT 'R5 - Overproduction',
    (SELECT COUNT(*) FROM (
        SELECT lr.job_number, SUM(lr.units_completed) - j.quantity_ordered AS overage
        FROM labour_records lr JOIN jobs j ON lr.job_number = j.job_number
        WHERE lr.stage = 'Finishing'
        GROUP BY lr.job_number, j.quantity_ordered HAVING overage > 0
    ))
UNION ALL
SELECT 'R6 - Duplicate records',
    (SELECT COUNT(*) FROM (
        SELECT date, job_number, stage, operator_id, hours_worked, units_completed, units_scrapped
        FROM labour_records
        GROUP BY date, job_number, stage, operator_id, hours_worked, units_completed, units_scrapped
        HAVING COUNT(*) > 1
    ))
UNION ALL
SELECT 'R8 - Product sold at loss',
    (SELECT COUNT(*) FROM products WHERE sell_price_per_unit <= cost_per_unit);
