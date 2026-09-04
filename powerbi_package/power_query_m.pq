
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
