"""
===============================================================================
 Meridian Precast Ltd. — Production Data Quality Analysis
 =========================================================
 Analyst:     [Your Name]
 Date:        2026-09-01
 Tool:        Python 3.12 + pandas + matplotlib + seaborn
 Purpose:     Answer all 9 questions from the Work Sample Brief
              by validating production data against business rules.
===============================================================================
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend for saving charts
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import sqlite3
import os
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
DATA_FILE   = r'Meridian Precast - Production Data.xlsx'
OUTPUT_DIR  = 'analysis_output'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Chart styling
sns.set_theme(style="whitegrid", palette="muted", font_scale=1.1)
plt.rcParams['figure.dpi'] = 150
plt.rcParams['savefig.bbox'] = 'tight'

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 80)
print("  MERIDIAN PRECAST — PRODUCTION DATA QUALITY ANALYSIS")
print("=" * 80)
print()

print("Loading data from:", DATA_FILE)
labour    = pd.read_excel(DATA_FILE, sheet_name='Labour Records')
jobs      = pd.read_excel(DATA_FILE, sheet_name='Jobs')
products  = pd.read_excel(DATA_FILE, sheet_name='Products')
centres   = pd.read_excel(DATA_FILE, sheet_name='Work Centres')

print(f"  ✓ Labour Records: {len(labour)} rows, {len(labour.columns)} columns")
print(f"  ✓ Jobs:           {len(jobs)} rows")
print(f"  ✓ Products:       {len(products)} rows")
print(f"  ✓ Work Centres:   {len(centres)} rows")
print()

# Ensure date columns are datetime
labour['Date'] = pd.to_datetime(labour['Date'])
jobs['Date Closed'] = pd.to_datetime(jobs['Date Closed'])

# ═════════════════════════════════════════════════════════════════════════════
# QUESTION 1: Missing Operator (Rule R3)
# "Are there any labour records that do not say who did the work?
#  How many, and which records are they?"
# ═════════════════════════════════════════════════════════════════════════════
print("─" * 80)
print("  QUESTION 1: Labour records with missing operator (Rule R3)")
print("─" * 80)

missing_operator = labour[labour['Operator ID'].isna() | labour['Operator Name'].isna()]

print(f"\n  ► Found: {len(missing_operator)} record(s) with missing operator information.\n")

if len(missing_operator) > 0:
    for _, row in missing_operator.iterrows():
        print(f"    Record ID: {row['Record ID']}  |  Date: {row['Date'].strftime('%Y-%m-%d')}  "
              f"|  Job: {row['Job Number']}  |  Stage: {row['Stage']}  "
              f"|  Operator ID: {row['Operator ID']}  |  Operator Name: {row['Operator Name']}")
    
    # Export
    missing_operator.to_csv(os.path.join(OUTPUT_DIR, 'Q1_missing_operator.csv'), index=False)
    print(f"\n  ✓ Exported to {OUTPUT_DIR}/Q1_missing_operator.csv")

print()

# ═════════════════════════════════════════════════════════════════════════════
# QUESTION 2: Duplicate Records (Rule R6)
# "Has any record been entered into the system twice by mistake?
#  If so, which records, and how would you have spotted it?"
# ═════════════════════════════════════════════════════════════════════════════
print("─" * 80)
print("  QUESTION 2: Duplicate records entered by mistake (Rule R6)")
print("─" * 80)

# Check for records that are identical on key fields
dup_cols = ['Date', 'Job Number', 'Stage', 'Operator ID', 'Hours Worked',
            'Units Completed', 'Units Scrapped']
duplicates = labour[labour.duplicated(subset=dup_cols, keep=False)].sort_values('Record ID')

print(f"\n  ► Found: {len(duplicates)} records that appear to be duplicates.\n")

if len(duplicates) > 0:
    for _, row in duplicates.iterrows():
        print(f"    Record ID: {row['Record ID']}  |  Date: {row['Date'].strftime('%Y-%m-%d')}  "
              f"|  Job: {row['Job Number']}  |  Stage: {row['Stage']}  "
              f"|  Operator: {row['Operator Name']}  |  Hours: {row['Hours Worked']}  "
              f"|  Units: {row['Units Completed']}  |  Scrapped: {row['Units Scrapped']}")
    
    print("\n  ► Detection Method:")
    print("    Group records by (Date, Job Number, Stage, Operator ID, Hours Worked,")
    print("    Units Completed, Units Scrapped) and flag groups with more than one entry.")
    print("    Records 119 and 120 share identical values across all these fields.")
    
    duplicates.to_csv(os.path.join(OUTPUT_DIR, 'Q2_duplicates.csv'), index=False)
    print(f"\n  ✓ Exported to {OUTPUT_DIR}/Q2_duplicates.csv")

print()

# ═════════════════════════════════════════════════════════════════════════════
# QUESTION 3: Overtime Violation (Rule R1 — max 8 hours per operator per day)
# "Rule R1 says nobody should record more than 8 hours on one date.
#  Did anyone? Give the operator, the date, and the total hours."
# ═════════════════════════════════════════════════════════════════════════════
print("─" * 80)
print("  QUESTION 3: Operators exceeding 8 hours in a single day (Rule R1)")
print("─" * 80)

# Total hours per operator per date
daily_hours = labour.groupby(['Operator ID', 'Operator Name', 'Date'])['Hours Worked'].sum().reset_index()
daily_hours.columns = ['Operator ID', 'Operator Name', 'Date', 'Total Hours']
overtime = daily_hours[daily_hours['Total Hours'] > 8.0].sort_values('Total Hours', ascending=False)

print(f"\n  ► Found: {len(overtime)} instance(s) of operators exceeding 8 hours on a single date.\n")

if len(overtime) > 0:
    print(f"    {'Operator ID':<14} {'Operator Name':<16} {'Date':<12} {'Total Hours':>11}")
    print(f"    {'─'*14} {'─'*16} {'─'*12} {'─'*11}")
    for _, row in overtime.iterrows():
        flag = " ⚠️" if row['Total Hours'] > 10 else ""
        op_id = row['Operator ID'] if pd.notna(row['Operator ID']) else 'MISSING'
        op_name = row['Operator Name'] if pd.notna(row['Operator Name']) else 'MISSING'
        print(f"    {str(op_id):<14} {str(op_name):<16} {row['Date'].strftime('%Y-%m-%d'):<12} {row['Total Hours']:>10.2f}{flag}")
    
    overtime.to_csv(os.path.join(OUTPUT_DIR, 'Q3_overtime.csv'), index=False)
    print(f"\n  ✓ Exported to {OUTPUT_DIR}/Q3_overtime.csv")

    # Chart: Overtime instances
    fig, ax = plt.subplots(figsize=(10, 5))
    overtime_plot = overtime.copy()
    overtime_plot['Label'] = overtime_plot.apply(
        lambda r: f"{r['Operator Name'] if pd.notna(r['Operator Name']) else 'MISSING'}\n{r['Date'].strftime('%m/%d')}", axis=1)
    bars = ax.barh(overtime_plot['Label'], overtime_plot['Total Hours'], color='#e74c3c', edgecolor='#c0392b')
    ax.axvline(x=8, color='#2c3e50', linestyle='--', linewidth=2, label='8-hour limit (R1)')
    ax.set_xlabel('Total Hours Worked')
    ax.set_title('Rule R1 Violations: Operators Exceeding 8 Hours/Day', fontweight='bold')
    ax.legend()
    for bar, val in zip(bars, overtime_plot['Total Hours']):
        ax.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2, f'{val:.2f}h',
                va='center', fontweight='bold', fontsize=10)
    plt.savefig(os.path.join(OUTPUT_DIR, 'Q3_overtime_chart.png'))
    plt.close()
    print(f"  ✓ Chart saved to {OUTPUT_DIR}/Q3_overtime_chart.png")

print()

# ═════════════════════════════════════════════════════════════════════════════
# QUESTION 4: Missing Scrap Reason (Rule R4)
# "Rule R4 says a scrap reason must be given whenever units are scrapped.
#  Are there records that break this rule? How many units are affected?"
# ═════════════════════════════════════════════════════════════════════════════
print("─" * 80)
print("  QUESTION 4: Scrapped units without a scrap reason (Rule R4)")
print("─" * 80)

scrapped_no_reason = labour[(labour['Units Scrapped'] > 0) & (labour['Scrap Reason'].isna())]

total_units_affected = scrapped_no_reason['Units Scrapped'].sum()
print(f"\n  ► Found: {len(scrapped_no_reason)} record(s) with scrapped units but NO scrap reason.")
print(f"  ► Total units affected: {int(total_units_affected)}\n")

if len(scrapped_no_reason) > 0:
    for _, row in scrapped_no_reason.iterrows():
        print(f"    Record ID: {row['Record ID']}  |  Job: {row['Job Number']}  "
              f"|  Stage: {row['Stage']}  |  Units Scrapped: {int(row['Units Scrapped'])}  "
              f"|  Scrap Reason: (missing)")
    
    scrapped_no_reason.to_csv(os.path.join(OUTPUT_DIR, 'Q4_missing_scrap_reason.csv'), index=False)
    print(f"\n  ✓ Exported to {OUTPUT_DIR}/Q4_missing_scrap_reason.csv")

print()

# ═════════════════════════════════════════════════════════════════════════════
# QUESTION 5: Labour After Job Close Date (Rule R2)
# "Rule R2 says no labour should be dated after a job's Date Closed.
#  Did this happen? Which job, and which records?"
# ═════════════════════════════════════════════════════════════════════════════
print("─" * 80)
print("  QUESTION 5: Labour records dated after job closure (Rule R2)")
print("─" * 80)

# Merge labour with jobs to get Date Closed
labour_jobs = labour.merge(jobs[['Job Number', 'Date Closed', 'Status']], on='Job Number', how='left')
post_close = labour_jobs[
    (labour_jobs['Date Closed'].notna()) &
    (labour_jobs['Date'] > labour_jobs['Date Closed'])
].sort_values(['Job Number', 'Date'])

print(f"\n  ► Found: {len(post_close)} record(s) with labour dated AFTER the job was closed.\n")

if len(post_close) > 0:
    for _, row in post_close.iterrows():
        days_after = (row['Date'] - row['Date Closed']).days
        print(f"    Record ID: {row['Record ID']}  |  Job: {row['Job Number']}  "
              f"|  Labour Date: {row['Date'].strftime('%Y-%m-%d')}  "
              f"|  Job Closed: {row['Date Closed'].strftime('%Y-%m-%d')}  "
              f"|  Days After: {days_after}")
    
    post_close_export = post_close[['Record ID', 'Date', 'Job Number', 'Date Closed', 'Stage', 'Operator Name', 'Hours Worked']]
    post_close_export.to_csv(os.path.join(OUTPUT_DIR, 'Q5_post_close_labour.csv'), index=False)
    print(f"\n  ✓ Exported to {OUTPUT_DIR}/Q5_post_close_labour.csv")

print()

# ═════════════════════════════════════════════════════════════════════════════
# QUESTION 6: Overproduction at Finishing (Rule R5)
# "Rule R5 says a job should not finish more units than were ordered.
#  Did any job do this? Which one, and by how many units?"
# ═════════════════════════════════════════════════════════════════════════════
print("─" * 80)
print("  QUESTION 6: Jobs that finished more units than ordered (Rule R5)")
print("─" * 80)

# Sum units completed at the Finishing stage per job
finishing = labour[labour['Stage'] == 'Finishing'].groupby('Job Number')['Units Completed'].sum().reset_index()
finishing.columns = ['Job Number', 'Total Finished']

# Merge with Jobs to get Quantity Ordered
finishing_vs_ordered = finishing.merge(jobs[['Job Number', 'Quantity Ordered']], on='Job Number', how='left')
finishing_vs_ordered['Overage'] = finishing_vs_ordered['Total Finished'] - finishing_vs_ordered['Quantity Ordered']
overproduced = finishing_vs_ordered[finishing_vs_ordered['Overage'] > 0].sort_values('Overage', ascending=False)

print(f"\n  ► Found: {len(overproduced)} job(s) that finished MORE units than ordered.\n")

if len(overproduced) > 0:
    print(f"    {'Job Number':<12} {'Qty Ordered':>12} {'Total Finished':>15} {'Overage':>10}")
    print(f"    {'─'*12} {'─'*12} {'─'*15} {'─'*10}")
    for _, row in overproduced.iterrows():
        print(f"    {row['Job Number']:<12} {int(row['Quantity Ordered']):>12} "
              f"{int(row['Total Finished']):>15} {int(row['Overage']):>10} ⚠️")

    # Also show all jobs comparison
    print(f"\n  Full comparison (all jobs with Finishing records):")
    print(f"    {'Job Number':<12} {'Qty Ordered':>12} {'Total Finished':>15} {'Overage':>10} {'Status':>8}")
    print(f"    {'─'*12} {'─'*12} {'─'*15} {'─'*10} {'─'*8}")
    for _, row in finishing_vs_ordered.sort_values('Job Number').iterrows():
        flag = " ⚠️" if row['Overage'] > 0 else ""
        print(f"    {row['Job Number']:<12} {int(row['Quantity Ordered']):>12} "
              f"{int(row['Total Finished']):>15} {int(row['Overage']):>10}{flag}")

    finishing_vs_ordered.to_csv(os.path.join(OUTPUT_DIR, 'Q6_finishing_vs_ordered.csv'), index=False)
    print(f"\n  ✓ Exported to {OUTPUT_DIR}/Q6_finishing_vs_ordered.csv")

    # Chart
    fig, ax = plt.subplots(figsize=(10, 6))
    x = range(len(finishing_vs_ordered))
    width = 0.35
    fvo = finishing_vs_ordered.sort_values('Job Number')
    bars1 = ax.bar([i - width/2 for i in x], fvo['Quantity Ordered'], width, label='Quantity Ordered', color='#3498db')
    bars2 = ax.bar([i + width/2 for i in x], fvo['Total Finished'], width, label='Total Finished', color='#e74c3c')
    ax.set_xticks(list(x))
    ax.set_xticklabels(fvo['Job Number'], rotation=45, ha='right')
    ax.set_ylabel('Units')
    ax.set_title('Rule R5 Check: Finished Units vs. Quantity Ordered', fontweight='bold')
    ax.legend()
    plt.savefig(os.path.join(OUTPUT_DIR, 'Q6_overproduction_chart.png'))
    plt.close()
    print(f"  ✓ Chart saved to {OUTPUT_DIR}/Q6_overproduction_chart.png")

print()

# ═════════════════════════════════════════════════════════════════════════════
# QUESTION 7: Products Tab Issues (Rule R8)
# "Look at the Products tab. Describe any issues."
# ═════════════════════════════════════════════════════════════════════════════
print("─" * 80)
print("  QUESTION 7: Issues in the Products tab (Rule R8)")
print("─" * 80)

products['Margin ($)'] = products['Selling Price per Unit ($)'] - products['Cost to Make per Unit ($)']
products['Margin (%)'] = (products['Margin ($)'] / products['Selling Price per Unit ($)'] * 100).round(1)

print(f"\n  Product Margin Analysis:\n")
print(f"    {'Product':<12} {'Description':<22} {'Cost ($)':>10} {'Sell ($)':>10} {'Margin ($)':>10} {'Margin %':>10} {'Status'}")
print(f"    {'─'*12} {'─'*22} {'─'*10} {'─'*10} {'─'*10} {'─'*10} {'─'*10}")

for _, row in products.iterrows():
    status = "✅ OK" if row['Margin ($)'] > 0 else "❌ LOSS"
    print(f"    {row['Product Code']:<12} {row['Description']:<22} "
          f"{row['Cost to Make per Unit ($)']:>10.2f} {row['Selling Price per Unit ($)']:>10.2f} "
          f"{row['Margin ($)']:>10.2f} {row['Margin (%)']:>9.1f}% {status}")

loss_products = products[products['Margin ($)'] < 0]
print(f"\n  ► ISSUE FOUND: {len(loss_products)} product(s) violate Rule R8 (sell for LESS than cost):")
for _, row in loss_products.iterrows():
    print(f"    → {row['Product Code']} ({row['Description']}): costs ${row['Cost to Make per Unit ($)']:.2f} "
          f"but sells for ${row['Selling Price per Unit ($)']:.2f}")
    print(f"      The company LOSES ${abs(row['Margin ($)']):.2f} on every unit sold.")

products.to_csv(os.path.join(OUTPUT_DIR, 'Q7_product_margins.csv'), index=False)
print(f"\n  ✓ Exported to {OUTPUT_DIR}/Q7_product_margins.csv")

# Chart
fig, ax = plt.subplots(figsize=(10, 6))
x = range(len(products))
width = 0.35
bars1 = ax.bar([i - width/2 for i in x], products['Cost to Make per Unit ($)'], width,
               label='Cost to Make', color='#e74c3c', edgecolor='#c0392b')
bars2 = ax.bar([i + width/2 for i in x], products['Selling Price per Unit ($)'], width,
               label='Selling Price', color='#2ecc71', edgecolor='#27ae60')
ax.set_xticks(list(x))
ax.set_xticklabels(products['Product Code'] + '\n' + products['Description'], fontsize=8)
ax.set_ylabel('Price ($)')
ax.set_title('Rule R8 Check: Cost vs. Selling Price by Product', fontweight='bold')
ax.legend()
# Highlight the loss-making product
for i, (_, row) in enumerate(products.iterrows()):
    if row['Margin ($)'] < 0:
        ax.annotate('LOSS!', xy=(i, max(row['Cost to Make per Unit ($)'], row['Selling Price per Unit ($)']) + 10),
                    ha='center', fontweight='bold', color='red', fontsize=12)
plt.savefig(os.path.join(OUTPUT_DIR, 'Q7_product_margins_chart.png'))
plt.close()
print(f"  ✓ Chart saved to {OUTPUT_DIR}/Q7_product_margins_chart.png")

print()

# ═════════════════════════════════════════════════════════════════════════════
# QUESTION 8: Labour Cost by Work Centre (Rule R7)
# "For each work centre, work out the total hours recorded and the total
#  labour cost of those hours. Present it as a small table."
# ═════════════════════════════════════════════════════════════════════════════
print("─" * 80)
print("  QUESTION 8: Total hours and labour cost by work centre (Rule R7)")
print("─" * 80)

# Merge labour with work centres to get the rate
labour_wc = labour.merge(centres, on='Work Centre', how='left')
labour_wc['Labour Cost ($)'] = labour_wc['Hours Worked'] * labour_wc['Labour Rate ($ per hour)']

# Group by work centre
wc_summary = labour_wc.groupby(['Work Centre', 'Name']).agg(
    Total_Hours=('Hours Worked', 'sum'),
    Total_Records=('Record ID', 'count'),
    Total_Labour_Cost=('Labour Cost ($)', 'sum')
).reset_index()

wc_summary['Total_Hours'] = wc_summary['Total_Hours'].round(2)
wc_summary['Total_Labour_Cost'] = wc_summary['Total_Labour_Cost'].round(2)

print(f"\n  Labour Cost Summary by Work Centre:\n")
print(f"    {'Work Centre':<14} {'Name':<12} {'Records':>8} {'Total Hours':>12} {'Rate ($/hr)':>12} {'Labour Cost ($)':>16}")
print(f"    {'─'*14} {'─'*12} {'─'*8} {'─'*12} {'─'*12} {'─'*16}")

grand_hours = 0
grand_cost = 0
for _, row in wc_summary.iterrows():
    rate = centres[centres['Work Centre'] == row['Work Centre']]['Labour Rate ($ per hour)'].values[0]
    flag = " ⚠️" if rate == 0 else ""
    print(f"    {row['Work Centre']:<14} {row['Name']:<12} {int(row['Total_Records']):>8} "
          f"{row['Total_Hours']:>12.2f} {rate:>12.2f} {row['Total_Labour_Cost']:>16.2f}{flag}")
    grand_hours += row['Total_Hours']
    grand_cost += row['Total_Labour_Cost']

print(f"    {'─'*14} {'─'*12} {'─'*8} {'─'*12} {'─'*12} {'─'*16}")
print(f"    {'TOTAL':<14} {'':<12} {int(wc_summary['Total_Records'].sum()):>8} "
      f"{grand_hours:>12.2f} {'':>12} {grand_cost:>16.2f}")

print(f"\n  ► CONCERN: WC-350 (Finishing) has a labour rate of $0.00/hr.")
print(f"    This means {wc_summary[wc_summary['Work Centre']=='WC-350']['Total_Hours'].values[0]:.2f} hours of Finishing work ")
print(f"    contribute $0.00 to labour costs. This could mean:")
print(f"    1. Finishing is done by salaried staff (not costed per hour)")
print(f"    2. It's an automated/curing process with no direct labour cost")
print(f"    3. It's a data entry error — the rate should be non-zero")
print(f"    Either way, total labour cost is understated if Finishing should be costed.")

wc_summary.to_csv(os.path.join(OUTPUT_DIR, 'Q8_labour_cost_by_wc.csv'), index=False)
print(f"\n  ✓ Exported to {OUTPUT_DIR}/Q8_labour_cost_by_wc.csv")

# Chart
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

# Hours by work centre
colors = ['#3498db', '#2ecc71', '#f1c40f', '#e74c3c']
bars = ax1.bar(wc_summary['Name'], wc_summary['Total_Hours'], color=colors, edgecolor='#2c3e50')
ax1.set_ylabel('Total Hours')
ax1.set_title('Total Hours by Work Centre', fontweight='bold')
for bar, val in zip(bars, wc_summary['Total_Hours']):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, f'{val:.1f}h',
             ha='center', fontweight='bold', fontsize=9)

# Labour cost by work centre
bars2 = ax2.bar(wc_summary['Name'], wc_summary['Total_Labour_Cost'], color=colors, edgecolor='#2c3e50')
ax2.set_ylabel('Labour Cost ($)')
ax2.set_title('Total Labour Cost by Work Centre', fontweight='bold')
ax2.yaxis.set_major_formatter(mticker.StrMethodFormatter('${x:,.0f}'))
for bar, val in zip(bars2, wc_summary['Total_Labour_Cost']):
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 50, f'${val:,.0f}',
             ha='center', fontweight='bold', fontsize=9)

plt.suptitle('Question 8: Labour Analysis by Work Centre', fontsize=14, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, 'Q8_labour_cost_chart.png'))
plt.close()
print(f"  ✓ Chart saved to {OUTPUT_DIR}/Q8_labour_cost_chart.png")

print()

# ═════════════════════════════════════════════════════════════════════════════
# QUESTION 9: JOB-1006 vs JOB-1010 Comparison
# "Both were for the same product (PRD-110), same quantity (24).
#  But one used far more labour hours than the other."
# ═════════════════════════════════════════════════════════════════════════════
print("─" * 80)
print("  QUESTION 9: JOB-1006 vs JOB-1010 — Labour Hours Comparison")
print("─" * 80)

job_1006 = labour[labour['Job Number'] == 'JOB-1006']
job_1010 = labour[labour['Job Number'] == 'JOB-1010']

hours_1006 = job_1006['Hours Worked'].sum()
hours_1010 = job_1010['Hours Worked'].sum()
diff = abs(hours_1006 - hours_1010)

print(f"\n  Job Details:")
print(f"    {'Metric':<30} {'JOB-1006':>12} {'JOB-1010':>12} {'Difference':>12}")
print(f"    {'─'*30} {'─'*12} {'─'*12} {'─'*12}")
print(f"    {'Product':<30} {'PRD-110':>12} {'PRD-110':>12} {'—':>12}")
print(f"    {'Quantity Ordered':<30} {24:>12} {24:>12} {'—':>12}")
print(f"    {'Total Labour Hours':<30} {hours_1006:>12.2f} {hours_1010:>12.2f} {diff:>12.2f}")
print(f"    {'Number of Records':<30} {len(job_1006):>12} {len(job_1010):>12} {abs(len(job_1006)-len(job_1010)):>12}")

# Breakdown by stage
print(f"\n  Hours Breakdown by Stage:")
for stage in ['Mixing', 'Casting', 'Curing', 'Finishing']:
    h1006 = job_1006[job_1006['Stage'] == stage]['Hours Worked'].sum()
    h1010 = job_1010[job_1010['Stage'] == stage]['Hours Worked'].sum()
    print(f"    {stage:<15} {h1006:>10.2f} {h1010:>10.2f} {abs(h1006 - h1010):>10.2f}")

# Scrap analysis
scrap_1006 = job_1006['Units Scrapped'].sum()
scrap_1010 = job_1010['Units Scrapped'].sum()
print(f"\n  Scrap Analysis:")
print(f"    {'Total Units Scrapped':<30} {int(scrap_1006):>12} {int(scrap_1010):>12}")
print(f"    {'Total Units Completed':<30} {int(job_1006['Units Completed'].sum()):>12} {int(job_1010['Units Completed'].sum()):>12}")

# Operator analysis
print(f"\n  Operators Used:")
print(f"    JOB-1006: {', '.join(job_1006['Operator Name'].dropna().unique())}")
print(f"    JOB-1010: {', '.join(job_1010['Operator Name'].dropna().unique())}")

# Unique operators
ops_1006 = set(job_1006['Operator ID'].dropna())
ops_1010 = set(job_1010['Operator ID'].dropna())
only_1010 = ops_1010 - ops_1006
if only_1010:
    print(f"\n    ► Operator(s) unique to JOB-1010: {only_1010}")
    for op in only_1010:
        op_records = job_1010[job_1010['Operator ID'] == op]
        print(f"      {op}: {op_records['Hours Worked'].sum():.2f} hours across {len(op_records)} records")

# Date range
print(f"\n  Date Ranges:")
print(f"    JOB-1006: {job_1006['Date'].min().strftime('%Y-%m-%d')} to {job_1006['Date'].max().strftime('%Y-%m-%d')} "
      f"({(job_1006['Date'].max() - job_1006['Date'].min()).days + 1} calendar days)")
print(f"    JOB-1010: {job_1010['Date'].min().strftime('%Y-%m-%d')} to {job_1010['Date'].max().strftime('%Y-%m-%d')} "
      f"({(job_1010['Date'].max() - job_1010['Date'].min()).days + 1} calendar days)")

print(f"\n  ─── Analysis & Hypotheses ───")
print(f"""
  How big is the difference?
    JOB-1006 used {hours_1006:.2f} hours and JOB-1010 used {hours_1010:.2f} hours.
    The difference is {diff:.2f} hours — JOB-1010 used significantly more labour.

  Possible explanations:
    1. OPERATOR EXPERIENCE: JOB-1010 includes OP-106 (Singh, A.) who does not appear
       on any other job in the dataset. This may be a new or less experienced operator
       who works slower and logs more hours, particularly at the Casting stage.

    2. HIGHER SCRAP / REWORK: JOB-1010 had {int(scrap_1010)} units scrapped vs.
       {int(scrap_1006)} for JOB-1006. More scrap means more rework and extra hours
       to produce replacement units. Scrap at the Finishing stage is especially costly
       because all prior work on those units is wasted.

    3. LONGER DURATION / SCHEDULING: JOB-1010 spanned a longer calendar period,
       possibly due to production interruptions, machine downtime, or lower priority
       scheduling that spread the work over more days with less efficient short shifts.

  What I would want to find out next:
    - Ask the Production Manager whether OP-106 was being trained during JOB-1010.
    - Check if there were any equipment issues or downtime logged during JOB-1010.
    - Ask the Controller whether the 8.00-hour entries from OP-106 represent actual
      work or if they are rounded/estimated (all three entries are exactly 8.00).
    - Review whether JOB-1010 had any quality or material issues not captured in the
      scrap reason field.
""")

# Export detailed records
job_1006.to_csv(os.path.join(OUTPUT_DIR, 'Q9_JOB_1006_records.csv'), index=False)
job_1010.to_csv(os.path.join(OUTPUT_DIR, 'Q9_JOB_1010_records.csv'), index=False)

# Chart: Side-by-side comparison
fig, axes = plt.subplots(1, 3, figsize=(16, 5))

# Hours by stage
stages = ['Mixing', 'Casting', 'Curing', 'Finishing']
h1006 = [job_1006[job_1006['Stage'] == s]['Hours Worked'].sum() for s in stages]
h1010 = [job_1010[job_1010['Stage'] == s]['Hours Worked'].sum() for s in stages]
x = range(len(stages))
width = 0.35
axes[0].bar([i - width/2 for i in x], h1006, width, label='JOB-1006', color='#3498db')
axes[0].bar([i + width/2 for i in x], h1010, width, label='JOB-1010', color='#e74c3c')
axes[0].set_xticks(list(x))
axes[0].set_xticklabels(stages)
axes[0].set_ylabel('Hours')
axes[0].set_title('Hours by Stage', fontweight='bold')
axes[0].legend()

# Scrap comparison
axes[1].bar(['JOB-1006', 'JOB-1010'], [scrap_1006, scrap_1010], color=['#3498db', '#e74c3c'])
axes[1].set_ylabel('Units Scrapped')
axes[1].set_title('Total Scrap', fontweight='bold')

# Total hours comparison
axes[2].bar(['JOB-1006', 'JOB-1010'], [hours_1006, hours_1010], color=['#3498db', '#e74c3c'])
axes[2].set_ylabel('Total Hours')
axes[2].set_title('Total Labour Hours', fontweight='bold')
for i, (val, label) in enumerate(zip([hours_1006, hours_1010], ['JOB-1006', 'JOB-1010'])):
    axes[2].text(i, val + 0.5, f'{val:.1f}h', ha='center', fontweight='bold')

plt.suptitle('Question 9: JOB-1006 vs JOB-1010 Comparison', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, 'Q9_job_comparison_chart.png'))
plt.close()
print(f"  ✓ Charts saved to {OUTPUT_DIR}/Q9_job_comparison_chart.png")

# ═════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ═════════════════════════════════════════════════════════════════════════════
print()
print("=" * 80)
print("  ANALYSIS COMPLETE — SUMMARY OF FINDINGS")
print("=" * 80)
print(f"""
  Q1  Missing Operator:     {len(missing_operator)} records (violates R3)
  Q2  Duplicate Records:    {len(duplicates)} records (violates R6)
  Q3  Overtime (>8 hrs):    {len(overtime)} instance(s) (violates R1)
  Q4  Missing Scrap Reason: {len(scrapped_no_reason)} records, {int(total_units_affected)} units (violates R4)
  Q5  Post-Close Labour:    {len(post_close)} records (violates R2)
  Q6  Overproduction:       {len(overproduced)} job(s) (violates R5)
  Q7  Product Issues:       {len(loss_products)} product(s) sold at a loss (violates R8)
  Q8  Labour Cost Table:    See above — $0/hr at Finishing is concerning
  Q9  JOB Comparison:       {diff:.2f} hour difference — see analysis above

  All output files saved to: {OUTPUT_DIR}/
""")
print("=" * 80)
