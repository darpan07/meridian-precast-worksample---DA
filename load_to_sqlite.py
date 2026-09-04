import pandas as pd
import sqlite3
import os

DATA_FILE = r'Meridian Precast - Production Data.xlsx'
DB_FILE   = r'meridian_precast.db'

# Remove existing database
if os.path.exists(DB_FILE):
    os.remove(DB_FILE)

# Connect to SQLite
conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()


labour = pd.read_excel(DATA_FILE, sheet_name='Labour Records')
labour.columns = [c.lower().replace(' ', '_') for c in labour.columns]
labour.to_sql('labour_records', conn, index=False, if_exists='replace')
print(f"Loaded {len(labour)} labour records")

# Load Jobs
jobs = pd.read_excel(DATA_FILE, sheet_name='Jobs')
jobs.columns = [c.lower().replace(' ', '_') for c in jobs.columns]
jobs.to_sql('jobs', conn, index=False, if_exists='replace')
print(f"Loaded {len(jobs)} jobs")

# Load Products
products = pd.read_excel(DATA_FILE, sheet_name='Products')
products.columns = ['product_code', 'description', 'cost_per_unit', 'sell_price_per_unit']
products.to_sql('products', conn, index=False, if_exists='replace')
print(f"Loaded {len(products)} products")

# Load Work Centres 
centres = pd.read_excel(DATA_FILE, sheet_name='Work Centres')
centres.columns = ['work_centre', 'name', 'labour_rate']
centres.to_sql('work_centres', conn, index=False, if_exists='replace')
print(f"Loaded {len(centres)} work centres")

conn.close()
print(f"Database saved as: {DB_FILE}")
print("You can now run the SQL queries from analysis.sql against this database.")
print("Example: sqlite3 meridian_precast.db < analysis.sql")
