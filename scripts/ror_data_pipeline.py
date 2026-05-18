import pandas as pd
import requests
import urllib.parse
from sqlalchemy import create_engine

# --- STEP 1: FETCH DATA (Ensuring it is flattened) ---
url = "https://api.ror.org/v2/organizations"
response = requests.get(url).json()
items = response.get('items', [])

# json_normalize is what creates the 'admin.created.date' columns
df = pd.json_normalize(items)

# --- STEP 2: CLEANING ---

# [Optional] Print columns if you still get errors to see the exact names:
# print(df.columns) 

# Extract Link
def extract_link(link_list):
    if isinstance(link_list, list) and len(link_list) > 0:
        return link_list[0].get('value')
    return None
df['website_url'] = df['links'].apply(extract_link)

# Extract Name
def extract_name(name_list):
    if isinstance(name_list, list) and len(name_list) > 0:
        for name in name_list:
            if 'ror_display' in name.get('types', []):
                return name.get('value')
        return name_list[0].get('value')
    return None
df['organization_name'] = df['names'].apply(extract_name)

# Extract Location
def extract_location(loc_list):
    if isinstance(loc_list, list) and len(loc_list) > 0:
        details = loc_list[0].get('geonames_details', {})
        return details.get('name')
    return None
df['city'] = df['locations'].apply(extract_location)

# Extract Type
df['org_type'] = df['types'].apply(lambda x: x[0] if isinstance(x, list) and len(x) > 0 else None)

# Fix Established Year
df['established'] = pd.to_numeric(df['established'], errors='coerce').fillna(0).astype(int)

# --- FIXING THE DATE ERROR ---
# We check if the column exists first. If not, we create a column of None/Null.

if 'admin.created.date' in df.columns:
    df['admin_created_date'] = pd.to_datetime(df['admin.created.date'])
else:
    df['admin_created_date'] = None

if 'admin.last_modified.date' in df.columns:
    # NOTE: Fixed typo here (you were assigning to admin_created_date again)
    df['admin_last_modified_date'] = pd.to_datetime(df['admin.last_modified.date'])
else:
    df['admin_last_modified_date'] = None

# --- STEP 3: FINAL SELECTION ---
columns_to_keep = [
    'id', 'organization_name', 'website_url', 'city', 'established', 
    'status', 'org_type', 'admin_created_date', 'admin_last_modified_date'
]

# Only keep columns that actually exist to avoid a new KeyError
final_cols = [c for c in columns_to_keep if c in df.columns]
df_clean = df[final_cols].copy()

print("Cleaning successful!")
print("Automatically visiting links to verify status...")

def visit_and_verify(url):
    if not url:
        return "No URL"
    try:
        # We use a 5-second timeout so one slow site doesn't hang the whole script
        # Using .head() is faster than .get() because it doesn't download the HTML content
        res = requests.head(url, timeout=5, allow_redirects=True)
        return f"Live (Status: {res.status_code})"
    except Exception as e:
        return f"Error: {type(e).__name__}"

# This will create a new column in your SQL table called 'link_check'
df_clean['link_check'] = df_clean['website_url'].apply(visit_and_verify)
print("Connecting to SQL Server...")

# Configuration - Change these!
SERVER_NAME = 'ICM-PK-ANDLEEBR'
DB_NAME = 'Warehouse'
TABLE_NAME = 'Cleaned_ROR_Data'

# Build the connection string (Windows Authentication)
params = urllib.parse.quote_plus(
    f'DRIVER={{ODBC Driver 17 for SQL Server}};'
    f'SERVER={SERVER_NAME};'
    f'DATABASE={DB_NAME};'
    f'Trusted_Connection=yes;'
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

# Upload the data
try:
    df_clean.to_sql(TABLE_NAME, engine, if_exists='replace', index=False)
    print(f"Success! {len(df_clean)} rows stored in table '{TABLE_NAME}'.")
except Exception as e:
    print(f"SQL Error: {e}")