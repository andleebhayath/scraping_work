import pandas as pd
import sqlalchemy
import urllib
# load dataset 
inputdata = 'data-raw/facebookgroupdata.csv'
df= pd.read_csv(inputdata)
# df.columns
# rename columns with 
df = df.rename(columns={
    'x1i10hfl href': 'tracking_url',
    'x1i10hfl': 'group_name',
    'x1i10hfl href 2': 'group_url',
    'x1lliihq': 'metadata',

})
# split the metadate columns into 3 parts 
metadata_split = df['metadata'].str.split(' · ', expand=True)
df['privacy'] = metadata_split[0]
df['member_count_raw'] = metadata_split[1]
df['posts_activity_raw'] = metadata_split[2]

# clean member count columns and extract only the numbers
def clean_members(val):
    if pd.isna(val): return 0
    val = val.replace(' members', '').replace(' member', '').strip()
    if 'K' in val:
        return int(float(val.replace('K', '')) * 1000)
    if 'M' in val:
        return int(float(val.replace('M', '')) * 1000000)
    return int(val.replace(',', ''))
# apply the cleaning function to the member count column
df['member_count'] = df['member_count_raw'].apply(clean_members)
# clean posts activity column and extract only the numbers
df['posts_per_day'] = df['posts_activity_raw'].str.extract('(\\d+)').astype(float)
# select only relevant columns for final output 
df_final = df[['group_name', 'group_url', 'privacy', 'member_count', 'posts_per_day']]
# save the cleaned dataset to a new csv file
#df_final.to_csv('data-processed/facebook_groups_cleaned.csv', index=False)
# save final dataset in excel 
#df_final.to_excel('data-processed/facebook_groups_cleaned.xlsx', index=False)
# create the connection string 
server = 'ICM-PK-ANDLEEBR'
database = 'Warehouse'
conn_str = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"Trusted_Connection=yes;"
)
# create the connection engine 
quoted_conn_str = urllib.parse.quote_plus(conn_str)
engine = sqlalchemy.create_engine(f"mssql+pyodbc:///?odbc_connect={quoted_conn_str}")
# write the dataframe to server 
try:
    df_final.to_sql(
        name='facebook_groups',
        con=engine,
        if_exists='replace',
        index=False
    )
    print("Data successfully written to the database.")
except Exception as e:
    print(f"An error occurred: {e}")
    