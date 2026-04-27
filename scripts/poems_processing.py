import pandas as pd

# 1. Load the input file
input_file = 'data-raw/urdu_potery.xlsx'
df = pd.read_excel(input_file)

# 2. Clean data: Remove empty rows or rows that only contain newline characters
def is_empty(val):
    if pd.isna(val): return True
    s = str(val).strip()
    return s == "" or s == "\n"

# Filter rows where 'column 1' or 'column 2' are not empty
df = df[~(df['column 1'].apply(is_empty) & df['column 2'].apply(is_empty))]
df = df.reset_index(drop=True)

# 3. Process each poem group
output_rows = []

# Group by 'column 3' while maintaining the original order
for poem_id, group in df.groupby('column 3', sort=False):
   
    # The new title is the first hemistich (column 1) of the poem
    new_title = str(group.iloc[0]['column 1']).strip()
    
    
    # Interleave column 1 and column 2 for the full content
    lines = []
    for _, row in group.iterrows():
        lines.append(str(row['column 1']).strip())
        lines.append(str(row['column 2']).strip())
    
    full_content = "\n".join(lines)
    output_rows.append({'title': new_title, 'content': full_content})

# 4. Create the final DataFrame and save to CSV
df_final = pd.DataFrame(output_rows)
df_final.to_csv("data-processed/transformed_poems.csv", index=False)
#Save as Excel
df_final.to_excel("data-processed/transformed_poems.xlsx", index=False)
