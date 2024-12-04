import pandas as pd
import os
import re

def extract_month_from_filename(file_path):
    """
    Extract and return the month (MM) in the format '2024MM' from the file path.
    """
    match = re.search(r"2024(\d{2})", file_path)
    if match:
        return match.group(1)  
    return None 

def determine_prefix_from_filename(file_path):
    """
    Check if the file path contains 'OG' or 'PR' and return the corresponding string.
    """
    if "PR" in file_path:
        return "PR"
    else:
        return "Default" 
    
file_path = r"C:\Users\kyeon\Captis_Programming\Compliance Tracker Rev.2 - 202408 PR.xlsm"
month = extract_month_from_filename(file_path)
prefix = determine_prefix_from_filename(file_path)

if not month:
    raise ValueError("Cannot find valid '2024MM' format month information in the file name.")

df = pd.read_excel(file_path, sheet_name=0, header=None)

# Keep columns corresponding to positions 34 to 41 (index range 34 to 42)
df = df.iloc[:, 34:42]

# Drop rows where all values are NaN and reset the index
df = df.dropna(how='all').reset_index(drop=True)

# Copy the first 9 rows
df_first_10 = df.iloc[:9].copy()

# Retrieve the 9th row (index 8) from the original data
row_10 = df.iloc[8, :].values 

# Backup the first row data
row_1 = df_first_10.iloc[0, :].values  

# Swap the first row and the 9th row data
df_first_10.iloc[0, :] = row_10  
df.iloc[8, :] = row_1 

# Add filtering conditions: Check for specific strings in the first column
filter_conditions = df.iloc[:, 0].astype(str).str.contains(
    'monitored equipment:|monthly data omission %:|monthly exceedance %:', na=False
)
filtered_df = df[filter_conditions].reset_index(drop=True)

# Move the last row to the first row
last_row = filtered_df.iloc[-1, :].copy()  

# Remove the last row and append it as the first row
filtered_df = filtered_df.iloc[:-1, :]  
filtered_df = pd.concat([pd.DataFrame([last_row.values], columns=filtered_df.columns), filtered_df], ignore_index=True)

# Set the first row as column names
filtered_df.columns = filtered_df.iloc[0]
filtered_df = filtered_df[1:].reset_index(drop=True) 

# Update column names
filtered_df.columns = [
    "Category",
    "Pre-Treatment System FT110",
    "Pre-Treatment System FT120",
    "TOX Flow",
    "TOX Combustion",
    "GM01",
    "GM02",
    "GM03",
]
# Select columns from the second to the seventh and unpivot
unpivoted_df = pd.melt(
    filtered_df,
    id_vars=["Category"], 
    value_vars=[
        "Pre-Treatment System FT110",
        "Pre-Treatment System FT120",
        "TOX Flow",
        "TOX Combustion",
        "GM01",
        "GM02",
        "GM03",
    ],
    var_name="System Name",  
    value_name="Value",      
)
# Replace "monthly data omission %:" with "Data Omission" and "monthly exceedance %:" with "Exceedance"
unpivoted_df["Category"] = unpivoted_df["Category"].replace(
    {"monthly data omission %:": "Data Omission", "monthly exceedance %:": "Exceedance"}
)

# Convert decimal values to percentages and append the '%' symbol
unpivoted_df["Value"] = (unpivoted_df["Value"].astype(float) * 100).round(2).astype(str) + "%"

# Split data into separate dataframes
data_omission_df = unpivoted_df[unpivoted_df["Category"] == "Data Omission"].reset_index(drop=True)
exceedance_df = unpivoted_df[unpivoted_df["Category"] == "Exceedance"].reset_index(drop=True)

# Save directory
save_dir = r"C:\Users\kyeon\Captis_Programming"

# Set filenames based on the month and OG/PR conditions
data_omission_filename = f"{month}_{prefix}_Data_Omission.csv"
exceedance_filename = f"{month}_{prefix}_Exceedance.csv"

# Generate unique file paths
output_path_omission = os.path.join(save_dir, data_omission_filename)
output_path_exceedance = os.path.join(save_dir, exceedance_filename)

# Save the first file: Data Omission
data_omission_df.to_csv(output_path_omission, index=False, encoding='utf-8-sig')

# Save the second file: Exceedance
exceedance_df.to_csv(output_path_exceedance, index=False, encoding='utf-8-sig')

print(f"Data Omission File Completed: {output_path_omission}")
print(f"Exceedance File Completed: {output_path_exceedance}")
