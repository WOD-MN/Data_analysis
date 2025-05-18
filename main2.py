import pandas as pd
import numpy as np
import os

# Load Excel file
file_path = "/Users/saimanpokhrel/Desktop/Data/Processed/final.xlsx"
xls = pd.ExcelFile(file_path)
df = xls.parse('Sheet1')

# Set the first row as header
df.columns = df.iloc[0]
df = df[1:]
df.columns.name = None

# Rename columns
df.columns = [
    "Class", "Chapter", "LO_Mean", "LO_Count",
    "Curiosity_Mean", "Curiosity_Count",
    "Alignment_Mean", "Alignment_Count"
]

# Fill missing values
df["Class"] = df["Class"].fillna(method='ffill')
df["Chapter"] = df["Chapter"].fillna(method='ffill')

# Convert rating columns to numeric
rating_cols = [
    "LO_Mean", "LO_Count",
    "Curiosity_Mean", "Curiosity_Count",
    "Alignment_Mean", "Alignment_Count"
]
df[rating_cols] = df[rating_cols].apply(pd.to_numeric, errors='coerce')

# Weighted average function
def safe_weighted_avg(group, mean_col, count_col):
    valid = group[count_col] > 0
    if valid.any():
        return np.average(group.loc[valid, mean_col], weights=group.loc[valid, count_col])
    else:
        return np.nan

# Group by Class and Chapter (original format retained)
summary = df.groupby(["Class", "Chapter"]).apply(
    lambda g: pd.Series({
        "LO_Mean": safe_weighted_avg(g, "LO_Mean", "LO_Count"),
        "LO_Count": g["LO_Count"].sum(),
        "Curiosity_Mean": safe_weighted_avg(g, "Curiosity_Mean", "Curiosity_Count"),
        "Curiosity_Count": g["Curiosity_Count"].sum(),
        "Alignment_Mean": safe_weighted_avg(g, "Alignment_Mean", "Alignment_Count"),
        "Alignment_Count": g["Alignment_Count"].sum()
    })
).reset_index()

# Define output path
output_path = os.path.join(os.path.dirname(file_path), "processed_chapter_summary.xlsx")

# Save the result
summary.to_excel(output_path, index=False)
print(f"✅ Processed file saved at: {output_path}")
