import pandas as pd
import os

# Set file paths
input_file = "/Users/saimanpokhrel/Desktop/Data/Book1.xlsx"
output_folder = "/Users/saimanpokhrel/Desktop/Data/Processed"

# Create output directory if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Load the Excel file
df = pd.read_excel(input_file)

# Clean column names (remove newline characters and extra whitespace)
df.columns = df.columns.str.replace(r'\s+', ' ', regex=True).str.strip()

# Define the rating map
rating_map = {
    "Strongly Disagree": 1,
    "Disagree": 2,
    "Neutral": 3,
    "Agree": 4,
    "Strongly Agree": 5
}

# Fill rating columns based on text responses
df["Rating of The learning outcome of the lesson was clearly defined."] = (
    df["The learning outcome of the lesson was clearly defined."].map(rating_map)
)

df["Rating of The curiosity section made students interested in the topic."] = (
    df["The curiosity section made students interested in the topic."].map(rating_map)
)

df["Rating of The learning outcome is aligned with the curriculum."] = (
    df["The learning outcome is aligned with the curriculum."].map(rating_map)
)

# Extract Class and Lesson ID
df["Class"] = df["Select the name of the Karkhana lesson"].str.extract(r'G(\d+)-')[0]
df["Lesson ID"] = df["Select the name of the Karkhana lesson"].str.extract(r'G\d+-(.+)')[0].str.strip()

# Create summary by Class and Lesson
summary = df.groupby(["Class", "Lesson ID"])[[
    "Rating of The learning outcome of the lesson was clearly defined.",
    "Rating of The curiosity section made students interested in the topic.",
    "Rating of The learning outcome is aligned with the curriculum."
]].agg(['mean', 'count'])

# Save to Excel files in the output folder
cleaned_file_path = os.path.join(output_folder, "Cleaned_Karkhana_Data.xlsx")
summary_file_path = os.path.join(output_folder, "Karkhana_Rating_Summary.xlsx")

df.to_excel(cleaned_file_path, index=False)
summary.to_excel(summary_file_path)

print("✅ Cleaned data saved to:", cleaned_file_path)
print("✅ Summary saved to:", summary_file_path)
