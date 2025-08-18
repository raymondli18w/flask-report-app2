import pandas as pd

# Load the Excel file
df = pd.read_excel("latest.xlsx")

# Convert 'Shipped Date' column to datetime, coercing errors to NaT (not a time)
df["Shipped Date"] = pd.to_datetime(df["Shipped Date"], errors='coerce')

# Filter rows where 'Shipped Date' is not null (i.e., has a valid date)
filtered_df = df[df["Shipped Date"].notnull()]

print("Filtered data preview:")
print(filtered_df.head())

# Save the filtered data to a new Excel file
filtered_df.to_excel("filtered_latest.xlsx", index=False)
print("Filtered data saved to filtered_latest.xlsx")
