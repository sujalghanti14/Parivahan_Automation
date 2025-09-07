import pandas as pd

# Load Excel file
file_path = r"C:\Users\sujal\Documents\Raam Group\Telangana market Analysis\TL 2-4W 2024.xlsx"  # change path if needed
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Rename columns for consistency
df = df.rename(columns={
    "makerName": "Maker Name",
    "OfficeCd": "RTA",
    "fuel": "Fuel Type",
    "vehicleClass": "Vehicle Type",
    "Month": "Month"
})

# Group by required columns and count
result = (
    df[["Maker Name", "RTA", "Fuel Type", "Vehicle Type", "Month"]]
    .groupby(["Maker Name", "RTA", "Fuel Type", "Vehicle Type", "Month"])
    .size()
    .reset_index(name="Count")
)

# Save to Excel
result.to_excel("Transformed_RTA_Data.xlsx", index=False)

print("✅ Transformation complete. File saved as Transformed_RTA_Data.xlsx")
