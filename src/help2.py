import pandas as pd

# Load the file
df = pd.read_excel(
    "out/ISCC_Certificates_03.02.2026_13.40.xlsx",
    sheet_name="Certificates Changed"
)

# Remove any rows where Value_Changed == "Company_Name, City"
df = df[df["Value_Changed"] != "Company_Name, City"].copy()

# OPTIONAL: remove rows with whitespace variations e.g. "Company_Name,   City"
# df = df[~df["Value_Changed"].str.contains(r"^Company_Name,\s*City$", na=False)].copy()

# Save back to Excel
df.to_excel(
    "out/ISCC_Certificates_03.02.2026_13.40.xlsx",
    index=False
)

print("Saved cleaned file ✔")