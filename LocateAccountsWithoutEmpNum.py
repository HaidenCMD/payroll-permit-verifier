# you see I want to take two very similar reports (FAM (Faculty) and SAM (Staff)) and combine them, and then verify that all of their accounts match up with another source

import pandas as pd

# Read input file
df = pd.read_csv("accounts.csv")

# Filter rows where EmployeeID is missing
missing_employee_id = df[df["EmployeeID"].isna()]

# Save those rows to a new CSV
missing_employee_id.to_csv("accounts_missing_employee_id.csv", index=False)

print("Done. Created accounts_missing_employee_id.csv")