import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).parent # where python file lives (cd ..)
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

OUTPUT_DIR.mkdir(exist_ok=True) # If not exist, create output folder. in brackets: don't crash if it exists already

FAM_DIR = BASE_DIR / "input" / "FAM"
SAM_DIR = BASE_DIR / "input" / "SAM"
LOCAL_DIR = BASE_DIR / "input" / "LOCAL"

def get_first_excel_file(folder, label): # helper function to get whatever excel file is in each folder
    files = list(folder.glob("*.xlsx"))

    if not files:
        raise FileNotFoundError(f"No {label} file found in {folder}")

    return files[0]


fam_file = get_first_excel_file(FAM_DIR, "FAM")
sam_file = get_first_excel_file(SAM_DIR, "SAM")

# Reads excel files into pandas dataframes, returned values as strings.
fam_df = pd.read_excel(fam_file, dtype=str)
sam_df = pd.read_excel(sam_file, dtype=str)

# Add source column so you know where each row came from.
fam_df["SourceReport"] = "FAM"
sam_df["SourceReport"] = "SAM"

# Combine the files into one, for better data comparison, vertical stack + reset record indexes.
combined_df = pd.concat([fam_df, sam_df], ignore_index=True)

# # put to excel and save in INPUT_DIR
# combined_df.to_excel(combined_output_file, index=False)

local_file = get_first_excel_file(LOCAL_DIR, "LOCAL")

# header=None => do not assume clean header row. Can fix later when fixing report style.
local_raw = pd.read_excel(local_file, header=None, dtype=str)

#          ( Grab first column ) (convert to: string) (rm trailing)
first_col = local_raw.iloc[:, 0].astype("string").str.strip()

# Keep only rows with a number (starts with one, at least)
local_df = local_raw[first_col.str.match(r"^\d", na=False)].copy()

# Either local_raw[EmployeeID"],
# or local_raw.iloc[:, 0] (if no header)


# # -----------------------------
# # File paths
# # -----------------------------
# INPUT_DIR = Path("input")
# OUTPUT_DIR = Path("output")
# OUTPUT_DIR.mkdir(exist_ok=True)

# fam_file = INPUT_DIR / "FAMTest1.xlsx"
# sam_file = INPUT_DIR / "SAMTest1.xlsx"
# local_file = INPUT_DIR / "local_active_permits_example.xlsx"

# combined_hr_output = OUTPUT_DIR / "combined_hr_accounts.xlsx"
# missing_from_hr_output = OUTPUT_DIR / "accounts_left_out_of_hr_db.xlsx"
# missing_from_local_output = OUTPUT_DIR / "accounts_left_out_of_local_db.xlsx"


# # -----------------------------
# # Settings
# # -----------------------------
# EMPLOYEE_ID_FIELD = "EmployeeID"

# # -----------------------------
# # Helper function
# # -----------------------------
# def clean_employee_id(series):
#     """
#     Converts employee numbers to clean strings.
#     This helps avoid issues like:
#     12345.0 vs 12345
#     spaces before/after values
#     missing values
#     """
#     return (
#         series
#         .astype("string")
#         .str.strip()
#         .str.replace(r"\.0$", "", regex=True)
#     )


# # -----------------------------
# # 1. Read FAM and SAM
# # -----------------------------
# fam_df = pd.read_excel(fam_file, dtype=str)
# sam_df = pd.read_excel(sam_file, dtype=str)

# # Add source column so you know where each row came from
# fam_df["SourceReport"] = "FAM"
# sam_df["SourceReport"] = "SAM"


# # -----------------------------
# # 2. Combine FAM + SAM
# # -----------------------------
# combined_hr_df = pd.concat(
#     [fam_df, sam_df],
#     ignore_index=True
# )

# # Clean employee IDs
# combined_hr_df[EMPLOYEE_ID_FIELD] = clean_employee_id(
#     combined_hr_df[EMPLOYEE_ID_FIELD]
# )

# # # Optional: remove duplicate employee accounts
# # combined_hr_df = combined_hr_df.drop_duplicates(
# #     subset=[EMPLOYEE_ID_FIELD]
# # )

# # Save combined HR file
# combined_hr_df.to_excel(combined_hr_output, index=False)


# # -----------------------------
# # 3. Read local active permit file
# # -----------------------------
# local_active_df = pd.read_excel(local_file, dtype=str)

# local_active_df[EMPLOYEE_ID_FIELD] = clean_employee_id(
#     local_active_df[EMPLOYEE_ID_FIELD]
# )

# # # If local file includes inactive permits too, filter active only
# # if STATUS_FIELD in local_df.columns:
# #     local_active_df = local_df[
# #         local_df[STATUS_FIELD].str.strip().str.lower() == ACTIVE_VALUE.lower()
# #     ]
# # else:
# #     local_active_df = local_df.copy()


# # -----------------------------
# # 4. Compare local vs HR using employee number
# # -----------------------------

# # Accounts in local permit system but NOT in HR file
# accounts_left_out_of_hr_db = local_active_df[
#     ~local_active_df[EMPLOYEE_ID_FIELD].isin(combined_hr_df[EMPLOYEE_ID_FIELD])
# ]

# # Accounts in HR file but NOT in local permit system
# accounts_left_out_of_local_db = combined_hr_df[
#     ~combined_hr_df[EMPLOYEE_ID_FIELD].isin(local_active_df[EMPLOYEE_ID_FIELD])
# ]


# # -----------------------------
# # 5. Save mismatch reports
# # -----------------------------
# accounts_left_out_of_hr_db.to_excel(
#     missing_from_hr_output,
#     index=False
# )

# accounts_left_out_of_local_db.to_excel(
#     missing_from_local_output,
#     index=False
# )


# # -----------------------------
# # 6. Summary
# # -----------------------------
# print("Done.")
# print(f"FAM rows: {len(fam_df)}")
# print(f"SAM rows: {len(sam_df)}")
# print(f"Combined HR rows: {len(combined_hr_df)}")
# print(f"Local active permit rows: {len(local_active_df)}")
# print(f"Missing from HR DB: {len(accounts_left_out_of_hr_db)}")
# print(f"Missing from Local DB: {len(accounts_left_out_of_local_db)}")