import pandas as pd

from payroll_permit_verifier.settings import REPORT_HEADER_ROW
from payroll_permit_verifier.excel_io import read_excel_file
from payroll_permit_verifier.cleaning import clean_column_names, clean_employee_id


def load_fsam_reports(fam_file, sam_file):
    # Reads excel files into pandas dataframes, returned values as strings (added header=2 to temp work for temp files).
    fam_df = read_excel_file(fam_file, header=REPORT_HEADER_ROW)
    sam_df = read_excel_file(sam_file, header=REPORT_HEADER_ROW)

    fam_df = clean_column_names(fam_df)
    sam_df = clean_column_names(sam_df)

    # Add source column so you know where each row came from.
    fam_df["SourceReport"] = "FAM"
    sam_df["SourceReport"] = "SAM"

    # Combine the files into one, for better data comparison, vertical stack + reset record indexes.
    combined_df = pd.concat([fam_df, sam_df], ignore_index=True)

    # Create standard EmployeeID column from FSAM's employee ID field
    combined_df["EmployeeID"] = clean_employee_id(combined_df["Empl ID"])

    report_summary_mask = combined_df["Name"].astype("string").str.contains("Sub-Total|REPORT TOTALS", case=False, na=False)

    fsam_report_summary_rows = combined_df[report_summary_mask].copy()
    combined_df = combined_df[~report_summary_mask].copy()

    # Check which rows have a real EmployeeID
    valid_employee_id_mask = combined_df["EmployeeID"].str.match(r"^\d", na=False)

    # Rows with no/invalid EmployeeID, so we can report them instead of silently removing them
    fsam_invalid_employee_ids = combined_df[
        ~valid_employee_id_mask
    ].copy()

    # Keep only rows where EmployeeID starts with a number
    combined_df = combined_df[
        valid_employee_id_mask
    ].copy()

    return combined_df, fsam_invalid_employee_ids, fsam_report_summary_rows


def load_local_report(local_file):
    # This LOCAL report has title/printed rows above the real header.
    # Excel row 6 is the real header, so Python uses header=5.
    local_df = read_excel_file(local_file, header=5)

    # Remove fully blank columns caused by weird Excel formatting / merged cells
    local_df = local_df.dropna(axis=1, how="all")

    # Remove fully blank spacer rows between records
    local_df = local_df.dropna(axis=0, how="all")

    # Clean field/header names, ex: "First Name:" -> "First Name"
    local_df.columns = (local_df.columns.astype(str).str.strip().str.replace(":", "", regex=False))

    # Keep only columns we actually care about
    local_df = local_df[
        [
            "Mac Emp No",
            "Mohawk Emp No",
            "First Name",
            "Last Name",
            "AIMs Acc",
            "Lot Code",
            "Loc",
            "Sec"
        ]
    ].copy()

    # Clean AIMs account number so LOCAL accounts can be compared against Mohawk accounts
    local_df["CleanAIMsAcc"] = clean_employee_id(local_df["AIMs Acc"])

    local_df["EmployeeID"] = clean_employee_id(local_df["Mac Emp No"])
    local_df["MohawkEmployeeID"] = clean_employee_id(local_df["Mohawk Emp No"])

    valid_mac_employee_id_mask = local_df["EmployeeID"].str.match(r"^\d", na=False)
    valid_mohawk_employee_id_mask = local_df["MohawkEmployeeID"].str.match(r"^\d", na=False)

    # Row is invalid only if it has neither Mac Emp No nor Mohawk Emp No
    valid_employee_id_mask = valid_mac_employee_id_mask | valid_mohawk_employee_id_mask
    invalid_employee_id_mask = ~valid_employee_id_mask

    # This is the no/invalid employee ID list
    local_invalid_employee_ids = local_df[invalid_employee_id_mask].copy()

    # Keep rows that have either Mac or Mohawk ID
    local_df = local_df[valid_employee_id_mask].copy()

    # Use Mac Emp No for comparison if it exists, otherwise use Mohawk Emp No
    local_df["CompareEmployeeID"] = local_df["EmployeeID"].where(
        local_df["EmployeeID"].str.match(r"^\d", na=False),
        local_df["MohawkEmployeeID"]
    )

    return local_df, local_invalid_employee_ids

def load_mohawk_report(mohawk_file):
    # Mohawk sheet only has AIMs account numbers, so this is used for account comparison only
    mohawk_df = read_excel_file(mohawk_file, header=0)

    mohawk_df = mohawk_df.dropna(axis=1, how="all")
    mohawk_df = mohawk_df.dropna(axis=0, how="all")
    mohawk_df = clean_column_names(mohawk_df)

    mohawk_df["CleanMohawkAcc"] = clean_employee_id(mohawk_df["ACCOUNT"])

    valid_account_mask = mohawk_df["CleanMohawkAcc"].str.match(r"^\d", na=False)
    mohawk_invalid_accounts = mohawk_df[~valid_account_mask].copy()
    mohawk_df = mohawk_df[valid_account_mask].copy()

    return mohawk_df, mohawk_invalid_accounts
