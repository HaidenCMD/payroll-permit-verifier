import pandas as pd
from pathlib import Path
from io import BytesIO
import getpass
import msoffcrypto

BASE_DIR = Path(__file__).parent # where python file lives (cd ..)
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

FAM_DIR = BASE_DIR / "input" / "FAM"
SAM_DIR = BASE_DIR / "input" / "SAM"
MOHAWK_DIR = BASE_DIR / "input" / "MOHAWK"
LOCAL_DIR = BASE_DIR / "input" / "LOCAL"

# Temporary because template files currently have the real header on row 3.
# Change to 0 later if the real files start directly with the header row.
REPORT_HEADER_ROW = 2

def read_excel_file(file, header=REPORT_HEADER_ROW):
    password = None

    # Try reading normally first, for files that are not password-protected
    try:
        return pd.read_excel(file, header=header, dtype=str)

    except Exception as normal_error:
        # If normal read fails and no password was given, stop with a clearer error
        if password is None:
            raise RuntimeError(
                f"Could not read {file.name}. It may be password-protected."
            ) from normal_error

        # Decrypt password-protected Excel file into memory, then read it with pandas
        decrypted_file = BytesIO()

        with open(file, "rb") as encrypted_file:
            office_file = msoffcrypto.OfficeFile(encrypted_file)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_file)

        decrypted_file.seek(0)

        return pd.read_excel(decrypted_file, header=header, dtype=str)
    
def get_first_excel_file(folder, label): # helper function to get whatever excel file is in each folder
    files = [
        file for file in folder.glob("*.xlsx")
    ]

    if not files:
        raise FileNotFoundError(f"No {label} file found in {folder}")

    print(f"{label} file being used: {files[0].name}")

    return files[0]


def clean_column_names(df):
    # Clean field/header names, ex: "Mac Emp No " -> "Mac Emp No"
    df.columns = df.columns.astype(str).str.strip()
    return df


def clean_employee_id(series):
    # Convert to string, remove trailing spaces, and fix IDs that accidentally become 12345.0
    return (
        series
        .astype("string")
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)
        .str.lstrip("0")
    )


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

def compare_employee_ids(combined_df, local_df):
    # In LOCAL, but not in FSAM
    # Uses Mac Emp No if it exists, otherwise Mohawk Emp No
    local_not_in_fsam = local_df[
        ~local_df["CompareEmployeeID"].isin(combined_df["EmployeeID"])
    ].copy()

    # In FSAM, but not in LOCAL
    # Compares FSAM EmployeeID against LOCAL's selected CompareEmployeeID
    fsam_not_in_local = combined_df[
        ~combined_df["EmployeeID"].isin(local_df["CompareEmployeeID"])
    ].copy()

    return local_not_in_fsam, fsam_not_in_local

def compare_aims_accounts(local_df, mohawk_df):
    # In LOCAL, but not in MOHAWK account sheet
    local_not_in_mohawk = local_df[
        ~local_df["CleanAIMsAcc"].isin(mohawk_df["CleanMohawkAcc"])
    ].copy()

    # In MOHAWK account sheet, but not in LOCAL
    mohawk_not_in_local = mohawk_df[
        ~mohawk_df["CleanMohawkAcc"].isin(local_df["CleanAIMsAcc"])
    ].copy()

    return local_not_in_mohawk, mohawk_not_in_local

def save_reports(combined_df, local_not_in_fsam, fsam_not_in_local, fsam_invalid_employee_ids, local_invalid_employee_ids, fsam_report_summary_rows, output_dir):
    output_dir.mkdir(exist_ok=True)

    combined_fsam_file = output_dir / "combined_fsam.xlsx"
    local_not_in_fsam_file = output_dir / "local_not_in_fsam.xlsx"
    fsam_not_in_local_file = output_dir / "fsam_not_in_local.xlsx"
    fsam_invalid_employee_ids_file = output_dir / "fsam_invalid_employee_ids.xlsx"
    local_invalid_employee_ids_file = output_dir / "local_invalid_employee_ids.xlsx"
    fsam_report_summary_rows_file = output_dir / "fsam_report_summary_rows_ignored.xlsx"

    # Save combined FAM + SAM report
    combined_df.to_excel(combined_fsam_file, index=False)

    # Save mismatch reports
    local_not_in_fsam.to_excel(local_not_in_fsam_file, index=False)
    fsam_not_in_local.to_excel(fsam_not_in_local_file, index=False)

    # Save invalid/no employee ID reports
    fsam_invalid_employee_ids.to_excel(fsam_invalid_employee_ids_file, index=False)
    local_invalid_employee_ids.to_excel(local_invalid_employee_ids_file, index=False)

    # Save subtotal/report total rows that were ignored
    fsam_report_summary_rows.to_excel(fsam_report_summary_rows_file, index=False)

    return {
        "combined_fsam": combined_fsam_file,
        "local_not_in_fsam": local_not_in_fsam_file,
        "fsam_not_in_local": fsam_not_in_local_file,
        "fsam_invalid_employee_ids": fsam_invalid_employee_ids_file,
        "local_invalid_employee_ids": local_invalid_employee_ids_file,
        "fsam_report_summary_rows_ignored": fsam_report_summary_rows_file
    }


def run_verification(fam_file, sam_file, mohawk_file, local_file, output_dir):
    # This is the main reusable function.
    # Local script, Microsoft Graph, or AWS can all call this same function.

    # Load reports + return invalid employeeIDs
    combined_df, fsam_invalid_employee_ids, fsam_report_summary_rows = load_fsam_reports(fam_file, sam_file)
    local_df, local_invalid_employee_ids = load_local_report(local_file)
    mohawk_df, mohawk_invalid_accounts = load_mohawk_report(mohawk_file)

    # Compare existing employeeIDs of FSAM vs Local
    local_not_in_fsam, fsam_not_in_local = compare_employee_ids(combined_df, local_df)
    local_not_in_mohawk, mohawk_not_in_local = compare_aims_accounts(local_df, mohawk_df)

    local_not_in_mohawk_file = output_dir / "local_not_in_mohawk_accounts.xlsx"
    mohawk_not_in_local_file = output_dir / "mohawk_not_in_local_accounts.xlsx"
    mohawk_invalid_accounts_file = output_dir / "mohawk_invalid_accounts.xlsx"

    local_not_in_mohawk.to_excel(local_not_in_mohawk_file, index=False)
    mohawk_not_in_local.to_excel(mohawk_not_in_local_file, index=False)
    mohawk_invalid_accounts.to_excel(mohawk_invalid_accounts_file, index=False)

    output_files = save_reports(combined_df, local_not_in_fsam, fsam_not_in_local, fsam_invalid_employee_ids, local_invalid_employee_ids, fsam_report_summary_rows, output_dir)

    summary = {
        "fsam_rows_checked": len(combined_df),
        "local_rows_checked": len(local_df),
        "local_not_in_fsam": len(local_not_in_fsam),
        "fsam_not_in_local": len(fsam_not_in_local),
        "fsam_invalid_employee_ids": len(fsam_invalid_employee_ids),
        "local_invalid_employee_ids": len(local_invalid_employee_ids),
        "fsam_report_summary_rows_ignored": len(fsam_report_summary_rows),
        "output_files": output_files,
        "local_not_in_mohawk": len(local_not_in_mohawk),
    "mohawk_not_in_local": len(mohawk_not_in_local),
    "mohawk_invalid_accounts": len(mohawk_invalid_accounts)
    }

    return summary


def main():
    OUTPUT_DIR.mkdir(exist_ok=True) # If not exist, create output folder. in brackets: don't crash if it exists already

    fam_file = get_first_excel_file(FAM_DIR, "FAM") 
    sam_file = get_first_excel_file(SAM_DIR, "SAM")
    mohawk_file = get_first_excel_file(MOHAWK_DIR, "MOHAWK")
    local_file = get_first_excel_file(LOCAL_DIR, "LOCAL")

    summary = run_verification(fam_file, sam_file, mohawk_file, local_file, OUTPUT_DIR)

    print("Comparison complete.")
    print(f"FSAM rows checked: {summary['fsam_rows_checked']}")
    print(f"LOCAL rows checked: {summary['local_rows_checked']}")
    print(f"LOCAL not in FSAM: {summary['local_not_in_fsam']}")
    print(f"FSAM not in LOCAL: {summary['fsam_not_in_local']}")
    print(f"fsam_invalid_employee_ids: {summary['fsam_invalid_employee_ids']}")
    print(f"local_invalid_employee_ids: {summary['local_invalid_employee_ids']}")
    print(f"fsam_report_summary_rows_ignored: {summary['fsam_report_summary_rows_ignored']}")
    print(f"LOCAL not in MOHAWK accounts: {summary['local_not_in_mohawk']}")
    print(f"MOHAWK accounts not in LOCAL: {summary['mohawk_not_in_local']}")
    print(f"MOHAWK invalid accounts: {summary['mohawk_invalid_accounts']}")

    print("Output files created:")
    for label, file in summary["output_files"].items():
        print(f"{label}: {file}")


if __name__ == "__main__":
    main()
