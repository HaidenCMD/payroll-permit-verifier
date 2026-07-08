from payroll_permit_verifier.paths import OUTPUT_DIR, FAM_DIR, SAM_DIR, MOHAWK_DIR, LOCAL_DIR
from payroll_permit_verifier.excel_io import get_first_excel_file
from payroll_permit_verifier.loaders import load_fsam_reports, load_local_report, load_mohawk_report
from payroll_permit_verifier.comparisons import compare_employee_ids, compare_aims_accounts
from payroll_permit_verifier.reporting import save_reports


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
