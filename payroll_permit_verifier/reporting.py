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
