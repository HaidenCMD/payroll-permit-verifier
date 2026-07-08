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
