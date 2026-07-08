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
