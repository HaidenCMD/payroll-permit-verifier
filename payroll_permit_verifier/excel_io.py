import pandas as pd
from io import BytesIO
import getpass
import msoffcrypto

from payroll_permit_verifier.settings import REPORT_HEADER_ROW


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
