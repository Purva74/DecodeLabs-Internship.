from pathlib import Path
import pandas as pd

# ------------------------------------------------------------
# DATA ANALYTICS PROJECT 1: DATA CLEANING & PREPARATION
# ------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "Dataset for Data Analytics.xlsx"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

CLEANED_FILE = OUTPUT_DIR / "Cleaned_Dataset.xlsx"
INVALID_DATES_FILE = OUTPUT_DIR / "Rows_With_Invalid_Dates.xlsx"
REPORT_FILE = OUTPUT_DIR / "Cleaning_Report.txt"


def main():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Dataset not found: {INPUT_FILE.name}. "
            "Place the Excel dataset in the project folder."
        )

    # 1. Load data
    df = pd.read_excel(INPUT_FILE)
    original_rows = len(df)
    original_columns = len(df.columns)
    missing_before = df.isna().sum()

    print("=" * 60)
    print("DATA ANALYTICS PROJECT 1 - DATA CLEANING & PREPARATION")
    print("=" * 60)
    print(f"Input rows: {original_rows}")
    print(f"Input columns: {original_columns}")
    print("\nMissing values before cleaning:")
    print(missing_before.to_string())

    # 2. Standardize column names and text whitespace
    df.columns = [str(col).strip() for col in df.columns]
    text_columns = df.select_dtypes(include=["object", "string"]).columns
    for col in text_columns:
        df[col] = df[col].apply(
            lambda value: value.strip() if isinstance(value, str) else value
        )

    # 3. Handle CouponCode blanks as specified for this dataset
    if "CouponCode" in df.columns:
        df["CouponCode"] = df["CouponCode"].fillna("No Coupon")
        # Also treat whitespace-only coupon values as missing.
        df["CouponCode"] = df["CouponCode"].replace(
            r"^\s*$", "No Coupon", regex=True
        )

    # 4. Remove completely duplicated rows
    duplicate_rows_removed = int(df.duplicated().sum())
    df = df.drop_duplicates().copy()

    # 5. Validate dates. Invalid/missing dates are written to a separate file
    # and excluded from the final clean dataset so all exported dates are valid.
    invalid_date_rows = pd.DataFrame()
    invalid_date_count = 0
    if "Date" in df.columns:
        original_date_values = df["Date"].copy()
        parsed_dates = pd.to_datetime(original_date_values, errors="coerce")
        invalid_mask = parsed_dates.isna()
        invalid_date_count = int(invalid_mask.sum())
        if invalid_date_count:
            invalid_date_rows = df.loc[invalid_mask].copy()
            invalid_date_rows.to_excel(INVALID_DATES_FILE, index=False)
        df["Date"] = parsed_dates
        df = df.loc[~invalid_mask].copy()
        df["Date"] = df["Date"].dt.normalize()

    # 6. Ensure OrderID values are present and unique.
    # Missing OrderIDs cannot be used as reliable identifiers, so they are
    # saved to the review file and excluded from the clean dataset.
    missing_id_count = 0
    duplicate_ids_removed = 0
    if "OrderID" in df.columns:
        missing_id_mask = df["OrderID"].isna() | (
            df["OrderID"].astype("string").str.strip() == ""
        )
        missing_id_count = int(missing_id_mask.sum())
        if missing_id_count:
            missing_ids_file = OUTPUT_DIR / "Rows_With_Missing_OrderID.xlsx"
            df.loc[missing_id_mask].to_excel(missing_ids_file, index=False)
            df = df.loc[~missing_id_mask].copy()

        duplicate_id_mask = df["OrderID"].duplicated(keep="first")
        duplicate_ids_removed = int(duplicate_id_mask.sum())
        df = df.loc[~duplicate_id_mask].copy()

    # 7. Convert expected numeric fields to numeric values.
    numeric_columns = [
        "Quantity", "UnitPrice", "ItemsInCart", "TotalPrice"
    ]
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 8. Save final clean dataset
    df.to_excel(CLEANED_FILE, index=False)

    # 9. Verify quality rules
    duplicate_ids_remaining = (
        int(df["OrderID"].duplicated().sum()) if "OrderID" in df.columns else 0
    )
    invalid_dates_remaining = (
        int(pd.to_datetime(df["Date"], errors="coerce").isna().sum())
        if "Date" in df.columns else 0
    )
    missing_after = df.isna().sum()

    report_lines = [
        "DATA ANALYTICS PROJECT 1",
        "Data Cleaning & Preparation",
        "=" * 40,
        f"Original rows: {original_rows}",
        f"Original columns: {original_columns}",
        f"Final clean rows: {len(df)}",
        f"Completely duplicate rows removed: {duplicate_rows_removed}",
        f"Duplicate OrderIDs removed: {duplicate_ids_removed}",
        f"Rows with invalid/missing dates moved out of clean dataset: {invalid_date_count}",
        f"Rows with missing OrderID moved out of clean dataset: {missing_id_count}",
        f"Duplicate OrderIDs remaining: {duplicate_ids_remaining}",
        f"Invalid/missing dates remaining: {invalid_dates_remaining}",
        "",
        "Missing values before cleaning:",
        missing_before.to_string(),
        "",
        "Missing values in final clean dataset:",
        missing_after.to_string(),
        "",
        "Verification:",
        f"Zero duplicate OrderIDs: {'PASS' if duplicate_ids_remaining == 0 else 'FAIL'}",
        f"Zero invalid/missing dates: {'PASS' if invalid_dates_remaining == 0 else 'FAIL'}",
        "",
        "Note: Other missing values are reported for review rather than filled with guessed values.",
        "CouponCode blanks are filled with 'No Coupon'.",
    ]
    REPORT_FILE.write_text("\n".join(report_lines), encoding="utf-8")

    print("\nCleaning completed.")
    print(f"Cleaned file: {CLEANED_FILE}")
    print(f"Report file: {REPORT_FILE}")
    print(f"Duplicate OrderIDs remaining: {duplicate_ids_remaining}")
    print(f"Invalid/missing dates remaining: {invalid_dates_remaining}")
    print("\nVerification:")
    print("Zero duplicate OrderIDs:", "PASS" if duplicate_ids_remaining == 0 else "FAIL")
    print("Zero invalid/missing dates:", "PASS" if invalid_dates_remaining == 0 else "FAIL")


if __name__ == "__main__":
    main()
