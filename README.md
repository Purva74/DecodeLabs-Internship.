# Data Analytics Project 1 — Data Cleaning & Preparation

## Objective
Clean the provided Excel dataset by checking missing values, removing duplicate rows,
ensuring unique OrderID values, validating date values, standardizing text whitespace,
and converting selected columns to numeric data types.

## Files
- `main.py` — Python cleaning script
- `Dataset for Data Analytics.xlsx` — input dataset
- `requirements.txt` — required Python libraries
- `output/` — generated results

## Requirements
- Python 3.10 or later
- Visual Studio Code
- Python extension for VS Code (recommended)

## Run in VS Code
1. Extract `Data_Analytics_Project_1.zip`.
2. Open VS Code.
3. Select **File → Open Folder** and choose `Data_Analytics_Project_1`.
4. Open **Terminal → New Terminal**.
5. Install dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

6. Run the project:

   ```bash
   python main.py
   ```

   If `python` does not work on Windows, try:

   ```bash
   py main.py
   ```

## Output files
- `output/Cleaned_Dataset.xlsx` — final cleaned dataset
- `output/Cleaning_Report.txt` — cleaning summary and validation checks
- `output/Rows_With_Invalid_Dates.xlsx` — rows that had invalid or missing dates, if any
- `output/Rows_With_Missing_OrderID.xlsx` — rows that had missing OrderID values, if any

## Cleaning decisions
- Blank `CouponCode` values are replaced with `No Coupon`.
- Fully duplicated rows are removed.
- Invalid or missing dates are excluded from the final clean dataset and saved separately for review.
- Missing OrderID rows are excluded from the final clean dataset and saved separately for review.
- Duplicate OrderID values are reduced to one record (the first occurrence is retained).
- Numeric fields `Quantity`, `UnitPrice`, `ItemsInCart`, and `TotalPrice` are converted to numeric values.
- Other missing values are reported, not guessed or automatically filled.

## Verification
The report shows whether the final output has:
- zero duplicate OrderID values
- zero invalid or missing dates

Review the report after running the script. Keep the original input dataset unchanged.
