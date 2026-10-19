# Online Retail Revenue Analysis

## About the project

This project uses the UCI Online Retail dataset to explore sales performance
and answer practical revenue questions for a retail business. A Python script
cleans the transaction data, calculates revenue, and produces charts and an
Excel workbook that can be reviewed or shared.

The analysis focuses on three questions:

1. How does revenue change from month to month?
2. Which 10 stock-code and description combinations have the highest revenue?
3. Which 10 countries have the highest revenue?

Revenue is calculated in GBP by multiplying each row's quantity by its unit
price. December 2011 is included and marked partial because the dataset ends
on 9 December 2011.

## Project files

| File or folder         | What it contains                                                                                   |
| ---------------------- | -------------------------------------------------------------------------------------------------- |
| `report.py`          | Command-line entry point for running the analysis on a supplied data file.                         |
| `retail_reports.py`  | Shared data validation, cleaning, revenue calculations, summaries, charts, and Excel export logic. |
| `datacleaning.ipynb` | Notebook used to explore and document the data-cleaning and analysis steps.                        |
| `data/`              | Source dataset files used by the project, including the raw CSV and Excel workbooks.               |
| `output/`            | Generated charts, cancellation review CSV, and Excel report workbook.                              |
| `tests/`             | Automated checks for the reporting and cleaning pipeline.                                          |
| `requirements.txt`   | Python packages required to run the report.                                                        |
| `.gitignore`         | Excludes local data, outputs, virtual environments, and generated files from Git.                  |

## Questions answered and results

The generated workbook, `output/online_retail_revenue_analysis.xlsx`, contains
three sheets: Monthly Revenue, Top 10 Products, and Top 10 Countries. The
corresponding charts are saved as PNG files in `output/`.

- **Revenue by month:** November 2011 is the highest full month in the saved
  report, at **£1,156,205.61**. December 2011 is included at **£517,190.44**
  and labelled partial.
- **Top products by revenue:** The highest-ranked stock-code/description row
  is `23843`, PAPER CRAFT, LITTLE BIRDIE, at **£168,469.60**.
- **Top countries by revenue:** The United Kingdom ranks first at
  **£7,285,024.64**.

These figures are from the currently saved workbook and reflect the cleaning
rules described below. The product ranking groups by stock code and
description, but it also includes non-physical transaction codes such as
`POST` (postage) and `M` (manual). Treat the results as revenue by recorded
line code, not exclusively as a ranking of physical merchandise, unless those
codes are separately classified.

## Process and cleaning choices

The report script checks that the input contains the expected columns:
`InvoiceNo`, `StockCode`, `Description`, `Quantity`, `InvoiceDate`,
`UnitPrice`, `CustomerID`, and `Country`. If required columns are missing, it
stops and reports which ones are absent.

For the current analysis, it:

1. Identifies cancellation rows whose invoice number starts with `C`, saves
   them separately for review, and excludes them from the revenue analysis.
2. Removes exact duplicate rows.
3. Removes rows with any missing value, including a missing `CustomerID`.
4. Removes rows with zero or negative quantity or unit price.
5. Calculates row revenue (`Quantity × UnitPrice`), groups dates by
   year-month, and calculates the three requested summaries.
6. Includes the final, incomplete month and labels it as partial.

For the full 541,909-row source file, the saved analysis records 9,288
cancellations, 5,231 duplicate rows, 134,658 rows with at least one missing
value, and 40 rows with zero or negative quantity or price removed. It retains
392,692 rows for the reports. Counts occur in sequence: for example, missing
values are counted after cancellations and duplicates have already been
removed.

Dropping every row with any missing value is a deliberate project rule, not a
claim that a customer with a blank `CustomerID` does not exist. It means
transactions without a recorded customer ID are excluded, which can affect
reported revenue.

## Reproducing the report

With Python and the packages in `requirements.txt` installed, run the script
from the project folder:

```powershell
python report.py data\online_retail_raw.csv
```

The script writes charts, the Excel workbook, and the cancellation review file
to `output/`. A full run also stores a cleaned compressed dataset and summary
information in a uniquely named folder under `processed_data/`. The source
files are read, not modified.

The source data is the [UCI Online Retail dataset](https://archive.ics.uci.edu/dataset/352/online-retail).
