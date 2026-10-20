# Online Retail Revenue Analysis

## About the project

This project uses the UCI Online Retail dataset to explore sales performance
and answer practical revenue questions for a retail business. The Python
reporting script validates and cleans transaction data, calculates net revenue,
and creates charts and an Excel workbook.

The analysis answers three questions:

1. How does monthly net revenue change over time?
2. Which 10 products have the highest net revenue?
3. Which 10 countries have the highest net revenue, excluding the UK so the
   other countries are easier to compare?

All three charts use the same net-revenue definition for confirmed product
codes: gross sales minus cancellations. Cancellations are assigned to the
month and country recorded on the cancellation row, i.e. when they were
processed. Gross sales and cancellation amounts are also shown separately in
the Excel summaries and chart footnotes. Amounts are in GBP.

The final month is included and labelled partial when the latest transaction
date falls before that month's end.

## Project files

| File or folder | What it contains |
| --- | --- |
| `report.py` | Command-line entry point for running the report on an input file. |
| `retail_reports.py` | Data validation, cleaning, cancellation netting, summaries, chart generation, and Excel export. |
| `datacleaning.ipynb` | Notebook used to explore and document the data-cleaning and analysis steps. |
| `data/` | Source dataset files used by the project, including the raw CSV and Excel workbooks. |
| `output/` | Generated charts, the Excel report, and review files from the latest saved run. |
| `tests/` | Automated tests, including a three-row sale/cancellation netting example. |
| `requirements.txt` | Python packages required to run the report. |
| `.gitignore` | Excludes local data, outputs, virtual environments, and generated files from Git. |

## Questions answered and results

Each report run creates an Excel workbook with sheets for Monthly Revenue,
Top 10 Products, and Top 10 Countries ex UK. Each summary shows gross sales,
cancellations, and net revenue. The PNG charts visualize net revenue.

In the current full-data run, using the source file with the cleaning rules
below:

- **Monthly revenue:** November 2011 is the highest full month at
  **£1,427,124.54 net**. December 2011 is marked partial and shows
  **£440,407.89 net**.
- **Top product:** StockCode `22423`, REGENCY CAKESTAND 3 TIER, leads at
  **£164,459.49 net**.
- **Top country outside the UK:** The Netherlands leads at
  **£283,479.54 net**.
- **Overall:** Gross product sales are **£10,247,905.13**, cancellations are
  **£475,901.16**, and net product revenue is **£9,772,003.97**.
- The UK accounts for **84.8% of net product-code revenue** in this run; the
  percentage is calculated from the data and appears in the country chart
  title.

These results reflect the cleaning rules below and the currently confirmed
non-product exclusions. In this run, 21 other stock codes did not fit the
expected five-digit-plus-optional-letters pattern; they remain included and
are listed for review, not excluded automatically.

## Process and cleaning choices

The script checks for `InvoiceNo`, `StockCode`, `Description`, `Quantity`,
`InvoiceDate`, `UnitPrice`, `CustomerID`, and `Country`. It stops with an
explicit error if a required column is missing.

For the current report, it:

1. Identifies invoice numbers beginning with `C` as cancellations and keeps
   those lines in the analysis with their negative revenue. It also writes a
   separate cancellation review CSV.
2. Removes exact duplicate rows from both sales and cancellation lines.
3. Removes rows missing a required transaction field, but keeps rows with a
   missing `CustomerID`. Customer IDs are not needed for these revenue
   summaries; they can be filtered for a later customer-specific analysis.
4. Removes sales rows with zero or negative quantity or unit price. Negative
   cancellation quantities are retained so the cancellation can reduce net
   revenue.
5. Excludes these confirmed service/adjustment codes from all three revenue
   charts: `AMAZONFEE`, `B`, `BANK CHARGES`, `C2`, `CRUK`, `D`, `DOT`, `M`,
   `m`, `POST`, and `S`. StockCode matching is case-sensitive.
6. Builds a review file for other source codes outside the expected product
   code pattern. It includes descriptions, source line counts, and signed
   value totals. These codes remain in the reports until individually
   reviewed and approved for exclusion.
7. Calculates line revenue as `Quantity × UnitPrice`, then groups by month,
   stock code, or country. A cancellation reduces the product's net revenue
   by stock code and is counted in the cancellation row's processing month
   and country.

For the full 541,909-row source file, the current run removed 5,268 exact
duplicates, 1,454 rows missing non-customer transaction fields, and 1,058
non-cancellation sales rows with non-positive quantity or price. It retained
9,251 cancellation rows for netting and excluded 2,890 rows with the confirmed
service/adjustment codes from the revenue charts. The three revenue summaries
use 531,239 rows before grouping.

The cancellation of `23843` PAPER CRAFT, LITTLE BIRDIE is a useful example:
the 9 December 2011 sale is £168,469.60 and the cancellation is -£168,469.60,
so the code's net revenue is £0.00. December's chart includes the cancellation
in December because that is when it was processed.

## Reproducing the report

With Python and the packages in `requirements.txt` installed, run this from
the project folder:

```powershell
python report.py
```

By default, the script reads `data\online_retail_raw.csv` and saves the charts,
Excel workbook, cancellation review CSV, and nonstandard-code review CSV to
`output/`. It does not create a separate `processed_data/` folder or retain a
cleaned copy unless you explicitly pass `--retain-dir`. The source file is
never modified. To analyze a different file, provide its path after the
command, for example:

```powershell
python report.py path\to\another_file.xlsx
```

To also retain the cleaned transactions and run summaries, explicitly choose a
separate folder:

```powershell
python report.py --retain-dir processed_data
```

The source data is the
[UCI Online Retail dataset](https://archive.ics.uci.edu/dataset/352/online-retail).
