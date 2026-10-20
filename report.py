import argparse
from pathlib import Path

from retail_reports import process_file


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Clean Online Retail data and create revenue reports."
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        type=Path,
        default=Path(__file__).resolve().parent / "data" / "online_retail_raw.csv",
        help=(
            "Path to a .csv, .csv.gz, or .xlsx input file "
            "(default: data/online_retail_raw.csv)."
        ),
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).resolve().parent / "output",
        help="Folder for charts, workbook, and cancellation review rows.",
    )
    parser.add_argument(
        "--retain-dir",
        type=Path,
        help="Optional folder where cleaned data and summaries are retained.",
    )
    parser.add_argument(
        "--rows",
        type=int,
        help="Process only the first N rows for a small test run.",
    )
    args = parser.parse_args()

    try:
        result = process_file(
            args.input_file,
            args.output_dir,
            rows=args.rows,
            retain_dir=args.retain_dir,
        )
    except (OSError, ValueError) as error:
        parser.error(str(error))

    counts = result["counts"]
    print(f"Rows loaded: {counts['rows_loaded']:,}")
    print(f"Cancellation rows included for netting: {counts['cancellation_rows_included']:,}")
    print(f"Duplicate rows removed: {counts['duplicate_rows_removed']:,}")
    print(
        "Rows missing required non-customer fields removed: "
        f"{counts['rows_with_missing_non_customer_fields_removed']:,}"
    )
    print(
        "Sales rows with zero/negative quantity or price removed: "
        f"{counts['unique_nonpositive_sales_rows_removed']:,}"
    )
    print(
        "Confirmed service/adjustment lines excluded from revenue reports: "
        f"{counts['service_adjustment_rows_excluded_from_revenue_reports']:,}"
    )
    print(f"Rows used in revenue reports: {counts['revenue_rows_in_reports']:,}")
    print(
        f"Gross product sales: £{result['gross_sales_total']:,.2f}; "
        f"cancellations: £{result['cancellations_total']:,.2f}; "
        f"net product revenue: £{result['net_revenue_total']:,.2f}"
    )
    print(f"Reports saved to: {args.output_dir}")
    print(f"Other nonstandard stock codes for review: {result['nonstandard_codes_file']}")
    if result["retained_dataset_dir"] is not None:
        print(
            f"Cleaned data and summaries retained at: "
            f"{result['retained_dataset_dir']}"
        )


if __name__ == "__main__":
    main()
