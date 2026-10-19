import argparse
from pathlib import Path

from retail_reports import process_file


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Clean Online Retail data and create revenue reports."
    )
    parser.add_argument(
        "input_file",
        type=Path,
        help="Path to a .csv, .csv.gz, or .xlsx input file.",
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
        default=Path(__file__).resolve().parent / "processed_data",
        help="Folder where cleaned data and summaries are retained.",
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
    print(f"Cancellation rows removed: {counts['cancellation_rows_removed']:,}")
    print(f"Duplicate rows removed: {counts['duplicate_rows_removed']:,}")
    print(
        "Rows with missing values removed: "
        f"{counts['rows_with_missing_values_removed']:,}"
    )
    print(
        "Unique rows with zero/negative quantity or price removed: "
        f"{counts['unique_nonpositive_rows_removed']:,}"
    )
    print(f"Rows remaining: {counts['rows_remaining']:,}")
    print(f"Reports saved to: {args.output_dir}")
    print(
        f"Cleaned data and summaries retained at: "
        f"{result['retained_dataset_dir']}"
    )


if __name__ == "__main__":
    main()
