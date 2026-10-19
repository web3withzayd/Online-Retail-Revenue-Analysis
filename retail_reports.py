from __future__ import annotations

import json
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.ticker import FuncFormatter


REQUIRED_COLUMNS = {
    "InvoiceNo",
    "StockCode",
    "Description",
    "Quantity",
    "InvoiceDate",
    "UnitPrice",
    "CustomerID",
    "Country",
}


def _read_input(input_file: Path, rows: int | None) -> pd.DataFrame:
    if not input_file.is_file():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    input_name = input_file.name.lower()
    if input_name.endswith(".csv") or input_name.endswith(".csv.gz"):
        df = pd.read_csv(
            input_file,
            nrows=rows,
            dtype={"InvoiceNo": "string", "StockCode": "string"},
            compression="infer",
        )
    elif input_name.endswith(".xlsx"):
        df = pd.read_excel(
            input_file,
            nrows=rows,
            dtype={"InvoiceNo": "string", "StockCode": "string"},
        )
    else:
        raise ValueError("Input must be a .csv, .csv.gz, or .xlsx file.")

    missing_columns = REQUIRED_COLUMNS - set(df.columns)
    if missing_columns:
        raise ValueError(
            f"Input is missing required columns: {', '.join(sorted(missing_columns))}"
        )

    return df


def _save_retained_copy(
    df: pd.DataFrame,
    workbook_file: Path,
    retain_dir: Path,
    source_name: str,
    counts: dict[str, int],
) -> tuple[str, Path]:
    retain_dir.mkdir(parents=True, exist_ok=True)
    dataset_id = (
        datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        + "-"
        + uuid4().hex[:8]
    )
    dataset_dir = retain_dir / dataset_id

    with tempfile.TemporaryDirectory(prefix=".staging-", dir=retain_dir) as staging:
        staging_dir = Path(staging)
        df.to_csv(
            staging_dir / "cleaned_transactions.csv.gz",
            index=False,
            compression="gzip",
        )
        shutil.copy2(workbook_file, staging_dir / "revenue_summaries.xlsx")
        metadata = {
            "dataset_id": dataset_id,
            "source_filename": Path(source_name.replace("\\", "/")).name,
            "processed_at_utc": datetime.now(timezone.utc).isoformat(),
            "counts": counts,
        }
        (staging_dir / "processing_summary.json").write_text(
            json.dumps(metadata, indent=2),
            encoding="utf-8",
        )
        staging_dir.rename(dataset_dir)

    return dataset_id, dataset_dir


def process_file(
    input_file: str | Path,
    output_dir: str | Path,
    *,
    rows: int | None = None,
    retain_dir: str | Path | None = None,
    source_name: str | None = None,
) -> dict[str, object]:
    """Clean one retail file, calculate revenue summaries, and save the reports."""
    if rows is not None and rows <= 0:
        raise ValueError("rows must be a positive integer.")

    input_path = Path(input_file)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    df = _read_input(input_path, rows)
    rows_loaded = len(df)
    if rows_loaded == 0:
        raise ValueError("The input file contains no data rows.")

    is_cancellation = (
        df["InvoiceNo"]
        .astype("string")
        .str.strip()
        .str.upper()
        .str.startswith("C", na=False)
    )
    cancellations = df.loc[is_cancellation].copy()
    cancellations_file = output_path / "cancellations_review.csv"
    cancellations.to_csv(cancellations_file, index=False)
    df = df.loc[~is_cancellation].copy()

    duplicate_rows = int(df.duplicated().sum())
    df = df.drop_duplicates().copy()

    missing_rows = int(df.isna().any(axis=1).sum())
    df = df.dropna().copy()

    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="raise")
    df["UnitPrice"] = pd.to_numeric(df["UnitPrice"], errors="raise")
    nonpositive_quantity = df["Quantity"] <= 0
    nonpositive_price = df["UnitPrice"] <= 0
    rows_to_remove = nonpositive_quantity | nonpositive_price
    nonpositive_rows = int(rows_to_remove.sum())

    counts = {
        "rows_loaded": rows_loaded,
        "cancellation_rows_removed": len(cancellations),
        "duplicate_rows_removed": duplicate_rows,
        "rows_with_missing_values_removed": missing_rows,
        "rows_with_nonpositive_quantity": int(nonpositive_quantity.sum()),
        "rows_with_nonpositive_price": int(nonpositive_price.sum()),
        "rows_meeting_both_nonpositive_rules": int(
            (nonpositive_quantity & nonpositive_price).sum()
        ),
        "unique_nonpositive_rows_removed": nonpositive_rows,
    }
    df = df.loc[~rows_to_remove].copy()
    counts["rows_remaining"] = len(df)

    if df.empty:
        raise ValueError("No rows remain after applying the cleaning rules.")

    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"], errors="raise")
    df["Revenue"] = df["Quantity"] * df["UnitPrice"]
    df["YearMonth"] = df["InvoiceDate"].dt.strftime("%Y-%m")

    last_month = df["InvoiceDate"].max().strftime("%Y-%m")
    monthly_revenue = (
        df.groupby("YearMonth", as_index=False)["Revenue"]
        .sum()
        .sort_values("YearMonth")
    )
    monthly_revenue["Month"] = monthly_revenue["YearMonth"]
    final_month_note = " (test slice)" if rows is not None else " (partial)"
    monthly_revenue.loc[monthly_revenue["YearMonth"] == last_month, "Month"] += (
        final_month_note
    )

    top_10_products = (
        df.groupby(["StockCode", "Description"], as_index=False, dropna=False)[
            "Revenue"
        ]
        .sum()
        .sort_values("Revenue", ascending=False)
        .head(10)
    )
    top_10_countries = (
        df.groupby("Country", as_index=False)["Revenue"]
        .sum()
        .sort_values("Revenue", ascending=False)
        .head(10)
    )

    monthly_excel = monthly_revenue[["Month", "Revenue"]].rename(
        columns={"Revenue": "Revenue (GBP)"}
    )
    products_excel = top_10_products[
        ["StockCode", "Description", "Revenue"]
    ].rename(columns={"Revenue": "Revenue (GBP)"})
    countries_excel = top_10_countries[["Country", "Revenue"]].rename(
        columns={"Revenue": "Revenue (GBP)"}
    )

    workbook_file = output_path / "online_retail_revenue_analysis.xlsx"
    with pd.ExcelWriter(workbook_file) as writer:
        monthly_excel.to_excel(writer, sheet_name="Monthly Revenue", index=False)
        products_excel.to_excel(writer, sheet_name="Top 10 Products", index=False)
        countries_excel.to_excel(writer, sheet_name="Top 10 Countries", index=False)

    gbp_format = FuncFormatter(lambda value, position: f"£{value:,.0f}")
    charts: dict[str, Path] = {}

    fig, ax = plt.subplots(figsize=(12, 6))
    ax.bar(monthly_revenue["Month"], monthly_revenue["Revenue"], color="steelblue")
    ax.yaxis.set_major_formatter(gbp_format)
    ax.set_title("Revenue per Month (GBP)")
    ax.set_xlabel("Month")
    ax.set_ylabel("Revenue (GBP)")
    plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
    fig.tight_layout()
    charts["monthly_revenue"] = output_path / "revenue_per_month.png"
    fig.savefig(charts["monthly_revenue"], dpi=300)
    plt.close(fig)

    products_plot = top_10_products.copy()
    products_plot["Product"] = (
        products_plot["StockCode"].astype("string")
        + " - "
        + products_plot["Description"].astype("string")
    )
    products_plot = products_plot.sort_values("Revenue")

    fig, ax = plt.subplots(figsize=(12, 7))
    ax.barh(products_plot["Product"], products_plot["Revenue"], color="seagreen")
    ax.xaxis.set_major_formatter(gbp_format)
    ax.set_title("Top 10 Products by Revenue (GBP)")
    ax.set_xlabel("Revenue (GBP)")
    ax.set_ylabel("Product")
    fig.tight_layout()
    charts["top_10_products"] = output_path / "top_10_products_by_revenue.png"
    fig.savefig(charts["top_10_products"], dpi=300)
    plt.close(fig)

    countries_plot = top_10_countries.sort_values("Revenue")
    fig, ax = plt.subplots(figsize=(10, 7))
    ax.barh(countries_plot["Country"], countries_plot["Revenue"], color="darkorange")
    ax.xaxis.set_major_formatter(gbp_format)
    ax.set_title("Top 10 Countries by Revenue (GBP)")
    ax.set_xlabel("Revenue (GBP)")
    ax.set_ylabel("Country")
    fig.tight_layout()
    charts["top_10_countries"] = output_path / "top_10_countries_by_revenue.png"
    fig.savefig(charts["top_10_countries"], dpi=300)
    plt.close(fig)

    dataset_id = None
    retained_dataset_dir = None
    if retain_dir is not None:
        dataset_id, retained_dataset_dir = _save_retained_copy(
            df,
            workbook_file,
            Path(retain_dir),
            source_name or input_path.name,
            counts,
        )

    return {
        "counts": counts,
        "cancellations_file": cancellations_file,
        "workbook_file": workbook_file,
        "charts": charts,
        "monthly_revenue": monthly_revenue,
        "top_10_products": top_10_products,
        "top_10_countries": top_10_countries,
        "dataset_id": dataset_id,
        "retained_dataset_dir": retained_dataset_dir,
    }
