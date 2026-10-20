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
REQUIRED_TRANSACTION_FIELDS = REQUIRED_COLUMNS - {"CustomerID"}
NON_PRODUCT_CODES = frozenset(
    {
        "AMAZONFEE",
        "B",
        "BANK CHARGES",
        "C2",
        "CRUK",
        "D",
        "DOT",
        "M",
        "m",
        "POST",
        "S",
    }
)
PRODUCT_CODE_PATTERN = r"\d{5}[A-Za-z]*"


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


def _revenue_components(
    df: pd.DataFrame,
    group_by: str | list[str],
) -> pd.DataFrame:
    rows = df.assign(
        **{
            "Gross Sales (GBP)": df["Revenue"].where(~df["IsCancellation"], 0),
            "Cancellations (GBP)": (-df["Revenue"]).where(df["IsCancellation"], 0),
        }
    )
    summary = (
        rows.groupby(group_by, as_index=False, dropna=False)[
            ["Gross Sales (GBP)", "Cancellations (GBP)"]
        ]
        .sum()
    )
    summary["Net Revenue (GBP)"] = (
        summary["Gross Sales (GBP)"] - summary["Cancellations (GBP)"]
    )
    return summary


def _nonstandard_code_review(df: pd.DataFrame) -> pd.DataFrame:
    has_code = df["StockCode"].notna()
    nonstandard = ~df["StockCode"].str.fullmatch(PRODUCT_CODE_PATTERN, na=False)
    not_already_confirmed = ~df["StockCode"].isin(NON_PRODUCT_CODES)
    review_rows = df.loc[has_code & nonstandard & not_already_confirmed].copy()
    if review_rows.empty:
        return pd.DataFrame(
            columns=[
                "StockCode",
                "Description",
                "Source Line Count",
                "Non-cancellation Value (GBP)",
                "Cancellations (GBP)",
                "Net Signed Value (GBP)",
            ]
        )

    review_rows["Non-cancellation Value (GBP)"] = review_rows["Revenue"].where(
        ~review_rows["IsCancellation"], 0
    )
    review_rows["Cancellations (GBP)"] = (-review_rows["Revenue"]).where(
        review_rows["IsCancellation"], 0
    )
    review_rows["Net Signed Value (GBP)"] = review_rows["Revenue"]

    def combine_descriptions(values: pd.Series) -> str:
        descriptions = sorted(values.dropna().astype(str).unique())
        return " | ".join(descriptions) if descriptions else "[missing description]"

    review = (
        review_rows.groupby("StockCode", as_index=False, dropna=False)
        .agg(
            Description=("Description", combine_descriptions),
            **{
                "Source Line Count": ("StockCode", "size"),
                "Non-cancellation Value (GBP)": (
                    "Non-cancellation Value (GBP)",
                    "sum",
                ),
                "Cancellations (GBP)": ("Cancellations (GBP)", "sum"),
                "Net Signed Value (GBP)": ("Net Signed Value (GBP)", "sum"),
            },
        )
        .sort_values("StockCode", key=lambda column: column.str.upper())
    )
    return review


def _save_retained_copy(
    df: pd.DataFrame,
    workbook_file: Path,
    retain_dir: Path,
    source_name: str,
    counts: dict[str, int | float],
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
    """Clean one retail file and create net-revenue reports."""
    if rows is not None and rows <= 0:
        raise ValueError("rows must be a positive integer.")

    input_path = Path(input_file)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    df = _read_input(input_path, rows)
    rows_loaded = len(df)
    if rows_loaded == 0:
        raise ValueError("The input file contains no data rows.")

    for column in ("InvoiceNo", "StockCode"):
        df[column] = df[column].astype("string").str.strip()
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="raise")
    df["UnitPrice"] = pd.to_numeric(df["UnitPrice"], errors="raise")
    df["IsCancellation"] = (
        df["InvoiceNo"].str.upper().str.startswith("C", na=False)
    )
    df["Revenue"] = df["Quantity"] * df["UnitPrice"]

    nonstandard_codes = _nonstandard_code_review(df)

    duplicate_rows = int(df.duplicated().sum())
    df = df.drop_duplicates().copy()

    missing_value_rows = int(df[list(REQUIRED_TRANSACTION_FIELDS)].isna().any(axis=1).sum())
    df = df.dropna(subset=list(REQUIRED_TRANSACTION_FIELDS)).copy()

    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"], errors="raise")
    df["YearMonth"] = df["InvoiceDate"].dt.strftime("%Y-%m")

    nonpositive_quantity = (~df["IsCancellation"]) & (df["Quantity"] <= 0)
    nonpositive_price = (~df["IsCancellation"]) & (df["UnitPrice"] <= 0)
    invalid_sales = nonpositive_quantity | nonpositive_price
    nonpositive_rows = int(invalid_sales.sum())
    df = df.loc[~invalid_sales].copy()

    cancellation_rows = df["IsCancellation"]
    cancellations = df.loc[cancellation_rows].copy()
    cancellations_file = output_path / "cancellations_review.csv"
    cancellations.to_csv(cancellations_file, index=False)

    excluded_service_rows = df["StockCode"].isin(NON_PRODUCT_CODES)
    product_rows = df.loc[~excluded_service_rows].copy()
    if product_rows.empty:
        raise ValueError("No product-code rows remain for revenue analysis.")

    nonstandard_codes_file = output_path / "nonstandard_stock_codes_review.csv"
    nonstandard_codes.to_csv(nonstandard_codes_file, index=False)

    latest_invoice_date = product_rows["InvoiceDate"].max()
    last_month = latest_invoice_date.strftime("%Y-%m")
    final_month_is_partial = latest_invoice_date.date() < (
        latest_invoice_date + pd.offsets.MonthEnd(0)
    ).date()

    monthly_revenue = _revenue_components(product_rows, "YearMonth").sort_values(
        "YearMonth"
    )
    monthly_revenue["Month"] = monthly_revenue["YearMonth"]
    partial_label = " (partial)" if final_month_is_partial else ""
    test_label = " (test slice)" if rows is not None else partial_label
    monthly_revenue.loc[
        monthly_revenue["YearMonth"] == last_month, "Month"
    ] += test_label

    top_10_products = _revenue_components(product_rows, "StockCode")
    descriptions = (
        product_rows.loc[~product_rows["IsCancellation"]]
        .groupby("StockCode")["Description"]
        .agg(lambda values: values.mode().iloc[0])
        .rename("Description")
    )
    top_10_products = (
        top_10_products.join(descriptions, on="StockCode")
        .sort_values("Net Revenue (GBP)", ascending=False)
        .head(10)
        [["StockCode", "Description", "Gross Sales (GBP)", "Cancellations (GBP)", "Net Revenue (GBP)"]]
    )

    country_revenue = _revenue_components(product_rows, "Country")
    net_total = float(country_revenue["Net Revenue (GBP)"].sum())
    uk_net_revenue = float(
        country_revenue.loc[
            country_revenue["Country"].str.casefold() == "united kingdom",
            "Net Revenue (GBP)",
        ].sum()
    )
    uk_net_share = uk_net_revenue / net_total if net_total else 0.0
    top_10_countries = (
        country_revenue.loc[
            country_revenue["Country"].str.casefold() != "united kingdom"
        ]
        .sort_values("Net Revenue (GBP)", ascending=False)
        .head(10)
    )

    gross_sales_total = float(product_rows.loc[~product_rows["IsCancellation"], "Revenue"].sum())
    cancellations_total = float(-product_rows.loc[product_rows["IsCancellation"], "Revenue"].sum())
    net_revenue_total = gross_sales_total - cancellations_total
    counts: dict[str, int | float] = {
        "rows_loaded": rows_loaded,
        "duplicate_rows_removed": duplicate_rows,
        "rows_with_missing_non_customer_fields_removed": missing_value_rows,
        "cancellation_rows_included": len(cancellations),
        "sales_rows_with_nonpositive_quantity": int(nonpositive_quantity.sum()),
        "sales_rows_with_nonpositive_price": int(nonpositive_price.sum()),
        "sales_rows_meeting_both_nonpositive_rules": int(
            (nonpositive_quantity & nonpositive_price).sum()
        ),
        "unique_nonpositive_sales_rows_removed": nonpositive_rows,
        "service_adjustment_rows_excluded_from_revenue_reports": int(
            excluded_service_rows.sum()
        ),
        "revenue_rows_in_reports": len(product_rows),
        "rows_remaining_after_sales_cleaning": len(df),
        "gross_sales_gbp_product_codes": gross_sales_total,
        "cancellations_gbp_product_codes": cancellations_total,
        "net_revenue_gbp_product_codes": net_revenue_total,
        "uk_share_of_net_revenue": uk_net_share,
    }

    workbook_file = output_path / "online_retail_revenue_analysis.xlsx"
    with pd.ExcelWriter(workbook_file) as writer:
        monthly_revenue[
            ["Month", "Gross Sales (GBP)", "Cancellations (GBP)", "Net Revenue (GBP)"]
        ].to_excel(writer, sheet_name="Monthly Revenue", index=False)
        top_10_products.to_excel(writer, sheet_name="Top 10 Products", index=False)
        top_10_countries.to_excel(writer, sheet_name="Top 10 Countries ex UK", index=False)

    gbp_format = FuncFormatter(lambda value, position: f"£{value:,.0f}")
    chart_summary = (
        f"Gross product sales £{gross_sales_total:,.2f} | "
        f"Cancellations £{cancellations_total:,.2f} | "
        f"Net £{net_revenue_total:,.2f}"
    )
    charts: dict[str, Path] = {}

    fig, ax = plt.subplots(figsize=(12, 6))
    ax.bar(monthly_revenue["Month"], monthly_revenue["Net Revenue (GBP)"], color="steelblue")
    ax.yaxis.set_major_formatter(gbp_format)
    partial_title = (
        f" — final month {last_month} is partial"
        if final_month_is_partial and rows is None
        else ""
    )
    test_title = " — test slice" if rows is not None else ""
    ax.set_title(
        f"Monthly Net Product Revenue (GBP){partial_title}{test_title}\n"
        "Cancellations grouped by processing month"
    )
    ax.set_xlabel("Month")
    ax.set_ylabel("Net Revenue (GBP)")
    fig.text(0.5, 0.02, chart_summary, ha="center")
    plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
    fig.tight_layout(rect=(0, 0.06, 1, 1))
    charts["monthly_revenue"] = output_path / "revenue_per_month.png"
    fig.savefig(charts["monthly_revenue"], dpi=300)
    plt.close(fig)

    products_plot = top_10_products.sort_values("Net Revenue (GBP)")
    products_plot = products_plot.copy()
    products_plot["Product"] = (
        products_plot["StockCode"].astype("string")
        + " - "
        + products_plot["Description"]
        .fillna("Unknown description")
        .astype("string")
    )
    fig, ax = plt.subplots(figsize=(12, 7))
    ax.barh(
        products_plot["Product"],
        products_plot["Net Revenue (GBP)"],
        color="seagreen",
    )
    ax.xaxis.set_major_formatter(gbp_format)
    ax.set_title("Top 10 Products by Net Revenue (GBP)")
    ax.set_xlabel("Net Revenue (GBP)")
    ax.set_ylabel("Product")
    fig.text(0.5, 0.02, chart_summary, ha="center")
    fig.tight_layout(rect=(0, 0.06, 1, 1))
    charts["top_10_products"] = output_path / "top_10_products_by_revenue.png"
    fig.savefig(charts["top_10_products"], dpi=300)
    plt.close(fig)

    countries_plot = top_10_countries.sort_values("Net Revenue (GBP)")
    fig, ax = plt.subplots(figsize=(10, 7))
    ax.barh(
        countries_plot["Country"],
        countries_plot["Net Revenue (GBP)"],
        color="darkorange",
    )
    ax.xaxis.set_major_formatter(gbp_format)
    ax.set_title(
        "Top 10 Countries by Net Product Revenue (GBP), excluding UK\n"
        f"UK share of net product revenue: {uk_net_share:.1%}"
    )
    ax.set_xlabel("Net Revenue (GBP)")
    ax.set_ylabel("Country")
    fig.text(0.5, 0.02, chart_summary, ha="center")
    fig.tight_layout(rect=(0, 0.06, 1, 1))
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
        "nonstandard_codes_file": nonstandard_codes_file,
        "nonstandard_codes_review": nonstandard_codes,
        "workbook_file": workbook_file,
        "charts": charts,
        "monthly_revenue": monthly_revenue,
        "top_10_products": top_10_products,
        "top_10_countries": top_10_countries,
        "gross_sales_total": gross_sales_total,
        "cancellations_total": cancellations_total,
        "net_revenue_total": net_revenue_total,
        "uk_net_share": uk_net_share,
        "dataset_id": dataset_id,
        "retained_dataset_dir": retained_dataset_dir,
    }
