import json
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from retail_reports import NON_PRODUCT_CODES, process_file


def sample_rows() -> pd.DataFrame:
    rows = [
        {
            "InvoiceNo": "100001",
            "StockCode": "11111",
            "Description": "Product A",
            "Quantity": 2,
            "InvoiceDate": "2011-11-10 10:00:00",
            "UnitPrice": 3.0,
            "CustomerID": 111,
            "Country": "United Kingdom",
        },
        {
            "InvoiceNo": "100002",
            "StockCode": "22222",
            "Description": "Product B",
            "Quantity": 1,
            "InvoiceDate": "2011-12-01 10:00:00",
            "UnitPrice": 4.0,
            "CustomerID": 222,
            "Country": "France",
        },
    ]
    rows.extend(
        [
            rows[1].copy(),
            {
                "InvoiceNo": "C100003",
                "StockCode": "33333",
                "Description": "Cancelled product",
                "Quantity": -1,
                "InvoiceDate": "2011-12-01 11:00:00",
                "UnitPrice": 5.0,
                "CustomerID": 333,
                "Country": "France",
            },
            {
                "InvoiceNo": "100004",
                "StockCode": "44444",
                "Description": "Missing customer",
                "Quantity": 1,
                "InvoiceDate": "2011-12-01 12:00:00",
                "UnitPrice": 2.0,
                "CustomerID": None,
                "Country": "France",
            },
            {
                "InvoiceNo": "100005",
                "StockCode": "55555",
                "Description": "Zero quantity",
                "Quantity": 0,
                "InvoiceDate": "2011-12-02 10:00:00",
                "UnitPrice": 2.0,
                "CustomerID": 555,
                "Country": "France",
            },
            {
                "InvoiceNo": "100006",
                "StockCode": "66666",
                "Description": "Negative quantity",
                "Quantity": -1,
                "InvoiceDate": "2011-12-02 11:00:00",
                "UnitPrice": 2.0,
                "CustomerID": 666,
                "Country": "France",
            },
            {
                "InvoiceNo": "100007",
                "StockCode": "77777",
                "Description": "Zero price",
                "Quantity": 1,
                "InvoiceDate": "2011-12-03 10:00:00",
                "UnitPrice": 0.0,
                "CustomerID": 777,
                "Country": "France",
            },
            {
                "InvoiceNo": "100008",
                "StockCode": "88888",
                "Description": "Negative price",
                "Quantity": 1,
                "InvoiceDate": "2011-12-03 11:00:00",
                "UnitPrice": -1.0,
                "CustomerID": 888,
                "Country": "France",
            },
        ]
    )
    return pd.DataFrame(rows)


class RetailReportTests(unittest.TestCase):
    def test_processing_does_not_retain_cleaned_data_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            input_file = root / "sample.csv"
            sample_rows().to_csv(input_file, index=False)

            result = process_file(input_file, root / "reports")

            self.assertIsNone(result["retained_dataset_dir"])
            self.assertFalse((root / "processed_data").exists())

    def test_processing_counts_reports_and_retention(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            input_file = root / "sample.csv"
            sample_rows().to_csv(input_file, index=False)

            result = process_file(
                input_file,
                root / "reports",
                rows=9,
                retain_dir=root / "retained",
            )

            counts = result["counts"]
            self.assertEqual(counts["rows_loaded"], 9)
            self.assertEqual(counts["cancellation_rows_included"], 1)
            self.assertEqual(counts["duplicate_rows_removed"], 1)
            self.assertEqual(counts["rows_with_missing_non_customer_fields_removed"], 0)
            self.assertEqual(counts["sales_rows_with_nonpositive_quantity"], 2)
            self.assertEqual(counts["sales_rows_with_nonpositive_price"], 2)
            self.assertEqual(counts["sales_rows_meeting_both_nonpositive_rules"], 0)
            self.assertEqual(counts["unique_nonpositive_sales_rows_removed"], 4)
            self.assertEqual(counts["rows_remaining_after_sales_cleaning"], 4)

            retained_dir = Path(result["retained_dataset_dir"])
            self.assertTrue((retained_dir / "cleaned_transactions.csv.gz").is_file())
            self.assertTrue((retained_dir / "revenue_summaries.xlsx").is_file())
            metadata = json.loads(
                (retained_dir / "processing_summary.json").read_text(
                    encoding="utf-8"
                )
            )
            self.assertEqual(metadata["counts"]["rows_remaining_after_sales_cleaning"], 4)
            self.assertFalse((retained_dir / "cancellations_review.csv").exists())

            cleaned = pd.read_csv(retained_dir / "cleaned_transactions.csv.gz")
            self.assertEqual(len(cleaned), 4)
            self.assertIn("C100003", cleaned["InvoiceNo"].tolist())
            self.assertEqual(cleaned["CustomerID"].isna().sum(), 1)

            with pd.ExcelFile(result["workbook_file"]) as workbook:
                self.assertEqual(
                    workbook.sheet_names,
                    [
                        "Monthly Revenue",
                        "Top 10 Products",
                        "Top 10 Countries ex UK",
                    ],
                )
            monthly = pd.read_excel(result["workbook_file"], sheet_name=0)
            self.assertIn("2011-12 (test slice)", monthly["Month"].tolist())
            self.assertIn("Net Revenue (GBP)", monthly.columns)

            self.assertTrue(Path(result["cancellations_file"]).is_file())
            self.assertEqual(
                len(pd.read_csv(result["cancellations_file"])),
                1,
            )
            for chart_path in result["charts"].values():
                self.assertTrue(Path(chart_path).is_file())
            self.assertTrue(Path(result["nonstandard_codes_file"]).is_file())

    def test_missing_required_columns_are_reported(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            input_file = root / "incomplete.csv"
            pd.DataFrame({"InvoiceNo": ["100001"]}).to_csv(input_file, index=False)

            with self.assertRaisesRegex(ValueError, "missing required columns"):
                process_file(input_file, root / "reports")

    def test_three_rows_net_matching_cancellation_to_zero(self) -> None:
        rows = [
            {
                "InvoiceNo": "500001",
                "StockCode": "23843",
                "Description": "Paper Craft Little Birdie",
                "Quantity": 10,
                "InvoiceDate": "2011-12-09 09:15:00",
                "UnitPrice": 2.08,
                "CustomerID": None,
                "Country": "United Kingdom",
            },
            {
                "InvoiceNo": "C500002",
                "StockCode": "23843",
                "Description": "Paper Craft Little Birdie",
                "Quantity": -10,
                "InvoiceDate": "2011-12-09 09:27:00",
                "UnitPrice": 2.08,
                "CustomerID": None,
                "Country": "United Kingdom",
            },
            {
                "InvoiceNo": "500003",
                "StockCode": "22423",
                "Description": "Regency Cake Stand",
                "Quantity": 1,
                "InvoiceDate": "2011-12-09 10:00:00",
                "UnitPrice": 15.0,
                "CustomerID": 123,
                "Country": "France",
            },
        ]

        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            input_file = root / "three_rows.csv"
            pd.DataFrame(rows).to_csv(input_file, index=False)

            result = process_file(
                input_file,
                root / "reports",
                retain_dir=root / "retained",
            )

            paper_craft = result["top_10_products"].set_index("StockCode").loc["23843"]
            self.assertEqual(paper_craft["Gross Sales (GBP)"], 20.8)
            self.assertEqual(paper_craft["Cancellations (GBP)"], 20.8)
            self.assertEqual(paper_craft["Net Revenue (GBP)"], 0)
            self.assertEqual(result["counts"]["cancellation_rows_included"], 1)
            self.assertEqual(result["counts"]["rows_remaining_after_sales_cleaning"], 3)
            self.assertEqual(result["counts"]["rows_with_missing_non_customer_fields_removed"], 0)

            monthly = result["monthly_revenue"].set_index("YearMonth")
            self.assertAlmostEqual(monthly.loc["2011-12", "Net Revenue (GBP)"], 15.0)
            self.assertAlmostEqual(
                monthly.loc["2011-12", "Cancellations (GBP)"],
                20.8,
            )

            retained = pd.read_csv(
                Path(result["retained_dataset_dir"])
                / "cleaned_transactions.csv.gz"
            )
            self.assertEqual(len(retained), 3)
            self.assertEqual(retained["CustomerID"].isna().sum(), 2)

    def test_approved_non_product_codes_include_case_sensitive_manual_code(self) -> None:
        self.assertIn("B", NON_PRODUCT_CODES)
        self.assertIn("M", NON_PRODUCT_CODES)
        self.assertIn("m", NON_PRODUCT_CODES)


if __name__ == "__main__":
    unittest.main()
