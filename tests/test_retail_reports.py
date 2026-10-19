import json
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from retail_reports import process_file


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
            self.assertEqual(counts["cancellation_rows_removed"], 1)
            self.assertEqual(counts["duplicate_rows_removed"], 1)
            self.assertEqual(counts["rows_with_missing_values_removed"], 1)
            self.assertEqual(counts["rows_with_nonpositive_quantity"], 2)
            self.assertEqual(counts["rows_with_nonpositive_price"], 2)
            self.assertEqual(counts["rows_meeting_both_nonpositive_rules"], 0)
            self.assertEqual(counts["unique_nonpositive_rows_removed"], 4)
            self.assertEqual(counts["rows_remaining"], 2)

            retained_dir = Path(result["retained_dataset_dir"])
            self.assertTrue((retained_dir / "cleaned_transactions.csv.gz").is_file())
            self.assertTrue((retained_dir / "revenue_summaries.xlsx").is_file())
            metadata = json.loads(
                (retained_dir / "processing_summary.json").read_text(
                    encoding="utf-8"
                )
            )
            self.assertEqual(metadata["counts"]["rows_remaining"], 2)
            self.assertFalse((retained_dir / "cancellations_review.csv").exists())

            cleaned = pd.read_csv(retained_dir / "cleaned_transactions.csv.gz")
            self.assertEqual(len(cleaned), 2)
            self.assertTrue((cleaned["Revenue"] > 0).all())
            self.assertNotIn("C100003", cleaned["InvoiceNo"].tolist())

            with pd.ExcelFile(result["workbook_file"]) as workbook:
                self.assertEqual(
                    workbook.sheet_names,
                    ["Monthly Revenue", "Top 10 Products", "Top 10 Countries"],
                )
            monthly = pd.read_excel(result["workbook_file"], sheet_name=0)
            self.assertIn("2011-12 (test slice)", monthly["Month"].tolist())

            self.assertTrue(Path(result["cancellations_file"]).is_file())
            self.assertEqual(
                len(pd.read_csv(result["cancellations_file"])),
                1,
            )
            for chart_path in result["charts"].values():
                self.assertTrue(Path(chart_path).is_file())

    def test_missing_required_columns_are_reported(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            input_file = root / "incomplete.csv"
            pd.DataFrame({"InvoiceNo": ["100001"]}).to_csv(input_file, index=False)

            with self.assertRaisesRegex(ValueError, "missing required columns"):
                process_file(input_file, root / "reports")


if __name__ == "__main__":
    unittest.main()
