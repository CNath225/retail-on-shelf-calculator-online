import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from step4_generate_report_preview import (
    apply_report_preview_formatting,
    style_report_preview,
    write_ttl_average_formulas,
)


def build_frame() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"Country": "AU", "Category": "W&D", "Channel": "HN", "SKU": "H15 Pro Heat", "MAY": 0.84, "JUN": 1.0, "Trend": "▲"},
            {"Country": "AU", "Category": "W&D", "Channel": "HN", "SKU": "H16 Pro Steam", "MAY": 0.19, "JUN": 0.24, "Trend": "▲"},
            {"Country": "AU", "Category": "W&D", "Channel": "HN", "SKU": "TTL", "MAY": "", "JUN": "", "Trend": "▲"},
            {"Country": "AU", "Category": "W&D", "Channel": "JB", "SKU": "H15 Pro Heat", "MAY": 0.26, "JUN": 0.20, "Trend": "▼"},
            {"Country": "AU", "Category": "W&D", "Channel": "BSR", "SKU": "H16 Pro Steam", "MAY": "not visited", "JUN": "not visited", "Trend": ""},
            {"Country": "AU", "Category": "W&D", "Channel": "", "SKU": "TTL", "MAY": "", "JUN": "", "Trend": "▼"},
        ]
    )


def write_workbook(path: Path, frame: pd.DataFrame) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        frame.to_excel(writer, sheet_name="Report Preview", index=False)
        wb = writer.book
        pct = wb.add_format({"num_format": "0%"})
        green = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        cat_text = wb.add_format({"bg_color": "#0070C0", "font_color": "#FFFFFF", "bold": True})
        cat_pct = wb.add_format({"bg_color": "#0070C0", "font_color": "#FFFFFF", "bold": True, "num_format": "0%"})
        ch_text = wb.add_format({"bg_color": "#A6CAEC", "bold": True})
        ch_pct = wb.add_format({"bg_color": "#A6CAEC", "bold": True, "num_format": "0%"})
        ws = writer.sheets["Report Preview"]
        write_ttl_average_formulas(ws, frame, ["MAY", "JUN"], pct, cat_pct, ch_pct)
        apply_report_preview_formatting(ws, frame, "JUN", "MAY", "Trend", green, red, cat_text, ch_text)


class ReportFormattingTests(unittest.TestCase):
    def test_excel_colours_and_conditional_formats(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "report.xlsx"
            write_workbook(path, build_frame())
            ws = load_workbook(path)["Report Preview"]

            # Category TTL row (position 5 -> excel row 7): strong blue fill, white.
            self.assertEqual(ws.cell(row=7, column=1).fill.fgColor.rgb, "FF0070C0")
            # Channel TTL row (position 2 -> excel row 4): lighter fill.
            self.assertEqual(ws.cell(row=4, column=1).fill.fgColor.rgb, "FFA6CAEC")
            # Detail row (position 0 -> excel row 2): no fill.
            self.assertIsNone(ws.cell(row=2, column=1).fill.patternType)
            # TTL JUN cell carries an AVERAGE formula and the blue fill.
            jun_cell = ws.cell(row=7, column=6)
            self.assertTrue(str(jun_cell.value).startswith("=IFERROR(AVERAGE"))
            self.assertEqual(jun_cell.fill.fgColor.rgb, "FF0070C0")
            # Conditional formatting present (value column + trend column).
            self.assertGreaterEqual(len(list(ws.conditional_formatting)), 2)

    def test_styler_marks_values_and_subtotals(self):
        frame = build_frame()
        rendered = style_report_preview(frame, "JUN").to_html()
        # Category TTL row shaded strong blue; green/red present for detail values.
        self.assertIn("#0070C0", rendered)
        self.assertIn("#C6EFCE", rendered)  # JUN 1.0 -> green
        self.assertIn("#FFC7CE", rendered)  # JUN 0.24 / 0.20 -> red


if __name__ == "__main__":
    unittest.main()
