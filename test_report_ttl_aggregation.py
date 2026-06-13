import unittest

import pandas as pd

from step4_generate_report_preview import (
    excel_average_formula,
    ttl_average_value,
    ttl_child_groups,
)


class ReportTtlAggregationTests(unittest.TestCase):
    def test_category_ttl_includes_single_sku_channels_and_excludes_not_visited(self):
        frame = pd.DataFrame(
            [
                {"Country": "AU", "Category": "W&D", "Channel": "HN", "SKU": "H15 Pro Heat", "JUN": 1.0},
                {"Country": "AU", "Category": "W&D", "Channel": "HN", "SKU": "H16 Pro Steam", "JUN": 0.24},
                {"Country": "AU", "Category": "W&D", "Channel": "HN", "SKU": "TTL", "JUN": 0.62},
                {"Country": "AU", "Category": "W&D", "Channel": "JB", "SKU": "H15 Pro Heat", "JUN": 0.20},
                {"Country": "AU", "Category": "W&D", "Channel": "BSR", "SKU": "H16 Pro Steam", "JUN": "not visited"},
                {"Country": "AU", "Category": "W&D", "Channel": None, "SKU": "TTL", "JUN": ""},
            ]
        )

        groups = ttl_child_groups(frame, 5)

        self.assertEqual(groups, [[2], [3], [4]])
        self.assertAlmostEqual(ttl_average_value(frame, "JUN", groups), 0.41)
        self.assertEqual(excel_average_formula(4, groups), '=IFERROR(AVERAGE(E4,E5,E6),"")')

    def test_channel_ttl_uses_detail_rows_inside_channel(self):
        frame = pd.DataFrame(
            [
                {"Country": "AU", "Category": "W&D", "Channel": "HN", "SKU": "H15 Pro Heat", "JUN": 1.0},
                {"Country": "AU", "Category": "W&D", "Channel": "HN", "SKU": "H16 Pro Steam", "JUN": 0.24},
                {"Country": "AU", "Category": "W&D", "Channel": "HN", "SKU": "TTL", "JUN": ""},
            ]
        )

        groups = ttl_child_groups(frame, 2)

        self.assertEqual(groups, [[0, 1]])
        self.assertAlmostEqual(ttl_average_value(frame, "JUN", groups), 0.62)
        self.assertEqual(excel_average_formula(4, groups), '=IFERROR(AVERAGE(E2:E3),"")')


if __name__ == "__main__":
    unittest.main()
