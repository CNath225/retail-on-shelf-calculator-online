import unittest

from identifier_matching import (
    CanonicalEntry,
    build_resolution_items,
    normalize_identifier,
)


class IdentifierMatchingTests(unittest.TestCase):
    def test_sku_channel_suffix_variants_auto_match(self):
        canonicals = [CanonicalEntry(domain="sku", canonical="Aqua10 Roller SE")]
        raws = ["Aqua10 Roller SE (JB)", "Aqua10 Roller SE (HN)", "aqua10  roller se"]

        items = build_resolution_items("sku", raws, canonicals)

        self.assertEqual({item.canonical for item in items}, {"Aqua10 Roller SE"})
        self.assertTrue(all(item.status == "auto_normalized" for item in items))

    def test_near_sku_is_confusable_not_merged(self):
        canonicals = [CanonicalEntry(domain="sku", canonical="Aqua10 Roller SE")]
        items = build_resolution_items("sku", ["Aqua10 Roller S"], canonicals)

        self.assertEqual(items[0].status, "confusable")
        self.assertEqual(items[0].canonical, "")
        self.assertEqual(items[0].confusable_with, ("Aqua10 Roller SE",))

    def test_near_sku_does_not_match_unrelated_suffix_sku(self):
        canonicals = [CanonicalEntry(domain="sku", canonical="Aqua10 Ultra Track S")]
        items = build_resolution_items("sku", ["Aqua10 Roller S"], canonicals)

        self.assertEqual(items[0].status, "unmatched")
        self.assertEqual(items[0].canonical, "")

    def test_old_model_bracket_is_preserved(self):
        old_variant = normalize_identifier("Aqua10 Roller SE (含老款)")
        new_variant = normalize_identifier("Aqua10 Roller SE")

        self.assertNotEqual(old_variant.key, new_variant.key)
        self.assertIn("含老款", old_variant.raw)

        canonicals = [CanonicalEntry(domain="sku", canonical="Aqua10 Roller SE")]
        items = build_resolution_items("sku", ["Aqua10 Roller SE (含老款)"], canonicals)
        self.assertNotEqual(items[0].status, "auto_normalized")

    def test_account_category_country_use_same_rules(self):
        cases = [
            ("account", "Harvey Norman (HN)", "Harvey Norman"),
            ("category", "Robot", "robot"),
            ("country", "AU", "au"),
        ]
        for domain, raw, canonical in cases:
            with self.subTest(domain=domain):
                items = build_resolution_items(
                    domain,
                    [raw],
                    [CanonicalEntry(domain=domain, canonical=canonical)],
                )
                self.assertEqual(items[0].status, "auto_normalized")
                self.assertEqual(items[0].canonical, canonical)


if __name__ == "__main__":
    unittest.main()
