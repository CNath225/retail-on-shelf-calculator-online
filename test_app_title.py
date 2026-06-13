import unittest

from app import APP_VERSION, app_title_html


class AppTitleTests(unittest.TestCase):
    def test_patch_notes_dropdown_versions_and_style_are_present(self):
        html = app_title_html()

        self.assertIn(APP_VERSION, html)
        self.assertIn('class="patch-notes"', html)
        for label in ["v1-beta", "v1.1-beta", "v1.2-beta", "future"]:
            self.assertIn(label, html)
        self.assertIn("Smart identifier matching beta, default off.", html)
        self.assertIn("W&amp;D matches 53%.", html)
        self.assertIn("Alias map is embedded in a hidden report sheet.", html)
        self.assertIn("font-size: 0.82rem", html)
        self.assertIn("color: #c8a760", html)


if __name__ == "__main__":
    unittest.main()
