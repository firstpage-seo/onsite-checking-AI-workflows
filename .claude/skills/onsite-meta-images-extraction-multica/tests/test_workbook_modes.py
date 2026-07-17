import contextlib
import csv
import importlib.util
import io
import os
import sys
import tempfile
import unittest
from unittest import mock

import openpyxl
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as SpreadsheetImage
from PIL import Image as PillowImage


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(ROOT, "scripts", "onsite_checklist.py")
SPEC = importlib.util.spec_from_file_location("onsite_checklist", SCRIPT)
onsite = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(onsite)


def empty_buckets():
    return {
        "13": {issue: [] for issue, _, _ in onsite.CANONICAL_ISSUES},
        **{key: [] for key in onsite.TABS},
    }


def checklist_row(ws, item):
    for row in range(3, ws.max_row + 1):
        try:
            key = f"{round(float(ws.cell(row, 2).value), 2):.1f}"
        except (TypeError, ValueError):
            continue
        if key == item:
            return row
    raise AssertionError(f"Checklist item not found: {item}")


class WorkbookModesTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.seo = onsite.default_template("seo")
        self.geo = onsite.default_template("seo_geo")
        self.meta = {
            "canonicals_available": True,
            "canonical_missing_reports": [],
            "images_crawled": True,
            "images_sizes_known": True,
        }

    def tearDown(self):
        self.tmp.cleanup()

    def path(self, name):
        return os.path.join(self.tmp.name, name)

    def write_csv(self, directory, name, fieldnames, rows):
        with open(os.path.join(directory, name), "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

    def test_detects_both_template_profiles(self):
        self.assertEqual(onsite.detect_profile(self.seo), "seo")
        self.assertEqual(onsite.detect_profile(self.geo), "seo_geo")

    def test_normalizes_approved_github_file_url(self):
        page = (
            "https://github.com/firstpage-seo/onsite-checking-AI-workflows/blob/main/"
            ".claude/skills/onsite-meta-images-extraction-multica/assets/templates/"
            "Onsite%20Checklist%20-%20TEMPLATE%20-%202026.xlsx?raw=1"
        )
        self.assertEqual(onsite.normalize_template_url(page), onsite.PROFILES["seo"]["url"])
        self.assertEqual(
            onsite.normalize_template_url(
                page.replace("onsite-checking-AI-workflows", "onsite-checking-ai-workflows")
            ),
            onsite.PROFILES["seo"]["url"],
        )
        with self.assertRaises(ValueError):
            onsite.normalize_template_url("https://github.com/another/repo/blob/main/template.xlsx")
        with self.assertRaises(ValueError):
            onsite.normalize_template_url(
                "http://github.com/firstpage-seo/onsite-checking-AI-workflows/"
                "blob/main/template.xlsx"
            )

    def test_downloads_and_validates_canonical_template_offline(self):
        with open(self.seo, "rb") as fh:
            payload = fh.read()

        class Response(io.BytesIO):
            def __init__(self, data):
                super().__init__(data)
                self.headers = {"Content-Length": str(len(data))}

            def __enter__(self):
                return self

            def __exit__(self, *args):
                self.close()

        with mock.patch.object(onsite.urllib.request, "urlopen", return_value=Response(payload)):
            downloaded = onsite.download_template(
                onsite.PROFILES["seo"]["url"], "seo", self.tmp.name,
                expected_sha256=onsite.PROFILES["seo"]["sha256"],
            )
        self.assertEqual(onsite.detect_profile(downloaded), "seo")

        with mock.patch.object(onsite.urllib.request, "urlopen", return_value=Response(payload)):
            with self.assertRaisesRegex(SystemExit, "checksum"):
                onsite.download_template(
                    onsite.PROFILES["seo"]["url"], "seo", self.tmp.name,
                    expected_sha256="0" * 64,
                )

        with mock.patch.object(onsite.urllib.request, "urlopen", return_value=Response(b"not xlsx")):
            with self.assertRaisesRegex(SystemExit, "ZIP signature"):
                onsite.download_template(
                    onsite.PROFILES["seo"]["url"], "seo", self.tmp.name,
                )

    def test_extracts_all_canonical_groups_and_numbered_elements(self):
        exports = self.path("canonical-exports")
        os.makedirs(exports)
        self.write_csv(
            exports, "page_titles_all.csv",
            ["Address", "Indexability", "Title 1 Pixel Width"],
            [{"Address": "https://example.com/", "Indexability": "Indexable",
              "Title 1 Pixel Width": "100"}],
        )
        canonical_fields = [
            "Address", "Canonical Link Element 1", "Canonical Link Element 2",
            "Canonical Link Element 3",
        ]
        self.write_csv(
            exports, "canonicals_all.csv", canonical_fields,
            [{"Address": "https://example.com/", "Canonical Link Element 1": "https://example.com/"}],
        )
        rows_by_file = {
            "canonicals_canonicalised.csv": [
                {"Address": "https://example.com/a", "Canonical Link Element 1": "https://example.com/"},
            ],
            "canonicals_missing.csv": [
                {"Address": "https://example.com/missing"},
            ],
            "canonicals_multiple.csv": [
                {"Address": "https://example.com/multiple",
                 "Canonical Link Element 1": "https://example.com/one",
                 "Canonical Link Element 2": "https://example.com/two",
                 "Canonical Link Element 3": "https://example.com/three"},
            ],
            "canonicals_multiple_conflicting.csv": [
                {"Address": "https://example.com/conflicting",
                 "Canonical Link Element 1": "https://example.com/one",
                 "Canonical Link Element 2": "https://example.com/two"},
            ],
            "canonicals_nonindexable_canonical.csv": [
                {"Address": "https://example.com/nonindexable",
                 "Canonical Link Element 1": "https://example.com/noindex"},
            ],
            "canonicals_invalid_attribute_in_annotation.csv": [
                {"Address": "https://example.com/invalid",
                 "Canonical Link Element 1": "https://example.com/invalid"},
            ],
            "canonicals_contains_fragment_url.csv": [
                {"Address": "https://example.com/fragment",
                 "Canonical Link Element 1": "https://example.com/fragment#part"},
            ],
            "canonicals_outside_head.csv": [
                {"Address": "https://example.com/outside",
                 "Canonical Link Element 1": "https://example.com/outside"},
            ],
        }
        for name, rows in rows_by_file.items():
            self.write_csv(exports, name, canonical_fields, rows)

        buckets, meta = onsite.build_buckets(exports)
        self.assertTrue(meta["canonicals_available"])
        self.assertEqual(
            list(buckets["13"]), [issue for issue, _, _ in onsite.CANONICAL_ISSUES],
        )
        multiple = buckets["13"]["Multiple"][0]
        self.assertEqual(multiple["Canonical Link Element"], "https://example.com/one")
        self.assertEqual(
            multiple["Additional Canonical Link Elements"],
            "https://example.com/two\nhttps://example.com/three",
        )

    def test_canonical_sheet_groups_rows_and_marks_checklist(self):
        buckets = empty_buckets()
        buckets["13"]["Canonicalised"] = [
            {"Address": "https://example.com/a", "Issue": "Canonicalised",
             "Canonical Link Element": "https://example.com/", "Additional Canonical Link Elements": None},
            {"Address": "https://example.com/b", "Issue": "Canonicalised",
             "Canonical Link Element": "https://example.com/", "Additional Canonical Link Elements": None},
        ]
        buckets["13"]["Missing"] = [
            {"Address": "https://example.com/missing", "Issue": "Missing",
             "Canonical Link Element": None, "Additional Canonical Link Elements": None},
        ]
        buckets["13"]["Multiple"] = [
            {"Address": "https://example.com/multiple", "Issue": "Multiple",
             "Canonical Link Element": "https://example.com/one",
             "Additional Canonical Link Elements": "https://example.com/two\nhttps://example.com/three"},
        ]

        output = self.path("canonical-groups.xlsx")
        onsite.build(self.seo, output, buckets, self.meta, mode="begin")
        wb = openpyxl.load_workbook(output)
        sheet = wb[onsite.CANONICAL_TAB]
        self.assertEqual(sheet.cell(3, 4).value, "Additional Canonical Link Element(s)")
        self.assertEqual(sheet.cell(3, 5).value, "Instructions")
        self.assertEqual(sheet.cell(3, 6).value, "Screenshots (If have)")
        self.assertEqual(sheet.cell(4, 2).value, "Canonicalised")
        self.assertEqual(sheet.cell(5, 2).value, "Canonicalised")
        self.assertTrue(all(sheet.cell(6, c).value is None for c in range(1, 7)))
        self.assertEqual(sheet.cell(7, 2).value, "Missing")
        self.assertTrue(all(sheet.cell(8, c).value is None for c in range(1, 7)))
        self.assertEqual(sheet.cell(9, 2).value, "Multiple")
        self.assertEqual(
            sheet.cell(9, 4).value,
            "https://example.com/two\nhttps://example.com/three",
        )
        checklist = wb[onsite.CHECK_TAB]
        self.assertEqual(checklist.cell(checklist_row(checklist, "13.1"), onsite.CHECK_COL).value, "✖")
        self.assertEqual(checklist.cell(checklist_row(checklist, "13.2"), onsite.CHECK_COL).value, "✖")
        for item in onsite.CANONICAL_MANUAL_ITEMS:
            self.assertEqual(
                checklist.cell(checklist_row(checklist, item), onsite.CHECK_COL).value, "N/A"
            )

    def test_canonical_review_preserves_instructions_screenshots_and_manual_marks(self):
        previous_buckets = empty_buckets()
        previous_buckets["13"]["Canonicalised"] = [
            {"Address": "https://example.com/persistent", "Issue": "Canonicalised",
             "Canonical Link Element": "https://example.com/", "Additional Canonical Link Elements": None},
            {"Address": "https://example.com/resolved", "Issue": "Canonicalised",
             "Canonical Link Element": "https://example.com/", "Additional Canonical Link Elements": None},
        ]
        previous = self.path("canonical-previous.xlsx")
        onsite.build(self.seo, previous, previous_buckets, self.meta, mode="begin")
        wb = openpyxl.load_workbook(previous)
        sheet = wb[onsite.CANONICAL_TAB]
        sheet.cell(4, 5).value = "Keep instruction"
        png = self.path("canonical-screenshot.png")
        PillowImage.new("RGB", (4, 4), "blue").save(png)
        sheet.add_image(SpreadsheetImage(png), "F4")
        sheet.add_image(SpreadsheetImage(png), "F5")
        checklist = wb[onsite.CHECK_TAB]
        checklist.cell(checklist_row(checklist, "13.3"), onsite.CHECK_COL).value = "✔"
        wb.save(previous)

        current_buckets = empty_buckets()
        current_buckets["13"]["Canonicalised"] = [
            {"Address": "https://example.com/persistent", "Issue": "Canonicalised",
             "Canonical Link Element": "https://example.com/", "Additional Canonical Link Elements": None},
            {"Address": "https://example.com/new", "Issue": "Canonicalised",
             "Canonical Link Element": "https://example.com/", "Additional Canonical Link Elements": None},
        ]
        refreshed = self.path("canonical-refreshed.xlsx")
        _, changes = onsite.build(
            previous, refreshed, current_buckets, self.meta, mode="review",
            blank_unverified=False,
        )
        wb = openpyxl.load_workbook(refreshed)
        sheet = wb[onsite.CANONICAL_TAB]
        self.assertEqual(sheet.cell(4, 5).value, "Keep instruction")
        self.assertIsNone(sheet.cell(5, 5).value)
        self.assertEqual(len(sheet._images), 1)
        self.assertEqual(onsite._image_position(sheet._images[0]), (4, 6))
        checklist = wb[onsite.CHECK_TAB]
        self.assertEqual(checklist.cell(checklist_row(checklist, "13.3"), onsite.CHECK_COL).value, "✔")
        self.assertEqual(changes["13.2"], {"new": 1, "persistent": 1, "resolved": 1})

    def test_known_nonmissing_canonical_issue_fails_13_2_with_incomplete_evidence(self):
        buckets = empty_buckets()
        buckets["13"]["Canonicalised"] = [
            {"Address": "https://example.com/old", "Issue": "Canonicalised",
             "Canonical Link Element": "https://example.com/new",
             "Additional Canonical Link Elements": None},
        ]
        output = self.path("canonical-incomplete.xlsx")
        onsite.build(
            self.seo, output, buckets,
            {"canonicals_available": False, "images_crawled": True,
             "images_sizes_known": True},
            mode="begin",
        )
        wb = openpyxl.load_workbook(output)
        checklist = wb[onsite.CHECK_TAB]
        self.assertEqual(
            checklist.cell(checklist_row(checklist, "13.1"), onsite.CHECK_COL).value, "?"
        )
        self.assertEqual(
            checklist.cell(checklist_row(checklist, "13.2"), onsite.CHECK_COL).value, "✖"
        )

    def test_review_from_blank_resolves_template_without_history(self):
        exports = self.path("exports")
        os.makedirs(exports)
        with open(os.path.join(exports, "page_titles_all.csv"), "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(
                fh, fieldnames=["Address", "Indexability", "Title 1 Pixel Width"],
            )
            writer.writeheader()
            writer.writerow({
                "Address": "https://example.com/",
                "Indexability": "Indexable",
                "Title 1 Pixel Width": "100",
            })

        argv = [
            SCRIPT,
            "--exports-dir", exports,
            "--client", "Fresh Review",
            "--checklist-type", "seo",
            "--mode", "review",
            "--review-base", "blank",
            "--date", "20260717",
            "--out-dir", self.tmp.name,
        ]
        output = io.StringIO()
        with mock.patch.object(onsite, "resolve_blank_template", return_value=self.seo) as resolver:
            with mock.patch.object(sys, "argv", argv), contextlib.redirect_stdout(output):
                onsite.main()

        resolver.assert_called_once()
        self.assertIn("historical new/persistent/resolved counts are not available", output.getvalue())
        workbook = self.path("Onsite Checklist - Fresh Review - 20260717.xlsx")
        self.assertEqual(onsite.detect_profile(workbook), "seo")
        wb = openpyxl.load_workbook(workbook)
        checklist = wb[onsite.CHECK_TAB]
        self.assertIsNone(checklist.cell(checklist_row(checklist, "10.1"), onsite.CHECK_COL).value)

    def test_review_from_blank_rejects_previous_workbook(self):
        argv = [
            SCRIPT,
            "--exports-dir", self.path("unused-exports"),
            "--client", "Invalid Review",
            "--checklist-type", "seo",
            "--mode", "review",
            "--review-base", "blank",
            "--previous-workbook", self.seo,
        ]
        with mock.patch.object(sys, "argv", argv), contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as raised:
                onsite.main()
        self.assertEqual(raised.exception.code, 2)

    def test_review_merges_rows_and_preserves_human_work(self):
        previous_buckets = empty_buckets()
        previous_buckets["17.1"] = [
            {"Address": "https://example.com/persistent"},
            {"Address": "https://example.com/resolved"},
        ]
        previous_buckets["20.1"] = [
            {"Source": "https://example.com/persistent", "Destination": "https://cdn.example/a.png"},
            {"Source": "https://example.com/resolved", "Destination": "https://cdn.example/b.png"},
        ]
        previous = self.path("previous.xlsx")
        onsite.build(self.seo, previous, previous_buckets, self.meta, mode="begin")

        wb = openpyxl.load_workbook(previous)
        missing = wb["17.1 Page titles - Missing"]
        missing.cell(4, 2).value = "Keep proposed title"
        missing.cell(4, 3).value = "Keep remark"
        missing.cell(4, 3).hyperlink = "https://notes.example/"
        missing.cell(4, 3).comment = Comment("Keep comment", "SEO")
        wb["17.2 Page titles - Duplicate"].sheet_state = "hidden"
        png = self.path("screenshot.png")
        PillowImage.new("RGB", (4, 4), "red").save(png)
        alt_sheet = wb["20.1 Image - Alt Text Missing"]
        alt_sheet.add_image(SpreadsheetImage(png), "C4")
        alt_sheet.add_image(SpreadsheetImage(png), "C5")
        checklist = wb[onsite.CHECK_TAB]
        unrelated = checklist_row(checklist, "10.1")
        checklist.cell(unrelated, onsite.CHECK_COL).value = "✖"
        wb.save(previous)

        current_buckets = empty_buckets()
        current_buckets["17.1"] = [
            {"Address": "https://example.com/persistent"},
            {"Address": "https://example.com/new"},
        ]
        current_buckets["17.2"] = [
            {"Address": "https://example.com/duplicate", "Title 1": "Duplicate title"}
        ]
        current_buckets["17.4"] = [
            {"Address": "https://example.com/multiple", "Title 1": "One", "Title 2": "Two"}
        ]
        current_buckets["20.1"] = [
            {"Source": "https://example.com/persistent", "Destination": "https://cdn.example/a.png"}
        ]

        refreshed = self.path("refreshed.xlsx")
        _, changes = onsite.build(
            previous, refreshed, current_buckets, self.meta, mode="review",
            blank_unverified=False,
        )

        wb = openpyxl.load_workbook(refreshed)
        missing = wb["17.1 Page titles - Missing"]
        self.assertEqual(missing.cell(4, 2).value, "Keep proposed title")
        self.assertEqual(missing.cell(4, 3).value, "Keep remark")
        self.assertEqual(missing.cell(4, 3).hyperlink.target, "https://notes.example/")
        self.assertEqual(missing.cell(4, 3).comment.text, "Keep comment")
        self.assertEqual(missing.cell(5, 1).value, "https://example.com/new")
        self.assertIsNone(missing.cell(6, 1).value)
        self.assertEqual(wb["17.2 Page titles - Duplicate"].sheet_state, "visible")
        self.assertEqual(wb["17.4 Page Titles - Multiple"].cell(4, 2).value, "One")
        self.assertEqual(wb["17.4 Page Titles - Multiple"].cell(4, 3).value, "Two")
        alt_sheet = wb["20.1 Image - Alt Text Missing"]
        self.assertEqual(len(alt_sheet._images), 1)
        self.assertEqual(onsite._image_position(alt_sheet._images[0]), (4, 3))
        self.assertEqual(wb[onsite.CHECK_TAB].cell(unrelated, onsite.CHECK_COL).value, "✖")
        self.assertEqual(changes["17.1"], {"new": 1, "persistent": 1, "resolved": 1})
        self.assertEqual(changes["17.2"], {"new": 1, "persistent": 0, "resolved": 0})
        self.assertEqual(changes["20.1"], {"new": 0, "persistent": 1, "resolved": 1})

    def test_geo_20_3_is_blank_for_begin_and_unchanged_for_review(self):
        begun = self.path("geo-begin.xlsx")
        onsite.build(self.geo, begun, empty_buckets(), self.meta, mode="begin")
        wb = openpyxl.load_workbook(begun)
        checklist = wb[onsite.CHECK_TAB]
        row = checklist_row(checklist, "20.3")
        self.assertIsNone(checklist.cell(row, onsite.CHECK_COL).value)
        checklist.cell(row, onsite.CHECK_COL).value = "✖"
        wb.save(begun)

        reviewed = self.path("geo-review.xlsx")
        onsite.build(
            begun, reviewed, empty_buckets(), self.meta, mode="review",
            blank_unverified=False,
        )
        wb = openpyxl.load_workbook(reviewed)
        self.assertEqual(wb[onsite.CHECK_TAB].cell(row, onsite.CHECK_COL).value, "✖")

    def test_unknown_image_evidence_writes_question_mark(self):
        output = self.path("unknown.xlsx")
        onsite.build(
            self.seo, output, empty_buckets(),
            {"images_crawled": False, "images_sizes_known": False}, mode="begin",
        )
        wb = openpyxl.load_workbook(output)
        checklist = wb[onsite.CHECK_TAB]
        self.assertEqual(checklist.cell(checklist_row(checklist, "13.1"), onsite.CHECK_COL).value, "?")
        self.assertEqual(checklist.cell(checklist_row(checklist, "13.2"), onsite.CHECK_COL).value, "?")
        self.assertEqual(checklist.cell(checklist_row(checklist, "20.1"), onsite.CHECK_COL).value, "?")
        self.assertEqual(checklist.cell(checklist_row(checklist, "20.2"), onsite.CHECK_COL).value, "?")


if __name__ == "__main__":
    unittest.main()
