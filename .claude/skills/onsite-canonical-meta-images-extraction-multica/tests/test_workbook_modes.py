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
            ".claude/skills/onsite-canonical-meta-images-extraction-multica/assets/templates/"
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
        # 13.1 (missing) is auto-marked; 13.2 and 13.3-13.7 are human-judgment -> left blank.
        self.assertEqual(checklist.cell(checklist_row(checklist, "13.1"), onsite.CHECK_COL).value, "✖")
        self.assertIsNone(checklist.cell(checklist_row(checklist, "13.2"), onsite.CHECK_COL).value)
        for item in onsite.CANONICAL_MANUAL_ITEMS:
            self.assertIsNone(
                checklist.cell(checklist_row(checklist, item), onsite.CHECK_COL).value
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

    def test_canonicalised_row_stored_but_13_2_left_for_human(self):
        buckets = empty_buckets()
        buckets["13"]["Canonicalised"] = [
            {"Address": "https://example.com/old", "Issue": "Canonicalised",
             "Canonical Link Element": "https://example.com/new",
             "Additional Canonical Link Elements": None},
        ]
        output = self.path("canonical-human.xlsx")
        onsite.build(
            self.seo, output, buckets,
            {"canonicals_available": False, "images_crawled": True,
             "images_sizes_known": True},
            mode="begin",
        )
        wb = openpyxl.load_workbook(output)
        checklist = wb[onsite.CHECK_TAB]
        # 13.1 unknown (incomplete evidence); 13.2 left BLANK for a human even though
        # a canonicalised URL exists - the skill can't judge if it's correct.
        self.assertEqual(
            checklist.cell(checklist_row(checklist, "13.1"), onsite.CHECK_COL).value, "?"
        )
        self.assertIsNone(
            checklist.cell(checklist_row(checklist, "13.2"), onsite.CHECK_COL).value
        )
        # ...but the canonicalised URL is still stored in the detail sheet as evidence.
        canon = wb[onsite.CANONICAL_TAB]
        addresses = [canon.cell(r, 1).value for r in range(onsite.DATA_ROW, canon.max_row + 1)]
        self.assertIn("https://example.com/old", addresses)

    def test_review_from_blank_base_is_rejected(self):
        # Review from a blank template was removed: Review now always builds
        # from the previous completed onsite so it can preserve the PreCheck block.
        argv = [
            SCRIPT,
            "--exports-dir", self.path("unused-exports"),
            "--client", "No Blank Review",
            "--checklist-type", "seo",
            "--mode", "review",
            "--review-base", "blank",
        ]
        with mock.patch.object(sys, "argv", argv), contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as raised:
                onsite.main()
        self.assertEqual(raised.exception.code, 2)

    def test_review_requires_previous_workbook(self):
        argv = [
            SCRIPT,
            "--exports-dir", self.path("unused-exports"),
            "--client", "No Previous",
            "--checklist-type", "seo",
            "--mode", "review",
        ]
        with mock.patch.object(sys, "argv", argv), contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as raised:
                onsite.main()
        self.assertEqual(raised.exception.code, 2)

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
        self.assertIsNone(checklist.cell(checklist_row(checklist, "13.2"), onsite.CHECK_COL).value)
        self.assertEqual(checklist.cell(checklist_row(checklist, "20.1"), onsite.CHECK_COL).value, "?")
        self.assertEqual(checklist.cell(checklist_row(checklist, "20.2"), onsite.CHECK_COL).value, "?")

    def test_human_review_list_covers_canonical_and_new_items(self):
        meta = dict(self.meta)
        onsite.build(self.seo, self.path("hr.xlsx"), empty_buckets(), meta,
                     mode="begin", family="seo", exports_dir=self.make_new_exports())
        items = {it for it, _ in meta.get("needs_review", [])}
        for expected in ("13.2", "13.3", "13.7", "14.1", "25.1", "25.2", "25.3", "25.4", "26.1", "27.1", "28.2"):
            self.assertIn(expected, items)
        # 25.3 (pages crawled/day) must be flagged as a number field
        why = dict(meta["needs_review"])["25.3"]
        self.assertIn("NUMBER", why)
        # 9.3 is auto-checked, so it must NOT be in the human-review list
        self.assertNotIn("9.3", items)

    def test_human_review_list_renumbers_for_seo_geo(self):
        meta = dict(self.meta)
        onsite.build(self.geo, self.path("hr_geo.xlsx"), empty_buckets(), meta,
                     mode="begin", family="seo_geo", exports_dir=self.make_new_exports())
        items = {it for it, _ in meta.get("needs_review", [])}
        for expected in ("13.2", "29.1", "30.1", "31.1", "32.2"):  # crawling/URL/status/speed renumbered
            self.assertIn(expected, items)

    # ---- new items (HTTPS/Pagination/URL/status codes/page speed) ---------

    def make_new_exports(self):
        d = self.path("new_exports")
        os.makedirs(d, exist_ok=True)
        self.write_csv(d, "page_titles_all.csv", ["Address", "Indexability", "Title 1 Pixel Width"], [
            {"Address": "https://example.com/", "Indexability": "Indexable", "Title 1 Pixel Width": "100"},
            {"Address": "https://example.com/blog?page=1", "Indexability": "Indexable", "Title 1 Pixel Width": "100"},
            {"Address": "https://www.example.com/", "Indexability": "Indexable", "Title 1 Pixel Width": "100"},  # www + non-www -> 9.3 fail
        ])
        self.write_csv(d, "security_http_urls.csv", ["Address"],
                       [{"Address": "http://example.com/insecure"}])          # 9.1 -> fail
        self.write_csv(d, "security_mixed_content.csv", ["Source", "Destination"], [])  # 9.2 -> pass
        self.write_csv(d, "url_parameters.csv", ["Address"],
                       [{"Address": "https://example.com/p?a=1&b=2"}])          # .2 -> fail
        self.write_csv(d, "url_underscores.csv", ["Address"], [])              # .3 -> pass
        self.write_csv(d, "url_uppercase.csv", ["Address"], [])                # .4 -> pass
        self.write_csv(d, "url_multiple_slashes.csv", ["Address"], [])         # .5 -> pass
        self.write_csv(d, "response_codes_server_error_(5xx).csv", ["Address"],
                       [{"Address": "https://example.com/boom"}])              # .2 5xx -> fail
        self.write_csv(d, "response_codes_redirection_(3xx).csv", ["Address", "Status Code", "Redirect URL"],
                       [{"Address": "https://example.com/old", "Status Code": "302",
                         "Redirect URL": "https://example.com/new"}])          # .3 302 -> fail
        self.write_csv(d, "response_codes_client_error_(4xx).csv", ["Address"],
                       [{"Address": "https://example.com/gone"}])
        self.write_csv(d, "response_codes_client_error_(4xx)_inlinks.csv", ["Source", "Destination"],
                       [{"Source": "https://example.com/", "Destination": "https://example.com/gone"}])  # broken -> fail
        self.write_csv(d, "pagination_contains_pagination.csv", ["Address"],
                       [{"Address": "https://example.com/blog?page=2"}])       # 14.3 -> pass
        self.write_csv(d, "pagespeed_all.csv", ["Address", "Performance Score"],
                       [{"Address": "https://example.com/", "Performance Score": "30"}])  # 28.1 -> fail
        return d

    def test_new_items_begin_marks_seo(self):
        output = self.path("new_seo.xlsx")
        onsite.build(self.seo, output, empty_buckets(), dict(self.meta),
                     mode="begin", family="seo", exports_dir=self.make_new_exports())
        wb = openpyxl.load_workbook(output)
        ws = wb[onsite.CHECK_TAB]
        m = lambda item: ws.cell(checklist_row(ws, item), onsite.CHECK_COL).value
        self.assertEqual(m("9.1"), "✖")   # internal HTTP URL
        self.assertEqual(m("9.2"), "✔")   # no mixed content
        self.assertEqual(m("9.3"), "✖")   # both www and non-www present
        # HTTPS detail sheet: http-vs-https URL sits in Source with no Destination
        https = wb["9. Implement HTTPS Protocol"]
        http_row = next(r for r in range(onsite.DATA_ROW, https.max_row + 1)
                        if https.cell(r, 4).value == "HTTP (insecure) URL")
        self.assertEqual(https.cell(http_row, 1).value, "http://example.com/insecure")  # Source
        self.assertIn(https.cell(http_row, 2).value, (None, ""))                          # no Destination
        self.assertIn(https.cell(http_row, 3).value, (None, ""))                          # no Anchor Text
        self.assertEqual(m("14.4"), "✖")  # page=1 exists
        self.assertEqual(m("14.3"), "✔")  # pagination markup present
        self.assertEqual(m("26.2"), "✖")  # parameters
        self.assertEqual(m("26.3"), "✔")  # underscores clean
        self.assertEqual(m("27.2"), "✖")  # 5xx
        self.assertEqual(m("27.3"), "✖")  # 302s
        self.assertEqual(m("24.8"), "✖")  # broken internal links
        self.assertEqual(m("28.1"), "✖")  # PSI mobile < 50

    def _write_file(self, name, text):
        p = self.path(name)
        with open(p, "w", encoding="utf-8") as fh:
            fh.write(text)
        return p

    def test_robots_and_sitemap_auto_checks(self):
        robots = self._write_file("robots.txt",
            "User-agent: *\nDisallow: /wp-admin/\nSitemap: https://www.example.com/sitemap.xml\n")
        sitemap = self._write_file("sitemap.xml",
            '<?xml version="1.0"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
            '  <url><loc>https://www.example.com/a</loc><lastmod>2026-01-01</lastmod></url>\n'
            '  <url><loc>https://www.example.com/shop?page=2</loc><lastmod>2026-01-01</lastmod></url>\n'
            '</urlset>\n')
        output = self.path("robots_sitemap.xlsx")
        onsite.build(self.seo, output, empty_buckets(), dict(self.meta),
                     mode="begin", family="seo", exports_dir=self.make_new_exports(),
                     robots_url=robots, sitemap_url=sitemap)
        wb = openpyxl.load_workbook(output)
        ws = wb[onsite.CHECK_TAB]
        m = lambda item: ws.cell(checklist_row(ws, item), onsite.CHECK_COL).value
        self.assertEqual(m("10.1"), "✔")   # robots.txt exists
        self.assertEqual(m("10.4"), "✔")   # css/js not blocked
        self.assertEqual(m("10.6"), "✔")   # lowercase robots.txt
        self.assertEqual(m("10.7"), "✔")   # sitemap link present
        self.assertEqual(m("11.1"), "✔")   # sitemap exists
        self.assertEqual(m("11.3"), "✔")   # https + single host
        self.assertEqual(m("11.4"), "✔")   # lastmod present
        self.assertEqual(m("11.6"), "✖")   # a paginated URL is in the sitemap
        # robots content and the paginated finding are written to the detail sheets
        self.assertTrue(any("wp-admin" in str(wb["10. Robots.txt"].cell(r, 1).value)
                            for r in range(4, wb["10. Robots.txt"].max_row + 1)))
        self.assertTrue(any("page=2" in str(wb["11. Sitemap"].cell(r, 2).value or "")
                            for r in range(4, wb["11. Sitemap"].max_row + 1)))

    def test_robots_sitemap_absent_falls_back_to_human_review(self):
        meta = dict(self.meta)
        onsite.build(self.seo, self.path("no_rs.xlsx"), empty_buckets(), meta,
                     mode="begin", family="seo", exports_dir=self.make_new_exports())
        items = {it for it, _ in meta.get("needs_review", [])}
        for expected in ("10.1", "10.7", "11.1", "11.6"):  # not readable -> human review
            self.assertIn(expected, items)

    def test_www_clean_redirect_passes_9_3(self):
        d = self.path("www_redirect_exports")
        os.makedirs(d)
        self.write_csv(d, "page_titles_all.csv", ["Address", "Indexability", "Title 1 Pixel Width"],
                       [{"Address": "https://example.com/", "Indexability": "Indexable", "Title 1 Pixel Width": "100"}])
        # the www variant exists only as a 301 redirect -> clean canonicalisation -> pass
        self.write_csv(d, "response_codes_redirection_(3xx).csv", ["Address", "Status Code", "Redirect URL"],
                       [{"Address": "https://www.example.com/", "Status Code": "301",
                         "Redirect URL": "https://example.com/"}])
        output = self.path("www_pass.xlsx")
        onsite.build(self.seo, output, empty_buckets(), dict(self.meta),
                     mode="begin", family="seo", exports_dir=d)
        ws = openpyxl.load_workbook(output)[onsite.CHECK_TAB]
        self.assertEqual(ws.cell(checklist_row(ws, "9.3"), onsite.CHECK_COL).value, "✔")

    def test_new_items_renumber_for_seo_geo(self):
        output = self.path("new_geo.xlsx")
        onsite.build(self.geo, output, empty_buckets(), dict(self.meta),
                     mode="begin", family="seo_geo", exports_dir=self.make_new_exports())
        ws = openpyxl.load_workbook(output)[onsite.CHECK_TAB]
        m = lambda item: ws.cell(checklist_row(ws, item), onsite.CHECK_COL).value
        self.assertEqual(m("30.2"), "✖")  # URL params (renumbered from 26.2)
        self.assertEqual(m("31.2"), "✖")  # 5xx (renumbered from 27.2)
        self.assertEqual(m("28.8"), "✖")  # broken internal links (from 24.8)
        self.assertEqual(m("32.1"), "✖")  # page speed (from 28.1)

    def test_page_speed_left_blank_without_psi_data(self):
        d = self.make_new_exports()
        os.remove(os.path.join(d, "pagespeed_all.csv"))
        output = self.path("no_psi.xlsx")
        onsite.build(self.seo, output, empty_buckets(), dict(self.meta),
                     mode="begin", family="seo", exports_dir=d)
        ws = openpyxl.load_workbook(output)[onsite.CHECK_TAB]
        self.assertIsNone(ws.cell(checklist_row(ws, "28.1"), onsite.CHECK_COL).value)

    def test_review_preserves_prechecks_and_reaudits_technical(self):
        previous = self.path("previous.xlsx")
        wb = openpyxl.load_workbook(self.seo)
        cl = wb[onsite.CHECK_TAB]
        cl.cell(checklist_row(cl, "1.4"), onsite.CHECK_COL).value = "✔"   # PreCheck §1 -> keep
        cl.cell(checklist_row(cl, "10.1"), onsite.CHECK_COL).value = "✖"  # §10 technical -> re-audit
        cl.cell(checklist_row(cl, "31.1"), onsite.CHECK_COL).value = "✔"  # §31 backlink -> keep
        wb.save(previous)

        output = self.path("reviewed.xlsx")
        onsite.build(previous, output, empty_buckets(), dict(self.meta),
                     mode="begin", blank_unverified=False, keep_prechecks=True,
                     family="seo", exports_dir=self.make_new_exports())
        ws = openpyxl.load_workbook(output)[onsite.CHECK_TAB]
        m = lambda item: ws.cell(checklist_row(ws, item), onsite.CHECK_COL).value
        self.assertEqual(m("1.4"), "✔")    # PreCheck preserved
        self.assertEqual(m("31.1"), "✔")   # backlink section preserved
        self.assertIsNone(m("10.1"))       # technical section blanked (re-audit fresh)
        self.assertEqual(m("9.1"), "✖")    # managed new item re-audited from crawl

    def test_review_records_previous_result_in_column_e(self):
        # Previous single-column onsite (production layout: D = check, E = detail links).
        previous = self.path("Onsite Checklist - acme - 2026.02.10.xlsx")
        wb = openpyxl.load_workbook(self.seo)
        cl = wb[onsite.CHECK_TAB]
        cl.cell(checklist_row(cl, "1.4"), onsite.CHECK_COL).value = "✔"    # PreCheck -> preserved in D
        cl.cell(checklist_row(cl, "10.1"), onsite.CHECK_COL).value = "✖"   # technical -> re-audited in D
        link_before = cl.cell(checklist_row(cl, "13.1"), onsite.RECORD_COL).value  # detail-link text in E
        wb.save(previous)

        output = self.path("Onsite Checklist - acme - 20260826.xlsx")
        onsite.build(previous, output, empty_buckets(), dict(self.meta),
                     mode="begin", blank_unverified=False, keep_prechecks=True,
                     family="seo", exports_dir=self.make_new_exports(),
                     record_previous=True)
        ws = openpyxl.load_workbook(output)[onsite.CHECK_TAB]
        hr = onsite._checklist_header_row(ws)
        # A fresh column E was inserted; detail links shifted to F.
        self.assertEqual(ws.cell(hr, onsite.CHECK_COL).value, "Check in Aug")   # D = current run month
        self.assertEqual(ws.cell(hr, onsite.RECORD_COL).value, "Check in Feb")  # E = previous file month
        self.assertEqual(ws.cell(hr, onsite.RECORD_COL + 1).value,
                         "Detailed Tab Links / Specific Instructions")
        self.assertEqual(ws.cell(checklist_row(ws, "13.1"), onsite.RECORD_COL + 1).value, link_before)
        # Column E keeps the previous audit verbatim as the record.
        self.assertEqual(ws.cell(checklist_row(ws, "1.4"), onsite.RECORD_COL).value, "✔")
        self.assertEqual(ws.cell(checklist_row(ws, "10.1"), onsite.RECORD_COL).value, "✖")
        # Column D shows the latest audit: PreCheck preserved, technical re-audited fresh.
        m = lambda item: ws.cell(checklist_row(ws, item), onsite.CHECK_COL).value
        self.assertEqual(m("1.4"), "✔")   # PreCheck preserved in current column
        self.assertIsNone(m("10.1"))            # technical blanked then re-audited
        self.assertEqual(m("9.1"), "✖")   # managed item re-audited from the new crawl


if __name__ == "__main__":
    unittest.main()
