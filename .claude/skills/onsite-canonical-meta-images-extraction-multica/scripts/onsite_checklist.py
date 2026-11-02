#!/usr/bin/env python3
"""
Get crawl data from Screaming Frog and produce either the SEO or SEO/GEO
FirstPage Onsite Checklist deliverable for items 13.1-13.2 and 17.1-20.2.

Two phases, independently runnable:
  crawl  -> Screaming Frog CLI writes CSV exports to a folder
  build  -> read those CSVs, apply house thresholds, write the .xlsx

Use --exports-dir to rebuild from an existing export folder without re-crawling.
"""

import argparse
import collections
import copy
import csv
import datetime
import glob
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

SF_BIN = "/Applications/Screaming Frog SEO Spider.app/Contents/MacOS/ScreamingFrogSEOSpiderLauncher"

# ---- House thresholds. These live here, NOT in Screaming Frog's config, so a
# ---- stale GUI setting can never silently change the numbers.
TITLE_PX = 580          # 17.3  (per checklist tab header)
DESC_PX = 880           # 18.3  (per tab note: Config > Spider > Preferences > 880 mobile max)
IMG_KB = 200            # 20.2
KB = 1024               # SF's kB convention is unconfirmed (1000 vs 1024); 100kB export's
                        # smallest row was 102876 bytes, which exceeds both. Affects only
                        # images within ~4.8KB of the cutoff.

EXPORT_TABS = [
    "Canonicals:All", "Canonicals:Canonicalised", "Canonicals:Missing",
    "Canonicals:Multiple", "Canonicals:Multiple Conflicting",
    "Canonicals:Non-Indexable Canonical", "Canonicals:Invalid Attribute In Annotation",
    "Canonicals:Contains Fragment URL", "Canonicals:Outside <head>",
    "Page Titles:All", "Page Titles:Missing", "Page Titles:Duplicate", "Page Titles:Multiple",
    "Meta Description:All", "Meta Description:Missing", "Meta Description:Duplicate",
    "H1:All", "H1:Missing", "H1:Duplicate", "H1:Multiple",
    "Images:All",
    "Security:HTTP URLs", "Security:Mixed Content",
    "Response Codes:Client Error (4xx)", "Response Codes:Server Error (5xx)",
    "Response Codes:Redirection (3xx)",
    "URL:Underscores", "URL:Uppercase", "URL:Parameters", "URL:Multiple Slashes",
    "URL:Over X Characters", "Pagination:Contains Pagination", "PageSpeed:All",
]

CANONICAL_ISSUES = (
    ("Canonicalised", ("canonicals", "canonicalised"), ()),
    ("Missing", ("canonicals", "missing"), ()),
    ("Multiple", ("canonicals", "multiple"), ("conflicting",)),
    ("Multiple Conflicting", ("canonicals", "multiple", "conflicting"), ()),
    ("Non-Indexable Canonical", ("canonicals", "nonindexable", "canonical"), ()),
    ("Invalid Attribute In Annotation", ("canonicals", "invalid", "attribute"), ()),
    ("Contains Fragment URL", ("canonicals", "contains", "fragment"), ()),
    ("Outside <head>", ("canonicals", "outside", "head"), ()),
)
BULK_EXPORTS = [
    "Images:Images Missing Alt Attribute & Text Inlinks",
    "Images:All Image Inlinks",
    "Response Codes:Client Error (4xx) Inlinks",
    "Response Codes:Server Error (5xx) Inlinks",
]


# --------------------------------------------------------------------------- crawl

def run_crawl(url, sitemap, outdir, config=None, client=None, date=None, crawl_file=None):
    outdir = os.path.abspath(outdir)          # MUST be absolute: the launcher resets cwd to
    os.makedirs(outdir, exist_ok=True)        # the app bundle, so relative paths hit /Applications
    if not os.path.exists(SF_BIN):
        sys.exit(f"Screaming Frog not found at {SF_BIN}")

    cmd = [SF_BIN, "--headless"]

    if crawl_file:
        # Export from a crawl the user already ran in the GUI. Their settings, their crawl
        # analysis, no config file needed. Loading from a FILE path works even when a crawl
        # is open in the GUI; loading by database ID does not.
        cmd += ["--load-crawl", os.path.abspath(crawl_file)]
        print(f"[load]  crawl <- {crawl_file}")
    else:
        # Always start a Spider crawl. --crawl-sitemap opens the saved crawl in List mode.
        # A client sitemap cannot be injected into "Crawl These Sitemaps" through the CLI,
        # so the local headless route deliberately accepts website-only crawls.
        if not url:
            sys.exit("The local headless route needs a website URL.")
        cmd += ["--crawl", url]
        if config:
            cmd += ["--config", os.path.abspath(config)]
        # --save-crawl puts the crawl in SF's DB so it can be opened in the GUI afterwards.
        # Name it, or it lands in the crawl list as a bare URL and is a pain to find later.
        if client:
            cmd += ["--task-name", f"{client} - {date}" if date else client,
                    "--project-name", client]
        cmd += ["--save-crawl"]
        print(f"[crawl] Spider mode from {url}")
        print(f"[crawl] config  -> {config or 'Screaming Frog DEFAULTS (not your GUI settings)'}")

    cmd += [
        "--output-folder", outdir,
        "--export-format", "csv",
        "--overwrite",
        "--export-tabs", ",".join(EXPORT_TABS),
        "--bulk-export", ",".join(BULK_EXPORTS),
    ]
    print(f"[run]   output -> {outdir}")
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        print(proc.stdout[-3000:])
        print(proc.stderr[-3000:], file=sys.stderr)
        sys.exit(f"Screaming Frog exited {proc.returncode}")
    for line in proc.stdout.splitlines():
        if "Completed the spider" in line or "completed urls through SpiderResults" in line:
            print("[run]   " + line.split("INFO  - ")[-1])
    return outdir


# --------------------------------------------------------------------------- io

def find_csv(outdir, *must, exclude=()):
    """Locate an export by keyword. Filenames are SF's own slugs, so match loosely."""
    for path in sorted(glob.glob(os.path.join(outdir, "*.csv"))):
        name = os.path.basename(path).lower()
        if all(m in name for m in must) and not any(x in name for x in exclude):
            return path
    return None


def load(path):
    if not path or not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def col(row, *names):
    """Fetch the first present column; SF header names vary between exports."""
    for n in names:
        if n in row and row[n] not in (None, ""):
            return row[n]
    return None


def num(v):
    try:
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return 0


def _addr(row):
    return col(row, "Address", "URL", "Source")


def indexable(rows):
    return [r for r in rows if (r.get("Indexability") or "").strip() == "Indexable"]


def group_by(rows, *keys):
    """Order duplicate rows so identical values sit together, biggest group first."""
    val = lambda r: (col(r, *keys) or "").strip()
    counts = collections.Counter(val(r) for r in rows)
    return sorted(rows, key=lambda r: (-counts[val(r)], val(r), col(r, "Address") or ""))


def canonical_elements(row):
    """Return numbered HTML canonical link elements in Screaming Frog export order."""
    elements = []
    for heading, value in row.items():
        match = re.fullmatch(r"Canonical Link Element(?: (\d+))?", heading or "")
        if match and value not in (None, ""):
            elements.append((int(match.group(1) or 1), str(value).strip()))
    return [value for _, value in sorted(elements)]


def canonical_record(row, issue):
    elements = canonical_elements(row)
    return {
        "Address": col(row, "Address"),
        "Issue": issue,
        "Canonical Link Element": elements[0] if elements else None,
        "Additional Canonical Link Elements": "\n".join(elements[1:]) or None,
    }


# --------------------------------------------------------------------------- buckets

def build_buckets(d):
    t_all = load(find_csv(d, "page_titles", "all"))
    d_all = load(find_csv(d, "meta_description", "all"))
    i_all = load(find_csv(d, "images_all"))

    # A failed crawl still writes header-only CSVs, which would sail through as
    # "zero issues everywhere" and produce an all-ticks deliverable for a site
    # nobody actually crawled. Refuse rather than report a false clean bill.
    if not t_all:
        sys.exit("No pages in Page Titles:All - the crawl produced nothing. "
                 "Refusing to build a deliverable that would read as an all-clear.")

    # 19.2 = cross-page H1-1 duplicates. Start from SF's Duplicate filter (which already
    # excludes noindex pages, handles pagination, etc. - do NOT reproduce that by grouping
    # h1_all, which pulls in noindex pages SF omits and inflates the count).
    #
    # SF's Duplicate ALSO flags pages whose H1-1 == H1-2 on the SAME page. Drop those - but
    # ONLY when that H1-1 value is unique within the set. A page can both repeat its H1 on
    # the page AND share that H1-1 with other pages (crefit /clearcardcredit/, H1-1
    # "健康貸度 由你掌控" on 4 pages); that is a real cross-page duplicate and must stay. The
    # earlier "drop every H1-1==H1-2 row" rule wrongly removed those. Pure same-page
    # collisions (unique H1-1) belong in 19.3 Multiple, where they already appear.
    h_dupe_raw = load(find_csv(d, "h1", "duplicate"))
    hnorm = lambda r: (col(r, "H1-1") or "").strip().lower()
    dup_h1_counts = collections.Counter(hnorm(r) for r in h_dupe_raw if hnorm(r))
    h_dupe = [r for r in h_dupe_raw
              if (col(r, "H1-1") or "").strip() != (col(r, "H1-2") or "").strip()
              or dup_h1_counts[hnorm(r)] > 1]

    # 20.2: read Size (Bytes) straight from All Image Inlinks - it carries a size on every
    # row, INCLUDING external-domain images. (An earlier version joined to Images:All for
    # size, but Images:All is internal-only, so it silently dropped every external-CDN image
    # - Prismic, Cloudinary, agentpro, etc. Reading size from the inlink row fixes that.)
    # Threshold stays in code (IMG_KB), not SF's configured "Over X kB".
    inlinks = load(find_csv(d, "all_image_inlinks"))
    alt_missing = load(find_csv(d, "missing_alt_attribute"))
    i_over = []
    sizes_seen = False
    for r in inlinks:
        size = num(col(r, "Size (Bytes)", "Size"))
        if size > 0:
            sizes_seen = True
        if size > IMG_KB * KB:
            i_over.append({"Source": col(r, "Source"),
                           "Destination": col(r, "Destination"), "Size (Bytes)": int(size)})
    i_over.sort(key=lambda x: -x["Size (Bytes)"])

    canonical_groups = {}
    canonical_paths = {}
    for issue, must, exclude in CANONICAL_ISSUES:
        path = find_csv(d, *must, exclude=exclude)
        canonical_paths[issue] = path
        canonical_groups[issue] = [canonical_record(row, issue) for row in load(path)]
    canonical_all_path = find_csv(d, "canonicals", "all")
    canonical_all = load(canonical_all_path)
    canonical_missing_reports = [
        issue for issue, path in canonical_paths.items() if not path
    ]
    if not canonical_all_path:
        canonical_missing_reports.append("All")
    elif not canonical_all:
        canonical_missing_reports.append("All (no data)")

    buckets = {
        "13": canonical_groups,
        "17.1": load(find_csv(d, "page_titles", "missing")),
        "17.2": group_by(load(find_csv(d, "page_titles", "duplicate")), "Title 1"),
        "17.3": sorted([r for r in indexable(t_all) if num(col(r, "Title 1 Pixel Width")) > TITLE_PX],
                       key=lambda r: -num(col(r, "Title 1 Pixel Width"))),
        "17.4": load(find_csv(d, "page_titles", "multiple")),
        "18.1": load(find_csv(d, "meta_description", "missing")),
        "18.2": group_by(load(find_csv(d, "meta_description", "duplicate")), "Meta Description 1"),
        "18.3": sorted([r for r in indexable(d_all) if num(col(r, "Meta Description 1 Pixel Width")) > DESC_PX],
                       key=lambda r: -num(col(r, "Meta Description 1 Pixel Width"))),
        "19.1": load(find_csv(d, "h1", "missing")),
        "19.2": group_by(h_dupe, "H1-1"),
        "19.3": load(find_csv(d, "h1", "multiple")),
        "20.1": alt_missing,
        "20.2": i_over,
    }
    # Did the crawl capture ANY images? Use every image export, not just Images:All -
    # Images:All lists INTERNAL images only, so a site on an external CDN has an empty
    # Images:All yet real image issues in the inlinks/alt exports (crefit: 0 in Images:All,
    # 17 missing-alt on prismic.io). If NOTHING across the image exports has rows, the crawl
    # genuinely captured no images (image crawling off, or JS rendering off on a JS-loaded
    # site) - flag 20.1/20.2 as UNKNOWN rather than report a false all-clear.
    #
    # images_sizes_known says whether All Image Inlinks supplied any byte sizes at all. If
    # images exist but that export has no sizes, a zero oversize result is unknown, not clean.
    meta = {
        "canonicals_available": bool(canonical_all) and all(canonical_paths.values()),
        "canonical_missing_reports": canonical_missing_reports,
        "images_crawled": bool(i_all or inlinks or alt_missing),
        "images_sizes_known": sizes_seen,
    }
    return buckets, meta


# --------------------------------------------------------------------------- workbook

s = lambda r, *k: ((col(r, *k) or "").strip() or None)

# Machine columns are refreshed from the crawl. Preserve columns contain human work that is
# carried forward for persistent issues in Review mode. key_cols identify an issue across runs.
TABS = {
    "17.1": {"tab": "17.1 Page titles - Missing",
             "rowfn": lambda r: [col(r, "Address")], "machine_cols": 1,
             "preserve_cols": (2, 3), "key_cols": (1,)},
    "17.2": {"tab": "17.2 Page titles - Duplicate",
             "rowfn": lambda r: [col(r, "Address"), s(r, "Title 1")], "machine_cols": 2,
             "preserve_cols": (3, 4), "key_cols": (1, 2)},
    "17.3": {"tab": "17.3 Page Titles - Too Long",
             "rowfn": lambda r: [col(r, "Address"), s(r, "Title 1"),
                                   int(num(col(r, "Title 1 Pixel Width")))], "machine_cols": 3,
             "preserve_cols": (4,), "key_cols": (1,)},
    "17.4": {"tab": "17.4 Page Titles - Multiple",
             "rowfn": lambda r: [col(r, "Address"), s(r, "Title 1"), s(r, "Title 2")],
             "machine_cols": 3, "preserve_cols": (4,), "key_cols": (1,)},
    "18.1": {"tab": "18. Descriptions - Missing",
             "rowfn": lambda r: [col(r, "Address")], "machine_cols": 1,
             "preserve_cols": (2,), "key_cols": (1,)},
    "18.2": {"tab": "18.2 Descriptions - Duplicate",
             "rowfn": lambda r: [col(r, "Address"), s(r, "Meta Description 1")],
             "machine_cols": 2, "preserve_cols": (), "key_cols": (1, 2)},
    "18.3": {"tab": "18.3 Descriptions - Too Long",
             "rowfn": lambda r: [col(r, "Address"), s(r, "Meta Description 1"),
                                   int(num(col(r, "Meta Description 1 Pixel Width")))],
             "machine_cols": 3, "preserve_cols": (), "key_cols": (1,)},
    "19.1": {"tab": "19.1 H1 - Missing",
             "rowfn": lambda r: [col(r, "Address")], "machine_cols": 1,
             "preserve_cols": (2,), "key_cols": (1,)},
    "19.2": {"tab": "19.2 H1 - Duplicate",
             "rowfn": lambda r: [col(r, "Address"), s(r, "H1-1")], "machine_cols": 2,
             "preserve_cols": (3,), "key_cols": (1, 2)},
    "19.3": {"tab": "19.3 H1 - Multiple",
             "rowfn": lambda r: [col(r, "Address"), s(r, "H1-1"), s(r, "H1-2")],
             "machine_cols": 3, "preserve_cols": (), "key_cols": (1,)},
    "20.1": {"tab": "20.1 Image - Alt Text Missing",
             "rowfn": lambda r: [col(r, "Source"), col(r, "Destination")], "machine_cols": 2,
             "preserve_cols": (3,), "key_cols": (1, 2)},
    "20.2": {"tab": "20.2 Image - Oversize",
             "rowfn": lambda r: [r["Source"], r["Destination"], r["Size (Bytes)"]],
             "machine_cols": 3, "preserve_cols": (), "key_cols": (1, 2)},
}

CANONICAL_TAB = "13. Canonical tag"
CANONICAL_CHECK_ITEMS = ("13.1", "13.2")
CANONICAL_MANUAL_ITEMS = ("13.3", "13.4", "13.5", "13.6", "13.7")
MANAGED_CHECK_ITEMS = set(TABS) | set(CANONICAL_CHECK_ITEMS) | set(CANONICAL_MANUAL_ITEMS)
CHANGE_ORDER = (*CANONICAL_CHECK_ITEMS, *TABS)

PROFILES = {
    "seo": {
        "template": "Onsite Checklist - TEMPLATE - 2026.xlsx",
        "output": "Onsite Checklist - {client} - {date}.xlsx",
        "url": "https://raw.githubusercontent.com/firstpage-seo/onsite-checking-AI-workflows/main/.claude/skills/onsite-canonical-meta-images-extraction-multica/assets/templates/Onsite%20Checklist%20-%20TEMPLATE%20-%202026.xlsx",
        "sha256": "69a5a6b0fb78602cbbf98093379d4d4fe0906c31eecb5f33f567d9e668eb4cd6",
    },
    "seo_geo": {
        "template": "[Working Draft_Internal SEO_GEO] Onsite Checklist - [TEMPLATE] - 20260714.xlsx",
        "output": "SEO_GEO Onsite Checklist - {client} - {date}.xlsx",
        "url": "https://raw.githubusercontent.com/firstpage-seo/onsite-checking-AI-workflows/main/.claude/skills/onsite-canonical-meta-images-extraction-multica/assets/templates/%5BWorking%20Draft_Internal%20SEO_GEO%5D%20Onsite%20Checklist%20-%20%5BTEMPLATE%5D%20-%2020260714.xlsx",
        "sha256": "4efcf9f3a6ea08b423e192f727049cf07a86649bb30659df1c10905d5d68ce12",
    },
}

APPROVED_GITHUB_OWNER = "firstpage-seo"
APPROVED_GITHUB_REPO = "onsite-checking-AI-workflows"
MAX_TEMPLATE_BYTES = 5 * 1024 * 1024

DATA_ROW = 4          # rows 1-2 metadata, row 3 headers, data from row 4
CHECK_TAB = "SEO Implementation Checklist"
CHECK_COL = 4         # column D, "Initial Check"
RECORD_COL = 5        # column E — Review keeps the previous audit here for the record

# The checklist tab was renamed across template revisions; accept either name and
# locate rows/tabs by wording, never by a hard-coded sheet name alone.
CHECK_TAB_ALIASES = ("SEO Implementation Checklist", "SEO Audit Checklist")

_MONTH_ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
_MONTH_FULL = ("January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December")


def resolve_check_tab(sheetnames):
    """Return the checklist tab actually present, tolerating template renames."""
    for name in CHECK_TAB_ALIASES:
        if name in sheetnames:
            return name
    return None


def _month_from_filename(path):
    """3-letter month from a YYYYMMDD / YYYY.MM.DD / YYYY-MM-DD date in the name."""
    name = os.path.basename(str(path or ""))
    m = re.search(r"20\d{2}[.\-_ ]?(\d{2})[.\-_ ]?\d{2}", name)
    if m:
        mm = int(m.group(1))
        if 1 <= mm <= 12:
            return _MONTH_ABBR[mm - 1]
    return None


def _month_from_header(text):
    """3-letter month named in an existing 'Check in Jul' / 'Check July' header."""
    t = str(text or "")
    for mon in _MONTH_ABBR:
        if re.search(rf"\b{mon}", t, re.IGNORECASE):
            return mon
    for i, mon in enumerate(_MONTH_FULL):
        if re.search(rf"\b{mon}\b", t, re.IGNORECASE):
            return _MONTH_ABBR[i]
    return None


def _checklist_header_row(ws):
    """Row carrying the '#/ITEM/Check ...' headers on the checklist tab (usually 1)."""
    for r in range(1, 5):
        d = str(ws.cell(r, CHECK_COL).value or "").lower()
        b = str(ws.cell(r, 2).value or "").strip()
        if "check" in d or b == "#":
            return r
    return 1


def record_previous_check_column(ws, prev_path, out_path):
    """Review only: preserve the previous audit's result column (D) into column E for
    the record, then relabel the headers 'Check in <prev>' / 'Check in <cur>'.

    - Two-column input (E is already a check column): D is copied over E; any older
      third audit already in E is dropped, so only the two latest audits are kept.
    - Single-column input (E is the detail-links column): a fresh column is inserted at
      E first, so the detail-links column shifts right and the previous D lands in E.
    The current run's marks are written into D afterwards by the normal audit path.
    Applies to the checklist tab only; every detail tab keeps latest-status-only.
    """
    header_row = _checklist_header_row(ws)
    e_header = str(ws.cell(header_row, RECORD_COL).value or "")
    if "check" not in e_header.lower():
        ws.insert_cols(RECORD_COL)          # single result column -> make room for E
    prev_month = (_month_from_header(ws.cell(header_row, CHECK_COL).value)
                  or _month_from_filename(prev_path))
    for r in range(1, ws.max_row + 1):
        dst = ws.cell(r, RECORD_COL)
        if isinstance(dst, MergedCell):
            continue
        src = ws.cell(r, CHECK_COL)
        dst.value = src.value
        dst._style = copy.copy(src._style)
    ws.cell(header_row, RECORD_COL).value = (
        f"Check in {prev_month}" if prev_month else "Previous Check (record)")
    cur_month = _month_from_filename(out_path)
    if cur_month:
        ws.cell(header_row, CHECK_COL).value = f"Check in {cur_month}"


def find_check_rows(ws):
    """Map '17.1' -> sheet row. Located by column B value, not hardcoded, because
    row positions shift whenever the template gains items."""
    out = {}
    for r in range(3, ws.max_row + 1):
        b = ws.cell(r, 2).value
        try:
            key = f"{round(float(b), 2):.1f}"     # col B carries float noise: 17.400000000000006
        except (TypeError, ValueError):
            continue
        if key in MANAGED_CHECK_ITEMS and key not in out:
            out[key] = r
    return out


def _norm_key(value):
    return str(value or "").strip()


def _issue_key(values, key_cols):
    return tuple(_norm_key(values[c - 1]) for c in key_cols)


def _change_counts(previous, current):
    old = collections.Counter(previous)
    new = collections.Counter(current)
    persistent = sum((old & new).values())
    return {
        "new": sum((new - old).values()),
        "persistent": persistent,
        "resolved": sum((old - new).values()),
    }


def _image_position(image):
    """Return a loaded image's one-based anchor row/column when available."""
    anchor = image.anchor
    if isinstance(anchor, str):
        cell = openpyxl.utils.cell.coordinate_from_string(anchor)
        return cell[1], openpyxl.utils.cell.column_index_from_string(cell[0])
    marker = getattr(anchor, "_from", None)
    if marker is None:
        return None
    return marker.row + 1, marker.col + 1


def _capture_cell(cell):
    return cell.value, copy.copy(cell.hyperlink), copy.copy(cell.comment)


def _restore_cell(cell, captured):
    value, hyperlink, comment = captured
    cell.value = value
    cell.hyperlink = hyperlink
    cell.comment = comment


def detect_profile(workbook_path):
    """Identify the template family from workbook structure, never its filename alone."""
    try:
        wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=False)
    except Exception as exc:
        sys.exit(f"Could not open base workbook as XLSX: {exc}")
    sheets = set(wb.sheetnames)
    wb.close()
    if not resolve_check_tab(sheets):
        sys.exit(f"Workbook is missing the checklist tab (one of {CHECK_TAB_ALIASES}).")
    required = {CANONICAL_TAB, *(spec["tab"] for spec in TABS.values())}
    missing = sorted(required - sheets)
    if missing:
        sys.exit(f"Workbook is missing required sheets: {missing}")
    return "seo_geo" if "20.3 Image - Next-gen" in sheets else "seo"


def normalize_template_url(url):
    """Convert an approved GitHub file page to its raw URL and reject other repositories."""
    parsed = urllib.parse.urlsplit(url.strip())
    if parsed.scheme.lower() != "https":
        raise ValueError("Template URL must use HTTPS")

    host = (parsed.hostname or "").lower()
    parts = [part for part in parsed.path.split("/") if part]
    owner_repo = [APPROVED_GITHUB_OWNER, APPROVED_GITHUB_REPO]
    approved = [part.lower() for part in owner_repo]

    if host == "github.com":
        if len(parts) < 5 or [part.lower() for part in parts[:2]] != approved or parts[2] != "blob":
            raise ValueError("Template URL must point to a file in the approved FirstPage repository")
        ref = parts[3]
        file_path = urllib.parse.quote(urllib.parse.unquote("/".join(parts[4:])), safe="/")
        return f"https://raw.githubusercontent.com/{owner_repo[0]}/{owner_repo[1]}/{ref}/{file_path}"

    if host == "raw.githubusercontent.com":
        if len(parts) < 4 or [part.lower() for part in parts[:2]] != approved:
            raise ValueError("Template URL must point to the approved FirstPage repository")
        raw_path = urllib.parse.quote(urllib.parse.unquote("/".join(parts)), safe="/")
        return urllib.parse.urlunsplit(("https", host, "/" + raw_path, "", ""))

    raise ValueError("Template URL host must be github.com or raw.githubusercontent.com")


def download_template(url, checklist_type, output_dir, expected_sha256=None):
    """Download and validate an approved template into a temporary task directory."""
    try:
        raw_url = normalize_template_url(url)
    except ValueError as exc:
        sys.exit(f"Invalid template URL: {exc}")

    request = urllib.request.Request(raw_url, headers={"User-Agent": "FirstPage-Onsite-Skill/1.0"})
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            length = response.headers.get("Content-Length")
            if length and int(length) > MAX_TEMPLATE_BYTES:
                sys.exit(f"Template download exceeds {MAX_TEMPLATE_BYTES // (1024 * 1024)} MB")
            chunks = []
            total = 0
            while True:
                chunk = response.read(64 * 1024)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_TEMPLATE_BYTES:
                    sys.exit(f"Template download exceeds {MAX_TEMPLATE_BYTES // (1024 * 1024)} MB")
                chunks.append(chunk)
    except (urllib.error.URLError, TimeoutError, OSError, ValueError) as exc:
        sys.exit(f"Could not download template: {exc}")

    payload = b"".join(chunks)
    if not payload.startswith(b"PK"):
        sys.exit("Downloaded template is not an XLSX file (ZIP signature missing)")

    digest = hashlib.sha256(payload).hexdigest()
    if expected_sha256 and digest != expected_sha256:
        sys.exit("Downloaded template checksum does not match the approved version")

    os.makedirs(output_dir, exist_ok=True)
    path = os.path.join(output_dir, PROFILES[checklist_type]["template"])
    with open(path, "wb") as fh:
        fh.write(payload)
    print(f"[template] downloaded {checklist_type} <- {raw_url}")
    return path


def resolve_blank_template(local_path, template_url, checklist_type, output_dir):
    if local_path:
        return local_path
    source = template_url or PROFILES[checklist_type]["url"]
    try:
        canonical = normalize_template_url(PROFILES[checklist_type]["url"])
        supplied = normalize_template_url(source)
    except ValueError as exc:
        sys.exit(f"Invalid template URL: {exc}")
    expected = PROFILES[checklist_type]["sha256"] if supplied.lower() == canonical.lower() else None
    return download_template(source, checklist_type, output_dir, expected_sha256=expected)


def ensure_canonical_schema(sheet):
    """Add the additional-elements column to legacy templates and previous workbooks."""
    additional_header = "Additional Canonical Link Element(s)"
    if sheet.cell(3, 4).value == additional_header:
        return
    if sheet.cell(3, 4).value != "Instructions":
        sys.exit("Canonical tab has an unsupported column layout; expected Instructions in D3")

    image_positions = [(image, _image_position(image)) for image in getattr(sheet, "_images", [])]
    old_instruction_width = sheet.column_dimensions["D"].width
    old_screenshot_width = sheet.column_dimensions["E"].width
    had_legacy_heading_merge = "B1:D2" in {str(rng) for rng in sheet.merged_cells.ranges}
    if had_legacy_heading_merge:
        sheet.unmerge_cells("B1:D2")
    sheet.insert_cols(4, 1)

    if had_legacy_heading_merge:
        sheet.merge_cells("B1:E2")

    sheet.cell(3, 4)._style = copy.copy(sheet.cell(3, 3)._style)
    sheet.cell(3, 4).value = additional_header
    sheet.cell(DATA_ROW, 4)._style = copy.copy(sheet.cell(DATA_ROW, 3)._style)
    sheet.column_dimensions["D"].width = sheet.column_dimensions["C"].width
    sheet.column_dimensions["E"].width = old_instruction_width
    sheet.column_dimensions["F"].width = old_screenshot_width

    for image, position in image_positions:
        if position and position[1] >= 4:
            image.anchor = f"{get_column_letter(position[1] + 1)}{position[0]}"


def canonical_values(record):
    return [
        record.get("Address"),
        record.get("Issue"),
        record.get("Canonical Link Element"),
        record.get("Additional Canonical Link Elements"),
    ]


def build_canonical_sheet(wb, groups, meta, check_rows, mode):
    sheet = wb[CANONICAL_TAB]
    ensure_canonical_schema(sheet)
    # Unmerge any merged ranges in the data area before rewriting cells. Previous
    # completed workbooks can carry merged cells in the data rows, which would make
    # the value writes below raise on read-only MergedCell objects. Mirrors the same
    # preamble already used by the meta/images and generic section builders.
    for rng in list(sheet.merged_cells.ranges):
        if rng.max_row >= DATA_ROW:
            sheet.unmerge_cells(str(rng))
    max_col = 6
    style = [copy.copy(sheet.cell(DATA_ROW, c)._style) for c in range(1, max_col + 1)]

    previous_keys = []
    previous_missing_keys = []
    row_images = collections.defaultdict(list)
    data_images = []
    for image in list(getattr(sheet, "_images", [])):
        position = _image_position(image)
        if position and position[0] >= DATA_ROW:
            row_images[position[0]].append((position[1], image))
            data_images.append(image)

    carry_forward = collections.defaultdict(collections.deque)
    for row in range(DATA_ROW, sheet.max_row + 1):
        machine = [sheet.cell(row, c).value for c in range(1, 5)]
        if not any(value not in (None, "") for value in machine):
            continue
        issue_key = _issue_key(machine, (1, 2, 3, 4))
        previous_keys.append(issue_key)
        if _norm_key(machine[1]) == "Missing":
            previous_missing_keys.append(issue_key)
        preserved = tuple(_capture_cell(sheet.cell(row, c)) for c in (5, 6))
        carry_forward[issue_key].append((preserved, row_images.get(row, [])))

    if data_images:
        sheet._images = [image for image in sheet._images if image not in data_images]
    for row in sheet.iter_rows(min_row=DATA_ROW, max_row=sheet.max_row, max_col=max_col):
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None
                cell.hyperlink = None
                cell.comment = None

    current_keys = []
    current_missing_keys = []
    output_row = DATA_ROW
    wrote_group = False
    for issue, _, _ in CANONICAL_ISSUES:
        records = groups.get(issue, [])
        if not records:
            continue
        if wrote_group:
            for c in range(1, max_col + 1):
                sheet.cell(output_row, c)._style = copy.copy(style[c - 1])
            output_row += 1
        wrote_group = True
        for record in records:
            values = canonical_values(record)
            issue_key = _issue_key(values, (1, 2, 3, 4))
            current_keys.append(issue_key)
            if issue == "Missing":
                current_missing_keys.append(issue_key)
            for c, value in enumerate(values, start=1):
                cell = sheet.cell(output_row, c)
                cell.value = value
                cell._style = copy.copy(style[c - 1])
            if mode == "review" and carry_forward[issue_key]:
                carried_cells, carried_images = carry_forward[issue_key].popleft()
                for c, captured in zip((5, 6), carried_cells):
                    cell = sheet.cell(output_row, c)
                    _restore_cell(cell, captured)
                    cell._style = copy.copy(style[c - 1])
                for image_col, image in carried_images:
                    image.anchor = f"{get_column_letter(image_col)}{output_row}"
                    sheet.add_image(image)
            output_row += 1

    complete = meta.get("canonicals_available", False)
    missing_count = len(current_missing_keys)
    total_count = len(current_keys)
    missing_mark = "✖" if missing_count else ("✔" if complete else "?")
    wb[resolve_check_tab(wb.sheetnames)].cell(check_rows["13.1"], CHECK_COL).value = missing_mark
    # 13.2 (correct version?) and 13.3-13.7 are human-judgment items: the skill
    # cannot tell whether a canonicalised URL is right or wrong, so it leaves these
    # BLANK for a human and lists them under NEEDS HUMAN REVIEW. The detail sheet
    # still stores every canonical group (including all Canonicalised URLs) as evidence.
    issue_mark = "—"
    if total_count:
        sheet.sheet_state = "visible"

    results = [
        ("13.1", "13. Canonical tag - Missing", missing_count, missing_mark),
        ("13.2", "13. Canonical tag - Issues", total_count, issue_mark),
    ]
    changes = {}
    if mode == "review":
        changes = {
            "13.1": _change_counts(previous_missing_keys, current_missing_keys),
            "13.2": _change_counts(previous_keys, current_keys),
        }
    return results, changes



# =========================================================================
# NEW ITEMS (extension 2026-07-30): HTTPS (§9), Pagination (§14),
# URL best practices (§26 SEO / §30 SEO-GEO), HTTP status codes & broken
# links (§27/§31 + 24.8/28.8), Page speed (§28/§32).
#
# Same crawl, same workbook. Auto-marks ONLY what the crawl / PSI proves;
# judgment sub-items are left blank for the analyst (like 13.3-13.7). A
# report that is absent leaves its item blank (never a false pass).
# =========================================================================

# HTTPS(9) and Pagination(14) share numbers across both families; URL,
# status codes and page speed are renumbered in the SEO/GEO template.
NEW_SECTIONS = {
    "seo":     {"url": "26", "status": "27", "pagespeed": "28", "broken": "24.8"},
    "seo_geo": {"url": "30", "status": "31", "pagespeed": "32", "broken": "28.8"},
}
HTTPS_TAB = "9. Implement HTTPS Protocol"
PAGINATION_TAB = "14. Pagination"   # SEO only; SEO/GEO has the items but no detail tab
PSI_FAIL_BELOW = 50                 # overall performance 0-49 => fail (house rule)

# Review preserves the PreCheck / account block from the previous onsite and
# re-audits everything else fresh (like Begin). Sections 1-6 by number, plus
# the three trailing account sections matched by header text (their numbers
# differ between SEO 30/31/32 and SEO/GEO 34/35/36).
PRECHECK_NUMERIC = {1, 2, 3, 4, 5, 6}
PRECHECK_NAMES = ("conversion tracking", "backlink profile", "crm protocol")


# =========================================================================
# NEEDS HUMAN REVIEW registry. Every item the skill deliberately leaves BLANK
# for a human to judge lives here, one row each: (seo_item, seo_geo_item, why).
# To move a NEW item to human review later, just add a line — it will be left
# blank and appear in the run's NEEDS HUMAN REVIEW summary automatically.
# HTTPS 9.3 (www vs non-www) is auto-checked, so it is NOT listed here.
# =========================================================================
HUMAN_REVIEW = [
    ("10.2", "10.2", "Robots.txt - are sections we WANT indexed being disallowed? (intent check; the Disallow rules are listed in the '10. Robots.txt' sheet)."),
    ("10.3", "10.3", "Robots.txt - are disallow blocks in place for URLs Google shouldn't crawl?"),
    ("10.5", "10.5", "Robots.txt - are cart / thank-you / admin / account / duplicate pages disallowed?"),
    ("11.2", "11.2", "Sitemap - is it a dynamic (auto-updating) sitemap?"),
    ("11.8", "11.8", "Sitemap - e-commerce: does it contain product images where appropriate?"),
    ("11.9", "11.9", "Sitemap - submitted to GSC and confirmed error-free? (GSC Sitemaps report - not in the crawl)."),
    ("13.2", "13.2", "Canonical - are canonicals referencing the correct version? Verify the Canonicalised URLs in the '13. Canonical tag' sheet."),
    ("13.3", "13.3", "Canonical - do tracking / query-string URLs correctly canonical to the base URL?"),
    ("13.4", "13.4", "Canonical - do e-commerce variant URLs canonical to the base product URL?"),
    ("13.5", "13.5", "Canonical - is internal search causing URL duplication?"),
    ("13.6", "13.6", "Canonical - are blog categories / tags / archives duplicating?"),
    ("13.7", "13.7", "Canonical - is Google respecting the canonicals? (GSC 'Duplicate, Google chose...')."),
    ("14.1", "14.1", "Pagination - are there on-page links to page 2 and onward?"),
    ("14.2", "14.2", "Pagination - a 'View All' page plus paginated pages without a correct rel=canonical?"),
    ("14.5", "14.5", "Pagination - has pagination caused content duplication?"),
    ("25.1", "29.1", "Crawling Issues - is the navigation bar clear/organised with links to important pages? (visual check)."),
    ("25.2", "29.2", "Crawling Issues - can Googlebot access the significant pages? (GSC > Crawl Stats > By Response > OK 200 - not in the API)."),
    ("25.3", "29.3", "Crawling Issues - how many pages does Google crawl per day? (GSC > Crawl Stats > Total crawl requests) - enter a NUMBER, not a tick/cross."),
    ("25.4", "29.4", "Crawling Issues - which URLs are wasting crawl budget? (GSC > Crawl Stats - not in the API)."),
    ("26.1", "30.1", "URL best practices - logical hierarchy + keywords in URLs? (over-length URLs are listed as evidence)."),
    ("26.6", "30.6", "URL best practices - are trailing slashes handled consistently?"),
    ("27.1", "31.1", "HTTP status - do error pages return a real 404 (not a soft 404)?"),
    ("27.4", "31.4", "HTTP status - do temporary redirects use 307 rather than 302?"),
    ("27.5", "31.5", "HTTP status - is there a custom 404 page with a CTA and links?"),
    ("28.2", "32.2", "Page speed - GTmetrix / Pingdom score."),
    ("28.4", "32.4", "Page speed - Core Web Vitals (GSC groups)."),
    ("28.5", "32.5", "Page speed - WordPress: does the site need a separate speed audit?"),
]


def human_review_list(family, all_rows):
    """Return [(item, why)] for the human-review items that exist in this workbook."""
    idx = 0 if family == "seo" else 1
    out = []
    for entry in HUMAN_REVIEW:
        item = entry[idx]
        if item in all_rows:
            out.append((item, entry[2]))
    return out


def find_all_rows(ws):
    """Map every checklist item number ('9.1', '26.2', ...) to its sheet row."""
    out = {}
    for r in range(3, ws.max_row + 1):
        b = ws.cell(r, 2).value
        try:
            key = f"{round(float(b), 2):.1f}"
        except (TypeError, ValueError):
            continue
        out.setdefault(key, r)
    return out


def keep_section_rows(ws):
    """Rows belonging to PreCheck sections (whole block: header + its items)."""
    headers = []
    for r in range(3, ws.max_row + 1):
        b = ws.cell(r, 2).value
        try:
            f = float(b)
        except (TypeError, ValueError):
            continue
        if abs(f - round(f)) < 1e-6:
            headers.append((r, int(round(f)), str(ws.cell(r, 3).value or "")))
    keep = set()
    for i, (r, sec, text) in enumerate(headers):
        end = headers[i + 1][0] if i + 1 < len(headers) else ws.max_row + 1
        tl = text.lower()
        if sec in PRECHECK_NUMERIC or any(n in tl for n in PRECHECK_NAMES):
            keep.update(range(r, end))
    return keep


def _is_page1(u):
    u = str(u or "").lower()
    return bool(re.search(r"[?&](page|p)=1(&|$)", u) or re.search(r"/page/1/?$", u))


def _is_paginated(u):
    u = str(u or "").lower()
    return bool(re.search(r"[?&](page|p)=\d+", u) or re.search(r"/page/\d+/?$", u))


def _reg_host(u, bases):
    """Record whether a URL's registrable host appeared with / without a www prefix."""
    from urllib.parse import urlsplit
    h = urlsplit(str(u or "")).netloc.lower()
    if not h:
        return
    base = h[4:] if h.startswith("www.") else h
    flags = bases.setdefault(base, [False, False])
    flags[0 if h.startswith("www.") else 1] = True


def build_new_buckets(d):
    """Parse the additional Screaming Frog exports. Each report is optional;
    `present` records which ones were actually supplied."""
    from urllib.parse import urlsplit

    def rpath(*must, exclude=()):
        return find_csv(d, *must, exclude=exclude)

    def rows(*must, exclude=()):
        return load(find_csv(d, *must, exclude=exclude))

    t_all = load(find_csv(d, "page_titles", "all"))
    hosts = collections.Counter()
    for r in t_all:
        h = urlsplit(str(col(r, "Address") or "")).netloc.lower()
        if h:
            hosts[h] += 1
    base_host = hosts.most_common(1)[0][0] if hosts else ""

    def is_internal(u):
        h = urlsplit(str(u or "")).netloc.lower()
        return h == "" or h == base_host

    http_internal = [r for r in rows("security", "http", exclude=("inlink",))
                     if is_internal(col(r, "Address", "URL"))]
    mixed = rows("security", "mixed")

    resp_4xx = rows("client_error", exclude=("inlink",))
    resp_5xx = rows("server_error", exclude=("inlink",))
    resp_3xx = rows("redirection", exclude=("inlink", "chain"))
    r302 = [r for r in resp_3xx if str(col(r, "Status Code") or "").strip() == "302"]

    broken_internal = []
    for r in rows("client_error", "inlink") + rows("server_error", "inlink"):
        dst = col(r, "Destination", "Address")
        if is_internal(dst):
            broken_internal.append((col(r, "Source"), dst))

    url_param = rows("url", "parameters")
    url_under = rows("url", "underscore")
    url_upper = rows("url", "uppercase")
    url_slash = rows("url", "multiple", "slash")
    url_long = rows("url", "over")

    pag = rows("pagination", exclude=("inlink",))
    page1 = [r for r in t_all if _is_page1(col(r, "Address"))]
    can_all = load(find_csv(d, "canonicals", "all"))
    can_map = {(col(r, "Address") or ""):
               (col(r, "Canonical Link Element 1", "Canonical Link Element") or "")
               for r in can_all}
    paginated = [col(r, "Address") for r in t_all if _is_paginated(col(r, "Address"))]
    pag_canon_mismatch = [a for a in paginated if can_map.get(a) and can_map.get(a) != a]
    pagination_rows = [(col(r, "Address"), "Pagination") for r in pag]

    # 9.3 www vs non-www: fail only when BOTH a www and a non-www variant serve LIVE
    # (200) pages for the same host. A clean redirect (one variant 301s to the other)
    # is correct and must pass, so count only live pages and skip any that are 3xx.
    redirect_urls = {str(col(r, "Address") or "").rstrip("/").lower() for r in resp_3xx}
    www_bases = {}
    for r in t_all:
        addr = col(r, "Address")
        if str(addr or "").rstrip("/").lower() in redirect_urls:
            continue
        _reg_host(addr, www_bases)
    www_conflict = any(w and n for w, n in www_bases.values())

    ps_rows, ps_fail = [], []
    for r in rows("pagespeed", exclude=("inlink",)):
        addr = col(r, "Address", "URL")
        sc = num(col(r, "Performance Score", "Performance"))
        if 0 < sc <= 1:
            sc *= 100
        ps_rows.append([addr, round(sc)])
        if 0 < sc < PSI_FAIL_BELOW:
            ps_fail.append(addr)

    present = {
        "security": rpath("security", "http", exclude=("inlink",)) is not None,
        "mixed": rpath("security", "mixed") is not None,
        "resp_5xx": rpath("server_error", exclude=("inlink",)) is not None,
        "resp_3xx": rpath("redirection", exclude=("inlink", "chain")) is not None,
        "resp_4xx": rpath("client_error", exclude=("inlink",)) is not None,
        "broken": rpath("client_error", "inlink") is not None or rpath("server_error", "inlink") is not None,
        "url_param": rpath("url", "parameters") is not None,
        "url_under": rpath("url", "underscore") is not None,
        "url_upper": rpath("url", "uppercase") is not None,
        "url_slash": rpath("url", "multiple", "slash") is not None,
        "pagination": rpath("pagination", exclude=("inlink",)) is not None,
        "pagespeed": rpath("pagespeed", exclude=("inlink",)) is not None,
    }
    data = {
        "http_urls_internal": http_internal, "mixed": mixed,
        "resp_4xx": resp_4xx, "resp_5xx": resp_5xx, "r302": r302,
        "broken_internal": broken_internal,
        "url_param": url_param, "url_under": url_under, "url_upper": url_upper,
        "url_slash": url_slash, "url_long": url_long,
        "page1": page1, "pagination": pag, "pagination_rows": pagination_rows,
        "pag_canon_mismatch": pag_canon_mismatch, "www_conflict": www_conflict,
        "ps_rows": ps_rows, "ps_fail": ps_fail,
    }
    return data, present


def _clear_tab(sheet):
    for row in sheet.iter_rows(min_row=DATA_ROW, max_row=sheet.max_row, max_col=sheet.max_column):
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None
                cell.hyperlink = None
                cell.comment = None


def _write_rows(wb, tab, data_rows):
    """Overwrite a detail tab's data area (row 4 down) with fresh rows."""
    if tab not in wb.sheetnames:
        return
    sheet = wb[tab]
    # Some template tabs vertically merge a column down the whole data area
    # (e.g. SEO/GEO B4:B33), which makes those cells read-only. Unmerge any
    # range that reaches into the data area so fresh rows can be written.
    for rng in list(sheet.merged_cells.ranges):
        if rng.max_row >= DATA_ROW:
            sheet.unmerge_cells(str(rng))
    _clear_tab(sheet)
    if not data_rows:
        return
    ncol = max(len(r) for r in data_rows)
    style = [copy.copy(sheet.cell(DATA_ROW, c)._style) for c in range(1, ncol + 1)]
    for i, vals in enumerate(data_rows):
        for c, v in enumerate(vals, start=1):
            cell = sheet.cell(DATA_ROW + i, c)
            cell.value = v
            cell._style = copy.copy(style[c - 1])
    sheet.sheet_state = "visible"


def apply_new_items(wb, family, exports_dir, all_rows, ws):
    """Set checklist marks and write detail tabs for the added items. Returns
    (results, notes)."""
    sec = NEW_SECTIONS[family]
    data, present = build_new_buckets(exports_dir)
    results, notes = [], []

    def setmark(key, mark, label, count):
        if mark is None:
            return
        r = all_rows.get(key)
        if r:
            ws.cell(r, CHECK_COL).value = mark
        results.append((f"{key} {label}", "", count, mark))

    def yn(flag, hits):
        return (("✖" if hits else "✔"), len(hits)) if flag else (None, 0)

    # ---- 9 HTTPS ----------------------------------------------------------
    m, n = yn(present["security"], data["http_urls_internal"]); setmark("9.1", m, "HTTPS - insecure/HTTP URLs", n)
    m, n = yn(present["mixed"], data["mixed"]);                 setmark("9.2", m, "HTTPS - mixed content", n)
    # 9.3 www vs non-www is a checklist mark only (no detail sheet): fail if the
    # crawl saw both a www and a non-www URL for the same host.
    setmark("9.3", ("✖" if data["www_conflict"] else "✔"), "HTTPS - www vs non-www",
            1 if data["www_conflict"] else 0)
    # Detail sheet (anchor text / screenshots are left for the human):
    #   http vs https -> the insecure URL goes in Source, no Destination.
    #   mixed content -> the HTTPS page in Source, the HTTP resource in Destination.
    https_rows = [[col(r, "Source"), col(r, "Destination", "Address"), "", "Mixed Content"]
                  for r in data["mixed"]]
    https_rows += [[col(r, "Address", "URL"), "", "", "HTTP (insecure) URL"]
                   for r in data["http_urls_internal"]]
    _write_rows(wb, HTTPS_TAB, https_rows)

    # ---- 14 Pagination ----------------------------------------------------
    setmark("14.4", ("✖" if data["page1"] else "✔"), "Pagination - page=1 exists", len(data["page1"]))
    if present["pagination"]:
        setmark("14.3", ("✔" if data["pagination"] else "✖"), "Pagination - rel=prev/next", len(data["pagination"]))
        setmark("14.6", ("✖" if data["pag_canon_mismatch"] else "✔"),
                "Pagination - canonical self-ref", len(data["pag_canon_mismatch"]))
    _write_rows(wb, PAGINATION_TAB, [[a, iss] for a, iss in data["pagination_rows"]])

    # ---- 26/30 URL best practices ----------------------------------------
    U = sec["url"]
    for suf, label, dk in (("2", "Parameters", "url_param"), ("3", "Underscores", "url_under"),
                           ("4", "Uppercase", "url_upper"), ("5", "Multiple Slashes", "url_slash")):
        m, n = yn(present[dk], data[dk]); setmark(f"{U}.{suf}", m, f"URL - {label}", n)
        _write_rows(wb, f"{U}.{suf} {label}", [[_addr(r)] for r in data[dk]])
    _write_rows(wb, f"{U}.1 URL best practices",
                [[_addr(r), f"Over {len(str(_addr(r)))} characters"] for r in data["url_long"]])

    # ---- 27/31 HTTP status codes & broken links --------------------------
    S = sec["status"]
    m, n = yn(present["resp_5xx"], data["resp_5xx"]); setmark(f"{S}.2", m, "5xx server errors", n)
    _write_rows(wb, f"{S}.2 5xx", [[_addr(r)] for r in data["resp_5xx"]])
    m, n = yn(present["resp_3xx"], data["r302"]); setmark(f"{S}.3", m, "302 redirects", n)
    _write_rows(wb, f"{S}.3 302s", [[col(r, "Address"), col(r, "Redirect URL", "Destination")] for r in data["r302"]])
    _write_rows(wb, f"{S}.1 404s", [[_addr(r), ""] for r in data["resp_4xx"]])
    bkey = sec["broken"]
    m, n = yn(present["broken"], data["broken_internal"]); setmark(bkey, m, "Broken internal links", n)
    _write_rows(wb, f"{bkey} Broken Internal Links", [[s_, d_] for s_, d_ in data["broken_internal"]])

    # ---- 28/32 Page speed (PSI-gated) ------------------------------------
    P = sec["pagespeed"]
    if present["pagespeed"] and data["ps_rows"]:
        mark = "✖" if data["ps_fail"] else "✔"
        setmark(f"{P}.1", mark, "Page speed - PSI mobile", len(data["ps_fail"]))
        setmark(f"{P}.3", mark, "Page speed - site speed", len(data["ps_fail"]))
        _write_rows(wb, f"{P}. Page speed", data["ps_rows"])
    else:
        notes.append(f"Page speed not evaluated ({P}.1/{P}.3 left blank): no PageSpeed Insights "
                     "data in the crawl. Re-run with SF > Config > API Access > PageSpeed Insights.")

    for label, flag in (("Security", present["security"]), ("Response Codes", present["resp_5xx"]),
                        ("URL filters", present["url_param"])):
        if not flag:
            notes.append(f"{label} export absent — related items left blank for the analyst.")
    return results, notes



# =========================================================================
# Robots.txt (§10) + Sitemap (§11) checks. These read the live robots.txt and
# sitemap files (URLs supplied by the user, since the runtime may not discover
# them) and cross-reference the crawl. Items that need intent/judgment stay in
# HUMAN_REVIEW; the file-checkable ones are auto-marked here. If a file is not
# supplied or cannot be read, its auto-items fall back to human review.
# =========================================================================
_ROBOTS_UA = "Mozilla/5.0 (compatible; FirstPage-Onsite/1.0)"


def _fetch_text(url_or_path, timeout=30):
    """Fetch a URL (browser UA) or read a local file. Returns text or None."""
    if not url_or_path:
        return None
    if os.path.exists(url_or_path):
        try:
            with open(url_or_path, encoding="utf-8", errors="replace") as fh:
                return fh.read()
        except OSError:
            return None
    try:
        req = urllib.request.Request(url_or_path, headers={"User-Agent": _ROBOTS_UA})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.read().decode("utf-8", "replace")
    except Exception:
        return None


def _parse_robots(text):
    sitemaps, disallows = [], []
    for line in (text or "").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        field, _, value = line.partition(":")
        field, value = field.strip().lower(), value.strip()
        if field == "sitemap":
            sitemaps.append(value)
        elif field == "disallow" and value:
            disallows.append(value)
    blocks_assets = any(re.search(r"\.(css|js)(\?|$)", d, re.I)
                        or re.search(r"/(wp-includes|assets|static|css|js|scripts|styles)/?$", d, re.I)
                        for d in disallows)
    return {"sitemaps": sitemaps, "disallows": disallows, "blocks_assets": blocks_assets}


def _path_disallowed(url, disallows):
    from urllib.parse import urlsplit
    path = urlsplit(str(url or "")).path or "/"
    return any(d and path.startswith(d) for d in disallows)


def _collect_sitemap_urls(sitemap_url, _depth=0, _seen=None):
    """Fetch a sitemap, following one level of index -> child sitemaps.
    Returns (loc_urls, has_lastmod, has_images, child_sitemaps, ok)."""
    _seen = _seen if _seen is not None else set()
    text = _fetch_text(sitemap_url)
    if text is None:
        return [], False, False, [], False
    has_lastmod = "<lastmod" in text
    has_images = "<image:" in text
    is_index = "<sitemapindex" in text.lower()
    found = [m.group(1).strip() for m in re.finditer(r"<loc>\s*([^<\s]+)\s*</loc>", text)]
    locs, children = ([], found) if is_index else (found, [])
    if is_index and _depth == 0:
        for child in children[:100]:
            if child in _seen:
                continue
            _seen.add(child)
            c_locs, c_lm, c_img, _, _ = _collect_sitemap_urls(child, _depth + 1, _seen)
            locs.extend(c_locs)
            has_lastmod = has_lastmod or c_lm
            has_images = has_images or c_img
    return locs, has_lastmod, has_images, children, True


def check_robots_sitemap(exports_dir, robots_url, sitemap_url):
    """Return (marks{item:mark}, detail{tab:rows}, needs[(item,why)], notes[])."""
    from urllib.parse import urlsplit
    marks, detail, needs, notes = {}, {}, [], []

    def norm(u):
        return str(u or "").rstrip("/").lower()

    # crawl cross-reference sets (empty if no exports)
    non_indexable, non200, canonicalised = set(), set(), set()
    if exports_dir:
        for r in load(find_csv(exports_dir, "page_titles", "all")):
            ix = (r.get("Indexability") or "").strip()
            if ix and ix != "Indexable":
                non_indexable.add(norm(col(r, "Address")))
        for grp in ("client_error", "server_error", "redirection"):
            for r in load(find_csv(exports_dir, grp, exclude=("inlink", "chain"))):
                non200.add(norm(col(r, "Address")))
        for r in load(find_csv(exports_dir, "canonicals", "all")):
            canon = col(r, "Canonical Link Element 1", "Canonical Link Element")
            if canon and norm(canon) != norm(col(r, "Address")):
                canonicalised.add(norm(col(r, "Address")))

    # ---- robots.txt (§10 auto items) ----
    robots = _fetch_text(robots_url) if robots_url else None
    rp = _parse_robots(robots) if robots else {"disallows": [], "sitemaps": [], "blocks_assets": False}
    if robots_url and robots is not None:
        marks["10.1"] = "✔"
        marks["10.4"] = "✖" if rp["blocks_assets"] else "✔"
        marks["10.6"] = "✔" if robots_url.strip().lower().endswith("/robots.txt") else "✖"
        marks["10.7"] = "✔" if rp["sitemaps"] else "✖"
        # §10 detail: a header row, then the ENTIRE current robots.txt pasted
        # into ONE cell (not one row per line), matching the approved template.
        robots_block = robots.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
        detail["10. Robots.txt"] = [
            ["Current Robots.txt", "Issue", "Instruction", "Remarks"],
            [robots_block],
        ]
        notes.append(f"Robots.txt: {len(rp['disallows'])} Disallow rule(s); "
                     f"{'sitemap linked' if rp['sitemaps'] else 'NO sitemap link'}; "
                     f"CSS/JS {'BLOCKED' if rp['blocks_assets'] else 'accessible'}.")
    else:
        if robots_url:
            notes.append(f"Could not read robots.txt ({robots_url}); attach the file if the server blocked the fetch.")
        for it in ("10.1", "10.4", "10.6", "10.7"):
            needs.append((it, "Robots.txt not supplied/readable - provide the robots.txt URL to auto-check."))

    # ---- sitemap (§11 auto items) ----
    if sitemap_url:
        locs, has_lastmod, _has_img, children, ok = _collect_sitemap_urls(sitemap_url)
    else:
        locs, has_lastmod, children, ok = [], False, [], False
    if sitemap_url and ok:
        hosts = {urlsplit(u).netloc.lower() for u in locs if u}
        schemes = {urlsplit(u).scheme.lower() for u in locs if u}
        bad_55 = [u for u in locs if _path_disallowed(u, rp["disallows"]) or norm(u) in non_indexable]
        bad_56 = [u for u in locs if _is_paginated(u)]
        bad_57 = [u for u in locs if norm(u) in non200 or norm(u) in canonicalised]
        marks["11.1"] = "✔"
        marks["11.3"] = "✔" if (locs and len(hosts) <= 1 and schemes <= {"https"}) else ("✖" if locs else "?")
        marks["11.4"] = "✔" if has_lastmod else "✖"
        marks["11.5"] = "✖" if bad_55 else "✔"
        marks["11.6"] = "✖" if bad_56 else "✔"
        marks["11.7"] = "✖" if bad_57 else "✔"
        rows = [["Sitemap index", sitemap_url]] + [["Child sitemap", c] for c in children]
        for lbl, bad in (("Blocked/noindex URL in sitemap (11.5)", bad_55),
                         ("Paginated URL in sitemap (11.6)", bad_56),
                         ("Non-200 / non-canonical URL in sitemap (11.7)", bad_57)):
            rows += [[lbl, u] for u in bad[:20]]
        detail["11. Sitemap"] = rows
        notes.append(f"Sitemap: {len(locs)} URL(s) across {len(children) or 1} sitemap(s); "
                     f"lastmod {'present' if has_lastmod else 'MISSING'}.")
    else:
        if sitemap_url:
            notes.append(f"Could not read sitemap ({sitemap_url}); attach the file if the server blocked the fetch.")
        for it in ("11.1", "11.3", "11.4", "11.5", "11.6", "11.7"):
            needs.append((it, "Sitemap not supplied/readable - provide the sitemap URL to auto-check."))
    return marks, detail, needs, notes


def build(base_workbook, out_path, buckets, meta=None, mode="begin",
          blank_unverified=True, keep_prechecks=False, family=None, exports_dir=None,
          robots_url=None, sitemap_url=None, record_previous=False):
    meta = meta or {}
    shutil.copy(base_workbook, out_path)
    wb = openpyxl.load_workbook(out_path)

    missing_tabs = [TABS[k]["tab"] for k in TABS if TABS[k]["tab"] not in wb.sheetnames]
    if CANONICAL_TAB not in wb.sheetnames:
        missing_tabs.insert(0, CANONICAL_TAB)
    if missing_tabs:
        sys.exit(f"Template is missing expected tabs: {missing_tabs}")

    ws = wb[resolve_check_tab(wb.sheetnames)]
    rows = find_check_rows(ws)
    missing_check_rows = sorted(MANAGED_CHECK_ITEMS - set(rows))
    if missing_check_rows:
        sys.exit(f"Could not locate checklist rows: {missing_check_rows}")
    all_rows = find_all_rows(ws)
    keep_rows = keep_section_rows(ws) if keep_prechecks else set()

    # Review only: roll the previous audit's result (column D) into column E for the
    # record before this run overwrites D. Checklist tab only; detail tabs unchanged.
    if record_previous:
        record_previous_check_column(ws, base_workbook, out_path)

    # The template ships '✔' pre-filled on every item. Left alone, the deliverable
    # would claim a clean pass on checks nobody ran.
    if blank_unverified or keep_rows:
        for r in range(3, ws.max_row + 1):
            if r in keep_rows:
                continue
            if ws.cell(r, CHECK_COL).value in ("✔", "✖", "?"):
                ws.cell(r, CHECK_COL).value = None

    results, changes = build_canonical_sheet(
        wb, buckets.get("13", {}), meta, rows, mode,
    )
    for key, spec in TABS.items():
        tab = spec["tab"]
        rowfn = spec["rowfn"]
        machine_cols = spec["machine_cols"]
        preserve_cols = spec["preserve_cols"]
        key_cols = spec["key_cols"]
        data = buckets.get(key, [])
        sheet = wb[tab]
        max_col = max((sheet.max_column, machine_cols, *preserve_cols))
        style = [copy.copy(sheet.cell(DATA_ROW, c)._style) for c in range(1, max_col + 1)]

        previous_keys = []
        row_images = collections.defaultdict(list)
        data_images = []
        for image in list(getattr(sheet, "_images", [])):
            position = _image_position(image)
            if position and position[0] >= DATA_ROW:
                row_images[position[0]].append((position[1], image))
                data_images.append(image)

        carry_forward = collections.defaultdict(collections.deque)
        for row in range(DATA_ROW, sheet.max_row + 1):
            machine = [sheet.cell(row, c).value for c in range(1, machine_cols + 1)]
            if not any(v not in (None, "") for v in machine):
                continue
            issue_key = _issue_key(machine, key_cols)
            previous_keys.append(issue_key)
            preserved_cells = tuple(_capture_cell(sheet.cell(row, c)) for c in preserve_cols)
            carry_forward[issue_key].append((preserved_cells, row_images.get(row, [])))

        if data_images:
            sheet._images = [image for image in sheet._images if image not in data_images]

        # Both modes replace the factual issue rows. Review mode first captured annotations
        # above; Begin mode also clears any accidental example data in a blank template.
        for row in sheet.iter_rows(min_row=DATA_ROW, max_row=sheet.max_row, max_col=max_col):
            for cell in row:
                if not isinstance(cell, MergedCell):
                    cell.value = None
                    cell.hyperlink = None
                    cell.comment = None

        current_keys = []
        for i, rec in enumerate(data):
            values = rowfn(rec)
            issue_key = _issue_key(values, key_cols)
            current_keys.append(issue_key)
            for c, val in enumerate(values, start=1):
                cell = sheet.cell(DATA_ROW + i, c)
                cell.value = val
                cell._style = copy.copy(style[c - 1])
            if mode == "review" and carry_forward[issue_key]:
                carried_cells, carried_images = carry_forward[issue_key].popleft()
                for c, captured in zip(preserve_cols, carried_cells):
                    cell = sheet.cell(DATA_ROW + i, c)
                    _restore_cell(cell, captured)
                    cell._style = copy.copy(style[c - 1])
                for image_col, image in carried_images:
                    image.anchor = f"{get_column_letter(image_col)}{DATA_ROW + i}"
                    sheet.add_image(image)

        if mode == "review":
            changes[key] = _change_counts(previous_keys, current_keys)
        if data:
            sheet.sheet_state = "visible"
        # Image items can't be judged clean when the evidence isn't there:
        #   - no images captured at all -> both 20.1 and 20.2 unknown
        #   - images captured but no size data (external CDN) -> 20.2 can't confirm "no
        #     oversize", so a 0 result is unknown, not clean (20.1 is still valid)
        no_images = key in ("20.1", "20.2") and not meta.get("images_crawled", True)
        no_sizes = (key == "20.2" and not data
                    and meta.get("images_crawled", True)
                    and not meta.get("images_sizes_known", True))
        if no_images or no_sizes:
            mark = "?"
        else:
            mark = "✖" if data else "✔"
        ws.cell(rows[key], CHECK_COL).value = mark
        results.append((key, tab, len(data), mark))

    if family and exports_dir:
        new_results, new_notes = apply_new_items(wb, family, exports_dir, all_rows, ws)
        results.extend(new_results)
        meta.setdefault("notes", []).extend(new_notes)
    if family:
        rs_marks, rs_detail, rs_needs, rs_notes = check_robots_sitemap(exports_dir, robots_url, sitemap_url)
        for item, mark in rs_marks.items():
            r = all_rows.get(item)
            if r:
                ws.cell(r, CHECK_COL).value = mark
            results.append((item, "robots/sitemap", 0, mark))
        for tab, rows in rs_detail.items():
            _write_rows(wb, tab, rows)
            if tab == "10. Robots.txt" and tab in wb.sheetnames:
                # The whole robots.txt sits in one cell under the "Current
                # Robots.txt" header — wrap it and size the row so the full
                # block is visible (matches the approved template).
                cell = wb[tab].cell(DATA_ROW + 1, 1)
                if cell.value:
                    cell.alignment = openpyxl.styles.Alignment(
                        wrap_text=True, vertical="top")
                    nlines = str(cell.value).count("\n") + 1
                    wb[tab].row_dimensions[DATA_ROW + 1].height = max(60, nlines * 15)
        meta.setdefault("notes", []).extend(rs_notes)
        meta["needs_review"] = human_review_list(family, all_rows) + rs_needs
    wb.save(out_path)
    return results, changes


# --------------------------------------------------------------------------- main

def default_template(checklist_type):
    """Return the canonical bundled blank template for a checklist family."""
    d = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "templates",
    )
    return os.path.join(d, PROFILES[checklist_type]["template"])


def main():
    p = argparse.ArgumentParser(
        description="Onsite checklist extraction (items 13.1-13.2 and 17.1-20.2)"
    )
    p.add_argument("--url", help="Start URL to crawl in Spider mode")
    p.add_argument("--sitemap", help="Sitemap URL for route selection only. The local headless "
                                      "route cannot inject it into Crawl These Sitemaps; use a "
                                      "saved crawl or MCP for a full sitemap audit.")
    p.add_argument("--client", required=True, help="Client name, used in the output filename")
    p.add_argument("--checklist-type", choices=sorted(PROFILES),
                   help="Checklist family: seo or seo_geo. Auto-detected from an attached base workbook.")
    p.add_argument("--mode", choices=("begin", "review"), default="begin",
                   help="begin starts from a blank template; review uses a previous onsite or blank template")
    p.add_argument("--review-base", choices=("previous", "blank"),
                   help="Review starting point: previous onsite or a fresh blank template")
    p.add_argument("--template", help="Blank checklist .xlsx for Begin or blank-template Review")
    p.add_argument("--template-url", help="Approved FirstPage GitHub template URL for a blank workbook")
    p.add_argument("--previous-workbook", help="Previous completed onsite .xlsx for Review mode")
    p.add_argument("--out-dir", default=".", help="Where to write the deliverable")
    p.add_argument("--config", help="Path to a .seospiderconfig (else SF defaults)")
    p.add_argument("--crawl-file", help="Export from a saved .dbseospider crawl instead of "
                                        "crawling. Uses the settings that crawl was run with.")
    p.add_argument("--exports-dir", help="Reuse existing CSV exports; skips crawling")
    p.add_argument("--robots-url", help="URL (or local path) of the site's robots.txt for the §10 checks")
    p.add_argument("--sitemap-url", help="URL (or local path) of the sitemap / sitemap index for the §11 checks")
    p.add_argument("--date", help="YYYYMMDD for the filename (default: today)")
    p.add_argument("--keep-template-ticks", action="store_true",
                   help="Do NOT blank the template's default ticks on unverified items")
    a = p.parse_args()

    if not (a.exports_dir or a.crawl_file or a.url or a.sitemap):
        p.error("give --url, --crawl-file, or --exports-dir")
    if a.sitemap and not (a.crawl_file or a.exports_dir):
        p.error("the local headless route does not support sitemap audits. Use a saved crawl "
                "configured with Crawl These Sitemaps, or use the Screaming Frog MCP route")

    if a.mode == "begin" and a.review_base:
        p.error("--review-base is only valid in Review mode")
    if a.mode == "begin" and a.previous_workbook:
        p.error("--previous-workbook is only valid in Review mode")

    if a.mode == "review":
        if a.review_base == "blank":
            p.error("Review from a blank template has been removed. Review now always builds from "
                    "the previous completed onsite; pass --previous-workbook.")
        if not a.previous_workbook:
            p.error("Review mode requires --previous-workbook.")
        if a.template or a.template_url:
            p.error("Review builds from --previous-workbook and cannot also use --template/--template-url.")

    uses_blank_template = a.mode == "begin"
    if uses_blank_template and not (a.template or a.checklist_type):
        p.error("A blank-template run requires --checklist-type when --template is not supplied")

    template_tmp = None
    if a.mode == "review":
        base_workbook = a.previous_workbook
    elif a.template:
        base_workbook = a.template
    else:
        template_tmp = tempfile.TemporaryDirectory(prefix="onsite-template-")
        base_workbook = resolve_blank_template(
            None, a.template_url, a.checklist_type, template_tmp.name,
        )
    if not os.path.exists(base_workbook):
        sys.exit(f"Base workbook not found: {base_workbook}")

    detected_type = detect_profile(base_workbook)
    if a.checklist_type and a.checklist_type != detected_type:
        sys.exit(f"Checklist type mismatch: selected {a.checklist_type}, but the workbook is {detected_type}")
    checklist_type = a.checklist_type or detected_type

    date = a.date or datetime.date.today().strftime("%Y%m%d")

    if a.exports_dir:
        exports = os.path.abspath(a.exports_dir)
        print(f"[reuse] exports <- {exports}")
    else:
        exports = run_crawl(a.url, a.sitemap,
                            os.path.join(os.path.abspath(a.out_dir), "_sf_exports"),
                            a.config, client=a.client, date=date, crawl_file=a.crawl_file)

    buckets, meta = build_buckets(exports)
    os.makedirs(os.path.abspath(a.out_dir), exist_ok=True)
    filename = PROFILES[checklist_type]["output"].format(client=a.client, date=date)
    out_path = os.path.join(os.path.abspath(a.out_dir), filename)
    if os.path.abspath(base_workbook) == os.path.abspath(out_path):
        sys.exit("Refusing to overwrite the source workbook. Choose a different output name, "
                 "date, or directory.")
    results, changes = build(
        base_workbook, out_path, buckets, meta, mode="begin",
        blank_unverified=(uses_blank_template and not a.keep_template_ticks),
        keep_prechecks=(a.mode == "review"),
        family=checklist_type, exports_dir=exports,
        robots_url=a.robots_url, sitemap_url=a.sitemap_url,
        record_previous=(a.mode == "review"),
    )

    print()
    print(f"{'ITEM':<36}{'ROWS':>6}   MARK")
    print("-" * 52)
    for key, tab, n, mark in results:
        print(f"{tab:<36}{n:>6}   {mark}")
    if a.mode == "review":
        print()
        print("Review: preserved PreCheck/account sections (1-6, conversion tracking, backlink "
              "profile, CRM) from the previous onsite; every other section was re-audited fresh.")
    for note in meta.get("notes", []):
        print("  - " + note)
    review_items = meta.get("needs_review", [])
    if review_items:
        print()
        print("NEEDS HUMAN REVIEW - left blank for a human to check:")
        for item, why in review_items:
            print(f"  - {item}: {why}")
    print()
    print("WROTE:", out_path)
    print(f"thresholds: title>{TITLE_PX}px  desc>{DESC_PX}px  image>{IMG_KB}kB")
    if not meta.get("canonicals_available", False):
        missing = ", ".join(meta.get("canonical_missing_reports", [])) or "All (no data)"
        print()
        print("!! WARNING: canonical exports are incomplete. Items 13.1/13.2 are marked '?' "
              "unless a known issue already proves a failure.\n"
              f"   Missing or unusable reports: {missing}")
    if not meta.get("images_crawled", True):
        print()
        print("!! WARNING: the crawl captured NO images. Items 20.1/20.2 are marked '?' "
              "(unknown), NOT clean.\n"
              "   Likely image crawling was off, or JS rendering was off on a site that "
              "loads images via\n"
              "   JavaScript. Re-crawl with images enabled to audit 20.1/20.2.")
    elif not meta.get("images_sizes_known", True):
        print()
        print("!! NOTE: images were found but none had size data - they are served from an "
              "external CDN\n"
              "   (e.g. Prismic/Cloudinary), which SF lists outside Images:All. 20.1 (missing "
              "alt) is valid,\n"
              "   but 20.2 (oversize) can't be judged and is marked '?'. To size external "
              "images, enable\n"
              "   'Check Images' / external resource crawling and re-crawl.")
    if template_tmp:
        template_tmp.cleanup()


if __name__ == "__main__":
    main()
