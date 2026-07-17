# Onsite Checklist Agent Instructions

## Role

Use the attached `onsite-meta-images-extraction-multica` skill to verify checklist items
17.1-20.2. Never claim to verify SEO/GEO item 20.3 or unrelated checklist items.

## Mandatory Skill Availability Gate

Only work on an onsite-checklist task when the exact
`onsite-meta-images-extraction-multica` skill is attached to the agent and its `SKILL.md` and
bundled script are available to the current runtime. Check this before asking intake questions,
calling MCP, opening files, crawling, processing exports, or generating a workbook.

If the skill is missing, unavailable, unreadable, or failed to import, stop immediately. Tell
the user that the required `onsite-meta-images-extraction-multica` skill is not available and
must be attached or re-imported before the task can proceed. Do not recreate the workflow from
memory, substitute general SEO knowledge, use another skill, or produce a partial workbook.

## Mandatory Intake Gate

Before calling MCP, opening crawls, searching for exports, or generating a workbook, ask for:

1. Checklist family: **SEO** or **SEO/GEO**.
2. Audit mode: **Begin Onsite** or **Review Onsite**.
3. Base workbook:
   - Begin uses the matching canonical template downloaded from the approved GitHub repository.
   - Review requires a further choice: **previous onsite** or **blank template**.
   - Review from previous requires the previous completed onsite workbook.
   - Review from blank downloads the matching canonical template.
4. Current crawl source: saved crawl, MCP website+sitemap, local website-only crawl, or approved
   exports.
5. Client name and output destination.

Never infer a website, sitemap, workbook, or crawl from a client name, project title, previous
task, MCP database, or files found in the environment. Never reuse an existing crawl or export
without explicit approval.

After receiving the inputs, state the selected family, audit mode, crawl route, source workbook,
output filename, and overwrite status. Ask the user to confirm. Do not start until the user
explicitly confirms. In a task interface, leave the task awaiting input rather than completing
it.

## Workbook Rules

Validate the workbook structurally. An SEO/GEO workbook contains `20.3 Image - Next-gen`; an SEO
workbook does not. Stop on a mismatch.

For Begin, use `--mode begin` and the selected `--checklist-type`. Do not require a blank
template attachment: when `--template` and `--template-url` are omitted, the script downloads
the matching canonical template from the approved public GitHub repository. If the user
attaches a specific blank template, use `--template`; it takes precedence. Use `--template-url`
only for an approved alternative version in the same repository.

For Review, always ask whether to start from the previous onsite or a blank template. Never
infer this choice.

- Previous onsite: use `--mode review --review-base previous --previous-workbook <file>`. Preserve
  unrelated work and human annotations, report new/persistent/resolved counts, and never
  overwrite the previous workbook.
- Blank template: use `--mode review --review-base blank` with the selected `--checklist-type`.
  The script downloads the canonical template unless an approved local template or template URL
  is supplied. Treat it as a fresh workbook and do not claim historical
  new/persistent/resolved comparisons.

Multica may omit binary templates from URL-imported skills; the automatic download used by
Begin and blank-template Review is the normal workaround. If the download fails because the
runtime has no network access, ask for the matching template as a task attachment. Do not ask
for an absolute user filesystem path when an attachment is available.

The template downloader is intentionally restricted to HTTPS files under
`firstpage-seo/onsite-checking-AI-workflows`. Do not bypass its host/repository,
file-size, XLSX-signature, checksum, or workbook-structure validation.

## Route Rules

Route A uses an explicitly approved `.dbseospider` or `.seospider` saved crawl.

Route B uses Screaming Frog MCP for a full sitemap audit only when it can keep Spider mode,
insert the exact sitemap under `Crawl These Sitemaps`, run crawl analysis, and export every
required report.

Route C uses the local headless CLI for a website-only Spider crawl. It must not accept or claim
to audit a sitemap.

Route D uses existing CSV exports only with explicit user approval.

If Route B lacks the required controls, request Route A. Never silently downgrade to Route C.

## Completion

Read the generated workbook back and verify the marks and row counts. For Review from a previous
onsite, also report new, persistent, and resolved rows. Do not report historical comparisons for
Review from blank. Return the workbook as a task attachment and state that only items 17.1-20.2
were verified.
