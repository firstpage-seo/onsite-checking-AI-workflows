# Onsite Checking AI Workflows

This repository is the shared home for First Page onsite-checking skills. Each skill is an
independent package under `.claude/skills/` so it can be used locally or imported into Multica
without depending on sibling skills.

## Skills

| Skill | Purpose |
|---|---|
| `onsite-meta-images-extraction-multica` | Builds or refreshes checklist items 17.1-20.2 from Screaming Frog data. Supports SEO and SEO/GEO, Begin and Review, and previous-workbook or blank-template Review. |

## Multica import

Import each skill from its own directory URL, not from the repository root:

```text
https://github.com/firstpage-seo/onsite-checking-AI-workflows/tree/main/.claude/skills/onsite-meta-images-extraction-multica
```

Copy the matching agent instructions from:

```text
.claude/skills/onsite-meta-images-extraction-multica/references/multica-agent-instructions.md
```

## Repository layout

```text
.claude/
├── SKILL_TEMPLATE.md
└── skills/
    └── <skill-name>/
        ├── SKILL.md
        ├── scripts/
        ├── references/
        ├── assets/
        ├── tests/
        └── requirements.txt
```

Keep every skill self-contained. Multica may import only the selected skill directory, so a
skill must not rely on scripts or runtime assets stored elsewhere in this repository.

## Adding a skill

1. Copy `.claude/SKILL_TEMPLATE.md` into `.claude/skills/<skill-name>/SKILL.md`.
2. Use lowercase hyphenated names for skill folders and frontmatter names.
3. Keep executable logic in `scripts/`, runtime files in `assets/`, and optional detail in
   `references/`.
4. Add tests inside the skill and document the new skill in the table above.
5. Import and test the specific skill-directory URL in Multica before team rollout.

Generated client deliverables and crawl files do not belong in this repository.
