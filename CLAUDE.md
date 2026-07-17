# Onsite Checking Workflows - Agent Guide

This repository contains independent onsite-checking skills under `.claude/skills/`. Route a
task using each skill's frontmatter `description`; that metadata is the source of truth.

## Required routing behavior

- Only perform a specialized onsite workflow when its exact skill exists and is readable.
- If the required skill is absent or failed to load, stop and identify the missing skill.
- Do not reconstruct a missing workflow from memory or substitute general SEO knowledge.
- Do not infer client URLs, sitemaps, crawls, templates, or prior workbooks from nearby files.
- Follow the selected skill's intake and confirmation gate before using tools or writing output.
- Write client outputs outside this repository unless the user explicitly chooses another safe
  destination.

## Available skills

| Skill | Trigger |
|---|---|
| `onsite-canonical-meta-images-extraction-multica` | Canonical, meta tags, and images portion of the SEO or SEO/GEO Onsite Checklist, items 13.1-13.2 and 17.1-20.2. |

## Adding skills

Start from `.claude/SKILL_TEMPLATE.md`. Skills are independent by default and must contain all
scripts, references, assets, dependencies, and tests needed by their own workflow. Only add a
cross-skill dependency when the workflow explicitly requires it and the dependency is available
in every target runtime.
