# Contributing

Use this checklist when adding or updating a skill.

## Skill Folders

- Put each skill in a top-level folder named exactly like the skill.
- Use lowercase letters, digits, and hyphens only.
- Keep `SKILL.md` as the only required instruction file inside the skill.
- Do not add README, changelog, install guide, or other repo-style docs inside a
  skill folder.
- Put reusable code in `scripts/`.
- Put optional UI metadata in `agents/openai.yaml`.

## SKILL.md

- Include only `name` and `description` in YAML frontmatter.
- Make `description` explicit about what the skill does and when it should be
  triggered.
- Keep body instructions concise and procedural.
- Move long references into `references/` only when the skill truly needs them.

## Before Commit

Run:

```bash
python3 scripts/validate_skills.py . --compile-scripts
```

Also run at least one realistic manual test for any script or workflow you
changed.

