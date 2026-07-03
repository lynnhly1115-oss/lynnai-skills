# Agent Guide

This repository is written for Codex, but any capable coding agent can use it.
Treat each `SKILL.md` file as the operating manual for that workflow.

## Routing

Use `duanju-schedule-table` when the user asks for:

- 顺场表
- 顺畅表
- 拍摄顺序表
- short-drama scene ordering
- role columns, day/night, interior/exterior, group extras, costume, prop, or
  note extraction from a script

Use `screen-art-department-breakdown` when the user asks for:

- 美术拆解
- 美术指导
- 分场表
- 场景表
- 置景道具执行表
- set dressing, hand props, purchase, fabrication, rental, borrowing, vehicles,
  or special props

## How To Run A Skill Without Native Codex Support

1. Clone this repository.
2. Read the relevant `SKILL.md` completely before touching files.
3. Use the helper script in that skill's `scripts/` directory.
4. Generate a draft workbook and, when useful, a JSON scene dump.
5. Review the workbook against the required checks in `SKILL.md`.
6. Return the final `.xlsx` path to the user.

Example:

```bash
python3 duanju-schedule-table/scripts/shunchangbiao_builder.py script.docx \
  --out-dir output \
  --dump-json output/scenes.json
```

## Dependencies

The helper scripts expect Python 3 with these packages available:

- `python-docx`
- `openpyxl`
- `pdfplumber` for text-based PDF input

If the default Python is missing packages, use the host agent's bundled
workspace Python or create a local virtual environment outside this repository.

## Quality Rules

- Do not trust the first draft blindly. Scripts are helpers, not the final
  production judgment.
- Check scene count, scene IDs, and scene order against the source script.
- Include malformed headings, flashbacks, montage inserts, phone-screen inserts,
  and extra scenes instead of silently dropping them.
- Keep generated workbooks, JSON dumps, screenshots, and temporary files out of
  this repository unless they are intentional fixtures.
- Do not send files to WeChat, Slack, email, or any chat app unless the user
  explicitly asks in the current task and gives the destination.

## Output Contract

For every completed run, report:

- source script name
- generated workbook path
- scene count
- any assumptions or source-number anomalies
- whether validation checks passed

If a file cannot be parsed because it is scanned/image-only, ask the user for a
Word export, text export, or OCR-ready copy.
