# Hermes Usage

Hermes does not need native Codex skill support to use this repository.
Use the repository as a set of executable instructions plus helper scripts.

## Quick Prompt For Hermes

```text
Clone or open https://github.com/lynnhly1115-oss/lynnai-skills.
Read AGENTS.md first.
For a short-drama 顺场表, read duanju-schedule-table/SKILL.md completely and run
duanju-schedule-table/scripts/shunchangbiao_builder.py.
For an AI真人剧资产统计表, read ai-live-action-asset-breakdown/SKILL.md and its
inventory schema completely, then run build_asset_workbook.mjs with Node.js and
@oai/artifact-tool.
For 美术拆解, read screen-art-department-breakdown/SKILL.md completely and run
screen-art-department-breakdown/scripts/art_department_breakdown.py.
Do not rely on the generated draft blindly. Check scene count, malformed
headings, role columns, groups, props, costume notes, and workbook sheets before
returning the final .xlsx path.
```

## Direct Commands

Create a short-drama schedule table:

```bash
python3 duanju-schedule-table/scripts/shunchangbiao_builder.py script.docx \
  --out-dir output \
  --dump-json output/scenes.json
```

Create an art department workbook:

```bash
python3 screen-art-department-breakdown/scripts/art_department_breakdown.py script.docx \
  --out-dir output \
  --dump-json output/scenes.json
```

Create an AI live-action asset workbook after preparing the reviewed JSON:

```bash
node ai-live-action-asset-breakdown/scripts/build_asset_workbook.mjs \
  --input assets.json \
  --output output/asset-inventory.xlsx \
  --preview-dir output/previews
```

## Smart Defaults

- Prefer `.docx` input when available.
- Convert `.doc` to `.docx` before parsing.
- Use `pdfplumber` only for text-based PDFs.
- Treat scanned PDFs as blocked until OCR or Word/text export is available.
- Keep source scene order unless the user gives a separate shooting order.
- For AI asset inventories, list only first appearances, preserve meaningful
  visual variants, and keep set dressing out of the actor-used prop category.
- Put generated files in an output folder, not inside the skill folders.

## What To Tell The User

When done, Hermes should report:

```text
Generated: /absolute/path/to/workbook.xlsx
Scenes: <count>
Checks: passed/failed
Notes: <malformed headings, inferred day/night, skipped OCR, or other assumptions>
```
