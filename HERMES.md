# Hermes Usage

Three skills in this repository are packaged for a computer that has Hermes Agent
but does not have Codex:

- `duanju-schedule-table`
- `ai-live-action-asset-breakdown`
- `screen-art-department-breakdown`

## Install

Run these commands in a terminal. Each command installs the complete skill
folder, including its scripts, references, and dependency manifest.

```bash
hermes skills install lynnhly1115-oss/lynnai-skills/duanju-schedule-table --yes
hermes skills install lynnhly1115-oss/lynnai-skills/ai-live-action-asset-breakdown --yes
hermes skills install lynnhly1115-oss/lynnai-skills/screen-art-department-breakdown --yes
```

Restart Hermes Desktop after installation. The skills can then trigger from a
normal request, or the user can invoke them explicitly:

```text
Use $duanju-schedule-table to turn this complete short-drama script into a checked Excel 顺场表.
```

```text
Use $ai-live-action-asset-breakdown to read this complete script and create a deduplicated Excel asset inventory with image placeholders.
```

```text
Use $screen-art-department-breakdown to turn this complete film, TV, or short-drama script into a checked folder containing four separate art-department Excel files.
```

## First Run

All three skills use Python. If `python-docx`, `openpyxl`, or `pdfplumber` is missing,
the bundled launcher downloads the packages from PyPI into the user's LynnAI
cache. It does not modify Hermes' own Python environment or the system Python.
An internet connection is therefore required the first time a skill runs with
missing packages. The skills share the same dependency cache when their
requirements match.

Supported source formats:

- `.docx` and `.docm`
- legacy `.doc` through macOS `textutil` or LibreOffice
- text-based `.pdf`
- `.txt` and `.md` for AI asset extraction

Image-only PDFs still require OCR or a Word/text export.

## Direct Commands

Hermes normally runs the scripts automatically. For manual diagnosis, resolve
the installed skill directory shown by Hermes and replace `<skill-dir>` below.
Use `python3` on macOS/Linux or `python` on Windows.

Create a screenplay schedule draft (one workbook with `顺场表` and `分场表`):

```bash
python3 "<skill-dir>/scripts/shunchangbiao_builder.py" script.docx \
  --out-dir work \
  --dump-json work/scenes.json \
  --report work/validation.json
```

After reviewing all scenes and their main locations, run the strict build.
It creates both sheets from one data set; cast categories follow the source
without a roster-confirmation question:

```bash
python3 "<skill-dir>/scripts/shunchangbiao_builder.py" script.docx \
  --out-dir output \
  --config reviewed.json \
  --report output/validation.json \
  --strict
```

Extract an AI live-action script and build the reviewed asset workbook:

```bash
python3 "<skill-dir>/scripts/build_asset_workbook.py" extract script.docx \
  --output work/script.txt

python3 "<skill-dir>/scripts/build_asset_workbook.py" build \
  --input work/assets.json \
  --output "output/《剧本名》_AI资产统计表.xlsx" \
  --preview-dir work/previews
```

Create a checked four-file art-department folder and an external scene dump:

```bash
python3 "<skill-dir>/scripts/art_department_breakdown.py" script.docx \
  --out-dir output \
  --folder-name "项目名-美术拆解" \
  --dump-json work/art-scenes.json
```

The art-department helper validates the fixed four-file contract, scene
coverage, and shared location/prop naming. Hermes must still complete the
semantic source review and visually inspect all four final workbooks before
delivery.

The asset build performs a structural workbook inspection. The optional preview
is a local HTML file for visual checking; only the final `.xlsx` is delivered.

## Boundaries

- Draft parsing is not final review. Check every scene or asset against the
  complete source before delivery.
- Do not silently drop malformed headings, missing day/night fields, variants,
  or scanned pages.
- Keep generated files outside the installed skill directory.
