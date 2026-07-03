# LynnAI Skills

Personal Codex skills for Chinese short-drama production workflows.

This repository packages reusable `SKILL.md`-based agents and helper scripts
that turn screenplay files into production-ready Excel workbooks. The current
focus is short-drama scheduling, scene ordering, and art-department prep.

![Skills](https://img.shields.io/badge/skills-2-1f4e5f)
![Codex](https://img.shields.io/badge/Codex-skills-2b6cb0)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776ab)

## What's Inside

| Skill | Output | Best For |
| --- | --- | --- |
| [`duanju-schedule-table`](duanju-schedule-table/SKILL.md) | One polished `顺场表` Excel workbook per script | 短剧顺场表、拍摄顺序表、角色出场、日夜内外、群演、梳化服、道具提示 |
| [`screen-art-department-breakdown`](screen-art-department-breakdown/SKILL.md) | A four-sheet art prep workbook: `顺场表`, `分场表`, `场景表`, `置景道具执行表` | 美术指导、场景拆解、置景道具、采买制作、特殊道具整理 |

## Quick Start

Clone the repository:

```bash
git clone https://github.com/lynnhly1115-oss/lynnai-skills.git
cd lynnai-skills
```

Install the skills into Codex by symlinking them into `~/.codex/skills`:

```bash
mkdir -p "$HOME/.codex/skills"
ln -s "$PWD/duanju-schedule-table" "$HOME/.codex/skills/duanju-schedule-table"
ln -s "$PWD/screen-art-department-breakdown" "$HOME/.codex/skills/screen-art-department-breakdown"
```

Restart Codex after installing. Then you can ask:

```text
Use $duanju-schedule-table to turn this short-drama script into a checked Excel 顺场表.
```

## Short-Drama Schedule Tables

`duanju-schedule-table` creates a clean Excel file with one sheet named
`顺场表`. It keeps scenes in script order and fills the production columns that
short-drama teams usually need:

- `拍摄顺序`
- `实际拍摄场地`
- `场次`
- `剧本中场景`
- `拍摄内容`
- `日/夜`
- `内/外`
- `页数`
- one column per role
- `群演`
- `梳化服提示`
- `道具提示`
- `备注/特殊道具/时间`

The skill handles common messy script formatting, including missing spaces,
day/night and interior/exterior markers written in different positions,
embedded flashbacks, montage headings, group characters, and practical prop
notes.

Run the helper script directly:

```bash
python3 duanju-schedule-table/scripts/shunchangbiao_builder.py script.docx --out-dir output
```

Useful options:

```bash
python3 duanju-schedule-table/scripts/shunchangbiao_builder.py script.docx \
  --out-dir output \
  --dump-json output/scenes.json
```

## Art Department Workbooks

`screen-art-department-breakdown` is for production prep beyond a basic
schedule table. It creates four sheets:

- `顺场表`: one scene per row in script order
- `分场表`: scenes grouped by practical location
- `场景表`: one row per main location
- `置景道具执行表`: prep list for set dressing, hand props, purchase, fabrication,
  rental, borrowing, and special props

Run the helper script directly:

```bash
python3 screen-art-department-breakdown/scripts/art_department_breakdown.py script.docx --out-dir output
```

## Supported Inputs

- `.docx`
- `.doc` after conversion to `.docx`
- text-based `.pdf`

Scanned or image-only PDFs need OCR or a Word/text export first. For best
results, use scripts with explicit scene headings and `人物` lines.

## Validate

Check skill structure:

```bash
python3 scripts/validate_skills.py .
```

Check structure and compile bundled Python scripts:

```bash
python3 scripts/validate_skills.py . --compile-scripts
```

## Repository Layout

```text
lynnai-skills/
├── duanju-schedule-table/
│   ├── SKILL.md
│   ├── agents/
│   │   └── openai.yaml
│   └── scripts/
│       └── shunchangbiao_builder.py
├── screen-art-department-breakdown/
│   ├── SKILL.md
│   ├── agents/
│   │   └── openai.yaml
│   └── scripts/
│       └── art_department_breakdown.py
└── scripts/
    └── validate_skills.py
```

## Notes For Agents

- Always read the relevant `SKILL.md` before running a helper script.
- Treat generated workbook drafts as drafts. Review scene counts, scene IDs,
  role columns, group characters, day/night, interior/exterior, props, costume
  hints, and notes before delivery.
- Keep generated `.xlsx`, `.json`, screenshots, and temporary script files out
  of the skill repository unless they are intentional fixtures.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for skill folder rules and the pre-commit
checklist.

## License

No license file has been added yet. Until a license is added, treat this as a
public source-available repository rather than a formally open-source package.
