# LynnAI Skills

Reusable Codex and Hermes skills for film, television, and short-drama production workflows.

This repository packages reusable `SKILL.md`-based agents and helper scripts
that turn screenplay files into production-ready Excel workbooks. The current
focus is short-drama scheduling, AI live-action asset inventories, and checked
art-department prep for scripts of any genre or production format.

![Skills](https://img.shields.io/badge/skills-3-1f4e5f)
![Codex](https://img.shields.io/badge/Codex-skills-2b6cb0)
![Hermes](https://img.shields.io/badge/Hermes-3%20native%20skills-7a2e2e)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776ab)
![License](https://img.shields.io/badge/license-MIT-green)

## What's Inside

| Skill | Output | Best For |
| --- | --- | --- |
| [`duanju-schedule-table`](duanju-schedule-table/SKILL.md) | One strictly checked `顺场表` Excel workbook per script | 短剧顺场表、全部有名角色分列、特约/群演、VO/OS、戏用车、同主场景连场标记 |
| [`ai-live-action-asset-breakdown`](ai-live-action-asset-breakdown/SKILL.md) | One `资产统计表` workbook per project, with reserved image cells | AI真人剧资产统计、按集累计去重、人物/场景/戏用道具、资产变体 |
| [`screen-art-department-breakdown`](screen-art-department-breakdown/SKILL.md) | One checked folder with four separate workbooks: `顺场表`, `分场表`, `场景表`, `置景道具执行表` | 电影、电视剧、短剧的美术指导、场景拆解、置景道具、采买制作、特殊道具整理 |

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
ln -s "$PWD/ai-live-action-asset-breakdown" "$HOME/.codex/skills/ai-live-action-asset-breakdown"
ln -s "$PWD/screen-art-department-breakdown" "$HOME/.codex/skills/screen-art-department-breakdown"
```

Restart Codex after installing. Then you can ask:

```text
Use $duanju-schedule-table to turn this short-drama script into a checked Excel 顺场表.
```

Or create an AI live-action asset inventory:

```text
Use $ai-live-action-asset-breakdown to read this complete script and create a deduplicated Excel asset inventory with image placeholders.
```

## Hermes Agent

All three skills are self-contained for Hermes Agent. Install their complete
bundles directly from GitHub:

```bash
hermes skills install lynnhly1115-oss/lynnai-skills/duanju-schedule-table --yes
hermes skills install lynnhly1115-oss/lynnai-skills/ai-live-action-asset-breakdown --yes
hermes skills install lynnhly1115-oss/lynnai-skills/screen-art-department-breakdown --yes
```

Restart Hermes Desktop after installation. Missing Python packages are installed
on first use into a separate LynnAI cache, without changing Hermes itself. See
[HERMES.md](HERMES.md) for supported inputs, direct diagnostics, and boundaries.

## Other Agents

These workflows are not limited to Codex. Any agent that can read files and run
the required Python helper can use the skills manually:

- Read [AGENTS.md](AGENTS.md) for generic agent instructions.
- Read [HERMES.md](HERMES.md) for a Hermes-ready prompt and command examples.
- Then read the target skill's `SKILL.md` before running a helper script.

## Short-Drama Schedule Tables

`duanju-schedule-table` creates a clean Excel file with one sheet named
`顺场表`. It keeps scenes in script order and fills the production columns that
short-drama teams usually need:

- `拍摄顺序`
- `实际拍摄场地`
- `提示`（如雨、烟机、闪回、武戏、AI做）
- `场次`
- `剧本中场景`
- `拍摄内容`
- `日/夜`
- `内/外`
- `页数`
- one column per named plot character, with childhood versions kept separate
- `特约演员`
- `群演`
- `VO`
- `OS`
- `梳化服提示`
- `道具提示`
- `戏用车`
- `备注/特殊道具/时间`

The skill handles common messy script formatting, including missing spaces,
day/night and interior/exterior markers written in different positions,
embedded flashbacks, montage headings, group characters, and practical prop
notes.

### 默认制作规则

- **演员自动分列**：读完整剧本后，所有有姓名的剧情角色各占一列，童年角色另列，不询问演员范围。用户主动指定名单时按用户要求执行。
- **特约与群演**：其他无姓名小角色，只要有个人台词或单独对手戏就归特约，不看场次多少；无个人台词或单独对手戏的背景人群归群演。
- **连场黑框**：必须同时满足“同一主场景”和“剧情连续”。同一宅院的客厅、走廊、附属门口可相连；换到警局等另一个主场景必须断框。只框场次列，组内横线保持细线。
- **页数**：从每场正文起按原排版计行，每3行计0.1，不足3行向上取整；1–3行为0.1，28–30行为1.0。
- **VO / OS**：VO为当前现场时空之外的旁白/内心独白；OS为当前场景空间内、镜头外发生的现场对白。
- **提示与车辆**：提示写简短拍摄需求；戏用车辆单列。所有内容均须以正文或已确定制作要求为依据。

Upgrading an existing reviewed config: add `body_line_count`, `hint`,
`on_camera`, `special_actors`, `vo`, `os`, and `vehicles` to every scene after
source review; explicitly fill `roles` and `continuity_groups`. Every continuity
group member also needs the same reviewed `main_location`. See the complete
[review config schema](duanju-schedule-table/references/review-config.md).
The helper validates recorded evidence; it does not replace the agent's full
script review or automatically establish cast identity and plot continuity.

Generate a draft parse for review:

```bash
python3 duanju-schedule-table/scripts/shunchangbiao_builder.py script.docx \
  --out-dir work \
  --dump-json work/scenes.json \
  --report work/validation.json
```

After reviewing every scene with the bundled config schema, run the strict build:

```bash
python3 duanju-schedule-table/scripts/shunchangbiao_builder.py script.docx \
  --out-dir output \
  --config reviewed.json \
  --report output/validation.json \
  --strict
```

## AI Live-Action Asset Inventories

`ai-live-action-asset-breakdown` reads a complete film, TV, or short-drama
script and creates one single-sheet asset inventory. It lists only the first
visible appearance of each character, playable scene, and actor-used prop.
Meaningful visual changes such as clothing, age, disguise, injury, damage,
day/night, screen content, or activation are recorded as variants.

The workbook keeps one sheet named `资产统计表`, groups assets by episode and
category, adds formula-driven subtotals and a cumulative total, and reserves
three image cells for characters and props or six for scenes.

Read the skill and its JSON schema before running the Python builder. The same
builder runs in Codex and Hermes, performs a structural workbook inspection,
and can create a local HTML preview for visual checking.

## Art Department Workbooks

`screen-art-department-breakdown` is for art-department preparation across
film, television, and short drama. It learns the current project's vocabulary
instead of assuming a fixed genre, then creates four separate checked files in
one folder:

- `顺场表`: one scene per row in script order
- `分场表`: scenes grouped by practical location
- `场景表`: one row per main location
- `置景道具执行表`: prep list for set dressing, hand props, purchase, fabrication,
  rental, borrowing, and special props

Run the helper script directly:

```bash
python3 screen-art-department-breakdown/scripts/art_department_breakdown.py script.docx \
  --out-dir output \
  --folder-name "项目名-美术拆解" \
  --dump-json work/art-scenes.json
```

The helper rejects silent scene-heading drops, checks scene coverage and
cross-file location/prop naming, reopens every workbook, and verifies the four
fixed file/header contracts. The Skill also requires semantic review against
the source and visual inspection of all four final workbooks.

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
├── AGENTS.md
├── HERMES.md
├── LICENSE
├── ai-live-action-asset-breakdown/
│   ├── SKILL.md
│   ├── agents/openai.yaml
│   ├── references/inventory-schema.md
│   ├── scripts/build_asset_workbook.py
│   ├── scripts/dependency_bootstrap.py
│   └── scripts/requirements.txt
├── duanju-schedule-table/
│   ├── SKILL.md
│   ├── agents/openai.yaml
│   ├── references/review-config.md
│   ├── scripts/shunchangbiao_builder.py
│   ├── scripts/dependency_bootstrap.py
│   ├── scripts/requirements.txt
│   └── tests/test_shunchangbiao_builder.py
├── screen-art-department-breakdown/
│   ├── SKILL.md
│   ├── agents/
│   │   └── openai.yaml
│   ├── scripts/
│   │   ├── art_department_breakdown.py
│   │   ├── dependency_bootstrap.py
│   │   └── requirements.txt
│   └── tests/
│       └── test_art_department_breakdown.py
└── scripts/
    └── validate_skills.py
```

## Notes For Agents

- Always read the relevant `SKILL.md` before running a helper script.
- Treat generated workbook drafts as drafts. Review scene counts, scene IDs,
  role columns, group characters, day/night, interior/exterior, props, costume
  hints, and notes before delivery.
- For AI asset inventories, read the complete script first, keep only first
  appearances, preserve meaningful variants, and exclude background dressing
  from the prop category.
- Keep generated `.xlsx`, `.json`, screenshots, and temporary script files out
  of the skill repository unless they are intentional fixtures.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for skill folder rules and the pre-commit
checklist.

## License

MIT. See [LICENSE](LICENSE).
