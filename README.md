# LynnAI Skills

LynnAI Skills is Lynn's personal collection of Codex skills. Each top-level
folder is a standalone skill that can be copied or symlinked into
`~/.codex/skills`.

## Skills

| Skill | What it does |
| --- | --- |
| `duanju-schedule-table` | Builds polished Excel 顺场表 / 顺畅表 / 拍摄顺序表 from short-drama scripts. |
| `screen-art-department-breakdown` | Builds art department prep workbooks from film, TV, or short-drama scripts. |

## Layout

Keep skill folders at the repository root:

```text
skill-name/
├── SKILL.md
├── agents/
│   └── openai.yaml
└── scripts/
    └── helper.py
```

Repository-level helpers live outside skill folders:

```text
scripts/
└── validate_skills.py
```

## Local Install

Copy or symlink the skills you want Codex to discover:

```bash
mkdir -p "$HOME/.codex/skills"
ln -s "$PWD/duanju-schedule-table" "$HOME/.codex/skills/duanju-schedule-table"
ln -s "$PWD/screen-art-department-breakdown" "$HOME/.codex/skills/screen-art-department-breakdown"
```

If a skill already exists in `~/.codex/skills`, remove or replace the old link
first.

## Validate

Run the repository validator before committing:

```bash
python3 scripts/validate_skills.py .
```

To also syntax-check bundled Python scripts:

```bash
python3 scripts/validate_skills.py . --compile-scripts
```

